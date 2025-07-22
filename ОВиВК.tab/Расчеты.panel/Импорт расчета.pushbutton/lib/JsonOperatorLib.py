#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr
import glob
import re
import sys
import json
import os
import ctypes
import codecs
from datetime import datetime, timedelta
from System import Environment
from collections import defaultdict
from System.Collections.Generic import List

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import dosymep
from pyrevit import forms, revit, script, HOST_APP, EXEC_PARAMS
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from dosymep_libs.bim4everyone import *

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)


class JsonAngleOperator:
    def __init__(self, doc, uiapp):
        self.doc = doc
        self.uiapp = uiapp

    def get_document_path(self, filename):
        """
        Возвращает путь к config.json в директории проекта.
        Создаёт директорию и файл при необходимости.

        Args:
            filename (str): Имя файла ('config.json' или 'version.json')

        Returns:
            str: Полный путь к файлу config.json.
        """
        plugin_name = 'Импорт из Audytor'
        version_number = self.uiapp.VersionNumber
        project_name = self.get_project_name()

        # Путь до папки
        base_dir = os.path.join(version_number, plugin_name, project_name)
        my_documents_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        full_dir_path = os.path.join(my_documents_path, 'dosymep', base_dir)

        # Создаём директорию при необходимости
        if not os.path.exists(full_dir_path):
            os.makedirs(full_dir_path)

        # Путь к config.json
        file_path = os.path.join(full_dir_path, filename)

        # Создаём файл, если его нет
        if not os.path.isfile(file_path):
            with codecs.open(file_path, 'w', encoding='utf-8') as f:
                json.dump(0.0, f, ensure_ascii=False)

        return file_path

    def send_json_data(self, data, filename):
        """
        Отправляет одно числовое значение в JSON файл.

        Args:
            data (float): Значение, которое нужно сохранить в JSON файл.
            filename (str): Имя файла ('config.json' или 'version.json')
        """
        new_file_path = self.get_document_path(filename)

        with codecs.open(new_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(data, json_file, ensure_ascii=False)


    def get_json_data(self, filename):
        """
        Получает данные из JSON файла.

        Args:
            filename (str): Имя файла ('config.json' или 'version.json')

        Returns:
            float: Значение из JSON-файла.
        """

        json_path = self.get_document_path(filename)

        if not os.path.isfile(json_path):
            return 0

        with codecs.open(json_path, 'r', encoding='utf-8') as json_file:
            value = json.load(json_file)
        return value

    def get_project_name(self):
        """
        Возвращает имя проекта.

        Returns:
            str: Имя проекта.
        """
        username = __revit__.Application.Username
        title = self.doc.Title
        username_upper = username.upper()
        title_upper = title.upper()

        if username_upper in title_upper:
            project_name = title_upper.replace('_' + username_upper, '').strip()
        else:
            project_name = title

        return project_name
