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


class EditedReport:
    edited_reports = []
    status_report = None
    edited_report = None

    def __init__(self, doc):
        self.doc = doc

    def get_element_editor_name(self, element):
        """
        Возвращает имя пользователя, занявшего элемент, или None.

        Args:
            element (Element): Элемент для проверки.

        Returns:
            str или None: Имя пользователя или None, если элемент не занят.
        """
        user_name = __revit__.Application.Username
        edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
        if edited_by is None:
            return None

        if edited_by.lower() == user_name.lower():
            return None
        return edited_by

    def is_elemet_edited(self, element):
        """
        Проверяет, заняты ли элементы другими пользователями.

        Args:
            elements (list): Список элементов для проверки.
        """
        update_status = WorksharingUtils.GetModelUpdatesStatus(self.doc, element.Id)

        if update_status == ModelUpdatesStatus.UpdatedInCentral:
            self.status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию. "

        name = self.get_element_editor_name(element)
        if name is not None and name not in self.edited_reports:
            self.edited_reports.append(name)

        if name is not None or update_status == ModelUpdatesStatus.UpdatedInCentral:
            return True

        return False

    def show_report(self):
        if len(self.edited_reports) > 0:
            self.edited_report = (
                "Работа не может быть продолжена. Часть элементов занята пользователями: {}".format(
                    ", ".join(self.edited_reports)
                )
            )

        if self.edited_report is not None or self.status_report is not None:
            report_message = ''
            if self.status_report is not None:
                report_message += self.status_report
            if self.edited_report is not None:
                if report_message:
                    report_message += '\n'
                report_message += self.edited_report

            forms.alert(report_message, "Ошибка", exitscript=True)

class LowVoltageSystemData:
    def __init__(self, id, creation_date=None, equipment_base_name=None,
                 autor_name=__revit__.Application.Username, json_name=None,
                 deletion_date=None, element=None, closed_algorithm=None, open_algorithm=None):
        """
        Инициализация объекта LowVoltageSystemData.

        Args:
            id (ElementId): Идентификатор элемента.
            creation_date (str, optional): Дата создания.
            equipment_base_name (str, optional): Базовое имя клапана.
            autor_name (str, optional): Имя автора. По умолчанию имя текущего пользователя.
            json_name (str, optional): Имя JSON файла.
            deletion_date (str, optional): Дата удаления.
        """
        self.id = id
        self.equipment_base_name = equipment_base_name
        self.autor_name = autor_name
        self.json_name = json_name
        self.creation_date = creation_date
        self.deletion_date = deletion_date
        self.element = element
        self.closed_algorithm = closed_algorithm
        self.open_algorithm = open_algorithm

    def to_dict(self):
        """
        Преобразует объект в словарь.

        Returns:
            dict: Словарь с данными объекта.
        """
        return {
            "id": str(self.id),
            "equipment_base_name": self.equipment_base_name,
            "autor_name": self.autor_name,
            "json_name": self.json_name,
            "creation_date": self.creation_date,
            "deletion_date": self.deletion_date,
            "closed_algorithm": self.closed_algorithm,
            "open_algorithm": self.open_algorithm
        }

    def insert(self, doc, time):
        """
        Вставляет данные в элемент документа.
        """
        element = doc.GetElement(self.id)

        if element is not None:
            element.SetParamValue(SharedParamsConfig.Instance.VISTaskSSMark, self.json_name)
            element.SetParamValue(SharedParamsConfig.Instance.VISTaskSSDate, self.creation_date)
        else:
            if not self.deletion_date and self.creation_date != time:
                self.deletion_date = time

class JsonOperator:
    def __init__(self, doc, uiapp):
        self.doc = doc
        self.uiapp = uiapp

    def show_dialog(self, instr, content=''):
        dialog = TaskDialog("Внимание")
        dialog.MainInstruction = instr
        dialog.MainContent = content
        dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No

        result = dialog.Show()

        if result == TaskDialogResult.Yes:
            return True
        elif result == TaskDialogResult.No:
            return False

    def get_document_path(self):
        """
        Возвращает путь к документу.

        Returns:
            str: Путь к документу.
        """

        plugin_name = 'Задания СС(СППЗ)'
        version_number = self.uiapp.VersionNumber
        project_name = self.get_project_name()
        base_root = os.path.join(version_number, plugin_name, project_name)

        network_directory = os.path.join(
            "W:/Проектный институт/Отд.стандарт.BIM и RD/BIM-Ресурсы/"
            "5-Надстройки/Bim4Everyone/A101/"
        )
        is_local = False # Флаг для открытия папки с данными в моих документах

        my_documents_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        network_path = os.path.join(network_directory, base_root)
        local_path = os.path.join(my_documents_path, 'dosymep', base_root)

        if not (os.path.exists(network_directory) and os.access(network_directory, os.R_OK | os.W_OK)):
            path = local_path
            is_local = True
        else:
            path = network_path

        if not os.path.exists(path):
            os.makedirs(path)

        return path, is_local

    def show_local_path(self, local_path):
        report = (
            'Нет доступа к сетевому диску. Файлы задания обрабатываются из папки: {} \n'
            'Открыть папку с файлами?'
        ).format(local_path)

        if self.show_dialog(report):
            os.startfile(local_path)

    def send_json_data(self, data, new_file_path):
        """
        Отправляет данные в JSON файл.

        Args:
            data (list): Список объектов LowVoltageSystemData.
            new_file_path (str): Путь к новому JSON файлу.
        """
        project_name = self.get_project_name()
        time = self.get_utc_date()
        new_file_path = new_file_path + "/СС(СППЗ)_" + project_name + "_" + time + ".json"

        data_dicts = [item.to_dict() for item in data]

        with codecs.open(new_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(data_dicts, json_file, ensure_ascii=False, indent=4)

    def get_json_data(self, project_path):
        """
        Получает данные из JSON файла.

        Args:
            project_path (str): Путь к папке проекта.

        Returns:
            list: Список объектов LowVoltageSystemData.
        """

        json_files = glob.glob(os.path.join(project_path, "*.json"))
        if not json_files:
            return {}

        latest_file = max(json_files, key=os.path.getmtime)

        if latest_file is not None and os.path.exists(latest_file):
            with codecs.open(latest_file, 'r', encoding='utf-8') as json_file:
                existing_data = json.load(json_file)

                return [
                    LowVoltageSystemData(
                        id=ElementId(int(item["id"])),
                        creation_date=item["creation_date"],
                        equipment_base_name=item["equipment_base_name"],
                        autor_name=item["autor_name"],
                        json_name=item["json_name"],
                        deletion_date=item["deletion_date"],
                        closed_algorithm=item["closed_algorithm"],
                        open_algorithm=item["open_algorithm"]
                    )
                    for item in existing_data
                ]
        return []

    def get_utc_date(self):
        """
        Возвращает текущую дату в часовом поясе Москвы (UTC+3).

        Returns:
            str: Текущая дата в формате "YYYY-MM-DD".
        """
        utc_time = datetime.utcnow()
        formatted_time = utc_time.strftime("%Y-%m-%d")

        return formatted_time

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

    def create_folder_if_not_exist(self, project_path):
        """
        Создает папку, если она не существует.

        Args:
            project_path (str): Путь к папке проекта.
        """
        if not os.path.exists(project_path):
            os.makedirs(project_path)