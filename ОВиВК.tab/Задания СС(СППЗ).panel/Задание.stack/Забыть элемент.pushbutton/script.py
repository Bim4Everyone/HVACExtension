#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import glob
import re
import sys
import json
import os
import ctypes
import codecs
from low_voltage_task_class_lib import JsonOperator, EditedReport, LowVoltageSystemData
from vacancy_numbers import read_vacancies, write_vacancies, vacancy_from_data, vacancy_key

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI.Selection import ISelectionFilter

from dosymep_libs.bim4everyone import *
from System.Collections.Generic import List

from datetime import datetime, timedelta
from System import Environment

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__.Application

categories = [BuiltInCategory.OST_MechanicalEquipment,
              BuiltInCategory.OST_DuctAccessory,
              BuiltInCategory.OST_PipeAccessory]

forget_all_method = "Элемента не было в задании СС(СППЗ)"
forget_id_method = "Элемент уже был в задании СС(СППЗ)"
TASK_SS_PARAM = SharedParamsConfig.Instance.VISTaskSSMark
DATE_SS_PARAM = SharedParamsConfig.Instance.VISTaskSSDate

operator = JsonOperator(doc, uiapp)

class VISElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        if element.InAnyCategory(categories):
            return True

        return False

    def AllowReference(self, reference, position):
        return True


def get_pre_selected():
    """
    Получение заранее выбранных элементов.
    """

    elements = [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]

    filter_result = []
    for element in elements:
        cat = element.Category
        if cat is None:
            continue

        if element.InAnyCategory(categories):
            filter_result.append(element)

    return filter_result


def get_selected():
    """
    Выделение воздуховодов, труб, арматуры и оборудования.
    """

    elements = [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]

    filter_result = []
    for element in elements:
        cat = element.Category
        if cat is None:
            continue

        if element.InAnyCategory(categories):
            filter_result.append(element)

    if len(filter_result) != 0:
        return filter_result

    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            VISElementsFilter(),
            "Выберите воздуховоды или трубы"
        )
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        sys.exit()

    elements = [doc.GetElement(r) for r in references]

    if not elements:
        sys.exit()

    return elements


def get_selected_mode():
    method = forms.alert("Выберите метод",
                      options=[forget_all_method,
                               forget_id_method])

    if method is False:
        script.exit()

    return method

def forget_elements(json_data, elements, method):
    if method == forget_all_method:
        selected_ids = set(element.Id.GetIdValue() for element in elements)
        return [
            data for data in json_data
            if abs(data.id.GetIdValue()) not in selected_ids
        ]

    if method == forget_id_method:
        deletion_date = operator.get_utc_date()
        for element in elements:
            for data in json_data:
                if element.Id == data.id:
                    data.id = ElementId(-abs(data.id.GetIdValue()))
                    data.deletion_date = deletion_date

    return json_data


def add_vacancies(json_data, elements, folder):
    """
    Сохраняет номера выбранных элементов как свободные для повторной выдачи.

    Учитывает обычные записи и записи с отрицательным ID из истории.
    Номера НО и оборудования группируются раздельно по типу и этажу, НЗ — по системе
    и этажу. Повторяющиеся номера добавляются один раз и сортируются.
    Данные записываются в файл «Вакантные номера.json» рядом с заданиями.

    Args:
        json_data (list[LowVoltageSystemData]): Записи заданий, из которых
            нужно получить номера выбранных элементов.
        elements (list[Element]): Выбранные элементы, которые не выдавались
            в задании и номера которых можно освободить.
        folder (str): Путь к папке проекта с JSON-файлами заданий.

    Raises:
        ValueError: Не удалось определить номер марки.
    """
    vacancies = read_vacancies(folder)
    existing = set(vacancy_key(item) for item in vacancies)
    selected_ids = set(element.Id.GetIdValue() for element in elements)
    for data in json_data:
        if abs(data.id.GetIdValue()) not in selected_ids:
            continue
        vacancy = vacancy_from_data(data)
        key = vacancy_key(vacancy)
        if key not in existing:
            vacancies.append(vacancy)
            existing.add(key)

    write_vacancies(folder, vacancies)


def clear_element_task_params(elements):
    with revit.Transaction("BIM: Забыть элемент"):
        for element in elements:
            element.SetParamValue(TASK_SS_PARAM, "")
            element.SetParamValue(DATE_SS_PARAM, "")


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

    file_folder_path, is_path_local = operator.get_document_path()

    if is_path_local:
        operator.show_local_path(file_folder_path)

    method = get_selected_mode()

    elements = get_pre_selected() or get_selected()

    old_json_file_path = operator.get_json_file_path(file_folder_path, is_today=False)
    old_json_data = operator.get_json_data(file_folder_path, is_today=False) or []

    today_json_file_path = operator.get_json_file_path(file_folder_path)
    today_json_data = []
    if today_json_file_path and operator.get_utc_date() in os.path.basename(today_json_file_path):
        today_json_data = operator.get_json_data(file_folder_path) or []

    if not old_json_data and not today_json_data:
        forms.alert("Данные заданий не были обнаружены", "Ошибка", exitscript=True)

    operator.create_backup(file_folder_path)

    if method == forget_all_method:
        add_vacancies(old_json_data + today_json_data, elements, file_folder_path)

    final_old_json = forget_elements(old_json_data, elements, method)
    if old_json_file_path:
        operator.write_json_data(final_old_json, old_json_file_path)
    """
    Удаляем или помечаем записи выбранных элементов в последнем прошлом
    и сегодняшнем файлах задания
    """
    if today_json_data:
        final_today_json = forget_elements(today_json_data, elements, method)
        operator.write_json_data(final_today_json, today_json_file_path)

    clear_element_task_params(elements)

script_execute()
