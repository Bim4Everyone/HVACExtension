#! /usr/bin/env python
# -*- coding: utf-8 -*-

from itertools import count

import clr
from System import DateTime

from unmodeling_class_library import UnmodelingFactory, MaterialCalculator, RowOfSpecification

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import glob

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from pyrevit import forms
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict
from rpw.ui.forms import SelectFromList
from unmodeling_class_library import *
from dosymep_libs.bim4everyone import *
import sys
from rpw.ui.forms import select_file, SelectFromList
import json
import os
import codecs
from datetime import datetime, timedelta
from System.Windows.Forms import FolderBrowserDialog, DialogResult

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument



def get_json_data(project_path):
    # Находим все JSON-файлы в директории
    json_files = glob.glob(os.path.join(project_path, "*.json"))
    if not json_files:
        return {}

    # Находим файл с самым поздним временем модификации
    latest_file = max(json_files, key=os.path.getmtime)
    old_data = []
    if os.path.exists(latest_file):
        with codecs.open(latest_file, 'r', encoding='utf-8') as json_file:
            name_mapping = json.load(json_file)
            name_mapping = {item['id']: item['name'] for item in name_mapping}

    return name_mapping

def send_json_data(data, new_file_path):
    project_name = doc.Title

    user_name = __revit__.Application.Username

    if user_name not in project_name:
        project_name = project_name + "_" + user_name

    time = get_moscow_time()

    new_file_path = new_file_path + "\Задание СПА_" + project_name + "_" + time + ".json"

    # Преобразование списка объектов в список словарей
    data_dicts = [item.to_dict() for item in data]

    # Запись в JSON
    with codecs.open(new_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(data_dicts, json_file, ensure_ascii=False, indent=4)


def get_old_data(old_file_path, data):
    # Проверяем, существует ли файл
    if old_file_path is not None and os.path.exists(old_file_path):

        # Читаем существующие данные
        with codecs.open(old_file_path, 'r', encoding='utf-8') as json_file:
            existing_data = json.load(json_file)

        # Объединяем существующие данные с новыми данными
        # Предполагается, что данные являются списками или словарями
        if isinstance(existing_data, list) and isinstance(data, list):
            combined_data = existing_data + data
        elif isinstance(existing_data, dict) and isinstance(data, dict):
            combined_data = dict(existing_data.items() + data.items())
        else:
            raise ValueError("Data types are not compatible for merging")
    else:
        combined_data = data
    return combined_data

def get_moscow_time():

    # Получаем текущее время в UTC
    utc_time = datetime.utcnow()

    # Добавляем 3 часа для перехода в часовой пояс Москвы (UTC+3)
    moscow_time = utc_time + timedelta(hours=3)

    # Форматируем время
    formatted_time = moscow_time.strftime("%Y_%m_%d_%H_%M")

    return formatted_time

def select_folder():
    dialog = FolderBrowserDialog()
    dialog.Description = "Выберите папку для сохранения задания"
    dialog.ShowNewFolderButton = True

    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.SelectedPath
    else:
        return None

def split_valves_by_floors(valves):
    valves_by_floors = {}

    for valve in valves:
        floor_name = valve.GetParamValueOrDefault("ФОП_Этаж")
        system_name = valve.GetParamValueOrDefault("ФОП_ВИС_Имя системы")

        if floor_name is None or (system_name is None or system_name == "!Нет системы"):
            continue

        if floor_name not in valves_by_floors:
            valves_by_floors[floor_name] = []


        valve_base_name = system_name + "-" + floor_name

        valves_by_floors[floor_name].append(ValveData(valve.Id, valve_base_name= valve_base_name))

    return valves_by_floors

def create_open_numbers(valves):
    valves_by_floors = split_valves_by_floors(valves)

    result = []
    for floor_name, valve_data_instances in valves_by_floors.items():
        count = 0
        for valve_data in valve_data_instances:
            count += 1

            valve_data.json_name = "НО-" + valve_data.valve_base_name + "-" + str(count)

            result.append(valve_data)

    return result

def create_closed_numbers(valves):
    valves_by_floors = split_valves_by_floors(valves)
    result = []

    for floor_name, valve_data_instances in valves_by_floors.items():
        # Группировка valve_data_instances по valve_base_name
        valve_groups = {}
        for valve_data in valve_data_instances:
            if valve_data.valve_base_name not in valve_groups:
                valve_groups[valve_data.valve_base_name] = []
            valve_groups[valve_data.valve_base_name].append(valve_data)

        # проходим по экземплярам НЗ сгруппированных по базовому имени(одинаковая система и этаж)
        for valve_base_name, valve_data_instances_group in valve_groups.items():
            count = 0
            for valve_data in valve_data_instances_group:
                count += 1
                valve_data.json_name = "НЗ-" + valve_data.valve_base_name + "-" + str(count)

                result.append(valve_data)

    return result

def get_project_name():
    # Получаем имя пользователя
    username = user_name = __revit__.Application.Username

    # Получаем заголовок документа
    title = doc.Title

    # Переводим все буквы в верхний регистр
    username_upper = username.upper()
    title_upper = title.upper()

    # Проверяем, является ли имя пользователя частью заголовка
    if username_upper in title_upper:
        # Убираем имя пользователя и подчеркивание перед ним
        project_name = title_upper.replace('_' + username_upper, '').strip()
    else:
        # Если имя пользователя не является частью заголовка, возвращаем заголовок как есть
        project_name = title

    return project_name

def create_folder_if_not_exist(project_path):
    if not os.path.exists(project_path):
        os.makedirs(project_path)

def split_collection(equipment_collection, old_data):
    open_valves = []
    closed_valves = []
    equipment_elements = []
    edited_report = EditedReport()

    for element in equipment_collection:
        if edited_report.is_elemet_edited(element):
            continue

        if str(element.Id) in old_data:
            name = old_data[str(element.Id)]
            element.SetParamValue("ADSK_Позиция", name)
            # если айди элемента упомянут в задании то заново его не рассматриваем, сразу ставим имя оттуда
            continue

        if element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
            equipment_elements.append(element)

        if element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
            mark = element.GetSharedParamValueOrDefault("ФОП_ВИС_Марка")
            if mark is not None and mark != "":
                if "НО" in mark:
                    open_valves.append(element)
                if "НЗ" in mark:
                    closed_valves.append(element)

    edited_report.show_report()
    return open_valves, closed_valves, equipment_elements

class EditedReport:
    edited_reports = []
    status_report = ''
    edited_report = ''

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

        if edited_by.lower() in user_name.lower():
            return None
        return edited_by

    def is_elemet_edited(self, element):
        """
        Проверяет, заняты ли элементы другими пользователями.

        Args:
            elements (list): Список элементов для проверки.
        """

        update_status = WorksharingUtils.GetModelUpdatesStatus(doc, element.Id)

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
            self.edited_report = \
                ("Часть элементов спецификации занята пользователями: {}".format(", ".join(self.edited_reports)))

        if self.edited_report != '' or self.status_report != '':
            report_message = (self.status_report +
                              ('\n' if (self.edited_report and self.status_report) else '') +
                              self.edited_report)
            forms.alert(report_message, "Ошибка", exitscript=True)

class ValveData:
    id = ''
    valve_base_name = ''
    autor_name = ''
    json_name = ''
    creation_date = ''

    def __init__(self,
                 id,
                 valve_base_name = '',
                 autor_name = __revit__.Application.Username,
                 json_name = '',
                 creation_date = get_moscow_time()):

        self.id = id
        self.valve_base_name = valve_base_name
        self.autor_name = autor_name
        self.json_name = json_name
        self.creation_date = creation_date

    def to_dict(self):
        return {
            "id": str(self.id),
            "valve_base_name": self.valve_base_name,
            "autor_name": self.autor_name,
            "json_name": self.json_name,
            "creation_date": self.creation_date
        }

    def insert(self):
        valve = doc.GetElement(self.id)
        valve.SetParamValue("ADSK_Позиция", self.json_name)

GLOBAL_PATH_CONST = \
    "W:/Проектный институт/Отд.стандарт.BIM и RD/BIM-Ресурсы/5-Надстройки/Bim4Everyone/A101/MEP/EquipmentNumbering/"

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    project_name = get_project_name()
    project_path = GLOBAL_PATH_CONST + project_name

    # Если папки еще не существует - создаем
    create_folder_if_not_exist(project_path)

    # old_data = get_json_data(project_path)
    old_data = []

    equipment_collection = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_DuctAccessory) \
        .WhereElementIsNotElementType() \
        .ToElements()

    with revit.Transaction("BIM: Задание СС"):
        open_valves, closed_valves, equipment_elements = split_collection(equipment_collection, old_data)

        open_valves_data = create_open_numbers(open_valves)
        closed_valves_data = create_closed_numbers(closed_valves)
        equipment_elements_data = create_open_numbers(equipment_elements)

        if open_valves_data or closed_valves_data or equipment_elements_data:
            json_data = []

            json_data.extend(open_valves_data)
            json_data.extend(closed_valves_data)
            json_data.extend(equipment_elements_data)

            for element in json_data:
                element.insert()

            send_json_data(json_data, project_path)

script_execute()
