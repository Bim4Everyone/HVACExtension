#! /usr/bin/env python
# -*- coding: utf-8 -*-
from tarfile import TUEXEC

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import re
import glob
import os
import json
import codecs
from low_voltage_task_class_lib import JsonOperator, EditedReport, LowVoltageSystemData

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
from dosymep.Bim4Everyone.Templates import ProjectParameters
from collections import defaultdict

from dosymep_libs.bim4everyone import *

from System.Collections.Generic import List
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__.Application

def split_equipment_by_floors(equipment_elements, opened = False, closed = False):
    ''' Делим список элементов на словарь, ключ - номер этажа, значение - экземпляр LowVoltageSystemData '''
    equpment_by_floors = {}

    for equipment in equipment_elements:
        floor_name = get_floor_value(equipment)
        system_name = equipment.GetParamValueOrDefault(SYSTEM_PARAM)

        equipment_base_name = system_name + "-" + floor_name

        if floor_name not in equpment_by_floors:
            equpment_by_floors[floor_name] = []

        equpment_by_floors[floor_name].append(LowVoltageSystemData
            (
                equipment.Id,
                creation_date=operator.get_utc_date(),
                equipment_base_name= equipment_base_name,
                element=equipment,
                open_algorithm=opened,
                closed_algorithm=closed
            )
        )

    return equpment_by_floors

def get_floor_value(element):
    value = element.GetParamValueOrDefault(FLOOR_PARAM)

    value = value.strip()

    if value:
        value = value.split()[0]

    if value.startswith(("-", "П")):
        suffix = value[1:]
        if suffix.isdigit():
            value = "П" + suffix.zfill(2)

    if value.isdigit():
        value = value.zfill(2)

    return value

def get_value_if_para_not_empty(element, param):
    """ Возвращает значения параметра, если он не пустой. Если пустой - падаем. """
    value = element.GetParamValueOrDefault(param)
    if value is None or value == "!Нет системы" or value == "":
        report_text = (("У части отмеченного для задания оборудования не заполнен параметр {}. "
                  "Выполните полное обновление и повторите запуск скрипта. Пример элемента с ошибкой: {}")
                  .format(param.Name, element.Id))
        forms.alert(report_text, "Ошибка", exitscript=True)

    return value

def take_number(vacancies, kind, key, max_numbers, assigned):
    used = assigned.setdefault((kind, key), set())
    available = vacancies.setdefault(kind, {}).setdefault(key, [])
    while available:
        number = available.pop(0)
        if number not in used:
            used.add(number)
            max_numbers[key] = max(max_numbers.get(key, 0), number)
            return number
    number = max_numbers.get(key, 0) + 1
    while number in used:
        number += 1
    max_numbers[key] = number
    used.add(number)
    return number


def use_open_algorithm(open_valves, old_data, vacancies, max_numbers, assigned):
    ''' Создаем имена для списка элементов используя алгоритм для открытых клапанов от СППЗ '''
    open_valves_by_floors = split_equipment_by_floors(open_valves, opened=True)

    result = []
    for floor_name, open_valves_instances in open_valves_by_floors.items():
        for open_valve in open_valves_instances:
            if open_valve.element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
                base_name  = open_valve.equipment_base_name
            else:
                base_name = "НО-" + open_valve.equipment_base_name

            # Проверяем, существует ли ключ в словаре max_numbers
            count = take_number(vacancies, "open", floor_name, max_numbers, assigned)

            open_valve.json_name = base_name + "-" + str(count)

            result.append(open_valve)

    return result

def use_closed_algorithm(valves, old_data, vacancies, max_numbers, assigned):
    ''' Создаем имена для списка элементов используя алгоритм для закрытых клапанов от СППЗ '''
    valves_by_floors = split_equipment_by_floors(valves, closed=True)

    result = []

    for floor_name, valve_data_instances in valves_by_floors.items():
        # Группировка valve_data_instances по equipment_base_name
        valve_groups = {}
        for valve_data in valve_data_instances:
            if valve_data.equipment_base_name not in valve_groups:
                valve_groups[valve_data.equipment_base_name] = []
            valve_groups[valve_data.equipment_base_name].append(valve_data)

        # проходим по экземплярам НЗ сгруппированных по базовому имени(одинаковая система и этаж)
        for equipment_base_name, valve_data_instances_group in valve_groups.items():
            for valve_data in valve_data_instances_group:
                json_base_name = "НЗ-" + valve_data.equipment_base_name
                count = take_number(vacancies, "closed", json_base_name, max_numbers, assigned)
                valve_data.json_name = json_base_name + "-" + str(count)

                result.append(valve_data)

    return result

def split_collection(equipment_collection):
    ''' Делим список элементов на открытые клапана, закрытые и оборудование. Если один из элементов '''
    open_valves = []
    closed_valves = []
    equipment_elements = []

    for element in equipment_collection:

        if element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
            equipment_elements.append(element)

        if element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
            mark = get_value_if_para_not_empty(element, MARK_PARAM)

            if "НО" in mark:
                open_valves.append(element)
            if "НЗ" in mark:
                closed_valves.append(element)


    return open_valves, closed_valves, equipment_elements

class TaskSelectionFilter(ISelectionFilter):
    def __init__(self, categories):
        self.categories = categories

    def AllowElement(self, element):
        return element.Category is not None and any(
            element.Category.IsId(category) for category in self.categories)

    def AllowReference(self, reference, position):
        return False


def get_elements():
    ''' Забираем только выбранные элементы нужных категорий. '''
    categories = [
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PipeAccessory
    ]

    selected_ids = uidoc.Selection.GetElementIds()
    if selected_ids.Count:
        elements = [doc.GetElement(element_id) for element_id in selected_ids]
    else:
        try:
            references = uidoc.Selection.PickObjects(
                ObjectType.Element, TaskSelectionFilter(categories),
                "Выберите арматуру и оборудование для обновления задания и нажмите «Готово»")
        except OperationCanceledException:
            script.exit()
        elements = [doc.GetElement(reference.ElementId) for reference in references]
    elements = [element for element in elements
                if element is not None and element.Category is not None
                and any(element.Category.IsId(category) for category in categories)]
    if not elements:
        forms.alert("Среди выбранных элементов нет арматуры или оборудования.",
                    "Ошибка", exitscript=True)

    return elements

def print_missing_task_params(elements):
    output = script.get_output()
    rows = []
    type_rows = []
    reported_types = set()
    for element in elements:
        missing = []
        element_type = element.GetElementType()
        if element_type is None or element_type.GetParamValueOrDefault(CREATE_TASK_SS_PARAM) != 1:
            family_name = element_type.FamilyName if element_type is not None else "Тип не найден"
            type_name = (element_type.GetParamValueOrDefault(BuiltInParameter.SYMBOL_NAME_PARAM)
                         if element_type is not None else "")
            type_key = (family_name, type_name)
            if type_key not in reported_types:
                reported_types.add(type_key)
                type_rows.append([family_name, type_name, output.linkify(element.Id),
                                  "ФОП_ВИС_СС Добавить к заданию: не включён или отсутствует"])
        for param, name in ((SYSTEM_PARAM, "ФОП_ВИС_Имя системы"),
                            (FLOOR_PARAM, "ФОП_Этаж СМР")):
            value = element.GetParamValueOrDefault(param)
            if value is None or not value.strip() or value.strip() == "!Нет системы":
                missing.append(name + " (экземпляр): не заполнен")
        if missing:
            rows.append([output.linkify(element.Id), "<br>".join(missing)])
    if type_rows:
        output.print_table(table_data=type_rows,
                           title="Параметры типов: добавление к заданию СС(СППЗ)",
                           columns=["Семейство", "Типоразмер", "Выбрать пример", "Замечание"])
    if rows:
        output.print_table(table_data=rows,
                           title="Параметры экземпляров: задание СС(СППЗ)",
                           columns=["Выбрать элемент", "Замечания"])


def get_elements_to_objective(elements):
    ''' Фильтруем, у каких элементов модели стоит галочка для добавления в задание
    и параметры системы этажа имеют значение '''
    filtered_elements = []
    for element in elements:
        system_name = element.GetParamValueOrDefault(SYSTEM_PARAM)
        floor_name = element.GetParamValueOrDefault(FLOOR_PARAM)
        if system_name is None or system_name == '':
            continue
        if floor_name is None or floor_name == '':
            continue

        if element.GetParamValueOrDefault(CREATE_TASK_SS_PARAM) == 1:
            filtered_elements.append(element)
        else:
            element_type = element.GetElementType()
            if element_type.GetParamValueOrDefault(CREATE_TASK_SS_PARAM) == 1:
                filtered_elements.append(element)

    return  filtered_elements

def math_elements_to_old_data(elements, old_data):
    ''' Сравниваем список элементов с данными из json-файла, возвращая только те которых в нем нет '''
    new_elements = []

    for element in elements:
        if not any(element.Id == data.id for data in old_data):
            new_elements.append(element)

    return new_elements

def get_closed_max_numbers(old_data):
    ''' Вычисляем для каждого базового имени в json-файле, созданного по закрытому алгоритму,
    максимальный номер, с которого будет продолжаться нумерация '''
    # Словарь для хранения максимальных порядковых номеров
    max_numbers = defaultdict(int)

    for data in old_data:
        if data.open_algorithm:
            continue

        match = re.search(r'^(.*)-([^-]+)$', data.json_name)
        # условное выражение делит имя образца "НЗ-ДП3.4-2эт ДОО-1" на "НЗ-ДП3.4-2эт ДОО" и "1", отрезая последний дефис

        if match:
            base_name = match.group(1)
            number = int(match.group(2))
            if number > max_numbers[base_name]:
                max_numbers[base_name] = number

    return max_numbers

def get_open_max_numbers(old_data):
    ''' Вычисляем для каждого базового имени, созданного по открытому алгоритму,
    в json-файле максимальный номер, с которого будет продолжаться нумерация '''
    # Словарь для хранения максимальных порядковых номеров
    max_numbers = defaultdict(int)

    for data in old_data:
        if data.closed_algorithm:
            continue

        parts = data.json_name.split('-')

        floor_name = parts[-2]
        number = int(parts[-1])

        if number > max_numbers[floor_name]:
            max_numbers[floor_name] = number

    return max_numbers


def get_assigned_numbers(old_data):
    assigned = {}
    for data in old_data:
        match = re.match(r'^(.+)-(\d+)$', data.json_name or '')
        if not match:
            continue
        number = int(match.group(2))
        if data.closed_algorithm:
            kind, key = "closed", "НЗ-" + data.equipment_base_name
        elif data.open_algorithm:
            kind, key = "open", data.equipment_base_name.rsplit('-', 1)[-1]
        else:
            continue
        assigned.setdefault((kind, key), set()).add(number)
    return assigned

def check_edited_elements(elements):
    ''' Проверяем занят ли элемент '''
    edited_report = EditedReport(doc)

    for element in elements:
        edited_report.is_elemet_edited(element)

    edited_report.show_report()

def clear_param_false_values(elements, json_data):
    ''' Проверяем что в данных которые мы получили/будем загружать в json существуют айди оборудования.
    Если их не существует - обнуляем значения марок, считается что это ошибка '''
    for element in elements:
        if not any(element.Id == data.id for data in json_data):
            element.SetParamValue(TASK_SS_PARAM, "")
            element.SetParamValue(DATE_SS_PARAM, "")

def del_latest_file(file_folder_path):
    # Находим все JSON-файлы в директории
    json_files = glob.glob(os.path.join(file_folder_path, "СС(СППЗ)_*.json"))
    if not json_files:
        return

    # Находим файл с самым поздним временем модификации
    latest_file = max(json_files, key=os.path.getmtime)

    date = operator.get_utc_date()

    if date not in latest_file:
        return
    else:
        os.remove(latest_file)

def setup_params():
    revit_params = [MARK_PARAM,
                    FLOOR_PARAM,
                    SYSTEM_PARAM,
                    TASK_SS_PARAM,
                    DATE_SS_PARAM,
                    CREATE_TASK_SS_PARAM
                    ]

    project_parameters = ProjectParameters.Create(doc.Application)
    project_parameters.SetupRevitParams(doc, revit_params)


MARK_PARAM = SharedParamsConfig.Instance.VISMarkNumber
FLOOR_PARAM = SharedParamsConfig.Instance.BuildingWorksLevel
SYSTEM_PARAM = SharedParamsConfig.Instance.VISSystemName
TASK_SS_PARAM = SharedParamsConfig.Instance.VISTaskSSMark
DATE_SS_PARAM = SharedParamsConfig.Instance.VISTaskSSDate
CREATE_TASK_SS_PARAM = SharedParamsConfig.Instance.VISTaskSSAdd

operator = JsonOperator(doc, uiapp)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

    setup_params()
    raw_collection = get_elements()
    print_missing_task_params(raw_collection)
    selected_ids = set(element.Id.GetIdValue() for element in raw_collection)
    file_folder_path, is_path_local = operator.get_document_path()

    if is_path_local:
        operator.show_local_path(file_folder_path)

    # Получаем данные из последнего по дате редактирования файла
    old_data = operator.get_json_data(file_folder_path)

    # Если часть элементов занята, то задание может выйти не актуальным. Сразу останавливаемся,  если так
    check_edited_elements(raw_collection)

    # Матчим с старыми элементами для фильтрации тех что являются новыми
    new_elements = math_elements_to_old_data(raw_collection, old_data)

    # Оставляем только те новые элементы, у которых стоит галочка о формировании задания и параметр этажа
    # и системы имеет значение
    elements_to_objective = get_elements_to_objective(new_elements)

    # если новых элементов для задания нет то на этом можно остановиться
    if len(elements_to_objective) == 0:
        with revit.Transaction("BIM: Обновление задания"):
            for data in old_data:
                if data.id.GetIdValue() in selected_ids:
                    data.insert(doc, operator.get_utc_date())

            # Проверяем все элементы на предмет ложно-заполненности
            # (У элемента есть марка, но его не было в старых заданиях)
            clear_param_false_values(raw_collection, old_data)

            # Записываем в json-файл
            operator.send_json_data(old_data, file_folder_path)
            return

    # Делим новые элементы по спискам для применения отдельных алгоритмов
    open_valves, closed_valves, equipment_elements = split_collection(elements_to_objective)
    vacancy_path = os.path.join(file_folder_path, "Вакантные номера.json")
    if os.path.exists(vacancy_path):
        with codecs.open(vacancy_path, 'r', encoding='utf-8') as source:
            vacancies = json.load(source)
    else:
        vacancies = {"open": {}, "closed": {}}
    assigned = get_assigned_numbers(old_data)
    open_max = get_open_max_numbers(old_data)
    closed_max = get_closed_max_numbers(old_data)
    open_valves_data = use_open_algorithm(open_valves, old_data, vacancies, open_max, assigned)
    closed_valves_data = use_closed_algorithm(closed_valves, old_data, vacancies, closed_max, assigned)
    equipment_elements_data = use_open_algorithm(equipment_elements, old_data, vacancies, open_max, assigned)

    with revit.Transaction("BIM: Обновление задания"):
        if open_valves_data or closed_valves_data or equipment_elements_data:
            json_data = []
            json_data.extend(old_data)
            json_data.extend(open_valves_data)
            json_data.extend(closed_valves_data)
            json_data.extend(equipment_elements_data)

            # вставляем новые данные в проект. Если элемента нет в проекте - указываем для него дату удаления.
            # Если дата удаления совпадает с сегодняшней - игнорируем, это просто рабочие правки
            for low_voltage_system_data in json_data:
                if low_voltage_system_data.id.GetIdValue() in selected_ids:
                    low_voltage_system_data.insert(doc, operator.get_utc_date())

            # Записываем в json-файл
            operator.send_json_data(json_data, file_folder_path)
            with codecs.open(vacancy_path, 'w', encoding='utf-8') as target:
                json.dump(vacancies, target, ensure_ascii=False, indent=4)

            # Проверяем все элементы на предмет ложно-заполненности
            # (У элемента есть марка, но его не было в старых заданиях или в новых данных для json)
            clear_param_false_values(raw_collection, json_data)

script_execute()
