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
    ''' Возвращает значение ФОП_Этаж, если в нем нет дефиса. Если есть - падаем с отчетом. '''
    value = element.GetParamValueOrDefault(FLOOR_PARAM)

    if "-" in value:
        report_text = ("У части отмеченного для задания оборудования в значении этажа указан символ дефиса (-). "
                       "Исключите такие обозначения. Для подземных этажей используйте обозначения П01, П02 и т.д")

        forms.alert(report_text, "Ошибка", exitscript=True)

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

def use_open_algorithm(open_valves, old_data):
    ''' Создаем имена для списка элементов используя алгоритм для открытых клапанов от СППЗ '''
    open_valves_by_floors = split_equipment_by_floors(open_valves, opened=True)
    max_numbers = get_open_max_numbers(old_data)

    result = []
    for floor_name, open_valves_instances in open_valves_by_floors.items():
        count = 0
        for open_valve in open_valves_instances:
            if open_valve.element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
                base_name  = open_valve.equipment_base_name
            else:
                base_name = "НО-" + open_valve.equipment_base_name

            # Проверяем, существует ли ключ в словаре max_numbers
            if floor_name in max_numbers and count < max_numbers[floor_name]:
                count = max_numbers[floor_name] + 1
            else:
                count += 1

            open_valve.json_name = base_name + "-" + str(count)

            result.append(open_valve)

    return result

def use_closed_algorithm(valves, old_data):
    ''' Создаем имена для списка элементов используя алгоритм для закрытых клапанов от СППЗ '''
    valves_by_floors = split_equipment_by_floors(valves, closed=True)
    max_numbers = get_closed_max_numbers(old_data)

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
            count = 0
            for valve_data in valve_data_instances_group:
                if valve_data.equipment_base_name in max_numbers and count < max_numbers[floor_name]:
                    count = max_numbers[floor_name] + 1
                else:
                    count += 1
                valve_data.json_name = "НЗ-" + valve_data.equipment_base_name + "-" + str(count)

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

def get_elements():
    ''' Забираем список элементов арматуры и оборудования '''
    categories = [
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PipeAccessory
    ]

    category_ids = List[ElementId]([ElementId(int(category)) for category in categories])

    multicategory_filter = ElementMulticategoryFilter(category_ids)

    elements = FilteredElementCollector(doc) \
        .WherePasses(multicategory_filter) \
        .WhereElementIsNotElementType() \
        .ToElements()

    return elements

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

        match = re.search(r'^(.*-\d+-)(\d+)$', data.json_name)

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
FLOOR_PARAM = SharedParamsConfig.Instance.Level
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
    file_folder_path, is_path_local = operator.get_document_path()

    if is_path_local:
        operator.show_local_path(file_folder_path)

    # Получаем данные из последнего по дате редактирования файла
    old_data = operator.get_json_data(file_folder_path)

    # Получаем список элементов, без фильтрации
    raw_collection = get_elements()

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
                data.insert(doc, operator.get_utc_date())

            # Проверяем все элементы на предмет ложно-заполненности
            # (У элемента есть марка, но его не было в старых заданиях)
            clear_param_false_values(raw_collection, old_data)

            # Записываем в json-файл
            operator.send_json_data(old_data, file_folder_path)
            return

    # Делим новые элементы по спискам для применения отдельных алгоритмов
    open_valves, closed_valves, equipment_elements = split_collection(elements_to_objective)
    open_valves_data = use_open_algorithm(open_valves, old_data)
    closed_valves_data = use_closed_algorithm(closed_valves, old_data)
    equipment_elements_data = use_open_algorithm(equipment_elements, old_data)

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
                low_voltage_system_data.insert(doc, operator.get_utc_date())

            # Записываем в json-файл
            operator.send_json_data(json_data, file_folder_path)

            # Проверяем все элементы на предмет ложно-заполненности
            # (У элемента есть марка, но его не было в старых заданиях или в новых данных для json)
            clear_param_false_values(raw_collection, json_data)

script_execute()