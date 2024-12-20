#! /usr/bin/env python
# -*- coding: utf-8 -*-

from itertools import count

import clr

from unmodeling_class_library import UnmodelingFactory, MaterialCalculator, RowOfSpecification

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from pyrevit import forms
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict
from unmodeling_class_library import  *
from dosymep_libs.bim4everyone import *
import sys
from rpw.ui.forms import select_file
import json
import os
import codecs


doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument

class ValveData:
    floor_name = ''
    system_name = ''
    id = ''
    valve_base_name = ''

    def __init__(self, floor_name, system_name, id):
        self.system_name = system_name
        self.floor_name = floor_name
        self.id = id
        self.valve_base_name = system_name + "-" + floor_name

def get_json_data(file_path):
    name_mapping = {}
    if file_path and os.path.exists(file_path):
        with codecs.open(file_path, 'r', encoding='utf-8') as json_file:
            name_mapping = json.load(json_file)
            name_mapping = {item['id']: item['name'] for item in name_mapping}

    return name_mapping

def send_json_data(data, desktop_path=os.path.join(os.path.expanduser("~"), "Desktop")):
    file_path = os.path.join(desktop_path, "Задание СС.json")

    # Проверяем, существует ли файл
    if os.path.exists(file_path):
        # Читаем существующие данные
        with codecs.open(file_path, 'r', encoding='utf-8') as json_file:
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

    # Запись данных в JSON-файл
    with codecs.open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(combined_data, json_file, ensure_ascii=False, indent=4)

def split_valves_by_floors(valves):
    valves_by_floors = {}
    for valve in valves:
        floor_name = valve.GetParamValueOrDefault("ФОП_Этаж")
        system_name = valve.GetParamValueOrDefault("ФОП_ВИС_Имя системы")

        if floor_name is None or (system_name is None or system_name == "!Нет системы"):
            continue

        if floor_name not in valves_by_floors:
            valves_by_floors[floor_name] = []

        valves_by_floors[floor_name].append(ValveData(floor_name, system_name, valve.Id))

    return valves_by_floors

def create_open_numbers(valves):
    valves_by_floors = split_valves_by_floors(valves)

    data = []

    for floor_name, valve_data_instances in valves_by_floors.items():
        count = 0
        for valve_data in valve_data_instances:
            valve = doc.GetElement(valve_data.id)
            count += 1

            name = "НО-" + valve_data.valve_base_name + "-" + str(count)

            valve.SetParamValue("ADSK_Позиция", name)

            data.append({"name": name, "id": str(valve_data.id)})

    return data

def create_closed_numbers(valves):
    valves_by_floors = split_valves_by_floors(valves)
    data = []

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
                valve = doc.GetElement(valve_data.id)
                count += 1

                name = "НЗ-" + valve_base_name + "-" + str(count)

                valve.SetParamValue("ADSK_Позиция", name)

                data.append({"name": name, "id": str(valve_data.id)})

    return data

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    open_valves = []
    closed_valves = []

    file_path = select_file(title="Выберите существующий файл задания или нажмите отменить для нового")

    duct_accessory_collection = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_DuctAccessory) \
            .WhereElementIsNotElementType() \
            .ToElements()

    with revit.Transaction("BIM: Задание СС"):
        name_mapping = get_json_data(file_path)

        for duct_accessory in duct_accessory_collection:
            mark = duct_accessory.GetSharedParamValueOrDefault("ФОП_ВИС_Марка")

            if mark is not None and mark != "":
                if "НО" in mark or "НЗ" in mark:
                    if str(duct_accessory.Id) in name_mapping:
                        name = name_mapping[str(duct_accessory.Id)]
                        duct_accessory.SetParamValue("ADSK_Позиция", name)
                        continue # если айди клапана упомянут в задании то заново его не рассматриваем, сразу ставим имя оттуда

                    if "НО" in mark:
                        open_valves.append(duct_accessory)
                    if "НЗ" in mark:
                        closed_valves.append(duct_accessory)


        json_data = []
        json_data.extend(create_open_numbers(open_valves))
        json_data.extend(create_closed_numbers(closed_valves))

        if file_path is None:
            send_json_data(json_data)


script_execute()