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

def create_open_numbers(valves, file_path=None):
    valves_by_floors = {}
    for valve in valves:
        floor_name = valve.GetParamValueOrDefault("ФОП_Этаж")

        if floor_name is None:
            continue

        valve_id = valve.Id

        if floor_name not in valves_by_floors:
            valves_by_floors[floor_name] = []

        valves_by_floors[floor_name].append(valve_id)

    data = []
    name_mapping = {}

    if file_path and os.path.exists(file_path):
        with codecs.open(file_path, 'r', encoding='utf-8') as json_file:
            name_mapping = json.load(json_file)
            name_mapping = {item['id']: item['name'] for item in name_mapping}

    for floor_name, valve_ids in valves_by_floors.items():
        count = 0
        for valve_id in valve_ids:
            valve = doc.GetElement(valve_id)
            system_name = valve.GetParamValueOrDefault("ФОП_ВИС_Имя системы")
            count += 1

            if str(valve_id) in name_mapping:
                name = name_mapping[str(valve_id)]
            else:
                name = "НО-" + system_name + "-" + floor_name + "-" + str(count)

            valve.SetParamValue("ADSK_Позиция", name)

            data.append({"name": name, "id": str(valve_id)})

    if file_path is None:
        # Путь к рабочему столу
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        file_path = os.path.join(desktop_path, "Задание СС.json")

        # Запись данных в JSON-файл
        with codecs.open(file_path, 'w', encoding='utf-8') as json_file:
            json.dump(data, json_file, ensure_ascii=False, indent=4)


def create_closed_numbers(valves):
    pass

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    open_valves = []
    closed_valves = []

    filepath = select_file(title="Выберите существующий файл задания или нажмите отменить для нового")

    duct_accessory_collection = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_DuctAccessory) \
            .WhereElementIsNotElementType() \
            .ToElements()

    for duct_accessory in duct_accessory_collection:
        mark = duct_accessory.GetSharedParamValueOrDefault("ФОП_ВИС_Марка")
        if mark is not None and mark != "":
            if "НО" in mark:
                open_valves.append(duct_accessory)
            if "НЗ" in mark:
                closed_valves.append(duct_accessory)

    with revit.Transaction("BIM: Задание СС"):
        create_open_numbers(open_valves, filepath)
        create_closed_numbers(closed_valves)




script_execute()