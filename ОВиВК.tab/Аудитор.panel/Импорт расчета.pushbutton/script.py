# -*- coding: utf-8 -*-
import sys
import clr


clr.AddReference('ProtoGeometry')
clr.AddReference("RevitNodes")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import Revit
import dosymep
import codecs
import math

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

import System
from System.Collections.Generic import *


from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import Selection
from Autodesk.DesignScript.Geometry import *

import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from rpw.ui.forms import select_file


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *


doc = __revit__.ActiveUIDocument.Document  # type: Document
uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application
uidoc = __revit__.ActiveUIDocument

class CylinderZ:
    def __init__(self, z_min, z_max):
        self.diameter = 1500
        self.z_min = z_min
        self.z_max = z_max
        self.len = z_max - z_min

class AuditorEquipment:
    processed = False
    level_cylinder = None

    def __init__(self,
                 connection_type= "",
                 x = 0.0,
                 y = 0.0,
                 z = 0.0,
                 len = 0,
                 code = "",
                 real_power = "",
                 nominal_power = "",
                 setting = 0.0,
                 maker = "",
                 full_name = ""):
        self.connection_type = connection_type
        self.x = x
        self.y = y
        self.z = z
        self.len = len
        self.code = code
        self.real_power = real_power
        self.nominal_power = nominal_power
        self.setting = setting
        self.maker = maker
        self.full_name = full_name

    def is_in_data_area(self, revit_equipment):
        xyz = revit_equipment.Location.Point
        bb = revit_equipment.GetBoundingBox()
        bb_center = get_bb_center(bb)

        revit_coords = RevitXYZmms(
            convert_to_mms(xyz.X),
            convert_to_mms(xyz.Y),
            convert_to_mms(xyz.Z)
        )

        revit_bb_coords = RevitXYZmms(
            convert_to_mms(bb_center.X),
            convert_to_mms(bb_center.Y),
            convert_to_mms(bb_center.Z)
        )

        radius = self.level_cylinder.diameter / 2.0

        epsilon = 1e-9
        if ((abs(self.level_cylinder.z_min - revit_coords.z) <= epsilon or self.level_cylinder.z_min < revit_coords.z)
                and (abs(revit_coords.z - self.level_cylinder.z_max) <= epsilon
                     or revit_coords.z < self.level_cylinder.z_max)):
            distance_to_location_center = math.sqrt((self.x - revit_coords.x) ** 2 + (self.y - revit_coords.y) ** 2)
            distance_to_bb_center = math.sqrt((self.x - revit_bb_coords.x) ** 2 + (self.y - revit_bb_coords.y) ** 2)

            distance = min(distance_to_bb_center, distance_to_location_center)

            return distance <= radius
        return False

    def set_level_cylinder(self, level_cylinders):
        for level_cylinder in level_cylinders:
            if level_cylinder.z_min <= self.z <= level_cylinder.z_max:
                self.level_cylinder = level_cylinder

    def print_debug_info(self, revit_equipment,
                         integer_id,
                         distance,
                         distance_to_bb_center,
                         distance_to_location_center,
                         radius,
                         revit_coords
                         ):
        '''
        self.print_debug_info(revit_equipment, 2335627, distance,
        distance_to_bb_center, distance_to_location_center, radius, revit_coords) - шпаргалка для вызова дебага
        в расчете, вставлять перед return distance <= radius
        '''
        if revit_equipment.Id.IntegerValue == integer_id:
            if distance <= radius:
                print('__Характеристика элемента__:')
                print('element_id: ' + str(revit_equipment.Id))
                print('distance: ' + str(distance))
                print('distance_to_bb_center: ' + str(distance_to_bb_center))
                print('distance_to_location_center: ' + str(distance_to_location_center))
                print('__Данные для расчета по z__:')
                print('level_cilinder_z_min ' + str(self.level_cylinder.z_min))
                print('level_cilinder_z_max ' + str(self.level_cylinder.z_max))
                print('revit_equipment_z ' + str(revit_coords.z))
                print('_________________________')

# class AuditorValves:
#     processed = False
#     level_cylinder = None
#
#     def __init__(self,
#                  valve_type,
#                  x,
#                  y,
#                  z,
#                  setting):
#         self.valve_type = valve_type
#         self.x = x
#         self.y = y
#         self.z = z
#         self.len = len
#         self.setting = setting

class ReadingRules:
    connection_type_index = 2
    x_index = 3
    y_index = 4
    z_index = 5
    len_index = 12
    code_index = 16
    real_power_index = 20
    nominal_power_index = 22
    setting_index = 28
    maker_index = 30
    full_name_index = 31

class ReadingRulesForValve:
    connection_type_index = 1
    maker_index = 2
    x_index = 3
    y_index = 4
    z_index = 5
    setting_index = 29

class RevitXYZmms:
    def __init__(self, x, y, z):
        self.x = x
        self.y = y
        self.z = z

def convert_to_mms(value):
    """Конвертирует из внутренних значений ревита в миллиметры"""
    result = UnitUtils.ConvertFromInternalUnits(value,
                                               UnitTypeId.Millimeters)
    return result

def get_setting_float_value(value):
    if value == 'N' or value == '' or value == 'Kvs':
        return 0
    else:
        return float(value)


def extract_heating_device_description(file_path):
    reading_rules_device = ReadingRules()

    with codecs.open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    equipment = []

    i = 0

    while i < len(lines):
        if "Отопительные приборы CO на плане" in lines[i]:
            description_start_index = i + 3
            i = description_start_index

            while i < len(lines) and lines[i].strip() != "":
                data = lines[i].strip().split(';')

                equipment.append(AuditorEquipment(
                    data[reading_rules_device.connection_type_index],
                    float(data[reading_rules_device.x_index].replace(',', '.')) * 1000,
                    float(data[reading_rules_device.y_index].replace(',', '.')) * 1000,
                    float(data[reading_rules_device.z_index].replace(',', '.')) * 1000,
                    float(data[reading_rules_device.len_index].replace(',', '.')),
                    data[reading_rules_device.code_index],
                    float(data[reading_rules_device.real_power_index]),
                    float(data[reading_rules_device.nominal_power_index]),
                    get_setting_float_value(data[reading_rules_device.setting_index].replace(',', '.')),
                    data[reading_rules_device.maker_index],
                    data[reading_rules_device.full_name_index]
                ))
                i += 1

        i += 1

    if not equipment:
        forms.alert("Строка 'Отопительные приборы CO на плане' не найдена в файле.", "Ошибка", exitscript=True)

    reading_rules_valve = ReadingRulesForValve()
    valves = []
    j = 0

    while j < len(lines):
        if "Арматура СО на плане" in lines[j]:
            description_start_index = j + 3
            j = description_start_index

            while j < len(lines) and lines[j].strip() != "":
                data = lines[j].strip().split(';')
                if data[reading_rules_valve.connection_type_index] == "ZAWTERM":
                    valves.append(AuditorEquipment(maker="2", x=3, y=4, z=5, setting=29))
                    # valves.append(AuditorValves(
                    #     data[reading_rules.valve_type],
                    #     data[reading_rules.valve_mark],
                    #     float(data[reading_rules.x_index].replace(',', '.')) * 1000,
                    #     float(data[reading_rules.y_index].replace(',', '.')) * 1000,
                    #     float(data[reading_rules.z_index].replace(',', '.')) * 1000,
                    #     get_setting_float_value(data[reading_rules.valve_setting_index].replace(',', '.'))
                    # ))
                    j += 1
        j += 1

    if not valves:
        forms.alert("Строка 'Арматура СО на плане' не найдена в файле.", "Ошибка", exitscript=True)

    equipment = equipment.extend(valves)

    return equipment

def extract_valve_description(file_path):
    reading_rules = ReadingRulesForValve()

    with codecs.open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    valves = []
    i = 0

    while i < len(lines):
        if "Арматура СО на плане" in lines[i]:
            description_start_index = i + 3
            i = description_start_index

            while i < len(lines) and lines[i].strip() != "":
                data = lines[i].strip().split(';')
                if data[reading_rules.valve_type] == "ZAWTERM":

                    valves.append(AuditorEquipment(maker=2, x=3, y=4, z=5, setting=29))

                    # valves.append(AuditorValves(
                    #     data[reading_rules.valve_type],
                    #     data[reading_rules.valve_mark],
                    #     float(data[reading_rules.x_index].replace(',', '.')) * 1000,
                    #     float(data[reading_rules.y_index].replace(',', '.')) * 1000,
                    #     float(data[reading_rules.z_index].replace(',', '.')) * 1000,
                    #     get_setting_float_value(data[reading_rules.valve_setting_index].replace(',', '.'))
                    # ))
                    i += 1
        i += 1

    if not valves:
        forms.alert("Строка 'Арматура СО на плане' не найдена в файле.", "Ошибка", exitscript=True)

    return valves

def get_elements_by_category(category):
    """ Возвращает коллекцию элементов по категории """
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

def insert_data(element, auditor_data):
    real_power_watts = UnitUtils.ConvertToInternalUnits(auditor_data.real_power, UnitTypeId.Watts)
    len_meters = UnitUtils.ConvertToInternalUnits(auditor_data.len, UnitTypeId.Millimeters)

    element.SetParamValue('ADSK_Размер_Длина', len_meters)
    element.SetParamValue('ADSK_Код изделия', auditor_data.code)
    element.SetParamValue('ADSK_Настройка', auditor_data.setting)
    element.SetParamValue('ADSK_Тепловая мощность', real_power_watts)

def get_bb_center(bb):
    minPoint = bb.Min
    maxPoint = bb.Max

    centroid = XYZ(
        (minPoint.X + maxPoint.X) / 2,
        (minPoint.Y + maxPoint.Y) / 2,
        (minPoint.Z + maxPoint.Z) / 2
    )
    return centroid

def get_level_cylinders(ayditror_equipment_elements):
    '''
    Формирование цилиндров идет по низу аудитор-оборудования, которое выше отметок уровней в ревите. Соответственно,
    для попадания отметок элементов ревита в эти цилиндры мы понижаем низ и верх цилиндров на небольшую величину
    '''
    unique_z_values = set()
    for ayditor_equipment in ayditror_equipment_elements:
        unique_z_values.add(ayditor_equipment.z)

    unique_z_values = sorted(unique_z_values)

    cylinder_list = []
    for i in range(len(unique_z_values)):
        z_min = unique_z_values[i] - 250
        if i < len(unique_z_values) - 1:
            z_max = unique_z_values[i + 1] - 250
        else:
            z_max = z_min + 2500

        cylinder = CylinderZ(z_min, z_max)
        cylinder_list.append(cylinder)
    return  cylinder_list

def print_area_overflow_report(ayditor_equipment, equipment_in_area):
    if len(equipment_in_area) > 1:  # Если в области данных дублирование элементов - данные из
        # аудитора могут перенестись идентично в несколько разных приборов
        print('В данные области попадает больше одного прибора:')
        print('Прибор х: {}, y: {}, z: {}'.format(
            ayditor_equipment.x,
            ayditor_equipment.y,
            ayditor_equipment.z))

        print('ID приборов:')
        for x in equipment_in_area:
            print(x.Id)


def print_not_found_report(audytor_equipment_elements):
    not_found_audytor_reports = []
    for audytor_equipment in ayditror_equipment_elements:
        if not audytor_equipment.processed:
            not_found_audytor_reports.append(audytor_equipment)

    if len(not_found_audytor_reports) > 0:
        print('Не найдено универсальное оборудование в областях:')
        for audytor_equipment in not_found_audytor_reports:
            print('Прибор х: {}, y: {}, z: {}'.format(
                audytor_equipment.x,
                audytor_equipment.y,
                audytor_equipment.z))

FAMILY_NAME_CONST = 'Обр_ОП_Универсальный'

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True )

    filepath = select_file('Файл расчетов (*.txt)|*.txt')

    if filepath is None:
        sys.exit()

    ayditror_equipment_elements = extract_heating_device_description(filepath)
    #ayditror_valve_elements = extract_valve_description(filepath)

    # собираем высоты цилиндров в которых будем искать данные
    level_cylinders = get_level_cylinders(ayditror_equipment_elements)

    for ayditor_equipment in ayditror_equipment_elements:
        ayditor_equipment.set_level_cylinder(level_cylinders)

    revit_equipment_elements = get_elements_by_category(BuiltInCategory.OST_MechanicalEquipment)

    with revit.Transaction("BIM: Импорт расчетов"):
        for ayditor_equipment in ayditror_equipment_elements:
            equipment_in_area = [] # Список данных которые попадают в текущую область
            for revit_equipment in revit_equipment_elements:
                family_name = revit_equipment.Symbol.Family.Name
                if FAMILY_NAME_CONST not in family_name:
                    continue

                if ayditor_equipment.is_in_data_area(revit_equipment):
                    equipment_in_area.append(revit_equipment)

            ayditor_equipment.processed = len(equipment_in_area) >= 1
            if len(equipment_in_area) == 1:
                insert_data(equipment_in_area[0], ayditor_equipment)

            print_area_overflow_report(ayditor_equipment, equipment_in_area)

        print_not_found_report(ayditror_equipment_elements)


script_execute()
