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
        self.diameter = 2000
        self.z_min = z_min
        self.z_max = z_max
        self.len = z_max - z_min

class AuditorEquipment:
    processed = False
    level_cylinder = None

    def __init__(self,
                 connection_type,
                 x,
                 y,
                 z,
                 len,
                 code,
                 real_power,
                 nominal_power,
                 setting,
                 maker,
                 full_name):
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

def extract_heating_device_description(file_path, reading_rules):
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
                    data[reading_rules.connection_type_index],
                    float(data[reading_rules.x_index].replace(',', '.')) * 1000,
                    float(data[reading_rules.y_index].replace(',', '.')) * 1000,
                    float(data[reading_rules.z_index].replace(',', '.')) * 1000,
                    float(data[reading_rules.len_index].replace(',', '.')),
                    data[reading_rules.code_index],
                    float(data[reading_rules.real_power_index]),
                    float(data[reading_rules.nominal_power_index]),
                    float(data[reading_rules.setting_index]),
                    data[reading_rules.maker_index],
                    data[reading_rules.full_name_index]
                ))
                i += 1

        i += 1

    if not equipment:
        forms.alert("Строка 'Отопительные приборы CO на плане' не найдена в файле.", "Ошибка", exitscript=True)

    return equipment

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

def get_level_cylinders(unique_z_values):
    cylinder_list = []
    for i in range(len(unique_z_values)):
        z_min = unique_z_values[i]
        if i < len(unique_z_values) - 1:
            z_max = unique_z_values[i + 1]
        else:
            z_max = z_min + 6000
        cylinder = CylinderZ(z_min, z_max)
        cylinder_list.append(cylinder)
    return  cylinder_list

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True )

    filepath = select_file('Файл расчетов (*.txt)|*.txt')

    if filepath is None:
        sys.exit()

    reading_rules = ReadingRules()

    ayditror_equipment_elements = extract_heating_device_description(filepath, reading_rules)

    unique_z_values = set()
    for ayditror_equipment in ayditror_equipment_elements:
        unique_z_values.add(ayditror_equipment.z)

    unique_z_values = sorted(unique_z_values)

    level_cylinders = get_level_cylinders(unique_z_values)

    for ayditror_equipment in ayditror_equipment_elements:
        for level_cylinder in level_cylinders:
            if level_cylinder.z_min <= ayditror_equipment.z <= level_cylinder.z_max:
                ayditror_equipment.level_cylinder = level_cylinder

    revit_equipment_elements = get_elements_by_category(BuiltInCategory.OST_MechanicalEquipment)

    not_found_ayditor_reports = [] # Отчеты о не найденных приборах для областей данных
    excess_in_data_area_reports = [] # Отчеты о превышениях количества приборов в областях даннных

    with revit.Transaction("BIM: Импорт расчетов"):
        for ayditror_equipment in ayditror_equipment_elements:


            # Если ранее было обработано можно не проверять
            if ayditror_equipment.processed:
                continue

            data_area = []
            for revit_equipment in revit_equipment_elements:
                family_name = revit_equipment.Symbol.Family.Name

                if 'Обр_ОП_Универсальный' not in family_name:
                    continue


                if ayditror_equipment.is_in_data_area(revit_equipment):
                    data_area.append(revit_equipment)

            if len(data_area) == 1:
                ayditror_equipment.processed = True
                insert_data(data_area[0], ayditror_equipment)
            if len(data_area) > 1: # Если в области данных дублирование элементов - данные из
                # аудитора могут перенестись идентично в несколько разных приборов
                print('В данные области попадает больше одного прибора:')
                print('Прибор х: {}, y: {}, z: {}'.format(
                    ayditror_equipment.x,
                    ayditror_equipment.y,
                    ayditror_equipment.z))
                print(ayditror_equipment.code)

                print('ID приборов:')
                for x in data_area:
                    print(x.Id)
                ayditror_equipment.processed = True
                excess_in_data_area_reports.append(ayditror_equipment)

    for ayditror_equipment in ayditror_equipment_elements:
        if not ayditror_equipment.processed:
            not_found_ayditor_reports.append(ayditror_equipment)


    if len(not_found_ayditor_reports) > 0:
        print('Оборудование аудитор, которое не было найдено в модели:')
        for ayditror_equipment in not_found_ayditor_reports:
            print('Прибор х: {}, y: {}, z: {}, артикул: {}, мощность: {}'.format(
                ayditror_equipment.x,
                ayditror_equipment.y,
                ayditror_equipment.z,
                ayditror_equipment.code,
                ayditror_equipment.real_power))


script_execute()
