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
import JsonOperatorLib
import DebugPlacerLib
from System.Collections.Generic import *


from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import InternalOrigin
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
uiapp = __revit__.Application
uidoc = __revit__.ActiveUIDocument

EQUIPMENT_TYPE_NAME = "Оборудование"
VALVE_TYPE_NAME = "Клапан"
OUTER_VALVE_NAME = "ZAWTERM"

class CylinderZ:
    def __init__(self, z_min, z_max):
        self.diameter = 2000
        self.z_min = z_min
        self.z_max = z_max
        self.len = z_max - z_min

class AuditorEquipment:
    processed = False
    level_cylinder = None
    type_name = None

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
                 full_name = "",
                 type_name = None):
        base_point = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_ProjectBasePoint) \
            .WhereElementIsNotElementType() \
            .FirstElement()

        self.base_point_z = base_point.GetParamValue(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)

        self.connection_type = connection_type
        self.x = x
        self.y = y
        self.z = z
        self.base_x = x
        self.base_y = y
        self.base_z = z
        self.len = len
        self.code = code
        self.real_power = real_power
        self.nominal_power = nominal_power
        self.setting = setting
        self.maker = maker
        self.full_name = full_name
        self.type_name = type_name



        self.debug_placer = DebugPlacerLib.DebugPlacer(doc, diameter=2000)

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

            #self.print_debug_info(revit_equipment, 1850772, distance,
            #                      distance_to_bb_center, distance_to_location_center, radius, revit_coords)
            return distance <= radius
        return False

    def set_level_cylinder(self, level_cylinders):
        for level_cylinder in level_cylinders:
            if level_cylinder.z_min <= self.z <= level_cylinder.z_max:
                self.level_cylinder = level_cylinder

                # self.level_cylinder.z_min += convert_to_mms(self.base_point_z)
                # self.level_cylinder.z_max += convert_to_mms(self.base_point_z)
                # self.level_cylinder.len = self.level_cylinder.z_max - self.level_cylinder.z_min

                self.debug_placer.place_symbol(
                    self.x,
                    self.y,
                    self.z,
                    self.level_cylinder.z_max - self.level_cylinder.z_min
                )

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
                print('__Данные для расчета по xy__:')
                print('revit_equipment_x ' + str(revit_coords.x))
                print('revit_equipment_y ' + str(revit_coords.y))
                print('__Данные для расчета по z__:')
                print('level_cilinder_z_min ' + str(self.level_cylinder.z_min))
                print('level_cilinder_z_max ' + str(self.level_cylinder.z_max))
                print('revit_equipment_z ' + str(revit_coords.z))
                print('_________________________')


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
    setting_index = 17

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

def rotate_point_around_origin(x, y, z, angle_degrees):
    if angle_degrees == 0:
        return x, y, z
    # Угол в радианах
    angle_radians = math.radians(angle_degrees)
    # Матрица поворота вокруг оси Z (в плоскости XY)
    cos_theta = math.cos(angle_radians)
    sin_theta = math.sin(angle_radians)
    x_new = x * cos_theta - y * sin_theta
    y_new = x * sin_theta + y * cos_theta
    z_new = z
    return x_new, y_new, z_new

def extract_heating_device_description(file_path, angle):
    def parse_float(value):
        return float(value.replace(',', '.'))

    def parse_equipment_section(lines, title, start_offset, parse_func):
        result = []
        i = 0
        while i < len(lines):
            if len(result) == 10:
                return result
            if title in lines[i]:
                i += start_offset
                while i < len(lines) and lines[i].strip():
                    parsed_item = parse_func(lines[i])
                    if parsed_item is not None:
                        result.append(parsed_item)
                    i += 1
            i += 1
        return result

    def parse_heating_device(line):
        data = line.strip().split(';')
        rr = reading_rules_device
        x = parse_float(data[rr.x_index]) * 1000
        y = parse_float(data[rr.y_index]) * 1000
        z = parse_float(data[rr.z_index]) * 1000

        z = z + z_correction
        x, y, z = rotate_point_around_origin(x, y, z, angle)

        return AuditorEquipment(
            data[rr.connection_type_index],
            x, y, z,
            parse_float(data[rr.len_index]),
            data[rr.code_index],
            parse_float(data[rr.real_power_index]),
            parse_float(data[rr.nominal_power_index]),
            get_setting_float_value(data[rr.setting_index].replace(',', '.')),
            data[rr.maker_index],
            data[rr.full_name_index],
            EQUIPMENT_TYPE_NAME
        )

    def parse_valve(line):
        data = line.strip().split(';')
        rr = reading_rules_valve
        x = parse_float(data[rr.x_index]) * 1000
        y = parse_float(data[rr.y_index]) * 1000
        z = parse_float(data[rr.z_index]) * 1000
        z = z + z_correction

        if data[rr.connection_type_index] != OUTER_VALVE_NAME:
            return None
        return AuditorEquipment(
            data[rr.maker_index],
            x,
            y,
            z,
            setting=get_setting_float_value(data[rr.setting_index].replace(',', '.')),
            type_name=VALVE_TYPE_NAME
        )

    with codecs.open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    internal_origin = InternalOrigin.Get(doc)

    base_point = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_ProjectBasePoint) \
        .WhereElementIsNotElementType() \
        .FirstElement()

    base_point_z = base_point.GetParamValue(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
    print internal_origin
    internal_origin_z = internal_origin.SharedPosition.Z
    z_correction = (base_point_z - internal_origin_z) * 304.8

    reading_rules_device = ReadingRules()
    reading_rules_valve = ReadingRulesForValve()

    equipment = parse_equipment_section(lines, "Отопительные приборы CO на плане", 3, parse_heating_device)
    valves = parse_equipment_section(lines, "Арматура СО на плане", 3, parse_valve)

    equipment.extend(valves)

    if not equipment:
        forms.alert("Не найдено оборудование в импортируемом файле.", "Ошибка", exitscript=True)

    return equipment

def get_elements_by_category(category):
    """ Возвращает коллекцию элементов по категории """
    revit_equipment_elements = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()

    filtered_equipment = [
        eq for eq in revit_equipment_elements
        if FAMILY_NAME_CONST in eq.Symbol.Family.Name
    ]
    return filtered_equipment

def insert_data(element, auditor_data):
    if auditor_data.type_name == EQUIPMENT_TYPE_NAME:
        real_power_watts = UnitUtils.ConvertToInternalUnits(auditor_data.real_power, UnitTypeId.Watts)
        len_meters = UnitUtils.ConvertToInternalUnits(auditor_data.len, UnitTypeId.Millimeters)
        element.SetParamValue('ADSK_Размер_Длина', len_meters)
        element.SetParamValue('ADSK_Код изделия', auditor_data.code)
        element.SetParamValue('ADSK_Настройка', auditor_data.setting)
        element.SetParamValue('ADSK_Тепловая мощность', real_power_watts)
    else:
        element.SetParamValue('ADSK_Настройка', auditor_data.setting)

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
        if ayditor_equipment.type_name == EQUIPMENT_TYPE_NAME:
            unique_z_values.add(ayditor_equipment.z)

    unique_z_values = sorted(unique_z_values)

    cylinder_list = []
    for i in range(len(unique_z_values)):
        z_min = unique_z_values[i] - 250
        if i < len(unique_z_values) - 1:
            z_max = unique_z_values[i + 1] - 250
            # if z_max - z_min < 850: #
            #     z_max = z_min+850
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
            ayditor_equipment.base_x,
            ayditor_equipment.base_y,
            ayditor_equipment.base_z))

        print('ID приборов:')
        for x in equipment_in_area:
            print(x.Id)

def print_not_found_report(audytor_equipment_elements):
    not_found_audytor_reports = []
    for audytor_equipment in audytor_equipment_elements:
        if not audytor_equipment.processed:
            not_found_audytor_reports.append(audytor_equipment)

    if len(not_found_audytor_reports) > 0:
        print('Не найдено универсальное оборудование в областях:')
        for audytor_equipment in not_found_audytor_reports:
            print('Прибор х: {}, y: {}, z: {}'.format(
                audytor_equipment.base_x,
                audytor_equipment.base_y,
                audytor_equipment.base_z))

def process_start_up():
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True )

    filepath = select_file('Файл расчетов (*.txt)|*.txt')

    if filepath is None:
        sys.exit()

    operator = JsonOperatorLib.JsonAngleOperator(doc, uiapp)

    # Получаем данные из последнего по дате редактирования файла
    old_angle = operator.get_json_data()

    angle = forms.ask_for_string(
        default=str(old_angle),
        prompt='Введите угол наклона модели в градусах:',
        title="Аудитор импорт"
    )

    try:
        angle = float(angle.replace(',', '.'))
    except ValueError:
        forms.alert(
            "Необходимо ввести число.",
            "Ошибка",
            exitscript=True
        )

    if angle is None:
        sys.exit()

    operator.send_json_data(angle)

    return angle, filepath

def process_audytor_revit_matching(ayditror_equipment_elements, filtered_equipment):
    for ayditor_equipment in ayditror_equipment_elements:
        equipment_in_area = [
            eq for eq in filtered_equipment if ayditor_equipment.is_in_data_area(eq)
        ]
        ayditor_equipment.processed = len(equipment_in_area) >= 1
        if len(equipment_in_area) == 1:
            insert_data(equipment_in_area[0], ayditor_equipment)
        # if len(equipment_in_area) == 2:
        #     print equipment_in_area[0].type_name
        #     print equipment_in_area[1].type_name

        print_area_overflow_report(ayditor_equipment, equipment_in_area)

    print_not_found_report(ayditror_equipment_elements)


FAMILY_NAME_CONST = 'Обр_ОП_Универсальный'

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    angle, filepath = process_start_up()

    ayditror_equipment_elements = extract_heating_device_description(filepath, angle)

    # собираем высоты цилиндров в которых будем искать данные
    level_cylinders = get_level_cylinders(ayditror_equipment_elements)

    with revit.Transaction("BIM: Импорт расчетов"):
        for ayditor_equipment in ayditror_equipment_elements:
            ayditor_equipment.set_level_cylinder(level_cylinders)

        filtered_equipment = get_elements_by_category(BuiltInCategory.OST_MechanicalEquipment)

        process_audytor_revit_matching(ayditror_equipment_elements, filtered_equipment)


script_execute()
