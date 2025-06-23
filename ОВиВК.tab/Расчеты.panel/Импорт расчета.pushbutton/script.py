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

class CylinderZ:
    def __init__(self, z_min, z_max):
        self.radius = 1000
        self.z_min = z_min
        self.z_max = z_max
        self.len = z_max - z_min

class UnitConverter:
    @staticmethod
    def to_millimeters(value):
        """Конвертирует внутренние единицы Revit в миллиметры."""
        return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters)
    @staticmethod
    def to_kilometers(value):
        """Конвертирует внутренние единицы Revit в километры."""
        return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Meters)/100

    @staticmethod
    def to_watts(value):
        """Конвертирует в Ватты."""
        return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Watts)

class GeometryHelper:
    @staticmethod
    def rotate_point(angle, x, y, z):
        '''
        Поворот координат из исходных данных на указанный угол вокруг начала координат из Аудитора
        '''
        if angle == 0:
            return x, y, z
        # Угол в радианах
        angle_radians = math.radians(angle)
        # Матрица поворота вокруг оси Z (в плоскости XY)
        cos_theta = math.cos(angle_radians)
        sin_theta = math.sin(angle_radians)
        x_new = x * cos_theta - y * sin_theta
        y_new = x * sin_theta + y * cos_theta
        return x_new, y_new, z

class TextParser:
    @staticmethod
    def parse_float(value):
        """Конвертирует строку в float, заменяя запятые на точки."""
        return float(value.replace(',', '.'))

    @staticmethod
    def parse_setting(value):
        """Обрабатывает значение настройки (например, 'N' → 0)."""
        if value in ('N', '', 'Kvs'):
            return 0
        return TextParser.parse_float(value)


class AuditorEquipment:

    '''
    Класс используется для хранения и обратки информации об элементах из Аудитора.

    ---
    processed : Bool
        Булева переменная предназначенная для недопущения дублирования элементов из Аудитор в циклах обработки

    level_cylinder : list
        Список, который содержит пары Z min и Z max для каждого элемента из Аудитор

    '''
    

    processed = False
    level_cylinder = None


    def __init__(self,
                 connection_type= "",
                 rotated_coords=None,
                 original_coords=None,
                 len = 0,
                 code = "",
                 real_power = "",
                 nominal_power = "",
                 setting = 0.0,
                 maker = "",
                 full_name = "",
                 type_name = None):
        '''
        Parametrs
        --------
        connection_type : str
            Тип обрабатываемого элемента в Аудиторе
        x_new, y_new, z_new : float
            Координаты после поворота
        x, y, z : float
            Координаты исходные

        '''
        base_point = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_ProjectBasePoint) \
            .WhereElementIsNotElementType() \
            .FirstElement()

        # base_point - Возвращает базовую точку проекта

        self.base_point_z = base_point.GetParamValue(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
        self.connection_type = connection_type

        # Используем XYZ для хранения координат
        self.original_coords = original_coords or XYZ.Zero
        self.rotated_coords = rotated_coords or XYZ.Zero

        self.len = len
        self.code = code
        self.real_power = real_power
        self.nominal_power = nominal_power
        self.setting = setting
        self.maker = maker
        self.full_name = full_name
        self.type_name = type_name

    def is_in_data_area(self, revit_equipment):
        '''
        Определяет, пересекаются ли области положений элемента в ревите и в аудиторе
        '''

        revit_location = revit_equipment.Location.Point
        revit_bb = revit_equipment.GetBoundingBox()
        revit_bb_center = BoundingBoxHelper.get_bb_center(revit_bb)

        revit_coords = XYZ(
            UnitConverter.to_millimeters(revit_location.X),
            UnitConverter.to_millimeters(revit_location.Y),
            UnitConverter.to_millimeters(revit_location.Z)
        )

        revit_bb_coords = XYZ(
            UnitConverter.to_millimeters(revit_bb_center.X),
            UnitConverter.to_millimeters(revit_bb_center.Y),
            UnitConverter.to_millimeters(revit_bb_center.Z)
        )

        radius = self.level_cylinder.radius

        epsilon = 1e-9

        if ((abs(self.level_cylinder.z_min - revit_coords.Z) <= epsilon or self.level_cylinder.z_min < revit_coords.Z)
                and (abs(revit_coords.Z - self.level_cylinder.z_max) <= epsilon
                     or revit_coords.Z < self.level_cylinder.z_max)):
            # Используем методы XYZ для вычисления расстояния
            distance_to_location_center = self.rotated_coords.DistanceTo(XYZ(revit_coords.X, revit_coords.Y, revit_coords.Z))
            distance_to_bb_center = self.rotated_coords.DistanceTo(XYZ(revit_bb_coords.X, revit_bb_coords.Y, revit_coords.Z))

            distance = min(distance_to_bb_center, distance_to_location_center)

            return distance <= radius
        return False

    def set_level_cylinder(self, level_cylinders):
        '''
        Вписывает в список свойств элемента из Аудитора минимальную и максимальную отметку проверочного цилиндра.
        При активации DEBUG_MODE создает в модели экземпляр Цилиндра по координатам элемента в Аудиторе.
        '''
        for level_cylinder in level_cylinders:
            if level_cylinder.z_min <= self.rotated_coords.Z <= level_cylinder.z_max:
                self.level_cylinder = level_cylinder

                if DEBUG_MODE:
                    comment = "{};{};{};{}".format(
                        self.type_name,
                        self.rotated_coords.X,
                        self.rotated_coords.Y,
                        self.rotated_coords.Z)
                    debug_placer.place_symbol(
                        self.rotated_coords.X,
                        self.rotated_coords.Y,
                        self.rotated_coords.Z,
                        self.level_cylinder.z_max - self.level_cylinder.z_min,
                        comment
                    )
                break

class EquipmentDataCache:
    def __init__(self):
        self._cache = {}

    def collect_data(self, element, auditor_data):
        """Собирает данные в кэш. Запись в Revit произойдёт позже."""
        if element.Id not in self._cache:
            self._cache[element.Id] = {
                "element": element,
                "data": auditor_data,
                "setting": auditor_data.setting or None
            }
        else:
            # Если это клапан, и есть новая настройка — обновим
            if auditor_data.setting:
                self._cache[element.Id]["setting"] = auditor_data.setting

    def write_all(self):
        """Пишет все данные в Revit — один раз для каждого элемента"""
        for item in self._cache.values():
            element = item["element"]
            data = item["data"]
            setting = item["setting"]

            if data.type_name == EQUIPMENT_TYPE_NAME:
                real_power_watts = UnitConverter.to_watts(data.real_power)
                len_meters = UnitConverter.to_kilometers(data.len)
                element.SetParamValue('ADSK_Размер_Длина', len_meters)
                element.SetParamValue('ADSK_Код изделия', data.code)
                element.SetParamValue('ADSK_Тепловая мощность', real_power_watts)

            # В любом случае, если есть настройка — записываем
            if setting:
                element.SetParamValue('ADSK_Настройка', setting)

class ReadingRulesForEquipment:
    '''
    Класс используется для интерпретиции данных по Приборам
    '''
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
    '''
    Класс используется для интерпретиции данных по Клапанам
    '''
    connection_type_index = 1
    maker_index = 2
    x_index = 3
    y_index = 4
    z_index = 5
    setting_index = 17

class BoundingBoxHelper:
    @staticmethod
    def get_bb_center(revit_bb):
        '''
        Получить центр Bounding Box

        Parametrs
        -------
        revit_coords : float
            Координаты центра вставки элемента в Ревите

        revit_bb_coords : float
            Координаты цента ВВ элемента в Ревите

        epsilon : float
            Погрешность
        '''
        minPoint = revit_bb.Min
        maxPoint = revit_bb.Max

        centroid = XYZ(
            (minPoint.X + maxPoint.X) / 2,
            (minPoint.Y + maxPoint.Y) / 2,
            (minPoint.Z + maxPoint.Z) / 2
        )
        return centroid

def get_setting_float_value(value):
    '''
    Корректировка настройки в исходных данных
    '''
    if value == 'N' or value == '' or value == 'Kvs':
        return 0
    else:
        return float(value)

def extract_heating_device_description(file_path, angle):
    '''
    Получение и обработка информации об элементов из исходных данных

    Parametrs
    ------
    equipment: list
        Список элементов и их свойств из Аудитора
    '''
    def parse_float(value):
        return float(value.replace(',', '.'))

    def parse_equipment_section(lines, title, start_offset, parse_func):
        '''
        Разбивка исходных данных на строки и отсечение лишней информации
        '''
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

        '''
        Чтение исходных данных для оборудования по указанными правилам, поворот координат на указанный угол
        и запись в параметры экземпляра класса
        '''
        data = line.strip().split(';')
        rr = reading_rules_device
        x = TextParser.parse_float(data[rr.x_index]) * 1000
        y = TextParser.parse_float(data[rr.y_index]) * 1000
        z = TextParser.parse_float(data[rr.z_index]) * 1000

        z = z + z_correction
        x_new, y_new, z_new = GeometryHelper.rotate_point(angle, x, y, z)

        return AuditorEquipment(
            connection_type=data[rr.connection_type_index],
            rotated_coords = XYZ(x_new,y_new,z_new),
            original_coords= XYZ(x,y,z),
            len=TextParser.parse_float(data[rr.len_index]),
            code=data[rr.code_index],
            real_power=TextParser.parse_float(data[rr.real_power_index]),
            nominal_power=TextParser.parse_float(data[rr.nominal_power_index]),
            setting=TextParser.parse_setting(data[rr.setting_index].replace(',', '.')),
            maker=data[rr.maker_index],
            full_name=data[rr.full_name_index],
            type_name=EQUIPMENT_TYPE_NAME
        )

    def parse_valve(line):
        '''
        Чтение исходных данных для клапанов по указанными правилам, поворот координат на указанный угол
        и запись в параметры экземпляра класса
        '''
        data = line.strip().split(';')
        rr = reading_rules_valve
        if data[rr.connection_type_index] != OUTER_VALVE_NAME:
            return None

        x = TextParser.parse_float(data[rr.x_index]) * 1000
        y = TextParser.parse_float(data[rr.y_index]) * 1000
        z = TextParser.parse_float(data[rr.z_index]) * 1000
        z = z + z_correction

        x_new, y_new, z_new = GeometryHelper.rotate_point(angle, x, y, z)

        return AuditorEquipment(
            maker=data[rr.maker_index],
            rotated_coords = XYZ(x_new,y_new,z_new),
            original_coords= XYZ(x,y,z),
            setting=TextParser.parse_setting(data[rr.setting_index].replace(',', '.')),
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

    internal_origin_z = internal_origin.SharedPosition.Z
    z_correction = (base_point_z - internal_origin_z) * 304.8

    reading_rules_device = ReadingRulesForEquipment()
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

def create_level_cylinders(ayditror_equipment_elements):
    '''
    Формирование цилиндров идет по низу аудитор-оборудования, которое выше отметок уровней в ревите. Соответственно,
    для попадания  отметок элементов ревита в эти цилиндры мы понижаем низ и верх цилиндров на небольшую величину
    '''

    max_z_offset = 2500 # Значение предельного смещения для Z-прибора. Обусловлено тем, что 2200 - предельная высота
    # установки приборов на лестничных клетках
    z_stock = 250 # Значение для понижения низа цилиндра, позволяющее проектировщику иметь погрешность по высоте

    unique_z_values = {
        eq.rotated_coords.Z for eq in ayditror_equipment_elements
        if eq.type_name == EQUIPMENT_TYPE_NAME
    }

    unique_z_values = sorted(unique_z_values)

    cylinder_list = []
    for i, z in enumerate(unique_z_values):
        z_min = z - z_stock

        if i + 1 < len(unique_z_values):
            next_z = unique_z_values[i + 1] - z_stock
            z_max = min(z_min + max_z_offset, next_z)
        else:
            z_max = z_min + max_z_offset

        cylinder = CylinderZ(z_min, z_max)
        cylinder_list.append(cylinder)
    return  cylinder_list

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
    def print_area_overflow_report(all_ayditor_equipment, all_equipment):
        from collections import defaultdict

        # Словарь: ключ — ID прибора, значение — список координат областей, в которые он попал
        equipment_to_areas = defaultdict(list)

        for ayditor_equipment in all_ayditor_equipment:
            equipment_in_area = [
                eq for eq in all_equipment if ayditor_equipment.is_in_data_area(eq)
            ]
            if len(equipment_in_area) > 1:
                for eq in equipment_in_area:
                    area_coords = (ayditor_equipment.original_coords.X,
                                   ayditor_equipment.original_coords.Y,
                                   ayditor_equipment.original_coords.Z)
                    equipment_to_areas[eq.Id].append(area_coords)

        if equipment_to_areas:
            print('Обнаружено переполнение данных областей:')
            for eq_id, areas in equipment_to_areas.iteritems():  # .iteritems() для Python 2
                print('\nID элемента: {}'.format(eq_id))
                print('Элемент попал в области:')
                for coords in areas:
                    print('  х: {}, y: {}, z: {}'.format(coords[0], coords[1], coords[2]))
            print '\n '

    def print_not_found_report(audytor_equipment_elements):
        not_found_audytor_reports = []
        for audytor_equipment in audytor_equipment_elements:
            if not audytor_equipment.processed:
                not_found_audytor_reports.append(audytor_equipment)

        if len(not_found_audytor_reports) > 0:
            print('Не найдено универсальное оборудование в областях:')
            for audytor_equipment in not_found_audytor_reports:
                print('Прибор х: {}, y: {}, z: {}'.format(
                    audytor_equipment.original_coords.X,
                    audytor_equipment.original_coords.Y,
                    audytor_equipment.original_coords.Z))

    data_cache = EquipmentDataCache()

    for ayditor_equipment in ayditror_equipment_elements:
        equipment_in_area = [
            eq for eq in filtered_equipment if ayditor_equipment.is_in_data_area(eq)
        ]
        ayditor_equipment.processed = len(equipment_in_area) >= 1
        if len(equipment_in_area) == 1:
            data_cache.collect_data(equipment_in_area[0], ayditor_equipment)

    # Новый группированный отчет
    print_area_overflow_report(ayditror_equipment_elements, filtered_equipment)

    # Финальный проход — запись параметров в Revit
    data_cache.write_all()

    print_not_found_report(ayditror_equipment_elements)

EQUIPMENT_TYPE_NAME = "Оборудование"
VALVE_TYPE_NAME = "Клапан"
OUTER_VALVE_NAME = "ZAWTERM"
FAMILY_NAME_CONST = 'Обр_ОП_Универсальный'
DEBUG_MODE = False

if DEBUG_MODE:
    debug_placer = DebugPlacerLib.DebugPlacer(doc, diameter=2000)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    angle, filepath = process_start_up()

    ayditror_equipment_elements = extract_heating_device_description(filepath, angle)

    # собираем высоты цилиндров в которых будем искать данные
    level_cylinders = create_level_cylinders(ayditror_equipment_elements)

    with revit.Transaction("BIM: Импорт приборов"):
        for ayditor_equipment in ayditror_equipment_elements:
            ayditor_equipment.set_level_cylinder(level_cylinders)

        equipment = get_elements_by_category(BuiltInCategory.OST_MechanicalEquipment)
        process_audytor_revit_matching(ayditror_equipment_elements, equipment)

script_execute()
