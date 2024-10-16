#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = "Отверстие по воздуховоду"
__doc__ = "Пересчитывает КМС соединительных деталей воздуховодов"

import clr
import datetime
from pyrevit.labs import target

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.DB import BuiltInCategory, ElementFilter, LogicalOrFilter, ElementCategoryFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from System.Collections.Generic import List
from System import Guid
from pyrevit import forms
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document # type: Document

uidoc = __revit__.ActiveUIDocument
uiapp = __revit__.Application

view = doc.ActiveView

class CurveSizes:
    curve_diameter = 0
    curve_width = 0
    curve_height = 0

# Создаем фильтры для нужных категорий
duct_filter = ElementCategoryFilter(BuiltInCategory.OST_DuctCurves)
pipe_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeCurves)

# Объединяем фильтры с помощью логического "или"
combined_filter = LogicalOrFilter(duct_filter, pipe_filter)

# Создаем пользовательский фильтр для выбора объектов
class CustomSelectionFilter(ISelectionFilter):
    def __init__(self, filter):
        self.filter = filter

    def AllowElement(self, element):
        return self.filter.PassesFilter(element)

    def AllowReference(self, reference, position):
        return True

# Функция для получения координат точки на воздуховоде
def get_point_coordinates():
    # Применяем пользовательский фильтр к выбору объектов
    selection_filter = CustomSelectionFilter(combined_filter)

    try:
        reference = uidoc.Selection.PickObject(ObjectType.Element, selection_filter)
        # Продолжайте выполнение кода, если объект был выбран
    except OperationCanceledException:
        # Обработка отмены операции выбора
        sys.exit()

    if reference:
        element = doc.GetElement(reference)
        # Получение координат точки
        point = reference.GlobalPoint
        coordinates = point.ToString()
        return element, point
    else:
        return None, None

# Функция для получения центра и направления воздуховода
def get_curve_direction(duct):
    options = Options()
    geometry = duct.get_Geometry(options)
    bbox = geometry.GetBoundingBox()
    # Получение направления воздуховода
    curve = next((geom for geom in geometry if isinstance(geom, Solid) and geom.Faces.Size > 0), None)
    if curve:
        face = curve.Faces.get_Item(0)
        normal = face.ComputeNormal(UV(0.5, 0.5))
        direction = normal.CrossProduct(XYZ.BasisZ)
    else:
        direction = XYZ.BasisX  # По умолчанию используем направление по оси X
    return direction

# Функция для поиска семейства в проекте
def find_family_symbol(family_name):
    family_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    family_symbol = next((fs for fs in family_symbols if fs.Family.Name == family_name), None)
    if not family_symbol:
        print("Семейство не найдено.")
        return None
    return family_symbol

# Функция для получения размеров элемента
def get_element_size(element):
    options = Options()
    geometry = element.get_Geometry(options)
    bbox = geometry.GetBoundingBox()
    size = bbox.Max - bbox.Min
    return size

def get_offset(element, point, direction, use_horizontal_projection):
    # Получаем коннекторы воздуховода
    connectors = element.ConnectorManager.Connectors

    # Проверяем, что у нас есть как минимум два коннектора
    if connectors.Size < 2:
        raise ValueError("Воздуховод должен иметь как минимум два коннектора.")

    # Получаем координаты начала и конца воздуховода через коннекторы
    start_point = None
    end_point = None
    for connector in connectors:
        if start_point is None:
            start_point = connector.Origin
        else:
            end_point = connector.Origin
            break

    if start_point is None or end_point is None:
        raise ValueError("Не удалось получить координаты коннекторов.")

    # Получаем координаты точки
    point_x = point.X
    point_y = point.Y
    point_z = point.Z

    # Получаем координаты начала и конца воздуховода
    start_x = start_point.X
    start_y = start_point.Y
    start_z = start_point.Z
    end_x = end_point.X
    end_y = end_point.Y
    end_z = end_point.Z

    if use_horizontal_projection:
        # Вычисляем длину нормали от точки до прямой в плоскости X-Y
        numerator = abs((end_x - start_x) * (start_y - point_y) - (start_x - point_x) * (end_y - start_y))
        denominator = math.sqrt((end_x - start_x) ** 2 + (end_y - start_y) ** 2)
    else:
        # Вычисляем длину нормали от точки до прямой в плоскости X-Z
        numerator = abs((end_x - start_x) * (start_z - point_z) - (start_x - point_x) * (end_z - start_z))
        denominator = math.sqrt((end_x - start_x) ** 2 + (end_z - start_z) ** 2)

    if denominator == 0:
        return 0  # Если длина прямой равна нулю, возвращаем 0

    distance = numerator / denominator

    # Вычисление vertical_offset
    vertical_offset = 0
    horizontal_offset = 0

    # Проверочная точка для выявления, проходит ли ось через нее
    if use_horizontal_projection:
        target = point + XYZ.BasisZ * vertical_offset + direction * distance
    else:
        target = point + XYZ.BasisZ * distance + direction * horizontal_offset

    # Проверка, проходит ли линия через точку target
    if use_horizontal_projection:
        if is_point_on_line(start_x, start_y, end_x, end_y, target.X, target.Y):
            return distance
        else:
            return distance * -1
    else:
        if is_point_on_line(start_x, start_z, end_x, end_z, target.X, target.Z):
            return distance
        else:
            return distance * -1

def is_point_on_line(start_x, start_y, end_x, end_y, target_x, target_y, epsilon=0.1):
    # Проверка, лежит ли точка на прямой с учетом погрешности
    if abs((end_y - start_y) * (target_x - start_x) - (end_x - start_x) * (target_y - start_y)) > epsilon:
        return False

    # Проверка, лежит ли точка в пределах отрезка
    if min(start_x, end_x) <= target_x <= max(start_x, end_x) and min(start_y, end_y) <= target_y <= max(start_y, end_y):
        return True

    return False

def get_parameter_if_exists(element, param_name):
    if element.IsExistsParam(param_name):
        return element.GetParam(param_name)
    else:
        return None

def setup_size(instance, curve):
    indent = (50 * 2) / 304.8

    instance_diameter_param = get_parameter_if_exists(instance, "ADSK_Размер_Диаметр")
    instance_height_param = get_parameter_if_exists(instance, "ADSK_Размер_Высота")
    instance_width_param = get_parameter_if_exists(instance, "ADSK_Размер_Ширина")

    curve_diameter = 0
    curve_width = 0
    curve_height = 0

    if curve.Category.IsId(BuiltInCategory.OST_PipeCurves):
        curve_diameter = curve.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
    elif curve.Category.IsId(BuiltInCategory.OST_DuctCurves):
        if curve.DuctType.Shape == ConnectorProfileType.Round:
            curve_diameter = curve.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        elif curve.DuctType.Shape == ConnectorProfileType.Rectangular:
            curve_width = curve.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            curve_height = curve.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)

    if curve_diameter != 0:
        if instance_diameter_param:
            instance_diameter_param.Set(curve_diameter + indent)
            correction = 0
        else:
            instance_width_param.Set(curve_diameter + indent)
            instance_height_param.Set(curve_diameter + indent)
            correction = (curve_diameter + indent) / 2
    else:
        if instance_diameter_param:
            instance_diameter_param.Set(math.sqrt(curve_width ** 2 + curve_height ** 2)  + indent)
            correction = 0
        else:
            instance_width_param.Set(curve_width + indent)
            instance_height_param.Set(curve_height + indent)
            correction = (curve_height + indent) / 2

    return correction

def get_curve_system(curve):
    system_name = curve.GetParamValueOrDefault("ФОП_ВИС_Имя системы")
    if system_name is None:
        system_name = curve.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
    return system_name

def setup_opening_instance(instance, curve):
    # Заполняем автора задания
    user_name = __revit__.Application.Username
    instance.SetParamValue("ФОП_Автор задания", user_name)

    # Заполняем айди линейного элемента
    instance.SetParamValue("ФОП_Описание", curve.Id.ToString())

    # Заполняем имя системы элемента
    instance.SetParamValue("ФОП_ВИС_Имя системы", get_curve_system(curve))

    # Заполняем время
    current_time = datetime.datetime.now()
    formatted_time = current_time.strftime("%Y-%m-%d %H:%M")
    instance.SetParamValue("ФОП_Дата", formatted_time)

    # Устанавливаем размер под воздуховод и получаем смещение относительно него
    correction = setup_size(instance, curve)
    instance_offset_param = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
    instance_current_offset = instance.GetParamValueOrDefault(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
    instance_offset_param.Set(instance_current_offset - correction)

# Функция для размещения семейства в заданных координатах
def place_family_at_coordinates(family_symbol, point, direction, element):
    # Создание экземпляра семейства
    transaction = Transaction(doc, "Размещение семейства")
    transaction.Start()

    # Вычисление horizontal_offset
    horizontal_offset = get_offset(element, point, direction, use_horizontal_projection=True)

    # Вычисление vertical_offset
    vertical_offset = get_offset(element, point, direction, use_horizontal_projection=False)

    # Сдвиг точки размещения на ось воздуховода по вертикали и горизонтали
    point = point + XYZ.BasisZ * vertical_offset + direction * horizontal_offset

    # Создание экземпляра
    family_symbol.Activate()
    instance = doc.Create.NewFamilyInstance(point, family_symbol, Structure.StructuralType.NonStructural)

    setup_opening_instance(instance, element)

    # Создание оси вращения, проходящей через точку размещения и направленной вдоль оси Z
    axis = Line.CreateBound(point, point + XYZ.BasisZ)

    # Вычисление угла поворота
    angle = math.atan2(direction.Y, direction.X)

    # Вращение экземпляра семейства вокруг оси Z
    instance.Location.Rotate(axis, angle)

    transaction.Commit()

#@notification()
#@log_plugin(EXEC_PARAMS.command_name)
def script_execute():
    element, point = get_point_coordinates()
    duct_direction = get_curve_direction(element)
    # family_name = "ОбщМд_Отв_Отверстие_Прямоугольное_В стене"
    family_name =  "ОбщМд_Отв_Отверстие_Круглое_В стене"
    family_symbol = find_family_symbol(family_name)

    if family_symbol:
        place_family_at_coordinates(family_symbol, point, duct_direction, element)

script_execute()