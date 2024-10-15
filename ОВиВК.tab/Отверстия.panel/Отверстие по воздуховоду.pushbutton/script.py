#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = "Отверстие по воздуховоду"
__doc__ = "Пересчитывает КМС соединительных деталей воздуховодов"

import clr

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
from System.Collections.Generic import List
from System import Guid

doc = __revit__.ActiveUIDocument.Document # type: Document

uidoc = __revit__.ActiveUIDocument
uiapp = __revit__.Application

view = doc.ActiveView

# Функция для получения координат точки на воздуховоде
def get_point_coordinates(uiapp):
    # Запрос на выбор элемента
    reference = uidoc.Selection.PickObject(ObjectType.Element)
    element = doc.GetElement(reference)
    # Получение координат точки
    point = reference.GlobalPoint
    coordinates = point.ToString()
    return element, point

# Функция для получения центра и направления воздуховода
def get_duct_center_and_direction(duct):
    options = Options()
    geometry = duct.get_Geometry(options)
    bbox = geometry.GetBoundingBox()
    center = (bbox.Min + bbox.Max) / 2
    # Получение направления воздуховода
    curve = next((geom for geom in geometry if isinstance(geom, Solid) and geom.Faces.Size > 0), None)
    if curve:
        face = curve.Faces.get_Item(0)
        normal = face.ComputeNormal(UV(0.5, 0.5))
        direction = normal.CrossProduct(XYZ.BasisZ)
    else:
        direction = XYZ.BasisX  # По умолчанию используем направление по оси X
    return center, direction

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
def get_horizontal_offset(element, point, direction):
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

    # Получаем координаты начала и конца воздуховода
    start_x = start_point.X
    start_y = start_point.Y
    end_x = end_point.X
    end_y = end_point.Y

    # Вычисляем длину нормали от точки до прямой
    # Формула для расстояния от точки до прямой:
    # d = |(x2 - x1)(y1 - y0) - (x1 - x0)(y2 - y1)| / sqrt((x2 - x1)^2 + (y2 - y1)^2)
    numerator = abs((end_x - start_x) * (start_y - point_y) - (start_x - point_x) * (end_y - start_y))
    denominator = math.sqrt((end_x - start_x) ** 2 + (end_y - start_y) ** 2)

    if denominator == 0:
        return 0  # Если длина прямой равна нулю, возвращаем 0

    distance = numerator / denominator

    # Вычисление vertical_offset
    vertical_offset = 0

    # Сдвиг точки размещения на ось воздуховода по вертикали и горизонтали
    target = point + XYZ.BasisZ * vertical_offset + direction * -1 * distance

    # Проверка, проходит ли линия через точку target
    if is_point_on_line(start_x, start_y, end_x, end_y, target.X, target.Y):
        return distance
    else:
        return distance * -1



def is_point_on_line(start_x, start_y, end_x, end_y, target_x, target_y, epsilon=1e-9):
    # Проверка, лежит ли точка на прямой с учетом погрешности
    if abs((end_x - start_x) * (target_y - start_y) - (end_y - start_y) * (target_x - start_x)) > epsilon:
        return False

    # Проверка, лежит ли точка в пределах отрезка
    if min(start_x, end_x) <= target_x <= max(start_x, end_x) and min(start_y, end_y) <= target_y <= max(start_y, end_y):
        return True

    return False

# Функция для размещения семейства в заданных координатах
def place_family_at_coordinates(family_symbol, point, direction, element):
    # Создание экземпляра семейства
    transaction = Transaction(doc, "Размещение семейства")
    transaction.Start()

    # Получение размеров воздуховода и семейства


    # Вычисление horizontal_offset
    horizontal_offset = get_horizontal_offset(element, point, direction)

    # Вычисление vertical_offset
    vertical_offset = 0

    # Сдвиг точки размещения на ось воздуховода по вертикали и горизонтали
    point = point + XYZ.BasisZ * vertical_offset + direction * -1 * horizontal_offset

    family_symbol.Activate()
    instance = doc.Create.NewFamilyInstance(point, family_symbol, Structure.StructuralType.NonStructural)

    # Создание оси вращения, проходящей через точку размещения и направленной вдоль оси Z
    axis = Line.CreateBound(point, point + XYZ.BasisZ)

    # Вычисление угла поворота
    angle = math.atan2(direction.Y, direction.X)

    # Вращение экземпляра семейства вокруг оси Z
    instance.Location.Rotate(axis, angle)

    transaction.Commit()

# Основная функция для запуска
def main():
    element, point = get_point_coordinates(uiapp)
    duct_center, duct_direction = get_duct_center_and_direction(element)
    family_name = "ОбщМд_Отв_Отверстие_Прямоугольное_В стене"
    family_symbol = find_family_symbol(family_name)
    if family_symbol:
        place_family_at_coordinates(family_symbol, point, duct_direction, element)

main()