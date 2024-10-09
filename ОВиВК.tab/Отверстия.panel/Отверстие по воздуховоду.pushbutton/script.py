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

# Функция для размещения семейства в заданных координатах
def place_family_at_coordinates(family_symbol, point, direction, element):
    # Создание экземпляра семейства
    transaction = Transaction(doc, "Размещение семейства")
    transaction.Start()

    # Получение размеров воздуховода и семейства
    duct_size = get_element_size(element)
    family_size = get_element_size(family_symbol)

    # Вычисление horizontal_offset
    horizontal_offset = duct_size.Y - family_size.Y

    # Вычисление vertical_offset
    vertical_offset = (duct_size.Z - family_size.Z) / 2

    # Сдвиг точки размещения на ось воздуховода по вертикали и горизонтали
    point = point + XYZ.BasisZ * vertical_offset + direction * horizontal_offset

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