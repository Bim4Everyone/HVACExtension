# -*- coding: utf-8 -*-
import sys
import clr
import codecs
import math

clr.AddReference('ProtoGeometry')
clr.AddReference("RevitNodes")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import Revit
import dosymep

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

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

from rpw.ui.forms import SelectFromList

from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI.Selection import ISelectionFilter

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import ElementFilter
from Autodesk.Revit.DB import LogicalOrFilter
from Autodesk.Revit.DB import ElementCategoryFilter
from Autodesk.Revit.DB import ElementMulticategoryFilter
from Autodesk.Revit.DB import Line
from Autodesk.Revit.DB import ElementTransformUtils
from Autodesk.Revit.DB import InternalOrigin

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView


class CustomSelectionFilter(ISelectionFilter):
    """Кастомный фильтр выбора, для отсечения не вертикальных элементов"""
    def __init__(self, filter):
        self.filter = filter

    def AllowElement(self, element):
        return self.filter.PassesFilter(element) and self.is_vertical(element)

    def AllowReference(self, reference, position):
        return True

    def is_vertical(self, element):
        start_xyz, end_xyz = get_connector_coordinates(element)

        # Вычисляем разности координат
        delta_x = round(end_xyz.X, 3) - round(start_xyz.X, 3)
        delta_y = round(end_xyz.Y, 3) - round(start_xyz.Y, 3)
        delta_z = round(end_xyz.Z, 3) - round(start_xyz.Z, 3)
        epsilon = 0.01 # Соответствует смещению в 3мм

        # Если линия вертикальна (delta_x < epsilon и delta_y < epsilon), возвращаем True
        if abs(delta_x) < epsilon and abs(delta_y) < epsilon:
            return True

        return False


def get_connector_coordinates(element):
    """
    Получение координат основных коннекторов элемента.
    """

    # Получаем коннекторы воздуховода
    connectors = element.ConnectorManager.Connectors

    # Получаем координаты начала и конца воздуховода через коннекторы
    start_point = None
    end_point = None

    for connector in connectors:
        # К линейному элементу может быть подключено произвольное количество врезок, каждая из которых попадет
        # в список коннекторов линейного элемента. Проверяем чтоб айди владельца был айди рабочего элемента
        if connector.Owner.Id == element.Id:
            if start_point is None:
                start_point = connector.Origin
            else:
                end_point = connector.Origin
                break

    if start_point is None or end_point is None:
        forms.alert("Не удалось получить координаты коннекторов.", "Ошибка", exitscript=True)

    # Получаем координаты начала и конца воздуховода
    start_xyz = XYZ(start_point.X, start_point.Y, start_point.Z)
    end_xyz = XYZ(end_point.X, end_point.Y, end_point.Z)

    return start_xyz, end_xyz


def get_selected():
    """
    Выделение воздуховодов и труб, размещенных вертикально.
    """
    categories = List[BuiltInCategory]()
    categories.Add(BuiltInCategory.OST_DuctCurves)
    categories.Add(BuiltInCategory.OST_PipeCurves)

    multi_category_filter = ElementMulticategoryFilter(categories)
    selection_filter = CustomSelectionFilter(multi_category_filter)

    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            selection_filter,
            "Выберите воздуховоды или трубы, нажмите Finish по окончании"
        )
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        sys.exit()

    elements = [doc.GetElement(r) for r in references]
    return elements


def get_reference_from_element(element):
    """
    Получаем референс из линейного элемента. Чтоб работало и на воздуховоды и на трубы - берем осевую линию.
    Если не получается ее найти - возвращаем None, чтоб потом вывести ошибку
    """
    options = Options()
    options.View = view
    options.ComputeReferences = True
    options.IncludeNonVisibleObjects = True

    geom_elem = element.get_Geometry(options)
    for geo_obj in geom_elem:
        if isinstance(geo_obj, Line):
            return geo_obj.Reference
    return None


def get_mark_endpoint(orientation_name, tag_point):
    """
    Получаем координаты марки самой марки и координаты перелома выноски. Вычисляем их на основе точки для которой
    считаем отметку.
    """
    end_epsilon = 3
    bend_epsilon = 2
    up_scale = 0.001

    right = view.RightDirection.Normalize()
    up = view.UpDirection.Normalize().Multiply(up_scale)
    left = -right

    if orientation_name == LEFT:
        direction = left + up
        end_point = tag_point + direction.Multiply(end_epsilon)
        bend_point = end_point + right.Multiply(bend_epsilon)
    else:
        direction = right + up
        end_point = tag_point + direction.Multiply(end_epsilon)
        bend_point = end_point - right.Multiply(bend_epsilon)

    return end_point, bend_point


def check_base_internal_diff():
    """
    Вычисляем корректировку по z на случай расхождения базовой точки и начала координат
    """
    internal_origin = InternalOrigin.Get(doc)

    base_point = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_ProjectBasePoint) \
        .WhereElementIsNotElementType() \
        .FirstElement()

    base_point_z = base_point.GetParamValue(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)

    internal_origin_z = internal_origin.SharedPosition.Z
    z_correction = (base_point_z - internal_origin_z)

    return z_correction


def start_up_checks():
    """Стартовые проверки"""
    if view.Category is None or not view.ViewType == ViewType.ThreeD:
        forms.alert(
            "Добавление отметок возможно только на 3D-Виде.",
            "Ошибка",
            exitscript=True)

    if not view.IsLocked:
        forms.alert(
            "3D‑вид должен быть заблокирован.",
            "Ошибка",
            exitscript=True)
        return


def get_levels_z():
    """Получение списка Z-координат уровней проекта"""
    levels = (FilteredElementCollector(doc).
              OfCategory(BuiltInCategory.OST_Levels).
              WhereElementIsNotElementType().ToElements())

    levels_z = []
    for level in levels:
        z = level.Elevation
        levels_z.append(z)
    return levels_z


LEFT = "Слева"
RIGHT = "Справа"

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    start_up_checks()

    elements = get_selected()

    orientation_name = SelectFromList('Выберите ориентацию марок', [LEFT, RIGHT])
    if orientation_name is None:
        sys.exit()

    levels_z = get_levels_z()

    with (revit.Transaction("BIM: Размещение отметок")):
        # z-коррекция нужна для случаев когда идет расхождение базовой точки и начала проекта. Оно сбоит при заборе
        # координат коннекторов, из них мы убираем коррекцию.
        # А так же сбой идет при передаче координат в ревит - тут мы наоборот добавляем коррекцию(tag_point)
        # Срабатывает и когда базовая точка ниже начала и выше
        z_correction = check_base_internal_diff()

        for element in elements:
            coord_list = get_connector_coordinates(element)
            connector_coord_1 = coord_list[0]
            connector_coord_2 = coord_list[1]
            min_z = min(connector_coord_1.Z, connector_coord_2.Z) - z_correction
            max_z = max(connector_coord_1.Z, connector_coord_2.Z) - z_correction

            for z in levels_z:
                ref = get_reference_from_element(element)
                if ref is None:
                    break
                if not (min_z < z < max_z):
                    continue

                tag_point = XYZ(connector_coord_1.X, connector_coord_1.Y, z + z_correction)
                end_point, bend_point = get_mark_endpoint(orientation_name, tag_point)

                doc.Create.NewSpotElevation(
                    view,
                    ref,
                    tag_point, # The point which the spot elevation evaluate.
                    bend_point, # The bend point for the spot elevation.
                    end_point, # The end point for the spot elevation.
                    tag_point, # The actual point on the reference which the spot elevation evaluate.
                    True
                )

    if ref is None:
        forms.alert(
            "На виде не найдены осевые линии. Отметки не были размещены.",
            "Ошибка",
            exitscript=True)


script_execute()