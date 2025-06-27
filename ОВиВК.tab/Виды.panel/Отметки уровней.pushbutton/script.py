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

class LevelDescription:
    def __init__(self, name, elevation):
        self.name = name
        self.elevation = elevation

class VISElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves,
                                  BuiltInCategory.OST_PipeCurves,
                                  BuiltInCategory.OST_MechanicalEquipment,
                                  BuiltInCategory.OST_DuctAccessory,
                                  BuiltInCategory.OST_PipeAccessory,
                                  BuiltInCategory.OST_DuctFitting,
                                  BuiltInCategory.OST_PipeFitting]):
            return True

        return False

    def AllowReference(self, reference, position):
        return True


class TypeAnnotationFilter(ISelectionFilter):
    def AllowElement(self, element):
        fam_name = element.Name

        if element.Category.IsId(BuiltInCategory.OST_GenericAnnotation) and fam_name == "ТипАн_Мрк_B4E_Уровень":
            return True
        return False

    def AllowReference(self, reference, position):
        return True


def get_connectors(element):
    """
    Получает коннекторы элемента. Если элемент воздуховод получит только его коннекторы, врезки будут игнорироваться.
    """
    connectors = []

    if isinstance(element, FamilyInstance) and element.MEPModel.ConnectorManager is not None:
        inst_cons = list(element.MEPModel.ConnectorManager.Connectors)

        min_z = min(c.Origin.Z for c in inst_cons)
        max_z = max(c.Origin.Z for c in inst_cons)

        filtered = [c for c in inst_cons if abs(c.Origin.Z - min_z) < 1e-6 or abs(c.Origin.Z - max_z) < 1e-6]
        connectors.extend(filtered)


    if element.InAnyCategory([BuiltInCategory.OST_DuctCurves,
                              BuiltInCategory.OST_PipeCurves,
                              BuiltInCategory.OST_FlexDuctCurves]) and \
            isinstance(element, MEPCurve) and element.ConnectorManager is not None:

        # Если это воздуховод — фильтруем только не Curve-коннекторы. Это завязано на врезки которые тоже падают в список
        # но с нулевым расходом и двунаправленным потоком
        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves]):
            for conn in element.ConnectorManager.Connectors:
                if conn.ConnectorType != ConnectorType.Curve:
                    connectors.append(conn)
        else:
            # Для других категорий (трубы и гибкие воздуховоды) — добавляем все
            connectors.extend(element.ConnectorManager.Connectors)

    return connectors

def get_connector_coordinates(element):
    """
    Получение координат основных коннекторов элемента.
    """

    # Получаем коннекторы
    connectors = get_connectors(element)

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


def get_selected(filter):
    """
    Выделение воздуховодов, труб, арматуры и оборудования.
    """
    # elements = [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]
    #
    # if len(elements) != 0:
    #     return elements

    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            filter,
            "Выберите воздуховоды или трубы, нажмите Finish по окончании"
        )
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        sys.exit()

    elements = [doc.GetElement(r) for r in references]

    if not elements:
        sys.exit()

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
    line = None
    lines = []
    for geo_obj in geom_elem:

        if isinstance(geo_obj, Line):
            line = geo_obj.Reference
            lines.append(line)

    return lines[0]


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


def get_levels_descriptions():
    """Получение списка Z-координат уровней проекта"""
    levels = (FilteredElementCollector(doc).
              OfCategory(BuiltInCategory.OST_Levels).
              WhereElementIsNotElementType().ToElements())

    levels_inst = []
    for level in levels:
        z = level.Elevation
        name = level.Name
        lev_el = LevelDescription(name, z)
        levels_inst.append(lev_el)
    return levels_inst

def interpolate_point_by_z(p1, p2, z_target):
    """Находит точку на отрезке p1-p2 с заданным z_target"""
    if abs(p2.Z - p1.Z) < 1e-6:
        # Практически горизонтальная труба — Z не изменяется
        return None

    t = (z_target - p1.Z) / (p2.Z - p1.Z)
    x = p1.X + t * (p2.X - p1.X)
    y = p1.Y + t * (p2.Y - p1.Y)
    return XYZ(x, y, z_target)


LEFT = "Слева"
RIGHT = "Справа"

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    start_up_checks()

    type_annotation = get_selected(TypeAnnotationFilter())[0]
    orig_pont = type_annotation.Location.Point

    elements = get_selected(VISElementsFilter())
    levels = get_levels_descriptions()
    z_correction = check_base_internal_diff()

    for element in elements:
        coord_list = get_connector_coordinates(element)
        connector_coord_1 = coord_list[0]
        connector_coord_2 = coord_list[1]
        min_z = min(connector_coord_1.Z, connector_coord_2.Z) - z_correction
        max_z = max(connector_coord_1.Z, connector_coord_2.Z) - z_correction

        with (revit.Transaction("BIM: Размещение отметок")):
            for level in levels:
                if not (min_z < level.elevation < max_z):
                    continue


                new_point = XYZ(connector_coord_1.X, connector_coord_1.Y, level.elevation + z_correction)

                translation = new_point - orig_pont
                new_tag_Id = ElementTransformUtils.CopyElement(doc, type_annotation.Id, translation)
                new_tag = doc.GetElement(new_tag_Id[0])
                new_tag.SetParamValue("Имя уровня", level.name)
                string_elevation = str(level.elevation * 304.8/1000)
                new_tag.SetParamValue("Отметка уровня", string_elevation)

    # elements = filter_elements(elements)
    # #print elements[0]
    # with (revit.Transaction("BIM: Размещение отметок")):
    #     type_annotation = type_annotation[0]
    #     orig_pont = type_annotation.Location.Point
    #     new_point = XYZ(orig_pont.X, orig_pont.Y, orig_pont.Z + 5)
    #     translation = new_point - orig_pont
    #     new_tag = ElementTransformUtils.CopyElement(doc, type_annotation.Id, translation)



    # with (revit.Transaction("BIM: Копирование")):
    #     print new_tag
    #
    #
    #     ElementTransformUtils.MoveElement(doc, new_tag[0], translation)


    # elements = filter_elements(elements)
    #
    # if len(elements) == 0:
    #     elements = get_selected()
    #
    # orientation_name = SelectFromList('Выберите ориентацию марок', [LEFT, RIGHT])
    # if orientation_name is None:
    #     sys.exit()
    #
    # levels_z = get_levels_z()
    #
    # with (revit.Transaction("BIM: Размещение отметок")):
    #
    #
    #     # z-коррекция нужна для случаев когда идет расхождение базовой точки и начала проекта. Оно сбоит при заборе
    #     # координат коннекторов, из них мы убираем коррекцию.
    #     # А так же сбой идет при передаче координат в ревит - тут мы наоборот добавляем коррекцию(tag_point)
    #     # Срабатывает и когда базовая точка ниже начала и выше
    #     z_correction = check_base_internal_diff()
    #
    #     for element in elements:
    #         coord_list = get_connector_coordinates(element)
    #         connector_coord_1 = coord_list[0]
    #         connector_coord_2 = coord_list[1]
    #
    #         min_z = min(connector_coord_1.Z, connector_coord_2.Z) - z_correction
    #         max_z = max(connector_coord_1.Z, connector_coord_2.Z) - z_correction
    #
    #         ref = get_reference_from_element(element)
    #         if ref is None:
    #             break
    #
    #         for z in levels_z:
    #             if not (min_z < z < max_z):
    #                 continue
    #
    #             tag_point = interpolate_point_by_z(connector_coord_1, connector_coord_2, z + z_correction)
    #             end_point, bend_point = get_mark_endpoint(orientation_name, tag_point)
    #
    #             doc.Create.NewSpotElevation(
    #                 view,
    #                 ref,
    #                 tag_point, # The point which the spot elevation evaluate.
    #                 bend_point, # The bend point for the spot elevation.
    #                 end_point, # The end point for the spot elevation.
    #                 tag_point, # The actual point on the reference which the spot elevation evaluate.
    #                 True
    #             )
    #
    # if ref is None:
    #     forms.alert(
    #         "На виде не найдены осевые линии. Отметки не были размещены.",
    #         "Ошибка",
    #         exitscript=True)


script_execute()