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
categories = [BuiltInCategory.OST_DuctCurves,
              BuiltInCategory.OST_PipeCurves,
              BuiltInCategory.OST_MechanicalEquipment,
              BuiltInCategory.OST_DuctAccessory,
              BuiltInCategory.OST_PipeAccessory,
              BuiltInCategory.OST_DuctFitting,
              BuiltInCategory.OST_PipeFitting]



class LevelDescription:
    def __init__(self, name, elevation):
        self.name = name
        self.elevation = elevation


class LevelOption(forms.TemplateListItem):
    @property
    def name(self):
        return self.item


class VISElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        if element.InAnyCategory(categories):
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
        filtered_cons = []
        for con in inst_cons:
            if con.Domain in (Domain.DomainHvac, Domain.DomainPiping):
                filtered_cons.append(con)

        min_z = min(c.Origin.Z for c in filtered_cons)
        max_z = max(c.Origin.Z for c in filtered_cons)

        filtered = [c for c in filtered_cons if abs(c.Origin.Z - min_z) < 1e-6 or abs(c.Origin.Z - max_z) < 1e-6]
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

    # Получаем координаты начала и конца элемента через коннекторы
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

    if start_point is None and end_point is None:
        forms.alert("Не удалось получить координаты коннекторов.", "Ошибка", exitscript=True)

    # Если у нас встречается элемент с одним коннектором - получает bbox и добавляем его высоту к стартовой точке
    if end_point is None:
        bbox = element.get_BoundingBox(None)
        min_pt = bbox.Min
        max_pt = bbox.Max
        z_diff = max_pt.Z - min_pt.Z
        end_point = XYZ(start_point.X, start_point.Y, start_point.Z + z_diff)

    start_xyz = XYZ(start_point.X, start_point.Y, start_point.Z)
    end_xyz = XYZ(end_point.X, end_point.Y, end_point.Z)

    return start_xyz, end_xyz


def get_selected():
    """
    Выделение воздуховодов, труб, арматуры и оборудования.
    """

    elements = [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]

    filter_result = []
    for element in elements:
        cat = element.Category
        if cat is None:
            continue

        if element.InAnyCategory(categories):
            filter_result.append(element)

    if len(filter_result) != 0:
        return filter_result

    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            VISElementsFilter(),
            "Выберите воздуховоды или трубы"
        )
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        sys.exit()

    elements = [doc.GetElement(r) for r in references]

    if not elements:
        sys.exit()

    return elements


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
    sorted_levels = sorted(
        levels,
        key=lambda lvl: (not lvl.Name[0].isalpha(), lvl.Name.lower())
    )

    options = [LevelOption(lvl.Name) for lvl in sorted_levels]

    selected = forms.SelectFromList.show(
        options,
        multiselect=True,
        name_attr='name',
        button_name='Выбрать уровни'
    )


    if selected is None:
        sys.exit()

    for level in levels:
        if level.Name not in selected:
            continue
        z = level.Elevation
        name = level.Name
        lower_name = name.lower()
        keyword = "этаж"

        index = lower_name.find(keyword)

        if index != -1:
            name = name[:index + len(keyword)]
        else:
            name = name
        lev_el = LevelDescription(name, z)
        levels_inst.append(lev_el)
    return levels_inst


def get_type_annotation():
    """Получаем размещенную на виде шаблонную ТА, если нет - прекращаем выполнение"""
    generic_annotations = FilteredElementCollector(doc, view.Id) \
        .OfCategory(BuiltInCategory.OST_GenericAnnotation) \
        .WhereElementIsNotElementType() \
        .ToElements()

    target_family_name = "ТипАн_Мрк_B4E_Уровень"

    filtered_annotations = [el for el in generic_annotations if el.Symbol.Family.Name == target_family_name]
    if len(filtered_annotations) == 0:
        forms.alert("Разместите на виде хотя бы одно шаблонное семейство ТипАн_Мрк_B4E_Уровень.",
                    "Ошибка",
                    exitscript=True)

    return filtered_annotations[0]


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    start_up_checks()
    type_annotation = get_type_annotation()
    elements = get_selected()
    orig_pont = type_annotation.Location.Point


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
                new_tag_id = ElementTransformUtils.CopyElement(doc, type_annotation.Id, translation)
                new_tag = doc.GetElement(new_tag_id[0])
                new_tag.SetParamValue("Имя уровня", level.name)
                string_elevation = "{:+.3f}".format(level.elevation * 304.8 / 1000)
                new_tag.SetParamValue("Отметка уровня", string_elevation)

script_execute()