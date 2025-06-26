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
from Autodesk.Revit.DB import BoundingBoxXYZ, XYZ, Outline, Transform
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


class DuctPipeVerticalFilter(ISelectionFilter):
    def AllowElement(self, element):
        cat = element.Category
        if not cat:
            return False

        if cat.Id.IntegerValue in [
            int(BuiltInCategory.OST_DuctCurves),
            int(BuiltInCategory.OST_PipeCurves),
            int(BuiltInCategory.OST_MechanicalEquipment),
            int(BuiltInCategory.OST_DuctAccessory),
            int(BuiltInCategory.OST_PipeAccessory),
            int(BuiltInCategory.OST_DuctFitting),
            int(BuiltInCategory.OST_PipeFitting),
        ]:
            return True
        return False

    def AllowReference(self, reference, position):
        return True


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
    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            DuctPipeVerticalFilter(),
            "Выберите воздуховоды или трубы, нажмите Finish по окончании"
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

def set_section_box_by_elements(view, elements):
    min_x = min_y = min_z = float('inf')
    max_x = max_y = max_z = float('-inf')

    for elem in elements:
        bbox = elem.get_BoundingBox(None)  # Лучше использовать None для получения общей границы
        if not bbox:
            continue

        min_pt = bbox.Min
        max_pt = bbox.Max

        min_x = min(min_x, min_pt.X)
        min_y = min(min_y, min_pt.Y)
        min_z = min(min_z, min_pt.Z)

        max_x = max(max_x, max_pt.X)
        max_y = max(max_y, max_pt.Y)
        max_z = max(max_z, max_pt.Z)

    # Если ни один элемент не дал границ — выходим
    if min_x == float('inf'):
        return

    section_box = BoundingBoxXYZ()
    section_box.Min = XYZ(min_x, min_y, min_z)
    section_box.Max = XYZ(max_x, max_y, max_z)
    section_box.Transform = Transform.Identity  # Обязательно!

    view.SetSectionBox(section_box)

def get_unique_name(name, elements):
    base_name = name
    counter = 1

    existing_names = set(el.Name for el in elements if hasattr(el, "Name"))

    while name in existing_names:
        name = "{}_копия {}".format(base_name, counter)
        counter += 1

    return name


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    start_up_checks()
    views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()

    elements = [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]

    if len(elements) == 0:
        elements = get_selected()
    new_name = get_unique_name(view.Name, views)

    with (revit.Transaction("BIM: Скопировать, обрезать")):
        new_view_id = view.Duplicate(ViewDuplicateOption.WithDetailing)
        new_view = doc.GetElement(new_view_id)
        new_view.Name = new_name

        new_view.IsSectionBoxActive = True
        set_section_box_by_elements(new_view, elements)

    uidoc.RequestViewChange(new_view)



script_execute()