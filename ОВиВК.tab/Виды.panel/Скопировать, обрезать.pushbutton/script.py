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

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView


class DuctPipeVerticalFilter(ISelectionFilter):
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


def get_selected():
    """
    Выделение воздуховодов, труб, арматуры и оборудования.
    """
    elements = [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]

    if len(elements) != 0:
        return elements

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


def start_up_checks():
    """Стартовые проверки"""
    if view.Category is None or not view.ViewType == ViewType.ThreeD:
        forms.alert(
            "Добавление отметок возможно только на 3D-Виде.",
            "Ошибка",
            exitscript=True)


def get_section_box_by_elements(elements):
    """Получаем BoundingBox по элементам для нового вида"""
    min_x = min_y = min_z = float('inf')
    max_x = max_y = max_z = float('-inf')

    for elem in elements:
        bbox = elem.get_BoundingBox(None)
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
        forms.alert(
            "Не удалось определить границы BoundingBox.",
            "Ошибка",
            exitscript=True)


    section_box = BoundingBoxXYZ()
    section_box.Min = XYZ(min_x, min_y, min_z)
    section_box.Max = XYZ(max_x, max_y, max_z)
    section_box.Transform = Transform.Identity  # Обязательно!

    return section_box


def get_unique_name():
    """Формируем уникальное имя нового вида"""
    name = view.Name
    views = (FilteredElementCollector(doc)
             .OfCategory(BuiltInCategory.OST_Views)
             .WhereElementIsNotElementType()
             .ToElements())

    base_name = name
    counter = 1

    existing_names = set(el.Name for el in views if hasattr(el, "Name"))

    while name in existing_names:
        name = "{}_копия {}".format(base_name, counter)
        counter += 1

    return name


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    start_up_checks()

    elements = get_selected()
    new_name = get_unique_name()
    section_box = get_section_box_by_elements(elements)

    with (revit.Transaction("BIM: Скопировать, обрезать")):
        new_view_id = view.Duplicate(ViewDuplicateOption.WithDetailing)
        new_view = doc.GetElement(new_view_id)
        new_view.Name = new_name

        new_view.IsSectionBoxActive = True
        new_view.SetSectionBox(section_box)


    uidoc.RequestViewChange(new_view)


script_execute()