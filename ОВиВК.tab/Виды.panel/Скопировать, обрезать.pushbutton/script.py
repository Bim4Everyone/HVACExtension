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
import re

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BoundingBoxXYZ, XYZ, Transform
from pyrevit import forms, revit, script, HOST_APP, EXEC_PARAMS

from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.DB import BuiltInCategory
from dosymep.Revit.Geometry import *

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView


class VisElementsFilter(ISelectionFilter):
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
            VisElementsFilter(),
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

    def expand_bounding_box(bbox, offset_ft=0.01):
        """Слегка расшриряем bbox чтоб избежать скрытия штриховки поверхности"""
        min_pt = bbox.Min
        max_pt = bbox.Max

        new_min = XYZ(min_pt.X - offset_ft, min_pt.Y - offset_ft, min_pt.Z - offset_ft)
        new_max = XYZ(max_pt.X + offset_ft, max_pt.Y + offset_ft, max_pt.Z + offset_ft)

        bbox.Min = new_min
        bbox.Max = new_max

        return bbox

    bboxes = List[BoundingBoxXYZ]()

    for elem in elements:
        bbox = None
        if elem.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves]):
            pipe_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
            duct_filter = ElementCategoryFilter(BuiltInCategory.OST_DuctInsulations)
            insulation_filter = LogicalOrFilter(pipe_filter, duct_filter)
            sub_elements_ids = elem.GetDependentElements(insulation_filter)
            for sub_element_id in sub_elements_ids:
                sub_element = doc.GetElement(sub_element_id)
                bbox = sub_element.GetBoundingBox()

        if bbox is None:
            bbox = elem.GetBoundingBox()
        if not bbox:
            continue
        bboxes.Add(bbox)

    # Если ни один элемент не дал границ — выходим
    if len(bboxes) == 0:
        forms.alert(
            "Не удалось определить границы BoundingBox.",
            "Ошибка",
            exitscript=True)

    section_box = BoundingBoxExtensions.CreateUnitedBoundingBox(bboxes)
    section_box.Transform = Transform.Identity
    section_box = expand_bounding_box(section_box)

    return section_box


def get_unique_name():
    """Формируем уникальное имя нового вида"""
    name = view.Name
    name = re.sub(r'[\\:{}\[\];<>?\`~]', '', name)
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
    uidoc.Selection.SetElementIds(List[ElementId]()) # сброс выделения. Если не выделить и скопировать один элемент
    # сохранится видимость коннекторов и временных размеров, видимо баг ревита

    with (revit.Transaction("BIM: Скопировать, обрезать")):
        new_view_id = view.Duplicate(ViewDuplicateOption.WithDetailing)
        new_view = doc.GetElement(new_view_id)
        new_view.Name = new_name

        new_view.IsSectionBoxActive = True
        new_view.SetSectionBox(section_box)
        doc.Regenerate()

    uidoc.ActiveView = new_view


script_execute()