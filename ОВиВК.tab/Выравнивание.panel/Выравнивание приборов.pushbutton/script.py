# coding=utf-8
import math
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from pyrevit import EXEC_PARAMS, revit, forms
from dosymep_libs.bim4everyone import *

_MIN_DISTANCE = 0.001
_MAX_DISTANCE = 4.92126  # 1.5 метра
_MAX_ANGLE = math.radians(30)

class RevitRepository:
    def __init__(self, doc):
        self.doc = doc

    def get_all_radiators(self):
        view = self.doc.ActiveView
        collector = FilteredElementCollector(self.doc, view.Id)
        return [e for e in collector.OfClass(FamilyInstance)
                if "обр_" in e.Symbol.FamilyName.lower() and "завес" not in e.Symbol.FamilyName.lower()]

    def get_radiator_center(self, radiator):
        bbox = radiator.GetBoundingBox()
        return (bbox.Min + bbox.Max) * 0.5 if bbox else radiator.Location.Point

    def get_alignment_lines(self):
        marker_lines = []
        options = Options()
        options.IncludeNonVisibleObjects = True

        for link in FilteredElementCollector(self.doc).OfClass(RevitLinkInstance):
            link_doc = link.GetLinkDocument()
            if not link_doc:
                continue
            transform = link.GetTransform()
            for inst in FilteredElementCollector(link_doc).OfClass(FamilyInstance):
                if not inst.Symbol or "Ант_Маркер_Низ" not in inst.Symbol.FamilyName:
                    continue
                geom = inst.get_Geometry(options)
                if not geom:
                    continue
                for g in geom:
                    if isinstance(g, GeometryInstance):
                        for obj in g.GetInstanceGeometry():
                            if isinstance(obj, Line):
                                style = link_doc.GetElement(obj.GraphicsStyleId)
                                if style and style.Name == "Отопит.приборы_Выравнивание":
                                    p1 = transform.OfPoint(obj.GetEndPoint(0))
                                    p2 = transform.OfPoint(obj.GetEndPoint(1))
                                    marker_lines.append(Line.CreateBound(p1, p2))
        return marker_lines

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    doc = __revit__.ActiveUIDocument.Document
    repo = RevitRepository(doc)

    radiators = repo.get_all_radiators()
    if not radiators:
        forms.alert("Не найдено ни одного радиатора в активном виде.", exitscript=True)

    marker_lines = repo.get_alignment_lines()
    if not marker_lines:
        forms.alert("Не найдено ни одной линии выравнивания.", exitscript=True)

    aligned = 0
    with revit.Transaction("Выравнивание радиаторов"):
        for radiator in radiators:
            center = repo.get_radiator_center(radiator)
            direction = radiator.GetTransform().BasisX.Normalize()

            nearest = None
            nearest_dist = float("inf")

            for line in marker_lines:
                marker_center = (line.GetEndPoint(0) + line.GetEndPoint(1)) * 0.5
                dist = center.DistanceTo(marker_center)
                if dist > _MAX_DISTANCE:
                    continue
                line_dir = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize()
                angle = min(direction.AngleTo(line_dir), abs(math.pi - direction.AngleTo(line_dir)))
                if angle < _MAX_ANGLE:
                    continue
                if dist < nearest_dist:
                    nearest = line
                    nearest_dist = dist

            if nearest:
                marker_center = (nearest.GetEndPoint(0) + nearest.GetEndPoint(1)) * 0.5
                to_marker = marker_center - center
                move_dist = to_marker.DotProduct(direction)
                offset = direction.Multiply(move_dist)
                if offset.GetLength() > _MIN_DISTANCE:
                    ElementTransformUtils.MoveElement(doc, radiator.Id, offset)
                    aligned += 1

    forms.alert("Готово! Выровнено радиаторов: {}".format(aligned), exitscript=True)

script_execute()
