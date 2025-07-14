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
from pyrevit import EXEC_PARAMS, revit, forms
from dosymep_libs.bim4everyone import *
from Autodesk.Revit.UI.Selection import ObjectType

_MIN_DISTANCE = 0.001
_MAX_DISTANCE = 9.84252  # 3 метра в футах
_MAX_ANGLE = math.radians(30)

class EditorReport:
    def __init__(self, doc):
        self.doc = doc
        self.edited_reports = []
        self.status_report = ''
        self.edited_report = ''

    def __get_element_editor_name(self, element):
        user_name = __revit__.Application.Username
        edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
        if edited_by is None:
            return None
        if edited_by.lower() in user_name.lower():
            return None
        return edited_by

    def is_element_edited(self, element):
        self.update_status = WorksharingUtils.GetModelUpdatesStatus(self.doc, element.Id)
        if self.update_status == ModelUpdatesStatus.UpdatedInCentral:
            self.status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию."

        name = self.__get_element_editor_name(element)
        if name is not None and name not in self.edited_reports:
            self.edited_reports.append(name)
            return True
        return False

    def show_report(self):
        if len(self.edited_reports) > 0:
            self.edited_report = (
                "Часть элементов занята пользователями: {}".format(", ".join(self.edited_reports))
            )
        if self.edited_report or self.status_report:
            message = self.status_report
            if self.edited_report and self.status_report:
                message += "\n"
            message += self.edited_report
            forms.alert(message, "Ошибка", exitscript=True)

class RevitRepository:
    def __init__(self, doc):
        self.doc = doc

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

    def get_windows(self):
        windows = []
        options = Options()
        options.IncludeNonVisibleObjects = True

        for link in FilteredElementCollector(self.doc).OfClass(RevitLinkInstance):
            link_doc = link.GetLinkDocument()
            if not link_doc:
                continue
            transform = link.GetTransform()
            for inst in FilteredElementCollector(link_doc).OfClass(FamilyInstance):
                if not inst.Symbol or "окн_" not in inst.Symbol.FamilyName.lower():
                    continue

                win_dir = None
                if hasattr(inst, "FacingOrientation"):
                    win_dir = transform.OfVector(inst.FacingOrientation).Normalize()
                elif hasattr(inst.Location, "Curve"):
                    curve = inst.Location.Curve
                    win_dir = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
                    win_dir = transform.OfVector(win_dir).Normalize()
                else:
                    continue
                bbox = inst.GetBoundingBox()
                if bbox:
                    center = (bbox.Min + bbox.Max) * 0.5
                elif hasattr(inst.Location, "Point"):
                    center = inst.Location.Point
                else:
                    continue
                center = transform.OfPoint(center)
                windows.append((center, win_dir, inst))
        return windows

def pick_radiators(doc):
    uidoc = __revit__.ActiveUIDocument
    selection = uidoc.Selection
    try:
        picked = selection.PickObjects(ObjectType.Element, "Выберите радиаторы для выравнивания")
        radiators = []
        for ref in picked:
            element = doc.GetElement(ref.ElementId)

            if isinstance(element, FamilyInstance) and "обр_" in element.Symbol.FamilyName.lower() and "завес" not in element.Symbol.FamilyName.lower():
                radiators.append(element)
        return radiators
    except Exception:
        forms.alert("Выбор элементов отменён.", exitscript=True)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    doc = __revit__.ActiveUIDocument.Document
    repo = RevitRepository(doc)
    report = EditorReport(doc)

    radiators = pick_radiators(doc)
    if not radiators:
        forms.alert("Не выбрано ни одного радиатора.", exitscript=True)

    marker_lines = repo.get_alignment_lines()
    windows = repo.get_windows()

    if not marker_lines and not windows:
        forms.alert("Не найдено ни одной линии выравнивания и ни одного окна.", exitscript=True)

    aligned = 0
    already_aligned_ids = set()

    with revit.Transaction("Выравнивание радиаторов"):
        for radiator in radiators:
            if report.is_element_edited(radiator):
                continue

            center = repo.get_radiator_center(radiator)
            direction = radiator.GetTransform().BasisX.Normalize()

            nearest = None
            nearest_dist = float("inf")
            aligned_type = None

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
                    nearest = marker_center
                    nearest_dir = line_dir
                    nearest_dist = dist
                    aligned_type = "marker"

            if not nearest:
                for win_center, win_dir, _ in windows:
                    dist = center.DistanceTo(win_center)
                    if dist > _MAX_DISTANCE:
                        continue
                    angle = min(direction.AngleTo(win_dir), abs(math.pi - direction.AngleTo(win_dir)))
                    if angle < _MAX_ANGLE:
                        continue
                    if dist < nearest_dist:
                        nearest = win_center
                        nearest_dir = win_dir
                        nearest_dist = dist
                        aligned_type = "window"

            if nearest and radiator.Id not in already_aligned_ids:
                to_marker = nearest - center
                move_dist = to_marker.DotProduct(direction)
                offset = direction.Multiply(move_dist)
                if offset.GetLength() > _MIN_DISTANCE:
                    ElementTransformUtils.MoveElement(doc, radiator.Id, offset)
                    aligned += 1
                    already_aligned_ids.add(radiator.Id)

    report.show_report()
    forms.alert("Готово! Выровнено радиаторов: {}".format(aligned), exitscript=True)

script_execute()