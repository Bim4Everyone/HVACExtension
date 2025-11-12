# -*- coding: utf-8 -*-
import sys
import clr

clr.AddReference('ProtoGeometry')
clr.AddReference("RevitNodes")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import Revit
import dosymep
import codecs
import math
from operator import attrgetter

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

import System
from System.Collections.Generic import *
from math import hypot

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import Selection
from Autodesk.DesignScript.Geometry import *

import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from pyrevit import forms, DB, revit, script, HOST_APP, EXEC_PARAMS
from rpw.ui.forms import SelectFromList

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig


doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument


class EditorReport:
    """
    Класс для отчета о редактировании элементов.

    Attributes:
        edited_reports (list): Список имен пользователей, редактирующих элементы.
        status_report (str): Сообщение о статусе редактирования.
        edited_report (str): Отчет о редактировании элементов.
    """

    def __init__(self):
        """Инициализация объекта EditorReport."""
        self.edited_reports = []
        self.status_report = ''
        self.edited_report = ''

    def __get_element_editor_name(self, element):
        """
        Возвращает имя пользователя, занявшего элемент, или None.

        Args:
            element (Element): Элемент для проверки.

        Returns:
            str или None: Имя пользователя или None, если элемент не занят.
        """
        user_name = __revit__.Application.Username
        edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
        if edited_by is None:
            return None
        if edited_by.lower() in user_name.lower():
            return None
        return edited_by

    def is_element_edited(self, element):
        """
        Проверяет, заняты ли элементы другими пользователями.

        Args:
            element (Element): Элемент для проверки.
        """
        self.update_status = WorksharingUtils.GetModelUpdatesStatus(doc, element.Id)
        if self.update_status == ModelUpdatesStatus.UpdatedInCentral:
            self.status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию."
        name = self.__get_element_editor_name(element)
        if name is not None and name not in self.edited_reports:
            self.edited_reports.append(name)
            return True
        return False

    def show_report(self):
        """Отображает отчет о редактировании элементов."""
        if len(self.edited_reports) > 0:
            self.edited_report = (
                "Часть элементов занята пользователями: {}".format(", ".join(self.edited_reports))
            )
        if self.edited_report != '' or self.status_report != '':
            report_message = (
                self.status_report +
                ('\n' if (self.edited_report and self.status_report) else '') +
                self.edited_report
            )
            forms.alert(report_message, "Ошибка", exitscript=True)


def select_start_cab(cabinets, point):
    """
    Возвращает шкаф, ближайший к point по координатам X и Y.
    """
    def distance_2d(cab):
        dx = cab.Location.Point.X - point.X
        dy = cab.Location.Point.Y - point.Y
        return math.sqrt(dx*dx + dy*dy)

    nearest_cab = min(cabinets, key=distance_2d)
    return nearest_cab


def sort_cabs(cabinets, point):
    """Группирует шкафы по удаленности от точки старта"""

    if not cabinets:
        return []
    start_cab = select_start_cab(cabinets, point)
    sorted_cabs = [start_cab]

    # Создаём множество/список необработанных шкафов
    remaining = [cab for cab in cabinets if cab != start_cab]

    current = start_cab
    while remaining:
        cx, cy = current.Location.Point.X, current.Location.Point.Y

        # Находим ближайший шкаф по евклидовой дистанции
        next_cab = min(remaining, key=lambda c: hypot(c.Location.Point.X - cx, c.Location.Point.Y - cy))

        sorted_cabs.append(next_cab)
        remaining.remove(next_cab)
        current = next_cab

    return sorted_cabs


def get_fire_cabinets_by_view(view):
    """
    Возвращает список элементов механического оборудования, название семейства которых содержит 'Обр_Шпк'.
    """
    editor_report = EditorReport()

    collector = FilteredElementCollector(doc, view.Id) \
        .OfClass(FamilyInstance) \
        .WherePasses(ElementCategoryFilter(BuiltInCategory.OST_MechanicalEquipment))

    result = []

    for el in collector:
        family_name = el.Symbol.Family.Name
        if "Обр_Шпк" in family_name:
            if not editor_report.is_element_edited(el):
                result.append(el)

    editor_report.show_report()

    return result


BY_VIEW = "Этажная"
CONTINUOUS = "Сквозная"
ADSK_POSITION_PARAM_NAME = "ADSK_Позиция"


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    views = [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]

    views = sorted(views, key=lambda v: v.Name)

    if not views:
        forms.alert("Выделите планы на которых выполняется нумерация шкафов.", "Ошибка", exitscript=True)

    invalid = [
        v for v in views
        if not isinstance(v, DB.View) or v.ViewType != DB.ViewType.FloorPlan
    ]

    if invalid:
        forms.alert(
            "Все выбранные элементы должны быть планами этажа.",
            "Ошибка",
            exitscript=True
        )

    uidoc.ActiveView = views[0]

    point = uidoc.Selection.PickPoint("Выберите стартовую точку для нумерации")

    if point is None:
        script.exit()

    selected_mode = SelectFromList('Вид нумерации:',
                                   [CONTINUOUS,
                                    BY_VIEW])

    if selected_mode is None:
        script.exit()

    with revit.Transaction("BIM: Нумерация шкафов"):
        number = 1
        exception_list = []

        for view_element in views:
            cabinets = get_fire_cabinets_by_view(view_element)

            sorted_cabs = sort_cabs(cabinets, point)

            for cabinet in sorted_cabs:
                try:
                    cabinet.SetParamValue(ADSK_POSITION_PARAM_NAME, str(number))
                    number += 1
                except Exception:
                    # Тут могут быть либо ридонли параметры, либо отсутствие параметра. В любом случае нужен список
                    # таких элементов, чтобы обратить внимание пользователя
                    exception_list.append(cabinet.Id.IntegerValue)

            if selected_mode == BY_VIEW:
                number = 1

    if exception_list:
        print "Следующим шкафам не удалось присвоить нумерацию:"
        print exception_list


script_execute()