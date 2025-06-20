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

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

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
from rpw.ui.forms import select_file


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document

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


class FireCabinet:
    position_number = 0
    def __init__(self, element):
        self.element = element
        revit_xyz = element.Location.Point
        x = convert_to_mms(revit_xyz.X)
        y = convert_to_mms(revit_xyz.Y)
        z = convert_to_mms(revit_xyz.Z)

        level_id = element.GetParamValue(BuiltInParameter.FAMILY_LEVEL_PARAM)

        self.level_name = doc.GetElement(level_id).Name
        self.xyz = XYZ(x, y, z)

def convert_to_mms(value):
    """Конвертирует из внутренних значений ревита в миллиметры"""
    result = UnitUtils.ConvertFromInternalUnits(value,
                                                UnitTypeId.Millimeters)
    return result

def group_by_rows_top_down(cabinets, y_tolerance=2000):
    """Разбивает шкафы на ряды по Y (сверху вниз) с учетом допуска"""
    sorted_cabs = sorted(cabinets, key=lambda c: (-c.xyz.Y, c.xyz.X))  # сверху вниз
    rows = []

    for cab in sorted_cabs:
        placed = False
        for row in rows:
            if abs(cab.xyz.Y - row[0].xyz.Y) <= y_tolerance:
                row.append(cab)
                placed = True
                break
        if not placed:
            rows.append([cab])
    return rows

def get_fire_cabinet_equipment():
    """
    Возвращает список элементов механического оборудования, название семейства которых содержит 'Обр_Шпк'.

    :param doc: Текущий документ Revit
    :return: Список элементов FamilyInstance
    """

    editor_report = EditorReport()

    collector = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_MechanicalEquipment) \
        .WhereElementIsNotElementType()

    result = []

    for el in collector:
        if isinstance(el, FamilyInstance):
            family_name = el.Symbol.Family.Name
            if "Обр_Шпк" in family_name:
                if not editor_report.is_element_edited(el):
                    result.append(el)
                else:
                    editor_report.show_report()

    return result

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    elements = get_fire_cabinet_equipment()
    fire_cabinets = []

    for element in elements:
        fire_cabinet = FireCabinet(element)
        fire_cabinets.append(fire_cabinet)

    cabinets_by_level = {}
    for cabinet in fire_cabinets:
        level = cabinet.level_name
        if level not in cabinets_by_level:
            cabinets_by_level[level] = []
        cabinets_by_level[level].append(cabinet)

    sorted_cabinets_by_level = dict(sorted(cabinets_by_level.items(), key=lambda item: item[0]))

    number = forms.ask_for_string(
        default='1',
        prompt='С какого числа стартует нумерация:',
        title="Нумерация шкафов"
    )

    try:
        number = int(number)
    except ValueError:
        forms.alert(
            "Нужно ввести число.",
            "Ошибка",
            exitscript=True)

    if number is None:
        sys.exit()

    with revit.Transaction("BIM: Нумерация шкафов"):
        for level_name in sorted(sorted_cabinets_by_level.keys()):
            cabinets = sorted_cabinets_by_level[level_name]
            rows = group_by_rows_top_down(cabinets)

            # Сортируем ряды сверху вниз (по средней Y, по убыванию)
            rows = sorted(rows, key=lambda row: -sum(c.xyz.Y for c in row) / len(row))

            for row in rows:
                # В ряду сортируем слева направо по X
                sorted_row = sorted(row, key=lambda c: c.xyz.X)
                for cabinet in sorted_row:
                    cabinet.element.SetParamValue("ADSK_Позиция", str(number))
                    number += 1

script_execute()