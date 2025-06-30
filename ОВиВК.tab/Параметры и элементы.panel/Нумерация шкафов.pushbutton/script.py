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

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import Selection
from Autodesk.DesignScript.Geometry import *

from collections import defaultdict
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import select_file

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig


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
    angle = 0
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


def group_by_rows(cabinets, selected_mode, y_tolerance=2000):
    """Группирует шкафы в ряды по Y, начиная строго слева от центра и далее по окружности"""

    if not cabinets:
        return []

    # Центр — по X середина, по Y верхняя граница
    min_x = min(cab.xyz.X for cab in cabinets)
    max_x = max(cab.xyz.X for cab in cabinets)
    min_y = min(cab.xyz.Y for cab in cabinets)
    max_y = max(cab.xyz.Y for cab in cabinets)
    center_x = (min_x + max_x) / 2
    center_y = (min_y + max_y) / 2

    has_y_elbow = max_y - min_y < 20000
    has_x_elbow = max_x - min_x < 20000

    if selected_mode in [MIN_Y_CLOCKWISE, MIN_Y_COUNTERCLOCKWISE]:
        # минимальный Y
        start_cab = min(cabinets, key=lambda cab: cab.xyz.Y)
    elif selected_mode in [MIN_X_CLOCKWISE, MIN_X_COUNTERCLOCKWISE]:
        # минимальный X
        start_cab = min(cabinets, key=lambda cab: cab.xyz.X)
    elif selected_mode in [MAX_X_CLOCKWISE, MAX_X_COUNTERCLOCKWISE]:
        # максимальный X
        start_cab = max(cabinets, key=lambda cab: cab.xyz.X)
    else:
        start_cab = max(cabinets, key=lambda cab: cab.xyz.Y)

    # Центр поворота
    if has_x_elbow and not has_y_elbow:
        center_point = XYZ(max_x + 10000, center_y, 0)
    elif not has_x_elbow and has_y_elbow:
        center_point = XYZ(center_x, max_y + 10000, 0)
    else:
        center_point = XYZ(center_x, center_y, 0)

    reverse = selected_mode not in [MIN_Y_COUNTERCLOCKWISE,
                                    MIN_X_COUNTERCLOCKWISE,
                                    MAX_Y_COUNTERCLOCKWISE,
                                    MAX_X_COUNTERCLOCKWISE]

    def get_relative_angle(cab):
        dx = cab.xyz.X - center_point.X
        dy = cab.xyz.Y - center_point.Y
        cab_angle = math.atan2(dy, dx)

        ref_dx = start_cab.xyz.X - center_point.X
        ref_dy = start_cab.xyz.Y - center_point.Y
        start_angle = math.atan2(ref_dy, ref_dx)

        # Относительный угол, приведение к [0, 2π)
        relative = (cab_angle - start_angle) % (2 * math.pi)
        return relative

    for cab in cabinets:
        cab.angle = get_relative_angle(cab)

    sorted_cabs = sorted(cabinets, key=get_relative_angle, reverse=reverse)
    if reverse:
        # Переносим последний элемент в начало
        sorted_cabs = [sorted_cabs[-1]] + sorted_cabs[:-1]

    # Группировка по расстоянию от центра (вдоль луча)
    rows = []
    for cab in sorted_cabs:
        placed = False
        for row in rows:
            ref_cab = row[0]
            vec = XYZ(cab.xyz.X - center_point.X, cab.xyz.Y - center_point.Y, 0)
            ref_vec = XYZ(ref_cab.xyz.X - center_point.X, ref_cab.xyz.Y - center_point.Y, 0)

            diff = vec - ref_vec
            dist = diff.GetLength()

            if dist <= y_tolerance:
                row.append(cab)
                placed = True
                break
        if not placed:
            rows.append([cab])

    return rows


def get_fire_cabinet_equipment():
    """
    Возвращает список элементов механического оборудования, название семейства которых содержит 'Обр_Шпк'.
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


def get_cabinets_by_levels(elements):
    fire_cabinets = []

    for element in elements:
        # Не хочется сначала выполнять IsExists а потом пытаться получить параметр для проверки на рид онли.
        param = element.LookupParameter("ADSK_Позиция")

        if param is None or param.IsReadOnly:
            forms.alert(
                "Параметр экземпляра ADSK_Позиция в части оборудования не существует "
                "или недоступен для редактирования. ID: {}".format(str(element.Id)),
                "Ошибка",
                exitscript=True)

        fire_cabinet = FireCabinet(element)
        fire_cabinets.append(fire_cabinet)

    cabinets_by_level = {}
    for cabinet in fire_cabinets:
        level = cabinet.level_name
        if level not in cabinets_by_level:
            cabinets_by_level[level] = []
        cabinets_by_level[level].append(cabinet)

    sorted_cabinets_by_level = dict(sorted(cabinets_by_level.items(), key=lambda item: item[0]))
    return sorted_cabinets_by_level


def split_elements_by_systems(elements):
    """Делим шкафы по системам"""
    system_para = SharedParamsConfig.Instance.VISSystemName

    systems_dict = defaultdict(list)

    for element in elements:
        system_name = element.GetParamValueOrDefault(system_para)
        if system_name is None:
            forms.alert(
                "У части шкафов не заполнен параметр ФОП_ВИС_Имя системы. Выполните полное обновление.",
                "Ошибка",
                exitscript=True)

        systems_dict[system_name].append(element)

    # Преобразуем словарь в список списков
    split_elements   = list(systems_dict.values())
    return split_elements


MIN_Y_CLOCKWISE = "Минимальным Y, по часовой стрелке"
MIN_Y_COUNTERCLOCKWISE = "Минимальным Y, против часовой стрелки"
MIN_X_CLOCKWISE = "Минимальным X, по часовой стрелке"
MIN_X_COUNTERCLOCKWISE = "Минимальным X, против часовой стрелки"
MAX_Y_CLOCKWISE = "Максимальным Y, по часовой стрелке"
MAX_Y_COUNTERCLOCKWISE = "Максимальным Y, против часовой стрелки"
MAX_X_CLOCKWISE = "Максимальным X, по часовой стрелке"
MAX_X_COUNTERCLOCKWISE = "Максимальным X, против часовой стрелки"


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selected_mode = SelectFromList('Нумерация от шкафа с:', [MIN_Y_CLOCKWISE,
                                                                      MIN_Y_COUNTERCLOCKWISE,
                                                                      MIN_X_CLOCKWISE,
                                                                      MIN_X_COUNTERCLOCKWISE,
                                                                      MAX_Y_CLOCKWISE,
                                                                      MAX_Y_COUNTERCLOCKWISE,
                                                                      MAX_X_CLOCKWISE,
                                                                      MAX_X_COUNTERCLOCKWISE])

    if selected_mode is None:
        sys.exit()

    elements = get_fire_cabinet_equipment()
    split_elements  = split_elements_by_systems(elements)

    with revit.Transaction("BIM: Нумерация шкафов"):
        for system_elements in split_elements:
            sorted_cabinets_by_level = get_cabinets_by_levels(system_elements)

            for level_name in sorted(sorted_cabinets_by_level.keys()):
                number = 1
                cabinets = sorted_cabinets_by_level[level_name]
                rows = group_by_rows(cabinets, selected_mode)

                for row in rows:
                    # В ряду сортируем слева направо по X
                    sorted_row = sorted(row, key=lambda c: c.xyz.X)
                    for cabinet in sorted_row:
                        cabinet.element.SetParamValue("ADSK_Позиция", str(number))
                        number += 1


script_execute()