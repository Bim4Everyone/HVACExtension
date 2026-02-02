#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig

from dosymep_libs.bim4everyone import *
from Autodesk.Revit.DB import *
from System.Collections.Generic import List

from pyrevit import revit
from pyrevit import forms
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS


doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView


class AttitudeParametersSet:
    absolute_mid = 0
    absolute_bot = 0
    level_elevation = 0
    from_level_offset = 0


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
            forms.alert(message, "Ошибка")


def setup_params():
    """Настраивает параметры проекта."""
    revit_params = [mark_bottom_to_zero_param, mark_axis_to_zero_param]
    project_parameters = ProjectParameters.Create(doc.Application)

    try:
        project_parameters.SetupRevitParams(doc, revit_params)
    except Exception as e:
        if "Copying one or more elements failed." in str(e):
            pass
        else:
            raise


def sort_parameters_to_group():
    """Сортирует все параметры в группу размеры"""
    group = GroupTypeId.Geometry

    param_names = [
        mark_bottom_to_zero_param.Name,
        mark_axis_to_zero_param.Name,
        ADSK_BOT_PARAM_NAME,
        ADSK_MID_PARAM_NAME,
        ADSK_HOLE_CUR_BOT_PARAM_NAME,
        ADSK_HOLE_CUR_OFFSET_PARAM_NAME,
        ADSK_LEVEL_CUR_OFFSET_PARAM_NAME,
        ADSK_HOLE_BOT_PARAM_NAME,
        ADSK_HOLE_OFFSET_PARAM_NAME,
        ADSK_LEVEL_OFFSET_PARAM_NAME
    ]

    with revit.Transaction("BIM: Настройка параметров"):
        for param_name in param_names:
            param = doc.GetSharedParam(param_name)
            if param is None:
                continue
            definition = param.GetDefinition()
            definition.ReInsertToGroup(doc, group)


def get_elements():
    """ Забирает список труб, воздуховодов и шаблонизированных семейств обобщенных моделей"""
    categories = [
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_GenericModel
    ]

    category_ids = List[ElementId]([ElementId(int(category)) for category in categories])

    multicategory_filter = ElementMulticategoryFilter(category_ids)

    elements = FilteredElementCollector(doc) \
        .WherePasses(multicategory_filter) \
        .WhereElementIsNotElementType() \
        .ToElements()

    filtered_elements = [
        el for el in elements
        if not (
                el.Category.IsId(BuiltInCategory.OST_GenericModel) and
                HOLE_NAME_KEY not in el.GetParam(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
        )
    ]

    return filtered_elements


def get_line_elevations(element):
    """Возвращает смещения воздуховодов и трубопроводов"""

    level_id = element.GetParam(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    level = doc.GetElement(level_id)
    level_elevation = level.GetParamValue(BuiltInParameter.LEVEL_ELEV)

    element_mid_elevation = max(
        element.GetParamValue(BuiltInParameter.RBS_START_OFFSET_PARAM),
        element.GetParamValue(BuiltInParameter.RBS_END_OFFSET_PARAM)
    )

    # Отметки низа работают неочевидно и могут меняться местами, лучше самому их вычислить
    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        outer_diameter = element.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
        element_bot_elevation = element_mid_elevation - outer_diameter / 2
    else:
        center_offset = (element.GetParamValue(BuiltInParameter.RBS_DUCT_TOP_ELEVATION)
                         - element_mid_elevation)

        element_bot_elevation = element_mid_elevation - center_offset

    return element_bot_elevation, element_mid_elevation, level_elevation


def get_generic_elevations(element):
    """Возвращает смещения обобщенных моделей"""

    level_id = element.GetParam(BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId()
    level = doc.GetElement(level_id)
    level_elevation = level.GetParamValue(BuiltInParameter.LEVEL_ELEV)

    is_in_wall = "В стене" in element.GetParam(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()

    # Для круглых отверстий смещаемся в центр
    element_bot_elevation = element.GetParamValue(BuiltInParameter.INSTANCE_ELEVATION_PARAM)

    if is_in_wall:
        diameter = element.GetParamValueOrDefault("ADSK_Размер_Диаметр", 0.0)
        element_bot_elevation + diameter/2

    # у отверстий нет параметра отметки оси
    return element_bot_elevation, element_bot_elevation, level_elevation


def from_millimeters(value):
    """Конвертация из внутренних значений в мм"""
    return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters)


def to_millimeters(value):
    """Конвертация во внутренние значения мм"""
    return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters)


def get_element_attitude(element):
    """Возвращает данные по отметке элементов"""
    if element.Category.IsId(BuiltInCategory.OST_GenericModel):
        element_bot_elevation, element_mid_elevation, level_elevation  = get_generic_elevations(element)
    else:
        element_bot_elevation, element_mid_elevation, level_elevation = get_line_elevations(element)

    level_elevation = from_millimeters(level_elevation)
    element_mid_elevation = from_millimeters(element_mid_elevation)
    element_bot_elevation = from_millimeters(element_bot_elevation)

    absolute_bot = level_elevation + element_bot_elevation
    absolute_mid = level_elevation + element_mid_elevation
    offset = element_mid_elevation # Смещение от уровня

    return (
        absolute_mid, # Абсолютная отметка центра
        absolute_bot, # Абсолютная отметка низа
        level_elevation, # Отметка уровня
        offset # Смещение от уровня
    )


def set_elevation_value(element, absolute_mid, absolute_bot, level_elevation, offset):
    """Устанавливает значения параметров отметок """

    # Имя параметра - функция преобразования
    convert_set = {
        ADSK_BOT_PARAM_NAME: lambda v: to_millimeters(v),
        ADSK_MID_PARAM_NAME: lambda v: to_millimeters(v),
        ADSK_HOLE_BOT_PARAM_NAME: lambda v: to_millimeters(v),
        ADSK_HOLE_OFFSET_PARAM_NAME: lambda v: to_millimeters(v),
        ADSK_LEVEL_OFFSET_PARAM_NAME: lambda v: to_millimeters(v),
    }

    # параметр - значение
    param_value_map = {
        ADSK_MID_PARAM_NAME: absolute_mid,
        mark_axis_to_zero_param: absolute_mid,

        ADSK_BOT_PARAM_NAME: absolute_bot,
        mark_bottom_to_zero_param: absolute_bot,
        ADSK_HOLE_CUR_BOT_PARAM_NAME: absolute_bot,
        ADSK_HOLE_BOT_PARAM_NAME: absolute_bot,

        ADSK_LEVEL_OFFSET_PARAM_NAME: level_elevation,
        ADSK_LEVEL_CUR_OFFSET_PARAM_NAME: level_elevation,

        ADSK_HOLE_OFFSET_PARAM_NAME: offset,
        ADSK_HOLE_CUR_OFFSET_PARAM_NAME: offset
    }

    for param_name, value in param_value_map.items():
        if not element.IsExistsParam(param_name):
            continue

        converter = convert_set.get(param_name, lambda x: x)
        param_value = converter(value)

        if can_set_param_value(element, param_name):
            element.SetParamValue(param_name, param_value)


def can_set_param_value(element, param_name):
    param = element.GetParam(param_name)
    if param.IsReadOnly:
        return False

    in_group = element.GroupId != ElementId.InvalidElementId
    if not in_group:
        return True

    definition = param.Definition
    if isinstance(definition, InternalDefinition):
        return definition.VariesAcrossGroups

    return False


mark_bottom_to_zero_param = SharedParamsConfig.Instance.VISMarkBottomToZero
mark_axis_to_zero_param = SharedParamsConfig.Instance.VISMarkAxisToZero

ADSK_BOT_PARAM_NAME = "ADSK_Отметка низа от нуля"
ADSK_MID_PARAM_NAME = "ADSK_Отметка оси от нуля"

ADSK_HOLE_CUR_BOT_PARAM_NAME = "ADSK_Отверстие_ОтметкаОтНуля"
ADSK_HOLE_CUR_OFFSET_PARAM_NAME = "ADSK_Отверстие_ОтметкаОтЭтажа"
ADSK_LEVEL_CUR_OFFSET_PARAM_NAME = "ADSK_Отверстие_ОтметкаЭтажа"
ADSK_HOLE_BOT_PARAM_NAME = "ADSK_Отверстие_Отметка от нуля"
ADSK_HOLE_OFFSET_PARAM_NAME = "ADSK_Отверстие_Отметка от этажа"
ADSK_LEVEL_OFFSET_PARAM_NAME = "ADSK_Отверстие_Отметка этажа"

HOLE_NAME_KEY = "ОбщМд_Отв_Отверстие_"


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    setup_params()
    sort_parameters_to_group()

    elements = get_elements()
    editor_report = EditorReport(doc)

    with revit.Transaction("BIM: Обновление абсолютной отметки"):
        for element in elements:
            if editor_report.is_element_edited(element):
                continue

            absolute_mid, absolute_bot, level_elevation, offset = get_element_attitude(element)
            set_elevation_value(element, absolute_mid, absolute_bot, level_elevation, offset)

    editor_report.show_report()

script_execute()

