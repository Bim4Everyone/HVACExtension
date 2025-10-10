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
    project_parameters.SetupRevitParams(doc, revit_params)

def get_elements():
    """ Забираем список элементов арматуры и оборудования """
    categories = [
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_DuctCurves
    ]

    category_ids = List[ElementId]([ElementId(int(category)) for category in categories])

    multicategory_filter = ElementMulticategoryFilter(category_ids)

    elements = FilteredElementCollector(doc) \
        .WherePasses(multicategory_filter) \
        .WhereElementIsNotElementType() \
        .ToElements()

    return elements

def get_element_attitude(element):
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

    def to_meters(value):
        return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Meters)

    level_elevation = to_meters(level_elevation)
    element_mid_elevation = to_meters(element_mid_elevation)
    element_bot_elevation = to_meters(element_bot_elevation)

    absolute_bot = level_elevation + element_bot_elevation
    absolute_mid = level_elevation + element_mid_elevation

    return absolute_mid, absolute_bot

def set_elevation_value(element, absolute_mid, absolute_bot):
    # ADSK версии идут в длине, ФОП_ВИС_ в числе. Нужно конвертировать
    convert_set = {
        ADSK_BOT_PARAM_NAME: UnitTypeId.Meters,
        ADSK_MID_PARAM_NAME: UnitTypeId.Meters
    }

    operations_set = {
        absolute_bot: [ADSK_BOT_PARAM_NAME, mark_bottom_to_zero_param],
        absolute_mid: [ADSK_MID_PARAM_NAME, mark_axis_to_zero_param]
    }

    for value, param_names in operations_set.items():
        for param_name in param_names:
            if not element.IsExistsParam(param_name):
                continue

            convert_type = convert_set.get(param_name)
            if convert_type is not None:
                param_value = UnitUtils.ConvertToInternalUnits(value, convert_type)
            else:
                param_value = value

            element.SetParamValue(param_name, param_value)

mark_bottom_to_zero_param = SharedParamsConfig.Instance.VISMarkBottomToZero
mark_axis_to_zero_param = SharedParamsConfig.Instance.VISMarkAxisToZero
ADSK_BOT_PARAM_NAME = "ADSK_Отметка низа от нуля"
ADSK_MID_PARAM_NAME = "ADSK_Отметка оси от нуля"

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    elements = get_elements()
    editor_report = EditorReport(doc)

    with revit.Transaction("BIM: Обновление абсолютной отметки"):
        for element in elements:
            if editor_report.is_element_edited(element):
                continue
            absolute_mid, absolute_bot = get_element_attitude(element)

            set_elevation_value(element, absolute_mid, absolute_bot)

    editor_report.show_report()

script_execute()

