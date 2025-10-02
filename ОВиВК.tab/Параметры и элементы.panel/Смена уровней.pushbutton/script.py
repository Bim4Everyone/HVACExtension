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
from rpw.ui.forms import SelectFromList


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *



doc = __revit__.ActiveUIDocument.Document  # type: Document
uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application
uidoc = __revit__.ActiveUIDocument

# типы параметров отвечающих за уровень
built_in_level_params = [BuiltInParameter.RBS_START_LEVEL_PARAM,
                         BuiltInParameter.FAMILY_LEVEL_PARAM,
                         BuiltInParameter.GROUP_LEVEL]

# типы параметров отвечающих за смещение от уровня
built_in_offset_params = [BuiltInParameter.INSTANCE_ELEVATION_PARAM,
                          BuiltInParameter.RBS_OFFSET_PARAM,
                          BuiltInParameter.GROUP_OFFSET_FROM_LEVEL,
                          BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM]

class TargetLevel:
    level_element = None
    level_elevation = None
    level_top_elevation = None

    def __init__(self, element, elevation, top_elevation):
        self.level_elevation = elevation
        self.level_element = element
        self.level_top_elevation = top_elevation

def get_elements_by_category(category):
    """ Возвращает коллекцию элементов по категории """
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

def convert(value):
    """ Преобразует дабл в миллиметры """
    unit_type = DisplayUnitType.DUT_MILLIMETERS
    new_v = UnitUtils.ConvertFromInternalUnits(value, unit_type)
    return new_v

def get_selected_elements(uidoc):
    """ Возвращает выбранные элементы """
    return [uidoc.Document.GetElement(elem_id) for elem_id in uidoc.Selection.GetElementIds()]

def check_is_nested(element):
    """ Проверяет, является ли вложением """
    if hasattr(element, "SuperComponent"):
        if not element.SuperComponent:
            return False
    if hasattr(element, "HostRailingId"):
        return True
    if hasattr(element, "GetStairs"):
        return True
    return False

def get_parameter_if_exist_not_ro(element, built_in_parameters):
    """ Получает параметр, если он существует и если он не ReadOnly """
    for built_in_parameter in built_in_parameters:
        parameter = element.get_Parameter(built_in_parameter)
        if parameter is not None and not parameter.IsReadOnly:
            return built_in_parameter

    return None

def filter_elements(elements):
    """Возвращает фильтрованный от вложений и от свободных от групп список элементов"""
    result = []
    for element in elements:
        if element.GroupId == ElementId.InvalidElementId:
            builtin_level_param = get_parameter_if_exist_not_ro(element, built_in_level_params)
            builtin_offset_param = get_parameter_if_exist_not_ro(element, built_in_offset_params)

            if builtin_offset_param is None:
                continue

            # у гибких элементов есть только базовый уровень, никакой отметки, поэтому дальнейшие фильтры они иначе не пройдут
            if element.InAnyCategory([BuiltInCategory.OST_FlexDuctCurves, BuiltInCategory.OST_FlexPipeCurves]):
                result.append(element)

            if builtin_level_param is None:
                continue

            # Даже если у элемента нашелся builtin - все равно просто параметра может и не быть.
            # Дело в том что для материалов изоляции мы находим RBS_START_LEVEL_PARAM и RBS_OFFSET_PARAM
            # Хотя таких параметров у них не существует
            # IsExistsParam по BuiltIn вернет будто параметр существует
            if not element.IsExistsParam(LabelUtils.GetLabelFor(builtin_level_param)):
                continue

            if not element.IsExistsParam(LabelUtils.GetLabelFor(builtin_offset_param)):
                continue

            # проверяем вложение или нет
            if not check_is_nested(element):
                result.append(element)

    return result

def get_real_height(doc, element, level_param_name, offset_param_name):
    """ Возвращает реальную абсолютную отметку элемента """
    level_id = element.GetParamValue(level_param_name)
    level = doc.GetElement(level_id)
    height_value = level.Elevation
    height_offset_value = element.GetParamValue(offset_param_name)
    real_height = height_value + height_offset_value
    return real_height

def get_height_by_element(doc, element):
    """ Возвращает абсолютную отметку, параметр смещения и параметр уровня """

    level_builtin_param = get_parameter_if_exist_not_ro(element, built_in_level_params)
    offset_builtin_param = get_parameter_if_exist_not_ro(element, built_in_offset_params)

    real_height = get_real_height(doc, element, level_builtin_param, offset_builtin_param)
    level_param = element.GetParam(level_builtin_param)
    offset_param = element.GetParam(offset_builtin_param)

    return [real_height, offset_param, level_param]

def find_new_level(height, target_levels):
    """ Ищем новый уровень. Здесь мы принимаем целевые уровни и смотрим в промежуток между отметками какого из них попадает
     наша отметка. Если дошли до самого верхнего - принимаем его"""

    for target_level in target_levels:
        element_offset = height - target_level.level_elevation
        # У самого верхнего уровня отметка верха - None. Он всегда будет последним из-за сортировки по отметке в методе где мы их собираем
        if target_level.level_top_elevation is None:
            return target_level.level_element, element_offset

        if target_level.level_elevation < height < target_level.level_top_elevation:
            return target_level.level_element, element_offset


def change_level(element, new_level, new_offset, offset_param, height_param):
    height_param.Set(new_level.Id)
    offset_param.Set(new_offset)
    return element

def get_selected_mode():
    method = forms.alert("Выберите метод привязки",
                      options=["Все элементы на активном виде к ближайшим уровням",
                               "Все элементы на активном виде к выбранному уровню",
                               "Выбранные элементы к выбранному уровню"])

    if method is False:
        script.exit()

    return method

def get_selected_level(method):
    """ Возвращаем выбранный уровень или False, если режим работы не подразумевает такого """
    if method != 'Все элементы на активном виде к ближайшим уровням':
        selected_view = True

        levelCol = get_elements_by_category(BuiltInCategory.OST_Levels)

        levels = []

        for levelEl in levelCol:
            levels.append(levelEl.Name)

        levels.sort()

        level_name = forms.SelectFromList.show(levels,
                                               title="Выберите уровень",
                                               button_name="Применить")
        if level_name is None:
            forms.alert("Уровень не выбран", "Ошибка", exitscript=True)

        for levelEl in levelCol:
            if levelEl.Name == level_name:
                level = levelEl
                return level

    return False

def get_list_of_elements(method):
    """ Возвращаем лист элементов в зависимости от выбранного режима работы """
    if method == 'Выбранные элементы к выбранному уровню':
        elements = get_selected_elements(uidoc)
    if (method == 'Все элементы на активном виде к выбранному уровню'
            or method == 'Все элементы на активном виде к ближайшим уровням'):
        elements = FilteredElementCollector(doc, doc.ActiveView.Id)

    filtered = filter_elements(elements)

    if len(filtered) == 0:
        forms.alert("Элементы не выбраны", "Ошибка", exitscript=True)

    return filtered

def get_target_levels_list():
    """ возвращает список целевых уровней, с отметками их низа и верха. Если верха нет - возвращает с None вместо отметки """
    all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    sorted_levels = sorted(all_levels, key=lambda level: level.GetParamValue(BuiltInParameter.LEVEL_ELEV))
    result = []

    for index, level in enumerate(sorted_levels):
        if index + 1 < len(sorted_levels):
            next_level_elevation = sorted_levels[index + 1].Elevation
        else:
            next_level_elevation = None

        result.append(TargetLevel(level, level.Elevation, next_level_elevation))

    return result

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    result_error = []
    result_ok = []
    target_levels = []

    method = get_selected_mode()
    elements = get_list_of_elements(method)
    level = get_selected_level(method)
    if not level:
        target_levels = get_target_levels_list()
    with revit.Transaction("Смена уровней"):
            for element in elements:
                height_result = get_height_by_element(doc, element)

                if height_result:
                    real_height = height_result[0]
                    offset_param = height_result[1]
                    height_param = height_result[2]

                    if level:
                        new_offset = real_height - level.Elevation
                        change_level(element, level, new_offset, offset_param, height_param)
                    else:
                        new_level, new_offset = find_new_level(real_height, target_levels)
                        change_level(element, new_level, new_offset, offset_param, height_param)
                    result_ok.append(element)
                else:
                    result_error.append(element)

if doc.IsFamilyDocument:
    forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True )

script_execute()