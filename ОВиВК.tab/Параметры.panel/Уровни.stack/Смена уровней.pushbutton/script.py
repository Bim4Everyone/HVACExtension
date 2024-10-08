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

import System
from System.Collections.Generic import *

import Revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import Selection
from Autodesk.DesignScript.Geometry import *
from Autodesk.Revit.UI import TaskDialog

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from pyrevit import revit
from pyrevit import forms
from rpw.ui.forms import SelectFromList

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters


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

def find_new_level(height):
    """ Ищем новый уровень. Здесь мы собираем лист из всех уровней, вычисляем у какого из них минимальное неотрцицательное(при наличии) смещение
    от нашей точки. Он и будет целевым """
    all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

    sorted_levels = sorted(all_levels, key=lambda level: level.GetParamValue(BuiltInParameter.LEVEL_ELEV))

    offsets = []
    for level in sorted_levels:
        offsets.append(height - level.Elevation)

    target_ind = -1
    # мы проходим по всем смещениям и ищем первое не отрицательное, т.к. просто минимальное смещение будет для уровней которые сильно выше реальной отметки(отрицательное)
    for offset in offsets:
        if offset > 0:
            target_ind = offsets.index(offset)

    # если целевой индекс остался -1 - значит мы не нашли нормальных уровней с верным смещением. Берем просто минимальный
    if target_ind == -1:
        target_ind = offsets.index(min(offsets))

    level = sorted_levels[target_ind]
    offset_from_new_level = offsets[target_ind]

    return level, offset_from_new_level

def change_level(element, new_level, new_offset, offset_param, height_param):
    height_param.Set(new_level.Id)
    offset_param.Set(new_offset)
    return element

def get_selected_mode():
    method = forms.SelectFromList.show(["Все элементы на активном виде к ближайшим уровням",
                                        "Все элементы на активном виде к выбранному уровню",
                                        "Выбранные элементы к выбранному уровню"],
                                       title="Выберите метод привязки",
                                       button_name="Применить")
    if method is None:
        sys.exit()
    return method

def get_selected_level(method):
    """ Возвращаем выбранный уровень или False, если режим работы не подразумевает такого """
    if method != 'Все элементы на активном виде к ближайшим уровням':
        selected_view = True

        levelCol = get_elements_by_category(BuiltInCategory.OST_Levels)

        levels = []

        for levelEl in levelCol:
            levels.append(levelEl.Name)

        level_name = forms.SelectFromList.show(levels,
                                               title="Выберите уровень",
                                               button_name="Применить")
        if level_name is None:
            sys.exit()

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
        TaskDialog.Show("Ошибка", "Элементы не выбраны")
        sys.exit()

    return filtered

def execute():
    result_error = []
    result_ok = []

    method = get_selected_mode()
    elements = get_list_of_elements(method)
    level = get_selected_level(method)

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
                        new_level, new_offset = find_new_level(real_height)
                        change_level(element, new_level, new_offset, offset_param, height_param)
                    result_ok.append(element)
                else:
                    result_error.append(element)

if doc.IsFamilyDocument:
    TaskDialog.Show("Ошибка", "Надстройка не предназначена для работы с семействами")
    sys.exit()

execute()