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

def get_element_name(element):
    try:
        return element.Name
    except AttributeError:
        pass

    try:
        return Element.Name.GetValue(element)
    except Exception:
        pass

    for built_in_parameter in [BuiltInParameter.ALL_MODEL_TYPE_NAME, BuiltInParameter.SYMBOL_NAME_PARAM]:
        parameter = element.get_Parameter(built_in_parameter)
        if parameter is not None:
            return parameter.AsString()

    raise Exception("Can not get element name")

def set_element_name(element, name):
    try:
        element.Name = name
        return
    except AttributeError:
        pass
    except Exception:
        pass

    try:
        Element.Name.SetValue(element, name)
        return
    except Exception:
        pass

    for built_in_parameter in [BuiltInParameter.ALL_MODEL_TYPE_NAME, BuiltInParameter.SYMBOL_NAME_PARAM]:
        parameter = element.get_Parameter(built_in_parameter)
        if parameter is not None and not parameter.IsReadOnly:
            parameter.Set(name)
            return

    raise Exception("Can not set element name")

def get_element_workset_id_value(element):
    try:
        workset_id = element.WorksetId
        if workset_id is not None:
            return workset_id.IntegerValue
    except Exception:
        pass

    parameter = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
    if parameter is not None:
        return parameter.AsInteger()

    return None

def set_element_workset(element, workset_id_value):
    if workset_id_value is None:
        return

    parameter = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
    if parameter is None or parameter.IsReadOnly:
        return

    try:
        parameter.Set(workset_id_value)
        return
    except Exception:
        pass

    try:
        parameter.Set(ElementId(workset_id_value))
    except Exception:
        pass

class UngroupedGroup:
    group_name = None
    group_type_id = None
    workset_id_value = None
    member_ids = None

    def __init__(self, group_name, group_type_id, workset_id_value, member_ids):
        self.group_name = group_name
        self.group_type_id = group_type_id
        self.workset_id_value = workset_id_value
        self.member_ids = member_ids

def is_model_group(element):
    """Returns True only for model group instances."""
    if not isinstance(element, Group):
        return False

    if hasattr(element, "IsAttached") and element.IsAttached:
        return False

    if element.Category is None:
        return True

    return element.Category.Id.IntegerValue == ElementId(BuiltInCategory.OST_IOSModelGroups).IntegerValue

def get_top_model_group(element):
    """Returns the outer model group for a selected/visible element."""
    if element is None:
        return None

    top_group = element if is_model_group(element) else None
    current_element = element

    while (current_element is not None
            and current_element.GroupId != ElementId.InvalidElementId):
        parent_group = doc.GetElement(current_element.GroupId)
        if parent_group is None:
            break

        if is_model_group(parent_group):
            top_group = parent_group

        current_element = parent_group

    return top_group

def get_model_groups(elements):
    """Collects unique selected/visible model group instances."""
    result = []
    processed_group_ids = set()

    for element in elements:
        group = get_top_model_group(element)
        if group is None:
            continue

        group_id = group.Id.IntegerValue
        if group_id in processed_group_ids:
            continue

        processed_group_ids.add(group_id)
        result.append(group)

    return result

def append_group_member_scope_ids(group, result):
    """Adds direct group members to the processing scope."""
    for member_id in group.GetMemberIds():
        result.add(member_id.IntegerValue)

def get_scope_element_ids(elements, include_group_members):
    """Returns ids of elements that are allowed to receive level updates."""
    result = set()

    for element in elements:
        if element is None:
            continue

        result.add(element.Id.IntegerValue)

        if include_group_members and is_model_group(element):
            append_group_member_scope_ids(element, result)

    return result

def get_active_view_element_ids():
    """Returns ids of elements visible on the active view at the current transaction state."""
    result = set()
    visible_element_ids = FilteredElementCollector(doc, doc.ActiveView.Id).ToElementIds()

    for element_id in visible_element_ids:
        result.add(element_id.IntegerValue)

    return result

def filter_elements_by_scope(elements, scope_element_ids):
    """Keeps only elements that belong to the selected/visible processing scope."""
    result = []

    for element in elements:
        if element is None:
            continue

        if element.Id.IntegerValue in scope_element_ids:
            result.append(element)

    return result

def get_group_instances_by_type_ids(group_type_ids):
    """Returns all model group instances by group type id."""
    result = {}

    if len(group_type_ids) == 0:
        return result

    all_groups = FilteredElementCollector(doc).OfClass(Group).ToElements()
    for group in all_groups:
        if not is_model_group(group):
            continue

        group_type_id = group.GroupType.Id.IntegerValue
        if group_type_id not in group_type_ids:
            continue

        if group_type_id not in result:
            result[group_type_id] = []

        result[group_type_id].append(group)

    return result

def get_group_type_ids(groups):
    result = set()

    for group in groups:
        result.add(group.GroupType.Id.IntegerValue)

    return result

def split_multiple_group_instances(groups):
    """Gives every instance of repeated group types its own numbered group type."""
    group_type_ids = get_group_type_ids(groups)
    groups_by_type = get_group_instances_by_type_ids(group_type_ids)

    for group_type_id, group_instances in groups_by_type.items():
        if len(group_instances) < 2:
            continue

        group_instances = sorted(group_instances, key=lambda group: group.Id.IntegerValue)
        source_group_type = group_instances[0].GroupType
        base_group_name = get_element_name(source_group_type)

        for index, group in enumerate(group_instances):
            group_name = get_unique_group_type_name("{} {}".format(base_group_name, index + 1))

            if hasattr(group, "Pinned") and group.Pinned:
                group.Pinned = False

            if index == 0:
                set_element_name(source_group_type, group_name)
                continue

            new_group_type = source_group_type.Duplicate(group_name)
            group.GroupType = new_group_type

def get_groups_with_elements_in_scope(groups, scope_element_ids):
    """Keeps groups only when their instance or members are in the processing scope."""
    result = []

    for group in groups:
        if group.Id.IntegerValue in scope_element_ids:
            result.append(group)
            continue

        for member_id in group.GetMemberIds():
            if member_id.IntegerValue in scope_element_ids:
                result.append(group)
                break

    return result

def get_unique_group_type_name(base_name):
    existing_names = set()
    group_types = FilteredElementCollector(doc).OfClass(GroupType).ToElements()

    for group_type in group_types:
        existing_names.add(get_element_name(group_type))

    result = base_name
    index = 1

    while result in existing_names:
        result = "{} {}".format(base_name, index)
        index += 1

    return result

def to_element_id_list(element_ids):
    """Converts Python/.NET enumerable to List[ElementId]."""
    result = List[ElementId]()

    for element_id in element_ids:
        if doc.GetElement(element_id) is not None:
            result.Add(element_id)

    return result

def ungroup_model_groups(groups):
    """Ungroups model group instances and stores data required to recreate them."""
    ungrouped_groups = []

    for group in groups:
        group_type = group.GroupType
        group_name = get_element_name(group_type)
        group_type_id = group_type.Id
        workset_id_value = get_element_workset_id_value(group)

        if hasattr(group, "Pinned") and group.Pinned:
            group.Pinned = False

        member_ids = list(group.UngroupMembers())
        ungrouped_groups.append(UngroupedGroup(group_name, group_type_id, workset_id_value, member_ids))

    return ungrouped_groups

def delete_old_group_types(ungrouped_groups):
    """Deletes old group types once all selected/visible instances are ungrouped."""
    processed_group_type_ids = set()

    for ungrouped_group in ungrouped_groups:
        group_type_id = ungrouped_group.group_type_id
        group_type_int_id = group_type_id.IntegerValue
        if group_type_int_id in processed_group_type_ids:
            continue

        processed_group_type_ids.add(group_type_int_id)
        if doc.GetElement(group_type_id) is not None:
            doc.Delete(group_type_id)

def set_group_name(new_group, group_name, recreated_group_types):
    """Restores the original group type name after regrouping."""
    if group_name in recreated_group_types:
        temp_group_type = new_group.GroupType
        try:
            new_group.GroupType = recreated_group_types[group_name]
            if doc.GetElement(temp_group_type.Id) is not None:
                doc.Delete(temp_group_type.Id)
            return
        except Exception:
            pass

    set_element_name(new_group.GroupType, group_name)
    recreated_group_types[group_name] = new_group.GroupType

def recreate_model_group(ungrouped_group, recreated_group_types):
    """Creates a model group from previously ungrouped member ids."""
    member_ids = to_element_id_list(ungrouped_group.member_ids)
    if member_ids.Count == 0:
        return None

    new_group = doc.Create.NewGroup(member_ids)
    set_element_workset(new_group, ungrouped_group.workset_id_value)
    set_group_name(new_group, ungrouped_group.group_name, recreated_group_types)
    return new_group

def filter_elements(elements):
    """Возвращает фильтрованный от вложений и от свободных от групп список элементов"""
    result = []
    for element in elements:
        if element is None or is_model_group(element):
            continue

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

def change_element_level(element, level, target_levels):
    height_result = get_height_by_element(doc, element)

    if not height_result:
        return False

    real_height = height_result[0]
    offset_param = height_result[1]
    height_param = height_result[2]

    if level:
        new_offset = real_height - level.Elevation
        change_level(element, level, new_offset, offset_param, height_param)
    else:
        new_level, new_offset = find_new_level(real_height, target_levels)
        change_level(element, new_level, new_offset, offset_param, height_param)

    return True

def change_elements_level(elements, level, target_levels, result_ok, result_error):
    for element in elements:
        if change_element_level(element, level, target_levels):
            result_ok.append(element)
        else:
            result_error.append(element)

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
    use_active_view_scope = False
    include_group_members_in_scope = False

    """ Возвращаем лист элементов в зависимости от выбранного режима работы """
    if method == 'Выбранные элементы к выбранному уровню':
        elements = get_selected_elements(uidoc)
        include_group_members_in_scope = True
    if (method == 'Все элементы на активном виде к выбранному уровню'
            or method == 'Все элементы на активном виде к ближайшим уровням'):
        elements = FilteredElementCollector(doc, doc.ActiveView.Id)
        use_active_view_scope = True

    elements = list(elements)
    scope_element_ids = get_scope_element_ids(elements, include_group_members_in_scope)
    groups = get_model_groups(elements)
    filtered = filter_elements(elements)

    if len(filtered) == 0 and len(groups) == 0:
        forms.alert("Элементы не выбраны", "Ошибка", exitscript=True)

    return filtered, groups, scope_element_ids, use_active_view_scope

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
    elements, groups, scope_element_ids, use_active_view_scope = get_list_of_elements(method)
    level = get_selected_level(method)
    if not level:
        target_levels = get_target_levels_list()
    with revit.Transaction("Смена уровней"):
        change_elements_level(elements, level, target_levels, result_ok, result_error)

        split_multiple_group_instances(groups)
        groups_to_process = get_groups_with_elements_in_scope(groups, scope_element_ids)
        ungrouped_groups = ungroup_model_groups(groups_to_process)
        delete_old_group_types(ungrouped_groups)

        if use_active_view_scope:
            doc.Regenerate()
            group_scope_element_ids = get_active_view_element_ids()
        else:
            group_scope_element_ids = scope_element_ids

        recreated_group_types = {}
        for ungrouped_group in ungrouped_groups:
            member_elements = [doc.GetElement(member_id) for member_id in ungrouped_group.member_ids]
            scoped_member_elements = filter_elements_by_scope(member_elements, group_scope_element_ids)
            filtered_members = filter_elements(scoped_member_elements)
            change_elements_level(filtered_members, level, target_levels, result_ok, result_error)

            new_group = recreate_model_group(ungrouped_group, recreated_group_types)
            if new_group is not None:
                result_ok.append(new_group)

if doc.IsFamilyDocument:
    forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True )

script_execute()
