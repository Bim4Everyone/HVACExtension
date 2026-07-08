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

# region Обработка групп

class UngroupedGroup:
    """Данные разгруппированной модели группы для последующей обратной сборки.

    Объект хранит исходный тип группы, временный тип, рабочий набор,
    исходную Location.Point, прямых участников группы и вложенные группы.
    Для вложенных групп используется древовидная структура child_groups.
    """

    group_id = None
    group_name = None
    group_type_id = None
    original_group_type_id = None
    temporary_group_type_id = None
    workset_id_value = None
    location_point = None
    member_ids = None
    child_groups = None

    def __init__(self, group_id, group_name, group_type_id, original_group_type_id,
                 temporary_group_type_id, workset_id_value, location_point,
                 member_ids, child_groups):
        self.group_id = group_id
        self.group_name = group_name
        self.group_type_id = group_type_id
        self.original_group_type_id = original_group_type_id
        self.temporary_group_type_id = temporary_group_type_id
        self.workset_id_value = workset_id_value
        self.location_point = location_point
        self.member_ids = member_ids
        self.child_groups = child_groups

class SplitGroupInfo:
    """Данные временного отделения экземпляра группы от исходного типоразмера.

    Перед разгруппированием экземпляр переводится на временный GroupType,
    чтобы не менять остальные экземпляры исходного типоразмера. После обратной
    сборки группа возвращается на original_group_type_id.
    """

    group_id = None
    original_group_type_id = None
    temporary_group_type_id = None
    location_point = None

    def __init__(self, group_id, original_group_type_id, temporary_group_type_id,
                 location_point):
        self.group_id = group_id
        self.original_group_type_id = original_group_type_id
        self.temporary_group_type_id = temporary_group_type_id
        self.location_point = location_point

def is_model_group(element):
    """Проверяет, что элемент является экземпляром модельной группы.

    Прикрепленные группы исключаются, потому что они не должны участвовать
    в алгоритме разгруппировки и обратной сборки как самостоятельные группы.
    """
    if not isinstance(element, Group):
        return False

    if hasattr(element, "IsAttached") and element.IsAttached:
        return False

    if element.Category is None:
        return True

    return element.Category.Id.IntegerValue == ElementId(BuiltInCategory.OST_IOSModelGroups).IntegerValue

def get_top_model_group(element):
    """Возвращает самую верхнюю модельную группу для элемента.

    Если выбран элемент внутри вложенной группы, функция поднимается по
    GroupId до внешнего контейнера. Это гарантирует, что обработка начинается
    с корневой группы, а вложенные группы разбираются рекурсивно.
    """
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
    """Собирает уникальные корневые модельные группы из набора элементов.

    На вход может прийти сам экземпляр группы или любой элемент внутри нее.
    Повторные попадания одной и той же группы отфильтровываются по ElementId.
    """
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
    """Добавляет участников группы всех уровней вложенности в scope.

    Используется для режима работы с выбранными группами: если пользователь
    выбрал группу, в область допустимой смены уровня попадают ее участники,
    включая элементы внутри вложенных групп.
    """
    for member_id in group.GetMemberIds():
        result.add(member_id.IntegerValue)
        member = doc.GetElement(member_id)
        if is_model_group(member):
            append_group_member_scope_ids(member, result)

def get_scope_element_ids(elements, include_group_members):
    """Возвращает Id элементов, которым разрешено менять уровень.

    Для выбранных групп при include_group_members=True в scope добавляются
    также участники групп. Для режимов по активному виду scope строится по
    видимым элементам и позднее уточняется после разгруппировки.
    """
    result = set()

    for element in elements:
        if element is None:
            continue

        result.add(element.Id.IntegerValue)

        if include_group_members and is_model_group(element):
            append_group_member_scope_ids(element, result)

    return result

def get_active_view_element_ids():
    """Возвращает Id элементов, видимых на активном виде в текущем состоянии транзакции.

    Вызывается после разгруппировки, потому что у участников групп появляются
    самостоятельные Id и их нужно повторно сверить с видимостью на виде.
    """
    result = set()
    visible_element_ids = FilteredElementCollector(doc, doc.ActiveView.Id).ToElementIds()

    for element_id in visible_element_ids:
        result.add(element_id.IntegerValue)

    return result

def filter_elements_by_scope(elements, scope_element_ids):
    """Оставляет только элементы, входящие в выбранный или видовой scope."""
    result = []

    for element in elements:
        if element is None:
            continue

        if element.Id.IntegerValue in scope_element_ids:
            result.append(element)

    return result

def get_element_location_point(element):
    """Возвращает копию Location.Point элемента.

    Точка копируется в новый XYZ, чтобы сохранить координату даже после
    удаления исходного экземпляра группы через UngroupMembers().
    """
    try:
        location = element.Location
        if location is None or not hasattr(location, "Point"):
            return None

        point = location.Point
        if point is None:
            return None

        return XYZ(point.X, point.Y, point.Z)
    except Exception:
        return None

def move_element_to_location_point(element, target_point):
    """Перемещает элемент так, чтобы его Location.Point совпал с target_point.

    Используется после смены GroupType и после NewGroup(), потому что Revit
    может изменить базовую точку группы при пересборке. Если восстановить
    точку не удалось, выбрасывается исключение и транзакция откатывается.
    """
    if element is None or target_point is None:
        return False

    current_point = get_element_location_point(element)
    if current_point is None:
        raise Exception("Can not get current group location point")

    translation = XYZ(
        target_point.X - current_point.X,
        target_point.Y - current_point.Y,
        target_point.Z - current_point.Z)
    if translation.GetLength() < 0.0000001:
        return True

    try:
        ElementTransformUtils.MoveElement(doc, element.Id, translation)
        doc.Regenerate()
    except Exception as error:
        raise Exception("Can not restore group location point: {0}".format(error))

    current_point = get_element_location_point(element)
    if current_point is None or current_point.DistanceTo(target_point) >= 0.0000001:
        raise Exception("Can not restore group location point")

    return True

def split_group_instance(group):
    """Переводит один экземпляр группы на временный типоразмер.

    Исходный GroupType остается нетронутым. Временный тип нужен, чтобы
    разгруппировать и менять элементы только в конкретном экземпляре группы,
    не затрагивая остальные экземпляры исходного типоразмера.
    """
    if group is None or not is_model_group(group):
        return None

    original_group_type = group.GroupType
    location_point = get_element_location_point(group)
    if location_point is None:
        raise Exception("Can not get source group location point")

    base_group_name = get_element_name(original_group_type)
    group_name = get_unique_group_type_name("{} TMP {}".format(
        base_group_name,
        group.Id.IntegerValue))

    if hasattr(group, "Pinned") and group.Pinned:
        group.Pinned = False

    temporary_group_type = original_group_type.Duplicate(group_name)
    group.GroupType = temporary_group_type

    return SplitGroupInfo(
        group.Id,
        original_group_type.Id,
        temporary_group_type.Id,
        location_point)

def group_has_elements_in_scope(group, scope_element_ids):
    """Проверяет, попадает ли группа или любой вложенный участник в scope."""
    if group.Id.IntegerValue in scope_element_ids:
        return True

    for member_id in group.GetMemberIds():
        if member_id.IntegerValue in scope_element_ids:
            return True

        member = doc.GetElement(member_id)
        if is_model_group(member) and group_has_elements_in_scope(member, scope_element_ids):
            return True

    return False

def get_groups_with_elements_in_scope(groups, scope_element_ids):
    """Отбирает группы, которые нужно разгруппировать и собрать обратно.

    Группа попадает в обработку, если в scope находится сам экземпляр группы,
    прямой участник или участник на любой глубине вложенности.
    """
    result = []

    for group in groups:
        if group_has_elements_in_scope(group, scope_element_ids):
            result.append(group)

    return result

def get_unique_group_type_name(base_name):
    """Возвращает уникальное имя временного типоразмера группы."""
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
    """Конвертирует набор ElementId в List[ElementId] для Revit API.

    В список попадают только Id элементов, которые существуют после
    разгруппировки и промежуточных преобразований.
    """
    result = List[ElementId]()

    for element_id in element_ids:
        if doc.GetElement(element_id) is not None:
            result.Add(element_id)

    return result

def ungroup_model_group_tree(group, temporary_group_type_ids):
    """Рекурсивно разгруппировывает группу и все вложенные модельные группы.

    Алгоритм:
    1. Переводит текущий экземпляр на временный GroupType.
    2. Сохраняет исходный тип, рабочий набор, Location.Point и member_ids.
    3. Выполняет UngroupMembers().
    4. Среди прямых участников ищет вложенные группы и повторяет процесс.

    Возвращает дерево UngroupedGroup, которое затем собирается снизу вверх.
    """
    split_group_info = split_group_instance(group)
    if split_group_info is None:
        return None

    if split_group_info.temporary_group_type_id is not None:
        temporary_group_type_ids.append(split_group_info.temporary_group_type_id)

    group_type = group.GroupType
    original_group_type_id = split_group_info.original_group_type_id
    original_group_type = doc.GetElement(original_group_type_id)
    group_name = get_element_name(original_group_type or group_type)
    group_type_id = group_type.Id
    workset_id_value = get_element_workset_id_value(group)

    if hasattr(group, "Pinned") and group.Pinned:
        group.Pinned = False

    member_ids = list(group.UngroupMembers())
    child_groups = []

    for member_id in member_ids:
        member = doc.GetElement(member_id)
        if not is_model_group(member):
            continue

        child_group = ungroup_model_group_tree(member, temporary_group_type_ids)
        if child_group is not None:
            child_groups.append(child_group)

    return UngroupedGroup(
        split_group_info.group_id,
        group_name,
        group_type_id,
        original_group_type_id,
        split_group_info.temporary_group_type_id,
        workset_id_value,
        split_group_info.location_point,
        member_ids,
        child_groups)

def ungroup_model_groups(groups, temporary_group_type_ids):
    """Разгруппировывает набор корневых групп и возвращает их деревья."""
    ungrouped_groups = []

    for group in groups:
        ungrouped_group = ungroup_model_group_tree(group, temporary_group_type_ids)
        if ungrouped_group is not None:
            ungrouped_groups.append(ungrouped_group)

    return ungrouped_groups

def delete_group_types(group_type_ids):
    """Удаляет временные типоразмеры групп после успешной обратной сборки.

    Если тип не удаляется, выбрасывается исключение. Это важно, чтобы
    служебные TMP-типы не оставались в модели незаметно.
    """
    processed_group_type_ids = set()

    for group_type_id in group_type_ids:
        if group_type_id is None:
            continue

        group_type_int_id = group_type_id.IntegerValue
        if group_type_int_id in processed_group_type_ids:
            continue

        processed_group_type_ids.add(group_type_int_id)
        if doc.GetElement(group_type_id) is not None:
            try:
                doc.Delete(group_type_id)
            except Exception as error:
                raise Exception("Can not delete temporary group type {0}: {1}".format(
                    group_type_int_id,
                    error))

def restore_group_original_type(group, original_group_type_id, target_location_point=None):
    """Возвращает пересозданную группу на исходный GroupType и Location.Point.

    NewGroup() создает новый служебный тип. После создания группа переводится
    на сохраненный исходный тип, а затем перемещается в сохраненную точку,
    чтобы координата группы до и после обработки совпадала.
    """
    if group is None:
        return None

    original_group_type = doc.GetElement(original_group_type_id)
    if original_group_type is None:
        raise Exception("Can not find original group type {0}".format(
            original_group_type_id.IntegerValue))

    temporary_group_type_id = group.GroupType.Id
    if temporary_group_type_id.IntegerValue == original_group_type_id.IntegerValue:
        move_element_to_location_point(group, target_location_point)
        return None

    if hasattr(group, "Pinned") and group.Pinned:
        group.Pinned = False

    try:
        group.GroupType = original_group_type
    except Exception as error:
        raise Exception("Can not restore original group type {0}: {1}".format(
            original_group_type_id.IntegerValue,
            error))

    doc.Regenerate()
    move_element_to_location_point(group, target_location_point)
    if group.GroupType.Id.IntegerValue != original_group_type_id.IntegerValue:
        raise Exception("Can not restore original group type {0}".format(
            original_group_type_id.IntegerValue))

    return temporary_group_type_id

def collect_ungrouped_member_elements(ungrouped_group):
    """Собирает все разгруппированные негрупповые элементы из дерева группы.

    Id вложенных групп исключаются из прямых участников родителя, потому что
    после рекурсивной разгруппировки сами группы удалены. Элементы из дочерних
    групп добавляются рекурсивно.
    """
    result = []
    child_group_ids = set()

    for child_group in ungrouped_group.child_groups:
        child_group_ids.add(child_group.group_id.IntegerValue)

    for member_id in ungrouped_group.member_ids:
        if member_id.IntegerValue in child_group_ids:
            continue

        member = doc.GetElement(member_id)
        if member is not None:
            result.append(member)

    for child_group in ungrouped_group.child_groups:
        result.extend(collect_ungrouped_member_elements(child_group))

    return result

def get_recreated_member_ids(ungrouped_group, recreated_child_group_ids):
    """Готовит список участников для обратной сборки группы.

    Старые Id вложенных групп заменяются на Id вновь созданных дочерних групп.
    Остальные существующие Id добавляются без изменений.
    """
    member_ids = List[ElementId]()

    for member_id in ungrouped_group.member_ids:
        member_int_id = member_id.IntegerValue
        if member_int_id in recreated_child_group_ids:
            member_ids.Add(recreated_child_group_ids[member_int_id])
            continue

        if doc.GetElement(member_id) is not None:
            member_ids.Add(member_id)

    return member_ids

def recreate_model_group_tree(ungrouped_group):
    """Рекурсивно собирает дерево групп обратно.

    Дочерние группы создаются первыми. Их новые Id подставляются в member_ids
    родительской группы, после чего создается родительская NewGroup().
    Каждая новая группа возвращается на исходный тип и исходную Location.Point.
    """
    temporary_group_type_ids = []
    recreated_child_group_ids = {}

    for child_group in ungrouped_group.child_groups:
        child_new_group, child_temporary_group_type_ids = recreate_model_group_tree(child_group)
        temporary_group_type_ids.extend(child_temporary_group_type_ids)

        if child_new_group is not None:
            recreated_child_group_ids[child_group.group_id.IntegerValue] = child_new_group.Id

    member_ids = get_recreated_member_ids(ungrouped_group, recreated_child_group_ids)
    if member_ids.Count == 0:
        return None, temporary_group_type_ids

    new_group = doc.Create.NewGroup(member_ids)
    temporary_group_type_id = restore_group_original_type(
        new_group,
        ungrouped_group.original_group_type_id,
        ungrouped_group.location_point)
    set_element_workset(new_group, ungrouped_group.workset_id_value)

    if temporary_group_type_id is not None:
        temporary_group_type_ids.append(temporary_group_type_id)

    return new_group, temporary_group_type_ids

# endregion Обработка групп

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

        groups_to_process = get_groups_with_elements_in_scope(groups, scope_element_ids)
        temporary_group_type_ids = []
        ungrouped_groups = ungroup_model_groups(groups_to_process, temporary_group_type_ids)

        if use_active_view_scope:
            doc.Regenerate()
            group_scope_element_ids = get_active_view_element_ids()
        else:
            group_scope_element_ids = scope_element_ids

        for ungrouped_group in ungrouped_groups:
            member_elements = collect_ungrouped_member_elements(ungrouped_group)
            scoped_member_elements = filter_elements_by_scope(member_elements, group_scope_element_ids)
            filtered_members = filter_elements(scoped_member_elements)
            change_elements_level(filtered_members, level, target_levels, result_ok, result_error)

            new_group, recreated_temporary_group_type_ids = recreate_model_group_tree(ungrouped_group)
            temporary_group_type_ids.extend(recreated_temporary_group_type_ids)
            if new_group is not None:
                result_ok.append(new_group)

        delete_group_types(temporary_group_type_ids)

if doc.IsFamilyDocument:
    forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True )

script_execute()
