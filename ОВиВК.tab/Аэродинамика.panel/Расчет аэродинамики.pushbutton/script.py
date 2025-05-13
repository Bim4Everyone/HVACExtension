#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import math
import CalculatorClassLib
import CrossTeeCalculator
import TransitionElbowCalculator
from pyrevit import forms
from pyrevit import script
from pyrevit import EXEC_PARAMS

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from Autodesk.Revit.DB.Mechanical import *
from pyrevit import revit
from collections import defaultdict, OrderedDict

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep_libs.bim4everyone import *

class CalculationMethod:
    """
    Класс для хранения информации о методе расчета.

    Attributes:
        name (str): Название метода расчета.
        server_id (Guid): Идентификатор сервера.
        server (ExternalService): Объект сервера.
        schema (Schema): Схема данных сервера.
        coefficient_field (Field): Поле коэффициента в схеме данных.
    """

    def __init__(self, name, server, server_id):
        """
        Инициализация объекта CalculationMethod.

        Args:
            name (str): Название метода расчета.
            server (ExternalService): Объект сервера.
            server_id (Guid): Идентификатор сервера.
        """
        self.name = name
        self.server = server
        self.server_id = server_id
        self.schema = server.GetDataSchema()
        self.coefficient_field = self.schema.GetField("Coefficient")

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
                "Часть элементов спецификации занята пользователями: {}".format(", ".join(self.edited_reports))
            )
        if self.edited_report != '' or self.status_report != '':
            report_message = (
                self.status_report +
                ('\n' if (self.edited_report and self.status_report) else '') +
                self.edited_report
            )
            forms.alert(report_message, "Ошибка", exitscript=True)

class SelectedSystem:
    """
    Класс для хранения информации о выбранной системе.

    Attributes:
        name (str): Название системы.
        elements (list): Список элементов системы.
        system (Element): Объект системы.
    """

    def __init__(self, name, elements, system):
        """
        Инициализация объекта SelectedSystem.

        Args:
            name (str): Название системы.
            elements (list): Список элементов системы.
            system (Element): Объект системы.
        """
        self.name = name
        self.system = system
        self.elements = elements

def get_system_elements():
    """
    Получает элементы выбранной системы воздуховодов.

    Returns:
        SelectedSystem: Объект выбранной системы.
    """
    selected_ids = uidoc.Selection.GetElementIds()
    if selected_ids.Count != 1:
        forms.alert(
            "Должна быть выделена одна система воздуховодов.",
            "Ошибка",
            exitscript=True
        )
    system = doc.GetElement(selected_ids[0])
    if not system.Category.IsId(BuiltInCategory.OST_DuctSystem):
        forms.alert(
            "Должна быть выделена одна система воздуховодов.",
            "Ошибка",
            exitscript=True
        )
    duct_elements = system.DuctNetwork
    system_name = system.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
    selected_system = SelectedSystem(system_name, duct_elements, system)
    return selected_system

def setup_params():
    """Настраивает параметры проекта."""
    revit_params = [cross_section_param, coefficient_param, pressure_loss_param]
    project_parameters = ProjectParameters.Create(doc.Application)
    project_parameters.SetupRevitParams(doc, revit_params)

def get_loss_methods():
    """
    Получает метод расчета потерь для фитингов и аксессуаров.

    Returns:
        CalculationMethod: Объект метода расчета.
    """
    service_id = ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService
    service = ExternalServiceRegistry.GetService(service_id)
    server_ids = service.GetRegisteredServerIds()
    for server_id in server_ids:
        server = get_server_by_id(server_id, service_id)
        name = server.GetName()
        if str(server_id) == calc_lib.COEFF_GUID_CONST:
            calculation_method = CalculationMethod(name, server, server_id)
            return calculation_method
    return None

def get_server_by_id(server_guid, service_id):
    """
    Получает сервер по его идентификатору.

    Args:
        server_guid (Guid): Идентификатор сервера.
        service_id (Guid): Идентификатор сервиса.

    Returns:
        ExternalService: Объект сервера или None, если сервер не найден.
    """
    service = ExternalServiceRegistry.GetService(service_id)
    if service is not None and server_guid is not None:
        server = service.GetServer(server_guid)
        if server is not None:
            return server
    return None

def set_calculation_method(element, method):
    """
    Устанавливает метод расчета для элемента.

    Args:
        element (Element): Элемент для установки метода расчета.
        method (CalculationMethod): Объект метода расчета.
    """
    param = element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
    current_guid = param.AsString()
    if current_guid != calc_lib.LOSS_GUID_CONST:
        param.Set(method.server_id.ToString())

def set_coefficient_value(element, method, element_coefficients):
    """
    Устанавливает значение коэффициента для элемента.

    Args:
        element (Element): Элемент для установки коэффициента.
        method (CalculationMethod): Объект метода расчета.
        element_coefficients (dict): Словарь коэффициентов элементов.
    """
    local_section_coefficient = 0
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        local_section_coefficient = element_coefficients[element.Id]
    if element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
        local_section_coefficient = element.GetParamValueOrDefault(coefficient_param, 0.0)
    param = element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
    current_guid = param.AsString()
    if local_section_coefficient != 0 and current_guid != calc_lib.LOSS_GUID_CONST:
        element = doc.GetElement(element.Id)
        entity = element.GetEntity(method.schema)
        entity.Set(method.coefficient_field, str(local_section_coefficient))
        element.SetEntity(entity)

def get_fittings_and_accessory(system_elements):
    """
    Получает фитинги и аксессуары из элементов системы.

    Args:
        system_elements (list): Список элементов системы.

    Returns:
        list: Список фитингов и аксессуаров.
    """
    elements = []
    for element in system_elements:
        editor_report.is_element_edited(element)
        if element.InAnyCategory([BuiltInCategory.OST_DuctFitting,
                                  BuiltInCategory.OST_DuctAccessory,
                                  BuiltInCategory.OST_DuctTerminal]):
            elements.append(element)
    return elements

def calculate_local_coefficient(element):
    """
    Высчитывает локальный коэффициент для фитинга.

    Args:
        element (Element): Фитинг для получения коэффициента.
        system (Element): Система, к которой принадлежит фитинг.

    Returns:
        float: Локальный коэффициент.
    """

    if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        local_section_coefficient = cross_tee_calculator.get_side_hole_coefficient(element)
        fitting_and_terminal_coefficient_cash[element.Id] = local_section_coefficient

        return local_section_coefficient

    part_type = element.MEPModel.PartType
    if part_type == element.MEPModel.PartType.Elbow:
        local_section_coefficient = transition_elbow_calculator.get_elbow_coefficient(element)
    elif part_type == element.MEPModel.PartType.Transition:
        local_section_coefficient = transition_elbow_calculator.get_transition_coefficient(element)
    elif part_type == element.MEPModel.PartType.Tee:
        local_section_coefficient = cross_tee_calculator.get_tee_coefficient(element)
    elif part_type == element.MEPModel.PartType.TapAdjustable:
        has_partner = cross_tee_calculator.get_tap_partner_if_exists(element)

        if has_partner:
            fitting_2, duct_element = has_partner

            if transition_elbow_calculator.is_tap_elbow(element) or transition_elbow_calculator.is_tap_elbow(fitting_2):
                local_section_coefficient = cross_tee_calculator.get_double_tap_tee_coefficient(element, fitting_2,
                                                                                                duct_element)
            else:
                local_section_coefficient = cross_tee_calculator.get_tap_cross_coefficient(element, fitting_2,
                                                                                           duct_element)
        elif transition_elbow_calculator.is_tap_elbow(element):
            local_section_coefficient = transition_elbow_calculator.get_tap_elbow_coefficient(element)
        else:
            local_section_coefficient = cross_tee_calculator.get_tap_tee_coefficient(element)
    elif part_type == element.MEPModel.PartType.Cross:
        local_section_coefficient = cross_tee_calculator.get_cross_coefficient(element)
    else:
        local_section_coefficient = 0
    fitting_and_terminal_coefficient_cash[element.Id] = local_section_coefficient

    return local_section_coefficient

def get_network_element_name(element):
    """
    Получает название элемента сети.

    Args:
        element (Element): Элемент сети.

    Returns:
        str: Название элемента.
    """

    def get_name_addon():
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            return ""

        mark = element.GetParamValueOrDefault("ADSK_Марка") \
               or element_type.GetParamValueOrDefault("ADSK_Марка", "")
        short_name = element.GetParamValueOrDefault("ADSK_Наименование краткое") \
                     or element_type.GetParamValueOrDefault("ADSK_Наименование краткое")

        return short_name or mark or ""

    element_type = element.GetElementType()
    element_name = transition_elbow_calculator.element_names.get(element.Id)
    name_addon = get_name_addon()
    if element_name is None:

        element_name = cross_tee_calculator.element_names.get(element.Id)

    if name_addon == "":
        name_addon = element_type.GetParamValueOrDefault("ADSK_Марка", "")

    if element_name is not None:
        return element_name + name_addon
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        return 'Воздуховод'
    if element.Category.IsId(BuiltInCategory.OST_FlexDuctCurves):
        return 'Гибкий воздуховод'
    if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        return 'Воздухораспределитель' + name_addon
    if element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
        return 'Оборудование ' + name_addon
    if element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
        return 'Арматура ' + name_addon
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        if element.MEPModel.PartType == PartType.Elbow:
            return 'Отвод воздуховода'
        if element.MEPModel.PartType == PartType.Transition:
            return 'Переход между сечениями'
        if element.MEPModel.PartType == PartType.Tee:
            return 'Тройник'
        if element.MEPModel.PartType == PartType.TapAdjustable:
            if transition_elbow_calculator.is_tap_elbow(element):
                return 'Отвод'
            return "Боковое ответвление"

    return "Неопознанный элемент"

def get_network_element_length(section, element_id):
    """
    Получает длину элемента сети.

    Args:
        section (Section): Секция системы.
        element_id (ElementId): Идентификатор элемента.

    Returns:
        float: Длина элемента в метрах.
    """
    try:
        length = section.GetSegmentLength(element_id) * 304.8 / 1000
        return float('{:.2f}'.format(length))
    except Exception:
        return '-'

def get_network_element_coefficient(section, element):
    """
    Получает коэффициент элемента сети.

    Args:
        section (Section): Секция системы.
        element (Element): Элемент сети.

    Returns:
        str: Коэффициент элемента.
    """
    coefficient = element.GetParamValueOrDefault(coefficient_param)

    if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_FlexDuctCurves]):
        return '-'
    if coefficient is None and element.InAnyCategory([
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_MechanicalEquipment]):
        return '0'

    if (coefficient is None or coefficient == 0) and element.InAnyCategory([
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_DuctTerminal]):
        coefficient = fitting_and_terminal_coefficient_cash.get(element.Id, 0)

    if isinstance(coefficient, (int, float)):
        return str(int(coefficient)) if coefficient == int(coefficient) else str(round(coefficient, 2))
    return str(coefficient)

def get_network_element_real_size(element, element_type):
    """
    Получает реальный размер элемента сети.

    Args:
        element (Element): Элемент сети.
        element_type (ElementType): Тип элемента.

    Returns:
        float: Реальный размер элемента в квадратных метрах.
    """

    if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        size = cross_tee_calculator.duct_terminals_sizes.get(element.Id, calc_lib.get_element_area(element))
        return size
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        if element.MEPModel.PartType in [PartType.TapAdjustable, PartType.Tee]:
            tee_params = cross_tee_calculator.cross_tee_params.get(element.Id)

            if tee_params is not None:
                if tee_params.name in [cross_tee_calculator.TEE_SUPPLY_PASS_NAME,
                                       cross_tee_calculator.TEE_EXHAUST_PASS_ROUND_NAME,
                                       cross_tee_calculator.TEE_EXHAUST_PASS_RECT_NAME,
                                       cross_tee_calculator.CROSS_SUPPLY_PASS_RECT_NAME,
                                       cross_tee_calculator.CROSS_EXHAUST_PASS_RECT_NAME,
                                       cross_tee_calculator.CROSS_SUPPLY_PASS_ROUND_NAME,
                                       cross_tee_calculator.CROSS_EXHAUST_PASS_ROUND_NAME]:
                    return tee_params.fp
                return tee_params.fo

        size = calc_lib.get_element_area(element)
        return size
    size = element.GetParamValueOrDefault(cross_section_param)
    if not size:
        size = element_type.GetParamValueOrDefault(cross_section_param)
    if not size:
        size = calc_lib.get_element_area(element)

    return size

def get_network_element_pressure_drop(section, element, density, velocity, coefficient):
    """
    Получает потери напора элемента сети.

    Args:
        section (Section): Секция системы.
        element (Element): Элемент сети.
        density (float): Плотность воздушной среды.
        velocity (float): Скорость воздуха.
        coefficient (str): Коэффициент элемента.

    Returns:
        float: Потери напора элемента в паскалях.
    """
    def calculate_pressure_drop():
        return float(coefficient) * (density * math.pow(velocity, 2)) / 2

    if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_FlexDuctCurves]):
        pressure_drop = section.GetPressureDrop(element.Id)
        return UnitUtils.ConvertFromInternalUnits(pressure_drop, UnitTypeId.Pascals)
    pressure_drop = element.GetParamValueOrDefault(pressure_loss_param)
    if pressure_drop is not None:
        return pressure_drop
    if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctTerminal):
        if coefficient and float(coefficient) != 0:
            return calculate_pressure_drop()
        return 10
    if element.InAnyCategory([BuiltInCategory.OST_DuctFitting,
                              BuiltInCategory.OST_DuctAccessory,
                              BuiltInCategory.OST_MechanicalEquipment]):
        return calculate_pressure_drop()
    return 0

def get_network_element_flow(section, element):
    """
    Получает расход воздуха для элемента сети.

    Args:
        section (Section): Секция системы.
        element (Element): Элемент сети.

    Returns:
        int: Расход воздуха в кубометрах в час.
    """
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        if element.MEPModel.PartType in [PartType.TapAdjustable, PartType.Tee]:
            tee_params = cross_tee_calculator.cross_tee_params.get(element.Id)
            if tee_params is not None:
                if tee_params.name in [cross_tee_calculator.TEE_SUPPLY_PASS_NAME,
                                       cross_tee_calculator.TEE_EXHAUST_PASS_ROUND_NAME,
                                       cross_tee_calculator.TEE_EXHAUST_PASS_RECT_NAME,
                                       cross_tee_calculator.CROSS_SUPPLY_PASS_RECT_NAME,
                                       cross_tee_calculator.CROSS_EXHAUST_PASS_RECT_NAME,
                                       cross_tee_calculator.CROSS_SUPPLY_PASS_ROUND_NAME,
                                       cross_tee_calculator.CROSS_EXHAUST_PASS_ROUND_NAME]:
                    return int(tee_params.Lc)
                return int(tee_params.Lp)
    if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        terminal_flow = cross_tee_calculator.duct_terminals_flows.get(element.Id)
        if terminal_flow is not None:
            return int(terminal_flow) # Возвращаем сразу, он уже в метрах кубических
        else:
            flow = element.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
    else:
        flow = section.Flow
    flow = UnitUtils.ConvertFromInternalUnits(flow, UnitTypeId.CubicMetersPerHour)
    return int(flow)

def get_network_element_velocity(element, flow, real_size):
    """
    Получает скорость воздуха в живом сечении элемента.

    Args:
        element (Element): Элемент сети.
        flow (int): Расход воздуха.
        real_size (float): Реальный размер элемента.

    Returns:
        float: Скорость воздуха в метрах в секунду.
    """
    return (float(flow) * 1000000) / (3600 * real_size * 1000000)

def show_network_report(data, selected_system, output, density):
    """
    Отображает отчет о расчете аэродинамики системы.

    Args:
        data (list): Данные для отчета.
        selected_system (SelectedSystem): Выбранная система.
        output (Output): Объект для вывода отчета.
        density (float): Плотность воздушной среды.
    """
    print('Плотность воздушной среды: ' + str(density) + ' кг/м3')
    output.print_table(
        table_data=data,
        title=("Отчет о расчете аэродинамики системы " + selected_system.name),
        columns=[
            "Номер участка",
            "Наименование элемента",
            "Длина, м.п.",
            "Размер, м2",
            "Расход, м3/ч",
            "Скорость, м/с",
            "КМС",
            "Потери напора элемента, Па",
            "Суммарные потери напора, Па",
            "Id элемента"
        ],
        formats=['', '', '']
    )

def prepare_data_to_demonstration(data):
    """
    Оптимизирует данные для демонстрации.

    Args:
        data (list): Данные для оптимизации.

    Returns:
        list: Оптимизированные данные.
    """

    data.sort(key=lambda row: float(row[4]) if isinstance(row, list) and len(row) > 4 else float('inf'))

    flow_ordered = OrderedDict()

    def find_similar_flow_key(flow):
        for key in flow_ordered:
            if abs(key - flow) <= 5:
                return key
        return None


    for row in data:
        if isinstance(row, list) and len(row) > 4:
            flow = float(row[4])
            similar_key = find_similar_flow_key(flow)
            if similar_key is not None:
                flow_ordered[similar_key].append(row)
            else:
                flow_ordered[flow] = [row]

    count = 1
    new_data = []
    for rows in flow_ordered.values():
        for row in rows:
            row[0] = count
            new_data.append(row)
        count += 1

    data[:] = new_data
    grouped = defaultdict(list)

    cumulative_loss = 0.0
    for row in data:
        element_loss = float(row[7])
        cumulative_loss += element_loss
        row[8] = str(cumulative_loss)

    i = 0
    old_count = None
    while i < len(data):
        row = data[i]
        if isinstance(row, list) and len(row) > 1:
            count = row[0]
            if old_count is not None and count != old_count:
                data.insert(i, ['Участок №' + str(count)])
                i += 1
            old_count = count
        i += 1

    for row in data:
        if isinstance(row, list) and len(row) > 1:
            count, name, length, size, flow, velocity, coefficient, element_loss, summ_loss, id = row
            if "Воздуховод" in name:
                grouped[(count, size)].append(row)

    for (count, size), group_rows in grouped.items():
        if len(group_rows) > 1:
            total_length = sum(float(row[2]) for row in group_rows)
            total_element_loss = sum(float(row[7]) for row in group_rows)
            max_summ_loss = max(float(row[8]) for row in group_rows)
            combined_ids = ",".join(str(row[9]) for row in group_rows)

            base_row = group_rows[0]
            base_row[2] = str(total_length)
            base_row[7] = str(total_element_loss)
            base_row[8] = str(max_summ_loss)
            base_row[9] = combined_ids

            for row in group_rows[1:]:
                data.remove(row)

    data.insert(0, ['Участок №1'])

    last_summ_loss = 0.0
    for row in reversed(data):
        if isinstance(row, list) and len(row) > 8:
            try:
                last_summ_loss = float(row[8])
                break
            except ValueError:
                continue

    # Создаём строку "Итого"
    total_row = [""]+["Итого, Па"] + [""] * 6 + [str(round(last_summ_loss, 2))] + [""]

    # Создаём строку "Итого + 15%"
    total_row_15 = [""]+ ["Итого, Па + 15%"] + [""] * 6 + [str(round(last_summ_loss * 1.15, 2))] + [""]

    data.append(total_row)
    data.append(total_row_15)

    return data

def prepare_section_elements(section):
    """
    Подготавливает элементы секции системы, сортируя их по категориям.

    Args:
        section (Section): Секция системы.

    Returns:
        list: Список элементов секции.
    """

    def sort_key(element):
        """
        Возвращает ключ сортировки для элемента.

        Args:
            element (Element): Элемент для сортировки.

        Returns:
            int: Ключ сортировки.
        """
        if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
            return 0
        if element.Category.IsId(BuiltInCategory.OST_FlexDuctCurves):
            return 1
        if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
            return 2
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            return 3
        return 4

    elements_ids = section.GetElementIds()
    segment_elements = []
    for element_id in elements_ids:
        if element_id in passed_elements:
            continue
        element = doc.GetElement(element_id)
        if not element.Category.IsId(BuiltInCategory.OST_DuctCurves):
            passed_elements.append(element_id)
        segment_elements.append(element)
    segment_elements.sort(key=sort_key)
    return segment_elements

def form_raw_data_list(system, density, output):
    """
    Формирует список данных для отчета.

    Args:
        system (Element): Система для формирования данных.
        density (float): Плотность воздушной среды.
        output (Output): Объект для вывода отчета.

    Returns:
        list: Список данных для отчета.
    """

    def round_floats(float_value):
        if isinstance(float_value, float):
            return round(float_value, 3)
        return float_value

    def get_data_by_element():
        element_type = element.GetElementType()
        length = get_network_element_length(section, element.Id)
        coefficient = get_network_element_coefficient(section, element)
        real_size = get_network_element_real_size(element, element_type)
        flow = get_network_element_flow(section, element)
        velocity = get_network_element_velocity(element, flow, real_size)
        name = get_network_element_name(element)
        pressure_drop = get_network_element_pressure_drop(section, element, density, velocity, coefficient)
        value = [
            0,
            name,
            length,
            real_size,
            flow,
            velocity,
            coefficient,
            pressure_drop,
            0,
            output.linkify(element.Id)
        ]
        rounded_value = [round_floats(item) for item in value]
        return rounded_value

    path_numbers = system.GetCriticalPathSectionNumbers()
    critical_path_numbers = list(path_numbers)
    if system.SystemType == DuctSystemType.SupplyAir:
        critical_path_numbers.reverse()

    data = []
    for number in critical_path_numbers:
        section = system.GetSectionByNumber(number)
        segment_elements = prepare_section_elements(section)
        for element in segment_elements:
            if not pass_data_filter(element, section):
                continue
            value = get_data_by_element()
            data.append(value)
    return data

def pass_data_filter(element, section):
    """
    Проверяет, проходит ли элемент фильтр данных.
    Он не должен быть соединением или заглушкой, не должен быть врезкой-партнером и
    посекционного воздуховода должен быть расход не 0.

    Args:
        element (Element): Элемент для проверки.
        section (Section): Секция системы.

    Returns:
        bool: True, если элемент проходит фильтр, иначе False.
    """
    if (element.Category.IsId(BuiltInCategory.OST_DuctFitting) and
            element.MEPModel.PartType in [element.MEPModel.PartType.Cap,
                                          element.MEPModel.PartType.Union]):
        return False
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and get_network_element_flow(section, element) == 0:
        return False
    if element.Id in cross_tee_calculator.tap_crosses_filtered:
        return False

    return True

def process_method_setup(selected_system):
    """
    Обрабатывает настройку метода расчета для выбранной системы.

    Args:
        selected_system (SelectedSystem): Выбранная система.
    """
    if selected_system.elements is None:
        forms.alert(
            "Не найдены элементы в системе.",
            "Ошибка",
            exitscript=True
        )

    network_elements = get_fittings_and_accessory(selected_system.elements)
    editor_report.show_report()
    specific_coefficient_method = get_loss_methods()
    with revit.Transaction("BIM: Установка метода расчета"):
        for element in network_elements:
            set_calculation_method(element, specific_coefficient_method)

    system = doc.GetElement(selected_system.system.Id)
    calc_lib.get_critical_path(system)
    cross_tee_calculator.get_critical_path(system)
    transition_elbow_calculator.get_critical_path(system)

    if len(calc_lib.critical_path_numbers) == 0:
        forms.alert(
            "Не найден диктующий путь, проверьте расчетность системы.",
            "Ошибка",
            exitscript=True
        )

    elements_coefficients = {}
    for element in network_elements:
        if element.InAnyCategory([BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctTerminal]):
            elements_coefficients[element.Id] = calculate_local_coefficient(element)

    with revit.Transaction("BIM: Установка коэффициентов"):
        for element in network_elements:
            if element.InAnyCategory([BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctAccessory]):
                set_coefficient_value(element, specific_coefficient_method, elements_coefficients)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

coefficient_param = SharedParamsConfig.Instance.VISLocalResistanceCoef
cross_section_param = SharedParamsConfig.Instance.VISCrossSection
pressure_loss_param = SharedParamsConfig.Instance.VISPressureLoss

calc_lib = CalculatorClassLib.AerodinamicCoefficientCalculator(doc, uidoc, view)
cross_tee_calculator = CrossTeeCalculator.CrossTeeCoefficientCalculator(doc, uidoc, view)
transition_elbow_calculator = TransitionElbowCalculator.TransitionElbowCoefficientCalculator(doc, uidoc, view)
editor_report = EditorReport()
fitting_and_terminal_coefficient_cash = {}
passed_elements = []

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    setup_params()
    selected_system = get_system_elements()
    process_method_setup(selected_system) # Ставим метод расчета Определенный коэффициент и заполняем его для фитингов

    # заново забираем систему  через ID, мы в прошлой транзакции обновили потери напора на элементах, поэтому данные
    # на системе могли измениться
    selected_system = get_system_elements()
    system = doc.GetElement(selected_system.system.Id)
    output = script.get_output()
    settings = DuctSettings.GetDuctSettings(doc)
    density = UnitUtils.ConvertFromInternalUnits(settings.AirDensity, UnitTypeId.KilogramsPerCubicMeter)
    raw_data = form_raw_data_list(system, density, output)
    data = prepare_data_to_demonstration(raw_data)

    show_network_report(data, selected_system, output, density)

    output.print_md('**<span style="color:red; text-decoration:underline;">'
                    'РАСЧЕТ НАХОДИТСЯ НА СТАДИИ ТЕСТИРОВАНИЯ. '
                    'ПЕРЕПРОВЕРЬТЕ РЕЗУЛЬТАТЫ.</span>**')

script_execute()