#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import sys
import System
import math
import CoefficientCalculator
from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from Autodesk.Revit.DB.Mechanical import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, Selection
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from collections import namedtuple
from collections import defaultdict, OrderedDict

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep_libs.bim4everyone import *

class CalculationMethod:
    name = None
    server_id = None
    server = None
    schema = None
    coefficient_field = None

    def __init__(self, name, server, server_id):
        self.name = name
        self.server = server
        self.server_id = server_id

        self.schema = server.GetDataSchema()
        self.coefficient_field = self.schema.GetField("Coefficient")

class EditorReport:
    edited_reports = []
    status_report = ''
    edited_report = ''

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
            element: Элемент для проверки.
        """

        self.update_status = WorksharingUtils.GetModelUpdatesStatus(doc, element.Id)

        if self.update_status == ModelUpdatesStatus.UpdatedInCentral:
            self.status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию. "

        name = self.__get_element_editor_name(element)
        if name is not None and name not in self.edited_reports:
            self.edited_reports.append(name)
            return True

    def show_report(self):
        if len(self.edited_reports) > 0:
            self.edited_report = ("Часть элементов спецификации занята пользователями: {}"
                                  .format(", ".join(self.edited_reports)))
        if self.edited_report != '' or self.status_report != '':
            report_message = (
                    self.status_report + ('\n' if (self.edited_report and self.status_report) else '')
                    + self.edited_report)
            forms.alert(report_message, "Ошибка", exitscript=True)

class SelectedSystem:
    name = None
    elements = None
    system = None

    def __init__(self, name, elements, system):
        self.name = name
        self.system = system
        self.elements = elements

def get_system_elements():
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count != 1:
        forms.alert(
            "Должна быть выделена одна система воздуховодов.",
            "Ошибка",
            exitscript=True)

    system = doc.GetElement(selected_ids[0])

    if system.Category.IsId(BuiltInCategory.OST_DuctSystem) == False:
        forms.alert(
            "Должна быть выделена одна система воздуховодов.",
            "Ошибка",
            exitscript=True)

    duct_elements = system.DuctNetwork
    system_name = system.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)

    selected_system = SelectedSystem(system_name, duct_elements, system)

    return selected_system

def setup_params():
    revit_params = [cross_section_param, coefficient_param]

    project_parameters = ProjectParameters.Create(doc.Application)
    project_parameters.SetupRevitParams(doc, revit_params)

def get_loss_methods():
    service_id = ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService

    service = ExternalServiceRegistry.GetService(service_id)
    server_ids = service.GetRegisteredServerIds()

    for server_id in server_ids:
        server = get_server_by_id(server_id, service_id)
        name = server.GetName()
        if str(server_id) == calculator.COEFF_GUID_CONST:
            calculation_method = CalculationMethod(name, server, server_id)
            return calculation_method

def get_server_by_id(server_guid, service_id):
    service = ExternalServiceRegistry.GetService(service_id)
    if service is not None and server_guid is not None:
        server = service.GetServer(server_guid)
        if server is not None:
            return server
    return None

def set_method(element, method):
    param = element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
    current_guid = param.AsString()

    if current_guid != calculator.LOSS_GUID_CONST:
        param.Set(method.server_id.ToString())

def set_method_value(element, method, element_coefficients):
    local_section_coefficient = 0

    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        local_section_coefficient = element_coefficients[element.Id]

    param = element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
    current_guid = param.AsString()

    if local_section_coefficient != 0 and current_guid != calculator.LOSS_GUID_CONST:
        element = doc.GetElement(element.Id)

        entity = element.GetEntity(method.schema)

        entity.Set(method.coefficient_field, str(local_section_coefficient))
        element.SetEntity(entity)

def split_elements(system_elements):
    elements = []

    for element in system_elements:
        editor_report.is_element_edited(element)
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting) or element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
            elements.append(element)

    return elements

def get_local_coefficient(fitting, system):
    part_type = fitting.MEPModel.PartType

    if part_type == fitting.MEPModel.PartType.Elbow:

        local_section_coefficient = calculator.get_elbow_coefficient(fitting)

    elif part_type == fitting.MEPModel.PartType.Transition:

        local_section_coefficient = calculator.get_transition_coefficient(fitting)

    elif part_type == fitting.MEPModel.PartType.Tee:

        local_section_coefficient = calculator.get_tee_coefficient(fitting)

    elif part_type == fitting.MEPModel.PartType.TapAdjustable:

        local_section_coefficient = calculator.get_tap_adjustable_coefficient(fitting)

    else:
        local_section_coefficient = 0

    fitting_coefficient_cash[fitting.Id.IntegerValue] = local_section_coefficient
    return local_section_coefficient

def get_network_element_name(element):
    element_name = calculator.element_names.get(element.Id)
    if element_name is not None:
        # Ключ найден, переменная tee_type_name содержит имя
        return element_name

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        name = 'Воздуховод'
    elif element.Category.IsId(BuiltInCategory.OST_FlexDuctCurves):
        name = 'Гибкий воздуховод'
    elif element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        name = 'Воздухораспределитель'
    elif element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
        name = 'Оборудование'
    elif element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        name = 'Фасонный элемент воздуховода'
        if element.MEPModel.PartType == PartType.Elbow:
            name = 'Отвод воздуховода'
        if element.MEPModel.PartType == PartType.Transition:
            name = 'Переход между сечениями'
        if element.MEPModel.PartType == PartType.Tee:
            name = 'Тройник'
        if element.MEPModel.PartType == PartType.TapAdjustable:
            if calculator.is_tap_elbow(element):
                name = 'Отвод'
            else:
                name = "Боковое ответвление"

    else:
        name = 'Арматура'

    return name

def get_network_element_length(section, element_id):
    length = '-'
    try:
        length = section.GetSegmentLength(element_id) * 304.8 / 1000
        length = float('{:.2f}'.format(length))
    except Exception:
        pass
    return length

def get_network_element_coefficient(section, element):
    coefficient = element.GetParamValueOrDefault(coefficient_param)

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        return '-'

    if coefficient is None and element.InAnyCategory([
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_DuctTerminal]):
        return 0

    if coefficient is None:
        coefficient = section.GetCoefficient(element.Id)
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        coefficient = fitting_coefficient_cash[element.Id.IntegerValue]  # КМС

    # Округляем, если есть цифры после запятой
    if isinstance(coefficient, (int, float)):
        coefficient = int(coefficient) if coefficient == int(coefficient) else round(coefficient, 2)

    return str(coefficient)

def get_network_element_real_size(element, element_type):
    def convert_to_meters(value):
        return UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.Meters)

    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        if element.MEPModel.PartType == PartType.TapAdjustable or element.MEPModel.PartType == PartType.Tee:
            tee_params = calculator.tee_params.get(element.Id)
            if tee_params is not None:
                if tee_params.name in [calculator.TEE_SUPPLY_PASS_NAME,
                                       calculator.TEE_EXHAUST_PASS_ROUND_NAME,
                                       calculator.TEE_EXHAUST_PASS_RECT_NAME]:
                    return tee_params.fp
                # Ключ найден, переменная tee_type_name содержит имя
                return tee_params.fo

    size = element.GetParamValueOrDefault(cross_section_param)
    if not size:
        size = element_type.GetParamValueOrDefault(cross_section_param)
    if not size:

        connectors = calculator.get_connectors(element)
        size_variants = []
        for connector in connectors:
            if connector.Shape == ConnectorProfileType.Rectangular:
                size_variants.append(convert_to_meters(connector.Height) * convert_to_meters(connector.Width))
            if connector.Shape == ConnectorProfileType.Round:
                size_variants.append(2 * convert_to_meters(connector.Radius) * math.pi)

        size = min(size_variants)
        UnitUtils.ConvertFromInternalUnits(
            size,
            UnitTypeId.SquareMeters)
    return size

def get_network_element_pressure_drop(section, element, density, velocity, coefficient):
    def calculate_pressure_drop():
        return float(coefficient) * (density * math.pow(velocity, 2)) / 2  # ΔP = ξ * ρ * v² / 2

    # 1. Воздуховоды и гибкие воздуховоды
    if element.InAnyCategory([BuiltInCategory.OST_DuctCurves,
                              BuiltInCategory.OST_FlexDuctCurves]):
        pressure_drop = section.GetPressureDrop(element.Id)
        return UnitUtils.ConvertFromInternalUnits(pressure_drop, UnitTypeId.Pascals)

    # 2. Пробуем взять значение из параметра
    pressure_drop = element.GetParamValueOrDefault("ФОП_ВИС_Потери давления")
    if pressure_drop is not None:
        return pressure_drop

    # 3. Воздухораспределители
    if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctTerminal):
        if coefficient and float(coefficient) != 0:
            return calculate_pressure_drop()
        return 10  # фиксированное значение

    # 4. Фитинги, аксессуары, оборудование
    if element.InAnyCategory([
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_MechanicalEquipment]):
        return calculate_pressure_drop()

    return 0

def show_network_report(data, selected_system, output, density):
    print ('Плотность воздушной среды: ' + str(density) + ' кг/м3')

    output.print_table(table_data=data,
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
                           "Id элемента"],
                       formats=['', '', ''],
                       )

def get_flow(section, element):
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        if element.MEPModel.PartType == PartType.TapAdjustable or element.MEPModel.PartType == PartType.Tee:
            tee_params = calculator.tee_params.get(element.Id)
            if tee_params is not None:

                if tee_params.name in [calculator.TEE_SUPPLY_PASS_NAME,
                                       calculator.TEE_EXHAUST_PASS_ROUND_NAME,
                                       calculator.TEE_EXHAUST_PASS_RECT_NAME]:
                    flow = tee_params.Lp
                else:
                    flow = tee_params.Lo
                return int(flow)

    if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        flow = element.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
    else:
        flow = section.Flow

    # Конвертация из внутренних единиц в м³/ч (кубометры в час)
    flow = UnitUtils.ConvertFromInternalUnits(flow, UnitTypeId.CubicMetersPerHour)

    return int(flow)

def get_velocity(element, flow, real_size):
    velocity = (float(flow) * 1000000)/(3600 * real_size *1000000) #скорость в живом сечении

    return velocity

def round_floats(value):
    if isinstance(value, float):
        return round(value, 3)
    return value

def sort_key(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        return 0
    elif element.Category.IsId(BuiltInCategory.OST_FlexDuctCurves):
        return 1
    elif element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        return 2
    elif element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        return 3
    return 4  # Все остальные

def optimise_data_to_demonstration(data):
    # Шаг 1: Группировка по flow с допуском ±5 и сохранением порядка
    flow_ordered = OrderedDict()

    def find_similar_flow_key(flow):
        for key in flow_ordered:
            if abs(key - flow) <= 5:
                return key
        return None

    for row in data:
        if isinstance(row, list) and len(row) > 4:
            flow = float(row[4])  # Явное приведение для уверенности
            similar_key = find_similar_flow_key(flow)
            if similar_key is not None:
                flow_ordered[similar_key].append(row)
            else:
                flow_ordered[flow] = [row]

    # Шаг 2: Назначаем count по порядку групп
    count = 1
    new_data = []
    for rows in flow_ordered.values():
        for row in rows:
            row[0] = count
            new_data.append(row)
        count += 1

    data[:] = new_data  # Обновляем основной список
    grouped = defaultdict(list)

    # Шаг 2.5: Обновляем max_summ_loss — накапливаемое значение element_loss
    cumulative_loss = 0.0
    for row in data:
        element_loss = float(row[7])
        cumulative_loss += element_loss
        row[8] = str(cumulative_loss)

    # Шаг 3: Вставка заголовков перед каждой группой count
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

    # Шаг 4: Агрегация воздуховодов по count и size
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

    # Шаг 5: Добавляем заголовок в самое начало
    data.insert(0, ['Участок №1'])

    return data

def prepare_section_elements(section):
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

def form_data_list(system, density, output ):
    def get_data_by_element():
        element_type = element.GetElementType()

        length = get_network_element_length(section, element.Id)

        coefficient = get_network_element_coefficient(section, element)

        real_size = get_network_element_real_size(element, element_type)

        flow = get_flow(section, element)

        velocity = get_velocity(element, flow, real_size)

        name = get_network_element_name(element)

        pressure_drop = get_network_element_pressure_drop(section, element, density, velocity, coefficient)

        value = [
            0,  # Count
            name,
            length,
            real_size,
            flow,
            velocity,
            coefficient,
            pressure_drop,
            0,  # Суммарные потери
            output.linkify(element.Id)]

        rounded_value = [round_floats(item) for item in value]

        return rounded_value

    path_numbers = system.GetCriticalPathSectionNumbers()

    critical_path_numbers = []
    for number in path_numbers:
        critical_path_numbers.append(number)
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
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting) and \
            (element.MEPModel.PartType == element.MEPModel.PartType.Cap or
             element.MEPModel.PartType == element.MEPModel.PartType.Union):
        return False
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and get_flow(section, element) == 0:
        return False
    return True

doc = __revit__.ActiveUIDocument.Document  # type: Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

coefficient_param = SharedParamsConfig.Instance.VISLocalResistanceCoef # ФОП_ВИС_КМС
cross_section_param = SharedParamsConfig.Instance.VISCrossSection # ФОП_ВИС_Живое сечение, м2

calculator = CoefficientCalculator.AerodinamicCoefficientCalculator(doc, uidoc, view)
editor_report = EditorReport()
fitting_coefficient_cash = {}
passed_elements = []

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    setup_params()

    selected_system = get_system_elements()

    if selected_system.elements is None:
        forms.alert(
            "Не найдены элементы в системе.",
            "Ошибка",
            exitscript=True)

    system = doc.GetElement(selected_system.system.Id)

    calculator.get_critical_path(system)

    if len(calculator.critical_path_numbers) == 0:
        forms.alert(
            "Не найден диктующий путь, проверьте расчетность системы.",
            "Ошибка",
            exitscript=True)

    network_elements = split_elements(selected_system.elements)

    editor_report.show_report()

    method = get_loss_methods()

    with revit.Transaction("BIM: Установка метода расчета"):
        # Необходимо сначала в отдельной транзакции переключиться на определенный коэффициент, где это нужно
        for element in network_elements:
            set_method(element, method)

    fittings_coefficients = {}
    for element in network_elements:
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            fittings_coefficients[element.Id] = get_local_coefficient(element, system)

    with revit.Transaction("BIM: Пересчет потерь напора"):
        for element in network_elements:
            # устанавливаем 0 на арматуру, чтоб она не убивала расчеты и считаем на фитинги
            set_method_value(element, method, fittings_coefficients)

    # заново забираем систему  через ID, мы в прошлой транзакции обновили потери напора на элементах, поэтому данные
    # на системе могли измениться
    system = doc.GetElement(selected_system.system.Id)
    output = script.get_output()

    settings = DuctSettings.GetDuctSettings(doc)

    density = UnitUtils.ConvertFromInternalUnits(settings.AirDensity, UnitTypeId.KilogramsPerCubicMeter)

    raw_data = form_data_list(system, density, output)

    data = optimise_data_to_demonstration(raw_data)

    show_network_report(data, selected_system, output, density)

script_execute()
