#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пересчет КМС'
__doc__ = "Пересчитывает КМС соединительных деталей воздуховодов"

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
        if name is not None and name not in edited_reports:
            self.edited_reports.append(name)
            return True

    def show_report(self):
        if len(self.edited_reports) > 0:
            self.edited_report = ("Часть элементов спецификации занята пользователями: {}"
                                  .format(", ".join(self.edited_reports)))
        if self.edited_report != '' or self.status_report != '':
            report_message = status_report + ('\n' if (edited_report and status_report) else '') + self.edited_report
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

def set_method_value(element, method, value = 0):
    local_section_coefficient = 0

    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        local_section_coefficient = get_local_coefficient(element)

    param = element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
    current_guid = param.AsString()

    if local_section_coefficient != 0 and current_guid != calculator.LOSS_GUID_CONST:
        element = doc.GetElement(element.Id)

        entity = element.GetEntity(method.schema)

        entity.Set(method.coefficient_field, str(value))
        element.SetEntity(entity)

def split_elements(system_elements):
    elements = []

    for element in system_elements:
        editor_report.is_element_edited(element)
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting) or element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
            elements.append(element)

    return elements

def get_local_coefficient(fitting):
    part_type = fitting.MEPModel.PartType

    if part_type == fitting.MEPModel.PartType.Elbow:
        local_section_coefficient = calculator.get_coef_elbow(fitting)
    elif part_type == fitting.MEPModel.PartType.Transition:
        local_section_coefficient = calculator.get_coef_transition(fitting)
    elif part_type == fitting.MEPModel.PartType.Tee:
        local_section_coefficient = calculator.get_coef_tee(fitting)
    elif part_type == fitting.MEPModel.PartType.TapAdjustable:
        local_section_coefficient = calculator.get_coef_tap_adjustable(fitting)
    else:
        local_section_coefficient = 0

    fitting_coefficient_cash[fitting.Id.IntegerValue] = local_section_coefficient
    return local_section_coefficient

def get_network_element_name(element, changing_flow):
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        name = 'Воздуховод'
    elif element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        name = 'Воздухораспределитель'
    elif element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
        name = 'Оборудование'
    elif element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        name = 'Фасонный элемент воздуховода'
        if str(element.MEPModel.PartType) == 'Elbow':
            name = 'Отвод воздуховода'
        if str(element.MEPModel.PartType) == 'Transition':
            name = 'Переход между сечениями'
        if str(element.MEPModel.PartType) == 'Tee':
            name = 'Тройник'
        if str(element.MEPModel.PartType) == 'TapAdjustable':
            if changing_flow:
                name = "Боковое ответвление"
            else:
                name = 'Врезка'
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
    coefficient = element.GetParamValueOrDefault('ФОП_ВИС_КМС')

    if coefficient is None:
        coefficient = section.GetCoefficient(element.Id)
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        coefficient = fitting_coefficient_cash[element.Id.IntegerValue]  # КМС

    return coefficient

def get_network_element_real_size(element, element_type):
    def convert_to_meters(value):
        return UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.Meters)

    size = element.GetParamValueOrDefault('ФОП_ВИС_Живое сечение, м2')
    if not size:
        size = element_type.GetParamValueOrDefault('ФОП_ВИС_Живое сечение, м2')
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

def get_network_element_pressure_drop(section, element, density, velocity):
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        Pd = (density * velocity * velocity) / 2  # Динамическое давление

        K = fitting_coefficient_cash[element.Id.IntegerValue]  # КМС
        pressure_drop = Pd * K
    else:
        pressure_drop = section.GetPressureDrop(element.Id) * 3.280839895

    return pressure_drop

def show_network_report(data, selected_system, output):

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

def get_flow(section):
    flow = section.Flow * 101.941317259
    flow = int(flow)

    return flow

def get_velocity(flow, real_size):

    velocity = (float(flow) * 1000000)/(3600 * real_size *1000000) #скорость в живом сечении

    return velocity

def round_floats(value):
    if isinstance(value, float):
        return round(value, 3)
    return value


doc = __revit__.ActiveUIDocument.Document  # type: Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

calculator = CoefficientCalculator.Aerodinamiccoefficientcalculator(doc, uidoc, view)
editor_report = EditorReport()
fitting_coefficient_cash = {}
passed_elements = []

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selected_system = get_system_elements()

    if selected_system.elements is None:
        forms.alert(
            "Не найдены элементы в системе.",
            "Ошибка",
            exitscript=True)

    network_elements = split_elements(selected_system.elements)

    editor_report.show_report()

    method = get_loss_methods()

    with revit.Transaction("BIM: Установка метода расчета"):
        # Необходимо сначала в отдельной транзакции переключиться на определенный коэффициент, где это нужно
        for element in network_elements:
            set_method(element, method)

    with revit.Transaction("BIM: Пересчет потерь напора"):
        for element in network_elements:
            # устанавливаем 0 на арматуру, чтоб она не убивала расчеты и считаем на фитинги
            set_method_value(element, method)

    with revit.Transaction("BIM: Вывод отчета"):
        # заново забираем систему  через ID, мы в прошлой транзакции обновили потери напора на элементах, поэтому данные
        # на системе могли измениться
        system = doc.GetElement(selected_system.system.Id)
        path_numbers = system.GetCriticalPathSectionNumbers()

        critical_path_numbers = []
        for number in path_numbers:
            critical_path_numbers.append(number)
        if system.SystemType == DuctSystemType.SupplyAir:
            critical_path_numbers.reverse()

        data = []
        count = 0

        output = script.get_output()

        old_flow = 0

        settings = DuctSettings.GetDuctSettings(doc)
        density = settings.AirDensity * 35.3146667215
        print 'Плотность воздушной среды: ' + str(density) + ' кг/м3'

        pressure_total = 0
        for number in critical_path_numbers:
            section = system.GetSectionByNumber(number)
            count += 1
            elements_ids = section.GetElementIds()
            for element_id in elements_ids:
                if element_id in passed_elements:
                    continue

                element = doc.GetElement(element_id)
                if not element.Category.IsId(BuiltInCategory.OST_DuctCurves):
                    passed_elements.append(element_id)

                element_type = element.GetElementType()

                length = get_network_element_length(section, element_id)

                coefficient = get_network_element_coefficient(section, element)

                real_size = get_network_element_real_size(element, element_type)

                flow = get_flow(section)

                velocity = get_velocity(flow, real_size)

                name = get_network_element_name(element, old_flow < flow)


                pressure_drop = get_network_element_pressure_drop(section, element, density, velocity)

                pressure_total += pressure_drop

                value = [
                    count,
                    name,
                    length,
                    real_size,
                    flow,
                    velocity,
                    coefficient,
                    pressure_drop,
                    pressure_total,
                    output.linkify(element_id)]

                rounded_value = [round_floats(item) for item in value]

                data.append(rounded_value)

    show_network_report(data, selected_system, output)

script_execute()
