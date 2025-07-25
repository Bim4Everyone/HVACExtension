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

import sys
import System
import math
from pyrevit import forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, Selection
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from collections import namedtuple
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from Autodesk.Revit.DB.Mechanical import *
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig

class ConnectorData:
    """Класс для хранения данных о коннекторе."""

    radius = None
    height = None
    width = None
    area = None
    angle = None
    connected_element = None
    flow = None

    def __init__(self, connector):
        """
        Инициализация объекта ConnectorData.

        Args:
            connector (Connector): Объект коннектора.
        """
        self.connector_element = connector
        self.shape = connector.Shape
        self.get_connected_element()
        self.flow = UnitUtils.ConvertFromInternalUnits(connector.Flow, UnitTypeId.CubicMetersPerHour)
        self.direction = connector.Direction
        self.angle = self.get_connector_angle()

        if connector.Shape == ConnectorProfileType.Round:
            self.radius = UnitUtils.ConvertFromInternalUnits(connector.Radius, UnitTypeId.Millimeters)
            self.area = math.pi * ((self.radius / 1000) ** 2)
        else:
            self.height = UnitUtils.ConvertFromInternalUnits(connector.Height, UnitTypeId.Millimeters)
            self.width = UnitUtils.ConvertFromInternalUnits(connector.Width, UnitTypeId.Millimeters)
            self.area = self.height / 1000 * self.width / 1000

    def get_connector_angle(self):
        """
        Вычисляет угол коннектора в градусах.

        Returns:
            float: Угол коннектора в градусах.
        """
        radians = self.connector_element.Angle
        angle = radians * (180 / math.pi)
        return angle

    def get_connected_element(self):
        """
        Определяет элемент, к которому подключен коннектор.
        """
        for reference in self.connector_element.AllRefs:
            if reference.Owner.InAnyCategory([BuiltInCategory.OST_DuctCurves,
                                              BuiltInCategory.OST_DuctFitting,
                                              BuiltInCategory.OST_MechanicalEquipment]):
                self.connected_element = reference.Owner

class MulticonElementCharacteristic:
    """Класс для хранения характеристик сложных фитингов."""

    def __init__(self, Lo, Lc, Lp, fo, fc, fp, name):
        """
        Инициализация объекта TapTeeCharacteristic.

        Args:
            Lo (float): Расход в ответвлении.
            Lc (float): Расход в основном потоке.
            Lp (float): Расход в проходном потоке.
            fo (float): Площадь ответвления.
            fc (float): Площадь основного потока.
            fp (float): Площадь проходного потока.
            name (str): Название типа тройника-крестовины.
        """
        self.name = name
        self.Lo = Lo
        self.Lc = Lc
        self.Lp = Lp
        self.fo = fo
        self.fc = fc
        self.fp = fp

class AerodinamicCoefficientCalculator(object):
    """Класс для расчета аэродинамических коэффициентов."""

    # GUIDы нужны для обращения в ExternalServiceRegistry, сверки с сервисом текущего элемента и его замены
    LOSS_GUID_CONST = "46245996-eebb-4536-ac17-9c1cd917d8cf"
    COEFF_GUID_CONST = "5a598293-1504-46cc-a9c0-de55c82848b9"

    doc = None
    uidoc = None
    view = None
    system = None
    system_is_supply = None
    all_sections_in_system = None
    section_indexes = None
    element_names = {}
    cross_tee_params = {}

    def __init__(self, doc, uidoc, view):
        self.doc = doc
        self.uidoc = uidoc
        self.view = view

    def __get_all_sections_in_system(self):
        """
        Возвращает список всех секций, к которым относятся элементы системы MEP. В стандартном API нет возможности
        посмотреть все номера секций, перебираем их до max_possible_sections.
        """
        # Получаем все элементы системы
        elements = self.system.DuctNetwork

        # Множество для хранения уникальных номеров секций
        found_section_indexes = set()

        # Пробуем пройтись по диапазону номеров секций
        max_possible_sections = 500  # можно увеличить при необходимости
        for number in range(0, max_possible_sections):
            try:
                section = self.system.GetSectionByIndex(number)
            except:
                section = None
            # Если нужной секции уже нет в текущей системе, а вариант еще остаются - вылетит ошибка.
            # Сразу в эксепт break включить не получится, ревит будет вылетать без отчета
            if section is None:
                break
            found_section_indexes.add(number)

        return sorted(found_section_indexes)

    def _is_rectangular_connector(self, connector):
        return connector.Shape == ConnectorProfileType.Rectangular

    def _is_rectangular_connector_data(self, connector_data):
        return connector_data.connector_element.Shape == ConnectorProfileType.Rectangular

    def _is_rectangular_element(self, element):
        connectors = self.get_connectors(element)
        return connectors[0].Shape == ConnectorProfileType.Rectangular

    def get_critical_path(self, system):
        """
        Получает критический путь системы.

        Args:
            system: Система воздуховодов
        """
        self.system = system

        path_numbers = system.GetCriticalPathSectionNumbers()
        self.critical_path_numbers = list(path_numbers)

        self.system_is_supply = system.SystemType == DuctSystemType.SupplyAir

        if self.system_is_supply:
            self.critical_path_numbers.reverse()

        self.section_indexes = self.__get_all_sections_in_system()

    def is_rectangular(self, element):
        """
        Возвращает True, если элемент прямоугольный.

        Args:
            element: Элемент системы воздуховодов
        Returns:
            True или False
        """
        if isinstance(element, Connector):
            return self._is_rectangular_connector(element)
        elif isinstance(element, ConnectorData):
            return self._is_rectangular_connector_data(element)
        else:
            return self._is_rectangular_element(element)

    def get_element_area(self, element):
        """
        Поиск площади любого элемента.

        Args:
            element: Элемент площадь которого нам нужна
        Returns:
            float: Площадь в м2
        """
        def get_connector_area(connector):
            area = None

            if connector.Shape == ConnectorProfileType.Round:
                radius = UnitUtils.ConvertFromInternalUnits(connector.Radius, UnitTypeId.Millimeters)
                area = math.pi * ((radius / 1000) ** 2)
            elif connector.Shape == ConnectorProfileType.Rectangular:
                height = UnitUtils.ConvertFromInternalUnits(connector.Height, UnitTypeId.Millimeters)
                width = UnitUtils.ConvertFromInternalUnits(connector.Width, UnitTypeId.Millimeters)
                area = height / 1000 * width / 1000

            return area

        if isinstance(element, ConnectorData):
            return element.area

        if isinstance(element, Connector):
            area = get_connector_area(element)
            return area

        else:
            connector_areas = []
            connectors = self.get_connectors(element)

            # Фильтруем только HVAC-коннекторы
            hvac_connectors = [conn for conn in connectors if conn.Domain == Domain.DomainHvac]

            if not hvac_connectors:
                forms.alert(
                    "Не удалось определить площадь одного из элементов. ID: " + str(element.Id),
                    "Ошибка",
                    exitscript=True
                )

            connector_areas = [get_connector_area(conn) for conn in hvac_connectors]
            min_area = min(connector_areas)
            return min_area

    def get_connectors(self, element):
        """
        Получает коннекторы элемента. Если элемент воздуховод получит только его коннекторы, врезки будут игнорироваться.
        """
        connectors = []

        if isinstance(element, FamilyInstance) and element.MEPModel.ConnectorManager is not None:
            connectors.extend(element.MEPModel.ConnectorManager.Connectors)

        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves,
                                  BuiltInCategory.OST_PipeCurves,
                                  BuiltInCategory.OST_FlexDuctCurves]) and \
                isinstance(element, MEPCurve) and element.ConnectorManager is not None:

            # Если это воздуховод — фильтруем только не Curve-коннекторы. Это завязано на врезки которые тоже падают в список
            # но с нулевым расходом и двунаправленным потоком
            if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
                for conn in element.ConnectorManager.Connectors:
                    if conn.ConnectorType != ConnectorType.Curve:
                        connectors.append(conn)
            else:
                # Для других категорий (трубы и гибкие воздуховоды) — добавляем все
                connectors.extend(element.ConnectorManager.Connectors)

        return connectors

    def remember_element_name(self, element, base_name, connector_data_elements, length=None, angle=None):
        """
        Сохраняет название элемента с учетом его размеров и угла.

        Args:
            element (Element): Элемент.
            base_name (str): Базовое название.
            connector_data_elements (list): Данные коннекторов.
            length (float, optional): Длина.
            angle (float, optional): Угол.
        """
        # Собираем размеры
        size_parts = []
        for c in connector_data_elements:
            if c.radius:
                size_parts.append(str(int(c.radius * 2)))
            else:
                size_parts.append('{0}x{1}'.format(int(c.width), int(c.height)))

        size = '-'.join(size_parts)

        # Добавляем длину, если она указана
        if length:
            base_name += ', ' + str(length)

        # Добавляем угол, если он указан
        if angle is not None:
            if angle <= 30:
                angle_value = 30
            elif angle <= 45:
                angle_value = 45
            elif angle <= 60:
                angle_value = 60
            else:
                angle_value = 90

            base_name += ' {0}°'.format(angle_value)


        self.element_names[element.Id] = base_name + ' ' + size

    def get_connector_data_instances(self, element):
        """
        Получает экземпляры ConnectorData для элемента.

        Args:
            element: Element у которого ищутся коннекторы
        Returns:
            Список коннекторов
        """
        connectors = self.get_connectors(element)
        connector_data_instances = []
        for connector in connectors:
            connector_data_instances.append(ConnectorData(connector))
        return connector_data_instances

    def find_input_output_connector(self, element):
        """
        Находит входной и выходной коннекторы элемента.

        Args:
            element: Элемент через который идет поток воздуха
        Returns:
            input_connector, output_connector - вход и выход
        """

        connector_data_instances = self.get_connector_data_instances(element)

        # Для поиска входа-выхода нужны два коннектора, у решеток он один
        if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
            return connector_data_instances[0], connector_data_instances[0]

        input_connector = None  # Первый на пути следования воздуха коннектор
        output_connector = None  # Второй на пути следования воздуха коннектор

        if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
            if self.system.SystemType == DuctSystemType.SupplyAir:
                input_connector = max(connector_data_instances, key=lambda c: c.flow)
                output_connector = min(connector_data_instances, key=lambda c: c.flow)
            else:
                input_connector = min(connector_data_instances, key=lambda c: c.flow)
                output_connector = max(connector_data_instances, key=lambda c: c.flow)

            return input_connector, output_connector

        # Поиск по критическому пути в системе
        passed_elements = []
        for number in self.critical_path_numbers:
            section = self.system.GetSectionByNumber(number)
            elements_ids = section.GetElementIds()

            for connector_data in connector_data_instances:
                if connector_data.connected_element is None:
                    continue

                if (connector_data.connected_element.Id in elements_ids and
                        connector_data.connected_element.Id not in passed_elements):
                    passed_elements.append(connector_data.connected_element.Id)

                    if input_connector is None and connector_data.direction == FlowDirectionType.In:
                        input_connector = connector_data
                    else:
                        output_connector = connector_data

            if input_connector is not None and output_connector is not None:
                break  # Нет смысла продолжать перебор сегментов, если нужный тройник уже обработан

        # Для элементов которые не найдены на критическом пути все равно нужно проверить КМС. Проверяем ориентировочно,
        # по направлением коннекторов
        flow_connectors = []  # Коннекторы, которые участвуют в поиске max flow
        if input_connector is None or output_connector is None:
            for connector_data in connector_data_instances:
                if self.system.SystemType == DuctSystemType.SupplyAir:
                    if connector_data.direction == FlowDirectionType.In:
                        input_connector = connector_data
                    else:
                        # Добавляем все Out, чтоб потом выбрать с максимальным расходом
                        flow_connectors.append(connector_data)

                if self.system.SystemType == DuctSystemType.ExhaustAir or self.system.SystemType == DuctSystemType.ReturnAir:
                    if connector_data.direction == FlowDirectionType.Out:
                        output_connector = connector_data
                    else:
                        # Добавляем все In, чтоб потом выбрать с максимальным расходом
                        flow_connectors.append(connector_data)

            # Если мы ищем output для SupplyAir, выбираем коннектор с максимальным flow
            if self.system.SystemType == DuctSystemType.SupplyAir:
                output_connector = max(flow_connectors, key=lambda c: c.flow)
            # Если мы ищем input для ExhaustAir/ReturnAir, выбираем коннектор с максимальным flow
            elif self.system.SystemType in [DuctSystemType.ExhaustAir, DuctSystemType.ReturnAir]:
                input_connector = max(flow_connectors, key=lambda c: c.flow)

        if input_connector is None or output_connector is None:
            forms.alert(
                "Не найден вход-выход в элемент. " + str(element.Id),
                "Ошибка",
                exitscript=True)

        return input_connector, output_connector

    def get_element_sections_flows(self, element, return_terminal_flow = True):
        """
        Получает все расходы секций, связанных с элементом.

        Args:
            element: Элемент для которого нужен список расходов
            return_terminal_flow: Если элемент терминал, то если не указано False - будет просто возвращен его расход
        Returns:
            Список расходов элемента
        """

        flows = []

        if element.Category.IsId(BuiltInCategory.OST_DuctTerminal) and return_terminal_flow:
            flow = element.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
            flow = UnitUtils.ConvertFromInternalUnits(flow, UnitTypeId.CubicMetersPerHour)
            return [flow]

        for section_index in self.section_indexes:
            section = self.system.GetSectionByIndex(section_index)
            section_elements = section.GetElementIds()

            if element.Id in section_elements:
                flow = UnitUtils.ConvertFromInternalUnits(section.Flow, UnitTypeId.CubicMetersPerHour)
                flows.append(flow)

        return flows



