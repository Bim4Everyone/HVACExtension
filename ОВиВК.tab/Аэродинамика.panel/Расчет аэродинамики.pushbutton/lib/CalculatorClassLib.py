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
        elif connector.Shape == ConnectorProfileType.Rectangular:
            self.height = UnitUtils.ConvertFromInternalUnits(connector.Height, UnitTypeId.Millimeters)
            self.width = UnitUtils.ConvertFromInternalUnits(connector.Width, UnitTypeId.Millimeters)
            self.area = self.height / 1000 * self.width / 1000
        else:
            forms.alert(
                "Не предусмотрена обработка овальных коннекторов.",
                "Ошибка",
                exitscript=True)

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
            if ((reference.Owner.Category.IsId(BuiltInCategory.OST_DuctCurves) or
                    reference.Owner.Category.IsId(BuiltInCategory.OST_DuctFitting)) or
                    reference.Owner.Category.IsId(BuiltInCategory.OST_MechanicalEquipment)):
                self.connected_element = reference.Owner

class TeeVariables:
    """Класс для хранения характеристик тройника."""

    def __init__(self,
                 input_output_angle,
                 input_branch_angle,
                 input_connector_data,
                 output_connector_data,
                 branch_connector_data):
        """
        Инициализация объекта TeeCharacteristic.

        Args:
            input_output_angle (float): Угол между входным и выходным коннекторами.
            input_branch_angle (float): Угол между входным и ответвляющимся коннекторами.
            input_connector_data (ConnectorData): Данные входного коннектора.
            output_connector_data (ConnectorData): Данные выходного коннектора.
            branch_connector_data (ConnectorData): Данные ответвляющегося коннектора.
        """
        self.input_output_angle = input_output_angle
        self.input_branch_angle = input_branch_angle
        self.input_connector_data = input_connector_data
        self.output_connector_data = output_connector_data
        self.branch_connector_data = branch_connector_data

class TapTeeCharacteristic:
    """Класс для хранения характеристик тройника с отводом."""

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
            name (str): Название типа тройника.
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
    tee_params = {}

    def __init__(self, doc, uidoc, view):
        """
        Инициализация объекта AerodinamicCoefficientCalculator.

        Args:
            doc (Document): Документ.
            uidoc (UIDocument): Интерфейс пользователя документа.
            view (View): Вид.
        """
        self.doc = doc
        self.uidoc = uidoc
        self.view = view

    def get_critical_path(self, system):
        """
        Получает критический путь системы.

        Args:
            system (System): Система.
        """
        self.system = system

        path_numbers = system.GetCriticalPathSectionNumbers()
        self.critical_path_numbers = list(path_numbers)

        self.system_is_supply = system.SystemType == DuctSystemType.SupplyAir

        if self.system_is_supply:
            self.critical_path_numbers.reverse()

        self.section_indexes = self.get_all_sections_in_system()

    def is_rectangular(self, element):
        if isinstance(element, Connector):
            return element.Shape == ConnectorProfileType.Rectangular
        if isinstance(element, ConnectorData):
            return element.ConnectorElement.Shape == ConnectorProfileType.Rectangular
        else:
            connectors = self.get_connectors(element)
            return connectors[0].Shape == ConnectorProfileType.Rectangular

    def get_area(self, element):
        def get_connector_area(connector):
            area = None
            if connector.Shape == ConnectorProfileType.Oval:
                forms.alert(
                    "Не предусмотрена обработка овальных коннекторов.",
                    "Ошибка",
                    exitscript=True)

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
            connectors = self.get_connectors(element)
            area = get_connector_area(connectors[0])
            return area

    def get_all_sections_in_system(self):
        """
        Возвращает список всех секций, к которым относятся элементы системы MEP.

        Returns:
            list: Список индексов секций.
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
                section = None  # Это делается для
            if section is None:
                continue
            found_section_indexes.add(number)

        return sorted(found_section_indexes)

    def get_connectors(self, element):
        """
        Получает коннекторы элемента.

        Args:
            element (Element): Элемент.

        Returns:
            list: Список коннекторов.
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
            if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
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
        Получает экземпляры данных коннекторов для элемента.

        Args:
            element (Element): Элемент.

        Returns:
            list: Список экземпляров данных коннекторов.
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
            element (Element): Элемент.

        Returns:
            tuple: Кортеж (входной коннектор, выходной коннектор).
        """
        connector_data_instances = self.get_connector_data_instances(element)

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

    def get_elbow_coefficient(self, element):
        """
        Вычисляет КМС отвода.

        Args:
            element (Element): Элемент.

        Returns:
            float: КМС.
        """
        connector_data = self.get_connector_data_instances(element)
        connector = connector_data[0]

        is_tap = element.MEPModel.PartType == PartType.TapAdjustable

        if is_tap:
            input_connector, output_connector = self.find_input_output_connector(element)
            input_element = input_connector.connected_element
            output_element = output_connector.connected_element

            main_element = input_element if self.system.SystemType == DuctSystemType.SupplyAir else output_element

            f = input_connector.area
            try:
                diameter = UnitUtils.ConvertFromInternalUnits(main_element.Diameter, UnitTypeId.Meters)
                F = math.pi * (diameter / 2) ** 2
            except:
                height = UnitUtils.ConvertFromInternalUnits(main_element.Height, UnitTypeId.Meters)
                width = UnitUtils.ConvertFromInternalUnits(main_element.Width, UnitTypeId.Meters)
                F = (height * width)

            if f != F:
                coefficient = ((f / F) ** 2 + 0.7 * (f / F) ** 2) if output_element == main_element else (
                    0.4 + 0.7 * (f / F) ** 2)
                base_name = 'Колено прямоугольное с изменением сечения'
                duct_input = self.find_input_output_connector(main_element)[0]
                self.remember_element_name(element, base_name, [input_connector, duct_input])
                return coefficient

            # Если площади равны — работаем как с обычным отводом
            connector.angle = 90

        element_type = element.GetElementType()
        # В стандартных семействах шаблона этот параметр есть. Для других вычислить почти невозможно, принимаем по ГОСТ
        rounding = element_type.GetParamValueOrDefault('Закругление', 150.0)
        if rounding != 150:
            rounding = UnitUtils.ConvertFromInternalUnits(rounding, UnitTypeId.Millimeters)

        if connector.shape == ConnectorProfileType.Rectangular:
            h, b = connector.height, connector.width
            coefficient = (0.25 * (b / h) ** 0.25) * (1.07 * math.exp(2 / (2 * (rounding + b / 2) / b + 1)) - 1) ** 2
            if connector.angle <= 60:
                coefficient *= 0.708
            base_name = 'Отвод прямоугольный'

        elif connector.shape == ConnectorProfileType.Round:
            coefficient = 0.33 if connector.angle > 85 else 0.18
            base_name = 'Отвод круглый'

        else:
            coefficient = 0
            base_name = 'Неизвестная форма отвода'

        self.remember_element_name(element, base_name, [connector, connector], angle=connector.angle)
        return coefficient

    def get_transition_coefficient(self, element):
        """
        Вычисляет коэффициент для диффузора или конфузора.

        Args:
            element (Element): Элемент.

        Returns:
            float: КМС.
        """
        def get_transition_variables(element):
            """
            Получает переменные для расчета коэффициента перехода.

            Args:
                element (Element): Элемент.

            Returns:
                tuple: Кортеж (входной коннектор, выходной коннектор, длина, угол).
            """
            input_conn, output_conn = self.find_input_output_connector(element)
            input_origin = input_conn.connector_element.Origin
            output_origin = output_conn.connector_element.Origin

            in_width = input_conn.radius * 2 if input_conn.radius else input_conn.width
            out_width = output_conn.radius * 2 if output_conn.radius else output_conn.width

            length = input_origin.DistanceTo(output_origin)
            length = UnitUtils.ConvertFromInternalUnits(length, UnitTypeId.Millimeters)

            angle_rad = math.atan(abs(in_width - out_width) / float(length))
            angle_deg = math.degrees(angle_rad)

            return input_conn, output_conn, length, angle_deg

        input_conn, output_conn, length, angle = get_transition_variables(element)

        base_name = 'Заужение' if input_conn.area > output_conn.area else 'Расширение'
        self.remember_element_name(element, base_name, [input_conn, output_conn])

        is_confuser = input_conn.area > output_conn.area
        is_circular = bool(output_conn.radius if is_confuser else input_conn.radius)

        if is_confuser:
            d = output_conn.radius * 2 if is_circular else (4.0 * output_conn.width * output_conn.height) / (
                2 * (output_conn.width + output_conn.height))
            l_d = length / float(d)

            thresholds = [(1, [(10, 0.41), (20, 0.34), (30, 0.27), (180, 0.24)]),
                          (0.15, [(10, 0.39), (20, 0.29), (30, 0.22), (180, 0.18)]),
                          (float('inf'), [(10, 0.29), (20, 0.20), (30, 0.15), (180, 0.13)])]

            for limit, table in thresholds:
                if l_d <= limit:
                    for angle_limit, coeff in table:
                        if angle <= angle_limit:
                            return coeff

        else:
            F = input_conn.area / float(output_conn.area)
            angle_limits = [(16, 0), (24, 0), (30 if is_circular else 32, 0), (180, 0)]

            values_by_F = {
                True: [  # Circular
                    (0.2, [0.19, 0.32, 0.43, 0.61]),
                    (0.25, [0.17, 0.28, 0.37, 0.49]),
                    (0.4, [0.12, 0.19, 0.25, 0.35]),
                    (float('inf'), [0.07, 0.1, 0.12, 0.17])
                ],
                False: [  # Rectangular
                    (0.2, [0.31, 0.4, 0.59, 0.69]),
                    (0.25, [0.27, 0.35, 0.52, 0.61]),
                    (0.4, [0.18, 0.23, 0.34, 0.4]),
                    (float('inf'), [0.09, 0.11, 0.16, 0.19])
                ]
            }

            for f_limit, coeffs in values_by_F[is_circular]:
                if F <= f_limit:
                    for (a_limit, coeff) in zip([l[0] for l in angle_limits], coeffs):
                        if angle <= a_limit:
                            return coeff

        return 0  # В случае равных сечений

    def get_element_sections_flows(self, element):
        flows = []

        for section_index in self.section_indexes:
            section = self.system.GetSectionByIndex(section_index)
            section_elements = section.GetElementIds()

            if element.Id in section_elements:
                flow = UnitUtils.ConvertFromInternalUnits(section.Flow, UnitTypeId.CubicMetersPerHour)
                flows.append(flow)

        return flows

    def get_section_flows_by_two_elements(self, element_1, element_2):
        """
        Получает расход в секции по двум элементам.

        Args:
            element_1 (Element): Первый элемент.
            element_2 (Element): Второй элемент.

        Returns:
            list: Список расходов.
        """

        flows = []

        for section_index in self.section_indexes:
            section = self.system.GetSectionByIndex(section_index)
            section_elements = section.GetElementIds()

            if element_1.Id in section_elements and element_2.Id in section_elements:
                flow = UnitUtils.ConvertFromInternalUnits(section.Flow, UnitTypeId.CubicMetersPerHour)
                flows.append(flow)

        return flows

