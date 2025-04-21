#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пересчет КМС'
__doc__ = "Пересчитывает КМС соединительных деталей воздуховодов"

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
    radius = None
    height = None
    width = None
    area = None
    angle = None
    connected_element = None
    flow = None


    def __init__(self, connector):
        self.connector_element = connector
        self.shape = connector.Shape
        self.get_connected_element()
        self.flow = UnitUtils.ConvertFromInternalUnits(connector.Flow, UnitTypeId.CubicMetersPerHour)
        self.direction = connector.Direction
        self.angle = self.get_connector_angle()

        if connector.Shape == ConnectorProfileType.Round:
            self.radius = UnitUtils.ConvertFromInternalUnits(connector.Radius, UnitTypeId.Millimeters)
            self.area = math.pi * ((self.radius/1000) ** 2)
        elif connector.Shape == ConnectorProfileType.Rectangular:
            self.height = UnitUtils.ConvertFromInternalUnits(connector.Height, UnitTypeId.Millimeters)
            self.width = UnitUtils.ConvertFromInternalUnits(connector.Width, UnitTypeId.Millimeters)
            self.area = self.height/1000 * self.width/1000
        else:
            forms.alert(
                "Не предусмотрена обработка овальных коннекторов.",
                "Ошибка",
                exitscript=True)

    def get_connector_angle(self):
        radians = self.connector_element.Angle
        angle = radians * (180 / math.pi)
        return angle

    def get_connected_element(self):
        for reference in self.connector_element.AllRefs:
            if ((reference.Owner.Category.IsId(BuiltInCategory.OST_DuctCurves) or
                    reference.Owner.Category.IsId(BuiltInCategory.OST_DuctFitting)) or
                    reference.Owner.Category.IsId(BuiltInCategory.OST_MechanicalEquipment)):
                self.connected_element = reference.Owner

class TeeCharacteristic:
    def __init__(self,
                 input_output_angle,
                 input_branch_angle,
                 input_connector_data,
                 output_connector_data,
                 branch_connector_data):
        self.input_output_angle = input_output_angle
        self.input_branch_angle = input_branch_angle
        self.input_connector_data = input_connector_data
        self.output_connector_data = output_connector_data
        self.branch_connector_data = branch_connector_data

class TapTeeCharacteristic:
    def __init__(self, Lo, Lc, Lp, fo, fc, fp, name):
        self.name = name
        self.Lo = Lo
        self.Lc = Lc
        self.Lp = Lp
        self.fo = fo
        self.fc = fc
        self.fp = fp

class ElementCharacteristic:
    def __init__(self, name, Lo = None, Lc = None, Lp = None, fo = None, fc = None, fp = None):
        self.name = name
        self.Lo = Lo
        self.Lc = Lc
        self.Lp = Lp
        self.fo = fo
        self.fc = fc
        self.fp = fp

class AerodinamicCoefficientCalculator:
    LOSS_GUID_CONST = "46245996-eebb-4536-ac17-9c1cd917d8cf"
    # Гуид для удельных потерь
    COEFF_GUID_CONST = "5a598293-1504-46cc-a9c0-de55c82848b9"
    # Это - Гуид "Определенный коэффициент". Вроде бы одинаков всегда

    TEE_SUPPLY_PASS_NAME = 'Тройник на проход нагнетание круглый/прямоугольный'
    TEE_SUPPLY_BRANCH_ROUND_NAME = 'Тройник нагнетание ответвление круглый'
    TEE_SUPPLY_BRANCH_RECT_NAME = 'Тройник нагнетание ответвление прямоугольный'
    TEE_SUPPLY_SEPARATION_NAME = 'Тройник симметричный разделение потока нагнетание'
    TEE_EXHAUST_PASS_ROUND_NAME ='Тройник всасывание на проход круглый'
    TEE_EXHAUST_PASS_RECT_NAME = 'Тройник всасывание на проход прямоугольный'
    TEE_EXHAUST_BRANCH_ROUND_NAME = 'Тройник всасывание ответвление круглый'
    TEE_EXHAUST_BRANCH_RECT_NAME = 'Тройник всасывание ответвление прямоугольный'
    TEE_EXHAUST_MERGER_NAME = 'Тройник симметричный слияние'

    doc = None
    uidoc = None
    view = None
    system = None
    all_sections_in_system = None
    element_names = {}
    tee_params = {}

    def __init__(self, doc, uidoc, view, system):
        self.doc = doc
        self.uidoc = uidoc
        self.view = view
        self.system = system

        path_numbers = system.GetCriticalPathSectionNumbers()
        self.critical_path_numbers = list(path_numbers)

        if system.SystemType == DuctSystemType.SupplyAir:
            self.critical_path_numbers.reverse()

    def get_connectors(self, element):
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
        connectors = self.get_connectors(element)
        connector_data_instances = []
        for connector in connectors:
            connector_data_instances.append(ConnectorData(connector))
        return connector_data_instances

    def find_input_output_connector(self, element):

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
                        #Добавляем все In, чтоб потом выбрать с максимальным расходом
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
        '''
        90 гр=0,25 * (b/h)^0,25 * ( 1,07 * e^(2/(2(R+b/2)/b+1)) -1 )^2
        45 гр=0,708*КМС90гр

        Непонятно откуда формула, но ее результаты сходятся с прил. 3 в ВСН и прил. 25.11 в учебнике Краснова.
        Формулы из ВСН и Посохина похожие, но они явно с опечатками, кривой результат

        '''

        element_type = element.GetElementType()

        rounding = element_type.GetParamValueOrDefault('Закругление', 150.0)
        if rounding != 150:
            rounding = UnitUtils.ConvertFromInternalUnits(rounding, UnitTypeId.Millimeters)
        # В стандартных семействах шаблона этот параметр есть. Для других вычислить почти невозможно, принимаем по ГОСТ

        connector_data = self.get_connector_data_instances(element)

        connector_data_element = connector_data[0]

        # Врезки работающие как отводы не дадут нам нормально свой угол забрать, но он все равно всегда 90
        if element.MEPModel.PartType == PartType.TapAdjustable:
            connector_data_element.angle = 90

        if connector_data_element.shape == ConnectorProfileType.Rectangular:
            h = connector_data_element.height
            b = connector_data_element.width

            coefficient = (0.25 * (b / h) ** 0.25) * (1.07 * math.e ** (2 / (2 * (rounding + b / 2) / b + 1)) - 1) ** 2

            if connector_data_element.angle <= 60:
                coefficient = coefficient * 0.708

        if connector_data_element.shape == ConnectorProfileType.Round:
            if connector_data_element.angle > 85:
                coefficient = 0.33
            else:
                coefficient = 0.18

        if connector_data_element.shape == ConnectorProfileType.Rectangular:
            base_name = 'Отвод прямоугольный'
        else:
            base_name = 'Отвод круглый'

        self.remember_element_name(element, base_name,
                                   [connector_data_element, connector_data_element],
                                   angle=connector_data_element.angle)

        return coefficient

    def get_transition_coefficient(self, element):
        '''
        Краснов Ю.С. Системы вентиляции и кондиционирования Прил. 25.1
        '''

        def get_transition_variables(element):
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

    def get_tee_coefficient(self, element):
        def get_tee_orientation(element):
            connector_data_instances = self.get_connector_data_instances(element)

            input_connector, output_connector = self.find_input_output_connector(element)
            branch_connector = None  # Коннектор-ответвление

            # Определяем branch_connector как оставшийся коннектор
            excluded_ids = {input_connector.connector_element.Id, output_connector.connector_element.Id}

            branch_connector = next(
                (cd for cd in connector_data_instances if cd.connector_element.Id not in excluded_ids),
                None
            )

            # Получаем координаты центров соединений
            input_origin = input_connector.connector_element.Origin
            output_origin = output_connector.connector_element.Origin
            branch_origin = branch_connector.connector_element.Origin

            # Получаем координату точки вставки тройника
            location = element.Location.Point

            # Создаем векторы направлений от точки вставки тройника
            vec_input_location = input_origin - location
            vec_output_location = output_origin - location
            vec_branch_location = branch_origin - location

            # Функция вычисления угла между векторами
            def calculate_angle(vec1, vec2):
                dot_product = vec1.DotProduct(vec2)
                norm1 = vec1.GetLength()
                norm2 = vec2.GetLength()

                cosine = dot_product / (norm1 * norm2)
                # Защита от выхода за границы из-за округления
                cosine = max(-1.0, min(1.0, cosine))

                return math.degrees(math.acos(cosine))

            # Вычисляем углы
            input_output_angle = calculate_angle(vec_input_location, vec_output_location)
            input_branch_angle = calculate_angle(vec_input_location, vec_branch_location)

            result = TeeCharacteristic(input_output_angle,
                                       input_branch_angle,
                                       input_connector,
                                       output_connector,
                                       branch_connector)

            return result

        def get_tap_tee_type_name(input_connector, output_connector):
            input_element = input_connector.connected_element
            output_element = output_connector.connected_element


            main_critical = False
            if self.system.SystemType == DuctSystemType.SupplyAir:
                main_element = output_element
            else:
                main_element = input_element

            for number in self.critical_path_numbers:
                section = self.system.GetSectionByNumber(number)
                elements_ids = section.GetElementIds()
                if main_element.Id in elements_ids:
                    main_critical = True
                    break

            if self.system.SystemType == DuctSystemType.SupplyAir and main_critical:
                if input_connector.shape == ConnectorProfileType.Rectangular:
                    return self.TEE_SUPPLY_BRANCH_RECT_NAME
                else:
                    return self.TEE_SUPPLY_BRANCH_ROUND_NAME

            elif self.system.SystemType == DuctSystemType.SupplyAir and not main_critical:
                return self.TEE_SUPPLY_PASS_NAME

            elif self.system.SystemType != DuctSystemType.SupplyAir and main_critical:
                if input_connector.shape == ConnectorProfileType.Rectangular:
                    return self.TEE_EXHAUST_BRANCH_RECT_NAME
                else:
                    return self.TEE_EXHAUST_BRANCH_ROUND_NAME

            elif self.system.SystemType != DuctSystemType.SupplyAir and not main_critical:
                if input_connector.shape == ConnectorProfileType.Rectangular:
                    return self.TEE_EXHAUST_PASS_RECT_NAME
                else:
                    return self.TEE_EXHAUST_PASS_ROUND_NAME

        def get_tee_type_name(tee_orientation, shape):
            flow_90_degree = tee_orientation.input_output_angle < 100
            branch_90_degree = tee_orientation.input_branch_angle < 100

            if self.system.SystemType == DuctSystemType.SupplyAir:
                if not flow_90_degree and branch_90_degree:
                    return self.TEE_SUPPLY_PASS_NAME

                if flow_90_degree and not branch_90_degree:
                    if shape == ConnectorProfileType.Rectangular:
                        return self.TEE_SUPPLY_BRANCH_RECT_NAME
                    else:
                        return self.TEE_SUPPLY_BRANCH_ROUND_NAME

                if flow_90_degree and branch_90_degree:
                    return self.TEE_SUPPLY_SEPARATION_NAME

            else:
                if not flow_90_degree and branch_90_degree:
                    if shape == ConnectorProfileType.Rectangular:
                        return self.TEE_EXHAUST_PASS_RECT_NAME
                    else:
                        return self.TEE_EXHAUST_PASS_ROUND_NAME

                if flow_90_degree and branch_90_degree:
                    if shape == ConnectorProfileType.Rectangular:
                        return self.TEE_EXHAUST_BRANCH_RECT_NAME
                    else:
                        return self.TEE_EXHAUST_BRANCH_ROUND_NAME

                if flow_90_degree and not branch_90_degree:
                    return self.TEE_EXHAUST_MERGER_NAME

        def get_tap_tee_variables(input_connector, output_connector, tee_type_name):
            input_element = input_connector.connected_element
            output_element = output_connector.connected_element

            if self.system.SystemType == DuctSystemType.SupplyAir:
                main_flows = self.get_flows_by_two_elements(input_element, element)
            else:
                main_flows = self.get_flows_by_two_elements(output_element, element)

            if len(main_flows) == 0:
                forms.alert(
                    "Невозможно обработать расходы на секциях. " + str(element.Id),
                    "Ошибка",
                    exitscript=True)

            if self.system.SystemType == DuctSystemType.SupplyAir:
                main_flows = self.get_flows_by_two_elements(input_element, element)
            else:
                main_flows = self.get_flows_by_two_elements(output_element, element)


            if self.system.SystemType == DuctSystemType.SupplyAir:
                Lo = output_connector.flow
            else:
                Lo = input_connector.flow

            if len(main_flows) == 2:
                Lc = max(main_flows)
                Lp = min(main_flows)
            if len(main_flows) == 1:
                Lp = max(main_flows)
                Lc = Lp + Lo

            if self.system.SystemType == DuctSystemType.SupplyAir:
                main_element = input_element
            else:
                main_element = output_element
            try:
                diameter = UnitUtils.ConvertFromInternalUnits(main_element.Diameter, UnitTypeId.Millimeters)
                area = math.pi * (diameter / 2) ** 2
            except Exception:
                height = UnitUtils.ConvertFromInternalUnits(main_element.Height, UnitTypeId.Millimeters)
                width = UnitUtils.ConvertFromInternalUnits(main_element.Width, UnitTypeId.Millimeters)

                area = height / 1000 * width / 1000

            fc = area
            fp = area
            fo = input_connector.area

            self.tee_params[element.Id] = TapTeeCharacteristic(Lo, Lc, Lp, fo, fc, fp, tee_type_name)

            return Lo, Lp, Lc, fo, fc, fp

        def get_tee_variables(tee_orientation, tee_type_name):
            if (tee_type_name == self.TEE_SUPPLY_PASS_NAME
                    or tee_type_name == self.TEE_SUPPLY_SEPARATION_NAME):
                Lc = tee_orientation.input_connector_data.flow
                Lp = tee_orientation.output_connector_data.flow
                Lo = tee_orientation.branch_connector_data.flow

                fc = tee_orientation.input_connector_data.area
                fp = tee_orientation.output_connector_data.area
                fo = tee_orientation.branch_connector_data.area

            if (tee_type_name == self.TEE_SUPPLY_BRANCH_ROUND_NAME
                    or tee_type_name == self.TEE_SUPPLY_BRANCH_RECT_NAME):
                Lc = tee_orientation.input_connector_data.flow
                Lp = tee_orientation.branch_connector_data.flow
                Lo = tee_orientation.output_connector_data.flow

                fc = tee_orientation.input_connector_data.area
                fp = tee_orientation.branch_connector_data.area
                fo = tee_orientation.output_connector_data.area

            if (tee_type_name == self.TEE_EXHAUST_PASS_RECT_NAME
                    or tee_type_name == self.TEE_EXHAUST_PASS_ROUND_NAME
                    or tee_type_name == self.TEE_EXHAUST_MERGER_NAME):
                Lc = tee_orientation.output_connector_data.flow
                Lp = tee_orientation.input_connector_data.flow
                Lo = tee_orientation.branch_connector_data.flow
                fc = tee_orientation.output_connector_data.area
                fp = tee_orientation.input_connector_data.area
                fo = tee_orientation.branch_connector_data.area


            if (tee_type_name == self.TEE_EXHAUST_BRANCH_RECT_NAME
                    or tee_type_name == self.TEE_EXHAUST_BRANCH_ROUND_NAME):
                Lc = tee_orientation.output_connector_data.flow
                Lp = tee_orientation.branch_connector_data.flow
                Lo = tee_orientation.input_connector_data.flow

                fc = tee_orientation.output_connector_data.area
                fp = tee_orientation.branch_connector_data.area
                fo = tee_orientation.input_connector_data.area

            self.tee_params[element.Id] = TapTeeCharacteristic(Lo, Lc, Lp, fo, fc, fp, tee_type_name)

            return Lo, Lp, Lc, fo, fc, fp

        def calculate_tee_coefficient(tee_type_name, Lo, Lp, Lc, fp, fo, fc):
            '''

            Расчетные формулы:

            Тройник на проход нагнетание круглый/прямоуг.
            ВСН прил.1 формула 3
            Посохин прил.2

            Тройник нагнетание ответвление круглый:
            ВСН прил.1 формула 4

            Тройник нагнетание ответвление прямоугольный:
            ВСН прил.1 формула 9

            Тройник симметричный разделение потока нагнетание
            Идельчик диаграмма 7-29, стр. 379

            Тройник всасывание на проход круглый/прямоугольный
            Посохин прил.2
            ВСН прил.1 формула 1

            Тройник всасывание ответвление круглый
            ВСН прил.1 формула 2

            Тройник всасывание ответвление прямоугольный
            Посохин прил.2

            '''


            fp_normed = fp / fc  # Нормированная площадь прохода
            fo_normed = fo / fc  # Нормированная площадь ответвления
            Lo_normed = Lo / Lc  # Нормированый расход в ответвлении
            Lp_normed = Lp / Lc  # Нормированный расход в проходе

            vo_normed = Lo_normed / fo_normed
            vp_normed = Lp_normed / fp_normed
            fn_sqrt = math.sqrt(fp_normed)


            if tee_type_name == self.TEE_SUPPLY_PASS_NAME:
                return (0.45*(fp_normed/(1-Lo_normed))**2+(0.6-1.7*fp_normed)*(fp_normed/(1-Lo_normed))
                        -(0.25-0.9*fp_normed**2)+0.19*((1-Lo_normed)/fp_normed))

            if tee_type_name == self.TEE_SUPPLY_BRANCH_ROUND_NAME:
                return ((fo_normed / Lo_normed) ** 2
                        - 0.58 * (fo_normed/Lo_normed) + 0.54
                        + 0.025 * (Lo_normed / fo_normed))

            if tee_type_name == self.TEE_SUPPLY_BRANCH_RECT_NAME:
                return ((fo_normed / Lo_normed) ** 2
                        - 0.42 * (fo_normed/Lo_normed) + 0.81
                        - 0.06 * (Lo_normed / fo_normed))

            if tee_type_name == self.TEE_SUPPLY_SEPARATION_NAME:
                return 1 + 0.3 * ((Lo_normed / fo_normed) ** 2)

            if tee_type_name == self.TEE_EXHAUST_PASS_ROUND_NAME:
                return (((1 - fn_sqrt) + 0.5 * Lo_normed + 0.05) *
                        ((1.7 + (1 / (2 * fo_normed) - 1) * Lo_normed - math.sqrt((fp_normed + fo_normed) * Lo_normed))
                        * ((fp_normed / (1 - Lo_normed)) ** 2)))

            if tee_type_name == self.TEE_EXHAUST_PASS_RECT_NAME:
                return (((1 - fn_sqrt) + 0.5 * Lo_normed + 0.05) *
                        (1.5 + (1 / (2 * fo_normed) - 1) * Lo_normed - math.sqrt((fp_normed + fo_normed) * Lo_normed))
                        * ((fp_normed / (1 - Lo_normed)) ** 2))

            if tee_type_name == self.TEE_EXHAUST_BRANCH_ROUND_NAME:
                return ((-0.7 - 6.05 * (1 - fp_normed) ** 3) * (fo_normed / Lo_normed) ** 2
                        + (1.32 + 3.23 * (1 - fp_normed) ** 2) * (fo_normed / Lo_normed)
                        + (0.5 + 0.42 * fp_normed) - 0.167 * (Lo_normed / fo_normed))

            if tee_type_name == self.TEE_EXHAUST_BRANCH_RECT_NAME:
                return (
                        (fo_normed / Lo_normed) ** 2) * (4.1 * ((fp_normed / fo_normed) ** 1.25) *
                                                  (Lo_normed**1.5) *
                                                  ( (fp_normed + fo_normed) **(
                                                          (0.3/ Lo_normed) * math.sqrt(fo_normed/fp_normed) - 2 ))
                                                         - 0.5 * (fp_normed/fo_normed)
                )

            if tee_type_name == self.TEE_EXHAUST_MERGER_NAME:
                if fo_normed <= 0.35:
                    return 1
                else:
                    if Lo_normed <= 0.4:
                        # Формула из Excel для случая Lo_normed <= 0.4
                        result = (
                                0.9 * (1 - Lo_normed) *
                                (1 + (Lo_normed / fo_normed) ** 2 - 2 * (
                                            1 - Lo_normed) ** 2 - 2 * Lo_normed ** 2 * 6.1257422745431E-17 / fo_normed) /
                                (Lo_normed / fo_normed) ** 2
                        )
                    else:
                        # Формула из Excel для случая Lo_normed > 0.4
                        result = (
                                0.55 *
                                (1 + (Lo_normed / fo_normed) ** 2 - 2 * (
                                            1 - Lo_normed) ** 2 - 2 * Lo_normed ** 2 * 6.1257422745431E-17 / fo_normed) /
                                (Lo_normed / fo_normed) ** 2
                        )
                    return result

            return None  # Если тип тройника не найден

        is_tap = element.MEPModel.PartType == PartType.TapAdjustable

        if is_tap:
            input_connector, output_connector = self.find_input_output_connector(element)
            tee_type_name = get_tap_tee_type_name(input_connector, output_connector)

            connected_element = input_connector.connected_element if self.system.SystemType == DuctSystemType.SupplyAir else output_connector.connected_element
            duct_input, duct_output = self.find_input_output_connector(connected_element)

            connector_data_list = [duct_input, duct_output, input_connector]
            get_variables = get_tap_tee_variables
            get_args = (input_connector, output_connector, tee_type_name)

        else:
            tee_orientation = get_tee_orientation(element)
            shape = tee_orientation.input_connector_data.shape
            tee_type_name = get_tee_type_name(tee_orientation, shape)

            connector_data_list = [
                tee_orientation.input_connector_data,
                tee_orientation.output_connector_data,
                tee_orientation.branch_connector_data
            ]
            get_variables = get_tee_variables
            get_args = (tee_orientation, tee_type_name)

        self.remember_element_name(element, tee_type_name, connector_data_list)

        if tee_type_name is None:
            forms.alert(
                "Не получилось обработать тройник. " + str(element.Id),
                "Ошибка",
                exitscript=True
            )

        Lo, Lp, Lc, fo, fc, fp = get_variables(*get_args)


        coefficient = calculate_tee_coefficient(tee_type_name, Lo, Lp, Lc, fp, fo, fc)


        return coefficient

    def get_flows_by_two_elements(self, element_1, element_2):
        section_indexes = self.get_all_sections_in_system()

        flows = []

        for section_index in section_indexes:
            section = self.system.GetSectionByIndex(section_index)
            section_elements = section.GetElementIds()

            if element_1.Id in section_elements and element_2.Id in section_elements:
                flow = UnitUtils.ConvertFromInternalUnits(section.Flow, UnitTypeId.CubicMetersPerHour)
                flows.append(flow)

        return flows

    def get_all_sections_in_system(self):
        """Возвращает список всех секций, к которым относятся элементы системы MEP"""

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
                section = None # Это делается для
            if section is None:
                continue
            found_section_indexes.add(number)

        return sorted(found_section_indexes)

    def is_tap_elbow(self, element):
        def get_zero_flow_section(element, section_indexes):
            for section_index in section_indexes:
                section = self.system.GetSectionByIndex(section_index)

                if section.Flow == 0:
                    section_elements = section.GetElementIds()

                    if element.Id in section_elements:
                        return section_index  # Возвращаем найденный номер секции
            return None

        indexes = self.get_all_sections_in_system()

        elbow_section_zero_flow = get_zero_flow_section(element, indexes)

        if elbow_section_zero_flow is None:
            return False

        return True

    def get_tap_adjustable_coefficient(self, element):

        if self.is_tap_elbow(element):
            coefficient = self.get_elbow_coefficient(element)
        else:
            coefficient = self.get_tee_coefficient(element)


        return coefficient
