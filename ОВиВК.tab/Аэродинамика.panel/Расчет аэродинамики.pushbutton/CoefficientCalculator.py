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

        if connector.Shape == ConnectorProfileType.Round:
            self.radius = UnitUtils.ConvertFromInternalUnits(connector.Radius, UnitTypeId.Millimeters)
            self.area = math.pi * self.radius ** 2
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

class TeeOrientationResult:
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

class Aerodinamiccoefficientcalculator:
    LOSS_GUID_CONST = "46245996-eebb-4536-ac17-9c1cd917d8cf" # Гуид для удельных потерь
    COEFF_GUID_CONST = "5a598293-1504-46cc-a9c0-de55c82848b9" # Это - Гуид "Определенный коэффициент". Вроде бы одинаков всегда
    TEE_SUPPLY_PASS_NAME = 'Тройник на проход нагнетание круглый/прямоуг'
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

    def __init__(self, doc, uidoc, view):
        self.doc = doc
        self.uidoc = uidoc
        self.view = view

    def is_supply_air(self, connector):
        return connector.DuctSystemType == DuctSystemType.SupplyAir

    def is_exhaust_air(self, connector):
        return (connector.DuctSystemType == DuctSystemType.ExhaustAir
                or connector.DuctSystemType == DuctSystemType.ReturnAir)

    def is_direction_inside(self, connector):
        return connector.Direction == FlowDirectionType.In

    def is_direction_bidirectonal(self, connector):
        return connector.Direction == FlowDirectionType.Bidirectional

    def is_direction_outside(self, connector):
        return connector.Direction == FlowDirectionType.Out

    def convert_to_milimeters(self, value):
        return  UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.Millimeters)

    def convert_to_square_meters(self, value):
        return  UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.SquareMeters)

    def get_connectors(self, element):
        connectors = []

        if isinstance(element, FamilyInstance) and element.MEPModel.ConnectorManager is not None:
            connectors.extend(element.MEPModel.ConnectorManager.Connectors)

        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves]) and \
                isinstance(element, MEPCurve) and element.ConnectorManager is not None:
            connectors.extend(element.ConnectorManager.Connectors)

        return connectors

    def get_connector_data_instances(self, element):
        connectors = self.get_connectors(element)
        connector_data_instances = []
        for connector in connectors:
            connector_data_instances.append(ConnectorData(connector))
        return connector_data_instances

    def get_con_coords(self, connector):
        a0 = connector.Origin.ToString()
        a0 = a0.replace("(", "")
        a0 = a0.replace(")", "")
        a0 = a0.split(",")
        for x in a0:
            x = float(x)
        return a0

    def get_connector_area(self, connector):
        if connector.Shape == ConnectorProfileType.Round:
            radius = self.convert_to_milimeters(connector.Radius)
            area = math.pi * radius ** 2
        else:
            height = self.convert_to_milimeters(connector.Height)
            width = self.convert_to_milimeters(connector.Width)
            area = height * width
        return area

    def get_duct_coords(self, in_tee_con, connector):
        main_con = []
        connector_set = connector.AllRefs.ForwardIterator()
        while connector_set.MoveNext():
            main_con.append(connector_set.Current)
        duct = main_con[0].Owner
        duct_cons = self.get_connectors(duct)
        for duct_con in duct_cons:
            if self.get_con_coords(duct_con) != in_tee_con:
                in_duct_con = self.get_con_coords(duct_con)
                return in_duct_con

    def get_connector_angle(self, connector):
        radians = connector.Angle
        angle = radians * (180 / math.pi)
        return angle

    def get_coef_elbow(self, element):
        '''
        90 гр=0,25 * (b/h)^0,25 * ( 1,07 * e^(2/(2(R+b/2)/b+1)) -1 )^2
        45 гр=0,708*КМС90гр

        Непонятно откуда формула, но ее результаты сходятся с прил. 3 в ВСН и прил. 25.11 в учебнике Краснова.
        Формулы из ВСН и Посохина похожие, но они явно с опечатками, кривой результат

        '''

        connector_data_element = self.get_connector_data_instances(element)[0]

        if connector_data_element.shape == ConnectorProfileType.Rectangular:
            h = connector_data_element.height
            b = connector_data_element.width

            coefficient = (0.25 * (b / h) ** 0.25) * (1.07 * math.e ** (2 / (2 * (100 + b / 2) / b + 1)) - 1) ** 2

            if connector_data_element.angle <= 60:
                coefficient = coefficient * 0.708

        if connector_data_element.shape == ConnectorProfileType.Round:
            if connector_data_element.angle > 85:
                coefficient = 0.33
            else:
                coefficient = 0.18

        return coefficient

    def get_coef_transition(self, element, system):
        '''
        Здесь используются формулы из
        Краснов Ю.С. Системы вентиляции и кондиционирования Прил. 25.1
        '''

        def get_transition_variables(element, system):
            connector_data_instances = self.get_connector_data_instances(element)

            path_numbers = system.GetCriticalPathSectionNumbers()
            critical_path_numbers = list(path_numbers)

            if system.SystemType == DuctSystemType.SupplyAir:
                critical_path_numbers.reverse()

            input_connector = None  # Первый на пути следования воздуха коннектор
            output_connector = None  # Второй на пути следования воздуха коннектор

            passed_elements = []
            for number in critical_path_numbers:
                section = system.GetSectionByNumber(number)
                elements_ids = section.GetElementIds()

                for connector_data in connector_data_instances:
                    if (connector_data.connected_element.Id in elements_ids and
                            connector_data.connected_element.Id not in passed_elements):
                        passed_elements.append(connector_data.connected_element.Id)

                        if input_connector is None:
                            input_connector = connector_data
                        else:
                            output_connector = connector_data

                if input_connector is not None and output_connector is not None:
                    break  # Нет смысла продолжать перебор сегментов, если нужный тройник уже обработан

            input_origin = input_connector.connector_element.Origin
            output_origin = output_connector.connector_element.Origin

            if input_connector.radius:
                input_width = input_connector.radius * 2
                output_width = output_connector.radius * 2
            else:
                input_width = input_connector.width
                output_width = output_connector.width

            transition_len = input_origin.DistanceTo(output_origin)
            transition_len= UnitUtils.ConvertFromInternalUnits(transition_len, UnitTypeId.Millimeters)

            R_in = input_width / 2
            R_out = output_width / 2

            transition_angle = math.atan(abs(R_in - R_out) / transition_len)
            transition_angle_degrees = math.degrees(transition_angle)
            print(transition_angle_degrees)

            return input_connector, output_connector, transition_len, transition_angle_degrees

        (input_connector,
         output_connector,
         transition_len,
         transition_angle) = get_transition_variables(element, system)

        # Конфузор
        if input_connector.area > output_connector.area:
            if output_connector.radius:
                diameter = output_connector.radius * 2
            else:
                width = output_connector.width
                height = output_connector.height
                diameter = (4 * (width * height)) / (2 * (width + height))  # Эквивалентный диаметр
            len_per_diameter = transition_len / diameter

            if len_per_diameter <= 1:
                if transition_angle <= 10:
                    return 0.41
                elif transition_angle <=20:
                    return 0.34
                elif transition_angle <= 30:
                    return 0.27
                else:
                    return 0.24
            if len_per_diameter <=0.15:
                if transition_angle <= 10:
                    return 0.39
                elif transition_angle <=20:
                    return 0.29
                elif transition_angle <= 30:
                    return 0.22
                else:
                    return 0.18
            else:
                if transition_angle <= 10:
                    return 0.29
                elif transition_angle <=20:
                    return 0.20
                elif transition_angle <= 30:
                    return 0.15
                else:
                    return 0.13
        #F = F0 / F1
        # диффузор
        if input_connector.area < output_connector.area:
            F = input_connector.area / output_connector.area

            if input_connector.radius:
                if F <= 0.2:
                    if transition_angle <=16:
                        return 0.19
                    elif transition_angle <=24:
                        return 0.32
                    elif transition_angle <= 30:
                        return 0.43
                    else:
                        return 0.61
                elif F <= 0.25:
                    if transition_angle <=16:
                        return 0.17
                    elif transition_angle <=24:
                        return 0.28
                    elif transition_angle <= 30:
                        return 0.37
                    else:
                        return 0.49
                elif F <= 0.4:
                    if transition_angle <=16:
                        return 0.12
                    elif transition_angle <=24:
                        return 0.19
                    elif transition_angle <= 30:
                        return 0.25
                    else:
                        return 0.35
                else:
                    if transition_angle <=16:
                        return 0.07
                    elif transition_angle <=24:
                        return 0.1
                    elif transition_angle <= 30:
                        return 0.12
                    else:
                        return 0.17
            else:
                if F <= 0.2:
                    if transition_angle <=20:
                        return 0.31
                    elif transition_angle <=24:
                        return 0.4
                    elif transition_angle <= 32:
                        return 0.59
                    else:
                        return 0.69
                elif F <= 0.25:
                    if transition_angle <=20:
                        return 0.27
                    elif transition_angle <=24:
                        return 0.35
                    elif transition_angle <= 32:
                        return 0.52
                    else:
                        return 0.61
                elif F <= 0.4:
                    if transition_angle <=20:
                        return 0.18
                    elif transition_angle <=24:
                        return 0.23
                    elif transition_angle <= 32:
                        return 0.34
                    else:
                        return 0.4
                else:
                    print(transition_angle)
                    print(F)
                    if transition_angle <=20:
                        return 0.09
                    elif transition_angle <=24:
                        return 0.11
                    elif transition_angle <= 32:
                        return 0.16
                    else:
                        return 0.19

        return 0 # Для случаев когда переход оказался с равными коннекторами

    def get_coef_tee(self, element, system):
        def get_tee_orientation(element, system):
            connector_data_instances = self.get_connector_data_instances(element)

            path_numbers = system.GetCriticalPathSectionNumbers()
            critical_path_numbers = list(path_numbers)

            if system.SystemType == DuctSystemType.SupplyAir:
                critical_path_numbers.reverse()

            input_connector = None  # Первый на пути следования воздуха коннектор
            output_connector = None  # Второй на пути следования воздуха коннектор
            branch_connector = None  # Коннектор-ответвление

            passed_elements = []
            for number in critical_path_numbers:
                section = system.GetSectionByNumber(number)
                elements_ids = section.GetElementIds()

                for connector_data in connector_data_instances:
                    if (connector_data.connected_element.Id in elements_ids and
                            connector_data.connected_element.Id not in passed_elements):
                        passed_elements.append(connector_data.connected_element.Id)

                        if input_connector is None:
                            input_connector = connector_data
                        else:
                            output_connector = connector_data

                if input_connector is not None and output_connector is not None:
                    break  # Нет смысла продолжать перебор сегментов, если нужный тройник уже обработан

            # Определяем branch_connector как оставшийся коннектор
            for connector_data in connector_data_instances:
                if connector_data != input_connector and connector_data != output_connector:
                    branch_connector = connector_data
                    break

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
                return math.degrees(math.acos(dot_product / (norm1 * norm2)))

            # Вычисляем углы
            input_output_angle = calculate_angle(vec_input_location, vec_output_location)
            input_branch_angle = calculate_angle(vec_input_location, vec_branch_location)

            result = TeeOrientationResult(input_output_angle,
                                          input_branch_angle,
                                          input_connector,
                                          output_connector,
                                          branch_connector)

            return result

        def get_tee_type_name(system_type, tee_orientation, shape):
            flow_90_degree = tee_orientation.input_output_angle < 100
            branch_90_degree = tee_orientation.input_branch_angle < 100

            if system_type == DuctSystemType.SupplyAir:
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

        def calculate_tee_coefficient(tee_type_name, Lo, Lp, Lc, fp, fo, fc):
            '''

            Расчетные формулы:

            Тройник на проход нагнетание круглый/прямоуг.
            0,45*(fп/(1-lо))^2+(0,6-1,7*fп)*fп/(1-lо)-(0,25-0,9*fп^2)+0,19*(1-lо)/fп
            ВСН прил.1 формула 3
            Посохин прил.2


            Тройник нагнетание ответвление круглый:
            (fо/lо)^2-0,58*fо/lо+0,54+0,025*lо/fо
            ВСН прил.1 формула 4


            Тройник нагнетание ответвление прямоугольный:
            (fо/lо)^2-0,42*fо/lо+0,81-0,06*lо/fо
            ВСН прил.1 формула 9


            Тройник симметричный разделение потока нагнетание
            1+0,3*(lо/fо)^2
            Идельчик диаграмма 7-29, стр. 379


            Тройник всасывание на проход круглый/прямоугольный
            ((1-fп^0,5)+0,5*lо+0,05)*(1,7+(1/(2*fо)-1)*lо-((fп+fо)*lо)^0,5)*(fп/(1-lо))^2 - круглый
            Посохин прил.2
            ВСН прил.1 формула 1

            (fп/(1-lо))^2*((1-fп)+0,5*lо+0,05)*(1,5+(1/(2*fо)-1)*lо-((fп+fо)*lо)^0,5) - прямоугольный
            Посохин прил.2
            ВСН прил.1 формула 7

            Тройник всасывание ответвление круглый
            (-0,7-6,05*(1-fп)^3)*(fо/lо)^2+(1,32+3,23*(1-fп)^2)*fо/lо+(0,5+0,42*fп)-0,167*lо/fо
            ВСН прил.1 формула 2

            Тройник всасывание ответвление прямоугольный
            (fо/lо)^2*(4,1*(fп/fо)^1,25*lо^1,5*(fп+fо)^(0,3*(fо/fп)^0,5/lо-2)-0,5*fп/fо)
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
                return (0.45 * (vo_normed / (1 - Lo_normed)) ** 2
                        + (0.6 - 1.7 * vo_normed) * (vo_normed / (1 - Lo_normed))
                        - (0.25 - 0.9 * vo_normed ** 2)
                        + 0.19 * (1 - Lo_normed) / vo_normed)

            if tee_type_name == self.TEE_SUPPLY_BRANCH_ROUND_NAME:
                return (fo_normed ** 2 - 0.58 * fo_normed + 0.54 + 0.025 * (Lo_normed / fo_normed))

            if tee_type_name == self.TEE_SUPPLY_BRANCH_RECT_NAME:
                return (fo_normed ** 2 - 0.42 * fo_normed + 0.81 - 0.06 * (Lo_normed / fo_normed) ** 2)

            if tee_type_name == self.TEE_SUPPLY_SEPARATION_NAME:
                return 1 + 0.3 * (Lo_normed / fo_normed) ** 2

            if tee_type_name == self.TEE_EXHAUST_PASS_ROUND_NAME:
                return ((1 - fn_sqrt) + 0.5 * Lo_normed + 0.05 * (
                        1.7 + (1 / (2 * fo_normed) - 1) * Lo_normed - math.sqrt((fp_normed + fo_normed) * Lo_normed))
                        * ((fp_normed / (1 - Lo_normed)) ** 2))

            if tee_type_name == self.TEE_EXHAUST_PASS_RECT_NAME:
                return ((1 - fn_sqrt) + 0.5 * Lo_normed + 0.05 * (
                        1.5 + (1 / (2 * fo_normed) - 1) * Lo_normed - math.sqrt((fp_normed + fo_normed) * Lo_normed))
                        * ((fp_normed / (1 - Lo_normed)) ** 2))

            if tee_type_name == self.TEE_EXHAUST_BRANCH_ROUND_NAME:
                return ((-0.7 - 6.05 * (1 - fp_normed) ** 3) * (fo_normed / Lo_normed) ** 2
                        + (1.32 + 3.23 * (1 - fp_normed) ** 2) * (fo_normed / Lo_normed)
                        + (0.5 + 0.42 * fp_normed) - 0.167 * (Lo_normed / fo_normed))

            if tee_type_name == self.TEE_EXHAUST_BRANCH_RECT_NAME:
                term_a = (fc / Lo_normed) ** 2
                term_b = 4.1 * (fp_normed / fo_normed) ** 1.25 * Lo_normed ** 1.5
                term_c = (fp_normed + fo_normed) ** (0.3 / Lo_normed)
                term_d = (fo_normed / fp_normed) ** 0.5
                term_e = -0.5 * (fp_normed / fo_normed)
                return term_a * (term_b * term_c * term_d ** (-2) + term_e)

            if tee_type_name == self.TEE_EXHAUST_MERGER_NAME:
                if fo <= 0.35:
                    return 1
                else:
                    if Lo / Lc <= 0.4:
                        0.9 * (1 - Lo / Lc) * (1 + (1 / fo) ^ 2 + 3 * (1 / fo) ^ 2 * ((Lo / Lc) ^ 2 - (Lo / Lc)))
                    else:
                        0.55 * (1 + (1 / fo) ^ 2 + 3 * (1 / fo) ^ 2 * ((Lo / Lc) ^ 2 - (Lo / Lc)))

            return None  # Если тип тройника не найден

        def get_tee_variables(tee_orientation, tee_type_name):
            if tee_type_name == self.TEE_SUPPLY_PASS_NAME or tee_type_name == self.TEE_SUPPLY_SEPARATION_NAME:
                Lc = tee_orientation.input_connector_data.flow
                Lp = tee_orientation.output_connector_data.flow
                Lo = tee_orientation.branch_connector_data.flow

                fc = tee_orientation.input_connector_data.area
                fp = tee_orientation.output_connector_data.area
                fo = tee_orientation.branch_connector_data.area

            if tee_type_name == self.TEE_SUPPLY_BRANCH_ROUND_NAME or tee_type_name == self.TEE_SUPPLY_BRANCH_RECT_NAME:
                Lc = tee_orientation.input_connector_data.flow
                Lp = tee_orientation.branch_connector_data.flow
                Lo = tee_orientation.output_connector_data.flow

                fc = tee_orientation.input_connector_data.area
                fp = tee_orientation.branch_connector_data.area
                fo = tee_orientation.output_connector_data.area

            if (tee_type_name == self.TEE_EXHAUST_PASS_RECT_NAME or tee_type_name == self.TEE_EXHAUST_PASS_ROUND_NAME
                    or tee_type_name == self.TEE_EXHAUST_MERGER_NAME):
                Lc = tee_orientation.output_connector_data.flow
                Lp = tee_orientation.input_connector_data.flow
                Lo = tee_orientation.branch_connector_data.flow
                fc = tee_orientation.output_connector_data.area
                fp = tee_orientation.input_connector_data.area
                fo = tee_orientation.branch_connector_data.area


            if tee_type_name == self.TEE_EXHAUST_BRANCH_RECT_NAME or tee_type_name == self.TEE_EXHAUST_BRANCH_RECT_NAME:
                Lc = tee_orientation.output_connector_data.flow
                Lp = tee_orientation.branch_connector_data.flow
                Lo = tee_orientation.input_connector_data.flow
                fc = tee_orientation.output_connector_data.area
                fp = tee_orientation.branch_connector_data.area
                fo = tee_orientation.input_connector_data.area


            return Lo, Lp, Lc, fo, fc, fp

        tee_orientation = get_tee_orientation(element, system)
        system_type = tee_orientation.input_connector_data.connector_element.DuctSystemType
        shape = tee_orientation.input_connector_data.shape

        tee_type_name = get_tee_type_name(system_type, tee_orientation, shape)

        # Lo  Расход воздуха в ответвлении, в формулах значит Lо
        # Lp  Расход воздуха в проходе, в формулах значит Lп
        # Lc  Расход воздуха в стволе, в формулах значит Lс
        #
        # fp  Площадь сечения прохода, в формулах fп
        # fo  площадь сечения ответвления, в формулах fo
        # fc  Площадь сечения ствола, в формулах fc

        Lo, Lp, Lc, fo, fc, fp = get_tee_variables(tee_orientation, tee_type_name)

        coefficient = calculate_tee_coefficient(tee_type_name, Lo, Lp, Lc, fp, fo, fc)

        return coefficient

    def get_coef_tap_adjustable(self, element):
        conSet = self.get_connectors(element)

        try:
            Fo = conSet[0].Height * 0.3048 * conSet[0].Width * 0.3048
            form = "Прямоугольный отвод"
        except:
            Fo = 3.14 * 0.3048 * 0.3048 * conSet[0].Radius ** 2
            form = "Круглый отвод"

        mainCon = []

        connectorSet_0 = conSet[0].AllRefs.ForwardIterator()

        connectorSet_1 = conSet[1].AllRefs.ForwardIterator()

        old_flow = 0
        for con in conSet:
            connectorSet = con.AllRefs.ForwardIterator()
            while connectorSet.MoveNext():
                try:
                    flow = connectorSet.Current.Owner.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
                except Exception:
                    flow = 0
                if flow > old_flow:
                    mainCon = []
                    mainCon.append(connectorSet.Current)
                    old_flow = flow

        duct = mainCon[0].Owner

        ductCons = self.get_connectors(duct)
        Flow = []

        for ductCon in ductCons:
            Flow.append(ductCon.Flow * 101.94)
            # try:
            #     Fc = conSet[0].Height * 0.3048 * ductCon.Width * 0.3048
            #     Fp = Fc
            # except:
            #     Fc = 3.14 * 0.3048 * 0.3048 * ductCon.Radius ** 2
            #     Fp = Fc

            if ductCon.Shape == ConnectorProfileType.Round:
                Fc = 3.14 * 0.3048 * 0.3048 * ductCon.Radius ** 2
                Fp = Fc

            elif ductCon.Shape == ConnectorProfileType.Rectangular:
                Fc = ductCon.Height * 0.3048 * ductCon.Width * 0.3048
                Fp = Fc

        Lc = max(Flow)
        Lo = conSet[0].Flow * 101.94

        f0 = Fo / Fc
        l0 = Lo / Lc
        fp = Fp / Fc

        if str(conSet[0].DuctSystemType) == "ExhaustAir" or str(conSet[0].DuctSystemType) == "ReturnAir":
            if form == "Круглый отвод":
                if Lc > Lo * 2:
                    coefficient = ((1 - fp ** 0.5) + 0.5 * l0 + 0.05) * (
                                1.7 + (1 / (2 * f0) - 1) * l0 - ((fp + f0) * l0) ** 0.5) * (fp / (1 - l0)) ** 2
                else:
                    coefficient = (-0.7 - 6.05 * (1 - fp) ** 3) * (f0 / l0) ** 2 + (1.32 + 3.23 * (1 - fp) ** 2) * f0 / l0 + (
                                0.5 + 0.42 * fp) - 0.167 * l0 / f0
            else:
                if Lc > Lo * 2:
                    coefficient = (fp / (1 - l0)) ** 2 * ((1 - fp) + 0.5 * l0 + 0.05) * (
                                1.5 + (1 / (2 * f0) - 1) * l0 - ((fp + f0) * l0) ** 0.5)
                else:
                    coefficient = (f0 / l0) ** 2 * (4.1 * (fp / f0) ** 1.25 * l0 ** 1.5 * (fp + f0) ** (
                                0.3 * (f0 / fp) ** 0.5 / l0 - 2) - 0.5 * fp / f0)

        if str(conSet[0].DuctSystemType) == "SupplyAir":
            if form == "Круглый отвод":
                if Lc > Lo * 2:
                    coefficient = 0.45 * (fp / (1 - l0)) ** 2 + (0.6 - 1.7 * fp) * fp / (1 - l0) - (
                                0.25 - 0.9 * fp ** 2) + 0.19 * (1 - l0) / fp
                else:
                    coefficient = (f0 / l0) ** 2 - 0.58 * f0 / l0 + 0.54 + 0.025 * l0 / f0
            else:
                if Lc > Lo * 2:
                    coefficient = 0.45 * (fp / (1 - l0)) ** 2 + (0.6 - 1.7 * fp) * fp / (1 - l0) - (
                                0.25 - 0.9 * fp ** 2) + 0.19 * (1 - l0) / fp
                else:
                    coefficient = (f0 / l0) ** 2 - 0.42 * f0 / l0 + 0.81 - 0.06 * l0 / f0

        return coefficient
