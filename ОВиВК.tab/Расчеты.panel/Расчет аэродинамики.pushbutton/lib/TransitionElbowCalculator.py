#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep
import CalculatorClassLib

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

class TransitionElbowCoefficientCalculator(CalculatorClassLib.AerodinamicCoefficientCalculator):
    def __calculate_elbow_coefficient(self, connector, rounding = 150):
        """
        Расчет КМС отводов для врезок и собственно отводов

        Args:
            connector: Диктующий коннектор отвода
            rounding: Закругление. По умолчанию принято 150, по ГОСТ.

        Returns:
            coefficient: КМС отвода
            base_name: Базовое имя для запоминания

        """

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

        return coefficient, base_name

    def get_transition_coefficient(self, element):
        """
        Вычисляет коэффициент для диффузора или конфузора.

        Args:
            element: Переход
        Returns:
            float: КМС
        """
        def get_transition_variables():
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

        input_conn, output_conn, length, angle = get_transition_variables()

        base_name = 'Заужение' if input_conn.area > output_conn.area else 'Расширение'
        self.remember_element_name(element, base_name, [input_conn, output_conn])

        is_confuser = input_conn.area > output_conn.area
        is_circular = bool(output_conn.radius if is_confuser else input_conn.radius)

        if is_confuser:
            d = output_conn.radius * 2 if is_circular else (4.0 * output_conn.width * output_conn.height) / (
                2 * (output_conn.width + output_conn.height)) # Для прямоугольных сечений берем эквивалентный D
            l_d = length / float(d)

            thresholds = [(0.1, [(10, 0.41), (20, 0.34), (30, 0.27), (180, 0.24)]),
                          (0.15, [(10, 0.39), (20, 0.29), (30, 0.22), (180, 0.18)]),
                          (float('inf'), [(10, 0.29), (20, 0.20), (30, 0.15), (180, 0.13)])]

            for limit, table in thresholds:
                if l_d <= limit:
                    for angle_limit, coeff in table:
                        if angle <= angle_limit:
                            return coeff


        else:
            F = input_conn.area / float(output_conn.area)
            angle_limits = (
                [16, 24, 30, 180] if is_circular else [20, 24, 32, 180]
            )

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
                    for a_limit, coeff in zip(angle_limits, coeffs):
                        if angle <= a_limit:
                            return coeff

        return 0  # В случае равных сечений

    def get_tap_elbow_coefficient(self, element):
        """
        Вычисляет коэффициент для колена исполненного через врезку.

        Args:
            element: Отвод или врезка
        Returns:
            float: КМС
        """
        connector_data = self.get_connector_data_instances(element)
        connector = connector_data[0]

        input_connector, output_connector = self.find_input_output_connector(element)
        input_element = input_connector.connected_element
        output_element = output_connector.connected_element

        main_element = input_element if self.system.SystemType == DuctSystemType.SupplyAir else output_element
        second_element = output_element if main_element == input_element else input_element

        f = self.get_element_area(second_element)
        F = self.get_element_area(main_element)

        if f != F:
            coefficient = ((f / F) ** 2 + 0.7 * (f / F) ** 2) if output_element == main_element else (
                    0.4 + 0.7 * (f / F) ** 2)
            base_name = 'Колено прямоугольное с изменением сечения'
            duct_input = self.find_input_output_connector(main_element)[0]
            duct_output = self.find_input_output_connector(second_element)[0]
            self.remember_element_name(element, base_name, [duct_output, duct_input])
            return coefficient

        # Если площади равны — работаем как с обычным отводом
        connector.angle = 90

        coefficient, base_name = self.__calculate_elbow_coefficient(connector)

        self.remember_element_name(element, base_name,
                                   [connector, connector],
                                   angle=connector.angle)

        return coefficient

    def get_elbow_coefficient(self, element):
        """
        Вычисляет коэффициент для колена.

        Args:
            element: Отвод или врезка
        Returns:
            float: КМС
        """
        connector_data = self.get_connector_data_instances(element)
        connector = connector_data[0]

        element_type = element.GetElementType()
        # В стандартных семействах шаблона этот параметр есть. Для других вычислить почти невозможно, принимаем по ГОСТ
        rounding = element_type.GetParamValueOrDefault('Закругление', 150.0)
        if rounding != 150:
            rounding = UnitUtils.ConvertFromInternalUnits(rounding, UnitTypeId.Millimeters)

        coefficient, base_name = self.__calculate_elbow_coefficient(connector, rounding)

        self.remember_element_name(element, base_name,
                                   [connector, connector],
                                   angle=connector.angle)

        return coefficient

    def is_tap_elbow(self, element):
        """
        Проверяет, является ли врезка отводом. Если среди расходов секций, в которых есть врезка имеется 0 - это отвод.
        Args:
            element: Врезка
        Returns:
            bool: True или False
        """

        flows = self.get_element_sections_flows(element)
        if 0 in flows:
            return True
        return False
