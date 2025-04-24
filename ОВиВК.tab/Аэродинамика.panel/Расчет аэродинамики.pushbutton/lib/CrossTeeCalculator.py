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



class CrossTeeCoefficientCalculator(CalculatorClassLib.AerodinamicCoefficientCalculator):
    TEE_SUPPLY_PASS_NAME = 'Тройник на проход нагнетание круглый/прямоугольный'
    TEE_SUPPLY_BRANCH_ROUND_NAME = 'Тройник нагнетание ответвление круглый'
    TEE_SUPPLY_BRANCH_RECT_NAME = 'Тройник нагнетание ответвление прямоугольный'
    TEE_SUPPLY_SEPARATION_NAME = 'Тройник симметричный разделение потока нагнетание'
    TEE_EXHAUST_PASS_ROUND_NAME = 'Тройник всасывание на проход круглый'
    TEE_EXHAUST_PASS_RECT_NAME = 'Тройник всасывание на проход прямоугольный'
    TEE_EXHAUST_BRANCH_ROUND_NAME = 'Тройник всасывание ответвление круглый'
    TEE_EXHAUST_BRANCH_RECT_NAME = 'Тройник всасывание ответвление прямоугольный'
    TEE_EXHAUST_MERGER_NAME = 'Тройник симметричный слияние'


    CROSS_SUPPLY_PASS_RECT_NAME = 'Крестовина на нагнетании проход прямоугольная'
    CROSS_SUPPLY_BRANCH_RECT_NAME = 'Крестовина на нагнетании ответвление прямоугольная'

    CROSS_SUPPLY_PASS_ROUND_NAME = 'Крестовина на нагнетании проход круглая'
    CROSS_SUPPLY_BRANCH_ROUND_NAME = 'Крестовина на нагнетании ответвление круглая'

    CROSS_EXHAUST_PASS_RECT_NAME = 'Крестовина на всасывании проход прямоугольная'
    CROSS_EXHAUST_PASS_ROUND_NAME = 'Крестовина на всасывании проход круглая'

    CROSS_EXHAUST_BRANCH_RECT_NAME = 'Крестовина на всасывании ответвление прямоугольная'
    CROSS_EXHAUST_BRANCH_ROUND_NAME = 'Крестовина на всасывании ответвление круглая'

    tap_crosses_filtered = []


    def __calculate_coefficient(self, tee_type_name, Lo, Lp, Lc, fp, fo, fc):
        """
        Рассчитывает коэффициент тройника.

        Args:
            tee_type_name (str): Название типа тройника.
            Lo (float): Расход в ответвлении.
            Lp (float): Расход в проходном потоке.
            Lc (float): Расход в основном потоке.
            fp (float): Площадь проходного потока.
            fo (float): Площадь ответвления.
            fc (float): Площадь основного потока.

        Returns:
            float: Коэффициент тройника.
        """
        fp_normed = fp / fc  # Нормированная площадь прохода
        fo_normed = fo / fc  # Нормированная площадь ответвления
        Lo_normed = Lo / Lc  # Нормированный расход в ответвлении
        Lp_normed = Lp / Lc  # Нормированный расход в проходе

        vo_normed = Lo_normed / fo_normed
        vp_normed = Lp_normed / fp_normed
        fn_sqrt = math.sqrt(fp_normed)

        if tee_type_name in [self.TEE_SUPPLY_PASS_NAME,
                             self.CROSS_SUPPLY_PASS_RECT_NAME,
                             self.CROSS_SUPPLY_PASS_ROUND_NAME]:
            return (0.45 * (fp_normed / (1 - Lo_normed)) ** 2 + (0.6 - 1.7 * fp_normed) * (fp_normed / (1 - Lo_normed))
                    - (0.25 - 0.9 * fp_normed ** 2) + 0.19 * ((1 - Lo_normed) / fp_normed))

        if tee_type_name in [self.TEE_SUPPLY_BRANCH_ROUND_NAME,
                             self.CROSS_SUPPLY_BRANCH_ROUND_NAME]:
            return ((fo_normed / Lo_normed) ** 2
                    - 0.58 * (fo_normed / Lo_normed) + 0.54
                    + 0.025 * (Lo_normed / fo_normed))

        if tee_type_name in [self.TEE_SUPPLY_BRANCH_RECT_NAME,
                             self.CROSS_SUPPLY_BRANCH_RECT_NAME]:
            return ((fo_normed / Lo_normed) ** 2
                    - 0.42 * (fo_normed / Lo_normed) + 0.81
                    - 0.06 * (Lo_normed / fo_normed))

        if tee_type_name == self.TEE_SUPPLY_SEPARATION_NAME:
            return 1 + 0.3 * ((Lo_normed / fo_normed) ** 2)

        if tee_type_name in [self.TEE_EXHAUST_PASS_ROUND_NAME,
                             self.CROSS_EXHAUST_PASS_ROUND_NAME]:
            return (((1 - fn_sqrt) + 0.5 * Lo_normed + 0.05) *
                    ((1.7 + (1 / (2 * fo_normed) - 1) * Lo_normed - math.sqrt((fp_normed + fo_normed) * Lo_normed))
                     * ((fp_normed / (1 - Lo_normed)) ** 2)))

        if tee_type_name in [self.TEE_EXHAUST_PASS_RECT_NAME,
                             self.CROSS_EXHAUST_PASS_RECT_NAME]:
            return (((1 - fn_sqrt) + 0.5 * Lo_normed + 0.05) *
                    (1.5 + (1 / (2 * fo_normed) - 1) * Lo_normed - math.sqrt((fp_normed + fo_normed) * Lo_normed))
                    * ((fp_normed / (1 - Lo_normed)) ** 2))

        if tee_type_name in [self.TEE_EXHAUST_BRANCH_ROUND_NAME,
                             self.CROSS_EXHAUST_BRANCH_ROUND_NAME]:
            return ((-0.7 - 6.05 * (1 - fp_normed) ** 3) * (fo_normed / Lo_normed) ** 2
                    + (1.32 + 3.23 * (1 - fp_normed) ** 2) * (fo_normed / Lo_normed)
                    + (0.5 + 0.42 * fp_normed) - 0.167 * (Lo_normed / fo_normed))

        if tee_type_name in [self.TEE_EXHAUST_BRANCH_RECT_NAME,
                             self.CROSS_EXHAUST_BRANCH_RECT_NAME]:
            return (
                    (fo_normed / Lo_normed) ** 2) * (4.1 * ((fp_normed / fo_normed) ** 1.25) *
                                                     (Lo_normed ** 1.5) *
                                                     ((fp_normed + fo_normed) ** (
                                                             (0.3 / Lo_normed) * math.sqrt(fo_normed / fp_normed) - 2))
                                                     - 0.5 * (fp_normed / fo_normed)
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

    def is_tap_cross(self, element):
        def angle_between_vectors(v1, v2):
            dot = v1.DotProduct(v2)
            len1 = v1.GetLength()
            len2 = v2.GetLength()
            if len1 == 0 or len2 == 0:
                return None
            cos_theta = dot / (len1 * len2)
            # Ограничим cos в [-1, 1] во избежание math domain error
            cos_theta = max(min(cos_theta, 1), -1)
            return math.acos(cos_theta)  # в радианах

        def find_right_angle_pair(tap_xyz, con_xyz, duct_connectors, tolerance_deg=10):
            if len(duct_connectors) != 2:
                return None  # Нужно ровно 2 точки

            base_vec = duct_connectors[1].Origin - duct_connectors[0].Origin

            vec = con_xyz - tap_xyz
            angle_rad = angle_between_vectors(vec, base_vec)
            if angle_rad is None:
                return None
            angle_deg = math.degrees(angle_rad)
            if abs(angle_deg - 90) <= tolerance_deg:
                return con_xyz

            return None

        input_connector, output_connector = self.find_input_output_connector(element)
        input_element = input_connector.connected_element
        output_element = output_connector.connected_element

        if self.system.SystemType == DuctSystemType.SupplyAir:
            tap_to_duct_connector = input_connector
            duct_element = input_element
        else:
            tap_to_duct_connector = output_connector
            duct_element = output_element

        duct_connectors = self.get_connectors(duct_element)
        tap_xyz = tap_to_duct_connector.connector_element.Origin

        for connector in duct_element.ConnectorManager.Connectors:
            con_xyz = connector.Origin
            connector_set = connector.AllRefs
            skip_connector = False  # флаг

            owner = None
            for ref_connector in connector_set:
                if ref_connector.Owner.Id == element.Id:
                    skip_connector = True
                    break  # прерываем внутренний цикл

                owner = ref_connector.Owner

            if skip_connector:
                continue  # переходим к следующему connector
            result = find_right_angle_pair(tap_xyz, con_xyz, duct_connectors)
            if result:
                return owner, duct_element

    def get_tee_coefficient(self, element):
        """
        Вычисляет коэффициент тройника для элемента.

        Args:
            element (Element): Элемент.

        Returns:
            float: Коэффициент тройника.
        """

        def get_tee_orientation(element):
            """
            Определяет ориентацию тройника.

            Args:
                element (Element): Элемент.

            Returns:
                TeeVariables: Объект TeeVariables с ориентацией тройника.
            """
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

            result = CalculatorClassLib.TeeVariables(input_output_angle,
                                  input_branch_angle,
                                  input_connector,
                                  output_connector,
                                  branch_connector)

            return result

        def get_tap_tee_type_name(input_connector, output_connector):
            """
            Определяет тип врезки-тройника.

            Args:
                input_connector (ConnectorData): Входной коннектор.
                output_connector (ConnectorData): Выходной коннектор.

            Returns:
                str: Название типа тройника.
            """
            input_element = input_connector.connected_element
            output_element = output_connector.connected_element

            duct_critical = False
            if self.system.SystemType == DuctSystemType.SupplyAir:
                duct_element = output_element
            else:
                duct_element = input_element

            for number in self.critical_path_numbers:
                section = self.system.GetSectionByNumber(number)
                elements_ids = section.GetElementIds()
                if duct_element.Id in elements_ids:
                    duct_critical = True
                    break

            if self.system.SystemType == DuctSystemType.SupplyAir and duct_critical:
                if input_connector.shape == ConnectorProfileType.Rectangular:
                    return self.TEE_SUPPLY_BRANCH_RECT_NAME
                else:
                    return self.TEE_SUPPLY_BRANCH_ROUND_NAME

            elif self.system.SystemType == DuctSystemType.SupplyAir and not duct_critical:
                return self.TEE_SUPPLY_PASS_NAME

            elif self.system.SystemType != DuctSystemType.SupplyAir and duct_critical:
                if input_connector.shape == ConnectorProfileType.Rectangular:
                    return self.TEE_EXHAUST_BRANCH_RECT_NAME
                else:
                    return self.TEE_EXHAUST_BRANCH_ROUND_NAME

            elif self.system.SystemType != DuctSystemType.SupplyAir and not duct_critical:
                if input_connector.shape == ConnectorProfileType.Rectangular:
                    return self.TEE_EXHAUST_PASS_RECT_NAME
                else:
                    return self.TEE_EXHAUST_PASS_ROUND_NAME

        def get_tee_type_name(tee_orientation, shape):
            """
            Определяет тип тройника.

            Args:
                tee_orientation (TeeVariables): Ориентация тройника.
                shape (ConnectorProfileType): Форма коннектора.

            Returns:
                str: Название типа тройника.
            """
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
            """
            Получает переменные для расчета КМС врезки-тройника

            Args:
                input_connector (ConnectorData): Входной коннектор.
                output_connector (ConnectorData): Выходной коннектор.
                tee_type_name (str): Название типа тройника.

            Returns:
                tuple: Кортеж (Lo, Lp, Lc, fo, fc, fp).
            """
            input_element = input_connector.connected_element
            output_element = output_connector.connected_element



            if self.system.SystemType == DuctSystemType.SupplyAir:
                main_flows = self.get_section_flows_by_two_elements(input_element, element)
            else:
                main_flows = self.get_section_flows_by_two_elements(output_element, element)

            if len(main_flows) == 0:
                forms.alert(
                    "Невозможно обработать расходы на секциях. " + str(element.Id),
                    "Ошибка",
                    exitscript=True)

            if self.system.SystemType == DuctSystemType.SupplyAir:
                Lo = output_connector.flow
                tap_to_duct_connector = input_connector
                duct_element = input_element
            else:
                Lo = input_connector.flow
                tap_to_duct_connector = output_connector
                duct_element = output_element



            if len(main_flows) == 2:
                Lc = max(main_flows)
                Lp = min(main_flows)
            if len(main_flows) == 1:
                # Если у нас нашелся только один расход, это значит, что на соседней секции нет нашей врезки
                # и определеить является наш расход проходом или стволом не представляется возможным.
                # Требуется получить все расходы врезок в воздуховод и перебором отыскать недостающий

                Lp = max(main_flows)
                Lc = Lp + Lo

            try:
                diameter = UnitUtils.ConvertFromInternalUnits(duct_element.Diameter, UnitTypeId.Meters)
                area = math.pi * (diameter / 2) ** 2
            except Exception:
                height = UnitUtils.ConvertFromInternalUnits(duct_element.Height, UnitTypeId.Meters)
                width = UnitUtils.ConvertFromInternalUnits(duct_element.Width, UnitTypeId.Meters)

                area = height * width

            fc = area
            fp = area
            fo = input_connector.area

            self.tee_params[element.Id] = CalculatorClassLib.TapTeeCharacteristic(Lo,
                                                                                     Lc,
                                                                                     Lp,
                                                                                     fo,
                                                                                     fc,
                                                                                     fp,
                                                                                     tee_type_name)

            return Lo, Lp, Lc, fo, fc, fp

        def get_tee_variables(tee_orientation, tee_type_name):
            """
            Получает переменные для расчета коэффициента тройника.

            Args:
                tee_orientation (TeeVariables): Ориентация тройника.
                tee_type_name (str): Название типа тройника.

            Returns:
                tuple: Кортеж (Lo, Lp, Lc, fo, fc, fp).
            """
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

            self.tee_params[element.Id] = CalculatorClassLib.TapTeeCharacteristic(Lo,
                                                                                     Lc,
                                                                                     Lp,
                                                                                     fo,
                                                                                     fc,
                                                                                     fp,
                                                                                     tee_type_name)

            return Lo, Lp, Lc, fo, fc, fp

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

        coefficient = self.__calculate_coefficient(tee_type_name, Lo, Lp, Lc, fp, fo, fc)

        return coefficient

    def get_cross_name(self, is_rectangular, duct_critical, branch_1_critical, branch_2_critical, Lo_1, Lo_2, Lp):
        name_map = {
            True: {  # Supply
                "PASS": {
                    True: self.CROSS_SUPPLY_PASS_RECT_NAME,
                    False: self.CROSS_SUPPLY_PASS_ROUND_NAME
                },
                "BRANCH": {
                    True: self.CROSS_SUPPLY_BRANCH_RECT_NAME,
                    False: self.CROSS_SUPPLY_BRANCH_ROUND_NAME
                }
            },
            False: {  # Exhaust
                "PASS": {
                    True: self.CROSS_EXHAUST_PASS_RECT_NAME,
                    False: self.CROSS_EXHAUST_PASS_ROUND_NAME
                },
                "BRANCH": {
                    True: self.CROSS_EXHAUST_BRANCH_RECT_NAME,
                    False: self.CROSS_EXHAUST_BRANCH_ROUND_NAME
                }
            }
        }

        if duct_critical:
            kind = "PASS"
        elif branch_1_critical or branch_2_critical:
            kind = "BRANCH"
        else:
            kind = "BRANCH" if (Lo_1 > Lp or Lo_2 > Lp) else "PASS"

        result_name = name_map[self.system_is_supply][kind][is_rectangular]

        return result_name

    def get_tap_cross_coefficient(self, element_1, element_2, duct):
        def get_tap_cross_variables():
            input_connector_1, output_connector_1 = self.find_input_output_connector(element_1)
            input_connector_2, output_connector_2 = self.find_input_output_connector(element_2)

            if element_1.Id not in self.tap_crosses_filtered and element_2.Id not in self.tap_crosses_filtered:
                self.tap_crosses_filtered.append(element_2.Id)

            input_element_1 = input_connector_1.connected_element
            output_element_1 = output_connector_1.connected_element

            input_element_2 = input_connector_2.connected_element
            output_element_2 = output_connector_2.connected_element


            if self.system.SystemType != DuctSystemType.SupplyAir:
                duct = output_element_1
                branch_duct_1 = input_element_1
                branch_duct_2 = input_element_2
            else:
                duct = input_element_1
                branch_duct_1 = output_element_1
                branch_duct_2 = output_element_2

            duct_connectors = self.get_connectors(duct)
            Lo_1 = max(self.get_element_sections_flows(branch_duct_1))
            Lo_2 = max(self.get_element_sections_flows(branch_duct_2))

            flows_1 = self.get_element_sections_flows(element_1)
            flows_2 = self.get_element_sections_flows(element_2)
            all_flows = flows_1 + flows_2
            excluded = [Lo_1, Lo_2]

            # Оставим только значения, которые не равны Lo_1 или Lo_2
            filtered_flows = [f for f in all_flows if f not in excluded]

            Lc = max(filtered_flows) if filtered_flows else None
            Lp = min(filtered_flows) if filtered_flows else None

            duct_critical = False
            branch_1_critical = False
            branch_2_critical = False
            for number in self.critical_path_numbers:
                section = self.system.GetSectionByNumber(number)
                elements_ids = section.GetElementIds()
                if duct.Id in elements_ids:
                    duct_critical = True
                    break
                if branch_duct_1.Id in elements_ids:
                    branch_1_critical = True
                    break
                if branch_duct_2.Id in elements_ids:
                    branch_2_critical = True
                    break

            is_rectangular = self.is_rectangular(duct_connectors[0])

            result_name = self.get_cross_name(is_rectangular,
                                             duct_critical,
                                             branch_1_critical,
                                             branch_2_critical,
                                             Lo_1,
                                             Lo_2,
                                             Lp)

            fc = self.get_area(duct)
            fp = fc
            fo_1 = self.get_area(input_connector_1)
            fo_2 = self.get_area(input_connector_2)

            if branch_1_critical or (not branch_2_critical and Lo_1 > Lo_2):
                fo_result = fo_1
                Lo_result = Lo_1
            else:
                fo_result = fo_2
                Lo_result = Lo_2


            return result_name, Lc, Lp, Lo_result, fc, fp, fo_result

        connector_data_instances_1 = self.get_connector_data_instances(element_1)
        connector_data_instances_2 = self.get_connector_data_instances(element_2)
        connector_data_instances_duct = self.get_connector_data_instances(duct)

        tap_cross_name, Lc, Lp, Lo, fc, fp, fo = get_tap_cross_variables()

        self.tee_params[element_1.Id] = CalculatorClassLib.TapTeeCharacteristic(Lo, Lc, Lp, fo, fc, fp, tap_cross_name)
        self.remember_element_name(element_1, tap_cross_name, [connector_data_instances_1[0],
                                                             connector_data_instances_2[0],
                                                             connector_data_instances_duct[0],
                                                             connector_data_instances_duct[0]])

        return self.__calculate_coefficient(tap_cross_name, Lo, Lp, Lc, fp, fo, fc)

    def get_cross_coefficient(self, element):
        def get_angle_between_connectors(connector_1, connector_2):

            # Получаем координаты центров соединений
            input_origin = connector_1.connector_element.Origin
            output_origin = connector_2.connector_element.Origin


            # Получаем координату точки вставки тройника
            location = element.Location.Point

            # Создаем векторы направлений от точки вставки тройника
            vec_input_location = input_origin - location
            vec_output_location = output_origin - location


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

            return input_output_angle

        def get_cross_variables():
            body_connector = max(connector_data_instances, key=lambda c: c.flow)

            other_connectors = [c for c in connector_data_instances if c != body_connector]

            pass_connector = next(
                (connector for connector in other_connectors
                 if abs(get_angle_between_connectors(body_connector, connector) - 180) <= 5),
                None
            )
            # Определяем branch_connector как оставшийся коннектор
            excluded_ids = {body_connector.connector_element.Id, pass_connector.connector_element.Id}
            # Отбираем все коннекторы, не входящие в excluded_ids
            branch_connectors = [cd for cd in connector_data_instances if cd.connector_element.Id not in excluded_ids]

            branch_connector_1 = branch_connectors[0] if len(branch_connectors) > 0 else None
            branch_connector_2 = branch_connectors[1] if len(branch_connectors) > 1 else None

            input_connector, output_connector = self.find_input_output_connector(element)

            Lc = body_connector.flow
            fc = body_connector.area
            Lp = pass_connector.flow
            fp = pass_connector.area

            if self.system_is_supply:
                if pass_connector.connector_element.Id == output_connector.connector_element.Id:
                    if self.is_rectangular(input_connector):
                        cross_name = self.CROSS_SUPPLY_PASS_RECT_NAME
                    else:
                        cross_name = self.CROSS_SUPPLY_PASS_ROUND_NAME

                    main_branch = max([branch_connector_1, branch_connector_2], key=lambda c: c.flow)
                    Lo = main_branch.flow
                    fo = main_branch.area

                if branch_connector_1.connector_element.Id == output_connector.connector_element.Id:
                    if self.is_rectangular(input_connector):
                        cross_name = self.CROSS_EXHAUST_BRANCH_RECT_NAME
                    else:
                        cross_name = self.CROSS_EXHAUST_BRANCH_ROUND_NAME

                    Lo = branch_connector_1.flow
                    fo = branch_connector_1.area

                if branch_connector_2.connector_element.Id == output_connector.connector_element.Id:
                    if self.is_rectangular(input_connector):
                        cross_name = self.CROSS_EXHAUST_BRANCH_RECT_NAME
                    else:
                        cross_name = self.CROSS_EXHAUST_BRANCH_ROUND_NAME

                    Lo = branch_connector_2.flow
                    fo = branch_connector_2.area

            else:
                if pass_connector.connector_element.Id == input_connector.connector_element.Id:
                    if self.is_rectangular(input_connector):
                        cross_name = self.CROSS_EXHAUST_PASS_RECT_NAME
                    else:
                        cross_name = self.CROSS_EXHAUST_PASS_ROUND_NAME

                    main_branch = max([branch_connector_1, branch_connector_2], key=lambda c: c.flow)
                    Lo = main_branch.flow
                    fo = main_branch.area

                if branch_connector_1.connector_element.Id == input_connector.connector_element.Id:
                    if self.is_rectangular(input_connector):
                        cross_name = self.CROSS_EXHAUST_BRANCH_RECT_NAME
                    else:
                        cross_name = self.CROSS_EXHAUST_BRANCH_ROUND_NAME

                    Lo = branch_connector_1.flow
                    fo = branch_connector_1.area
                if branch_connector_2.connector_element.Id == input_connector.connector_element.Id:
                    if self.is_rectangular(input_connector):
                        cross_name = self.CROSS_EXHAUST_BRANCH_RECT_NAME
                    else:
                        cross_name = self.CROSS_EXHAUST_BRANCH_ROUND_NAME

                    Lo = branch_connector_2.flow
                    fo = branch_connector_2.area

            return cross_name, Lc, Lp, Lo, fc, fp, fo

        connector_data_instances = self.get_connector_data_instances(element)

        cross_name, Lc, Lp, Lo, fc, fp, fo = get_cross_variables()

        self.tee_params[element.Id] = CalculatorClassLib.TapTeeCharacteristic(Lo, Lc, Lp, fo, fc, fp, cross_name)
        self.remember_element_name(element, cross_name, connector_data_instances)
        return self.__calculate_coefficient(cross_name, Lo, Lp, Lc, fp, fo, fc)