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

    HOLE_NAME = 'Боковое отверстие '
    START_TERMINAL_NAME = 'Воздухораспределитель '
    END_TERMINAL_NAME_SUPPLY = 'Воздухозабор '
    END_TERMINAL_NAME_EXHAUST = 'Выброс '

    tap_crosses_filtered = []
    duct_terminals_flows = {}
    duct_terminals_sizes = {}

    def __calculate_coefficient(self, tee_type_name, Lo, Lp, Lc, fp, fo, fc):
        """
        Рассчитывает КМС тройника или крестовины.

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
                result = (
                        1 * (
                        1 + (1 / fo_normed) ** 2 +
                        3 * (1 / fo_normed) ** 2 * (Lo_normed ** 2 - Lo_normed)
                )
                )
            else:
                if Lo_normed <= 0.4:
                    result = (
                            0.9 * (1 - Lo_normed) *
                            (
                                    1 + (1 / fo_normed) ** 2 +
                                    3 * (1 / fo_normed) ** 2 * (Lo_normed ** 2 - Lo_normed)
                            )
                    )
                else:
                    result = (
                            0.55 *
                            (
                                    1 + (1 / fo_normed) ** 2 +
                                    3 * (1 / fo_normed) ** 2 * (Lo_normed ** 2 - Lo_normed)
                            )
                    )
            return result

        return None  # Если тип тройника не найден

    def __get_angle_between_connectors(self, element, connector_1, connector_2):
        """
        Возвращает угол в градусах между линиями опущенными из коннекторов на точку вставки элемента

        Args:
            element: Элемент точка вставки которого будет использована
            connector_1: 1-ый коннектор
            connector_2: 2-ой коннектор
        Returns:
            float: Угол в градусах
        """
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

    def get_tap_partner_if_exists(self, element):
        """
        Проверяется наличие "партнёрской" врезки. Если у воздуховода, к которому подключена текущая врезка, есть ещё
        одна врезка, и через её центр и центр исходной врезки можно провести прямую, образующую угол 90° с
        осью воздуховода, то эта врезка считается партнёрской. В этом случае возвращаются найденная врезка и
        соответствующий воздуховод.
        В противном случае возвращается None.

        Args:
            element: Врезка

        Returns:
            None - если партнера нет;
            Element, Element - врезка-партнер и воздуховод-владелец, если партнер есть
        """

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

    def get_double_tap_tee_coefficient(self, element_1, element_2, duct):
        """
        Определяем КМС тройника состоящего из воздуховода и двух врезок.

        Args:
            element_1: Врезка
            element_2: Врезка
            duct: Воздуховод-хозяин обеих врезок

        Returns:
            float: Коэффициент местного сопротивления

        """

        def get_double_tap_tee_variables():
            """
            Требуется определить расход во всех позициях, размеры воздуховодов и тип тройника.

            Ищем входной-выходной элемент для врезки. Если приток, то магистральный воздуховод на входе, отвод на выходе
            для вытяжки наоборот.
            Ищем расход на ответвлениях через поиск максимального среди всех расходов на воздуховоде ответвлений.
            Собираем все расходы имеющие отношения к врезке. Фильтруем с нее расходы ответвления. Расход ствола - максимальный.
            Расход прохода - минимальный. Скорее всего на проходе здесь всегда будет 0.

            Таким образом определены все расходы и размеры отверстий.

            После этого остается определить тип тройника

            Проверяем лежит ли магистраль на критическом пути. Если да это "на проход"
            Если отвод на критическом пути то это "на отвод"
            Если ни то ни то, проверяем где расход больше, на ответвлении или на проходе. Куда больше туда и принимаем поток.

            Сам расход диктующего ответвления берем или максимальный, если идем на проход
            или с того ответвления куда поворачиваем.

            Returns:
                Имя тройника + все габаритные размеры и расходы
            """

            input_connector_1, output_connector_1 = self.find_input_output_connector(element_1)
            input_connector_2, output_connector_2 = self.find_input_output_connector(element_2)


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

            branch_1_critical = False
            branch_2_critical = False
            for number in self.critical_path_numbers:
                section = self.system.GetSectionByNumber(number)
                elements_ids = section.GetElementIds()
                if branch_duct_1.Id in elements_ids:
                    branch_1_critical = True
                    break
                if branch_duct_2.Id in elements_ids:
                    branch_2_critical = True
                    break

            if self.system_is_supply:
                result_name = self.TEE_SUPPLY_SEPARATION_NAME
            else:
                result_name = self.TEE_EXHAUST_MERGER_NAME

            fc = self.get_element_area(duct)
            fp = fc
            fo_1 = self.get_element_area(input_connector_1)
            fo_2 = self.get_element_area(input_connector_2)

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


        if element_1.Id not in self.tap_crosses_filtered and element_2.Id not in self.tap_crosses_filtered:
            self.tap_crosses_filtered.append(element_2.Id)

        double_tap_tee_name, Lc, Lp, Lo, fc, fp, fo = get_double_tap_tee_variables()

        self.cross_tee_params[element_1.Id] =  CalculatorClassLib.MulticonElementCharacteristic(Lo,
                                                                                                Lc,
                                                                                                Lp,
                                                                                                fo,
                                                                                                fc,
                                                                                                fp,
                                                                                                double_tap_tee_name)

        self.remember_element_name(element_1, double_tap_tee_name, [connector_data_instances_1[0],
                                                             connector_data_instances_2[0],
                                                             connector_data_instances_duct[0]])

        return self.__calculate_coefficient(double_tap_tee_name, Lo, Lp, Lc, fp, fo, fc)

    def get_tap_cross_coefficient(self, element_1, element_2, duct):
        """
        Определяем КМС для крестовины состоящей из воздуховода и двух врезок, или одной врезки и одного терминала.

        Args:
            element_1: Врезка или терминал
            element_2: Врезка или терминал
            duct: Воздуховод-хозяин обеих врезок

        Returns:
            float: Коэффициент местного сопротивления

        """

        def get_tap_cross_variables():
            """
            Требуется определить расход во всех позициях, размеры воздуховодов и тип крестовины.

            Ищем входной-выходной элемент для врезки. Если приток, то магистральный воздуховод на входе, отвод на выходе
            для вытяжки наоборот.
            Ищем расход на ответвлениях через поиск максимального среди всех расходов на воздуховоде ответвлений.
            Собираем все расходы имеющие отношения к врезке. Фильтруем с нее расходы ответвления. Расход ствола - максимальный.
            Расход прохода - минимальный.

            Таким образом определены все расходы и размеры отверстий.

            После этого остается определить тип тройника

            Проверяем лежит ли магистраль на критическом пути. Если да это "на проход"
            Если отвод на критическом пути то это "на отвод"
            Если ни то ни то, проверяем где расход больше, на ответвлении или на проходе. Куда больше туда и принимаем поток.

            Сам расход диктующего ответвления берем или максимальный, если идем на проход
            или с того ответвления куда поворачиваем.

            Returns:
                Имя крестовины + все габаритные размеры и расходы
            """

            input_connector_1, output_connector_1 = self.find_input_output_connector(element_1)
            input_connector_2, output_connector_2 = self.find_input_output_connector(element_2)

            input_element_1 = input_connector_1.connected_element
            output_element_1 = output_connector_1.connected_element
            input_element_2 = input_connector_2.connected_element
            output_element_2 = output_connector_2.connected_element

            is_supply_air = self.system.SystemType == DuctSystemType.SupplyAir

            duct = input_element_1 if is_supply_air else output_element_1

            def get_branch_duct(element, input_element, output_element):
                if element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
                    return element
                return output_element if is_supply_air else input_element

            branch_duct_1 = get_branch_duct(element_1, input_element_1, output_element_1)
            branch_duct_2 = get_branch_duct(element_2, input_element_2, output_element_2)

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

            fc = self.get_element_area(duct)
            fp = fc
            fo_1 = self.get_element_area(input_connector_1)
            fo_2 = self.get_element_area(input_connector_2)

            if branch_1_critical or (not branch_2_critical and Lo_1 > Lo_2):
                fo_result = fo_1
                Lo_result = Lo_1
            else:
                fo_result = fo_2
                Lo_result = Lo_2


            return result_name, Lc, Lp, Lo_result, fc, fp, fo_result

        if element_1.Id not in self.tap_crosses_filtered and element_2.Id not in self.tap_crosses_filtered:
            self.tap_crosses_filtered.append(element_2.Id)

        connector_data_instances_1 = self.get_connector_data_instances(element_1)
        connector_data_instances_2 = self.get_connector_data_instances(element_2)
        connector_data_instances_duct = self.get_connector_data_instances(duct)

        tap_cross_name, Lc, Lp, Lo, fc, fp, fo = get_tap_cross_variables()

        self.cross_tee_params[element_1.Id] = CalculatorClassLib.MulticonElementCharacteristic(Lo,
                                                                                               Lc,
                                                                                               Lp,
                                                                                               fo,
                                                                                               fc,
                                                                                               fp,
                                                                                               tap_cross_name)
        self.remember_element_name(element_1, tap_cross_name, [connector_data_instances_1[0],
                                                             connector_data_instances_2[0],
                                                             connector_data_instances_duct[0],
                                                             connector_data_instances_duct[0]])

        return self.__calculate_coefficient(tap_cross_name, Lo, Lp, Lc, fp, fo, fc)

    def get_cross_coefficient(self, element):
        """
        Определяем КМС для крестовины.

        Args:
            element: Крестовина

        Returns:
            float: Коэффициент местного сопротивления

        """
        def get_cross_variables():
            """
            Требуется определить расход во всех позициях, размеры воздуховодов и тип крестовины.

            Коннектор, отвечающий за ствол - тот у которого максимальный коннектор. Проход - на 180 градусов от него.
            Ответвлениями назначаем два оставшихся коннектора.

            Проверяем откуда выходит воздух, если через проход - главное ответвление это с максимальным расходом.
            Если через отвод - главное ответвление это тот, через что выходим.

            Returns:
                Имя крестовины + все габаритные размеры и расходы
            """

            body_connector = max(connector_data_instances, key=lambda c: c.flow)

            other_connectors = [c for c in connector_data_instances if c != body_connector]

            pass_connector = next(
                (connector for connector in other_connectors
                 if abs(self.__get_angle_between_connectors(element, body_connector, connector) - 180) <= 5),
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

        self.cross_tee_params[element.Id] = CalculatorClassLib.MulticonElementCharacteristic(Lo, Lc, Lp, fo, fc, fp, cross_name)
        self.remember_element_name(element, cross_name, connector_data_instances)
        return self.__calculate_coefficient(cross_name, Lo, Lp, Lc, fp, fo, fc)

    def get_tee_coefficient(self, element):
        """
        Определяем КМС для тройника.

        Args:
            element: Тройник

        Returns:
            float: Коэффициент местного сопротивления

        """
        def get_tee_variables():
            """
            Требуется определить расход во всех позициях, размеры воздуховодов и тип тройника.

            Сразу же проверяем, есть ли среди расходов нулевые. Если один из трех концов идет на 0 - значит там заглушка.
            Тогда теоретически нужно пересчитывать его как отвод, но это оставляем на проектировщика. В такой ситуации
            принимаем стволом тот коннектор у которого площадь больше, проходом тот у которого меньше. Если равны то значения
            не имеет.

            Если нет нулевых расходов, ствол - наибольший расход. Проход - тот коннектор который на 180 градусов от него.
            Когда нет коннектора на 180 - это разветвление или слияние, тут только ответвления. Берем за ответвление
            максимальный расход и оставшийся назначаем проходом.

            Если проход на 180 был найден, берем ответвление как оставшийся коннектор.

            Таким образом все расходы и размеры были найдены.

            Остается найти тип тройника:
            Ищем вход-выход воздуха из тройника и сверяем все три коннектора. Если выход это проход - "на проход",
            если выход это отвод - "на отвод".

            Returns:
                  Имя тройника + все габаритные размеры и расходы
            """

            body_connector = branch_connector = pass_connector = None
            other_connectors = None

            has_zero_flow = any(c.flow == 0 for c in connector_data_instances)

            if has_zero_flow:
                pass_connector = next((c for c in connector_data_instances if c.flow == 0), None)
                non_zero_connectors = sorted(
                    (c for c in connector_data_instances if c.flow > 0),
                    key=lambda c: c.area,
                    reverse=True
                )
                body_connector, branch_connector = non_zero_connectors[:2]
            else:
                body_connector = max(connector_data_instances, key=lambda c: c.flow)
                other_connectors = [c for c in connector_data_instances if c != body_connector]
                pass_connector = next(
                    (
                        c for c in other_connectors
                        if abs(self.__get_angle_between_connectors(element, body_connector, c) - 180) <= 5
                    ),
                    None
                )

            # Предустановка имени тройника
            tee_name = None
            is_supply = self.system_is_supply
            is_exhaust = not is_supply
            pass_is_zero = pass_connector is None or pass_connector.flow == 0

            if is_supply and pass_is_zero:
                tee_name = self.TEE_SUPPLY_SEPARATION_NAME
            elif is_exhaust and pass_is_zero:
                tee_name = self.TEE_EXHAUST_MERGER_NAME

            if pass_connector is None and other_connectors:
                branch_connector = max(connector_data_instances, key=lambda c: c.flow)
                pass_connector = next(c for c in other_connectors if c != branch_connector)
            elif branch_connector is None and other_connectors:
                branch_connector = next(c for c in other_connectors if c != pass_connector)

            input_connector, output_connector = self.find_input_output_connector(element)

            # Расчет значений
            Lc, fc = body_connector.flow, body_connector.area
            Lp, fp = pass_connector.flow, pass_connector.area
            Lo, fo = branch_connector.flow, branch_connector.area

            if tee_name is None:
                output_id = output_connector.connector_element.Id
                input_id = input_connector.connector_element.Id
                pass_id = pass_connector.connector_element.Id
                branch_id = branch_connector.connector_element.Id
                is_rect = self.is_rectangular(input_connector)

                if is_supply:
                    if pass_id == output_id:
                        tee_name = self.TEE_SUPPLY_PASS_NAME
                    elif branch_id == output_id:
                        tee_name = self.TEE_EXHAUST_BRANCH_RECT_NAME if is_rect else self.TEE_EXHAUST_BRANCH_ROUND_NAME
                else:
                    if pass_id == input_id:
                        tee_name = self.TEE_EXHAUST_PASS_RECT_NAME if is_rect else self.TEE_EXHAUST_PASS_ROUND_NAME
                    elif branch_id == input_id:
                        tee_name = self.TEE_EXHAUST_BRANCH_RECT_NAME if is_rect else self.TEE_EXHAUST_BRANCH_ROUND_NAME

            return tee_name, Lc, Lp, Lo, fc, fp, fo

        connector_data_instances = self.get_connector_data_instances(element)

        tee_name, Lc, Lp, Lo, fc, fp, fo = get_tee_variables()

        self.cross_tee_params[element.Id] = CalculatorClassLib.MulticonElementCharacteristic(Lo, Lc, Lp, fo, fc, fp, tee_name)

        self.remember_element_name(element, tee_name, connector_data_instances)

        return self.__calculate_coefficient(tee_name, Lo, Lp, Lc, fp, fo, fc)

    def get_tap_tee_coefficient(self, element):
        """
        Определение КМС тройника, состоящего из воздуховода и врезки.

        Args:
            element: Врезка.

        Returns:
            float: Коэффициент местного сопротивления

        """

        def get_tap_tee_variables():
            """
            Требуется определить расход во всех позициях, размеры воздуховодов и тип тройника.

            Ищем входной-выходной элемент для врезки. Если приток, то магистральный воздуховод на входе, отвод на выходе
            для вытяжки наоборот.
            Ищем расход на ответвлении через поиск максимального среди всех расходов на воздуховоде ответвления
            Собираем все расходы имеющие отношения к врезке. Фильтруем с нее расходы ответвления. Расход ствола - максимальный.
            Расход прохода - минимальный.

            Таким образом определены все расходы и размеры отверстий.

            После этого остается определить тип тройника

            Проверяем лежит ли магистраль на критическом пути. Если да это "на проход"
            Если отвод на критическом пути то это "на отвод"
            Если ни то ни то, проверяем где расход больше, на ответвлении или на проходе. Куда больше туда и принимаем поток.

            Returns:
                  Имя тройника + все габаритные размеры и расходы
            """

            input_connector_1, output_connector_1 = self.find_input_output_connector(element)

            input_element = input_connector_1.connected_element
            output_element = output_connector_1.connected_element

            if self.system.SystemType != DuctSystemType.SupplyAir:
                duct = output_element
                branch_duct = input_element

            else:
                duct = input_element
                branch_duct = output_element


            duct_connectors = self.get_connectors(duct)
            Lo = max(self.get_element_sections_flows(branch_duct))


            all_flows = self.get_element_sections_flows(element)
            excluded = [Lo]

            filtered_flows = [f for f in all_flows if f not in excluded]

            Lc = max(filtered_flows) if filtered_flows else None
            Lp = min(filtered_flows) if filtered_flows else None

            duct_critical = False
            branch_critical = False

            for number in self.critical_path_numbers:
                section = self.system.GetSectionByNumber(number)
                elements_ids = section.GetElementIds()
                if duct.Id in elements_ids:
                    duct_critical = True
                    break
                if branch_duct.Id in elements_ids:
                    branch_critical = True
                    break

            is_rectangular = self.is_rectangular(duct_connectors[0])

            name_map = {
                True: {  # Supply
                    "PASS": {
                        True: self.TEE_SUPPLY_PASS_NAME,
                        False: self.TEE_SUPPLY_PASS_NAME
                    },
                    "BRANCH": {
                        True: self.TEE_SUPPLY_BRANCH_RECT_NAME,
                        False: self.TEE_SUPPLY_BRANCH_ROUND_NAME
                    }
                },
                False: {  # Exhaust
                    "PASS": {
                        True: self.TEE_EXHAUST_PASS_RECT_NAME,
                        False: self.TEE_EXHAUST_PASS_ROUND_NAME
                    },
                    "BRANCH": {
                        True: self.TEE_EXHAUST_BRANCH_RECT_NAME,
                        False: self.TEE_EXHAUST_BRANCH_ROUND_NAME
                    }
                }
            }

            if duct_critical:
                kind = "PASS"
            elif branch_critical:
                kind = "BRANCH"
            else:
                kind = "BRANCH" if (Lo > Lp) else "PASS"

            result_name = name_map[self.system_is_supply][kind][is_rectangular]

            fc = self.get_element_area(duct)
            fp = fc
            fo = self.get_element_area(input_connector_1)

            return result_name, Lc, Lp, Lo, fc, fp, fo, duct

        connector_data_instances = self.get_connector_data_instances(element)

        tap_tee_name, Lc, Lp, Lo, fc, fp, fo, duct = get_tap_tee_variables()

        connector_data_instances_duct = self.get_connector_data_instances(duct)

        self.cross_tee_params[element.Id] = CalculatorClassLib.MulticonElementCharacteristic(Lo, Lc, Lp, fo, fc, fp,
                                                                                             tap_tee_name)

        self.remember_element_name(element, tap_tee_name, [connector_data_instances[0],
                                                               connector_data_instances_duct[0],
                                                               connector_data_instances_duct[0]])

        return self.__calculate_coefficient(tap_tee_name, Lo, Lp, Lc, fp, fo, fc)

    def get_side_hole_coefficient(self, terminal):
        """
        Требуется определить решетка это или среднее боковое отверстие.

        Если на первом участке или на последнем участке - решетка или забор-выброс. Тогда КМС берем из параметра при наличии
        и возвращаем без расчета.

        Если не находится на критическом пути то же самое.

        Иначе считаем как среднее отверстие.

        Args:
            terminal: Воздухораспределитель.

        Returns:
            float: Коэффициент местного сопротивления

        """

        element_id = terminal.Id

        terminal_critical = False
        for number in self.critical_path_numbers:
            section = self.system.GetSectionByNumber(number)
            elements_ids = section.GetElementIds()
            if element_id in elements_ids:
                terminal_critical = True

        first_section = self.system.GetSectionByNumber(self.critical_path_numbers[0])
        last_section = self.system.GetSectionByNumber(self.critical_path_numbers[-1])


        first_elements_ids = first_section.GetElementIds()
        last_elements_ids = last_section.GetElementIds()

        local_coefficient = terminal.GetParamValueOrDefault(SharedParamsConfig.Instance.VISLocalResistanceCoef, 0.0)


        connector_element = self.get_connectors(terminal)[0]

        # Если терминал на первом участке или не является
        # частью критического пути - не рассматриваем его и возвращаем КМС
        if element_id in first_elements_ids or not terminal_critical:
            # элемент есть в первом сечении
            self.element_names[terminal.Id] =  self.START_TERMINAL_NAME
            return local_coefficient

        # Если терминал на последнем участке - это выбор или забор, а не боковое отверстие
        if element_id in last_elements_ids:
            # элемент есть в последнем сечении
            if self.system_is_supply:
                name = self.END_TERMINAL_NAME_SUPPLY
            else:
                name = self.END_TERMINAL_NAME_EXHAUST

            self.element_names[terminal.Id] = name
            return local_coefficient

        self.element_names[terminal.Id] = self.HOLE_NAME

        duct = None
        for reference in connector_element.AllRefs:
            if reference.Owner.Category.IsId(BuiltInCategory.OST_DuctCurves):
                duct = reference.Owner

        duct_area = self.get_element_area(duct)
        terminal_area = self.get_element_area(terminal)

        duct_flow = max(self.get_element_sections_flows(terminal, return_terminal_flow=False))

        terminal_flow = terminal.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
        terminal_flow = UnitUtils.ConvertFromInternalUnits(terminal_flow, UnitTypeId.CubicMetersPerHour)



        area_criteria = terminal_area / duct_area
        flow_criteria = terminal_flow / duct_flow

        table = [
            # area_criteria <= 0.1
            [(0.1, 0.1), (0.2, -0.1), (0.3, -0.8), (0.4, -2.6), (float('inf'), -6.6)],
            # area_criteria <= 0.2
            [(0.1, 0.1), (0.2, 0.2), (0.3, -0.01), (0.4, -0.6), (float('inf'), -2.1)],
            # area_criteria <= 0.4
            [(0.1, 0.2), (0.2, 0.3), (0.3, 0.3), (0.4, 0.2), (float('inf'), -0.2)],
            # area_criteria > 0.4
            [(0.1, 0.2), (0.2, 0.3), (0.3, 0.4), (0.4, 0.4), (float('inf'), 0.3)],
        ]

        if area_criteria <= 0.1:
            row = table[0]
        elif area_criteria <= 0.2:
            row = table[1]
        elif area_criteria <= 0.4:
            row = table[2]
        else:
            row = table[3]

        for limit, local_coefficient in row:
            if flow_criteria <= limit:
                self.duct_terminals_flows[terminal.Id] = duct_flow
                self.duct_terminals_sizes[terminal.Id] = duct_area
                return local_coefficient