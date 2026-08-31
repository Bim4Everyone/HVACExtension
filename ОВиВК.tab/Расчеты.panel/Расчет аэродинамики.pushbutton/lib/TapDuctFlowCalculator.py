#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPI")
clr.AddReference("dosymep.Revit.dll")
import dosymep

clr.ImportExtensions(dosymep.Revit)

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Mechanical import *


class TapDuctFlowCalculator(object):
    """Определяет локальные расходы магистрали для врезок одного воздуховода."""

    def __init__(self, calculator, duct, additional_branches=None):
        """
        Инициализирует расчётчик расходов для одного магистрального воздуховода.

        Args:
            calculator: Основной расчётчик, предоставляющий методы работы с секциями и коннекторами.
            duct: Магистральный воздуховод с рассчитываемыми врезками.
            additional_branches: Дополнительные элементы ответвлений, например терминал крестовины.
        """
        self.calculator = calculator
        self.duct = duct
        self.additional_branches = additional_branches or []
        self._tap_flows = None

    def _get_end_connectors(self):
        """
        Находит два торцевых коннектора по концам осевой линии воздуховода.

        Returns:
            list: Два торцевых коннектора воздуховода.

        Raises:
            ValueError: Если осевая линия или два торцевых коннектора не найдены.
        """
        connectors = self.calculator.get_connectors(self.duct)
        location = self.duct.Location
        if location is None or not hasattr(location, "Curve") or location.Curve is None:
            raise ValueError("У воздуховода ID {} не найдена осевая линия".format(self.duct.Id))

        curve = location.Curve
        endpoints = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
        result = []
        for endpoint in endpoints:
            connector = min(connectors, key=lambda item: item.Origin.DistanceTo(endpoint))
            if connector not in result:
                result.append(connector)

        if len(result) != 2:
            raise ValueError("У воздуховода ID {} не найдены два торцевых коннектора".format(self.duct.Id))
        return result

    def _get_connected_taps(self):
        """
        Собирает уникальные врезки, боковые терминалы и дополнительные ответвления.

        Терминалы на торцах воздуховода не включаются: их расход уже учтён
        расходом соответствующего торцевого коннектора. Терминал, подключённый
        непосредственно к боковой поверхности воздуховода, считается отдельным
        ответвлением и должен уменьшать локальный расход магистрали.

        Returns:
            list: Элементы ответвлений, расходы которых изменяют расход магистрали.
        """
        taps = {}
        end_connectors = self._get_end_connectors()
        for connector in self.duct.ConnectorManager.Connectors:
            for reference in connector.AllRefs:
                owner = reference.Owner
                if owner.Id == self.duct.Id or owner.Category is None:
                    continue

                is_tap = (owner.Category.IsId(BuiltInCategory.OST_DuctFitting)
                          and owner.MEPModel.PartType == PartType.TapAdjustable)
                is_end_connector = any(
                    connector.Origin.IsAlmostEqualTo(item.Origin)
                    for item in end_connectors)
                is_side_terminal = (owner.Category.IsId(BuiltInCategory.OST_DuctTerminal)
                                    and not is_end_connector)
                if not is_tap and not is_side_terminal:
                    continue

                taps[owner.Id.GetIdValue()] = owner
        for element in self.additional_branches:
            taps[element.Id.GetIdValue()] = element
        return list(taps.values())

    def _get_tap_connection_point(self, tap):
        """
        Возвращает координату подключения ответвления к магистрали.

        Args:
            tap: Врезка или дополнительный элемент ответвления.

        Returns:
            XYZ: Координата коннектора со стороны магистрального воздуховода.

        Raises:
            ValueError: Если соединение с магистральным воздуховодом не найдено.
        """
        for connector in self.calculator.get_connectors(tap):
            for reference in connector.AllRefs:
                if reference.Owner.Id == self.duct.Id:
                    return connector.Origin
        raise ValueError("Врезка ID {} не подключена к воздуховоду ID {}".format(tap.Id, self.duct.Id))

    def _get_branch_flow(self, tap):
        """
        Определяет расход конкретного ответвления существующим секционным алгоритмом.

        Для терминала используется его собственный расход. Для врезки сначала
        проверяются секции подключённого воздуховода ответвления, затем секции
        самой врезки используются как резервный источник.

        Args:
            tap: Врезка или терминал ответвления.

        Returns:
            float: Расход ответвления в кубических метрах в час.
        """
        if tap.Category.IsId(BuiltInCategory.OST_DuctTerminal):
            flows = self.calculator.get_element_sections_flows(tap)
            return max(flows)

        input_connector, output_connector = self.calculator.find_input_output_connector(tap)
        if self.calculator.system.SystemType == DuctSystemType.SupplyAir:
            branch_element = output_connector.connected_element
        else:
            branch_element = input_connector.connected_element

        flows = self.calculator.get_element_sections_flows(branch_element) if branch_element else []
        if not flows:
            flows = self.calculator.get_element_sections_flows(tap)

        return max(flows)

    def _calculate(self):
        """
        Рассчитывает локальные расходы до и после каждой позиции врезок.

        Врезки сортируются от торцевого коннектора с наибольшим расходом.
        Элементы в одной позиции группируются и одновременно вычитаются из
        текущего расхода. Для каждой врезки сохраняется пара (Lc, Lp), где Lc
        является расходом до позиции, а Lp - расходом после неё.

        Returns:
            dict: Соответствие числового ElementId врезки паре (Lc, Lp).

        Raises:
            ValueError: Если расход ответвлений превышает расход магистрали
                или не сходится баланс между торцевыми коннекторами.
        """
        end_connectors = self._get_end_connectors()
        start_connector = max(end_connectors, key=lambda item: item.Flow)
        start_flow = UnitUtils.ConvertFromInternalUnits(
            start_connector.Flow,
            UnitTypeId.CubicMetersPerHour)

        # Для притока это направление потока, для вытяжки - обратное направление
        # от общего расхода к меньшему. В обоих случаях Lc находится со стороны
        # торцевого коннектора с наибольшим расходом.
        tap_data = []
        for tap in self._get_connected_taps():
            point = self._get_tap_connection_point(tap)
            tap_data.append({
                "tap": tap,
                "distance": point.DistanceTo(start_connector.Origin),
                "branch_flow": self._get_branch_flow(tap)
            })
        tap_data.sort(key=lambda item: item["distance"])

        station_tolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters)
        stations = []
        for item in tap_data:
            if not stations or abs(stations[-1][0]["distance"] - item["distance"]) > station_tolerance:
                stations.append([item])
            else:
                stations[-1].append(item)

        result = {}
        current_flow = start_flow
        for station in stations:
            station_branch_flow = sum(item["branch_flow"] for item in station)
            if station_branch_flow > current_flow + 0.01:
                raise ValueError(
                    "Расход ответвлений {} превышает расход магистрали {} у воздуховода ID {}".format(
                        station_branch_flow, current_flow, self.duct.Id))
            next_flow = max(0.0, current_flow - station_branch_flow)
            for item in station:
                result[item["tap"].Id.GetIdValue()] = (current_flow, next_flow)
            current_flow = next_flow

        end_connector = min(end_connectors, key=lambda item: item.Flow)
        end_flow = UnitUtils.ConvertFromInternalUnits(
            end_connector.Flow,
            UnitTypeId.CubicMetersPerHour)
        balance_tolerance = max(1.0, start_flow * 0.005)
        if abs(current_flow - end_flow) > balance_tolerance:
            raise ValueError(
                "Не сходится баланс расходов воздуховода ID {}: расчёт {}, коннектор {}".format(
                    self.duct.Id, current_flow, end_flow))
        return result

    def get_flows(self, tap):
        """
        Возвращает расходы магистрали непосредственно до и после заданной врезки.

        Результат для всех врезок воздуховода вычисляется один раз и кешируется
        в экземпляре расчётчика.

        Args:
            tap: Врезка, для которой требуются локальные расходы.

        Returns:
            tuple: Пара (Lc, Lp) в кубических метрах в час.

        Raises:
            ValueError: Если врезка не найдена среди ответвлений воздуховода.
        """
        if self._tap_flows is None:
            self._tap_flows = self._calculate()
        tap_id = tap.Id.GetIdValue()
        if tap_id not in self._tap_flows:
            raise ValueError("Врезка ID {} не найдена на воздуховоде ID {}".format(tap.Id, self.duct.Id))
        return self._tap_flows[tap_id]
