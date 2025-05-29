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

from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict

from unmodeling_class_library import *
from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
material_calculator = MaterialCalculator(doc)
unmodeling_factory = UnmodelingFactory(doc)
element_stocks = ElementStocks(doc)

def get_material_hosts(element_types, calculation_name, builtin_category):
    """ Проверяет для типов элементов можно ли на их базе создать расходники

    Args:
        element_types: Типы элементов
        calculation_name: Имя расчета
        builtin_category: Категория для которой ведется расчет

    Returns:
        list: Лист элементов-основ на базе которых будут считаться расходники
    """
    result_list = []

    for element_type in element_types:
        if element_type.GetSharedParamValueOrDefault(calculation_name) == 1:
            for el_id in element_type.GetDependentElements(None):
                element = doc.GetElement(el_id)
                category = element.Category
                if category and category.IsId(builtin_category) and element.GetTypeId() != ElementId.InvalidElementId:
                    result_list.append(element)

    return result_list

def split_calculation_elements_list(elements):
    """ Разделяем список элементов на подсписки из тех элементов, у которых одинаковая функция, система,
    блок, секция и этаж

    Args:
        elements: Список элементов которые нужно поделить по функции-системе

    Returns:
        Массив из списков с уникальным значение функции-системы
    """

    # Создаем словарь для группировки элементов по ключу
    grouped_elements = defaultdict(list)

    for element in elements:
        shared_function = element.GetSharedParamValueOrDefault(
            SharedParamsConfig.Instance.EconomicFunction.Name, unmodeling_factory.OUT_OF_FUNCTION_VALUE)
        shared_system = element.GetSharedParamValueOrDefault(
            SharedParamsConfig.Instance.VISSystemName.Name, unmodeling_factory.OUT_OF_SYSTEM_VALUE)
        shared_block = element.GetParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksBlock, '')
        shared_section = element.GetParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksSection, '')
        shared_floor = element.GetParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksLevel, '')

        uniq_group_key = (shared_function + "_" +
                               shared_system + "_" +
                               shared_block + "_" +
                               shared_section + "_" +
                               shared_floor)

        # Добавляем элемент в соответствующий список в словаре
        grouped_elements[uniq_group_key].append(element)

    # Преобразуем значения словаря в список списков
    lists = list(grouped_elements.values())

    return lists

def get_material_number_value(element, operation_name):
    """Вычисление количественного значения расходника

    Args:
        element: Элемент, для которого можно вычислить количество расходников
        operation_name: Имя операции текущего вычисления, в зависимости от которого выбирается расчет

    Returns:
        double: Количественное значение для расходного материала
    """
    diameter = 0
    width = 0
    height = 0
    outer_diameter = 0

    length, area = material_calculator.get_curve_len_area_parameters_values(element)

    stock = element_stocks.get_stock(element)

    length = length * stock
    area = area * stock

    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        outer_diameter = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER),
            UnitTypeId.Millimeters)
        diameter = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM),
            UnitTypeId.Millimeters)

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and element.DuctType.Shape == ConnectorProfileType.Round:
        diameter = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM),
            UnitTypeId.Millimeters)

    if element.Category.IsId(
            BuiltInCategory.OST_DuctCurves) and element.DuctType.Shape == ConnectorProfileType.Rectangular:
        width = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM),
            UnitTypeId.Millimeters)

        height = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM),
            UnitTypeId.Millimeters)

    if (operation_name == unmodeling_factory.PIPE_METAL_RULE_NAME
            and element.Category.IsId(BuiltInCategory.OST_PipeCurves)):

        result = CalculationResult(material_calculator.get_pipe_material_mass(length, diameter),
                                   area)
        return result
    if (operation_name == unmodeling_factory.DUCT_METAL_RULE_NAME
            and element.Category.IsId(BuiltInCategory.OST_DuctCurves)):
        result = CalculationResult(
            material_calculator.get_duct_material_mass(element, diameter, width, height, area),
            area)
        return result
    if (operation_name == unmodeling_factory.COLOR_RULE_NAME
            and element.Category.IsId(BuiltInCategory.OST_PipeCurves)):
        result = CalculationResult(material_calculator.get_color_mass(area), area)
        return result
    if (operation_name == unmodeling_factory.GRUNT_RULE_NAME
            and element.Category.IsId(BuiltInCategory.OST_PipeCurves)):
        result = CalculationResult(material_calculator.get_grunt_mass(area), area)
        return result
    if (operation_name in [unmodeling_factory.CLAMPS_RULE_NAME, unmodeling_factory.PIN_RULE_NAME]
            and element.Category.IsId(BuiltInCategory.OST_PipeCurves)):
        result = CalculationResult(material_calculator.get_collars_and_pins_number(element, diameter, length), area)
        return result
    if (operation_name == unmodeling_factory.RAPID_CLAMPS_RULE_NAME
            and element.Category.IsId(BuiltInCategory.OST_PipeCurves)):
        result = CalculationResult(material_calculator.get_rapid_collars_number(diameter, length), area)
        return result

    return CalculationResult(0, 0)

def remove_old_models():
    """ Удаление уже размещенных в модели расходников и материалов перед новой генерацией"""
    unmodeling_factory.remove_models(unmodeling_factory.MATERIAL_DESCRIPTION)
    unmodeling_factory.remove_models(unmodeling_factory.CONSUMABLE_DESCRIPTION)

def process_materials(family_symbol, material_description):
    """ Обработка предопределенного списка материалов

    Args:
        family_symbol: Символ семейства якорного элемента для создания новых экземпляров
        material_description: Описание расходника с которым он будет создан и по которому будет удален
    """

    def process_pipe_clamps(elements,
                            system,
                            function,
                            block,
                            section,
                            floor,
                            rule_set,
                            material_description,
                            family_symbol):
        pipes = []
        pipe_dict = {}

        for element in elements:
            if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
                pipes.append(element)

        for pipe in pipes:
            full_diameter = UnitUtils.ConvertFromInternalUnits(
                pipe.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER),
                UnitTypeId.Millimeters)

            pipe_nominal_diameter = UnitUtils.ConvertFromInternalUnits(
                pipe.GetParamValue(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM),
                UnitTypeId.Millimeters)

            key = (full_diameter, pipe_nominal_diameter)

            if key not in pipe_dict:
                pipe_dict[key] = []

            pipe_dict[key].append(pipe)

        material_location = unmodeling_factory.get_base_location()

        for (pipe_outer_diameter, pipe_nominal_diameter) in pipe_dict:
            new_row = unmodeling_factory.create_material_row_class_instance(
                system,
                function,
                block,
                section,
                floor,
                rule_set,
                material_description)

            if rule_set.name == unmodeling_factory.RAPID_CLAMPS_RULE_NAME:
                name_d = " DN" + "{:.15g}".format(pipe_nominal_diameter)
            else:
                name_d = " D=" + "{:.15g}".format(pipe_outer_diameter)

            new_row.name = new_row.name + name_d

            material_location = unmodeling_factory.update_location(material_location)

            for element in pipe_dict[pipe_outer_diameter, pipe_nominal_diameter]:
                value = get_material_number_value(element, rule_set.name)
                new_row.number += value.number
            unmodeling_factory.create_new_position(new_row, family_symbol, material_description, material_location)

    def process_other_rules(elements,
                            system,
                            function,
                            block,
                            section,
                            floor,
                            rule_set,
                            material_description,
                            material_location,
                            family_symbol):
        new_row = unmodeling_factory.create_material_row_class_instance(system,
                                                                        function,
                                                                        block,
                                                                        section,
                                                                        floor,
                                                                        rule_set,
                                                                        material_description)
        area = 0
        for element in elements:
            value = get_material_number_value(element, rule_set.name)
            new_row.number += value.number
            area += value.area

        if rule_set.name in [unmodeling_factory.GRUNT_RULE_NAME, unmodeling_factory.COLOR_RULE_NAME]:
            round_area = round(area, 2)
            new_row.note = str(round_area) + ' м², площадь за раз'

        unmodeling_factory.create_new_position(new_row, family_symbol, material_description, material_location)

    material_location = unmodeling_factory.get_base_location()
    generation_rules_list = unmodeling_factory.get_ruleset()

    for rule_set in generation_rules_list:
        elem_types = unmodeling_factory.get_elements_types_by_category(rule_set.category)
        calculation_elements = get_material_hosts(elem_types, rule_set.method_name, rule_set.category)

        split_lists = split_calculation_elements_list(calculation_elements)

        for elements in split_lists:
            (system,
             function,
             block,
             section,
             floor) = unmodeling_factory.get_element_charactristic(elements[0])

            if rule_set.name in [unmodeling_factory.RAPID_CLAMPS_RULE_NAME, unmodeling_factory.CLAMPS_RULE_NAME]:
                process_pipe_clamps(elements,
                                    system,
                                    function,
                                    block,
                                    section,
                                    floor,
                                    rule_set,
                                    material_description,
                                    family_symbol)
                material_location = unmodeling_factory.get_base_location()
            else:
                material_location = unmodeling_factory.update_location(material_location)

                process_other_rules(elements,
                                    system,
                                    function,
                                    block,
                                    section,
                                    floor,
                                    rule_set,
                                    material_description,
                                    material_location,
                                    family_symbol)

def process_insulation_consumables(family_symbol, consumable_description):
    """ Обработка расходников изоляции

    Args:
        family_symbol: Символ семейства якорного элемента для создания новых экземпляров
        consumable_description: Описание расходника с которым он будет создан и по которому будет удален
    """
    consumable_location = unmodeling_factory.get_base_location()
    insulation_list = get_insulation_elements_list()
    split_insulation_lists = split_calculation_elements_list(insulation_list)

    consumables_by_insulation_type = {}

    insulation_types = unmodeling_factory.get_pipe_duct_insulation_types()

    # кэшируем данные по расходникам изоляции для ее типов
    for insulation_type in insulation_types:
        if insulation_type not in consumables_by_insulation_type:
            consumables_by_insulation_type[insulation_type.Id] = material_calculator.get_consumables_class_instances(
                insulation_type)

    # Разбили изоляцию по системе-функции и дробим по типам чтоб их сопоставить
    for insulation_elements in split_insulation_lists:
        system, function, block, section, floor = unmodeling_factory.get_element_charactristic(insulation_elements[0])

        insulation_elements_by_type = {}

        # Наполняем словарь по типу изоляции
        for insulation_element in insulation_elements:
            insulation_type = insulation_element.GetElementType()

            if insulation_type.Id not in insulation_elements_by_type:
                insulation_elements_by_type[insulation_type.Id] = []
            insulation_elements_by_type[insulation_type.Id].append(insulation_element)

        # Сравниваем словарь по расходникам изоляции и словарь по типам изоляции в этой функции-системе
        for insulation_type_id, elements in insulation_elements_by_type.items():

            if insulation_type_id in consumables_by_insulation_type:
                # Получение классов расходников для этого айди типа изоляции
                consumables = consumables_by_insulation_type[insulation_type_id]

                # для каждого расходника генерируем строку
                for consumable in consumables:
                    new_consumable_row = unmodeling_factory.create_consumable_row_class_instance(
                        system,
                        function,
                        block,
                        section,
                        floor,
                        consumable,
                        consumable_description)

                    for element in elements:
                        length, area = material_calculator.get_curve_len_area_parameters_values(element)

                        stock = element_stocks.get_stock(element)

                        length = length * stock
                        area = area * stock

                        if (consumable.is_expenditure_by_linear_meter == 0
                                or consumable.is_expenditure_by_linear_meter is None):
                            value = consumable.expenditure * area
                            new_consumable_row.number += value
                        else:
                            value = consumable.expenditure * length

                            new_consumable_row.number += value

                    consumable_location = unmodeling_factory.update_location(consumable_location)

                    unmodeling_factory.create_new_position(new_consumable_row, family_symbol,
                                                           consumable_description, consumable_location)

def get_insulation_elements_list():
    """
    Получаем список элементов изоляции труб и воздуховодов

    Returns:
        list: Лист из элементов изоляции
    """
    insulations = []
    insulations += unmodeling_factory.get_elements_by_category(BuiltInCategory.OST_PipeInsulations)
    insulations += unmodeling_factory.get_elements_by_category(BuiltInCategory.OST_DuctInsulations)
    return insulations



@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    family_symbol = unmodeling_factory.startup_checks()

    # При каждом запуске затираем расходники с соответствующим описанием и генерируем заново
    remove_old_models()

    with revit.Transaction("BIM: Добавление расчетных элементов"):
        family_symbol.Activate()

        process_materials(family_symbol, unmodeling_factory.MATERIAL_DESCRIPTION)
        process_insulation_consumables(family_symbol, unmodeling_factory.CONSUMABLE_DESCRIPTION)


script_execute()
