#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = "Расчет краски и креплений"
__doc__ = "Генерирует в модели элементы с расчетом количества соответствующих материалов"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import sys
import System
import paraSpec
import checkAnchor
import math

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from Autodesk.Revit.DB import *

from System.Collections.Generic import List
from System import Guid
from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from Redomine import *

from dosymep_libs.bim4everyone import *

#Исходные данные
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

def get_elements_by_category(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col

col_pipes = get_elements_by_category(BuiltInCategory.OST_PipeCurves)
col_curves = get_elements_by_category(BuiltInCategory.OST_DuctCurves)
col_model = get_elements_by_category(BuiltInCategory.OST_GenericModel)
col_systems = get_elements_by_category(BuiltInCategory.OST_DuctSystem)
col_insulation = get_elements_by_category(BuiltInCategory.OST_DuctInsulations)

name_of_model = "_Якорный элемент"
description = "Расчет краски и креплений"

# Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
col_model = \
    [elem for elem in col_model if elem.GetElementType()
    .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == name_of_model]

class GenerationElement:
    def __init__(self, group, name, mark, art, maker, unit, method, collection, isType):
        self.group = group
        self.name = name
        self.mark = mark
        self.maker = maker
        self.unit = unit
        self.collection = collection
        self.method = method
        self.isType = isType
        self.art = art

genList = [
    GenerationElement(
        group = "12. Расчетные элементы",
        name = "Металлические крепления для воздуховодов",
        mark = "",
        art = "",
        unit = "кг.",
        maker = "",
        method = SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name,
        collection=col_curves,
        isType= False),
    GenerationElement(
        group = "12. Расчетные элементы",
        name = "Металлические крепления для трубопроводов",
        mark = "",
        art = "",
        unit = "кг.",
        maker = "",
        method =  SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name,
        collection= col_pipes,
        isType= False),

    GenerationElement(
        group = "12. Расчетные элементы",
        name = "Краска антикоррозионная за два раза",
        mark = "БТ-177",
        art = "",
        unit = "кг.",
        maker = "",
        method =  SharedParamsConfig.Instance.VISIsPaintCalculation.Name,
        collection= col_pipes,
        isType= False),
    GenerationElement(
        group = "12. Расчетные элементы",
        name = "Грунтовка для стальных труб",
        mark = "ГФ-031",
        art = "",
        unit = "кг.",
        maker = "",
        method =  SharedParamsConfig.Instance.VISIsPaintCalculation.Name,
        collection= col_pipes,
        isType= False),
    GenerationElement(

        group = "12. Расчетные элементы",
        name = "Хомут трубный под шпильку М8",
        mark = "",
        art = "",
        unit = "шт.",
        maker = "",
        method =  SharedParamsConfig.Instance.VISIsClampsCalculation.Name,
        collection= col_pipes,
        isType= False),
    GenerationElement(
        group = "12. Расчетные элементы",
        name = "Шпилька М8 1м/1шт",
        mark = "",
        art = "",
        unit = "шт.",
        maker = "",
        method =  SharedParamsConfig.Instance.VISIsClampsCalculation.Name,
        collection= col_pipes,
        isType= False)
]

def roundup(divider, number):
    x = number/divider
    y = int(number/divider)
    if x - y > 0.2:
        return int(number) + 1
    else:
        return int(number)

class CalculationElement:
    pipe_insulation_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
    def __init__(self, element, collection, parameter, name, mark, maker):
        self.local_description = description

        self.corp = element.GetSharedParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksBlock.Name, "")
        self.sec = element.GetSharedParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksSection.Name, "")
        self.floor = element.GetSharedParamValueOrDefault(SharedParamsConfig.Instance.Level.Name, "")

        self.length = UnitUtils.ConvertFromInternalUnits(element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH),
                                                         UnitTypeId.Meters)

        if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
            self.pipe_diametr = UnitUtils.ConvertFromInternalUnits(
                element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM),
                UnitTypeId.Millimeters)

        if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and element.DuctType.Shape == ConnectorProfileType.Round:
            self.duct_diametr = UnitUtils.ConvertFromInternalUnits(
                element.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM),
                UnitTypeId.Millimeters)

        if element.IsExistsParam(SharedParamsConfig.Instance.VISSystemName.Name):
            self.system = element.GetSharedParamValueOrDefault(SharedParamsConfig.Instance.VISSystemName.Name, "")

        # Этот параметр не вызываем с платформы и удаляем из всех шаблонов
        if element.IsExistsParam("ADSK_Имя системы"):
            self.system = element.GetSharedParamValueOrDefault("ADSK_Имя системы", "")

        self.group ="12. Расчетные элементы"
        self.name = name
        self.mark = mark
        self.art = ""
        self.maker = maker
        self.unit = "None"
        self.number = self.get_number(element, self.name)
        self.mass = ""
        self.comment = ""
        self.EF = element.GetSharedParamValueOrDefault(SharedParamsConfig.Instance.VISEconomicFunction.Name, "")
        self.parentId = element.Id.IntegerValue

        for gen in genList:
            if gen.collection == collection and parameter == gen.method:
                self.unit = gen.unit
                isType = gen.isType

        self.key = self.EF + self.corp + self.sec + self.floor + self.system + \
                   self.group + self.name + self.mark + self.art + \
                   self.maker + self.local_description

    def is_pipe_insulated(self, element):
        dependent_elements = element.GetDependentElements(self.pipe_insulation_filter)
        return len(dependent_elements) > 0

    def mid_calculation_fix(self, coeff):
        number = self.length / coeff
        if number < 1:
            number = 1
        return int(number)

    def pins(self, element):
        self.local_description = "{0} {1}, Ду{2}".format(self.local_description, self.name,self.pipe_diametr)
        dict_var_pins = {15: [2, 1.5], 20: [3, 2], 25: [3.5, 2], 32: [4, 2.5], 40: [4.5, 3], 50: [5, 3], 65: [6, 4],
                            80: [6, 4], 100: [6, 4.5], 125: [7, 5]}

        # Мы не считаем крепление труб до 0.5 м
        if self.length < 0.5:
            return 0

        if self.is_pipe_insulated(element):
            if self.pipe_diametr in dict_var_pins:
                return self.mid_calculation_fix(dict_var_pins[self.pipe_diametr][0])
            else:
                return self.mid_calculation_fix(7)
        else:
            if self.pipe_diametr in dict_var_pins:
                return self.mid_calculation_fix(dict_var_pins[self.pipe_diametr][1])
            else:
                return self.mid_calculation_fix(5)

    def collars(self, element):
        self.name = "{0}, Ду{1}".format(self.name, int(self.pipe_diametr))
        self.local_description = "{0} {1}".format(self.local_description, self.name)
        dict_var_collars = {15:[2, 1.5], 20:[3, 2], 25:[3.5, 2], 32:[4, 2.5], 40:[4.5, 3], 50:[5, 3], 65:[6, 4],
                            80:[6, 4], 100:[6, 4.5], 125:[7, 5]}

        if self.length < 0.5:
            return 0

        if self.is_pipe_insulated(element):
            if self.pipe_diametr in dict_var_collars:
                return self.mid_calculation_fix(dict_var_collars[self.pipe_diametr][0])
            else:
                return self.mid_calculation_fix(7)
        else:
            if self.pipe_diametr in dict_var_collars:
                return self.mid_calculation_fix(dict_var_collars[self.pipe_diametr][1])
            else:
                return self.mid_calculation_fix(5)

    def duct_material(self, element):
        area = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA),
            UnitTypeId.SquareMeters)  # Преобразование в квадратные метры

        if element.DuctType.Shape == ConnectorProfileType.Round:
            diameter = self.duct_diametr
            perimeter = 3.14 * diameter

        if element.DuctType.Shape == ConnectorProfileType.Rectangular:
            width = UnitUtils.ConvertFromInternalUnits(
                element.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM),
                UnitTypeId.Millimeters)  # Преобразование в метры

            height = UnitUtils.ConvertFromInternalUnits(
                element.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM),
                UnitTypeId.Millimeters)  # Преобразование в метры

            perimeter = 2 * (width + height)


        if perimeter < 1001:
            kg = area * 0.65
        elif perimeter < 1801:
            kg = area * 1.22
        else:
            kg = area * 2.25

        return kg

    def pipe_material(self, element):
        dict_var_p_mat = {15: 0.14, 20: 0.12, 25: 0.11, 32: 0.1, 40: 0.11, 50: 0.144, 65: 0.195,
                            80: 0.233, 100: 0.37, 125: 0.53}
        up_coeff = 1.7
        # Запас 70% задан по согласованию.

        # Сортируем ключи словаря
        sorted_keys = sorted(dict_var_p_mat.keys())

        # Ищем первый ключ, который больше или равен self.pipe_diametr
        for key in sorted_keys:
            if self.pipe_diametr <= key:
                key_up = dict_var_p_mat[key] * up_coeff
                return key_up * self.length
        else:
            return 0.62*up_coeff*self.length

    def grunt(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area / 10
        return number

    def colorBT(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area * 0.2 * 2
        return number

    def get_number(self, element, name):
        number = 1
        if name == "Металлические крепления для трубопроводов" and element in col_pipes:
            number = self.pipe_material(element)
        if name == "Металлические крепления для воздуховодов" and element in col_curves:
            number = self.duct_material(element)
        if name == "Краска антикоррозионная за два раза" and element in col_pipes:
            number = self.colorBT(element)
        if name == "Грунтовка для стальных труб" and element in col_pipes:
            number = self.grunt(element)
        if name == "Хомут трубный под шпильку М8" and element in col_pipes:
            number = self.collars(element)
        if name == "Шпилька М8 1м/1шт" and element in col_pipes:
            number = self.pins(element)
        return number

def is_object_to_generate(element, gen_col, collection, parameter, gen_list = genList):
    if element in gen_col:
        for gen in gen_list:
            if gen.collection == collection and parameter == gen.method:
                try:
                    elem_type = doc.GetElement(element.GetTypeId())
                    if elem_type.GetSharedParamValueOrDefault(parameter) == 1:
                        return True
                except Exception:
                    print parameter
                    if element.GetSharedParamValueOrDefault(parameter) == 1:
                        return True

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with revit.Transaction("Добавление расчетных элементов"):
        # при каждом повторе расчета удаляем старые версии
        remove_models(col_model, name_of_model, description)

        collections = [col_insulation, col_pipes, col_curves]

        elements_to_generate = []

        #перебираем элементы и выясняем какие из них подлежат генерации
        for collection in collections:
            for element in collection:
                for gen in genList:
                    binding_name = gen.name
                    binding_mark = gen.mark
                    binding_maker = gen.maker
                    parameter = gen.method
                    genCol = gen.collection
                    if is_object_to_generate(element, genCol, collection, parameter):
                        definition = CalculationElement(element, collection, parameter, binding_name, binding_mark, binding_maker)

                        key = definition.EF + definition.corp + definition.sec + definition.floor + definition.system + \
                              definition.group + definition.name + definition.mark + definition.art + \
                              definition.maker + definition.local_description

                        toAppend = True
                        for element_to_generate in elements_to_generate:
                            if element_to_generate.key == key:
                                toAppend = False
                                element_to_generate.number = element_to_generate.number + definition.number

                        if toAppend:
                            elements_to_generate.append(definition)

        #иначе шпилек получится дробное число, а они в штуках
        for el in elements_to_generate:
            if el.name == "Шпилька М8 1м/1шт":
                el.number = int(math.ceil(el.number))

        new_position(elements_to_generate, temporary, name_of_model, description)

temporary = isFamilyIn(BuiltInCategory.OST_GenericModel, name_of_model)

if isItFamily():
    forms.alert(
        "Надстройка не предназначена для работы с семействами",
        "Ошибка",
        exitscript=True
        )

if temporary is None:
    forms.alert(
        "Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.",
        "Ошибка",
        exitscript=True
        )

status = paraSpec.check_parameters()
if not status:
    anchor = checkAnchor.check_anchor(showText = False)
    if anchor:
        script_execute()

