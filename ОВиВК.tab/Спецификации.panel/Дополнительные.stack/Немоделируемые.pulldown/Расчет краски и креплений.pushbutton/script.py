#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Расчет краски и креплений'
__doc__ = "Генерирует в модели элементы с расчетом количества соответствующих материалов"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")


import sys
import System
import dosymep
import paraSpec
import checkAnchor
import math

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)


from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from Autodesk.Revit.DB import *

from Autodesk.Revit.UI import TaskDialog

from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from Redomine import *


from System.Runtime.InteropServices import Marshal
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert




#Исходные данные
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
colPipes = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colModel = make_col(BuiltInCategory.OST_GenericModel)
colSystems = make_col(BuiltInCategory.OST_DuctSystem)
colInsul = make_col(BuiltInCategory.OST_DuctInsulations)
nameOfModel = '_Якорный элемент'
description = 'Расчет краски и креплений'


class generationElement:
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
    generationElement(group = '12. Расчетные элементы', name = "Металлические крепления для воздуховодов", mark = '', art = '', unit = 'кг.', maker = '',method = 'ФОП_ВИС_Расчет металла для креплений', collection=colCurves,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Металлические крепления для трубопроводов", mark = '', art = '', unit = 'кг.', maker = '', method =  'ФОП_ВИС_Расчет металла для креплений', collection= colPipes,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Изоляция для фланцев и стыков", mark = '', art = '', unit = 'м².', maker = '', method =  'ФОП_ВИС_Совместно с воздуховодом', collection= colInsul,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Краска антикоррозионная за два раза", mark = 'БТ-177', art = '', unit = 'кг.', maker = '', method =  'ФОП_ВИС_Расчет краски и грунтовки', collection= colPipes,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Грунтовка для стальных труб", mark = 'ГФ-031', art = '', unit = 'кг.', maker = '', method =  'ФОП_ВИС_Расчет краски и грунтовки', collection= colPipes,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Хомут трубный под шпильку М8", mark = '', art = '', unit = 'шт.', maker = '', method =  'ФОП_ВИС_Расчет хомутов', collection= colPipes,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Шпилька М8 1м/1шт", mark = '', art = '', unit = 'шт.', maker = '', method =  'ФОП_ВИС_Расчет хомутов', collection= colPipes,isType= False)
]

def roundup(divider, number):
    x = number/divider
    y = int(number/divider)
    if x - y > 0.2:
        return int(number) + 1
    else:
        return int(number)


class collar_variant:
    def __init__(self, diameter, isInsulated):
        self.diameter = diameter
        self.isInsulated = isInsulated



class calculation_element:
    pipe_insulation_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
    def __init__(self, element, collection, parameter, Name, Mark, Maker):
        self.local_description = description
        self.corp = str(element.LookupParameter('ФОП_Блок СМР').AsString())
        self.sec = str(element.LookupParameter('ФОП_Секция СМР').AsString())
        self.floor = str(element.LookupParameter('ФОП_Этаж').AsString())
        self.length = UnitUtils.ConvertFromInternalUnits(element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH),
                                                         UnitTypeId.Meters)
        self.diametr = UnitUtils.ConvertFromInternalUnits(element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM),
                                                     UnitTypeId.Millimeters)

        if element.LookupParameter('ФОП_ВИС_Имя системы'):
            self.system = str(element.LookupParameter('ФОП_ВИС_Имя системы').AsString())
        else:
            try:
                self.system = str(element.LookupParameter('ADSK_Имя системы').AsString())
            except:
                self.system = 'None'
        self.group ='12. Расчетные элементы'
        self.name = Name
        self.mark = Mark
        self.art = ''
        self.maker = Maker
        self.unit = 'None'
        self.number = self.get_number(element, self.name)
        self.mass = ''
        self.comment = ''
        self.EF = str(element.LookupParameter('ФОП_Экономическая функция').AsString())
        self.parentId = element.Id.IntegerValue


        for gen in genList:
            if gen.collection == collection and parameter == gen.method:
                self.unit = gen.unit
                isType = gen.isType


        if parameter == 'ФОП_ВИС_Совместно с воздуховодом':
            pass

        #self.number = self.get_number(element, self.name)

        elemType = doc.GetElement(element.GetTypeId())
        if element in colInsul and elemType.LookupParameter('ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
            self.name = 'Изоляция для фланцев и стыков (' + get_ADSK_Name(element) + ')'

        self.key = self.EF + self.corp + self.sec + self.floor + self.system + \
                   self.group + self.name + self.mark + self.art + \
                   self.maker + self.local_description

    def is_pipe_insulated(self, element):
        dependent_elements = element.GetDependentElements(self.pipe_insulation_filter)
        return len(dependent_elements) > 0

    def mid_calculation_fix(self, coeff):
        num = self.length / coeff
        if num < 1:
            num = 1
        return int(num)

    def pins(self, element):
        self.local_description = '{0} {1}, Ду{2}'.format(self.local_description, self.name,self.diametr)
        dict_var_pins = {15: [2, 1.5], 20: [3, 2], 25: [3.5, 2], 32: [4, 2.5], 40: [4.5, 3], 50: [5, 3], 65: [6, 4],
                            80: [6, 4], 100: [6, 4.5], 125: [7, 5]}

        # Мы не считаем крепление труб до 0.5 м
        if self.length < 0.5:
            return 0

        if self.is_pipe_insulated(element):
            if self.diametr in dict_var_pins:
                return self.mid_calculation_fix(dict_var_pins[self.diametr][0])
            else:
                return self.mid_calculation_fix(7)
        else:
            if self.diametr in dict_var_pins:
                return self.mid_calculation_fix(dict_var_pins[self.diametr][1])
            else:
                return self.mid_calculation_fix(5)

    def collars(self, element):
        self.name = '{0}, Ду{1}'.format(self.name, int(self.diametr))
        self.local_description = '{0} {1}'.format(self.local_description, self.name)
        dict_var_collars = {15:[2, 1.5], 20:[3, 2], 25:[3.5, 2], 32:[4, 2.5], 40:[4.5, 3], 50:[5, 3], 65:[6, 4],
                            80:[6, 4], 100:[6, 4.5], 125:[7, 5]}

        if self.length < 0.5:
            return 0

        if self.is_pipe_insulated(element):
            if self.diametr in dict_var_collars:
                return self.mid_calculation_fix(dict_var_collars[self.diametr][0])
            else:
                return self.mid_calculation_fix(7)
        else:
            if self.diametr in dict_var_collars:
                return self.mid_calculation_fix(dict_var_collars[self.diametr][1])
            else:
                return self.mid_calculation_fix(5)

    def duct_material(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903) / 100

        if str(element.DuctType.Shape) == "Round":
            D = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            P = 3.14 * D
        if str(element.DuctType.Shape) == "Rectangular":
            A = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            B = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            P = 2 * (A + B)


        if P < 1001:
            kg = area * 65
        elif P < 1801:
            kg = area * 122
        else:
            kg = area * 225

        return kg

    def pipe_material(self, element):
        dict_var_p_mat = {15: 0.14, 20: 0.12, 25: 0.11, 32: 0.1, 40: 0.11, 50: 0.144, 65: 0.195,
                            80: 0.233, 100: 0.37, 125: 0.53}
        up_coeff = 1.7
        # Запас 70% задан по согласованию.
        if self.diametr in dict_var_p_mat:
            key_up = dict_var_p_mat[self.diametr] * up_coeff
            return key_up*self.length
        else:
            return 0.62*up_coeff*self.length

    def insul_stock(self, element):
        area = element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA)
        if area == None:
            area = 0
        area = area * 0.092903 * 0.03
        return area

    def grunt(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area / 10
        return number

    def colorBT(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area * 0.2 * 2
        return number

    def get_number(self, element, name):
        Number = 1
        if name == "Металлические крепления для трубопроводов" and element in colPipes:
            Number = self.pipe_material(element)
        if name == "Металлические крепления для воздуховодов" and element in colCurves:
            Number = self.duct_material(element)
        if name == "Изоляция для фланцев и стыков" and element in colInsul:
            Number = self.insul_stock(element)
        if name == "Краска антикоррозионная за два раза" and element in colPipes:
            Number = self.colorBT(element)
        if name == "Грунтовка для стальных труб" and element in colPipes:
            Number = self.grunt(element)
        if name == "Хомут трубный под шпильку М8" and element in colPipes:
            Number = self.collars(element)
        if name == "Шпилька М8 1м/1шт" and element in colPipes:
            Number = self.pins(element)


        return Number



def is_object_to_generate(element, genCol, collection, parameter, genList = genList):
    if element in genCol:
        for gen in genList:
            if gen.collection == collection and parameter == gen.method:
                try:
                    elemType = doc.GetElement(element.GetTypeId())
                    if elemType.LookupParameter(parameter).AsInteger() == 1:
                        return True
                except Exception:
                    print parameter
                    if element.LookupParameter(parameter).AsInteger() == 1:
                        return True

def script_execute():
    with revit.Transaction("Добавление расчетных элементов"):
        # при каждом повторе расчета удаляем старые версии
        remove_models(colModel, nameOfModel, description)

        #список элементов которые будут сгенерированы
        calculation_elements = []

        collpasing_objects = []

        collections = [colInsul, colPipes, colCurves]

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
                        definition = calculation_element(element, collection, parameter, binding_name, binding_mark, binding_maker)

                        #

                        key = definition.EF + definition.corp + definition.sec + definition.floor + definition.system + \
                                          definition.group + definition.name + definition.mark + definition.art + \
                                          definition.maker + definition.local_description

                        # key = definition.corp + definition.sec + definition.floor + definition.system + \
                        #                   definition.group + definition.name + definition.mark + definition.art + \
                        #                   definition.maker + definition.local_description


                        toAppend = True
                        for element_to_generate in elements_to_generate:
                            if element_to_generate.key == key:
                                toAppend = False
                                element_to_generate.number = element_to_generate.number + definition.number

                        if toAppend:
                            elements_to_generate.append(definition)

        #иначе шпилек получится дробное число, а они в штуках
        for el in elements_to_generate:
            if el.name == 'Шпилька М8 1м/1шт':
                el.number = int(math.ceil(el.number))

        new_position(elements_to_generate, temporary, nameOfModel, description)



temporary = isFamilyIn(BuiltInCategory.OST_GenericModel, nameOfModel)

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

if temporary == None:
    print 'Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()



status = paraSpec.check_parameters()
if not status:
    anchor = checkAnchor.check_anchor(showText = False)
    if anchor:
        script_execute()

