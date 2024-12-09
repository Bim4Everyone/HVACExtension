#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Библиотека стандартных функций'
__doc__ = "pass"

import os.path as op
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Revit.Geometry import *
import sys
from Autodesk.Revit.DB import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output
from pyrevit import script


doc = __revit__.ActiveUIDocument.Document  # type: Document
import os.path as op
import clr

output = script.get_output()


def getParameter(element, paraName):
    try:
        parameter = element.LookupParameter(paraName)
        return parameter
    except:
        return None

def setIfNotRO(parameter, value):
    if not parameter.IsReadOnly:
        parameter.Set(value)



def getDefCols():
    colFittings = make_col(BuiltInCategory.OST_DuctFitting)
    colPipeFittings = make_col(BuiltInCategory.OST_PipeFitting)
    colPipeCurves = make_col(BuiltInCategory.OST_PipeCurves)
    colCurves = make_col(BuiltInCategory.OST_DuctCurves)
    colFlexCurves = make_col(BuiltInCategory.OST_FlexDuctCurves)
    colFlexPipeCurves = make_col(BuiltInCategory.OST_FlexPipeCurves)
    colTerminals = make_col(BuiltInCategory.OST_DuctTerminal)
    colAccessory = make_col(BuiltInCategory.OST_DuctAccessory)
    colPipeAccessory = make_col(BuiltInCategory.OST_PipeAccessory)
    colEquipment = make_col(BuiltInCategory.OST_MechanicalEquipment)
    colInsulations = make_col(BuiltInCategory.OST_DuctInsulations)
    colPipeInsulations = make_col(BuiltInCategory.OST_PipeInsulations)
    colPlumbingFixtures = make_col(BuiltInCategory.OST_PlumbingFixtures)
    colSprinklers = make_col(BuiltInCategory.OST_Sprinklers)

    collections = [colFittings, colPipeFittings, colCurves, colFlexCurves, colFlexPipeCurves, colInsulations,
                   colPipeInsulations, colPipeCurves, colSprinklers, colAccessory,
                   colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures]

    return collections




output = script.get_output()

def isElementEditedBy(element):
    user_name = __revit__.Application.Username
    edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
    return edited_by and edited_by.lower() != user_name.lower()

def fillReportRows(element, report_rows):
    edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
    if edited_by:
        report_rows.add(edited_by.lower())
    return report_rows


def isItFamily():
    try:
        manager = doc.FamilyManager
        return True
    except Exception:
        pass

def getDocIfItsWorkshared():
    doc = __revit__.ActiveUIDocument.Document  # type: Document
    if not doc.IsWorkshared:
        print "Документ не является файлом для общей работы. Сохраните его как проект и повторите процедуру."
        sys.exit()
    return doc

def lookupCheck(elementOrType, paraName, isExit = True):
    try:
        if element.GetTypeId():
            type = 'экземпляра '
    except:
        type = 'типа '

    parameter = getParameter(elementOrType, paraName)

    if parameter:
        return parameter
    else:
        if isExit:
            print 'Параметр ' + type + paraName + ' не назначен для категории ' + elementOrType.Category.Name + ' (ID элемента на котором найдена ошибка ' + output.linkify(elementOrType.Id) +")"
            sys.exit()
        else:
            return None



def getSharedParameter (element, paraName, replaceName, return_para = False):
    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        if paraName == 'ADSK_Наименование':
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)
            if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента').AsInteger() == 1:

                parameter = element.LookupParameter('Описание сегмента').AsString()
                return parameter

    report = paraName.split('_')[1]
    replaceDef = isReplasing(replaceName)
    name = paraName
    if replaceDef:
            name = replaceIsValid(element, paraName, replaceName)
    try:
        parameter = element.LookupParameter(name).AsString()
        if return_para:
            parameter = element.LookupParameter(name)
            return parameter
        if parameter == None:
            parameter = 'None'
    except Exception:
        #try:

        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if return_para:
            parameter = ElemType.LookupParameter(name)
            return parameter

        if ElemType.LookupParameter(name) == None:
            parameter = "None"
        else:
            parameter = ElemType.LookupParameter(name).AsString()

    nullParas = ['ADSK_Завод-изготовитель', 'ADSK_Марка', 'ADSK_Код изделия']

    if not parameter:
        parameter = 'None'

    if parameter == 'None' or parameter == None:
        if paraName in nullParas:
            parameter = ''
        else:
            parameter = 'None'

    parameter = str(parameter)

    return parameter

def fromRevitToMeters(number):
    meters = (number * 304.8)/1000
    return meters

def fromRevitToMilimeters(number):
    mms = number * 304.8
    return mms
def fromRevitToSquareMeters(number):
    square = number * 0.092903
    return square


#создает коллекцию из экземпляров категории
def make_col(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col

def isReplasing(replaceName ):
    replaceDefiniton = doc.ProjectInformation.LookupParameter(replaceName).AsString()
    if replaceDefiniton == None:
        return False
    if replaceDefiniton == 'None':
        return False
    if replaceDefiniton == '':
        return False
    else:
        return replaceDefiniton

def replaceIsValid (element, paraName, replaceName):
    replaceparameter = doc.ProjectInformation.LookupParameter(replaceName).AsString()
    if not element.LookupParameter(replaceparameter):
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if not ElemType.LookupParameter(replaceparameter):
            print output.linkify(element.Id)
            print 'Назначеного параметра замены ' + replaceparameter + ' нет у одной из категорий, назначенных для исходного параметр ' + paraName
            sys.exit()

    return replaceparameter

def get_ADSK_Izm(element):
    paraName = 'ADSK_Единица измерения'
    replaceName = 'ФОП_ВИС_Замена параметра_Единица измерения'
    ADSK_Izm = getSharedParameter(element, paraName, replaceName)
    return ADSK_Izm

def get_ADSK_Name(element, return_para = False):
    paraName = 'ADSK_Наименование'
    replaceName = 'ФОП_ВИС_Замена параметра_Наименование'
    ADSK_Name = getSharedParameter(element, paraName, replaceName, return_para)
    return ADSK_Name

def get_ADSK_Maker(element):
    paraName = 'ADSK_Завод-изготовитель'
    replaceName = 'ФОП_ВИС_Замена параметра_Завод-изготовитель'
    ADSK_Maker = getSharedParameter(element, paraName, replaceName)
    return ADSK_Maker

def get_ADSK_Mark(element, return_para = False):
    paraName = 'ADSK_Марка'
    replaceName = 'ФОП_ВИС_Замена параметра_Марка'
    ADSK_Mark = getSharedParameter(element, paraName, replaceName, return_para)
    return ADSK_Mark

def get_ADSK_Code(element):
    paraName = 'ADSK_Код изделия'
    replaceName = 'ФОП_ВИС_Замена параметра_Код изделия'
    ADSK_Code = getSharedParameter(element, paraName, replaceName)
    return ADSK_Code

def getConnectors(element):
    connectors = []
    try:
        a = element.ConnectorManager.Connectors.ForwardIterator()
        while a.MoveNext():
            connectors.append(a.Current)
    except:
        try:
            a = element.MEPModel.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
        except:
            a = element.MEPSystem.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
    return connectors


#возвращает площадь фитинга воздуховода
def get_fitting_area(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        options = Options()
        geoms = element.get_Geometry(options)

        for g in geoms:
            solids = g.GetInstanceGeometry()
        area = 0

        for solid in solids:
            if isinstance(solid, Line) or isinstance(solid, Arc):
                continue
            for face in solid.Faces:
                area = area + face.Area

        area = fromRevitToSquareMeters(area)
        connectors = getConnectors(element)

        for connector in connectors:
            try:
                H = connector.Height
                B = connector.Width
                S = (H * B)
                S = fromRevitToSquareMeters(S)
                area = area - S
            except Exception:
                R = connector.Radius
                S = (3.14 * R * R)
                S = fromRevitToSquareMeters(S)
                area = (area - S)
        area = round(area, 2)
    return area


#заполняет ячейки в сгенерированном немоделируемом
def setElement(element, name, setting):

    if setting == 'None':
        setting = ''
    if setting == None:
        setting = ''

    if name == 'ФОП_Экономическая функция':
        if setting == '':
            setting = 'None'
        if setting == None:
            setting = 'None'

    if name == 'ФОП_ВИС_Имя системы':
        if setting == '':
            setting = 'None'
        if setting == None:
            setting = 'None'


    if name == 'ФОП_ВИС_Число' or name == 'ФОП_ВИС_Масса':
        element.LookupParameter(name).Set(setting)
    else:
        try:
            element.LookupParameter(name).Set(str(setting))
        except:
            element.LookupParameter(name).Set(setting)


# Генерирует пустые элементы в рабочем наборе немоделируемых
# def new_position(calculation_elements, temporary, famName, description):
#     # Создаем заглушки по элементам, собранным из таблицы
#     loc = XYZ(0, 0, 0)
#
#     temporary.Activate()
#     col_model = []
#
#     # Находим рабочий набор "99_Немоделируемые элементы"
#     fws = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
#     workset_id = None
#     for ws in fws:
#         if ws.Name == '99_Немоделируемые элементы':
#             workset_id = ws.Id
#             break
#
#     if workset_id is None:
#         print('Не удалось найти рабочий набор "99_Немоделируемые элементы", проверьте список наборов')
#         return
#
#     # Создаем элементы и добавляем их в colModel
#     for _ in calculation_elements:
#         family_inst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)
#         col_model.append(family_inst)
#
#     # Фильтруем элементы и присваиваем рабочий набор
#     for element in col_model:
#         try:
#             elem_type = doc.GetElement(element.GetTypeId())
#             if elem_type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString() == famName:
#                 if not element.LookupParameter('ФОП_ВИС_Назначение').AsString():
#                     ews = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
#                     ews.Set(workset_id.IntegerValue)
#         except Exception as e:
#             print 'Ошибка при присвоении рабочего набора'
#
#     index = 1
#     # Для первого элемента списка заглушек присваиваем все параметры, после чего удаляем его из списка
#     for position in calculation_elements:
#         group = position.group
#         if description != 'Пустая строка':
#             posGroup = f'{position.group}_{position.name}_{position.mark}_{index}'
#             index += 1
#             if description in ['Расходники изоляции', 'Расчет краски и креплений']:
#                 posGroup = f'{position.group}_{position.name}_{position.mark}'
#             group = posGroup
#
#         dummy = col_model.pop(0) if description != 'Пустая строка' else family_inst
#
#         setElement(dummy, 'ФОП_Блок СМР', position.corp)
#         setElement(dummy, 'ФОП_Секция СМР', position.sec)
#         setElement(dummy, 'ФОП_Этаж', position.floor)
#         setElement(dummy, 'ФОП_ВИС_Имя системы', position.system)
#         setElement(dummy, 'ФОП_ВИС_Группирование', group)
#         setElement(dummy, 'ФОП_ВИС_Наименование комбинированное', position.name)
#         setElement(dummy, 'ФОП_ВИС_Марка', position.mark)
#         setElement(dummy, 'ФОП_ВИС_Код изделия', position.art)
#         setElement(dummy, 'ФОП_ВИС_Завод-изготовитель', position.maker)
#         setElement(dummy, 'ФОП_ВИС_Единица измерения', position.unit)
#         setElement(dummy, 'ФОП_ВИС_Число', position.number)
#         setElement(dummy, 'ФОП_ВИС_Масса', position.mass)
#         setElement(dummy, 'ФОП_ВИС_Примечание', position.comment)
#         setElement(dummy, 'ФОП_Экономическая функция', position.EF)
#
#         # Фильтрация шпилек под разные диаметры
#         setElement(dummy, 'ФОП_ВИС_Назначение',
#                    description if description != 'Расчет краски и креплений' else position.local_description)
#
#         if description != 'Пустая строка':
#             col_model.pop(0)

#для прогона новых ревизий генерации немоделируемых: стирает элмент с переданным именем модели
def remove_models(colModel, famName, description):
    try:
        for element in colModel:
            edited_by = isElementEditedBy(element)
            if edited_by:
                print "Якорные элементы не были обработаны, так как были заняты пользователями:"
                print edited_by
                sys.exit()
    except Exception:
        pass

    for element in colModel:
        if element.LookupParameter('ФОП_ВИС_Назначение'):
            elemType = doc.GetElement(element.GetTypeId())
            currentName = elemType.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString()
            currentDescription = element.LookupParameter('ФОП_ВИС_Назначение').AsString()
            #
            # print description
            # print currentDescription
            # print currentDescription == description
            if  currentName == famName:
                if description in currentDescription:
                    doc.Delete(element.Id)

#класс содержащий все ячейки типовой спецификации
class rowOfSpecification:
    def __init__(self, corp, sec, floor, system, group, name, mark, art, maker, unit, number, mass, comment, EF):
        self.corp = corp
        self.sec = sec
        self.floor = floor
        self.system = system
        self.group = group
        self.name = name
        self.mark = mark
        self.art = art
        self.maker = maker
        self.unit = unit
        self.number = number
        self.mass = mass
        self.comment = comment
        self.EF = EF

# Возвращает FamilySymbol, если семейство есть в проекте, None если нет
def isFamilyIn(builtin, name):
    # Создаем фильтрованный коллектор для категории OST_Mass и класса FamilySymbol
    collector = FilteredElementCollector(doc).OfCategory(builtin).OfClass(FamilySymbol)

    # Итерируемся по элементам коллектора
    for element in collector:
        if element.Family.Name == name:
            return element

    return None