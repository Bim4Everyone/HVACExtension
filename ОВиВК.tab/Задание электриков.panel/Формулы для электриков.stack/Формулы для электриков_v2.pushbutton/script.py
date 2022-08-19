#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление формул'
__doc__ = "Добавление формул"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from rpw.ui.forms import SelectFromList
from System import Guid
from pyrevit import revit

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView


def make_col(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col


connectorCol = make_col(BuiltInCategory.OST_ConnectorElem)
loadsCol = make_col(BuiltInCategory.OST_ElectricalLoadClassifications)

t = Transaction(doc, 'Добавление формул')

try:
    manager = doc.FamilyManager
except Exception:
    print "Надстройка предназначена для работы с семействами"
    sys.exit()

def associate(param, famparam):
    manager.AssociateElementParameterToFamilyParameter(param, famparam)


spFile = doc.Application.OpenSharedParameterFile()

set = doc.FamilyManager.Parameters

paraNames = ['ADSK_Полная мощность', 'ADSK_Коэффициент мощности', 'ADSK_Количество фаз', 'ADSK_Напряжение',
             'ADSK_Классификация нагрузок', 'ADSK_Не нагреватель_Не шкаф', 'ADSK_Номинальная мощность', 'ADSK_Без частотного регулятора', 'mS_Имя нагрузки']

for param in set:
    if str(param.Definition.Name) in paraNames:
        paraNames.remove(param.Definition.Name)
if len(paraNames) > 0:
    print 'Необходимо добавить параметры'
    for name in paraNames:
        print name
    sys.exit()


connectorNum = 0
try:
    for connector in connectorCol:
        if str(connector.Domain) == "DomainElectrical":
            connectorNum = connectorNum + 1
except Exception:

    print "Не найдено электрических коннекторов, должен быть один"
    sys.exit()


if connectorNum > 1:
    print "Электрических коннекторов больше одного, удалите лишние"
    sys.exit()
if connectorNum == 0:
    print "Не найдено электрических коннекторов, должен быть один"
    sys.exit()


with revit.Transaction("Добавление формул"):
    #если не присвоить значение, то потом в процессе получается деление на ноль
    for param in set:
        if str(param.Definition.Name) == 'ADSK_Коэффициент мощности':
            manager.SetFormula(param, "1")
    #нельзя присвоить значение в коннекторе, если классификатор обозначен в типе никак.
    for element in loadsCol:
        if element.Name == "HVAC" or element.Name == "Прочее" or element.Name == "Other" or element.Name == "ОВ":
            for param in set:
                if str(param.Definition.Name) == 'ADSK_Классификация нагрузок':
                    manager.Set(param, element.Id)

    for connector in connectorCol:
        if str(connector.Domain) == "DomainElectrical":
            connector.SystemClassification = MEPSystemClassification.PowerBalanced
            connector.SetParamValue(BuiltInParameter.RBS_ELEC_POWER_FACTOR_STATE, 1)
    for param in set:
        if str(param.Definition.Name) == 'ADSK_Количество фаз':
            ADSK_phase = param
            manager.SetFormula(param, "if(ADSK_Напряжение < 250 В, 1, 3)")

        if str(param.Definition.Name) == 'ADSK_Напряжение':
            ADSK_U = param
        if str(param.Definition.Name) == 'ADSK_Классификация нагрузок':
            ADSK_Class = param
        if str(param.Definition.Name) == 'ADSK_Коэффициент мощности':
            ADSK_K = param
            manager.SetFormula(param, "if(ADSK_Без частотного регулятора, if(ADSK_Не нагреватель_Не шкаф, if(ADSK_Номинальная мощность < 1000 Вт, 0.65, if(ADSK_Номинальная мощность < 4000 Вт, 0.75, 0.85)), 1), 0.95)")

        if str(param.Definition.Name) == 'ADSK_Полная мощность':
            manager.SetFormula(param, "ADSK_Номинальная мощность / ADSK_Коэффициент мощности")
            ADSK_power = param

        if str(param.Definition.Name) == 'mS_Имя нагрузки':
            manager.SetFormula(param, "ADSK_Наименование")



    for connector in connectorCol:
        params = connector.GetOrderedParameters()
        for param in params:
            if param.Definition.Name == 'Коэффициент мощности':
                associate(param, ADSK_K)
            if param.Definition.Name == 'Напряжение':
                associate(param, ADSK_U)
            if param.Definition.Name == 'Полная установленная мощность':
                associate(param, ADSK_power)
            if param.Definition.Name == 'Количество полюсов':
                associate(param, ADSK_phase)
            if param.Definition.Name == 'Классификация нагрузок':
                associate(param, ADSK_Class)