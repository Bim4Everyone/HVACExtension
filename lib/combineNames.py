#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Сформировать имя'
__doc__ = "Генерирует имена и марки арматуры воздуховодов из их маски, и ADSK_ размеров"

import os.path as op
import clr

import Redomine

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
import sys
import paraSpec
from Autodesk.Revit.DB import *
from Redomine import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output

doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView
colAccessory = make_col(BuiltInCategory.OST_DuctAccessory)



report_rows = []
def isElementEditedBy(element):
    try:
        edited_by = element.GetParamValue(BuiltInParameter.EDITED_BY)
    except Exception:
        edited_by = __revit__.Application.Username

    if edited_by and edited_by != __revit__.Application.Username:
        if edited_by not in report_rows:
            report_rows.append(edited_by)
        return True
    return False


def check_name_mas(element, elemType):
    pass

def check_mask(paraName, element, elemType):
    mark_mask = None
    if elemType.LookupParameter(paraName):
        mask_para = elemType.LookupParameter(paraName)

        if mask_para.AsString() == None or mask_para.AsString() == "":
            return None
        else:
            return mask_para.AsString()



def script_execute():
    for element in colAccessory:
        if not isElementEditedBy(element):
            name_para = Redomine.get_ADSK_Name(element, return_para=True)

            mark_para = Redomine.get_ADSK_Mark(element, return_para=True)

            if element.LookupParameter("ADSK_Размер_Ширина"):
                width_para = element.LookupParameter("ADSK_Размер_Ширина")
            else:
                width_para = None
            if element.LookupParameter("ADSK_Размер_Высота"):
                height_para = element.LookupParameter("ADSK_Размер_Высота")
            else:
                height_para = None
            if element.LookupParameter("ADSK_Размер_Длина"):
                length_para = element.LookupParameter("ADSK_Размер_Длина")
            else:
                length_para = None
            if element.LookupParameter("ADSK_Размер_Диаметр"):
                diameter_para = element.LookupParameter("ADSK_Размер_Диаметр")
            else:
                diameter_para = None


            if name_para and mark_para and width_para and height_para and length_para or \
                    name_para and mark_para and diameter_para and length_para:
                elemType = doc.GetElement(element.GetTypeId())
                mark_mask = check_mask("ФОП_ВИС_Маска марки", element, elemType)
                name_mask = check_mask("ФОП_ВИС_Маска наименования", element, elemType)

                if mark_mask:
                    if "ВЫСОТА" in mark_mask:
                        mark_mask = mark_mask.replace("ВЫСОТА", height_para.AsValueString())
                    if "ДЛИНА" in mark_mask:
                        mark_mask = mark_mask.replace("ДЛИНА", length_para.AsValueString())
                    if "ШИРИНА" in mark_mask:
                        mark_mask = mark_mask.replace("ШИРИНА", width_para.AsValueString())
                    if "ДИАМЕТР" in mark_mask:
                        mark_mask = mark_mask.replace("ДИАМЕТР", diameter_para.AsValueString())
                        print int(fromRevitToMilimeters(diameter_para.AsDouble()))
                    mark_para.Set(mark_mask)

                if name_mask:
                    if "ВЫСОТА" in name_mask:
                        name_mask = name_mask.replace("ВЫСОТА", height_para.AsValueString())
                    if "ДЛИНА" in name_mask:
                        name_mask = name_mask.replace("ДЛИНА", length_para.AsValueString())
                    if "ШИРИНА" in name_mask:
                        name_mask = name_mask.replace("ШИРИНА", width_para.AsValueString())
                    if "ДИАМЕТР" in name_mask:
                        name_mask = name_mask.replace("ДИАМЕТР", diameter_para.AsValueString())
                    name_para.Set(name_mask)



def combine_names():
    if isItFamily():
        print 'Надстройка не предназначена для работы с семействами'
        sys.exit()

    parametersAdded = paraSpec.check_parameters()


    if not parametersAdded:
        information = doc.ProjectInformation
        with revit.Transaction("Формирование имен и марок"):
            script_execute()
            for report in report_rows:
                print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report

def solo_combine():
    with revit.Transaction("Формирование имен и марок"):
        script_execute()
        for report in report_rows:
            print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report