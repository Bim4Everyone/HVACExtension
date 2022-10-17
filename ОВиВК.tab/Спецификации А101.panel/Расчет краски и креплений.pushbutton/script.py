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


from System.Runtime.InteropServices import Marshal
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert





doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
colPipes = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colModel = make_col(BuiltInCategory.OST_GenericModel)
colSystems = make_col(BuiltInCategory.OST_DuctSystem)
# create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_GenericModel)
collector.OfClass(FamilySymbol)
famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()





is_temporary_in = False

for element in famtypeitr:
    famtypeID = element
    famsymb = doc.GetElement(famtypeID)

    if famsymb.Family.Name == '_Якорный элемен(металл и краска)':
        temporary = famsymb
        is_temporary_in = True

if is_temporary_in == False:
    print 'Не обнаружен якорный элемен(металл и краска). Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()

paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Единица измерения' ,'ФОП_ВИС_Масса', 'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число', 'ФОП_ВИС_Узел', 'ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка',
             'ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ТИП_Назначение', 'ФОП_ТИП_Число', 'ФОП_ТИП_Единица измерения',
             'ФОП_ТИП_Код', 'ФОП_ТИП_Наименование работы', 'ФОП_ВИС_Имя трубы из сегмента', 'ФОП_ВИС_Позиция', 'ФОП_ВИС_Площади воздуховодов в примечания',
             'ФОП_ВИС_Нумерация позиций', 'ФОП_ВИС_Расчет комплектов заделки', 'ФОП_ВИС_Расчет краски и грунтовки', 'ФОП_ВИС_Расчет металла для креплений']


def setElement(element, name, setting):
    if name == "ФОП_ВИС_Масса":
        element.LookupParameter(name).Set(str(setting))



    if name == 'ADSK_Единица измерения':
        element.LookupParameter('ФОП_ТИП_Единица измерения').Set(str(setting))
        element.LookupParameter('ФОП_ВИС_Единица измерения').Set(str(setting))
    try:
        if setting == None:
            pass
        else:
            element.LookupParameter(name).Set(setting)
            if name == 'ФОП_ВИС_Число':
                element.LookupParameter('ФОП_ТИП_Число').Set(str(setting))
            if name == 'ФОП_ВИС_Наименование комбинированное':
                element.LookupParameter('ФОП_ТИП_Назначение').Set(setting)

    except Exception:
        pass


def duct_material(element):
    area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903) / 100
    if element.GetParamValue(BuiltInParameter.RBS_EQ_DIAMETER_PARAM) == element.GetParamValue(BuiltInParameter.RBS_HYDRAULIC_DIAMETER_PARAM):
        D = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        P = 3.14 * D
    else:
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

def new_position(calculation_elements):
    # создаем заглушки по элементов собранных из таблицы

    loc = XYZ(0, 0, 0)

    temporary.Activate()
    for element in calculation_elements:
        familyInst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)

    # собираем список из созданных заглушек
    colModel = make_col(BuiltInCategory.OST_GenericModel)
    Models = []
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемен(металл и краска)':
            try:
                element.CreatedPhaseId = phaseid
            except Exception:
                print
                'Не удалось присвоить стадию спецификация, проверьте список стадий'

            Models.append(element)


    # для первого элмента списка заглушек присваиваем все параметры, после чего удаляем его из списка
    for element in calculation_elements:
        group = str(element[3]) + str(element[4]) + str(element[5])
        dummy = Models[0]
        setElement(dummy, 'ADSK_Имя системы', element[0])
        setElement(dummy, 'ФОП_ТИП_Код', element[1])
        setElement(dummy, 'ФОП_ТИП_Наименование работы', element[2])
        setElement(dummy, 'ФОП_ВИС_Группирование', group)
        setElement(dummy, 'ФОП_ВИС_Наименование комбинированное', element[4])
        setElement(dummy, 'ADSK_Марка', element[5])
        setElement(dummy, 'ADSK_Код изделия', element[6])
        setElement(dummy, 'ADSK_Завод-изготовитель', element[7])
        setElement(dummy, 'ADSK_Единица измерения', element[8])
        setElement(dummy, 'ФОП_ВИС_Число', element[9])
        setElement(dummy, 'ФОП_ВИС_Масса', element[10])
        setElement(dummy, 'ADSK_Примечание', element[11])
        setElement(dummy, 'ФОП_Экономическая функция', element[12])
        Models.pop(0)



#проверка на наличие нужных параметров
map = doc.ParameterBindings
it = map.ForwardIterator()
while it.MoveNext():
    newProjectParameterData = it.Key.Name
    if str(newProjectParameterData) in paraNames:
        paraNames.remove(str(newProjectParameterData))
if len(paraNames) > 0:
    try:
        print 'Были добавлен параметры, перезапустите скрипт'
        import paraSpec

    except Exception:
        print 'Не удалось добавить параметры'

else:


    for element in colModel:
        edited_by = element.LookupParameter('Редактирует').AsString()
        if edited_by and edited_by != __revit__.Application.Username:
            print "Якорные элементы не были обработаны, так как были заняты пользователями:"
            print edited_by
            sys.exit()

    with revit.Transaction("Добавление расчетных элементов"):
        #при каждом повторе расчета удаляем старые версии
        for element in colModel:
            if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемен(металл и краска)':
                doc.Delete(element.Id)

        calculation_elements = []

        for phase in doc.Phases:
            if phase.Name == 'Спецификация':
                phaseid = phase.Id

        duct_metal = []
        duct_dict = {}
        for element in colCurves:
            Class = ''
            Work = ''
            Group = '8. Расчетные элементы'
            Name = "Металлические крепления для воздуховодов"
            Mark = ''
            Art = ''
            Maker = ''
            Izm = 'кг.'
            Mass = ''
            Comment = ''

            elemType = doc.GetElement(element.GetTypeId())
            if elemType.LookupParameter('ФОП_ВИС_Расчет металла для креплений').AsInteger() == 1:

                EF = str(element.LookupParameter('ФОП_Экономическая функция').AsString())
                System = str(element.LookupParameter('ADSK_Имя системы').AsString())
                Key = EF + " " + System
                Number = duct_material(element)

                if Key not in duct_dict:
                    duct_dict[Key] = Number
                else:
                    duct_dict[Key] = duct_dict[Key] + Number



        for duct in duct_dict:
            key = str(duct).split()
            EF = key[0]

            duct_metal.append([key[1], Class, Work, Group, Name, Mark, Art, Maker, Izm, duct_dict[duct], Mass, Comment, EF])

        # в следующем блоке генерируем новые экземпляры пустых семейств куда уйдут расчеты
        #new_position(duct_metal)


        if doc.ProjectInformation.LookupParameter('ФОП_ВИС_Расчет комплектов заделки').AsInteger() == 1:
            Class = ''
            Work = ''
            Group = '8. Расчетные элементы'
            Name = "Комплект заделки отверстий с восстановлением предела огнестойкости"
            Mark = ''
            Art = ''
            Maker = ''
            Izm = 'компл.'
            Mass = ''
            Comment = ''

            EF_dict = {}
            system_list = []
            for element in colCurves:
                EF = str(element.LookupParameter('ФОП_Экономическая функция').AsString())
                System = str(element.LookupParameter('ADSK_Имя системы').AsString())

                if System not in EF_dict:
                    EF_dict[System] = EF

            for system in EF_dict:
                system_list.append(
                    [system, Class, Work, Group, Name, Mark, Art, Maker, Izm, 1, Mass, Comment, EF_dict[system]])

            # в следующем блоке генерируем новые экземпляры пустых семейств куда уйдут расчеты
            #new_position(system_list)

        calculation_elements = system_list + duct_metal
        new_position(calculation_elements)









