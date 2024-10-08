#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление параметров'
__doc__ = "Добавление параметров в семейство для за полнения спецификации и экономической функции"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import checkAnchor
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)


from Autodesk.Revit.DB import *
import os
from pyrevit import revit
import sys

doc = __revit__.ActiveUIDocument.Document  # type: Document
parameters_added = False


class shared_parameter:
    def insert(self, eDef, set, istype, group):

        if istype:
            newBinding = doc.Application.Create.NewTypeBinding(set)
        else:
            newBinding = doc.Application.Create.NewInstanceBinding(set)

        doc.ParameterBindings.Insert(eDef, newBinding, group)

    def default_values(self, name):
        pipeparaNames = ['ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка', 'ФОП_ВИС_Имя трубы из сегмента',
                         'ФОП_ВИС_Расчет краски и грунтовки', 'ФОП_ВИС_Расчет хомутов']
        if name in pipeparaNames:
            colPipeTypes = FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements()
            for element in colPipeTypes:
                if name == 'ФОП_ВИС_Днар х Стенка':
                    element.LookupParameter(name).Set(1)
                else:
                    element.LookupParameter(name).Set(0)

        if name == 'ФОП_ВИС_Запас воздуховодов/труб':
            inf = doc.ProjectInformation.LookupParameter(name)
            inf.Set(10)

        if name == 'ФОП_ВИС_Запас изоляции':
            inf = doc.ProjectInformation.LookupParameter(name)
            inf.Set(15)

        if name == 'ФОП_ВИС_Запас изоляции воздуховодов':
            inf = doc.ProjectInformation.LookupParameter(name)
            inf.Set(20)

        if name == 'ФОП_ВИС_Запас изоляции труб':
            inf = doc.ProjectInformation.LookupParameter(name)
            inf.Set(8)


    def __init__(self, name, definition, set, istype = False, group = BuiltInParameterGroup.PG_DATA):
        if not is_exists_params(name):
            self.eDef = definition.get_Item(name)
            self.insert(self.eDef, set, istype, group)
            self.default_values(name)
            global parameters_added
            parameters_added = True

def make_type_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsElementType()\
                            .ToElements()
    return col

def get_cats(list_of_cats):
    cats = []
    for cat in list_of_cats:
        cats.append(doc.Settings.Categories.get_Item(cat))
    set = doc.Application.Create.NewCategorySet()
    for cat in cats:
        set.Insert(cat)
    return set

def check_spfile(group):
    try:
        doc.Application.SharedParametersFilename = str(
            os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\ФОП_v1.txt"
        spFile = doc.Application.OpenSharedParameterFile()
    except Exception:
        print
        'По стандартному пути не найден файл общих параметров, обратитесь в BIM-отдел или замените вручную на ФОП_v1.txt'
        sys.exit()

    for dG in spFile.Groups:
        if str(dG.Name) == group:
            visDefinitions = dG.Definitions
            return visDefinitions

def is_exists_params(Name):
    if not doc.IsExistsSharedParam(Name):
        return False
    return True

def script_execute():
    # сеты категорий идущие по стандарту
    defaultCatSet = get_cats(
        [BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_PipeCurves,
         BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_FlexDuctCurves,
         BuiltInCategory.OST_FlexPipeCurves,
         BuiltInCategory.OST_DuctTerminal, BuiltInCategory.OST_DuctAccessory,
         BuiltInCategory.OST_PipeAccessory,
         BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_DuctInsulations,
         BuiltInCategory.OST_PipeInsulations,
         BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_Sprinklers,
         BuiltInCategory.OST_GenericModel])

    EFCatSet = get_cats(
        [BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_PipeCurves,
         BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_FlexDuctCurves,
         BuiltInCategory.OST_FlexPipeCurves,
         BuiltInCategory.OST_DuctTerminal, BuiltInCategory.OST_DuctAccessory,
         BuiltInCategory.OST_PipeAccessory,
         BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_DuctInsulations,
         BuiltInCategory.OST_PipeInsulations,
         BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_Sprinklers, BuiltInCategory.OST_ProjectInformation,
         BuiltInCategory.OST_GenericModel])


    # нестандартные сеты категорий
    ductCatSet = get_cats([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_DuctInsulations])
    ductInsCatSet = get_cats([BuiltInCategory.OST_DuctInsulations])
    nodeCatSet = get_cats([BuiltInCategory.OST_DuctTerminal, BuiltInCategory.OST_DuctAccessory,
                           BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment,
                           BuiltInCategory.OST_PlumbingFixtures])
    pipeCatSet = get_cats([BuiltInCategory.OST_PipeCurves])
    projectCatSet = get_cats([BuiltInCategory.OST_ProjectInformation])

    ductandpipeCatSet = get_cats([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves])
    insulandliveCatSet = get_cats([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctInsulations, BuiltInCategory.OST_PipeInsulations])
    systemsCatSet = get_cats(([BuiltInCategory.OST_DuctSystem, BuiltInCategory.OST_PipingSystem]))
    insulsCatSet = get_cats([BuiltInCategory.OST_DuctInsulations, BuiltInCategory.OST_PipeInsulations])
    accessoryCatSet = get_cats([BuiltInCategory.OST_DuctAccessory])

    #получаем группу из которой выбираем параметры
    visDefinitions = check_spfile("03_ВИС")
    arDefinitions = check_spfile("01_АР")
    genDefinitions  = check_spfile("00_Общие")

    with revit.Transaction("Добавление параметров"):
        # Параметры выводим из работы
        #shared_parameter('ФОП_ВИС_Совместно с воздуховодом', visDefinitions, ductInsCatSet, istype=True)
        # shared_parameter('ФОП_Код работы', genDefinitions, defaultCatSet)
        # shared_parameter('ФОП_Привязка к помещениям', genDefinitions, projectCatSet)
        # shared_parameter('ФОП_Помещение', arDefinitions, defaultCatSet)


        shared_parameter('ФОП_ВИС_Изол_Расходник 1_Наименование', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 1_Марка', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 1_Изготовитель', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 1_Ед. изм.', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 1_Расход на м2', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 1_Расход по м.п.', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)

        shared_parameter('ФОП_ВИС_Изол_Расходник 2_Наименование', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 2_Марка', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 2_Изготовитель', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 2_Ед. изм.', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 2_Расход на м2', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 2_Расход по м.п.', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)

        shared_parameter('ФОП_ВИС_Изол_Расходник 3_Наименование', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 3_Марка', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 3_Изготовитель', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 3_Ед. изм.', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 3_Расход на м2', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)
        shared_parameter('ФОП_ВИС_Изол_Расходник 3_Расход по м.п.', visDefinitions, insulsCatSet, istype=True,
                         group=BuiltInParameterGroup.PG_MATERIALS)

        shared_parameter('ФОП_ВИС_Живое сечение, м2', visDefinitions, accessoryCatSet, group=BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW)
        shared_parameter('ФОП_ВИС_КМС', visDefinitions, accessoryCatSet,
                         group=BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW)


        shared_parameter('ФОП_Экономическая функция', genDefinitions, EFCatSet)
        shared_parameter('ФОП_ВИС_Экономическая функция', visDefinitions, defaultCatSet, istype=True)
        shared_parameter('ФОП_ВИС_ЭФ для системы', visDefinitions, systemsCatSet, istype=True)
        shared_parameter('ФОП_Блок СМР', genDefinitions, defaultCatSet)


        shared_parameter('ФОП_Секция СМР', genDefinitions, defaultCatSet)
        shared_parameter('ФОП_Этаж', genDefinitions, defaultCatSet)

        shared_parameter('ФОП_ВИС_Имя системы', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Имя системы принудительное', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Код изделия', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Завод-изготовитель', visDefinitions, defaultCatSet)

        shared_parameter('ФОП_ВИС_Группирование', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Группирование принудительное', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Единица измерения', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Масса', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Наименование комбинированное', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Наименование принудительное', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Дополнение к имени', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Отметка оси от нуля', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Отметка низа от нуля', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Позиция', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Марка', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Число', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Примечание', visDefinitions, defaultCatSet)
        shared_parameter('ФОП_ВИС_Индивидуальный запас', visDefinitions, insulandliveCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Сокращение для системы', visDefinitions, systemsCatSet, istype=True)


        shared_parameter('ФОП_ВИС_Минимальная толщина воздуховода', visDefinitions, ductCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Максимальная толщина воздуховода', visDefinitions, ductCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Узел', visDefinitions, nodeCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Исключить из узла', visDefinitions, nodeCatSet, istype=True)

        shared_parameter('ФОП_ВИС_Ду', visDefinitions, pipeCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Ду х Стенка', visDefinitions, pipeCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Днар х Стенка', visDefinitions, pipeCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Имя трубы из сегмента', visDefinitions, pipeCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Расчет краски и грунтовки', visDefinitions, pipeCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Расчет хомутов', visDefinitions, pipeCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Расчет металла для креплений', visDefinitions, ductandpipeCatSet, istype=True)


        shared_parameter('ФОП_ВИС_Учитывать фитинги труб по типу трубы', visDefinitions, pipeCatSet, istype=True)
        shared_parameter('ФОП_ВИС_Учитывать фитинги воздуховодов', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Имя внесистемных элементов', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Учитывать фитинги труб', visDefinitions, projectCatSet)
        #Запас изоляции нужно удалить после перехода на новую спеку
        shared_parameter('ФОП_ВИС_Запас изоляции', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Запас изоляции воздуховодов', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Запас изоляции труб', visDefinitions, projectCatSet)


        shared_parameter('ФОП_ВИС_Запас воздуховодов/труб', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Замена параметра_Единица измерения', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Замена параметра_Завод-изготовитель', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Замена параметра_Код изделия', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Замена параметра_Марка', visDefinitions, projectCatSet)
        shared_parameter('ФОП_ВИС_Замена параметра_Наименование', visDefinitions, projectCatSet)



    if parameters_added:
        print 'Были добавлены параметры, перезапустите скрипт'

def check_parameters():
    checkAnchor.check_anchor()

    script_execute()
    return parameters_added
