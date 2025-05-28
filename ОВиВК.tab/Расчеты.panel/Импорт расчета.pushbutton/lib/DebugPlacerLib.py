#! /usr/bin/env python
# -*- coding: utf-8 -*-


import clr
import datetime
import os
import json
import codecs
from pyrevit.labs import target
from pyrevit.revit.db.query import get_family_symbol
from unicodedata import category

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import UnitTypeId, UnitUtils, XYZ, Structure
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.DB import (BuiltInCategory,
                               ElementFilter,
                               LogicalOrFilter,
                               ElementCategoryFilter,
                               ElementMulticategoryFilter,
                               Line,
                               ElementTransformUtils)

from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List
from System import Guid
from System import Environment
from pyrevit import forms
from pyrevit import revit
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

import dosymep
from pyrevit import forms, revit, script, HOST_APP, EXEC_PARAMS
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from dosymep_libs.bim4everyone import *

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

class DebugPlacer:
    FAMILY_NAME = "B4E_Дебаг цилиндры"

    def __init__(self, doc, diameter):
        self.doc = doc
        self.symbol = self.__find_family_symbol(self.FAMILY_NAME)
        self.diameter = UnitUtils.ConvertToInternalUnits(diameter, UnitTypeId.Millimeters)

    def __find_family_symbol(self, family_name):
        family_symbols = FilteredElementCollector(self.doc).OfClass(FamilySymbol).ToElements()
        family_symbol = next((fs for fs in family_symbols if fs.Family.Name == family_name), None)
        if not family_symbol:
            forms.alert("Семейство цилиндра не найдено", "Ошибка", exitscript=True)

        return family_symbol

    def __get_elements_by_category(self, category):
        """
        Возвращает список элементов по их категории.

        Args:
            category: Категория элементов.

        Returns:
            List[Element]: Список элементов.
        """
        return FilteredElementCollector(self.doc) \
            .OfCategory(category) \
            .WhereElementIsNotElementType() \
            .ToElements()

    def remove_models(self):
        """
        Удаляет элементы с переданным описанием.

        Args:
            description: Описание элемента.
        """
        # Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
        generic_model_collection = \
            [elem for elem in self.__get_elements_by_category(BuiltInCategory.OST_GenericModel) if elem.GetElementType()
            .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == self.FAMILY_NAME]

        for element in generic_model_collection:
            self.doc.Delete(element.Id)

    def place_symbol(self, x, y, z, height):
        self.symbol.Activate()

        # Конвертация из миллиметров в внутренние единицы (футы)
        x = UnitUtils.ConvertToInternalUnits(x, UnitTypeId.Millimeters)
        y = UnitUtils.ConvertToInternalUnits(y, UnitTypeId.Millimeters)
        z = UnitUtils.ConvertToInternalUnits(z, UnitTypeId.Millimeters)
        height = UnitUtils.ConvertToInternalUnits(height, UnitTypeId.Millimeters)

        location = XYZ(x, y, z)

        instance = self.doc.Create.NewFamilyInstance(
            location,
            self.symbol,
            Structure.StructuralType.NonStructural)

        instance.SetParamValue("Диаметр", self.diameter)
        instance.SetParamValue("Высота", height)
