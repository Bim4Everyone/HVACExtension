#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пересчет КМС'
__doc__ = "Пересчитывает КМС соединительных деталей воздуховодов"

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
from pyrevit import forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, Selection
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from collections import namedtuple

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig

class Aerodinamiccoefficientcalculator:
    LOSS_GUID_CONST = "46245996-eebb-4536-ac17-9c1cd917d8cf" # Гуид для удельных потерь
    COEFF_GUID_CONST = "5a598293-1504-46cc-a9c0-de55c82848b9" # Это - Гуид "Определенный коэффициент". Вроде бы одинаков всегда

    doc = None
    uidoc = None
    view = None

    def __init__(self, doc, uidoc, view):
        self.doc = doc
        self.uidoc = uidoc
        self.view = view

    def is_supply_air(self, connector):
        return connector.DuctSystemType == DuctSystemType.SupplyAir

    def is_exhaust_air(self, connector):
        return (connector.DuctSystemType == DuctSystemType.ExhaustAir
                or connector.DuctSystemType == DuctSystemType.ReturnAir)

    def is_direction_inside(self, connector):
        return connector.Direction == FlowDirectionType.In

    def is_direction_bidirectonal(self, connector):
        return connector.Direction == FlowDirectionType.Bidirectional

    def is_direction_outside(self, connector):
        return connector.Direction == FlowDirectionType.Out

    def convert_to_milimeters(self, value):
        return  UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.Millimeters)

    def convert_to_square_meters(self, value):
        return  UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.SquareMeters)

    def get_connectors(self, element):
        connectors = []

        if isinstance(element, FamilyInstance) and element.MEPModel.ConnectorManager is not None:
            connectors.extend(element.MEPModel.ConnectorManager.Connectors)

        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves]) and \
                isinstance(element, MEPCurve) and element.ConnectorManager is not None:
            connectors.extend(element.ConnectorManager.Connectors)

        return connectors

    def get_con_coords(self, connector):
        a0 = connector.Origin.ToString()
        a0 = a0.replace("(", "")
        a0 = a0.replace(")", "")
        a0 = a0.split(",")
        for x in a0:
            x = float(x)
        return a0

    def get_connector_area(self, connector):
        if connector.Shape == ConnectorProfileType.Round:
            radius = self.convert_to_milimeters(connector.Radius)
            area = math.pi * radius ** 2
        else:
            height = self.convert_to_milimeters(connector.Height)
            width = self.convert_to_milimeters(connector.Width)
            area = height * width
        return area

    def get_duct_coords(self, in_tee_con, connector):
        main_con = []
        connector_set = connector.AllRefs.ForwardIterator()
        while connector_set.MoveNext():
            main_con.append(connector_set.Current)
        duct = main_con[0].Owner
        duct_cons = self.get_connectors(duct)
        for duct_con in duct_cons:
            if self.get_con_coords(duct_con) != in_tee_con:
                in_duct_con = self.get_con_coords(duct_con)
                return in_duct_con

    def get_coef_elbow(self, element):
        coefficient = 3

        return coefficient

    def get_coef_transition(self, element):
        coefficient = 3

        return coefficient

    def get_coef_tee(self, element):
        coefficient = 3

        return coefficient

    def get_coef_tap_adjustable(self, element):
        coefficient = 3

        return coefficient
