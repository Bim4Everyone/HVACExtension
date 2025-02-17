#! /usr/bin/env python
# -*- coding: utf-8 -*-

from itertools import count

import clr

from unmodeling_class_library import UnmodelingFactory, MaterialCalculator, RowOfSpecification

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from pyrevit import forms
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict
from unmodeling_class_library import  *
from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
selected_ids = uidoc.Selection.GetElementIds()
unmodeling_factory = UnmodelingFactory()


def process_new_position(family_symbol, rows_number):
    """
    Создает новые позиции в спецификации

    Args:
        family_symbol: Символ семейства, для которого мы создадим экземпляры
        rows_number: Количество строк в спецификации

    """
    element = doc.GetElement(selected_ids[0])
    location = unmodeling_factory.get_base_location(doc)

    parent_system, parent_function = unmodeling_factory.get_system_function(element)
    parent_group = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISGrouping, '')

    for count in range(1, rows_number + 1):
        new_group = "{}{}".format(parent_group, '_' + str(count))

        new_position = RowOfSpecification(
            parent_system,
            parent_function,
            new_group
        )

        location = unmodeling_factory.update_location(location)

        unmodeling_factory.create_new_position(doc,
                                               new_position,
                                               family_symbol,
                                               unmodeling_factory.EMPTY_DESCRIPTION,
                                               location)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    family_symbol = unmodeling_factory.startup_checks(doc)

    if view.Category == None or not view.Category.IsId(BuiltInCategory.OST_Schedules):
        forms.alert(
            "Добавление пустого элемента возможно только на целевой спецификации.",
            "Ошибка",
            exitscript=True)

    if 0 == selected_ids.Count:
        forms.alert(
            "Выделите целевой элемент",
            "Ошибка",
            exitscript=True)

    rows_number = forms.ask_for_string(
        default='1',
        prompt='Введите количество пустых строк:',
        title=unmodeling_factory.EMPTY_DESCRIPTION
    )

    try:
        rows_number = int(rows_number)
    except ValueError:
        forms.alert(
            "Нужно ввести число.",
            "Ошибка",
            exitscript=True)

    with revit.Transaction("BIM: Добавление пустого элемента"):
        family_symbol.Activate()

        process_new_position(family_symbol, rows_number)

script_execute()