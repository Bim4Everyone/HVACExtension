#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Импорт немоделируемых'
__doc__ = "Генерирует в модели элементы в соответствии с их ведомостью"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import sys
import System
import dosymep

from UnmodelingClassLibrary import UnmodelingFactory, MaterialCalculator, RowOfSpecification

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict
from UnmodelingClassLibrary import  *
from dosymep_libs.bim4everyone import *

from Microsoft.Office.Interop import Excel
from Redomine import *
from rpw.ui.forms import select_file


#Исходные данные
doc = __revit__.ActiveUIDocument.Document
unmodeling_factory = UnmodelingFactory()
view = doc.ActiveView
description = 'Импорт немоделируемых'

def find_column(worksheet, search_value):
    found_cell = worksheet.Cells.Find(What=search_value)
    if found_cell is not None:
        return found_cell.Column
    else:
        error = "Ячейка с содержимым '{}' не найдена.".format(search_value)

        forms.alert(
            error,
            "Ошибка",
            exitscript=True)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    family_symbol = unmodeling_factory.startup_checks(doc)
    exel = Excel.ApplicationClass()
    # Создание объекта TextInput

    filepath = select_file()

    if filepath is None:
        sys.exit()

    workbook = exel.Workbooks.Open(filepath, ReadOnly=True)
    sheet_name = 'Импорт'

    try:
        worksheet = workbook.Sheets[sheet_name]
    except Exception:
        forms.alert(
            "Добавление пустого элемента возможно только на целевой спецификации.",
            "Ошибка",
            exitscript=True)

    # Находим последнюю заполненную строку
    used_range = worksheet.UsedRange
    last_row = used_range.Rows.Count

    function_column = find_column(worksheet, 'Экономическая функция')
    system_name_column = find_column(worksheet, 'Имя системы')
    group_column = find_column(worksheet, 'Группирование')
    name_column = find_column(worksheet, 'Наименование')
    mark_column = find_column(worksheet, 'Марка')
    code_column = find_column(worksheet, 'Код')
    maker_column = find_column(worksheet, 'Завод-изготовитель')
    unit_column = find_column(worksheet, 'Единица измерения')
    number_column = find_column(worksheet, 'Число')
    mass_column = find_column(worksheet, 'Масса')
    note_column = find_column(worksheet, 'Примечание')

    elements_to_generate = []

    found_cell = worksheet.Cells.Find('Экономическая функция')
    row = found_cell.Row + 1

    while row <= last_row:
        function = worksheet.Cells(row, function_column).value2
        system = worksheet.Cells(row, system_name_column).value2
        group = worksheet.Cells(row, group_column).value2
        name = worksheet.Cells(row, name_column).value2
        mark = worksheet.Cells(row, mark_column).value2
        code = worksheet.Cells(row, code_column).value2
        maker = worksheet.Cells(row, maker_column).value2
        unit = worksheet.Cells(row, unit_column).value2
        number = worksheet.Cells(row, number_column).value2
        mass = worksheet.Cells(row, mass_column).value2
        note = worksheet.Cells(row, note_column).value2

        if name is None:
            break

        if function is None:
            function = unmodeling_factory.out_of_function_value

        if system is None:
            system = unmodeling_factory.out_of_system_value

        try:
            number = float(number)
        except ValueError:
            error = "Значение количества в строке '{}' не является числом.".format(row)
            forms.alert(
                error,
                "Ошибка",
                exitscript=True)

        if type(mass) is not str:
            error = "Значение массы в строке '{}' не является текстом.".format(row)
            forms.alert(
                error,
                "Ошибка",
                exitscript=True)

        elements_to_generate.append(
            RowOfSpecification(
                system,
                function,
                group,
                name,
                mark,
                code,
                maker,
                unit,
                description,
                number,
                mass,
                note
            )
        )

        row += 1

    with revit.Transaction("Добавление расчетных элементов"):
        family_symbol.Activate()

        # при каждом повторе расчета удаляем старые версии
        unmodeling_factory.remove_models(doc, description)

        element_location = unmodeling_factory.get_base_location(doc)

        for element in elements_to_generate:
            element_location = unmodeling_factory.update_location(element_location)

            unmodeling_factory.create_new_position(doc, element, family_symbol, description, element_location)

    # Закрываем рабочую книгу без сохранения изменений
    workbook.Close(SaveChanges=False)

    # Закрываем приложение Excel
    exel.Quit()

script_execute()