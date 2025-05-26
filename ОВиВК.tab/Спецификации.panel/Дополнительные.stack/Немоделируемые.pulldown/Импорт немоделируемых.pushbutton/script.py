#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import sys
import System
import dosymep

from unmodeling_class_library import UnmodelingFactory, MaterialCalculator, RowOfSpecification

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict
from unmodeling_class_library import  *
from dosymep_libs.bim4everyone import *

from Microsoft.Office.Interop import Excel
from rpw.ui.forms import select_file

doc = __revit__.ActiveUIDocument.Document
unmodeling_factory = UnmodelingFactory(doc)
view = doc.ActiveView

def find_column(worksheet, search_value):
    """
    Ищет номер колонки на листе

    Args:
        worksheet: Рабочий лист
        search_value: Заголовок колонки. Скрипт прервется, если не найдет такую.
    Returns:
        int: Номер колонки
    """
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
    family_symbol = unmodeling_factory.startup_checks()
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
    block_column = find_column(worksheet, 'Блок СМР')
    section_column = find_column(worksheet, 'Секция СМР')
    floor_column = find_column(worksheet, 'Этаж СМР')

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
        block = worksheet.Cells(row, block_column).value2
        section = worksheet.Cells(row, section_column).value2
        floor = worksheet.Cells(row, floor_column).value2

        if name is None:
            break

        if function is None:
            function = unmodeling_factory.OUT_OF_FUNCTION_VALUE

        if system is None:
            system = unmodeling_factory.OUT_OF_SYSTEM_VALUE

        try:
            number = float(number)
        except ValueError:
            error = "Значение количества в строке '{}' не является числом.".format(row)
            forms.alert(
                error,
                "Ошибка",
                exitscript=True)

        if type(mass) is not str and mass is not None:
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
                unmodeling_factory.IMPORT_DESCRIPTION,
                number,
                mass,
                note,
                block,
                section,
                floor
            )
        )

        row += 1

    # при каждом повторе расчета удаляем старые версии
    unmodeling_factory.remove_models(unmodeling_factory.IMPORT_DESCRIPTION)

    with revit.Transaction("BIM: Импорт немоделируемых"):
        family_symbol.Activate()

        element_location = unmodeling_factory.get_base_location()

        for element in elements_to_generate:
            element_location = unmodeling_factory.update_location(element_location)

            unmodeling_factory.create_new_position(element,
                                                   family_symbol,
                                                   unmodeling_factory.IMPORT_DESCRIPTION,
                                                   element_location)

    # Закрываем рабочую книгу без сохранения изменений
    workbook.Close(SaveChanges=False)

    # Закрываем приложение Excel
    exel.Quit()

script_execute()