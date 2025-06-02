#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from Autodesk.Revit.DB import *
from System.Collections.Generic import List
from pyrevit import revit

from dosymep_libs.bim4everyone import *

class EditorReport:
    edited_reports = []
    status_report = ''
    edited_report = ''

    def __get_element_editor_name(self, element):
        """
        Возвращает имя пользователя, занявшего элемент, или None.

        Args:
            element (Element): Элемент для проверки.

        Returns:
            str или None: Имя пользователя или None, если элемент не занят.
        """
        user_name = __revit__.Application.Username
        edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
        if edited_by is None:
            return None

        if edited_by.lower() in user_name.lower():
            return None
        return edited_by

    def is_element_edited(self, element):
        """
        Проверяет, заняты ли элементы другими пользователями.

        Args:
            element: Элемент для проверки.
        """

        self.update_status = WorksharingUtils.GetModelUpdatesStatus(doc, element.Id)

        if self.update_status == ModelUpdatesStatus.UpdatedInCentral:
            self.status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию. "

        name = self.__get_element_editor_name(element)
        if name is not None and name not in self.edited_reports:
            self.edited_reports.append(name)
            return True

    def show_report(self):
        if len(self.edited_reports) > 0:
            self.edited_report = ("Часть элементов спецификации занята пользователями: {}"
                                  .format(", ".join(self.edited_reports)))
        if self.edited_report != '' or self.status_report != '':
            report_message =(self.status_report +
                             ('\n' if (self.edited_report and self.status_report) else '') + self.edited_report)
            forms.alert(report_message, "Ошибка", exitscript=True)

doc = __revit__.ActiveUIDocument.Document  # type: Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

def get_insulation_elements():
    """
    Забираем список элементов изоляции

    Returns:
        List(Elements)
    """
    categories = [
        BuiltInCategory.OST_DuctInsulations,
        BuiltInCategory.OST_PipeInsulations
    ]

    category_ids = List[ElementId]([ElementId(int(category)) for category in categories])

    multicategory_filter = ElementMulticategoryFilter(category_ids)

    elements = FilteredElementCollector(doc) \
        .WherePasses(multicategory_filter) \
        .WhereElementIsNotElementType() \
        .ToElements()

    return elements

def plural_element_form(n):
    n = abs(n)
    if n % 10 == 1 and n % 100 != 11:
        return "элемент"
    elif 2 <= n % 10 <= 4 and not (12 <= n % 100 <= 14):
        return "элемента"
    else:
        return "элементов"

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    elements = get_insulation_elements()
    editor_report = EditorReport()

    insulation_to_delete = []
    for element in elements:
        host_element_id = element.HostElementId

        if host_element_id is not None:
            host_element = doc.GetElement(host_element_id)

            if not host_element:
                continue

            if host_element.InAnyCategory([BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_DuctAccessory]):
                if not editor_report.is_element_edited(element):
                    insulation_to_delete.append(element)

    insulation_number = len(insulation_to_delete)

    # Нужно вынести удаление в отдельный цикл от проверки занятости,
    # иначе при необходимости синхрона будет сбрасывать транзакцию с сообщением о устаревшей версии
    with revit.Transaction("BIM: Удаление изоляции"):
        for element in insulation_to_delete:
            doc.Delete(element.Id)

    word = plural_element_form(insulation_number)
    forms.alert("Было удалено {} {}".format(insulation_number, word), "Очистка изоляции")

    editor_report.show_report()


script_execute()





