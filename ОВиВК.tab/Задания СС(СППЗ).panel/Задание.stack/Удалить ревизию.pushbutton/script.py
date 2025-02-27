#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

from low_voltage_task_class_lib import JsonOperator, EditedReport, LowVoltageSystemData

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")


import dosymep
import glob
import os
import datetime
import pytz

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__.Application

operator = JsonOperator(doc, uiapp)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

    file_folder_path, is_path_local = operator.get_document_path()

    if is_path_local:
        operator.show_local_path(file_folder_path)

    # Находим все JSON-файлы в директории
    json_files = glob.glob(os.path.join(file_folder_path, "*.json"))
    if not json_files:
        forms.alert("Данные не найдены, обновите задание.", "Ошибка", exitscript=True)

    # Находим файл с самым поздним временем модификации
    latest_file = max(json_files, key=os.path.getmtime)

    date = operator.get_utc_date()

    if date not in latest_file:
        forms.alert('Данные за сегодняшнее число не существуют.',
                    "Ошибка", exitscript=True)
    else:
        report = ("Данное действие невозможно отменить. "
                  "Данные сегодняшней ревизии будут удалены и могут быть только добавлены заново. Удалить файл?")

        result = operator.show_dialog(report)

        if result:
            os.remove(latest_file)

script_execute()
