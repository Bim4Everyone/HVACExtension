#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import glob
import re
import sys
import json
import os
import ctypes
import codecs
from low_voltage_task_class_lib import JsonOperator, EditedReport, LowVoltageSystemData

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
from collections import defaultdict

from dosymep_libs.bim4everyone import *
from System.Collections.Generic import List

from datetime import datetime, timedelta
from System import Environment

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

    # Получаем данные из последнего по дате редактирования файла
    old_data = operator.get_json_data(file_folder_path)

    old_data = sorted(old_data, key=lambda x: x.deletion_date if x.deletion_date else "")

    output_data = []

    # Добавляем только элементы у которых заполнена дата удаления. Заполняется при обновлении задания
    for data in old_data:
        if data.deletion_date is not None:
            output_data.append([data.deletion_date, data.creation_date, data.json_name])

    if len(output_data) == 0:
        forms.alert('Удаленные элементы в прошлые ревизии не найдены. Обновите задание.',
                    "Ошибка", exitscript=True)

    output = script.get_output()

    output.print_table(table_data=output_data,
                       title=("Удаленные элементы"),
                       columns=["Дата удаления", "Дата создания", "Имя элемента"],
                       formats=['', '', ''],
                       )

script_execute()
