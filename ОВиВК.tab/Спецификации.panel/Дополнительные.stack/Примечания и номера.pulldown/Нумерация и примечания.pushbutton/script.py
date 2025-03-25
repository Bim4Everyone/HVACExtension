#! /usr/bin/env python
# -*- coding: utf-8 -*-
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from table_class_library import *
from pyrevit import EXEC_PARAMS

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    document = __revit__.ActiveUIDocument.Document  # type: Document
    view = document.ActiveView
    specification_filler = SpecificationFiller(document, view)

    specification_filler.fill_position_and_notes(fill_numbers=True, fill_areas=True)

script_execute()
