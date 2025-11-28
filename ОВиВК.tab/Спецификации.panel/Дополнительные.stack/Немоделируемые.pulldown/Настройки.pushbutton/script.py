#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr
import re
import System
from collections import defaultdict

from unmodeling_class_library import UnmodelingFactory, MaterialCalculator, RowOfSpecification

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")

from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, FilteredElementCollector

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from pyrevit import forms
from pyrevit import script
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from unmodeling_class_library import *
from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
unmodeling_factory = UnmodelingFactory(doc)

class SettingsWindow(forms.WPFWindow):
    """Simple window with tabs and tables."""

    def __init__(self, xaml_path):
        super(SettingsWindow, self).__init__(xaml_path)
        self.ok_button = self.FindName("OkButton")
        self.cancel_button = self.FindName("CancelButton")
        self.ducts_grid = self.FindName("DuctsGrid")
        self.pipes_grid = self.FindName("PipesGrid")

        if self.ok_button:
            self.ok_button.Click += self._on_ok
        if self.cancel_button:
            self.cancel_button.Click += self._on_cancel
        if self.ducts_grid:
            self.ducts_grid.PreviewMouseLeftButtonUp += self._on_grid_checkbox_click
        if self.pipes_grid:
            self.pipes_grid.PreviewMouseLeftButtonUp += self._on_grid_checkbox_click

    def _on_ok(self, sender, args):
        self.DialogResult = True

    def _on_cancel(self, sender, args):
        self.Close()

    def set_type_rows(self, ducts, pipes):
        if self.ducts_grid is not None:
            self.ducts_grid.ItemsSource = ducts
        if self.pipes_grid is not None:
            self.pipes_grid.ItemsSource = pipes

    def _on_grid_checkbox_click(self, sender, args):
        checkbox = getattr(args, "OriginalSource", None)
        if not self._is_checkbox(checkbox):
            return
        prop_name = getattr(checkbox, "Tag", None)
        if not prop_name:
            return
        current = getattr(checkbox.DataContext, prop_name, False)
        new_value = not bool(current)
        self._apply_checkbox_to_selection(sender, prop_name, new_value, fallback_item=checkbox.DataContext)

    @staticmethod
    def _is_checkbox(obj):
        try:
            return obj is not None and obj.__class__.__name__ == "CheckBox"
        except Exception:
            return False

    def _apply_checkbox_to_selection(self, grid, prop_name, new_value, fallback_item=None):
        if grid is None:
            return
        try:
            items = list(grid.SelectedItems)
        except Exception:
            items = []
        current_item = getattr(grid, "CurrentItem", None)
        if not items and current_item:
            items = [current_item]
        if not items and fallback_item is not None:
            items = [fallback_item]
        for row in [item for item in items if item is not None]:
            setattr(row, prop_name, bool(new_value))
        try:
            grid.Items.Refresh()
        except Exception:
            pass


def _parse_settings_region(raw_text):
    """Extracts values between UNMODELING markers into a dict."""
    if not raw_text:
        return {}

    start_marker = "##UNMODELING_REGION_START##"
    finish_marker = "##UNMODELING_REGION_FINISH##"

    start_index = raw_text.find(start_marker)
    finish_index = raw_text.find(finish_marker)
    if start_index == -1 or finish_index == -1 or finish_index <= start_index:
        return {}

    region = raw_text[start_index + len(start_marker):finish_index]
    pattern = re.compile(r"##\s*([A-Z_]+)\s*#+", re.MULTILINE)
    matches = list(pattern.finditer(region))

    values = {}
    for idx, match in enumerate(matches):
        key = match.group(1)
        value_start = match.end()
        value_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(region)
        values[key] = region[value_start:value_end].strip()

    return values


class TypeRow(object):
    def __init__(self, element):
        self.Id = str(element.Id)
        self.Name = element.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        self.CalcMetal = False
        self.CalcPaintFixings = False
        self.CalcClamps = False


def get_types_by_category(category):
    return FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsElementType() \
        .ToElements()


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    unmodeling_factory.startup_checks()
    info = doc.ProjectInformation
    pipe_types = get_types_by_category(BuiltInCategory.OST_PipeCurves)
    duct_types = get_types_by_category(BuiltInCategory.OST_DuctCurves)

    xaml_path = script.get_bundle_file("settings.xaml")
    if not xaml_path:
        forms.alert("settings.xaml was not found next to the script.", exitscript=True)

    window = SettingsWindow(xaml_path)
    window.set_type_rows(
        ducts=[TypeRow(el) for el in duct_types],
        pipes=[TypeRow(el) for el in pipe_types]
    )

    def first_or_empty(lst):
        return lst[0] if lst else ""

    window.EnamelNameTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("ENAMEL_NAME"))
    window.EnamelBrandTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("ENAMEL_MARK"))
    window.EnamelCodeTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("ENAMEL_CODE"))
    window.EnamelFactoryTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("ENAMEL_CREATOR"))
    window.EnamelUnitTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("ENAMEL_UNIT"))

    window.PrimerNameTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("PRIMER_NAME"))
    window.PrimerBrandTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("PRIMER_MARK"))
    window.PrimerCodeTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("PRIMER_CODE"))
    window.PrimerFactoryTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("PRIMER_CREATOR"))
    window.PrimerUnitTextBox.Text = first_or_empty(unmodeling_factory.get_setting_region("PRIMER_UNIT"))

    result = window.ShowDialog()

    corrected_settings = ""
    if result:
        settings = unmodeling_factory.info.GetParamValueOrDefault("ФОП_ВИС_Настройки немоделируемых", "")

        # Список регионов и соответствующих значений (каждое значение — список)
        regions = [
            ("ENAMEL_NAME", [window.EnamelNameTextBox.Text]),
            ("ENAMEL_MARK", [window.EnamelBrandTextBox.Text]),
            ("ENAMEL_CODE", [window.EnamelCodeTextBox.Text]),
            ("ENAMEL_CREATOR", [window.EnamelFactoryTextBox.Text]),
            ("ENAMEL_UNIT", [window.EnamelUnitTextBox.Text]),
            ("PRIMER_NAME", [window.PrimerNameTextBox.Text]),
            ("PRIMER_MARK", [window.PrimerBrandTextBox.Text]),
            ("PRIMER_CODE", [window.PrimerCodeTextBox.Text]),
            ("PRIMER_UNIT", [window.PrimerUnitTextBox.Text]),
            ("PRIMER_CREATOR", [window.PrimerFactoryTextBox.Text]),
        ]

        # Обрабатываем каждый регион
        base_settings = unmodeling_factory.info.GetParamValue("ФОП_ВИС_Настройки немоделируемых")

        for region_name, values in regions:
            base_settings = unmodeling_factory.edit_setting_region(base_settings, region_name, values, append=False)

        corrected_settings = base_settings


        with revit.Transaction("BIM: Обновление настроек"):
            print corrected_settings
            unmodeling_factory.info.SetParamValue("ФОП_ВИС_Настройки немоделируемых", corrected_settings)



script_execute()
