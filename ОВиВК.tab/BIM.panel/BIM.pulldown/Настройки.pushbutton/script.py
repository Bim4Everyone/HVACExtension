#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import clr
import re
import System
from collections import defaultdict


def _add_unmodeling_lib_path():
    """Add unmodeling library folder to sys.path (relative to this script)."""
    try:
        from System.IO import Path  # type: ignore

        script_dir = Path.GetDirectoryName(__file__)
        tab_dir = Path.GetFullPath(Path.Combine(script_dir, "..", "..", ".."))
        lib_dir = Path.GetFullPath(
            Path.Combine(tab_dir, u"Спецификации.panel", u"Дополнительные.stack", "lib")
        )
    except Exception:
        script_dir = os.path.dirname(__file__)
        tab_dir = os.path.abspath(os.path.join(script_dir, "..", "..", ".."))
        lib_dir = os.path.abspath(
            os.path.join(tab_dir, u"Спецификации.panel", u"Дополнительные.stack", "lib")
        )

    if lib_dir and lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)


_add_unmodeling_lib_path()

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

    def set_insulation_rows(self, pipe_rows, duct_rows):
        pipe_ins_items = self.FindName("PipeInsulationItems")
        duct_ins_items = self.FindName("DuctInsulationItems")
        if pipe_ins_items is not None:
            pipe_ins_items.ItemsSource = pipe_rows
        if duct_ins_items is not None:
            duct_ins_items.ItemsSource = duct_rows

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


def _set_window_values_from_settings(window):
    def _safe_set_text(box, value):
        if box is not None:
            box.Text = value if value is not None else ""

    _safe_set_text(
        getattr(window, "EnamelNameTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "ENAMEL", "NAME"]),
    )
    _safe_set_text(
        getattr(window, "EnamelBrandTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "ENAMEL", "MARK"]),
    )
    _safe_set_text(
        getattr(window, "EnamelCodeTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "ENAMEL", "CODE"]),
    )
    _safe_set_text(
        getattr(window, "EnamelFactoryTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "ENAMEL", "CREATOR"]),
    )
    _safe_set_text(
        getattr(window, "EnamelUnitTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "ENAMEL", "UNIT"]),
    )
    _safe_set_text(
        getattr(window, "PrimerNameTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "PRIMER", "NAME"]),
    )
    _safe_set_text(
        getattr(window, "PrimerBrandTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "PRIMER", "MARK"]),
    )
    _safe_set_text(
        getattr(window, "PrimerCodeTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "PRIMER", "CODE"]),
    )
    _safe_set_text(
        getattr(window, "PrimerFactoryTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "PRIMER", "CREATOR"]),
    )
    _safe_set_text(
        getattr(window, "PrimerUnitTextBox", None),
        unmodeling_factory.get_setting_value(["UNMODELING", "PRIMER", "UNIT"]),
    )


def _get_corrected_settings_from_window(window):
    def _read_text(box):
        try:
            return box.Text
        except Exception:
            return ""

    setting_values = [
        (["UNMODELING", "ENAMEL", "NAME"], _read_text(getattr(window, "EnamelNameTextBox", None))),
        (["UNMODELING", "ENAMEL", "MARK"], _read_text(getattr(window, "EnamelBrandTextBox", None))),
        (["UNMODELING", "ENAMEL", "CODE"], _read_text(getattr(window, "EnamelCodeTextBox", None))),
        (["UNMODELING", "ENAMEL", "CREATOR"], _read_text(getattr(window, "EnamelFactoryTextBox", None))),
        (["UNMODELING", "ENAMEL", "UNIT"], _read_text(getattr(window, "EnamelUnitTextBox", None))),
        (["UNMODELING", "PRIMER", "NAME"], _read_text(getattr(window, "PrimerNameTextBox", None))),
        (["UNMODELING", "PRIMER", "MARK"], _read_text(getattr(window, "PrimerBrandTextBox", None))),
        (["UNMODELING", "PRIMER", "CODE"], _read_text(getattr(window, "PrimerCodeTextBox", None))),
        (["UNMODELING", "PRIMER", "UNIT"], _read_text(getattr(window, "PrimerUnitTextBox", None))),
        (["UNMODELING", "PRIMER", "CREATOR"], _read_text(getattr(window, "PrimerFactoryTextBox", None))),
    ]

    base_settings = doc.ProjectInformation.GetParamValueOrDefault(SharedParamsConfig.Instance.VISSettings, "")
    for setting_key, value in setting_values:
        base_settings = unmodeling_factory.set_setting_value(base_settings, setting_key, value)
    return base_settings



def script_execute():
    unmodeling_factory.startup_checks()

    xaml_path = script.get_bundle_file("settings.xaml")
    if not xaml_path:
        forms.alert("Не найдено окно настроек.", exitscript=True)

    window = SettingsWindow(xaml_path)

    _set_window_values_from_settings(window)

    result = window.ShowDialog()

    if result:
        corrected_settings = _get_corrected_settings_from_window(window)

        with revit.Transaction("BIM: Обновление настроек"):
            doc.ProjectInformation.SetParamValue(SharedParamsConfig.Instance.VISSettings, corrected_settings)



script_execute()
