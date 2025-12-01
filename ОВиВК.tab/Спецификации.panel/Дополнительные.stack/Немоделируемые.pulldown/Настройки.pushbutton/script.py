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


class TypeRow(object):
    def __init__(self, element):
        self.Id = str(element.Id)
        self.Name = element.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        self.CalcMetal = False
        self.CalcPaintFixings = False
        self.CalcClamps = False


class ConsumableRow(object):
    def __init__(self, index):
        self.Index = index
        self.Name = ""
        self.Mark = ""
        self.Code = ""
        self.Article = ""
        self.Factory = ""
        self.Unit = ""
        self.RatePerMeter = ""
        self.RatePerSqm = False


class InsulationTypeRow(object):
    def __init__(self, element):
        self.Id = str(element.Id)
        self.TypeName = element.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        self.Consumables = [ConsumableRow(i) for i in range(1, 5)]


def _parse_bool_region(entries):
    truthy = ("true")
    result = set()
    for entry in entries:
        if not entry:
            continue
        parts = [p.strip() for p in entry.split("-", 1)]
        if len(parts) != 2:
            continue
        elem_id, value = parts
        if value.lower() in truthy:
            result.add(elem_id)
    return result


def _apply_region_to_rows(rows, true_ids, attr_name):
    for row in rows:
        setattr(row, attr_name, row.Id in true_ids)


def _region_values_from_rows(rows, attr_name):
    values = ["{0} - True".format(r.Id) for r in rows if getattr(r, attr_name, False)]
    return values if values else [""]


def _collect_consumable_regions(ins_rows):
    regions = defaultdict(list)

    def _has_value(val):
        try:
            return bool(val) and bool(val.strip())
        except Exception:
            return bool(val)

    for ins in ins_rows:
        base_id = ins.Id

        for idx, cons in enumerate(ins.Consumables, start=1):
            prefix = "CONSUMABLE{0}".format(idx)

            if _has_value(cons.Name):
                regions[prefix + "_NAMES"].append("{0} - {1}".format(base_id, cons.Name))
            if _has_value(cons.Mark):
                regions[prefix + "_MARKS"].append("{0} - {1}".format(base_id, cons.Mark))
            if _has_value(cons.Code):
                regions[prefix + "_CODES"].append("{0} - {1}".format(base_id, cons.Code))
            if _has_value(cons.Factory):
                regions[prefix + "_MAKER"].append("{0} - {1}".format(base_id, cons.Factory))
            if _has_value(cons.Unit):
                regions[prefix + "_UNIT"].append("{0} - {1}".format(base_id, cons.Unit))
            if _has_value(cons.RatePerMeter):
                regions[prefix + "_RATE"].append("{0} - {1}".format(base_id, cons.RatePerMeter))
            if cons.RatePerSqm:
                regions[prefix + "_RATE_BY_SQUARE"].append("{0} - {1}".format(base_id, cons.RatePerSqm))

    for idx in range(1, 5):
        prefix = "CONSUMABLE{0}".format(idx)
        for suffix in ("_NAMES", "_MARKS", "_CODES", "_MAKER", "_UNIT", "_RATE", "_RATE_BY_SQUARE"):
            key = prefix + suffix
            if not regions.get(key):
                regions[key] = [""]

    return regions


def _apply_consumable_region(rows, region_entries, attr_name, cons_index, is_bool=False):
    truthy = ("true", "1", "yes", u"да", u"истина")
    values = {}
    for entry in region_entries:
        if not entry:
            continue
        normalized = entry.strip()
        if " - " in normalized:
            elem_id, value = [p.strip() for p in normalized.split(" - ", 1)]
        elif "-" in normalized:
            elem_id, value = [p.strip() for p in normalized.split("-", 1)]
        else:
            parts = normalized.split(None, 1)
            elem_id = parts[0].strip() if parts else ""
            value = parts[1].strip() if len(parts) > 1 else ""
        if not elem_id:
            continue
        if isinstance(value, basestring) and not value.strip():
            continue
        values[elem_id] = value

    for row in rows:
        val = values.get(row.Id)
        if val is None:
            continue
        if isinstance(val, basestring) and not val.strip():
            continue
        if cons_index < 1 or cons_index > len(row.Consumables):
            continue
        target = row.Consumables[cons_index - 1]
        if is_bool:
            setattr(target, attr_name, val.strip().lower() in truthy)
        else:
            setattr(target, attr_name, val.strip())


def get_types_by_category(category):
    return FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsElementType() \
        .ToElements()


def script_execute():
    unmodeling_factory.startup_checks()

    xaml_path = script.get_bundle_file("settings.xaml")
    if not xaml_path:
        forms.alert("settings.xaml was not found next to the script.", exitscript=True)

    window = SettingsWindow(xaml_path)

    print unmodeling_factory.info.GetParamValueOrDefault("ФОП_ВИС_Настройки немоделируемых", "")

    window.EnamelNameTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "ENAMEL", "NAME"])
    window.EnamelBrandTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "ENAMEL", "MARK"])
    window.EnamelCodeTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "ENAMEL", "CODE"])
    window.EnamelFactoryTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "ENAMEL", "CREATOR"])
    window.EnamelUnitTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "ENAMEL", "UNIT"])

    window.PrimerNameTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "PRIMER", "NAME"])
    window.PrimerBrandTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "PRIMER", "MARK"])
    window.PrimerCodeTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "PRIMER", "CODE"])
    window.PrimerFactoryTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "PRIMER", "CREATOR"])
    window.PrimerUnitTextBox.Text = unmodeling_factory.get_setting_value(
        ["UNMODELING", "PRIMER", "UNIT"])

    result = window.ShowDialog()
    
    corrected_settings = ""
    if result:
        setting_values = [
            (["UNMODELING", "ENAMEL", "NAME"], window.EnamelNameTextBox.Text),
            (["UNMODELING", "ENAMEL", "MARK"], window.EnamelBrandTextBox.Text),
            (["UNMODELING", "ENAMEL", "CODE"], window.EnamelCodeTextBox.Text),
            (["UNMODELING", "ENAMEL", "CREATOR"], window.EnamelFactoryTextBox.Text),
            (["UNMODELING", "ENAMEL", "UNIT"], window.EnamelUnitTextBox.Text),

            (["UNMODELING", "PRIMER", "NAME"], window.PrimerNameTextBox.Text),
            (["UNMODELING", "PRIMER", "MARK"], window.PrimerBrandTextBox.Text),
            (["UNMODELING", "PRIMER", "CODE"], window.PrimerCodeTextBox.Text),
            (["UNMODELING", "PRIMER", "UNIT"], window.PrimerUnitTextBox.Text),
            (["UNMODELING", "PRIMER", "CREATOR"], window.PrimerFactoryTextBox.Text),
        ]

        base_settings = unmodeling_factory.info.GetParamValueOrDefault("ФОП_ВИС_Настройки немоделируемых", "")

        for setting_key, values in setting_values:
            base_settings = unmodeling_factory.set_setting_value(base_settings, setting_key, values)

        corrected_settings = base_settings

        with revit.Transaction("BIM: Обновление настроек немоделируемых"):
            unmodeling_factory.info.SetParamValue("ФОП_ВИС_Настройки немоделируемых", corrected_settings)


script_execute()
