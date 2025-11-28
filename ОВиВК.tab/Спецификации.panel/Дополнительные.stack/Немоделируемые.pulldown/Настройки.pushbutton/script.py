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

    for ins in ins_rows:
        base_id = ins.Id

        for idx, cons in enumerate(ins.Consumables, start=1):
            prefix = "CONSUMABLE{0}".format(idx)

            regions[prefix + "_NAMES"].append("{0} - {1}".format(base_id, cons.Name))
            regions[prefix + "_MARKS"].append("{0} - {1}".format(base_id, cons.Mark))
            regions[prefix + "_CODES"].append("{0} - {1}".format(base_id, cons.Code))
            regions[prefix + "_MAKER"].append("{0} - {1}".format(base_id, cons.Factory))
            regions[prefix + "_UNIT"].append("{0} - {1}".format(base_id, cons.Unit))
            regions[prefix + "_RATE"].append("{0} - {1}".format(base_id, cons.RatePerMeter))
            regions[prefix + "_RATE_BY_SQUARE"].append("{0} - {1}".format(base_id, cons.RatePerSqm))

    return regions


def _apply_consumable_region(rows, region_entries, attr_name, is_bool=False):
    truthy = ("true", "1", "yes", u"да", u"истина")
    values = {}
    for entry in region_entries:
        if not entry:
            continue
        normalized = entry.strip()
        # поддерживаем форматы "Id - value", "Id value", "Id\tvalue"
        if " - " in normalized:
            elem_id, value = [p.strip() for p in normalized.split(" - ", 1)]
        elif "-" in normalized:
            elem_id, value = [p.strip() for p in normalized.split("-", 1)]
        else:
            first_split = normalized.split(None, 1)
            elem_id = first_split[0].strip() if first_split else ""
            value = first_split[1].strip() if len(first_split) > 1 else ""

        if not elem_id:
            continue
        values[elem_id] = value
    for row in rows:
        val = values.get(row.Id)
        if val is None:
            continue
        if isinstance(val, basestring) and not val.strip():
            continue
        if is_bool:
            setattr(row, attr_name, val.strip().lower() in truthy)
        else:
            setattr(row, attr_name, val.strip())


def _ensure_regions(settings, region_names):
    finish_marker = "##UNMODELING_REGION_FINISH##"
    finish_index = settings.find(finish_marker)
    if finish_index == -1:
        return settings
    missing = []
    for name in region_names:
        tag = "##{0}##".format(name)
        if tag not in settings:
            missing.append(tag + "\n")
    if not missing:
        return settings
    insert = "".join(missing)
    return settings[:finish_index] + insert + settings[finish_index:]


def get_types_by_category(category):
    return FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsElementType() \
        .ToElements()


def script_execute():
    unmodeling_factory.startup_checks()


    pipe_types = get_types_by_category(BuiltInCategory.OST_PipeCurves)
    duct_types = get_types_by_category(BuiltInCategory.OST_DuctCurves)
    pipe_insulation_types = get_types_by_category(BuiltInCategory.OST_PipeInsulations)
    duct_insulation_types = get_types_by_category(BuiltInCategory.OST_DuctInsulations)

    duct_rows = [TypeRow(el) for el in duct_types]
    pipe_rows = [TypeRow(el) for el in pipe_types]
    pipe_ins_rows = [InsulationTypeRow(el) for el in pipe_insulation_types]
    duct_ins_rows = [InsulationTypeRow(el) for el in duct_insulation_types]

    xaml_path = script.get_bundle_file("settings.xaml")
    if not xaml_path:
        forms.alert("settings.xaml was not found next to the script.", exitscript=True)

    window = SettingsWindow(xaml_path)

    _apply_region_to_rows(duct_rows, _parse_bool_region(unmodeling_factory.get_setting_region("DUCT_TYPES_METALL")), "CalcMetal")
    _apply_region_to_rows(pipe_rows, _parse_bool_region(unmodeling_factory.get_setting_region("PIPE_TYPES_METALL")), "CalcMetal")
    _apply_region_to_rows(pipe_rows, _parse_bool_region(unmodeling_factory.get_setting_region("PIPE_TYPES_COLOR")), "CalcPaintFixings")
    _apply_region_to_rows(pipe_rows, _parse_bool_region(unmodeling_factory.get_setting_region("PIPE_TYPES_CLAMPS")), "CalcClamps")

    all_ins_rows = pipe_ins_rows + duct_ins_rows
    for idx in range(1, 5):
        _apply_consumable_region(all_ins_rows, unmodeling_factory.get_setting_region("CONSUMABLE{0}_NAMES".format(idx)), "Name")
        _apply_consumable_region(all_ins_rows, unmodeling_factory.get_setting_region("CONSUMABLE{0}_MARKS".format(idx)), "Mark")
        _apply_consumable_region(all_ins_rows, unmodeling_factory.get_setting_region("CONSUMABLE{0}_CODES".format(idx)), "Code")
        _apply_consumable_region(all_ins_rows, unmodeling_factory.get_setting_region("CONSUMABLE{0}_MAKER".format(idx)), "Factory")
        _apply_consumable_region(all_ins_rows, unmodeling_factory.get_setting_region("CONSUMABLE{0}_UNIT".format(idx)), "Unit")
        _apply_consumable_region(all_ins_rows, unmodeling_factory.get_setting_region("CONSUMABLE{0}_RATE".format(idx)), "RatePerMeter")
        _apply_consumable_region(all_ins_rows, unmodeling_factory.get_setting_region("CONSUMABLE{0}_RATE_BY_SQUARE".format(idx)), "RatePerSqm", is_bool=True)

    window.set_type_rows(ducts=duct_rows, pipes=pipe_rows)
    window.set_insulation_rows(pipe_ins_rows, duct_ins_rows)

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
        duct_region_values = _region_values_from_rows(duct_rows, "CalcMetal")
        pipe_region_metall = _region_values_from_rows(pipe_rows, "CalcMetal")
        pipe_region_color = _region_values_from_rows(pipe_rows, "CalcPaintFixings")
        pipe_region_clamps = _region_values_from_rows(pipe_rows, "CalcClamps")

        consumable_regions = _collect_consumable_regions(pipe_ins_rows + duct_ins_rows)

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
            ("DUCT_TYPES_METALL", duct_region_values),
            ("PIPE_TYPES_METALL", pipe_region_metall),
            ("PIPE_TYPES_COLOR", pipe_region_color),
            ("PIPE_TYPES_CLAMPS", pipe_region_clamps),
        ]

        regions.extend([(region_name, values) for region_name, values in consumable_regions.items()])
        #print regions

        print "Изначальные настройки"
        base_settings = unmodeling_factory.info.GetParamValueOrDefault(u"ФОП_ВИС_Настройки немоделируемых", "")
        print base_settings
        print "Конец изначальных настроек"
        #base_settings = _ensure_regions(base_settings, [r[0] for r in regions] + consumable_region_names)

        for region_name, values in regions:
            base_settings = unmodeling_factory.edit_setting_region(base_settings, region_name, values, append=False)

        corrected_settings = base_settings

        print corrected_settings

        with revit.Transaction("BIM: запись настроек"):
            unmodeling_factory.info.SetParamValue(u"ФОП_ВИС_Настройки немоделируемых", corrected_settings)


script_execute()
