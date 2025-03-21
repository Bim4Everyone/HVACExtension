# -*- coding: utf-8 -*-
import math
import sys

import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *
import Autodesk.Revit.Exceptions
import form_class

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from dosymep.Bim4Everyone.SharedParams import *
from dosymep.Bim4Everyone import ElementExtensions

from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *
from dosymep.Revit import *
from dosymep.Revit.Geometry import *

position_param = SharedParamsConfig.Instance.VISPosition
note_param = SharedParamsConfig.Instance.VISNote
group_param = SharedParamsConfig.Instance.VISGrouping
individual_stock_param = SharedParamsConfig.Instance.VISIndividualStock
duct_stock_param = SharedParamsConfig.Instance.VISPipeDuctReserve
use_duct_fittings_param = SharedParamsConfig.Instance.VISConsiderDuctFittings

class SpecificationSettings:
    """
    Класс для настройки спецификаций в документе.

    Атрибуты:
        definition (ScheduleDefinition): Определение спецификации.
        rollback_itemized (bool): Флаг для возврата состояния развернутости спецификации.
        rollback_header (bool): Флаг для возврата состояния отображения заголовков.
        hidden (list): Список скрытых столбцов.
        sort_para_group_indexes (list): Индексы параметров сортировки.
        position_index (int): Индекс параметра позиции.
        group_index (int): Индекс параметра группы.
    """

    definition = None
    rollback_itemized = False
    rollback_header = False
    hidden = []
    sort_para_group_indexes = None
    position_index = None
    group_index = None

    def __init__(self, definition):
        """
        Инициализация экземпляра класса SpecificationSettings.

        Args:
            definition (ScheduleDefinition): Определение спецификации.
        """
        self.definition = definition

        self.position_index = self.__get_schedule_parameter_index(position_param.Name)
        self.group_index = self.__get_schedule_parameter_index(group_param.Name)
        self.sort_para_group_indexes = self.__get_sorting_params_indexes()

        # Если заголовки показаны изначально или если спека изначально развернута - сворачивать назад не нужно
        self.rollback_itemized = not definition.IsItemized
        self.rollback_header = not definition.ShowHeaders

    def show_all_specification(self):
        """
        Разворачивает спецификацию полностью.
        """
        self.definition.IsItemized = True
        self.definition.ShowHeaders = True

        # Собираем номера скрытых столбцов и разворачиваем их
        i = 0
        while i < self.definition.GetFieldCount():
            if self.definition.GetField(i).IsHidden:
                self.hidden.append(i)
            self.definition.GetField(i).IsHidden = False
            i += 1

    def repair_specification(self):
        """
        Возвращает спецификацию в исходное состояние.
        """
        self.definition.IsItemized = not self.rollback_itemized
        self.definition.ShowHeaders = not self.rollback_header

        i = 0
        while i < self.definition.GetFieldCount():
            if i in self.hidden:
                self.definition.GetField(i).IsHidden = True
            i += 1

    def __get_sorting_params_indexes(self):
        """
        Возвращает индексы параметров сортировки.

        Returns:
            list: Список индексов параметров сортировки.
        """
        sorting_params_indexes = []
        index = 0
        for field in self.definition.GetFieldOrder():
            for scheduleSortGroupField in self.definition.GetSortGroupFields():
                if scheduleSortGroupField.FieldId.ToString() == field.ToString():
                    sorting_params_indexes.append(index)

            index += 1

        return sorting_params_indexes

    def __get_schedule_parameter_index(self, name):
        """
        Возвращает индекс параметра спецификации по имени.

        Args:
            name (str): Имя параметра.

        Returns:
            int: Индекс параметра.
        """
        index = 0

        for field in self.definition.GetFieldOrder():
            field = self.definition.GetField(field)
            if field.GetName() == name:
                return index
            index += 1

        forms.alert('В таблице нет параметра {}. Выполнение невозможно.'.format(name),
                    "Ошибка", exitscript=True)

class SpecificationFiller:
    """
    Класс для заполнения спецификаций в документе.

    Атрибуты:
        doc (Document): Документ, в котором происходит заполнение спецификаций.
        active_view (View): Активный вид, используемый для заполнения спецификаций.
    """

    doc = None
    active_view = None
    duct_stock = 0  # Запас изоляции из сведений о проекте
    types_cash = {}  # Кэш данных по индивидуальным запасов из типов воздуховодов-фитингов

    mechanical_categories = [
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_FlexDuctCurves,
        BuiltInCategory.OST_FlexPipeCurves,
        BuiltInCategory.OST_DuctTerminal,
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_PipeAccessory,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_DuctInsulations,
        BuiltInCategory.OST_PipeInsulations,
        BuiltInCategory.OST_PlumbingFixtures,
        BuiltInCategory.OST_Sprinklers,
        BuiltInCategory.OST_CableTray
    ]

    def __init__(self, doc, active_view):
        """
        Инициализация экземпляра класса SpecificationFiller.

        Args:
            doc (Document): Документ, в котором происходит заполнение спецификаций.
            active_view (View): Активный вид, используемый для заполнения спецификаций.
        """
        self.doc = doc
        self.active_view = active_view
        info = self.doc.ProjectInformation
        self.duct_stock = float(info.GetParamValueOrDefault(duct_stock_param))

    def __get_sort_rule_string(self, row, specification_settings, vs):
        """
        Возвращает значение данных по которым идет сортировка слитыми в единую строку.

        Args:
            row (int): Номер строки.
            specification_settings (SpecificationSettings): Настройки спецификации.
            vs (ViewSchedule): Активный вид спецификации.

        Returns:
            str: Строка, содержащая значения данных для сортировки.
        """
        sort_texts = [vs.GetCellText(SectionType.Body, row, ind) for ind in
                      specification_settings.sort_para_group_indexes]

        result = ''.join(sort_texts)

        return result

    def __process_row(self, row, specification_settings, old_schedule_string, position_number):
        """
        Обрабатывает строку спецификации.

        Args:
            row (int): Номер строки.
            specification_settings (SpecificationSettings): Настройки спецификации.
            old_schedule_string (str): Старая строка сортировки.
            position_number (int): Текущий номер позиции.

        Returns:
            tuple: Новая строка сортировки и обновленный номер позиции.
        """
        new_sort_rule = self.__get_sort_rule_string(row, specification_settings, self.active_view)
        element_id = self.active_view.GetCellText(SectionType.Body, row, specification_settings.position_index)
        if element_id.isdigit():
            group = self.active_view.GetCellText(SectionType.Body, row, specification_settings.group_index)
            element = self.doc.GetElement(ElementId(int(element_id)))

            not_manifold = '_Узел_' not in group
            spec_by_instance = not specification_settings.rollback_itemized

            if spec_by_instance and not_manifold:
                position_number += 1
            elif new_sort_rule != old_schedule_string and not_manifold:
                position_number += 1

            if not_manifold:
                self.__set_if_not_ro(element, position_param, str(position_number))
            else:
                self.__set_if_not_ro(element, position_param, '')

        return new_sort_rule, position_number

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

    def __check_edited_elements(self, elements):
        """
        Проверяет, заняты ли элементы другими пользователями.

        Args:
            elements (list): Список элементов для проверки.
        """
        edited_reports = []
        status_report = ''
        edited_report = ''

        for element in elements:
            update_status = WorksharingUtils.GetModelUpdatesStatus(self.doc, element.Id)

            if update_status == ModelUpdatesStatus.UpdatedInCentral:
                status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию. "

            name = self.__get_element_editor_name(element)
            if name is not None and name not in edited_reports:
                edited_reports.append(name)
        if len(edited_reports) > 0:
            edited_report = "Часть элементов спецификации занята пользователями: {}".format(", ".join(edited_reports))
        if edited_report != '' or status_report != '':
            report_message = status_report + ('\n' if (edited_report and status_report) else '') + edited_report
            forms.alert(report_message, "Ошибка", exitscript=True)

    def __get_connectors(self, element):
        connectors = []

        if isinstance(element, FamilyInstance) and element.MEPModel.ConnectorManager is not None:
            connectors.extend(element.MEPModel.ConnectorManager.Connectors)

        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves]) and \
                isinstance(element, MEPCurve) and element.ConnectorManager is not None:
            connectors.extend(element.ConnectorManager.Connectors)

        return connectors

    def __get_fitting_area(self, element):
        area = 0

        for solid in dosymep.Revit.Geometry.ElementExtensions.GetSolids(element):
            for face in solid.Faces:
                area += face.Area

        area = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters)

        if area > 0:
            false_area = 0
            connectors = self.__get_connectors(element)
            for connector in connectors:
                if connector.Shape == ConnectorProfileType.Rectangular:
                    false_area += UnitUtils.ConvertFromInternalUnits(
                        connector.Height * connector.Width, UnitTypeId.SquareMeters)
                if connector.Shape == ConnectorProfileType.Round:
                    false_area += UnitUtils.ConvertFromInternalUnits(
                        connector.Radius * connector.Radius * math.pi, UnitTypeId.SquareMeters)
                if connector.Shape == ConnectorProfileType.Oval:
                    false_area += 0

            # Вычитаем площадь пустоты на местах коннекторов
            area -= false_area

        return area

    def __get_duct_area(self, duct):
        return UnitUtils.ConvertFromInternalUnits(
            duct.GetParamValueOrDefault(BuiltInParameter.RBS_CURVE_SURFACE_AREA),
            UnitTypeId.SquareMeters)

    def __process_areas(self):
        """
        Обрабатывает площади воздуховодов, их фитингов, и обновляет параметр VISNote.
        """
        # Заполняем площади фитингов, если включен их учет. Иначе - металл и так идет в м2
        info = self.doc.ProjectInformation
        fill_fitting_areas = info.GetParamValueOrDefault(use_duct_fittings_param) == 1

        area_elements = []

        area_elements.extend(FilteredElementCollector(
            self.doc,
            self.active_view.Id).OfCategory(BuiltInCategory.OST_DuctCurves).ToElements())

        if fill_fitting_areas:
            duct_fittings = FilteredElementCollector(
                self.doc,
                self.active_view.Id).OfCategory(BuiltInCategory.OST_DuctFitting).ToElements()

            area_elements.extend(duct_fittings)

        duct_dict = {}

        for area_element in area_elements:
            element_position = area_element.GetParamValue(position_param)
            if area_element.Category.IsId(BuiltInCategory.OST_DuctCurves):
                element_area = self.__get_duct_area(area_element)
            else:
                element_area = self.__get_fitting_area(area_element)

            if element_position in duct_dict:
                duct_dict[element_position] += element_area
            else:
                duct_dict[element_position] = element_area

        for area_element in area_elements:
            element_position = area_element.GetParamValue(position_param)

            value = duct_dict[element_position]
            formated_value = self.__format_area_value(area_element, value)

            self.__set_if_not_ro(area_element, note_param, formated_value)

    def __format_area_value(self, element, value):
        # Если у нас заполняется примечание, проверяем индивидуальный запас и общий запас на воздуховоды. Увеличиваем
        # значение на него и форматируем под м2

        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_DuctFitting]):
            element_type = element.GetElementType()

            # Проверяем, существует ли уже айди типа в кэше
            if element_type.Id in self.types_cash:
                individual_stock = self.types_cash[element_type.Id]
            else:
                individual_stock = element_type.GetParamValueOrDefault(
                    individual_stock_param)
                self.types_cash[element_type.Id] = individual_stock

            if ((individual_stock == 0 or individual_stock is None)
                    and (self.duct_stock != 0 and self.duct_stock is not None)):
                value = value + value * (self.duct_stock / 100)

            if individual_stock != 0 and individual_stock is not None:
                value = value + value * (individual_stock / 100)

            value = "{:.2f}".format(value).rstrip('0').rstrip('.') + ' м²'

        return value

    def __set_if_not_ro(self, element, shared_param, value):
        """
        Заполняет значение параметра, если он не ридонли.

        Args:
            element (Element): Элемент спецификации
            shared_param: RevitParam с платформы
            value: Устанавливаемое значение
        """
        if not element.IsExistsParam(shared_param):
            return

        param = element.GetParam(shared_param)
        if not param.IsReadOnly:
            element.SetParamValue(shared_param, value)

    def __fill_id_to_schedule_param(self, specification_settings, elements):
        """
        Заполняет параметр позиции айди элементов.

        Args:
            specification_settings (SpecificationSettings): Настройки спецификации.
            elements (list): Список элементов для заполнения.
        """
        with revit.Transaction("BIM: Запись айди"):
            specification_settings.show_all_specification()

            for element in elements:
                self.__set_if_not_ro(element, position_param, str(element.Id.IntegerValue))

    def __fill_values(self, specification_settings, elements, fill_areas, fill_numbers, first_index):
        """
        Заполняет позицию и примечания для элементов.

        Args:
            specification_settings (SpecificationSettings): Настройки спецификации.
            elements (list): Список элементов для заполнения.
            fill_areas (bool): Флаг для заполнения площадей.
            fill_numbers (bool): Флаг для заполнения номеров.
        """
        with revit.Transaction("BIM: Запись номера"):
            section_data = self.active_view.GetTableData().GetSectionData(SectionType.Body)
            row_number = section_data.FirstRowNumber

            position_number = first_index - 1 # Вычитать единицу нужно для первой строки.
            # Скрипт сразу увеличивает значение на 1, т.к. предыдущая строка данных не имеет значения

            old_sort_rule = ''
            while row_number <= section_data.LastRowNumber:
                old_sort_rule, position_number = self.__process_row(row_number,
                                                                    specification_settings,
                                                                    old_sort_rule,
                                                                    position_number)
                row_number += 1

            specification_settings.repair_specification()

            if fill_areas:
                self.__process_areas()

            if not fill_numbers:
                for element in elements:
                    self.__set_if_not_ro(element, position_param, '')

    def __check_param(self, elements, revit_param):
        """ Проверяем все элементы из списка на наличие параметра

        Args:
            elements: Список элементов
            revit_param: Параметр для проверки
        Returns:
            Имя параметра или None если все в норме
        """
        for element in elements:
            if not element.InAnyCategory(self.mechanical_categories):
                continue

            if not element.IsExistsSharedParam(revit_param.Name):
                return revit_param.Name
        return None

    def __check_params_instance(self, elements):
        """
        Проверяем наличие параметров в экземпляре. Это может быть сделано специально, поэтому не падаем

        Args:
            elements: Все элементы с активного вида
        """
        position_report = self.__check_param(elements, position_param)
        note_report = self.__check_param(elements, note_param)

        results = [result for result in [position_report, note_report] if result is not None]

        if results:
            forms.alert(
                'Параметры {} найдены не у всех элементов спецификации.'.format(', '.join(results)),
                "Внимание")

    def __setup_params(self):
        revit_params = [note_param,
                        position_param,
                        duct_stock_param,
                        individual_stock_param,
                        use_duct_fittings_param,
                        ]

        project_parameters = ProjectParameters.Create(self.doc.Application)
        project_parameters.SetupRevitParams(self.doc, revit_params)

    def __get_first_index(self, fill_numbers):
        """
        Получение начала нумерации. Если окошко было закрыто - прерываем скрипт.

        Args:
            fill_numbers: Флаг отвечающий за вызов нумерации
        """
        first_index = 1
        if fill_numbers:
            first_index = forms.ask_for_string(
                default='1',
                prompt='С какого числа стартует нумерация:',
                title="Нумерация"
            )

            if first_index is None or not first_index.isdigit():
                forms.alert(
                    "Нужно ввести число.",
                    "Ошибка",
                    exitscript=True)

        return int(first_index)

    def __statup_checks(self):
        if self.doc.IsFamilyDocument:
            forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

        if self.active_view.Category is None or not self.active_view.Category.IsId(BuiltInCategory.OST_Schedules):
            forms.alert(
                "Надстройка предназначена только для работы на активном виде целевой спецификации",
                "Ошибка", exitscript=True)

        # На всякий случай выполняем настройку параметров - в теории уже должны быть на месте, но лучше продублировать
        self.__setup_params()

        elements = FilteredElementCollector(self.doc, self.doc.ActiveView.Id)

        # Проверяем параметры. Они могли быть добавлены ранее или сидеть в семействах в типе
        self.__check_params_instance(elements)

        # Если хоть один элемент спеки на редактировании - отменяем выполнение, нужно освобождать
        self.__check_edited_elements(elements)

        # Получаем правила по которым работает спека
        specification_settings = SpecificationSettings(self.active_view.Definition)

        return elements, specification_settings

    def __get_table_params(self, definition):
        paramList = []
        for scheduleGroupField in definition.GetFieldOrder():
            scheduleField = definition.GetField(scheduleGroupField)
            paramList.append(scheduleField.GetName())
        return paramList

    def fill_position_and_notes(self, fill_numbers=False, fill_areas=False):
        """
        Основной метод для заполнения позиций и примечаний в спецификации.

        Args:
            fill_numbers (bool): Флаг для заполнения номеров
            fill_areas (bool): Флаг для заполнения площадей
        """

        elements, specification_settings = self.__statup_checks()

        # Выясняем с какого числа стартует нумерация
        first_index = self.__get_first_index(fill_numbers)

        # Заполняем айди в параметр позиции элементов для их чтения
        self.__fill_id_to_schedule_param(specification_settings, elements)

        # заполняем значения нумерации и, для воздуховодов их фитингов, примечаний
        self.__fill_values(specification_settings, elements, fill_areas, fill_numbers, first_index)

    def replace_parameter_values(self):
        elements, specification_settings = self.__statup_checks()

        table_params = self.__get_table_params(specification_settings.definition)
        # print(table_params)

        form = form_class.SelectParametersForm(table_params)

        selected_param_name_in, selected_param_name_out = form.show_form()

        # selected_option1, selected_option2 = form.show_form()
