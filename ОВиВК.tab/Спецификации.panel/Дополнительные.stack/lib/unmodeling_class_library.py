# -*- coding: utf-8 -*-
import math

from System import Environment
import os
import clr
import re

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *

from System.Collections.Generic import List
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


class ElementStocks:
    """ Класс для вычисления запасов на элемент """
    types_cash = {}
    pipe_insulation_stock = None
    duct_insulation_stock = None

    def __init__(self, doc):
        info = doc.ProjectInformation
        pipe_insulatin_stock_param = SharedParamsConfig.Instance.VISPipeInsulationReserve
        duct_insulation_stock_param = SharedParamsConfig.Instance.VISDuctInsulationReserve
        pipe_and_duct_stock_param = SharedParamsConfig.Instance.VISPipeDuctReserve

        self.pipe_insulation_stock = self.get_stock_value(info
                                      .GetParamValueOrDefault(pipe_insulatin_stock_param))

        self.duct_insulation_stock = self.get_stock_value(info
                                      .GetParamValueOrDefault(duct_insulation_stock_param))

        self.duct_and_pipe_stock = self.get_stock_value(info
                                      .GetParamValueOrDefault(pipe_and_duct_stock_param))

    def get_stock_value(self, param):
        # На питоне не получается задать дефолтное значение 0, вылетает "Заданное приведение не является допустимым".
        # Чтоб не ломаться о None - приводим к нулю руками
        if param is None:
            param = 0

        return param/100

    def get_individual_stock(self, element):
        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves,
                                  BuiltInCategory.OST_PipeCurves,
                                  BuiltInCategory.OST_DuctInsulations,
                                  BuiltInCategory.OST_PipeInsulations]):
            element_type = element.GetElementType()

            # Проверяем, существует ли уже айди типа в кэше
            if element_type.Id in self.types_cash:
                individual_stock = self.types_cash[element_type.Id]
            else:
                individual_stock = self.get_stock_value(element_type.GetParamValueOrDefault(
                    SharedParamsConfig.Instance.VISIndividualStock, 0.0))
                self.types_cash[element_type.Id] = individual_stock

            if individual_stock is None:
                return 0

            return individual_stock

    def get_stock(self, element):
        individual_stock = self.get_individual_stock(element)

        if individual_stock != 0 and individual_stock is not None:
            return 1 + individual_stock
        if element.Category.IsId(BuiltInCategory.OST_PipeInsulations):
            return 1 + self.pipe_insulation_stock
        if element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
            return 1 + self.duct_insulation_stock
        if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
            return 1 + self.duct_and_pipe_stock
        if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
            return 1 + self.duct_and_pipe_stock

        return 0

class CalculationResult:
    """ Класс для передачи результата расчетов материалов """
    def __init__(self, number, area):
        self.number = number
        self.area = area

class GenerationRuleSet:
    """
    Класс правило для генерации элементов, содержит имя метода, категорию и описание материала
    """

    def __init__(self, group, name, mark, code, maker, unit, method_name, category):
        """
        Инициализация класса GenerationRuleSet.

        Args:
            group: Имя группирования для спецификации
            name: Имя расходника
            mark: Марка расходника
            maker: Изготовитель расходника
            unit: Единица измерения расходника
            category: BuiltInCategory для расчета
            code: Код изделия расходника(обычно пустует)
            method_name: Имя метода по которому будем выбирать расчет

        """
        self.group = group
        self.name = name
        self.mark = mark
        self.maker = maker
        self.unit = unit
        self.category = category
        self.method_name = method_name
        self.code = code

class MaterialVariants:
    """
    Класс содержащий расчетные данные для создаваемых расходников
    """
    def __init__(self, diameter, insulated_rate, not_insulated_rate):
        """
        Инициализация класса MaterialVariants

        Args:
            diameter: диаметр линейного элемента под который идет расчет
            insulated_rate: Расход материала на изолированный элемент
            not_insulated_rate: Расход материала на неизолированный элемент
        """
        self.diameter = diameter
        self.insulated_rate = insulated_rate
        self.not_insulated_rate = not_insulated_rate

class RowOfSpecification:
    """
    Класс, описывающий строку спецификации
    """
    def __init__(self,
                 system,
                 function,
                 group,
                 name = '',
                 mark = '',
                 code = '',
                 maker = '',
                 unit = '',
                 local_description = '',
                 number = 0,
                 mass = '',
                 note = '',
                 smr_block = '',
                 smr_section = '',
                 smr_floor = ''):

        """
        Инициализация класса строки спецификации

        Args:
            system: Имя системы
            function: Имя функции
            group: Группирование
            name: Наименование
            mark: Маркировка
            code: Код изделия
            maker: Завод-изготовитель
            unit: Единица измерения
            local_description: Назначение, по которому ищем якорный элемент и удаляем
            number: Число
            mass: Масса в текстовом формате
            note: Примечание
        """

        self.system = system
        self.function = function
        self.group = group

        self.name = name
        self.mark = mark
        self.code = code
        self.maker = maker
        self.unit = unit
        self.number = number
        self.mass = mass
        self.note = note
        self.smr_block = smr_block
        self.smr_section = smr_section
        self.smr_floor = smr_floor

        self.local_description = local_description
        self.diameter = 0
        self.parentId = 0

class InsulationConsumables:
    """ Класс описывающий расходники изоляции """
    def __init__(self, name, mark, maker, unit, expenditure, is_expenditure_by_linear_meter):
        """
        Инициализация класса расходника изоляции

        Args:
            name: Наименование
            mark: Марка
            maker: Завод-изготовитель
            unit: Единицы измерения
            expenditure: Расход
            is_expenditure_by_linear_meter: Считается ли расход по метру погонному. Если нет - считаем по площади.
        """
        self.name = name
        self.mark = mark
        self.maker = maker
        self.unit = unit
        self.expenditure = expenditure
        self.is_expenditure_by_linear_meter = is_expenditure_by_linear_meter

class UnmodelingFactory:
    """ Класс, оперирующий созданием немоделируемых элементов """
    doc = None
    info = None

    COORDINATE_STEP = 0.01  # Шаг координаты на который разносим немоделируемые. ~3 мм, чтоб они не стояли в одном месте и чтоб не растягивали чертеж своим существованием
    DESCRIPTION_PARAM_NAME = 'ФОП_ВИС_Назначение'  # Пока нет в платформе, будет добавлено и перенесено в RevitParams

    # Значения параметра "ФОП_ВИС_Назначение" по которому определяется удалять элемент или нет
    EMPTY_DESCRIPTION = 'Пустая строка'
    IMPORT_DESCRIPTION = 'Импорт немоделируемых'
    MATERIAL_DESCRIPTION = 'Расчет краски и креплений'
    CONSUMABLE_DESCRIPTION = 'Расходники изоляции'
    AI_DESCRIPTION = 'Элементы АИ'

    # Значение группирования для элементов
    CONSUMABLE_GROUP = '12. Расходники изоляции'
    MATERIAL_GROUP = "12. Расчетные элементы"

    # Имена расчетов
    PIPE_METAL_RULE_NAME = 'Металлические крепления для трубопроводов'
    DUCT_METAL_RULE_NAME = 'Металлические крепления для воздуховодов'
    COLOR_RULE_NAME = 'Краска антикоррозионная, покрытие в два слоя. Расход - 0.2 кг на м²'
    GRUNT_RULE_NAME = 'Грунтовка для стальных труб, покрытие в один слой. Расход - 0.1 кг на м²'
    CLAMPS_RULE_NAME = 'Хомут трубный под шпильку М8'
    RAPID_CLAMPS_RULE_NAME = 'Хомут соединительный Rapid'
    PIN_RULE_NAME = 'Шпилька М8 1м/1шт'

    FAMILY_NAME = '_Якорный элемент'
    OUT_OF_SYSTEM_VALUE = '!Нет системы'
    OUT_OF_FUNCTION_VALUE = '!Нет функции'
    ws_id = None

    edited_reports = [] # Перчень редакторов элементов
    sync_status_report = None # Отчет о статусе необходимости синхронизации
    edited_status_report = None # Отчет о статусе занятых элементов

    # Максимальная встреченная координата в проекте. Обновляется в первый раз в get_base_location, далее обновляется в
    # при создании экземпляра якоря
    max_location_y = 0

    def __init__(self, doc):
        self.doc = doc
        self.info = self.doc.ProjectInformation

    def get_elements_types_by_category(self, category):
        """
        Получает типы элементов по их категории.

        Args:
            category: Категория элементов.

        Returns:
            List[Element]: Список типов элементов.
        """
        col = FilteredElementCollector(self.doc) \
            .OfCategory(category) \
            .WhereElementIsElementType() \
            .ToElements()
        return col

    def get_pipe_duct_insulation_types(self):
        """
        Получает типы изоляции труб и воздуховодов.

        Args:
            doc: Документ Revit.

        Returns:
            List[Element]: Список типов изоляции труб и воздуховодов.
        """
        # Создаем список категорий
        categories = List[BuiltInCategory]()
        categories.Add(BuiltInCategory.OST_PipeInsulations)
        categories.Add(BuiltInCategory.OST_DuctInsulations)

        multicategory_filter = ElementMulticategoryFilter(categories)

        return (FilteredElementCollector(self.doc).WherePasses(multicategory_filter)
                .WhereElementIsElementType().ToElements())

    def create_consumable_row_class_instance(self,
                                             system,
                                             function,
                                             block,
                                             section,
                                             floor,
                                             consumable,
                                             consumable_description):
        """
        Создает экземпляр класса расходника изоляции для генерации строки.

        Args:
            system: Система.
            function: Функция.
            consumable: Расходник.
            consumable_description: Описание расходника.

        Returns:
            RowOfSpecification: Экземпляр класса строки спецификации.
        """
        return RowOfSpecification(
            system=system,
            function=function,
            group=self.CONSUMABLE_GROUP,
            name=consumable.name,
            mark=consumable.mark,
            code='',  # У расходников не будет кода изделия
            maker=consumable.maker,
            unit=consumable.unit,
            local_description=consumable_description,
            smr_block=block,
            smr_section=section,
            smr_floor=floor
        )

    def create_material_row_class_instance(self,
                                           system,
                                           function,
                                           block,
                                           section,
                                           floor,
                                           rule_set,
                                           material_description):
        """
        Создает экземпляр класса материала для генерации строки.

        Args:
            system: Система.
            function: Функция.
            rule_set: Набор правил.
            material_description: Описание материала.

        Returns:
            RowOfSpecification: Экземпляр класса строки спецификации.
        """
        return RowOfSpecification(
            system= system,
            function=function,
            group=rule_set.group,
            name=rule_set.name,
            mark=rule_set.mark,
            code = rule_set.code,
            maker=rule_set.maker,
            unit=rule_set.unit,
            local_description=material_description,
            smr_block=block,
            smr_section=section,
            smr_floor=floor
        )

    def get_element_charactristic(self, element):
        """
        Получает значения параметров функции, системы, блока, секции и этажа из элемента.

        Args:
            element: Элемент Revit.

        Returns:
            Tuple[str, str, str, str, str]: Кортеж из значений функции, системы, блока, секции и этажа.
        """
        system = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISSystemName,
                                                self.OUT_OF_SYSTEM_VALUE)
        function = element.GetParamValueOrDefault(SharedParamsConfig.Instance.EconomicFunction,
                                                  self.OUT_OF_FUNCTION_VALUE)
        block = element.GetParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksBlock, '')
        section = element.GetParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksSection, '')
        floor = element.GetParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksLevel, '')
        return system, function, block, section, floor

    def get_base_location(self):
        """
        Получает базовую локацию для вставки первого из элементов.

        Returns:
            XYZ: Базовая локация.
        """
        if self.max_location_y == 0:
            # Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
            generic_models = self.get_elements_by_category(BuiltInCategory.OST_GenericModel)
            filtered_generics = [elem for elem in generic_models if elem.GetElementType()
                                 .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == self.FAMILY_NAME]

            if len(filtered_generics) == 0:
                return XYZ(0, 0, 0)

            max_y = None
            base_location_point = None

            for elem in filtered_generics:
                # Получаем LocationPoint элемента
                location_point = elem.Location.Point
                # Получаем значение Y из LocationPoint
                y_value = location_point.Y
                # Проверяем, является ли текущее значение Y максимальным
                if max_y is None or y_value > max_y:
                    max_y = y_value
                    base_location_point = location_point

            return XYZ(0, self.COORDINATE_STEP + max_y, 0)

        return XYZ(0, self.COORDINATE_STEP + self.max_location_y, 0)

    def update_location(self, loc):
        """
        Обновляет локацию, слегка увеличивая ее.

        Args:
            loc: Текущая локация.

        Returns:
            XYZ: Обновленная локация.
        """
        return XYZ(0, loc.Y + self.COORDINATE_STEP, 0)

    def get_ruleset(self):
        """
        Получает список правил для генерации материалов.

        Returns:
            List[GenerationRuleSet]: Список правил для генерации материалов.
        """

        gen_list = [
            GenerationRuleSet(
                group=self.MATERIAL_GROUP,
                name=self.DUCT_METAL_RULE_NAME,
                mark="",
                code="",
                unit="кг.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name,
                category=BuiltInCategory.OST_DuctCurves),
            GenerationRuleSet(
                group=self.MATERIAL_GROUP,
                name=self.PIPE_METAL_RULE_NAME,
                mark="",
                code="",
                unit="кг.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves),
            GenerationRuleSet(
                group=self.MATERIAL_GROUP,
                name=self.COLOR_RULE_NAME,
                mark="БТ-177",
                code="",
                unit="кг.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsPaintCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves),
            GenerationRuleSet(
                group=self.MATERIAL_GROUP,
                name=self.GRUNT_RULE_NAME,
                mark="ГФ-031",
                code="",
                unit="кг.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsPaintCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves),
            GenerationRuleSet(
                group=self.MATERIAL_GROUP,
                name=self.CLAMPS_RULE_NAME,
                mark="",
                code="",
                unit="шт.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsClampsCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves),
            GenerationRuleSet(
                group=self.MATERIAL_GROUP,
                name=self.PIN_RULE_NAME,
                mark="",
                code="",
                unit="шт.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsClampsCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves),
            GenerationRuleSet(
                group=self.MATERIAL_GROUP,
                name=self.RAPID_CLAMPS_RULE_NAME,
                mark="",
                code="",
                unit="шт.",
                maker="",
                method_name="ФОП_ВИС_Расчет хомутов Rapid",
                category=BuiltInCategory.OST_PipeCurves)
        ]
        return gen_list

    def get_element_editor_name(self, element):
        """
        Возвращает имя пользователя, который последним редактировал элемент.

        Args:
            element: Элемент Revit.

        Returns:
            str: Имя пользователя или None, если элемент не на редактировании.
        """
        user_name = __revit__.Application.Username
        edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
        if edited_by is None:
            return None

        if edited_by.lower() == user_name.lower():
            return None
        return edited_by

    def is_elemet_edited(self, element):
        """
        Проверяет, заняты ли элементы другими пользователями.
        """
        update_status = WorksharingUtils.GetModelUpdatesStatus(self.doc, element.Id)

        if update_status == ModelUpdatesStatus.UpdatedInCentral or update_status == ModelUpdatesStatus.DeletedInCentral:
            self.sync_status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию. "

        name = self.get_element_editor_name(element)
        if name is not None and name not in self.edited_reports:
            self.edited_reports.append(name)

        if name is not None or update_status == ModelUpdatesStatus.UpdatedInCentral:
            return True

        return False

    def show_report(self, exit_on_report = False):
        if len(self.edited_reports) > 0:
            self.edited_status_report = (
                "Часть элементов занята пользователями: {}".format(
                    ", ".join(self.edited_reports)
                )
            )

        if self.edited_status_report is not None or self.sync_status_report is not None:
            report_message = ''
            if self.sync_status_report is not None:
                report_message += self.sync_status_report
            if self.edited_status_report is not None:
                if report_message:
                    report_message += '\n'
                report_message += self.edited_status_report

            if exit_on_report:
                forms.alert(report_message, "Ошибка", exitscript=True)
            else:
                forms.alert(report_message, "Ошибка")

    def find_family_symbol(self):
        collector = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(FamilySymbol)

        for element in collector:
            if element.Family.Name == self.FAMILY_NAME:
                return element

        return None

    def is_family_in(self):
        """
        Проверяет, есть ли семейство в проекте.

        Args:
            doc: Документ Revit.

        Returns:
            FamilySymbol: Символ семейства, если оно есть в проекте, иначе None.
        """

        symbol = self.find_family_symbol()

        if not symbol:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            local_path = os.path.join(script_dir, self.FAMILY_NAME + '.rfa')

            with revit.Transaction("BIM: Загрузка семейства"):
                self.doc.LoadFamily(local_path)

            return self.find_family_symbol()

        return symbol

    def get_elements_by_category(self, category):
        """
        Возвращает список элементов по их категории.

        Args:
            category: Категория элементов.

        Returns:
            List[Element]: Список элементов.
        """
        return FilteredElementCollector(self.doc) \
            .OfCategory(category) \
            .WhereElementIsNotElementType() \
            .ToElements()

    def remove_models(self, description):
        """
        Удаляет элементы с переданным описанием.

        Args:
            description: Описание элемента.
        """
        # Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
        generic_model_collection = \
            [elem for elem in self.get_elements_by_category(BuiltInCategory.OST_GenericModel) if elem.GetElementType()
            .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == self.FAMILY_NAME]

        for element in generic_model_collection:
            self.is_elemet_edited(element)

        self.show_report(exit_on_report=True)

        with revit.Transaction("BIM: Очистка немоделируемых"):
            for element in generic_model_collection:
                if element.IsExistsParam(self.DESCRIPTION_PARAM_NAME):
                    elem_type = self.doc.GetElement(element.GetTypeId())
                    current_name = elem_type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString()
                    current_description = element.GetParamValueOrDefault(self.DESCRIPTION_PARAM_NAME)

                    if current_name == self.FAMILY_NAME:
                        if current_description is None or description in current_description:
                            self.doc.Delete(element.Id)

    def create_new_position(self, new_row_data, family_symbol, description, loc):
        """
        Генерирует пустые элементы в рабочем наборе немоделируемых.

        Args:
            new_row_data: Данные новой строки.
            family_symbol: Символ семейства.
            description: Описание.
            loc: Локация.
        """
        def set_param_value(shared_param, param_value):
            if param_value is not None:
                family_inst.SetParamValue(shared_param, param_value)

        new_row_data.number = round(new_row_data.number, 2) # заранее округляем, на случай значений типа 0.001. Для этого
        # можно не генерировать строки

        if new_row_data.number == 0 and description != self.EMPTY_DESCRIPTION:
            return

        self.max_location_y = loc.Y

        if self.ws_id is None:
            forms.alert('Не удалось найти рабочий набор "99_Немоделируемые элементы"', "Ошибка", exitscript=True)

        # Создаем элемент и назначаем рабочий набор
        family_inst = self.doc.Create.NewFamilyInstance(loc, family_symbol, Structure.StructuralType.NonStructural)

        family_inst_workset = family_inst.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        family_inst_workset.Set(self.ws_id.IntegerValue)

        group = '{}_{}_{}_{}_{}'.format(
            new_row_data.group, new_row_data.name, new_row_data.mark, new_row_data.maker, new_row_data.code)

        if self.doc.IsExistsParam(SharedParamsConfig.Instance.VISSpecNumbersCurrency):
            number_param = SharedParamsConfig.Instance.VISSpecNumbersCurrency
        else:
            number_param = SharedParamsConfig.Instance.VISSpecNumbers

        set_param_value(SharedParamsConfig.Instance.VISSystemName, new_row_data.system)
        set_param_value(SharedParamsConfig.Instance.VISGrouping, group)
        set_param_value(SharedParamsConfig.Instance.VISCombinedName, new_row_data.name)
        set_param_value(SharedParamsConfig.Instance.VISMarkNumber, new_row_data.mark)
        set_param_value(SharedParamsConfig.Instance.VISItemCode, new_row_data.code)
        set_param_value(SharedParamsConfig.Instance.VISManufacturer, new_row_data.maker)
        set_param_value(SharedParamsConfig.Instance.VISUnit, new_row_data.unit)
        set_param_value(number_param, new_row_data.number)
        set_param_value(SharedParamsConfig.Instance.VISMass, new_row_data.mass)
        set_param_value(SharedParamsConfig.Instance.VISNote, new_row_data.note)
        set_param_value(SharedParamsConfig.Instance.EconomicFunction, new_row_data.function)
        set_param_value(SharedParamsConfig.Instance.BuildingWorksBlock, new_row_data.smr_block)
        set_param_value(SharedParamsConfig.Instance.BuildingWorksSection, new_row_data.smr_section)
        set_param_value(SharedParamsConfig.Instance.BuildingWorksLevel, new_row_data.smr_floor)
        description_param = family_inst.GetParam(self.DESCRIPTION_PARAM_NAME)
        description_param.Set(description)

    def prepare_settings(self):
        settings =  self.info.GetParamValueOrDefault("ФОП_ВИС_Настройки немоделируемых", "")

        if settings == "":
            print "СБРОС НАСТРОЕК"
            def_sets = """
            ##UNMODELING_REGION_START##
            ##ENAMEL_NAME##
            Краска антикоррозионная, покрытие в два слоя. Расход - 0.2 кг на м²;
            ##ENAMEL_MARK##
            БТ-177;
            ##ENAMEL_CODE##
            ##ENAMEL_CREATOR##
            ##ENAMEL_UNIT##
            кг.;
            ##PRIMER_NAME##
            Грунтовка для стальных труб, покрытие в один слой. Расход - 0.1 кг на м²;
            ##PRIMER_MARK##
            ГФ-031;
            ##PRIMER_CODE##
            ##PRIMER_UNIT##
            кг.;
            ##PRIMER_CREATOR##
            ##PIPE_TYPES_METALL##
            ##PIPE_TYPES_COLOR##
            ##PIPE_TYPES_CLAMPS##
            ##DUCT_TYPES_METALL##
            ##CONSUMABLE1_NAMES##
            ##CONSUMABLE2_NAMES##
            ##CONSUMABLE3_NAMES##
            ##CONSUMABLE4_NAMES##
            ##CONSUMABLE1_MARKS##
            ##CONSUMABLE2_MARKS##
            ##CONSUMABLE3_MARKS##
            ##CONSUMABLE4_MARKS##
            ##CONSUMABLE1_CODES##
            ##CONSUMABLE2_CODES##
            ##CONSUMABLE3_CODES##
            ##CONSUMABLE4_CODES##
            ##CONSUMABLE1_MAKER##
            ##CONSUMABLE2_MAKER##
            ##CONSUMABLE3_MAKER##
            ##CONSUMABLE4_MAKER##
            ##CONSUMABLE1_UNIT##
            ##CONSUMABLE2_UNIT##
            ##CONSUMABLE3_UNIT##
            ##CONSUMABLE4_UNIT##
            ##CONSUMABLE1_RATE##
            ##CONSUMABLE2_RATE##
            ##CONSUMABLE3_RATE##
            ##CONSUMABLE4_RATE##
            ##CONSUMABLE1_RATE_BY_SQUARE##
            ##CONSUMABLE2_RATE_BY_SQUARE##
            ##CONSUMABLE3_RATE_BY_SQUARE##
            ##CONSUMABLE4_RATE_BY_SQUARE##
            ##UNMODELING_REGION_END##
            """
            self.info.SetParamValue("ФОП_ВИС_Настройки немоделируемых", def_sets)


    def get_setting_region(self, region_name):
        settings = self.info.GetParamValueOrDefault("ФОП_ВИС_Настройки немоделируемых", "")

        # Ищем текст между двумя тегами вида ##REGION_NAME##
        pattern = r"##" + re.escape(region_name) + r"##\s*(.*?)\s*##"
        match = re.search(pattern, settings, flags=re.DOTALL)

        if not match:
            return []

        region_text = match.group(1)

        # Разбиваем по ';'
        parts = [p.strip() for p in region_text.split(";")]

        # Убираем пустые элементы
        return [p for p in parts if p]

    def edit_setting_region(self, base_settings, region_name, new_value_list, append=True):
        if not new_value_list:
            return base_settings

        pattern = r"(##" + re.escape(region_name) + r"##\s*)(.*?)(?=\s*##[^#]+##)"
        match = re.search(pattern, base_settings, flags=re.DOTALL)

        if not match:
            return base_settings

        start_tag = match.group(1)
        old_value = match.group(2).strip()

        old_list = [x.strip() for x in old_value.split(";") if x.strip()] if old_value else []
        new_list = [x.strip() for x in new_value_list if x.strip()]

        final_list = old_list + new_list if append else new_list
        combined = "; ".join(final_list)

        new_settings = (
                base_settings[:match.start()] +
                start_tag +
                combined +
                base_settings[match.end():]
        )

        return new_settings

    def startup_checks(self):
        """
        Выполняет начальные проверки файла и семейства.

        Returns:
            FamilySymbol: Символ семейства.
        """
        if self.doc.IsFamilyDocument:
            forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

        # На всякий случай выполняем настройку параметров - в теории уже должны быть на месте, но лучше продублировать
        revit_params = [SharedParamsConfig.Instance.EconomicFunction,
                        SharedParamsConfig.Instance.VISSystemName,
                        SharedParamsConfig.Instance.BuildingWorksBlock,
                        SharedParamsConfig.Instance.BuildingWorksSection,
                        SharedParamsConfig.Instance.BuildingWorksLevel]


        with revit.Transaction("BIM: Настройка параметров"):
            for param in revit_params:
                self.sort_parameter_to_group(param.Name, GroupTypeId.Data)

            self.prepare_settings()


        project_parameters = ProjectParameters.Create(self.doc.Application)
        project_parameters.SetupRevitParams(self.doc, revit_params)

        family_symbol = self.is_family_in()

        if family_symbol is None:
            forms.alert(
                "Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.",
                "Ошибка",
                exitscript=True)

        self.check_family(family_symbol)
        self.check_worksets()

        return family_symbol

    def sort_parameter_to_group(self, paraname, group):
        # Получаем параметр по имени
        param = self.doc.GetSharedParam(paraname)

        if param is None:
            return

        definition = param.GetDefinition()
        definition.ReInsertToGroup(self.doc, group)

    def check_worksets(self):
        """
        Проверяет наличие рабочего набора немоделируемых элементов.

        """
        if WorksetTable.IsWorksetNameUnique(self.doc, '99_Немоделируемые элементы'):
            with revit.Transaction("BIM: Добавление рабочего набора"):
                new_ws = Workset.Create(self.doc, '99_Немоделируемые элементы')
                forms.alert('Был создан рабочий набор "99_Немоделируемые элементы". '
                            'Откройте диспетчер рабочих наборов и снимите галочку с параметра "Видимый на всех видах". '
                            'В данном рабочем наборе будут создаваться немоделируемые элементы '
                            'и требуется исключить их видимость.',
                            "Рабочие наборы")
                self.ws_id = new_ws.Id
        else:
            fws = FilteredWorksetCollector(self.doc).OfKind(WorksetKind.UserWorkset)
            for ws in fws:
                if ws.Name == '99_Немоделируемые элементы':
                    self.ws_id = ws.Id

                if ws.Name == '99_Немоделируемые элементы' and ws.IsVisibleByDefault:
                    forms.alert('Рабочий набор "99_Немоделируемые элементы" на данный момент отображается на всех видах.'
                                ' Откройте диспетчер рабочих наборов и снимите галочку с параметра "Видимый на всех видах".'
                                ' В данном рабочем наборе будут создаваться немоделируемые элементы '
                                'и требуется исключить их видимость.',
                                "Рабочие наборы")
                    self.ws_id = ws.Id
                    return

    def check_family(self, family_symbol):
        """
        Проверяет семейство на наличие необходимых параметров.

        Args:
            family_symbol: Символ семейства.

        Returns:
            List: Список отсутствующих параметров.
        """
        param_names_list = [
            self.DESCRIPTION_PARAM_NAME,
            SharedParamsConfig.Instance.VISNote.Name,
            SharedParamsConfig.Instance.VISMass.Name,
            SharedParamsConfig.Instance.VISPosition.Name,
            SharedParamsConfig.Instance.VISGrouping.Name,
            SharedParamsConfig.Instance.EconomicFunction.Name,
            SharedParamsConfig.Instance.VISSystemName.Name,
            SharedParamsConfig.Instance.VISCombinedName.Name,
            SharedParamsConfig.Instance.VISMarkNumber.Name,
            SharedParamsConfig.Instance.VISItemCode.Name,
            SharedParamsConfig.Instance.VISUnit.Name,
            SharedParamsConfig.Instance.VISManufacturer.Name,
            SharedParamsConfig.Instance.BuildingWorksBlock.Name,
            SharedParamsConfig.Instance.BuildingWorksSection.Name,
            SharedParamsConfig.Instance.BuildingWorksLevel.Name
            ]

        family = family_symbol.Family
        symbol_params = self.get_family_shared_parameter_names(family)

        result = []
        missing_params = [param for param in param_names_list if param not in symbol_params]

        if missing_params:
            missing_params_str = ", ".join(missing_params)
            forms.alert('Обновите семейство якорного элемента. Параметры {} отсутствуют.'.format(missing_params_str),
                        "Ошибка", exitscript=True)

        return result

    def get_family_shared_parameter_names(self, family):
        """
        Получает список имен общих параметров семейства.

        Args:
            family: Семейство.

        Returns:
            List[str]: Список имен общих параметров.
        """
        # Открываем документ семейства для редактирования
        family_doc = self.doc.EditFamily(family)

        shared_parameters = []
        try:
            # Получаем менеджер семейства
            family_manager = family_doc.FamilyManager

            # Получаем все параметры семейства
            parameters = family_manager.GetParameters()

            # Фильтруем параметры, чтобы оставить только общие
            shared_parameters = [param.Definition.Name for param in parameters if param.IsShared]

            return shared_parameters
        finally:
            # Закрываем документ семейства без сохранения изменений
            family_doc.Close(False)

class MaterialCalculator:
    """
    Класс-калькулятор для расходных элементов труб и воздуховодов.
    """
    doc = None

    def __init__(self, doc):
        self.doc = doc

    def get_connectors(self, element):
        connectors = []

        if isinstance(element, FamilyInstance) and element.MEPModel.ConnectorManager is not None:
            connectors.extend(element.MEPModel.ConnectorManager.Connectors)

        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves]) and \
                isinstance(element, MEPCurve) and element.ConnectorManager is not None:
            connectors.extend(element.ConnectorManager.Connectors)

        return connectors

    def get_fitting_insulation_area(self, element, host):
        area = 0

        for solid in dosymep.Revit.Geometry.ElementExtensions.GetSolids(host):
            for face in solid.Faces:
                area += face.Area

        # Складываем площадь коннекторов хоста
        if area > 0:
            false_area = 0
            connectors = self.get_connectors(host)
            for connector in connectors:
                if connector.Shape == ConnectorProfileType.Rectangular:
                    height = connector.Height
                    width = connector.Width

                    false_area += height * width
                if connector.Shape == ConnectorProfileType.Round:
                    radius = connector.Radius
                    false_area += radius * radius * math.pi
                if connector.Shape == ConnectorProfileType.Oval:
                    false_area += 0

            # Вычитаем площадь пустоты на местах коннекторов
            area -= false_area

        return area

    def get_curve_len_area_parameters_values(self, element):
        """
        Получает значения длины и площади поверхности элемента.

        Args:
            element: Элемент, для которого требуется получить параметры.

        Returns:
            tuple: Длина и площадь поверхности элемента в метрах и квадратных метрах соответственно.
        """
        length = element.GetParamValueOrDefault(BuiltInParameter.CURVE_ELEM_LENGTH)

        if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
            outer_diameter = element.GetParamValueOrDefault(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
            area = math.pi * outer_diameter * length
        else:
            area = element.GetParamValueOrDefault(BuiltInParameter.RBS_CURVE_SURFACE_AREA)

        if element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
            host = self.doc.GetElement(element.HostElementId)
            if host.Category.IsId(BuiltInCategory.OST_DuctFitting):
                # Для залагавшей изоляции
                if host is None:
                    return 0, 0

                area = self.get_fitting_insulation_area(element, host)

        if length is None:
            length = 0
        if area is None:
            area = 0

        length = UnitUtils.ConvertFromInternalUnits(length, UnitTypeId.Meters)
        area = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters)

        return length, area

    def get_pipe_material_class_instances(self):
        """
        Возвращает коллекцию вариантов расхода металла по диаметрам для изолированных труб.

        Returns:
            list: Список экземпляров MaterialVariants, отсортированный по диаметру.
        """
        dict_var_p_mat = {
            15: 0.238, 20: 0.204, 25: 0.187, 32: 0.170, 40: 0.187, 50: 0.2448, 65: 0.3315,
            80: 0.3791, 100: 0.629, 125: 0.901, 150: 1.054, 200: 1.309, 999: 0.1564
        }

        variants = []
        for diameter, insulated_rate in dict_var_p_mat.items():
            variant = MaterialVariants(diameter, insulated_rate, 0)
            variants.append(variant)

        variants_sorted = sorted(variants, key=lambda x: x.diameter)
        return variants_sorted

    def get_collar_material_class_instances(self):
        """
        Возвращает коллекцию вариантов расхода хомутов по диаметрам для изолированных и неизолированных труб.

        Returns:
            list: Список экземпляров MaterialVariants, отсортированный по диаметру.
        """
        dict_var_collars = {
            15: [2, 1.5], 20: [3, 2], 25: [3.5, 2], 32: [4, 2.5], 40: [4.5, 3], 50: [5, 3], 65: [6, 4],
            80: [6, 4], 100: [6, 4.5], 125: [7, 5], 999: [7, 5]
        }

        variants = []

        for diameter, rates in dict_var_collars.items():
            insulated_rate, not_insulated_rate = rates
            variant = MaterialVariants(diameter, insulated_rate, not_insulated_rate)
            variants.append(variant)

        variants_sorted = sorted(variants, key=lambda x: x.diameter)
        return variants_sorted

    def is_pipe_insulated(self, pipe):
        """
        Проверяет, изолирована ли труба.

        Args:
            pipe: Элемент трубы.

        Returns:
            bool: True, если труба изолирована, иначе False.
        """
        pipe_insulation_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
        dependent_elements = pipe.GetDependentElements(pipe_insulation_filter)
        return len(dependent_elements) > 0

    def get_material_value_by_rate(self, material_rate, curve_length):
        """
        Возвращает количество материала в зависимости от его расхода на длину.

        Args:
            material_rate: Расход материала.
            curve_length: Длина кривой.

        Returns:
            int: Количество материала.
        """
        number = curve_length / material_rate
        if number < 1:
            number = 1
        return int(number)

    def get_rapid_collars_number(self, pipe_diameter, pipe_length):
        """Возвращает число хомутов Rapid на трубу. Если размер трубы не подходит - возвращает 0 и хомут не создается
        Args:
            pipe_diameter: Диаметр трубы.
            pipe_length: Длина трубы.

        Returns:
            int: Количество хомутов.
        """
        if pipe_length <= 3:
            return 0

        d_variants = [50, 70, 80, 100, 125, 150, 200]

        if pipe_diameter not in d_variants:
            return 0
        collars_count = int(math.ceil(pipe_length / 3.0)) - 1

        return collars_count

    def get_collars_and_pins_number(self, pipe, pipe_diameter, pipe_length):
        """
        Возвращает число хомутов и шпилек.

        Args:
            pipe: Элемент трубы.
            pipe_diameter: Диаметр трубы.
            pipe_length: Длина трубы.

        Returns:
            int: Количество хомутов и шпилек.
        """
        collar_materials = self.get_collar_material_class_instances()

        if pipe_length < 0.5:
            return 0

        for collar_material in collar_materials:
            if pipe_diameter <= collar_material.diameter:
                if self.is_pipe_insulated(pipe):
                    return self.get_material_value_by_rate(collar_material.insulated_rate, pipe_length)
                else:
                    return self.get_material_value_by_rate(collar_material.not_insulated_rate, pipe_length)

    def get_duct_material_mass(self, duct, duct_diameter, duct_width, duct_height, duct_area):
        """
        Возвращает массу металла воздуховодов.

        Args:
            duct: Элемент воздуховода.
            duct_diameter: Диаметр воздуховода.
            duct_width: Ширина воздуховода.
            duct_height: Высота воздуховода.
            duct_area: Площадь поверхности воздуховода.

        Returns:
            float: Масса металла воздуховодов.
        """
        perimeter = 0
        if duct.DuctType.Shape == ConnectorProfileType.Round:
            perimeter = math.pi * duct_diameter

        if duct.DuctType.Shape == ConnectorProfileType.Rectangular:
            duct_width = UnitUtils.ConvertFromInternalUnits(
                duct.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM),
                UnitTypeId.Millimeters)
            duct_height = UnitUtils.ConvertFromInternalUnits(
                duct.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM),
                UnitTypeId.Millimeters)
            perimeter = 2 * (duct_width + duct_height)

        if perimeter < 1001:
            mass = duct_area * 0.65
        elif perimeter < 1801:
            mass = duct_area * 1.22
        else:
            mass = duct_area * 2.25

        return mass

    def get_pipe_material_mass(self, pipe_length, pipe_diameter):
        """
        Возвращает массу металла для труб.

        Args:
            pipe_length: Длина трубы.
            pipe_diameter: Диаметр трубы.

        Returns:
            float: Масса металла для труб.
        """
        pipe_materials = self.get_pipe_material_class_instances()

        for pipe_material in pipe_materials:
            if pipe_diameter <= pipe_material.diameter:
                return pipe_material.insulated_rate * pipe_length

    def get_grunt_mass(self, pipe_area):
        """
        Возвращает массу грунтовки.

        Args:
            pipe_area: Площадь поверхности трубы.

        Returns:
            float: Масса грунтовки.
        """
        number = pipe_area * 0.1
        return number

    def get_color_mass(self, pipe_area):
        """
        Возвращает массу краски.

        Args:
            pipe_area: Площадь поверхности трубы.

        Returns:
            float: Масса краски.
        """
        number = pipe_area * 0.2 * 2
        return number

    def get_consumables_class_instances(self, insulation_element_type):
        """
        Возвращает список экземпляров расходников изоляции для конкретных ее типов.

        Args:
            insulation_element_type: Тип элемента изоляции.

        Returns:
            list: Список экземпляров InsulationConsumables.
        """
        def is_name_value_exists(shared_param):
            value = insulation_element_type.GetParamValueOrDefault(shared_param)
            return value is not None and value != ""

        def is_expenditure_value_exist(shared_param):
            value = insulation_element_type.GetParamValueOrDefault(shared_param)
            return value is not None and value != 0

        consumables_name_1 = SharedParamsConfig.Instance.VISInsulationConsumable1Name
        consumables_mark_1 = SharedParamsConfig.Instance.VISInsulationConsumable1MarkNumber
        consumables_maker_1 = SharedParamsConfig.Instance.VISInsulationConsumable1Manufacturer
        consumables_unit_1 = SharedParamsConfig.Instance.VISInsulationConsumable1Unit
        consumables_expenditure_1 = SharedParamsConfig.Instance.VISInsulationConsumable1ConsumptionPerSqM
        is_expenditure_by_linear_meter_1 = SharedParamsConfig.Instance.VISInsulationConsumable1ConsumptionPerMetr

        consumables_name_2 = SharedParamsConfig.Instance.VISInsulationConsumable2Name
        consumables_mark_2 = SharedParamsConfig.Instance.VISInsulationConsumable2MarkNumber
        consumables_maker_2 = SharedParamsConfig.Instance.VISInsulationConsumable2Manufacturer
        consumables_unit_2 = SharedParamsConfig.Instance.VISInsulationConsumable2Unit
        consumables_expenditure_2 = SharedParamsConfig.Instance.VISInsulationConsumable2ConsumptionPerSqM
        is_expenditure_by_linear_meter_2 = SharedParamsConfig.Instance.VISInsulationConsumable2ConsumptionPerMetr

        consumables_name_3 = SharedParamsConfig.Instance.VISInsulationConsumable3Name
        consumables_mark_3 = SharedParamsConfig.Instance.VISInsulationConsumable3MarkNumber
        consumables_maker_3 = SharedParamsConfig.Instance.VISInsulationConsumable3Manufacturer
        consumables_unit_3 = SharedParamsConfig.Instance.VISInsulationConsumable3Unit
        consumables_expenditure_3 = SharedParamsConfig.Instance.VISInsulationConsumable3ConsumptionPerSqM
        is_expenditure_by_linear_meter_3 = SharedParamsConfig.Instance.VISInsulationConsumable3ConsumptionPerMetr

        result = []
        if is_name_value_exists(consumables_name_1) and is_expenditure_value_exist(consumables_expenditure_1):
            result.append(
                InsulationConsumables(
                insulation_element_type.GetParamValueOrDefault(consumables_name_1),
                insulation_element_type.GetParamValueOrDefault(consumables_mark_1),
                insulation_element_type.GetParamValueOrDefault(consumables_maker_1),
                insulation_element_type.GetParamValueOrDefault(consumables_unit_1),
                insulation_element_type.GetParamValueOrDefault(consumables_expenditure_1),
                insulation_element_type.GetParamValueOrDefault(is_expenditure_by_linear_meter_1))
            )

        if is_name_value_exists(consumables_name_2) and is_expenditure_value_exist(consumables_expenditure_2):
            result.append(
                InsulationConsumables(
                insulation_element_type.GetParamValueOrDefault(consumables_name_2),
                insulation_element_type.GetParamValueOrDefault(consumables_mark_2),
                insulation_element_type.GetParamValueOrDefault(consumables_maker_2),
                insulation_element_type.GetParamValueOrDefault(consumables_unit_2),
                insulation_element_type.GetParamValueOrDefault(consumables_expenditure_2),
                insulation_element_type.GetParamValueOrDefault(is_expenditure_by_linear_meter_2))
            )

        if is_name_value_exists(consumables_name_3) and is_expenditure_value_exist(consumables_expenditure_3):
            result.append(
                InsulationConsumables(
                insulation_element_type.GetParamValueOrDefault(consumables_name_3),
                insulation_element_type.GetParamValueOrDefault(consumables_mark_3),
                insulation_element_type.GetParamValueOrDefault(consumables_maker_3),
                insulation_element_type.GetParamValueOrDefault(consumables_unit_3),
                insulation_element_type.GetParamValueOrDefault(consumables_expenditure_3),
                insulation_element_type.GetParamValueOrDefault(is_expenditure_by_linear_meter_3))
            )

        return result