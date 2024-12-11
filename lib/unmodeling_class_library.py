# -*- coding: utf-8 -*-
import math

import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *

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

# класс-правило для генерации элементов, содержит имя метода, категорию и описание материала
class GenerationRuleSet:
    def __init__(self, group, name, mark, code, maker, unit, method_name, category, is_type):
        self.group = group
        self.name = name
        self.mark = mark
        self.maker = maker
        self.unit = unit
        self.category = category
        self.method_name = method_name
        self.is_type = is_type
        self.code = code

# класс содержащий расчетные данные для создаваемых расходников
class MaterialVariants:
    def __init__(self, diameter, insulated_rate, not_insulated_rate):
        self.diameter = diameter
        self.insulated_rate = insulated_rate
        self.not_insulated_rate = not_insulated_rate

# класс содержащий все ячейки типовой спецификации
class RowOfSpecification:
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
                 note = ''):
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

        self.local_description = local_description
        self.diameter = 0
        self.parentId = 0

# класс описывающий расходники изоляции
class InsulationConsumables:
    def __init__(self, name, mark, maker, unit, expenditure, is_expenditure_by_linear_meter):
        self.name = name
        self.mark = mark
        self.maker = maker
        self.unit = unit
        self.expenditure = expenditure
        self.is_expenditure_by_linear_meter = is_expenditure_by_linear_meter

# класс оперирующий созданием немоделируемых элементов
class UnmodelingFactory:
    coordinate_step = 0.01 # Шаг координаты на который разносим немоделируемые. ~3 мм, чтоб они не стояли в одном месте и чтоб не растягивали чертеж своим существованием
    description_param_name = 'ФОП_ВИС_Назначение' # Пока нет в платформе, будет добавлено и перенесено в RevitParams

    # Значения параметра "ФОП_ВИС_Назначение" по которому определеяется удалять элемент или нет
    empty_description = 'Пустая строка'
    import_description = 'Импорт немоделируемых'
    material_description = 'Расчет краски и креплений'
    consumable_description = 'Расходники изоляции'

    # Значение группирования для элементов
    consumable_group = '12. Расходники изоляции'

    family_name = '_Якорный элемент'
    out_of_system_value = '!Нет системы'
    out_of_function_value = '!Нет функции'
    ws_id = None

    # Максимальная встреченная координата в проекте. Обновляется в первый раз в get_base_location, далее обновляется в
    # при создании экземпляра якоря
    max_location_y = 0

    # Получает типы элементов по их категории
    def get_elements_types_by_category(self, doc, category):
        col = FilteredElementCollector(doc) \
            .OfCategory(category) \
            .WhereElementIsElementType() \
            .ToElements()
        return col

    def get_pipe_duct_insulation_types(self, doc):
        result = []

        result.extend(self.get_elements_types_by_category(doc, BuiltInCategory.OST_PipeInsulations))
        result.extend(self.get_elements_types_by_category(doc, BuiltInCategory.OST_DuctInsulations))
        return result

    # Создаем экземпляр класса расходника изоляции для генерации строки
    def create_consumable_row_class_instance(self, system, function, consumable, consumable_description):
        return RowOfSpecification(
            system,
            function,
            self.consumable_group,
            consumable.name,
            consumable.mark,
            '',  # У расходников не будет кода изделия
            consumable.maker,
            consumable.unit,
            consumable_description
        )

    # Создаем экземпляр класса материала для генерации строки
    def create_material_row_class_instance(self, system, function, rule_set, material_description):
        return RowOfSpecification(
            system,
            function,
            rule_set.group,
            rule_set.name,
            rule_set.mark,
            rule_set.code,
            rule_set.maker,
            rule_set.unit,
            material_description
        )

    # Забираем из элемента значение параметров функции и системы
    def get_system_function(self, element):
        system = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISSystemName,
                                                self.out_of_system_value)
        function = element.GetParamValueOrDefault(SharedParamsConfig.Instance.EconomicFunction,
                                                  self.out_of_function_value)
        return system, function

    # Получаем базовую локацию для вставки первого из элементов. Проверяем якоря проекта на самый большой Y и возвращаем его, немного увеличив.
    # Если уже max_location_y не ноль - возвращает его с небольшим прирощением
    def get_base_location(self, doc):
        if self.max_location_y == 0:
            # Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
            generic_models = self.get_elements_by_category(doc, BuiltInCategory.OST_GenericModel)

            filtered_generics = \
                [elem for elem in generic_models if elem.GetElementType()
                .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == self.family_name]

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

            return XYZ(0, self.coordinate_step + max_y, 0)

        return XYZ(0, self.coordinate_step + self.max_location_y, 0)

    # Получает базовую локацию и слегка ее увеличивает
    def update_location(self, loc):
        return XYZ(0, loc.Y + self.coordinate_step, 0)

    # Получаем список правил для генерации материалов
    def get_ruleset(self):
        gen_list = [
            GenerationRuleSet(
                group="12. Расчетные элементы",
                name="Металлические крепления для воздуховодов",
                mark="",
                code="",
                unit="кг.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name,
                category=BuiltInCategory.OST_DuctCurves,
                is_type=False),
            GenerationRuleSet(
                group="12. Расчетные элементы",
                name="Металлические крепления для трубопроводов",
                mark="",
                code="",
                unit="кг.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves,
                is_type=False),

            GenerationRuleSet(
                group="12. Расчетные элементы",
                name="Краска антикоррозионная за два раза",
                mark="БТ-177",
                code="",
                unit="кг.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsPaintCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves,
                is_type=False),
            GenerationRuleSet(
                group="12. Расчетные элементы",
                name="Грунтовка для стальных труб",
                mark="ГФ-031",
                code="",
                unit="кг.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsPaintCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves,
                is_type=False),
            GenerationRuleSet(

                group="12. Расчетные элементы",
                name="Хомут трубный под шпильку М8",
                mark="",
                code="",
                unit="шт.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsClampsCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves,
                is_type=False),
            GenerationRuleSet(
                group="12. Расчетные элементы",
                name="Шпилька М8 1м/1шт",
                mark="",
                code="",
                unit="шт.",
                maker="",
                method_name=SharedParamsConfig.Instance.VISIsClampsCalculation.Name,
                category=BuiltInCategory.OST_PipeCurves,
                is_type=False)
        ]
        return gen_list

    # Возвращает имя занявшего элемент или None
    def get_element_editor_name(self, element):
        user_name = __revit__.Application.Username
        edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
        if edited_by is None:
            return None

        if edited_by.lower() == user_name.lower():
            return None
        return edited_by

    # Возвращает FamilySymbol, если семейство есть в проекте, None если нет
    def is_family_in(self, doc):
        collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(FamilySymbol)

        for element in collector:
            if element.Family.Name == self.family_name:
                return element

        return None

    # Возвращает список элементов по их категории
    def get_elements_by_category(self, doc, category):
        return FilteredElementCollector(doc) \
            .OfCategory(category) \
            .WhereElementIsNotElementType() \
            .ToElements()

    #для прогона новых ревизий генерации немоделируемых: стирает элемент с переданным именем модели
    def remove_models(self, doc, description):
        user_name = __revit__.Application.Username
        # Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
        generic_model_collection = \
            [elem for elem in self.get_elements_by_category(doc, BuiltInCategory.OST_GenericModel) if elem.GetElementType()
            .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == self.family_name]

        for element in generic_model_collection:
            edited_by = self.get_element_editor_name(element)
            if edited_by:
                forms.alert("Якорные элементы не были обработаны, так как были заняты пользователями:" + edited_by,
                            "Ошибка",
                            exitscript=True)

        for element in generic_model_collection:
            if element.IsExistsParam(self.description_param_name):
                elem_type = doc.GetElement(element.GetTypeId())
                current_name = elem_type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString()
                current_description = element.GetParamValueOrDefault(self.description_param_name)

                if  current_name == self.family_name:
                    if description in current_description:
                        doc.Delete(element.Id)

    # Генерирует пустые элементы в рабочем наборе немоделируемых
    def create_new_position(self, doc, new_row_data, family_symbol, description, loc):
        def set_param_value(shared_param, param_value):
            if param_value is not None:
                family_inst.SetParamValue(shared_param, param_value)

        if new_row_data.number == 0 and description != self.empty_description:
            return

        self.max_location_y = loc.Y

        if self.ws_id is None:
            forms.alert('Не удалось найти рабочий набор "99_Немоделируемые элементы"', "Ошибка", exitscript=True)


        # Создаем элемент и назначаем рабочий набор
        family_inst = doc.Create.NewFamilyInstance(loc, family_symbol, Structure.StructuralType.NonStructural)

        family_inst_workset = family_inst.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        family_inst_workset.Set(self.ws_id.IntegerValue)

        group = '{}_{}_{}_{}_{}'.format(
            new_row_data.group, new_row_data.name, new_row_data.mark, new_row_data.maker, new_row_data.code)


        set_param_value(SharedParamsConfig.Instance.VISSystemName, new_row_data.system)
        set_param_value(SharedParamsConfig.Instance.VISGrouping, group)
        set_param_value(SharedParamsConfig.Instance.VISCombinedName, new_row_data.name)
        set_param_value(SharedParamsConfig.Instance.VISMarkNumber, new_row_data.mark)
        set_param_value(SharedParamsConfig.Instance.VISItemCode, new_row_data.code)
        set_param_value(SharedParamsConfig.Instance.VISManufacturer, new_row_data.maker)
        set_param_value(SharedParamsConfig.Instance.VISUnit, new_row_data.unit)
        set_param_value(SharedParamsConfig.Instance.VISSpecNumbers, new_row_data.number)
        set_param_value(SharedParamsConfig.Instance.VISMass, new_row_data.mass)
        set_param_value(SharedParamsConfig.Instance.VISNote, new_row_data.note)
        set_param_value(SharedParamsConfig.Instance.EconomicFunction, new_row_data.function)
        description_param = family_inst.GetParam(self.description_param_name)
        description_param.Set(description)

    # Проверяем файл, семейство
    def startup_checks(self, doc):
        if doc.IsFamilyDocument:
            forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

        family_symbol = self.is_family_in(doc)

        if family_symbol is None:
            forms.alert(
                "Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.",
                "Ошибка",
                exitscript=True)

        self.check_family(family_symbol, doc)

        self.check_worksets(doc)

        # На всякий случай выполняем настройку параметров - в теории уже должны быть на месте, но лучше продублировать
        revit_params = [SharedParamsConfig.Instance.EconomicFunction,
                        SharedParamsConfig.Instance.VISSystemName]

        project_parameters = ProjectParameters.Create(doc.Application)
        project_parameters.SetupRevitParams(doc, revit_params)

        return family_symbol

    # Проверяет наличие рабочего набора немоделируемых. Если нет - создаем и просим выключить видимость, если есть но видим -
    # просто просим ее выключить
    def check_worksets(self, doc):
        if WorksetTable.IsWorksetNameUnique(doc, '99_Немоделируемые элементы'):
            with revit.Transaction("Добавление рабочего набора"):
                new_ws = Workset.Create(doc, '99_Немоделируемые элементы')
                forms.alert('Был создан рабочий набор "99_Немоделируемые элементы". '
                            'Откройте диспетчер рабочих наборов и снимите галочку с параметра "Видимый на всех видах". '
                            'В данном рабочем наборе будут создаваться немоделируемые элементы '
                            'и требуется исключить их видимость.',
                            "Рабочие наборы")
                self.ws_id = new_ws.Id
        else:
            fws = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
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

    # Проверяем семейство на наличие параметров. Если чего-то нет - останавливаем скрипт
    def check_family(self, family_symbol, doc):
        param_names_list = [
            self.description_param_name,
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
            SharedParamsConfig.Instance.VISManufacturer.Name
            ]

        family = family_symbol.Family
        symbol_params = self.get_family_shared_parameter_names(doc, family)

        result = []
        missing_params = [param for param in param_names_list if param not in symbol_params]

        if missing_params:
            missing_params_str = ", ".join(missing_params)
            forms.alert('Обновите семейство якорного элемента. Параметры {} отсутствуют.'.format(missing_params_str),
                        "Ошибка", exitscript=True)

        return result

    # Получаем список имен общих параметров семейства
    def get_family_shared_parameter_names(self, doc, family):
        # Открываем документ семейства для редактирования
        family_doc = doc.EditFamily(family)

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

# класс-калькулятор для расходных элементов труб и воздуховодов
class MaterialCalculator:
    # Получаем значения длины и площади хоста изоляции. Для фитингов не сработает
    def get_curve_len_area_parameters_values(self, element):
        length = element.GetParamValueOrDefault(BuiltInParameter.CURVE_ELEM_LENGTH)

        # Если просто брать площадь поверхности для труб - ревит возвращает площадь по условному диаметру
        if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
            outer_diameter = element.GetParamValueOrDefault(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
            area = math.pi * outer_diameter * length
        else:
            area = element.GetParamValueOrDefault(BuiltInParameter.RBS_CURVE_SURFACE_AREA)

        # Это для соединительных деталей в случае если идет обсчет хостов изоляции может сработать
        if length is None or area is None:
            return 0, 0

        length = UnitUtils.ConvertFromInternalUnits(
            length,
            UnitTypeId.Meters)
        area = UnitUtils.ConvertFromInternalUnits(
            area,
            UnitTypeId.SquareMeters)

        return length, area

    # Получаем экземпляры класса материала для металла труб
    def get_pipe_material_class_instances(self):
        """ Возвращает коллекцию вариантов расхода металла по диаметрам для изолированных труб. Для неизолированных 0 """

        # Ключ словаря - диаметр. Первое значение списка по ключу - значение для изолированной трубы, второе - для неизолированной
        # для труб согласовано использование в расчетах только изолированных трубопроводов, поэтому в качестве неизолированной при создании
        # экземпляра варианта используем 0
        # Запас 70% задан по согласованию. В этих значениях он уже учтен, в нормативке он соответственно меньше.
        # Цифры приняты по http://slpl.ru/node/218 с запасом, по согласованию с инженерами
        dict_var_p_mat = {15: 0.238, 20: 0.204, 25: 0.187, 32: 0.170, 40: 0.187, 50: 0.2448, 65: 0.3315,
                          80: 0.3791, 100: 0.629, 125: 0.901, 150: 1.054, 200: 1.309, 999: 0.1564}

        variants = []
        for diameter, insulated_rate in dict_var_p_mat.items():
            variant = MaterialVariants(diameter, insulated_rate, 0)
            variants.append(variant)

        variants_sorted = sorted(variants, key=lambda x: x.diameter)
        return variants_sorted

    # Получаем экземпляры класса материала для хомутов или шпилек
    def get_collar_material_class_instances(self):
        """ Возвращает коллекцию вариантов расхода хомутов по диаметрам для изолированных и неизолированных труб """

        # Ключ словаря - диаметр. Первое значение списка по ключу - значение для изолированной трубы, второе - для неизолированной
        # Количество хомутов на метр по диаметру принято по согласованию с инженерами
        dict_var_collars = {15: [2, 1.5], 20: [3, 2], 25: [3.5, 2], 32: [4, 2.5], 40: [4.5, 3], 50: [5, 3], 65: [6, 4],
                            80: [6, 4], 100: [6, 4.5], 125: [7, 5], 999: [7, 5]}

        variants = []

        for diameter, rates in dict_var_collars.items():
            insulated_rate, not_insulated_rate = rates
            variant = MaterialVariants(diameter, insulated_rate, not_insulated_rate)
            variants.append(variant)

        variants_sorted = sorted(variants, key=lambda x: x.diameter)
        return variants_sorted

    # Если труба изолирована - возвращает True
    def is_pipe_insulated(self, pipe):
        pipe_insulation_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
        dependent_elements = pipe.GetDependentElements(pipe_insulation_filter)
        return len(dependent_elements) > 0

    # возвращает количество материала в зависимости от его расхода на длину, если количество меньше 1 - возвращает 1
    def get_material_value_by_rate(self, material_rate, curve_length):
        number = curve_length / material_rate
        if number < 1:
            number = 1
        return int(number)

    # Возвращает число хомутов и шпилек
    def get_collars_and_pins_number(self, pipe, pipe_diameter, pipe_length):
        collar_materials = self.get_collar_material_class_instances()

        if pipe_length < 0.5:
            return 0

        for collar_material in collar_materials:
            if pipe_diameter <= collar_material.diameter:
                if self.is_pipe_insulated(pipe):
                    return self.get_material_value_by_rate(collar_material.insulated_rate, pipe_length)
                else:
                    return self.get_material_value_by_rate(collar_material.not_insulated_rate, pipe_length)

    # Возвращает массу металла воздуховодов
    def get_duct_material_mass(self, duct, duct_diameter, duct_width, duct_height, duct_area):
        perimeter = 0
        if duct.DuctType.Shape == ConnectorProfileType.Round:
            perimeter = math.pi * duct_diameter

        if duct.DuctType.Shape == ConnectorProfileType.Rectangular:
            duct_width = UnitUtils.ConvertFromInternalUnits(
                duct.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM),
                UnitTypeId.Millimeters)  # Преобразование в метры

            duct_height = UnitUtils.ConvertFromInternalUnits(
                duct.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM),
                UnitTypeId.Millimeters)  # Преобразование в метры

            perimeter = 2 * (duct_width + duct_height)

        # Расходы массы на площади идут из СНиП IV-2-82 Сборник 20. Вентиляция и
        # кондиционирование воздуха
        if perimeter < 1001:
            mass = duct_area * 0.65
        elif perimeter < 1801:
            mass = duct_area * 1.22
        else:
            mass = duct_area * 2.25

        return mass

    # Возвращает массу металла для труб
    def get_pipe_material_mass(self, pipe_length, pipe_diameter):
        pipe_materials = self.get_pipe_material_class_instances()

        for pipe_material in pipe_materials:
            if pipe_diameter <= pipe_material.diameter:
                return pipe_material.insulated_rate * pipe_length

    # Возвращает массу грунтовки
    def get_grunt_mass(self, pipe_area):
        number = pipe_area / 10
        return number

    # Возвращает массу краски
    def get_color_mass(self, pipe_area):
        number = pipe_area * 0.2 * 2
        return number

    # Возвращает список экземпляров расходников изоляции для конкретных ее типов
    def get_consumables_class_instances(self, insulation_element_type):
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
        # Если у 1 расходника и имя и расход не равны None, то добавляем расходник в результаты
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

        # Если у 2 расходника и имя и расход не равны None, то добавляем расходник в результаты
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

        # Если у 3 расходника и имя и расход не равны None, то добавляем расходник в результаты
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