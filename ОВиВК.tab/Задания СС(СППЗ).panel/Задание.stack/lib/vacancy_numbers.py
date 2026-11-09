# -*- coding: utf-8 -*-
"""Хранение и подбор освобождённых номеров задания СС(СППЗ)."""
import codecs
import json
import os


def read_vacancies(folder):
    """Читает единый список вакансий из папки задания."""
    path = os.path.join(folder, "Вакантные номера.json")
    if not os.path.exists(path):
        return []
    with codecs.open(path, 'r', encoding='utf-8') as source:
        return json.load(source)


def write_vacancies(folder, vacancies):
    """Сортирует и сохраняет текущий список вакансий."""
    path = os.path.join(folder, "Вакантные номера.json")
    vacancies.sort(key=lambda item: (item['system'], item['floor'], item['number'], item['type']))
    with codecs.open(path, 'w', encoding='utf-8') as target:
        json.dump(vacancies, target, ensure_ascii=False, indent=4)


def vacancy_from_data(data):
    """Получает вакансию для записи в json из сохранённой марки, системы и этажа задания."""
    system, floor = data.equipment_base_name.rsplit('-', 1)
    number = int(data.json_name.rsplit('-', 1)[1])
    if data.closed_algorithm:
        kind = 'НЗ'
    else:
        kind = 'НО' if data.json_name.startswith('НО-') else 'Оборудование'
    return {'type': kind, 'floor': floor, 'system': system, 'number': number}


def vacancy_key(item):
    """
    Возвращает ключ вакансии для поиска дубликатов при добавлении в список.

    Ключ — кортеж (тип, этаж, система или None, номер). Две вакансии
    с одинаковым ключом считаются одной и той же свободной позицией.
    Для НЗ учитывается система, для НО и оборудования вместо неё ставится
    None, поскольку их номера выдаются независимо от системы.

    """
    return (item['type'], item['floor'],
            item['system'] if item['type'] == 'НЗ' else None, item['number'])


def allocate_number(vacancies, kind, floor, system, next_number):
    """Занимает минимальную подходящую вакансию либо следующий новый номер."""
    candidates = sorted([item for item in vacancies
                         if item['type'] == kind and item['floor'] == floor
                         and (kind != 'НЗ' or item['system'] == system)],
                        key=lambda item: item['number'])
    if candidates:
        item = candidates[0]
        vacancies.remove(item)
        return item['number']
    return next_number
