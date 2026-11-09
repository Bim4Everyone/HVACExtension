# -*- coding: utf-8 -*-
"""Хранение и подбор освобождённых номеров задания СС(СППЗ)."""
import codecs
import json
import os


def read_vacancies(folder):
    """Читает единый список вакансий из папки задания."""
    path = os.path.join(folder, u"Вакантные номера.json")
    if not os.path.exists(path):
        return []
    with codecs.open(path, 'r', encoding='utf-8') as source:
        return json.load(source)


def write_vacancies(folder, vacancies):
    """Сортирует и сохраняет текущий список вакансий."""
    path = os.path.join(folder, u"Вакантные номера.json")
    vacancies.sort(key=lambda item: (item['system'], item['floor'], item['number'], item['type']))
    with codecs.open(path, 'w', encoding='utf-8') as target:
        json.dump(vacancies, target, ensure_ascii=False, indent=4)


def vacancy_from_data(data):
    """Получает вакансию из сохранённой марки, системы и этажа задания."""
    system, floor = data.equipment_base_name.rsplit('-', 1)
    number = int(data.json_name.rsplit('-', 1)[1])
    if data.closed_algorithm:
        kind = u'НЗ'
    else:
        kind = u'НО' if data.json_name.startswith(u'НО-') else u'Оборудование'
    return {'type': kind, 'floor': floor, 'system': system, 'number': number}


def vacancy_key(item):
    """Возвращает ключ без системы для НО и оборудования."""
    return (item['type'], item['floor'],
            item['system'] if item['type'] == u'НЗ' else None, item['number'])


def allocate_number(vacancies, kind, floor, system, used, next_number):
    """Занимает минимальную подходящую вакансию либо следующий новый номер."""
    candidates = sorted([item for item in vacancies
                         if item['type'] == kind and item['floor'] == floor
                         and (kind != u'НЗ' or item['system'] == system)],
                        key=lambda item: item['number'])
    for item in candidates:
        vacancies.remove(item)
        if item['number'] not in used:
            number = item['number']
            used.add(number)
            return number
    while next_number in used:
        next_number += 1
    used.add(next_number)
    return next_number
