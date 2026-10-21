"""Модуль очистки и обогащения строк (шаг 1 исходного руководства)."""

import re

import pandas as pd

from remover import remove_street_info
from settings import (
    back_replacements,
    commas_replacements,
    replacements,
    separator,
    small_subject_replacements,
    subject_replacements,
    town_region,
)
from string_lib import (
    add_commas_to_addresses_prepared,
    add_region_to_town_prepared,
    apply_order_to_region,
    g_gorod_maker,
    has_no_comma,
    normalize_city,
    prepare_commas,
    prepare_replacements,
    prepare_town_region,
    remove_settlement,
    replace_ignore_case_prepared,
)

# --- Всё, что можно посчитать заранее, считаем один раз при импорте ---
_COMMAS = prepare_commas(commas_replacements)
_REPLACEMENTS = prepare_replacements(replacements)
_BACK_REPLACEMENTS = prepare_replacements(back_replacements)
_TOWN_REGION = prepare_town_region(town_region)

_ALL_SUBJECT_VALUES_LOWER = [
    v.lower() for v in list(subject_replacements.values()) + list(small_subject_replacements.values())
]
_SUBJECT_ITEMS = [(k, k.lower(), v) for k, v in subject_replacements.items()]

_RE_ZIP8 = re.compile(r"\b\d{8},\s*")
_RE_ZIP7 = re.compile(r"\b\d{7},\s*")
_RE_ZIP6C = re.compile(r"\b\d{6},\s*")
_RE_ZIP6 = re.compile(r"\b\d{6}\s*")
_RE_ZIP3 = re.compile(r"\b\d{3},\s*")
_RE_QUOTE = re.compile(r'^"\s*')
_RE_SPACES = re.compile(r"\s{2,}")


def process_address(address):
    """Очистка и обогащение одного адреса."""
    if not isinstance(address, str):
        if pd.isna(address):
            return ""
        address = str(address)

    # Расставляем запятые в строках "без запятых"
    if has_no_comma(address):
        address = add_commas_to_addresses_prepared(address, _COMMAS)

    # Серединное удаление "сельсоветов", "сельских поселений" и т.д.
    address = remove_settlement(address)

    # Замены независимо от регистра
    address = replace_ignore_case_prepared(address, _REPLACEMENTS)

    # Удаление почтового индекса и следующей за ним запятой
    address = _RE_ZIP8.sub("", address)
    address = _RE_ZIP7.sub("", address)
    address = _RE_ZIP6C.sub("", address)
    address = _RE_ZIP6.sub("", address)

    # Удаляем странные "643,"
    address = _RE_ZIP3.sub("", address)

    # Запятые / точки в начале строки
    if address.startswith(","):
        address = address.lstrip(", ").lstrip()
    if address.startswith("."):
        address = address.lstrip(". ").lstrip()

    # Двойная кавычка в начале строки
    address = _RE_QUOTE.sub("", address)

    # Лишние пробелы
    address = _RE_SPACES.sub(" ", address).strip()

    # Торцевое удаление (улицы, дома и т.п.)
    address = remove_street_info(address)

    # "область Амурская" -> "Амурская область"
    address = apply_order_to_region(address)

    address_lower = address.lower()
    found = False

    # "Воронежская," в начале строки -> "Воронежская область,"
    for key, key_lower, value in _SUBJECT_ITEMS:
        if address_lower.startswith(key_lower + ","):
            address = address.replace(key, value)
            found = True
            break

    # Прямое включение названия субъекта РФ
    for value_lower in _ALL_SUBJECT_VALUES_LOWER:
        if value_lower in address_lower:
            found = True
            break

    # Если не нашли — ищем по уникальному ключу, затем по списку исключений
    if not found:
        for key, key_lower, value in _SUBJECT_ITEMS:
            if key_lower in address_lower:
                address = value + separator + address
                break
            elif ("алтай" in address_lower) and ("респ" in address_lower):
                address = "Республика Алтай" + separator + address
                break
            elif ("крым" in address_lower) and ("респ" in address_lower):
                address = "Республика Крым" + separator + address
                break
            elif ("коми" in address_lower) and ("респ" in address_lower):
                address = "Республика Коми" + separator + address
                break
            elif ("алтайский" in address_lower) and ("кр" in address_lower):
                address = "Алтайский край" + separator + address
                break
            elif "костромская" in address_lower:
                address = "Костромская область" + separator + address
                break
            elif "томская" in address_lower:
                address = "Томская область" + separator + address
                break
            elif "омская" in address_lower:
                address = "Омская область" + separator + address
                break
            elif "ямало" in address_lower:
                address = "Ямало-Ненецкий автономный округ" + separator + address
                break
            elif "ненецкий" in address_lower:
                address = "Ненецкий автономный округ" + separator + address
                break

    # Добавляем субъект РФ перед городом
    address = add_region_to_town_prepared(address, _TOWN_REGION, separator)

    # Меняем «спасённые» названия НП обратно на реальные
    address = replace_ignore_case_prepared(address, _BACK_REPLACEMENTS)

    if "санкт-петербург" in address.lower():
        address = "Санкт-Петербург"

    # Нормализация названий городов
    address = normalize_city(address)

    # "Ачинск г" -> "г Ачинск"
    address = g_gorod_maker(address)

    return address
