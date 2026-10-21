"""Функции работы со строками (логика сохранена, добавлены «подготовленные» быстрые версии)."""

import re


def has_no_comma(address):
    return "," not in address


# ---------------------------------------------------------------------------
# Замены с расстановкой запятых в адресах без запятых
# ---------------------------------------------------------------------------
def prepare_commas(commas_replacements):
    """Заранее считаем нижний регистр для ключей (делается один раз)."""
    return [(old, old.lower(), new) for old, new in commas_replacements]


def add_commas_to_addresses_prepared(address, prepared):
    """Выполняет замены для добавления запятых в адресах без запятых, игнорируя регистр."""
    address_lower = address.lower()
    result = address
    for old, old_lower, new in prepared:
        if old_lower in address_lower:
            result = result.replace(old, new)
            address_lower = result.lower()
    return result


def add_commas_to_addresses(address, commas_replacements):
    return add_commas_to_addresses_prepared(address, prepare_commas(commas_replacements))


# ---------------------------------------------------------------------------
# Замены независимо от регистра
# ---------------------------------------------------------------------------
def prepare_replacements(replacements):
    return [(old, old.lower(), new, new.lower(), len(old)) for old, new in replacements]


def replace_ignore_case_prepared(address, prepared):
    lower_address = address.lower()
    for old, old_lower, new, new_lower, n_old in prepared:
        idx = lower_address.find(old_lower)
        if idx == -1:
            continue
        len_new = len(new)
        while idx != -1:
            address = address[:idx] + new + address[idx + n_old :]
            lower_address = lower_address[:idx] + new_lower + lower_address[idx + n_old :]
            idx = lower_address.find(old_lower, idx + len_new)
    return address


def replace_ignore_case_fast(address, replacements):
    return replace_ignore_case_prepared(address, prepare_replacements(replacements))


# ---------------------------------------------------------------------------
# Удаление «сельсоветов», «сельских поселений» и т.д.
# ---------------------------------------------------------------------------
unwanted = (
    "сельское поселение",
    "поссовет",
    "сельсовет",
    "СЕЛЬСОВЕТ",
    "с.п.",
    "г.п.",
    "С.П.",
    "г П.",
    "г П ",
    "ПОССОВЕТ",
    "гп рп /пгт /",
)


def remove_settlement(address):
    # Разбиваем строку по запятым
    parts = address.split(",")
    # Фильтруем части, исключая те, что содержат ключевые слова
    filtered_parts = [part for part in parts if not any(kw in part for kw in unwanted)]
    return ",".join(filtered_parts).strip(",")


# "Ачинск г" -> "г Ачинск" (это уменьшает вероятность ошибок при распознавании)
_G_KEYWORDS = ("г", "г.", "город")


def g_gorod_maker(address: str) -> str:
    if not address:
        return address

    segments = address.split(",")
    changed = False

    for i, segment in enumerate(segments):
        seg = segment.strip()
        if not seg:
            continue

        last_space = seg.rfind(" ")
        if last_space == -1:
            continue

        last_word = seg[last_space + 1 :].lower()
        if last_word not in _G_KEYWORDS:
            continue

        city_part = seg[:last_space].strip()

        if city_part.count(" ") > 1:
            continue

        segments[i] = "г " + city_part
        changed = True

    if changed:
        return ",".join(segments)
    return address


# ---------------------------------------------------------------------------
# Нормализация городов
# г.о.Тейково,г.Тейково - > г Тейково, и т.п.
# ---------------------------------------------------------------------------
_BAD_PREFIXES = tuple(
    p.lower()
    for p in (
        "г.о.",
        "г о.",
        "г п г ",
        "гп г ",
        "г п. г ",
        "городской округ г ",
        "г о.город ",
        "город ",
    )
)


def normalize_city(address: str) -> str:
    bad_prefixes = _BAD_PREFIXES
    seen = set()
    result = []

    for seg in address.split(","):
        original = seg.strip()
        s = original
        s_low = s.lower()

        is_city = (
            s_low.startswith("г ")
            or s_low.startswith("г.")
            or any(s_low.startswith(p) for p in bad_prefixes)
        )

        if is_city:
            changed = True
            while changed:
                changed = False
                for p in bad_prefixes:
                    if s_low.startswith(p):
                        s = s[len(p) :].lstrip(" .")
                        s_low = s.lower()
                        changed = True
                        break

            if s_low.startswith("г."):
                s = "г " + s[2:].lstrip(" .")
            elif not s_low.startswith("г "):
                s = "г " + s.lstrip(" .")

            if len(s) > 2:
                s = "г " + s[2:3].upper() + s[3:].lower()
        else:
            s = original

        if s not in seen:
            seen.add(s)
            result.append(s)

    return ", ".join(result)


# ---------------------------------------------------------------------------
# Добавление субъекта РФ перед городом
# ---------------------------------------------------------------------------
def prepare_town_region(town_region):
    """[(ключ_lower, имя_города_lower, длина_имени, значение), ...] — считается один раз."""
    prepared = []
    for key, value in town_region.items():
        city_name = key[2:]
        prepared.append((key.lower(), city_name.lower(), len(city_name), value))
    return prepared


def add_region_to_town_prepared(address, prepared, separator):
    if separator in address:
        return address

    address_lower = address.lower()
    n_addr = len(address_lower)

    # Шаг 1: если строка начинается с названия города без "г ", добавляем "г "
    for _key_lower, city_lower, n_city, _value in prepared:
        if address_lower == city_lower or (
            address_lower.startswith(city_lower)
            and n_addr > n_city
            and address_lower[n_city : n_city + 1] in (",", " ")
        ):
            address = "г " + address
            break

    address_lower = address.lower()
    # Шаг 2: если строка начинается с ключа из словаря, добавляем значение перед ключом
    for key_lower, _city_lower, _n_city, value in prepared:
        if address_lower.startswith(key_lower):
            address = value + separator + address
            break

    return address


def add_region_to_town(address, town_region, separator):
    return add_region_to_town_prepared(address, prepare_town_region(town_region), separator)


def apply_order_to_region(address):
    """
    Если в какой-либо запятой-подстроке на первом месте стоит "область" или "край"
    и в подстроке ровно два слова — меняем слова местами.
    """
    if "," not in address:
        return address

    segments = address.split(",")

    for i, segment in enumerate(segments):
        words = segment.strip().split()
        if len(words) == 2 and words[0].lower() in ("область", "край"):
            segments[i] = f"{words[1]} {words[0]}"

    return ", ".join(s.strip() for s in segments)
