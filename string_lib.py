"""Функции работы со строками (логика сохранена, добавлены «подготовленные» быстрые версии)."""

import heapq
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
_ANCHOR = 3  # длина «якоря» — характерного фрагмента ключа для быстрого предфильтра


class ReplacementTable(list):
    """Список замен (old, old_lower, new, new_lower, len(old)) + индекс «якорей».

    Исходный алгоритм для КАЖДОЙ строки перебирал все ключи (сотни str.find), хотя в строке
    обычно не встречается ни один. Ключ может сработать, только если в строке есть любой его
    фрагмент — берём самый редкий 3-символьный фрагмент ключа («якорь») и проверяем только те
    ключи, чьи якоря встретились. Порядок ключей и каскадные замены сохранены: после реальной
    замены к кандидатам добавляются ключи, якоря которых появились в новом тексте.
    Результат полностью совпадает с прежним алгоритмом (предфильтр лишь отсеивает заведомо
    несработавшие ключи; каждый кандидат по-прежнему проверяется через str.find).
    """

    def __init__(self, items):
        super().__init__(items)
        keys = [item[1] for item in self]
        freq = {}
        for k in keys:
            for g in {k[i : i + _ANCHOR] for i in range(len(k) - _ANCHOR + 1)}:
                freq[g] = freq.get(g, 0) + 1
        anchors = {}
        always = []
        for i, k in enumerate(keys):
            if len(k) < _ANCHOR:
                always.append(i)  # слишком короткий ключ — проверяем всегда
                continue
            grams = [k[j : j + _ANCHOR] for j in range(len(k) - _ANCHOR + 1)]
            best = min(range(len(grams)), key=lambda j: (freq[grams[j]], j))
            anchors.setdefault(grams[best], []).append(i)
        self._anchors = anchors
        self._anchor_keys = frozenset(anchors)
        self._always = always

    def hits(self, text):
        """Индексы ключей, якоря которых встречаются в text (надмножество реально способных сработать)."""
        n = len(text) - _ANCHOR + 1
        if n <= 0:
            return set()
        grams = {text[i : i + _ANCHOR] for i in range(n)}
        found = set()
        anchors = self._anchors
        for g in self._anchor_keys.intersection(grams):
            found.update(anchors[g])
        return found


def prepare_replacements(replacements):
    return ReplacementTable(
        [(old, old.lower(), new, new.lower(), len(old)) for old, new in replacements]
    )


def _replace_ignore_case_fast(address, table):
    lower_address = address.lower()
    queued = table.hits(lower_address)
    queued.update(table._always)
    heap = list(queued)
    heapq.heapify(heap)
    items = table
    while heap:
        i = heapq.heappop(heap)
        old, old_lower, new, new_lower, n_old = items[i]
        idx = lower_address.find(old_lower)
        if idx == -1:
            continue
        len_new = len(new)
        n_new_lower = len(new_lower)
        touched = []
        while idx != -1:
            address = address[:idx] + new + address[idx + n_old :]
            lower_address = lower_address[:idx] + new_lower + lower_address[idx + n_old :]
            touched.append(idx)
            idx = lower_address.find(old_lower, idx + len_new)
        # Текст изменился: новые ключи-кандидаты могут появиться только рядом с вставленным
        # фрагментом (с запасом в 2 символа с каждой стороны для «пограничных» фрагментов).
        for pos in touched:
            region = lower_address[max(0, pos - _ANCHOR + 1) : pos + n_new_lower + _ANCHOR - 1]
            for j in table.hits(region):
                if j > i and j not in queued:
                    queued.add(j)
                    heapq.heappush(heap, j)
    return address


def replace_ignore_case_prepared(address, prepared):
    if isinstance(prepared, ReplacementTable):
        return _replace_ignore_case_fast(address, prepared)
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


_UNWANTED_RE = re.compile("|".join(re.escape(kw) for kw in unwanted))


def remove_settlement(address):
    # Быстрый путь: ни одно ключевое слово не встречается в строке целиком — удалять нечего
    # (ключевые слова не содержат запятой, поэтому поиск по всей строке эквивалентен поиску по частям)
    if _UNWANTED_RE.search(address) is None:
        return address.strip(",")
    # Разбиваем строку по запятым
    parts = address.split(",")
    # Фильтруем части, исключая те, что содержат ключевые слова
    filtered_parts = [part for part in parts if _UNWANTED_RE.search(part) is None]
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
class TownTable(list):
    """[(ключ_lower, имя_города_lower, длина_имени, значение), ...] + индексы по первым символам.

    Вместо перебора всех городов для каждой строки смотрим только тех, чьё начало совпадает
    с началом строки. Порядок городов сохраняется — результат тот же, что и раньше.
    """

    def __init__(self, items):
        super().__init__(items)
        self.by_city = {}  # первые 2 символа названия города -> [элементы по порядку]
        self.by_key = {}  # первые 3 символа ключа «г Город» -> [элементы по порядку]
        self.simple = True  # False, если есть слишком короткие названия -> безопасный полный перебор
        for item in self:
            key_lower, city_lower = item[0], item[1]
            if len(city_lower) < 2 or len(key_lower) < 3:
                self.simple = False
            self.by_city.setdefault(city_lower[:2], []).append(item)
            self.by_key.setdefault(key_lower[:3], []).append(item)


def prepare_town_region(town_region):
    """[(ключ_lower, имя_города_lower, длина_имени, значение), ...] — считается один раз."""
    prepared = []
    for key, value in town_region.items():
        city_name = key[2:]
        prepared.append((key.lower(), city_name.lower(), len(city_name), value))
    return TownTable(prepared)


def add_region_to_town_prepared(address, prepared, separator):
    if separator in address:
        return address

    address_lower = address.lower()
    n_addr = len(address_lower)
    fast = isinstance(prepared, TownTable) and prepared.simple

    # Шаг 1: если строка начинается с названия города без "г ", добавляем "г "
    for _key_lower, city_lower, n_city, _value in (
        prepared.by_city.get(address_lower[:2], ()) if fast else prepared
    ):
        if address_lower == city_lower or (
            address_lower.startswith(city_lower)
            and n_addr > n_city
            and address_lower[n_city : n_city + 1] in (",", " ")
        ):
            address = "г " + address
            break

    address_lower = address.lower()
    # Шаг 2: если строка начинается с ключа из словаря, добавляем значение перед ключом
    for key_lower, _city_lower, _n_city, value in (
        prepared.by_key.get(address_lower[:3], ()) if fast else prepared
    ):
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