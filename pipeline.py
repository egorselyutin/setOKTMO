"""Конвейер обработки: очистка -> распознавание -> метрики -> Excel.

Модуль не зависит от Streamlit (его можно импортировать из дочерних процессов).
Логика очистки и распознавания полностью повторяет cleaner+enhancer.py и recognizer.py,
но работает быстрее: дедупликация адресов, предвычисленные индексы, кэш, параллелизм.
"""

from __future__ import annotations

import io
import math
import multiprocessing as mp
import os
import re
import time
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
import xlsxwriter
from rapidfuzz import fuzz, process

from stringsCleaner import process_address

# Названия колонок по умолчанию (используются, если в файле колонка называется «Местонахождение»)
COL_SRC = "Местонахождение"
COL_CLEAN = "Местонахождение_Обработано"
COL_CLEAN_ADDR = "Адрес_Обработано"  # имя новой колонки, если исходная колонка называется «Адрес…»

# Слова, при наличии которых колонка с «адрес» в названии точно НЕ является почтовым адресом
_NOT_ADDRESS_WORDS = ("почт", "mail", "сайт", "url", "ip")


def find_address_column(columns):
    """Определяет колонку с адресами и имя колонки для результата очистки.

    Порядок поиска (берётся первая подходящая колонка слева направо):
      1. «Местонахождение» (точное совпадение, регистр не важен);
      2. «Адрес» (точное совпадение);
      3. любое название, СОДЕРЖАЩЕЕ «адрес» («Адрес места», «Адрес местонахождения» и т.п.);
      4. любое название, содержащее «местонахождение».
    Колонки вроде «Адрес электронной почты» пропускаются.

    Возвращает (колонка_с_адресами, имя_новой_колонки) или (None, None), если ничего не найдено.
      - если в названии есть «адрес»  -> «Адрес_Обработано»;
      - иначе                         -> «Местонахождение_Обработано».
    """
    cols = [(c, str(c).strip().lower()) for c in columns]

    def first(pred):
        for c, name in cols:
            if pred(name):
                return c
        return None

    found = (
        first(lambda n: n == "местонахождение")
        or first(lambda n: n == "адрес")
        or first(lambda n: "адрес" in n and not any(w in n for w in _NOT_ADDRESS_WORDS))
        or first(lambda n: "местонахождение" in n)
    )
    if found is None:
        return None, None
    clean_name = COL_CLEAN_ADDR if "адрес" in str(found).strip().lower() else COL_CLEAN
    return found, clean_name
OUT_COLS = ["Субъект", "Район", "НП", "ОКТМО"]
EXCEL_MAX_ROWS = 1_048_575  # 1 строка занята заголовком

# Режим структурного поиска: "0" (по умолчанию) — только если основной поиск ничего не нашёл;
# "1" — при уверенном совпадении он главнее основного (исправляет и случаи, когда нечёткий поиск выбрал не тот НП).
STRUCT_FIRST = os.environ.get("OKTMO_STRUCT_FIRST", "0") == "1"

# Порог, ниже которого параллельный запуск не окупается (запуск процессов ~1-3 с)
MIN_PARALLEL_CLEAN = 4000
MIN_PARALLEL_MATCH = 400

COL_SUBJ_REF = "Субъект Российской Федерации"
COL_DIST_REF = "Район"
COL_NP_REF = "Наименование населенного пункта"
COL_OKTMO_REF = "ОКТМО"


# ---------------------------------------------------------------------------
# Чтение Excel
# ---------------------------------------------------------------------------
_XLS_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # OLE2-контейнер: настоящий .xls (BIFF)
_ZIP_MAGIC = b"PK\x03\x04"  # ZIP: .xlsx / .xlsm / .xlsb


def sniff_excel_format(src) -> str:
    """Определяет формат по первым байтам файла, а не по расширению (оно часто врёт).

    'xls'  — старый двоичный Excel 97-2003;
    'xlsx' — современный Excel (ZIP/OOXML);
    'html' — «.xls», который на самом деле HTML/XML/текст (так часто выгружают из веб-систем);
    ''     — не удалось определить (пробуем читать как есть).
    """
    try:
        if hasattr(src, "read"):
            pos = src.tell() if hasattr(src, "tell") else 0
            head = src.read(8)
            src.seek(pos)
        else:
            with open(src, "rb") as f:
                head = f.read(8)
    except Exception:
        return ""
    if head.startswith(_XLS_MAGIC):
        return "xls"
    if head.startswith(_ZIP_MAGIC):
        return "xlsx"
    if head.lstrip(b"\xef\xbb\xbf \t\r\n")[:1] == b"<":
        return "html"
    return ""


def read_excel_fast(src, **kw) -> pd.DataFrame:
    """Читает .xlsx и .xls. calamine (Rust) в разы быстрее openpyxl/xlrd и понимает оба формата;
    если он не установлен или не справился — запасной движок: xlrd для .xls, openpyxl для .xlsx."""
    fmt = sniff_excel_format(src)
    if fmt == "html":
        raise ValueError(
            "Файл с расширением .xls на самом деле не книга Excel (это HTML/текст). "
            "Откройте его в Excel и сохраните как «Книга Excel (.xlsx)»."
        )
    engines = ["calamine", "xlrd"] if fmt == "xls" else ["calamine", "openpyxl"]
    problems = []
    for engine in engines:
        if hasattr(src, "seek"):
            src.seek(0)
        try:
            return pd.read_excel(src, engine=engine, **kw)
        except ImportError:
            problems.append(f"{engine}: не установлен")
        except Exception as e:  # повреждённый файл, пароль и т.п. — пробуем следующий движок
            problems.append(f"{engine}: {e}")
    if fmt == "xls" and all("не установлен" in p for p in problems):
        raise ValueError("Для чтения .xls установите пакет python-calamine или xlrd (pip install python-calamine xlrd).")
    raise ValueError(
        "Не удалось прочитать файл как Excel (.xlsx/.xls). Файл повреждён или защищён паролем? "
        + "; ".join(problems)
    )


# ---------------------------------------------------------------------------
# Ресурсы сервера: сколько процессов можно запускать
# ---------------------------------------------------------------------------
# Проблема: на облачном хостинге (контейнер/cgroup) os.cpu_count() возвращает число ядер ВСЕГО
# физического сервера (десятки), а реально контейнеру доступна доля ядра и ограниченная память.
# Запуск 8 процессов (каждый ~0,2-0,3 ГБ) там приводит к нехватке памяти и «throttle» (урезанию CPU),
# и обработка обрывается. Поэтому смотрим на лимиты контейнера и подбираем число процессов под них.
# На обычном компьютере (Windows/macOS/Linux без лимитов) поведение прежнее: min(ядер, 8).
WORKER_MB = int(os.environ.get("OKTMO_WORKER_MB", "300"))     # память одного процесса-воркера
RESERVE_MB = int(os.environ.get("OKTMO_RESERVE_MB", "1200"))  # запас главному процессу (файл, таблицы, Excel)


def _read_text(path: str):
    try:
        with open(path, encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return None


def parse_cpu_max(text):
    """cgroup v2, файл cpu.max: 'max 100000' (без лимита) или '50000 100000' (= 0,5 ядра)."""
    parts = (text or "").split()
    if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit() and int(parts[1]) > 0:
        return int(parts[0]) / int(parts[1])
    return None


def parse_cfs(quota_text, period_text):
    """cgroup v1: cpu.cfs_quota_us / cpu.cfs_period_us (квота -1 = без лимита)."""
    try:
        q, per = int(quota_text), int(period_text)
    except (TypeError, ValueError):
        return None
    return q / per if q > 0 and per > 0 else None


def parse_mem_bytes(text):
    """Значение лимита/использования памяти из cgroup; 'max' и «практически бесконечность» -> None."""
    try:
        v = int((text or "").strip())
    except ValueError:
        return None
    return v if 0 < v < (1 << 60) else None


def cpu_limit() -> int:
    """Сколько ядер реально доступно: доступные процессу ядра и квота контейнера (cgroup)."""
    n = os.cpu_count() or 1
    try:
        n = min(n, len(os.sched_getaffinity(0)))
    except AttributeError:  # Windows/macOS
        pass
    for quota in (
        parse_cpu_max(_read_text("/sys/fs/cgroup/cpu.max")),
        parse_cfs(_read_text("/sys/fs/cgroup/cpu/cpu.cfs_quota_us"), _read_text("/sys/fs/cgroup/cpu/cpu.cfs_period_us")),
    ):
        if quota is not None:
            n = min(n, max(1, int(quota)))
    return max(1, n)


def memory_budget_mb():
    """Сколько МБ памяти ещё свободно для нас (с учётом лимита контейнера). None — не удалось определить."""
    free = []
    limit = parse_mem_bytes(_read_text("/sys/fs/cgroup/memory.max")) or parse_mem_bytes(
        _read_text("/sys/fs/cgroup/memory/memory.limit_in_bytes")
    )
    used = parse_mem_bytes(_read_text("/sys/fs/cgroup/memory.current")) or parse_mem_bytes(
        _read_text("/sys/fs/cgroup/memory/memory.usage_in_bytes")
    )
    if limit:
        free.append((limit - (used or 0)) / 2**20)
    meminfo = _read_text("/proc/meminfo")
    if meminfo:
        for line in meminfo.splitlines():
            if line.startswith("MemAvailable:"):
                free.append(int(line.split()[1]) / 1024)
    return min(free) if free else None


def default_workers() -> int:
    """Число процессов: OKTMO_WORKERS, если задано; иначе min(ядра, 8), урезанное по памяти."""
    env = os.environ.get("OKTMO_WORKERS")
    if env and env.isdigit() and int(env) > 0:
        return int(env)
    n = min(cpu_limit(), 8)
    budget = memory_budget_mb()
    if budget is not None:
        n = min(n, int((budget - RESERVE_MB) // WORKER_MB))
    return max(1, n)


def resource_summary() -> str:
    """Строка для логов сервера: на какие ресурсы рассчитываем."""
    b = memory_budget_mb()
    return (
        f"процессов={default_workers()}, ядер доступно={cpu_limit()}, "
        f"свободной памяти={'?' if b is None else f'{b:.0f} МБ'}"
    )


# ---------------------------------------------------------------------------
# Справочник ОКТМО и поиск (логика recognizer.py)
# ---------------------------------------------------------------------------
def normalize_text(text) -> str:
    if text is None or (isinstance(text, float) and text != text):
        return ""
    if not isinstance(text, str) and pd.isna(text):
        return ""
    return str(text).strip().lower()


def _fmt_oktmo(v) -> str:
    """Код ОКТМО строкой с лидирующими нулями (8 или 11 цифр).

    Если справочник прочитан правильно (колонка как текст), значение приходит строкой и не меняется.
    Страховка на случай, когда код всё же превратился в число (старый файл, другой движок чтения):
    возвращаем нули — 8 цифр для муниципальных образований, 11 для населённых пунктов."""
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    if s.isdigit() and len(s) not in (8, 11):
        s = s.zfill(8 if len(s) < 8 else 11)
    return s


class OktmoIndex:
    """Неизменяемый индекс справочника. Легко пиклится и передаётся в процессы."""

    def __init__(self, ref: pd.DataFrame):
        for c in (COL_SUBJ_REF, COL_DIST_REF, COL_NP_REF, COL_OKTMO_REF):
            if c not in ref.columns:
                raise ValueError(f"В справочнике ОКТМО нет колонки «{c}».")

        def norm(col):
            return ref[col].fillna("").astype(str).str.strip().str.lower()

        subj_n, dist_n, np_n = norm(COL_SUBJ_REF), norm(COL_DIST_REF), norm(COL_NP_REF)
        search = subj_n + " " + dist_n + " " + np_n

        df = pd.DataFrame(
            {
                "ss": search,
                "subj_n": subj_n,
                "np_n": np_n,
                "subj": ref[COL_SUBJ_REF],
                "dist": ref[COL_DIST_REF],
                "np": ref[COL_NP_REF],
                "oktmo": ref[COL_OKTMO_REF],
            }
        ).drop_duplicates(subset="ss")  # как в оригинале: оставляем первое вхождение

        self.keys: list[str] = df["ss"].tolist()
        self.np_norm: dict[str, str] = dict(zip(self.keys, df["np_n"].tolist()))
        self.lookup: dict[str, tuple] = {}
        for k, s, d, n, o in zip(
            self.keys, df["subj"].tolist(), df["dist"].tolist(), df["np"].tolist(), df["oktmo"].tolist()
        ):
            self.lookup[k] = (
                "" if pd.isna(s) else s,
                "" if pd.isna(d) else d,
                "" if pd.isna(n) else n,
                _fmt_oktmo(o),
            )

        sub = defaultdict(list)
        for k, s in zip(self.keys, df["subj_n"].tolist()):
            sub[s].append(k)
            words = s.split()
            if words:
                sub[words[0]].append(k)
            if "ханты" in s:
                sub["ханты"].append(k)
        # убираем повторы (порядок сохраняется -> результат extractOne тот же)
        self.subject_index: dict[str, list[str]] = {k: list(dict.fromkeys(v)) for k, v in sub.items()}


def _index_getstate(self):
    st = dict(self.__dict__)
    st.pop("_fb_cache", None)  # запасной индекс строится заново в каждом процессе, не передаём его по каналу
    return st


OktmoIndex.__getstate__ = _index_getstate


def load_oktmo_index(path) -> OktmoIndex:
    # dtype=str для ОКТМО: иначе pandas превращает текст «01701000001» в число 1701000001 и теряет ведущие нули
    return OktmoIndex(read_excel_fast(path, dtype={COL_OKTMO_REF: str}))


_SPECIAL = (
    ("москва", ("Москва", "Москва", "Москва", "45000000000")),
    ("санкт-петербург", ("Санкт-Петербург", "Санкт-Петербург", "Санкт-Петербург", "40000000000")),
    ("байконур", ("Байконур", "Байконур", "Байконур", "55000000000")),
    ("севастополь", ("Севастополь", "Севастополь", "Севастополь", "67000000000")),
)
_CITY_PATTERNS = ("г ", "г.", "город ")


def _extract_city(normalized: str) -> str:
    for pattern in _CITY_PATTERNS:
        if pattern in normalized:
            city_part = normalized.split(pattern)[1].split(",")[0].strip()
            words = city_part.split()
            for w in words:
                if not w.isdigit() and len(w) > 1:
                    return w
            return words[0] if words else ""
    return ""


class Matcher:
    """Поиск лучшего совпадения. Кэши живут внутри объекта (по одному на процесс)."""

    def __init__(self, index: OktmoIndex, threshold: int = 70):
        self.ix = index
        self.threshold = threshold
        self._pool_cache: dict[str, str | None] = {}
        self._variants: dict[tuple, list[str]] = {}
        self._fb: _Fallback | None = None

    def _pool_key(self, subject_norm: str):
        """Ключ пула кандидатов (None = весь справочник). Повторяет цепочку из recognizer.py."""
        if subject_norm in self._pool_cache:
            return self._pool_cache[subject_norm]
        si = self.ix.subject_index
        key = subject_norm if si.get(subject_norm) else None
        if key is None and " " in subject_norm:
            fw = subject_norm.split()[0]
            if si.get(fw):
                key = fw
        if key is None and "ханты" in subject_norm and si.get("ханты"):
            key = "ханты"
        self._pool_cache[subject_norm] = key
        return key

    def _fallback(self, norm: str):
        if self._fb is None:
            self._fb = getattr(self.ix, "_fb_cache", None) or _Fallback(self.ix)
            try:
                self.ix._fb_cache = self._fb  # в главном процессе индекс живёт долго — строим один раз
            except Exception:
                pass
        return self._fb.find(norm)

    def match(self, address: str):
        norm = normalize_text(address)
        if not norm:
            return None
        for needle, res in _SPECIAL:
            if needle in norm:
                return res

        ix = self.ix
        hit = ix.lookup.get(norm)
        if hit is not None:
            return hit

        if STRUCT_FIRST:
            row, sure = self._fallback(norm)
            if row is not None and sure:
                return row

        city = _extract_city(norm)
        subject_norm = normalize_text(norm.split(",")[0].strip()) if norm else ""
        pool_key = self._pool_key(subject_norm) if subject_norm else None
        pool = ix.subject_index[pool_key] if pool_key is not None else ix.keys

        if city:
            vk = (pool_key, city)
            variants = self._variants.get(vk)
            if variants is None:
                npn = ix.np_norm
                variants = [k for k in pool if city in npn[k]]
                self._variants[vk] = variants
            if variants:
                res = process.extractOne(
                    norm, variants, scorer=fuzz.token_set_ratio, score_cutoff=self.threshold - 10
                )
                if res:
                    return ix.lookup[res[0]]

        if pool:
            res = process.extractOne(norm, pool, scorer=fuzz.token_set_ratio, score_cutoff=self.threshold)
            if res:
                return ix.lookup[res[0]]
        return self._fallback(norm)[0]  # основной поиск ничего не нашёл — пробуем структурный



# ---------------------------------------------------------------------------
# Запасной структурный поиск
# ---------------------------------------------------------------------------
# Срабатывает ТОЛЬКО когда основной поиск (нечёткое сравнение всей строки) ничего не нашёл, поэтому уже
# найденные результаты он не меняет. Идея: разобрать адрес на части (субъект, район, населённый пункт с типом),
# найти НП в справочнике по точному названию внутри субъекта (ё = е, без типа «г/п/с/д…»), а похожие НП
# разных районов различить по названию района. Если вариантов несколько и отличить их нечем — не гадаем.
_NONALNUM = re.compile(r"[^а-яa-z0-9]+")
_BRACKETS = re.compile(r"\([^)]*\)|/[^/]*/")
_RE_GO = re.compile(r"\bг\s*\.\s*о\b\.?")
_RE_MR = re.compile(r"\bм\s*\.\s*р\s*-?\s*н\b\.?")
_RE_MO = re.compile(r"\bм\s*\.\s*о\b\.?")
_RE_SS = re.compile(r"\bс\s*/\s*с\b")
_RE_SP = re.compile(r"\bс\s*\.\s*п\b\.?")
_RE_GPT = re.compile(r"\bг\s*\.\s*п\b\.?|\bг\s+п\b\.?")

# Слова-типы населённых пунктов (как в начале/конце названия: «г Пермь», «Южный рп», «поселок городского типа Карабаш»)
_NP_TYPES = frozenset(
    "г город гор пгт рп п пос поселок посел городок дп с село д дер деревня х хутор ст станица ца аул аал "
    "нп сл слобода кордон свх совхоз совхоза городского типа рабочий дачный курортный".split()
)
_NAME_DROP = _NP_TYPES | {"им", "имени", "при", "станции", "станция", "разъезда", "разъезд", "железнодорожной", "железнодорожного", "железнодорожный", "железнодорожная"}
# Слова, по которым часть адреса считается названием района/муниципального образования, а не НП
_DIST_WORDS = frozenset(
    "район района округ округа сельсовет волость поселение межселенные муниципальный муниципального "
    "внутригородской внутригородская внутригородское образование".split()
)
_STOP_WORDS = frozenset("ул улица пер переулок пр проспект ш шоссе тер территория км дом стр строение кв литера лит корп".split())
_CITY_T = frozenset({"г", "город", "гор"})
_TYPE_CLASS = {}
for _cls, _ws in (
    ("city", "г город гор"),
    ("urban", "пгт рп п пос поселок посел городок дп"),
    ("village", "с село ст станица ца аул аал сл слобода"),
    ("hamlet", "д дер деревня х хутор нп свх совхоз совхоза"),
):
    for _w in _ws.split():
        _TYPE_CLASS[_w] = _cls


def _fold(s) -> str:
    """Нижний регистр, ё -> е, всё кроме букв и цифр -> пробел."""
    return _NONALNUM.sub(" ", str(s).lower().replace("ё", "е")).strip()


def _fold_series(s: pd.Series) -> list:
    return (
        s.fillna("").astype(str).str.lower().str.replace("ё", "е", regex=False)
        .str.replace(r"[^а-яa-z0-9]+", " ", regex=True).str.strip().tolist()
    )


def _core(tokens) -> str:
    return " ".join(t for t in tokens if t not in _NAME_DROP)


def _stem_eq(a: str, b: str) -> bool:
    """Слова совпадают с точностью до окончания («Воскресенск» ~ «Воскресенский», «Пермь» ~ «Пермский»),
    но «Красноглинский» и «Красноярский» — разные."""
    if a == b:
        return True
    n = 0
    for x, y in zip(a, b):
        if x != y:
            break
        n += 1
    return n >= max(4, min(len(a), len(b)) - 3)


class _Fallback:
    """Индексы справочника для структурного поиска. Строятся лениво (один раз на процесс)."""

    def __init__(self, ix: "OktmoIndex"):
        keys = ix.keys
        rows = [ix.lookup[k] for k in keys]
        self.rows = rows
        subj = _fold_series(pd.Series([r[0] for r in rows], dtype=object))
        dist = _fold_series(pd.Series([r[1] for r in rows], dtype=object))
        npf = _fold_series(pd.Series([r[2] for r in rows], dtype=object))
        self.subj_names = {}  # свёрнутое название субъекта -> оно же (для проверки принадлежности)
        self.dist_tok = []
        self.np_type = []
        self.by_core = defaultdict(list)    # (субъект, ядро названия НП) -> [номера строк]
        self.by_subj = defaultdict(list)    # субъект -> [номера строк]
        self.core_of = []
        for i, (sf, df_, nf) in enumerate(zip(subj, dist, npf)):
            toks = nf.split()
            core = _core(toks)
            self.subj_names[sf] = True
            self.core_of.append(core)
            self.np_type.append(frozenset(t for t in toks if t in _NP_TYPES))
            self.dist_tok.append(tuple(t for t in df_.split() if len(t) >= 3 and t not in _DIST_WORDS))
            if core:
                self.by_core[(sf, core)].append(i)
            self.by_subj[sf].append(i)
        self.subj_of = subj
        self._subj_cores: dict = {}

    # --- разбор адреса ----------------------------------------------------
    @staticmethod
    def _segments(address: str):
        out = []
        for seg in address.split(","):
            s = seg.lower()
            s = _RE_GO.sub(" городской округ ", s)
            s = _RE_MR.sub(" муниципальный район ", s)
            s = _RE_MO.sub(" муниципальный округ ", s)
            s = _RE_SS.sub(" сельсовет ", s)
            s = _RE_SP.sub(" сельское поселение ", s)
            s = _RE_GPT.sub(" городское поселение ", s)
            s = _BRACKETS.sub(" ", s)
            out.append(_fold(s).split())
        return out

    def _subject(self, segs):
        found = None
        for i, toks in enumerate(segs[:4]):
            if toks and " ".join(toks) in self.subj_names:
                found = (i, " ".join(toks))  # берём последний (вложенный: «Тюменская…, Ханты-Мансийский…»)
        return found

    @staticmethod
    def _name_tokens(toks):
        """Обрезает название на цифрах и «улично-домовых» словах."""
        out = []
        for t in toks:
            if t in _STOP_WORDS or any(c.isdigit() for c in t):
                break
            out.append(t)
        return out

    def _candidates(self, segs, subj_idx):
        """[(сила, номер части, ядро, типы, можно_обрезать)]; сила: 0 — с явным типом, 1 — «гп …», 2 — без типа."""
        cands = []
        for i, toks in enumerate(segs):
            if i == subj_idx or not toks:
                continue
            if any(t in _DIST_WORDS for t in toks):
                continue  # район/округ — подсказка, не НП
            lead = 0
            while lead < len(toks) and (toks[lead] in _NP_TYPES or toks[lead] == "гп"):
                lead += 1
            if lead:
                types = frozenset(toks[:lead]) & _NP_TYPES
                name = self._name_tokens(toks[lead:])
                weak = toks[0] == "гп"
                core = _core(name)
                if core and len(name[0]) > 1:
                    cands.append((1 if weak else 0, i, core, types, weak))
                continue
            tail = 0
            while tail < len(toks) and toks[-1 - tail] in _NP_TYPES:
                tail += 1
            if tail:  # «Южный рп», «Киров г»
                core = _core(self._name_tokens(toks[: len(toks) - tail]))
                if core:
                    cands.append((0, i, core, frozenset(toks[len(toks) - tail:]) & _NP_TYPES, False))
                continue
            mid = [k for k, t in enumerate(toks) if t in _NP_TYPES and 0 < k]
            if mid:  # «Управленческий п Гипсовый карьер»
                k = mid[0]
                core = _core(self._name_tokens(toks[:k]))
                if core:
                    cands.append((0, i, core, frozenset({toks[k]}), False))
                continue
            name = self._name_tokens(toks)
            core = _core(name)
            if core and len(name) == len(toks):
                cands.append((2, i, core, frozenset({"г"}), False))
        return cands

    # --- выбор строки справочника ----------------------------------------
    def _score(self, row, hints, qtypes):
        rt = self.dist_tok[row]
        sc = sum(1 for h in hints if any(_stem_eq(h, r) for r in rt))
        qc = {_TYPE_CLASS.get(t) for t in qtypes}
        tb = 1 if qc & {_TYPE_CLASS.get(t) for t in self.np_type[row]} else 0
        return (sc, tb)

    def _choose(self, rows, hints, qtypes):
        """(лучшая строка, число совпавших подсказок о районе); строка None, если лучших несколько."""
        if len(rows) == 1:
            return rows[0], self._score(rows[0], hints, qtypes)[0]
        scored = sorted(((self._score(r, hints, qtypes), r) for r in rows), reverse=True)
        if scored[0][0] == scored[1][0]:
            return None, 0
        return scored[0][1], scored[0][0][0]

    def _fuzzy_rows(self, subj, core):
        cores = self._subj_cores.get(subj)
        if cores is None:
            idx = self.by_subj.get(subj, [])
            cores = ([self.core_of[i] for i in idx], idx)
            self._subj_cores[subj] = cores
        names, idx = cores
        res = process.extract(core, names, scorer=fuzz.ratio, score_cutoff=88, limit=8)
        qt = [t for t in core.split() if len(t) >= 4]
        rows = []
        for _name, _sc, pos in res:
            rt = self.core_of[idx[pos]].split()
            if all(any(_stem_eq(q, r) for r in rt) for q in qt):
                rows.append(idx[pos])
        return rows

    def _subset_rows(self, subj, core, hints):
        """Строки, в названии НП которых есть все слова запроса целиком («Лащи» -> «железнодорожный разъезд Лащи»).
        Берём только подтверждённые названием района — иначе слишком много случайных совпадений."""
        cores = self._subj_cores.get(subj)
        if cores is None:
            idx = self.by_subj.get(subj, [])
            cores = ([self.core_of[i] for i in idx], idx)
            self._subj_cores[subj] = cores
        names, idx = cores
        q = set(core.split())
        out = []
        for _n, _sc, pos in process.extract(core, names, scorer=fuzz.token_set_ratio, score_cutoff=100, limit=30):
            row = idx[pos]
            if q <= set(self.core_of[row].split()) and any(
                _stem_eq(h, d) for h in hints for d in self.dist_tok[row]
            ):
                out.append(row)
        return out

    def _district_stem_rows(self, subj, core, dist_hints):
        """Строки субъекта в названном районе, где каждое слово запроса совпадает со словом названия НП с точностью до окончания."""
        q = [t for t in core.split() if len(t) >= 4]
        if not q:
            return []
        out = []
        for i in self.by_subj.get(subj, ()):
            if not any(_stem_eq(h, d) for h in dist_hints for d in self.dist_tok[i]):
                continue
            rt = self.core_of[i].split()
            if rt and all(any(_stem_eq(x, r) for r in rt) for x in q):
                out.append(i)
        return out

    def find(self, address: str):
        """(строка справочника | None, уверенно). «Уверенно» — НП с явным типом найден по точному названию
        и выбран однозначно (без нечёткого сравнения и без отката к более крупному НП)."""
        segs = self._segments(address)
        sj = self._subject(segs)
        if sj is None:
            return None, False
        first_strong = True
        subj_idx, subj = sj
        cands = sorted(self._candidates(segs, subj_idx), key=lambda c: (c[0], -c[1]))

        def hints_for(i):
            return [
                t for k, toks in enumerate(segs) if k not in (i, subj_idx)
                for t in toks if len(t) >= 3 and t not in _DIST_WORDS and t not in _NAME_DROP
            ]

        def resolve(core, types, strength, trim):
            """(строка, совпало подсказок, точное ли имя, имя целиком) | 'ambiguous' | None."""
            variants = [core]
            if trim:  # «гп Рощинский Военторг»: название муниципалитета + лишнее слово
                parts = core.split()
                variants += [" ".join(parts[:n]) for n in range(len(parts) - 1, 0, -1)]
            for v in variants:
                rows = self.by_core.get((subj, v))
                exact = bool(rows)
                if not rows and strength == 0 and not (types & _CITY_T) and len(v) >= 4:
                    rows = self._fuzzy_rows(subj, v)
                    if not rows:  # слова запроса целиком внутри названия + район подтверждён — считаем надёжным
                        rows = self._subset_rows(subj, v, hints_for_cur[0])
                        exact = bool(rows)
                if rows:
                    pick, dscore = self._choose(rows, hints_for_cur[0], types)
                    if pick is None:
                        return "ambiguous"
                    if dscore == 0 and explicit_dist[0]:
                        # район назван в адресе прямо, а найденный НП лежит в другом: ищем в названном районе
                        # НП с тем же названием с точностью до окончания («Нагорское» ~ «Нагорск»)
                        alt = self._district_stem_rows(subj, v, explicit_dist[0])
                        if len(alt) == 1:
                            return alt[0], 1, False, v == core
                    return pick, dscore, exact, v == core
            return None

        hints_for_cur = [None]
        explicit_dist = [[
            t for toks in segs if any(w in _DIST_WORDS for w in toks)
            for t in toks if len(t) >= 3 and t not in _DIST_WORDS and t not in _NAME_DROP
        ]]
        # 1) НП с явным типом: собираем все найденные и берём самый конкретный
        resolved = []
        fails = 0
        for strength, i, core, types, trim in cands:
            if strength != 0:
                break
            hints_for_cur[0] = hints_for(i)
            r = resolve(core, types, 0, False)
            if r == "ambiguous":
                if not resolved:
                    return None, False  # самый конкретный НП не отличить от двойников — не гадаем
                continue
            if r is None:
                fails += 1
                continue
            resolved.append((r, core, fails))
        if resolved:
            def inside(x):  # сколько других найденных НП названы в районе этого НП (значит, он внутри них)
                rt = self.dist_tok[x[0][0]]
                return sum(
                    1 for y in resolved if y is not x
                    and any(_stem_eq(t, d) for t in y[1].split() for d in rt)
                )
            best = max(resolved, key=lambda x: (inside(x), -resolved.index(x)))
            (pick, dscore, exact, whole), core, nfail = best
            return self.rows[pick], bool(exact and whole and nfail == 0)

        # 2) «гп …» и названия без типа — по одному, от конца адреса
        for strength, i, core, types, trim in cands:
            if strength == 0:
                continue
            hints_for_cur[0] = hints_for(i)
            r = resolve(core, types, strength, trim)
            if r == "ambiguous":
                return None, False
            if r is not None:
                pick, dscore, exact, whole = r
                return self.rows[pick], bool(exact and whole and fails == 0 and strength == 2 and dscore >= 1)
        return None, False


# --- функции для процессов -------------------------------------------------
_MATCHER: Matcher | None = None


def _init_matcher(index: OktmoIndex):
    global _MATCHER
    _MATCHER = Matcher(index)


def _clean_chunk(chunk: list[str]) -> list[str]:
    return [process_address(a) for a in chunk]


def _match_chunk(chunk: list[str]):
    m = _MATCHER
    return [m.match(a) for a in chunk]


class SharedPool:
    """Один пул процессов на оба этапа (очистка и распознавание).

    Запуск процессов (spawn) и передача справочника в каждый из них занимают секунды, поэтому
    пул создаётся лениво — только если какому-то этапу реально нужен параллелизм — и
    закрывается один раз в конце обработки файла.
    """

    def __init__(self, workers: int, index: OktmoIndex):
        self.workers, self.index, self._ex = workers, index, None

    def get(self):
        if self._ex is None:
            self._ex = ProcessPoolExecutor(
                max_workers=self.workers,
                mp_context=mp.get_context("spawn"),
                initializer=_init_matcher,
                initargs=(self.index,),
            )
        return self._ex

    def reset(self):  # после сбоя процессов — выбросить сломанный пул, при необходимости создадим новый
        if self._ex is not None:
            self._ex.shutdown(wait=False, cancel_futures=True)
            self._ex = None

    def close(self):
        if self._ex is not None:
            # cancel_futures: при отмене задания не ждём все оставшиеся части этапа
            self._ex.shutdown(wait=True, cancel_futures=True)
            self._ex = None


def run_parallel(func, items, workers, min_parallel, on_progress=None, initializer=None, initargs=(), weights=None, pool=None, stats=None):
    """Применяет func к чанкам items; возвращает плоский список результатов в исходном порядке.

    weights — сколько строк исходного файла соответствует каждому элементу (уникальному адресу);
    прогресс считается в строках файла, а не в уникальных значениях.
    pool — общий SharedPool; если не задан, пул создаётся и закрывается внутри вызова.
    stats — необязательный словарь для замеров: mode (serial/parallel/fallback), n_items, n_chunks,
            first_s — сколько секунд прошло до первого готового результата (в параллельном режиме сюда входит
            запуск процессов и передача им справочника).
    """
    n = len(items)
    if stats is not None:
        stats.update(mode="empty" if n == 0 else "", n_items=n, n_chunks=0, first_s=None)
    if n == 0:
        return []
    if weights is None:
        weights = [1] * n
    total_w = sum(weights) or 1
    prefix = [0]
    for w in weights:
        prefix.append(prefix[-1] + w)

    t_begin = time.time()

    def note_first():
        if stats is not None and stats["first_s"] is None:
            stats["first_s"] = time.time() - t_begin

    if workers <= 1 or n < min_parallel:
        if stats is not None:
            stats["mode"] = "serial"
        if initializer:
            initializer(*initargs)
        out, step = [], max(50, n // 100)
        for i in range(0, n, step):
            out.extend(func(items[i : i + step]))
            note_first()
            if on_progress:
                on_progress(prefix[min(i + step, n)], total_w)
        return out

    chunk = max(100, math.ceil(n / (workers * 24)))
    starts = list(range(0, n, chunk))
    chunks = [items[i : i + chunk] for i in starts]
    chunk_w = [prefix[min(i + chunk, n)] - prefix[i] for i in starts]
    if stats is not None:
        stats.update(mode="parallel", n_chunks=len(chunks))
    results: list = [None] * len(chunks)
    done = 0
    own = None
    try:
        if pool is not None:
            ex = pool.get()
        else:
            ex = own = ProcessPoolExecutor(
                max_workers=workers,
                mp_context=mp.get_context("spawn"),
                initializer=initializer,
                initargs=initargs,
            )
        futs = {ex.submit(func, c): i for i, c in enumerate(chunks)}
        for f in as_completed(futs):
            i = futs[f]
            results[i] = f.result()
            note_first()
            done += chunk_w[i]
            if on_progress:
                on_progress(done, total_w)
    except Exception as e:  # BrokenProcessPool и т.п. (например, процесс убит из-за нехватки памяти)
        if pool is not None:
            pool.reset()
        if isinstance(e, (ValueError, KeyError, TypeError)):
            raise
        if stats is not None:
            stats["mode"] = "fallback"
        # Досчитываем в этом процессе только те части, которые ещё не готовы (готовые не пересчитываем,
        # прогресс не откатывается назад). Раньше этап начинался заново — на сервере это очень долго.
        if initializer:
            initializer(*initargs)
        for i, c in enumerate(chunks):
            if results[i] is None:
                results[i] = func(c)
                done += chunk_w[i]
                if on_progress:
                    on_progress(done, total_w)
    finally:
        if own is not None:
            own.shutdown(wait=True, cancel_futures=True)
    return [r for part in results for r in part]


# ---------------------------------------------------------------------------
# Метрики (логика recognizer.py, но векторно)
# ---------------------------------------------------------------------------
_REMOVE_PREFIXES = [
    "аал", "аул", "г", "город", "гп", "городское поселение", "д", "дп", "деревня", "дер",
    "железнодорожная станция", "кордон", "м", "нп", "п", "поселок", "дачный поселок", "пгт",
    "п г т", "р.п.", "рп", "с", "село", "сл", "ст", "станица", "ст-ца", "х", "хутор", "свх",
]
_PREFIX_TOKENS = set()
for _t in _REMOVE_PREFIXES:
    _PREFIX_TOKENS.update(re.sub(r"[^а-яё]+", " ", _t.lower().replace("ё", "е")).split())


def _norm_letters_series(s: pd.Series) -> pd.Series:
    s = s.fillna("").astype(str).str.lower().str.replace("ё", "е", regex=False)
    return s.str.replace(r"[^а-я]+", " ", regex=True).str.strip()


def _clean_np(norm: str) -> str:
    return " ".join(t for t in norm.split() if t not in _PREFIX_TOKENS)


def compute_status(df: pd.DataFrame, src_col: str = COL_SRC):
    """Возвращает (recognized, correct, status) — status: 0 нет заливки, 1 зелёная, 2 красная."""
    recognized = (df["ОКТМО"].astype(str).str.strip() != "").to_numpy()
    np_norm = _norm_letters_series(df["НП"])
    np_clean = np_norm.map({v: _clean_np(v) for v in np_norm.unique()})
    loc_norm = _norm_letters_series(df[src_col]).tolist()
    correct = np.fromiter(
        (bool(n) and f" {n} " in f" {l} " for n, l in zip(np_clean.tolist(), loc_norm)),
        dtype=bool,
        count=len(df),
    )
    has_np = (df["НП"].astype(str) != "").to_numpy()
    has_loc = (df[src_col].astype(str) != "").to_numpy()
    status = np.where(has_np & has_loc & recognized, np.where(correct, 1, 2), 0)
    return recognized, correct & recognized, status


# ---------------------------------------------------------------------------
# Запись Excel (xlsxwriter, потоковый режим — быстро и экономно по памяти)
# ---------------------------------------------------------------------------


def build_xlsx(
    df: pd.DataFrame, status: np.ndarray, stats_rows: list[list], on_progress=None,
    src_col: str = COL_SRC, clean_col: str = COL_CLEAN, timing: dict | None = None,
) -> bytes:
    t_func = time.time()
    buf = io.BytesIO()
    # strings_to_*: False — не проверять каждую строку на «ссылку/формулу/число» (быстрее) и не превращать
    # адреса вида «http…», «=…» в гиперссылки/формулы (у Excel есть лимит ~65 тыс. гиперссылок)
    wb = xlsxwriter.Workbook(
        buf,
        {
            "constant_memory": True,
            "nan_inf_to_errors": True,
            "strings_to_urls": False,
            "strings_to_formulas": False,
            "strings_to_numbers": False,
        },
    )
    ws = wb.add_worksheet("Sheet1")
    f_head = wb.add_format({"bold": True, "border": 1})
    f_green = wb.add_format({"bg_color": "#92D050"})
    f_red = wb.add_format({"bg_color": "#FF0000"})
    f_center = wb.add_format({"align": "center"})

    cols = list(df.columns)
    # ширина колонок: адрес — 80, обработанный адрес — 60, субъект — 43, остальные — 25
    widths_by_name = {src_col: 80, clean_col: 60, "Субъект": 43}
    for c in cols:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            df[c] = df[c].astype(str)
    data = [df[c].astype(object).where(df[c].notna(), "").tolist() for c in cols]
    np_i, ok_i = cols.index("НП"), cols.index("ОКТМО")

    for i, c in enumerate(cols):
        ws.set_column(i, i, widths_by_name.get(c, 25))
    ws.write_row(0, 0, cols, f_head)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), len(cols) - 1)

    n = len(df)
    st = status.tolist()
    step = max(1000, n // 50)
    t_rows = time.time()
    if timing is not None:
        timing["write.prep"] = t_rows - t_func  # подготовка: списки значений, ширина колонок, заголовок
    for r, row in enumerate(zip(*data), start=1):
        ws.write_row(r, 0, row)
        s = st[r - 1]
        if s:
            ws.write_string(r, np_i, row[np_i], f_green if s == 1 else f_red)
        if row[ok_i] != "":
            ws.write_string(r, ok_i, str(row[ok_i]), f_center)
        if on_progress and r % step == 0:
            on_progress(r, n)

    t_close = time.time()
    if timing is not None:
        timing["write.rows"] = t_close - t_rows  # запись строк (по этому участку идёт индикатор)
    ws2 = wb.add_worksheet("Статистика")
    widths = [0] * len(stats_rows[0])
    for r, row in enumerate(stats_rows):
        for c, v in enumerate(row):
            ws2.write(r, c, v)
            widths[c] = max(widths[c], len(str(v)))
    for c, w in enumerate(widths):
        ws2.set_column(c, c, w + 2)
    wb.close()
    out = buf.getvalue()
    if timing is not None:
        timing["write.close"] = time.time() - t_close  # лист «Статистика», сжатие и сборка файла (индикатор в это время стоит на ~100%)
    return out


# ---------------------------------------------------------------------------
# Главный конвейер
# ---------------------------------------------------------------------------
@dataclass
class Result:
    xlsx: bytes
    stats: dict
    preview: pd.DataFrame
    preview_status: np.ndarray  # 0 — без заливки, 1 — зелёная, 2 — красная (для первых 100 строк)
    elapsed: float
    timings: dict = field(default_factory=dict)
    info: dict = field(default_factory=dict)  # условия замеров: строки, уникальные адреса, процессы, точки прогресса


# (ключ, название, доля общего прогресса)
# Веса здесь — для параллельного режима (workers >= 2). Калибровка на сервере, файл 343 627 строк
# (почти все адреса уникальны), workers=2: итого 289 с — чтение 2,23; очистка 53,18; распознавание 195,72;
# показатели 2,09; Excel 35,90 с.
# Сумма весов = 1. Названия этапов используются и в интерфейсе (app.py).
STAGES = [
    ("read", "Чтение файла", 0.008),
    ("clean", "Очистка и обогащение адресов", 0.184),
    ("match", "Распознавание населённых пунктов", 0.677),
    ("stats", "Расчёт показателей", 0.007),
    ("write", "Формирование Excel-файла", 0.124),
]

# Веса для однопроцессного режима (workers = 1): очистка и распознавание не ускоряются процессами,
# поэтому их доля больше, а доля чтения/записи — меньше. Калибровка на сервере, тот же файл, workers=1:
# итого 482 с — чтение 1,97; очистка 80,42; распознавание 362,30; показатели 1,85; Excel 35,58 с.
STAGE_WEIGHTS_SERIAL = {
    "read": 0.004,
    "clean": 0.167,
    "match": 0.751,
    "stats": 0.004,
    "write": 0.074,
}


def stage_weights(workers: int) -> dict:
    """Веса этапов для заданного числа процессов (сумма = 1)."""
    if workers <= 1:
        return dict(STAGE_WEIGHTS_SERIAL)
    return {k: w for k, _, w in STAGES}


def run_pipeline(file_obj, index: OktmoIndex, progress=None, workers: int | None = None) -> Result:
    """progress(frac, stage_title, done, total) — необязательный колбэк для индикатора.

    Принимает .xlsx и .xls (формат определяется по содержимому); результат — всегда .xlsx.
    """
    workers = workers or default_workers()
    pool = SharedPool(workers, index) if workers > 1 else None  # один пул на очистку и распознавание
    try:
        return _run_pipeline(file_obj, index, progress, workers, pool)
    finally:
        if pool is not None:
            pool.close()


def _run_pipeline(file_obj, index: OktmoIndex, progress, workers: int, pool) -> Result:
    t0 = time.time()
    timings: dict = {}
    weight = stage_weights(workers)
    base, _acc = {}, 0.0
    for _k, _t, _w in STAGES:
        base[_k] = _acc
        _acc += weight[_k]
    title = {k: t for k, t, _ in STAGES}

    stage_t0: dict = {}  # когда начался каждый этап
    marks: dict = {}     # "clean@50" -> сколько секунд от начала этапа прошло до 50% его прогресса
    info: dict = {
        "workers": workers, "cpus": cpu_limit(), "parallel": pool is not None, "marks": marks,
        "weights": weight,
    }

    def report(key, done, total):
        f = done / total if total else 1.0
        t_stage = stage_t0.get(key)
        if t_stage is not None:
            for m in (25, 50, 75):
                if f >= m / 100 and f"{key}@{m}" not in marks:
                    marks[f"{key}@{m}"] = time.time() - t_stage
        if progress:
            frac = base[key] + weight[key] * f
            progress(min(frac, 1.0), title[key], done, total)

    def stage_cb(key):
        return lambda d, t: report(key, d, t)

    # 1. Чтение
    try:  # размер файла в байтах (для замеров)
        file_obj.seek(0, 2)
        info["file_bytes"] = file_obj.tell()
        file_obj.seek(0)
    except Exception:
        info["file_bytes"] = None
    ts = time.time()
    stage_t0["read"] = ts
    report("read", 0, 1)
    df = read_excel_fast(file_obj)
    # Ищем колонку с адресами: «Местонахождение», «Адрес» или любая, содержащая «Адрес»
    src_col, clean_col = find_address_column(df.columns)
    if src_col is None:
        raise ValueError(
            "В файле не найдена колонка с адресами: её название должно быть «Местонахождение», «Адрес» "
            f"или содержать слово «Адрес». Найдены колонки: {', '.join(map(str, df.columns)) or '—'}."
        )
    if len(df) > EXCEL_MAX_ROWS:
        raise ValueError(f"В файле {len(df):,} строк — больше лимита Excel ({EXCEL_MAX_ROWS:,}).".replace(",", " "))
    df = df.reset_index(drop=True)
    src = df[src_col].fillna("").astype(str)
    timings["read"] = time.time() - ts
    report("read", 1, 1)

    # 2. Очистка (только уникальные адреса)
    ts = time.time()
    stage_t0["clean"] = ts
    uniq = src.unique().tolist()
    w_clean = src.value_counts().reindex(uniq).tolist()  # строк файла на каждый уникальный адрес
    t_run = time.time()
    timings["clean.pre"] = t_run - ts  # уникальные адреса и их веса
    clean_stats: dict = {}
    cleaned = run_parallel(
        _clean_chunk, uniq, workers, MIN_PARALLEL_CLEAN, stage_cb("clean"), weights=w_clean, pool=pool,
        stats=clean_stats,
    )
    t_post = time.time()
    timings["clean.run"] = t_post - t_run  # собственно очистка (в параллельном режиме — вместе с запуском процессов)
    df[clean_col] = src.map(dict(zip(uniq, cleaned)))
    timings["clean.post"] = time.time() - t_post  # раскладка результата по всем строкам файла
    timings["clean"] = time.time() - ts
    info["clean_run"] = clean_stats
    report("clean", 1, 1)

    # 3. Распознавание (только уникальные очищенные адреса, отсортированные для локальности кэша)
    ts = time.time()
    stage_t0["match"] = ts
    uniq_c = sorted(df[clean_col].unique().tolist())
    w_match = df[clean_col].value_counts().reindex(uniq_c).tolist()
    t_run = time.time()
    timings["match.pre"] = t_run - ts
    match_stats: dict = {}
    matched = run_parallel(
        _match_chunk, uniq_c, workers, MIN_PARALLEL_MATCH, stage_cb("match"),
        initializer=_init_matcher, initargs=(index,), weights=w_match, pool=pool, stats=match_stats,
    )
    t_post = time.time()
    timings["match.run"] = t_post - t_run
    mp_ = dict(zip(uniq_c, matched))
    empty = ("", "", "", "")
    res = df[clean_col].map(lambda a: mp_.get(a) or empty)
    out = pd.DataFrame(res.tolist(), columns=OUT_COLS, index=df.index)
    for c in OUT_COLS:
        df[c] = out[c]
    timings["match.post"] = time.time() - t_post
    timings["match"] = time.time() - ts
    info["match_run"] = match_stats
    report("match", 1, 1)

    # 4. Метрики
    ts = time.time()
    stage_t0["stats"] = ts
    recognized, correct, status = compute_status(df, src_col)
    total = len(df)
    rec_n, cor_n = int(recognized.sum()), int(correct.sum())
    rec_lvl = rec_n / total * 100 if total else 0
    acc_lvl = cor_n / total * 100 if total else 0
    stats = {
        "total": total,
        "recognized": rec_n,
        "unrecognized": total - rec_n,
        "correct": cor_n,
        "incorrect": rec_n - cor_n,
        "recognition_level": rec_lvl,
        "accuracy_level": acc_lvl,
        "address_column": str(src_col),   # из какой колонки взяты адреса
        "clean_column": str(clean_col),   # как названа колонка с обработанными адресами
        "unique_raw": len(uniq),
        "unique_clean": len(uniq_c),
    }
    stats_rows = [
        ["Показатель", "Значение"],
        ["Уровень распознавания (%)", f"{rec_lvl:.2f}"],
        ["Нераспознанных НП", total - rec_n],
        ["Уровень достоверности (%)", f"{acc_lvl:.2f}"],
        ["Возможно неверно распознанных НП", rec_n - cor_n],
    ]
    timings["stats"] = time.time() - ts
    report("stats", 1, 1)

    # 5. Excel
    ts = time.time()
    stage_t0["write"] = ts
    preview = df.head(100).copy()
    preview_status = status[:100].copy()
    report("write", 0, 1)
    xlsx = build_xlsx(df, status, stats_rows, stage_cb("write"), src_col=src_col, clean_col=clean_col, timing=timings)
    timings["write"] = time.time() - ts
    report("write", 1, 1)

    elapsed = time.time() - t0
    stats["elapsed"] = elapsed
    info.update(rows=total, unique_raw=len(uniq), unique_clean=len(uniq_c))
    return Result(
        xlsx=xlsx, stats=stats, preview=preview, preview_status=preview_status,
        elapsed=elapsed, timings=timings, info=info,
    )


# ---------------------------------------------------------------------------
# Отчёт о замерах (для калибровки весов этапов и оценки оставшегося времени)
# ---------------------------------------------------------------------------
def timing_report(res: Result):
    """Замеры одного запуска.

    Возвращает (rows, header, values):
      rows   — таблица для экрана: (этап, секунд, фактическая доля %, текущий вес в коде %);
      header — названия столбцов через табуляцию, values — значения через табуляцию
               (две строки удобно вставлять в Excel и присылать для расчёта новых весов).
    """
    t, info = res.timings, res.info
    used_w = info.get("weights") or {k: w for k, _, w in STAGES}  # веса, применённые в этом запуске
    stage_sum = sum(t.get(k, 0.0) for k, _, _ in STAGES) or 1e-9
    rows = [(title, t.get(k, 0.0), t.get(k, 0.0) / stage_sum * 100, used_w.get(k, 0.0) * 100) for k, title, _ in STAGES]

    cr, mr = info.get("clean_run") or {}, info.get("match_run") or {}
    fb = info.get("file_bytes")
    cols = [
        ("rows", info.get("rows")), ("unique_raw", info.get("unique_raw")), ("unique_clean", info.get("unique_clean")),
        ("workers", info.get("workers")), ("cpus", info.get("cpus")),
        ("file_mb", None if fb is None else fb / 1e6),
    ]
    cols += [(k, t.get(k)) for k, _, _ in STAGES] + [("total", res.elapsed)]
    cols += [(k, t.get(k)) for k in (
        "clean.pre", "clean.run", "clean.post", "match.pre", "match.run", "match.post",
        "write.prep", "write.rows", "write.close",
    )]
    cols += [("clean_mode", cr.get("mode")), ("clean_first_s", cr.get("first_s")),
             ("match_mode", mr.get("mode")), ("match_first_s", mr.get("first_s"))]
    marks = info.get("marks") or {}
    cols += [(f"{k}@{m}", marks.get(f"{k}@{m}")) for k in ("clean", "match", "write") for m in (25, 50, 75)]

    def fmt(v):
        if v is None:
            return ""
        return f"{v:.2f}" if isinstance(v, float) else str(v)

    return rows, "\t".join(k for k, _ in cols), "\t".join(fmt(v) for _, v in cols)