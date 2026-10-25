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
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


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


def load_oktmo_index(path) -> OktmoIndex:
    return OktmoIndex(read_excel_fast(path))


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
        return None


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