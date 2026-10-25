"""Streamlit-сайт: автоматическая простановка кодов ОКТМО по адресам из Excel.

Как устроен файл (сверху вниз):
  1. Константы: пути, цвета, SVG-иконка.
  2. CSS — весь внешний вид страницы (зона загрузки, кнопка, прогресс-бар, скрытый плеер).
  3. Вспомогательные функции (формат времени, таблица предпросмотра, вывод результата).
  4. main() — сама логика страницы: загрузка файла -> автообработка -> результат.
  5. Защитный вызов main() (нужен для Windows, см. комментарий в самом конце).
"""

from __future__ import annotations  # аннотации не вычисляются при импорте (нужно для условных импортов ниже)

import io
import multiprocessing as mp
import os
import time
from pathlib import Path

# Дочерние процессы (spawn) заново импортируют этот файл как «__mp_main__». Интерфейс им не нужен,
# поэтому тяжёлые библиотеки интерфейса (streamlit, pandas) в них не загружаем: это экономит память
# и секунды запуска каждого процесса.
# Признак надёжнее всего — имя модуля: при импорте в дочернем процессе mp.parent_process() ещё None.
_IS_CHILD = __name__ == "__mp_main__" or mp.parent_process() is not None
if not _IS_CHILD:
    import pandas as pd
    import streamlit as st
    from streamlit.runtime.scriptrunner import get_script_run_ctx

    import jobs  # фоновые задания: обработка не прерывается «перезапусками» страницы
    import pipeline  # здесь вся «математика»: очистка, распознавание, формирование Excel

# ---------------------------------------------------------------------------
# 1. КОНСТАНТЫ
# ---------------------------------------------------------------------------
BASE = Path(__file__).parent  # папка, где лежит app.py
# Справочник НП/ОКТМО. Путь можно переопределить переменной окружения OKTMO_FILE.
OKTMO_FILE = Path(os.environ.get("OKTMO_FILE", BASE / "НП_ОКТМО_01012025.xlsx"))
# Звук завершения. Если файла нет — просто тишина, ошибки не будет.
SOUND_FILE = BASE / "win-xp-sound.wav"

# >>> ИЗМЕНЕНО: единый акцентный цвет сайта (насыщенный синий).
#     Им красятся кнопка «СКАЧАТЬ ОБРАБОТАННЫЙ ФАЙЛ» и заливка прогресс-бара.
#     Хотите другой оттенок — меняйте только эти две строки.
ACCENT = "#2B60E2"        # основной синий
ACCENT_HOVER = "#1F4FC4"  # чуть темнее — для кнопки при наведении мыши

# >>> ДОБАВЛЕНО: бирюзовый цвет — названия и галочки завершённых этапов обработки
TEAL = "#34B1AA"

# >>> ДОБАВЛЕНО: цветные акценты карточек показателей (левая полоса и цифра)
#   распознавание — зелёный #28a745; нераспознанные — нейтральный серо-синий;
#   достоверность — бирюзовый; сомнительные — красный (как красная заливка НП в таблице)
METRIC_COLORS = {
    "recognition": "#28a745",  # >>> ИЗМЕНЕНО: был синий ACCENT
    "unrecognized": "#8795AE",
    "accuracy": TEAL,
    "doubtful": "#E5484D",
}

# >>> ДОБАВЛЕНО: подсказки при наведении на значок «?» у карточек (поддерживают Markdown)
METRIC_HELP = {
    "recognition": (
        "**Доля адресов, для которых найден населённый пункт** и проставлен код ОКТМО, "
        "от общего числа строк.\n\n"
        "Показывает, сколько адресов программа сумела обработать, но не гарантирует, что результат верный "
        "— для этого есть «Уровень достоверности»."
    ),
    "unrecognized": (
        "Адреса, для которых населённый пункт найти не удалось: ячейки «НП» и «ОКТМО» в файле "
        "остаются пустыми.\n\n"
        "Вместе с распознанными дают общее число строк."
    ),
    "accuracy": (
        "**Доля адресов, где найденный населённый пункт действительно встречается в исходном адресе** "
        "(зелёная заливка в колонке «НП»), от общего числа строк.\n\n"
        "Чем выше, тем меньше результатов придётся проверять вручную. "
    ),
    "doubtful": (
        "Код проставлен, но название найденного населённого пункта **не встречается в исходном адресе** "
        "(красная заливка в колонке «НП»).\n\n"
        "Часто результат всё равно верный — например, «Н. Ломов» → «Нижний Ломов», — "
        "но такие строки лучше проверить. "
        "Сомнительные = распознанные − достоверные."
    ),
}

# Цвета заливки колонки «НП» в таблице предпросмотра (те же, что в Excel-файле):
FILL_GREEN = "background-color: #92D050; color: #1f1f1f"  # НП найден в исходном адресе
FILL_RED = "background-color: #FF0000; color: #ffffff"    # НП не найден в исходном адресе

# SVG-иконка в центре зоны загрузки. Цвет «листа» в центре (%232B60E2) — это ACCENT,
# закодированный для URL (символ # записывается как %23).

ICON = (
    "data:image/svg+xml;utf8,"
    "<svg xmlns='http://www.w3.org/2000/svg' width='120' height='90' viewBox='0 0 120 90' fill='none' "
    "stroke='white' stroke-width='1.4'>"
    "<rect x='6' y='12' width='34' height='46' rx='2'/><path d='M12 24h22M12 32h22M12 40h22M20 12v46'/>"
    "<rect x='78' y='4' width='34' height='46' rx='2'/><circle cx='98' cy='28' r='8'/><path d='M98 28V20'/>"
    "<rect x='30' y='16' width='42' height='62' rx='2' fill='%232B60E2'/>"
    "<path d='M38 26h26M38 34h26M38 42h18'/>"
    
    # --- ДОБАВЛЕНА СТРЕЛКА СКАЧИВАНИЯ ---
    "<path d='M51 46v12m0 0l-4-4m4 4l4-4' stroke-linecap='round' stroke-linejoin='round'/>"
    
    "<rect x='36' y='64' width='22' height='10' fill='white' stroke='none'/>"
    "<text x='47' y='72' font-size='8' font-family='Arial' font-weight='bold' fill='%232B60E2' text-anchor='middle'>XLSX</text>"
    "</svg>"
)


# ---------------------------------------------------------------------------
# 2. CSS (внешний вид)
# Это f-строка, поэтому фигурные скобки CSS удвоены: {{ ... }}.
# Селекторы вида [data-testid="..."] — это «имена» элементов внутри Streamlit.
# ---------------------------------------------------------------------------
CSS = f"""
<style>
/* ===== Чистый сплошной фон без клеточек ===== */
.stApp {{
    background-color: #F4F6F9;
}}

/* Прячем стандартное меню Streamlit и подвал */
#MainMenu, footer {{visibility: hidden;}}

/* Максимальная ширина содержимого страницы и отступ сверху */
/* >>> ИЗМЕНЕНО: контейнер страницы растянут на всю ширину окна (max-width:none), а колонка содержимого шириной
       до 1500px получается боковыми отступами --side (на узких экранах --side = 1rem, как и было). Это нужно, чтобы футер
       мог дотянуться до краёв окна: его отрицательные поля равны --side. Нижний отступ 0 — футер вплотную к низу. */
.block-container {{
    max-width: none !important;
    --side: max(1rem, calc((100vw - 1500px) / 2));
    padding: 2.5rem var(--side) 0 var(--side) !important;
}}
/* Блок содержимого не ниже окна: на коротких страницах (например, до загрузки файла) футер опускается к низу окна */
.block-container > [data-testid="stVerticalBlock"] {{ min-height: calc(100vh - 2.5rem); }}

/* Заголовок и подзаголовок по центру */
h1.hero {{text-align:center; font-weight:800; font-size:2.6rem; color:#1f1f1f; margin:0 0 0 0;}}
h2.hero {{text-align:center; font-weight:600; font-size:2.0rem; color:#1f1f1f; margin:0 0 .4rem 0;}}
p.sub {{text-align:center; color:#666; margin-bottom:1.6rem;}}

/* ===== Синяя зона загрузки файла ===== */
[data-testid="stFileUploader"] section,
[data-testid="stFileUploaderDropzone"] {{
    background:#0A3E85 !important;                          /* фон зоны */
    border:2px dashed rgba(255,255,255,.85) !important;    /* белый пунктир внутри */
    outline:10px solid #1656B8; outline-offset:-1px;       /* широкая внешняя рамка */
    border-radius:10px;
    padding:3.2rem 1rem !important;
    display:flex; flex-direction:column; align-items:center; gap:1rem;
}}
/* Иконка над кнопкой (рисуется псевдоэлементом ::before) */
[data-testid="stFileUploaderDropzone"]::before {{
    content:""; width:120px; height:90px; background:url("{ICON}") center/contain no-repeat;
}}
/* Подсказка под кнопкой: прячем английский текст и подставляем свой через ::after */
[data-testid="stFileUploaderDropzoneInstructions"] {{ order:3; text-align:center; }}
[data-testid="stFileUploaderDropzoneInstructions"] * {{ display:none !important; }}
[data-testid="stFileUploaderDropzoneInstructions"]::after {{
    content:"или перетащите файл .xlsx / .xls сюда"; color:#e0efff; font-size:1rem;
}}
/* Белая кнопка выбора файла с русской подписью.
    Красим только stBaseButton-secondary, чтобы не задеть кнопки «удалить»/«добавить». */
[data-testid="stFileUploaderDropzone"] button[data-testid="stBaseButton-secondary"] {{
    order:2; background:#fff !important; color:#1f1f1f !important; border:none !important;
    font-size:0 !important;                 /* прячем исходную английскую подпись */
    padding:.9rem 2rem !important; border-radius:6px; font-weight:700;
}}
[data-testid="stFileUploaderDropzone"] button[data-testid="stBaseButton-secondary"]::after {{
    content:"ВЫБЕРИТЕ ФАЙЛ"; font-size:1rem; letter-spacing:.02em; color:#1f1f1f;
}}
[data-testid="stFileUploaderDropzone"] button[data-testid="stBaseButton-secondary"] * {{ display:none; }}
/* Кнопка «добавить ещё файл» не нужна: файл всегда один */
[data-testid="stFileUploaderDropzone"] button[aria-label="Add files"] {{ display:none !important; }}

/* ===== Кнопка «СКАЧАТЬ ОБРАБОТАННЫЙ ФАЙЛ» (Парящая) ===== */
.stElementContainer:has([data-testid="stDownloadButton"]) {{
    width:100% !important;
    display:flex !important;
    justify-content:center !important;
    margin: 20px 0;
}}
[data-testid="stDownloadButton"] {{
    display:flex; justify-content:center;
}}
/* Размер кнопки: 300 x 60 px, эффект парения (тень + анимация) */
[data-testid="stDownloadButton"] button {{
    width:300px !important; height:60px !important; font-weight:700;
    background-color:{ACCENT} !important; border:1px solid {ACCENT} !important; color:#fff !important;
    border-radius: 12px !important;
    box-shadow: 0 6px 20px rgba(43, 96, 226, 0.35) !important;
    transition: all 0.25s ease-in-out !important;
}}
/* Эффект при наведении: кнопка плавно приподнимается и тень становится объемнее */
[data-testid="stDownloadButton"] button:hover {{
    background-color:{ACCENT_HOVER} !important; border-color:{ACCENT_HOVER} !important; color:#fff !important;
    transform: translateY(-3px) !important;
    box-shadow: 0 10px 25px rgba(43, 96, 226, 0.5) !important;
}}
[data-testid="stDownloadButton"] button:active,
[data-testid="stDownloadButton"] button:focus {{
    background-color:{ACCENT_HOVER} !important; border-color:{ACCENT_HOVER} !important; color:#fff !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(43, 96, 226, 0.4) !important;
}}

/* ===== Прогресс-бар ===== */
[data-testid="stProgressBarTrack"] > div {{ background-color:{ACCENT} !important; }}

/* ===== Звук завершения ===== */
.stElementContainer:has([data-testid="stAudio"]), [data-testid="stAudio"] {{
    position:absolute !important; left:-9999px; width:1px; height:1px; overflow:hidden; opacity:0;
    margin:0 !important; pointer-events:none;
}}

/* ===== Карточки показателей ===== */
.st-key-m_recognition {{ --accent:{METRIC_COLORS["recognition"]}; }}
.st-key-m_unrecognized {{ --accent:{METRIC_COLORS["unrecognized"]}; }}
.st-key-m_accuracy {{ --accent:{METRIC_COLORS["accuracy"]}; }}
.st-key-m_doubtful {{ --accent:{METRIC_COLORS["doubtful"]}; }}
.st-key-metrics [data-testid="stMetric"] {{
    background:#fff; border:1px solid #E3E8F0; border-left:6px solid var(--accent, {ACCENT});
    border-radius:12px; padding:1rem 1.2rem; box-shadow:0 1px 4px rgba(16,24,40,.06);
    text-align:center; max-width:260px; margin-inline:auto;
}}
.st-key-metrics [data-testid="stMetricLabel"] {{
    display:flex !important; justify-content:center; align-items:center; gap:.4rem;
    color:#5b6472; font-weight:600; text-align:center;
}}
.st-key-metrics [data-testid="stMetricValue"] {{
    color:var(--accent, {ACCENT}); font-weight:800; text-align:center; justify-content:center;
}}
.st-key-metrics [data-testid="stMetricValue"] * {{ text-align:center; justify-content:center; }}

/* ===== Этапы обработки ===== */
.stages {{ display:flex; flex-wrap:wrap; justify-content:space-between; gap:.7rem 2rem; margin:.4rem 0 1.2rem 0; }}
.stage {{ display:flex; align-items:center; gap:.55rem; color:#9AA3B2; font-size:.97rem; }}
.stage .mark {{
    width:24px; height:24px; border-radius:50%; border:2px solid #C9D0DC; background:#fff;
    display:inline-flex; align-items:center; justify-content:center; font-size:.85rem; font-weight:800; flex:none;
}}
.stage.done {{ color:{TEAL}; font-weight:600; }}
.stage.done .mark {{ border-color:{TEAL}; color:{TEAL}; background:#E8F7F6; }}
.stage.active {{ color:#1f1f1f; font-weight:700; }}
.stage.active .mark {{ border-color:{TEAL}; }}
.stage.active .mark::after {{
    content:""; width:10px; height:10px; border-radius:50%; background:{TEAL}; animation:stagePulse 1.1s ease-in-out infinite;
}}
@keyframes stagePulse {{ 0%,100% {{ transform:scale(.55); opacity:.55; }} 50% {{ transform:scale(1); opacity:1; }} }}

/* ===== ФУТЕР ===== */
/* >>> ИЗМЕНЕНО: футер больше не «приклеен» к окну. Он стоит в потоке страницы ПОСЛЕДНИМ элементом
       (order:999), поэтому ничего не закрывает: при прокрутке появляется после всех данных.
       Если данных мало, margin-top:auto прижимает его к низу окна (см. min-height у блока содержимого выше).
       Во всю ширину окна растянут отрицательными полями, равными боковым отступам страницы (--side). */
.stElementContainer:has(.footer) {{ order: 999; margin-top: auto; }}
.footer {{
    margin: 2rem calc(-1 * var(--side, 1rem)) 0 calc(-1 * var(--side, 1rem));
    box-sizing: border-box;
    background-color: #002b5d;
    text-align: center;
    padding: 14px 20px;
    font-size: 15px;
    color: #cbd5e1;
    border-top: 1px solid #334155;
    font-family: var(--font-text); font-variant-numeric: lining-nums tabular-nums;
}}

/* «Контакты» — раскрывающийся блок (<details>): по клику показывает текст, по повторному — скрывает.
   Работает без JavaScript и без перезапуска страницы. */
.footer details {{ margin: 0; }}
.footer summary {{
    display: inline-block; cursor: pointer; user-select: none; list-style: none;
    color: #38bdf8; font-weight: 600; padding: 3px 12px; border-radius: 6px;
    background: rgba(56, 189, 248, 0.15); border: 1px solid rgba(56, 189, 248, 0.4);
    font-family: var(--font-ui);
}}
.footer summary::-webkit-details-marker {{ display: none; }}  /* прячем стандартный треугольник (Chrome/Safari) */
.footer summary::after {{ content: " ▾"; }}                   /* ▾ — закрыто */
.footer details[open] summary::after {{ content: " ▴"; }}      /* ▴ — открыто */
.footer summary:hover {{ background: rgba(56, 189, 248, 0.28); }}
.footer .contacts-text {{ margin-top: 10px; color: #e2e8f0; }}
</style>
<h1 class="hero">Простановка кодов ОКТМО</h1>
<h2 class="hero">Справочник ОКТМО актуален на 01.01.2025</h2>
<p class="sub">Загрузите Excel-файл (.xlsx или .xls) с колонкой адресов («Местонахождение», «Адрес» или любой заголовок со словом «Адрес») — адреса будут очищены,
распознаны и дополнены субъектом, районом, населённым пунктом и кодом ОКТМО.</p>
"""

# >>> ИЗМЕНЕНО: в футере остались только «Контакты» (<details>/<summary> — раскрывается по клику).
# Текст контактов меняется здесь, в блоке contacts-text.
# Строки внутри — без отступов в начале и без пустых строк: иначе Markdown примет их за блок кода.
FOOTER_HTML = """<div class="footer">
<details>
<summary>Контакты 📞</summary>
<div class="contacts-text">Селютин Егор Геннадьевич, Управление платежных систем и расчетов СГУ, 📞 ВТС (232)-1235, 📞 FMC 62-48-19</div>
</details>
</div>"""

# ---------------------------------------------------------------------------
# 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ---------------------------------------------------------------------------
def fmt_time(sec: float) -> str:
    """Секунды -> '5 мин 47 с' или '12 с'."""
    sec = int(sec)
    return f"{sec // 60} мин {sec % 60} с" if sec >= 60 else f"{sec} с"


def stages_html(active: int, all_done: bool = False) -> str:
    """HTML ряда этапов обработки.

    active   — номер текущего этапа (0..4), -1 — обработка ещё не началась (ждём очереди);
    all_done — все этапы завершены (все с галочкой).
    Этапы до текущего считаются завершёнными. Названия берутся из pipeline.STAGES.
    """
    items = []
    for i, (_key, title, _weight) in enumerate(pipeline.STAGES):
        if all_done or i < active:
            cls, mark = "done", "✓"
        elif i == active:
            cls, mark = "active", ""  # внутри — пульсирующая точка (рисуется CSS)
        else:
            cls, mark = "", ""
        items.append(f'<div class="stage {cls}"><span class="mark">{mark}</span><span>{title}</span></div>')
    return f'<div class="stages">{"".join(items)}</div>'


def styled_preview(res: pipeline.Result):
    """Таблица предпросмотра: ячейки колонки «НП» закрашены по достоверности распознавания.

    res.preview_status — список из 100 значений: 0 — без заливки, 1 — зелёная, 2 — красная.
    """
    df = res.preview.reset_index(drop=True)
    status = list(res.preview_status)

    def paint(col: pd.Series):
        # Для каждой ячейки колонки «НП» возвращаем CSS-строку с цветом (или пустую строку)
        css = {1: FILL_GREEN, 2: FILL_RED}
        return [css.get(s, "") for s in status[: len(col)]]

    # subset=["НП"] — красим только колонку «НП», остальные колонки не трогаем
    return df.style.apply(paint, subset=["НП"])


def show_result(res: pipeline.Result, out_name: str):
    """Выводит итог: сообщение, кнопку скачивания, показатели и таблицу предпросмотра."""
    s = res.stats
    st.success(f"Готово! Обработано {s['total']:,} строк за {fmt_time(res.elapsed)}.".replace(",", " "))
    # Какая колонка была распознана как адресная и как названа колонка с результатом очистки
    st.caption(
        f"Адреса взяты из колонки «{s['address_column']}»; "
        f"очищенные адреса записаны в колонку «{s['clean_column']}»."
    )

    # Кнопка скачивания. Её цвет и размер задаёт CSS выше; расположение по центру — тоже CSS.
    # width="content" — не растягивать кнопку на всю ширину (ширина 300px задана в CSS).
    st.download_button(
        "📥 СКАЧАТЬ ОБРАБОТАННЫЙ ФАЙЛ",
        data=res.xlsx,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="content",
    )

    # Четыре показателя в ряд — цветные карточки. Объяснение каждого — в подсказке (значок «?» у названия).
    # Контейнеры с key= нужны, чтобы CSS мог покрасить каждую карточку в свой цвет (см. METRIC_COLORS).
    with st.container(key="metrics"):
        c1, c2, c3, c4 = st.columns(4, gap="medium")  # >>> ИЗМЕНЕНО: gap="medium" — больше базовый отступ между колонками
        with c1.container(key="m_recognition"):
            st.metric("Уровень распознавания", f"{s['recognition_level']:.2f}%", help=METRIC_HELP["recognition"])
        with c2.container(key="m_unrecognized"):
            st.metric("Нераспознанных НП", f"{s['unrecognized']:,}".replace(",", " "), help=METRIC_HELP["unrecognized"])
        with c3.container(key="m_accuracy"):
            st.metric("Уровень достоверности", f"{s['accuracy_level']:.2f}%", help=METRIC_HELP["accuracy"])
        with c4.container(key="m_doubtful"):
            st.metric("Сомнительных НП", f"{s['incorrect']:,}".replace(",", " "), help=METRIC_HELP["doubtful"])
    
    st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)

    # Таблица первых 100 строк с цветной колонкой «НП»
    with st.expander("Предпросмотр результата (первые 100 строк)", expanded=True):
        st.dataframe(styled_preview(res), width="stretch", hide_index=True)
        st.caption(
            "Колонка «НП»: зелёный — название найдено в исходном адресе (достоверно), "
            "красный — не найдено; пустые ячейки без заливки. "
            "Тот же цвет — в скачанном файле. Показатели также на листе «Статистика»."
        )


# ---------------------------------------------------------------------------
# 4. ГЛАВНАЯ ЛОГИКА СТРАНИЦЫ
# ---------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Простановка кодов ОКТМО", page_icon="favicon.svg", layout="centered")
    st.markdown(CSS, unsafe_allow_html=True)  # подключаем стили и выводим заголовок
    # >>> ДОБАВЛЕНО: футер. Выводим сразу (до возможных st.stop() ниже); CSS ставит его последним элементом страницы.
    st.markdown(FOOTER_HTML, unsafe_allow_html=True)

    # Справочник ОКТМО читается один раз и кэшируется (повторные загрузки файлов его не перечитывают).
    # mtime в аргументах: если вы замените файл справочника — кэш обновится сам.
    @st.cache_resource(show_spinner="Загрузка справочника ОКТМО…")
    def get_index(path: str, mtime: float) -> pipeline.OktmoIndex:
        return pipeline.load_oktmo_index(path)

    # Без справочника работать нельзя — сообщаем и останавливаемся
    if not OKTMO_FILE.exists():
        st.error(f"Не найден справочник ОКТМО: {OKTMO_FILE.name}. Положите его рядом с app.py.")
        st.stop()

    # Диспетчер фоновых заданий — один на весь сервер (кэшируется между перезапусками страницы)
    @st.cache_resource
    def get_manager() -> jobs.JobManager:
        return jobs.JobManager()

    mgr = get_manager()
    # Идентификатор сессии браузера: нужен, чтобы результаты разных пользователей не смешивались
    ctx = get_script_run_ctx()
    session = ctx.session_id if ctx else "local"

    # Синяя зона загрузки. Подпись скрыта (label_visibility="collapsed"); принимаем .xlsx и .xls
    uploaded = st.file_uploader("Excel-файл", type=["xlsx", "xls"], label_visibility="collapsed")

    # Файла нет (или его удалили крестиком) — останавливаем и забываем задание, освобождаем память
    if uploaded is None:
        mgr.cancel_session(session)
        st.stop()

    mgr.purge()  # заодно убираем старые завершённые задания (экономим память сервера)

    # Ключ задания: сессия + имя + размер + внутренний id файла.
    # По нему отличаем «новый файл» от «страница перезапустилась, а работа уже идёт/закончена».
    job_id = f"{uploaded.name}:{uploaded.size}:{getattr(uploaded, 'file_id', '')}"
    key = f"{session}|{job_id}"
    out_name = f"{Path(uploaded.name).stem}_cleaned_processed.xlsx"  # имя скачиваемого файла

    job = mgr.get(key)
    if job is None:
        # Задания ещё нет (новый файл или сервер перезапускался) — запускаем обработку В ФОНЕ.
        try:
            index = get_index(str(OKTMO_FILE), OKTMO_FILE.stat().st_mtime)  # справочник (из кэша)
        except Exception as e:
            st.error(f"Не удалось загрузить справочник ОКТМО: {e}")
            st.stop()
        data = uploaded.getvalue()  # байты файла: поток не должен обращаться к объектам Streamlit

        def target(j):
            return pipeline.run_pipeline(io.BytesIO(data), index, progress=j.report)

        # Строка в логах сервера («Manage app» -> Logs): помогает понять, чем ограничен сервер
        print(f"[oktmo] старт: «{uploaded.name}», {uploaded.size / 1e6:.1f} МБ; {pipeline.resource_summary()}", flush=True)
        job = mgr.start(key, session, target)

    # Обработка ещё идёт: показываем прогресс. Если страница перезапустится, она подключится
    # к этому же заданию (работа не теряется и не начинается заново).
    was_running = not job.done
    if was_running:
        st.info(f"Файл «{uploaded.name}» загружен ({uploaded.size / 1024 / 1024:.1f} МБ). Идёт обработка…")
        bar = st.progress(0.0, text="Подготовка…")  # прогресс-бар (цвет — см. CSS)
        # >>> ДОБАВЛЕНО: под полосой — пять этапов; завершённые отмечены бирюзовой галочкой.
        stages_box = st.empty()
        titles = [t for _k, t, _w in pipeline.STAGES]
        shown = None  # какой набор этапов уже нарисован (перерисовываем только при смене этапа)
        while not job.done:
            if job.queued:
                bar.progress(0.0, text="В очереди: сервер сейчас обрабатывает другой файл…")
                active = -1
            else:
                frac, stage, done, total = job.progress
                active = titles.index(stage) if stage in titles else 0
                elapsed = time.time() - (job.started or time.time())
                # Оценка оставшегося времени по пройденной доле (показываем после первых 3%)
                eta = f" • осталось ≈ {fmt_time(elapsed / frac * (1 - frac))}" if frac > 0.03 else ""
                # Счётчик «обработано/всего» — только если этап считает строки
                detail = f" ({done:,}/{total:,})".replace(",", " ") if total > 1 else ""
                bar.progress(min(frac, 1.0), text=f"{stage}{detail} — {frac * 100:.0f}%{eta}")
            if active != shown:
                stages_box.markdown(stages_html(active), unsafe_allow_html=True)
                shown = active
            time.sleep(0.4)
        if job.error:
            bar.empty()
            stages_box.empty()
        else:
            bar.progress(1.0, text="Обработка завершена — 100%")
            stages_box.markdown(stages_html(-1, all_done=True), unsafe_allow_html=True)  # все этапы с галочками

    # Итог: ошибка или результат
    if job.error:
        st.error(job.error)
        st.stop()
    if job.result is None:  # задание было отменено/прервано без результата
        st.warning("Обработка была прервана. Загрузите файл ещё раз.")
        st.stop()

    show_result(job.result, out_name)

    # Звук завершения. Играет один раз — только если обработка закончилась на наших глазах
    # (при последующих перерисовках, например после «Скачать», звука нет). Плеер скрыт CSS.
    if was_running and SOUND_FILE.exists():
        try:
            st.audio(str(SOUND_FILE), autoplay=True)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 5. ЗАПУСК
# На Windows/macOS multiprocessing запускает дочерние процессы через spawn и заново импортирует
# этот файл под именем «__mp_main__». Строить интерфейс в дочерних процессах нельзя — иначе они
# падают (uploaded = None) и обработка зависает. Поэтому main() вызывается только в главном процессе.
# ---------------------------------------------------------------------------
if __name__ != "__mp_main__" and mp.parent_process() is None:
    main()