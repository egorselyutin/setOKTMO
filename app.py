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
/* Прячем стандартное меню Streamlit и подвал */
#MainMenu, footer {{visibility: hidden;}}

/* Максимальная ширина содержимого страницы и отступ сверху */
.block-container {{max-width: 1500px; padding-top: 2.5rem;}}

/* Заголовок и подзаголовок по центру */
h1.hero {{text-align:center; font-weight:800; font-size:2.6rem; color:#1f1f1f; margin:0 0 0 0;}}
h2.hero {{text-align:center; font-weight:600; font-size:2.0rem; color:#1f1f1f; margin:0 0 .4rem 0;}}
p.sub {{text-align:center; color:#666; margin-bottom:1.6rem;}}

/* ===== Синяя зона загрузки файла ===== */
[data-testid="stFileUploader"] section,
[data-testid="stFileUploaderDropzone"] {{
    background:#0A3E85 !important;                         /* фон зоны */
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

/* ===== Кнопка «СКАЧАТЬ ОБРАБОТАННЫЙ ФАЙЛ» ===== */
/* >>> ИЗМЕНЕНО (центрирование): контейнер кнопки растягиваем на всю ширину страницы
       и выравниваем содержимое по центру. Саму кнопку при этом не растягиваем. */
.stElementContainer:has([data-testid="stDownloadButton"]) {{
    width:100% !important;
    display:flex !important;
    justify-content:center !important;
}}
[data-testid="stDownloadButton"] {{
    display:flex; justify-content:center;
}}
/* Размер кнопки: 300 x 60 px */
/* >>> ИЗМЕНЕНО (цвет): фон и рамка кнопки — ACCENT (#2B60E2), белый текст.
       !important нужен, чтобы перебить красный цвет из темы Streamlit. */
[data-testid="stDownloadButton"] button {{
    width:300px !important; height:60px !important; font-weight:700;
    background-color:{ACCENT} !important; border:1px solid {ACCENT} !important; color:#fff !important;
}}
/* Цвет кнопки при наведении и нажатии мыши — чуть темнее */
[data-testid="stDownloadButton"] button:hover,
[data-testid="stDownloadButton"] button:active,
[data-testid="stDownloadButton"] button:focus {{
    background-color:{ACCENT_HOVER} !important; border-color:{ACCENT_HOVER} !important; color:#fff !important;
}}

/* ===== Прогресс-бар ===== */
/* >>> ИЗМЕНЕНО: заливка полосы (внутренний div внутри «дорожки» stProgressBarTrack) — ACCENT.
       Раньше селектор целился в другие вложенные div и не срабатывал, поэтому полоса была красной. */
[data-testid="stProgressBarTrack"] > div {{ background-color:{ACCENT} !important; }}

/* ===== Звук завершения ===== */
/* Плеер уводим за пределы экрана и делаем 1x1 px: звук играет сам, а самого плеера не видно */
.stElementContainer:has([data-testid="stAudio"]), [data-testid="stAudio"] {{
    position:absolute !important; left:-9999px; width:1px; height:1px; overflow:hidden; opacity:0;
    margin:0 !important; pointer-events:none;
}}
</style>
<h1 class="hero">Простановка кодов ОКТМО</h1>
<h2 class="hero">Справочник ОКТМО актуален на 01.01.2025</h2>
<p class="sub">Загрузите Excel-файл (.xlsx или .xls) с колонкой адресов («Местонахождение», «Адрес» или любой заголовок со словом «Адрес») — адреса будут очищены,
распознаны и дополнены субъектом, районом, населённым пунктом и кодом ОКТМО.</p>
"""


# ---------------------------------------------------------------------------
# 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ---------------------------------------------------------------------------
def fmt_time(sec: float) -> str:
    """Секунды -> '5 мин 47 с' или '12 с'."""
    sec = int(sec)
    return f"{sec // 60} мин {sec % 60} с" if sec >= 60 else f"{sec} с"


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

    # Четыре показателя в ряд
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Уровень распознавания", f"{s['recognition_level']:.2f}%")
    c2.metric("Нераспознанных НП", f"{s['unrecognized']:,}".replace(",", " "))
    c3.metric("Уровень достоверности", f"{s['accuracy_level']:.2f}%")
    c4.metric("Возможно неверно распознанных НП", f"{s['incorrect']:,}".replace(",", " "))

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
    st.set_page_config(page_title="Простановка кодов ОКТМО", page_icon="📍", layout="centered")
    st.markdown(CSS, unsafe_allow_html=True)  # подключаем стили и выводим заголовок

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
        while not job.done:
            if job.queued:
                bar.progress(0.0, text="В очереди: сервер сейчас обрабатывает другой файл…")
            else:
                frac, stage, done, total = job.progress
                elapsed = time.time() - (job.started or time.time())
                # Оценка оставшегося времени по пройденной доле (показываем после первых 3%)
                eta = f" • осталось ≈ {fmt_time(elapsed / frac * (1 - frac))}" if frac > 0.03 else ""
                # Счётчик «обработано/всего» — только если этап считает строки
                detail = f" ({done:,}/{total:,})".replace(",", " ") if total > 1 else ""
                bar.progress(min(frac, 1.0), text=f"{stage}{detail} — {frac * 100:.0f}%{eta}")
            time.sleep(0.4)
        if job.error:
            bar.empty()
        else:
            bar.progress(1.0, text="Обработка завершена — 100%")

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