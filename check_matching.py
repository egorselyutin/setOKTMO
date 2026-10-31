"""Проверка поиска НП: сравнивает прежний алгоритм с новым на вашем файле и вашем справочнике.

Запуск (из папки с программой):
    python check_matching.py Adr_105.xlsx
    python check_matching.py Adr_105.xlsx НП_ОКТМО_01012025.xlsx

Результат — файл <имя>_check.xlsx: по строке на каждый адрес, с тремя вариантами ответа и пометкой:
  «новое»       — раньше НП не находился, теперь найден;
  «изменено»    — результаты разных вариантов отличаются (смотрите глазами — какой верный);
  «без изменений», «не найдено» — как есть.
Варианты:  «Было» — прежний алгоритм; «Стало» — новый (запасной поиск только если прежний ничего не нашёл);
           «Стало+» — новый, режим OKTMO_STRUCT_FIRST=1 (структурный поиск главнее при уверенном совпадении).
"""
import sys
from pathlib import Path

import pandas as pd

import pipeline as pl

HERE = Path(__file__).parent


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    src = Path(sys.argv[1])
    ref = Path(sys.argv[2]) if len(sys.argv) > 2 else HERE / "НП_ОКТМО_01012025.xlsx"
    print("Читаю справочник…", ref)
    ix = pl.load_oktmo_index(str(ref))
    df = pl.read_excel_fast(str(src))
    col, _ = pl.find_address_column(df.columns)
    if col is None:
        raise SystemExit("Не найдена колонка с адресами.")
    raw = df[col].fillna("").astype(str).tolist()
    clean = {a: pl.process_address(a) for a in dict.fromkeys(raw)}

    def run(mode):
        pl.STRUCT_FIRST = mode == "first"
        m = pl.Matcher(ix)
        if mode == "old":
            m._fallback = lambda norm: (None, False)  # без запасного поиска = прежнее поведение
        return {c: m.match(c) for c in dict.fromkeys(clean.values())}

    old, new, first = run("old"), run("new"), run("first")
    pl.STRUCT_FIRST = False

    def fmt(r):
        return "" if r is None else f"{r[0]} | {r[1]} | {r[2]} | {r[3]}"

    rows = []
    for a in raw:
        c = clean[a]
        o, n, f = old[c], new[c], first[c]
        if o is None and (n is not None or f is not None):
            mark = "новое"
        elif o is None:
            mark = "не найдено"
        elif fmt(o) != fmt(n) or fmt(o) != fmt(f):
            mark = "изменено"
        else:
            mark = "без изменений"
        rows.append({"Адрес": a, "Очищенный": c, "Было": fmt(o), "Стало": fmt(n), "Стало+": fmt(f), "Отметка": mark})
    out = pd.DataFrame(rows)
    path = src.with_name(src.stem + "_check.xlsx")
    out.to_excel(path, index=False)
    print(out["Отметка"].value_counts().to_string())
    print("Распознано: было", sum(old[clean[a]] is not None for a in raw), "| стало", sum(new[clean[a]] is not None for a in raw),
          "| стало+", sum(first[clean[a]] is not None for a in raw), "из", len(raw))
    print("Файл:", path)


if __name__ == "__main__":
    main()
