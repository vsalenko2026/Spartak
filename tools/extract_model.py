# -*- coding: utf-8 -*-
"""Экстрактор финмодели: .xlsx -> карта, которую можно читать как код.

    python3 tools/extract_model.py                     полная выгрузка в build/
    python3 tools/extract_model.py --cell 02_Модель!B12 формула, от кого зависит,
                                                        кто зависит от неё
    python3 tools/extract_model.py --sheet 07_Мерч      карта одного листа
    python3 tools/extract_model.py --grep "роялти"      где встречается текст

Выгрузка:
    build/model_map.json          всё машиночитаемо
    build/sheets/<лист>.txt       формулы листа, свёрнутые по протяжке
    build/model_map.md            оглавление и сводка

Ключевая идея свёртки: протянутая по месяцам формула в R1C1 у всех ячеек
строки одинаковая. Печатаем её один раз как «B5:BI5», а любое отклонение
внутри строки — отдельной строкой с пометкой. Так разрыв в протяжке видно
глазом, а карта помещается в контекст.
"""
import argparse
import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlsx_lib as X

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_MODEL = ROOT / "model" / "Spartak_V8_9_понятная.xlsx"
BUILD = ROOT / "build"


def load(path: Path):
    wb = openpyxl.load_workbook(path, data_only=False)
    wbv = openpyxl.load_workbook(path, data_only=True)   # кэш значений
    return wb, wbv


def cell_kind(cell) -> str:
    v = cell.value
    if v is None:
        return "empty"
    if isinstance(v, str) and v.startswith("="):
        return "formula"
    if isinstance(v, str):
        return "text"
    return "const"


def scan(wb, wbv):
    """Собирает все непустые ячейки книги."""
    known_names = set(wb.defined_names.keys())
    sheets = {}
    for ws in wb.worksheets:
        wsv = wbv[ws.title]
        cells = {}
        for row in ws.iter_rows():
            for c in row:
                kind = cell_kind(c)
                if kind == "empty":
                    continue
                addr = c.coordinate
                rec = {"kind": kind, "row": c.row, "col": c.column}
                if kind == "formula":
                    f = c.value
                    rec["f"] = f
                    rec["r1c1"] = X.to_r1c1(f, c.column, c.row)
                    rec["refs"] = X.parse_refs(f, ws.title)
                    rec["names"] = X.parse_names(f, known_names)
                    cached = wsv[addr].value
                    if cached is not None:
                        rec["v"] = cached if not isinstance(cached, (bytes,)) else str(cached)
                else:
                    rec["v"] = c.value if not hasattr(c.value, "isoformat") else c.value.isoformat()
                if c.number_format and c.number_format != "General":
                    rec["fmt"] = c.number_format
                if c.comment is not None:
                    rec["note"] = c.comment.text
                cells[addr] = rec
        sheets[ws.title] = {
            "title": ws.title,
            "state": ws.sheet_state,
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "cells": cells,
        }
    names = {}
    for name, dn in wb.defined_names.items():
        try:
            names[name] = dn.value
        except Exception:
            names[name] = "?"
    return sheets, names


def json_safe(v):
    if v is None or isinstance(v, (int, float, str, bool)):
        return v
    return str(v)


NUM_LIT_RE = re.compile(r"(?<![A-Za-z_$])\d+(?:\.\d+)?")


def template(r1c1: str) -> str:
    """R1C1 с числовыми литералами, заменёнными на #. Строка вида
    =DATE(...+0,1), =DATE(...+1,1) ... даёт один шаблон."""
    return NUM_LIT_RE.sub("#", X.strip_strings(r1c1))


def fold_row(sheet, row_cells):
    """Свёртка строки: подряд идущие ячейки с одинаковой R1C1-формулой
    объединяются в диапазон. Возвращает список (диапазон, формула, отклонение)."""
    items = sorted(row_cells, key=lambda a: column_index_from_string(re.match(r"[A-Z]+", a).group(0)))
    groups, cur = [], None
    for addr in items:
        rec = sheet["cells"][addr]
        if rec["kind"] == "formula":
            sig = rec.get("r1c1")
            tpl = template(rec.get("r1c1", ""))
        else:
            sig = tpl = f"\0const:{rec.get('v')!r}"
        ci = rec["col"]
        same = cur and ci == cur["last_col"] + 1 and (
            cur["sig"] == sig or (cur["tpl"] == tpl and rec["kind"] == "formula"))
        if same:
            cur["last"], cur["last_col"], cur["n"] = addr, ci, cur["n"] + 1
            cur["last_rec"] = rec
            if cur["sig"] != sig:
                cur["varies"] = True
        else:
            cur = {"sig": sig, "tpl": tpl, "first": addr, "last": addr, "last_col": ci,
                   "n": 1, "rec": rec, "last_rec": rec, "varies": False}
            groups.append(cur)
    return groups


def render_sheet(sheet) -> str:
    """Человекочитаемая карта листа."""
    out = [f"# {sheet['title']}   ({sheet['max_row']} строк x {sheet['max_col']} столбцов, "
           f"{len(sheet['cells'])} заполненных ячеек)", ""]
    by_row = defaultdict(list)
    for addr, rec in sheet["cells"].items():
        by_row[rec["row"]].append(addr)

    for r in sorted(by_row):
        groups = fold_row(sheet, by_row[r])
        # подпись строки — первый текст слева
        label = ""
        for g in groups:
            if g["rec"]["kind"] == "text" and isinstance(g["rec"].get("v"), str):
                label = g["rec"]["v"].strip()
                break
        head = f"{r:>5} | {label[:60]}"
        parts = []
        for g in groups:
            rec = g["rec"]
            span = g["first"] if g["n"] == 1 else f"{g['first']}:{g['last']}"
            if rec["kind"] == "formula":
                val = rec.get("v")
                vs = f"  -> {val}" if isinstance(val, (int, float)) else ""
                parts.append(f"      {span:<14} {rec['f']}{vs}")
                if g["varies"]:
                    parts.append(f"      {'':<14} ...последняя ({g['last']}): {g['last_rec']['f']}")
            elif rec["kind"] == "const":
                parts.append(f"      {span:<14} = {rec.get('v')}")
            elif rec["kind"] == "text" and rec.get("v") != label:
                parts.append(f"      {span:<14} «{rec.get('v')}»")
        out.append(head)
        out.extend(parts)
    return "\n".join(out)


def build_dependents_index(sheets):
    """Обратный индекс: (лист, ref-потребитель) — по нему ищем, кто ссылается."""
    idx = defaultdict(list)          # sheet -> [(addr_потребителя, sheet_потр, ref)]
    for sname, sheet in sheets.items():
        for addr, rec in sheet["cells"].items():
            if rec["kind"] != "formula":
                continue
            for ref in rec.get("refs", []):
                idx[ref["sheet"]].append((sname, addr, ref["ref"]))
    return idx


def find_dependents(sheets, idx, target_sheet, addr):
    m = re.match(r"([A-Z]+)(\d+)", addr)
    col, row = m.group(1), int(m.group(2))
    sh = sheets.get(target_sheet, {})
    out = []
    for sname, caddr, ref in idx.get(target_sheet, []):
        if X.covers(ref, col, row, sh.get("max_row", 0), sh.get("max_col", 0)):
            out.append((sname, caddr, ref))
    return out


def cmd_dump(sheets, names, model_path):
    BUILD.mkdir(exist_ok=True)
    (BUILD / "sheets").mkdir(exist_ok=True)

    payload = {
        "model": str(model_path.relative_to(ROOT)),
        "generated": datetime.now().isoformat(timespec="seconds"),
        "defined_names": names,
        "sheets": {},
    }
    total_cells = total_f = 0
    for sname, sheet in sheets.items():
        cells = {}
        nf = 0
        for addr, rec in sheet["cells"].items():
            r = {k: v for k, v in rec.items() if k != "r1c1"}
            if "v" in r:
                r["v"] = json_safe(r["v"])
            cells[addr] = r
            if rec["kind"] == "formula":
                nf += 1
        payload["sheets"][sname] = {
            "state": sheet["state"], "max_row": sheet["max_row"],
            "max_col": sheet["max_col"], "formulas": nf,
            "filled": len(cells), "cells": cells,
        }
        total_cells += len(cells)
        total_f += nf
        safe = re.sub(r"[^0-9A-Za-zА-Яа-яЁё_-]", "_", sname)
        (BUILD / "sheets" / f"{safe}.txt").write_text(render_sheet(sheet), encoding="utf-8")

    (BUILD / "model_map.json").write_text(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    md = [f"# Карта модели — {model_path.name}", "",
          f"Собрано {datetime.now():%d.%m.%Y %H:%M}. "
          f"{len(sheets)} листов, {total_cells:,} заполненных ячеек, "
          f"{total_f:,} формул.".replace(",", " "), "",
          "| Лист | Строк | Столбцов | Заполнено | Формул |",
          "|---|---:|---:|---:|---:|"]
    for sname, sheet in sheets.items():
        nf = sum(1 for r in sheet["cells"].values() if r["kind"] == "formula")
        md.append(f"| {sname} | {sheet['max_row']} | {sheet['max_col']} | "
                  f"{len(sheet['cells'])} | {nf} |")
    md += ["", f"## Именованные диапазоны ({len(names)})", ""]
    for n, v in sorted(names.items()):
        md.append(f"- `{n}` → `{v}`")
    md += ["", "## Файлы", "",
           "- `build/model_map.json` — вся книга машиночитаемо",
           "- `build/sheets/<лист>.txt` — формулы листа, свёрнутые по протяжке"]
    (BUILD / "model_map.md").write_text("\n".join(md), encoding="utf-8")

    print(f"листов          : {len(sheets)}")
    print(f"ячеек           : {total_cells:,}".replace(",", " "))
    print(f"формул          : {total_f:,}".replace(",", " "))
    print(f"имён            : {len(names)}")
    print(f"записано в      : {BUILD}")


def cmd_cell(sheets, names, target):
    if "!" not in target:
        sys.exit("нужно вида 02_Модель!B12")
    sname, addr = target.rsplit("!", 1)
    sname = X.unquote_sheet(sname)
    if sname not in sheets:
        sys.exit(f"нет листа {sname}. Есть: {', '.join(sheets)}")
    rec = sheets[sname]["cells"].get(addr)
    print(f"=== {sname}!{addr} ===")
    if not rec:
        print("ячейка пуста")
    else:
        print(f"тип       : {rec['kind']}")
        if rec["kind"] == "formula":
            print(f"формула   : {rec['f']}")
            print(f"R1C1      : {rec['r1c1']}")
        if "v" in rec:
            print(f"значение  : {rec['v']}")
        if "fmt" in rec:
            print(f"формат    : {rec['fmt']}")
        if "note" in rec:
            print(f"примечание: {rec['note'][:400]}")
        if rec.get("names"):
            print(f"имена     : {', '.join(rec['names'])}")
        if rec.get("refs"):
            print("зависит от:")
            for r in rec["refs"]:
                print(f"   {r['sheet']}!{r['ref']}")
    idx = build_dependents_index(sheets)
    deps = find_dependents(sheets, idx, sname, addr)
    print(f"на неё ссылаются ({len(deps)}):")
    for s, a, ref in deps[:60]:
        print(f"   {s}!{a}   (через {ref})")
    if len(deps) > 60:
        print(f"   ... ещё {len(deps) - 60}")


def cmd_sheet(sheets, sname):
    if sname not in sheets:
        sys.exit(f"нет листа {sname}. Есть: {', '.join(sheets)}")
    print(render_sheet(sheets[sname]))


def cmd_grep(sheets, pattern):
    rx = re.compile(pattern, re.I)
    n = 0
    for sname, sheet in sheets.items():
        for addr, rec in sheet["cells"].items():
            hay = " ".join(str(rec.get(k, "")) for k in ("f", "v", "note"))
            if rx.search(hay):
                what = rec.get("f") or rec.get("v")
                print(f"{sname}!{addr:<8} {str(what)[:150]}")
                n += 1
    print(f"--- найдено {n}")


def main():
    ap = argparse.ArgumentParser(description="Экстрактор финмодели «Спартак»")
    ap.add_argument("--model", default=str(DEFAULT_MODEL))
    ap.add_argument("--cell", help="показать ячейку и её связи, напр. 02_Модель!B12")
    ap.add_argument("--sheet", help="карта одного листа")
    ap.add_argument("--grep", help="искать текст по формулам, значениям и примечаниям")
    a = ap.parse_args()

    path = Path(a.model)
    if not path.is_absolute():
        path = ROOT / path
    wb, wbv = load(path)
    sheets, names = scan(wb, wbv)

    if a.cell:
        cmd_cell(sheets, names, a.cell)
    elif a.sheet:
        cmd_sheet(sheets, a.sheet)
    elif a.grep:
        cmd_grep(sheets, a.grep)
    else:
        cmd_dump(sheets, names, path)


if __name__ == "__main__":
    main()
