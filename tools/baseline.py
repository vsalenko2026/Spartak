# -*- coding: utf-8 -*-
"""Эталон ключевых цифр модели.

Смысл простой: после любой правки цифры либо не двинулись, либо двинулись
осознанно — и тогда эталон обновляется вместе с объяснением, что именно
их сдвинуло.

    python3 tools/baseline.py            показать текущий эталон и расхождения
    python3 tools/baseline.py --update   принять новые цифры
"""
import argparse
import json
import subprocess
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BASELINE = ROOT / "tests" / "baseline.json"
RUNS = ROOT / "build" / "runs"

# что считаем эталонным: показатель -> (блок, подпись, как сводить)
KEYS = [
    ("Выручка — всего, без НДС", "модель", "Выручка — всего", "сумма"),
    ("EBITDA", "модель", "EBITDA", "сумма"),
    ("Прибыль после УСН", "модель", "Прибыль после УСН", "сумма"),
    ("Деньги на конец горизонта", "модель", "Денежный остаток на конец", "последнее"),
    ("Минимальный остаток денег", "модель", "Денежный остаток на конец", "минимум"),
    ("Активных детей на конец", "модель", "Всего активных детей", "последнее"),
    ("Новых школ на конец", "модель", "Новые школы — активные школы", "последнее"),
    ("Активных юрлиц на конец", "модель", "Активные юрлица", "последнее"),
    ("Паушальные, без НДС", "модель", "Паушальные взносы", "сумма"),
    ("Роялти новых школ, без НДС", "модель", "Роялти новых школ", "сумма"),
    ("Мерч, без НДС", "модель", "Мерч, без НДС", "сумма"),
    ("Выплаты ФК — роялти", "модель", "ФК — роялти", "сумма"),
    ("Выплаты ФК — паушальные", "модель", "ФК — паушальные", "сумма"),
    ("Погашение займа", "модель", "Погашение займа", "сумма"),
]


def series(snapshot, block, label):
    rows = snapshot["blocks"][block]["rows"]
    want = label.lower().strip()
    best = None
    for r, rec in sorted(rows.items(), key=lambda kv: int(kv[0])):
        got = (rec["label"] or "").lower().strip()
        if got == want:
            best = rec
            break
        if best is None and got.startswith(want):
            best = rec
    if best is None:
        raise SystemExit(f"нет строки «{label}» в блоке «{block}»")
    return [v if isinstance(v, (int, float)) else 0.0 for v in best["v"]]


def collect(tag="база"):
    path = RUNS / f"{tag}.json"
    if not path.exists():
        subprocess.run([sys.executable, str(ROOT / "tools" / "run_model.py"),
                        "--tag", tag, "--keep"], check=True, cwd=str(ROOT))
    snap = json.loads(path.read_text(encoding="utf-8"))
    out = {}
    for name, block, label, how in KEYS:
        vals = series(snap, block, label)
        out[name] = {"сумма": sum(vals), "последнее": vals[-1],
                     "минимум": min(vals)}[how]
    return out, snap


def main():
    ap = argparse.ArgumentParser(description="Эталон ключевых цифр модели")
    ap.add_argument("--update", action="store_true", help="принять новые цифры")
    ap.add_argument("--tag", default="база")
    a = ap.parse_args()

    now, snap = collect(a.tag)
    old = json.loads(BASELINE.read_text(encoding="utf-8"))["показатели"] \
        if BASELINE.exists() else {}

    print(f"{'показатель':<32}{'эталон':>20}{'сейчас':>20}{'дельта':>18}")
    сдвинулись = []
    for name in now:
        было = old.get(name)
        стало = now[name]
        if было is None:
            print(f"{name:<32}{'—':>20}{стало:>20,.0f}{'новый':>18}".replace(",", " "))
            сдвинулись.append(name)
            continue
        д = стало - было
        метка = "" if abs(д) <= max(1.0, abs(было) * 1e-9) else f"{д:>18,.0f}"
        if метка:
            сдвинулись.append(name)
        print(f"{name:<32}{было:>20,.0f}{стало:>20,.0f}{метка:>18}".replace(",", " "))

    if a.update:
        BASELINE.write_text(json.dumps(
            {"обновлён": date.today().isoformat(),
             "модель": snap["model"], "сценарий": snap.get("scenario_no"),
             "показатели": now}, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"\nэталон обновлён: {BASELINE.relative_to(ROOT)}")
        return 0

    if сдвинулись:
        print(f"\nсдвинулись: {', '.join(сдвинулись)}")
        print("если это ожидаемо — python3 tools/baseline.py --update "
              "и запись в docs/Журнал_правок.md")
        return 1
    print("\nвсе ключевые цифры совпали с эталоном")
    return 0


if __name__ == "__main__":
    sys.exit(main())
