# -*- coding: utf-8 -*-
"""Самопроверка контура: ловят ли тесты и аудит те ошибки, ради которых
написаны.

Зелёные тесты сами по себе ничего не доказывают — они могли бы молчать и
на сломанной модели. Поэтому мы ломаем модель нарочно, по одному дефекту
за раз, и требуем, чтобы контур сработал.

    python3 tools/selfcheck.py            все дефекты
    python3 tools/selfcheck.py --only бддс

Каждый дефект описан тем, что он изображает: не «поменять формулу в AA44»,
а «кто-то вбил в EBITDA ноль за один месяц».
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
MODEL = ROOT / "model" / "Spartak_V8_9_понятная.xlsx"
MUTANTS = ROOT / "build" / "mutants"

# дефект: (описание, правки {лист!ячейка: формула/значение}, что обязано упасть)
DEFECTS = {
    "ебитда": (
        "в EBITDA за один месяц вбит ноль вместо формулы",
        {"02_Модель!AA44": 0},
        ["test_ebitda"],
    ),
    "бддс": (
        "денежный поток подменён EBITDA — БДР и БДДС смешаны",
        {f"02_Модель!{c}71": f"={c}44" for c in
         ("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M")},
        ["test_поток_после_погашения_займа"],
    ),
    "выручка": (
        "из суммы выручки выпало слагаемое «мерч»",
        {"02_Модель!AA23": "=AA18+AA19+AA20+AA22"},
        ["test_выручка_равна_сумме_источников"],
    ),
    "набор": (
        "выключен постепенный набор — школа сразу на полной загрузке",
        {"01_Вводные!B295": "Нет"},
        ["test_новые_школы_набирают_детей_постепенно"],
    ),
    "закрытие": (
        "коэффициент сохранения оторван от ставки закрытия",
        {"01_Вводные!B14": 1.0},
        ["test_закрытие_школ_применяется_как_задано"],
    ),
    "партнёры": (
        "партнёров приравняли к школам — 1 партнёр = 1 школа",
        {"01_Вводные!B9": 51},
        ["test_на_партнёра_приходится_больше_одной_школы"],
    ),
    "остатки": (
        "остаток на начало месяца оторван от конца предыдущего",
        {"02_Модель!AB51": "=AA51"},
        ["test_остатки_состыкованы_по_месяцам"],
    ),
}


def build_mutant(name: str, edits: dict) -> Path:
    MUTANTS.mkdir(parents=True, exist_ok=True)
    dst = MUTANTS / f"{name}.xlsx"
    wb = openpyxl.load_workbook(MODEL, data_only=False)
    for target, value in edits.items():
        sheet, addr = target.rsplit("!", 1)
        wb[sheet][addr] = value
    wb.save(dst)
    return dst


def run_model(model: Path, tag: str):
    cmd = [sys.executable, str(ROOT / "tools" / "run_model.py"),
           "--model", str(model), "--tag", tag]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=2400)
    if not (ROOT / "build" / "runs" / f"{tag}.json").exists():
        raise RuntimeError(f"прогон «{tag}» не состоялся:\n{res.stdout[-1500:]}"
                           f"\n{res.stderr[-1500:]}")


def run_tests(tag: str):
    """Тесты против подменённого среза. Возвращает множество упавших тестов."""
    env = dict(os.environ, SPARTAK_BASE_TAG=tag)
    res = subprocess.run(
        [sys.executable, "-m", "pytest", str(ROOT / "tests"), "-q", "--no-header",
         "-p", "no:cacheprovider", "--tb=no"],
        capture_output=True, text=True, env=env, cwd=str(ROOT), timeout=2400)
    failed = set()
    for line in res.stdout.splitlines():
        if line.startswith("FAILED"):
            # FAILED tests/test_бдр.py::test_ebitda - AssertionError: ...
            node = line.split()[1] if len(line.split()) > 1 else ""
            name = node.split("::")[-1].split("[")[0]
            if name:
                failed.add(name)
    return failed, res.stdout


def main():
    ap = argparse.ArgumentParser(description="Самопроверка контура тестов")
    ap.add_argument("--only", action="append", help="проверить только этот дефект")
    a = ap.parse_args()

    names = a.only or list(DEFECTS)
    ok = True
    print(f"{'дефект':<12}{'что изображает':<52}{'итог'}")
    print("-" * 76)
    for name in names:
        if name not in DEFECTS:
            sys.exit(f"нет дефекта {name!r}. Есть: {', '.join(DEFECTS)}")
        описание, edits, expect = DEFECTS[name]
        tag = f"дефект_{name}"
        model = build_mutant(name, edits)
        run_model(model, tag)
        failed, out = run_tests(tag)
        missed = [t for t in expect if t not in failed]
        if missed:
            ok = False
            print(f"{name:<12}{описание:<52}НЕ ПОЙМАН")
            print(f"{'':12}ожидали падения: {', '.join(missed)}")
            print(f"{'':12}упали фактически: {', '.join(sorted(failed)) or '—'}")
        else:
            print(f"{name:<12}{описание:<52}пойман ({len(failed)} тестов)")

    print("-" * 76)
    if ok:
        print("контур рабочий: каждый подсаженный дефект поймали тесты")
    else:
        print("ЕСТЬ ДЫРЫ: часть дефектов проходит мимо тестов")
    # чистим срезы дефектов, чтобы не путались с настоящими прогонами
    for name in names:
        (ROOT / "build" / "runs" / f"дефект_{name}.json").unlink(missing_ok=True)
    shutil.rmtree(MUTANTS, ignore_errors=True)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
