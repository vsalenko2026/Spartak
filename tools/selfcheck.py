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
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
MODEL = ROOT / "model" / "Spartak_V8_9_понятная.xlsx"
MUTANTS = ROOT / "build" / "mutants"

# дефект: (описание, правки {лист!ячейка: формула/значение}, что обязано упасть)
СТОЛБЦЫ_МЕСЯЦЕВ = [get_column_letter(c) for c in range(2, 62)]

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
    "котёл": (
        "роялти действующей сети вернулось к котлу «школы x 15 000»",
        {f"06_Сеть_и_дети!{c}28":
         (f"=MAX({c}20*'18_Календарь_тарифы'!{c}8*"
          f"IF({c}$3>='01_Вводные'!$B$245,'01_Вводные'!$B$57,"
          f"'01_Вводные'!$B$55),{c}11*15000)")
         for c in СТОЛБЦЫ_МЕСЯЦЕВ},
        ["test_фикс_берётся_с_точки_начисления_а_не_со_школы"],
    ),
    "календарь": (
        "веса плана продаж снова берутся по номеру месяца в сезоне",
        {f"04_Продажи!{get_column_letter(c)}15":
         (f"='01_Вводные'!$B${35 + min(5, (c - 2) // 12 + 1)}"
          f"*'01_Вводные'!${get_column_letter(3 + (c - 2) % 12)}"
          f"${35 + min(5, (c - 2) // 12 + 1)}"
          f"/SUM('01_Вводные'!$C${35 + min(5, (c - 2) // 12 + 1)}"
          f":$N${35 + min(5, (c - 2) // 12 + 1)})")
         for c in range(2, 62)},
        ["test_пик_продаж_стоит_на_сентябре_а_не_на_первом_месяце"],
    ),
    "пустой счёт": (
        "порог остатка снят — клубу снова платят до нуля на счёте",
        {"01_Вводные!B237": 0},
        ["test_порог_остатка_при_возврате_займа_не_нулевой",
         "test_касса_не_проваливается_глубже_месячных_расходов",
         "test_первый_месяц_не_уходит_в_ноль"],
    ),
    "разом": (
        "формула открытий вернулась к «весь коэффициент в один месяц»",
        {f"06_Сеть_и_дети!{get_column_letter(c)}8":
         (f"=IF(OR({c - 1}<='01_Вводные'!$B$15,"
          f"{c - 1}-'01_Вводные'!$B$15>60),0,"
          f"INDEX('04_Продажи'!$B$15:$BI$15,1,{c - 1}-'01_Вводные'!$B$15))"
          f"*'01_Вводные'!$B$298*'01_Вводные'!$B$22")
         for c in range(2, 62)},
        ["test_первая_школа_партнёра_одна"],
    ),
    "партнёры": (
        "партнёров приравняли к школам — 1 партнёр = 1 школа",
        {"01_Вводные!B9": 51},
        ["test_на_партнёра_приходится_больше_одной_школы",
         "test_партнёров_не_больше_чем_школ"],
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


def run_tests(tag: str, model: Path):
    """Тесты против подменённого среза. Возвращает множество упавших тестов.

    SPARTAK_MODEL нужен тестам, которые считают собственный прогон
    (например, с другой датой старта): без него они посчитались бы с
    целой модели и дефект бы не увидели."""
    env = dict(os.environ, SPARTAK_BASE_TAG=tag, SPARTAK_MODEL=str(model))
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
        failed, out = run_tests(tag, model)
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
        for срез in (ROOT / "build" / "runs").glob(f"дефект_{name}*.json"):
            срез.unlink(missing_ok=True)
    shutil.rmtree(MUTANTS, ignore_errors=True)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
