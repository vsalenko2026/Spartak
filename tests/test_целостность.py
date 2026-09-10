# -*- coding: utf-8 -*-
"""Целостность книги: ошибки в ячейках и собственные проверки модели.

Лист 13_Проверки — контроли, которые автор заложил сам. Тесты не заменяют
их, а следят, что они не превратились в декорацию.
"""
import json
import subprocess
import sys
from pathlib import Path

import pytest
from conftest import ROOT

ИЗВЕСТНЫЕ_РИСКИ = ("риск плана",)      # не ошибки формул, а плановые допущения


def test_нет_ошибок_в_ячейках(base):
    errors = base.data.get("cell_errors", [])
    assert not errors, ("в пересчитанной книге ошибки ячеек:\n  "
                        + "\n  ".join(errors[:15]))


def test_собственные_проверки_модели(base):
    """Лист 13_Проверки не должен содержать провалов, кроме заранее
    признанных плановых рисков."""
    rows = base.rows("проверки")
    провалы = []
    for r, rec in sorted(rows.items(), key=lambda kv: int(kv[0])):
        for v in rec["v"]:
            if not isinstance(v, str):
                continue
            низ = v.lower()
            if "провал" in низ or "проверить" in низ:
                if any(з in низ for з in ИЗВЕСТНЫЕ_РИСКИ):
                    continue
                провалы.append(f"строка {r}: {rec['label'][:60]} → {v}")
                break
    assert not провалы, ("собственные проверки модели не пройдены:\n  "
                         + "\n  ".join(провалы[:10]))


def test_статический_аудит_без_важных_замечаний():
    """Аудит формул не должен находить ничего с приоритетом «важно»:
    ошибки ячеек, разрывы протяжки в середине ряда, съехавшие ссылки."""
    карта = ROOT / "build" / "model_map.json"
    if not карта.exists():
        pytest.skip("нет build/model_map.json — сначала tools/extract_model.py")
    res = subprocess.run([sys.executable, str(ROOT / "tools" / "audit_model.py"),
                          "--show", "40"],
                         capture_output=True, text=True, cwd=str(ROOT), timeout=1200)
    важных = None
    for line in res.stdout.splitlines():
        if line.startswith("ИТОГО"):
            важных = int(line.split()[-1])
    assert важных is not None, f"аудит не отработал:\n{res.stdout[-1500:]}\n{res.stderr[-800:]}"
    assert важных == 0, (f"аудит нашёл {важных} важных замечаний:\n"
                         + "\n".join(res.stdout.splitlines()[-25:]))
