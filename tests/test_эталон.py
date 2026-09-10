# -*- coding: utf-8 -*-
"""Регресс: ключевые цифры не двигаются молча.

Тест не утверждает, что цифры правильные. Он утверждает, что они те же,
что были в прошлый раз. Любое расхождение обязано быть объяснено в
docs/Журнал_правок.md и принято через tools/baseline.py --update.
"""
import json

import pytest
from conftest import ROOT

BASELINE = ROOT / "tests" / "baseline.json"


def test_ключевые_цифры_совпадают_с_эталоном(base):
    if not BASELINE.exists():
        pytest.skip("эталона ещё нет: python3 tools/baseline.py --update")
    эталон = json.loads(BASELINE.read_text(encoding="utf-8"))["показатели"]

    import sys
    sys.path.insert(0, str(ROOT / "tools"))
    from baseline import KEYS, series

    сдвиги = []
    for name, block, label, how in KEYS:
        if name not in эталон:
            continue
        vals = series(base.data, block, label)
        стало = {"сумма": sum(vals), "последнее": vals[-1], "минимум": min(vals)}[how]
        было = эталон[name]
        if abs(стало - было) > max(1.0, abs(было) * 1e-9):
            сдвиги.append(f"{name}: было {было:,.0f}, стало {стало:,.0f}, "
                          f"дельта {стало - было:+,.0f}".replace(",", " "))
    assert not сдвиги, (
        "ключевые цифры разошлись с эталоном:\n  " + "\n  ".join(сдвиги)
        + "\n\nЕсли это ожидаемое следствие правки — опишите её в "
          "docs/Журнал_правок.md и примите: python3 tools/baseline.py --update")
