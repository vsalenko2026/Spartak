# -*- coding: utf-8 -*-
"""Общая обвязка тестов: срез прогона и доступ к строкам модели.

Тесты работают не с файлом Excel, а со срезом прогона
(build/runs/<тег>.json), который делает tools/run_model.py. Если среза нет,
он считается на месте — тесты самодостаточны.
"""
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
RUNS = ROOT / "build" / "runs"
RUB = 1.0          # допуск на копейки при сверке сумм
REL = 1e-6         # относительный допуск


def ensure_run(tag: str, args=()):
    """Срез прогона; считается на месте, если его ещё нет.

    SPARTAK_MODEL подменяет саму книгу — этим пользуется tools/selfcheck.py,
    чтобы прогнать против намеренно испорченной модели не только базовый
    срез, но и производные от него (например, прогон с другой датой старта)."""
    path = RUNS / f"{tag}.json"
    if not path.exists():
        модель = os.environ.get("SPARTAK_MODEL")
        cmd = [sys.executable, str(ROOT / "tools" / "run_model.py"), "--tag", tag,
               *(["--model", модель] if модель else []), *args]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=2400)
        if not path.exists():
            pytest.fail(f"не удалось прогнать модель «{tag}»:\n"
                        f"{res.stdout[-2000:]}\n{res.stderr[-2000:]}")
    return json.loads(path.read_text(encoding="utf-8"))


class Run:
    """Срез одного прогона с удобным доступом к строкам."""

    def __init__(self, data):
        self.data = data

    @property
    def tag(self):
        return self.data["tag"]

    def rows(self, block="модель"):
        return self.data["blocks"][block]["rows"]

    def find(self, label, block="модель", exact=False):
        """Строка по подписи. Совпадение по началу подписи, чтобы
        «EBITDA» не цеплял «EBITDA margin»."""
        want = label.lower().strip()
        best = None
        for r, rec in sorted(self.rows(block).items(), key=lambda kv: int(kv[0])):
            got = (rec["label"] or "").lower().strip()
            if got == want:
                return rec
            if not exact and best is None and got.startswith(want):
                best = rec
        if best is None:
            raise AssertionError(
                f"нет строки «{label}» в блоке «{block}» прогона «{self.tag}»")
        return best

    def series(self, label, block="модель"):
        """Помесячный ряд: числа, None -> 0."""
        rec = self.find(label, block)
        return [v if isinstance(v, (int, float)) else 0.0 for v in rec["v"]]

    def total(self, label, block="модель"):
        return sum(self.series(label, block))


@pytest.fixture(scope="session")
def base():
    """Базовый прогон. SPARTAK_BASE_TAG подменяет его другим срезом —
    этим пользуется tools/selfcheck.py, чтобы прогнать тесты против
    намеренно испорченной модели."""
    tag = os.environ.get("SPARTAK_BASE_TAG", "база")
    return Run(ensure_run(tag, ["--keep"]))


@pytest.fixture(scope="session")
def negative():
    return Run(ensure_run("негатив", ["--scenario", "негативный"]))


@pytest.fixture(scope="session")
def positive():
    return Run(ensure_run("позитив", ["--scenario", "позитивный"]))


def close(a, b, note="", rel=REL, abs_=RUB):
    """Сверка с допуском, с внятным сообщением при расхождении."""
    diff = abs(a - b)
    if diff <= max(abs_, abs(a) * rel, abs(b) * rel):
        return
    raise AssertionError(
        f"{note}: {a:,.2f} против {b:,.2f}, расхождение {diff:,.2f}".replace(",", " "))


def close_series(xs, ys, note="", rel=REL, abs_=RUB):
    """Помесячная сверка: сообщает первый разошедшийся месяц, а не только факт."""
    assert len(xs) == len(ys), f"{note}: разная длина рядов {len(xs)} и {len(ys)}"
    bad = []
    for i, (a, b) in enumerate(zip(xs, ys), start=1):
        if abs(a - b) > max(abs_, abs(a) * rel, abs(b) * rel):
            bad.append(f"месяц {i}: {a:,.2f} против {b:,.2f}".replace(",", " "))
    if bad:
        raise AssertionError(f"{note}: расходится в {len(bad)} мес. из {len(xs)}\n  "
                             + "\n  ".join(bad[:6]))
