# -*- coding: utf-8 -*-
"""Сценарии: негативный обязан быть хуже базового, позитивный — лучше.

Проверка дешёвая, а ловит целый класс ошибок: перепутанные местами
коэффициенты, сценарий, не доехавший до расчёта, и переключатель,
который молча возвращается к базовому.
"""
import pytest


ORDERED = ["Выручка — всего", "EBITDA", "Всего активных детей"]


@pytest.mark.parametrize("label", ORDERED)
def test_негативный_хуже_базового(negative, base, label):
    n, b = negative.total(label), base.total(label)
    assert n < b, (f"негативный сценарий даёт «{label}» {n:,.0f}, "
                   f"а базовый {b:,.0f} — сценарии перепутаны".replace(",", " "))


@pytest.mark.parametrize("label", ORDERED)
def test_позитивный_лучше_базового(positive, base, label):
    p, b = positive.total(label), base.total(label)
    assert p > b, (f"позитивный сценарий даёт «{label}» {p:,.0f}, "
                   f"а базовый {b:,.0f} — сценарии перепутаны".replace(",", " "))


def test_сценарий_действительно_переключился(negative, base, positive):
    assert (negative.data["scenario_no"], base.data["scenario_no"],
            positive.data["scenario_no"]) == (1, 2, 3), (
        "модель посчитала не те сценарии: "
        f"{negative.data['scenario_no']}, {base.data['scenario_no']}, "
        f"{positive.data['scenario_no']}")


def test_сценарии_различаются_помесячно(negative, base):
    """Разница должна быть видна по месяцам, а не только в итоге: иначе
    сценарий приложен к результату задним числом одним множителем."""
    n = negative.series("EBITDA")
    b = base.series("EBITDA")
    same = sum(1 for a, c in zip(n, b) if abs(a - c) < 1)
    assert same < len(b) * 0.5, (
        f"{same} месяцев из {len(b)} совпали до рубля — "
        f"сценарий, похоже, не влияет на помесячный расчёт")
