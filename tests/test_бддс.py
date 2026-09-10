# -*- coding: utf-8 -*-
"""БДДС: сходится ли движение денег, и не подменён ли он БДР.

Главное тождество кассы: остаток на конец = остаток на начало + поток.
Если оно нарушено хоть в одном месяце, деньги в модели берутся из воздуха.
"""
from conftest import close, close_series


def test_остаток_на_конец(base):
    got = [a + b for a, b in zip(base.series("Денежный остаток на начало"),
                                 base.series("Чистый денежный поток"))]
    close_series(base.series("Денежный остаток на конец"), got,
                 "Остаток на конец ≠ остаток на начало + чистый поток")


def test_остатки_состыкованы_по_месяцам(base):
    end = base.series("Денежный остаток на конец")
    start = base.series("Денежный остаток на начало")
    close_series(start[1:], end[:-1],
                 "Остаток на начало месяца ≠ остатку на конец предыдущего")


def test_поступления_равны_сумме(base):
    parts = ["Паушальные поступления", "Роялти поступления", "Мерч поступления",
             "Лагеря поступления"]
    got = [sum(v) for v in zip(*(base.series(p) for p in parts))]
    close_series(base.series("Денежные поступления — всего"), got,
                 "Поступления ≠ сумме своих источников")


def test_поток_до_займа(base):
    got = [a - b for a, b in zip(base.series("Денежные поступления — всего"),
                                 base.series("Денежные расходы до займа"))]
    close_series(base.series("Чистый поток до займа"), got,
                 "Поток до займа ≠ поступления − расходы")


def test_поток_после_погашения_займа(base):
    got = [a - b for a, b in zip(base.series("Чистый поток до займа"),
                                 base.series("Погашение займа"))]
    close_series(base.series("Чистый денежный поток"), got,
                 "Чистый поток ≠ поток до займа − погашение займа")


def test_бдр_и_бддс_не_совпадают(base):
    """Начисление и деньги обязаны расходиться: есть склад, займ и авансы.
    Полное совпадение означает, что один отчёт подменили другим."""
    eb = base.total("EBITDA")
    cf = base.total("Чистый денежный поток")
    assert abs(eb - cf) > max(1.0, abs(eb) * 1e-6), (
        f"EBITDA и денежный поток совпали до копейки ({eb:,.0f}) — "
        f"признак того, что БДДС повторяет БДР".replace(",", " "))


def test_погашение_займа_не_превышает_займ(base):
    repaid = base.total("Погашение займа")
    loan_end = base.series("Займ на конец")
    start_loan = loan_end[0] + base.series("Погашение займа")[0]
    assert repaid <= start_loan + 1, (
        f"погашено {repaid:,.0f} при займе {start_loan:,.0f}".replace(",", " "))


def test_займ_не_отрицательный(base):
    bad = [(i, v) for i, v in enumerate(base.series("Займ на конец"), 1) if v < -1]
    assert not bad, f"займ ушёл в минус (переплата) в месяцах {bad[:5]}"


def test_склад_не_отрицательный(base):
    bad = [(i, v) for i, v in enumerate(base.series("Склад на конец"), 1) if v < -1]
    assert not bad, f"склад ушёл в минус в месяцах {bad[:5]}"


def test_кассовый_разрыв_объявлен(base):
    """Модель допускает уход в минус — но тогда это должно быть видно на
    дашборде как пиковый дефицит, а не всплывать при чтении БДДС."""
    end = base.series("Денежный остаток на конец")
    worst = min(end)
    if worst >= 0:
        return
    declared = base.series("Пиковый кассовый дефицит", "дашборд")
    nums = [v for v in declared if isinstance(v, (int, float)) and v != 0]
    assert nums, (f"минимальный остаток {worst:,.0f}, "
                  f"а на дашборде дефицит не показан".replace(",", " "))
    # дашборд разбит по сезонам, последнее число — итог за все 60 мес.
    close(max(abs(v) for v in nums), abs(worst),
          "пиковый дефицит на дашборде не совпадает с минимумом остатка",
          rel=1e-3)
