# -*- coding: utf-8 -*-
"""БДР: сходится ли отчёт о прибылях сам с собой.

Каждый итог должен быть равен сумме своих слагаемых — помесячно, а не
только в целом за 60 месяцев. Итог, сходящийся «в сумме», но разъезжающийся
по месяцам, прячет ошибку в сезонности.
"""
from conftest import close, close_series


def test_выручка_равна_сумме_источников(base):
    parts = ["Паушальные взносы", "Роялти действующей сети", "Роялти новых школ",
             "Мерч, без НДС", "Лагеря, без НДС"]
    got = [sum(v) for v in zip(*(base.series(p) for p in parts))]
    close_series(base.series("Выручка — всего"), got,
                 "Выручка ≠ паушальные + роялти + мерч + лагеря")


def test_прямые_расходы_равны_сумме(base):
    parts = ["COGS мерча", "Логистика/возвраты/брак", "ФК — 35% прибыли мерча",
             "ФК — паушальные", "ФК — роялти", "Лагеря — расходы + партнёр"]
    got = [sum(v) for v in zip(*(base.series(p) for p in parts))]
    close_series(base.series("Прямые расходы — всего"), got,
                 "Прямые расходы ≠ сумме своих статей")


def test_валовая_прибыль(base):
    got = [a - b for a, b in zip(base.series("Выручка — всего"),
                                 base.series("Прямые расходы — всего"))]
    close_series(base.series("Валовая прибыль"), got,
                 "Валовая прибыль ≠ выручка − прямые расходы")


def test_opex_равен_сумме(base):
    parts = ["ФОТ и кадровые расходы", "ERP", "Офис / аренда / уборка / парковка",
             "Юристы / аудит / бухгалтерия", "Прочие IT", "SMM / подрядчики",
             "Маркетинг привлечения франшиз"]
    got = [sum(v) for v in zip(*(base.series(p) for p in parts))]
    close_series(base.series("OPEX — всего"), got, "OPEX ≠ сумме своих статей")


def test_ebitda(base):
    got = [a - b for a, b in zip(base.series("Валовая прибыль"),
                                 base.series("OPEX — всего"))]
    close_series(base.series("EBITDA"), got, "EBITDA ≠ валовая прибыль − OPEX")


def test_маржа_ebitda(base):
    rev = base.series("Выручка — всего")
    eb = base.series("EBITDA")
    got = base.series("EBITDA margin")
    for i, (r, e, m) in enumerate(zip(rev, eb, got), start=1):
        if abs(r) < 1:
            continue
        close(m, e / r, f"месяц {i}: маржа не равна EBITDA/выручка", abs_=1e-6)


def test_прибыль_после_усн(base):
    got = [a - b for a, b in zip(base.series("EBITDA"),
                                 base.series("УСН / минимальный налог"))]
    close_series(base.series("Прибыль после УСН"), got,
                 "Прибыль после УСН ≠ EBITDA − налог")


def test_налог_не_отрицательный(base):
    bad = [(i, v) for i, v in enumerate(base.series("УСН / минимальный налог"), 1) if v < -1]
    assert not bad, f"отрицательный налог (возврат из бюджета) в месяцах: {bad[:5]}"


def test_выручка_не_отрицательная(base):
    for label in ("Паушальные взносы", "Роялти действующей сети", "Роялти новых школ",
                  "Мерч, без НДС", "Выручка — всего"):
        bad = [(i, v) for i, v in enumerate(base.series(label), 1) if v < -1]
        assert not bad, f"«{label}» уходит в минус в месяцах {bad[:5]}"
