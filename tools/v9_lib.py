# -*- coding: utf-8 -*-
"""Строительные блоки книги V9.0: стили, раскладка, запись строк.

Раскладка каждого расчётного листа одна и та же:
    A  — номер блока / номер строки расчёта
    B  — название показателя
    C  — единица измерения
    D  — откуда берётся (короткое пояснение формулы словами)
    E:I — итоги по годам 1..5
    J   — итого за горизонт
    L:  — 60 месяцев

Годы слева, месяцы справа: руководитель видит пятилетку, не прокручивая,
а аналитик уходит вправо за помесячной раскладкой.
"""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

RED = "9E1B31"
GREY = "F2F2F2"
YELLOW = "FFF2CC"
BLUE = "DDEBF7"
GREEN = "E2F0D9"
WHITE = "FFFFFF"

КОЛ_ГОД1 = 5           # E
КОЛ_ИТОГО = 10         # J
КОЛ_МЕС1 = 12          # L
ЛЕТ = 5
МЕСЯЦЕВ = 60

ДЕНЬГИ = '# ##0;[Red]-# ##0;-'
ДЕНЬГИ_МЛН = '# ##0,0;[Red]-# ##0,0;-'
ШТУКИ = '# ##0,0;[Red]-# ##0,0;-'
ЦЕЛЫЕ = '# ##0;[Red]-# ##0;-'
ПРОЦЕНТ = '0,0%;[Red]-0,0%;-'
ДАТА = 'ммм гггг'


def мес(i):
    """Буква столбца месяца i (1..60)."""
    return L(КОЛ_МЕС1 + i - 1)


def год(y):
    """Буква столбца года y (1..5)."""
    return L(КОЛ_ГОД1 + y - 1)


class Лист:
    """Обёртка над листом: пишет строки расчёта в единой раскладке."""

    def __init__(self, ws, заголовок, подзаголовок=""):
        self.ws = ws
        self.строка = 1
        ws.sheet_view.showGridLines = False
        ws["B1"] = заголовок
        ws["B1"].font = Font(bold=True, size=14, color=RED)
        if подзаголовок:
            ws["B2"] = подзаголовок
            ws["B2"].font = Font(italic=True, size=10, color="595959")
        self.строка = 4
        self._ширины()

    def _ширины(self):
        ws = self.ws
        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 44
        ws.column_dimensions["C"].width = 11
        ws.column_dimensions["D"].width = 46
        for y in range(ЛЕТ):
            ws.column_dimensions[L(КОЛ_ГОД1 + y)].width = 15
        ws.column_dimensions[L(КОЛ_ИТОГО)].width = 16
        ws.column_dimensions[L(КОЛ_ИТОГО + 1)].width = 3
        for i in range(1, МЕСЯЦЕВ + 1):
            ws.column_dimensions[мес(i)].width = 13

    def шапка(self, дата_строка=None):
        """Строка заголовков: годы и месяцы."""
        ws, r = self.ws, self.строка
        ws[f"B{r}"] = "Показатель"
        ws[f"C{r}"] = "Ед."
        ws[f"D{r}"] = "Как считается"
        for y in range(1, ЛЕТ + 1):
            ws[f"{год(y)}{r}"] = f"Год {y}"
        ws[f"{L(КОЛ_ИТОГО)}{r}"] = "Итого 5 лет"
        for i in range(1, МЕСЯЦЕВ + 1):
            c = мес(i)
            if дата_строка:
                ws[f"{c}{r}"] = f"={дата_строка}!{c}$4"
                ws[f"{c}{r}"].number_format = ДАТА
            else:
                ws[f"{c}{r}"] = f"М{i}"
        for col in range(2, КОЛ_МЕС1 + МЕСЯЦЕВ):
            c = ws.cell(row=r, column=col)
            c.font = Font(bold=True, color=WHITE, size=10)
            c.fill = PatternFill("solid", fgColor=RED)
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
        ws.freeze_panes = f"{L(КОЛ_МЕС1)}{r + 1}"
        ws.row_dimensions[r].height = 30
        self.строка += 1
        return r

    def блок(self, название):
        ws, r = self.ws, self.строка
        ws[f"B{r}"] = название
        ws[f"B{r}"].font = Font(bold=True, size=11, color=RED)
        ws[f"B{r}"].fill = PatternFill("solid", fgColor=YELLOW)
        for col in range(2, КОЛ_МЕС1 + МЕСЯЦЕВ):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=YELLOW)
        self.строка += 1
        return r

    def ряд(self, название, формула, ед="₽", как="", формат=ДЕНЬГИ,
            жирный=False, первый=None, итог="сумма"):
        """Строка расчёта: месяцы считаются формулой, годы — сводкой.

        `формула` — шаблон с {м} вместо буквы столбца месяца и {пред} вместо
        предыдущего. `первый` — отдельная формула для первого месяца."""
        ws, r = self.ws, self.строка
        ws[f"B{r}"] = название
        ws[f"C{r}"] = ед
        ws[f"D{r}"] = как
        ws[f"D{r}"].font = Font(italic=True, size=9, color="595959")
        ws[f"D{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        for i in range(1, МЕСЯЦЕВ + 1):
            c = мес(i)
            ш = первый if (i == 1 and первый is not None) else формула
            if ш is None:
                continue
            знач = ш.format(м=c, пред=мес(i - 1) if i > 1 else None, i=i, r=r)
            ws[f"{c}{r}"] = знач if str(знач).startswith("=") else знач
            ws[f"{c}{r}"].number_format = формат
        # годы и итог
        for y in range(1, ЛЕТ + 1):
            a, b = мес((y - 1) * 12 + 1), мес(y * 12)
            ф = (f"=SUM({a}{r}:{b}{r})" if итог == "сумма"
                 else f"={b}{r}" if итог == "конец"
                 else f"=MIN({a}{r}:{b}{r})" if итог == "минимум"
                 else f"=AVERAGE({a}{r}:{b}{r})")
            ws[f"{год(y)}{r}"] = ф
            ws[f"{год(y)}{r}"].number_format = формат
        итого = (f"=SUM({мес(1)}{r}:{мес(МЕСЯЦЕВ)}{r})" if итог == "сумма"
                 else f"={мес(МЕСЯЦЕВ)}{r}" if итог == "конец"
                 else f"=MIN({мес(1)}{r}:{мес(МЕСЯЦЕВ)}{r})" if итог == "минимум"
                 else f"=AVERAGE({мес(1)}{r}:{мес(МЕСЯЦЕВ)}{r})")
        ws[f"{L(КОЛ_ИТОГО)}{r}"] = итого
        ws[f"{L(КОЛ_ИТОГО)}{r}"].number_format = формат
        if жирный:
            for col in list(range(2, КОЛ_ИТОГО + 1)) + \
                       list(range(КОЛ_МЕС1, КОЛ_МЕС1 + МЕСЯЦЕВ)):
                ws.cell(row=r, column=col).font = Font(bold=True)
        for y in range(ЛЕТ):
            ws.cell(row=r, column=КОЛ_ГОД1 + y).fill = PatternFill("solid", fgColor=GREY)
        ws.cell(row=r, column=КОЛ_ИТОГО).fill = PatternFill("solid", fgColor=BLUE)
        self.строка += 1
        return r

    def пусто(self, n=1):
        self.строка += n
