#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Одностраничная шпаргалка лектора к Лекции 4-доп (А4, две колонки).

Только смысловые опоры по слайдам 3-6: аббревиатуры, формулы и их расшифровки.
Без таймингов, схем, таблиц и пересказа того, что и так видно на слайдах
(титул, план, итоговые тезисы, заключение).

Запуск:  python build_cheatsheet.py
"""

import os

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT
from reportlab.lib import colors as rl_colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
)

from build_lecture_4dop import (
    BASE, PDF_FONT, PDF_FONT_BOLD, pdf_markup, SCRIPT,
)

OUT_PATH = os.path.join(BASE, "Лекция_4доп_Шпаргалка_лектора_А4.pdf")

ACCENT = rl_colors.HexColor("#00695C")
SLIDE_COLOR = rl_colors.HexColor("#BF360C")
FORMULA_COLOR = rl_colors.HexColor("#0D47A1")

# Кегль и межстрочный интервал — крутим, чтобы влезло на один лист
BODY_SIZE = 9.8
BODY_LEAD = 11.7

styles = getSampleStyleSheet()
S_TITLE = ParagraphStyle(
    "CtTitle", parent=styles["Normal"], fontName=PDF_FONT_BOLD,
    fontSize=11, leading=13, textColor=ACCENT, spaceAfter=2,
)
S_ANCHORS = ParagraphStyle(
    "CtAnchors", parent=styles["Normal"], fontName=PDF_FONT,
    fontSize=BODY_SIZE - 0.6, leading=BODY_LEAD - 1.0,
    textColor=rl_colors.HexColor("#37474F"),
)
S_SLIDE = ParagraphStyle(
    "CtSlide", parent=styles["Normal"], fontName=PDF_FONT_BOLD,
    fontSize=BODY_SIZE + 1.2, leading=BODY_LEAD + 1.2, textColor=SLIDE_COLOR,
    spaceBefore=6, spaceAfter=2,
)
S_BODY = ParagraphStyle(
    "CtBody", parent=styles["Normal"], fontName=PDF_FONT,
    fontSize=BODY_SIZE, leading=BODY_LEAD, alignment=TA_LEFT, spaceAfter=1.6,
)
S_FORMULA = ParagraphStyle(
    "CtFormula", parent=S_BODY, fontName=PDF_FONT_BOLD,
    leftIndent=6, textColor=FORMULA_COLOR, spaceBefore=1.4,
)


# ── Содержимое: опоры по слайдам 3-6 ──────────────────────────────────────
# Ключ — номер слайда (сверяется с раскадровкой). "f" — формула, "b" — обычный пункт.
SLIDES = [
    (3, "Медь — проблема для ИК-лазеров", [
        ("b", "A(λ) — спектральная поглощательная способность (absorbance): доля энергии луча, поглощённая материалом, а не отражённая."),
        ("b", "A(1070 нм) ≈ 5% → 95% отражается; 5% греют лишь тонкий поверхностный слой → плавление нестабильно → поры, брызги, несплавления."),
        ("b", "A(535 нм) = 40–60% → в 8–12 раз больше энергии в зоне при той же мощности → стабильный шов."),
        ("b", "Где нужен: сварка медных шин в электромобилях, 3D-печать медных теплообменников."),
        ("b", "Вывод: физика идеальна — проблема в экономике."),
    ]),
    (4, "Прямая генерация: почему не работает", [
        ("b", "InGaN — нитрид индия-галлия. Диодные лазеры на нём есть, но упираются в проблему зелёного диапазона (green gap)."),
        ("f", "E = hc/λ — энергия фотона: h — постоянная Планка (Дж·с), c = 3·10^{8} м/с, λ — длина волны (м)."),
        ("b", "535 нм → E ≈ 3,7·10^{−19} Дж ≈ 2,3 эВ → нужна запрещённая зона 2,3 эВ → >30% индия в решётке."),
        ("b", "Атом индия крупнее галлия → при >30% решётка нестабильна → V-дефекты: ловушки, энергия → тепло вместо излучения."),
        ("b", "QCSE — квантово-размерный эффект Штарка: пьезоэлектрическое поле в квантовой яме разводит электрон и дырку → излучательная рекомбинация падает."),
        ("b", "Итог: ≤ 0,18 Вт (180 мВт) в лаборатории против нужных киловатт — разрыв 10 000 раз."),
    ]),
    (5, "Нелинейная оптика: вторая гармоника", [
        ("b", "SHG (Second Harmonic Generation) — генерация второй гармоники: 1070 нм (ω) → 535 нм (2ω)."),
        ("f", "E(t) = E_{0}·cos(ω·t) — волна, падающая на кристалл: E_{0} — амплитуда поля, ω = 2πc/λ — циклическая частота (рад/с), t — время (с)."),
        ("f", "P = ε_{0}·(χ^{(1)}E + χ^{(2)}E^{2} + χ^{(3)}E^{3} + …)"),
        ("b", "P — поляризация среды (Кл/м^{2}), дипольный момент единицы объёма; ε_{0} = 8,85·10^{−12} Ф/м, Ф (фарад) = Кл/В; E — напряжённость поля (В/м); χ^{(1)} — линейная, безразмерная, только преломление; χ^{(2)} — квадратичная, м/В, удвоение частоты — оставляем её; χ^{(3)} — кубическая, м^{2}/В^{2}, на порядки меньше — пренебрегаем. Индекс в скобках — порядок нелинейности, а не степень."),
        ("b", "Откуда нелинейность: шарик в чаше — симметричная яма (стекло, вода) → колебания на частоте поля, новых частот нет; асимметричная яма (ниобат лития: один край крутой, другой пологий) → колебания несимметричны → кратные частоты."),
        ("f", "P^{(2)} ∝ χ^{(2)}·E_{0}^{2}·cos^{2}(ω·t) — подставили E(t) во второй член; «∝» — пропорционально, постоянные множители опущены."),
        ("f", "cos^{2}(ω·t) = (1 + cos(2ω·t))/2 — тригонометрическое тождество."),
        ("b", "1/2 — оптическое выпрямление: постоянное поле, не излучает → бесполезно. cos(2ωt) — среда колеблется на 2ω → длина волны вдвое меньше."),
        ("f", "λ_{2ω} = 1070/2 = 535 нм"),
        ("b", "Проблема: из-за дисперсии ω и 2ω идут с разными скоростями → зелёные волны из разных точек кристалла гасят друг друга."),
        ("b", "PPLN (Periodically Poled Lithium Niobate) — периодически поляризованный ниобат лития: каждые Λ кристаллическая ось перевёрнута на 180° → знак χ^{(2)} меняется → набег фазы компенсируется. Λ ≈ 6,96 мкм."),
        ("f", "Δk = k_{2ω} − 2k_{ω} = 2π/Λ — условие синхронизма: k — волновые числа, Λ — период решётки переполяризации, Δk — рассогласование фаз."),
    ]),
    (6, "Моё предложение: двухпроходная схема", [
        ("b", "Однопроход: КПД 10–15%. Предложение: двухпроходная схема с активной клиновидной термокомпенсацией (подвижный клин из плавленого кварца, fused silica) для волоконных лазеров ≥1 кВт под медь. Прецедент: Imeshev, Fejer (1998)."),
        ("b", "Печь ±0,1 °C нужна всегда; клин снимает фазовый сдвиг между прямым и обратным проходом (отражение от зеркала)."),
        ("f", "η = P_{2ω}/P_{ω} · (2ω^{2}·d_{eff}^{2}·L^{2}) / (ε_{0}·c^{3}·n_{ω}^{2}·n_{2ω}) · (P_{ω}/A) · sinc^{2}(Δk·L/2)"),
        ("b", "d_{eff} ≈ 25 пм/В (MgO:PPLN, d_{33}); L = 50 мм; n_{ω} ≈ 2,13, n_{2ω} ≈ 2,23; A — площадь пятна; sinc(x) = sin(x)/x. Формула — из уравнений связанных волн, приближение неистощённой накачки."),
        ("b", "Преимущества перед Имешевым (клин = сам PPLN): 1) тепловой дефазинг — печь и внешний клин решают задачи независимо; 2) кварц дёшев в обработке; 3) пластину меняют, не трогая PPLN; 4) подвижный клин правит фазу в реальном времени."),
    ]),
]

ANCHORS = (
    "5% при 1070 нм · 40–60% при 535 нм (в 8–12 раз) · 2,3 эВ, >30% In · ≤ 0,18 Вт против кВт "
    "(разрыв 10 000 раз) · ω = 2πc/λ · Λ ≈ 6,96 мкм · КПД 10–15% один проход / 15–20% два · "
    "печь ±0,1 °C · dn/dT = +9·10^{−6} К^{−1} · d_{33} ≈ 25 пм/В · L = 50 мм · Imeshev, Fejer (1998)"
)


def _check_slides():
    """Шпаргалка не должна разъезжаться с раскадровкой: нужные слайды должны быть на месте."""
    known = {title.split(".")[0] for title, _timing, _paras in SCRIPT}
    for num, name, _lines in SLIDES:
        if str(num) not in known:
            raise RuntimeError(
                f"В раскадровке нет слайда {num} ({name}) — обновите SLIDES в build_cheatsheet.py"
            )


def build_flowables():
    _check_slides()
    flow = [
        Paragraph(pdf_markup("Шпаргалка лектора · Лекция 4-доп · Генерация зелёного излучения (535 нм)"), S_TITLE),
        Spacer(1, 2),
    ]
    for num, name, lines in SLIDES:
        flow.append(Paragraph(pdf_markup(f"Слайд {num}. {name}"), S_SLIDE))
        for kind, text in lines:
            flow.append(Paragraph(pdf_markup(text), S_FORMULA if kind == "f" else S_BODY))
    return flow


def build_pdf():
    page_w, page_h = A4
    margin = 0.9 * cm
    top = 0.8 * cm
    bottom = 0.6 * cm
    gutter = 0.55 * cm
    head_h = 1.1 * cm

    head_frame = Frame(
        margin, page_h - top - head_h, page_w - 2 * margin, head_h,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="head",
    )
    col_w = (page_w - 2 * margin - gutter) / 2.0
    col_h = page_h - top - head_h - bottom
    left_frame = Frame(
        margin, bottom, col_w, col_h,
        leftPadding=0, rightPadding=3, topPadding=0, bottomPadding=0, id="left",
    )
    right_frame = Frame(
        margin + col_w + gutter, bottom, col_w, col_h,
        leftPadding=3, rightPadding=0, topPadding=0, bottomPadding=0, id="right",
    )

    def decorate(canvas, _doc):
        canvas.saveState()
        canvas.setStrokeColor(rl_colors.HexColor("#B0BEC5"))
        canvas.setLineWidth(0.4)
        x_line = margin + col_w + gutter / 2.0
        canvas.line(x_line, bottom, x_line, page_h - top - head_h + 0.35 * cm)
        canvas.line(margin, page_h - top - head_h, page_w - margin, page_h - top - head_h)
        canvas.restoreState()

    doc = BaseDocTemplate(
        OUT_PATH, pagesize=A4,
        leftMargin=margin, rightMargin=margin, topMargin=top, bottomMargin=bottom,
        title="Шпаргалка лектора — Лекция 4-доп",
        author="Лекция 4-доп",
    )
    doc.addPageTemplates([PageTemplate(id="two-col", frames=[head_frame, left_frame, right_frame], onPage=decorate)])
    doc.build(build_flowables())

    from pypdf import PdfReader
    pages = len(PdfReader(OUT_PATH).pages)
    print(f"OK: {OUT_PATH} — страниц: {pages}")
    if pages != 1:
        print(f"ВНИМАНИЕ: не помещается на один лист (страниц {pages}) — уменьшите BODY_SIZE/BODY_LEAD")
    return doc


if __name__ == "__main__":
    build_pdf()
