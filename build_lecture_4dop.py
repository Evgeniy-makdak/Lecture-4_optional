#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Лекция 4-доп (внеплановая): Генерация зелёного излучения для аддитивных технологий.
Источник: скриншоты полной текстовки раскадровки (папка screenshots/).
Готовит: PPTX-презентацию (8 слайдов, картинки + таблицы) и PDF-раскадровку для лектора.
"""

import math
import os
import re
from datetime import datetime

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib.colors import darkgray, gray
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, KeepTogether, Table, TableStyle
from reportlab.lib import colors as rl_colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from PIL import Image, ImageDraw, ImageFont


BASE = os.path.dirname(os.path.abspath(__file__))
PPTX_PATH = os.path.join(BASE, "Лекция_4доп_Генерация_зелёного_излучения.pptx")
PDF_PATH = os.path.join(BASE, "Лекция_4доп_Раскадровка_для_лектора.pdf")


def _detect_font():
    # Кандидаты: Windows -> Linux (Liberation/DejaVu) -> macOS. Все — с кириллицей.
    windir = os.environ.get("WINDIR", r"C:\Windows")
    cands = [
        (os.path.join(windir, "Fonts", "arial.ttf"),
         os.path.join(windir, "Fonts", "arialbd.ttf")),
        (os.path.join(windir, "Fonts", "segoeui.ttf"),
         os.path.join(windir, "Fonts", "segoeuib.ttf")),
        ("/usr/share/fonts/liberation-sans/LiberationSans-Regular.ttf",
         "/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
         "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
        ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
        ("/System/Library/Fonts/Supplemental/Arial.ttf",
         "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    ]
    for p, bold in cands:
        if os.path.exists(p):
            return p, bold if os.path.exists(bold) else p
    return None, None


def _detect_math_font():
    windir = os.environ.get("WINDIR", r"C:\Windows")
    for p in [
        os.path.join(windir, "Fonts", "timesi.ttf"),
        os.path.join(windir, "Fonts", "cambriai.ttf"),
        os.path.join(windir, "Fonts", "times.ttf"),
        "/usr/share/fonts/liberation-serif/LiberationSerif-Italic.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Italic.ttf",
        "/usr/share/fonts/dejavu/DejaVuSerif-Italic.ttf",
        "/Library/Fonts/Times Italic.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf",
    ]:
        if os.path.exists(p):
            return p
    return None


FONT_PATH, FONT_BOLD_PATH = _detect_font()
MATH_FONT_PATH = _detect_math_font()
if FONT_PATH:
    pdfmetrics.registerFont(TTFont("CyrFont", FONT_PATH))
    pdfmetrics.registerFont(TTFont("CyrFontBold", FONT_BOLD_PATH or FONT_PATH))
    PDF_FONT, PDF_FONT_BOLD = "CyrFont", "CyrFontBold"
else:
    print(
        "WARNING: не найден TTF-шрифт с кириллицей (Arial/Segoe/DejaVu/Liberation). "
        "PDF будет собран на Helvetica — кириллица отобразится квадратиками! "
        "Установите шрифт с кириллицей и перезапустите сборку."
    )
    PDF_FONT, PDF_FONT_BOLD = "Helvetica", "Helvetica-Bold"

if MATH_FONT_PATH:
    pdfmetrics.registerFont(TTFont("MathFont", MATH_FONT_PATH))
    PDF_MATH = "MathFont"
else:
    PDF_MATH = PDF_FONT


def _detect_symbol_font():
    """Шрифт с ∝, → и др. математическими символами (в Times Italic/Arial их нет)."""
    windir = os.environ.get("WINDIR", r"C:\Windows")
    for name in ("seguisym.ttf", "arialuni.ttf"):
        p = os.path.join(windir, "Fonts", name)
        if os.path.exists(p):
            return p
    return None


SYM_FONT_PATH = _detect_symbol_font()
if SYM_FONT_PATH:
    pdfmetrics.registerFont(TTFont("SymFont", SYM_FONT_PATH))
    PDF_SYM = "SymFont"
else:
    PDF_SYM = PDF_MATH


# ── мини-разметка: ^{...} степень (надстрочно), _{...} индекс (подстрочно) ──

MARKUP_RE = re.compile(r"\^\{([^{}]*)\}|_\{([^{}]*)\}")

GREEK_CHARS = "αβγδεζηθικλμνξπρστυφχψωΑΒΓΔΕΖΗΘΛΞΠΣΦΨΩ∂"


def parse_markup(text):
    """Разбивает строку на токены (текст, режим): режим 'n' | 'sup' | 'sub'."""
    tokens = []
    pos = 0
    for m in MARKUP_RE.finditer(text):
        if m.start() > pos:
            tokens.append((text[pos:m.start()], "n"))
        if m.group(1) is not None:
            tokens.append((m.group(1), "sup"))
        else:
            tokens.append((m.group(2), "sub"))
        pos = m.end()
    if pos < len(text):
        tokens.append((text[pos:], "n"))
    return tokens if tokens else [("", "n")]


def has_greek(text):
    return any(ch in GREEK_CHARS for ch in text)


def pdf_markup(text):
    """Мини-разметка -> ReportLab: ^{x} -> <sup>x</sup>, _{x} -> <sub>x</sub>; греческие — MathFont."""
    if not text:
        return text
    out = []
    for txt, kind in parse_markup(text):
        t = txt
        # греческие и математические символы — отдельным шрифтом
        if any(ch in GREEK_CHARS + "·×" for ch in t):
            t = re.sub(
                r"([%s])" % GREEK_CHARS,
                lambda m: f'<font name="{PDF_MATH}">{m.group(1)}</font>',
                t,
            )
            t = t.replace("·", f'<font name="{PDF_MATH}">·</font>')
        # ∝ и → отсутствуют в Times Italic/Arial — Segoe UI Symbol
        t = t.replace("∝", f'<font name="{PDF_SYM}">∝</font>')
        t = t.replace("→", f'<font name="{PDF_SYM}">→</font>')
        if kind == "sup":
            t = f"<sup>{t}</sup>"
        elif kind == "sub":
            t = f"<sub>{t}</sub>"
        out.append(t)
    return "".join(out)


def pil_safe(text):
    """Для PIL-схем: заменяем только символы, отсутствующие в TTF (остальное рисуем по токенам)."""
    if not text:
        return text
    return str(text)


# ── PIL-шрифты и примитивы ────────────────────────────────────────────────

def _pil_font(size=14, bold=False):
    cands = []
    if bold:
        cands += [
            os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arialbd.ttf"),
            os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "segoeuib.ttf"),
        ]
    cands += [
        FONT_PATH,
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf"),
    ]
    for p in cands:
        if p and os.path.exists(p):
            try:
                return ImageFont.truetype(p, size=size)
            except OSError:
                pass
    return ImageFont.load_default()


def _pil_math_font(size=14, bold=False):
    windir = os.environ.get("WINDIR", r"C:\Windows")
    cands = []
    if bold:
        cands += [os.path.join(windir, "Fonts", "timesbi.ttf")]
    cands += [
        os.path.join(windir, "Fonts", "timesi.ttf"),
        os.path.join(windir, "Fonts", "times.ttf"),
        os.path.join(windir, "Fonts", "cambriai.ttf"),
    ]
    for p in cands:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size=size)
            except OSError:
                continue
    return _pil_font(size)


def _pil_symbol_font(size=12):
    windir = os.environ.get("WINDIR", r"C:\Windows")
    for name in ("seguisym.ttf", "segoeui.ttf", "arial.ttf"):
        path = os.path.join(windir, "Fonts", name)
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size)
            except OSError:
                continue
    return _pil_font(size)


def _tw(d, text, font):
    try:
        return int(d.textlength(text, font=font))
    except Exception:
        return len(text) * 7


def _draw_text(d, xy, text, fill, font):
    d.text(xy, pil_safe(text), fill=fill, font=font)


def _draw_sup(d, x, y, base, sup, font, color="#333"):
    """Простая пара «основание + надстрочный текст» (степени)."""
    fs = _pil_font(max(6, getattr(font, "size", 12) - 5))
    d.text((x, y), base, fill=color, font=font)
    ox = x + _tw(d, base, font)
    d.text((ox, y - 4), sup, fill=color, font=fs)
    return ox + _tw(d, sup, fs)


def _has_glyph(font, ch):
    try:
        m = font.getmask(ch)
        return m.getbbox() is not None
    except Exception:
        return False


def _font_for_text(ch, size, italic_math=True):
    """Шрифт, в котором есть нужный символ (греческие/математика — Times Italic или Segoe UI Symbol)."""
    if italic_math:
        f = _pil_math_font(size)
        if _has_glyph(f, ch):
            return f
    f = _pil_font(size)
    if _has_glyph(f, ch):
        return f
    return _pil_symbol_font(size)


def _markup_width(d, markup, size):
    w = 0
    for txt, kind in parse_markup(markup):
        s = max(8, int(size * (0.62 if kind in ("sup", "sub") else 1.0)))
        f = _fit_font(d, txt, s)
        w += _tw(d, txt, f)
    return w


def _fit_font(d, txt, size):
    """Подбирает шрифт, в котором весь txt читаем (по всем символам)."""
    if italic_like(txt):
        f = _pil_math_font(size)
        if all(_has_glyph(f, ch) or ch == " " for ch in txt):
            return f
    f = _pil_font(size)
    if all(_has_glyph(f, ch) or ch == " " for ch in txt):
        return f
    return _pil_symbol_font(size)


def italic_like(txt):
    return any(ch in GREEK_CHARS for ch in txt)


SPECIAL_PIL = "∝→←↔⇒"


def draw_markup(d, x, y_base, markup, size, color):
    """Рисует строку с мини-разметкой от точки (x, y_base) — базовая линия. Возвращает новый x."""
    for txt, kind in parse_markup(markup):
        if kind in ("sup", "sub"):
            s = max(8, int(size * 0.62))
        else:
            s = size
        # спецсимволы (∝, →, …) рисуем Segoe UI Symbol: в Times Italic/Arial их нет
        for seg in re.findall("[%s]|[^%s]+" % (SPECIAL_PIL, SPECIAL_PIL), txt):
            if len(seg) == 1 and seg in SPECIAL_PIL:
                f = _pil_symbol_font(s)
            else:
                f = _fit_font(d, seg, s)
            if kind == "sup":
                d.text((x, y_base - size * 0.45), seg, fill=color, font=f, anchor="ls")
            elif kind == "sub":
                d.text((x, y_base + size * 0.18), seg, fill=color, font=f, anchor="ls")
            else:
                d.text((x, y_base), seg, fill=color, font=f, anchor="ls")
            x += _tw(d, seg, f)
    return x


def draw_frac(d, x, cy, num, den, size, color):
    """Дробь: num над чертой, den под чертой; cy — вертикальная ось черты. Возвращает новый x."""
    w1 = _markup_width(d, num, size * 0.8)
    w2 = _markup_width(d, den, size * 0.8)
    w = max(w1, w2) + 6
    draw_markup(d, x + (w - w1) / 2, cy - size * 0.35, num, size * 0.8, color)
    draw_markup(d, x + (w - w2) / 2, cy + size * 0.85, den, size * 0.8, color)
    d.line([(x, cy), (x + w, cy)], fill=color, width=2)
    return x + w + 10


def _new_canvas(w, h, bg="white"):
    """RGB с заливкой bg; bg=None → прозрачный RGBA (тёмные слайды)."""
    if bg is None:
        return Image.new("RGBA", (int(w), int(h)), (0, 0, 0, 0))
    return Image.new("RGB", (int(w), int(h)), bg)


def render_formula_strip(markup, size=30, fg="#0d47a1", bg=None, pad=14):
    """Формула-картинка (одна строка). bg=None -> прозрачный фон (для тёмных слайдов)."""
    tmp = Image.new("RGB", (10, 10), "white")
    td = ImageDraw.Draw(tmp)
    w = _markup_width(td, markup, size) + 2 * pad
    h = int(size * 2.1)
    img = _new_canvas(max(20, int(w)), int(h), bg)
    d = ImageDraw.Draw(img)
    draw_markup(d, pad, size * 1.45, markup, size, fg)
    return img


def save_pil_image(img, path):
    img.save(path, "PNG")
    return path


# ── Диаграммы (белые панели; вставляются и в PPTX, и в PDF) ───────────────

def draw_copper_absorption_pil():
    """Медь: A(1070 нм) ~5% против A(535 нм) = 40–60%; справа — схематичная кривая A(λ)."""
    img = Image.new("RGB", (880, 560), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs, fxs = _pil_font(20, True), _pil_font(16), _pil_font(14), _pil_font(12)

    _draw_text(d, (14, 8), "Спектральная поглощательная способность меди A(λ)", fill="#333", font=fh)

    # ── Левая панель: столбики ──
    d.rectangle([(20, 60), (420, 470)], outline="#bf360c", width=2)
    _draw_text(d, (32, 70), "Два окна обработки", fill="#bf360c", font=f)
    # ось выше, подписи — с зазором под ней, чтобы не наползали на горизонталь
    base_y = 392
    scale = 300.0  # 60% -> 300 px
    # 1070 нм
    h1 = 0.05 * scale
    d.rectangle([(70, base_y - h1), (170, base_y)], fill="#e64a19", outline="#870000")
    _draw_text(d, (75, base_y + 18), "1070 нм (ИК)", fill="#870000", font=fs)
    _draw_text(d, (70, base_y - h1 - 26), "≈ 5%", fill="#870000", font=f)
    # 535 нм
    h2 = 0.60 * scale
    d.rectangle([(250, base_y - h2), (350, base_y)], fill="#2e7d32")
    d.line([(250, base_y - 0.40 * scale), (350, base_y - 0.40 * scale)], fill="#a5d6a7", width=2)
    # подпись НАД столбцом (на белом фоне), а не внутри зелёного
    _draw_text(d, (250, base_y - h2 - 26), "40–60%", fill="#1b5e20", font=f)
    _draw_text(d, (255, base_y + 18), "535 нм (зелёный)", fill="#1b5e20", font=fs)
    # ось
    d.line([(50, 140), (50, base_y)], fill="#666", width=2)
    d.line([(50, base_y), (400, base_y)], fill="#666", width=2)
    _draw_text(d, (26, 128), "A, %", fill="#666", font=fxs)
    _draw_text(d, (48, base_y + 40), "в 8–12 раз больше энергии в зону обработки", fill="#555", font=fxs)

    # ── Правая панель: схематичная кривая ──
    d.rectangle([(440, 60), (860, 470)], outline="#1565c0", width=2)
    _draw_text(d, (452, 70), "Кривая A(λ) меди (схематично)", fill="#1565c0", font=f)
    ox, oy = 500, 420  # оси
    d.line([(ox, 100), (ox, oy)], fill="#666", width=2)
    d.line([(ox, oy), (ox + 330, oy)], fill="#666", width=2)
    _draw_text(d, (ox - 44, 96), "A, %", fill="#666", font=fxs)
    _draw_text(d, (ox + 240, oy + 20), "λ, нм", fill="#666", font=fxs)
    # засечки длин волн
    pts = []
    curve = [(450, 42), (535, 50), (600, 38), (700, 22), (800, 12), (1070, 5)]
    def px(lam):
        lam_lo, lam_hi = 400.0, 1100.0
        return ox + (lam - lam_lo) / (lam_hi - lam_lo) * 320
    def py(a):
        return oy - a / 60.0 * 290
    for i in range(len(curve) - 1):
        (l1, a1), (l2, a2) = curve[i], curve[i + 1]
        steps = 24
        for s in range(steps + 1):
            t = s / steps
            lam = l1 + (l2 - l1) * t
            a = a1 + (a2 - a1) * t
            pts.append((px(lam), py(a)))
    d.line(pts, fill="#1565c0", width=3)
    # маркеры
    d.ellipse([(px(1070) - 5, py(5) - 5), (px(1070) + 5, py(5) + 5)], fill="#e64a19")
    _draw_text(d, (px(1070) - 66, py(5) + 10), "1070 нм: ~5%", fill="#870000", font=fs)
    d.ellipse([(px(535) - 5, py(50) - 5), (px(535) + 5, py(50) + 5)], fill="#2e7d32")
    _draw_text(d, (px(535) - 78, py(50) - 30), "535 нм: 40–60%", fill="#1b5e20", font=fs)
    # сетка 5% / 60%
    for a, lab in ((5, "5%"), (60, "60%")):
        d.line([(ox, py(a)), (ox + 330, py(a))], fill="#cfd8dc", width=1)
        _draw_text(d, (ox - 40, py(a) - 8), lab, fill="#90a4ae", font=fxs)

    _draw_text(
        d, (20, 500),
        "Вывод: на 1070 нм 95% мощности отражается; на 535 нм медь поглощает в 8–12 раз больше.",
        fill="#c62828", font=fs,
    )
    return img


def draw_green_gap_pil():
    """«Проблема зелёного диапазона»: зонная диаграмма InGaN + разрыв мощности 0,18 Вт vs 1000 Вт."""
    img = Image.new("RGB", (880, 560), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs, fxs = _pil_font(20, True), _pil_font(15), _pil_font(13), _pil_font(11)

    _draw_text(d, (14, 8), "Прямая генерация 535 нм на InGaN: почему не работает", fill="#333", font=fh)

    # ── Левая панель: зонная диаграмма ──
    d.rectangle([(20, 55), (470, 470)], outline="#1565c0", width=2)
    _draw_text(d, (32, 63), "Зонная диаграмма InGaN (>30% In)", fill="#1565c0", font=f)
    # зона проводимости / валентная зона
    d.rectangle([(60, 110), (430, 150)], fill="#b3e5fc", outline="#0277bd", width=1)
    _draw_text(d, (70, 118), "зона проводимости", fill="#01579b", font=fs)
    d.rectangle([(60, 380), (430, 420)], fill="#ffcdd2", outline="#c62828", width=1)
    _draw_text(d, (70, 388), "валентная зона", fill="#b71c1c", font=fs)
    # стрелка запрещённой зоны
    d.line([(240, 160), (240, 370)], fill="#6a1b9a", width=3)
    d.polygon([(240, 155), (233, 174), (247, 174)], fill="#6a1b9a")
    d.polygon([(240, 375), (233, 356), (247, 356)], fill="#6a1b9a")
    _draw_text(d, (252, 250), "E", fill="#6a1b9a", font=f)
    x = draw_markup(d, 252 + _tw(d, "E", f) + 4, 262, "g = 2,3 эВ", 15, "#6a1b9a")
    # V-дефекты (ловушки)
    for yy in (215, 275, 330):
        d.ellipse([(330, yy - 6), (346, yy + 7)], outline="#e65100", width=2)
    _draw_text(d, (356, 210), "V-дефекты —", fill="#e65100", font=fs)
    _draw_text(d, (356, 228), "ловушки:", fill="#e65100", font=fs)
    _draw_text(d, (356, 246), "энергия → тепло", fill="#e65100", font=fs)
    # электрон и дырка, разведённые пьезополем
    d.ellipse([(150, 165), (172, 187)], fill="#0277bd")
    d.ellipse([(150, 345), (172, 367)], outline="#c62828", width=3)
    draw_markup(d, 176, 181, "e^{-}", 13, "#01579b")
    draw_markup(d, 176, 361, "h^{+}", 13, "#b71c1c")
    d.line([(140, 195), (140, 340)], fill="#9e9e9e", width=2)
    _draw_text(d, (60, 245), "пьезоэлектри-",
               fill="#555", font=fxs)
    _draw_text(d, (60, 262), "ческое поле", fill="#555", font=fxs)
    _draw_text(d, (60, 279), "(QCSE): e и h", fill="#555", font=fxs)
    _draw_text(d, (60, 296), "разведены →", fill="#555", font=fxs)
    _draw_text(d, (60, 313), "↓ рекомбинация", fill="#555", font=fxs)

    # ── Правая панель: разрыв мощности ──
    d.rectangle([(490, 55), (860, 470)], outline="#e64a19", width=2)
    _draw_text(d, (502, 63), "Разрыв мощности — в 10 000 раз", fill="#e64a19", font=f)
    base_y = 420
    # лог-шкала: 0.1 Вт ... 2000 Вт
    def pw(watts):
        return 60 + (math.log10(watts) - math.log10(0.1)) / (math.log10(2000) - math.log10(0.1)) * 260
    # 0,18 Вт
    d.rectangle([(520, base_y - 40), (520 + pw(0.18), base_y)], fill="#bf360c")
    _draw_text(d, (524, base_y - 62), "0,18 Вт (лабораторный образец)", fill="#870000", font=fs)
    _draw_text(d, (524, base_y + 8), "зелёный диодный лазер", fill="#555", font=fxs)
    # 1000 Вт
    d.rectangle([(520, base_y - 130), (520 + pw(1000), base_y - 78)], fill="#1565c0")
    _draw_text(d, (524, base_y - 152), "1000 Вт (нужно для 3D-печати меди)", fill="#0d47a1", font=fs)
    _draw_text(d, (524, base_y + 26), "иттербиевое волокно 1070 нм", fill="#555", font=fxs)
    _draw_text(d, (520, 90), "Разрыв: 10 000 раз", fill="#c62828", font=fh)

    _draw_text(
        d, (20, 500),
        "«Проблема зелёного диапазона» (green gap): >30% индия → V-дефекты + QCSE → мощность ≤ 0,18 Вт. Нужен другой путь.",
        fill="#c62828", font=fs,
    )
    return img


def draw_asym_well_pil():
    """Потенциальная яма электрона: симметричная (линейный отклик) vs асимметричная (нелинейность)."""
    img = Image.new("RGB", (880, 530), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs, fxs = _pil_font(19, True), _pil_font(15), _pil_font(13), _pil_font(11)

    _draw_text(d, (14, 8), "Почему в ниобате лития отклик нелинейный", fill="#333", font=fh)

    # ── Левая панель: симметричная чаша ──
    d.rectangle([(20, 50), (420, 445)], outline="#2e7d32", width=2)
    _draw_text(d, (32, 58), "Обычная среда (стекло, вода)", fill="#2e7d32", font=f)
    cx, cy = 220, 290
    # симметричная парабола
    pts = [(cx - 150 + i * 3, cy - (i * 3 - 150) ** 2 / 90) for i in range(0, 101)]
    d.line(pts, fill="#2e7d32", width=3)
    d.ellipse([(cx - 10, cy - 60), (cx + 10, cy - 40)], fill="#333")
    _draw_text(d, (60, 120), "E пот.", fill="#555", font=fxs)
    _draw_text(d, (30, 445 - 22), "смещение электрона →", fill="#555", font=fxs)
    for i, line in enumerate([
        "яма симметрична — колебания",
        "электрона повторяют частоту ω",
        "поля. Только преломление:",
        "новых частот не возникает.",
    ]):
        _draw_text(d, (32, 350 + i * 22), line, fill="#555", font=fxs)

    # ── Правая панель: асимметричная яма ──
    d.rectangle([(450, 50), (860, 445)], outline="#e64a19", width=2)
    _draw_text(d, (462, 58), "Нелинейный кристалл (ниобат лития)", fill="#e64a19", font=f)
    cx2, cy2 = 655, 290
    # асимметричная чаша: левый край крутой, правый пологий
    pts = []
    for i in range(-150, 151, 3):
        x = cx2 + i
        if i < 0:
            y = cy - (i / 150.0) ** 2 * 150 * 1.6
        else:
            y = cy - (i / 150.0) ** 2 * 150 * 0.8
        pts.append((x, y))
    d.line(pts, fill="#e64a19", width=3)
    d.ellipse([(cx2 - 22, cy2 - 52), (cx2 - 2, cy2 - 32)], fill="#333")
    _draw_text(d, (470, 100), "E пот.", fill="#555", font=fxs)
    _draw_text(d, (462, 130), "круче", fill="#c62828", font=fxs)
    _draw_text(d, (800, 155), "пологий", fill="#c62828", font=fxs)
    for i, line in enumerate([
        "яма асимметрична —",
        "колебания несимметричны,",
        "в отклике появляются",
        "гармоники: 2ω, 3ω, …",
    ]):
        _draw_text(d, (462, 350 + i * 22), line, fill="#555", font=fxs)

    _draw_text(d, (20, 478), "Шарик в круглой чаше → линейный отклик; шарик в несимметричной чаше → гармоники.", fill="#555", font=fs)
    return img


def draw_shg_ppln_pil():
    """Схема SHG: ИК 1070 нм -> PPLN (период Λ) -> зелёный 535 нм; волны ω и 2ω."""
    img = Image.new("RGB", (880, 520), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs, fxs = _pil_font(19, True), _pil_font(15), _pil_font(13), _pil_font(11)

    _draw_text(d, (14, 8), "Генерация второй гармоники (SHG) в кристалле PPLN", fill="#333", font=fh)

    # ИК-лазер
    d.rectangle([(30, 120), (170, 190)], outline="#870000", width=2)
    _draw_text(d, (40, 138), "ИК-лазер", fill="#870000", font=f)
    x = draw_markup(d, 40, 186, "λ = 1070 нм (ω)", 13, "#870000")
    # стрелка накачки
    d.line([(175, 155), (300, 155)], fill="#c62828", width=5)
    d.polygon([(300, 155), (284, 147), (284, 163)], fill="#c62828")

    # кристалл PPLN
    d.rectangle([(305, 90), (600, 230)], outline="#6a1b9a", width=3)
    _draw_text(d, (318, 98), "PPLN — периодически поляризованный", fill="#6a1b9a", font=fs)
    _draw_text(d, (318, 114), "ниобат лития", fill="#6a1b9a", font=fs)
    # домены: чередующиеся стрелки
    dom_x = 320
    dom_w = 32
    yy = 190
    for i in range(8):
        col = "#ce93d8" if i % 2 == 0 else "#b39ddb"
        d.rectangle([(dom_x + i * dom_w, yy), (dom_x + (i + 1) * dom_w, yy + 90)], fill=col, outline="#6a1b9a", width=1)
        ax = dom_x + i * dom_w + dom_w // 2
        if i % 2 == 0:
            d.line([(ax, yy + 42), (ax, yy + 12)], fill="#4a148c", width=3)
            d.polygon([(ax, yy + 6), (ax - 6, yy + 18), (ax + 6, yy + 18)], fill="#4a148c")
        else:
            d.line([(ax, yy + 12), (ax, yy + 42)], fill="#4a148c", width=3)
            d.polygon([(ax, yy + 48), (ax - 6, yy + 36), (ax + 6, yy + 36)], fill="#4a148c")
    # период Λ
    d.line([(dom_x, yy + 66), (dom_x + 2 * dom_w, yy + 66)], fill="#4a148c", width=1)
    d.line([(dom_x, yy + 60), (dom_x, yy + 72)], fill="#4a148c", width=1)
    d.line([(dom_x + 2 * dom_w, yy + 60), (dom_x + 2 * dom_w, yy + 72)], fill="#4a148c", width=1)
    draw_markup(d, dom_x + 6, yy + 84, "Λ ≈ 6,96 мкм", 13, "#4a148c")
    _draw_text(d, (318, 300), "знак χ^{(2)} перевёрнут каждые Λ микрон", fill="#555", font=fxs)

    # зелёный выход
    d.line([(605, 155), (730, 155)], fill="#2e7d32", width=5)
    d.polygon([(742, 155), (726, 147), (726, 163)], fill="#2e7d32")
    d.rectangle([(745, 120), (865, 190)], outline="#1b5e20", width=2)
    _draw_text(d, (753, 138), "зелёный", fill="#1b5e20", font=f)
    draw_markup(d, 753, 178, "λ = 535 нм (2ω)", 13, "#1b5e20")

    # нижняя панель: волны ω и 2ω
    d.rectangle([(30, 330), (860, 460)], outline="#90a4ae", width=1)
    draw_markup(d, 40, 344, "Поле волны: E(t) = E_{0}·cos(ω·t)  →  отклик на 2ω: cos(2ω·t)", 13, "#333")
    # красная синусоида ω
    pts = []
    for i in range(0, 320):
        t = i / 320.0
        pts.append((70 + i, 415 - 28 * math.sin(t * 2 * math.pi * 3)))
    d.line(pts, fill="#c62828", width=2)
    _draw_text(d, (400, 402), "накачка ω (1070 нм)", fill="#870000", font=fxs)
    # зелёная синусоида 2ω
    pts = []
    for i in range(0, 320):
        t = i / 320.0
        pts.append((70 + i, 445 - 16 * math.sin(t * 2 * math.pi * 6)))
    d.line(pts, fill="#2e7d32", width=2)
    _draw_text(d, (400, 438), "вторая гармоника 2ω (535 нм)", fill="#1b5e20", font=fxs)

    return img


def draw_twopass_pil():
    """Двухпроходная схема с активной термокомпенсацией: лазер -> PPLN (печь) -> зеркало -> клин."""
    img = Image.new("RGB", (900, 560), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs, fxs = _pil_font(19, True), _pil_font(15), _pil_font(13), _pil_font(11)

    _draw_text(d, (14, 8), "Моё предложение: двухпроходная схема с активной клиновидной термокомпенсацией", fill="#333", font=fh)

    # ── Волоконный ИК-лазер ──
    d.rectangle([(30, 120), (150, 200)], outline="#870000", width=2)
    _draw_text(d, (38, 140), "волоконный", fill="#870000", font=fs)
    _draw_text(d, (38, 158), "лазер", fill="#870000", font=fs)
    _draw_text(d, (38, 176), "1070 нм, 1 кВт", fill="#870000", font=fxs)

    # прямой проход (верхняя линия)
    d.line([(155, 150), (720, 150)], fill="#c62828", width=4)
    d.polygon([(732, 150), (716, 142), (716, 158)], fill="#c62828")

    # коллиматор/линза
    d.rectangle([(190, 120), (215, 200)], fill="#ffe0b2", outline="#e65100", width=2)
    _draw_text(d, (176, 104), "линза", fill="#e65100", font=fxs)

    # PPLN в печи
    d.rectangle([(300, 110), (470, 220)], outline="#6a1b9a", width=3)
    d.rectangle([(290, 96), (480, 234)], outline="#9e9d24", width=2)
    _draw_text(d, (310, 128), "PPLN", fill="#6a1b9a", font=f)
    draw_markup(d, 310, 156, "L = 5 см, Λ ≈ 6,96 мкм", 12, "#6a1b9a")
    _draw_text(d, (298, 76), "печь ±0,1 °C", fill="#827717", font=fs)
    _draw_text(d, (310, 170), "1-й проход: КПД 10–15%", fill="#555", font=fxs)
    _draw_text(d, (310, 186), "2-й проход: КПД до 15–20%", fill="#555", font=fxs)

    # зеркало
    d.rectangle([(720, 110), (745, 230)], fill="#b0bec5", outline="#455a64", width=2)
    _draw_text(d, (700, 236), "зеркало (R > 99% @ 1070)", fill="#455a64", font=fxs)

    # обратный проход (нижняя линия, справа налево)
    d.line([(732, 210), (732, 260), (240, 260)], fill="#ef6c00", width=4)
    d.polygon([(228, 260), (244, 252), (228, 268)], fill="#ef6c00")
    _draw_text(d, (500, 264), "обратный проход: удвоение длины взаимодействия", fill="#e65100", font=fxs)

    # кварцевый клин в обратном ходе
    d.polygon([(150, 236), (230, 236), (150, 284)], fill="#b3e5fc", outline="#0277bd", width=2)
    _draw_text(d, (60, 234), "подвижный", fill="#01579b", font=fxs)
    _draw_text(d, (60, 252), "кварцевый клин", fill="#01579b", font=fxs)
    _draw_text(d, (60, 270), "dn/dT > 0,", fill="#01579b", font=fxs)
    _draw_text(d, (60, 288), "1070 и 535 нм", fill="#01579b", font=fxs)
    _draw_text(d, (234, 302), "компенсация фазового сдвига между проходами", fill="#01579b", font=fxs)

    # зелёный выход вниз к обработке
    d.line([(150, 290), (150, 360)], fill="#2e7d32", width=5)
    d.polygon([(150, 372), (142, 356), (158, 356)], fill="#2e7d32")
    d.rectangle([(90, 375), (330, 445)], outline="#1b5e20", width=2)
    _draw_text(d, (100, 393), "зелёный 535 нм", fill="#1b5e20", font=f)
    _draw_text(d, (100, 418), "→ 3D-печать / сварка меди", fill="#1b5e20", font=fs)

    # ── нижняя сводка (без рамки) ──
    for i, line in enumerate([
        "Печь: термостабилизация квазифазового синхронизма (иначе КПД падает почти до нуля).",
        "Кварцевый клин: активная компенсация набега фазы зеркала — печь и клин дополняют друг друга.",
        "Прецедент: Imeshev, Fejer (1998); новизна — киловаттный уровень мощности и внешняя термокомпенсация.",
    ]):
        _draw_text(d, (42, 462 + i * 20), line, fill="#555", font=fxs)
    return img


def draw_sinc2_pil():
    """График sinc²(Δk·L/2): почему критично условие синхронизма."""
    img = Image.new("RGB", (560, 355), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(17, True), _pil_font(13), _pil_font(11)

    _draw_text(d, (14, 8), "Почему появляется sinc²(Δk·L/2)", fill="#333", font=fh)
    ox, oy = 90, 200
    d.line([(ox, 60), (ox, oy)], fill="#666", width=2)
    d.line([(ox - 30, oy), (ox + 420, oy)], fill="#666", width=2)
    _draw_text(d, (ox - 56, 54), "sinc²", fill="#666", font=f)
    draw_markup(d, ox + 340, oy + 20, "Δk·L/2", 13, "#666")

    pts = []
    for i in range(0, 800):
        x = -6.0 + 12.0 * i / 800.0
        if abs(x) < 1e-9:
            v = 1.0
        else:
            v = (math.sin(x) / x) ** 2
        pts.append((ox + (x + 6) / 12.0 * 400, oy - v * 120))
    d.line(pts, fill="#1565c0", width=3)

    # максимум
    d.ellipse([(ox + 200 - 4, oy - 124), (ox + 200 + 4, oy - 116)], fill="#2e7d32")
    _draw_text(d, (ox + 150, oy - 145), "sinc²(0) = 1 — максимум", fill="#1b5e20", font=fs)
    # нули
    for k in (-2, -1, 1, 2):
        xx = ox + (math.pi * k + 6) / 12.0 * 400
        d.ellipse([(xx - 3, oy - 3), (xx + 3, oy + 3)], fill="#c62828")
    _draw_text(d, (ox + 150, oy + 30), "рассогласование фаз → гашение", fill="#555", font=fs)

    for i, line in enumerate([
        "Волны, рождённые в разных точках кристалла,",
        "складываются синфазно только при Δk → 0.",
        "Точный подбор Λ (или подстройка температуры)",
        "держит sinc² близким к 1.",
    ]):
        _draw_text(d, (22, 238 + i * 18), line, fill="#555", font=fs)
    return img


def draw_economy_pil():
    """Экономика: готовый зелёный лазер против нашего модуля."""
    img = Image.new("RGB", (760, 380), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(18, True), _pil_font(14), _pil_font(12)

    _draw_text(d, (14, 8), "Экономика решения", fill="#333", font=fh)

    base = 150.0  # тыс. $
    def x_of(v):
        return 320 + v / 320.0 * 400
    d.line([(320, 60), (320, 280)], fill="#666", width=2)
    # готовый лазер: подпись белым ВНУТРИ столбца (столбец широкий)
    d.rectangle([(320, 80), (x_of(300), 130)], fill="#bf360c")
    _draw_text(d, (330, 94), "150–300 тыс. $", fill="white", font=f)
    _draw_text(d, (40, 96), "промышленный зелёный лазер", fill="#870000", font=fs)
    _draw_text(d, (40, 114), "100–200 Вт (готовый станок)", fill="#555", font=fs)
    # наш модуль: столбец узкий — подпись СПРАВА от него, тёмно-синим на белом
    d.rectangle([(320, 180), (x_of(30), 230)], fill="#1565c0")
    _draw_text(d, (368, 192), "на порядок ниже", fill="#0d47a1", font=f)
    _draw_text(d, (40, 196), "внешний SHG-модуль", fill="#0d47a1", font=fs)
    _draw_text(d, (40, 214), "модернизация волоконного лазера", fill="#555", font=fs)
    d.line([(320, 280), (720, 280)], fill="#666", width=2)
    _draw_text(d, (640, 288), "тыс. $", fill="#666", font=fs)

    _draw_text(
        d, (40, 320),
        "Та же медь, те же 535 нм — но на порядок ниже стоимость владения.",
        fill="#1b5e20", font=fs,
    )
    return img


# ── Формульные картинки ───────────────────────────────────────────────────

def formula_photon_energy(fg="#0d47a1", bg="white"):
    """E = hc/λ и расчёт для 535 нм."""
    img = _new_canvas(760, 170, bg)
    d = ImageDraw.Draw(img)
    draw_markup(d, 20, 60, "E = h·c / λ", 34, fg)
    draw_markup(d, 260, 60, "для λ = 535 нм:  E ≈ 3,7·10^{-19} Дж ≈ 2,3 эВ", 21, fg)
    draw_markup(d, 20, 130, "(h — постоянная Планка,  c — скорость света,  λ — длина волны)", 16, "#8899aa" if bg is None else "#666")
    return img


def formula_polarization_series(fg="#0d47a1", bg="white"):
    """P = ε₀(χ⁽¹⁾E + χ⁽²⁾E² + χ⁽³⁾E³ + …)."""
    return render_formula_strip(
        "P = ε_{0}·( χ^{(1)}·E  +  χ^{(2)}·E^{2}  +  χ^{(3)}·E^{3}  +  … )",
        size=30, fg=fg, bg=bg,
    )


def formula_cos2(fg="#0d47a1", bg="white"):
    """Вывод через cos²(ωt)."""
    img = _new_canvas(820, 250, bg)
    d = ImageDraw.Draw(img)
    draw_markup(d, 20, 50, "E(t) = E_{0}·cos(ω·t),      ω = 2π·c / λ", 24, fg)
    draw_markup(d, 20, 120, "P^{(2)}(t) ∝ χ^{(2)}·E_{0}^{2}·cos^{2}(ω·t)", 26, fg)
    draw_markup(d, 20, 195, "cos^{2}(ω·t) = (1 + cos(2ω·t)) / 2", 26, fg)
    return img


def formula_sync_condition(fg="#0d47a1", bg="white"):
    return render_formula_strip("Δk = k_{2ω} − 2·k_{ω} = 2π / Λ", size=30, fg=fg, bg=bg)


def formula_efficiency(fg="#0d47a1", bg="white"):
    """η = P₂ω/P_ω · (2ω²d²eff L²)/(ε₀c³n²ω n₂ω) · (P_ω/A) · sinc²(ΔkL/2)."""
    w, h = 900, 185
    img = _new_canvas(w, h, bg)
    d = ImageDraw.Draw(img)
    cy = 82
    x = draw_markup(d, 16, cy + 4, "η  =", 30, fg)
    x = draw_frac(d, x, cy, "P_{2ω}", "P_{ω}", 26, fg)
    x = draw_markup(d, x, cy, "·", 26, fg)
    x = draw_frac(d, x, cy, "2ω^{2}·d^{2}_{eff}·L^{2}", "ε_{0}·c^{3}·n^{2}_{ω}·n_{2ω}", 26, fg)
    x = draw_markup(d, x, cy, "·", 26, fg)
    x = draw_frac(d, x, cy, "P_{ω}", "A", 26, fg)
    x = draw_markup(d, x, cy, "·  sinc²( Δk·L/2 )", 26, fg)
    draw_markup(d, 16, 162, "(уравнения связанных волн, приближение неистощённой накачки)", 14, "#8899aa" if bg is None else "#666")
    return img


def formula_lambda_half(fg="#0d47a1", bg="white"):
    return render_formula_strip("λ_{2ω} = λ_{ω} / 2 = 1070 нм / 2 = 535 нм", size=30, fg=fg, bg=bg)


# ── Раскадровка ───────────────────────────────────────────────────────────

READING_WPM = 130

SECTION_STAGE = {
    "1. ТИТУЛЬНЫЙ": (0.17, "Пауза на вход: слайд титула. Вывести на экран тему."),
    "2. ПЛАН ЛЕКЦИИ": (0.17, "Показать пять пунктов плана на слайде."),
    "3. ПОЧЕМУ МЕДЬ — ПРОБЛЕМА ДЛЯ ИК-ЛАЗЕРОВ": (
        2.5,
        "Доска: таблица поглощения меди на разных длинах волн; график спектральной поглощательной способности.",
    ),
    "4. ПРЯМАЯ ГЕНЕРАЦИЯ ЗЕЛЁНОГО СВЕТА: ПОЧЕМУ НЕ РАБОТАЕТ": (
        2.8,
        "Доска: формула энергии фотона E = hc/λ; схема зонной диаграммы InGaN; таблица «проблема зелёного диапазона».",
    ),
    "5. НЕЛИНЕЙНАЯ ОПТИКА: ГЕНЕРАЦИЯ ВТОРОЙ ГАРМОНИКИ": (
        3.7,
        "Доска: разложение поляризации в ряд; вывод формулы cos²(ωt) = (1 + cos(2ωt))/2; схема PPLN.",
    ),
    "6. МОЁ ПРЕДЛОЖЕНИЕ: ДВУХПРОХОДНАЯ СХЕМА С ТЕРМОКОМПЕНСАЦИЕЙ": (
        3.1,
        "Доска: схема двухпроходной установки; формула для КПД; расчёт для 1 кВт.",
    ),
    "7. ИТОГОВЫЕ ТЕЗИСЫ": (0.0, "Вывести пять тезисов на слайд."),
    "8. ЗАКЛЮЧЕНИЕ": (0.0, "Вывести на слайд основные выводы."),
}


def _count_words(text):
    return len(re.findall(r"[A-Za-zА-Яа-яЁё0-9]+(?:[-'][A-Za-zА-Яа-яЁё0-9]+)*", text))


def _fmt_clock(minutes):
    sec = max(0, int(round(minutes * 60)))
    return f"{sec // 60:02d}:{sec % 60:02d}"


def _to_min(x):
    mm, ss = x.split(":")
    return int(mm) + int(ss) / 60.0


def _time_script(blocks):
    out = []
    cum = 0.0
    total_words = 0
    for title, timing, paragraphs in blocks:
        words = sum(_count_words(p) for p in paragraphs)
        total_words += words
        pause_min, cue = SECTION_STAGE.get(title, (0.0, ""))
        m = re.search(r"(\d\d:\d\d)\s*[–-]\s*(\d\d:\d\d)", timing)
        wall_min = None
        if m:
            wall_min = max(0.0, _to_min(m.group(2)) - _to_min(m.group(1)))
        if wall_min is None:
            wall_min = words / READING_WPM + float(pause_min)
        start, end = cum, cum + wall_min
        label = f"{timing} | тайминг от начала: {_fmt_clock(start)} – {_fmt_clock(end)}"
        out.append((title, label, paragraphs, cue))
        cum = end
    return out, total_words, cum


SCRIPT = [
    ("1. ТИТУЛЬНЫЙ", "00:00 – 00:46 (стенка 46 сек = чтение 36 сек + доска/паузы 10 сек)", [
        "Добрый день. Мы продолжаем курс. В предыдущих лекциях мы разобрали, как оптика управляет плотностью мощности — тем, что определяет, плавится металл или нет.",
        "Сегодня мы разберём проблему, которая стоит перед всей отраслью аддитивных технологий: почему медь — один из самых востребованных металлов в электронике и энергетике — так плохо поддаётся лазерной обработке.",
        "И главное: как мы предлагаем эту проблему решить. Речь пойдёт о генерации зелёного излучения с длиной волны 535 нанометров — и о моём предложении, которое делает эту технологию экономически доступной.",
    ]),
    ("2. ПЛАН ЛЕКЦИИ", "00:46 – 01:53 (стенка 67 сек = чтение 57 сек + доска/паузы 10 сек)", [
        "План будет таким.",
        "Первое. Почему медь — проблема для инфракрасных лазеров. Разберём спектральную поглощательную способность и покажем цифры.",
        "Второе. Почему зелёный свет решает проблему. Сравним поглощение на длинах волн 1070 и 535 нанометров.",
        "Третье. Прямая генерация зелёного света: почему полупроводниковые лазеры на нитриде галлия не могут дать нужную мощность. Это фундаментальное ограничение, называемое «проблемой зелёного диапазона».",
        "Четвёртое. Нелинейная оптика: как получить зелёный свет из инфракрасного. Генерация второй гармоники — базовая формула.",
        "Пятое. Моё предложение: двухпроходная схема с активной термокомпенсацией. Почему это позволяет снизить стоимость на порядок.",
    ]),
    ("3. ПОЧЕМУ МЕДЬ — ПРОБЛЕМА ДЛЯ ИК-ЛАЗЕРОВ", "01:53 – 06:41 (стенка 4,8 мин = чтение 2,3 мин + доска/паузы 2,5 мин)", [
        "Давайте начнём с физики. Почему медь так плохо обрабатывается волоконными лазерами?",
        "Вспомним ключевое понятие — спектральная поглощательная способность. Это доля энергии лазерного луча, которая поглощается материалом, а не отражается. Обозначим её буквой A (от английского absorbance).",
        "Для меди ситуация драматична. На длине волны 1070 нанометров — это стандарт для мощных иттербиевых волоконных лазеров — поглощательная способность меди составляет всего около 5 процентов. Это означает, что 95 процентов энергии просто отражается от поверхности.",
        "Что происходит с этими пятью процентами? Они нагревают тонкий поверхностный слой. Но этого недостаточно для стабильного расплава. Процесс становится нестабильным: то есть плавление, то нет. Появляются поры, брызги, несплавления.",
        "Теперь посмотрим на зелёный диапазон. На длине волны 535 нанометров поглощательная способность меди возрастает до 40–60 процентов. Это в 8–12 раз больше!",
        "Что это означает на практике? При той же мощности лазера в зону обработки попадает в 8–12 раз больше энергии. Процесс становится стабильным. Шов формируется без разрывов. Именно поэтому зелёные лазеры так востребованы для сварки медных шин в электромобилях и для 3D-печати медных теплообменников.",
        "Вывод: технически зелёный лазер идеален для меди. Проблема не в физике, а в экономике.",
    ]),
    ("4. ПРЯМАЯ ГЕНЕРАЦИЯ ЗЕЛЁНОГО СВЕТА: ПОЧЕМУ НЕ РАБОТАЕТ", "06:41 – 12:08 (стенка 5,4 мин = чтение 2,6 мин + доска/паузы 2,8 мин)", [
        "Закономерный вопрос: почему бы не создать лазер, который сразу генерирует зелёный свет? Такие лазеры существуют — это полупроводниковые диодные лазеры на основе нитрида индия-галлия, InGaN.",
        "Но здесь мы сталкиваемся с фундаментальным ограничением, которое в научной литературе называется «проблемой зелёного диапазона» (green gap).",
        "Объясню физику. Энергия фотона определяется формулой:",
        "E = hc/λ. Расшифровка: E — энергия фотона (Дж); h — постоянная Планка (Дж·с); c — скорость света (3·10^{8} м/с); λ — длина волны (м). Для зелёного света 535 нм энергия фотона составляет примерно 3,7·10^{-19} Дж, или 2,3 электрон-вольта.",
        "Чтобы получить такую энергию в полупроводниковом лазере, нужно создать материал с шириной запрещённой зоны 2,3 эВ. Для InGaN это требует добавления более 30 процентов индия в кристаллическую решётку.",
        "И вот здесь начинаются проблемы. Атомы индия значительно крупнее атомов галлия. При такой высокой концентрации кристаллическая решётка становится нестабильной. Возникают дефекты, которые называются V-дефектами. Они работают как ловушки для электронов — вместо излучения света энергия превращается в тепло.",
        "Дополнительный фактор — квантово-размерный эффект Штарка (QCSE). Пьезоэлектрическое поле внутри квантовой ямы разводит электрон и дырку в пространстве, снижая вероятность их рекомбинации с излучением фотона.",
        "Результат: эффективность падает, пороговый ток растёт, а мощность ограничивается. Современные лабораторные образцы зелёных диодных лазеров демонстрируют мощность 180 милливатт — это 0,18 ватта. Для 3D-печати меди нужны киловатты. Разрыв в 10 000 раз.",
        "Вывод: прямая генерация зелёного света на полупроводниках не может конкурировать с мощными волоконными лазерами. Нужен другой путь.",
    ]),
    ("5. НЕЛИНЕЙНАЯ ОПТИКА: ГЕНЕРАЦИЯ ВТОРОЙ ГАРМОНИКИ", "12:08 – 17:50 (стенка 5,7 мин = чтение 2,3 мин + доска/паузы 3,4 мин)", [
        "Идея: взять мощный и дешёвый инфракрасный лазер 1070 нм и «превратить» его излучение в зелёное 535 нм с помощью нелинейной оптики — генерации второй гармоники (Second Harmonic Generation, SHG).",
        "Откуда нелинейность? Отклик среды на поле описывается степенным рядом: P = ε_{0}·(χ^{(1)}·E + χ^{(2)}·E^{2} + χ^{(3)}·E^{3} + …).",
        "Расшифровка символов: P — поляризация среды (Кл/м^{2}), дипольный момент единицы объёма; ε_{0} — электрическая постоянная, 8,85·10^{-12} Ф/м, Ф (фарад) = Кл/В; E — напряжённость поля волны (В/м); χ^{(1)} — линейная восприимчивость, безразмерная, отвечает за преломление; χ^{(2)} — квадратичная (м/В), отвечает за удвоение частоты; χ^{(3)} — кубическая (м^{2}/В^{2}), мала — пренебрегаем. Индекс в скобках — порядок нелинейности, а не степень числа.",
        "Линейный член не создаёт новых частот — он лишь обычное преломление. Кубический для ниобата лития на порядки меньше квадратичного. Значит, вторую гармонику генерирует только член χ^{(2)}·E^{2} — его и оставляем.",
        "Пусть на кристалл падает волна E(t) = E_{0}·cos(ω·t): E_{0} — амплитуда поля; ω = 2πc/λ — циклическая частота (рад/с); t — время (с). Подставляем её во второй член: P^{(2)}(t) ∝ χ^{(2)}·E_{0}^{2}·cos^{2}(ω·t); знак «∝» — «пропорционально», постоянные множители опущены.",
        "Тригонометрическое тождество: cos^{2}(ω·t) = (1 + cos(2ω·t))/2. Постоянная часть 1/2 — оптическое выпрямление: постоянное поле, не колеблется и не излучает — для SHG бесполезна. А часть cos(2ω·t) — колебание с удвоенной частотой: длина волны по ω = 2πc/λ уменьшается ровно вдвое, λ_{2ω} = 1070/2 = 535 нм.",
        "Ключевая проблема: из-за дисперсии волны ω и 2ω идут в кристалле с разными скоростями, рассинхронизируются — и зелёные волны, рождённые в разных точках, гасят друг друга.",
        "Решение — периодически поляризованный ниобат лития (PPLN — Periodically Poled Lithium Niobate): каждые Λ микрон кристаллическая ось перевёрнута на 180°, знак χ^{(2)} меняется, и набег фазы компенсируется. Условие квазисинхронизма: Δk = k_{2ω} − 2k_{ω} = 2π/Λ, где k — волновые числа, Λ — период решётки переполяризации. Для нашей задачи Λ ≈ 6,96 мкм.",
    ]),
    ("6. МОЁ ПРЕДЛОЖЕНИЕ: ДВУХПРОХОДНАЯ СХЕМА С ТЕРМОКОМПЕНСАЦИЕЙ", "17:50 – 26:14 (стенка 8,4 мин = чтение 5,3 мин + доска/паузы 3,1 мин)", [
        "Теперь — самое важное. Почему зелёные такие дорогие и что я предлагаю.",
        "Стандартная схема — однопроходная. Луч проходит через кристалл один раз. КПД преобразования при этом составляет около 10–15 процентов. Для получения 150 ватт зелёного света нужен киловатт инфракрасного. Остальное уходит в тепло.",
        "Моё предложение — двухпроходная схема с активной термокомпенсацией.",
        "Прежде чем описывать схему, я должен быть честен: двухпроходная схема сама по себе не является уникальной. Исследования в этой области ведутся с 1998 года. В классической работе Имешева и Феджера (Imeshev, Fejer) была предложена и экспериментально проверена двухпроходная схема с клиновидным кристаллом для компенсации фазовых сдвигов, вносимых зеркалами. В их эксперименте с 5-сантиметровым кристаллом PPLN было показано ожидаемое увеличение эффективности преобразования.",
        "Что именно является моим предложением? Моё предложение — это применение двухпроходной схемы с активной клиновидной термокомпенсацией для мощных непрерывных волоконных лазеров (1 кВт и выше) с целью обработки меди в аддитивных технологиях. Вот что здесь нового:",
        "1. Масштаб мощности: существующие работы с двухпроходными схемами выполнены на маломощных лазерах. При переходе к киловаттному уровню возникает проблема тепловой деградации кристалла, которую стандартные схемы не решают.",
        "2. Активная клиновидная компенсация: вместо пассивной термостабилизации кристалла в печи я предлагаю использовать подвижную клиновидную пластину из плавленого кварца, которая механически подстраивает фазу в реальном времени, компенсируя тепловые искажения.",
        "Целевое применение: фокус именно на аддитивных технологиях для меди, где требуется высокая стабильность мощности в течение длительного времени.",
        "Важное уточнение: печь и клин — это не альтернативы, а дополнения.",
        "Печь (термостабилизация кристалла) нужна в любом случае. Квазифазовый синхронизм в PPLN критически зависит от температуры: при изменении температуры на 1 °C меняется показатель преломления, и условие синхронизма нарушается. Поэтому кристалл обязательно помещают в термостабилизированную печь с точностью ±0,1 °C. Без неё КПД упадёт с 15% до почти нуля.",
        "Кварцевый клин решает другую задачу — компенсацию фазового сдвига между прямым и обратным проходом в двухпроходной схеме. Когда луч отражается от зеркала и идёт обратно через кристалл, он накапливает дополнительный фазовый сдвиг. Клин позволяет этот сдвиг скомпенсировать, чтобы оба прохода работали синфазно.",
        "Почему именно плавленый кварц? Выбор материала для клиновидной пластины — это не произвольное решение. Вот моя логика как физика-теоретика:",
        "Аргумент 1: Известные оптические свойства. Плавленый кварц (fused silica) — это один из самых изученных оптических материалов. Его показатель преломления и его температурная зависимость (термооптический коэффициент dn/dT) измерены с высокой точностью. Например, в работе NASA Goddard Space Flight Center (2013) абсолютный показатель преломления плавленого кварца измерен с точностью до ±1·10^{-5} в диапазоне от 0,4 до 2,6 мкм. Это означает, что я могу точно рассчитать, как пластина будет влиять на фазу.",
        "Аргумент 2: Положительный термооптический коэффициент. Исследования 1971 года показали, что термооптический коэффициент плавленого кварца положителен и составляет около +9·10^{-6} К^{-1} при комнатной температуре. Это критически важно: при нагреве показатель преломления увеличивается, что позволяет компенсировать тепловые искажения в кристалле PPLN.",
        "Аргумент 3: Прозрачность в нужном диапазоне. Плавленый кварц прозрачен как для инфракрасного излучения 1070 нм, так и для зелёного 535 нм. Это означает, что клиновидная пластина может стоять на пути луча без существенных потерь.",
        "Аргумент 4: Технологичность. Плавленый кварц легко обрабатывается, шлифуется и полируется. Из него можно изготовить клин с высокой точностью угла.",
        "Аргумент 5: Прецедент в литературе. В работе Имешева и Феджера (1998) для компенсации фазовых сдвигов использовался именно клиновидный элемент. Хотя в той работе клин был изготовлен из самого нелинейного кристалла, сама идея использования клина для фазовой компенсации подтверждена экспериментально. Моё предложение — использовать отдельный клин из плавленого кварца, что технологически проще и дешевле.",
        "Таким образом, выбор плавленого кварца обоснован как с точки зрения его известных оптических свойств, так и с точки зрения технологичности и наличия прецедентов в литературе.",
        "Теперь подробно сравню мою схему с работой Имешева и Феджера. Цель и там, и там одна: скомпенсировать фазовый сдвиг, который накапливается между прямым и обратным проходом из-за отражения от зеркала. Это одна и та же физическая задача. Различия — в реализации, и они принципиальны.",
        "Различие первое — где стоит компенсирующий клин. У Имешева клин — это сам нелинейный кристалл PPLN (периодически поляризованный ниобат лития). У меня клин — отдельная пластина из плавленого кварца, стоящая вне кристалла.",
        "Различие второе — активность компенсации. У Имешева клин пассивный: угол задан при изготовлении, и после этого изменить его нельзя. У меня клин подвижный: его можно смещать в реальном времени, подстраивая фазу.",
        "Почему при мощностях 1 киловатт и выше выигрывает моя схема? Четыре причины.",
        "Причина первая — тепловая деградация PPLN (периодически поляризованного ниобата лития). При киловаттных мощностях кристалл сильно нагревается: даже доли процента поглощения дают десятки ватт тепла в объёме кристалла. Возникает тепловой дефазинг: показатель преломления меняется, условие квазифазового синхронизма нарушается, КПД падает. Если клин — сам PPLN (периодически поляризованный ниобат лития), как у Имешева, то тепловые искажения одновременно портят и нелинейный процесс, и работу клина: два эффекта накладываются, и разделить их невозможно. Если же клин — отдельная кварцевая пластина, как у меня, то кристалл термостабилизируется в печи с точностью ±0,1 °C — стандартная практика для PPLN (периодически поляризованный ниобат лития), — а фазу компенсирует внешний подвижный клин. Печь и клин работают независимо, каждый решает свою задачу.",
        "Причина вторая — технологичность. Шлифовка PPLN (периодически поляризованного ниобата лития) под углом с микронной точностью — дорогая и сложная операция. Плавленый кварц, напротив, дёшев, легко шлифуется и полируется, его оптические свойства хорошо изучены.",
        "Причина третья — ремонтопригодность. Если клин встроен в кристалл, выход из строя любого элемента означает замену всего узла. Отдельную кварцевую пластину можно заменить, не трогая дорогой кристалл.",
        "Причина четвёртая — динамическая подстройка. По мере нагрева кристалла фаза уходит, и я смещаю клин, возвращая синфазность в реальном времени. У Имешева такой возможности нет: его клин настроен один раз и навсегда.",
        "Формула эффективности преобразования: η = P_{2ω}/P_{ω} · (2ω^{2}·d_{eff}^{2}·L^{2}) / (ε_{0}·c^{3}·n_{ω}^{2}·n_{2ω}) · (P_{ω}/A) · sinc^{2}(Δk·L/2).",
        "Откуда взялась эта формула? Она выводится из системы уравнений связанных волн (coupled-wave equations) — фундаментальных уравнений нелинейной оптики, впервые полученных в 1960-х годах. Эти уравнения описывают, как амплитуды двух волн — накачки и второй гармоники — изменяются по мере распространения через нелинейную среду.",
        "В приближении неистощённой накачки (когда мощность инфракрасного луча >> мощности зелёного) решение этих уравнений даёт приведённую формулу. Расшифровка символов — в таблице на слайде.",
        "Почему здесь появляется sinc? Функция sinc(x) определяется как sinc(x) = sin(x)/x. Она возникает из-за фазового рассогласования между волнами накачки и второй гармоники.",
        "Представьте: по мере распространения через кристалл зелёная волна, рождённая в начале, и зелёная волна, рождённая в конце, накапливают разность фаз. Если эта разность мала (Δk·L/2 → 0), все волны складываются синфазно, и sinc^{2}(0) = 1 — максимальная эффективность. Если разность фаз растёт, волны начинают гасить друг друга, и sinc^{2} падает.",
        "Именно поэтому для эффективной генерации нужно точное выполнение условия синхронизма — либо через точный подбор периода решётки Λ, либо через температурную подстройку.",
        "Замечание об экономике. Реализация данной установки будет стоить на порядок меньше, чем покупка полноценного мощного станка зелёного лазера. По данным рынка, промышленные зелёные лазеры мощностью 100–200 Вт стоят от 150 000 до 300 000 долларов. Наша модернизация существующего волоконного лазера потребует лишь внешнего модуля, стоимость которого на порядок ниже.",
    ]),
    ("7. ИТОГОВЫЕ ТЕЗИСЫ", "26:14 – 28:00 (стенка 1,8 мин = чтение 1,8 мин)", [
        "Сфиксируем пять тезисов.",
        "Первое. Медь плохо поглощает инфракрасное излучение (5%), но хорошо поглощает зелёное (40–60%). Это делает зелёные лазеры идеальными для обработки меди.",
        "Второе. Прямая генерация зелёного света на полупроводниках ограничена «проблемой зелёного диапазона»: мощность не превышает 180 милливатт. Для промышленности нужны киловатты.",
        "Третье. Генерация второй гармоники в кристалле PPLN позволяет преобразовать инфракрасное излучение 1070 нм в зелёное 535 нм с КПД до 15–20%.",
        "Четвёртое. Моё предложение — двухпроходная схема с активной клиновидной термокомпенсацией из плавленого кварца для мощных волоконных лазеров. Существующие работы с двухпроходными схемами выполнены на маломощных лазерах, и их масштабирование на киловаттный уровень является новой инженерной задачей.",
        "Пятое. Реализация данной установки будет стоить на порядок меньше, чем покупка полноценного зелёного лазера.",
    ]),
    ("8. ЗАКЛЮЧЕНИЕ", "28:00 – 28:30 (стенка 30 сек = чтение 21 сек)", [
        "Итак, мы разобрали физику генерации зелёного излучения для аддитивных технологий. Ключевые моменты: медь требует зелёного света, прямая генерация ограничена «проблемой зелёного диапазона», а нелинейная оптика позволяет обойти это ограничение.",
        "Моё предложение — двухпроходная схема с клиновидной термокомпенсацией — открывает путь к созданию доступных зелёных лазеров для промышленности.",
    ]),
]


def _full_script():
    timed, *_rest = _time_script(SCRIPT)
    return timed


# ── Таблицы (один источник для PPTX и PDF) ────────────────────────────────

CHI_TABLE = {
    "headers": ["Символ", "Название", "Единица", "Физический смысл"],
    "rows": [
        ["P", "Поляризация среды", "Кл/м^{2}", "Дипольный момент единицы объёма"],
        ["ε_{0}", "Электрическая постоянная", "Ф/м", "8,85·10^{-12} Ф/м; Ф (фарад) = Кл/В"],
        ["E", "Напряжённость поля волны", "В/м", "Поле лазерного луча"],
        ["χ^{(1)}", "Линейная восприимчивость", "безразм.", "Отвечает за преломление"],
        ["χ^{(2)}", "Квадратичная восприимчивость", "м/В", "Отвечает за удвоение частоты"],
        ["χ^{(3)}", "Кубическая восприимчивость", "м^{2}/В^{2}", "Мала, пренебрегаем"],
    ],
}

COPPER_TABLE = {
    "headers": ["Длина волны", "A(λ) меди", "Следствие"],
    "rows": [
        ["1070 нм (ИК, волокно)", "≈ 5%", "95% отражается; поры, брызги, несплавления"],
        ["535 нм (зелёный)", "40–60%", "В 8–12 раз больше энергии в зоне; стабильный шов"],
    ],
}

GREEN_GAP_TABLE = {
    "headers": ["Параметр", "Значение"],
    "rows": [
        ["Энергия фотона 535 нм", "≈ 3,7·10^{-19} Дж ≈ 2,3 эВ"],
        ["Нужная ширина запрещённой зоны", "2,3 эВ"],
        ["Концентрация индия в InGaN", "> 30%"],
        ["Дефекты решётки", "V-дефекты (ловушки электронов)"],
        ["Квантово-размерный эффект Штарка", "QCSE: e и h разведены полем"],
        ["Мощность лабораторных образцов", "≤ 0,18 Вт (нужны кВт)"],
    ],
}

ETA_TABLE = {
    "headers": ["Символ", "Название", "Физический смысл"],
    "rows": [
        ["η", "Эффективность", "Отношение P_{2ω} / P_{ω}"],
        ["d_{eff}", "Эффективная нелинейность", "для MgO:PPLN d_{33} ≈ 25 пм/В"],
        ["L", "Длина кристалла", "мм (здесь 50 мм)"],
        ["n_{ω}, n_{2ω}", "Показатели преломления", "≈ 2,13 и ≈ 2,23"],
        ["A", "Площадь пятна", "мм^{2}"],
    ],
}


# ── PPTX ──────────────────────────────────────────────────────────────────

def build_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    ACCENT = RGBColor(0x00, 0x96, 0x88)
    ACCENT2 = RGBColor(0xE6, 0x4A, 0x19)
    GREEN = RGBColor(0x2E, 0x7D, 0x32)
    LT = RGBColor(0xEE, 0xEE, 0xEE)
    GT = RGBColor(0x75, 0x75, 0x75)
    FLIGHT = RGBColor(0xBD, 0xE0, 0xFF)

    tmpdir = os.path.join(BASE, "_diagrams_tmp_lecture4dop")
    os.makedirs(tmpdir, exist_ok=True)

    def abg(s):
        f = s.background.fill
        f.solid()
        f.fore_color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

    def _fill_rich(p, text, size, color, bold=False):
        """Абзац с мини-разметкой: ^{x} — надстрочное, _{x} — подстрочное (настоящие степени)."""
        for txt, kind in parse_markup(text):
            if not txt:
                continue
            r = p.add_run()
            r.text = txt
            f = r.font
            f.size = Pt(size * (0.65 if kind in ("sup", "sub") else 1.0))
            f.bold = bold
            f.color.rgb = color
            f.name = "Cambria Math" if has_greek(txt) else "Calibri"
            if kind in ("sup", "sub"):
                rPr = r._r.get_or_add_rPr()
                rPr.set("baseline", "30000" if kind == "sup" else "-25000")

    def at(s, l, t, w, h, tx, sz=18, b=False, c=LT, a=PP_ALIGN.LEFT):
        tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = a
        _fill_rich(p, tx, sz, LT if c is None else c, b)
        return tb

    def bullets(s, items, left=1.5, top=1.55, w=6.2, sz=17, space=10, height=5.6):
        tb = s.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(height))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            _fill_rich(p, it, sz, LT)
            p.space_after = Pt(space)
        return tb

    def hdr(s, title, accent=ACCENT, sz=26):
        abg(s)
        at(s, 1.5, 0.30, 10.3, 0.9, title, sz, True)
        ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(1.22), Inches(10.3), Pt(3))
        ln.fill.solid()
        ln.fill.fore_color.rgb = accent
        ln.line.fill.background()

    def add_table(s, table_def, left, top, width, height, font=11, col_widths=None):
        headers, rows = table_def["headers"], table_def["rows"]
        tbl = s.shapes.add_table(1 + len(rows), len(headers), Inches(left), Inches(top), Inches(width), Inches(height)).table
        if col_widths:
            for ci, w in enumerate(col_widths):
                tbl.columns[ci].width = Inches(w)
        for ci, htxt in enumerate(headers):
            cell = tbl.cell(0, ci)
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            _fill_rich(p, htxt, font, LT, True)
            p.alignment = PP_ALIGN.CENTER
            cell.fill.solid()
            cell.fill.fore_color.rgb = ACCENT
        for ri, row in enumerate(rows, start=1):
            bgc = RGBColor(0x24, 0x24, 0x3E) if ri % 2 else RGBColor(0x1E, 0x1E, 0x32)
            for ci, val in enumerate(row):
                cell = tbl.cell(ri, ci)
                p = cell.text_frame.paragraphs[0]
                p.text = ""
                _fill_rich(p, val, font, LT)
                p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
                cell.fill.solid()
                cell.fill.fore_color.rgb = bgc
        return tbl

    def add_pic(s, img, left, top, width, max_right=13.05, max_bottom=7.28):
        """Вставка PNG с сохранением пропорций; не выходит за край слайда 13,333×7,5″."""
        w_px, h_px = img.size
        height = width * h_px / float(w_px)
        if left + width > max_right:
            width = max(0.4, max_right - left)
            height = width * h_px / float(w_px)
        if top + height > max_bottom:
            height = max(0.4, max_bottom - top)
            width = height * w_px / float(h_px)
        png = os.path.join(tmpdir, f"pic_{abs(hash((left, top, width, id(img)))) % 99999}.png")
        save_pil_image(img, png)
        return s.shapes.add_picture(
            png, Inches(left), Inches(top), width=Inches(width), height=Inches(height),
        )

    # ── Слайд 1: титульный ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    abg(s)
    at(s, 0.9, 1.1, 11.5, 1.6, "ЛЕКЦИЯ 4-ДОП", 20, True, GREEN)
    at(s, 0.9, 1.7, 11.5, 1.6, "Генерация зелёного излучения (535 нм)", 34, True, LT)
    at(s, 0.9, 2.9, 11.5, 1.0, "для аддитивных технологий", 26, False, FLIGHT)
    at(s, 0.9, 4.1, 11.5, 1.0, "Медь: от проблемы ИК-обработки к двухпроходной SHG-схеме с термокомпенсацией", 18, False, GT)
    at(s, 0.9, 5.6, 11.5, 0.8, "Раскадровка: 8 слайдов • ~28 минут", 15, False, GT)

    # ── Слайд 2: план ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    hdr(s, "2. ПЛАН ЛЕКЦИИ")
    bullets(s, [
        "1. Почему медь — проблема для ИК-лазеров (A(λ) в цифрах)",
        "2. Почему зелёный свет решает проблему: 1070 vs 535 нм",
        "3. Прямая генерация: проблема зелёного диапазона InGaN",
        "4. Нелинейная оптика: генерация второй гармоники (SHG)",
        "5. Моё предложение: двухпроходная схема с активной термокомпенсацией",
    ], 1.5, 1.7, 10.3, 20, 16)

    # ── Слайд 3: медь — проблема для ИК-лазеров ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    hdr(s, "3. ПОЧЕМУ МЕДЬ — ПРОБЛЕМА ДЛЯ ИК-ЛАЗЕРОВ", ACCENT2, 24)
    bullets(s, [
        "A(λ) — доля энергии луча, поглощённая материалом (absorbance)",
        "1070 нм (иттербиевое волокно): A ≈ 5% — 95% отражается",
        "5% хватает лишь на тонкий поверхностный слой",
        "Результат: поры, брызги, несплавления",
        "535 нм: A = 40–60% — в 8–12 раз больше",
        "При той же мощности — в 8–12 раз больше энергии в зоне",
        "Вывод: проблема не в физике, а в экономике",
    ], 0.55, 1.45, 6.1, 14, 8)
    add_table(s, COPPER_TABLE, 0.55, 5.55, 6.1, 1.6, font=11, col_widths=[1.9, 1.2, 3.7])
    add_pic(s, draw_copper_absorption_pil(), 6.9, 1.5, 6.2)

    # ── Слайд 4: проблема зелёного диапазона ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    hdr(s, "4. ПРЯМАЯ ГЕНЕРАЦИЯ: ПОЧЕМУ НЕ РАБОТАЕТ", ACCENT2, 24)
    bullets(s, [
        "InGaN-диоды существуют — но упираются в проблему зелёного диапазона (green gap)",
        "E = hc/λ; для 535 нм E ≈ 3,7·10^{-19} Дж ≈ 2,3 эВ",
        "Нужна ширина запрещённой зоны 2,3 эВ → >30% индия",
        "Крупные атомы In → V-дефекты (ловушки: энергия → тепло)",
        "QCSE: пьезополе разводит электрон и дырку — рекомбинация падает",
        "Итог: ≤ 0,18 Вт против нужных кВт — разрыв 10 000 раз",
    ], 0.55, 1.45, 5.9, 14, 12)
    add_pic(s, formula_photon_energy(fg="#bde0ff", bg=None), 0.55, 5.9, 5.9)
    add_pic(s, draw_green_gap_pil(), 6.7, 1.5, 6.2)
    add_table(s, GREEN_GAP_TABLE, 6.7, 5.6, 6.3, 1.75, font=9, col_widths=[3.3, 3.0])

    # ── Слайд 5: нелинейная оптика ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    hdr(s, "5. НЕЛИНЕЙНАЯ ОПТИКА: ВТОРАЯ ГАРМОНИКА", ACCENT2, 24)
    bullets(s, [
        "SHG: 1070 нм (ω) → 535 нм (2ω) в нелинейном кристалле",
        "Отклик среды: P = ε_{0}·(χ^{(1)}E + χ^{(2)}E^{2} + χ^{(3)}E^{3} + …)",
        "Оставляем член χ^{(2)}·E^{2} — удвоение частоты",
        "cos^{2}(ω·t) = (1 + cos(2ω·t))/2: постоянная часть — оптическое выпрямление, колебательная — 2ω",
        "λ_{2ω} = 1070/2 = 535 нм",
        "Дисперсия → рассинхронизация → гашение",
        "PPLN: знак χ^{(2)} меняется каждые Λ ≈ 6,96 мкм",
        "Условие синхронизма: Δk = k_{2ω} − 2k_{ω} = 2π/Λ",
    ], 0.5, 1.42, 5.5, 12.5, 8)
    add_table(s, CHI_TABLE, 6.25, 1.4, 6.9, 2.4, font=10, col_widths=[0.9, 2.5, 1.0, 2.4])
    add_pic(s, draw_shg_ppln_pil(), 6.25, 3.95, 5.9)

    # ── Слайд 6: двухпроходная схема ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    hdr(s, "6. МОЁ ПРЕДЛОЖЕНИЕ: ДВУХПРОХОДНАЯ СХЕМА С ТЕРМОКОМПЕНСАЦИЕЙ", ACCENT2, 21)
    bullets(s, [
        "Однопроход: КПД 10–15%; 150 Вт зелёного ← 1 кВт ИК",
        "Два прохода: удвоение длины, КПД до 15–20%",
        "Печь ±0,1 °C — квазифазовый синхронизм в PPLN",
        "Кварцевый клин — компенсация фазы обратного прохода",
        "dn/dT ≈ +9·10^{-6} К^{-1}; прозрачен на 1070 и 535 нм",
        "Прецедент: Imeshev, Fejer (1998); новизна — киловаттный масштаб",
        "Экономика: на порядок дешевле готового зелёного лазера",
    ], 0.45, 1.38, 4.7, 12, 6, 3.05)
    add_pic(s, draw_sinc2_pil(), 0.45, 4.42, 4.55)
    add_pic(s, draw_twopass_pil(), 5.28, 1.36, 6.65)
    add_pic(s, formula_efficiency(fg="#bde0ff", bg=None), 5.28, 5.62, 7.65)

    # ── Слайд 7: итоговые тезисы ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    hdr(s, "7. ИТОГОВЫЕ ТЕЗИСЫ", ACCENT2, 26)
    bullets(s, [
        "1. Медь: ИК — 5% поглощения, зелёный — 40–60% → зелёный идеален для меди",
        "2. Прямая генерация: проблема зелёного диапазона, ≤ 0,18 Вт; нужны киловатты",
        "3. SHG в PPLN: 1070 нм → 535 нм с КПД до 15–20%",
        "4. Двухпроходная схема с активной клиновидной термокомпенсацией из плавленого кварца — новая инженерная задача на киловаттном уровне",
        "5. Реализация на порядок дешевле покупки готового зелёного лазера",
    ], 0.7, 1.45, 11.9, 15, 8, 2.55)
    add_pic(s, draw_economy_pil(), 3.35, 4.05, 6.6)

    # ── Слайд 8: заключение ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    hdr(s, "8. ЗАКЛЮЧЕНИЕ", ACCENT2, 26)
    bullets(s, [
        "Медь требует зелёного света: A(535 нм) в 8–12 раз выше, чем на 1070 нм",
        "Прямая генерация ограничена «проблемой зелёного диапазона» InGaN",
        "Нелинейная оптика (SHG в PPLN, Λ ≈ 6,96 мкм) обходит ограничение",
        "Моё предложение: двухпроходная схема с клиновидной термокомпенсацией",
        "Итог: доступные зелёные лазеры (535 нм) для аддитивных технологий",
    ], 0.9, 1.7, 11.5, 18, 14)

    prs.save(PPTX_PATH)
    print(f"OK: {PPTX_PATH}")


# ── PDF (раскадровка для лектора) ─────────────────────────────────────────

def build_pdf():
    doc = SimpleDocTemplate(
        PDF_PATH, pagesize=A4,
        leftMargin=1.6 * cm, rightMargin=1.6 * cm,
        topMargin=1.2 * cm, bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()
    sh2 = ParagraphStyle("H2", parent=styles["Heading2"], fontName=PDF_FONT_BOLD, fontSize=13, leading=17)
    sh3 = ParagraphStyle("H3", parent=styles["Heading3"], fontName=PDF_FONT_BOLD, fontSize=11, leading=14)
    body = ParagraphStyle("B", parent=styles["Normal"], fontName=PDF_FONT, fontSize=10, leading=13, alignment=TA_JUSTIFY)
    hint = ParagraphStyle("H", parent=styles["Normal"], fontName=PDF_FONT, fontSize=9, leading=12, textColor=darkgray, alignment=TA_JUSTIFY)
    stage = ParagraphStyle(
        "Stage", parent=styles["Normal"], fontName=PDF_FONT, fontSize=9,
        leading=12, textColor=darkgray, alignment=TA_JUSTIFY, leftIndent=6,
    )
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontName=PDF_FONT, fontSize=8.5, leading=11)
    cell_style_c = ParagraphStyle("CellC", parent=cell_style, alignment=TA_CENTER)
    cell_head = ParagraphStyle("CellH", parent=cell_style, fontName=PDF_FONT_BOLD, textColor=rl_colors.HexColor("#FFFFFF"))

    def rl_table(table_def, col_widths_cm):
        headers, rows = table_def["headers"], table_def["rows"]
        data = [[Paragraph(pdf_markup(h), cell_head) for h in headers]]
        for row in rows:
            data.append([
                Paragraph(pdf_markup(v), cell_style if ci == 0 else cell_style_c)
                for ci, v in enumerate(row)
            ])
        t = Table(data, colWidths=[w * cm for w in col_widths_cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#009688")),
            ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.HexColor("#90A4AE")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.HexColor("#FFFFFF"), rl_colors.HexColor("#ECEFF1")]),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    pdf_tmp = os.path.join(BASE, "_diagrams_tmp_lecture4dop_pdf")
    os.makedirs(pdf_tmp, exist_ok=True)

    diagram_for_slide = {
        "3. ПОЧЕМУ МЕДЬ — ПРОБЛЕМА ДЛЯ ИК-ЛАЗЕРОВ": [draw_copper_absorption_pil],
        "4. ПРЯМАЯ ГЕНЕРАЦИЯ ЗЕЛЁНОГО СВЕТА: ПОЧЕМУ НЕ РАБОТАЕТ": [draw_green_gap_pil],
        "5. НЕЛИНЕЙНАЯ ОПТИКА: ГЕНЕРАЦИЯ ВТОРОЙ ГАРМОНИКИ": [draw_shg_ppln_pil, draw_asym_well_pil],
        "6. МОЁ ПРЕДЛОЖЕНИЕ: ДВУХПРОХОДНАЯ СХЕМА С ТЕРМОКОМПЕНСАЦИЕЙ": [draw_twopass_pil],
        "7. ИТОГОВЫЕ ТЕЗИСЫ": [draw_economy_pil],
    }

    # Дополнительные материалы по слайдам: формулы и таблицы
    formula_by_slide = {
        "4. ПРЯМАЯ ГЕНЕРАЦИЯ ЗЕЛЁНОГО СВЕТА: ПОЧЕМУ НЕ РАБОТАЕТ": formula_photon_energy,
        "5. НЕЛИНЕЙНАЯ ОПТИКА: ГЕНЕРАЦИЯ ВТОРОЙ ГАРМОНИКИ": lambda: formula_cos2(),
        "6. МОЁ ПРЕДЛОЖЕНИЕ: ДВУХПРОХОДНАЯ СХЕМА С ТЕРМОКОМПЕНСАЦИЕЙ": formula_efficiency,
    }
    table_by_slide = {
        "3. ПОЧЕМУ МЕДЬ — ПРОБЛЕМА ДЛЯ ИК-ЛАЗЕРОВ": (COPPER_TABLE, [4.5, 2.5, 9.0]),
        "4. ПРЯМАЯ ГЕНЕРАЦИЯ ЗЕЛЁНОГО СВЕТА: ПОЧЕМУ НЕ РАБОТАЕТ": (GREEN_GAP_TABLE, [7.5, 8.5]),
        "5. НЕЛИНЕЙНАЯ ОПТИКА: ГЕНЕРАЦИЯ ВТОРОЙ ГАРМОНИКИ": (CHI_TABLE, [2.0, 5.5, 2.5, 6.0]),
        "6. МОЁ ПРЕДЛОЖЕНИЕ: ДВУХПРОХОДНАЯ СХЕМА С ТЕРМОКОМПЕНСАЦИЕЙ": (ETA_TABLE, [2.5, 5.5, 8.0]),
    }

    script = _full_script()
    total_words = sum(sum(_count_words(p) for p in paras) for _, _, paras, _ in script)
    total_wall_min = _time_script(SCRIPT)[2]
    wall_min_label = f"{total_wall_min:.1f}".replace(".", ",")

    elements = []
    elements.append(Paragraph(pdf_markup("ЛЕКЦИЯ 4-ДОП: ГЕНЕРАЦИЯ ЗЕЛЁНОГО ИЗЛУЧЕНИЯ (535 нм) ДЛЯ АДДИТИВНЫХ ТЕХНОЛОГИЙ"), sh2))
    elements.append(Paragraph(
        f"Раскадровка • (8 слайдов) • ~{total_words} слов • {datetime.now().strftime('%d.%m.%Y')}",
        hint,
    ))
    elements.append(Paragraph(
        "Тайминг «стенка» = произнесение текста (~130 слов/мин) + работа у доски/со схемой. "
        "Строки «Доска / пауза:» — ремарки лектору, не читаются вслух. "
        f"Целевой хронометраж лекции: {wall_min_label} минуты.",
        hint,
    ))
    elements.append(Spacer(1, 8))

    for si, (title, timing, paragraphs, cue) in enumerate(script):
        block = []
        block.append(Paragraph(f"СЛАЙД: {title} | {timing}", sh3))
        if cue:
            block.append(Paragraph(pdf_markup(f"Доска / пауза: {cue}"), stage))
        for para in paragraphs:
            block.append(Paragraph(pdf_markup(para), body))
        block.append(Spacer(1, 5))

        diagrams = diagram_for_slide.get(title) or []
        for di, draw_fn in enumerate(diagrams):
            img = draw_fn()
            safe = re.sub(r"[^\w\-]+", "_", title)[:44]
            png_path = os.path.join(pdf_tmp, f"pdf_{safe}_{di}.png")
            img.save(png_path, "PNG")
            w_px, h_px = img.size
            img_w = 16.6 * cm if w_px >= 800 else 13.5 * cm
            block.append(RLImage(png_path, width=img_w, height=img_w * (h_px / w_px)))
            block.append(Spacer(1, 6))

        if title in formula_by_slide:
            fimg = formula_by_slide[title]()
            safe = re.sub(r"[^\w\-]+", "_", title)[:30] + "_formula"
            fpath = os.path.join(pdf_tmp, f"{safe}.png")
            fimg.save(fpath, "PNG")
            w_px, h_px = fimg.size
            fw = 14.5 * cm
            block.append(RLImage(fpath, width=fw, height=fw * (h_px / w_px)))
            block.append(Spacer(1, 4))

        if title in table_by_slide:
            tdef, widths = table_by_slide[title]
            block.append(rl_table(tdef, widths))
            block.append(Spacer(1, 6))

        elements.extend(block)

    def _page_num(canvas, _doc):
        canvas.saveState()
        canvas.setFont(PDF_FONT, 10)
        canvas.setFillColor(gray)
        canvas.drawCentredString(A4[0] / 2, 0.85 * cm, str(canvas.getPageNumber()))
        canvas.restoreState()

    doc.build(elements, onFirstPage=_page_num, onLaterPages=_page_num)
    print(f"OK: {PDF_PATH} (~{total_words} слов)")


if __name__ == "__main__":
    print("=" * 70)
    print("СБОРКА ЛЕКЦИИ 4-ДОП")
    print("=" * 70)
    build_pptx()
    build_pdf()
    print("=" * 70)
    print("ГОТОВО")
    print("=" * 70)
