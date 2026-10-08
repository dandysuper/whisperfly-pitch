#!/usr/bin/env python3
"""Build WhisperFly-Pitch-RU.pptx — the editable PowerPoint version of the deck.

Nine slides, in the order the competition asks for:

    01 титульный лист   02 введение       03 уникальность
    04 КПД              05 план реализации 06 ожидаемые результаты
    07 финансовая часть 08 команда         09 заключение

The palette is not hand-picked: it is computed from the same VARSITY token formula
the CSS uses (oklch, --tp-h: 262), so the .pptx and the HTML deck stay in sync.

    python3 build_pptx.py

Requires: python-pptx  (pip install python-pptx)
"""

import math

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ── Palette, derived from the theme's oklch tokens ───────────────────────────

HUE = 262


def _encode(channel: float) -> int:
    """Linear-light sRGB channel -> 0..255 gamma-encoded byte."""
    channel = max(0.0, min(1.0, channel))
    if channel <= 0.0031308:
        value = 12.92 * channel
    else:
        value = 1.055 * (channel ** (1 / 2.4)) - 0.055
    return int(round(value * 255))


def oklch(lightness: float, chroma: float, hue: float) -> RGBColor:
    """oklch(...) -> RGBColor, via Oklab and linear sRGB."""
    radians = math.radians(hue)
    a = chroma * math.cos(radians)
    b = chroma * math.sin(radians)

    l_ = lightness + 0.3963377774 * a + 0.2158037573 * b
    m_ = lightness - 0.1055613458 * a - 0.0638541728 * b
    s_ = lightness - 0.0894841775 * a - 1.2914855480 * b

    ll, mm, ss = l_**3, m_**3, s_**3

    red = 4.0767416621 * ll - 3.3077115913 * mm + 0.2309699292 * ss
    green = -1.2684380046 * ll + 2.6097574011 * mm - 0.3413193965 * ss
    blue = -0.0041960863 * ll - 0.7034186147 * mm + 1.7076147010 * ss

    return RGBColor(_encode(red), _encode(green), _encode(blue))


BG = RGBColor(0xF3, 0xF1, 0xEA)        # --tp-bg
SURF = RGBColor(0xFD, 0xFC, 0xF8)      # --tp-surf
SURF2 = RGBColor(0xE5, 0xE1, 0xD4)     # --tp-surf-2
INK = RGBColor(0x1C, 0x24, 0x47)       # --tp-ink
MUTE = RGBColor(0x5A, 0x61, 0x80)      # --tp-mute
ACC_INK = RGBColor(0xFD, 0xFC, 0xF8)   # --tp-acc-ink
ACC = oklch(0.44, 0.16, HUE)
ACC2 = oklch(0.78, 0.14, HUE + 155)

# ── Type ─────────────────────────────────────────────────────────────────────

FONT_DISPLAY = "Unbounded"
FONT_BODY = "Rubik"
FONT_MONO = "JetBrains Mono"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.62)
CONTENT_W = SLIDE_W - MARGIN * 2
BORDER = Pt(2.25)

# ── Shape / text helpers ─────────────────────────────────────────────────────


def blank(prs):
    """A slide with no placeholders."""
    return prs.slides.add_slide(prs.slide_layouts[6])


def box(slide, x, y, w, h, fill=None, line=None, radius=None,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, line_w=BORDER):
    """Add a shape. `radius` (0..0.5) tunes the corner rounding on rounded rects."""
    shp = slide.shapes.add_shape(shape, int(x), int(y), int(w), int(h))
    shp.shadow.inherit = False

    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = radius

    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill

    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = line_w

    return shp


def label(slide, x, y, w, h, runs, *, size=14, font=FONT_BODY, color=INK,
          bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
          line_spacing=1.1, space_after=0, caps=False):
    """A text box. `runs` is a string, or a list of (text, {overrides}) tuples."""
    tb = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor

    if isinstance(runs, str):
        runs = [(runs, {})]

    paragraph = tf.paragraphs[0]
    paragraph.alignment = align
    paragraph.line_spacing = line_spacing
    paragraph.space_after = Pt(space_after)

    for text, override in runs:
        run = paragraph.add_run()
        run.text = text.upper() if override.get("caps", caps) else text
        run.font.size = Pt(override.get("size", size))
        run.font.name = override.get("font", font)
        run.font.bold = override.get("bold", bold)
        run.font.color.rgb = override.get("color", color)
        run.font.italic = override.get("italic", False)

    return tb


def fitted(slide, x, y, w, h, text, **kwargs):
    """A text box that may need a second, shorter line set — kept for symmetry."""
    return label(slide, x, y, w, h, text, **kwargs)


def shape_label(shape, text, *, size=10, font=FONT_MONO, color=INK, bold=True,
                caps=False, align=PP_ALIGN.CENTER):
    """Write centred text inside an existing autoshape."""
    tf = shape.text_frame
    tf.word_wrap = False
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text.upper() if caps else text
    run.font.size = Pt(size)
    run.font.name = font
    run.font.bold = bold
    run.font.color.rgb = color
    return shape


def pill(slide, x, y, w, h, text, *, fill=SURF, color=INK, line=INK,
         size=9, font=FONT_MONO, bold=True, caps=False, radius=0.5):
    """A rounded chip with centred text."""
    shp = box(slide, x, y, w, h, fill=fill, line=line,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)
    return shape_label(shp, text, size=size, font=font, color=color,
                       bold=bold, caps=caps)


def kicker(slide, text, x=None, y=Inches(0.36), gold=False):
    """The small uppercase chip that opens most slides."""
    x = MARGIN if x is None else x
    width = Inches(0.28) + Inches(0.088) * len(text)
    return pill(slide, x, y, width, Inches(0.32), text,
                fill=ACC2 if gold else ACC,
                color=INK if gold else ACC_INK, size=8, caps=True)


def title(slide, runs, y=Inches(0.76), size=27, width=None):
    """Slide headline. `runs` accepts the same shapes as `label`."""
    width = width or (CONTENT_W - Inches(1.6))
    return label(slide, MARGIN, y, width, Inches(0.9), runs,
                 size=size, font=FONT_DISPLAY, color=INK, bold=True,
                 line_spacing=1.0)


def subtitle(slide, text, y=Inches(1.42), width=None, size=10.5):
    """The one-line explainer under a headline."""
    width = width or (CONTENT_W - Inches(1.4))
    return label(slide, MARGIN, y, width, Inches(0.6), text,
                 size=size, color=MUTE, line_spacing=1.32)


def slide_number(slide, text, x=None, y=Inches(0.34)):
    """The mono pill in the top-right corner, mirroring the HTML deck."""
    width = Inches(0.42)
    x = (SLIDE_W - MARGIN - width) if x is None else x
    return pill(slide, x, y, width, Inches(0.3), text, fill=SURF, color=INK,
                size=8)


def spine(slide):
    """Recolour the slide background and draw the letterman stripe down the left edge."""
    bg = box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=BG, shape=MSO_SHAPE.RECTANGLE)
    bg.shadow.inherit = False

    segment = Inches(0.62)
    y = 0.0
    index = 0
    while y < SLIDE_H:
        height = min(segment, SLIDE_H - y)
        stripe = box(slide, 0, Emu(int(y)), Inches(0.15), Emu(int(height)),
                     fill=ACC if index % 2 == 0 else ACC2, shape=MSO_SHAPE.RECTANGLE)
        stripe.shadow.inherit = False
        y += segment
        index += 1


# ── Layout helpers ───────────────────────────────────────────────────────────

GAP = Inches(0.20)
HEAD_TOP = Inches(0.36)
BODY_TOP = Inches(1.86)


def columns(count, gap=GAP, total=None, margin=None):
    """Even columns across the content width. Returns (width, [x, ...])."""
    total = CONTENT_W if total is None else total
    margin = MARGIN if margin is None else margin
    width = int((total - gap * (count - 1)) / count)
    xs = [int(margin + i * (width + gap)) for i in range(count)]
    return width, xs


def card(slide, x, y, w, h, *, fill=SURF, line=INK, radius=0.09):
    return box(slide, x, y, w, h, fill=fill, line=line,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)


def card_head(slide, x, y, w, text, color=INK, size=12.5):
    return label(slide, x, y, w, Inches(0.40), text, size=size, font=FONT_DISPLAY,
                 color=color, bold=True, line_spacing=1.02, caps=True)


def card_body(slide, x, y, w, h, text, color=MUTE, size=9):
    return label(slide, x, y, w, h, text, size=size, font=FONT_BODY,
                 color=color, line_spacing=1.2)


def icon_badge(slide, x, y, size, glyph, *, fill=SURF2, color=INK, line=INK):
    """The small rounded square that opens a card, holding a single glyph."""
    shp = box(slide, x, y, size, size, fill=fill, line=line,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.16)
    return shape_label(shp, glyph, size=13, font=FONT_DISPLAY, color=color)


def stat_block(slide, x, y, w, number, caption, *, number_color=INK,
               caption_color=MUTE, number_size=26, caption_size=8):
    label(slide, x, y, w, Inches(0.52), number, size=number_size,
          font=FONT_DISPLAY, color=number_color, bold=True, line_spacing=1.0)
    label(slide, x, y + Inches(0.52), w, Inches(0.5), caption, size=caption_size,
          font=FONT_MONO, color=caption_color, bold=True, line_spacing=1.16, caps=True)


def bullet(slide, x, y, w, h, head, tail, *, head_color=INK, body_color=MUTE,
           size=8.5, mark_color=ACC, line_spacing=1.2):
    """One ✓ row: a check glyph plus a bold lead-in and a plain continuation."""
    label(slide, x, y, Inches(0.17), h, "✓", size=size + 1.5, font=FONT_BODY,
          color=mark_color, bold=True, line_spacing=line_spacing)
    return label(slide, x + Inches(0.20), y, w - Inches(0.20), h,
                 [(head + " ", {"bold": True, "color": head_color}),
                  (tail, {"color": body_color})],
                 size=size, line_spacing=line_spacing)


def fact_column(slide, x, y, w, items, *, head_color=INK, body_color=MUTE,
                mark_color=ACC, size=8.5, step=0.30, height=0.80):
    """A vertical run of ✓ rows inside a card or a band."""
    for i, (head, tail) in enumerate(items):
        bullet(slide, x, y + Inches(step) * i, w, Inches(height), head, tail,
               head_color=head_color, body_color=body_color,
               mark_color=mark_color, size=size)


def strip(slide, y, heading, items, *, columns_count=3, height=None,
          item_step=None):
    """The wide sunken band that closes a slide: a heading above N columns of ✓ rows."""
    if height is None:
        height = Inches(1.30) if columns_count <= 2 else Inches(1.34)
    band = card(slide, MARGIN, y, CONTENT_W, height, fill=SURF2)

    label(slide, MARGIN + Inches(0.26), y + Inches(0.20),
          CONTENT_W - Inches(0.52), Inches(0.24), heading,
          size=8, font=FONT_MONO, color=MUTE, bold=True, caps=True)

    gap = Inches(0.34)
    col_w = int((CONTENT_W - Inches(0.52) - gap * (columns_count - 1)) / columns_count)
    rows = (len(items) + columns_count - 1) // columns_count
    step = item_step if item_step is not None else Inches(0.30)

    for i, (head, tail) in enumerate(items):
        row = i // columns_count
        col = i % columns_count
        bullet(slide, MARGIN + Inches(0.26) + (col_w + gap) * col,
               y + Inches(0.52) + step * row, col_w, Inches(0.34),
               head, tail, size=8.5)
    return band


def chip_row(slide, x, y, items, *, fill_for=None, max_w=None):
    """A wrapping row of pill chips."""
    limit = (SLIDE_W - MARGIN) if max_w is None else max_w
    height = Inches(0.28)
    for i, text in enumerate(items):
        width = Inches(0.30) + Inches(0.070) * len(text)
        if x + width > limit:
            break
        fill = SURF if fill_for is None else fill_for(i)
        pill(slide, x, y, width, height, text, fill=fill, color=INK, size=8)
        x += width + Inches(0.09)


def badge_row(slide, y, items, width=None):
    """Centred badge pills for the title slide."""
    widths = [Inches(0.30) + Inches(0.075) * len(t) for t in items]
    total = sum(widths) + Inches(0.11) * (len(items) - 1)
    start = (SLIDE_W - total) / 2 if width is None else width
    x = int(start)
    for i, text in enumerate(items):
        pill(slide, x, y, widths[i], Inches(0.32), text,
             fill=ACC2 if i == 0 else SURF, color=INK, size=8, caps=True)
        x += widths[i] + Inches(0.11)
# ── 01 · Титульный лист ──────────────────────────────────────────────────────


def slide_title(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "01/09")

    # logo: the favicon rebuilt from shapes — navy tile, gold ring, waveform bars
    logo_x, logo_y, logo_s = Inches(5.72), Inches(0.52), Inches(1.72)
    box(slide, logo_x, logo_y, logo_s, logo_s, fill=INK, line=INK,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.22)

    ring = box(slide, logo_x + Inches(0.20), logo_y + Inches(0.20),
               logo_s - Inches(0.40), logo_s - Inches(0.40),
               fill=None, line=ACC2, shape=MSO_SHAPE.OVAL, line_w=Pt(2))
    ring.shadow.inherit = False

    bar_w = Inches(0.135)
    bar_gap = Inches(0.052)
    bars_w = bar_w * 5 + bar_gap * 4
    bar_x = logo_x + (logo_s - bars_w) / 2
    centre_y = logo_y + logo_s / 2
    heights = [0.30, 0.50, 0.66, 0.52, 0.34]
    for i, h_in in enumerate(heights):
        h = Inches(h_in)
        bar = box(slide, bar_x + (bar_w + bar_gap) * i, centre_y - h / 2,
                  bar_w, h, fill=ACC2 if i == 2 else ACC_INK, line=None,
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        bar.shadow.inherit = False

    label(slide, MARGIN, Inches(2.52), CONTENT_W, Inches(1.1),
          [("WHISPER", {}), ("FLY", {"color": ACC})],
          size=52, font=FONT_DISPLAY, color=INK, bold=True,
          align=PP_ALIGN.CENTER, line_spacing=1.0)

    label(slide, Inches(3.4), Inches(3.62), SLIDE_W - Inches(6.8), Inches(0.8),
          "Бесплатная диктовка нажатием клавиши для macOS. Голос — в текст: "
          "мгновенно, в любом приложении и без единого тенге за подписку.",
          size=12, color=MUTE, align=PP_ALIGN.CENTER, line_spacing=1.3)

    meta_w = Inches(10.6)
    meta_x = (SLIDE_W - meta_w) / 2
    meta_gap = Inches(0.18)
    meta_cell = int((meta_w - meta_gap * 2) / 3)
    meta = [
        ("Проект", "WhisperFly — голосовой ввод для macOS"),
        ("Команда", "WhisperFly · независимый open-source проект"),
        ("Спикер", "Даниил Бурыкин, основатель"),
    ]
    for i, (term, value) in enumerate(meta):
        x = meta_x + (meta_cell + meta_gap) * i
        card(slide, x, Inches(4.60), meta_cell, Inches(0.92))
        label(slide, x + Inches(0.18), Inches(4.74), meta_cell - Inches(0.36),
              Inches(0.22), term, size=7.5, font=FONT_MONO, color=MUTE,
              bold=True, caps=True)
        label(slide, x + Inches(0.18), Inches(5.00), meta_cell - Inches(0.36),
              Inches(0.44), value, size=9.5, bold=True, line_spacing=1.24)

    badge_row(slide, Inches(5.92), [
        "Защита проекта · 2026", "macOS 14+", "11 языков", "Открытый код", "0 ₸",
    ])
    return slide


# ── 02 · Введение ────────────────────────────────────────────────────────────


def slide_intro(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "02/09")
    kicker(slide, "Введение")
    title(slide, [("Что такое ", {}), ("WhisperFly", {"color": ACC})])
    subtitle(slide,
             "WhisperFly — приложение для macOS, которое превращает речь в текст в любом "
             "приложении по одному нажатию ⌘⇧Space. Оно пишет и микрофон, и системный звук, "
             "распознаёт речь в облаке и вставляет готовый текст туда, где стоит курсор. "
             "Сегодня это работающий продукт версии 2.0 с открытым исходным кодом — "
             "и он полностью бесплатен.")

    left_w = Inches(7.0)
    card(slide, MARGIN, Inches(2.46), left_w, Inches(4.10), fill=INK)

    icon = box(slide, MARGIN + Inches(0.30), Inches(2.76), Inches(0.46),
               Inches(0.46), fill=ACC2, line=ACC2,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.16)
    shape_label(icon, "◎", size=13, font=FONT_DISPLAY, color=INK)

    card_head(slide, MARGIN + Inches(0.30), Inches(3.38),
              left_w - Inches(0.60), "Главная цель", color=ACC_INK, size=13)

    label(slide, MARGIN + Inches(0.30), Inches(3.86), left_w - Inches(0.60),
          Inches(1.60),
          "Убрать разрыв между скоростью мысли и скоростью набора текста — сделать "
          "голосовой ввод бесплатным, мгновенным и доступным каждому, кто работает за Mac.",
          size=12, color=ACC_INK, line_spacing=1.34)

    chips = ["Любое приложение macOS", "Без локальных моделей", "0 ₸"]
    chip_x = MARGIN + Inches(0.30)
    for i, text in enumerate(chips):
        width = Inches(0.30) + Inches(0.070) * len(text)
        pill(slide, chip_x, Inches(5.62), width, Inches(0.28), text,
             fill=ACC2 if i == 2 else None, color=INK if i == 2 else ACC_INK,
             line=ACC2 if i == 2 else ACC_INK, size=8)
        chip_x += width + Inches(0.09)

    right_x = MARGIN + left_w + Inches(0.20)
    right_w = CONTENT_W - left_w - Inches(0.20)
    label(slide, right_x, Inches(2.46), right_w, Inches(0.24),
          "Конкретные задачи", size=8, font=FONT_MONO, color=MUTE,
          bold=True, caps=True)

    tasks = [
        ("Сделать ввод мгновенным.",
         "Один хоткей, вставка текста в активное поле, отклик — секунды."),
        ("Охватить все сценарии звука.",
         "Микрофон, системный звук и готовые аудио- и видеофайлы."),
        ("Снять языковой барьер.",
         "Интерфейс на 11 языках, включая казахский; распознавание — более 100 языков."),
        ("Удержать нулевую цену.",
         "Бесплатное ядро и открытый код как основа продукта, а не как акция."),
    ]
    y = Inches(2.86)
    for i, (head, tail) in enumerate(tasks):
        num = box(slide, right_x, y, Inches(0.34), Inches(0.34), fill=ACC2,
                  line=INK, shape=MSO_SHAPE.OVAL)
        shape_label(num, str(i + 1), size=9, font=FONT_MONO, color=INK)
        label(slide, right_x + Inches(0.46), y, right_w - Inches(0.46),
              Inches(0.80),
              [(head + " ", {"bold": True, "color": INK}), (tail, {"color": MUTE})],
              size=9, line_spacing=1.24)
        y += Inches(0.92)
    return slide


# ── 03 · Уникальность ────────────────────────────────────────────────────────


COMPETITORS = [
    ("WhisperFly", "0 ₸", True, True, True, True),
    ("Superwhisper", "от $8,99/мес", False, True, False, False),
    ("MacWhisper", "от €59", False, True, False, False),
    ("Wispr Flow", "от $12/мес", False, False, False, False),
    ("Apple «Диктовка»", "0 ₸", False, False, False, False),
]


def slide_uniqueness(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "03/09")
    kicker(slide, "Уникальность")
    title(slide, [("Бесплатно и ", {}), ("открыто", {"color": ACC})])

    headers = ["Продукт", "Цена", "Системный звук", "Файлы", "Открытый код",
               "Казахский интерфейс"]
    fractions = [0.215, 0.155, 0.160, 0.120, 0.160, 0.190]
    widths = [int(CONTENT_W * f) for f in fractions]
    widths[-1] = int(CONTENT_W - sum(widths[:-1]))

    top = Inches(1.66)
    header_h = Inches(0.42)
    row_h = Inches(0.50)
    tint = RGBColor(0xF7, 0xE7, 0xC4)

    x = MARGIN
    for i, text in enumerate(headers):
        fill = ACC if i != 0 else INK
        cell = box(slide, x, top, widths[i], header_h, fill=fill, line=None,
                   shape=MSO_SHAPE.RECTANGLE)
        cell.shadow.inherit = False
        tf = cell.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.14)
        tf.margin_right = Inches(0.06)
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        run = para.add_run()
        run.text = text.upper()
        run.font.size = Pt(8)
        run.font.name = FONT_DISPLAY
        run.font.bold = True
        run.font.color.rgb = ACC_INK
        x += widths[i]

    for r, row in enumerate(COMPETITORS):
        y = int(top + header_h + row_h * r)
        ours = r == 0
        x = MARGIN
        values = [row[0], row[1]] + list(row[2:])
        for c, value in enumerate(values):
            fill = tint if ours else (SURF2 if r % 2 else SURF)
            cell = box(slide, x, y, widths[c], row_h, fill=fill, line=None,
                       shape=MSO_SHAPE.RECTANGLE)
            cell.shadow.inherit = False
            tf = cell.text_frame
            tf.word_wrap = False
            tf.margin_left = Inches(0.14)
            tf.margin_right = Inches(0.06)
            tf.margin_top = 0
            tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = tf.paragraphs[0]
            run = para.add_run()
            if isinstance(value, bool):
                run.text = "✓" if value else "✗"
                run.font.size = Pt(12)
                run.font.name = FONT_BODY
                run.font.bold = True
                run.font.color.rgb = ACC if value else MUTE
            else:
                run.text = value
                run.font.size = Pt(9.5)
                run.font.name = FONT_BODY
                run.font.bold = ours
                run.font.color.rgb = INK
            x += widths[c]

    frame = box(slide, MARGIN, top, CONTENT_W, header_h + row_h * len(COMPETITORS),
                fill=None, line=INK, shape=MSO_SHAPE.RECTANGLE)
    frame.shadow.inherit = False

    strip(slide, Inches(5.06), "В чём инновационность", [
        ("Облачный ASR вместо локальной модели.",
         "Приложение весит мегабайты, а не гигабайты."),
        ("Вставка через Accessibility API:",
         "текст появляется прямо там, где стоит курсор."),
        ("Системный звук через ScreenCaptureKit:",
         "созвоны и видео транскрибируются прямо с дорожки Mac."),
    ], columns_count=3, height=Inches(1.44), item_step=Inches(0.0))
    return slide
# ── 04 · КПД ─────────────────────────────────────────────────────────────────


EFFECT_STEPS = [
    ("10 000", "ПОЛЬЗОВАТЕЛЕЙ К КОНЦУ ГОДА"),
    ("40 мин", "ЭКОНОМИИ В ДЕНЬ НА ЧЕЛОВЕКА"),
    ("220", "РАБОЧИХ ДНЕЙ"),
    ("≈ 1,5 млн", "ЧЕЛОВЕКО-ЧАСОВ В ГОД"),
]


def slide_efficiency(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "04/09")
    kicker(slide, "КПД")
    title(slide, [("Выгоды ", {}), ("обществу и региону", {"color": ACC})])

    width, xs = columns(3)
    cards = [
        ("✓", "Обществу", SURF, INK, MUTE, [
            ("Доступность.", "Голосовой ввод возвращает работу с текстом людям "
                             "с нарушениями моторики и зрения."),
            ("Язык и культура.", "Казахский интерфейс — редкий случай для нишевого "
                                 "macOS-софта; распознавание — более 100 языков."),
            ("Бесплатный доступ.", "Студенты, учителя и малый бизнес получают "
                                   "инструмент без подписки."),
        ]),
        ("▲", "Экономике региона", ACC, ACC_INK, ACC_INK, [
            ("Экономия рабочего времени.", "Часы, которые специалисты тратят на набор, "
                                           "а не на работу."),
            ("Импортозамещение.", "Платные зарубежные подписки заменяются бесплатным "
                                  "локальным продуктом."),
            ("Налоговая база.", "Платная версия и корпоративные внедрения оформляются "
                                "как резидент Казахстана."),
        ]),
        ("◆", "Компетенциям", ACC2, INK, INK, [
            ("Открытый код как учебный материал.", "3 000+ строк Swift на GitHub — "
                                                   "площадка для студентов."),
            ("Экспортный потенциал.", "Продукт для глобального macOS-рынка, "
                                      "созданный в регионе."),
            ("Рабочие места.", "План найма: продуктовый маркетинг, продажи "
                               "и поддержка в 2026 году."),
        ]),
    ]

    top = Inches(1.62)
    height = Inches(3.14)
    for i, (glyph, head, fill, head_color, body_color, facts) in enumerate(cards):
        x = xs[i]
        card(slide, x, top, width, height, fill=fill)
        icon_badge(slide, x + Inches(0.26), top + Inches(0.26), Inches(0.42),
                   glyph,
                   fill=ACC_INK if fill == ACC else (INK if fill == ACC2 else SURF2),
                   color=ACC if fill == ACC else (ACC2 if fill == ACC2 else INK),
                   line=ACC_INK if fill == ACC else (INK if fill == ACC2 else INK))
        card_head(slide, x + Inches(0.26), top + Inches(0.86),
                  width - Inches(0.52), head, color=head_color, size=12)
        fact_column(slide, x + Inches(0.26), top + Inches(1.32),
                    width - Inches(0.52), facts,
                    head_color=head_color, body_color=body_color,
                    mark_color=head_color, size=8, step=0.54, height=0.80)

    band_y = Inches(4.96)
    band_h = Inches(1.90)
    card(slide, MARGIN, band_y, CONTENT_W, band_h, fill=SURF2)
    label(slide, MARGIN + Inches(0.26), band_y + Inches(0.18),
          CONTENT_W - Inches(0.52), Inches(0.22),
          "Расчёт эффекта · 12 месяцев", size=8, font=FONT_MONO, color=MUTE,
          bold=True, caps=True)

    op_zone = Inches(0.36)
    inner = CONTENT_W - Inches(0.52)
    item_w = int((inner - op_zone * 3) / 4)
    for i, (number, caption) in enumerate(EFFECT_STEPS):
        x = MARGIN + Inches(0.26) + (item_w + op_zone) * i
        last = i == len(EFFECT_STEPS) - 1
        label(slide, x, band_y + Inches(0.50), item_w, Inches(0.40), number,
              size=18, font=FONT_DISPLAY, color=ACC if last else INK, bold=True,
              line_spacing=1.0)
        label(slide, x, band_y + Inches(0.94), item_w, Inches(0.28), caption,
              size=7, font=FONT_MONO, color=MUTE, bold=True, line_spacing=1.2)
        if not last:
            label(slide, x + item_w, band_y + Inches(0.50), op_zone, Inches(0.40),
                  "×", size=15, font=FONT_DISPLAY, color=MUTE, align=PP_ALIGN.CENTER)

    label(slide, MARGIN + Inches(0.26), band_y + Inches(1.32),
          CONTENT_W - Inches(0.52), Inches(0.46),
          "Это эквивалент примерно 800 полных ставок в год — время, которое регион "
          "получает без единого тенге бюджетных затрат.",
          size=8.5, color=MUTE, line_spacing=1.3)
    return slide


# ── 05 · План реализации ─────────────────────────────────────────────────────


ROADMAP = [
    ("Q4 2025", "готово", ACC, [
        "Релиз 2.0: системный звук",
        "Транскрипция файлов",
        "История транскрипций",
        "11 языков интерфейса",
    ]),
    ("Q1 2026", "в работе", ACC, [
        "Пилот: 1 000 установок",
        "Метрики и обратная связь",
        "Подготовка Pro-версии",
        "Стабилизация macOS 26",
    ]),
    ("Q2 2026", None, SURF2, [
        "WhisperFly Pro",
        "Оплата и лицензии",
        "Свой словарь",
        "Найм: маркетинг",
    ]),
    ("Q3 2026", None, SURF2, [
        "Потоковая запись",
        "iOS-компаньон",
        "Версия для Windows",
        "Публичный API",
    ]),
    ("Q4 2026", None, SURF2, [
        "Админ-панель и SSO",
        "Общие словари",
        "Развёртывание on-prem",
        "Аудит безопасности",
    ]),
]


def slide_roadmap(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "05/09")
    kicker(slide, "План реализации")
    title(slide, [("От продукта к ", {}), ("платформе", {"color": ACC})])
    subtitle(slide, "Пять этапов на 15 месяцев. Первый уже закрыт, остальные ложатся "
                    "на работающий продукт, а не на прототип.", y=Inches(1.34))

    width, xs = columns(5)
    top = Inches(1.98)
    for i, (quarter, tag, bar_fill, items) in enumerate(ROADMAP):
        x = xs[i]
        bar = box(slide, x, top, width, Inches(0.15), fill=bar_fill, line=INK,
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        bar.shadow.inherit = False

        label(slide, x, top + Inches(0.26), width, Inches(0.30), quarter,
              size=11, font=FONT_DISPLAY, color=INK, bold=True, line_spacing=1.0)

        if tag:
            tag_w = Inches(0.24) + Inches(0.075) * len(tag)
            pill(slide, x, top + Inches(0.60), tag_w, Inches(0.24), tag,
                 fill=ACC if tag == "в работе" else SURF2, color=INK, size=7,
                 caps=True)

        for j, item in enumerate(items):
            y = top + Inches(0.98) + Inches(0.36) * j
            dot = box(slide, x + Inches(0.01), y + Inches(0.05), Inches(0.08),
                      Inches(0.08), fill=SURF, line=INK, shape=MSO_SHAPE.OVAL,
                      line_w=Pt(1.25))
            dot.shadow.inherit = False
            label(slide, x + Inches(0.20), y, width - Inches(0.20), Inches(0.34),
                  item, size=8, color=MUTE, line_spacing=1.24)

    strip(slide, Inches(4.78), "Почему план реалистичен", [
        ("Продукт уже работает.", "Ядро и интерфейс выпущены — предстоит развивать, "
                                  "а не начинать."),
        ("Нулевая инфраструктура.", "Движки на бесплатных тарифах, платить "
                                    "за серверы не нужно."),
        ("Один разработчик, весь стек.", "Дизайн, код, сборки и релизы — "
                                         "в одних руках."),
        ("Бюджет — только на вывод на рынок.", "Разработка ядра оплачена временем "
                                               "основателя."),
    ], columns_count=4, height=Inches(1.72), item_step=Inches(0.0))
    return slide


# ── 06 · Ожидаемые результаты ────────────────────────────────────────────────


def slide_results(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "06/09")
    kicker(slide, "Ожидаемые результаты")
    title(slide, [("Результаты ", {}), ("первого года", {"color": ACC})])

    width, xs = columns(4)
    top = Inches(1.62)
    height = Inches(2.20)
    stats = [
        ("10 000", "установок за 12 месяцев", INK, ACC_INK),
        ("500", "платящих пользователей Pro", ACC, ACC_INK),
        ("3", "корпоративных внедрения", ACC2, INK),
        ("1,5 млн", "сэкономленных человеко-часов в год", SURF, INK),
    ]
    for i, (number, caption, fill, fg) in enumerate(stats):
        card(slide, xs[i], top, width, height, fill=fill)
        stat_block(slide, xs[i] + Inches(0.28), top + Inches(0.50),
                   width - Inches(0.56), number, caption,
                   number_color=fg, caption_color=fg, number_size=24)

    strip(slide, Inches(4.06), "Как мы оцениваем успех проекта", [
        ("Охват.", "Не менее 10 000 активных установок за 12 месяцев и не менее "
                   "1 000 в первом квартале."),
        ("Конверсия.", "Переход Free → Pro не ниже 5 % — при сохранении полностью "
                       "бесплатного ядра."),
        ("Удержание.", "Не менее 35 % пользователей возвращаются к продукту "
                       "на 30-й день после установки."),
        ("Внедрения.", "Три организации региона — вуз, IT-компания или госструктура — "
                       "используют WhisperFly в ежедневной работе."),
    ], columns_count=2, height=Inches(2.32), item_step=Inches(0.66))
    return slide
# ── 07 · Финансовая часть ────────────────────────────────────────────────────


BUDGET = [
    ("Оплата труда основателя", "200 000 ₸ × 12 мес", "2 400 000 ₸"),
    ("Маркетинг и вывод продукта на рынок", "", "600 000 ₸"),
    ("Тестовое оборудование и устройства", "", "500 000 ₸"),
    ("Облачные API распознавания и AI-обработки", "", "360 000 ₸"),
    ("Юридические и бухгалтерские услуги", "", "300 000 ₸"),
    ("Сертификаты и подпись приложения Apple", "", "54 000 ₸"),
    ("Резерв на непредвиденные расходы", "", "286 000 ₸"),
]


def slide_finance(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "07/09")
    kicker(slide, "Финансовая часть")
    title(slide, [("Смета на ", {}), ("12 месяцев", {"color": ACC})])

    left_w = int(CONTENT_W * 0.60)
    col_a = int(left_w * 0.62)
    col_b = int(left_w - col_a)

    top = Inches(1.66)
    header_h = Inches(0.40)
    row_h = Inches(0.46)

    head_a = box(slide, MARGIN, top, col_a, header_h, fill=INK, line=None,
                 shape=MSO_SHAPE.RECTANGLE)
    head_a.shadow.inherit = False
    tf = head_a.text_frame
    tf.word_wrap = False
    tf.margin_left = Inches(0.16)
    tf.margin_right = Inches(0.08)
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = "Статья расходов"
    run.font.size = Pt(8)
    run.font.name = FONT_DISPLAY
    run.font.bold = True
    run.font.color.rgb = ACC_INK

    head_b = box(slide, MARGIN + col_a, top, col_b, header_h, fill=INK, line=None,
                 shape=MSO_SHAPE.RECTANGLE)
    head_b.shadow.inherit = False
    tf = head_b.text_frame
    tf.word_wrap = False
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.16)
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.RIGHT
    run = para.add_run()
    run.text = "Сумма"
    run.font.size = Pt(8)
    run.font.name = FONT_DISPLAY
    run.font.bold = True
    run.font.color.rgb = ACC_INK

    for r, (article, note, amount) in enumerate(BUDGET):
        y = int(top + header_h + row_h * r)
        fill = SURF2 if r % 2 else SURF

        cell = box(slide, MARGIN, y, col_a, row_h, fill=fill, line=None,
                   shape=MSO_SHAPE.RECTANGLE)
        cell.shadow.inherit = False
        label(slide, MARGIN + Inches(0.16), y + (Inches(0.06) if note else Inches(0.13)),
              col_a - Inches(0.24), Inches(0.20), article, size=8.5, line_spacing=1.2)
        if note:
            label(slide, MARGIN + Inches(0.16), y + Inches(0.25),
                  col_a - Inches(0.24), Inches(0.18), note, size=7,
                  font=FONT_MONO, color=MUTE, line_spacing=1.2)

        cell = box(slide, MARGIN + col_a, y, col_b, row_h, fill=fill, line=None,
                   shape=MSO_SHAPE.RECTANGLE)
        cell.shadow.inherit = False
        label(slide, MARGIN + col_a, y + Inches(0.13), col_b - Inches(0.16),
              Inches(0.22), amount, size=8.5, font=FONT_MONO, color=INK,
              bold=True, align=PP_ALIGN.RIGHT, line_spacing=1.2)

    total_y = int(top + header_h + row_h * len(BUDGET))
    row_total = box(slide, MARGIN, total_y, left_w, row_h, fill=INK, line=None,
                    shape=MSO_SHAPE.RECTANGLE)
    row_total.shadow.inherit = False
    label(slide, MARGIN + Inches(0.16), total_y + Inches(0.13),
          col_a - Inches(0.24), Inches(0.22), "Итого по смете", size=9,
          font=FONT_DISPLAY, color=ACC_INK, bold=True, line_spacing=1.2)
    label(slide, MARGIN + col_a, total_y + Inches(0.13), col_b - Inches(0.16),
          Inches(0.22), "4 500 000 ₸", size=9, font=FONT_DISPLAY, color=ACC_INK,
          bold=True, align=PP_ALIGN.RIGHT, line_spacing=1.2)

    frame = box(slide, MARGIN, top, left_w, header_h + row_h * (len(BUDGET) + 1),
                fill=None, line=INK, shape=MSO_SHAPE.RECTANGLE)
    frame.shadow.inherit = False

    right_x = MARGIN + left_w + Inches(0.20)
    right_w = CONTENT_W - left_w - Inches(0.20)

    card(slide, right_x, top, right_w, Inches(2.10), fill=ACC)
    icon = box(slide, right_x + Inches(0.28), top + Inches(0.26), Inches(0.42),
               Inches(0.42), fill=ACC_INK, line=ACC_INK,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.16)
    shape_label(icon, "₸", size=13, font=FONT_DISPLAY, color=ACC)
    label(slide, right_x + Inches(0.28), top + Inches(0.86),
          right_w - Inches(0.56), Inches(0.22), "Запрашиваемая сумма", size=7.5,
          font=FONT_MONO, color=ACC_INK, bold=True, caps=True)
    label(slide, right_x + Inches(0.28), top + Inches(1.12),
          right_w - Inches(0.56), Inches(0.48), "4 500 000 ₸", size=23,
          font=FONT_DISPLAY, color=ACC_INK, bold=True, line_spacing=1.0)
    label(slide, right_x + Inches(0.28), top + Inches(1.62),
          right_w - Inches(0.56), Inches(0.40),
          "Четыре миллиона пятьсот тысяч тенге на 12 месяцев — вся смета целиком.",
          size=8.5, color=ACC_INK, line_spacing=1.26)

    fund_y = top + Inches(2.34)
    fund_h = Inches(2.10)
    card(slide, right_x, fund_y, right_w, fund_h, fill=SURF2)
    fact_column(slide, right_x + Inches(0.28), fund_y + Inches(0.26),
                right_w - Inches(0.56), [
                    ("Собственный вклад.", "Выпущенный продукт, инфраструктура и время "
                                           "основателя — эквивалент 2 000 000 ₸."),
                    ("Только вывод на рынок.", "Разработка ядра в смете отсутствует: "
                                               "средства идут на масштабирование."),
                    ("Возвратность.", "Выход на выручку — Q3 2026, самоокупаемость "
                                      "к концу года."),
                ], size=8, step=0.62, height=0.80)
    return slide


# ── 08 · Команда ─────────────────────────────────────────────────────────────


CREW = [
    ("ДБ", "Даниил Бурыкин", "Основатель · разработка · продукт",
     "Полный цикл: архитектура на Swift, интерфейс, конвейер распознавания, "
     "сборки, релизы и дистрибуция.", SURF, ACC, ACC_INK, INK, MUTE),
    ("<>", "Открытое сообщество", "Контрибьюторы · тестирование",
     "Репозиторий на GitHub открыт: вклад в код, переводы интерфейса "
     "и сообщения об ошибках приветствуются.", ACC2, INK, ACC2, INK, INK),
    ("HK", "hukopo", "Автор основы проекта",
     "Автор qwenwishper — проекта, из которого вырос WhisperFly; его код стал "
     "отправной точкой разработки.", SURF, ACC, ACC_INK, INK, MUTE),
]


def slide_team(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "08/09")
    kicker(slide, "Команда")
    title(slide, [("Весь стек ", {}), ("в одних руках", {"color": ACC})])

    width, xs = columns(3)
    top = Inches(1.62)
    height = Inches(3.06)
    for i, (avatar, name, role, desc, fill, av_fill, av_fg, name_fg, body_fg) in enumerate(CREW):
        x = xs[i]
        card(slide, x, top, width, height, fill=fill)

        size = Inches(0.86)
        circle = box(slide, int(x + (width - size) / 2), top + Inches(0.30),
                     size, size, fill=av_fill, line=INK, shape=MSO_SHAPE.OVAL)
        shape_label(circle, avatar, size=15 if avatar != "<>" else 12,
                    font=FONT_DISPLAY, color=av_fg)

        label(slide, x, top + Inches(1.34), width, Inches(0.34), name,
              size=12, font=FONT_DISPLAY, color=name_fg, bold=True,
              align=PP_ALIGN.CENTER, line_spacing=1.05)
        label(slide, x + Inches(0.18), top + Inches(1.72), width - Inches(0.36),
              Inches(0.36), role, size=7.5, font=FONT_MONO, color=body_fg,
              bold=True, align=PP_ALIGN.CENTER, line_spacing=1.22, caps=True)
        label(slide, x + Inches(0.26), top + Inches(2.14), width - Inches(0.52),
              Inches(0.80), desc, size=8.5, color=body_fg,
              align=PP_ALIGN.CENTER, line_spacing=1.26)

    strip(slide, Inches(4.86), "Роль каждого в реализации", [
        ("Основатель", "— ведёт разработку, отвечает за продукт, релизы и дистрибуцию, "
                       "представляет проект."),
        ("Сообщество", "— тестирование на разных конфигурациях, локализация интерфейса, "
                       "обратная связь."),
        ("Планируемые роли", "— продуктовый маркетинг и продажи: найм запланирован "
                             "на Q2 2026 года."),
    ], columns_count=3, height=Inches(1.66), item_step=Inches(0.0))
    return slide


# ── 09 · Заключение ──────────────────────────────────────────────────────────


def slide_closing(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "09/09")
    kicker(slide, "Заключение")

    title(slide, [("Голос быстрее клавиатуры. Мы сделали это ", {}),
                  ("бесплатным", {"color": ACC}), (".", {})],
          y=Inches(0.80), size=25, width=CONTENT_W - Inches(0.4))

    label(slide, MARGIN, Inches(1.86), CONTENT_W - Inches(0.6), Inches(0.70),
          "WhisperFly уже работает и уже бесплатен: голосовой ввод в любом приложении "
          "macOS, собственный код, открытый исходный код и ноль затрат для пользователя. "
          "Осталось превратить готовый продукт в устойчивую платформу — и вернуть "
          "региону миллионы сэкономленных часов.",
          size=10, color=MUTE, line_spacing=1.34)

    width, xs = columns(3)
    top = Inches(2.78)
    height = Inches(2.10)
    closing = [
        ("★", "Уже работает",
         "Версия 2.0 выпущена: системный звук, транскрипция файлов, история "
         "и 11 языков интерфейса."),
        ("◆", "Бесплатно и открыто",
         "Ни подписки, ни локальных моделей, ни закрытых исходников — продукт "
         "можно проверить и собрать самому."),
        ("▲", "Эффект для региона",
         "1,5 млн человеко-часов в год, бесплатный доступ для образования "
         "и малого бизнеса, рабочие места в 2026 году."),
    ]
    for i, (glyph, head, body) in enumerate(closing):
        x = xs[i]
        card(slide, x, top, width, height)
        icon_badge(slide, x + Inches(0.28), top + Inches(0.28), Inches(0.44),
                   glyph, fill=SURF2, color=INK, line=INK)
        card_head(slide, x + Inches(0.28), top + Inches(0.92),
                  width - Inches(0.56), head, size=11.5)
        card_body(slide, x + Inches(0.28), top + Inches(1.34),
                  width - Inches(0.56), Inches(0.80), body, size=8.5)

    cta_y = Inches(5.24)
    ctas = [
        ("Сайт презентации", ACC, ACC_INK),
        ("Скачать бесплатно", SURF, INK),
        ("GitHub", SURF, INK),
    ]
    cta_x = MARGIN
    for text, fill, fg in ctas:
        w = Inches(0.54) + Inches(0.105) * len(text)
        pill(slide, cta_x, cta_y, w, Inches(0.50), text, fill=fill, color=fg,
             size=10, font=FONT_DISPLAY, caps=True)
        cta_x += w + Inches(0.16)

    label(slide, MARGIN, Inches(6.02), CONTENT_W, Inches(0.28),
          "Открытый исходный код · macOS 14+ · 11 языков интерфейса",
          size=8, font=FONT_MONO, color=MUTE, bold=True, caps=True)
    return slide


# ── Entry point ──────────────────────────────────────────────────────────────

BUILDERS = [
    slide_title,
    slide_intro,
    slide_uniqueness,
    slide_efficiency,
    slide_roadmap,
    slide_results,
    slide_finance,
    slide_team,
    slide_closing,
]

OUTPUT = "WhisperFly-Pitch-RU.pptx"


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    for build in BUILDERS:
        build(prs)

    prs.save(OUTPUT)
    print(f"wrote {OUTPUT} — {len(prs.slides._sldIdLst)} slides")
    print(f"accent oklch(0.44 0.16 {HUE}) -> #{ACC}")


if __name__ == "__main__":
    main()
