#!/usr/bin/env python3
"""Build WhisperFly-Pitch-RU.pptx — the editable PowerPoint version of the deck.

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


BG = RGBColor(0xF3, 0xF1, 0xEA)      # --tp-bg
SURF = RGBColor(0xFD, 0xFC, 0xF8)    # --tp-surf
SURF2 = RGBColor(0xE5, 0xE1, 0xD4)   # --tp-surf-2
INK = RGBColor(0x1C, 0x24, 0x47)     # --tp-ink
MUTE = RGBColor(0x5A, 0x61, 0x80)    # --tp-mute
ACC_INK = RGBColor(0xFD, 0xFC, 0xF8)  # --tp-acc-ink
ACC = oklch(0.44, 0.16, HUE)
ACC2 = oklch(0.78, 0.14, HUE + 155)

# ── Type ─────────────────────────────────────────────────────────────────────

FONT_DISPLAY = "Unbounded"
FONT_BODY = "Rubik"
FONT_MONO = "JetBrains Mono"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.75)
BORDER = Pt(2.25)

# ── Shape / text helpers ─────────────────────────────────────────────────────


def blank(prs):
    """A slide with no placeholders."""
    return prs.slides.add_slide(prs.slide_layouts[6])


def box(slide, x, y, w, h, fill=None, line=None, radius=None,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, line_w=BORDER):
    """Add a shape. `radius` (0..0.5) tunes the corner rounding on rounded rects."""
    shp = slide.shapes.add_shape(shape, x, y, w, h)
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
    tb = slide.shapes.add_textbox(x, y, w, h)
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


def kicker(slide, text, x=None, y=Inches(0.46), gold=False):
    """The small uppercase chip that opens most slides."""
    x = MARGIN if x is None else x
    width = Inches(0.30) + Inches(0.105) * len(text)

    chip = box(slide, x, y, width, Inches(0.36),
               fill=ACC2 if gold else ACC, line=INK, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
               radius=0.5)
    tf = chip.text_frame
    tf.word_wrap = False
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = text.upper()
    run.font.size = Pt(9)
    run.font.name = FONT_MONO
    run.font.bold = True
    run.font.color.rgb = INK if gold else ACC_INK
    return chip


def title(slide, runs, y=Inches(0.92), size=30, width=None):
    """Slide headline. `runs` accepts the same shapes as `label`."""
    width = width or (SLIDE_W - MARGIN * 2)
    return label(slide, MARGIN, y, width, Inches(1.0), runs,
                 size=size, font=FONT_DISPLAY, color=INK, bold=True,
                 line_spacing=0.98)


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

CONTENT_W = SLIDE_W - MARGIN * 2
BODY_TOP = Inches(1.98)
GAP = Inches(0.22)


def columns(count, gap=GAP, total=None, margin=None):
    """Even columns across the content width. Returns (width, [x, ...])."""
    total = CONTENT_W if total is None else total
    margin = MARGIN if margin is None else margin
    width = int((total - gap * (count - 1)) / count)
    xs = [int(margin + i * (width + gap)) for i in range(count)]
    return width, xs


def card(slide, x, y, w, h, *, fill=SURF, line=INK, radius=0.11):
    return box(slide, x, y, w, h, fill=fill, line=line,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)


def card_head(slide, x, y, w, text, color=INK, size=14):
    return label(slide, x, y, w, Inches(0.5), text, size=size, font=FONT_DISPLAY,
                 color=color, bold=True, line_spacing=1.02, caps=True)


def card_body(slide, x, y, w, h, text, color=MUTE, size=10.5):
    return label(slide, x, y, w, h, text, size=size, font=FONT_BODY,
                 color=color, line_spacing=1.22)


def stat_block(slide, x, y, w, number, caption, *, number_color=INK,
               caption_color=MUTE, number_size=30):
    label(slide, x, y, w, Inches(0.62), number, size=number_size,
          font=FONT_DISPLAY, color=number_color, bold=True, line_spacing=1.0)
    label(slide, x, y + Inches(0.60), w, Inches(0.5), caption, size=9,
          font=FONT_MONO, color=caption_color, bold=True, line_spacing=1.12, caps=True)


def chip_row(slide, y, items, *, fill_for=None):
    """A wrapping row of pill chips. `fill_for(i)` may return an override fill."""
    x = MARGIN
    height = Inches(0.34)
    for i, text in enumerate(items):
        width = Inches(0.34) + Inches(0.078) * len(text)
        if x + width > SLIDE_W - MARGIN:
            break
        fill = SURF if fill_for is None else fill_for(i)
        shape = box(slide, x, y, width, height, fill=fill, line=INK,
                    shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        tf = shape.text_frame
        tf.word_wrap = False
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = text
        run.font.size = Pt(9)
        run.font.name = FONT_MONO
        run.font.bold = True
        run.font.color.rgb = INK
        x += width + Inches(0.11)


def badge_row(slide, y, items):
    """Centred badge pills for the title slide."""
    widths = [Inches(0.34) + Inches(0.082) * len(t) for t in items]
    total = sum(widths) + Inches(0.12) * (len(items) - 1)
    x = int((SLIDE_W - total) / 2)
    for i, text in enumerate(items):
        shape = box(slide, x, y, widths[i], Inches(0.36),
                    fill=ACC2 if i == 0 else SURF, line=INK,
                    shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        tf = shape.text_frame
        tf.word_wrap = False
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = text.upper()
        run.font.size = Pt(9)
        run.font.name = FONT_MONO
        run.font.bold = True
        run.font.color.rgb = INK
        x += widths[i] + Inches(0.12)


# ── 01 · Титул ───────────────────────────────────────────────────────────────


def slide_title(prs):
    slide = blank(prs)
    spine(slide)

    crest_size = Inches(1.62)
    crest_x = int((SLIDE_W - crest_size) / 2)
    crest = box(slide, crest_x, Inches(0.62), crest_size, crest_size,
                fill=ACC, line=INK, shape=MSO_SHAPE.OVAL)

    tf = crest.text_frame
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = "WF"
    run.font.size = Pt(30)
    run.font.name = FONT_DISPLAY
    run.font.bold = True
    run.font.color.rgb = ACC_INK

    label(slide, MARGIN, Inches(2.62), CONTENT_W, Inches(1.3),
          [("WHISPER", {}), ("FLY", {"color": ACC})],
          size=62, font=FONT_DISPLAY, color=INK, bold=True,
          align=PP_ALIGN.CENTER, line_spacing=1.0)

    label(slide, Inches(2.4), Inches(4.32), SLIDE_W - Inches(4.8), Inches(1.0),
          "Диктовка нажатием клавиши для macOS. Голос — в текст: мгновенно, "
          "в любом приложении и без единого тенге за подписку.",
          size=13.5, color=MUTE, align=PP_ALIGN.CENTER, line_spacing=1.3)

    badge_row(slide, Inches(5.62), [
        "Конкурс стартапов · 2026", "macOS 14+", "10 языков",
        "Открытый код", "0 ₸",
    ])
    return slide


# ── 02 · Проблема ────────────────────────────────────────────────────────────


def slide_problem(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Проблема")
    title(slide, [
        ("Печатать — ", {}),
        ("в 4 раза медленнее", {"color": ACC}),
        (", чем говорить", {}),
    ])
    label(slide, MARGIN, Inches(1.66), CONTENT_W - Inches(1.2), Inches(0.7),
          "Мы думаем со скоростью речи, но вводим текст со скоростью клавиатуры. "
          "Разрыв между мыслью и набором съедает часы рабочего времени каждый день — "
          "и почти все существующие решения просят за это ежемесячную подписку.",
          size=11.5, color=MUTE, line_spacing=1.3)

    width, xs = columns(4)
    cards = [
        ("40", "слов в минуту", "Средняя скорость набора на клавиатуре у офисного сотрудника.", INK, SURF),
        ("150", "слов в минуту", "Средняя скорость речи. Почти в четыре раза быстрее набора.", ACC, ACC_INK),
        ("≈ 3 ч", "в день", "Уходит у человека, работа которого — текст: письма, заметки, документы, код.", ACC2, INK),
        ("$15–30", "в месяц", "Стоит подписка на существующие приложения для диктовки.", INK, SURF),
    ]
    for i, (number, caption, body, fill, fg) in enumerate(cards):
        top = int(BODY_TOP + Inches(0.35))
        card(slide, xs[i], top, width, Inches(2.65), fill=fill)
        stat_block(slide, xs[i] + Inches(0.28), top + Inches(0.36),
                   width - Inches(0.56), number, caption,
                   number_color=fg,
                   caption_color=fg)
        card_body(slide, xs[i] + Inches(0.28), top + Inches(1.62),
                  width - Inches(0.56), Inches(0.9), body, color=fg)
    return slide


# ── 03 · Решение ─────────────────────────────────────────────────────────────


def slide_solution(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Решение")
    title(slide, [
        ("Нажал. ", {}),
        ("Сказал.", {"color": ACC}),
        (" Готово.", {}),
    ])
    label(slide, MARGIN, Inches(1.66), CONTENT_W - Inches(1.2), Inches(0.7),
          "WhisperFly превращает речь в текст в любой программе macOS — по одному "
          "нажатию ⌘⇧Space. Без подписки, без загрузки моделей и без требований к железу.",
          size=11.5, color=MUTE, line_spacing=1.3)

    left_w = Inches(6.9)
    right_x = int(MARGIN + left_w + GAP)
    right_w = int(CONTENT_W - left_w - GAP)

    points = [
        ("Одной клавишей.", "Глобальное сочетание ⌘⇧Space запускает и останавливает запись из любого приложения."),
        ("Текст там, где курсор.", "Вставка через Accessibility API, при отказе — надёжный откат на буфер обмена."),
        ("Ноль локальных моделей.", "Распознавание идёт в облаке: не нужен GPU и гигабайты загрузок."),
        ("Бесплатно.", "Оба движка распознавания работают на бесплатных тарифах провайдеров."),
    ]
    y = int(BODY_TOP + Inches(0.30))
    for i, (head, body) in enumerate(points):
        num = box(slide, MARGIN, y, Inches(0.46), Inches(0.46), fill=ACC2, line=INK,
                  shape=MSO_SHAPE.OVAL)
        tf = num.text_frame
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = str(i + 1)
        run.font.size = Pt(11)
        run.font.name = FONT_MONO
        run.font.bold = True
        run.font.color.rgb = INK

        label(slide, MARGIN + Inches(0.70), y - Inches(0.04), left_w - Inches(0.70), Inches(0.9),
              [(head + " ", {"bold": True, "color": INK}), (body, {"color": MUTE})],
              size=10.5, line_spacing=1.24)
        y += Inches(1.16)

    card(slide, right_x, BODY_TOP, right_w, Inches(4.30), fill=INK)
    label(slide, right_x + Inches(0.36), BODY_TOP + Inches(0.55),
          right_w - Inches(0.72), Inches(1.0), "⌘⇧Space",
          size=34, font=FONT_DISPLAY, color=ACC2, bold=True, line_spacing=1.0)
    label(slide, right_x + Inches(0.36), BODY_TOP + Inches(1.55),
          right_w - Inches(0.72), Inches(0.4), "одно сочетание на всё",
          size=9, font=FONT_MONO, color=SURF2, bold=True, caps=True)
    label(slide, right_x + Inches(0.36), BODY_TOP + Inches(2.20),
          right_w - Inches(0.72), Inches(1.6),
          "Нажать — сказать — отпустить. Результат появляется в поле ввода "
          "за считанные секунды.",
          size=10.5, color=SURF2, line_spacing=1.28)
    return slide


# ── 04 · Как это работает ────────────────────────────────────────────────────


def slide_pipeline(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Как это работает")
    title(slide, [
        ("Пять шагов за ", {}),
        ("секунды", {"color": ACC}),
    ])
    label(slide, MARGIN, Inches(1.66), CONTENT_W - Inches(1.2), Inches(0.5),
          "От нажатия клавиши до готового текста в поле ввода — один короткий конвейер.",
          size=11.5, color=MUTE, line_spacing=1.3)

    width, xs = columns(5, gap=Inches(0.0))
    steps = [
        ("01", "Запись", "AVFoundation пишет микрофон. ScreenCaptureKit — всё, что звучит на Mac."),
        ("02", "Конвертация", "Аудио приводится к WAV 16 кГц — формату, который понимает движок."),
        ("03", "Распознавание", "Groq Whisper Large V3 или Google Gemini 2.5 Flash. Более 100 языков."),
        ("04", "Обработка", "Необязательная AI-переформулировка: грамматика, пунктуация, перевод."),
        ("05", "Вставка", "Готовый текст появляется прямо в активном поле ввода."),
    ]
    top = int(BODY_TOP + Inches(0.45))
    height = Inches(3.35)
    for i, (num, head, body) in enumerate(steps):
        fill = SURF2 if i % 2 else SURF
        card(slide, xs[i], top, width, height, fill=fill,
             radius=0.06 if 0 < i < 4 else 0.16)
        label(slide, xs[i] + Inches(0.26), top + Inches(0.30), width - Inches(0.52),
              Inches(0.3), num, size=10, font=FONT_MONO, color=ACC, bold=True)
        label(slide, xs[i] + Inches(0.26), top + Inches(0.72), width - Inches(0.52),
              Inches(0.7), head, size=12.5, font=FONT_DISPLAY, color=INK,
              bold=True, line_spacing=1.04, caps=True)
        card_body(slide, xs[i] + Inches(0.26), top + Inches(1.62),
                  width - Inches(0.52), Inches(1.6), body)

    label(slide, MARGIN, Inches(6.62), CONTENT_W, Inches(0.4),
          "Временные аудиофайлы удаляются сразу после распознавания. Результат попадает "
          "в историю, его можно скопировать или озвучить системным голосом.",
          size=9.5, color=MUTE, line_spacing=1.3)
    return slide


# ── 05 · Возможности ─────────────────────────────────────────────────────────


def slide_features(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Возможности")
    title(slide, [
        ("Всё, что нужно для ", {}),
        ("голосового ввода", {"color": ACC}),
    ])

    tiles = [
        ("Два источника звука", "Микрофон или системный звук — транскрибируйте созвоны и видео прямо с дорожки Mac."),
        ("Два бесплатных движка", "Groq Whisper Large V3 и Google Gemini 2.5 Flash через OpenRouter — оба на бесплатном тарифе."),
        ("Транскрипция файлов", "MP3, M4A, WAV, FLAC, MP4, MOV — результат сразу уходит в буфер обмена."),
        ("AI-переформулировка", "Чистка грамматики, расстановка пунктуации или перевод на английский — одним переключателем."),
        ("История транскрипций", "Последние 100 результатов: просмотр, копирование и повторное открытие в отдельном окне."),
        ("10 языков интерфейса", "Русский, английский, немецкий, французский, испанский, японский, китайский, корейский, итальянский, хинди."),
    ]
    width, xs = columns(3)
    row_top = [int(BODY_TOP + Inches(0.10)), int(BODY_TOP + Inches(2.35))]
    row_h = Inches(2.05)

    for i, (head, body) in enumerate(tiles):
        x = xs[i % 3]
        y = row_top[i // 3]
        card(slide, x, y, width, row_h)
        card_head(slide, x + Inches(0.28), y + Inches(0.26), width - Inches(0.56), head, size=12.5)
        card_body(slide, x + Inches(0.28), y + Inches(0.98), width - Inches(0.56),
                  Inches(0.95), body, size=10)
    return slide


# ── 06 · Технологии ──────────────────────────────────────────────────────────


def slide_stack(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Технологии")
    title(slide, [
        ("Нативный стек, ", {}),
        ("без компромиссов", {"color": ACC}),
    ])

    chip_row(slide, Inches(1.72), [
        "Swift 6", "SwiftUI", "AppKit", "ScreenCaptureKit", "AVFoundation",
        "Carbon", "Groq API", "OpenRouter",
    ], fill_for=lambda i: ACC if i == 0 else (ACC2 if i >= 6 else SURF))

    left_w = Inches(6.9)
    right_x = int(MARGIN + left_w + GAP)
    right_w = int(CONTENT_W - left_w - GAP)
    top = Inches(2.42)

    card(slide, MARGIN, top, left_w, Inches(3.95))
    card_head(slide, MARGIN + Inches(0.30), top + Inches(0.28),
              left_w - Inches(0.60), "Пять модулей", size=13)
    layers = [
        ("App", "Конвейер, плавающая панель, HUD-окна"),
        ("Services", "Запись, конвертация, движки, хоткей, вставка"),
        ("Core", "Протоколы и статусы конвейера"),
        ("Models", "Настройки и хранилище истории"),
        ("Views", "SwiftUI-интерфейс и строка меню"),
    ]
    y = top + Inches(0.95)
    for name, desc in layers:
        label(slide, MARGIN + Inches(0.30), y, Inches(1.55), Inches(0.34), name,
              size=11, font=FONT_DISPLAY, color=INK, bold=True, caps=True)
        label(slide, MARGIN + Inches(1.95), y + Inches(0.02), left_w - Inches(2.30),
              Inches(0.4), desc, size=10, color=MUTE, line_spacing=1.2)
        y += Inches(0.57)

    stat_w, stat_xs = columns(2, total=right_w, margin=right_x)
    stat_h = Inches(1.90)
    stat_cards = [
        ("3 000+", "строк Swift", INK, ACC_INK),
        ("10", "локализаций", ACC, ACC_INK),
        ("3", "сборки DMG", ACC2, INK),
        ("0 ГБ", "загрузок моделей", INK, ACC_INK),
    ]
    for i, (number, caption, fill, fg) in enumerate(stat_cards):
        x = stat_xs[i % 2]
        y = int(top + (stat_h + GAP) * (i // 2))
        card(slide, x, y, stat_w, stat_h, fill=fill)
        stat_block(slide, x + Inches(0.26), y + Inches(0.42),
                   stat_w - Inches(0.52), number, caption,
                   number_color=fg, caption_color=fg, number_size=26)
    return slide
# ── 07 · Рынок ───────────────────────────────────────────────────────────────


def slide_market(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Рынок")
    title(slide, [
        ("Растущий рынок ", {}),
        ("голосовых интерфейсов", {"color": ACC}),
    ])

    width, xs = columns(3)
    top = int(BODY_TOP + Inches(0.10))
    cards = [
        ("$5,2 млрд", "TAM · весь рынок", "Мировой рынок распознавания речи и ПО для диктовки.", INK, ACC_INK, SURF2),
        ("$780 млн", "SAM · наш сегмент", "Пользователи macOS, чья работа — постоянно вводить текст.", ACC, ACC_INK, ACC_INK),
        ("$12 млн", "SOM · цель за 3 года", "Русскоязычный сегмент и малый бизнес — рынок, который мы можем занять.", ACC2, INK, INK),
    ]
    for i, (number, caption, body, fill, fg, body_fg) in enumerate(cards):
        card(slide, xs[i], top, width, Inches(2.55), fill=fill)
        stat_block(slide, xs[i] + Inches(0.30), top + Inches(0.38),
                   width - Inches(0.60), number, caption,
                   number_color=fg, caption_color=fg, number_size=25)
        card_body(slide, xs[i] + Inches(0.30), top + Inches(1.58),
                  width - Inches(0.60), Inches(0.85), body, color=body_fg)

    chip_row(slide, Inches(5.18), [
        "≈ 20 % в год — рост рынка",
        "100+ млн — активных Mac",
        "250 млн — носителей русского",
        "0 ₸ — порог входа",
    ])

    label(slide, MARGIN, Inches(6.28), CONTENT_W, Inches(0.6),
          "Оценка построена на публичных данных о рынке ASR и голосового ПО; SOM — "
          "целевая доля русскоязычного направления за три года.",
          size=9.5, color=MUTE, line_spacing=1.3)
    return slide


# ── 08 · Бизнес-модель ───────────────────────────────────────────────────────


def slide_business(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Бизнес-модель")
    title(slide, [
        ("Бесплатное ядро. ", {}),
        ("Платная скорость.", {"color": ACC}),
    ])

    width, xs = columns(3)
    top = int(BODY_TOP + Inches(0.10))
    plans = [
        ("Free · 0 ₸", "Всё ядро без ограничений: два движка, транскрипция файлов, "
                       "история, 10 языков. Исходный код открыт.", SURF, INK, MUTE),
        ("Pro · 3 990 ₸/мес", "Приоритетные модели, безлимитный объём, пользовательский "
                            "словарь, экспорт и ранний доступ к новым функциям.", ACC, ACC_INK, ACC_INK),
        ("Team · от 2 990 ₸", "За сотрудника в месяц: админ-панель, единый вход, общие "
                            "словари и развёртывание внутри компании.", SURF, INK, MUTE),
    ]
    for i, (head, body, fill, head_fg, body_fg) in enumerate(plans):
        card(slide, xs[i], top, width, Inches(2.70), fill=fill)
        card_head(slide, xs[i] + Inches(0.30), top + Inches(0.40),
                  width - Inches(0.60), head, color=head_fg, size=13.5)
        card_body(slide, xs[i] + Inches(0.30), top + Inches(1.28),
                  width - Inches(0.60), Inches(1.20), body, color=body_fg)

    chip_row(slide, Inches(5.32), [
        "Маржа > 80 %",
        "Себестоимость распознавания — центы в месяц",
        "Привлечение через открытый код",
        "Цель LTV/CAC > 3×",
    ], fill_for=lambda i: ACC if i == 0 else SURF)

    label(slide, MARGIN, Inches(6.34), CONTENT_W, Inches(0.6),
          "Монетизация — план на 2026 год. Бесплатное ядро и открытый исходный код "
          "остаются основой продукта: они и есть канал привлечения.",
          size=9.5, color=MUTE, line_spacing=1.3)
    return slide


# ── 09 · Конкуренты ──────────────────────────────────────────────────────────


def _mix(a, b, t):
    return RGBColor(*(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3)))


def slide_competition(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Конкурентная среда")
    title(slide, [
        ("Бесплатно — и с ", {}),
        ("открытым кодом", {"color": ACC}),
    ])

    headers = ["Продукт", "Цена", "Системный звук", "Файлы", "Открытый код"]
    rows = [
        ("WhisperFly", "0 ₸", True, True, True),
        ("Superwhisper", "от $8,99/мес", False, True, False),
        ("MacWhisper", "от €59", False, True, False),
        ("Wispr Flow", "от $12/мес", False, False, False),
        ("Apple «Диктовка»", "0 ₸", False, False, False),
    ]

    total = CONTENT_W
    fractions = [0.30, 0.19, 0.18, 0.15, 0.18]
    widths = [int(total * f) for f in fractions]
    widths[-1] = int(total - sum(widths[:-1]))

    x0 = MARGIN
    y = int(BODY_TOP + Inches(0.18))
    header_h = Inches(0.60)
    row_h = Inches(0.74)
    tint = _mix(ACC, SURF, 0.86)

    # header
    x = x0
    for i, text in enumerate(headers):
        fill = ACC if i == 4 else INK
        cell = box(slide, x, y, widths[i], header_h, fill=fill, line=None,
                   shape=MSO_SHAPE.RECTANGLE)
        cell.shadow.inherit = False
        tf = cell.text_frame
        tf.margin_left = Inches(0.16)
        tf.margin_right = Inches(0.08)
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        run = para.add_run()
        run.text = text.upper()
        run.font.size = Pt(9.5)
        run.font.name = FONT_DISPLAY
        run.font.bold = True
        run.font.color.rgb = ACC_INK
        x += widths[i]

    # body rows
    for r, (name, price, sys_audio, files, open_src) in enumerate(rows):
        y = int(y + (header_h if r == 0 else row_h))
        ours = r == 0
        x = x0
        values = [name, price, sys_audio, files, open_src]

        for c, value in enumerate(values):
            fill = tint if ours and c == 4 else (tint if ours else (
                SURF2 if (r - 1) % 2 else SURF))
            if ours:
                fill = tint
            cell = box(slide, x, y, widths[c], row_h, fill=fill, line=None,
                       shape=MSO_SHAPE.RECTANGLE)
            cell.shadow.inherit = False
            tf = cell.text_frame
            tf.margin_left = Inches(0.16)
            tf.margin_right = Inches(0.08)
            tf.margin_top = 0
            tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = tf.paragraphs[0]

            run = para.add_run()
            if isinstance(value, bool):
                run.text = "✓" if value else "✗"
                run.font.size = Pt(15)
                run.font.name = FONT_BODY
                run.font.bold = True
                run.font.color.rgb = ACC if value else MUTE
            else:
                run.text = value
                run.font.size = Pt(11)
                run.font.name = FONT_BODY
                run.font.bold = ours
                run.font.color.rgb = INK
            x += widths[c]

    # outer frame
    frame = box(slide, x0, int(BODY_TOP + Inches(0.18)), total,
                int(header_h + row_h * len(rows)),
                fill=None, line=INK, shape=MSO_SHAPE.RECTANGLE)
    frame.shadow.inherit = False

    label(slide, MARGIN, Inches(6.40), CONTENT_W, Inches(0.6),
          "Публичные тарифы и возможности продуктов на начало 2026 года. WhisperFly "
          "сочетает то, что конкуренты предлагают только по отдельности: нулевую цену, "
          "системный звук и открытый исходный код.",
          size=9.5, color=MUTE, line_spacing=1.3)
    return slide


# ── 10 · Трекшн ──────────────────────────────────────────────────────────────


def slide_traction(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Трекшн")
    title(slide, [
        ("Готовый продукт, ", {}),
        ("не прототип", {"color": ACC}),
    ])

    width, xs = columns(4)
    top = int(BODY_TOP + Inches(0.05))
    stats = [
        ("2.0", "текущая версия", INK, ACC_INK),
        ("3", "сборки: Apple Silicon, Intel, универсальная", ACC, ACC_INK),
        ("10", "языков интерфейса", ACC2, INK),
        ("32", "коммита в основной ветке", INK, ACC_INK),
    ]
    for i, (number, caption, fill, fg) in enumerate(stats):
        card(slide, xs[i], top, width, Inches(1.72), fill=fill)
        stat_block(slide, xs[i] + Inches(0.26), top + Inches(0.34),
                   width - Inches(0.52), number, caption,
                   number_color=fg, caption_color=fg, number_size=26)

    facts_top = int(top + Inches(1.97))
    card(slide, MARGIN, facts_top, CONTENT_W, Inches(2.30), fill=SURF2)
    facts = [
        ("Установка одной командой", "через Homebrew или готовый DMG."),
        ("Совместимость", "с macOS 14 (Sonoma) вплоть до macOS 26 (Tahoe)."),
        ("Три проблемы macOS 26 решены:", "крэш панели, ложные отказы прав, тихая запись."),
        ("Открытый исходный код", "на GitHub — продукт можно проверить и собрать самому."),
    ]
    col_w = int((CONTENT_W - Inches(0.6) - Inches(0.5)) / 2)
    for i, (head, tail) in enumerate(facts):
        x = int(MARGIN + Inches(0.32) + (col_w + Inches(0.5)) * (i % 2))
        y = int(facts_top + Inches(0.40) + Inches(0.92) * (i // 2))
        mark = box(slide, x, y + Inches(0.03), Inches(0.22), Inches(0.22),
                   fill=None, line=ACC, shape=MSO_SHAPE.OVAL)
        mark.shadow.inherit = False
        label(slide, x + Inches(0.38), y, col_w - Inches(0.38), Inches(0.8),
              [(head + " ", {"bold": True, "color": INK}), (tail, {"color": MUTE})],
              size=10, line_spacing=1.24)
    return slide


# ── 11 · Роадмап ─────────────────────────────────────────────────────────────


def slide_roadmap(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Роадмап")
    title(slide, [
        ("От приложения к ", {}),
        ("платформе", {"color": ACC}),
    ])

    width, xs = columns(4)
    top = int(BODY_TOP + Inches(0.20))
    phases = [
        ("Q1 2026 · сейчас", ACC, [
            "Релиз 2.0: системный звук и файлы",
            "История транскрипций",
            "Совместимость с macOS 26",
            "Сборки и Homebrew",
        ]),
        ("Q2 2026", ACC2, [
            "WhisperFly Pro и подписка",
            "Пользовательский словарь",
            "Оплата и лицензирование",
            "Усиление команды: GTM",
        ]),
        ("Q3 2026", SURF2, [
            "Потоковая транскрипция",
            "Компаньон для iOS",
            "Версия для Windows",
            "Публичный API",
        ]),
        ("Q4 2026", SURF2, [
            "Админ-панель и единый вход",
            "Общие словари компании",
            "Развёртывание on-prem",
            "Сертификация безопасности",
        ]),
    ]
    for i, (label_text, bar_fill, items) in enumerate(phases):
        x = xs[i]
        bar = box(slide, x, top, width, Inches(0.28), fill=bar_fill, line=INK,
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        bar.shadow.inherit = False
        label(slide, x, top + Inches(0.48), width, Inches(0.4), label_text,
              size=13, font=FONT_DISPLAY, color=INK, bold=True, line_spacing=1.05)
        y = top + Inches(1.05)
        for item in items:
            dot = box(slide, x + Inches(0.02), y + Inches(0.07), Inches(0.13), Inches(0.13),
                      fill=SURF, line=INK, shape=MSO_SHAPE.OVAL, line_w=Pt(1.25))
            dot.shadow.inherit = False
            label(slide, x + Inches(0.30), y, width - Inches(0.30), Inches(0.7),
                  item, size=10, color=MUTE, line_spacing=1.22)
            y += Inches(0.80)
    return slide


# ── 12 · Команда ─────────────────────────────────────────────────────────────


def slide_team(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Команда")
    title(slide, [
        ("Один разработчик. ", {}),
        ("Весь стек.", {"color": ACC}),
    ])

    width, xs = columns(3)
    top = int(BODY_TOP + Inches(0.15))
    crew = [
        ("DS", "dandysuper", "Основатель · разработка · продукт",
         "Полный цикл: архитектура на Swift, интерфейс, дизайн, релизы и дистрибуция.",
         SURF, ACC, ACC_INK, INK, MUTE),
        ("<>", "Открытый код", "Сообщество",
         "Репозиторий на GitHub: код открыт, вклад и обратная связь приветствуются.",
         ACC2, INK, ACC2, INK, INK),
        ("HK", "hukopo", "Автор основы",
         "Автор проекта qwenwishper, из которого вырос WhisperFly.",
         SURF, ACC, ACC_INK, INK, MUTE),
    ]
    for i, (avatar, name, role, desc, fill, av_fill, av_fg, name_fg, role_fg) in enumerate(crew):
        x = xs[i]
        card(slide, x, top, width, Inches(3.55), fill=fill)
        size = Inches(1.15)
        circle = box(slide, int(x + (width - size) / 2), top + Inches(0.42), size, size,
                     fill=av_fill, line=INK, shape=MSO_SHAPE.OVAL)
        tf = circle.text_frame
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = avatar
        run.font.size = Pt(20 if avatar != "<>" else 16)
        run.font.name = FONT_DISPLAY
        run.font.bold = True
        run.font.color.rgb = av_fg

        label(slide, x, top + Inches(1.80), width, Inches(0.42), name,
              size=15, font=FONT_DISPLAY, color=name_fg, bold=True,
              align=PP_ALIGN.CENTER, line_spacing=1.05)
        label(slide, x + Inches(0.20), top + Inches(2.28), width - Inches(0.40),
              Inches(0.42), role, size=8.5, font=FONT_MONO, color=role_fg,
              bold=True, align=PP_ALIGN.CENTER, line_spacing=1.2, caps=True)
        label(slide, x + Inches(0.28), top + Inches(2.78), width - Inches(0.56),
              Inches(0.9), desc, size=10, color=role_fg if fill == ACC2 else MUTE,
              align=PP_ALIGN.CENTER, line_spacing=1.24)

    label(slide, MARGIN, Inches(6.42), CONTENT_W, Inches(0.5),
          "В планах — усиление команды: продуктовый маркетинг и продажи во втором "
          "квартале 2026 года.",
          size=9.5, color=MUTE, line_spacing=1.3)
    return slide


# ── 13 · Призыв ──────────────────────────────────────────────────────────────


def slide_ask(prs):
    slide = blank(prs)
    spine(slide)
    kicker(slide, "Что мы ищем")
    title(slide, [
        ("Помогите нам ", {}),
        ("взлететь", {"color": ACC}),
    ])

    width, xs = columns(3)
    top = int(BODY_TOP + Inches(0.15))
    asks = [
        ("Инвестиции", "Посевной раунд для запуска Pro-версии и найма специалиста по выходу на рынок."),
        ("Менторы", "Опыт вывода инструментов для продуктивности на массовый рынок и в B2B."),
        ("Партнёры", "Дистрибуция через сообщества разработчиков, IT-компании и вузы."),
    ]
    for i, (head, body) in enumerate(asks):
        card(slide, xs[i], top, width, Inches(2.35))
        card_head(slide, xs[i] + Inches(0.30), top + Inches(0.40),
                  width - Inches(0.60), head, size=14)
        card_body(slide, xs[i] + Inches(0.30), top + Inches(1.18),
                  width - Inches(0.60), Inches(1.00), body)

    cta_y = int(top + Inches(2.75))
    for i, (text, fill, fg) in enumerate([
        ("↓  Скачать бесплатно", ACC, ACC_INK),
        ("</>  GitHub", SURF, INK),
    ]):
        w = Inches(0.70) + Inches(0.115) * len(text)
        x = MARGIN if i == 0 else int(MARGIN + w + Inches(0.28))
        pill = box(slide, x, cta_y, w, Inches(0.62), fill=fill, line=INK,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        tf = pill.text_frame
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = text
        run.font.size = Pt(12)
        run.font.name = FONT_DISPLAY
        run.font.bold = True
        run.font.color.rgb = fg

    label(slide, MARGIN, cta_y + Inches(0.82), CONTENT_W, Inches(0.4),
          "Открытый исходный код · macOS 14+ · 10 языков",
          size=9.5, font=FONT_MONO, color=MUTE, bold=True, caps=True)
    return slide


# ── Entry point ──────────────────────────────────────────────────────────────

BUILDERS = [
    slide_title,
    slide_problem,
    slide_solution,
    slide_pipeline,
    slide_features,
    slide_stack,
    slide_market,
    slide_business,
    slide_competition,
    slide_traction,
    slide_roadmap,
    slide_team,
    slide_ask,
]

OUTPUT = "WhisperFly-Pitch-RU.pptx"


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    for build in BUILDERS:
        build(prs)

    prs.save(OUTPUT)
    print(f"wrote {OUTPUT} — {len(prs.slides.__iter__.__self__._sldIdLst)} slides")
    print(f"accent oklch({0.44} {0.16} {HUE}) -> #{ACC}")


if __name__ == "__main__":
    main()
