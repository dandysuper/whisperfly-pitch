#!/usr/bin/env python3
"""Build WhisperFly-Pitch-RU.pptx — the editable PowerPoint version of the deck.

Nine slides in Apple Cupertino / Keynote aesthetic, matching index.html:

    01 титульный лист   02 введение (+ скриншоты)  03 уникальность (+ скриншот)
    04 КПД              05 план реализации         06 ожидаемые результаты
    07 финансовая часть 08 команда                 09 заключение

Works with `python-pptx` if installed, or falls back to a built-in zero-dependency
OpenXML writer so `/usr/bin/python3 build_pptx.py` always succeeds.
"""

import math
import os
import zipfile
from xml.sax.saxutils import escape

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Inches, Pt
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

    def Inches(val):
        return int(round(val * 914400))

    def Pt(val):
        return int(round(val * 12700))

    def Emu(val):
        return int(val)

    class RGBColor:
        def __init__(self, r, g, b):
            self.r, self.g, self.b = int(r), int(g), int(b)
        def __str__(self):
            return f"{self.r:02X}{self.g:02X}{self.b:02X}"
        def __eq__(self, other):
            return isinstance(other, RGBColor) and (self.r, self.g, self.b) == (other.r, other.g, other.b)

    class MSO_SHAPE:
        RECTANGLE = "rect"
        ROUNDED_RECTANGLE = "roundRect"
        OVAL = "ellipse"

    class MSO_ANCHOR:
        TOP = "t"
        MIDDLE = "ctr"
        BOTTOM = "b"

    class PP_ALIGN:
        LEFT = "l"
        CENTER = "ctr"
        RIGHT = "r"

    class _Font:
        def __init__(self):
            self.size = Pt(12)
            self.name = "Inter"
            self.bold = False
            self.italic = False
            self.color = type("Color", (), {"rgb": RGBColor(0, 0, 0)})()

    class _Run:
        def __init__(self):
            self.text = ""
            self.font = _Font()

    class _Paragraph:
        def __init__(self):
            self.alignment = PP_ALIGN.LEFT
            self.line_spacing = 1.1
            self.space_after = 0
            self.runs = []
        def add_run(self):
            r = _Run()
            self.runs.append(r)
            return r

    class _TextFrame:
        def __init__(self):
            self.word_wrap = True
            self.margin_left = 0
            self.margin_right = 0
            self.margin_top = 0
            self.margin_bottom = 0
            self.vertical_anchor = MSO_ANCHOR.TOP
            self.paragraphs = [_Paragraph()]

    class _Fill:
        def __init__(self):
            self.mode = "none"
            self.fore_color = type("Color", (), {"rgb": None})()
        def background(self):
            self.mode = "none"
        def solid(self):
            self.mode = "solid"

    class _Line:
        def __init__(self):
            self.fill = _Fill()
            self.color = type("Color", (), {"rgb": None})()
            self.width = Pt(1)

    class _Shape:
        def __init__(self, kind, shape_type, x, y, w, h, pic_path=None):
            self.kind = kind  # "shape", "textbox", "pic"
            self.shape_type = shape_type
            self.x, self.y, self.w, self.h = int(x), int(y), int(w), int(h)
            self.pic_path = pic_path
            self.shadow = type("Shadow", (), {"inherit": False})()
            self.adjustments = [0.1]
            self.fill = _Fill()
            self.line = _Line()
            self.text_frame = _TextFrame()

    class _Shapes:
        def __init__(self):
            self.items = []
        def add_shape(self, shape_type, x, y, w, h):
            s = _Shape("shape", shape_type, x, y, w, h)
            self.items.append(s)
            return s
        def add_textbox(self, x, y, w, h):
            s = _Shape("textbox", MSO_SHAPE.RECTANGLE, x, y, w, h)
            self.items.append(s)
            return s
        def add_picture(self, path, x, y, w, h):
            s = _Shape("pic", MSO_SHAPE.RECTANGLE, x, y, w, h, pic_path=path)
            self.items.append(s)
            return s

    class _Slide:
        def __init__(self):
            self.shapes = _Shapes()

    class _SlideList:
        def __init__(self):
            self._slides = []
            self._sldIdLst = self._slides
        def add_slide(self, layout):
            s = _Slide()
            self._slides.append(s)
            return s

    class Presentation:
        def __init__(self):
            self.slide_width = Inches(13.333)
            self.slide_height = Inches(7.5)
            self.slide_layouts = [None] * 7
            self.slides = _SlideList()

        def save(self, out_path):
            _save_openxml_pptx(self, out_path)


def _save_openxml_pptx(prs, out_path):
    """Write a compliant OpenXML .pptx archive from our slide shapes."""
    # Read base skeleton from existing pptx if available, or keep non-slide parts
    base_parts = {}
    if os.path.exists(out_path):
        with zipfile.ZipFile(out_path, "r") as zin:
            for name in zin.namelist():
                if not name.startswith("ppt/slides/") and not name.startswith("ppt/media/"):
                    base_parts[name] = zin.read(name)

    media_map = {}  # path -> media_name (e.g. image1.png)
    media_bytes = {}

    def get_media(path):
        if path not in media_map:
            idx = len(media_map) + 1
            ext = os.path.splitext(path)[1].lower() or ".png"
            mname = f"image{idx}{ext}"
            media_map[path] = mname
            with open(path, "rb") as f:
                media_bytes[f"ppt/media/{mname}"] = f.read()
        return media_map[path]

    slide_xmls = {}
    slide_rels = {}

    for s_idx, slide in enumerate(prs.slides._slides, start=1):
        rels = [
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>'
        ]
        pic_rids = {}
        parts = [
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>",
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>',
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
        ]
        for obj_id, shp in enumerate(slide.shapes.items, start=2):
            if shp.kind == "pic":
                mname = get_media(shp.pic_path)
                if mname not in pic_rids:
                    rid = f"rId{len(rels) + 1}"
                    pic_rids[mname] = rid
                    rels.append(
                        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{mname}"/>'
                    )
                rid = pic_rids[mname]
                descr = escape(os.path.basename(shp.pic_path))
                parts.append(
                    f'<p:pic><p:nvPicPr><p:cNvPr id="{obj_id}" name="Picture {obj_id}" descr="{descr}"/>'
                    f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
                    f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
                    f'<p:spPr><a:xfrm><a:off x="{shp.x}" y="{shp.y}"/><a:ext cx="{shp.w}" cy="{shp.h}"/></a:xfrm>'
                    f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
                )
            else:
                is_tb = shp.kind == "textbox"
                tx_attr = ' txBox="1"' if is_tb else ""
                prst = shp.shape_type if isinstance(shp.shape_type, str) else "rect"
                av_xml = "<a:avLst/>"
                if prst == "roundRect" and shp.adjustments:
                    val = int(round(max(0.0, min(0.5, shp.adjustments[0])) * 100000))
                    av_xml = f'<a:avLst><a:gd name="adj" fmla="val {val}"/></a:avLst>'

                if shp.fill.mode == "solid" and shp.fill.fore_color.rgb is not None:
                    fill_xml = f'<a:solidFill><a:srgbClr val="{shp.fill.fore_color.rgb}"/></a:solidFill>'
                else:
                    fill_xml = "<a:noFill/>"

                if shp.line.color.rgb is not None and shp.line.fill.mode != "none":
                    ln_xml = f'<a:ln w="{int(shp.line.width)}"><a:solidFill><a:srgbClr val="{shp.line.color.rgb}"/></a:solidFill></a:ln>'
                else:
                    ln_xml = "<a:ln><a:noFill/></a:ln>" if not is_tb else ""

                tf = shp.text_frame
                wrap = "square" if tf.word_wrap else "none"
                anchor = tf.vertical_anchor if isinstance(tf.vertical_anchor, str) else "t"
                body_pr = f'<a:bodyPr wrap="{wrap}" lIns="{int(tf.margin_left)}" rIns="{int(tf.margin_right)}" tIns="{int(tf.margin_top)}" bIns="{int(tf.margin_bottom)}" anchor="{anchor}"/>'

                p_xmls = []
                for para in tf.paragraphs:
                    algn = para.alignment if isinstance(para.alignment, str) else "l"
                    spc_pct = int(round((para.line_spacing or 1.1) * 100000))
                    p_pr = f'<a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{spc_pct}"/></a:lnSpc></a:pPr>'
                    r_xmls = []
                    for r in para.runs:
                        sz = int(round((r.font.size / 12700.0) * 100))
                        b = "1" if r.font.bold else "0"
                        it = "1" if r.font.italic else "0"
                        col = str(r.font.color.rgb or RGBColor(29, 29, 31))
                        fname = escape(r.font.name or "Inter")
                        txt = escape(r.text or "")
                        r_xmls.append(
                            f'<a:r><a:rPr sz="{sz}" b="{b}" i="{it}"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
                            f'<a:latin typeface="{fname}"/><a:cs typeface="{fname}"/></a:rPr><a:t>{txt}</a:t></a:r>'
                        )
                    p_xmls.append(f"<a:p>{p_pr}{''.join(r_xmls)}</a:p>")

                parts.append(
                    f'<p:sp><p:nvSpPr><p:cNvPr id="{obj_id}" name="Shape {obj_id}"/><p:cNvSpPr{tx_attr}/><p:nvPr/></p:nvSpPr>'
                    f'<p:spPr><a:xfrm><a:off x="{shp.x}" y="{shp.y}"/><a:ext cx="{shp.w}" cy="{shp.h}"/></a:xfrm>'
                    f'<a:prstGeom prst="{prst}">{av_xml}</a:prstGeom>{fill_xml}{ln_xml}</p:spPr>'
                    f'<p:txBody>{body_pr}<a:lstStyle/>{"".join(p_xmls)}</p:txBody></p:sp>'
                )

        parts.append("</p:spTree></p:cSld></p:sld>")
        slide_xmls[f"ppt/slides/slide{s_idx}.xml"] = "".join(parts).encode("utf-8")
        slide_rels[f"ppt/slides/_rels/slide{s_idx}.xml.rels"] = (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{"".join(rels)}</Relationships>'
        ).encode("utf-8")

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for k, v in base_parts.items():
            zout.writestr(k, v)
        for k, v in media_bytes.items():
            zout.writestr(k, v)
        for k, v in slide_xmls.items():
            zout.writestr(k, v)
        for k, v in slide_rels.items():
            zout.writestr(k, v)


# ── Apple Cupertino Palette ──────────────────────────────────────────────────

HUE = 211

BG = RGBColor(0xF5, 0xF5, 0xF7)        # Apple system grouped light grey
SURF = RGBColor(0xFF, 0xFF, 0xFF)      # Pure white card surface
SURF2 = RGBColor(0xEC, 0xEC, 0xF1)     # Secondary soft grey surface
INK = RGBColor(0x1D, 0x1D, 0x1F)       # Apple Space Black text / dark card
DARK2 = RGBColor(0x28, 0x2A, 0x30)     # Elevated dark surface
MUTE = RGBColor(0x6E, 0x6E, 0x73)      # Apple secondary label grey
LINE = RGBColor(0xDC, 0xDC, 0xE2)      # Subtle hairline divider
ACC_INK = RGBColor(0xFF, 0xFF, 0xFF)   # White text on blue/dark
ACC = RGBColor(0x00, 0x71, 0xE3)       # Cupertino Blue
ACC2 = RGBColor(0xE8, 0xF2, 0xFF)      # Soft blue pill background
GREEN = RGBColor(0x28, 0xCD, 0x41)     # macOS green

# ── Type ─────────────────────────────────────────────────────────────────────

FONT_DISPLAY = "Inter"
FONT_BODY = "Inter"
FONT_MONO = "JetBrains Mono"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.60)
CONTENT_W = SLIDE_W - MARGIN * 2
BORDER = Pt(0.75)

APP_ICON = "assets/appicon.png"
SHOT_READY = "assets/screenshot-ready.png"
SHOT_REC = "assets/screenshot-recording.png"
SHOT_PERM = "assets/screenshot-permissions.png"
SHOT_PILL = "assets/screenshot-listening.png"


# ── Shape / text helpers ─────────────────────────────────────────────────────

def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def box(slide, x, y, w, h, fill=None, line=None, radius=None,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, line_w=BORDER):
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
          line_spacing=1.15, space_after=0, caps=False):
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


def shape_label(shape, text, *, size=10, font=FONT_MONO, color=INK, bold=True,
                caps=False, align=PP_ALIGN.CENTER):
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


def pill(slide, x, y, w, h, text, *, fill=SURF, color=INK, line=LINE,
         size=9, font=FONT_BODY, bold=True, caps=False, radius=0.5):
    shp = box(slide, x, y, w, h, fill=fill, line=line,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)
    return shape_label(shp, text, size=size, font=font, color=color,
                       bold=bold, caps=caps)


def kicker(slide, text, x=None, y=Inches(0.36), gold=False):
    x = MARGIN if x is None else x
    width = Inches(0.32) + Inches(0.082) * len(text)
    return pill(slide, x, y, width, Inches(0.30), text,
                fill=ACC2, color=ACC, line=None, size=8, font=FONT_BODY, caps=True)


def title(slide, runs, y=Inches(0.74), size=28, width=None):
    width = width or (CONTENT_W - Inches(1.4))
    return label(slide, MARGIN, y, width, Inches(0.85), runs,
                 size=size, font=FONT_DISPLAY, color=INK, bold=True,
                 line_spacing=1.04)


def subtitle(slide, text, y=Inches(1.36), width=None, size=10.5):
    width = width or (CONTENT_W - Inches(1.2))
    return label(slide, MARGIN, y, width, Inches(0.58), text,
                 size=size, color=MUTE, line_spacing=1.32)


def slide_number(slide, text, x=None, y=Inches(0.34)):
    width = Inches(0.52)
    x = (SLIDE_W - MARGIN - width) if x is None else x
    return pill(slide, x, y, width, Inches(0.28), text, fill=SURF, color=MUTE,
                line=LINE, size=8, font=FONT_MONO)


def spine(slide):
    """Apple Keynote clean light canvas with a subtle top accent bar."""
    bg = box(slide, 0, 0, SLIDE_W, SLIDE_H, fill=BG, shape=MSO_SHAPE.RECTANGLE)
    bg.shadow.inherit = False
    top_bar = box(slide, 0, 0, SLIDE_W, Inches(0.05), fill=ACC, shape=MSO_SHAPE.RECTANGLE)
    top_bar.shadow.inherit = False


GAP = Inches(0.20)


def columns(count, gap=GAP, total=None, margin=None):
    total = CONTENT_W if total is None else total
    margin = MARGIN if margin is None else margin
    width = int((total - gap * (count - 1)) / count)
    xs = [int(margin + i * (width + gap)) for i in range(count)]
    return width, xs


def card(slide, x, y, w, h, *, fill=SURF, line=LINE, radius=0.07):
    return box(slide, x, y, w, h, fill=fill, line=line,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)


def card_head(slide, x, y, w, text, color=INK, size=13):
    return label(slide, x, y, w, Inches(0.38), text, size=size, font=FONT_DISPLAY,
                 color=color, bold=True, line_spacing=1.05)


def card_body(slide, x, y, w, h, text, color=MUTE, size=9):
    return label(slide, x, y, w, h, text, size=size, font=FONT_BODY,
                 color=color, line_spacing=1.25)


def icon_badge(slide, x, y, size, glyph, *, fill=ACC2, color=ACC, line=None):
    shp = box(slide, x, y, size, size, fill=fill, line=line,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.22)
    return shape_label(shp, glyph, size=12, font=FONT_DISPLAY, color=color)


def stat_block(slide, x, y, w, number, caption, *, number_color=INK,
               caption_color=MUTE, number_size=26, caption_size=8):
    label(slide, x, y, w, Inches(0.52), number, size=number_size,
          font=FONT_DISPLAY, color=number_color, bold=True, line_spacing=1.0)
    label(slide, x, y + Inches(0.54), w, Inches(0.48), caption, size=caption_size,
          font=FONT_BODY, color=caption_color, bold=True, line_spacing=1.18, caps=True)


def bullet(slide, x, y, w, h, head, tail, *, head_color=INK, body_color=MUTE,
           size=8.5, mark_color=ACC, line_spacing=1.22):
    label(slide, x, y, Inches(0.18), h, "✓", size=size + 1, font=FONT_BODY,
          color=mark_color, bold=True, line_spacing=line_spacing)
    return label(slide, x + Inches(0.20), y, w - Inches(0.20), h,
                 [(head + " ", {"bold": True, "color": head_color}),
                  (tail, {"color": body_color})],
                 size=size, line_spacing=line_spacing)


def fact_column(slide, x, y, w, items, *, head_color=INK, body_color=MUTE,
                mark_color=ACC, size=8.5, step=0.30, height=0.80):
    for i, (head, tail) in enumerate(items):
        bullet(slide, x, y + Inches(step) * i, w, Inches(height), head, tail,
               head_color=head_color, body_color=body_color,
               mark_color=mark_color, size=size)


def strip(slide, y, heading, items, *, columns_count=3, height=None,
          item_step=None):
    if height is None:
        height = Inches(1.30) if columns_count <= 2 else Inches(1.34)
    band = card(slide, MARGIN, y, CONTENT_W, height, fill=SURF2, line=LINE)

    label(slide, MARGIN + Inches(0.26), y + Inches(0.18),
          CONTENT_W - Inches(0.52), Inches(0.24), heading,
          size=8, font=FONT_BODY, color=MUTE, bold=True, caps=True)

    gap = Inches(0.32)
    col_w = int((CONTENT_W - Inches(0.52) - gap * (columns_count - 1)) / columns_count)
    step = item_step if item_step is not None else Inches(0.30)

    for i, (head, tail) in enumerate(items):
        row = i // columns_count
        col = i % columns_count
        bullet(slide, MARGIN + Inches(0.26) + (col_w + gap) * col,
               y + Inches(0.48) + step * row, col_w, Inches(0.36),
               head, tail, size=8.5)
    return band


def badge_row(slide, y, items, width=None):
    widths = [Inches(0.30) + Inches(0.072) * len(t) for t in items]
    total = sum(widths) + Inches(0.10) * (len(items) - 1)
    start = (SLIDE_W - total) / 2 if width is None else width
    x = int(start)
    for i, text in enumerate(items):
        pill(slide, x, y, widths[i], Inches(0.30), text,
             fill=ACC if i == 0 else SURF,
             color=ACC_INK if i == 0 else INK,
             line=None if i == 0 else LINE,
             size=8, font=FONT_BODY)
        x += widths[i] + Inches(0.10)


# ── 01 · Титульный лист ──────────────────────────────────────────────────────

def slide_title(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "01/09")

    logo_s = Inches(1.65)
    slide.shapes.add_picture(APP_ICON, int((SLIDE_W - logo_s) / 2),
                             Inches(0.52), logo_s, logo_s)

    label(slide, MARGIN, Inches(2.38), CONTENT_W, Inches(1.05),
          [("Whisper", {}), ("Fly", {"color": ACC})],
          size=54, font=FONT_DISPLAY, color=INK, bold=True,
          align=PP_ALIGN.CENTER, line_spacing=1.0)

    label(slide, Inches(3.1), Inches(3.48), SLIDE_W - Inches(6.2), Inches(0.75),
          "Бесплатная диктовка нажатием клавиши для macOS. Голос — в текст: "
          "мгновенно, в любом приложении и без единого тенге за подписку.",
          size=12, color=MUTE, align=PP_ALIGN.CENTER, line_spacing=1.32)

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
        card(slide, x, Inches(4.48), meta_cell, Inches(0.92))
        label(slide, x + Inches(0.18), Inches(4.62), meta_cell - Inches(0.36),
              Inches(0.22), term, size=7.5, font=FONT_MONO, color=MUTE,
              bold=True, caps=True)
        label(slide, x + Inches(0.18), Inches(4.88), meta_cell - Inches(0.36),
              Inches(0.44), value, size=9.5, bold=True, line_spacing=1.24)

    badge_row(slide, Inches(5.78), [
        "PIZZA PITCH · SKO Hub", "16 октября 2026 · Петропавловск", "Готовый MVP 2.0",
        "macOS 14+", "11 языков", "0 ₸",
    ])
    return slide


# ── 02 · Введение (+ живые скриншоты macOS) ──────────────────────────────────

def slide_intro(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "02/09")
    kicker(slide, "Введение · Продукт в работе")
    title(slide, [("Что такое ", {}), ("WhisperFly", {"color": ACC})])
    subtitle(slide,
             "WhisperFly — нативное приложение для macOS, которое превращает речь в текст "
             "в любом приложении по нажатию ⌘⇧Space. Записывает микрофон и системный звук, "
             "распознаёт речь через Groq Whisper и вставляет чистый текст прямо к курсору.")

    top_y = Inches(2.12)
    col_h = Inches(4.65)
    left_w = Inches(3.80)
    mid_w = Inches(4.15)
    right_w = CONTENT_W - left_w - mid_w - Inches(0.36)

    # Left: Main Goal Card (Space Black)
    card(slide, MARGIN, top_y, left_w, col_h, fill=INK, line=None)
    icon_badge(slide, MARGIN + Inches(0.26), top_y + Inches(0.26), Inches(0.42),
               "◎", fill=DARK2, color=ACC)
    card_head(slide, MARGIN + Inches(0.26), top_y + Inches(0.84),
              left_w - Inches(0.52), "Главная цель", color=ACC_INK, size=13)
    label(slide, MARGIN + Inches(0.26), top_y + Inches(1.28), left_w - Inches(0.52),
          Inches(2.10),
          "Убрать разрыв между скоростью мысли и скоростью набора текста — сделать "
          "голосовой ввод бесплатным, мгновенным и доступным каждому на Mac.",
          size=10.5, color=ACC_INK, line_spacing=1.36)

    chips = ["macOS 14+", "Groq Whisper", "0 ₸"]
    chip_x = MARGIN + Inches(0.26)
    for i, text in enumerate(chips):
        w = Inches(0.26) + Inches(0.068) * len(text)
        pill(slide, chip_x, top_y + Inches(3.98), w, Inches(0.28), text,
             fill=ACC if i == 2 else DARK2, color=ACC_INK, line=None, size=7.5)
        chip_x += w + Inches(0.08)

    # Middle: 4 Tasks
    mid_x = MARGIN + left_w + Inches(0.18)
    label(slide, mid_x, top_y, mid_w, Inches(0.24),
          "Конкретные задачи", size=8, font=FONT_BODY, color=MUTE,
          bold=True, caps=True)
    tasks = [
        ("Мгновенный ввод.",
         "Один хоткей, авто-вставка текста в активное поле, отклик — доли секунды."),
        ("Все сценарии звука.",
         "Микрофон, системный звук macOS и готовые аудио/видеофайлы."),
        ("Без языкового барьера.",
         "Интерфейс на 11 языках (включая казахский), распознавание 100+ языков."),
        ("Выход на рынок.",
         "Лендинг, первые 1 000 установок и грант GitHub на AI-модели."),
    ]
    ty = top_y + Inches(0.32)
    for i, (head, tail) in enumerate(tasks):
        card(slide, mid_x, ty, mid_w, Inches(0.98), fill=SURF, line=LINE)
        num = box(slide, mid_x + Inches(0.14), ty + Inches(0.16), Inches(0.28), Inches(0.28),
                  fill=ACC2, line=None, shape=MSO_SHAPE.OVAL)
        shape_label(num, str(i + 1), size=8.5, font=FONT_MONO, color=ACC)
        label(slide, mid_x + Inches(0.50), ty + Inches(0.14), mid_w - Inches(0.62),
              Inches(0.74),
              [(head + " ", {"bold": True, "color": INK}), (tail, {"color": MUTE})],
              size=8.3, line_spacing=1.22)
        ty += Inches(1.08)

    # Right: Real macOS Screenshots Bento Card
    right_x = mid_x + mid_w + Inches(0.18)
    card(slide, right_x, top_y, right_w, col_h, fill=INK, line=None)
    label(slide, right_x + Inches(0.18), top_y + Inches(0.14), right_w - Inches(0.36),
          Inches(0.22), "ИНТЕРФЕЙС MACOS 2.0", size=7.5, font=FONT_MONO,
          color=ACC_INK, bold=True)

    shot_w = Inches(1.66)
    shot_h = Inches(2.14)
    if os.path.exists(SHOT_READY):
        slide.shapes.add_picture(SHOT_READY, int(right_x + Inches(0.18)),
                                 int(top_y + Inches(0.42)), shot_w, shot_h)
    if os.path.exists(SHOT_REC):
        slide.shapes.add_picture(SHOT_REC, int(right_x + Inches(0.28) + shot_w),
                                 int(top_y + Inches(0.42)), shot_w, shot_h)

    if os.path.exists(SHOT_PILL):
        pill_w = right_w - Inches(0.36)
        pill_h = int(pill_w * 96 / 360)
        slide.shapes.add_picture(SHOT_PILL, int(right_x + Inches(0.18)),
                                 int(top_y + Inches(2.74)), pill_w, pill_h)

    label(slide, right_x + Inches(0.18), top_y + Inches(3.82), right_w - Inches(0.36),
          Inches(0.68),
          "Строка меню Ready / Recording и плавающий индикатор Listening… прямо у курсора.",
          size=8, color=MUTE, line_spacing=1.24)
    return slide


# ── 03 · Уникальность (+ скриншот системных разрешений) ──────────────────────

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
    title(slide, [("Бесплатно, нативно и ", {}), ("открыто", {"color": ACC})])

    top = Inches(1.56)
    table_w = Inches(8.15)
    headers = ["Продукт", "Цена", "Системный звук", "Файлы", "Открытый код", "Казахский UI"]
    fractions = [0.21, 0.16, 0.17, 0.12, 0.16, 0.18]
    widths = [int(table_w * f) for f in fractions]
    widths[-1] = int(table_w - sum(widths[:-1]))

    header_h = Inches(0.42)
    row_h = Inches(0.52)

    x = MARGIN
    for i, text in enumerate(headers):
        fill = ACC if i == 0 else SURF2
        fg = ACC_INK if i == 0 else MUTE
        cell = box(slide, x, top, widths[i], header_h, fill=fill, line=LINE,
                   shape=MSO_SHAPE.RECTANGLE)
        tf = cell.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.12)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        run = para.add_run()
        run.text = text.upper()
        run.font.size = Pt(7.5)
        run.font.name = FONT_BODY
        run.font.bold = True
        run.font.color.rgb = fg
        x += widths[i]

    for r, row in enumerate(COMPETITORS):
        y = int(top + header_h + row_h * r)
        ours = r == 0
        x = MARGIN
        values = [row[0], row[1]] + list(row[2:])
        for c, value in enumerate(values):
            fill = ACC2 if ours else (SURF2 if r % 2 else SURF)
            cell = box(slide, x, y, widths[c], row_h, fill=fill, line=LINE,
                       shape=MSO_SHAPE.RECTANGLE)
            tf = cell.text_frame
            tf.word_wrap = False
            tf.margin_left = Inches(0.12)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = tf.paragraphs[0]
            run = para.add_run()
            if isinstance(value, bool):
                run.text = "✓" if value else "✗"
                run.font.size = Pt(11.5)
                run.font.name = FONT_BODY
                run.font.bold = True
                run.font.color.rgb = ACC if value else MUTE
            else:
                run.text = value
                run.font.size = Pt(9)
                run.font.name = FONT_BODY
                run.font.bold = ours
                run.font.color.rgb = INK
            x += widths[c]

    # Right: Permissions & Code Signature Screenshot Card
    perm_x = MARGIN + table_w + Inches(0.20)
    perm_w = CONTENT_W - table_w - Inches(0.20)
    perm_h = header_h + row_h * len(COMPETITORS)
    card(slide, perm_x, top, perm_w, perm_h, fill=INK, line=None)
    if os.path.exists(SHOT_PERM):
        img_h = Inches(2.26)
        img_w = int(img_h * 540 / 560)
        slide.shapes.add_picture(SHOT_PERM, int(perm_x + (perm_w - img_w) / 2),
                                 int(top + Inches(0.14)), img_w, img_h)
    label(slide, perm_x + Inches(0.16), top + Inches(2.48), perm_w - Inches(0.32),
          Inches(0.46),
          "ScreenCaptureKit, Accessibility API и подпись SR9LWJ7XG9",
          size=7.8, color=ACC_INK, align=PP_ALIGN.CENTER, line_spacing=1.2)

    strip(slide, Inches(4.88), "В чём инновационность", [
        ("Облачный ASR вместо локальной модели.",
         "Приложение весит мегабайты, а не гигабайты, и не греет Mac."),
        ("Вставка через Accessibility API:",
         "текст печатается прямо у курсора и очищается режимом Cleanup."),
        ("Системный звук через ScreenCaptureKit:",
         "созвоны и видео транскрибируются прямо с аудиодорожки Mac."),
    ], columns_count=3, height=Inches(1.62), item_step=Inches(0.0))
    return slide


# ── 04 · КПД ─────────────────────────────────────────────────────────────────

EFFECT_STEPS = [
    ("5 000", "ПОЛЬЗОВАТЕЛЕЙ ЗА ГОД"),
    ("40 мин", "ЭКОНОМИИ В ДЕНЬ НА ЧЕЛОВЕКА"),
    ("220", "РАБОЧИХ ДНЕЙ"),
    ("≈ 0,7 млн", "ЧЕЛОВЕКО-ЧАСОВ В ГОД"),
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
        ("◆", "Компетенциям", SURF, INK, MUTE, [
            ("Открытый код как учебный материал.", "3 000+ строк Swift на GitHub — "
                                                   "площадка для студентов."),
            ("Экспортный потенциал.", "Продукт для глобального macOS-рынка, "
                                      "созданный в регионе."),
            ("Рабочие места.", "План найма: продуктовый маркетинг, продажи "
                               "и поддержка в 2027 году."),
        ]),
    ]

    top = Inches(1.56)
    height = Inches(3.16)
    for i, (glyph, head, fill, head_color, body_color, facts) in enumerate(cards):
        x = xs[i]
        card(slide, x, top, width, height, fill=fill, line=None if fill == ACC else LINE)
        icon_badge(slide, x + Inches(0.26), top + Inches(0.24), Inches(0.40),
                   glyph,
                   fill=ACC_INK if fill == ACC else ACC2,
                   color=ACC)
        card_head(slide, x + Inches(0.26), top + Inches(0.80),
                  width - Inches(0.52), head, color=head_color, size=12.5)
        fact_column(slide, x + Inches(0.26), top + Inches(1.26),
                    width - Inches(0.52), facts,
                    head_color=head_color, body_color=body_color,
                    mark_color=ACC_INK if fill == ACC else ACC, size=8.2, step=0.56, height=0.80)

    band_y = Inches(4.92)
    band_h = Inches(1.92)
    card(slide, MARGIN, band_y, CONTENT_W, band_h, fill=SURF2, line=LINE)
    label(slide, MARGIN + Inches(0.26), band_y + Inches(0.18),
          CONTENT_W - Inches(0.52), Inches(0.22),
          "Расчёт эффекта · 12 месяцев", size=8, font=FONT_BODY, color=MUTE,
          bold=True, caps=True)

    op_zone = Inches(0.36)
    inner = CONTENT_W - Inches(0.52)
    item_w = int((inner - op_zone * 3) / 4)
    for i, (number, caption) in enumerate(EFFECT_STEPS):
        x = MARGIN + Inches(0.26) + (item_w + op_zone) * i
        last = i == len(EFFECT_STEPS) - 1
        label(slide, x, band_y + Inches(0.48), item_w, Inches(0.42), number,
              size=19, font=FONT_DISPLAY, color=ACC if last else INK, bold=True,
              line_spacing=1.0)
        label(slide, x, band_y + Inches(0.94), item_w, Inches(0.28), caption,
              size=7, font=FONT_MONO, color=MUTE, bold=True, line_spacing=1.2)
        if not last:
            label(slide, x + item_w, band_y + Inches(0.48), op_zone, Inches(0.40),
                  "×", size=15, font=FONT_DISPLAY, color=MUTE, align=PP_ALIGN.CENTER)

    label(slide, MARGIN + Inches(0.26), band_y + Inches(1.32),
          CONTENT_W - Inches(0.52), Inches(0.46),
          "Это эквивалент примерно 400 полных ставок в год — время, которое регион "
          "получает без единого тенге бюджетных затрат.",
          size=8.5, color=MUTE, line_spacing=1.3)
    return slide


# ── 05 · План реализации ─────────────────────────────────────────────────────

ROADMAP = [
    ("Q4 2026", "готово", GREEN, [
        "Релиз 2.0 выпущен",
        "Системный звук и файлы",
        "История и 11 языков",
        "Сборки: ARM, Intel, универсальная",
    ]),
    ("Q1 2027", "в работе", ACC, [
        "Лендинг продукта",
        "Заявка на грант GitHub",
        "Первые 1 000 установок",
        "Обратная связь и метрики",
    ]),
    ("Q2 2027", None, LINE, [
        "WhisperFly Pro",
        "Оплата и лицензии",
        "Свой словарь",
        "Найм: маркетинг",
    ]),
    ("Q3 2027", None, LINE, [
        "Потоковая запись",
        "iOS-компаньон",
        "Версия для Windows",
        "Публичный API",
    ]),
    ("Q4 2027", None, LINE, [
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
    subtitle(slide, "Пять этапов на 12 месяцев. Первый уже закрыт, остальные опираются "
                    "на работающий продукт, а не на прототип.", y=Inches(1.34))

    width, xs = columns(5)
    top = Inches(1.94)
    col_h = Inches(2.68)
    for i, (quarter, tag, bar_fill, items) in enumerate(ROADMAP):
        x = xs[i]
        card(slide, x, top, width, col_h, fill=SURF, line=LINE)
        bar = box(slide, x, top, width, Inches(0.08), fill=bar_fill, line=None,
                  shape=MSO_SHAPE.RECTANGLE)
        bar.shadow.inherit = False

        label(slide, x + Inches(0.18), top + Inches(0.22), width - Inches(0.36), Inches(0.30),
              quarter, size=12, font=FONT_DISPLAY, color=INK, bold=True, line_spacing=1.0)

        if tag:
            tag_w = Inches(0.26) + Inches(0.075) * len(tag)
            pill(slide, x + Inches(0.18), top + Inches(0.56), tag_w, Inches(0.24), tag,
                 fill=ACC if tag == "в работе" else ACC2,
                 color=ACC_INK if tag == "в работе" else ACC,
                 line=None, size=7, caps=True)

        for j, item in enumerate(items):
            y = top + Inches(0.94) + Inches(0.38) * j
            dot = box(slide, x + Inches(0.18), y + Inches(0.06), Inches(0.07),
                      Inches(0.07), fill=ACC, line=None, shape=MSO_SHAPE.OVAL)
            dot.shadow.inherit = False
            label(slide, x + Inches(0.32), y, width - Inches(0.44), Inches(0.34),
                  item, size=8.2, color=MUTE, line_spacing=1.22)

    strip(slide, Inches(4.82), "Почему план реалистичен", [
        ("Продукт уже работает.", "Ядро и интерфейс выпущены — предстоит развивать, "
                                  "а не начинать."),
        ("Нулевая инфраструктура.", "Движки на бесплатных тарифах, платить "
                                    "за серверы не нужно."),
        ("Один разработчик, весь стек.", "Дизайн, код, сборки и релизы — "
                                         "в одних руках."),
        ("Финансирование под контролем.", "Грант GitHub покрывает смету — "
                                          "инвестиции не нужны."),
    ], columns_count=4, height=Inches(1.68), item_step=Inches(0.0))
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
        ("5 000", "установок за 12 месяцев", INK, ACC_INK),
        ("300", "платящих пользователей Pro", ACC, ACC_INK),
        ("$10 000", "грант GitHub для стартапов", ACC2, ACC),
        ("0,7 млн", "сэкономленных человеко-часов в год", SURF, INK),
    ]
    for i, (number, caption, fill, fg) in enumerate(stats):
        card(slide, xs[i], top, width, height, fill=fill, line=LINE if fill == SURF else None)
        stat_block(slide, xs[i] + Inches(0.28), top + Inches(0.50),
                   width - Inches(0.56), number, caption,
                   number_color=fg, caption_color=fg if fill in (INK, ACC) else MUTE,
                   number_size=25)

    strip(slide, Inches(4.06), "Как мы оцениваем успех проекта", [
        ("Охват.", "Не менее 5 000 активных установок за 12 месяцев и не менее "
                   "1 000 в первом квартале."),
        ("Конверсия.", "Переход Free → Pro не ниже 5 % — при сохранении полностью "
                       "бесплатного ядра."),
        ("Удержание.", "Не менее 35 % пользователей возвращаются к продукту "
                       "на 30-й день после установки."),
        ("Финансирование.", "Грант GitHub для стартапов получен и покрывает смету "
                            "на 6 месяцев — без внешних инвестиций."),
    ], columns_count=2, height=Inches(2.32), item_step=Inches(0.66))
    return slide


# ── 07 · Финансовая часть ────────────────────────────────────────────────────

BUDGET = [
    ("Оплата труда основателя", "200 000 ₸ × 6 мес", "1 200 000 ₸"),
    ("AI-модели и облачные API", "$1 000 · уже оплачено", "530 000 ₸"),
    ("Продвижение лендинга и первые установки", "", "250 000 ₸"),
    ("Домен, хостинг и аналитика лендинга", "", "60 000 ₸"),
    ("Apple Developer Program", "$99 в год", "54 000 ₸"),
    ("Резерв на непредвиденные расходы", "", "106 000 ₸"),
]


def slide_finance(prs):
    slide = blank(prs)
    spine(slide)
    slide_number(slide, "07/09")
    kicker(slide, "Финансовая часть")
    title(slide, [("Что нужно и ", {}), ("откуда", {"color": ACC})])

    left_w = int(CONTENT_W * 0.60)
    col_a = int(left_w * 0.64)
    col_b = int(left_w - col_a)

    top = Inches(1.64)
    header_h = Inches(0.40)
    row_h = Inches(0.50)

    for idx, (head_text, w, x_off, align) in enumerate([
        ("Смета на 6 месяцев", col_a, MARGIN, PP_ALIGN.LEFT),
        ("Сумма", col_b, MARGIN + col_a, PP_ALIGN.RIGHT),
    ]):
        cell = box(slide, x_off, top, w, header_h, fill=SURF2, line=LINE,
                   shape=MSO_SHAPE.RECTANGLE)
        tf = cell.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.14)
        tf.margin_right = Inches(0.14)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.alignment = align
        run = para.add_run()
        run.text = head_text.upper()
        run.font.size = Pt(8)
        run.font.name = FONT_BODY
        run.font.bold = True
        run.font.color.rgb = MUTE

    for r, (item, note, amount) in enumerate(BUDGET):
        y = int(top + header_h + row_h * r)
        fill = SURF2 if r % 2 else SURF

        cell_a = box(slide, MARGIN, y, col_a, row_h, fill=fill, line=LINE,
                     shape=MSO_SHAPE.RECTANGLE)
        tf_a = cell_a.text_frame
        tf_a.word_wrap = True
        tf_a.margin_left = Inches(0.14)
        tf_a.vertical_anchor = MSO_ANCHOR.MIDDLE
        para_a = tf_a.paragraphs[0]
        run_a = para_a.add_run()
        run_a.text = item + (f"  ({note})" if note else "")
        run_a.font.size = Pt(8.8)
        run_a.font.name = FONT_BODY
        run_a.font.color.rgb = INK

        cell_b = box(slide, MARGIN + col_a, y, col_b, row_h, fill=fill, line=LINE,
                     shape=MSO_SHAPE.RECTANGLE)
        tf_b = cell_b.text_frame
        tf_b.word_wrap = False
        tf_b.margin_right = Inches(0.14)
        tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
        para_b = tf_b.paragraphs[0]
        para_b.alignment = PP_ALIGN.RIGHT
        run_b = para_b.add_run()
        run_b.text = amount
        run_b.font.size = Pt(9)
        run_b.font.name = FONT_MONO
        run_b.font.bold = True
        run_b.font.color.rgb = INK

    total_y = int(top + header_h + row_h * len(BUDGET))
    for text, w, x_off, align in [
        ("ИТОГО НА 6 МЕСЯЦЕВ", col_a, MARGIN, PP_ALIGN.LEFT),
        ("2 200 000 ₸", col_b, MARGIN + col_a, PP_ALIGN.RIGHT),
    ]:
        cell = box(slide, x_off, total_y, w, header_h, fill=INK, line=None,
                   shape=MSO_SHAPE.RECTANGLE)
        tf = cell.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.14)
        tf.margin_right = Inches(0.14)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.alignment = align
        run = para.add_run()
        run.text = text
        run.font.size = Pt(9.5)
        run.font.name = FONT_DISPLAY
        run.font.bold = True
        run.font.color.rgb = ACC_INK

    right_x = MARGIN + left_w + Inches(0.22)
    right_w = CONTENT_W - left_w - Inches(0.22)

    card(slide, right_x, top, right_w, Inches(2.14), fill=ACC, line=None)
    icon_badge(slide, right_x + Inches(0.28), top + Inches(0.24), Inches(0.40),
               "$", fill=ACC_INK, color=ACC)
    label(slide, right_x + Inches(0.28), top + Inches(0.78),
          right_w - Inches(0.56), Inches(0.22), "Запрашиваемая сумма", size=7.5,
          font=FONT_BODY, color=ACC_INK, bold=True, caps=True)
    label(slide, right_x + Inches(0.28), top + Inches(1.04),
          right_w - Inches(0.56), Inches(0.48), "$10 000", size=24,
          font=FONT_DISPLAY, color=ACC_INK, bold=True, line_spacing=1.0)
    label(slide, right_x + Inches(0.28), top + Inches(1.54),
          right_w - Inches(0.56), Inches(0.46),
          "Грант GitHub для стартапов — на AI-модели и вывод лендинга на рынок. "
          "Около 5 300 000 ₸: смета покрывается с запасом.",
          size=8.3, color=ACC_INK, line_spacing=1.24)

    fund_y = top + Inches(2.30)
    fund_h = Inches(2.14)
    card(slide, right_x, fund_y, right_w, fund_h, fill=SURF2, line=LINE)
    fact_column(slide, right_x + Inches(0.26), fund_y + Inches(0.24),
                right_w - Inches(0.52), [
                    ("Инвестиции не привлекаем.", "Призовой фонд PIZZA PITCH — "
                                                  "150 000 ₸; приз закрывает домен и продвижение."),
                    ("Уже вложено.", "Продукт 2.0, открытый код и подписка на топовые "
                                     "AI-модели на 6 месяцев — $1 000."),
                    ("Возвратность.", "WhisperFly Pro — с Q2 2027; себестоимость "
                                      "распознавания — центы на пользователя."),
                ], size=8, step=0.60, height=0.80)
    return slide


# ── 08 · Команда ─────────────────────────────────────────────────────────────

CREW = [
    ("ДБ", "Даниил Бурыкин", "Основатель · разработка · продукт",
     "Полный цикл: архитектура на Swift, интерфейс, конвейер распознавания, "
     "сборки, релизы и дистрибуция.", SURF, ACC, ACC_INK, INK, MUTE),
    ("<>", "Открытое сообщество", "Контрибьюторы · тестирование",
     "Репозиторий на GitHub открыт: вклад в код, переводы интерфейса "
     "и сообщения об ошибках приветствуются.", ACC2, INK, ACC_INK, INK, MUTE),
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
        card(slide, x, top, width, height, fill=fill, line=LINE)

        size = Inches(0.84)
        circle = box(slide, int(x + (width - size) / 2), top + Inches(0.30),
                     size, size, fill=av_fill, line=None, shape=MSO_SHAPE.OVAL)
        shape_label(circle, avatar, size=14 if avatar != "<>" else 12,
                    font=FONT_DISPLAY, color=av_fg)

        label(slide, x, top + Inches(1.32), width, Inches(0.34), name,
              size=13, font=FONT_DISPLAY, color=name_fg, bold=True,
              align=PP_ALIGN.CENTER, line_spacing=1.05)
        label(slide, x + Inches(0.18), top + Inches(1.70), width - Inches(0.36),
              Inches(0.34), role, size=7.5, font=FONT_MONO, color=ACC,
              bold=True, align=PP_ALIGN.CENTER, line_spacing=1.22, caps=True)
        label(slide, x + Inches(0.26), top + Inches(2.12), width - Inches(0.52),
              Inches(0.80), desc, size=8.5, color=body_fg,
              align=PP_ALIGN.CENTER, line_spacing=1.26)

    strip(slide, Inches(4.86), "Роль каждого в реализации", [
        ("Основатель", "— ведёт разработку, отвечает за продукт, релизы и дистрибуцию, "
                       "представляет проект."),
        ("Сообщество", "— тестирование на разных конфигурациях, локализация интерфейса, "
                       "обратная связь."),
        ("Планируемые роли", "— продуктовый маркетинг и продажи: найм запланирован "
                             "на Q2 2027 года."),
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
          y=Inches(0.80), size=26, width=CONTENT_W - Inches(0.4))

    label(slide, MARGIN, Inches(1.82), CONTENT_W - Inches(0.6), Inches(0.70),
          "WhisperFly уже работает и уже бесплатен: голосовой ввод в любом приложении "
          "macOS, собственный код, открытый исходный код и ноль затрат для пользователя. "
          "Осталось превратить готовый продукт в устойчивую платформу — и вернуть "
          "региону миллионы сэкономленных часов.",
          size=10, color=MUTE, line_spacing=1.34)

    width, xs = columns(3)
    top = Inches(2.74)
    height = Inches(2.12)
    closing = [
        ("★", "Уже работает",
         "Версия 2.0 выпущена: системный звук, транскрипция файлов, история "
         "и 11 языков интерфейса."),
        ("◆", "Бесплатно и открыто",
         "Ни подписки, ни локальных моделей, ни закрытых исходников — продукт "
         "можно проверить и собрать самому."),
        ("▲", "Эффект для региона",
         "0,7 млн человеко-часов в год, бесплатный доступ для образования "
         "и малого бизнеса, рабочие места в 2027 году."),
    ]
    for i, (glyph, head, body) in enumerate(closing):
        x = xs[i]
        card(slide, x, top, width, height)
        icon_badge(slide, x + Inches(0.28), top + Inches(0.26), Inches(0.42),
                   glyph, fill=ACC2, color=ACC)
        card_head(slide, x + Inches(0.28), top + Inches(0.88),
                  width - Inches(0.56), head, size=12)
        card_body(slide, x + Inches(0.28), top + Inches(1.30),
                  width - Inches(0.56), Inches(0.80), body, size=8.8)

    cta_y = Inches(5.20)
    ctas = [
        ("Сайт презентации", ACC, ACC_INK),
        ("Скачать бесплатно", SURF, INK),
        ("GitHub", SURF, INK),
    ]
    cta_x = MARGIN
    for text, fill, fg in ctas:
        w = Inches(0.50) + Inches(0.10) * len(text)
        pill(slide, cta_x, cta_y, w, Inches(0.46), text, fill=fill, color=fg,
             line=None if fill == ACC else LINE, size=9.5, font=FONT_DISPLAY)
        cta_x += w + Inches(0.16)

    label(slide, MARGIN, Inches(5.96), CONTENT_W, Inches(0.28),
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
    print(f"wrote {OUTPUT} — {len(prs.slides._sldIdLst)} slides (Apple Cupertino theme)")


if __name__ == "__main__":
    main()
