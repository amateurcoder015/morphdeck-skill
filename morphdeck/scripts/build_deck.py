#!/usr/bin/env python3
"""Build a cinematic PowerPoint deck from a morphdeck JSON spec.

Every slide shares a "stage" of decorative shapes named with the `!!` prefix.
PowerPoint's Morph transition matches shapes by that name and glides them
between their per-layout poses, so the whole deck plays like one continuous
camera move. Content on each slide then builds in with timed entrance
animations (rise, fade, zoom, wipe, letter-by-letter) that start automatically
after the transition.

Usage:
    python3 build_deck.py spec.json out.pptx [--theme NAME]
    python3 build_deck.py --list-themes
"""
import argparse
import itertools
import json
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
SW, SH = 13.333, 7.5  # 16:9 slide, inches

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_DECL = f'xmlns:p="{NS_P}" xmlns:a="{NS_A}"'

LAYOUTS = ("title", "section", "statement", "bullets", "stat", "cards",
           "timeline", "compare", "split", "closing")


def emu(inches):
    return Emu(int(round(inches * 914400)))


def load_themes():
    with open(os.path.join(HERE, "themes.json")) as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------

FILL_TAGS = ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill")


def _strip_style(shape):
    """Drop the theme style python-pptx attaches (it adds a drop shadow)."""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)


def _set_fill_xml(shape, fill_xml):
    sppr = shape._element.spPr
    for tag in FILL_TAGS:
        for el in sppr.findall(qn(tag)):
            sppr.remove(el)
    fill = etree.fromstring(f'<root xmlns:a="{NS_A}">{fill_xml}</root>')[0]
    geom = sppr.find(qn("a:prstGeom"))
    if geom is None:
        geom = sppr.find(qn("a:xfrm"))
    geom.addnext(fill)


def _clr(hex_, alpha=1.0):
    if alpha >= 1:
        return f'<a:srgbClr val="{hex_}"/>'
    return f'<a:srgbClr val="{hex_}"><a:alpha val="{int(alpha * 100000)}"/></a:srgbClr>'


def solid(shape, hex_, alpha=1.0):
    _set_fill_xml(shape, f"<a:solidFill>{_clr(hex_, alpha)}</a:solidFill>")


def glow(shape, hex_, alpha):
    """Radial gradient from `alpha` at the centre to transparent at the edge."""
    _set_fill_xml(shape, (
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        f'<a:gs pos="0">{_clr(hex_, alpha)}</a:gs>'
        f'<a:gs pos="45000">{_clr(hex_, alpha * 0.45)}</a:gs>'
        f'<a:gs pos="100000">{_clr(hex_, 0.0001)}</a:gs>'
        '</a:gsLst><a:path path="circle">'
        '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>'))


def no_line(shape):
    shape.line.fill.background()


def line(shape, hex_, width_pt, dash=None, alpha=1.0):
    shape.line.color.rgb = RGBColor.from_string(hex_)
    shape.line.width = Pt(width_pt)
    if dash:
        shape.line.dash_style = dash
    if alpha < 1:
        clr = shape._element.spPr.find(qn("a:ln")).find(".//" + qn("a:srgbClr"))
        etree.SubElement(clr, qn("a:alpha")).set("val", str(int(alpha * 100000)))


# ---------------------------------------------------------------------------
# Text
# ---------------------------------------------------------------------------

def _wrap_lines(text, chars_per_line):
    lines = 0
    for para in text.split("\n"):
        cur = 0
        lines += 1
        for word in para.split():
            need = len(word) + (1 if cur else 0)
            if cur and cur + need > chars_per_line:
                lines += 1
                cur = len(word)
            else:
                cur += need
    return lines


def fit(text, w, h, max_pt, min_pt=10, wf=0.64, lh=1.18):
    """Largest point size whose estimated wrap fits a w x h inch box.

    `wf` is the average glyph width as a fraction of the point size; Unbounded
    is a wide face, so the defaults are deliberately conservative.
    """
    for pt in range(int(max_pt), int(min_pt) - 1, -1):
        cpl = max(1, int(w / (pt / 72 * wf)))
        if max((len(word) for word in text.split()), default=0) > cpl:
            continue  # never let a single word break mid-way
        if _wrap_lines(text, cpl) * pt / 72 * lh <= h:
            return pt
    return min_pt


class Deck:
    def __init__(self, spec, theme):
        self.spec = spec
        self.t = theme
        fonts = spec.get("fonts", {})
        self.head_font = fonts.get("heading", "Unbounded")
        self.body_font = fonts.get("body", "Unbounded")
        self.prs = Presentation()
        self.prs.slide_width = emu(SW)
        self.prs.slide_height = emu(SH)
        self.blank = self.prs.slide_layouts[6]
        self.layout_count = {}

    # -- shapes -------------------------------------------------------------

    def shape(self, slide, kind, name, x, y, w, h, rot=0):
        sh = slide.shapes.add_shape(kind, emu(x), emu(y), emu(w), emu(h))
        _strip_style(sh)
        sh.name = name
        sh.rotation = rot
        return sh

    def text(self, slide, name, x, y, w, h, paras, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, shape=None):
        """paras: list of dicts {text, size, bold, color, font, spc, alpha, lh, after}."""
        if shape is None:
            shape = slide.shapes.add_textbox(emu(x), emu(y), emu(w), emu(h))
            shape.name = name
        tf = shape.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for i, p in enumerate(paras):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            para.line_spacing = p.get("lh", 1.0)
            if p.get("after"):
                para.space_after = Pt(p["after"])
            run = para.add_run()
            run.text = p["text"]
            f = run.font
            f.name = p.get("font", self.body_font)
            f.size = Pt(p["size"])
            f.bold = p.get("bold", False)
            f.color.rgb = RGBColor.from_string(p.get("color", self.t["text"]))
            rpr = run._r.get_or_add_rPr()
            if p.get("spc"):
                rpr.set("spc", str(p["spc"]))
            if p.get("alpha", 1) < 1:
                clr = rpr.find(".//" + qn("a:srgbClr"))
                etree.SubElement(clr, qn("a:alpha")).set("val", str(int(p["alpha"] * 100000)))
        return shape

    def head(self, text, w, h, max_pt, min_pt=18, color=None, lh=1.02):
        size = fit(text, w, h, max_pt, min_pt, wf=0.74, lh=lh * 1.12)
        return {"text": text, "size": size, "bold": True, "font": self.head_font,
                "color": color or self.t["text"], "lh": lh}

    def body(self, text, w, h, max_pt=20, min_pt=11, color=None):
        size = fit(text, w, h, max_pt, min_pt, wf=0.60, lh=1.45)
        return {"text": text, "size": size, "color": color or self.t["muted"], "lh": 1.25}

    def kicker(self, text, color=None):
        return {"text": text.upper(), "size": 12, "bold": True, "spc": 400,
                "color": color or self.t["accent"], "font": self.head_font}

    def card(self, slide, name, x, y, w, h, paras, fill=None, alpha=1.0, border=True,
             anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, pad=0.32):
        sh = self.shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, name, x, y, w, h)
        sh.adjustments[0] = min(0.12, 0.25 / min(w, h))
        solid(sh, fill or self.t["surface"], alpha)
        if border:
            line(sh, self.t["accent"], 0.75, alpha=0.35)
        else:
            no_line(sh)
        self.text(slide, name, 0, 0, 0, 0, paras, align=align, anchor=anchor, shape=sh)
        tf = sh.text_frame
        tf.margin_left = tf.margin_right = emu(pad)
        tf.margin_top = tf.margin_bottom = emu(pad)
        return sh

    # -- stage --------------------------------------------------------------

    def stage(self, slide, layout, idx, n):
        """Persistent `!!` shapes. Their pose per layout is what Morph animates."""
        t = self.t
        k = self.layout_count.get(layout, 0)
        spin = 70 * idx  # keeps the dashed ring turning even across repeated layouts
        P = {
            # name: (x, y, w, h, rotation)
            "title":     dict(ga=(6.6, -1.8, 8.6, 8.6), gb=(-2.6, 3.6, 5.6, 5.6), ring=(8.4, 0.7, 5.4, 5.4),
                              disc=(11.6, 5.7, 0.6, 0.6), panel=(0, 0, 0.16, SH), bar=(0.9, 5.05, 1.5, 0.09)),
            "section":   dict(ga=(-2.5, -2.0, 9.5, 9.5), gb=(9.5, 4.0, 5.0, 5.0), ring=(0.9, 1.25, 5.0, 5.0),
                              disc=(5.4, 1.35, 0.5, 0.5), panel=(0, 0, 0.16, SH), bar=(6.6, 4.25, 1.2, 0.09)),
            "statement": dict(ga=(-3.2, -3.2, 9.0, 9.0), gb=(9.2, 3.4, 6.0, 6.0), ring=(-1.9, 4.2, 4.6, 4.6),
                              disc=(12.0, 0.9, 0.45, 0.45), panel=(1.3, 1.15, 10.7, 5.2), bar=(1.95, 1.75, 0.9, 0.09)),
            "bullets":   dict(ga=(7.6, 0.6, 8.0, 8.0), gb=(-3.0, -3.0, 5.5, 5.5), ring=(9.1, 1.45, 4.6, 4.6),
                              disc=(9.0, 1.2, 0.55, 0.55), panel=(8.9, 0, SW - 8.9, SH), bar=(0.9, 2.0, 1.0, 0.09)),
            "stat":      dict(ga=(6.8, -0.6, 8.0, 8.0), gb=(-3.0, 4.2, 5.5, 5.5), ring=(7.9, 0.95, 5.6, 5.6),
                              disc=(12.1, 0.8, 0.5, 0.5), panel=(0, 0, 0.16, SH), bar=(0.9, 5.05, 1.2, 0.09)),
            "cards":     dict(ga=(8.5, -4.0, 7.5, 7.5), gb=(-2.5, 4.4, 5.5, 5.5), ring=(10.9, -1.6, 3.4, 3.4),
                              disc=(12.25, 0.55, 0.4, 0.4), panel=(0, 6.95, SW, 0.55), bar=(0.9, 1.85, 1.0, 0.09)),
            "timeline":  dict(ga=(-3.5, -4.5, 8.0, 8.0), gb=(9.6, 3.9, 5.8, 5.8), ring=(11.2, -1.4, 3.4, 3.4),
                              disc=(0.62, 4.13, 0.3, 0.3), panel=(0, 0, 0.16, SH), bar=(0.9, 4.24, 11.6, 0.05)),
            "compare":   dict(ga=(3.6, -4.5, 6.0, 6.0), gb=(3.6, 5.0, 6.0, 6.0), ring=(5.97, 3.37, 1.4, 1.4),
                              disc=(6.47, 3.87, 0.4, 0.4), panel=(6.635, 2.05, 0.06, 4.7), bar=(0.9, 1.85, 1.0, 0.09)),
            "split":     dict(ga=(-3.0, 3.5, 6.5, 6.5), gb=(8.0, -1.5, 7.5, 7.5), ring=(7.6, 1.5, 4.5, 4.5),
                              disc=(7.25, 5.9, 0.5, 0.5), panel=(7.0, 0, SW - 7.0, SH), bar=(0.9, 3.3, 1.0, 0.09)),
            "closing":   dict(ga=(2.6, -2.0, 8.2, 8.2), gb=(-2.5, 4.5, 5.0, 5.0), ring=(3.67, 0.25, 6.0, 6.0),
                              disc=(9.3, 1.0, 0.55, 0.55), panel=(0, 0, 0.16, SH), bar=(6.17, 4.4, 1.0, 0.09)),
        }[layout]
        g = t["glow"]
        x, y, w, h = P["ga"]
        glow(self.shape(slide, MSO_SHAPE.OVAL, "!!glow_a", x, y, w, h), t["accent"], g)
        x, y, w, h = P["gb"]
        glow(self.shape(slide, MSO_SHAPE.OVAL, "!!glow_b", x, y, w, h), t["accent2"], g * 0.8)
        x, y, w, h = P["panel"]
        panel = self.shape(slide, MSO_SHAPE.RECTANGLE, "!!panel", x, y, w, h)
        solid(panel, t["accent"] if layout in ("compare", "title", "section", "stat", "timeline", "closing")
              else t["surface"], 0.9 if layout != "cards" else 0.6)
        no_line(panel)
        x, y, w, h = P["ring"]
        ring = self.shape(slide, MSO_SHAPE.OVAL, "!!ring", x, y, w, h, rot=(spin + 37 * k) % 360)
        _set_fill_xml(ring, "<a:noFill/>")
        line(ring, t["accent2"], 1.5, dash=MSO_LINE.DASH, alpha=0.7)
        x, y, w, h = P["disc"]
        disc = self.shape(slide, MSO_SHAPE.OVAL, "!!disc", x, y, w, h)
        solid(disc, t["accent2"])
        no_line(disc)
        x, y, w, h = P["bar"]
        bar = self.shape(slide, MSO_SHAPE.RECTANGLE, "!!bar", x, y, w, h)
        solid(bar, t["accent"])
        no_line(bar)
        prog = self.shape(slide, MSO_SHAPE.RECTANGLE, "!!progress", 0, SH - 0.07,
                          max(0.05, SW * (idx + 1) / n), 0.07)
        solid(prog, t["accent2"], 0.85)
        no_line(prog)
        self.layout_count[layout] = k + 1
        return {"disc": disc, "ring": ring, "bar": bar}

    # -- layouts ------------------------------------------------------------

    def build(self):
        slides = self.spec["slides"]
        n = len(slides)
        for idx, s in enumerate(slides):
            layout = s.get("layout", "bullets")
            if layout not in LAYOUTS:
                sys.exit(f"slide {idx + 1}: unknown layout '{layout}' (use one of {', '.join(LAYOUTS)})")
            slide = self.prs.slides.add_slide(self.blank)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor.from_string(self.t["bg"])
            stage = self.stage(slide, layout, idx, n)
            anim = Anim()
            getattr(self, "L_" + layout)(slide, s, anim, stage)
            # ambient loops on the stage: gentle float + a slow sway
            anim.add("float", stage["disc"], 200, 2600, dy=-0.025)
            anim.add("sway", stage["ring"], 0, 6000, deg=12)
            if s.get("notes"):
                slide.notes_slide.notes_text_frame.text = s["notes"]
            dur = int(self.t["tempo"] * 1000) if idx else 0
            if idx:
                add_morph(slide, dur)
            anim.attach(slide)
        return self.prs

    def L_title(self, slide, s, a, st):
        t = self.t
        if s.get("kicker"):
            k = self.text(slide, "kicker", 0.9, 1.55, 8.0, 0.4, [self.kicker(s["kicker"])])
            a.add("fade", k, 0, 600)
        h = self.text(slide, "title", 0.9, 1.95, 8.3, 2.95, [self.head(s["title"], 8.3, 2.95, 66)],
                      anchor=MSO_ANCHOR.BOTTOM)
        a.add("letters", h, 150, 450)
        if s.get("subtitle"):
            sub = self.text(slide, "subtitle", 0.9, 5.4, 7.6, 1.3, [self.body(s["subtitle"], 7.6, 1.3, 20)])
            a.add("rise", sub, 900, 700)
        if s.get("byline"):
            b = self.text(slide, "byline", 0.9, 6.75, 7.6, 0.4,
                          [{"text": s["byline"], "size": 11, "color": t["muted"]}])
            a.add("fade", b, 1300, 600)

    def L_section(self, slide, s, a, st):
        t = self.t
        num = s.get("number", "")
        if num:
            nb = self.text(slide, "number", 0.9, 1.25, 5.0, 5.0,
                           [{"text": num, "size": fit(num, 4.2, 3.0, 150, 60, wf=0.8), "bold": True,
                             "font": self.head_font, "color": t["accent"]}],
                           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            a.add("zoom", nb, 0, 700)
        h = self.text(slide, "title", 6.6, 1.3, 6.0, 2.8, [self.head(s["title"], 6.0, 2.8, 50)],
                      anchor=MSO_ANCHOR.BOTTOM)
        a.add("letters", h, 300, 450)
        if s.get("subtitle"):
            sub = self.text(slide, "subtitle", 6.6, 4.6, 5.9, 1.6, [self.body(s["subtitle"], 5.9, 1.6, 18)])
            a.add("rise", sub, 1000, 700)

    def L_statement(self, slide, s, a, st):
        t = self.t
        q = self.text(slide, "statement", 1.95, 2.05, 9.6, 3.2,
                      [self.head(s["text"], 9.6, 3.1, 42, 20, lh=1.1)], anchor=MSO_ANCHOR.MIDDLE)
        a.add("rise", q, 0, 900, by_word=True)
        if s.get("attribution"):
            at = self.text(slide, "attribution", 1.95, 5.45, 9.4, 0.5,
                           [{"text": "— " + s["attribution"], "size": 14, "color": t["accent"],
                             "font": self.head_font}])
            a.add("fade", at, 1300, 700)

    def L_bullets(self, slide, s, a, st):
        t = self.t
        h = self.text(slide, "title", 0.9, 0.7, 7.6, 1.2, [self.head(s["title"], 7.6, 1.2, 36)],
                      anchor=MSO_ANCHOR.BOTTOM)
        a.add("rise", h, 0, 700)
        pts = s.get("points", [])[:6]
        top, bottom = 2.45, 6.85
        row = (bottom - top) / max(1, len(pts))
        size = min(fit(p, 6.9, row - 0.15, 20, 11, wf=0.60, lh=1.4) for p in pts) if pts else 18
        for i, p in enumerate(pts):
            y = top + i * row
            dot = self.shape(slide, MSO_SHAPE.OVAL, f"dot{i}", 0.92, y + size / 72 * 0.42, 0.14, 0.14)
            solid(dot, t["accent2"] if i % 2 else t["accent"])
            no_line(dot)
            tx = self.text(slide, f"point{i}", 1.35, y, 6.9, row - 0.1,
                           [{"text": p, "size": size, "color": t["text"], "lh": 1.2}])
            a.add("zoom", dot, 500 + i * 220, 400)
            a.add("rise", tx, 550 + i * 220, 600)
        if s.get("aside"):
            asd = self.text(slide, "aside", 9.5, 2.6, 3.4, 2.4,
                            [self.head(s["aside"], 3.4, 2.4, 30, 14, color=t["text"])],
                            anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
            a.add("zoom", asd, 400 + len(pts) * 220, 700)

    def L_stat(self, slide, s, a, st):
        t = self.t
        value, suffix = str(s.get("value", "")), str(s.get("suffix", ""))
        size = fit(value + suffix, 6.8, 3.0, 170, 48, wf=0.82, lh=1.0)
        v = self.text(slide, "value", 0.9, 1.0, 7.2, 3.9, [], anchor=MSO_ANCHOR.BOTTOM)
        para = v.text_frame.paragraphs[0]
        para.line_spacing = 0.9
        for txt, col, sz in ((value, t["text"], size), (suffix, t["accent"], int(size * 0.7))):
            if not txt:
                continue
            r = para.add_run()
            r.text = txt
            r.font.name, r.font.size, r.font.bold = self.head_font, Pt(sz), True
            r.font.color.rgb = RGBColor.from_string(col)
        a.add("zoom", v, 0, 800)
        lbl = self.text(slide, "label", 0.9, 5.3, 6.8, 0.75, [self.head(s.get("label", ""), 6.8, 0.75, 24, 14)])
        a.add("rise", lbl, 600, 600)
        if s.get("note"):
            nt = self.text(slide, "note", 0.9, 6.05, 6.8, 1.0, [self.body(s["note"], 6.8, 1.0, 15)])
            a.add("fade", nt, 1000, 700)
        if s.get("context"):
            c = self.text(slide, "context", 8.9, 2.75, 3.6, 2.0,
                          [self.body(s["context"], 3.6, 2.0, 18, 11, color=t["text"])],
                          anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
            a.add("fade", c, 1200, 800)

    def L_cards(self, slide, s, a, st):
        t = self.t
        h = self.text(slide, "title", 0.9, 0.55, 11.5, 1.2, [self.head(s["title"], 11.5, 1.2, 36)],
                      anchor=MSO_ANCHOR.BOTTOM)
        a.add("rise", h, 0, 700)
        cards = s.get("cards", [])[:4]
        gap, x0, x1, y, hgt = 0.3, 0.9, SW - 0.9, 2.35, 4.15
        w = (x1 - x0 - gap * (len(cards) - 1)) / max(1, len(cards))
        inner = w - 0.64
        tsize = min(fit(c.get("title", ""), inner, 1.1, 22, 13, wf=0.70) for c in cards) if cards else 20
        bsize = min(fit(c.get("text", ""), inner, 2.4, 18, 10, wf=0.60, lh=1.45) for c in cards) if cards else 14
        for i, c in enumerate(cards):
            paras = [{"text": f"{i + 1:02d}", "size": 14, "bold": True, "font": self.head_font,
                      "color": t["accent"], "after": 14},
                     {"text": c.get("title", ""), "size": tsize, "bold": True, "font": self.head_font,
                      "color": t["text"], "after": 10, "lh": 1.05},
                     {"text": c.get("text", ""), "size": bsize, "color": t["muted"], "lh": 1.25}]
            cd = self.card(slide, f"card{i}", x0 + i * (w + gap), y, w, hgt, paras)
            a.add("rise", cd, 500 + i * 250, 700)

    def L_timeline(self, slide, s, a, st):
        t = self.t
        h = self.text(slide, "title", 0.9, 0.55, 11.5, 1.2, [self.head(s["title"], 11.5, 1.2, 36)],
                      anchor=MSO_ANCHOR.BOTTOM)
        a.add("rise", h, 0, 700)
        steps = s.get("steps", [])[:6]
        n = max(1, len(steps))
        span = 11.6 / n
        lsize = min(fit(st_.get("label", ""), span - 0.3, 0.9, 30, 12, wf=0.80) for st_ in steps) if steps else 24
        bsize = min(fit(st_.get("text", ""), span - 0.3, 2.0, 15, 10, wf=0.60, lh=1.45) for st_ in steps) if steps else 13
        for i, st_ in enumerate(steps):
            cx = 0.9 + span * i + 0.1
            node = self.shape(slide, MSO_SHAPE.OVAL, f"node{i}", cx, 4.115, 0.3, 0.3)
            solid(node, t["bg"])
            line(node, t["accent"], 2.25)
            lb = self.text(slide, f"label{i}", cx, 2.85, span - 0.3, 1.0,
                           [{"text": st_.get("label", ""), "size": lsize, "bold": True,
                             "font": self.head_font, "color": t["accent"] if i % 2 == 0 else t["accent2"]}],
                           anchor=MSO_ANCHOR.BOTTOM)
            tx = self.text(slide, f"step{i}", cx, 4.75, span - 0.3, 2.1,
                           [{"text": st_.get("text", ""), "size": bsize, "color": t["text"], "lh": 1.25}])
            d = 500 + i * 380
            a.add("zoom", node, d, 400)
            a.add("rise", lb, d + 120, 600)
            a.add("fade", tx, d + 250, 700)

    def L_compare(self, slide, s, a, st):
        t = self.t
        h = self.text(slide, "title", 0.9, 0.55, 11.5, 1.2, [self.head(s["title"], 11.5, 1.2, 36)],
                      anchor=MSO_ANCHOR.BOTTOM)
        a.add("rise", h, 0, 700)
        for side, x, col, d in (("left", 0.9, t["accent"], 500), ("right", 7.15, t["accent2"], 800)):
            c = s.get(side, {})
            pts = c.get("points", [])
            bsize = min([fit(p, 4.7, 3.4 / max(1, len(pts)), 17, 10, wf=0.60, lh=1.4) for p in pts] or [16])
            paras = [{"text": c.get("title", "").upper(), "size": 14, "bold": True, "spc": 300,
                      "font": self.head_font, "color": col, "after": 16}]
            paras += [{"text": "—  " + p, "size": bsize, "color": t["text"], "after": 10, "lh": 1.2} for p in pts]
            cd = self.card(slide, side, x, 2.15, 5.3, 4.5, paras, border=False, pad=0.4)
            a.add("rise", cd, d, 700)

    def L_split(self, slide, s, a, st):
        t = self.t
        if s.get("kicker"):
            k = self.text(slide, "kicker", 0.9, 1.1, 5.6, 0.4, [self.kicker(s["kicker"])])
            a.add("fade", k, 0, 500)
        h = self.text(slide, "title", 0.9, 1.5, 5.6, 1.7, [self.head(s["title"], 5.6, 1.7, 36)],
                      anchor=MSO_ANCHOR.BOTTOM)
        a.add("rise", h, 100, 700)
        if s.get("text"):
            b = self.text(slide, "body", 0.9, 3.65, 5.6, 3.2, [self.body(s["text"], 5.6, 3.2, 18, 11)])
            a.add("rise", b, 600, 700)
        img = s.get("image")
        if img:
            img = img if os.path.isabs(img) else os.path.join(self.spec.get("_dir", "."), img)
            pic = slide.shapes.add_picture(img, emu(7.0), 0, emu(SW - 7.0), emu(SH))
            pic.name = "image"
            _cover_crop(pic, (SW - 7.0) / SH)
            a.add("fade", pic, 300, 1000)
        elif s.get("aside"):
            asd = self.text(slide, "aside", 8.0, 2.4, 3.6, 2.7,
                            [self.head(s["aside"], 3.6, 2.7, 34, 14)],
                            anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
            a.add("zoom", asd, 900, 700)

    def L_closing(self, slide, s, a, st):
        t = self.t
        h = self.text(slide, "title", 1.4, 1.6, 10.53, 2.6, [self.head(s["title"], 10.5, 2.6, 60)],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
        a.add("letters", h, 100, 450)
        if s.get("subtitle"):
            sub = self.text(slide, "subtitle", 2.2, 4.75, 8.93, 1.3, [self.body(s["subtitle"], 8.9, 1.3, 20)],
                            align=PP_ALIGN.CENTER)
            a.add("rise", sub, 900, 700)


def _cover_crop(pic, target_ratio):
    iw, ih = pic.image.size
    ratio = iw / ih
    if ratio > target_ratio:
        c = (1 - target_ratio / ratio) / 2
        pic.crop_left = pic.crop_right = c
    else:
        c = (1 - ratio / target_ratio) / 2
        pic.crop_top = pic.crop_bottom = c


# ---------------------------------------------------------------------------
# Transition + animation XML
# ---------------------------------------------------------------------------

def _insert_before_timing(slide, el):
    sld = slide._element
    anchor = sld.find(qn("p:timing"))
    if anchor is None:
        anchor = sld.find(qn("p:extLst"))
    if anchor is not None:
        anchor.addprevious(el)
    else:
        sld.append(el)


def add_morph(slide, dur_ms):
    xml = (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
        f'<p:transition {NS_DECL} xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
        f'spd="slow" p14:dur="{dur_ms}"><p159:morph option="byObject"/></p:transition>'
        '</mc:Choice><mc:Fallback>'
        f'<p:transition {NS_DECL} spd="slow"><p:fade/></p:transition>'
        '</mc:Fallback></mc:AlternateContent>')
    _insert_before_timing(slide, etree.fromstring(xml))


class Anim:
    """Collects effects; all start automatically once the slide (and its Morph) lands.

    Each effect is "With Previous" plus its own delay, which is how PowerPoint
    stores a staggered After-Previous chain, so it stays editable in the
    Animation Pane.
    """

    PRESETS = {  # kind: (presetClass, presetID, presetSubtype)
        "fade": ("entr", 10, 0), "rise": ("entr", 42, 0), "letters": ("entr", 42, 0),
        "zoom": ("entr", 53, 16), "wipe": ("entr", 22, 8), "float": ("path", 42, 0),
        "sway": ("emph", 8, 0),
    }

    def __init__(self):
        self.effects = []

    def add(self, kind, shape, delay, dur=600, **opts):
        if shape is not None:
            self.effects.append((kind, shape, int(delay), int(dur), opts))

    def attach(self, slide):
        if not self.effects:
            return
        ids = itertools.count(5)
        pars, blds = [], []
        for kind, shape, delay, dur, o in self.effects:
            pars.append(self._effect(kind, shape.shape_id, delay, dur, o, ids,
                                     has_text=shape.has_text_frame if hasattr(shape, "has_text_frame") else False))
            if kind in ("fade", "rise", "letters", "zoom", "wipe") and getattr(shape, "has_text_frame", False) \
                    and shape.text_frame.text.strip():
                anim_bg = "" if kind == "letters" or o.get("by_word") else ' animBg="1"'
                blds.append(f'<p:bldP spid="{shape.shape_id}" grpId="0"{anim_bg}/>')
        bld = f'<p:bldLst>{"".join(dict.fromkeys(blds))}</p:bldLst>' if blds else ""
        xml = (
            f'<p:timing {NS_DECL}><p:tnLst><p:par>'
            '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            '<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            + "".join(pars) +
            '</p:childTnLst></p:cTn></p:par>'
            '</p:childTnLst></p:cTn></p:par>'
            '</p:childTnLst></p:cTn>'
            '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
            '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>' + bld + '</p:timing>')
        old = slide._element.find(qn("p:timing"))
        if old is not None:
            slide._element.remove(old)
        el = etree.fromstring(xml)
        ext = slide._element.find(qn("p:extLst"))
        if ext is not None:
            ext.addprevious(el)
        else:
            slide._element.append(el)

    @staticmethod
    def _tgt(spid):
        return f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'

    def _set_visible(self, spid, ids):
        return (f'<p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold"><p:stCondLst>'
                f'<p:cond delay="0"/></p:stCondLst></p:cTn>{self._tgt(spid)}'
                '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                '<p:to><p:strVal val="visible"/></p:to></p:set>')

    def _fade(self, spid, dur, ids):
        return (f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{next(ids)}" dur="{dur}"/>'
                f'{self._tgt(spid)}</p:cBhvr></p:animEffect>')

    def _anim(self, spid, attr, frm, to, dur, ids):
        return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
                f'<p:cTn id="{next(ids)}" dur="{dur}" decel="100000" fill="hold"/>{self._tgt(spid)}'
                f'<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>'
                f'<p:tav tm="0"><p:val><p:strVal val="{frm}"/></p:val></p:tav>'
                f'<p:tav tm="100000"><p:val><p:strVal val="{to}"/></p:val></p:tav></p:tavLst></p:anim>')

    def _effect(self, kind, spid, delay, dur, o, ids, has_text=False):
        cls, pid, sub = self.PRESETS[kind]
        eid = next(ids)
        extra, iterate = "", ""
        if kind == "letters" or o.get("by_word"):
            kind_ = "lt" if kind == "letters" else "wd"
            iterate = f'<p:iterate type="{kind_}"><p:tmPct val="{o.get("pct", 6000 if kind_ == "lt" else 12000)}"/></p:iterate>'
        if kind == "float":
            extra = ' accel="50000" decel="50000" repeatCount="indefinite" autoRev="1"'
            body = (f'<p:animMotion origin="layout" path="M 0 0 L 0 {o.get("dy", -0.02)} E" pathEditMode="relative" '
                    f'ptsTypes=""><p:cBhvr><p:cTn id="{next(ids)}" dur="{dur}" fill="hold"/>{self._tgt(spid)}'
                    '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
                    '</p:cBhvr></p:animMotion>')
        elif kind == "sway":
            extra = ' accel="50000" decel="50000" repeatCount="indefinite" autoRev="1"'
            body = (f'<p:animRot by="{int(o.get("deg", 2) * 60000)}"><p:cBhvr><p:cTn id="{next(ids)}" dur="{dur}" '
                    f'fill="hold"/>{self._tgt(spid)}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>'
                    '</p:cBhvr></p:animRot>')
        else:
            body = self._set_visible(spid, ids)
            if kind == "wipe":
                body += (f'<p:animEffect transition="in" filter="wipe(left)"><p:cBhvr>'
                         f'<p:cTn id="{next(ids)}" dur="{dur}"/>{self._tgt(spid)}</p:cBhvr></p:animEffect>')
            else:
                body += self._fade(spid, dur, ids)
            if kind in ("rise", "letters"):
                body += self._anim(spid, "ppt_y", f"#ppt_y+{o.get('dy', 0.06)}", "#ppt_y", dur, ids)
            if kind == "zoom":
                body += self._anim(spid, "ppt_w", "#ppt_w*0.55", "#ppt_w", dur, ids)
                body += self._anim(spid, "ppt_h", "#ppt_h*0.55", "#ppt_h", dur, ids)
        return (f'<p:par><p:cTn id="{eid}" presetID="{pid}" presetClass="{cls}" presetSubtype="{sub}"{extra} '
                f'fill="hold" grpId="0" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
                f'{iterate}<p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", nargs="?")
    ap.add_argument("out", nargs="?")
    ap.add_argument("--theme", help="override the theme named in the spec")
    ap.add_argument("--list-themes", action="store_true")
    args = ap.parse_args()
    themes = load_themes()
    if args.list_themes:
        for k, v in themes.items():
            print(f"{k:10s} {'dark ' if v['dark'] else 'light'}  {v['mood']}")
        return
    if not (args.spec and args.out):
        ap.error("spec and out are required")
    with open(args.spec) as f:
        spec = json.load(f)
    spec["_dir"] = os.path.dirname(os.path.abspath(args.spec))
    name = args.theme or spec.get("theme", "midnight")
    if name not in themes:
        sys.exit(f"unknown theme '{name}'. Available: {', '.join(themes)}")
    theme = dict(themes[name], **spec.get("theme_overrides", {}))
    prs = Deck(spec, theme).build()
    prs.core_properties.title = spec.get("title", "")
    prs.save(args.out)
    print(f"wrote {args.out}: {len(spec['slides'])} slides, theme '{name}'")


if __name__ == "__main__":
    main()
