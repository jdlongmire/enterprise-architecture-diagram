"""archdiagram: spec-driven, fully editable PowerPoint architecture diagrams.

Engine used by make_diagram.py (call that instead). Direct use: python3 archdiagram.py spec.json out.pptx
Spec:   {"slides":[{rows...}]}  or a single slide {"rows":[...]}.  See SKILL.md for the block vocabulary.
Every card, label, arrow and icon is a native object; logical units are grouped and named.
Icons are embedded as SVG with a PNG fallback (crisp, recolorable, Convert-to-Shape capable).
"""
import json, os, subprocess, sys, tempfile, shutil, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from lxml import etree

ICONS = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons.json")))

THEME = {
    "font": "Calibri",
    "ink": "2A3546", "ink2": "4A5568", "ink3": "6B7686",
    "accent": "2F6FD6", "accent2": "2A63C4", "on_accent": "FFFFFF", "on_accent_sub": "EAF1FC",
    "canvas": None, "panel": "FFFFFF", "panel_line": "DDE3EC",
    "sub_fill": "F8FAFC", "sub_line": "E4E9F0", "bar_fill": "F3F8FE", "bar_line": "CFE0F6", "bar_ink": "2A4F8C",
    "arrow": "6D8FC9", "step_fill": "FBFCFE",
    # hue: [tint, line, icon]
    "hues": {
        "blue": ["EAF2FC", "CFE0F6", "2F6FD6"], "green": ["EBF7EF", "CFE9D7", "2F8A55"],
        "violet": ["F2EEFC", "DDD3F5", "6B4FC7"], "amber": ["FDF4E4", "F4DFB6", "C27A12"],
        "rose": ["FCECEE", "F4CFD4", "C43D50"], "teal": ["E8F6F6", "C9E7E7", "1F8A8A"],
        "slate": ["F1F3F6", "DCE1E8", "56637A"],
    },
    "hue_order": ["blue", "green", "violet", "amber", "rose", "teal"],
}

LH = 1.22 / 72  # inches per point of font size, per line
PAD_X, PAD_Y, GAP = 0.2, 0.14, 0.07


# Calibri advance widths (em) for accurate wrap prediction; unknown chars use 0.5
_CW = dict(zip("abcdefghijklmnopqrstuvwxyz", [.479,.525,.423,.525,.498,.305,.471,.525,.229,.267,.455,.229,.799,.525,.527,.525,.525,.349,.391,.335,.525,.452,.715,.433,.453,.395]))
_CW.update(zip("ABCDEFGHIJKLMNOPQRSTUVWXYZ", [.579,.544,.533,.615,.488,.459,.631,.623,.252,.319,.520,.420,.855,.646,.662,.517,.673,.543,.459,.487,.642,.567,.890,.519,.487,.468]))
_CW.update({c: .507 for c in "0123456789"}); _CW.update({" ": .226, ".": .252, ",": .250, "-": .306, "&": .682, "/": .386, "(": .303, ")": .303, ":": .268, "'": .221, "?": .463, "+": .498, "·": .268})


# relative width of common fonts vs Calibri (set from theme["font"], or override with theme["font_width"])
FONT_WIDTH = {"calibri": 1.0, "carlito": 1.0, "arial": 1.12, "helvetica": 1.12, "liberation sans": 1.12, "segoe ui": 1.1,
              "aptos": 1.08, "verdana": 1.3, "tahoma": 1.1, "trebuchet ms": 1.1, "cambria": 1.06, "georgia": 1.15, "times new roman": 0.98}
_FW = [1.0]
_LONG = set()  # words predicted to break mid-word


def _tbl(lower, upper, extra):
    t = dict(zip("abcdefghijklmnopqrstuvwxyz", [v / 1000 for v in lower]))
    t.update(zip("ABCDEFGHIJKLMNOPQRSTUVWXYZ", [v / 1000 for v in upper]))
    t.update({c: extra[0] / 1000 for c in "0123456789"})
    t.update(zip(" .,-&/():'?+·", [v / 1000 for v in extra[1:]]))
    return t
_UP = [667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611]
_ARIAL = _tbl([556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500], _UP,
              [556,278,278,278,333,667,278,333,333,278,191,556,584,278])
_ARIAL_B = _tbl([556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500],
                [722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611],
                [556,278,278,278,333,722,278,333,333,333,238,611,584,278])
_METRIC = ["calibri"]


def text_w(text, size, bold=False):
    """Rendered width in inches. Exact tables for Calibri and Arial-family fonts; others scale Calibri."""
    if _METRIC[0] == "arial":
        t = _ARIAL_B if bold else _ARIAL
        return sum(t.get(ch, 0.556) for ch in str(text)) * size / 72
    return sum(_CW.get(ch, 0.5) for ch in str(text)) * size / 72 * (1.05 if bold else 1.0) * _FW[0]


def est_lines(text, size, width, bold=False):
    if not text:
        return 0
    avail = width * 0.97  # small safety margin
    sp = text_w(" ", size, bold)
    lines = 0
    for para in str(text).split("\n"):
        lines += 1; cur = 0.0
        for w in para.split():
            ww = text_w(w, size, bold)
            if ww > avail:
                _LONG.add(w); lines += int(ww / avail)  # will break mid-word
            if cur == 0: cur = ww
            elif cur + sp + ww <= avail: cur += sp + ww
            else: lines += 1; cur = ww
    return lines


def th(text, size, width, bold=False):
    return est_lines(text, size, width, bold) * size * LH


class Diagram:
    def __init__(self, theme=None, width=13.333, height=7.5):
        self.T = copy.deepcopy(THEME)
        for k, v in (theme or {}).items():
            if k == "hues": self.T["hues"].update(v)
            else: self.T[k] = v
        _FW[0] = self.T.get("font_width") or FONT_WIDTH.get(self.T["font"].lower(), 1.1)
        _METRIC[0] = "arial" if self.T["font"].lower() in ("arial", "helvetica", "liberation sans", "arimo") and not self.T.get("font_width") else "calibri"
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(width), Inches(height)
        self.W, self.H = width, height
        self.pending, self.warnings, self._svg_n = [], [], 0
        self.tmp = tempfile.mkdtemp(prefix="archdiagram_")
        self.s = 1.0  # font scale
        self.sp = 1.0  # spacing scale
        self.issues = []  # structured problems: {severity, slide, path, problem, fix}
        self.si = -1; self.bpath = ""
        self.min_scale = 1.0

    # ---------- quality checks ----------
    def issue(self, severity, path, problem, fix):
        self.issues.append({"severity": severity, "slide": self.si + 1, "path": path, "problem": problem, "fix": fix})

    def max_chars(self, size, width, lines, bold=False):
        avg = text_w("abcdefghijklmnopqrstuvwxyz eeee aaaa", size, bold) / 36
        return max(4, int(width * 0.97 * lines / avg) - 2 * lines)

    def check(self, path, what, text, size, width, max_lines, bold=False, severity="error"):
        if not text: return
        n = est_lines(text, size, width, bold)
        if n > max_lines:
            lim = self.max_chars(size, width, max_lines, bold)
            self.issue(severity, path, f'{what} "{text}" wraps to {n} lines (limit {max_lines})',
                       f"shorten it to at most {lim} characters" + ("" if len(str(text).split()) > 1 else " or use a shorter word"))

    # ---------- primitives ----------
    def fs(self, size): return size * self.s
    def rgb(self, h): return RGBColor.from_string(h)

    def rrect(self, shapes, x, y, w, h, fill=None, line=None, radius=0.1, shadow=False, grad=None, name=None):
        s = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        st = s._element.find(qn("p:style"))
        if st is not None: s._element.remove(st)
        s.adjustments[0] = min(0.5, radius / max(0.01, min(w, h)))
        if grad:
            s.fill.gradient(); s.fill.gradient_angle = 90
            g = s.fill.gradient_stops
            g[0].color.rgb = self.rgb(grad[0]); g[0].position = 0
            g[1].color.rgb = self.rgb(grad[1]); g[1].position = 1.0
        else:
            s.fill.solid(); s.fill.fore_color.rgb = self.rgb(fill)
        if line: s.line.color.rgb = self.rgb(line); s.line.width = Pt(0.75)
        else: s.line.fill.background()
        if shadow:
            eff = etree.SubElement(s._element.spPr, qn("a:effectLst"))
            sh = etree.SubElement(eff, qn("a:outerShdw"), blurRad=str(Pt(10)), dist=str(Pt(2)), dir="5400000", algn="t", rotWithShape="0")
            c = etree.SubElement(sh, qn("a:srgbClr"), val="14284F"); etree.SubElement(c, qn("a:alpha"), val="8000")
        if name: s.name = name
        return s

    def text(self, shapes, x, y, w, h, paras, align="l", anchor="t", name=None):
        """paras: list of (text, size, bold, color[, space_before_pt])"""
        tb = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(max(h, 0.05)))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        first = True
        for p in paras:
            if not p[0]: continue
            para = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
            para.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
            if len(p) > 4: para.space_before = Pt(p[4])
            r = para.add_run(); r.text = str(p[0])
            f = r.font; f.name = self.T["font"]; f.size = Pt(p[1]); f.bold = p[2]; f.color.rgb = self.rgb(p[3])
        if name: tb.name = name
        return tb

    def icon(self, shapes, name, color, x, y, size):
        if not name: return
        if name not in ICONS:
            self.issue("warning", self.bpath or f"slides[{self.si}]", f"unknown icon '{name}' skipped", "use a name from reference/icons.md"); return
        self.pending.append((shapes, self.slide, name, color, x, y, size))

    def group(self, shapes, name):
        g = shapes.add_group_shape(); g.name = name[:80]
        return g.shapes

    def _flush_icons(self):
        jobs, files = [], {}
        for _, _, name, color, *_ in self.pending:
            key = (name, color)
            if key in files: continue
            svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24" fill="none" '
                   f'stroke="#{color}" color="#{color}" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">{ICONS[name]}</svg>')
            base = os.path.join(self.tmp, f"{name}_{color}")
            open(base + ".svg", "w").write(svg)
            files[key] = (base + ".svg", base + ".png"); jobs.append({"svg": svg, "out": base + ".png"})
        if jobs:
            how = rasterize(jobs, self.tmp)
            if how == "svg-only" and not getattr(self, "_svg_warned", False):
                self._svg_warned = True
                self.issue("warning", "", "no SVG rasterizer found; icons are SVG-only (fine in PowerPoint 365/2019+, invisible in older viewers and in the preview)",
                           "run: pip install cairosvg --break-system-packages")
        for shapes, slide, name, color, x, y, size in self.pending:
            svgp, pngp = files[(name, color)]
            pic = shapes.add_picture(pngp, Inches(x), Inches(y), Inches(size), Inches(size))
            pic.name = f"icon-{name}"
            self._embed_svg(slide, pic, svgp)
        self.pending = []

    def _embed_svg(self, slide, pic, svgp):
        self._svg_n += 1
        part = Part(PackURI(f"/ppt/media/archicon{self._svg_n}.svg"), "image/svg+xml", slide.part.package, open(svgp, "rb").read())
        rId = slide.part.relate_to(part, RT.IMAGE)
        blip = pic._element.find(".//" + qn("a:blip"))
        ext = etree.SubElement(etree.SubElement(blip, qn("a:extLst")), qn("a:ext"), uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
        ns = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
        el = etree.SubElement(ext, f"{{{ns}}}svgBlip", nsmap={"asvg": ns}); el.set(qn("r:embed"), rId)

    def hue(self, item, i):
        h = item.get("hue") or self.T["hue_order"][i % len(self.T["hue_order"])]
        return self.T["hues"][h]

    # ---------- blocks: measure(b, w) / draw(b, shapes, x, y, w, h) ----------
    def m_heading(self, b, w): return th(b["text"], self.fs(b.get("size", 15)), w, True)
    def d_heading(self, b, S, x, y, w, h, al):
        self.check(self.bpath + ".text", "Heading", b["text"], self.fs(b.get("size", 15)), w, 2, True)
        self.text(S, x, y, w, h, [(b["text"], self.fs(b.get("size", 15)), True, self.T["ink"])], b.get("align", al), name=f"Heading - {b['text']}")

    def m_subheading(self, b, w): return th(b["text"], self.fs(b.get("size", 10.5)), w, True)
    def d_subheading(self, b, S, x, y, w, h, al):
        self.text(S, x, y, w, h, [(b["text"], self.fs(b.get("size", 10.5)), True, self.T["ink"])], b.get("align", al), name=f"Subheading - {b['text']}")

    def m_text(self, b, w): return th(b["text"], self.fs(b.get("size", 9.5)), w)
    def d_text(self, b, S, x, y, w, h, al):
        self.text(S, x, y, w, h, [(b["text"], self.fs(b.get("size", 9.5)), False, self.T["ink2"])], b.get("align", al), name="Text")

    def _banner_sizes(self, b):
        return (self.fs(11.5), self.fs(8)) if b.get("size") == "sm" else (self.fs(15), self.fs(9.5))
    def m_banner(self, b, w):
        t, s = self._banner_sizes(b)
        return th(b["title"], t, w - 0.2, True) + (th(b.get("subtitle"), s, w - 0.2) + 0.02 if b.get("subtitle") else 0) + 0.15
    def d_banner(self, b, S, x, y, w, h, al):
        t, s = self._banner_sizes(b)
        self.check(self.bpath + ".title", "Banner title", b["title"], t, w - 0.2, 1, True)
        self.check(self.bpath + ".subtitle", "Banner subtitle", b.get("subtitle"), s, w - 0.2, 2)
        g = self.group(S, f"Banner - {b['title']}")
        self.rrect(g, x, y, w, h, radius=0.1, grad=(self.T["accent"], self.T["accent2"]))
        self.text(g, x + 0.1, y, w - 0.2, h, [(b["title"], t, True, self.T["on_accent"]), (b.get("subtitle"), s, False, self.T["on_accent_sub"])], "c", "m")

    def m_chip(self, b, w): return th(b["text"], self.fs(8), w - 0.1) + 0.1
    def d_chip(self, b, S, x, y, w, h, al):
        self.check(self.bpath + ".text", "Chip", b["text"], self.fs(8), w - 0.1, 2)
        g = self.group(S, f"Chip - {b['text'][:40]}")
        self.rrect(g, x, y, w, h, self.T["sub_fill"], self.T["sub_line"], radius=0.07)
        self.text(g, x + 0.05, y, w - 0.1, h, [(b["text"], self.fs(8), False, self.T["ink3"])], "c", "m")

    def m_bar(self, b, w): return th(b["text"], self.fs(10.5), w - 0.2, True) + 0.13
    def d_bar(self, b, S, x, y, w, h, al):
        self.check(self.bpath + ".text", "Bar", b["text"], self.fs(10.5), w - 0.2, 1, True, "warning")
        g = self.group(S, f"Bar - {b['text'][:40]}")
        self.rrect(g, x, y, w, h, self.T["bar_fill"], self.T["bar_line"], radius=0.08)
        self.text(g, x + 0.1, y, w - 0.2, h, [(b["text"], self.fs(10.5), True, self.T["bar_ink"])], "c", "m")

    def _band_inline_widths(self, b, w):
        return [text_w(it["label"], self.fs(9)) * 1.05 + 0.34 for it in b["items"]]
    def _band_layout(self, b, w):
        lay = b.get("layout", "stack")
        if lay == "inline" and sum(self._band_inline_widths(b, w)) > w - 0.2: lay = "stack"
        return lay
    def m_band(self, b, w):
        title = th(b.get("title"), self.fs(10), w - 0.2, True) + (0.05 if b.get("title") else 0)
        n = len(b["items"]); slot = w / n
        if self._band_layout(b, w) == "inline": return 0.08 + title + 0.2 + 0.08
        lab = max(th(it["label"], self.fs(9), slot - 0.08) for it in b["items"])
        return 0.08 + title + 0.22 + 0.04 + lab + 0.08
    def d_band(self, b, S, x, y, w, h, al):
        g = self.group(S, f"Band - {b.get('title') or 'items'}")
        self.rrect(g, x, y, w, h, self.T["sub_fill"], self.T["sub_line"], radius=0.1)
        ty = y + 0.08
        if b.get("title"):
            tH = th(b["title"], self.fs(10), w - 0.2, True)
            self.text(g, x + 0.1, ty, w - 0.2, tH, [(b["title"], self.fs(10), True, self.T["ink"])], "c"); ty += tH + 0.05
        items = b["items"]; col = self.T["accent"]
        if self._band_layout(b, w) == "inline":
            ws = self._band_inline_widths(b, w); gx = (w - sum(ws)) / (len(ws) + 1); cx = x + gx
            ry = y + h - 0.08 - 0.2
            for it, iw in zip(items, ws):
                ig = self.group(g, f"Item - {it['label']}")
                self.icon(ig, it.get("icon"), col, cx, ry + 0.015, 0.17)
                self.text(ig, cx + 0.24, ry, iw - 0.24, 0.2, [(it["label"], self.fs(9), False, self.T["ink2"])], "l", "m")
                cx += iw + gx
        else:
            slot = w / len(items); avail = y + h - 0.08 - ty
            content = 0.22 + 0.04 + max(th(it["label"], self.fs(9), slot - 0.08) for it in items)
            iy = ty + max(0, (avail - content) / 2)
            for i, it in enumerate(items):
                self.check(f"{self.bpath}.items[{i}].label", "Band label", it["label"], self.fs(9), slot - 0.08, 2)
                cx = x + slot * i + slot / 2
                ig = self.group(g, f"Item - {it['label']}")
                self.icon(ig, it.get("icon"), col, cx - 0.11, iy, 0.22)
                self.text(ig, cx - slot / 2 + 0.04, iy + 0.26, slot - 0.08, content - 0.26, [(it["label"], self.fs(9), False, self.T["ink2"])], "c")

    def _cards_grid(self, b, w):
        n = len(b["items"]); cols = b.get("columns") or min(n, 6 if b.get("layout", "vertical") == "vertical" else 5)
        rows = -(-n // cols); g = 0.08 if b.get("layout", "vertical") == "vertical" else 0.14
        cw = (w - g * (cols - 1)) / cols
        return cols, rows, g, cw
    def _card_h(self, it, cw, vertical):
        if vertical:
            iw = cw - 0.12
            return 0.1 + 0.32 + 0.06 + th(it["title"], self.fs(10), iw, True) + (th(it.get("desc"), self.fs(8.5), iw) + 0.03 if it.get("desc") else 0) + 0.1
        iw = cw - 0.56
        return max(0.46, th(it["title"], self.fs(10.5), iw, True) + th(it.get("desc"), self.fs(8.5), iw) + 0.16)
    def m_cards(self, b, w):
        cols, rows, g, cw = self._cards_grid(b, w); v = b.get("layout", "vertical") == "vertical"
        ch = max(self._card_h(it, cw, v) for it in b["items"])
        return rows * ch + (rows - 1) * g
    def d_cards(self, b, S, x, y, w, h, al):
        cols, rows, g, cw = self._cards_grid(b, w); v = b.get("layout", "vertical") == "vertical"
        ch = (h - (rows - 1) * g) / rows
        for i, it in enumerate(b["items"]):
            r, c = divmod(i, cols); cx = x + c * (cw + g); cy = y + r * (ch + g)
            tint, line, icol = self.hue(it, i)
            ip = f"{self.bpath}.items[{i}]"
            if v:
                self.check(ip + ".title", "Card title", it["title"], self.fs(10), cw - 0.12, 2, True)
                self.check(ip + ".desc", "Card desc", it.get("desc"), self.fs(8.5), cw - 0.12, 3)
            else:
                self.check(ip + ".title", "Card title", it["title"], self.fs(10.5), cw - 0.56, 1, True)
                self.check(ip + ".desc", "Card desc", it.get("desc"), self.fs(8.5), cw - 0.56, 2)
            grp = self.group(S, f"Card - {it['title']}")
            self.rrect(grp, cx, cy, cw, ch, tint, line, radius=0.1 if v else 0.08)
            if v:
                off = max(0, (ch - max(self._card_h(o, cw, True) for o in b["items"])) / 2)
                self.icon(grp, it.get("icon"), icol, cx + cw / 2 - 0.16, cy + 0.1 + off, 0.32)
                self.text(grp, cx + 0.06, cy + 0.48 + off, cw - 0.12, ch - 0.54 - off,
                          [(it["title"], self.fs(10), True, self.T["ink"]), (it.get("desc"), self.fs(8.5), False, self.T["ink2"], 1)], "c")
            else:
                self.icon(grp, it.get("icon"), icol, cx + 0.12, cy + ch / 2 - 0.14, 0.28)
                self.text(grp, cx + 0.5, cy + 0.06, cw - 0.56, ch - 0.12,
                          [(it["title"], self.fs(10.5), True, self.T["ink"]), (it.get("desc"), self.fs(8.5), False, self.T["ink2"])], "l", "m")

    def _li_h(self, it, w):
        iw = w - 0.32
        return th(it["title"], self.fs(10), iw, True) + th(it.get("desc"), self.fs(8.5), iw)
    def m_list(self, b, w):
        hs = [self._li_h(it, w) for it in b["items"]]
        return sum(hs) + 0.06 * self.sp * (len(hs) - 1)
    def d_list(self, b, S, x, y, w, h, al):
        hs = [self._li_h(it, w) for it in b["items"]]
        gap = (h - sum(hs)) / max(1, len(hs) - 1) if len(hs) > 1 else 0
        cy = y
        for i, (it, ih) in enumerate(zip(b["items"], hs)):
            col = self.hue(it, i)[2] if b.get("hued") else self.T["accent"]
            self.check(f"{self.bpath}.items[{i}].title", "List title", it["title"], self.fs(10), w - 0.32, 1, True)
            self.check(f"{self.bpath}.items[{i}].desc", "List desc", it.get("desc"), self.fs(8.5), w - 0.32, 2)
            grp = self.group(S, f"Item - {it['title']}")
            self.icon(grp, it.get("icon"), col, x, cy + 0.01, 0.2)
            self.text(grp, x + 0.32, cy, w - 0.32, ih, [(it["title"], self.fs(10), True, self.T["ink"]), (it.get("desc"), self.fs(8.5), False, self.T["ink3"])])
            cy += ih + gap

    def _flow_geom(self, b, w):
        n = len(b["items"]); aw = 0.28
        return n, aw, (w - aw * (n - 1)) / n
    def m_flow(self, b, w):
        n, aw, sw = self._flow_geom(b, w)
        if sw < 1.25 and not getattr(self, "_flow_warned", False):
            self._flow_warned = True
            self.issue("warning", self.bpath or f"slides[{self.si}]", f"flow steps are only {sw:.2f} in wide", "use at most 6 steps and put the flow in a full-width row")
        return max(0.46, max(th(it["title"], self.fs(10.5), sw - 0.48, True) + th(it.get("desc"), self.fs(8.5), sw - 0.48) for it in b["items"]) + 0.14)
    def d_flow(self, b, S, x, y, w, h, al):
        n, aw, sw = self._flow_geom(b, w)
        for i, it in enumerate(b["items"]):
            sx = x + i * (sw + aw)
            col = self.hue(it, i)[2] if b.get("hued") else self.T["accent"]
            self.check(f"{self.bpath}.items[{i}].title", "Flow step title", it["title"], self.fs(10.5), sw - 0.48, 2, True)
            self.check(f"{self.bpath}.items[{i}].desc", "Flow step desc", it.get("desc"), self.fs(8.5), sw - 0.48, 2)
            grp = self.group(S, f"Step - {it['title']}")
            self.rrect(grp, sx, y, sw, h, self.T["step_fill"], self.T["sub_line"], radius=0.08)
            self.icon(grp, it.get("icon"), col, sx + 0.1, y + h / 2 - 0.13, 0.26)
            self.text(grp, sx + 0.43, y + 0.05, sw - 0.48, h - 0.1, [(it["title"], self.fs(10.5), True, self.T["ink"]), (it.get("desc"), self.fs(8.5), False, self.T["ink3"])], "l", "m")
            if i < n - 1:
                self.icon(S, "arrow", self.T["arrow"], sx + sw + (aw - 0.18) / 2, y + h / 2 - 0.09, 0.18)

    def m_spacer(self, b, w): return b.get("h", 0.1)
    def d_spacer(self, *a): pass

    # ---------- panels / rows ----------
    def _gap_after(self, b): return (0.04 if b["type"] in ("heading", "subheading") else GAP) * self.sp
    def panel_natural(self, p, w):
        bl = p["blocks"]; iw = w - 2 * PAD_X
        hs = [getattr(self, "m_" + b["type"])(b, iw) for b in bl]
        gaps = [self._gap_after(b) for b in bl[:-1]]
        return 2 * PAD_Y * self.sp + sum(hs) + sum(gaps), hs, gaps

    def draw_panel(self, p, x, y, w, h, path=""):
        S = self.slide.shapes; bare = p.get("style") == "none"
        if not bare:
            self.rrect(S, x, y, w, h, self.T["panel"], self.T["panel_line"], radius=0.18, shadow=True,
                       name=f"Panel - {next((b.get('text') for b in p['blocks'] if b['type']=='heading'), 'panel')}")
        nat, hs, gaps = self.panel_natural(p, w)
        spare = max(0, h - nat); bl = p["blocks"]
        if p.get("distribute", True) and spare > 0.01:
            # 1) lists open up (capped per gap)  2) cards/flows grow (capped)  3) block gaps widen (capped)
            lists = [i for i, b in enumerate(bl) if b["type"] == "list"]
            for i in lists:
                add = min(spare / len(lists), 0.5 * max(1, len(bl[i]["items"]) - 1)); hs[i] += add; spare -= add
            grow = [i for i, b in enumerate(bl) if b["type"] in ("cards", "flow")]
            for i in grow:
                add = min(spare / len(grow), hs[i] * 0.3); hs[i] += add; spare -= add
            if gaps:
                per = min(spare / len(gaps), 0.22); gaps = [g + per for g in gaps]; spare -= per * len(gaps)
        # leftover: keep the leading heading (+ text) pinned, center the rest
        pin = 0
        while pin < len(bl) and bl[pin]["type"] in ("heading", "subheading", "text") and pin < 2: pin += 1
        if pin == len(bl): pin = 0
        offset = [0.0] * len(bl)
        for i in range(pin, len(bl)): offset[i] = max(0, spare) / 2
        al = "c" if p.get("align") == "center" else "l"
        cy = y + PAD_Y * self.sp; ix = x + PAD_X; iw = w - 2 * PAD_X
        for i, b in enumerate(bl):
            self.bpath = f"{path}.blocks[{i}]"
            getattr(self, "d_" + b["type"])(b, S, ix, cy + offset[i], iw, hs[i], al)
            cy += hs[i] + (gaps[i] if i < len(gaps) else 0)

    def layout(self, spec):
        M = spec.get("margin", 0.3); G = spec.get("gap", 0.15)
        aw = self.W - 2 * M; ah = self.H - 2 * M
        plan = []
        for r in spec["rows"]:
            ps = r["panels"]; tw = sum(p.get("w", 1) for p in ps)
            pws = [(aw - G * (len(ps) - 1)) * p.get("w", 1) / tw for p in ps]
            nat = max(self.panel_natural(p, pw)[0] for p, pw in zip(ps, pws))
            plan.append((r, ps, pws, max(nat, r.get("min_h", 0))))
        total = sum(n for *_, n in plan) + G * (len(plan) - 1)
        return plan, total, ah, M, G

    def add_slide(self, spec):
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6]); self.si = len(self.prs.slides) - 1
        cv = spec.get("canvas", self.T["canvas"])
        if cv:
            bg = self.slide.background.fill; bg.solid()
            bg.fore_color.rgb = self.rgb(cv)
        self.s = spec.get("font_scale", 1.0); self.sp = 1.0
        while True:  # fit: compress spacing first, then fonts (floor 85%)
            plan, total, ah, M, G = self.layout(spec)
            if total <= ah + 1e-6: break
            if self.sp > 0.55: self.sp = round(self.sp - 0.15, 2)
            elif self.s > 0.85: self.s = round(self.s - 0.03, 2)
            else: break
        _LONG.clear(); self.layout(spec)
        if _LONG: self.issue("error", f"slides[{self.si}]", f"words too long for their box: {', '.join(sorted(_LONG))}", "replace them with shorter words or abbreviations")
        _LONG.clear()
        if total > ah + 1e-6:
            self.issue("error", f"slides[{self.si}]", f"content is {total - ah:.1f} in too tall for the slide even at 85% text size",
                       "remove one block or list item, shorten long descriptions, or move the bottom row to a new slide")
        elif self.s < 0.91:
            self.issue("warning", f"slides[{self.si}]", f"text shrunk to {int(self.s*100)}% to fit",
                       "shorten the longest descriptions on this slide, or widen the most crowded panel (raise its w)")
        self.min_scale = min(self.min_scale, self.s)
        spare = max(0, ah - total)
        wt = lambda r, ps, n: r.get("weight", 0 if all(p.get("style") == "none" for p in ps) else n)
        tw = sum(wt(r, ps, n) for r, ps, _, n in plan) or 1
        rhs = [nat + min(spare * wt(r, ps, nat) / tw, nat * spec.get("max_stretch", 0.25)) for r, ps, _, nat in plan]
        y = M + max(0, ah - sum(rhs) - G * (len(plan) - 1)) / 2  # sparse content: center on slide
        for ri, ((r, ps, pws, nat), rh) in enumerate(zip(plan, rhs)):
            x = M
            for pi, (p, pw) in enumerate(zip(ps, pws)):
                self.draw_panel(p, x, y, pw, rh, f"slides[{self.si}].rows[{ri}].panels[{pi}]"); x += pw + G
            y += rh + G
        if spec.get("notes"): self.slide.notes_slide.notes_text_frame.text = spec["notes"]
        self._flush_icons()

    # ---------- connected diagrams: nodes + glued connectors ----------
    def _node_text_h(self, n, w, icon):
        iw = w - (0.62 if icon else 0.24) - 0.07  # margins + rounded-corner text inset
        return th(n["title"], self.fs(10.5), iw, True) + th(n.get("desc"), self.fs(8.5), iw)

    def add_graph_slide(self, spec):
        from pptx.enum.shapes import MSO_CONNECTOR
        from pptx.enum.dml import MSO_LINE
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6]); self.si = len(self.prs.slides) - 1
        cv = spec.get("canvas", self.T["canvas"])
        if cv:
            bg = self.slide.background.fill; bg.solid(); bg.fore_color.rgb = self.rgb(cv)
        self.s = spec.get("font_scale", 1.0)
        S = self.slide.shapes; T = self.T
        M = spec.get("margin", 0.35); y0 = M
        if spec.get("title"):
            th_ = th(spec["title"], 22, self.W - 2 * M, True)
            self.text(S, M, y0, self.W - 2 * M, th_, [(spec["title"], 22, True, T["ink"])], spec.get("title_align", "l")[0], name="Title"); y0 += th_ + 0.04
        if spec.get("subtitle"):
            sh_ = th(spec["subtitle"], 11, self.W - 2 * M)
            self.text(S, M, y0, self.W - 2 * M, sh_, [(spec["subtitle"], 11, False, T["ink2"])], spec.get("title_align", "l")[0], name="Subtitle"); y0 += sh_
        y0 += 0.2
        nodes = spec["nodes"]; lanes = spec.get("lanes")
        cols = spec.get("cols") or max(n["col"] for n in nodes) + 1
        rows = len(lanes) if lanes else (spec.get("rows") or max(n["row"] for n in nodes) + 1)
        lane_w = 1.05 if lanes else 0
        ax, ay, aw, ah = M + lane_w, y0, self.W - 2 * M - lane_w, self.H - M - y0
        cw, ch = aw / cols, ah / rows
        nw = min(cw * spec.get("node_w", 0.82 if lanes else 0.78), spec.get("node_max_w", 2.6))
        boxes = [n for n in nodes if n.get("shape", "card") in ("card", "banner")]
        def fit():
            _LONG.clear()
            tl = max([est_lines(n["title"], self.fs(10.5), nw - (0.62 if n.get("icon") else 0.24) - 0.07, True) for n in boxes] or [1])
            for n in nodes:
                if n.get("shape") in ("pill", "decision"):
                    avail = nw * 0.8 - 0.45 if n.get("shape") == "pill" else min(nw * 1.05, 1.8) * 0.5 - 0.04  # inscribed text area
                    est_lines(n["title"], self.fs(10.5 if n.get("shape") == "pill" else 9.5), avail, True)
            if _LONG: tl = max(tl, 3)
            h = max(spec.get("node_h", 0.62), max([self._node_text_h(n, nw, n.get("icon")) + 0.2 for n in boxes] or [0]))
            return tl, h
        tl, nh = fit()
        while (tl > 2 or nh > ch * 0.8) and self.s > 0.85:
            self.s = round(self.s - 0.03, 2); tl, nh = fit()
        _LONG.clear()
        if nh > ch * 0.8: self.issue("error", f"slides[{self.si}]", f"boxes are too tall for their rows ({nh:.2f} in vs {ch:.2f} in rows)",
                                     "shorten node descriptions, or use fewer rows/lanes")
        if self.s < 0.91: self.issue("warning", f"slides[{self.si}]", f"box text shrunk to {int(self.s*100)}%", "shorten node titles and descriptions, or use fewer columns")
        self.min_scale = min(self.min_scale, self.s)
        cx = lambda c: ax + cw * (c + 0.5)
        cy = lambda r: ay + ch * (r + 0.5)
        # lanes
        if lanes:
            for i, name in enumerate(lanes):
                g = self.group(S, f"Lane - {name}")
                self.rrect(g, M, ay + i * ch + 0.04, self.W - 2 * M, ch - 0.08, T["panel"] if i % 2 == 0 else T["sub_fill"], T["panel_line"], radius=0.12)
                lsz = 11
                while lsz > 8 and max(text_w(w, lsz, True) for w in name.split()) > lane_w - 0.27: lsz -= 0.5
                self.text(g, M + 0.15, ay + i * ch, lane_w - 0.25, ch, [(name, lsz, True, T["ink"])], "l", "m")
        # zones (boundaries) behind nodes
        for z in spec.get("zones", []):
            (c0, c1), (r0, r1) = z["cols"], z["rows"]
            zx, zy = ax + cw * c0 + 0.08, ay + ch * r0 + 0.04
            zw, zh = cw * (c1 - c0 + 1) - 0.16, ch * (r1 - r0 + 1) - 0.08
            tint, line, icol = T["hues"][z.get("hue", "slate")]
            g = self.group(S, f"Zone - {z['title']}")
            r = self.rrect(g, zx, zy, zw, zh, tint, line, radius=0.16)
            r.line.dash_style = MSO_LINE.DASH; r.line.width = Pt(1)
            self.text(g, zx + 0.14, zy + 0.07, zw - 0.28, 0.2, [(z["title"].upper(), self.fs(8.5), True, icol)])
        # nodes
        shp = {}
        for i, n in enumerate(nodes):
            kind = n.get("shape", "card")
            npth = f"slides[{self.si}].nodes[{i}]"
            if kind in ("card", "banner"):
                tw_ = nw - (0.62 if n.get("icon") else 0.24) - 0.07
                self.check(npth + ".title", "Box title", n["title"], self.fs(10.5), tw_, 2, True)
                self.check(npth + ".desc", "Box desc", n.get("desc"), self.fs(8.5), tw_, 2)
            elif kind == "pill":
                self.check(npth + ".title", "Start/end title", n["title"], self.fs(10.5), nw * 0.8 - 0.45, 2, True)
            else:
                self.check(npth + ".title", "Decision text", n["title"], self.fs(9.5), min(nw * 1.05, 1.8) * 0.5 - 0.04, 2, True)
            w_ = {"decision": min(nw * 1.05, 1.8), "pill": nw * 0.8}.get(kind, nw) * n.get("w", 1)
            h_ = {"decision": max(nh, 0.95), "pill": min(nh, 0.5)}.get(kind, nh)
            x, y = cx(n["col"]) - w_ / 2, cy(n["row"]) - h_ / 2
            tint, line, icol = self.hue(n if (n.get("hue") or kind != "decision") else dict(n, hue="amber"), i)
            g = self.group(S, f"Node - {n['title']}")
            if kind == "decision":
                s = g.add_shape(MSO_SHAPE.DIAMOND, Inches(x), Inches(y), Inches(w_), Inches(h_))
            else:
                s = g.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w_), Inches(h_))
                s.adjustments[0] = 0.5 if kind == "pill" else min(0.5, 0.1 / min(w_, h_))
            st = s._element.find(qn("p:style"))
            if st is not None: s._element.remove(st)
            banner = kind == "banner"
            if banner:
                s.fill.gradient(); s.fill.gradient_angle = 90
                gs = s.fill.gradient_stops; gs[0].color.rgb = self.rgb(T["accent"]); gs[0].position = 0
                gs[1].color.rgb = self.rgb(T["accent2"]); gs[1].position = 1.0; s.line.fill.background()
            else:
                s.fill.solid(); s.fill.fore_color.rgb = self.rgb(tint if kind != "pill" else T["panel"])
                s.line.color.rgb = self.rgb(line if kind != "pill" else T["accent"]); s.line.width = Pt(1 if kind != "card" else 0.75)
            s.name = f"Box - {n['title']}"
            tf = s.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            has_icon = bool(n.get("icon")) and kind in ("card", "banner")
            pad = 0.02 if kind == "decision" else 0.12
            tf.margin_left = Inches(0.5 if has_icon else pad); tf.margin_right = Inches(pad)
            tf.margin_top = tf.margin_bottom = Inches(0.04)
            ink, ink2 = (T["on_accent"], T["on_accent_sub"]) if banner else (T["ink"], T["ink2"])
            align = PP_ALIGN.LEFT if has_icon else PP_ALIGN.CENTER
            for j, (txt, sz, b, col) in enumerate([(n["title"], 10.5 if kind != "decision" else 9.5, True, ink), (n.get("desc"), 8.5, False, ink2)]):
                if not txt: continue
                p = tf.paragraphs[0] if j == 0 else tf.add_paragraph(); p.alignment = align
                r = p.add_run(); r.text = txt; f = r.font; f.name = T["font"]; f.size = Pt(self.fs(sz)); f.bold = b; f.color.rgb = self.rgb(col)
            if has_icon:
                self.icon(g, n["icon"], T["on_accent"] if banner else icol, x + 0.13, y + h_ / 2 - 0.15, 0.3)
            shp[n["id"]] = (s, n)
        # edges: real connectors glued to connection sites (0 top, 1 left, 2 bottom, 3 right),
        # routed by trying candidate shapes and scoring box collisions and shared ports
        E = 914400
        def rect(sh): return (sh.left / E, sh.top / E, sh.width / E, sh.height / E)
        def port(sh, side):
            x, y, w, h = rect(sh)
            return {0: (x + w / 2, y), 1: (x, y + h / 2), 2: (x + w / 2, y + h), 3: (x + w, y + h / 2)}[side]
        def seg_hits(p, q, r):
            x, y, w, h = r; x += 0.03; y += 0.03; w -= 0.06; h -= 0.06
            (x1, y1), (x2, y2) = p, q
            if abs(y1 - y2) < 1e-6:  # horizontal
                return y < y1 < y + h and max(min(x1, x2), x) < min(max(x1, x2), x + w)
            if abs(x1 - x2) < 1e-6:  # vertical
                return x < x1 < x + w and max(min(y1, y2), y) < min(max(y1, y2), y + h)
            for t in [i / 20 for i in range(1, 20)]:  # diagonal: sample
                px, py = x1 + (x2 - x1) * t, y1 + (y2 - y1) * t
                if x < px < x + w and y < py < y + h: return True
            return False
        def crosses(p1, p2, q1, q2):
            """count perpendicular crossings / collinear overlaps between two axis-aligned segments"""
            h1, h2 = abs(p1[1] - p2[1]) < 1e-6, abs(q1[1] - q2[1]) < 1e-6
            v1, v2 = abs(p1[0] - p2[0]) < 1e-6, abs(q1[0] - q2[0]) < 1e-6
            if h1 and v2: (hp, hq), (vp, vq) = (p1, p2), (q1, q2)
            elif v1 and h2: (hp, hq), (vp, vq) = (q1, q2), (p1, p2)
            elif h1 and h2 and abs(p1[1] - q1[1]) < 0.02:
                return 1 if min(max(p1[0], p2[0]), max(q1[0], q2[0])) - max(min(p1[0], p2[0]), min(q1[0], q2[0])) > 0.08 else 0
            elif v1 and v2 and abs(p1[0] - q1[0]) < 0.02:
                return 1 if min(max(p1[1], p2[1]), max(q1[1], q2[1])) - max(min(p1[1], p2[1]), min(q1[1], q2[1])) > 0.08 else 0
            else: return 0
            x, y = vp[0], hp[1]
            return 1 if (min(hp[0], hq[0]) + 0.02 < x < max(hp[0], hq[0]) - 0.02 and min(vp[1], vq[1]) + 0.02 < y < max(vp[1], vq[1]) - 0.02) else 0
        routed = []      # (edge key, [segments])
        label_jobs = []
        used = {}
        for e in spec.get("edges", []):
            for k in ("from", "to"):
                if e[k] not in shp: raise SystemExit(f"edge refers to unknown node id '{e[k]}'")
        def flexible(e):
            na, nb = shp[e["from"]][1], shp[e["to"]][1]
            return not ("from_side" in e or "to_side" in e) and na["col"] != nb["col"] and na["row"] != nb["row"]
        eidx = {id(e): k for k, e in enumerate(spec.get("edges", []))}
        for e in sorted(spec.get("edges", []), key=flexible):
            epth = f"slides[{self.si}].edges[{eidx[id(e)]}]"
            (a, na), (b, nb) = shp[e["from"]], shp[e["to"]]
            dc, dr = nb["col"] - na["col"], nb["row"] - na["row"]
            hs_ = 3 if dc > 0 else 1; ht_ = 1 if dc > 0 else 3      # horizontal source / target sides
            vs_ = 2 if dr > 0 else 0; vt_ = 0 if dr > 0 else 2      # vertical source / target sides
            if "from_side" in e or "to_side" in e:
                side = {"top": 0, "left": 1, "bottom": 2, "right": 3}
                sa, sb = side[e.get("from_side", "right")], side[e.get("to_side", "left")]
                same_axis = (sa in (1, 3)) == (sb in (1, 3))
                cands = [("straight" if e.get("style") == "straight" else ((("z" if sa in (1, 3) else "zv")) if same_axis else "l"), sa, sb)]
            elif dc == 0:
                cands = [("straight", vs_, vt_)]
            elif dr == 0:
                cands = [("straight", hs_, ht_)]
            else:
                cands = [("z", hs_, ht_), ("l", hs_, vt_), ("l", vs_, ht_), ("zv", vs_, vt_)]
            others = [rect(s2) for k2, (s2, _) in shp.items() if k2 not in (e["from"], e["to"])]
            best = None
            for rank, (kind, sa, sb) in enumerate(cands):
                B, Q = port(a, sa), port(b, sb)
                if kind == "straight": pts = [B, Q]
                elif kind == "z": mx = (B[0] + Q[0]) / 2; pts = [B, (mx, B[1]), (mx, Q[1]), Q]
                elif kind == "zv": my = (B[1] + Q[1]) / 2; pts = [B, (B[0], my), (Q[0], my), Q]
                else: pts = [B, (Q[0], B[1]), Q] if sa in (1, 3) else [B, (B[0], Q[1]), Q]
                hits = sum(seg_hits(pts[i], pts[i + 1], r) for i in range(len(pts) - 1) for r in others)
                shared = used.get((e["from"], sa), 0) + used.get((e["to"], sb), 0)
                segs_ = list(zip(pts, pts[1:]))
                xing = sum(crosses(a1, a2, b1, b2) for _, osegs in routed for (b1, b2) in osegs for (a1, a2) in segs_)
                score = hits * 10 + shared * 4 + xing * 2.5 + rank
                if best is None or score < best[0]: best = (score, kind, sa, sb, pts, hits, shared)
            _, kind, sa, sb, pts, hits, shared_ = best
            routed.append((epth, list(zip(pts, pts[1:]))))
            if hits: self.issue("error", epth, f'line "{na["title"]}" -> "{nb["title"]}" passes through another box',
                                "move one of the two boxes so nothing sits between them (change col/row), or remove col/row from all nodes to use auto-layout")
            elif shared_: self.issue("warning", epth, f'line "{na["title"]}" -> "{nb["title"]}" shares a box side with another line',
                                          "acceptable for a deliberate fork or merge; otherwise move one box to a different row")
            used[(e["from"], sa)] = used.get((e["from"], sa), 0) + 1; used[(e["to"], sb)] = used.get((e["to"], sb), 0) + 1
            if e.get("style") == "curve": mso = MSO_CONNECTOR.CURVE
            elif kind == "straight": mso = MSO_CONNECTOR.STRAIGHT
            else: mso = MSO_CONNECTOR.ELBOW
            c = S.add_connector(mso, 0, 0, Inches(1), Inches(1))
            c.begin_connect(a, sa); c.end_connect(b, sb)
            if kind in ("l", "zv"): self._orient_connector(c, kind, vertical_first=(sa in (0, 2)))
            c.name = f"Connector - {na['title']} to {nb['title']}"
            c.line.color.rgb = self.rgb(e.get("color", T["arrow"] if not e.get("dashed") else T["ink3"]))
            c.line.width = Pt(e.get("width", 1.5))
            if e.get("dashed"): c.line.dash_style = MSO_LINE.DASH
            ln = c.line._get_or_add_ln()
            if e.get("arrow", "end") in ("end", "both"):
                etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
            if e.get("arrow") == "both":
                etree.SubElement(ln, qn("a:headEnd"), type="triangle", w="med", len="med")
            if e.get("label"): label_jobs.append((epth, e, pts))
        # labels last, so they can avoid every line and box
        node_rects = [rect(s2) for s2, _ in shp.values()]
        placed = []
        for epth, e, pts in label_jobs:
            lw = text_w(e["label"], self.fs(8)) + 0.16; lh = 0.22
            own = [sg for k, sg in routed if k == epth]
            other_segs = [sg for k, segs_ in routed if k != epth for sg in segs_]
            segs = sorted(zip(pts, pts[1:]), key=lambda sg: -(abs(sg[0][0] - sg[1][0]) + abs(sg[0][1] - sg[1][1])))
            cands = []
            for p1, p2 in segs:
                horiz = abs(p1[1] - p2[1]) < 1e-6; L = abs(p1[0] - p2[0]) + abs(p1[1] - p2[1])
                for t in (0.5, 0.35, 0.65, 0.2, 0.8):
                    mx, my = p1[0] + (p2[0] - p1[0]) * t, p1[1] + (p2[1] - p1[1]) * t
                    if horiz and L < lw + 0.12: cands.append((mx, my - 0.17)); break
                    cands.append((mx, my))
            def bad(cx_, cy_):
                r = (cx_ - lw / 2 - 0.02, cy_ - lh / 2 - 0.02, lw + 0.04, lh + 0.04)
                if any(seg_hits(a1, a2, (r[0] - 0.03, r[1] - 0.03, r[2] + 0.06, r[3] + 0.06)) for a1, a2 in other_segs): return True
                if any(not (r[0] + r[2] < nx + 0.02 or nx + nw2 - 0.02 < r[0] or r[1] + r[3] < ny + 0.02 or ny + nh2 - 0.02 < r[1]) for nx, ny, nw2, nh2 in node_rects): return True
                if any(not (r[0] + r[2] < q[0] or q[0] + q[2] < r[0] or r[1] + r[3] < q[1] or q[1] + q[3] < r[1]) for q in placed): return True
                return False
            pos = next(((x_, y_) for x_, y_ in cands if not bad(x_, y_)), None)
            if pos is None:
                pos = cands[0]
                self.issue("warning", epth + ".label", f'line label "{e["label"]}" touches another line or box', "shorten the label, or move a box to open more space around this line")
            mx, my = pos; placed.append((mx - lw / 2, my - lh / 2, lw, lh))
            seg_len = max(abs(a[0] - b[0]) + abs(a[1] - b[1]) for a, b in zip(pts, pts[1:]))
            if lw > max(seg_len, 0.4) + 0.5:
                self.issue("warning", epth + ".label", f'line label "{e["label"]}" is much longer than its line', "use a 1 to 2 word label")
            lb = self.rrect(S, mx - lw / 2, my - lh / 2, lw, lh, (spec.get("canvas", T["canvas"]) or T["panel"]) if not lanes else T["panel"], None, radius=0.11, name=f"Label - {e['label']}")
            tf = lb.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
            r = p.add_run(); r.text = e["label"]; f = r.font; f.name = T["font"]; f.size = Pt(self.fs(8)); f.color.rgb = self.rgb(T["ink3"])
        n_x = sum(crosses(a1, a2, b1, b2) for i, (_, sa_) in enumerate(routed) for _, sb_ in routed[i + 1:] for a1, a2 in sa_ for b1, b2 in sb_)
        if n_x > max(2, len(routed) // 4):
            self.issue("warning", f"slides[{self.si}]", f"{n_x} line crossings", "reorder nodes so connected boxes sit next to each other, or delete col/row to auto-layout")
        if spec.get("notes"): self.slide.notes_slide.notes_text_frame.text = spec["notes"]
        self._flush_icons()

    def _orient_connector(self, c, kind, vertical_first):
        """Rewrite connector geometry: 'l' = bentConnector2 (one bend), 'zv' = vertical bentConnector3.
        Vertical-first paths use a 90/270 degree rotation with optional vertical flip."""
        E = 914400
        bx, by, ex, ey = c.begin_x, c.begin_y, c.end_x, c.end_y
        sp = c._element.spPr; xfrm = sp.find(qn("a:xfrm")); geom = sp.find(qn("a:prstGeom"))
        geom.set("prst", "bentConnector2" if kind == "l" else "bentConnector3")
        av = geom.find(qn("a:avLst"))
        if av is not None:
            for g in list(av): av.remove(g)
        if not vertical_first: return  # python-pptx flips already describe a horizontal-first path
        dx, dy = ex - bx, ey - by
        W, H = abs(dx), abs(dy); cxm, cym = (bx + ex) / 2, (by + ey) / 2
        rot, flipv = {(False, True): (5400000, False), (True, True): (5400000, True),
                      (True, False): (16200000, False), (False, False): (16200000, True)}[(dx > 0, dy > 0)]
        for k in ("flipH", "flipV", "rot"):
            if k in xfrm.attrib: del xfrm.attrib[k]
        xfrm.set("rot", str(rot))
        if flipv: xfrm.set("flipV", "1")
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        off.set("x", str(int(cxm - H / 2))); off.set("y", str(int(cym - W / 2)))
        ext.set("cx", str(int(H))); ext.set("cy", str(int(W)))

    def save(self, path):
        self.prs.save(path); shutil.rmtree(self.tmp, ignore_errors=True)


_TRANSPARENT_PNG = bytes.fromhex("89504e470d0a1a0a0000000d4948445200000001000000010806000000"
                                 "1f15c4890000000d49444154789c63000100000500010d0a2db40000000049454e44ae426082")


def rasterize(jobs, tmp):
    """PNG fallbacks for the embedded SVG icons. Tries cairosvg, then node+sharp; otherwise writes
    transparent placeholders (PowerPoint 2016+/365 still shows the SVG; older viewers show no icon)."""
    try:
        import cairosvg
        for j in jobs: cairosvg.svg2png(bytestring=j["svg"].encode(), write_to=j["out"], output_width=256, output_height=256)
        return "cairosvg"
    except Exception:
        pass
    try:
        js = os.path.join(tmp, "r.js"); jf = os.path.join(tmp, "jobs.json")
        json.dump(jobs, open(jf, "w"))
        open(js, "w").write("const sharp=require('sharp'),fs=require('fs');(async()=>{for(const j of JSON.parse(fs.readFileSync(process.argv[2],'utf8')))"
                            "await sharp(Buffer.from(j.svg),{density:1200}).resize(256,256).png().toFile(j.out);})();")
        env = dict(os.environ); gm = subprocess.run(["npm", "root", "-g"], capture_output=True, text=True).stdout.strip()
        env["NODE_PATH"] = gm + os.pathsep + env.get("NODE_PATH", "")
        subprocess.run(["node", js, jf], check=True, env=env, capture_output=True, timeout=120)
        if all(os.path.exists(j["out"]) for j in jobs): return "sharp"
    except Exception:
        pass
    for j in jobs: open(j["out"], "wb").write(_TRANSPARENT_PNG)
    return "svg-only"


def build(spec, out, png=None):
    slides = spec["slides"] if "slides" in spec else [spec]
    d = Diagram(spec.get("theme"), spec.get("width", 13.333), spec.get("height", 7.5))
    for s in slides: (d.add_graph_slide if "nodes" in s else d.add_slide)(s)
    d.save(out)
    if png: export_png(out, png)
    return d


def export_png(pptx, png):
    outdir = tempfile.mkdtemp()
    env = dict(os.environ, HOME=outdir)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx], check=True, env=env, capture_output=True, timeout=180)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    base = os.path.splitext(png)[0]
    subprocess.run(["pdftoppm", "-png", "-r", "150", pdf, base], check=True)
    print("rendered", *sorted(f for f in os.listdir(os.path.dirname(os.path.abspath(png)) or ".") if f.startswith(os.path.basename(base)) and f.endswith(".png")))


if __name__ == "__main__":
    a = sys.argv[1:]
    if len(a) < 2: sys.exit(__doc__)
    png = a[a.index("--png") + 1] if "--png" in a else None
    d = build(json.load(open(a[0])), a[1], png)
    print("saved", a[1]); [print(i["severity"].upper(), i["path"], i["problem"], "->", i["fix"]) for i in d.issues]
