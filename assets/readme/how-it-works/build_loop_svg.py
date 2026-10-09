#!/usr/bin/env python3
"""Build loop-light.svg and loop-dark.svg: the pptx-ts "how it works" loop.

Text is converted to outlines with fontTools (an <img> cannot load web fonts).
Each glyph outline is stored once in <defs> and placed with <use>, which keeps
the files small. Run: python3 build_loop_svg.py  (writes next to this script).
"""

import re
from pathlib import Path

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

HERE = Path(__file__).resolve().parent
FONT_DIR = Path("/usr/share/fonts/noto")
FONTS = {
    "sans": FONT_DIR / "NotoSans-Regular.ttf",
    "sans-sb": FONT_DIR / "NotoSans-SemiBold.ttf",
    "mono": FONT_DIR / "NotoSansMono-Regular.ttf",
    "mono-sb": FONT_DIR / "NotoSansMono-SemiBold.ttf",
}

W, H = 1200, 432

THEMES = {
    "light": {
        "bg": "#ffffff",
        "text": "#1f2328",
        "muted": "#59636e",
        "faint": "#818b98",
        "card": "#f6f8fa",
        "border": "#d1d9e0",
        "accent": "#b4451f",
        "accent_fill": "#fcf1ec",
        "accent_border": "#e7b9a5",
        "code": "#953800",
    },
    "dark": {
        "bg": "#0d1117",
        "text": "#e6edf3",
        "muted": "#9198a1",
        "faint": "#6e7681",
        "card": "#151b23",
        "border": "#3d444d",
        "accent": "#e9855c",
        "accent_fill": "#21160f",
        "accent_border": "#7a3f27",
        "code": "#f0a07e",
    },
}


def short_id(n):
    letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    out = ""
    while True:
        out = letters[n % 52] + out
        n = n // 52 - 1
        if n < 0:
            return out


def to_relative(d):
    """Rewrite an absolute M/L/H/V/Q/C/Z path as relative commands (smaller)."""
    tokens = re.findall(r"[MLHVQCZ]|-?\d+", d)
    out, i, cx, cy, sx, sy = [], 0, 0, 0, 0, 0
    sizes = {"M": 2, "L": 2, "H": 1, "V": 1, "Q": 4, "C": 6, "Z": 0}
    cmd = None
    while i < len(tokens):
        tok = tokens[i]
        if tok in sizes:
            cmd = tok
            i += 1
            if cmd == "Z":
                out.append("z")
                cx, cy = sx, sy
                continue
        n = sizes[cmd]
        args = [int(v) for v in tokens[i:i + n]]
        i += n
        if cmd == "H":
            rel = [args[0] - cx]
            cx = args[0]
        elif cmd == "V":
            rel = [args[0] - cy]
            cy = args[0]
        else:
            rel = []
            for k in range(0, n, 2):
                rel += [args[k] - cx, args[k + 1] - cy]
            cx, cy = args[-2], args[-1]
            if cmd == "M":
                sx, sy = cx, cy
        body = ""
        for v in rel:
            body += str(v) if (not body or v < 0) else " " + str(v)
        out.append(cmd.lower() + body)
        if cmd == "M":
            cmd = "L"  # pairs after a moveto are implicit linetos
    return "".join(out)


class Glyphs:
    """Turns strings into <use> runs over a shared glyph table."""

    def __init__(self):
        self.fonts = {k: TTFont(p) for k, p in FONTS.items()}
        self.defs = {}  # id -> path d
        self.ids = {}  # (font key, glyph name) -> short id

    def _glyph(self, key, ch):
        font = self.fonts[key]
        cmap = font.getBestCmap()
        name = cmap.get(ord(ch))
        if name is None:
            raise ValueError(f"{key} has no glyph for {ch!r}")
        adv = font["hmtx"][name][0]
        if (key, name) not in self.ids:
            self.ids[(key, name)] = short_id(len(self.ids))
        gid = self.ids[(key, name)]
        if gid not in self.defs:
            gs = font.getGlyphSet()
            pen = SVGPathPen(gs, ntos=lambda v: str(round(v)))
            # flip y here so <use> needs no transform per glyph
            gs[name].draw(TransformPen(pen, (1, 0, 0, -1, 0, 0)))
            self.defs[gid] = to_relative(pen.getCommands())
        return gid, adv

    def width(self, s, key, size):
        upm = self.fonts[key]["head"].unitsPerEm
        return sum(self._glyph(key, c)[1] for c in s) * size / upm

    def text(self, s, x, y, size, key="sans", fill="#000", anchor="start", spacing=0):
        upm = self.fonts[key]["head"].unitsPerEm
        scale = size / upm
        track = spacing * upm / size  # extra letter spacing, px -> font units
        total = self.width(s, key, size) + spacing * max(len(s) - 1, 0)
        if anchor == "middle":
            x -= total / 2
        elif anchor == "end":
            x -= total
        uses, pen_x = [], 0.0
        for c in s:
            gid, adv = self._glyph(key, c)
            if c != " " and self.defs[gid]:
                uses.append(f'<use href="#{gid}" x="{round(pen_x)}"/>')
            pen_x += adv + track
        return (
            f'<g fill="{fill}" transform="translate({x:.1f} {y:.1f}) scale({scale:.5f})">'
            + "".join(uses)
            + "</g>"
        )

    def defs_xml(self):
        return "".join(f'<path id="{k}" d="{d}"/>' for k, d in self.defs.items() if d)


def build(theme_name):
    t = THEMES[theme_name]
    g = Glyphs()
    out = []
    add = out.append

    def card(x, y, w, h, accent=False, dashed=False):
        fill = t["accent_fill"] if accent else t["card"]
        stroke = t["accent"] if accent else t["border"]
        sw = 1.6 if accent else 1.2
        dash = ' stroke-dasharray="5 4"' if dashed else ""
        add(
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="12" fill="{fill}" '
            f'stroke="{stroke}" stroke-width="{sw}"{dash}/>'
        )

    def arrow(d, color, dashed=False, marker="_m"):
        dash = ' stroke-dasharray="5 5"' if dashed else ""
        add(
            f'<path d="{d}" fill="none" stroke="{color}" stroke-width="1.8" '
            f'stroke-linecap="round" stroke-linejoin="round"{dash} marker-end="url(#{marker})"/>'
        )

    def label_pill(s, cx, cy, color):
        w = g.width(s, "sans-sb", 12) + 14
        add(f'<rect x="{cx - w / 2:.1f}" y="{cy - 10}" width="{w:.1f}" height="20" rx="10" fill="{t["bg"]}"/>')
        add(g.text(s, cx, cy + 4.2, 12, "sans-sb", color, "middle"))

    # Section labels
    add(g.text("WRITE", 40, 36, 11.5, "sans-sb", t["faint"], spacing=1.6))
    add(g.text("READ BACK", 40, 262, 11.5, "sans-sb", t["faint"], spacing=1.6))

    # Row 1 geometry
    y1, h1 = 54, 136
    A = (40, 220)    # x, w
    B = (315, 240)
    C = (610, 220)
    D = (950, 210)
    mid1 = y1 + 52

    # A: your code
    card(A[0], y1, A[1], h1)
    add(g.text("Your TypeScript", A[0] + 20, y1 + 32, 17, "sans-sb", t["text"]))
    for i, line in enumerate(["const pptx = new TsPptx()", "pptx.addSlide()", "  .addText(…)", "  .addChart(…)"]):
        add(g.text(line, A[0] + 20, y1 + 62 + i * 19, 12.5, "mono", t["muted"]))

    # B: pptx-ts
    card(B[0], y1, B[1], h1, accent=True)
    add(g.text("pptx-ts", B[0] + 20, y1 + 32, 18, "mono-sb", t["accent"]))
    add(g.text("builds every part in memory", B[0] + 20, y1 + 56, 13, "sans", t["muted"]))
    rows = [("Node", "writeFile() to disk"), ("Browser", "writeFile() downloads"), ("Anywhere", "toBytes()")]
    for i, (k, v) in enumerate(rows):
        yy = y1 + 82 + i * 19
        add(g.text(k, B[0] + 20, yy, 12.5, "sans-sb", t["text"]))
        add(g.text(v, B[0] + 92, yy, 12.5, "sans", t["muted"]))

    # C: the .pptx package
    card(C[0], y1, C[1], h1, dashed=True)
    add(g.text(".pptx", C[0] + 20, y1 + 32, 18, "mono-sb", t["text"]))
    add(g.text("a zip of XML parts", C[0] + 84, y1 + 31, 13, "sans", t["muted"]))
    parts = ["[Content_Types].xml", "ppt/presentation.xml", "ppt/slides/slide1.xml", "ppt/charts/chart1.xml"]
    for i, p in enumerate(parts):
        yy = y1 + 58 + i * 19
        add(f'<rect x="{C[0] + 20}" y="{yy - 9}" width="8" height="10" rx="1.5" fill="none" stroke="{t["faint"]}" stroke-width="1.1"/>')
        add(g.text(p, C[0] + 36, yy, 12, "mono", t["muted"]))

    # D: office apps
    card(D[0], y1, D[1], h1)
    add(g.text("PowerPoint", D[0] + 20, y1 + 32, 17, "sans-sb", t["text"]))
    add(g.text("opens it cleanly", D[0] + 20, y1 + 56, 13, "sans", t["muted"]))
    for i, line in enumerate(["Keynote, Google Slides and", "LibreOffice Impress", "import it too"]):
        add(g.text(line, D[0] + 20, y1 + 84 + i * 19, 12.5, "sans", t["faint"]))

    # Row 1 arrows
    arrow(f"M{A[0] + A[1] + 6} {mid1} H{B[0] - 8}", t["muted"])
    arrow(f"M{B[0] + B[1] + 6} {mid1} H{C[0] - 8}", t["accent"], marker="_a")
    arrow(f"M{C[0] + C[1] + 6} {mid1} H{D[0] - 8}", t["muted"])
    label_pill("opens in", (C[0] + C[1] + D[0]) / 2, mid1 - 18, t["muted"])
    back_y = y1 + 104
    arrow(f"M{D[0] - 6} {back_y} H{C[0] + C[1] + 8}", t["faint"], dashed=True)
    label_pill("or made by hand", (C[0] + C[1] + D[0]) / 2, back_y + 18, t["faint"])

    # Row 2 geometry
    y2, h2 = 286, 116
    S = (315, 240)   # script, under B
    R = (610, 220)   # read, under C
    I = (950, 210)   # inspect, under D
    mid2 = y2 + 58

    def lib_card(box, name, call, lines):
        x, w = box
        card(x, y2, w, h2, accent=True)
        add(g.text(name, x + 20, y2 + 30, 16, "mono-sb", t["accent"]))
        add(g.text(call, x + 20, y2 + 54, 12, "mono", t["code"]))
        for i, line in enumerate(lines):
            add(g.text(line, x + 20, y2 + 78 + i * 18, 12.5, "sans", t["muted"]))

    lib_card(S, "pptx-ts/script", "readModelToIr, printScript",
             ["TypeScript that rebuilds the deck,", "with notes on what it could not"])
    lib_card(R, "pptx-ts/read", "Presentation.load(bytes)",
             ["edit text, shapes, notes; save()", "rewrites only what changed"])
    lib_card(I, "pptx-ts/inspect", "inspectPptx(bytes)",
             ["every element on every", "slide, with its box in inches"])

    # .pptx <-> read
    yb = y1 + h1
    arrow(f"M{C[0] + 70} {yb + 6} V{y2 - 8}", t["accent"], marker="_a")
    label_pill("load", C[0] + 70, (yb + y2) / 2, t["accent"])
    arrow(f"M{C[0] + 150} {y2 - 6} V{yb + 8}", t["accent"], marker="_a")
    label_pill("save", C[0] + 150, (yb + y2) / 2, t["accent"])

    # .pptx -> inspect
    arrow(f"M{C[0] + C[1] - 30} {yb + 6} V{yb + 26} Q{C[0] + C[1] - 30} {yb + 40} {C[0] + C[1] - 16} {yb + 40} "
          f"H{I[0] + I[1] / 2 - 14} Q{I[0] + I[1] / 2} {yb + 40} {I[0] + I[1] / 2} {yb + 54} V{y2 - 8}", t["accent"], marker="_a")

    # read -> script
    arrow(f"M{R[0] - 6} {mid2} H{S[0] + S[1] + 8}", t["accent"], marker="_a")

    # script -> your code
    ax = A[0] + A[1] / 2
    arrow(f"M{S[0] - 6} {mid2} H{ax + 14} Q{ax} {mid2} {ax} {mid2 - 14} V{yb + 8}", t["accent"], marker="_a")
    label_pill("deck.ts", (S[0] + ax) / 2 + 6, mid2, t["accent"])

    markers = (
        f'<marker id="_m" viewBox="0 0 10 10" refX="7" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">'
        f'<path d="M1 1L8 5L1 9z" fill="{t["muted"]}"/></marker>'
        f'<marker id="_a" viewBox="0 0 10 10" refX="7" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">'
        f'<path d="M1 1L8 5L1 9z" fill="{t["accent"]}"/></marker>'
    )

    title = "How pptx-ts works: write a deck from TypeScript, read it back, inspect it, or turn it into code"
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}" role="img" '
        f'aria-label="{title}"><title>{title}</title>'
        f"<defs>{markers}{g.defs_xml()}</defs>"
        f'<rect width="{W}" height="{H}" fill="{t["bg"]}"/>'
        + "".join(out)
        + "</svg>\n"
    )
    return svg


def main():
    for name in THEMES:
        path = HERE / f"loop-{name}.svg"
        path.write_text(build(name))
        print(f"{path.name}: {path.stat().st_size} bytes")


if __name__ == "__main__":
    main()
