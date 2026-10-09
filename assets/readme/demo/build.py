#!/usr/bin/env python3
"""Build the README hero demo: demo-light.svg and demo-dark.svg.

Pipeline:
  1. `node report.mts` writes report.pptx with the real library.
  2. LibreOffice converts it to PDF, pdftoppm rasterises slide 1.
  3. The slide raster is embedded as JPEG; the code shown on the left is
     report.mts itself, so the animation shows exactly the program that ran.
  4. All text is converted to glyph outlines (fontTools), so no <text> remains.

Run:  python3 build.py            (regenerates the deck and the raster)
      python3 build.py --no-deck  (reuses slide.png)
Animation is CSS only. Every element's resting style is the final frame.
"""

import base64
import io
import os
import re
import subprocess
import sys

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.ttLib import TTFont
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = "/usr/share/fonts/TTF/CaskaydiaMonoNerdFontMono-Regular.ttf"
FONT_BOLD = "/usr/share/fonts/TTF/CaskaydiaMonoNerdFontMono-Bold.ttf"

# ---------------------------------------------------------------- deck


def build_deck():
    run = lambda *a: subprocess.run(a, cwd=HERE, check=True, capture_output=True)
    run("node", "report.mts")
    run("soffice", "--headless", "--convert-to", "pdf", "report.pptx")
    run("pdftoppm", "-png", "-f", "1", "-l", "1", "-singlefile", "-scale-to-x", "1152",
        "-scale-to-y", "648", "report.pdf", "slide")


# ---------------------------------------------------------------- layout

W, H = 1200, 640
FS = 14  # code font size, px
LH = 19  # code line height
CODE_X, CODE_Y, CODE_W = 20, 20, 524
CHROME = 38
GUTTER = 40
PAD = 14
RIGHT_X = 588
RIGHT_W = 592
TERM_Y, TERM_H = 20, 84
CHIP_Y = 140
SLIDE_W = RIGHT_W
SLIDE_H = SLIDE_W * 9 / 16
SLIDE_Y = 216

T = 23.0  # loop length, seconds

THEMES = {
    "light": dict(
        bg="#ffffff", panel="#f6f8fa", chrome="#eaeef2", border="#d0d7de", text="#1f2328",
        muted="#6e7781", accent="#b4451f", kw="#b4451f", str="#0a3069", num="#0550ae",
        fn="#8250df", prop="#1f2328", punct="#57606a", hl="#b4451f14", shadow="#1f232822",
        dots=("#d0d7de", "#d0d7de", "#d0d7de"),
    ),
    "dark": dict(
        bg="#0d1117", panel="#161b22", chrome="#1c2128", border="#30363d", text="#e6edf3",
        muted="#7d8590", accent="#e9855c", kw="#e9855c", str="#a5d6ff", num="#79c0ff",
        fn="#d2a8ff", prop="#e6edf3", punct="#9198a1", hl="#e9855c1f", shadow="#00000066",
        dots=("#30363d", "#30363d", "#30363d"),
    ),
}

# ---------------------------------------------------------------- glyphs


class Outliner:
    """Turns strings into <use> references to glyph outlines held in <defs>."""

    def __init__(self, path, prefix):
        self.font = TTFont(path)
        self.gs = self.font.getGlyphSet()
        self.cmap = self.font.getBestCmap()
        self.upm = self.font["head"].unitsPerEm
        self.prefix = prefix
        self.used = {}

    def adv(self, size):
        return self.font["hmtx"][self.cmap[ord("M")]][0] * size / self.upm

    def gid(self, ch):
        name = self.cmap[ord(ch)]
        if name not in self.used:
            pen = SVGPathPen(self.gs)
            self.gs[name].draw(pen)
            d = pen.getCommands()
            self.used[name] = (f"{self.prefix}{len(self.used)}", d)
        return self.used[name][0]

    def run(self, text, x, y, size, fill, extra=""):
        """One horizontal run of text in one colour, baseline at y."""
        s = size / self.upm
        aw = self.font["hmtx"][self.cmap[ord("M")]][0]
        uses = []
        for i, ch in enumerate(text):
            if ch == " ":
                continue
            gid = self.gid(ch)
            if not self.used[self.cmap[ord(ch)]][1]:
                continue
            uses.append(f'<use href="#{gid}" x="{i * aw}"/>')
        if not uses:
            return ""
        return (f'<g fill="{fill}" transform="translate({x:.2f} {y:.2f}) scale({s:.6f} {-s:.6f})"'
                f'{extra}>{"".join(uses)}</g>')

    def defs(self):
        return "".join(f'<path id="{g}" d="{d}"/>' for g, d in self.used.values() if d)


# ---------------------------------------------------------------- syntax

KEYWORDS = {"import", "from", "const", "new", "await", "true", "false", "return"}
TOKEN = re.compile(r'(?P<str>"[^"]*")|(?P<num>\b\d+(?:\.\d+)?\b)|(?P<id>[A-Za-z_$][\w$]*)'
                   r'|(?P<ws>\s+)|(?P<p>.)')


def tokens(line):
    out = []
    for m in TOKEN.finditer(line):
        kind, val = m.lastgroup, m.group()
        if kind == "id":
            rest = line[m.end():]
            if val in KEYWORDS:
                kind = "kw"
            elif rest.startswith("("):
                kind = "fn"
            elif rest.startswith(":") and not rest.startswith("::"):
                kind = "prop"
            elif val == "TsPptx":
                kind = "fn"
            else:
                kind = "text"
        elif kind == "p" or kind == "ws":
            kind = "punct"
        out.append((kind, val))
    return out


# ---------------------------------------------------------------- css helpers


def pct(t):
    return f"{max(0.0, min(100.0, t / T * 100)):.3f}%"


FADE_OUT = (T - 0.9, T - 0.15)  # everything dynamic fades before the loop restarts


class Css:
    def __init__(self):
        self.rules = []
        self.n = 0

    def anim(self, frames, rest, timing="linear"):
        """frames: list of (time_s, css, optional timing for the next segment)."""
        self.n += 1
        name = f"a{self.n}"
        body = []
        for f in frames:
            t, css = f[0], f[1]
            tf = f"animation-timing-function:{f[2]};" if len(f) > 2 else ""
            body.append(f"{pct(t)}{{{css.rstrip(';')};{tf}}}")
        self.rules.append(f"@keyframes {name}{{{''.join(body)}}}")
        self.rules.append(f".{name}{{{rest}animation:{name} {T}s {timing} infinite}}")
        return name

    def text(self):
        return "".join(self.rules)


# ---------------------------------------------------------------- scene


def detect_bars(img):
    """Columns of the accent fill in the slide raster: (x0, x1, top, base)."""
    px = img.load()
    w, h = img.size
    acc = (0xB4, 0x45, 0x1F)
    near = lambda c: sum(abs(a - b) for a, b in zip(c[:3], acc)) < 60
    cols = []
    for x in range(w):
        ys = [y for y in range(h) if near(px[x, y])]
        cols.append((min(ys), max(ys)) if len(ys) > 20 else None)
    bars, start = [], None
    for x in range(w + 1):
        c = cols[x] if x < w else None
        if c and start is None:
            start = x
        if not c and start is not None:
            seg = [cols[i] for i in range(start, x)]
            bars.append((start, x - 1, min(s[0] for s in seg), max(s[1] for s in seg)))
            start = None
    return [b for b in bars if b[1] - b[0] > 10]


def build(theme_name, code_lines, slide_png, pptx_bytes):
    th = THEMES[theme_name]
    mono = Outliner(FONT, "g")
    bold = Outliner(FONT_BOLD, "b")
    adv = mono.adv(FS)
    css = Css()
    parts = []
    P = parts.append

    # ---- code window
    P(f'<rect x="{CODE_X + .5}" y="{CODE_Y + .5}" width="{CODE_W - 1}" height="{H - 2 * CODE_Y - 1}"'
      f' rx="10" fill="{th["panel"]}" stroke="{th["border"]}"/>')
    P(f'<path d="M{CODE_X + .5} {CODE_Y + CHROME}V{CODE_Y + 10.5}a10 10 0 0 1 10 -10H{CODE_X + CODE_W - 10.5}'
      f'a10 10 0 0 1 10 10V{CODE_Y + CHROME}Z" fill="{th["chrome"]}" stroke="{th["border"]}"/>')
    for i, c in enumerate(th["dots"]):
        P(f'<circle cx="{CODE_X + 20 + i * 18}" cy="{CODE_Y + CHROME / 2}" r="5.5" fill="{c}"/>')
    P(mono.run("report.mts", CODE_X + CODE_W / 2 - 5 * adv * 12.5 / FS, CODE_Y + CHROME / 2 + 4.5,
               12.5, th["muted"]))

    # typing schedule
    t = 0.6
    per_char, per_line = 0.0085, 0.05
    sched = []
    for line in code_lines:
        n = len(line.rstrip())
        dur = n * per_char
        sched.append((t, t + dur))
        t += dur + (per_line if n else 0.12)
    typing_end = t

    top = CODE_Y + CHROME + PAD + FS
    text_x = CODE_X + GUTTER + 6
    write_idx = next(i for i, l in enumerate(code_lines) if "writeFile" in l)
    run_t = typing_end + 0.35

    code = []
    for i, line in enumerate(code_lines):
        y = top + i * LH
        t0, t1 = sched[i]
        # writeFile highlight: persistent at rest, appears when the program runs
        if i == write_idx:
            a = css.anim([(0, "opacity:0"), (run_t, "opacity:0"), (run_t + 0.25, "opacity:1"),
                          (FADE_OUT[0], "opacity:1"), (FADE_OUT[1], "opacity:0"), (T, "opacity:0")],
                         "")
            code.append(f'<g class="{a}"><rect x="{CODE_X + 1}" y="{y - FS - 2.5}" width="{CODE_W - 2}"'
                        f' height="{LH + 1}" fill="{th["hl"]}"/><rect x="{CODE_X + 1}" y="{y - FS - 2.5}"'
                        f' width="2.5" height="{LH + 1}" fill="{th["accent"]}"/></g>')
        # line number
        num = str(i + 1)
        a = css.anim([(0, "opacity:0"), (t0, "opacity:0", "step-end"), (t0 + 0.001, "opacity:1"),
                      (FADE_OUT[0], "opacity:1"), (FADE_OUT[1], "opacity:0"), (T, "opacity:0")], "")
        code.append(mono.run(num, CODE_X + GUTTER - 6 - len(num) * adv, y, FS, th["muted"],
                             f' class="{a}"'))
        # tokens
        x = text_x
        for kind, val in tokens(line):
            colour = {"kw": th["kw"], "str": th["str"], "num": th["num"], "fn": th["fn"],
                      "prop": th["prop"], "punct": th["punct"], "text": th["text"]}[kind]
            code.append(mono.run(val, x, y, FS, colour))
            x += len(val) * adv
        n = len(line.rstrip())
        if n:
            # cover that shrinks toward the right edge in n steps: one char per step
            cw = n * adv + 1
            a = css.anim([(0, "opacity:1;transform:scaleX(1)"),
                          (t0, "opacity:1;transform:scaleX(1)", f"steps({n},end)"),
                          (t1, "opacity:1;transform:scaleX(0)"),
                          (T, "opacity:1;transform:scaleX(0)")],
                         "opacity:0;transform-box:fill-box;transform-origin:right center;")
            code.append(f'<rect class="{a}" x="{text_x - .5:.2f}" y="{y - FS - 1:.2f}" width="{cw:.2f}"'
                        f' height="{LH:.2f}" fill="{th["panel"]}"/>')
            # cursor that rides the typing edge
            a = css.anim([(0, "opacity:0;transform:translateX(0)", "step-end"),
                          (t0, "opacity:1;transform:translateX(0)", f"steps({n},end)"),
                          (t1, f"opacity:1;transform:translateX({n * adv:.2f}px)", "step-end"),
                          (t1 + 0.06, "opacity:0"), (T, "opacity:0")],
                         "opacity:0;")
            code.append(f'<rect class="{a}" x="{text_x:.2f}" y="{y - FS + 0.5:.2f}" width="2"'
                        f' height="{FS + 3}" fill="{th["accent"]}"/>')
    a = css.anim([(0, "opacity:1"), (FADE_OUT[0], "opacity:1"), (FADE_OUT[1], "opacity:0"),
                  (T, "opacity:0")], "")
    P(f'<g class="{a}">{"".join(code)}</g>')

    # ---- terminal
    P(f'<rect x="{RIGHT_X + .5}" y="{TERM_Y + .5}" width="{RIGHT_W - 1}" height="{TERM_H - 1}" rx="10"'
      f' fill="{th["panel"]}" stroke="{th["border"]}"/>')
    tfs = 14
    tadv = mono.adv(tfs)
    ty = TERM_Y + 34
    tx = RIGHT_X + 20
    cmd = "node report.mts"
    term = [mono.run("$", tx, ty, tfs, th["accent"])]
    term.append(mono.run(cmd, tx + 2 * tadv, ty, tfs, th["text"]))
    c0 = run_t
    c1 = c0 + len(cmd) * 0.035
    a = css.anim([(0, "opacity:1;transform:scaleX(1)"),
                  (c0, "opacity:1;transform:scaleX(1)", f"steps({len(cmd)},end)"),
                  (c1, "opacity:1;transform:scaleX(0)"), (T, "opacity:1;transform:scaleX(0)")],
                 "opacity:0;transform-box:fill-box;transform-origin:right center;")
    term.append(f'<rect class="{a}" x="{tx + 2 * tadv - .5:.2f}" y="{ty - tfs - 1}" width="{len(cmd) * tadv + 1:.2f}"'
                f' height="{tfs + 6}" fill="{th["panel"]}"/>')
    # second line: the prompt returns, the write is done
    done = c1 + 0.45
    a = css.anim([(0, "opacity:0"), (done, "opacity:0", "step-end"), (done + 0.001, "opacity:1"),
                  (T, "opacity:1")], "")
    term.append(f'<g class="{a}">{mono.run("$", tx, ty + 26, tfs, th["accent"])}'
                f'<rect x="{tx + 2 * tadv:.2f}" y="{ty + 26 - tfs + .5}" width="{tadv * .9:.2f}"'
                f' height="{tfs + 3}" fill="{th["muted"]}" opacity=".7"/></g>')
    a = css.anim([(0, "opacity:1"), (FADE_OUT[0], "opacity:1"), (FADE_OUT[1], "opacity:0"),
                  (T, "opacity:0")], "")
    P(f'<g class="{a}">{"".join(term)}</g>')

    # ---- file chip: report.pptx appears on disk
    kb = round(pptx_bytes / 1024)
    chip_label = "report.pptx"
    size_label = f"{kb} KB"
    cfs = 14
    cadv = mono.adv(cfs)
    chip_w = 46 + len(chip_label) * cadv + 14 + len(size_label) * cadv + 16
    cx = RIGHT_X
    ch = 40
    icon = (f'<path d="M{cx + 16} {CHIP_Y + 9}h11l6 6v16h-17z" fill="none" stroke="{th["accent"]}"'
            f' stroke-width="1.6" stroke-linejoin="round"/><path d="M{cx + 27} {CHIP_Y + 9}v6h6"'
            f' fill="none" stroke="{th["accent"]}" stroke-width="1.6" stroke-linejoin="round"/>'
            f'<rect x="{cx + 19.5}" y="{CHIP_Y + 20}" width="10" height="7" rx="1" fill="{th["accent"]}"/>')
    chip = (f'<rect x="{cx + .5}" y="{CHIP_Y + .5}" width="{chip_w:.1f}" height="{ch}" rx="20"'
            f' fill="{th["bg"]}" stroke="{th["accent"]}"/>{icon}'
            + mono.run(chip_label, cx + 46, CHIP_Y + 25, cfs, th["text"])
            + mono.run(size_label, cx + 46 + (len(chip_label) + 2) * cadv, CHIP_Y + 25, cfs, th["muted"]))
    a = css.anim([(0, "opacity:0;transform:translateY(-8px) scale(.94)"),
                  (done, "opacity:0;transform:translateY(-8px) scale(.94)", "cubic-bezier(.2,.8,.3,1.2)"),
                  (done + 0.4, "opacity:1;transform:none"),
                  (FADE_OUT[0], "opacity:1;transform:none"), (FADE_OUT[1], "opacity:0;transform:none"),
                  (T, "opacity:0")],
                 "transform-box:fill-box;transform-origin:left center;")
    P(f'<g class="{a}">{chip}</g>')

    # connector chip -> slide
    lx = cx + 24.5
    l0, l1 = CHIP_Y + ch + 8, SLIDE_Y - 10
    ln = l1 - l0
    arrow = (f'<path d="M{lx} {l0}V{l1}" stroke="{th["accent"]}" stroke-width="1.6" fill="none"'
             f' stroke-dasharray="{ln}" class="CLS"/>')
    d0 = done + 0.35
    d1 = d0 + 0.35
    a = css.anim([(0, f"stroke-dashoffset:{ln}"), (d0, f"stroke-dashoffset:{ln}", "ease-in"),
                  (d1, "stroke-dashoffset:0"), (T, "stroke-dashoffset:0")], "")
    head = f'<path d="M{lx - 5} {l1 - 6}L{lx} {l1}L{lx + 5} {l1 - 6}" stroke="{th["accent"]}" stroke-width="1.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/>'
    b = css.anim([(0, "opacity:0"), (d1 - 0.05, "opacity:0"), (d1, "opacity:1"),
                  (FADE_OUT[0], "opacity:1"), (FADE_OUT[1], "opacity:0"), (T, "opacity:0")], "")
    c = css.anim([(0, "opacity:1"), (FADE_OUT[0], "opacity:1"), (FADE_OUT[1], "opacity:0"),
                  (T, "opacity:0")], "")
    P(f'<g class="{c}">{arrow.replace("CLS", a)}</g><g class="{b}">{head}</g>')

    # ---- slide
    img = Image.open(slide_png).convert("RGB")
    iw, ih = img.size
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=84, optimize=True, progressive=True)
    b64 = base64.b64encode(buf.getvalue()).decode()
    sx, sy = RIGHT_X, SLIDE_Y
    k = SLIDE_W / iw
    s0 = d1 + 0.05
    s1 = s0 + 0.55
    slide = [
        f'<rect x="{sx}" y="{sy + 3}" width="{SLIDE_W}" height="{SLIDE_H:.2f}" rx="4" fill="{th["shadow"]}"'
        f' filter="url(#blur)"/>',
        f'<image x="{sx}" y="{sy}" width="{SLIDE_W}" height="{SLIDE_H:.2f}" clip-path="url(#slideclip)"'
        f' preserveAspectRatio="none" href="data:image/jpeg;base64,{b64}"/>',
        f'<rect x="{sx + .5}" y="{sy + .5}" width="{SLIDE_W - 1}" height="{SLIDE_H - 1:.2f}" rx="4"'
        f' fill="none" stroke="{th["border"]}"/>',
    ]
    # bars grow: white covers over each bar (and its label) that retract upward
    bars = detect_bars(img)
    for j, (x0, x1, btop, base) in enumerate(bars):
        g0 = s1 - 0.1 + j * 0.14
        g1 = g0 + 0.55
        ry = max(btop - ih * 0.075, 0)
        a = css.anim([(0, "opacity:1;transform:scaleY(1)"),
                      (g0, "opacity:1;transform:scaleY(1)", "cubic-bezier(.3,0,.2,1)"),
                      (g1, "opacity:1;transform:scaleY(0)"), (T, "opacity:1;transform:scaleY(0)")],
                     "opacity:0;transform-box:fill-box;transform-origin:center top;")
        slide.append(f'<rect class="{a}" x="{sx + (x0 - 6) * k:.2f}" y="{sy + ry * k:.2f}"'
                     f' width="{(x1 - x0 + 13) * k:.2f}" height="{(base - ry + 0.5) * k:.2f}" fill="#ffffff"/>')
    a = css.anim([(0, "opacity:0;transform:translateY(14px) scale(.97)"),
                  (s0, "opacity:0;transform:translateY(14px) scale(.97)", "cubic-bezier(.2,.7,.2,1)"),
                  (s1, "opacity:1;transform:none"),
                  (FADE_OUT[0], "opacity:1;transform:none"), (FADE_OUT[1], "opacity:0;transform:none"),
                  (T, "opacity:0")],
                 "transform-box:fill-box;transform-origin:center top;")
    # caption under the slide
    cap = "Slide 1 of report.pptx, rendered by LibreOffice"
    slide.append(mono.run(cap, sx, sy + SLIDE_H + 26, 12, th["muted"]))
    P(f'<g class="{a}">{"".join(slide)}</g>')

    defs = (f'<defs>{mono.defs()}{bold.defs()}'
            f'<clipPath id="slideclip"><rect x="{sx}" y="{sy}" width="{SLIDE_W}" height="{SLIDE_H:.2f}" rx="4"/></clipPath>'
            f'<filter id="blur" x="-5%" y="-5%" width="110%" height="115%"><feGaussianBlur stdDeviation="6"/></filter></defs>')
    style = (f"<style>{css.text()}"
             "@media (prefers-reduced-motion:reduce){*{animation:none!important}}</style>")
    title = "<title>A TypeScript program runs and writes report.pptx, a slide with a bar chart</title>"
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}"'
            f' role="img">{title}{style}{defs}<rect width="{W}" height="{H}" fill="{th["bg"]}"/>'
            f'{"".join(parts)}</svg>')


def main():
    if "--no-deck" not in sys.argv:
        build_deck()
    code_lines = open(os.path.join(HERE, "report.mts")).read().rstrip("\n").split("\n")
    size = os.path.getsize(os.path.join(HERE, "report.pptx"))
    for name in THEMES:
        svg = build(name, code_lines, os.path.join(HERE, "slide.png"), size)
        assert "<text" not in svg
        out = os.path.join(HERE, f"demo-{name}.svg")
        with open(out, "w") as f:
            f.write(svg)
        print(f"{out}: {len(svg) / 1024:.0f} KB")


if __name__ == "__main__":
    main()
