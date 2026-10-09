#!/usr/bin/env python3
"""Generate the "slide as canvas" README banners for pptx-ts.

Run: python3 generate.py   (needs fontTools, rsvg-convert, and the Adwaita fonts)
Writes banner-light.svg, banner-dark.svg, their PNG previews and preview.png.
All text is emitted as glyph outlines, so the SVGs need no fonts at view time.
"""
import os
import subprocess
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont
from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen

HERE = os.path.dirname(os.path.abspath(__file__))
SANS = "/usr/share/fonts/Adwaita/AdwaitaSans-Regular.ttf"  # Inter-derived, OFL
MONO = "/usr/share/fonts/Adwaita/AdwaitaMono-Regular.ttf"  # Iosevka-derived, OFL

W, H = 1280, 320
_fonts = {}


def font(path, **axes):
    key = (path, tuple(sorted(axes.items())))
    if key not in _fonts:
        f = TTFont(path)
        if axes and "fvar" in f:
            f = instantiateVariableFont(f, axes)
        _fonts[key] = f
    return _fonts[key]


def fmt(v):
    s = f"{v:.1f}"
    return s[:-2] if s.endswith(".0") else s


_glyph_defs = {}


def text_path(f, text, x, y, size, tracking=0.0, anchor="start", fill="#000", inline=False):
    """Lay out text on a baseline as <use> refs to shared glyph outlines.

    Returns (svg markup, advance width in px). Glyph outlines are stored once in
    font units in _glyph_defs and placed with a scale/flip transform.
    """
    upm = f["head"].unitsPerEm
    cmap = f.getBestCmap()
    gs = f.getGlyphSet()
    hmtx = f["hmtx"]
    s = size / upm
    names = [cmap[ord(c)] for c in text]
    adv = [hmtx[n][0] * s + tracking * size for n in names]
    total = sum(adv) - tracking * size
    if anchor == "middle":
        x -= total / 2
    elif anchor == "end":
        x -= total
    fid = _font_ids.setdefault(id(f), "f%d" % len(_font_ids))
    out = []
    cx = x
    if inline:  # one absolute path, so a userSpaceOnUse gradient spans the run
        d = []
        for n, a in zip(names, adv):
            pen = SVGPathPen(gs, ntos=fmt)
            gs[n].draw(TransformPen(pen, (s, 0, 0, -s, cx, y)))
            d.append(pen.getCommands())
            cx += a
        return f'<path fill="{fill}" d="{"".join(d)}"/>', total
    for n, a in zip(names, adv):
        gid = f"{fid}{n}".replace(".", "_")
        if gid not in _glyph_defs:
            pen = SVGPathPen(gs, ntos=lambda v: str(round(v)))
            gs[n].draw(pen)
            _glyph_defs[gid] = pen.getCommands()
        if _glyph_defs[gid]:
            out.append(f'<use href="#{gid}" transform="matrix({s:.5g} 0 0 {-s:.5g} {fmt(cx)} {fmt(y)})"/>')
        cx += a
    return f'<g fill="{fill}">' + "".join(out) + "</g>", total


_font_ids = {}


THEMES = {
    "light": dict(
        accent="#b4451f", g0="#b4451f", g1="#d9743f",
        ink="#1f2328", muted="#59636e", faint="#818b98",
        slide="#ffffff", slide_stroke="#d1d9e0", chrome="#f6f8fa",
        rail_stroke="#d1d9e0", ph="#d1d9e0", ph_fill="#eff2f5",
        shadow="#1f2328", shadow_op="0.10", bar="#e6d5cc", handle_fill="#ffffff",
    ),
    "dark": dict(
        accent="#e9855c", g0="#e9855c", g1="#f3b48f",
        ink="#f0f6fc", muted="#9198a1", faint="#656c76",
        slide="#161b22", slide_stroke="#3d444d", chrome="#151b23",
        rail_stroke="#3d444d", ph="#3d444d", ph_fill="#212830",
        shadow="#000000", shadow_op="0.0", bar="#3a3330", handle_fill="#161b22",
    ),
}


def build(theme):
    _glyph_defs.clear()
    t = THEMES[theme]
    bold = font(SANS, wght=720, opsz=32)
    medium = font(SANS, wght=500, opsz=20)
    regular = font(SANS, wght=480, opsz=14)
    mono = font(MONO)

    el = []
    a = el.append

    # ---- thumbnail rail -------------------------------------------------
    tw, th = 128, 72
    tx = 36
    ty0 = (H - (3 * th + 2 * 16)) / 2
    for i in range(3):
        y = ty0 + i * (th + 16)
        num, _ = text_path(regular, str(i + 1), tx - 12, y + 13, 12, anchor="end",
                           fill=t["accent"] if i == 0 else t["faint"])
        a(num)
        if i == 0:
            a(f'<rect x="{tx-4}" y="{fmt(y-4)}" width="{tw+8}" height="{th+8}" rx="7" '
              f'fill="none" stroke="{t["accent"]}" stroke-width="2.5"/>')
        a(f'<rect x="{tx}" y="{fmt(y)}" width="{tw}" height="{th}" rx="4" fill="{t["slide"]}" '
          f'stroke="{t["rail_stroke"]}"/>')
        if i == 0:  # mini title slide
            a(f'<rect x="{tx+14}" y="{fmt(y+24)}" width="52" height="9" rx="2" fill="{t["ink"]}"/>')
            a(f'<rect x="{tx+68}" y="{fmt(y+24)}" width="20" height="9" rx="2" fill="url(#g)"/>')
            a(f'<rect x="{tx+14}" y="{fmt(y+40)}" width="70" height="4" rx="2" fill="{t["ph"]}"/>')
            for k, hgt in enumerate((12, 19, 15, 27)):
                a(f'<rect x="{tx+90+k*7}" y="{fmt(y+51-hgt)}" width="5" height="{hgt}" rx="1" '
                  f'fill="{"url(#gb)" if k == 3 else t["bar"]}"/>')
        elif i == 1:  # chart slide
            a(f'<rect x="{tx+12}" y="{fmt(y+10)}" width="44" height="5" rx="2" fill="{t["muted"]}"/>')
            for k, hgt in enumerate((18, 30, 24, 40)):
                a(f'<rect x="{tx+16+k*26}" y="{fmt(y+62-hgt)}" width="16" height="{hgt}" rx="2" '
                  f'fill="{"url(#gb)" if k == 3 else t["bar"]}"/>')
            a(f'<rect x="{tx+12}" y="{fmt(y+62)}" width="104" height="1" fill="{t["ph"]}"/>')
        else:  # table slide
            a(f'<rect x="{tx+12}" y="{fmt(y+10)}" width="44" height="5" rx="2" fill="{t["muted"]}"/>')
            a(f'<rect x="{tx+12}" y="{fmt(y+22)}" width="104" height="9" rx="1.5" fill="url(#g)" opacity=".8"/>')
            for r in range(3):
                a(f'<rect x="{tx+12}" y="{fmt(y+34+r*11)}" width="104" height="8" rx="1.5" '
                  f'fill="{t["ph_fill"]}"/>')

    # ---- main slide -----------------------------------------------------
    sx, sy, sw, sh = 196, 22, 1062, 276
    a(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh}" rx="10" fill="{t["slide"]}" '
      f'stroke="{t["slide_stroke"]}" filter="url(#s)"/>')

    # wordmark in a selected title placeholder
    size = 112
    base = 165
    wx = 262
    d1, w1 = text_path(bold, "pptx", wx, base, size, tracking=-0.025, fill=t["ink"])
    d2, w2 = text_path(bold, "-ts", wx + w1 - 0.025 * size + 2, base, size, tracking=-0.025,
                       fill="url(#gw)", inline=True)
    a(d1)
    a(d2)
    word_w = w1 + w2 + 2 - 0.025 * size
    bx, by = wx - 18, 57
    bw, bh = word_w + 36, 138
    a(f'<rect x="{fmt(bx)}" y="{by}" width="{fmt(bw)}" height="{bh}" fill="none" '
      f'stroke="{t["accent"]}" stroke-width="1.5" stroke-dasharray="5 4"/>')
    # rotate handle
    cxm = bx + bw / 2
    a(f'<path d="M{fmt(cxm)} {by} v-16" stroke="{t["accent"]}" stroke-width="1.5"/>')
    a(f'<circle cx="{fmt(cxm)}" cy="{by-20}" r="5" fill="{t["handle_fill"]}" stroke="{t["accent"]}" stroke-width="1.5"/>')
    # eight resize handles
    for hx in (bx, bx + bw / 2, bx + bw):
        for hy in (by, by + bh / 2, by + bh):
            if hx == bx + bw / 2 and hy == by + bh / 2:
                continue
            a(f'<rect x="{fmt(hx-4.5)}" y="{fmt(hy-4.5)}" width="9" height="9" rx="1.5" '
              f'fill="{t["handle_fill"]}" stroke="{t["accent"]}" stroke-width="1.5"/>')

    # subtitle placeholder text
    tag, _ = text_path(medium, "Write a program, get a PowerPoint file.", wx, 250, 29,
                       tracking=-0.01, fill=t["muted"])
    a(tag)

    # ---- right side: chart object on the slide --------------------------
    cx0, cbase = 940, 236
    bars = (64, 104, 84, 150)
    for k, hgt in enumerate(bars):
        x = cx0 + k * 62
        fill = "url(#gb)" if k == len(bars) - 1 else t["bar"]
        a(f'<rect x="{x}" y="{cbase-hgt}" width="40" height="{hgt}" rx="5" fill="{fill}"/>')
    a(f'<rect x="{cx0-14}" y="{cbase}" width="{62*3+40+28}" height="1.5" fill="{t["ph"]}"/>')
    # small mono caption under chart, like a code-generated label
    cap, cw = text_path(mono, "slide.addChart()", cx0 - 14 + (62 * 3 + 40 + 28) / 2, 266, 14,
                        anchor="middle", fill=t["faint"])
    a(cap)

    defs = (
        f'<defs>'
        f'<linearGradient id="g" x1="0" y1="0" x2="1" y2="0"><stop stop-color="{t["g0"]}"/>'
        f'<stop offset="1" stop-color="{t["g1"]}"/></linearGradient>'
        f'<linearGradient id="gw" gradientUnits="userSpaceOnUse" x1="{fmt(wx+w1)}" y1="0" '
        f'x2="{fmt(wx+word_w)}" y2="0"><stop stop-color="{t["g0"]}"/>'
        f'<stop offset="1" stop-color="{t["g1"]}"/></linearGradient>'
        f'<linearGradient id="gb" x1="0" y1="1" x2="0" y2="0"><stop stop-color="{t["g0"]}"/>'
        f'<stop offset="1" stop-color="{t["g1"]}"/></linearGradient>'
        f'<filter id="s" x="-5%" y="-20%" width="110%" height="140%">'
        f'<feDropShadow dx="0" dy="3" stdDeviation="6" flood-color="{t["shadow"]}" '
        f'flood-opacity="{t["shadow_op"]}"/></filter>'
        + "".join(f'<path id="{k}" d="{v}"/>' for k, v in _glyph_defs.items() if v)
        + f'</defs>'
    )
    title = "pptx-ts: Write a program, get a PowerPoint file."
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}" '
        f'role="img" aria-label="{title}"><title>{title}</title>{defs}' + "".join(el) + "</svg>\n"
    )


def main():
    for theme in THEMES:
        svg = os.path.join(HERE, f"banner-{theme}.svg")
        with open(svg, "w") as fh:
            fh.write(build(theme))
        subprocess.run(["rsvg-convert", "-z", "1", "-o",
                        os.path.join(HERE, f"banner-{theme}.png"), svg], check=True)
    # combined preview: light on #ffffff over dark on #0d1117
    pad = 40
    comp = (
        f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'width="{W+2*pad}" height="{2*(H+2*pad)}">'
        f'<rect width="100%" height="{H+2*pad}" fill="#ffffff"/>'
        f'<rect y="{H+2*pad}" width="100%" height="{H+2*pad}" fill="#0d1117"/>'
        f'<image x="{pad}" y="{pad}" width="{W}" height="{H}" xlink:href="banner-light.svg"/>'
        f'<image x="{pad}" y="{H+3*pad}" width="{W}" height="{H}" xlink:href="banner-dark.svg"/>'
        f'</svg>'
    )
    comp_path = os.path.join(HERE, ".preview.svg")
    with open(comp_path, "w") as fh:
        fh.write(comp)
    subprocess.run(["rsvg-convert", "-o", os.path.join(HERE, "preview.png"), comp_path],
                   check=True, cwd=HERE)
    os.remove(comp_path)


if __name__ == "__main__":
    main()
