#!/usr/bin/env python3
"""Generate the outlines of exam-zh's congruence and similarity signs for the layout engine.

The Chinese textbooks' congruence sign (全等, U+224C, \\backcong) and similarity sign (相似, U+223D, \\backsim) and their
negations are drawn as exam-zh draws them: exam-zh-symbols.sty (v0.3.8, 2026-08-21, commit afffaeac068a,
https://github.com/xkwxdyy/exam-zh, LPPL 1.3c, by Kangwei Xia) defines \\cong, \\ncong, \\sim and \\nsim as TikZ
pictures. The paths below are its coordinates (em and ex of the text font), line widths (TikZ's default 0.4pt; 0.6pt for
the similarity signs), round line caps and baselines; the box of each picture is what TikZ makes of it: the bounding box of
the points and Bezier control points, half a line width bigger on every side, on the given baseline.

TikZ strokes with an absolute line width at every size (exam-zh draws the picture at the text, script and scriptscript
font sizes through \\mathchoice), so each sign has an outline for 10, 7 and 5 pt, as the TeX fonts have. The strokes are
turned into filled outlines (skia), so the engine paints them like font glyphs. ex is Computer Modern's x-height,
0.430555 em, which exam-zh's default font (New Computer Modern) shares.

The boxes agree with TeX's: \\setbox0\\hbox{$\\cong$} under exam-zh-symbols gives 11.19986pt x 7.09601pt + 0.2pt at 10pt
(Latin Modern, ex 4.31pt), the numbers this script checks against.

usage: gen-exam-zh-signs.py [--output <file>]     (needs skia-python: python3 -m pip install skia-python)
"""

from __future__ import annotations

import argparse
import math
from pathlib import Path

try:
    import skia
except ImportError as error:  # pragma: no cover
    raise SystemExit("skia-python is required: python3 -m pip install skia-python") from error

EX_PER_EM = 0.430555
SIZES = (10, 7, 5)
UNITS_PER_EM = 1000

# A sign: its stroked paths, each a list of segments in (em, ex) coordinates (TikZ's: x right, y up), the line width
# in pt, and where its baseline is: "zero" (TikZ's `baseline`, the y = 0 of the picture) or the bounding box's south
# moved by a number of pt (`baseline = {([yshift = d]current bounding box.south)}`).
CONG_CURVE = [
    ("M", (-0.20, 1.35)),
    ("C", (-0.46, 1.60), (-0.54, 0.65), (-0.25, 0.65)),
    ("C", (-0.06, 0.65), (0.06, 1.35), (0.25, 1.35)),
    ("C", (0.54, 1.35), (0.46, 0.40), (0.20, 0.65)),
]
CONG_LINES = [
    [("M", (-0.46, 0.40)), ("L", (0.46, 0.40))],
    [("M", (-0.46, 0.0)), ("L", (0.46, 0.0))],
]
SIM_CURVE = [
    ("M", (-0.25, 1.15)),
    ("C", (-0.55, 1.15), (-0.51, 0.23), (-0.275, 0.23)),
    ("C", (0.0, 0.23), (0.0, 1.15), (0.275, 1.15)),
    ("C", (0.51, 1.15), (0.55, 0.23), (0.25, 0.23)),
]


def slash(start: tuple[float, float], angle_deg: float, length_em: float) -> list:
    """TikZ `(x,y) --++ (angle:length)`: a line from start, length in em; the end is given in em, so mark it."""
    return [("M", start), ("L*", start, angle_deg, length_em)]


SIGNS = {
    # \cong of exam-zh (全等): the S with curled ends over two lines
    "backcong": {"paths": [CONG_CURVE, *CONG_LINES], "line_width": 0.4, "baseline": ("zero", 0.0)},
    # \ncong: the same with a slash (-0.22em,-0.4ex) --++ (66:1em)
    "notbackcong": {"paths": [CONG_CURVE, *CONG_LINES, slash((-0.22, -0.4), 66, 1.0)], "line_width": 0.4, "baseline": ("zero", 0.0)},
    # \sim of exam-zh (相似): line width 0.6pt, baseline 1pt under the box
    "backsim": {"paths": [SIM_CURVE], "line_width": 0.6, "baseline": ("south", -1.0)},
    # \nsim: with a slash (-0.15em,-0.2ex) --++ (70:0.8em), baseline 1pt over the bottom of the box
    "notbacksim": {"paths": [SIM_CURVE, slash((-0.15, -0.2), 70, 0.8)], "line_width": 0.6, "baseline": ("south", 1.0)},
}

# TeX's boxes for the text, script and scriptscript sizes (exam-zh-symbols v0.3.8, xelatex, Latin Modern: ex = 0.431 em)
TEX_BOXES = {
    "backcong": [(11.19986, 7.09601, 0.2), (7.9599, 5.0272, 0.2), (5.79993, 3.648, 0.2)],
    "notbackcong": [(11.19986, 7.61147, 1.92397), (7.9599, 5.38803, 1.40677), (5.79993, 3.90573, 1.06198)],
    "backsim": [(11.60007, 5.5652, 0.0), (8.30005, 4.37564, 0.0), (6.10004, 3.5826, 0.0)],
    "notbacksim": [(11.60007, 7.1176, 1.0), (8.30005, 4.86232, 1.0), (6.10004, 3.3588, 1.0)],
}


def points_in_pt(segments: list, size: float, ex_per_em: float) -> list:
    """The segments with (em, ex) turned into pt at the font size."""
    em, ex = size, ex_per_em * size
    out = []
    for seg in segments:
        kind = seg[0]
        if kind == "L*":
            (sx, sy), angle, length = seg[1], seg[2], seg[3]
            x0, y0 = sx * em, sy * ex
            out.append(("L", (x0 + math.cos(math.radians(angle)) * length * em, y0 + math.sin(math.radians(angle)) * length * em)))
        else:
            out.append((kind, *[(x * em, y * ex) for x, y in seg[1:]]))
    return out


def build(sign: dict, size: float, ex_per_em: float):
    """The filled outline (skia.Path, pt, y up) and the box (left, bottom, right, top of the ink box TikZ makes, pt)."""
    line_width = sign["line_width"]
    paint = skia.Paint()
    paint.setStyle(skia.Paint.kStroke_Style)
    paint.setStrokeWidth(line_width)
    paint.setStrokeCap(skia.Paint.kRound_Cap)
    paint.setStrokeJoin(skia.Paint.kMiter_Join)
    paint.setStrokeMiter(10)
    union = skia.Path()
    xs: list[float] = []
    ys: list[float] = []
    for segments in sign["paths"]:
        path = skia.Path()
        for seg in points_in_pt(segments, size, ex_per_em):
            pts = seg[1:]
            for x, y in pts:
                xs.append(x)
                ys.append(y)
            # skia's y goes down: draw with y negated, flip back when writing
            if seg[0] == "M":
                path.moveTo(pts[0][0], -pts[0][1])
            elif seg[0] == "L":
                path.lineTo(pts[0][0], -pts[0][1])
            elif seg[0] == "C":
                path.cubicTo(pts[0][0], -pts[0][1], pts[1][0], -pts[1][1], pts[2][0], -pts[2][1])
        stroked = skia.Path()
        paint.getFillPath(path, stroked)
        union = skia.Op(union, stroked, skia.PathOp.kUnion_PathOp)
    union = skia.Simplify(union)
    half = line_width / 2
    box = (min(xs) - half, min(ys) - half, max(xs) + half, max(ys) + half)
    return union, box


def svg_d(path, dx: float, dy: float, scale: float) -> str:
    """SVG path data in font units (y up): x' = (x + dx) * scale, y' = (-y + dy) * scale."""
    def f(v: float) -> str:
        s = f"{v:.1f}"
        return "0" if s in ("0.0", "-0.0") else s.rstrip("0").rstrip(".")

    out = []
    it = skia.Path.Iter(path, False)
    while True:
        verb, pts = it.next()
        if verb == skia.Path.kDone_Verb:
            break
        if verb == skia.Path.kConic_Verb:
            # the round caps come as conic arcs: quadratic pieces (4 per arc) are what SVG has
            quads = skia.Path.ConvertConicToQuads(pts[0], pts[1], pts[2], it.conicWeight(), 2)
            for i in range(1, len(quads) - 1, 2):
                (x1, y1), (x2, y2) = quads[i], quads[i + 1]
                out.append(f"Q{f((x1 + dx) * scale)} {f((-y1 + dy) * scale)} {f((x2 + dx) * scale)} {f((-y2 + dy) * scale)}")
            continue
        if verb == skia.Path.kMove_Verb:
            x, y = pts[0]
            out.append(f"M{f((x + dx) * scale)} {f((-y + dy) * scale)}")
        elif verb == skia.Path.kLine_Verb:
            x, y = pts[1]
            out.append(f"L{f((x + dx) * scale)} {f((-y + dy) * scale)}")
        elif verb == skia.Path.kQuad_Verb:
            (x1, y1), (x2, y2) = pts[1], pts[2]
            out.append(f"Q{f((x1 + dx) * scale)} {f((-y1 + dy) * scale)} {f((x2 + dx) * scale)} {f((-y2 + dy) * scale)}")
        elif verb == skia.Path.kCubic_Verb:
            (x1, y1), (x2, y2), (x3, y3) = pts[1], pts[2], pts[3]
            out.append(
                f"C{f((x1 + dx) * scale)} {f((-y1 + dy) * scale)} {f((x2 + dx) * scale)} {f((-y2 + dy) * scale)} "
                f"{f((x3 + dx) * scale)} {f((-y3 + dy) * scale)}"
            )
        elif verb == skia.Path.kConic_Verb:
            raise SystemExit("conic segments are not expected in a stroked outline")
        elif verb == skia.Path.kClose_Verb:
            out.append("Z")
    return "".join(out)


def glyph(sign: dict, size: float, ex_per_em: float) -> dict:
    outline, (left, bottom, right, top) = build(sign, size, ex_per_em)
    mode, shift = sign["baseline"]
    baseline = 0.0 if mode == "zero" else bottom + shift
    width, height, depth = right - left, top - baseline, max(0.0, baseline - bottom)
    scale = UNITS_PER_EM / size
    return {
        "path": svg_d(outline, -left, -baseline, scale),
        "advanceWidth": round(width * scale, 2),
        "height": round(height * scale, 2),
        "depth": round(depth * scale, 2),
        "box_pt": (width, height, depth),
    }


def check_against_tex() -> None:
    """The boxes with TeX's x-height of the measurement (0.431 em) are TeX's own, to 0.001 pt."""
    for name, sign in SIGNS.items():
        for size, (wd, ht, dp) in zip(SIZES, TEX_BOXES[name]):
            w, h, d = glyph(sign, size, 0.431)["box_pt"]
            if max(abs(w - wd), abs(h - ht), abs(d - dp)) > 0.002:
                raise SystemExit(f"{name} at {size}pt: {w:.5f} x {h:.5f} + {d:.5f}, TeX has {wd} x {ht} + {dp}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--output", type=Path, default=Path(__file__).resolve().parent.parent / "src" / "exam-zh-signs.generated.ts")
    args = parser.parse_args()
    check_against_tex()
    lines = [
        "// Generated by scripts/gen-exam-zh-signs.py from exam-zh-symbols.sty v0.3.8 (2026-08-21); do not edit.",
        "// The congruence and similarity signs and their negations as exam-zh draws them (TikZ strokes turned into",
        "// outlines): for each sign an outline at 10, 7 and 5 pt, in thousandths of that size, y up, origin on the baseline.",
        "export const EXAM_ZH_SIGNS = {",
        f"  unitsPerEm: {UNITS_PER_EM},",
        f"  sizes: [{', '.join(str(s) for s in SIZES)}],",
        "  glyphs: {",
    ]
    for name, sign in SIGNS.items():
        lines.append(f"    {name}: [")
        for size in SIZES:
            g = glyph(sign, size, EX_PER_EM)
            lines.append(
                f"      {{ size: {size}, advanceWidth: {g['advanceWidth']}, height: {g['height']}, depth: {g['depth']}, "
                f"path: \"{g['path']}\" }},"
            )
        lines.append("    ],")
    lines += ["  },", "} as const;", ""]
    args.output.write_text("\n".join(lines), encoding="utf-8")
    print(f"wrote {args.output}")


if __name__ == "__main__":
    main()
