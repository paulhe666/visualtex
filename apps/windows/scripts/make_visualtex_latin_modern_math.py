#!/usr/bin/env python3
"""Build VisualTeX Latin Modern Math from the pinned Latin Modern Math 1.959.

Microsoft Word's built-in PDF export cannot embed fonts with PostScript (CFF)
outlines. It rasterizes such text, and while doing so it drops the parentheses
of native equation numbers (#(n)). This script converts the glyph outlines to
TrueType (quadratic) outlines, which Word embeds, and renames the font as the
GUST Font License requests for derived works. Glyphs, metrics and the MATH,
GSUB and GPOS tables are otherwise unchanged.

usage: make_visualtex_latin_modern_math.py latinmodern-math.otf VisualTeXLatinModernMath.ttf
"""
import hashlib
import sys

from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable

SOURCE_SHA256 = "6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F"
FAMILY = "VisualTeX Latin Modern Math"
POSTSCRIPT = "VisualTeXLatinModernMath-Regular"
CHANGE = ("Modified by VisualTeX: the PostScript (CFF) outlines of Latin Modern Math 1.959 "
          "were converted to TrueType outlines (fontTools cu2qu, maximum error 1 font unit) "
          "and the font was renamed; nothing else was changed. This modified version is not "
          "supported by the authors of Latin Modern Math.")
LICENSE = ("This work is released under the GUST Font License (LaTeX Project Public License 1.3c "
           "with a renaming request). The unmodified original is available from "
           "https://ctan.org/pkg/lm-math.")
LICENSE_URL = "http://tug.org/fonts/licenses/GUST-FONT-LICENSE.txt"


def sha256(path):
    with open(path, "rb") as stream:
        return hashlib.sha256(stream.read()).hexdigest().upper()


def main(source, target):
    actual = sha256(source)
    if actual != SOURCE_SHA256:
        raise SystemExit(f"unexpected source font {actual}; expected Latin Modern Math 1.959 {SOURCE_SHA256}")

    # recalcTimestamp=False keeps head.modified from the source, so the same
    # fontTools version reproduces the same bytes.
    font = TTFont(source, recalcTimestamp=False)
    if font.sfntVersion != "OTTO" or "CFF " not in font or "MATH" not in font:
        raise SystemExit("the source is not the CFF OpenType math font")
    order = font.getGlyphOrder()
    glyphset = font.getGlyphSet()

    glyphs = {}
    for name in order:
        pen = TTGlyphPen(glyphset)
        glyphset[name].draw(Cu2QuPen(pen, max_err=1.0, reverse_direction=True))
        glyphs[name] = pen.glyph()

    font["loca"] = newTable("loca")
    glyf = font["glyf"] = newTable("glyf")
    glyf.glyphOrder = order
    glyf.glyphs = glyphs
    del font["CFF "]
    if "VORG" in font:
        del font["VORG"]
    glyf.compile(font)

    hmtx = font["hmtx"]
    for name, glyph in glyf.glyphs.items():
        advance, _ = hmtx.metrics[name]
        hmtx.metrics[name] = (advance, glyph.xMin if glyph.numberOfContours else 0)

    maxp = font["maxp"] = newTable("maxp")
    maxp.tableVersion = 0x00010000
    for field in ("maxZones", "maxTwilightPoints", "maxStorage", "maxFunctionDefs",
                  "maxInstructionDefs", "maxStackElements", "maxSizeOfInstructions",
                  "maxComponentElements", "maxComponentDepth"):
        setattr(maxp, field, 0)
    maxp.maxZones = 1

    post = font["post"]
    post.formatType = 2.0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = order
    font.sfntVersion = "\x00\x01\x00\x00"

    names = font["name"]
    version = names.getDebugName(5) or "Version 1.959"
    for record in list(names.names):
        if record.nameID in (1, 4, 16, 21):
            record.string = FAMILY
        elif record.nameID == 3:
            record.string = f"{FAMILY} 1.959 (TrueType outlines)"
        elif record.nameID == 6:
            record.string = POSTSCRIPT
        elif record.nameID == 5:
            record.string = f"{version}; VisualTeX TrueType outlines"
        elif record.nameID == 0:
            record.string = f"{record.toUnicode()} {CHANGE}"
    for name_id, text in ((10, CHANGE), (13, LICENSE), (14, LICENSE_URL)):
        names.setName(text, name_id, 3, 1, 0x409)

    font.save(target)

    check = TTFont(target)
    if "MATH" not in check or "glyf" not in check or len(check.getGlyphOrder()) != len(order):
        raise SystemExit("the converted font lost its MATH table or glyphs")
    print(f"{target}: {len(order)} glyphs, family '{check['name'].getDebugName(1)}', sha256 {sha256(target)}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])
