#!/usr/bin/env python3
"""Build the TrueType math fonts VisualTeX installs for Word from their pinned CFF originals.

Microsoft Word's built-in PDF export cannot embed fonts with PostScript (CFF)
outlines. It rasterizes such text, and while doing so it drops the parentheses
of native equation numbers (#(n)). This script converts the glyph outlines to
TrueType (quadratic) outlines, which Word embeds, and renames the font as the
GUST Font License requests for derived works. Glyphs, metrics and the MATH,
GSUB and GPOS tables are otherwise unchanged.

usage: make_visualtex_math_font.py <font> source.otf target.ttf
       <font>: latin-modern (latinmodern-math.otf 1.959) or termes (texgyretermes-math.otf 1.543)
"""
import hashlib
import sys

from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable

LICENSE_URL = "http://tug.org/fonts/licenses/GUST-FONT-LICENSE.txt"

# Each font: its pinned source, its new names, and the original's name, version and home.
FONTS = {
    "latin-modern": {
        "source_sha256": "6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F",
        "family": "VisualTeX Latin Modern Math",
        "postscript": "VisualTeXLatinModernMath-Regular",
        "original": "Latin Modern Math",
        "version": "1.959",
        "home": "https://ctan.org/pkg/lm-math",
    },
    "termes": {
        "source_sha256": "E5CE42B4FE826B7EA4C0874359DECF0B294439355C511A4794107A9924E2CD5A",
        "family": "VisualTeX TeX Gyre Termes Math",
        "postscript": "VisualTeXTeXGyreTermesMath-Regular",
        "original": "TeX Gyre Termes Math",
        "version": "1.543",
        "home": "https://ctan.org/pkg/tex-gyre-math",
    },
}


def sha256(path):
    with open(path, "rb") as stream:
        return hashlib.sha256(stream.read()).hexdigest().upper()


def convert(spec, source, target):
    family, version, original = spec["family"], spec["version"], spec["original"]
    change = (f"Modified by VisualTeX: the PostScript (CFF) outlines of {original} {version} "
              "were converted to TrueType outlines (fontTools cu2qu, maximum error 1 font unit) "
              "and the font was renamed; nothing else was changed. This modified version is not "
              f"supported by the authors of {original}.")
    license_text = ("This work is released under the GUST Font License (LaTeX Project Public License 1.3c "
                    f"with a renaming request). The unmodified original is available from {spec['home']}.")

    actual = sha256(source)
    if actual != spec["source_sha256"]:
        raise SystemExit(f"unexpected source font {actual}; expected {original} {version} {spec['source_sha256']}")

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
    source_version = names.getDebugName(5) or f"Version {version}"
    for record in list(names.names):
        if record.nameID in (1, 4, 16, 21):
            record.string = family
        elif record.nameID == 3:
            record.string = f"{family} {version} (TrueType outlines)"
        elif record.nameID == 6:
            record.string = spec["postscript"]
        elif record.nameID == 5:
            record.string = f"{source_version}; VisualTeX TrueType outlines"
        elif record.nameID == 0:
            record.string = f"{record.toUnicode()} {change}"
    for name_id, text in ((10, change), (13, license_text), (14, LICENSE_URL)):
        names.setName(text, name_id, 3, 1, 0x409)

    font.save(target)

    check = TTFont(target)
    if "MATH" not in check or "glyf" not in check or len(check.getGlyphOrder()) != len(order):
        raise SystemExit("the converted font lost its MATH table or glyphs")
    print(f"{target}: {len(order)} glyphs, family '{check['name'].getDebugName(1)}', sha256 {sha256(target)}")


if __name__ == "__main__":
    if len(sys.argv) != 4 or sys.argv[1] not in FONTS:
        raise SystemExit(__doc__)
    convert(FONTS[sys.argv[1]], sys.argv[2], sys.argv[3])
