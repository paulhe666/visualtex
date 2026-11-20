#!/usr/bin/env python3
"""Build VisualTeX Latin Modern Math from the pinned Latin Modern Math 1.959
(make_visualtex_math_font.py latin-modern; kept for the README's reproduction note).

usage: make_visualtex_latin_modern_math.py latinmodern-math.otf VisualTeXLatinModernMath.ttf
"""
import sys

from make_visualtex_math_font import FONTS, convert

if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    convert(FONTS["latin-modern"], sys.argv[1], sys.argv[2])
