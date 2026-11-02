VisualTeX Latin Modern Math
===========================

VisualTeX Latin Modern Math is a modified version of Latin Modern Math 1.959
(CTAN package lm-math), Copyright 2012--2014 B. Jackowski, P. Strzelczyk and
P. Pianowski, on behalf of TeX users groups. It is distributed under the GUST
Font License (LaTeX Project Public License 1.3c with a renaming request); see
GUST-FONT-LICENSE.txt.

What was changed
----------------
- The PostScript (CFF) glyph outlines were converted to TrueType (quadratic)
  outlines with fontTools cu2qu (maximum error 1 font unit).
- The font was renamed (family "VisualTeX Latin Modern Math", PostScript name
  VisualTeXLatinModernMath-Regular, file VisualTeXLatinModernMath.ttf) and the
  name table states the change.
- Nothing else: glyphs, metrics and the OpenType MATH, GSUB and GPOS tables are
  those of Latin Modern Math 1.959.

The conversion is reproduced by apps/windows/scripts/make_visualtex_latin_modern_math.py
(fontTools 4.66.1) from latinmodern-math.otf with SHA-256
6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F.

Why
---
Microsoft Word's built-in PDF export cannot embed fonts with PostScript
outlines. It rasterizes such text, and in doing so it drops the parentheses of
native equation numbers (#(n)). With TrueType outlines Word embeds the font, so
equation numbers keep their parentheses and the math stays vector text.

This modified version is not supported by the authors of Latin Modern Math.
The unmodified original is installed alongside it (latinmodern-math.otf) and is
available from https://ctan.org/pkg/lm-math.

Manifest
--------
VisualTeXLatinModernMath.ttf             font family "VisualTeX Latin Modern Math"
  SHA-256 BB3646F9296C10C9B5CCD95BF5680297A4E35B3BC25EA5F4F7593A9C405E0C93
README-VisualTeX-Latin-Modern-Math.txt   this file
GUST-FONT-LICENSE.txt                    the license of Latin Modern Math
