VisualTeX TeX Gyre Termes Math
==============================

VisualTeX TeX Gyre Termes Math is a modified version of TeX Gyre Termes Math
1.543 (CTAN package tex-gyre-math), Copyright 2012--2014 B. Jackowski,
P. Strzelczyk and P. Pianowski, on behalf of TeX users groups. It is
distributed under the GUST Font License (LaTeX Project Public License 1.3c
with a renaming request); see GUST-FONT-LICENSE.txt.

What was changed
----------------
- The PostScript (CFF) glyph outlines were converted to TrueType (quadratic)
  outlines with fontTools cu2qu (maximum error 1 font unit).
- The font was renamed (family "VisualTeX TeX Gyre Termes Math", PostScript
  name VisualTeXTeXGyreTermesMath-Regular, file VisualTeXTeXGyreTermesMath.ttf)
  and the name table states the change.
- Nothing else: glyphs, metrics and the OpenType MATH, GSUB and GPOS tables are
  those of TeX Gyre Termes Math 1.543.

The conversion is reproduced by apps/windows/scripts/make_visualtex_math_font.py
termes (fontTools 4.66.1) from texgyretermes-math.otf with SHA-256
E5CE42B4FE826B7EA4C0874359DECF0B294439355C511A4794107A9924E2CD5A.

Why
---
Microsoft Word's built-in PDF export cannot embed fonts with PostScript
outlines. It rasterizes such text, and in doing so it drops the parentheses of
native equation numbers (#(n)). With TrueType outlines Word embeds the font, so
equation numbers keep their parentheses and the math stays vector text.

This modified version is not supported by the authors of TeX Gyre Termes Math.
The unmodified original is available from https://ctan.org/pkg/tex-gyre-math.

Manifest
--------
VisualTeXTeXGyreTermesMath.ttf             font family "VisualTeX TeX Gyre Termes Math"
  SHA-256 9F58AD9C3C92D18BB439FF31535FBA7A8038624E5E5E7605E2681FFDB9C8B1EE
README-VisualTeX-TeX-Gyre-Termes-Math.txt  this file
GUST-FONT-LICENSE.txt                      the license of TeX Gyre Termes Math
