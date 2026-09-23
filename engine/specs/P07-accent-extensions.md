# P07 extended decorations

Supported additional spellings: overgroup, undergroup, overleftharpoon,
overrightharpoon, overlinesegment and underlinesegment. Each is a real Accent node
with an editable body; source/selection/copy/delete/history share the same owner.
No new opaque preview node, macro expansion or third-party renderer was added.

## Marks and metrics

The group marks use Unicode U+23E0/U+23E1 (top/bottom tortoise-shell brackets).
The harpoons use U+21BC/U+21C0. These codepoint meanings were cross-checked with
Unicode's names list and the existing public command-to-character data (read-only
inspection only; no old renderer executed or path/layout implementation copied).

Font-based marks use real horizontal OpenType MATH variants and connector
assemblies, retaining each piece's identity and uniform scale. Explicitly
registered fallback fonts are consulted only when the primary glyph is absent;
there is no implicit system-font or different-codepoint substitution. Missing
stretch data still fails explicitly under the existing variant/assembly policy.

Line-segment marks are authored geometric rules with two perpendicular endpoint
ticks. They are not top/bottom square brackets, a plain overline, or replacement
font outlines. The horizontal rule and tick thickness use active overbar/underbar
MATH rule thickness. Tick full height is max(4*ruleThickness, 0.18 local em), a
finite engine policy. The nearest tick edge is offset by the font's overbar or
underbar gap. Three rectangles comprise the mark; natural bounds account for ink
overhang and signed mathematical advance. Extra ascender/descender values continue
to come from the font. These defaults are documented policy rather than screenshot
coordinates or a claim of pixel identity with every external editor.

All six commands inherit color, respect phantom paint suppression, allow nested
structures and retain exact authored spelling/trivia. Unknown/unclosed sources
stay recoverable and cannot become finished exports. A completed accent chain
exits by existing structural navigation without adding source or undo history.

## Evidence to collect

accent_extensions.rs covers registry/source semantics, blank/unknown input,
actual structural edits, ID-preserving wrapping, independent copy/edit/delete,
undo/redo, all four mathematical styles and long genuine assemblies, oriented
mark geometry, segment tick thickness and gaps, natural dimensions, colors,
phantoms, cache, and 360 source disturbances (seed 0xACCE2026).

The existing layout-parity runner gains all six short/long/nested/colored samples
and native/WASM editor actions for each new spelling. The original accent-browser
runner retains its previous tests and adds actual controls/keyboard/pointer,
body edits and source/undo checks for each spelling. Its separate extension
screenshot is target/accent-browser/accent-extensions.png.

Original K0 and later fraction-form tests remain in place. The original production
P00 guard remains strict and may still flag external Office edits. This slice
establishes neither general TeX/macros, the whole P07 language matrix, P08
Shift+Enter, Office integration, physical OS IME reruns nor final performance.

Public reference data:
https://www.unicode.org/charts/nameslist/n_2300.html
https://www.unicode.org/charts/nameslist/n_2100.html
https://mathlive.io/mathfield/reference/commands/
https://learn.microsoft.com/en-us/typography/opentype/spec/math
