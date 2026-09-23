# P07 colors — bounded authored paint and group scope

Scope: independent core and sandbox. Production UI, Office DTOs, original Word
objects, installers and MathLive/MathJax runtime entrypoints are unchanged.

## Supported source

- `\textcolor[model]{value}{body}` is a real editable Color node with a scoped
  body. A direct UI color action wraps a selection or inserts an empty editable
  body, using `\textcolor[HTML]{RRGGBB}{...}`. It never rewrites all source.
- `\color[model]{value}` is a zero-width ColorChange declaration in the current
  mathematical row/group. It is not a symbol or script base. Paint reaches the
  following siblings, and nested row declarations do not leak to outer siblings.
- Supported explicit models: `rgb` (three 0..1 decimals), `RGB` (three 0..255
  integers), `HTML` (six hexadecimal digits, no CSS '#'), `gray` (0..1 decimal),
  and the finite named/default expression set. Final paint channels are 8-bit RGB.
- Named colors: black, white, red, green, blue, cyan, magenta, yellow, gray,
  darkgray, lightgray, brown, lime, olive, orange, pink, purple, teal and violet.
  These use xcolor base definitions rather than CSS color-name substitutions.
- Bounded named mixtures such as `red!50!blue` and `red!25` are resolved without
  code execution. At most 16 mixing stages and 1024 color-source bytes are accepted.
  Each weight must be a finite decimal in 0..100. An omitted final color is white.
- Command spelling, model/value spelling, comments and trivia are retained in
  source. ColorSpec separately stores validated resolved RGB; model validation
  rejects any mismatch between source color and mathematical paint.

Unsupported models/expressions, out-of-range values, arbitrary CSS/URLs, non-finite
numbers, macros/redefinitions and unrecognized names remain explicit unsupported
source. Unclosed headers/bodies remain incomplete drafts. No black fallback hides
unsupported color. Pure color declarations and empty Color bodies cannot be
published as complete formulas.

## Layout, editing and export

Color does not change glyph identity, math style, semantic font size or geometry.
ColorChange is skipped when classifying neighboring math atoms, so introducing a
paint directive does not change unary/binary spacing. Ordinary component geometry
can be cached while final paint is resolved through the current scene traversal.

Color wrappers and row-local declarations carry separate scope information. All
visible child glyphs and rules inherit the effective paint, including fraction
bars, radical rules, accents and boxed borders. Phantom paint suppression wins:
hidden colors never reintroduce glyphs, rules or font-outline resources.

The typed scene adds an optional {r,g,b} to glyph and rule instances. Omitted paint
retains the previous foreground/export behavior and uncolored SVG serialization.
Only derived '#rrggbb' fill attributes are output. Author strings never enter SVG
style, URL, event handler or resource-reference attributes. The browser's rgbHex
validator checks exact channel keys and integer range before creating the new SVG
subtree; malformed paint leaves the old valid SVG intact. Editor overlays are not
recolored as formula contents.

Structural commands, source patches, projected ranges, node copies, delete,
undo/redo and history use the existing EngineSession and PositionMap. Color source
edits inside styled/boxed/hidden bodies recurse through every wrapper. Optional
model brackets inside a root degree are brace-protected by source projection.

## Evidence and limitations

`tests/colors.rs`: 14 parser/structure/font/scene/editor tests, including 640 seeded
source disturbances (seed 0xc0102026); two paint module tests cover values and
faults. Tests verify actual colors, unchanged dimensions, source/structure agreement,
all source identities, copied-node independence, cache freshness and undo paint.

`tools/layout-parity.mjs`: actual Native/WASM comparison includes the supported
names, numerical models, mixtures, nested declarations and wrappers, all four math
styles, hidden colored content, rule colors, source modifications and editor undo.

`tools/color-browser.mjs`: real owned-browser button/key/pointer interaction from
an empty formula, source input, real mathematical selection, color changes, undo,
zoom and invalid-source recovery. It compares each rendered glyph/rule fill with
its typed scene. Sixteen deliberately synthetic View-paint faults are explicitly
separated from the real input steps; all must leave the original SVG unchanged.
The screenshot is a sandbox result, not an installed-product or Word acceptance.

The full legacy SVG View runner initially failed because it aliased svg-view.js
as /view.js while omitting its new relative paint.js dependency. The runner now
serves current /host modules at their real paths. The actual 18 zoom/DPR/scroll
checks and the retained fixed-fixture vector/raster comparison were rerun without
changing their assertions. This is recorded as a test resource-closure correction,
not an absent Devspace result or a waived geometry regression.

This slice is not all of xcolor. Custom color definitions, current-color '.',
CMYK/HSB/named-palette imports, scientific-notation numbers, transparency, gradients,
color profiles, page color, arbitrary colorbox/fcolorbox text boxes and document
preambles are not covered. Broader math syntax/typography remains subject to P07
capability inventory and later integration tests, not a universal TeX guarantee.

Primary semantic reference consulted 2026-09-13: LaTeX Project xcolor source,
https://github.com/latex3/xcolor/blob/main/xcolor.dtx, especially color group scope,
textcolor's implicit grouping, base colors and standard color-expression mixing.
This reference defines expected semantics; it is not evidence of implemented
product coverage. No third-party math editor/parser/layout engine is invoked.
