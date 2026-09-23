# P07.6 — explicit spacing, phantom boxes and boxed formulas

Scope: isolated `engine/`, source-preserving mathematical editing and SVG. This
is not permission to switch production UI, change Office, install or publish.

## Supported language and geometry

- Explicit mu glue: `\,`, `\:`, `\;`, `\!`, `\thinspace`, `\medspace`,
  `\thickspace`, `\negthinspace`, `\negmedspace`, `\negthickspace`.
  Natural widths are respectively 3, 4, 5, -3, 3, 4, 5, -3, -4, -5 mu.
  One mu is 1/18 of the current mathematical style's em.
- Explicit em skips: `\enspace`, `\enskip`, `\quad`, `\qquad` retain the
  surrounding semantic font size, rather than being scaled as script mu glue.
  The current engine uses fixed natural width; stretch/shrink glue, line-breaking
  penalties and user-redefined skips are not claimed.
- Glue is a distinct Space node. It is not an ordinary math symbol, does not
  change binary/unary classification, and adds to existing math atom spacing.
  Source spellings and comments remain lossless. Negative advance is retained;
  the export rectangle is normalized without clamping the mathematical advance.
- `\phantom`, `\hphantom`, `\vphantom` use a real editable body and the same
  math/font layout as visible content. Full phantom retains advance, ascent and
  descent. Horizontal phantom zeros ascent/descent; vertical phantom zeros advance.
  A paint-suppression flag is inherited during scene projection: no hidden glyph,
  glyph outline resource, fraction rule or nested box border is added to drawing
  or SVG. Hidden subtree regions have null ink. No opacity trick is used.
- Hidden content keeps source, mathematical node IDs, caret stops, accessible
  structural descriptions, keyboard navigation and shared history. The present
  UI does not draw an editor-only ghost preview of hidden content. Zero-axis
  boxes may have internal editing geometry outside their natural export box;
  such editing geometry does not enlarge publication/export dimensions.
- `\boxed{...}` has an editable body, display-style math and four vector rules.
  Its explicit engine border policy is 0.3 em inner padding and active MATH-font
  fraction-rule thickness. This is not a promise of byte/pixel identity with
  an arbitrary LaTeX preamble's `\fboxsep`/`\fboxrule` settings. Geometry includes
  real ink bearings to prevent border clipping. Zoom does not change this policy.

## Transactions and source integrity

InsertSpace replaces the selected range; InsertPhantom and InsertBoxed wrap the
selection or create an editable empty body. Empty wrappers remain non-exportable.
All use the original EngineSession/PositionMap/history, with atomic validation,
fresh copied IDs, undo/redo, no-op/replay rules and no old-engine fallback.
Unsupported commands and incomplete source remain drafts or explicit failures.

A prior P07.5 defect was reproduced by strengthening the source-patch regression:
`\mathbf{x}` patched to `y` retained `x` in the math tree even though its source
changed. The subtree replacement traversal omitted Style. Traversal now covers
Style, Phantom and Boxed; tests compare actual canonical structure, facts, every
source buffer identity/range and unaffected siblings, not only wrapper IDs.

A preexisting negative-advance corruption test initially failed after signed
spacing was enabled. The test was retained unchanged. Scene validation now binds
signed advance to both root logical width and directed root caret displacement;
legitimate negative glue is accepted, fabricated dimensional changes are not.

## Verification layers

`crates/core/tests/spacing_phantom_boxed.rs` covers roundtrip syntax, dimensions,
paint/resource exclusion, nested visibility, signed glue, class spacing, script
units, border geometry, copy/delete/history, wrapper typing, atomic failures,
source patching and 480 deterministic source disturbances (seed 0x706e2026).
The strengthened existing styles test is retained in `tests/styles.rs`.

`tools/layout-parity.mjs` adds all supported spacing commands, all phantom axes,
nested enclosures, legitimate negative extents, style-source content changes,
scene queries and operation/history comparisons on actual Native and WASM builds.

`tools/spacing-browser.mjs` uses owned browser controls and real key/pointer input.
It starts from an empty document, edits hidden content, wraps a true Ctrl+A
selection, inserts/deletes glue, checks undo/redo, changes zoom, exports SVG and
recovers invalid source. Inspection hooks are read-only. It does not claim native
Windows IME, WebView2, Word/Ribbon or installed-product testing. An initial test
used unsupported Ctrl+Home/End chords; this was corrected to the supported Ctrl+A
path and now explicitly checks that a real mathematical selection exists.

## Reference semantics (consulted 2026-09-13)

- LaTeX Project kernel, `base/ltmath.dtx`, phantom axis/empty-box mechanism:
  https://github.com/latex3/latex2e/blob/develop/base/ltmath.dtx
- LaTeX Project / AMS, `required/amsmath/amsmath.dtx`, math mu spaces and boxed
  display body: https://github.com/latex3/latex2e/blob/develop/required/amsmath/amsmath.dtx

These are semantic references, not evidence that the entire TeX language is
implemented. The Rust representation, navigation, layout and projection are
independent implementations. Colors, arbitrary dimensions/skips, redefinitions,
full macro support and broader P07/P08 structures are separate gates.
