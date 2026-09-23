# CP07.8 — authored symbols and explicit math-font fallback

Captured 2026-09-13 after 20:01 +08:00. Checkout remains reference-3d207d7 at
00b48e17b736c7e57a0a72e19e271df82906f89d. This is a non-Git source checkpoint,
not an installed application or release. No engine staging/commit/push.

## Status

Twelve build/engine/browser steps PASS on unchanged input sources. The final
strict P00 production-protection step remains FAILED because the working checkout
contains separate Office/source/evidence modifications. Original protection
assertions and baseline remain untouched; this is not a globally green gate and
does not authorize Office integration or claim completion of all P07.

## Implemented slice and regression

175 additional authored symbol spellings (Greek variants, relations, sets,
arrows, delimiters and physics), a dedicated InsertSymbol structural transaction,
and 22 Unicode-defined integral commands with editable upper/lower slots. Exact
source spelling, atom class, safe command boundaries, source edits, copied node
independence, deletion, PositionMap and shared undo/redo are exercised.

The first font test run had 7 passes and 3 failures: Latin Modern Math did not
contain digamma/Diamond or several integral codepoints. No test was skipped and no
symbol was changed to a visually similar replacement. Explicit ordered verified
fallback font IDs are now part of internal layout options and cache context.
Primary glyphs always win; selected fallback metrics and genuine outlines are
used before layout, and glyph-resource provenance is retained in the scene.
Empty fallback options preserve the old strict missing-font failure behavior.

The pinned optional STIX Two Math v2.13 fixture is held only in engine/target,
not installed or added to production. It and all other font files are excluded
from this source archive. Its preparation tool validates bytes and the existing
fixture hash without overwriting unexpected content. See P07-symbols.md for exact
hashes, source and fallback scope. ESINT-specific outline aliases remain explicit
unsupported cases, not presumed Unicode substitutes or completed features.

## Actual verification

Receipt: engine/target/verification/2026-09-13T11-59-47-320Z/report.json
Verified source hash: 053e923105d883d85745f13bd3c7314fca6f479556daabce0ce78262f6cffccf
Source unchanged during verification: true. Input files: 133.

- 311 native Rust tests passed, including 11 new symbol/integral integration tests.
- Existing actual Native/WASM comparisons passed: protocol 2151, source-session
  222, source-edit 440, editor/history 473 requests including 100 history states.
- Existing fixed-resource layout comparison: 3945 requests / 2689 scenes PASS.
- New authored-symbol/fallback comparison: 2383 requests / 1755 scenes PASS,
  175 atom command spellings and 22 integral spellings, all four mathematical
  styles and operator placement modes. Numeric difference 0 and exact SVG parity
  with the identical supplied resources.
- Existing K0/root/fence/accent/operator/style/spacing/color browser tests PASS.
- New symbol browser test PASS: 16 trusted browser input events, controls from
  empty input, glyph provenance, bounds, undo/redo, zoom, SVG and invalid-source
  recovery. It does not substitute a scripted editor callback for those actions.
- Input fault injection and SVG zoom/DPR/scroll checks PASS.
- Separate independent layout review: 519 requests / 360 scenes PASS, relative
  1e-9 numeric tolerance, exact SVG and identities.
- Separate Office contract snapshot: 11 sources, 3 interfaces, 10 routes and
  3 negative mutation tests PASS; writes 0.

Browser report: engine/target/symbol-browser/report.json.
Screenshot: engine/target/symbol-browser/symbols.png; actually opened and inspected.
At 200% zoom the long formula scrolls horizontally; viewport framing is not a
claim of every formula fitting on screen. Natural vector bounds are validated.
Browser build: sandbox-aafe48d2d9fb0a24dd731a0f, Edge 152.0.4191.62.

Native CLI SHA256:
d4b2ae828d603315264824a389765088a43c6670fca2bd5dae26a7f2a9f81e00
Node WASM SHA256:
2cc77ea1bb483d7a28eb7e140d65d9c023281b4fd3eadb51102139c3b556b9fd
Web WASM SHA256:
891aeddd3c4e3daae132f6d6625d0af7e1a12ebac6dcc95ca5a9c3490c96fa27
Symbol parity transcript SHA256:
299c4fd53366db8d3ccbe811e4d0bf1cfa6d4cd010c376733b3d52fa3957ff09

## P08 user requirement recorded, not implemented here

The master project plan P08 section and engine/specs/P08-contained-multiline.md
now specify default Shift+Enter to promote or add rows within ONE formula frame.
Do not merge neighboring frames; preserve outer session/formula identity and
single-transaction undo. IME/command confirmation comes first. Two-frame isolation,
nested-structure behavior, source save/reopen and independent copy tests are
explicitly required and currently NOT_RUN. No P08 keyboard shortcut or environment
node is claimed implemented by this checkpoint.

## Protection and remaining gates

The guard at 12:01:18Z reported WordEquationNumbering, ThisAddIn,
WordFormulaFormatConversion, WordFormulaService.CopyPaste/FormatConversion/main,
target-word.json, ui-actions.ndjson, ui-uia.json, windowsSource/wordVsto/vstoShared
and historical untracked path differences. They are outside this continuation's
write scope; no writer identity is inferred and no external content was removed.
The authorized historical Git transition was acknowledged normally.

No new physical Windows IME, Word/Ribbon, PowerPoint, installed WebView2,
MathML/Office compatibility, main application switch, installer or final latency
acceptance is asserted. Continue remaining P07 compatibility coverage, then P08
contained environments under the newly recorded constraint. Keep this archive
immutable; later source checkpoints need a different directory.
