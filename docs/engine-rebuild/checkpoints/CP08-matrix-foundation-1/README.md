# CP08-matrix-foundation-1

Status: **ENGINE_COMPONENTS_PASS_PROTECTION_ALERT**.
This is the first Matrix foundation checkpoint, not all Matrix/P08 acceptance.
No engine Git commit or push was made. Existing CP08-contained-review-1 and
all earlier checkpoints remain unchanged. No user Word document was accessed.

## Identity and stable source

- Checkout: `C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7`
- Branch: `reference-3d207d7`
- Base HEAD: `00b48e17b736c7e57a0a72e19e271df82906f89d`
- Devspace returned workspace: `ws_fc27a8d3d3` for the same actual checkout.
- Verified source SHA256: `959dc6b7b26c5e8955d052b9d0628523f6cc378421c317e29830a75d03200a24`
- Matrix verification: `engine/target/matrix-verification/2026-09-13T15-01-16-621Z/report.json`
- Existing strict verification: `engine/target/verification/2026-09-13T15-02-19-191Z/report.json`
- Both runs report unchanged, matching source inputs. Archive hash/size and
  file-by-file audit are recorded in `audit.json` after actual package audit.

## Implemented

`GridKind::Matrix` extends the existing single-session Grid/GridRow foundation.
Rows, bare Row cells and typed columns have independent identities. Authored
rectangular matrix source uses token-level separators, preserving comments,
escaped ampersands, nested environments and empty cells.

The shared Rust layout solves column widths from real cell extents and emits
natural SVG plus explicit typed anchors. It does not create per-cell editors,
use CSS margins for alignment, or modify production UI/Office protocols.

Real transactions support inserting a matrix, inserting/deleting rows and
columns, row-major Tab/Shift+Tab, same-column vertical navigation, independent
whole-matrix copying/pasting/duplication, source editing and exact history.
Shift+Enter adds a row to the nearest current matrix, including from an inner
fraction; the fraction is not split. Nested gathered retains its own policy.
Ordinary Enter remains an outer callback; IME/preedit has priority.

Matrix cell-edge Backspace/Delete never merges unrelated columns. Until typed
rectangular selection is implemented, partial cross-cell copy/edit is refused
atomically, rather than promoting to the entire matrix. A real browser drag
negative test found a repeated refusal via native beforeinput; the host now
consumes rejected Backspace/Delete once and remains immediately editable.

## Actual acceptance

- Rust workspace: **398 passed**, including **13 new Matrix foundation tests**.
- Existing protocol/source-session/source-edit/editor parities: 2151/222/440/473.
- Existing layout: 4546 requests / 3063 scenes, max numerical difference 0,
  exact Native/WASM SVG parity. Symbol parity: 2570 requests / 1891 scenes.
- Matrix Native/WASM: **724 requests, 227 scenes, 142 fresh-parser oracles**,
  max numerical difference **0**, exact SVG parity.
  `engine/target/matrix-parity-review/2026-09-13T15-01-40-788Z/report.json`
- Matrix browser: **13 checkpoints**, **81 cell-pointer/zoom/DPR cases**,
  37 recorded input events / 36 trusted, Edge `152.0.4191.62`.
  One expected cross-cell refusal, zero unexpected host errors. Frame B's entire
  snapshot remains unchanged; two sessions remain two. This uses a separately
  owned browser, not the user's browser or Office.
  `engine/target/matrix-browser-review/2026-09-13T15-01-48-721Z/report.json`
- Browser screenshot was read and visually inspected:
  `engine/target/matrix-browser-review/2026-09-13T15-01-48-721Z/matrix-two-independent-frames.png`
- Existing gathered parity/browser and 22 grid host-boundary faults reran PASS
  in the Matrix verification bundle. Broad P07 browser/input/SVG regressions
  also passed in the strict existing runner.
- Chromium-protocol composition was exercised. Physical Windows IME and Office
  UI were not exercised and are not claimed here.

## P00 protection remains red

The existing strict runner is **FAILED**, solely at its final protection step.
It records original working-byte differences in WordEquationNumbering.cs,
WordOmmlFormulaStore.cs, ThisAddIn.cs, WordFormulaFormatConversion.cs,
WordFormulaService.CopyPaste.cs, WordFormulaService.FormatConversion.cs,
WordFormulaService.cs, and remediation evidence target-word.json,
ui-actions.ndjson, ui-uia.json, plus corresponding group/untracked differences.
These are the parallel Office/VSTO/evidence changes identified at handoff.
None was restored, deleted, overwritten or incorporated into this engine archive.
The baseline, protection implementation and Git-transition acknowledgement were
not changed to make the guard green.

## Still required before Matrix/P08 completion

Rectangular cell selection/clipboard/paste, implicit ragged-cell source policy,
broader matrix variants, then Cases and Aligned including shared `&`/`&&`
anchors, selection and row split/merge. Long-document viewport requirements
remain part of P08. Do not report gathered + Matrix foundation as all P08 PASS.
See `engine/specs/P08-matrix-foundation.md` and `P08-contained-multiline.md`.
Production engine switch remains P10; MathLive/MathJax removal remains P15.

## Retained failures during this slice

The first Rust compile found a column-index usize/u32 mismatch and the first
full regression found the old matrix-unsupported corpus expectation. Both were
corrected, with the corpus strengthened by dedicated Matrix tests. The first
new cross-cell browser test failed with three error reports instead of one:
`engine/target/matrix-browser-review/2026-09-13T15-00-21-412Z/report.json`.
Its receipt is retained, and the final passing browser test includes the exact
refusal, unchanged state and subsequent successful editing assertions.

## Archive boundaries

The tgz is an engine-only source archive. Compiler output, WASM/EXE binaries,
node_modules, fonts, user configuration and Office/VSTO/evidence are excluded.
The explicit STIX Two Math 2.13 test fallback remains under engine/target with
SHA256 `f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72`;
no font was installed or included in the source archive.
