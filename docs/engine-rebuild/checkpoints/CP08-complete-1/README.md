# CP08 — contained grids, multiline editing and long-document virtualization complete

Captured 2026-09-14 +08:00. Checkout `reference-3d207d7` remains at local HEAD
`00b48e17b736c7e57a0a72e19e271df82906f89d`. This is a non-Git source checkpoint,
not a production editor switch, Office release, commit or push.

## Status

P08 is complete for the isolated Rust/WASM engine and sandbox host scope defined by
`VISUALTEX_ENGINE_PROJECT_PLAN.md` and the P08 specs. Gathered, Matrix, identified
cell rectangles, Cases, Aligned/align/align*, explicit alignment anchors, physical
Windows IME priority and 1000-formula viewport virtualization all have actual
passing evidence.

The production VisualTeX main editor is still unchanged and remains a later P10
gate. Office/VSTO/COM/OLE/session contracts were not modified or exercised by this
P08 completion. The original P00 protection remains red because of unrelated
parallel Office/evidence workspace changes; no guard or historical baseline was
relaxed, reverted or overwritten.

Stable qualifying receipt:
`engine/target/matrix-verification/2026-09-14T00-26-43-612Z/report.json`

Source SHA256:
`046a954731268fca68d80675631488bdfec56d671fcc278f1d9500aff483ab46`

`sourceUnchanged=true`; all 9 P08 runner steps PASS.

## Implemented P08 behavior

A contained multiline formula always has one outer formula frame and one
EngineSession/document/formula identity. Internal Grid/GridRow/cell/column IDs are
mathematical editor identities only. No row uses a separate editor owner and no
adjacent formula frame is borrowed for alignment.

- Gathered: root Shift+Enter promotion and continued internal rows.
- Matrix: authored rectangular cells, 2x2/3x3/large matrices, empty cells, row and
  column CRUD, Tab/navigation, rectangular selection, clear/paste and exact history.
- Cases: fixed two-column schema, real math-font left brace, internal ordinary Enter,
  Shift+Enter and nearest-environment insertion without cutting nested structures.
- Aligned/align/align*: paired shared anchors for `&` and multiple groups including
  `&&`, per-column split/merge, nested fraction/root structures and same-owner rows.
- Clipboard: rectangle source carries no live session/node IDs; copied descendants
  receive fresh IDs while surviving destination cell/column IDs remain stable.
- Source: physical newlines remain trivia; environment boundaries are token scoped;
  comments, escaped ampersands, text and nested environments do not regex-split rows.
- SVG/layout: Rust owns widths, row baselines, hit regions and explicit grid anchors;
  zoom/DPR does not rewrite source or natural geometry.

## Long-document virtualization

`EngineEditorSessionOwner` separates core state ownership from mounted
`EngineEditorHost` DOM. `parkSession` / `resumeSession` move the same complete
EngineSession between active and parked runtime maps without reconstructing from
LaTeX. NodeIds, selection, history, event receipts and handle stay intact. Layout
and scene caches are dropped while parked.

Active sessions remain capped at 64. Parked sessions use a separate bounded
count/weight budget. An active composition cannot be parked. Destroy releases both
active and parked ownership.

Actual real Edge review:
`engine/target/virtual-viewport-browser-review/2026-09-14T00-27-39-840Z/report.json`

- 1000 passive formula shells.
- Every formula actually materialized at least once.
- 219 real scroll samples.
- Peak mounted hosts = 11.
- Peak textarea inputs = 11.
- Peak SVG views = 11.
- Peak active sessions = 11.
- Peak parked sessions = 992.
- Formula 0 resumed with the same handle and full editor snapshot.
- Undo/redo after resume restored the same mathematical identities.
- Composition pins an offscreen formula until preedit ends.
- No formula focus => zero visible carets; one focused formula => one visible caret.
- Final dispose returned active/parked sessions and editable DOM counts to zero.

## Actual executed acceptance

1. Full engine build and regression: 436 Rust tests PASS. Native CLI/WASM builds and
   bindings rebuilt. Existing protocol/source/edit/history/layout/symbol regression
   remained green. Protocol parity includes park/resume and is 2156 cases.
2. P08 Native/WASM parity:
   `engine/target/matrix-parity-review/2026-09-14T00-27-08-442Z/report.json`
   2423 requests, 498 scenes, 586 fresh-parser oracles, maximum numeric difference
   0 and exact SVG parity. Includes large Matrix, Cases and Aligned/align/align*.
3. Real Edge Grid browser acceptance:
   `engine/target/matrix-browser-review/2026-09-14T00-27-24-054Z/report.json`
   Edge 152.0.4191.62, 115 input events / 112 trusted, frame B unchanged and zero
   unexpected host errors. Cases and Aligned zoom/DPR cell hit checks passed.
4. Real browser rectangle clipboard acceptance:
   `engine/target/grid-rectangle-browser-review/2026-09-14T00-27-33-251Z/report.json`
   19/19 trusted events, frame B unchanged, real copy/cut/paste and dimension refusal.
5. Real Edge 1000-formula viewport acceptance:
   `engine/target/virtual-viewport-browser-review/2026-09-14T00-27-39-840Z/report.json`.
6. Physical Windows IME acceptance:
   `engine/target/grid-native-ime-review/2026-09-14T00-27-51-098Z/report.json`
   PASS in UUID/PID-owned Edge using Win32 SendInput ASCII pinyin and guarded
   Shift+Enter. No CDP imeSetComposition, Unicode injection or clipboard. Cases and
   Aligned both passed preedit priority and environment-specific Enter behavior.
7. Gathered Native/WASM and two-host Edge acceptance remain green.
8. Grid view fault-injection boundary runner remains green.

## Preserved failures

Earlier unique receipts remain and are not rewritten as successes. They include:
inline-module CSP rejection in the first native-IME runner; Windows foreground-lock
rejections where the helper explicitly sent no keys; an intentionally too-strict
IME candidate-source oracle; a wrong test comparison across full replaceSource ID
regeneration; and virtual-scroll fixture/sampling mistakes. The final PASS did not
remove or reinterpret those failures.

The owned IME helper was extended with guarded Shift+Enter and a PID/UUID-validated
window activation fallback. SendInput still refuses to run unless the exact owned
Edge window is the actual foreground window. It never targeted the user's Word or
browser window.

## Phase boundary

P08 completion does not claim P09 command-state/macros/custom-symbol integration,
P10 production main-app switching, P11 MathML/Office compatibility, P12/P13 Office
session integration, P15 MathLive/MathJax runtime removal, or installed Office UI
acceptance. Those remain separate later phases.

No Git add/commit/push, branch switch, pull/rebase/merge, reset/clean/stash, Office
document save/close or production engine switch occurred in this checkpoint.
