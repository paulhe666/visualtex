# CP08 — Matrix cell rectangles and clipboard checkpoint 1

Non-Git checkpoint captured after actual verification on 2026-09-13.
Branch reference-3d207d7; base HEAD 00b48e17b736c7e57a0a72e19e271df82906f89d.
This follows CP08-matrix-foundation-1 without overwriting that archive or any
previous checkpoint. It is not all-P08 completion or a production engine switch.

## Implemented slice

Identified GridCells selection uses gridId + anchorCell + focusCell. The core
resolves rectangle bounds and rejects wrong owners, nested descendants and stale
IDs. Same-grid cross-cell pointer dragging and Ctrl+Shift+arrow expansion select
whole cells. Single-cell math/text ranges remain ordinary selections. Escape,
horizontal navigation and Tab return to real cell carets without content edits.

Delete/Backspace clear cell contents only. Copy emits a source-only rectangular
matrix through the existing schemaVersion=1 envelope. Browser Ctrl+C/X/V are
actually exercised. Pasting into a rectangle requires an exact dimension match;
explicit caret-targeted cell paste must fit the current grid. No silent growth,
truncation, replication, nested-owner creation or whole-grid normalization occurs.
Destination grid/row/column/cell IDs remain; pasted descendants have fresh IDs.
Undo/redo restores source, selection and IDs exactly. Empty rectangle clear is a
no-op. Wrong-size, unsupported and malformed input fail atomically. A second
formula frame B stays byte-for-byte unchanged at each browser and parity step.

Scene selection returns one highlight per selected cell from shared column and
cell geometry, including empty cells. Export stays one SVG and excludes selection
paint. The old raw cross-cell destructive-range guard remains tested; only the
new explicit rectangular path is accepted. The prior browser refusal case was
updated to assert actual rectangle clearing, survivor IDs and exact undo instead
of treating the now-supported pointer gesture as an error.

## Actual evidence

Stable inputs SHA256:
3d01afaf954ec13d352cbac746174e8fac95b729cb5437d9c19eecda40dbb1a7

Matrix verification:
engine/target/matrix-verification/2026-09-13T15-55-50-846Z/report.json
7 steps PASS, sourceUnchanged=true. Full Rust suite 410 PASS, including 12 new
rectangle tests. Native cases include all 81 ordered 3x3 corner combinations,
16x16 cells, exact history, nested ownership, source edits, malformed shapes,
composition refusal, replay and independent scene queries.

Native/WASM parity:
engine/target/matrix-parity-review/2026-09-13T15-56-15-739Z/report.json
1204 requests, 248 scenes, 287 fresh-parser oracles; max numeric difference 0,
exact SVG. Includes 20 cell-clear/paste/history cycles in addition to existing
matrix sizes, nested mathematics, 32 row additions, 12 column cycles, 24 source
edits and separate formula B.

New owned Edge browser:
engine/target/grid-rectangle-browser-review/2026-09-13T15-56-33-107Z/report.json
9 checkpoints, 19/19 trusted recorded input events, actual trusted copy/paste/cut
clipboard events; 9 zoom/DPR rectangle cases. One deliberate wrong-dimension
paste is refused. No unexpected error. Screenshot cell-rectangle-two-frames.png
was read and visually inspected: four selected cells in A, separate z_2 in B.
Physical Windows IME is not claimed by this browser test.

Existing matrix browser:
engine/target/matrix-browser-review/2026-09-13T15-56-25-809Z/report.json
13 checkpoints and 81 cell pointer/zoom/DPR cases PASS. Gathered parity/browser
and 22 invalid Grid host-boundary injections PASS in the same bundle.

Strict original runner:
engine/target/verification/2026-09-13T15-56-50-545Z/report.json
12 engine/build/browser steps PASS; only original P00 protection FAILED due to
external Office/VSTO/evidence changes. Stable engine inputs match both runners.
The guard and baseline were not changed and the failure remains red. No Office
acceptance or installation is inferred from engine-only test success.

## Limitations / next stage

Cases and Aligned are next; their nearest-environment behavior is not implemented
by this checkpoint. Implicit ragged cells, matrix delimiter variants, arbitrary
spreadsheet HTML/TSV import, inherited-format clipboard expansion, long-document
virtualization and remaining P09-P15 integration are not claimed. Formal Windows
IME/product and Office acceptance remain separate. The main application still
uses the original engine. No user Word document was opened/saved/closed here.

No Git add/commit/push, branch/ref change, production edit, dependency install or
font installation occurred. Only engine source and non-Git checkpoint artifacts
were written. Archive excludes fonts, generated target output and binaries.
Compilation/test harness failures preceding the final run were not deleted.
