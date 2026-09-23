# CP08 Cases — two-column contained environment

Non-Git source checkpoint. The main application still uses its original engine;
this is not installed-product acceptance or completion of P08. Baseline supplied
by the handoff is reference-3d207d7 / 00b48e17b736c7e57a0a72e19e271df82906f89d.
The original protection runner acknowledges that Git transition and reports
external Office/VSTO/evidence differences. No baseline/guard was relaxed.

## Implemented

Cases reuses the actual Grid/GridRow/Row and typed column IDs. Two left-aligned
math columns share widths and geometry; an actual font stretch construction
supplies the left brace, owned by the same grid. InsertCases and row edits retain
one formula/session identity. Tab navigates cells. Shift+Enter operates in the
nearest current environment without cutting nested fractions or borrowing an
adjacent formula. Ordinary Enter inside Cases adds an internal row; outside Cases
it retains the separate outer callback. Pending input remains higher priority.

Identified rectangular copy/clear/paste preserves the Cases grid, rows, columns,
and destination cells; pasted descendants get independent IDs. Copying the cell
rectangle emits a matrix fragment, without a second Cases brace. Exact history
restores source, selection and internal identity. Raw cross-cell ranges are
rejected, preventing accidental promotion into deleting the whole Cases object.
Cases column insertion/deletion fails atomically because its schema has two columns.

## Actual receipts

Stable source hash:
222ebe613ffc3c9dafd59958517c5cc0b47c48a8eb5653c14c6718d76e7240d5

Full original runner:
engine/target/verification/2026-09-13T16-48-38-166Z/report.json
420 Rust tests passed, including 10 new Cases tests. All 12 engine/build/browser
steps passed. The final P00 protection step remains FAILED due to original working
bytes in Office/VSTO/evidence and historical-untracked changes. This is preserved
as ENGINE_COMPONENTS_PASS_PROTECTION_ALERT, never converted to a full green gate.

Grid acceptance runner:
engine/target/matrix-verification/2026-09-13T16-47-24-495Z/report.json
All seven steps passed. Existing Matrix/Gathered/rectangle checks were retained.

Native/WASM:
engine/target/matrix-parity-review/2026-09-13T16-47-50-116Z/report.json
1659 requests, 362 scenes, 388 fresh-parser oracles; maximum numerical difference
0; SVG exact parity. Includes separately asserted Cases geometry (64 rows,
nested/hidden/color/Chinese inputs, four styles and two natural sizes), Cases
insertion/history, 32 row additions, fixed-column refusal, rectangular clipboard,
20 source edits and independent copies. These counts include retained Matrix tests.

Browser:
engine/target/matrix-browser-review/2026-09-13T16-48-02-851Z/report.json
Real isolated Edge browser with two EngineEditorHost instances for TWO independent
formulas. A alone contains Cases; there are not multiple owners per internal row.
Cases checks cover the visible Cases button, Tab, Shift+Enter, ordinary Enter,
exact undo/redo, nested Matrix/Cases/fraction, actual trusted browser copy/paste,
36 Cases cell-pointer cases across zoom .5/1/2 and DPR 1/1.25/2, protocol composition
and a single vector export. B's full snapshot remains unchanged at checkpoints.
The same run retains 81 Matrix cell-pointer cases and previous Matrix behavior.
Screenshot matrix-two-independent-frames.png in that receipt folder was read and
visually inspected: f(x) and a two-row Cases in A, untouched z_2 in B.

## Failures preserved and resolved

cases_review.rs first compiled with an incorrect raw integer revision assertion;
this was corrected to the actual Revision value. Its first executable run then
reproduced the real raw-cross-Cases deletion bug (9 passed, 1 failed). The generic
fragment guard was extended to Cases; the unchanged failing assertion now passes.
A browser pointer assertion initially assumed the cell-root caret must win over a
coincident scripted-base caret. The actual hit was inside the intended cell.
That test now checks the exact intended cell's subtree membership and records the
returned position, while still forbidding other cells, grids or formula owners.
Initial failed receipts remain under their original unique target directories.

Some tool calls were blocked by the tool safety layer; they are not execution
or verification evidence. The uncreated standalone parity script was not used.
Cases-specific coverage is in the maintained common grid verification scripts.

## Remaining scope

Cases currently requires two explicitly authored cells per row. Missing implicit
conditions, row-gap/placement options and dcases/rcases remain preserved unsupported
source. Aligned/align have not been implemented or validated at this checkpoint.
Physical Windows IME, long-document virtualization, MathML/Office output and
installed UI acceptance remain separate. No production switch, Office document
action, font installation, Git commit/push or remote release happened here.
The archive excludes fonts, binaries and generated outputs. Audit verifies each
source input; prior checkpoints are not overwritten.
