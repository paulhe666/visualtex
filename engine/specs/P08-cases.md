# P08 Cases — one-owner two-column environment

This describes the Cases slice built on the existing Grid/GridRow/cell model.
It is not a declaration that all P08 or installed Office acceptance is complete.
Actual test receipts and archive provenance belong in the independent checkpoint.

## Structure and source

`\begin{cases} value & condition \\ value & condition \end{cases}` owns one
Grid identity, two typed column IDs, independently identified GridRows, and two
editable Row cells per row. Both columns use text-style math and left alignment.
Natural geometry and anchors come from Rust. The left brace uses an actual
stretch construction from the registered math font and is owned by the Grid;
there is no additional editor, implicit Fence owner, CSS border or text glyph.

The current accepted source has two explicitly authored cells per row. Empty
values and conditions are valid. Missing implicit conditions, dcases/rcases,
placement options and row-gap options remain preserved but unsupported. This
finite restriction must not be reported as complete amsmath compatibility.
Physical source newlines and comments do not create rows or formula frames.

## Editing

InsertCases(rows) creates the two-column schema in one history transaction.
Selected mathematical children move into its first value cell with their IDs
preserved. InsertGridRow/DeleteGridRow work inside the nearest current Cases;
Shift+Enter also inserts an empty two-cell row, preserving the current column.
Nested fractions/text are not split to perform this row insertion. A nearer
Matrix or Gathered retains its own command behavior instead of modifying an
outer Cases. Column insertion/deletion is rejected because Cases has two columns.
Tab/Shift+Tab navigate its real cells; cell-edge Backspace/Delete do not merge
values or conditions. The first/last boundaries cannot borrow adjacent frames.

Ordinary Enter inside Cases retains the environment's internal row behavior.
Ordinary Enter outside Cases retains the outer callback. Shift+Enter never calls
the outer callback, including when the structural action is refused. Browser
paragraph input follows the same policy. IME/pending input remains higher priority
in the existing input bridge. This does not imply installed Office Apply testing.

Identified rectangular selection uses the existing one-owner GridCells contract.
Copy emits a matrix-shaped source fragment (no duplicated Cases brace); clear,
cut and paste keep destination Grid/row/column/cell IDs and create fresh copied
child IDs. Unsupported raw cross-cell ranges are refused rather than promoted
into deleting the whole Cases. Ordinary content within one cell is still editable.
Undo/redo restore exact source, selection and internal identity, not old revision
numbers. Source edits are compared to a fresh parser, not only serialized strings.

## Required acceptance

Run cases_review.rs for source, shape, IDs, row editing, nearest-environment,
rectangles, invalid raw ranges, genuine font brace, left column anchors and
natural size scaling. Existing matrix-parity-review.mjs and matrix-browser-review.mjs
now include separately asserted Cases scenarios in addition to their unchanged
Matrix/Gathered assertions. Cases-specific checks must actually pass before the
checkpoint may claim Cases parity or trusted browser acceptance.

The Native/WASM suite must include nested and hidden/colored Cases, Unicode,
64 rows, four mathematical styles and two font sizes, repeat edits/history and
fresh-parser oracles. Browser evidence must include the visible Cases button,
Tab, Shift+Enter, ordinary Cases Enter, nested Matrix/Cases, trusted rectangle
copy/paste, two independent hosts and 36 cell pointer checks at zoom .5/1/2 and
DPR 1/1.25/2. Chromium-protocol composition is separate from physical Windows IME.

P00 protection stays unchanged and may remain red due to parallel Office/evidence
changes. Do not alter that baseline. No production UI or Office integration is
switched by this slice; no Git commit/push is authorized.
