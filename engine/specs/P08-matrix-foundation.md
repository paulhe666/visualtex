# P08 Matrix foundation — implementation contract and remaining scope

This extends P08-contained-multiline.md M05 for **matrix**. It does not mark
Matrix, Cases, Aligned, all of P08, or production integration as complete.
Actual results live in timestamped target receipts and independent checkpoints.

## Ownership and structure

One Grid(kind=Matrix) belongs to the existing EngineSession/document/formula.
A GridRow owns independently identified bare Row cells. GridColumnId is typed
metadata allocated monotonically with node IDs; it is not a caret or session ID.
No row or cell creates another EngineEditorHost/Mathfield/Office session.
Unchanged rows, cells, columns and mathematical children retain their identities.
Duplicate whole matrices allocate new node and column IDs. Undo/redo restores
original source, selection and internal IDs through the existing history owner.

## Parser/source scope

The implemented spelling is `\begin{matrix}...\end{matrix}`. Token-scoped `&`
and `\\` delimit cells and rows only at the current environment level. Comments,
escaped `\&`, text, brace groups and nested environments do not split outer cells.
Physical source newlines remain trivia; they never create another formula frame.

This foundation requires rectangular **authored** rows. Explicit empty cells
(e.g. `a&&b`) have real carets. Ragged rows requiring implicit cells remain
lossless unsupported source; no separators are invented in saved input. Row-gap
options, starred breaks, pmatrix/bmatrix/etc., smallmatrix, cases and aligned
remain outside this slice. Initial literal `[a]` in a matrix is not an option.

Visual edits update the mathematical tree and clear changed origins. Existing
projection serializes changed ancestors while preserving untouched spelling and
comments. No regex whole-document row split or alignment normalization is used.
Every acceptance edit is compared to a fresh parser oracle.

## Editing policy

- InsertMatrix(rows, columns) creates a rectangular grid; selected row content
  moves into its first cell retaining mathematical child identities.
- Shift+Enter in a matrix inserts an empty row **after the current row**, with
  the same column count, and places the caret in the same column. It does not
  divide the active cell or damage a nested fraction, radical, script or text.
- The nearest current environment owns Shift+Enter. An inner matrix changes
  independently of its enclosing matrix. An inner gathered retains its own
  gathered split policy; it does not add a row to the outer matrix.
- Insert/delete row and column are atomic transactions. The final row or column
  cannot be deleted by these commands; select the complete matrix to remove it.
- Tab/Shift+Tab navigate in row-major order, clamped at the environment edges;
  they never insert a row, change source/history, or focus another formula.
  Up/Down targets the same column of the adjacent row using actual scene carets.
- Backspace at a cell start and Delete at a cell end are contained no-ops. Matrix
  columns never inherit gathered's single-cell line-concatenation behavior.
- IME/preedit and pending command confirmation take priority. Plain Enter keeps
  the existing outer callback, separate from contained row insertion.

Rows are bounded at 256, columns at 64, subject to the shared 8192-identity and
64-depth budgets. Invalid dimensions, stale positions and exhausted resources
must fail atomically without consuming live identities or history.

## Selection and clipboard boundary — still incomplete

Single-cell mathematical content uses existing range editing/clipboard. A whole
matrix can be explicitly copied, pasted or duplicated as one LaTeX fragment.
**Rectangular cell selection and rectangular paste are not implemented.**
Partial cross-cell edits/copies are rejected rather than silently promoted to
operating on the entire matrix. The refusal must leave source, selection,
revision, history and the second formula unchanged, and the host stays usable.
This guard is a safety boundary, not a substitute for the required rectangle
selection/clipboard implementation in the next Matrix slice.

## Layout policy

The grid measures all cell logical/ink extents and solves shared column widths.
Cells are text-style math. Each row has minimum 0.8/0.2em ascent/descent struts;
clear row gap is 0.3em; column gap is 1em. These are documented engine defaults,
not fabricated font MATH constants. Taller mathematical content controls its row.
The complete matrix is centered around the ambient mathematical axis.

Column anchors carry grid/column identity, index, x and width in natural em
coordinates. SVG, hit testing and navigation use the same layout. Font size
converts natural units once; zoom/DPR never change exported dimensions. Cell
subtrees are independently cached; constrained row fragments are assembled by
the grid, not reused with obsolete context-free column widths. No CSS margins,
DOM table measurement, synthetic glyphs or per-cell rendering engine is used.

## Required evidence and explicit non-coverage

`matrix_foundation_review.rs` covers real core actions, identities, source
projection/fresh parse, dimensions, boundaries, copies, history and layout.
`matrix-parity-review.mjs` compares actual Native/WASM requests and exact SVG.
`matrix-browser-review.mjs` exercises two actual browser hosts with trusted
keyboard/pointer and visible controls, nested matrix/fraction insertion, empty
cells, row/column operations, same-column navigation, zoom/DPR and protocol IME.
`matrix-verify.mjs` bundles these with old gathered and grid-host regressions;
it does not override P00 protection or claim physical Windows IME/Office UI.

Next required slices: rectangular selection/clipboard, implicit ragged-cell
source policy and broader matrix variants, then Cases and Aligned/M05 including
shared `&`/`&&` anchors, cross-row split/merge, copy/selection and all parity/UI
acceptance. P10 production switch and P15 dependency removal remain later gates.
