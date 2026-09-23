# P08 — identified cell rectangles and source-only clipboard

This extends the matrix foundation without changing any production/Office host.
The current cell rectangle is an internal Selection variant with gridId,
anchorCell and focusCell. Endpoints are the IDs of actual bare cell rows under
ONE grid; inner descendants, grid rows, another environment and zero/stale IDs
are rejected. Direction is preserved. The normalized rectangle uses inclusive
row/column bounds derived by the core, never cached source offsets.

## Interaction and ownership

Dragging across cells of the same nearest grid produces a cell rectangle.
Dragging within one cell retains ordinary math/text selection. Cross-environment
raw ranges still fail the existing destructive-fragment guard; a rectangle does
not promote to deletion of the outer matrix. Ctrl+Shift+arrow extends the current
cell selection. At grid edges it is bounded; a one-cell rectangle is legal.
Escape collapses to the focus cell. Plain horizontal arrows collapse/extend using
core positions. Tab collapses the rectangle to its focus before cell navigation.
All pointer and keyboard actions use the same session and revision. Selection
changes do not create content history. Pending browser input/IME remains first.

The scene query produces one highlight rectangle per selected cell using its
shared column anchor and actual row extent, including empty cells. The host only
paints that geometry. Selection paint is not part of SVG export.

## Transactions and clipboard

Delete/Backspace/ClearGridCells clear the selected cell contents, retaining grid,
row, column and cell IDs. Unselected source spelling and mathematical IDs are
unchanged; source origins are correctly rebased. Empty content is a no-op.
Other structural/typing commands on a rectangle are explicitly refused until a
specific replacement policy is defined. They never panic or flatten the grid.

Copy returns source-only `\\begin{matrix}...\\end{matrix}` cell data using the
existing schemaVersion=1 clipboard envelope. It contains no live edit identities.
Cut copies first and then clears through the normal revision/selection guard.
PasteLatex into a cell rectangle dispatches to PasteGridCells. Payload dimensions
must exactly match the rectangle. Explicit PasteGridCells from a caret starts at
its nearest cell and must fit the existing dimensions; it does not grow columns,
truncate content, repeat values or wrap a second matrix around the destination.
A scalar expression is a 1x1 payload. Destination cell IDs remain fixed; pasted
mathematical descendants and nested column IDs are fresh. Unsupported, malformed,
overflowing and wrong-size payloads fail atomically. Outer clipboard comments
outside a matrix payload are refused rather than silently discarded.

Undo/redo restores source, selection and all internal IDs in one content item.
Source edits retain the rectangle when endpoint identities survive; if they do
not, the shared PositionMap safely maps its focus to a surviving valid position.
An event replay cannot apply a second paste/clear. A second session is not read or
modified by any rectangle transaction.

## Evidence and limits

Actual tests include all 81 ordered corner pairs of a 3x3 matrix, a 16x16 matrix,
empty cells, reversed rectangles, nested fractions/grids, source-parser oracles,
wrong-owner/shape failures, exact history, source edit mapping, browser clipboard
copy/paste/cut, keyboard expansion, and nine zoom/DPR combinations. Detailed
counts and immutable receipts are in the next checkpoint, not inferred here.

This does not claim Excel/HTML/table clipboard import, arbitrary ragged source
normalization, native physical Windows IME acceptance, Cases/Aligned acceptance,
long-document virtualization or installed Office/product integration.
