# P08 Aligned — shared column anchors inside one formula

This is the isolated engine slice for authored aligned/align/align* environments.
Checkpoint receipts, not this specification, establish executed acceptance.
Production UI, numbering and Office Apply integration remain separate gates.

## Source, structure and scope

GridKind preserves the authored environment spelling: aligned, align or align*.
Grid/row/cell/typed-column identity remains the existing single-session model.
Each top-level authored & delimits a real cell. Consecutive && therefore includes
an empty cell; nested grids, braces, comments and text cannot split outer cells.
Physical source newlines are trivia, not outer formula records or extra hosts.

This slice requires equally many explicitly authored cells in each row. Odd cell
counts and explicit empty cells are accepted; implicit padding of ragged rows is
not implemented. Unsupported placement/row-gap options, alignat, tags and other
TeX page-level behavior are not silently converted. Source remains preserved with
explicit diagnostics. Supporting align spellings here does not imply automatic
TeX equation numbering or full amsmath compatibility; existing Office numbering
continues to live outside this engine.

## Layout

All rows participate in one Rust-owned width calculation. Column pairs alternate
right and left alignment in display math. The right cell is laid out with an
initial empty ordinary-atom context so leading relation/binary glue has the
correct mathematical role; the source tree is not modified to manufacture that
spacing. Contextual row fragments are built by the grid; ordinary child subtrees
retain normal identity/style/hash caches. Ordinary rows and Matrix/Cases retain
their own context-free math spacing.

Pair members meet at their shared boundary. A documented 2em natural gap separates
pairs. This is bounded intrinsic layout, not page-width justification. Every
column retains a typed ID and an explicit GridColumnAnchor with scene-em x/width.
Positions, SVG export and hit testing use that same geometry; zoom and DPR cannot
change these natural anchors or semantic font sizes. No CSS margin alignment,
per-row Mathfield, or per-cell editor owner is introduced.

## Editing and history

InsertAligned(rows,pairs) creates an even-column aligned environment in one edit.
Selected mathematical children move into the first cell without changing their
IDs. Each environment keeps one document/session/formula identity.

At the end of an aligned row, Shift+Enter adds a row and places its caret in the
first cell. At a direct cell boundary inside a row, entries left of the caret
remain in that row; the suffix moves to the same column of the next row. Later
columns move to corresponding new-row cells, earlier new cells remain empty.
Existing cells/rows/columns and moved mathematical child identities survive.
Color declarations and comments retain explicit scope-safe source boundaries.
A caret inside a nested fraction, script or text inserts a blank row in the
nearest aligned environment without cutting that enclosing mathematical object.
A nearer Matrix, Cases or Gathered retains its own rule.

Backspace at the first cell's start merges with the previous internal row;
Delete at the last cell's end merges with the next internal row. Content joins
column-by-column, retaining the left cells and moved mathematical children. Other
cell-edge deletes are contained no-ops. First/last environment boundaries never
read or merge another formula frame. Undo/redo restores exact source, selection,
and internal identity. PositionMap covers all old carets in native acceptance.

Ordinary Enter in Aligned remains the outer callback; Shift+Enter never calls it.
Cases ordinary Enter retains its explicitly different internal-row behavior.
IME/pending input has precedence in the existing input bridge. Tab navigates real
cells. Rectangular selection/copy/clear/paste uses the existing GridCells contract;
raw ambiguous cross-cell ranges are rejected rather than deleting a whole grid.
Individual column resize is Matrix-only in this slice. Aligned pair resize is
not yet implemented and must not be advertised as tested.

## Acceptance

aligned_review.rs checks all split boundaries and both merge directions, node
identity, exact history, multi-pair/empty cells, aliases, nested fractions,
rectangular clipboard, safe failures, color/comments, RHS spacing and shared
anchors in four styles with natural-size scaling and caching.

The maintained grid parity and browser runners contain separately asserted
Aligned scenarios in addition to retained Matrix/Cases/Gathered coverage. Parity
must run actual Native and WASM binaries and compare fresh-parser oracles and SVG.
Browser actions must use visible controls, trusted keyboard/pointer and clipboard;
expected coverage includes 72 two-pair cell-pointer checks across nine zoom/DPR
combinations, one-owner history, three-row Shift+Enter, nested environments,
source reopen, and the other formula frame's entire unchanged snapshot.
Chromium-protocol composition is not physical Windows IME acceptance.

P08 is not globally complete: implicit cell policies, pair-column editing,
long-document virtualization and later Office/MathML acceptance remain distinct.
No Git mutation, production engine switch or user Office document action is
permitted by this specification.
