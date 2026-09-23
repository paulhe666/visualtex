# CP08 — Cases + Aligned contained-environment checkpoint 1

Captured 2026-09-14 +08:00 from checkout `reference-3d207d7` at base HEAD
`00b48e17b736c7e57a0a72e19e271df82906f89d`. This is a non-Git source
checkpoint. No product switch, Office integration, commit, push or release is
claimed.

## Stable source receipt

The reproducible P08 grid runner passed on unchanged engine inputs:

- verification: `engine/target/matrix-verification/2026-09-13T23-46-14-833Z/report.json`
- status: PASS
- source SHA256: `09b5a54aca8d81b131cb2c8653b6e71e94b022131916df5fa525661aaa431321`
- sourceUnchanged: true
- seven runner steps PASS
- full Rust workspace total reported by the build step: 431 tests PASS

P00 protection remains a separate red gate because tracked/untracked Office/VSTO
and evidence files are being changed by another task. No protection baseline or
Office file was changed to make that guard green.

## Cases capability actually verified

`cases` is a real two-column Grid environment in one formula/session. It has one
stretchable left brace built from font geometry, identified rows/cells/columns,
real empty cells and lossless source round-trip. `InsertCases`, internal row
insert/delete, Tab navigation, Shift+Enter nearest-environment insertion, ordinary
Cases Enter behavior, rectangular selection/copy/paste, undo/redo and source edits
all use the same engine owner. Nested fractions, matrices and gathered formulas are
not cut to create rows; the nearest Cases receives the operation.

Column insertion/deletion is rejected for Cases because its two-column schema is
fixed. Raw cross-cell ranges cannot silently promote into deleting/copying the
whole Cases object.

## Aligned / align / align* capability actually verified

Authored `aligned`, `align` and `align*` spellings parse into the same shared grid
machinery while preserving environment identity/spelling. Columns are paired by
alignment markers. Each row shares the same column identities and computed widths;
layout does not use CSS margins or separate editor owners.

The common form `a &= b+c \\ d &= e+f` has one right edge for the left column and
one left edge for the right column. Multiple alignment groups including authored
empty cells from `&&` are supported. The right-hand column receives the required
math context without changing unrelated Matrix/Cases rows.

Shift+Enter at an aligned row end creates another row inside the same environment.
At an interior cell boundary it structurally splits the current aligned row,
preserving corresponding suffix columns and mathematical node identities. Row
boundary Backspace/Delete merges only within the same aligned owner. A caret inside
a nested fraction/root-like structure adds to the nearest aligned environment
without cutting that structure. Undo/redo restores exact source, selection and
internal identities.

## Actual Native/WASM evidence

`engine/target/matrix-parity-review/2026-09-13T23-46-40-001Z/report.json` PASS:

- 2423 actual Native/WASM requests
- 498 validated scenes
- 586 fresh-parser source oracles
- maximum numeric difference: 0
- exact SVG parity: true
- 24 grid source families, including 7 aligned/align/align* and 5 Cases formulas
- 32-row aligned and 64-row Cases samples
- nested/hidden/color/Unicode/comment inputs and four math styles
- Cases: one-owner insertion/history, 32 internal row additions, fixed-column
  rejection, rectangular clipboard and independent-copy identity checks
- Aligned: 20 new rows with identical equality anchors, per-column split/merge,
  10 interior split/merge cycles, rectangular clipboard, source edits and copied
  identity independence
- independent frame B remained unchanged throughout the dedicated transactions

Native test CLI SHA256:
`adc3e728c77b683eb6ccb746e91fb093806e2f38df97cb3a6702256867cc5967`

Node WASM SHA256:
`88a76d6e38fdd7c7f4486d8fd67f1042503ad69abc4abd3509ae064454d8988d`

## Actual Edge browser evidence

`engine/target/matrix-browser-review/2026-09-13T23-46-56-411Z/report.json` PASS:

- Edge `152.0.4191.62`
- sandbox host build `sandbox-03f6894e487a47063f76a5df`
- 115 recorded input events, 112 trusted
- Cases: 36 real cell-pointer checks across zoom .5/1/2 and DPR 1/1.25/2
- Aligned: 72 real cell-pointer checks across the same zoom/DPR matrix
- trusted keyboard Tab, Shift+Enter, ordinary Enter, undo/redo and row merge/split
- nested Cases/Aligned nearest-environment insertion
- actual browser clipboard replacement for aligned/cell selections
- protocol composition cannot create a duplicate aligned row
- frame B entire editor snapshot remained unchanged
- `unexpectedHostErrors = 0`
- rectangular drag/clear/undo passed
- vector export remained one SVG object with no selection overlay
- screenshot: `engine/target/matrix-browser-review/2026-09-13T23-46-56-411Z/matrix-two-independent-frames.png`

The separate real rectangle browser review also passed:
`engine/target/grid-rectangle-browser-review/2026-09-13T23-47-05-552Z/report.json`
with 19/19 trusted input events, real copy/cut/paste, expected dimension refusal and
unchanged frame B.

Gathered parity/browser and Grid host-boundary fault injection also passed in the
same stable-source runner. The original Matrix and rectangle tests remain active;
Cases/Aligned were added rather than replacing them.

## What is NOT claimed

This checkpoint completes the implemented M05 nearest-environment behavior for
Gathered, Matrix, Cases and Aligned-family environments in the isolated engine.
It does **not** mark all P08 or the product migration complete.

Still outside this checkpoint: implicit ragged cells, aligned placement and row-gap
options, P08 long-document viewport virtualization, physical Windows IME acceptance
for these new environment operations, installed VisualTeX/main-application switch,
real Office Ribbon acceptance, P10 integration, P15 removal of MathLive/MathJax, or
any Office/VSTO modification. Production continues using the old engine path.
