# CP08 — contained gathered formula, independent review checkpoint 1

Captured 2026-09-13 around 22:12 +08:00. Checkout reference-3d207d7 at
00b48e17b736c7e57a0a72e19e271df82906f89d. This is a non-Git source checkpoint,
not a product release, installed editor switch, or completion of all P08.

## Status

The supported gathered/Shift+Enter slice has actual native core, Native/WASM and
two-frame browser evidence. The strict original verification runner remains FAILED
only at its P00 protection step due to existing external Office/evidence changes.
Its source inventory is stable in this final run. No guard or historical baseline
was relaxed. Office integration remains gated.

Stable-source receipt:
engine/target/verification/2026-09-13T14-07-24-796Z/report.json
sourceSha256: 15c20bde3e153aaa94a494c17b36756f61e5d14159b4b4e990c256588b6e975b
sourceUnchanged: true; 158 recorded inputs. Separate rehash after supplemental
checks found no changed, added or removed engine inputs. No compiler warnings.

## User requirement actually exercised

One formula frame owns one session. Shift+Enter at a root mathematical boundary
promotes the current formula into a gathered environment and creates the second
internal row in one content transaction. Further breaks extend the same group.
The root/document identity and existing mathematical children are retained.
Undo/redo restores exact source, selection and internal identities. Breaks at
unsupported fraction/script/text interiors and noncollapsed selections fail
atomically, rather than slicing LaTeX or crossing into another frame.

The browser host consumes Shift+Enter separately from ordinary Enter/Apply, even
when the requested interior split is rejected, so native textarea newline fallback
does not produce another edit. Pending input is handled by the existing input
bridge. Row-start Backspace and row-end Delete merge contained rows only; first/
last-row boundaries never borrow an adjacent formula. Empty cells have real caret
positions. Full gathered source with physical newlines reopens as one formula.

## Implementation and review ownership

The concurrent P08 implementation lane added Grid/GridRow, typed column identities,
source parsing/projection, row split/merge PositionMaps, source edits, same-scene
layout and row navigation. This review reused that implementation, not a competing
math engine or a second host per line.

This review added 31 native tests across grid_foundation_contract_review.rs,
grid_scene_contract_review.rs, contained_line_contract_review.rs and
grid_source_property_review.rs. It reproduced and fixed three Rust scene-column
validation gaps via private grid_scene_validation.rs plus a one-line call in
scene_validation.rs. It also wired the actual Shift+Enter branch and TS commands.
The new host grid-geometry.ts received matching required-owner/anchor/identity
checks after three independent fault-injection failures. Existing assertions were
kept; four duplicate semantic test cases caused by concurrent additions were
removed only from the review's own duplicate additions, leaving exhaustive Grid
semantics in place.

## Actual executed verification

1. Full build: 385 native Rust tests PASS; native CLI, wasm32 and both bindgen
   targets rebuilt. Existing protocol 2151, source-session 222, source-edit 440,
   editor/history 473 comparisons (100 history states), broad layout 4546 requests
   / 3063 scenes and symbol/style 2570 requests / 1891 scenes PASS. Broad P07 tests
   are not substituted for the P08-specific results below.
2. All twelve existing engine/build/browser runner steps PASS on the stable-source
   receipt, including old K0/root/fence/accent/operator/style/spacing/color/symbol
   interactions, input faults and SVG zoom/DPR/raster checks.
3. New P08 Native/WASM parity PASS:
   engine/target/contained-multiline-parity-review/2026-09-13T14-09-20-113Z/report.json
   714 requests, 238 scenes, 144 fresh-parser oracles, maximum numeric difference 0,
   exact SVG. Covers nested/hidden/color/Chinese/comment inputs, four mathematical
   styles/two natural sizes, 40 line additions, 40 source edits, duplicate-event
   rejection, disjoint copied column/node IDs, history and independent frame B.
4. New real two-host browser review PASS:
   engine/target/contained-multiline-review/2026-09-13T14-09-42-239Z/report.json
   13 checkpoints; 41 recorded input events, 40 trusted; 18 pointer checks across
   zoom .5/1/2, DPR 1/1.25/2 and first/third internal rows. The complete B snapshot
   stayed unchanged at every checkpoint. Source, natural geometry and column
   anchors did not change during zoom. Screenshot two-independent-frames.png in
   that directory was read and visually inspected: three rows in A, z_2 in B.
5. New host grid-boundary review PASS:
   engine/target/grid-view-contract-review/2026-09-13T14-09-24-440Z/report.json
   22 cases, explicitly synthetic invalid-data injection. Missing/empty anchor
   arrays, NaN/Infinity, wrong owners, duplicate IDs/regions, bad indices and
   out-of-owner coordinates are rejected. Real hidden phantom grids remain valid.
6. Native source disturbances: 640 generated patches, seed 0x08002026, compare the
   actual edited tree against fresh parsing, not just its displayed source string.
7. Separate Office contract snapshot PASS: 11 sources, 3 interfaces, 10 routes,
   3 negative mutation cases, writes=0. This does NOT override the P00 alert.

Native CLI SHA256:
5b477cbcbad1f0c8534732d8d1c92f81f4a50e978c96b0c76acf415909c9df3b
Node WASM SHA256:
35136e711d616deca1c1365fb4e42fe871d1736f3539bc2722485193ec238754
Web WASM SHA256:
d4f31b22bf1b331f86f4c6d7c2e7764457a6d63b358ee47eab832fc193dac328
Browser: Edge 152.0.4191.62; host sandbox-9f2054b6616425a9b2548592.

## Acceptance limits and preserved failures

M01/M02/M03/M06/M08 and the gathered portions of M05/M10 have actual passing
coverage. M04 uses explicit interior refusal, not arbitrary inner-structure line
splitting. M07 covers the pure core and Chromium protocol composition; native
Windows IME and unfinished command-state acceptance remain separate. M09 covers
source reopen and native independent copying, not all system clipboard/file-save
routes. M11 verifies ordinary Enter separation, not installed Office Apply.
M12 has actual Native/WASM, browser pointer/zoom/DPR and same-scene SVG, not Office.

Matrix/cases/aligned, rectangular cell selection, arbitrary column operations,
long-document virtualization, complete command/macros/ESINT compatibility, MathML/
Office output and installed-product acceptance are still outstanding. The main
VisualTeX UI still uses its original engine. No new physical Windows IME run is
claimed; CP06 native-IME evidence is historical, not a test of this new chord.

Earlier failures remain in their unique output directories: initial native compile
while Grid wiring was incomplete; a test-harness wrong layout-option field; actual
Shift+Enter falling through to outer Enter; missing Rust/TS grid validation; and
sourceUnchanged=false in the 14:01 runner because a new TS file appeared during
that run. None was represented as a lost Devspace connection or silently removed.

The archive excludes fonts, compiled binaries and generated target outputs. Its
content hashes and correspondence to every verified source input are recorded in
audit.json. No Git add/commit/push, reset/clean/stash, branch changes, installation,
user Office document save/close or production source writes occurred in this review.
Next P08 slices should extend the existing grid, not create per-line editor owners.
