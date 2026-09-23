# P08 independent contained-grid review

Observed 2026-09-13 21:38–21:43 +08:00 in ws_7e5a0b84a3 at reference-3d207d7 / 00b48e17b736c7e57a0a72e19e271df82906f89d. The implementation lane is actively introducing grids; this review lane has NOT edited its core or host files.

## Review-owned files

- engine/crates/core/tests/grid_foundation_contract_review.rs
- engine/tools/contained-multiline-browser-review.mjs

These are real public-core assertions and a two-host browser acceptance runner, not an alternative mathematical implementation. Preserve failures. They are new and NOT yet claimed to pass.

The browser runner compiles the current TypeScript host using createSandboxAssets. It creates two independent EngineEditorHosts only in a dedicated test page. Setup is scripted; edits use actual controls, browser keyboard and pointer input. window.__containedReview exposes read-only inspections. M01/M02/M03/M04/M06/M08/M10/M11 and parts of M09/M12 are asserted; composition uses Chromium protocol and must not be called physical Windows IME. Unsupported nested interior breaks are expected to fail atomically. Empty-line merge is required, not silently skipped.

## Executed results and exact boundary

The stable-source verification runner started at 13:38:31Z and failed in native compilation because source_edit.rs, layout/boxes.rs and layout.rs did not yet cover Grid/GridRow. Receipt: engine/target/verification/2026-09-13T13-38-31-722Z/report.json. A subsequent focused native attempt found only layout/boxes.rs and layout.rs still missing. This is an in-progress implementation result, not a rollback of CP07-general-operators. No P08 test pass is asserted.

At 21:42 the layout fingerprint and scene GridColumnAnchor fields had been added by the implementation lane. The current host still sent Shift+Enter to the ordinary onEnter callback; it was not yet wired to a contained-row transaction. These observations are timestamped, not assumptions about later code.

## Cases requiring attention

1. validate_tree should reject GridRow as a freestanding root/ordinary-row child. A row separator has meaning only under its owning Grid; otherwise projected source can look like a valid formula while losing the environment. The new orphan-row test makes this explicit.
2. A GridColumnId must share the monotonic allocation namespace without becoming an editable NodeId/caret. Repeated/colliding column IDs (including root and cell IDs) must be rejected.
3. Reopening all of begin/gathered/end including physical source newlines must use one document. Changing/undoing frame A must leave B's entire editor snapshot unchanged, not only its source text.
4. First Shift+Enter must be a single undo entry. Middle-row splits preserve both mathematical child identities and insertion affinity; row-start deletion cannot merge external frames.
5. New scene column anchors must be validated for finite coordinates, valid owning Grid/column identity and natural-unit consistency. Do not weaken existing negative-scene validation tests to accommodate them.

The two earlier attempts to write a P07-middle coordination note and modify model.rs were explicitly blocked by tool safety and did not land. No alternate mechanism was used to apply those changes. Independent review test creation did succeed. A combined syntax-check/focused-test shell invocation was also blocked before execution. Do not describe these isolated denials as Devspace connection failure or count them as executed tests.

No Git stage/commit/push, Office actions, package installation, original-document save/close or production switch was performed by this review. The original P00 guard and immutable checkpoints remain unchanged.

## Update 21:56 +08:00 — executed fixes and passing focused checks

Review now also owns grid_scene_contract_review.rs, contained_line_contract_review.rs, and the new private grid_scene_validation.rs. The only validator wiring is mod grid_scene_validation in lib.rs and its call at the start of scene_validation::validate_snapshot. Initial grid scene tests reproduced three failures: nonfinite x accepted, nonexistent grid owner accepted, and anchors outside the owner accepted. All six scene tests pass after validation; hidden grid anchors are constrained to their OWN grid, not incorrectly to the visible phantom export frame.

Focused native execution passed 12 contained-line + 9 grid-foundation + 6 grid-scene + 7 existing scene tests. The orphan-row ownership condition was fixed by the active implementation lane and passes without weakening the new assertion. Full build/native/WASM then passed existing suites, with temporary duplicate Grid match warnings after both lanes extended the same four semantic test helpers. This review removed ONLY its own added duplicate cases; the other lane's semantic cases remain. Rerun before checkpoint to verify clean warnings.

A real browser failure was reproduced after correcting this review harness's incorrect layout option (mathFallbacks, not symbolFallbacks). Correct input x+y=1 reached frame A, but Shift+Enter incremented the outer Enter callback and did not create a grid. This review added splitLine/mergeLine TS command types and a contained Shift+Enter branch in editor-host.ts. It consumes the chord even when an interior core split is rejected, preventing a fallback native textarea newline. Ordinary Enter/modified Apply behavior stays separate. Composition gating remains in BrowserInputBridge.

Actual two-frame browser receipt PASS:
engine/target/contained-multiline-review/2026-09-13T13-53-16-260Z/report.json
Host sandbox-52d0e746a07eb6693ce3cbef; browser WASM 74af0e7d59ae0a3c4b7894fd3e6b43821650fee3b5dfea8faab1f480bc6edcf5.
41 recorded input events, 40 trusted; one protocol-composition event is not physical Windows IME. A built three rows from blank via keys, undid/redid initial promotion with exact identities, split by real pointer at a middle root boundary, merged internal rows, rejected interior fraction split atomically, kept ordinary Enter separate and reopened complete source. Frame B's entire editor snapshot and the two-session count stayed unchanged at every checkpoint. The screenshot two-independent-frames.png was read and visually inspected: three rows in A, original z_2 in B.

The failure receipts at 13-51-15 and 13-52-05 remain; first is test-harness initialization mismatch, second is genuine missing host routing. No failed condition was simply deleted. New P08-specific Native/WASM parity and expanded zoom/hidden-grid checks are next; existing P07 layout parity alone is NOT proof of P08 parity. No immutable P08 checkpoint is claimed yet.

## Update 22:07 +08:00

P08-specific Native/WASM review PASSED at target/contained-multiline-parity-review/2026-09-13T13-58-01-892Z/report.json: 714 requests, 238 scenes, 144 fresh-parser comparisons, numerical difference 0, exact SVG. Includes 13 nested/hidden/color/Unicode/comment inputs in four styles/two sizes, 40 row additions and 40 source edits with second-frame state isolated. Both binary hashes were checked before/after the run.

Expanded two-host browser review PASSED at target/contained-multiline-review/2026-09-13T13-58-47-769Z/report.json. Eighteen actual pointer cases span zoom .5/1/2, DPR 1/1.25/2, and rows 0/2. Source/revision/natural dimensions/column anchors and all of frame B stay unchanged. This remains Chromium protocol composition, not physical Windows IME.

Four additional native tests in grid_source_property_review.rs passed, including 640 source disturbances (seed 0x08002026), physical newline distinction, unknown nested environment budgets and explicit unsupported row options. No source assertion was weakened.

Full verify-current started at 14:01:37Z: 385 Rust tests plus all 12 engine/build/browser steps PASS, strict protection FAILED on external production changes as before. sourceUnchanged was false because the implementation lane ADDED web/host/grid-geometry.ts during the run; a read-only rehash found ZERO changes among the original 156 recorded inputs and no removals. That run must NOT qualify a stable full-source checkpoint.

Independent grid-view-contract-review.mjs then reproduced three missing host-boundary checks: omitted anchor array on a grid, empty array on a grid, and duplicate grid owner regions were accepted. This review targeted grid-geometry.ts to require every grid owner to have anchors, validate unique region identities and avoid per-anchor linear lookups; it also checks ordered nonoverlapping columns. Corrected test receipt target/grid-view-contract-review/2026-09-13T14-05-57-662Z/report.json is 22/22 PASS. This is explicitly synthetic fault injection, not real keyboard input. Current host module SHA256 6b6fb7088784ca4b86bbc25eb18f986033a5b81e131bd1128fa79a93a009047d.

A new stable-source pass and checkpoint are now being prepared. Please retain both validators and independent tests; the current review is not modifying the ongoing parser/split/merge implementation.
