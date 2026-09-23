# CP07.4 — tested operators, with an external-evidence protection alert

Captured 2026-09-13 around 16:21 +08:00. Checkout: `reference-3d207d7`, product baseline `00b48e17b736c7e57a0a72e19e271df82906f89d`. This is a non-Git source checkpoint. No engine files were staged/committed/pushed and no installed product was switched.

## Status — keep the distinction

Operator implementation and executed core/Native/WASM/browser checks PASS. The strict full verification runner is **FAILED at its protection step**, not globally PASS. Its alert was traced to new Word UIA evidence files outside this engine task's output paths. The protection script and historical baseline were not changed; no external file was deleted. This checkpoint is a recoverable tested engine slice, not permission to enter production/Office integration or claim all P07 complete.

## Implemented slice

Twenty-five authored large/named operators with real lower/upper row identities; interleaved original `limits`, `nolimits` and `displaylimits` spelling/order/trivia; explicit mode changes and Auto restoration; lower-upper-body navigation; same-operator vertical bound movement; copy independence; grouped powers of an explicitly selected whole operator; shared undo/redo; actual font display variants and MATH-based side/stacked geometry. Browser sandbox controls preserve the existing EngineEditorHost owner and use the same layout for display and SVG. Ambient display/text changes do not change source or content history.

Reproduced and corrected issues: selected whole operator powers previously became duplicate upper-bound source, vertical bound movement landed in the root row, and the sandbox retained an old export preview after a mathematical-style change. The corresponding regression assertions remain in the code. A geometry assertion was corrected only after measured coordinates proved that it double-counted the glyph's left ink bearing; final tests assert exact logical placement of both side bounds.

## Executed verification

Receipt: `engine/target/verification/2026-09-13T08-17-20-341Z/report.json`.
Verified source SHA256: `51e282c801425f805925398e4543ddfd437824b01ad143ab5d21e8ef204dbde3`.
All 112 recorded source inputs were separately rehashed at checkpoint time and remained identical.

- 258 Rust tests PASS, including 17 operator tests and 8 scene-query tests. Operator corpus covers 25 commands, four styles/four placement modes, source directives/trivia, fixed-font glyph/geometry, identities, history, resource failures and 600 seeded source disturbances.
- Real Native/WASM: 2151 protocol cases, 222 source-session cases, 440 source-edit requests, 473 editor/history requests and 100 history states PASS.
- Expanded layout parity: 2844 requests, 1944 validated scenes, maximum numeric difference 0 at the identical fixed resources, exact SVG PASS.
- Existing sandbox, roots, fences, input fault injection and SVG geometry checks all PASS in this receipt.
- Independent layout parity, executed separately: 519 requests / 360 valid scenes PASS, relative 1e-9 numeric tolerance and exact SVG/identities.
- Office contract snapshot, separately: 11 sources, 3 interfaces, 10 routes and 3 negative mutation tests PASS; no writes.
- `node tools/operator-browser.mjs`: PASS, 35 trusted browser events with actual controls/key/pointer input from blank, upper/lower selection, math-style and limits switching, history, vector dimensions and invalid-draft recovery. Report: `engine/target/operator-browser/report.json`.
- `node tools/accent-browser.mjs`: PASS again, 26 trusted browser events.
- Operator display/inline screenshots were opened and inspected. These are independent sandbox browser results, not actual Office or a new physical Windows IME run.

The earlier receipt `engine/target/verification/2026-09-13T08-14-51-148Z/report.json` failed a root pointer test and is retained. Diagnosis logged a click at y=526.605 in a viewport only 485 px high after toolbar expansion. `radical-browser.mjs` now scrolls the actual SVG into view, verifies `elementFromPoint` is in the editor, and sends the same real mouse input. The corrected click y=228.605 hit the SVG, and root degree edit/undo/redo all passed. No formula output assertion was weakened and no mathematical editor callback replaced the mouse action.

## Protection alert evidence

The strict protection step at 08:18:09Z reported only `historicalUntracked paths`; its 12 source/resource groups, 12 artifacts and 65 original dirty-byte checks had no reported mismatch. Later at 08:20:35Z, the read-only historical path reconstruction found exactly four additions:

- `docs/remediation-3d207d7/evidence/user-mt-baseline-ribbon-uia.json` (08:16:37Z)
- `docs/remediation-3d207d7/evidence/user-mt-edit-before-uia.json` (08:18:20Z)
- `docs/remediation-3d207d7/evidence/user-mt-edit-ready-uia.json` (08:19:37Z)
- `docs/remediation-3d207d7/evidence/user-mt-edit-ribbon-uia.json` (08:17:59Z)

Removing those four paths only from an in-memory comparison reconstructed exactly all 11629 historical paths and SHA256 `38044b8f9b9d54ab14e4f1b7bf83f55c6ff6a2ac30493b1bdb63f56e2e718d0e`. No filesystem removal occurred. This proves additive evidence at that observation time, not an absent returned tool result or a missing historical file. It does not identify which external process wrote the logs, and it does not waive future guard changes. Other concurrent work may continue adding or changing its evidence. Keep the strict alert visible and recheck source/Office protections before any integration.

## Build identities and remaining work

Native CLI: `f53687b2ecd6ec4f5b8c2da293b6352182796866cfcef2bee13216b9c5706201`.
Node WASM: `466f9f548770512390ed44940c27683739e67d3c2eb5195e4e2e39b56c8bd49b`.
Browser WASM: `472beac9f962e922c7804051cf7d673c497ff9c8283a9a2d65e129af1fd07a15`.
Current tested host: `sandbox-c19899f603d2b336f9fded44`, Edge `152.0.4191.62`.
No font or compiled binary belongs in the source archive.

Next independent P07 work: mathematical alphabet/style and function commands, then spacing/phantom/paint and remaining symbols; matrices/aligned/multi-line and complete command/clipboard/macros remain P08/P09. Full production UI integration, MathML/Office output, actual Ribbon acceptance, installers and MathLive removal remain later gates. Do not treat this checkpoint's passing component tests as a globally green protection gate.
