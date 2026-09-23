# CP13 Office batch complete 1

Status: `complete`.

`p14Permitted=true`.

This checkpoint freezes P13: the Windows hidden Office conversion path, redraw path and bulk-import formula renderer now use non-interactive VisualTeX core ownership while the P12 interactive Office editor remains core-owned and unchanged in behavior.

## Final verification

Final receipt:

`engine/target/p13-verification/2026-09-16T02-52-24-976Z/report.json`

Receipt SHA256:

`11ad3c491dd5c82958ad52ea848519b7602b2e69f753351d30e6508daf9c6f56`

Verified P13 source digest:

`dd278aead97f5a8a3f6aa6c7ca63c75ca219ad10a63b40b4e98d339e6d1e4713`

The verifier passed 9/9 gates with `sourceUnchanged=true`: Windows TypeScript, production Office bundle, P13 interactive-core guard, P13 headless-converter browser acceptance, P13 headless-converter performance, existing Office formula-document regression, existing Office Session safety regression, existing Office Apply-shortcut regression, and existing document-import parser coverage. The protected Office boundary was empty at verifier start and end.

The checkpoint uses the verifier's SHA256 manifest as its source freeze. It records all 10 verified source/input hashes and 31 real-Office evidence hashes in the receipt. No project binary archive was generated through a shell write path.

## Real Word acceptance

The frozen record is:

`engine/specs/P13-real-office-acceptance.json`

SHA256:

`4b4088a7931f0aa3eb250692b8de1f29236913bea0ff38465f0a6eb9855d4ef0`

The 100-formula corpus passed VisualTeX redraw and all six full-document format directions with zero failures. Structural audits stayed at 101 paragraphs and showed exactly the intended 100 objects after every leg: Word OMath for OMML, `Equation.DSMT4` for MathType and `VisualTeX.Formula.1` for VisualTeX. Selected VisualTeX → OMML, OMML → MathType and MathType → VisualTeX conversions also passed.

A separate mixed-content document passed 146/146 MathType → VisualTeX while retaining two unrelated inline shapes and one unrelated OMML equation. The same 146 VisualTeX formulas were then restored through the real Ribbon `重绘全文 → 全文 VisualTeX OLE 转为 LaTeX 代码` path; the unrelated objects remained.

The 1000-formula Word redraw passed with `unique=1000`, `skipped=0`, average core render 8 ms, max 49 ms, insert total 7233 ms and max 127 ms. A later strict pass verified all 1000 OMath objects were physically materialized. The subsequent return to LaTeX was an explicit Undo acceptance step captured by the `p13-core-office-1000-undo-*` evidence, not a spontaneous rollback.

Real bulk import used the visible `批量导入` window and `导入到 Word` button with `engine/fixtures/p13-office-bulk-import.md`. The preview identified 7 blocks, 1 inline formula and 2 display formulas. Word received 3 OMath objects, zero OLE objects and 8 paragraphs; the Chinese prose remained present in source order.

## Browser/headless boundary

`apps/windows/scripts/p13_office_interactive_core_guard.mjs` verifies that the interactive Office editor still mounts the VisualTeX core editor, has zero MathLive ownership, performs real-browser input/export, uses core Undo/Redo, rejects stale Session exports and commits through the existing Office Session handshake.

`apps/windows/scripts/p13_office_core_converter_browser_smoke.mjs` verifies that `vsto-convert` mounts no MathLive field, no core input element and no interactive editor. Native OLE, Word OMML, MathType OLE and cross-platform picture use the required headless core export contracts. OMML remains MathML-only. Unsupported input fails only its Session with no partial export, following Sessions continue, terminal writes remain ordered and unique, and equation tags are applied by the Office core compatibility wrapper.

`apps/windows/scripts/p13_office_core_converter_performance.mjs` passes 100 mixed formulas and 1000 OMML formulas with zero failures and shared core runtime/font resources.

## Frozen Office boundary

P13 does not modify WordVsto, PowerPointVsto, Windows Office Contracts, FormulaOleServer, VstoShared, WindowsOleBridge, Tauri Office Session/server/lifecycle/state/backend/pipe code, `src/office/shared`, `src/office/api`, or `src/office/dialog/useOfficeSession.ts`. At final verification, `git status --porcelain --untracked-files=all` was empty for all of those protected paths.

The baseline remains the P12 complete checkpoint:

`docs/engine-rebuild/checkpoints/CP12-office-interactive-complete-1/`

The branch is `reference-3d207d7` and HEAD remains `d1f0aff8db9275c90dc18c7318dc2a4d7bbbd760`.

## Runtime observation carried to P14

During exploratory acceptance, three non-deterministic debug companion failures were recorded with Windows exception `0xc000001d` at fault offset `0x114eace`. Binary mapping places the trap in a Tao `EventLoopRunner` `Rc` clone used by native `Window::new`. It was not reproducible under exact 1000-Session batching, idle prewarm, concurrent first window creation, 30 converter close/recreate rounds, or a final 20-round lifecycle stress on a Word-started background companion. Final full-document, selected, 146-formula and bulk-import acceptance completed without another crash.

The observation is therefore carried into P14 fault/runtime hardening. P13 does not claim that this rare runtime issue has been root-cause-fixed; it only records that the formula-core export, Session identity/order and real Word mutation acceptance all pass independently.

## Repository safety

No reset, clean, stash, branch switch, pull, rebase, merge, commit, push, installer overwrite or release action was performed while finalizing P13. Existing dirty/untracked files were preserved. No user Word document was saved, closed or modified by P13 acceptance; only dedicated owned test Word processes were used.
