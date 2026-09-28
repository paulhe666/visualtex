# Independent coexistence checkpoint — 2026-09-28 08:38–08:41 +08:00

This file is intentionally separate from `mathtype-coexistence-checkpoint-20260928.md`, which has received concurrent edits. It records directly verified evidence and a still-failing production path; it does not overwrite the other execution's CP4/CP5 or claim final acceptance.

Workspace: `C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7`, workspaceId `ws_6e742f6c2d`. Preserve all existing changes. No commit/push/reset/stash/branch changes. Evidence paths below use `E = artifacts/mathtype-install-20260927/coexistence-20260928`.

## Coordination gate: concurrent source and Word operations

At 08:35–08:36 this execution attempted an exact targeted edit against the previously read `WordFormulaHostResolver.cs`. Both attempts failed at the first oldText match and applied nothing. Re-reading showed an independently written paragraph-expansion implementation, new `WordOmmlLocalBoundaryChecks.cs` tests, a built R4 DLL, and additional CP4/CP5 text in the shared checkpoint. None of those changes were made by the failed parent-navigation edits in this execution. Their originating actor is not identified.

This execution then stopped product edits, builds, deployment and Word UI input. It only continued read-only inspection, independent offline tests and saved-evidence comparison. The user was asked to pause the other writer before resuming single-owner source/UI work. Do not overwrite or revert the concurrently changed file, and do not reuse the previous Word identity blindly.

Latest read-only state at 08:40:51:
- Installed Word DLL SHA256 `700460A0AA50FD4D1879D990A5DD6EABF3C12378EB64BF01B0AEDA4442317A5D` (R3).
- Built Word DLL SHA256 `5DAE6D55D0BCE336715A677BE465D4B72ECD003757C2883365B1C8FC47D9D35A` (concurrent R4).
- `WordFormulaHostResolver.cs` SHA256 `6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`, modified 08:35:07.
- Prior own test Word PID 528 is no longer present. A new Word PID 41416 started 08:39:54, titled `文档2 - Word`; it was not launched or claimed by this execution. Do not operate it based on the old state file.
- Protected Word 93104 still has its original 2026-09-27 23:57:22 start time. Native MathType 38260 and MathTypeLib 22688 still have their original 2026-09-28 00:48:22 start times. No termination of these processes was performed here.

## R3 work actually completed by this execution

R3 fixed two measured Ribbon-state defects while retaining the proven passive-event boundary: a selected OLE of unresolved identity may be explicitly inspected, but passive callbacks may not export/activate it or claim its native gesture. Target/window changes invalidate controls even when semantic font size is unknown. Ordinary prose still avoids redundant refreshes. The actual installed R3 was verified by hash and loaded MVID `9d2ffe9e-f9cd-4084-b073-b8235da9dfb7`.

Focused tests: seven existing passive/command cases and fifteen actual compiled Ribbon/full-event cases passed. Authoritative reports: `E/passive-boundary-r3-20260928.json` and `E/ribbon-boundary-r3-outargs-20260928.json`. The unchanged R2 exposed six Ribbon/target-state failures first. No historical test expectations were loosened.

Actual installed R3 Word, no debugger, no external document polling: Display/Right/Left insertions x=1..20, with Ctrl+F4 equation-only return. Eighteen of these are a continuous repeated native Ribbon sequence. `E/candidate-r3-no-debugger-20260928-081521/native-20-inspection.json` confirms twenty genuine decodable Equation.DSMT4 objects, thirteen native numbered paragraphs, and no VT bookmarks in `candidate-r3-native-20.docx`.

Actual VisualTeX UI then completed LaTeX -> MathType for `y=1/2+alpha`, explicit font increase to 11 pt, and MathType -> OMML. OMML -> VisualTeX FAILED; the failure is not a crash and is not a pass. The saved rollback fixture retains the complete OMML and the original twenty Equation Native payloads/native field instructions unchanged. See `E/candidate-r3-no-debugger-20260928-081521/omml-failure-inspection.json`.

## New confirmed defect: range-scoped OMath views are clipped

`E/candidate-r3-no-debugger-20260928-081521/omml-boundary-idle-snapshot.json` is one explicit idle-Word boundary diagnostic, not a UI acceptance or polling loop. It shows a complete physical equation at 2839:2850 but a selection at 2840:2851 whose `.OMaths[1].Range` is 2840:2850. That starts inside the mathematical italic y surrogate pair. Other local scopes similarly truncate the first or last symbol. Exporting the clipped scope gave one syntactically valid OMath while silently dropping y; clearing that same partial range correctly triggered the surviving-host guard.

The semantic/deletion defect is caused by treating a range-scoped OMath view as a complete physical host. Keep the deletion assertion and transactional rollback. Re-reading `.OMaths[1].Range` from the same clipped scope cannot establish completeness.

### The concurrent R4 does NOT close the actual noncollapsed-selection path

The concurrent R4 fixes `ResolveLocalOmml` by using full owning paragraphs, but leaves `CaptureOmmlIndex` reading the range-scoped view. `WordFormulaService.FormatConversionCore.cs` routes a noncollapsed selection through `CaptureScopeKindIndex`, not `ResolveLocal`.

Independent actual-compiled-DLL tests in `apps/windows/tests/WordOmmlBoundaryChecks.cs` / `Test-WordOmmlBoundary.ps1` model both paths. The model includes correct paragraph Expand behavior, so the R4 local path is tested fairly. Report `E/omml-boundary-r4-concurrent-scope-review-20260928.json` has FOUR failures:

| Case | Required physical host | R4 returned |
| --- | --- | --- |
| scope-start-inside-surrogate | 30:55 | 31:55 |
| scope-end-inside-surrogate | 30:55 | 30:54 |
| scope-deep-interior-excludes-neighbors | 30:55 | 41:43 |
| scope-two-partial-endpoints | 10:20,30:55 | 15:20,30:42 |

Local tests and full-scope controls pass. The unchanged R3 failed eight of ten (`E/omml-boundary-r3-before-20260928.json`). Therefore the other execution's thirty-seven passing local/passive/Ribbon tests are not evidence that the noncollapsed conversion is fixed. Do not ship R4 as the final repair or relax the four failures.

### More direct candidate established by an actual Word probe

`E/candidate-r3-no-debugger-20260928-081521/omml-boundary-parent-snapshot.json` adds `ParentOMath.Range` to the same six measured scopes. Every one returns the complete physical 2839:2850 equation, including when the child view starts inside a surrogate or ends before the final surrogate. Even the clipped child reports NestingLevel 0; nesting alone cannot tell whether its Range is complete.

This supports using a single local physical-OMath resolver based on the parent element in BOTH ResolveLocal and CaptureOmmlIndex, avoiding paragraph-wide discovery and guessed character padding. It is a candidate supported by this Word version's measured behavior, not a claim that all nested/root cases are already validated. Before deployment test nested equations, duplicate scoped views, adjacent hosts, REF exclusions, header/table stories, ambiguous selections and ordinary prose. Explicitly own/release returned COM references; do not use FinalReleaseComObject or weaken failure handling to force success.

No parent-based production edit was applied by this execution. Once single-writer ownership is resolved, merge only the necessary boundary correction with the current source, use a distinct candidate/version label, rerun the focused local+scope tests, and re-run the exact failed UI sequence.

## Native isolation evidence for the actual VisualTeX operations

`E/r3-visualtex-public-state-review-20260928.json` was generated at 08:38:26 by comparing saved snapshots before/after the VisualTeX operations, without intervening native insertion/edit:
- 1,382 public MathType/COM setting hashes unchanged, no added/removed values.
- 506 installed native files unchanged.
- 24 font registrations/file hashes unchanged.
- Four recorded Word/MathType process identities unchanged across the snapshot interval.

The bounded out-of-context SHOW/HIDE observer `E/candidate-r3-no-debugger-20260928-081521/vt-native-window-observer-1.json` ran 08:22:43–08:31:09 and recorded no native MathType window events; initial/final snapshots also show none. It uses no Office object model. Its own stop signal was written and it ended. A positive-control calibration of this observer has not yet been performed, so this is a bounded no-event observation, not an unlimited guarantee against every popup.

## Remaining gates

Single-writer/source/UI coordination first. Then fix both physical-OMML resolution paths, repeat actual installed native/VisualTeX interleaving, native and VisualTeX edit after save/reopen, and preserve unrelated native objects/numbering. Final installer build, actual complete-package install/retest with a new NSIS path/hash, completely font-clean/no-native environment, both installation orders, and the original native installer disappearance remain unfinished. There is no new final NSIS and no commit/push from this execution.
