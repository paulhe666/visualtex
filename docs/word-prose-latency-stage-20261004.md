# Word ordinary-text interaction latency — CP264/CP265, 2026-10-04

2026-10-06 follow-up: bulk text ReplaceAll is investigated separately in [word-bulk-text-replace-stage-20261006.md](word-bulk-text-replace-stage-20261006.md), CP312–314. Those real native replacement and in-process add-in results must not be conflated with the older out-of-process caret measurements below.

## CP265 installed-file comparison and count-scaling follow-up

The registered installed Word DLL (5702CC88C7CF5C795FE92D110AD336495CEE2E9F88FC77E5F207561CEC5D6575) was inspected read-only with the existing Mono.Cecil comparison tool. Its compiled OnWindowSelectionChange, caret normalization, selected-OLE classification and relevant passive numbering checks match the pre-fix source DLLFB031CF2. Entire HostResolver22 methods, MetadataReader15, HostLayout12 and SelectionPerformance5 also match. Differences elsewhere in ThisAddIn/WordFormulaService concern explicit operations; the only WordEquationNumbering difference is explicit PrepareVisualTeXDisplayHostLayout. This is compiled-method evidence, not a claim that every assembly byte/dependency is identical.

Registry-free isolated runs then copied the installed DLL and its installed dependency DLLs into an owned test bundle, without changing the installed directory. The20-OLE run passed11:15:30. Installed20 callback medians (ordinary arithmetic medians): empty37.1355ms, distant33.904, leading47.794, trailing63.783, nearby prose31.395. This confirms the unnecessary boundary path is present in the registered file as well as the source baseline.

A single owned Word process then grew0→20→100 real VisualTeX OLEs and passed all selected caret/typing/inventory/save assertions. This was local creation for a passive-interaction fixture, not formula-conversion benchmarking. In that same process:
- Ordinary distant prose:0OLE36.053ms median (n8),20OLE31.309ms (n12),100OLE28.860ms (n12)
- Inline leading boundary:20OLE48.7975ms,100OLE47.5385ms (n10 each)
- Inline trailing boundary:20OLE62.812ms,100OLE62.236ms (n10 each)
-100OLE nearby prose:28.7665ms (n10)

All samples are retained, including first/cold events; no outliers were excluded. These results do not demonstrate worsening latency with total formula count. They also do not explain every user document: fixture formulas/paragraph structure, Word layout state and actual physical UI differ. The out-of-process callback/COM/logging caveat remains. Paste-key and Delete-key hook installation fails in this harness, so these runs are not physical typing-hook or in-process add-in acceptance.

Before/after key values, registered DLL hashes and original Word process identities are exactly equal for both runs: installed20 11:15:10.1751996Z→11:15:30.4668773Z; installed100 11:17:11.5566337Z→11:18:06.4624084Z. Registry-suppression scopes explicitly remained skipped. All owned processes exited; original88900/93828/95356 keep the same starts/responding state. Read-only Windows module enumeration did not expose managed VisualTeX paths and therefore does not establish which DLL those existing user processes loaded.

No additional production fix was made: CP264's final Word DLL remainsF9B5B796F0B182A7E2B2346C87F9829DBF1F4915FDCB2CB1E204CD58119AAE60. New frozen-R3 test build cp265-prose-scaling reproduces that hash; only the test probe's optional100-OLE path and runner's copied-installed-DLL option were added. Held R4 hash remains unchanged. No install, registration, user-document edit, git commit/push or larger generic-count test followed.

Evidence under artifacts/word-prose-latency-20261004:
- installed-code/summary.json and seven per-type compiled-method comparisons
- installed20/: exact copied DLL hash, trace, saved20-object document, before/after state and passing completion
- installed100/: same evidence plus saved100-object document and0/20/100 trace
- installed100-summary.json: complete selected phase blocks, expected sample counts/unique completion IDs, arithmetic medians and exact state comparison
- Inspect-InstalledPassive.ps1 and the updated Run-ProseProbe.ps1

Next action is pending the user's affected copy/path and exact triggering interaction. Do not add more generic counts or infer a root cause that these measurements did not reproduce.

## Result

Ordinary collapsed Word selections do invoke VisualTeX's formula-related event checks. The audited current/frozen-R3 paths do **not** explicitly scan every formula or serialize the document merely because the caret is in ordinary body text. A confirmed unnecessary boundary path was narrowed: a caret immediately before a VisualTeX OLE previously activated/read its payload and requested layout-backed table membership, although caret normalization later rejected that leading boundary. At the trailing boundary, table-membership/descriptor reconstruction was also unused work.

Production change is confined to WordFormulaService.cs, NormalizeTypingCaretAfterInlineFormulaCore. It retains the same local ±2-character probe, exact VisualTeX ownership, inclusive touching-host ambiguity rejection, static reentrancy guard, and authoritative block/inline decision. It now proves the exact trailing boundary before reading metadata, keeps the post-read boundary check, and passes the retained exact local range directly to RestoreInlineCaret. Native OMML and third-party OLE remain outside this metadata path. Frozen/held WordFormulaHostResolver files are unchanged. This does not claim that all reported user-visible lag is fixed.

## Code evidence

- ThisAddIn.OnWindowSelectionChange: Ribbon owner discovery, numbered-host deletion/Enter handling, caret normalization, and selected OLE classification
- NormalizeTypingCaretAfterInlineFormula previously called WordFormulaHostResolver.ResolveLocal(VisualTeX). Both held R4 and frozen R3 have the same VisualTeX path: local shapes, inclusive boundary ownership, TryReadEmbeddedNativeOle, then IsWithinTable
- IsWithinTable uses Range.Information(wdWithInTable). The passive caller consumes only DisplayMode and Range, not this layout-backed descriptor field
- TryReadEmbeddedNativeOle reaches WordOleObjectAccessor.GetRunningObject, format.Object and possible DoVerb/345 ms of explicit retry sleeps. Retry/activation cost is a source-derived possibility, not an observed 345 ms delay in the real-Word samples
- GetFormulaJson only returns that OLE's metadata JSON. This path does not enumerate document Variables/CustomXML parts/bookmarks or export WordOpenXML
- RestoreInlineCaret still scans its owning paragraph for a nearby ordinary typing character. It is local paragraph work, potentially O(paragraph characters), and remains unchanged
- The direct numbered-OMML right-cell Enter repair has later document-bookmark/CustomXML lookup, but only after narrow structural gates. Ordinary body prose does not reach it
- VisualTeX copy/paste repair is armed by an actual pending paste; there was no unconditional VisualTeX document inventory scan on this passive path

## Verification and limits

Baseline Word DLL: FB031CF2A98E8C2266C90BF20730606985956220A7610A4CC88EC26763885EF2.
Final frozen-R3 Word DLL: F9B5B796F0B182A7E2B2346C87F9829DBF1F4915FDCB2CB1E204CD58119AAE60.
Final build: artifacts/omml-edit-audit-20261001/build-cp264-prose-final.
Command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File artifacts/omml-edit-audit-20261001/Build-Candidate.ps1 -Label cp264-prose-final.
Compiler: Visual Studio 18 BuildTools MSBuild, SDK 8.0.423, Release/x64/net472 with FreezeR3.targets.
Held R4 SHA256 remains 6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC.

Compiled managed interface-double tests: 9/9 passive checks pass. The new leading-caret case fails against the untouched baseline DLL with exactly one OLE activation attempt and one layout query; the final DLL reaches neither. Existing MathType, unknown/missing-registration observation boundaries and explicit-command content reachability pass. The ambiguity test also remains nonmutating. Old assertions were not relaxed.

Owned real Word final probe passed: 0 and20 real OLE fixtures,12 distant text selections,10 inline leading and10 trailing boundaries,10 nearby prose positions,5 block leading and5 block trailing boundaries, ten text insertions, unchanged20-object inventory and save. Every requested caret remains collapsed at the requested position; inline trailing carets retain zero baseline offset. This is COM selection/event testing, not physical mouse/keyboard or full Ribbon/editor lifecycle acceptance.

The add-in is manually hosted **out of process** in this harness. Sub-millisecond Selection.SetRange return times are not end-to-end UI latency; callback work occurs separately. The optional callback trace includes cross-process COM overhead and tracing I/O and must not be presented as installed in-process user latency.

Final event callback medians:
- Empty-document ordinary prose:35.893 ms (8 events)
-20-OLE distant ordinary prose:33.670 ms (12)
-20-OLE inline leading boundary:37.656 ms (10)
-20-OLE inline trailing boundary:52.360 ms (10)
-20-OLE nearby prose:32.608 ms (10)
- Block leading/trailing:36.711/42.744 ms (5 each)

Initial untouched-baseline trace shows leading-boundary49–55 ms and trailing-boundary62–73 ms, compared with final leading36–44 and trailing51–58 ms. These are bounded same-machine observations, not a formal statistical performance guarantee. No0→20 total-formula-count slowdown was reproduced. Ordinary prose still takes local checks; the remaining user symptom may involve installed version, layout/document structure or another component.

Existing Ribbon-double suite remains11/15 passing, with the same four unmodelled Range.Expand cases on baseline and candidate. Behavioral assertions for all15 pass; no expectations were changed to hide the coverage gap. Existing word-selection-change-idle fails before its interaction checks because its numbered fixture metadata is missing; it is not a passing numbering acceptance and its baseline was not rerun. No broad conversion/format/numbering suite or x86 runtime was run. Existing CP255/full-tree failure boundaries remain unchanged.

## Test harness registry incident and final isolation

The first baseline probe (~10:54:55–10:55:14 UTC), first candidate probe (~10:59:24–10:59:43), and attempted legacy idle regression (~11:03–11:04) reused the acceptance program's existing startup. That startup writes HKCU Software/Microsoft/Office/{Word,PowerPoint}/Addins/VisualTeX.{Word,PowerPoint}Vsto LoadBehavior=0, then restores the captured values on disposal. This write was outside the task's no-registry-change constraint. It must not be described as no registry writes.

No per-run pre-key snapshot was captured for those initial runs, so restoration of the original Word value cannot be independently proven from their artifacts. Read-only audit at11:06:10 found both keys present with LoadBehavior0, FriendlyName VisualTeX and the existing acceptance-suppression description (CP263 already documented PPT0). No separate restoration write was attempted.

The new prose-probe mode now explicitly skips both registry-suppression scopes. The final run11:07:49–11:08:09 captures read-only before/after keys, registered DLL hashes and Word process identities; all three comparisons are equal. All owned test processes exited through normal disposal. Original Word88900/93828/95356 keep their original start times and remain responsive. No user documents were opened or closed.

Registered product Word DLL is5702CC88C7CF5C795FE92D110AD336495CEE2E9F88FC77E5F207561CEC5D6575; registered PPT DLL isAFE19183CBCB0F3E9EE418F95C45E704078EF7AD801AE67892DC01B314192505. These files are unchanged by the final run and differ from source candidates. Actual loaded DLLs in the three existing user Word processes were not verified. No installation/deployment, git commit/push, reset/stash or source baseline replacement was performed.

## Files and evidence

Production:
- apps/windows/src-windows/VisualTeX.WordVsto/WordFormulaService.cs

Tests:
- New WordProseSelectionCostAcceptance.cs, plus new mode and registry-free branch in existing Program.cs (other prior PPT edits preserved)
- WordPassiveObservationChecks.cs adds leading-boundary/ambiguity checks without weakening existing checks
- artifacts/word-prose-latency-20261004/Run-ProseProbe.ps1 and Inspect-State.ps1

Evidence:
- artifacts/word-prose-latency-20261004/baseline and after-r2: original test evidence
- final: registry-free passing run, process seed/identity, saved document, detailed samples, callback trace, exact tested DLL and before/after state
- final-summary.json: phase timing summaries and equal state comparisons
- passive-baseline.json / passive-after.json: compiled before/after negative/positive proof
- ribbon-baseline.json / ribbon-after.json: unchanged suite gaps
- idle-regression/acceptance.log: retained fixture-precondition failure

Next step, if the user still observes lag: verify the loaded installed Word add-in version and profile the exact affected document in an authorized owned copy. Do not infer that current-source checks or this small test have diagnosed an unseen user's document, and do not expand into the frozen formula-conversion performance work.
