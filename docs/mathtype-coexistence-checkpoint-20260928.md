# VisualTeX / native MathType coexistence — continuation checkpoints

## CP13 — 2026-09-28 12:28 +08:00 — user accepts installed coexistence build; commit before apply-latency work

User confirms the new version does not crash and authorizes a local commit of the completed coexistence work, followed by root-cause investigation of the 1–2 second pause on VisualTeX MathType insert/edit Apply ONLY. Redraw/conversion are reported unaffected. Preserve the private service/account/process boundary and passive OLE observation policy. User explicitly permits closing unsaved documents; nevertheless use an owned test Word whenever practical.

Live inspection confirms installed Word DLL is now the packaged R3 `0FF0A17EEE048C58E21AC2B19826621B4C25453A5DBC4BF2E67752C8FC4BE59C` (not the earlier manually deployed hash). Current Word PID 41624 started 12:21:09; VisualTeX PID 89912 started 12:21:05. Native MathType 38260 / MathTypeLib 22688 retain their original 00:48:22 start identities. HEAD before commit is `7da7f6320026cf8d1418b9d932030510b2bd381f`.

Commit scope: accepted coexistence architecture, Word passive/Ribbon fixes, related packaging/source tests and checkpoint documentation. Exclude generated caches, crash dumps, installers, unrelated UI diagnostics and the explicitly held R4 OMML source/tests. No source optimization has been applied yet. Next compare explicit insert/edit native-render preparation against redraw/conversion, measure stage timings before changing code, then make one evidence-backed scheduling/render-preparation correction and rerun only targeted coexistence/apply checks through the installed Ribbon.

## CP12 — 2026-09-28 09:52 +08:00 — new R3 No-OCR NSIS built and actual payload verified; NOT installed

Named candidate now EXISTS, not just a planned filename:
`artifacts/mathtype-install-20260927/coexistence-20260928/release-inputs-r3/VisualTeX_1.2.8_x64_MathType-Coexistence-R3_NoOCR_20260928.exe`
Bytes: `148908491`. SHA256: `D9E93D1E91576857FEC4C16EAD5E8A29DA9EA3252C011071EBA2E4267B23880B`. Completed build at 09:48:16, rehashed at 09:51:59. No installation or installation smoke executed; open user Word documents remain protected.

Build discrepancy was identified rather than bypassed: standard Office build chose VS2022 MSBuild/C# 17.14, while the accepted installed R3 and standalone Word build used VS18 MSBuild 18.7. The frozen source WAS applied in both runs (compiler command/marker in transcript), so this was not evidence that R4 leaked into the package. Pinned the build-only driver PATH to the same exact VS18 MSBuild as the accepted R3 and reran Office onward. No comparison exclusion/old test expectation was relaxed. `packaged-word-build-identity-48180.json` proves the rebuilt x64 Word (`0FF0A17EEE048C58E21AC2B19826621B4C25453A5DBC4BF2E67752C8FC4BE59C`) matches installed R3 executable/resource/metadata bytes; only actual build/debug identity fields differ. The failed 17.14 build/comparison is preserved, not counted as a pass.

Release verification evidence under `.../release-inputs-r3/`:
- `package-build-state.json`: BuiltCandidateNotInstalled; Office, Desktop, Rust, Bundle, Verify complete. Frontend and embedded-asset checks passed, No-OCR configuration/size check passed.
- `x64-Passive-packaged-20260928.json`, `x64-Ribbon-packaged-20260928.json`: 7+15 checks passed on the actual x64 package build. Equivalent x86 reports pass 7+15 under 32-bit Windows PowerShell using explicit x86 dependencies. x86 DLL SHA256 `F56F7CB5F2B769964E6A42CD8733C0A8A89F45399C483B6BABB8A9F0CF6218BE`. These are compiled-code checks, NOT x86 Office UI acceptance. Test runners gained only an optional dependency-root parameter; assertion meanings unchanged.
- `embedded-payload-verification/report.json`: independently extracted this exact NSIS and decompiled/exported all THREE actual embedded MSIs without installing/executing payloads. Each Office architecture has 17 manifest-listed file payload hashes checked; all 83 native-render manifest payloads checked. No loose interactive MathType runtime in the outer NSIS; native engine resides inside the private-service MSI. Actual extracted x64 Word also passes installed-R3 equivalence.
- Embedded MSI hashes: x64 Office `0BF81A09AA1D929919BCEE22DF59AF876F6CB8965E15D86EA5FCE5E1B4E7C537`; x86 Office `7F507E2AABE06D537E0EC4AA3C51414276CDA6954E2AA041271A777D908BF1ED`; native renderer `E44F57DFD0A34D7180E95C2165B3BF7BB6CE05ED712963856616F4B6FA270E0F`.
- WiX decompiler emitted WIX1059 while interpreting permission-table relationships. Do NOT use its omission as the ACL authority. Independent raw queries of the embedded MSI (`embedded-payload-verification/service-raw-tables-verified.json`) confirm both LockObject values are NativeRenderData/Jobs, Table is CreateFolder, both referenced folder rows exist, exact protected service SID ACLs match, and service uses `NT SERVICE\VisualTeX.NativeRender`, automatic own-process/noninteractive service. The first raw-table helper accidentally output Execute()'s null return as an extra row; corrected that TEST helper with [void] and retained the failed evidence; no MSI change was made.

Controlled UI/content evidence is consolidated at `.../r3-controlled-mathtype-20260928-085914/controlled-sequence-summary.json`: all eleven specified step assertions pass, with native serialization differences and the exactly undone harness error explicitly retained. This is not eleven new tests in addition to CP9/CP10. At 09:51:59 the Application log has zero WINWORD/wwlib events 1000/1001 since the controlled test began at 08:59:14.

Final local state: installed DLL STILL original accepted R3 `700460A0AA50FD4D1879D990A5DD6EABF3C12378EB64BF01B0AEDA4442317A5D`; working-copy R4 source still `6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`, not reverted or shipped. HEAD unchanged `7da7f6320026cf8d1418b9d932030510b2bd381f`; no Git commit/push. Word 41416/mixed-r3.docx and protected 93104 preserved; owned test 55580/01-controlled.docx still open and saved. Native MathType 38260/MathTypeLib 22688 retain 00:48:22 starts; service Running. No user documents or unrelated processes closed for packaging.

Still incomplete: actual installation of this newly built NSIS followed by Ribbon UI acceptance; real x86 Office UI; both installation orders; never-MathType-installed/font-clean environment; original native installer disappearance. Older uncontrolled x=20 loss is NOT retroactively attributed to the user or to the harness mechanism found later. Do not label the candidate as full final acceptance or run installation while user Word documents remain open. The next gate is saving/closing user Office work with permission, then installing the named package and checking loaded hashes plus native/VisualTeX alternating operations. No unrelated OMML work is needed for that gate.

## CP11 — 2026-09-28 09:37 +08:00 — MathType-only source inputs frozen; candidate package building

To exclude the unrelated held R4 without reverting anyone's working copy, wrote an immutable build input under `artifacts/mathtype-install-20260927/coexistence-20260928/release-inputs-r3/WordFormulaHostResolver.cs`. Its Git blob exactly equals HEAD/R3 (`d082edfdf56412d53aa125248f72fc58919df560`), SHA256 `FCE1BB7534EC0C9BF518C06818FA7EDA6816EBA5A0266D203CD84B8A84EBD045`. Working-copy R4 source remains `6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`; original R4 DLL/PDB preserved under `release-inputs-r3/held-r4-build/`.

A process-scoped standard MSBuild import (`FreezeR3.targets`) replaces only that Compile input, preserving source order. It is build isolation, not a runtime compatibility branch and is not installed. Fresh x64 Word build SHA256 `0FF0A17EEE048C58E21AC2B19826621B4C25453A5DBC4BF2E67752C8FC4BE59C`. `word-build-identity.json` proves EXACT equality to installed/accepted R3 of executable bytes, resources and metadata after excluding only actual MVID/timestamp/checksum and CodeView/PDB debug records. Shared normalized hash `BA32910592B66F1A406AEF73D0000551B6F996D3F4AA4E5AECAEAE8B0E4F4969`. Installed Word was NOT replaced.

`Build-CoexistenceCandidate.ps1` started at 09:37:03, build driver PID 80384. It prepares the existing verified NSIS template, builds both Office architectures and native renderer MSI using frozen R3, builds/verifies frontend and Rust, bundles No-OCR NSIS and verifies it. No git operation, release cleanup script, installation smoke, user-process termination, or installation is requested. The ordinary Office build's build-tree OLE-server stop path is preflighted empty; fail instead of stopping an active build-tree server. Old generated no-OCR installer is backed up before reuse of its output path.

Resume/read actual `release-inputs-r3/package-build-state.json` and transcript before any retry; the script has explicit ResumeAt stages to avoid rerunning finished work. At this checkpoint Status=Running, Stage=Office. A planned delivery filename is NOT evidence the package already exists. Final installed-package UI acceptance and clean-font/install-order/original-native-installer issues remain pending.

## CP10 — 2026-09-28 09:28 +08:00 — controlled R3 alternating native/VisualTeX operations and reopen

All evidence below is in `artifacts/mathtype-install-20260927/coexistence-20260928/r3-controlled-mathtype-20260928-085914/`, actual installed R3 Word PID 55580. No product source/build/deployment change this turn. Unrelated R4 remains held, not installed. Original twenty-equation baseline and the user's earlier mixed fixture were not overwritten.

After CP9:
- Actual VisualTeX edit x=20 -> x=220, with right numbering and 12 pt retained: `04-after-vt-apply.json` has twenty OLEs; only index20 MTEF changed, every native field instruction and other nineteen MTEFs remained exact.
- Starting an empty create session on the selected equation then using explicit Cancel (`05-after-blank-create-cancel.json`) and title-bar X (`06-after-blank-create-close.json`) left all twenty MTEFs/field instructions unchanged. No missing object reproduced.
- Actual VisualTeX Ribbon insertion at an ordinary document-end caret created z=3, 10.5 pt, Equation.DSMT4. `07-after-vt-new.json` has twenty-one OLEs; original twenty content/number fields unchanged. One VTMATH bookmark belongs to this self-created object; no zero-bookmark claim is made for the mixed document. Earlier fraction text-entry attempts in the source UI were not committed and are NOT fraction/layout acceptance.
- VT-only interval `vt-window-interval.json`, 09:08:34–09:13:49, recorded no native top-level SHOW/HIDE events. Observer used only OS WinEvent/EnumWindows, not Office COM. `vt-public-state-comparison.json` confirms all 1,382 public settings, 506 installed files, 24 font registrations and five recorded process identities unchanged across the VT interval, BEFORE resuming native actions.
- Actual native Display/Right/Left x=22..30, nine Ctrl+F4 insert-return cycles after the private renderer was used: `08-after-native-resume.json` has thirty OLEs, all prior twenty-one content/field instructions unchanged. Closed only this saved test document and reopened it through Word UI. `09-after-reopen.json` has all thirty exact MTEFs/fields unchanged.
- Reopened VisualTeX-created z=3 is natively editable through Equation > Edit. Changed it to z=33 and saved: `10-native-edit-vt-created.json` has thirty OLEs, only index21 MTEF changed; 10.5 pt retained. Native editing changed how the native MTPlaceRef/SEQ instruction strings are materialized (some nested SEQ vs cached label forms); this is reported, not treated as byte-identical numbering. Screenshots still show the original native labels.
- Correctly clicked native Equation > Open on the VisualTeX-rewritten x=220: actual external native editor screenshot `screenshot-0-092557.png` shows x=220. Changed to x=320, Ctrl+F4, saved. `12-after-native-open-return.json` still contains thirty OLEs, target20 is x=320/12pt; previously natively edited z=33 retains its meaning/10.5pt but its serialized Equation Native hash changed again. Do NOT claim all non-target MTEF bytes unchanged during this last native serialization step. Other twenty-eight MTEF hashes unchanged.

Test-method defect isolated and corrected, not a product repair: while initially inspecting the native Open submenu, the helper activated the Word document before sending Enter. That dismissed the nonactivating popup and inserted a paragraph over the selected x=220. It reproduced the appearance "formula missing, number remains" in this NEW test copy. A single Undo restored all thirty exact MTEF/field instructions (`11-after-harness-input-undo.json`, compared with step10). This does not by itself prove the cause of the older uncontrolled x=20 loss. Helper physical popup clicks now retain the existing owner-foreground state and hit-test the exact popup; the successful native Open was a physical click, not the failed key sequence. No user document was altered.

At the 09:20 event check, zero WINWORD/wwlib events 1000/1001 since this controlled Word started. Native MathType 38260 / MathTypeLib 22688 retain original 00:48:22 starts. Word 41416 and protected 93104 untouched. This is targeted installed-R3 coexistence evidence, NOT final NSIS, every formula/layout, both installation orders or a font-clean machine acceptance. Next freeze MathType-only package inputs without incorporating held R4, verify produced package contents/hashes and keep uncompleted environment/installer cases explicit.

## CP9 — 2026-09-28 09:08 +08:00 — controlled native edit and VisualTeX cancel preserve all objects

New actual installed R3 Word PID 55580, start `2026-09-28T00:59:14.9297995Z`, trace/evidence directory `artifacts/mathtype-install-20260927/coexistence-20260928/r3-controlled-mathtype-20260928-085914/`. No debugger, no external Word document polling. Prior mixed fixture in Word 41416 and protected Word 93104 were not operated or closed. R4 remains uninstalled.

`00-baseline.docx` is an immutable Save As copy of the original native twenty-equation document. `01-controlled.docx` was saved separately and verified before editing: `01-before-edit.json` proves twenty OLEs, no OMML/VT bookmarks, and exact MTEF/field-instruction equality with the new baseline.

1. Native Equation context-menu Edit of x=15: native in-place UI screenshot confirms only x=15 selected, then x=115 typed, clicked outside and saved. `02-after-native-edit.json` has twenty objects; only MTEF index 15 changed to x=115, all native field instructions unchanged, other nineteen MTEFs unchanged including x=20. This controlled native-edit step passed, not a claim about every edit path.
2. Clicked the visible x=20 and actual VisualTeX Ribbon Edit Selected Formula. The real installed editor screenshot `screenshot-0-090651.png` shows x=20, MathType OLE, 12 pt, right numbering. Clicked its explicit Cancel button, returned to Word and saved. `03-after-vt-edit-cancel.json` has all twenty MTEFs and field instructions exactly unchanged from step 1; no VT bookmarks.

Thus the earlier missing x=20 has not reproduced in either isolated step. Do not blame the user's manual clicks or declare the earlier sequence explained. Next test actual VisualTeX changed-content Apply, blank create+cancel/close on an already-selected native equation, native edits after these actions, and save/reopen with stepwise saved-MTEF checks. Native-window/settings isolation gets its own bounded observation interval.

Test-only helper changes: resize only the owned new test window for readable screenshots; explicitly bind the actual installed VisualTeX editor HWND/PID/start identity for UI testing. A PowerShell 5 UTF-8-without-BOM title literal caused a parser failure before any input; replaced that literal with exact Unicode codepoints in the test helper. No product change/build/deployment.

## CP8 — 2026-09-28 08:59 +08:00 — user clarified manual input; controlled MathType-only continuation

The user explicitly confirms the extra VisualTeX clicks in CP7 were their manual input and authorizes continuing. They are not evidence of another automated UI writer. This resolves attribution of the extra clicks, NOT the cause or exact time of x=20 disappearing. The mixed-r3 fixture remains an uncontrolled sequence and must not be labelled a native-edit failure in isolation.

Rechecked at 08:58:35: installed DLL still R3 `700460A0AA50FD4D1879D990A5DD6EABF3C12378EB64BF01B0AEDA4442317A5D`; workspace DLL still held R4 `5DAE6D55D0BCE336715A677BE465D4B72ECD003757C2883365B1C8FC47D9D35A`. HEAD unchanged. Word 41416 and protected 93104 remain; native MathType 38260/MathTypeLib 22688 retain their 00:48:22 starts; NativeRender Running. No production/deployment change.

Next single-variable sequence: leave the saved missing-object fixture untouched; launch a separately identified actual installed R3 test Word, open the unchanged twenty-object baseline and Save As a new test copy, verify twenty objects before any edit. Test native edit and VisualTeX explicit edit open/cancel/apply separately, saving and decoding MTEF between each step. Never infer that an editor opened correctly from exit code alone. Use OS/UI inspection only during live insert/edit; decode the saved ZIP after operations, with the explicitly recorded installed R3 parser. Continue to preserve existing native process/settings/numbering boundaries. Do not deploy or develop the unrelated OMML R4.

## CP7 — 2026-09-28 08:51 +08:00 — R3 reopen/native-edit observation; mixed sequence NOT passed

Scope remains MathType only. No production source/build/deployment change was made this turn; the unrelated R4 OMML change is still held, not installed.

Actual test: `coexistence-20260928/r3-mathtype-only-reopen-20260928-083954/`, Word PID 41416/start `2026-09-28T00:39:54.7291533Z`, real installed R3, no debugger or Word document-model polling. The prior owned Word 528 was allowed to close only after visually selecting Yes/review recovered files later; recovery files were not deleted. Protected Word 93104 was untouched.

Opened `candidate-r3-native-20.docx` via Word Open UI. Screenshot `screenshot-0-084142.png` shows all twenty native equations with left/right numbers. Saved a separate `mixed-r3.docx` through Save As before editing. Original baseline SHA256 was rechecked at 08:51 and is still `C2E4C439A80A60183B25B77755E9D03F4DC3D1BE4940C1A7FFD2412AC9C0099F`.

Native edit actually exercised: the native Equation context-menu **Edit** verb opens MathType IN-PLACE, not an external EQNWINCLASS window. Screen `screenshot-0-084813.png` identifies the actual selected object as x=15 (do not claim x=1 from an older viewport). `screenshot-0-084845.png` confirms Ctrl+A selected only its native editor content. Typed x=115, clicked outside to commit/deactivate, then saved through Ctrl+S. No VisualTeX command was intentionally issued by this test script during that edit.

Important failure/attribution boundary:
- `native-edit-saved.json`, produced at 08:49:43 with the installed R3 parser DLL, contains **19**, not 20, Equation.DSMT4 objects. x=15 is correctly x=115 / 12 pt, but x=20 is missing; screen `screenshot-0-084943.png` retains its right number (1.13). Saved SHA256 at inspection: `23223C8E9CB3356C6B44999052CE4F879302E8EE94362E89B824DE67C1F393DF`. This run is NOT a passing coexistence or content-preservation test.
- x=20 is not visible in the screenshot at entry into native in-place editing, before the deliberate x=115 input. That visual state alone does not establish when the stored OLE was removed; native in-place rendering can differ from the final saved package. Neither the earlier extra actions nor the later edit/commit has yet been isolated as the cause.
- The hook log has additional physical clicks and three VisualTeX OpenSession sequences at 08:42:59, 08:43:29 and 08:46:06 (session IDs 5cdfc104-b7c8-4bff-8633-a4dcedffadaf, f3a0d837-9897-4aac-b037-be01696efa92, 357bdeaf-7dae-49e2-a16d-405863bc7498). These are NOT in this run's scripted `ui-actions.ndjson`. The first is preceded by a click on x=20, then the VisualTeX tab and button. Actor/cause is not known: possible concurrent manual/test input must be clarified, not silently assumed or blamed. Therefore this interval is not a single-variable native-edit control.
- Test-window focus guards repeatedly stopped input while Chrome PID 21116 was foreground. No command was sent on those failures. The diagnostic helper now logs the actual foreground identity and temporarily exposes only the exact owned HWND's caption, physically clicks Windows-confirmed HTCAPTION space, and restores its original topmost state in finally. It does not minimize/click Chrome, alter global foreground policy or change product hooks. The invalid SendKeys `{SPACE}` attempt was a harness error, not a product test; subsequent system-menu restoration used the literal space.
- Offline inspector now accepts an explicit parser assembly and records its path/hash; this run used installed R3 `700460A0AA50FD4D1879D990A5DD6EABF3C12378EB64BF01B0AEDA4442317A5D`, NOT the workspace R4 build. Ordinary Get-FileHash on the live open mixed file was sharing-denied; the inspector's FileShare.ReadWrite hash above is the actual saved-file evidence.

At 08:51 Word 41416 and protected Word 93104 are alive; native MathType 38260 / MathTypeLib 22688 retain original start identities. Application events 1000/1001 since this test's 08:39:54 start contain zero matching WINWORD/wwlib crash records. This is not proof against all native exceptions, and does not turn the missing-object result into a pass.

Next: clarify extra live UI input before sending more actions to the shared desktop. Preserve current saved fixture and logs. On a fresh copy of the unchanged baseline, verify all twenty objects immediately after Save As and before editing, then isolate native Edit/Open and explicit VisualTeX edit/cancel/apply as separate recorded steps. Investigate a newly reproduced loss in MathType command/lifecycle handling, not in OMML. No final NSIS, install-order or font-clean acceptance claim.

## CP6 — 2026-09-28 08:38 +08:00 — scope corrected to MathType coexistence; R4 NOT deployed

The user explicitly challenged the diversion into OMML. The canonical-OMML-boundary issue found during mixed conversion is a separate defect, not a prerequisite for proving/fixing the native MathType conflict. Stop that branch of work. Its existing source/test changes are preserved for separate review; do not silently delete them, deploy R4, or include them in a MathType-only acceptance claim.

Actual state verified this turn, not inferred from the duplicate/out-of-order CP4 entries below:
- Installed Word DLL remains R3 SHA256 `700460A0AA50FD4D1879D990A5DD6EABF3C12378EB64BF01B0AEDA4442317A5D`; current build output is R4 `5DAE6D55D0BCE336715A677BE465D4B72ECD003757C2883365B1C8FC47D9D35A`. They are NOT interchangeable. No deployment/build/product edit this turn.
- HEAD remains `7da7f6320026cf8d1418b9d932030510b2bd381f`. R4's `WordFormulaHostResolver.cs` diff is the unrelated local OMML change. It stays on hold.
- Owned test Word PID 528/start `2026-09-28T00:15:21.6274298Z` and protected Word 93104 remain. MathType 38260 / MathTypeLib 22688 have the same 00:48:22 local starts; NativeRender is Running.
- Current owned Word is at a modal Microsoft Word dialog (HWND 16058052, owner 12200866, PID 528). Inspect it before input; never manipulate protected Word 93104.
- Read persisted native-20 inspection: twenty Equation.DSMT4 OLEs, thirteen native numbered equations, no VT bookmarks. This is previous R3 UI/save evidence, not twenty new cycles this turn.
- Read `vt-native-window-observer-1.json`: 08:22:43–08:31:09 out-of-context WinEvent observation, no recorded native SHOW/HIDE events, native process identities unchanged. It covers that interval only.
- Ran saved-snapshot comparison `r3-public-state-scope6-audit.json`: all 1,382 public-setting hashes, 506 installed-file records, 24 font registrations and four recorded process identities are unchanged across the prior VisualTeX interval. This is an audit of saved snapshots, not a new insertion test.

Next strictly MathType-only gate: use the installed R3 and a verified owned test Word, reopen the saved native document, interleave native Display/Left/Right and native equation edits with VisualTeX LaTeX-to-MathType and explicit MathType editing/font operations; save/reopen and decode actual MTEF/native numbering. Observe native windows only during VisualTeX actions and compare native public state separately from legitimate native edits. No OMML-to-VisualTeX work, global Office polling, extra delay/catch patches, or whole-suite reruns. Final NSIS/install-order/font-clean acceptance remains uncompleted.

## CP0 — 2026-09-28 07:32 +08:00 — handover verified, before any product change

Request: continue the crash/root-cause work, report progress during execution, and keep explicit checkpoints. Do not label an old installer or a partial test as the final fix.

Workspace: `C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7`; reuse `ws_6e742f6c2d`.
Git HEAD (read-only query): `7da7f6320026cf8d1418b9d932030510b2bd381f`. Existing uncommitted changes belong to previous work and must remain intact. This checkpoint is documentation, not a Git commit or a rollback of the worktree.

### Directly rechecked this turn

- Read `docs/mathtype-word-event-reentrancy-20260928.md` in full. Its last completed stage remains R2 installation plus seven compiled-DLL boundary checks; final R2 physical UI acceptance is not recorded as passed.
- Actual installed Word DLL and workspace `bin/x64/Release/net472` DLL both SHA256 `8E63D124CBA192D0055C692DCDC60D91794DCEDA29741B8F7B06CE224FE48FB5` (R2).
- Initial source SHA256: `ThisAddIn.cs` = `0007AF4E8311E8B7D3B9596783271D69A316E41201C865164DC5485256E8960C`; `WordFormulaService.cs` = `F06402C98D5EB2FB0761E081454E9E9DD0ED1E1F2CDB47E83E806D33DD3CA598`.
- OS process inspection only, no Word COM access: Word PID 68756 started 2026-09-28 01:10:21 +08:00 (prior R2 test); Word PID 93104 started 2026-09-27 23:57:22 +08:00 (unrelated/protected).
- Native MathType PID 38260 and MathTypeLib PID 22688 both started 2026-09-28 00:48:22 +08:00 and remain in Session 1.
- `VisualTeX.NativeRender` reports Running. No service changes made.
- No product source, installed DLL, registry, processes or documents changed during CP0. Only this new checkpoint file was written.

### Established previous evidence (not new tests this turn)

Two local full Word dumps have recorded the same initiating call route: `OnWindowSelectionChange -> NormalizeTypingCaretAfterInlineFormula -> WordMathTypeHostAdapter.ResolveLocal -> ReadMetadata -> MathTypeWordOpenXml.Read -> Range.WordOpenXML -> wwlib+0x2ee4e8`. Second baseline uses normal Ctrl+F4 equation close and has no external document COM polling. Exact private Word internal lifetime mechanism remains unproven.

R1 has nine successful actual native Ribbon insertion cycles; saved content contains eleven decodable MathType OLE objects. R2 removes additional passive identity-fallback export. Six passive + one explicit-command compiled-DLL tests passed. Those are not the same as final R2 UI acceptance.

### Next gated steps

1. Read the actual stack evidence and inspect all relevant passive selection/Ribbon/hook call routes, not just methods covered by the seven tests. Distinguish confirmed unsafe content access from unproven native crash hypotheses.
2. Determine whether R2 already closes the proven path or a narrow additional ownership-boundary correction is necessary. No timer/sleep, catch-based masking, process-name ownership guesses, or wholesale rewrite solely to satisfy old tests.
3. Add/run focused checks only for concrete gaps. Preserve explicit-command access to third-party MTEF and standalone MathType OLE capability.
4. Check the prior test window without modifying any unrelated document/dialog. Actual UI actions must be scoped to verified test PID/start identity; no external COM document polling. Do not bypass another document's Save dialog.
5. Record build/test/deployment evidence separately. Final NSIS, native/VisualTeX interleaving, save/reopen/edit, font-clean environment and both installation orders remain pending until actually demonstrated.

Constraints: no reset/clean/stash/branch changes/pull/rebase/merge; no unrequested commit/push; no deleting unknown files; do not close or manipulate the user's other Word documents; keep crash dumps local.

## CP1 — recorded 2026-09-28 08:00 +08:00 — connection recovered; R2 unchanged

This records the interrupted 07:32–07:40 review, not a new completed fix. Read/write and readback now work. At 08:00 the installed DLL is still R2 (`8E63D124CBA192D0055C692DCDC60D91794DCEDA29741B8F7B06CE224FE48FB5`), both product source hashes are unchanged from CP0, and HEAD is unchanged. Word 68756 / protected Word 93104, MathType 38260 / MathTypeLib 22688 retain their recorded start identities; NativeRender is Running.

Re-read the actual `baseline-ctrlf4-20260928-010917/managed-stack.txt`: it confirms the passive caret-to-MathType-WordOpenXML initiating chain. No claim is made about private Word internal lifetime details.

The persisted `candidate-r2-no-debugger-20260928-011021/native-cycles-0-073521.json` records ONE completed actual Ribbon Display insertion, x=1, Ctrl+F4 return at 07:35:23. The second Right-numbered click did not yield a unique editor before the timeout. Prior live inspection identified a native Insert Equation Number dialog in the owned test Word; its exact present state must be inspected again before input. This is not three passed cycles and not evidence of a new crash. The screen-crop image from 07:36 contains an overlapping ChatGPT window and is not valid visual evidence of the target dialog. Do not infer target content from that crop.

Two code-review gaps need focused tests before changing production:
1. `GetFormulaFontSizeEnabled` uses known semantic size OR positively recognized MathType ProgID. With missing/unrecognized registration, R2 correctly avoids content reads but may prevent the user from reaching an otherwise valid explicit command. The existing explicit-command test invokes the internal method directly and does not prove Ribbon reachability.
2. `ScheduleFormulaFontControlsInvalidation` skips invalidation on prose when the font cache is NaN. R2 now permits an enabled target with unknown semantic size, so that implication is no longer valid. Need a target-state transition test, not restoration of passive payload reads.

Only diagnostic `Scoped-Ui.ps1` gained a foreground-only Focus action in the interrupted turn. No production source, installed binary, native settings/service or unrelated Word documents were changed. No new build/test/NSIS/commit was completed. Next: test actual compiled Ribbon callbacks and transitions, make only evidence-backed state-boundary changes, then finish real installed UI validation with exact PID/start ownership and no document COM polling.

## CP2 — 2026-09-28 08:12 +08:00 — R3 compiled and focused checks passed; before deployment

Evidence-backed production changes are confined to the two existing Word files:
- `WordFormulaService.cs`: distinguish a confirmed OLE with unresolved ProgID from no OLE. A font command may inspect such a candidate; positive MathType/VisualTeX recognition is still required for passive native gesture routing. Explicit semantic validation remains mandatory before mutation. No passive content/activation fallback was restored.
- `ThisAddIn.cs`: command enablement uses that candidate distinction; Ribbon invalidation uses loss of target/window context as well as known font size. Unknown size no longer falsely implies disabled controls. Prose-to-prose movement still avoids redundant invalidation.

New `WordRibbonObservationChecks.cs` / `Test-WordRibbonObservation.ps1` load the actual compiled add-in with managed Word-interface doubles. The unchanged R2 produced six failing cases covering inaccessible unknown/missing-registration commands and stale selection/window font state (`ribbon-boundary-r2-before-20260928.json`). This is a failed baseline, not a passing product test.

R3 build succeeded (`docs/remediation-3d207d7/evidence/mathtype-coexistence-r3-ribbon-state-20260928-build.log`), retaining existing compiler warnings. Built DLL SHA256: `700460A0AA50FD4D1879D990A5DD6EABF3C12378EB64BF01B0AEDA4442317A5D`.
- Existing seven passive/explicit-command checks pass (`passive-boundary-r3-20260928.json`).
- Fifteen Ribbon/full-event/negative-control checks pass (`ribbon-boundary-r3-outargs-20260928.json`). All passive cases have zero content exports/activations/clipboard capture/selection writes and no unmodelled calls. Known MathType still caches geometry; unknown OLEs do not acquire gesture ownership. An unrelated non-MathType payload is rejected by explicit semantic inspection, and an ordinary picture remains disabled.
- The managed proxy initially did not preserve GetPoint out-arguments; the test proxy was corrected to return the same argument array and a positive geometry assertion was added. The final outargs report is the authoritative expanded-check result. No production code was changed to accommodate this test-model issue.

Actual R2 UI continuation (no debugger, installed DLL, no document COM polling): inspected the owned chapter/section initialization dialog using PrintWindow, confirmed values 1/1 and the global-default checkbox unchecked, accepted only for this test document. Completed the pending x=2 Right insertion with Ctrl+F4, then nine physical Left/Display/Right cycles x=3..11 (`native-cycles-2-080734.json`). Saved through Word Save As UI to `candidate-r2-no-debugger-20260928-011021/candidate-r2-native-20260928.docx`. Offline inspection confirms 11 decodable Equation.DSMT4 OLEs, all 12 pt, seven native numbers, and no VT bookmarks. A foreground screenshot at 100% (`screenshot-0-080931.png`) shows x=1..11 and alternating native left/right numbering; the earlier 230% crop was not full-layout evidence.

The owned test Word was closed normally after saving. At its recovery prompt, explicitly selected Yes/review later; no recovery files were deleted. Protected Word 93104 and native MathType processes were not closed. The diagnostic focus helper now clicks only a Windows-hit-tested caption of the exact owned window when background foreground activation is denied; no arbitrary input, injected Office code or global focus setting changes.

R3 is not yet installed at this checkpoint. Next: verified Word-only diagnostic deployment, fresh actual installed /x /w Word without debugger, native repetition and VisualTeX interleaving, save/reopen/edit, then full package build/install gates. No final NSIS or Git commit exists for R3.

## CP3 — 2026-09-28 08:20 +08:00 — R3 installed; twenty actual native insertions, interleaving next

First deployment was correctly stopped before changes because rebuilt Contracts SHA256 differed from the installed file (`candidate-r3-deployment/failed.txt`). A conservative PE/CLI comparison then proved all executable, resource and metadata bytes identical after excluding only the actual Module.Mvid, COFF/checksum and CodeView/PDB debug records (`contracts-build-identity-20260928.json`; normalized hash `C8BBD98AFD4BC3E01F33E611DDF76A941308B02E1F7B2808687DDEE6993851E4`). The installed Contracts DLL was NOT replaced, and remains `74EFF692736E10BCD7B858EF804AED858EF663C4CBCF88365385E23F4DA3E9AA`. This is a byte-level build-identity comparison, not a version-string/API-compatibility assumption.

Verified deployment at 08:15:00 replaced only the known Word DLL/PDB with R3 `700460A0AA50FD4D1879D990A5DD6EABF3C12378EB64BF01B0AEDA4442317A5D`, preserving native processes/settings/service and protected Word 93104. Backups/proof/report: `candidate-r3-verified-deployment/`.

Actual installed R3 Word PID 528, started `2026-09-28T00:15:21.6274298Z`, /x /w, no debugger and no external Word document polling. Evidence root `candidate-r3-no-debugger-20260928-081521/`. Its `word-hook.log` confirms installed path, MVID `9d2ffe9e-f9cd-4084-b073-b8235da9dfb7`, and normal hooks. Native Display x=1, Right x=2 after its own 1/1 chapter/section initialization (global-default checkbox left unchecked), then 18 physical Left/Display/Right insert-return cycles x=3..20 all returned normally using Ctrl+F4. Reports `native-cycles-0-081545.json`, `native-cycles-2-081624.json`, plus scoped UI actions. Native MathType PID 38260 and MathTypeLib 22688 remain unchanged. This is not yet final mixed-product acceptance; the 20-equation test document has not yet been saved/reopened/edited.

Before the first VisualTeX operation, captured `artifacts/mathtype-install-20260927/local-machine-r3-pre-visualtex-20260928/`: 1,382 public setting hashes, 506 installed file hashes, 24 font registrations, process identities. Compare immediately after VisualTeX operations, before another native operation can legitimately update native settings. Current own Word is maximized, VisualTeX tab selected, caret in prose after x=20. Next use actual Ribbon actions for interleaving and verify no native editor popup.

## CP4 — resumed after response failure, 2026-09-28 08:30+08:00 — actual R3 evidence rechecked

The failed chat response is not an execution checkpoint. Re-read this file and inspected the current installed/workspace DLLs: both remain R3 SHA256 `700460A0AA50FD4D1879D990A5DD6EABF3C12378EB64BF01B0AEDA4442317A5D`. HEAD remains `7da7f6320026cf8d1418b9d932030510b2bd381f`. Owned Word 528, protected Word 93104, MathType 38260 and MathTypeLib 22688 retain the recorded identities; renderer is Running. Do not restart/reproduce on unrelated Word.

Persisted evidence after CP3 was ahead of the narrative:
- `candidate-r3-native-20.docx` and `native-20-inspection.json` confirm 20 decodable native OLEs / 13 native numbered equations / no VisualTeX bookmarks.
- `word-hook.log` records actual VisualTeX LaTeX-to-MathType conversion for `y=\\frac{1}{2}+\\alpha`, successful explicit font-size change to 11, and MathType-to-OMML conversion at 08:26:07. These logged operations are not yet complete interleaving/reopen acceptance.
- At 08:26:17 OMML-to-VisualTeX failed with `The exact OMML source host survived text clearing.` The pre-mutation render request already omitted the initial y. Word did not crash. Saved failure file `candidate-r3-omml-failure.docx` retains the complete OMML and all 20 original Equation Native payloads/native fields unchanged, verified by `omml-failure-inspection.json`.
- The single explicit idle diagnostic `omml-boundary-idle-snapshot.json` proves `selectedRange.OMaths[1].Range` is CLIPPED to its source Range: selection 2840:2851 returns math 2840:2850 and XML beginning `=`, whereas owning paragraph 2839:2851 returns math 2839:2850 and complete XML beginning `y=`. A second clipped probe also truncates the final alpha. Thus Count==1 plus agreement with another equally clipped OMath is not proof of the complete host boundary. This is independent of the passive third-party OLE export crash.

Next focused work: fix canonical local OMML host discovery rather than deleting the safety assertion or broadening deletion blindly. Preserve paragraph/table/text and all third-party OLEs; do not use document-wide scans or export third-party OLE in passive paths. Add partial-selection/caret/multiple-formula negative tests against the compiled resolver, then real installed Ribbon conversion checks. Current test document is the saved failure fixture in owned PID 528. No new product change or build has yet been made in this resumed turn. Final NSIS, both install orders and font-clean validation remain incomplete.

## CP5 — R4 local OMML boundary candidate built, before deployment

New focused tests `WordOmmlLocalBoundaryChecks.cs` / `Test-WordOmmlLocalBoundary.ps1` reproduce Word's observed clipped OMath views with managed interface doubles. Against the unchanged installed R3 they expose ten actual resolver failures (prefix/suffix fragments, interior/boundary carets, second formula, display, table/header and distant document positions). Negative controls for prose, ambiguous multi-formula selection and REF-result exclusion pass. Evidence: `omml-local-boundary-r3-before.json`; the failing baseline is not labelled a product pass.

The sole new product edit in this resumed turn is `WordFormulaHostResolver.cs`: local selection/probe OMaths are presence evidence only; acquire complete candidates from a duplicate expanded to the owning paragraph(s), filter against the ORIGINAL selection, and preserve unique/exact host validation. No Selection mutation, whole-document OMath enumeration, paragraph XML export, blind character offset, metadata ownership, or deletion-assertion removal was added. Existing OMML numbering/reference logic remains in place.

R4 build succeeded (`mathtype-coexistence-r4-omml-boundary-20260928-build.log`). DLL SHA256 `5DAE6D55D0BCE336715A677BE465D4B72ECD003757C2883365B1C8FC47D9D35A`.
All 37 targeted checks pass on that DLL: 15 local OMML boundary cases (`omml-local-boundary-r4.json`), 7 third-party passive/explicit-command cases (`passive-boundary-r4.json`), 15 Ribbon/full-selection-event cases (`ribbon-boundary-r4.json`). No global reads or exports outside an exact OMML host occur in the new cases. Ordinary prose avoids paragraph expansion. These are compiled-code model tests, not final live Office acceptance.

R3 is still installed; R4 is only built at CP5. Next preserve/close only the saved own R3 fixture, deploy via verified dependency/hash gate, launch a fresh actual installed Word, reproduce the same clipped selection through actual Ribbon, and continue native/VisualTeX alternating edit/save/reopen. Protected Word 93104, native MathType processes and renderer must remain untouched.

## CP4 — 2026-09-28 08:33 +08:00 — native cycles saved; mixed UI exposed clipped OMML host, not a crash

Saved actual R3 native-only document through Word UI as `candidate-r3-no-debugger-20260928-081521/candidate-r3-native-20.docx`. `native-20-inspection.json` confirms 20 valid Equation.DSMT4 OLEs, 13 native numbered paragraphs, no VT bookmarks; saved SHA256 `C2E4C439A80A60183B25B77755E9D03F4DC3D1BE4940C1A7FFD2412AC9C0099F`.

Mixed actual Ribbon actions in Word 528:
- Literal `\\[y=\\frac{1}{2}+\\alpha\\]` entered by Unicode input, selected locally, redrawn as MathType through the actual Ribbon. Completion status reports one MathType. The first keyboard menu navigation selected VisualTeX OLE instead; its confirmation was inspected and cancelled BEFORE mutation, then the visually verified MathType menu item was clicked. Do not count the cancelled dialog as a conversion.
- Actual Increase font button completed; semantic size 11 pt. Trace records native preview, CFB rewrite and successful completion, not just a direct helper call.
- Actual MathType -> OMML completed, one formula.
- Actual OMML -> VisualTeX failed with `The exact OMML source host survived text clearing.` Word remained alive. This is a real incomplete acceptance case, not masked as a pass.

Saved failure after rollback separately as `candidate-r3-omml-failure.docx`; `omml-failure-inspection.json` reads only the saved ZIP. It shows the complete y=1/2+alpha OMath and the original 20 native MathType Equation Native hashes / native numbering instructions unchanged. No VT bookmarks. Original native-only baseline file was not overwritten.

A single explicitly scoped idle-Word diagnostic (NOT UI acceptance, NOT polling, no concurrent native insert/edit) compared the current selection's OMath with the containing paragraph's OMath. `omml-boundary-idle-snapshot.json` proves:
- Owning paragraph OMath full range 2839:2850, starts with mathematical italic y (a UTF-16 surrogate pair).
- Current selection 2840:2851 yields `.OMaths[1].Range` 2840:2850, beginning inside that surrogate pair.
- Independently constructed 2841:2850 yields an OMath range starting at '='; ending at 2849 similarly cuts the final mathematical alpha surrogate.
Thus Word's range-scoped OMath collection returns clipped equation views; `OMath.Range` equality against the same clipped scope does NOT prove a complete physical host. Both `ResolveLocalOmml` and `CaptureOmmlIndex` currently trust such a view. This explains partial semantic capture and the deletion guard correctly detecting a surviving OMath. The fix must resolve a full canonical local host independently of the selection, not remove the guard or widen by a guessed number of characters.

Diagnostic helper corrections were limited to test infrastructure (visible owned OpusApp selection, PowerShell array-expression parentheses, FileShare.ReadWrite hashing); no product change for this new defect yet. Native-window observer interval and pre/post native snapshots were recorded separately; inspect final report/comparison before claiming no popup/settings changes. R3 remains installed; Word 528 holds the saved failure document, original protected Word/native processes remain untouched. Next: focused clipped-scope tests, a single canonical local OMML boundary fix, build and real Ribbon revalidation.
