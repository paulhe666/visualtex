# MathType / Word selection-event coexistence investigation — 2026-09-28

## Status

Two full dumps of the **actual installed** old add-in confirm the same passive-event OLE export call chain, including a second reproduction using normal Ctrl+F4 equation close (no native MathType exit). **R1** completed nine normal native Ribbon insertion/edit-return cycles and saved eleven valid native OLEs. **R2 is now the installed candidate**, closing an additional hidden identity-probe export path; all seven compiled-DLL observation/command boundary checks pass. R2's no-debugger UI test is **blocked before insertion** because its window cannot acquire foreground while another Word document's Save dialog is visible. The unrelated document was not changed or closed. This is **not final acceptance**: R2 native UI repetition, saved-document reopen/edit, VisualTeX/native interleaving, and final NSIS remain pending. Original native installer disappearance and a never-MathType-installed/font-clean environment are still unverified.

Workspace: `C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7`; current workspace ID `ws_6e742f6c2d`. No branch/reset/stash/commit/push. Existing unrelated worktree changes were retained.

## Previous evidence supplied by the user

Completely disabling the VisualTeX add-in, background programs and renderer together stopped the user's observed native MathType Word crashes. Restoring them and opening a real `/x /w` Word reproduced the crash after two equations. The 2026-09-28 00:14:29 event identifies Word PID 58524, `wwlib.dll 16.0.14334.20918`, AV `0xc0000005`, offset `0x2ee4e8`. That test did not run external Word COM polling. The multi-component disabled control alone does not isolate which component is responsible.

## Actual installed baseline dump

Evidence directory:
`artifacts/mathtype-install-20260927/coexistence-20260928/baseline-20260928-003241/`

- Word PID 27508, OS main thread 84460, started with `/x /w` at 00:32:41. Existing hidden Word PID 93104 was left untouched.
- Installed assembly: `C:\Program Files\VisualTeX\WindowsOffice\VSTO\VisualTeX.WordVsto.dll`.
- Installed baseline SHA256: `8F85DB98059CB38D90839B1CB5763C372F795722A3425673172652A9D9307AF4`.
- Loaded MVID: `309ee06e-07b1-46af-9dfb-757e8997ee8e`; add-in trace confirms the installed DLL and all ordinary hooks.
- Native renderer remained Running. No native MathType registration/language/settings changes were made for this trial.
- Physical MathType Ribbon: Display, type `x=1`, return, Right-numbered. Word crashed again at 00:34:15.
- `capture.log`: AV `wwlib+0x2ee4e8`, read address `0x3ac`.
- `word-av-1.dmp`: full local dump, 1,125,153,250 bytes. Kept local; not uploaded.
- `managed-stack.txt`: offline ClrMD analysis using the matching local .NET Framework DAC.

Actual managed main-thread stack leading to the native fault:

```text
ThisAddIn.OnWindowSelectionChange(Selection)
  WordFormulaService.NormalizeTypingCaretAfterInlineFormula(Selection)
    WordFormulaService.NormalizeTypingCaretAfterInlineFormulaCore(Selection)
      WordMathTypeHostAdapter.ResolveLocal(Application, Document, Range)
        MathTypeOleInterop.ReadMetadata(Application, InlineShape, ...)
          MathTypeOleStorage.CaptureCompoundFile(InlineShape)
            MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(...)
              MathTypeWordOpenXml.Read(InlineShape)
                Range.WordOpenXML
                  wwlib.dll + 0x2ee4e8
```

`MathTypeWordOpenXml.cs:308` confirms `Read(InlineShape)` obtains the shape's Range and reads `range.WordOpenXML`. Selection event id 9 logged `ribbon-owner` but never completed `caret-normalization`, corroborating the dump. Exporting the third-party OLE payload was being triggered by a passive selection notification inside native insertion. This initiating call chain is confirmed; the exact private Word internal null-pointer/lifetime implementation is **not** established by export-symbol guesses.

The older diagnostic launch script pointed to a missing `crash-tools\WordCrashCapture.exe`. A separately compiled scoped helper was used instead. Its textual native stack contains nearest exports and cannot resolve JIT frames; use `managed-stack.txt` for managed identities. `dump=1` means the dump succeeded; its printed `err=57` is a stale GetLastError after success.

### Test-method correction

The first baseline used Alt+F4 in the native editor. Wiris documents that as exiting the entire MathType application; Ctrl+F4 closes only the equation and returns to Word. The baseline dump remains evidence of an actual unsafe selection-event call and matching crash, but is labelled an application-exit/relaunch sequence, not a normal close-equation acceptance cycle. The subsequent normal insertion cycles use Ctrl+F4. The stricter old-DLL Ctrl+F4 reproduction was subsequently completed at 01:09 and captured the same crash/call chain; see the completed comparison section below.

Official references:
- https://docs.wiris.com/en_US/menus — File menu distinguishes Ctrl+F4 and Alt+F4.
- https://docs.wiris.com/editing-formatting-equations/mathtype-preferences-problem-when-inserting-equations-into-microsoft-word — known non-fatal Windows preferences-send warning, not evidence by itself of a VisualTeX configuration modification.

## Narrow candidate changes

Only two product files changed in this investigation:

1. `apps/windows/src-windows/VisualTeX.WordVsto/WordFormulaService.cs`
   - Removed MathType semantic-resolution fallback from passive caret normalization. Third-party MathType lifecycle/caret belongs to its active owner. Explicit VisualTeX insertion/edit/conversion paths already restore their own typing caret on completion.
   - Split passive `GetSelectedFormulaFontSize()` from `ReadSelectedFormulaFontSizeForCommand()`. The passive path does not export MathType OLE; exact third-party semantic size remains unknown until an explicit command reads MTEF.
2. `apps/windows/src-windows/VisualTeX.WordVsto/ThisAddIn.cs`
   - Explicit increase/decrease font-size commands use the command-only read.
   - A MathType OLE remains an enabled font-size command target based on its host kind, without reading its payload during Ribbon refresh.

No catch/delay/native-process-name workaround was added. The explicit MathType import/convert adapter and native engine remain intact. Other suspected event structure reads have not been changed without fault evidence.

## R1 build and historical installed-candidate results

Build succeeded via `docs/remediation-3d207d7/build-word.ps1 -Label mathtype-coexistence-event-boundary-20260928`. Existing nullable/interop warnings remain; no unrelated warning cleanup.

`coexistence-20260928/Install-ScopedWordCandidate.ps1` deployed only the built Word DLL/PDB, after checking the baseline hash, equal shared-contracts hash, and no running Word having that DLL loaded. It did not close any Word, touch native MathType, or modify registration/service configuration. This is a diagnostic candidate deployment to the real installed path, **not a newly released MSI/NSIS**.

- Candidate SHA256: `C1A9BCD0539F8778CBAE000185FD9C89E0D175874C290D763473047BC7562118`.
- Original DLL backup and deployment report: `artifacts/mathtype-install-20260927/coexistence-20260928/candidate-r1-deployment/`.
- Candidate Word PID 77076, started 00:44:26, crash debugger attached; separate from untouched Word 93104.
- Evidence: `coexistence-20260928/candidate-r1-20260928-004426/`.

Initial test control pauses are retained as failed/incomplete harness runs, not counted as passed inserts: native preferences-send warning following whole-application exits; the native new-document chapter/section initialization dialog; a UI Automation label containing padding whitespace. The native initialization accepted chapter 1/section 1 only for the new test document and did not select the always-use-default checkbox. Test code now identifies the unique trimmed Ribbon label and uses Ctrl+F4; no product behavior was changed for these test issues.

Normal sequence `native-cycles-2-005153.json`: nine actual physical Ribbon insertions, Left / Display / Right repeated three times, entering x=3 through x=11 and closing the equation with Ctrl+F4. All nine returned to Word. No AV was recorded during the attached capture interval. Use the report's NativeBefore/NativeAfter and per-cycle PIDs for actual process continuity; do not infer process continuity from a generic deployment flag.

## Saved native content validation

Saved by Word's actual Save As UI:
`candidate-r1-20260928-004426/candidate-r1-native.docx`

`Inspect-SavedNativeDocx.ps1` reads the saved ZIP with read sharing and decodes embedded compound-file Equation Native streams offline; it does not inspect live Word or activate OLE.

`saved-native-inspection.json` confirms:
- SHA256 `3F41B640F0807F3BFCF43455EBB89B55B8C53A4558B9B15641BC9722202FE90A`.
- 11 `Equation.DSMT4` compound OLE payloads; all Equation Native streams parse as x=1 through x=11.
- 12 pt MTEF full font size for every object.
- Seven native MathType numbered paragraphs, retaining MTPlaceRef/MTEqn and the native chapter/section structure.
- No VisualTeX-prefixed bookmarks added to the native-only document.

This checks content/structure, not yet full on-screen layout, save/reopen editing, or final coexistence stability.

## Completed stricter old-DLL normal-close reproduction (01:09)

`baseline-close-equation-deployment/installed.json` records restoring the exact saved old DLL solely for comparison, with native MathType PID 38260 and MathTypeLib PID 22688 unchanged. Only the known installed Word DLL was changed; service/registry/native files were not changed. R1's saved test Word was closed normally; the recovery-files prompt was explicitly set to review later, not delete. Original Word PID 93104 was left untouched.

`baseline-ctrlf4-20260928-010917/`:
- Actual installed Word PID 30020, start 01:09:17, debugger attached, no external Word document model calls or polling.
- Physical Display, type x=1, **Ctrl+F4**, returned successfully with native MathType process retained. Physical Right-numbered then crashed Word before another editor could open.
- `capture.log`: thread 48208, `wwlib+0x2ee4e8`, read `0x3ac`.
- `word-av-1.dmp`: 1,161,972,779 bytes, retained locally.
- `managed-stack.txt`: exactly the same `OnWindowSelectionChange → caret normalization → MathType ResolveLocal → ReadMetadata → Read(InlineShape) → WordOpenXML` initiating chain as the first dump.

Thus the initiating fault does not depend on Alt+F4 or on restarting/terminating the native MathType process. Private native Word implementation details remain unknown; the offending add-in call route is proven twice.

## R2 ownership boundary and targeted compiled-DLL checks

Source audit found `MathTypeOleInterop.IsMathTypeOle(InlineShape)` itself falls back to WordOpenXML for missing or unrecognized ProgID. R1 still called this from passive font inspection. R2 reuses a strictly host-identity-only `ReadOleHostKindWithoutContent` helper, and passive font inspection returns unknown for third-party/unknown OLE before any semantic fallback. The explicit-command path still uses the complete identity/MTEF parser, preserving the ability to work without a native MathType registration. It does not guess semantic font size from adjacent Word text.

New focused tests (not modifications of old acceptance behavior):
- `apps/windows/tests/WordPassiveObservationChecks.cs`
- `apps/windows/tests/Test-WordPassiveObservation.ps1`

They load each actual compiled DLL with managed Office-interface doubles. No live Word, COM activation, or user documents are used. Three host identities (known MathType, unknown third-party, missing registration) exercise passive font and caret operations. Assertions require zero content exports, OLE activations, clipboard captures, selection writes, or unmodelled calls. An explicit command must still reach a controlled content-read barrier. This last check proves the command route remains available, not actual successful font editing.

Evidence under `coexistence-20260928/`:
- `passive-boundary-baseline.json`: six passive boundary violations reproduced; known MathType caret also attempts clipboard fallback after the controlled export exception. `Passed=true` here means **expected unsafe old behavior was detected**, not product success.
- `passive-boundary-r1.json`: known MathType fixed, but unknown/missing-registration passive font queries still export. Again run as expected-unsafe diagnostic.
- `passive-boundary-r2.json`: all six passive cases have zero exports/activations/copies/selection writes, no unmodelled calls; explicit command reaches content once. Default safe-mode runner returns **Passed=true** for all seven checks.

R2 compiled successfully with VS2022 MSBuild, Release x64/net472. Existing nullable/interop warnings remain. SHA256: `8E63D124CBA192D0055C692DCDC60D91794DCEDA29741B8F7B06CE224FE48FB5`.

## Current installed state and UI blocker (01:10–01:11)

`candidate-r2-deployment/installed.json`: R2 deployed at 01:10:11 to `C:\Program Files\VisualTeX\WindowsOffice\VSTO\VisualTeX.WordVsto.dll`; old baseline no longer installed. Known previous DLL/PDB backed up. Native MathType PID 38260 / MathTypeLib PID 22688 retained; service Running. No registration changes or process termination by the deployment script.

`candidate-r2-no-debugger-20260928-011021/state.json`:
- Own Word PID 68756, started `2026-09-27T17:10:21.4346851Z`, **no debugger**, no document COM polling.
- `word-hook.log` confirms actual installed R2 loaded at 01:10:22, MVID `c2e6692a-3dd0-4fbc-a875-10f9eb002d4e`.
- Both native-cycle attempts stopped at the foreground check **before sending an insertion input**. Do not count them as passed cycles or as crashes.
- `screenshot-0-011108.png` is a screen crop of the test window's rectangle, not an isolated render: the covering foreground is another document titled `文档1 - MathType Native Numbere...`, showing a `保存此文件` dialog with a Doc1 filename. It is not the current blank `文档2` test. No save/cancel/close command was sent to that document.
- User must resolve that existing dialog before more physical UI tests. Do not bypass it by manipulating the unrelated Word document, guessing its owner, or restarting Word globally.

Current user-visible installation is R2, not old baseline or R1. No new NSIS was built; no installer should be labelled as including this final fix yet.

## Next work / incomplete acceptance

1. After the unrelated Save dialog is handled, resume own Word PID 68756 only if PID/start identity still matches. Run R2 native Display/Right/Left cycles without a debugger, repeated orders, save/reopen and native edit, retaining the native process across Ctrl+F4.
2. Complete actual VisualTeX Ribbon conversion/insertion/edit interleaving, no native popup on VisualTeX operations, exact native OLE/content and numbering, unchanged native process/preferences/registration. Check native font-size command UI because its passive display now intentionally stays unknown.
3. Verify saved-document rendered layout and editing after reopen, not just ZIP/MTEF structure. R1's nine-cycle result does not substitute for final R2 UI acceptance.
4. Build the full final installer package and actually install/retest, with explicit new NSIS path/hash. No old NSIS represents R2. Focused compiled-DLL regression is already passing; do not rerun all old suites after each edit.
5. Clean environment without residual MathType application/fonts, both installation orders, and original downloaded MathType installer exit issue remain open. Do not delete shared fonts or reboot the user's workstation unasked.

All paths in the comparison/test sections, unless otherwise qualified, are relative to `artifacts/mathtype-install-20260927/coexistence-20260928/`. No commit/push was performed.
