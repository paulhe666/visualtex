# Word shortcut conflicts — CP266, 2026-10-04

## Final bounded result, 14:01 UTC

A concrete conflict is confirmed. Current HKCU Software/VisualTeX/OfficeIntegration settings are WordInlineFormulaShortcut=Ctrl+D and WordDisplayFormulaShortcut=Ctrl+F, both OLE. Actual owned Word returns Ctrl+D→FormatFont, Ctrl+F→SmartFind and Ctrl+B→Bold. The installed registered Word DLL's keyboard hook matches the old source hook: read-only Mono.Cecil comparison reports13 methods on each side,0 differences.

That hook matches the VisualTeX chord before Word, invokes its callback and unconditionally returns1, including caught callback exceptions. It does not consult Word key bindings or document-input focus. This explains the configured shortcuts being intercepted. It is not evidence that all Word shortcuts are overwritten merely by installation: formula shortcuts default to null. No production KeyBindings mutation was found in the searched Word/shared/backend code.

The final source candidate protects existing Word bindings and has passed35 compiled managed-double/simulated-routing checks plus9 actual-Word binding checks. It is not installed, and physical keyboard/Ribbon/editor behavior is not accepted.

## Source repair

Changed:
- apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/OfficeFormulaShortcutKeyboardHook.cs
- apps/windows/src-windows/VisualTeX.WordVsto/WordFormulaShortcutPolicy.cs (new)
- apps/windows/src-windows/VisualTeX.WordVsto/ThisAddIn.cs
- apps/windows/src/components/OfficeFormulaShortcutSettings.tsx

The Word policy admits only wdKeyCategoryNil. Commands, macros, styles, fonts, symbols, autotext, disabled keys and prefixes remain with Word. An unreadable key/context is not treated as unassigned.

Actual host testing found that FindKey is dependent on CustomizationContext. The initial AA1C303C candidate missed a document-only binding when the original customization context was restored. Both failing runs are retained, and the second explicitly verifies the fixture is the active document. The final policy checks the original context, current document, attached template, Normal template and loaded global templates. It temporarily selects each lookup context, then restores the original context in finally. It never adds, clears, disables or saves a production key binding. Restoration failure rejects the VisualTeX launch.

Other narrow boundaries:
- Actual focused _WwG must be within the active Word window and its foreground root, in the current process/UI thread
- Alt-only menu accelerators remain with Word; new Word settings require Ctrl. Documented F1–F16 mapping is supported; unknown mappings yield
- AltGr and Windows combinations do not launch formulas
- Same-focus preceding passed-through nonmodifier input is checked for a Word two-stroke prefix. This is conservative history tracking, not a Word prefix-state API
- Only HC_ACTION initial key-down can launch. Queue peeks, repeats and key-up do not
- A boolean callback consumes input only after synchronous launch acceptance. Missing/cancelled lifetime, selection-capture failure, policy denial and exceptions pass through
- True means asynchronous launch accepted, not eventual editor completion
- Rejected Ctrl+V still reaches native paste observation; Ctrl+Alt+V is not ordinary paste observation
- Existing PowerPoint Action constructor remains compatible
- Empty shortcut tokens no longer index an empty string
- UI explains Word priority and how to clear shortcuts. Existing saved values remain unchanged

Separate silent OCR is unchanged. It registers global Ctrl+Alt+O, with Ctrl+Alt+Shift+O fallback, only when enabled; its default is false. This task did not inspect the user's current OCR enablement.

## Build and verification

Final frozen-R3 x64/net472 build:
artifacts/omml-edit-audit-20261001/build-cp266-word-shortcuts-r3

Word DLL SHA256:
827AAE47EE26B13BE515C72E24B4ED226CE54900C88090DAEEE538953D151CEF

Held R4 remains:
6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC

Build command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File artifacts/omml-edit-audit-20261001/Build-Candidate.ps1 -Label cp266-word-shortcuts-r3

Compiled simulated routing35/35 passed again against the final DLL. Cases cover all binding categories, unassigned chords, prefix second stroke, focus changes, unknown F24, F12 modifier mapping, lookup/dispatch exceptions, native paste, AltGr/Windows combinations, HC_NOREMOVE/repeat/key-up, PowerPoint constructor compatibility and BeginSession early rejection. These tests create no Word process or keyboard hook.

Actual owned Word9/9 passed:
1. Ctrl+D remains FormatFont
2. Ctrl+F remains SmartFind
3. Ctrl+B remains Bold
4. Truly unassigned Ctrl+Alt+Shift+A is eligible
5. A document-only Bold binding is protected
6. That document binding remains protected when the original customization context is restored
7. A document's explicitly disabled binding is protected
8. Native equation selection keeps Ctrl+D with Word
9. Native equation selection keeps Ctrl+F with Word

Every actual case independently verifies that CustomizationContext is restored. Native tests call the final compiled binding policy with simulated document-focus eligibility; they do not inject or claim physical input.

Existing Office shortcut frontend regression passed after the settings change. Final targeted git diff --check passed, with line-ending warnings only. No whole-tree or x86 claim; prior CP255 acceptance limits are unchanged.

## Evidence and isolation

- apps/windows/tests/WordShortcutRoutingChecks.cs
- apps/windows/tests/Test-WordShortcutRouting.ps1
- final build/shortcut-routing-final.json
- artifacts/word-shortcut-conflicts-20261004/OwnedShortcutBindings.cs
- artifacts/word-shortcut-conflicts-20261004/Run-OwnedShortcutBindings.ps1
- owned-binding-r5 and r6 retain the initial real document-context failure
- owned-binding-final-r8/summary.json, bindings.json, before.json, after.json and owned-word-process.json retain final evidence
- installed-hook-comparison.json records the13-method installed/source-baseline equality

The original required read returned “user cancelled MCP tool call” at13:01. Work paused without retry/alternate read. After explicit user continuation at13:48, the exact call was retried once successfully.

The standalone native runner invokes only the audited compiled fresh-owned helper, never Program.Main or either registry-suppression scope. It launches /x /Automation with a private seed and acquires COM only from that new PID's own _WwG window. All key-binding test mutations are confined to a throwaway document, discarded on close. No user document or Normal-template binding is edited.

Initial runner preparation failures are retained: shared-read hashing was required for Normal.dotm; Add-Type required explicit COM method interfaces for Close/Quit. A resolver variable bug caused startup logging to fail after launching owned empty-seed PID111652. After verifying exact PID, millisecond start identity and private /x /Automation seed command line, only that owned startup process was terminated. The first cleanup refused a submillisecond timestamp comparison and made no change; the corrected identity check succeeded. This was not a user Word process. Later owned probes close normally; final PID107880 exited.

Final before/after comparisons all match:
- HKCU Word/PPT registration values
- The user's Ctrl+D/Ctrl+F formula settings
- Registered installed Word/PPT DLL hashes
- Normal.dotm SHA256 C2FBAE48829D444F39D42593CC23AF719C9211F628B9862E2EE97271FFDD2CE3
- Original Word88900/93828/95356 process identities and responding state

Registered Word remains5702CC88C7CF5C795FE92D110AD336495CEE2E9F88FC77E5F207561CEC5D6575. No installation, registry write, user-template save, commit/push, reset/clean/stash/branch/pull/rebase/merge occurred. Program.cs and conversion sources were not changed.

## Limits and next step

Physical keyboard dispatch, native _WwG focus/ancestry, dialog/search/backstage boundaries, full editor lifecycle and x86 remain unrun. The native two-stroke fixture was rejected by Word's KeyBindings.Add before that fixture existed; unestablished-prefix-fixture.txt is retained. It is explicitly unestablished, not counted as a native pass. Prefix coverage remains simulated.

Distinct user-owned attached/global-template customizations were not created/tested; final code traverses those contexts conservatively. Native equation tests use an actual selected OMath; caret-inside and immediate boundaries are not accepted. No speculative OMaths.Count exclusion was added.

Prefix history may conservatively suppress one attempt after nonkeyboard cancellation; a changed focused HWND or subsequent nonmodifier key resets it. Unknown preceding mappings also yield. These safety limitations are not full multistroke acceptance.

Immediate manual workaround: clear the two Word formula shortcuts to “None” in VisualTeX, or choose unassigned combinations. There is no need to reset Word's own customizations. The final source candidate must still be deployed through an authorized installation path before it affects the user's installed add-in.

## Primary references

- https://learn.microsoft.com/en-us/office/vba/api/word.application.findkey
- https://learn.microsoft.com/en-us/office/vba/api/word.application.buildkeycode
- https://learn.microsoft.com/en-us/office/vba/api/word.wdkeycategory
- https://learn.microsoft.com/en-us/windows/win32/winmsg/keyboardproc
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-registerhotkey
