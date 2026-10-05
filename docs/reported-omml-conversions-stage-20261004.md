# CP267 — Reported historical OMML conversions, 2026-10-04

## Request and actual screenshots

The user supplied two screenshots from a version preceding the current repairs and asked whether OMML→OLE and OMML→LaTeX problems really exist, suspecting color.

Both original PNGs were materialized through their exact Library identities and their actual pixels inspected. The originals were not edited:
- libfile_ba98e04f54b881918d54998338b9ccd8, 09949B8063B53205AF3FE34067C51941.png: VisualTeX 公式格式转换, “公式格式转换失败: The cross-format replacement boundary is no longer the expected temporary ordinary text character.”
- libfile_bcaf982a65488191a8ba9c223fa921a5, 041B623401DB2868BC25CD4700B00504.png: VisualTeX 公式转为 LaTeX, formula 075aa357-017a-40a8-a12e-5b8721d1c2e6 failed LaTeX writeback verification because actual Word content differed from the expected source.

The screenshots show error dialogs, not the complete input formula, source document topology, conversion selection or exact installed version. Neither alone establishes color as a cause.

## Exact source trace and previous changes

- The first exception is WordOmmlHostWriter.RemovePlainTextBoundaryGuard. Its sole current caller is WordVisualTeXHostWriter.Insert, establishing the VisualTeX OLE branch rather than the separate MathType writer.
- The guard requires the owned temporary x immediately before the replacement OLE, allowing exactly the generated display TAB plus x structure. The handler does not read colors.
- OMML→LaTeX uses WordFormulaService.FormulaToLatexCore → WordFormulaHostMutationKernel.ReplaceSourceWithExternalTargetInActiveTransaction → literal insertion → VerifyLatexSourceRange. The second exception is exact source-text mismatch after normalizing only line endings.
- These checks still intentionally exist. They were not removed or relaxed.
- The earlier commit 74144e8a consolidated display external deletion onto DeleteExact instead of its older partial duplicate. Current source includes that repair. CP257 changed color contracts/converters, not these guard/writeback checks.
- Without the historical version and original input, a matched before/after reproduction of the screenshots is unavailable. CP257's existing color evidence is relevant background, not proof that it fixed either reported dialog.

## Current verified behavior

A new focused acceptance partial class exercises the public conversion APIs in explicitly owned isolated Word, invoking the method directly without Program.Main or either registry-writing auto-load-suppression scope. No user Word instance, user document, template or binding is attached or changed.

Frozen-R3 build: artifacts/omml-edit-audit-20261001/build-cp267-reported-omml-r3
Word DLL SHA256: 827AAE47EE26B13BE515C72E24B4ED226CE54900C88090DAEEE538953D151CEF
Held R4 SHA256: 6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC
HEAD: 67e57fef62a4abcfa9577fac8849c53817f83893

Eight basic cases pass:
- x+y, uncolored and exact #0070C0 blue
- inline and display native OMML
- each converted to VisualTeX OLE and literal LaTeX

Each passed production conversion, target/source object inventory, exact surrounding text, one-step Undo/Redo, and saved-document reopen within the same owned Word process. OLE embedded LaTeX was read back exactly. Colored source and output retain exact #0070C0 LaTeX. No source OMath or temporary x survives successful conversion. The stricter independent artifact audit compares all whitespace and paragraph marks exactly, without stripping them.

Evidence under artifacts/reported-omml-conversions-20261004/:
- current-basic-r3/results.json, summary.json, before-state.json, after-state.json
- source/converted XML and source/converted DOCX per case
- basic-exact-audit.json, Audit-Basic.ps1
- Run-Probe.ps1
- owned-crash-cleanup.json

Native OLE previews in this structural test are explicitly synthetic fixtures. This is not production-renderer, pixel-fidelity, physical Ribbon-click or separate-Word-process reopen acceptance.

## Retained failed attempts and limits

- Initial focused test compilation lacked the WordBulkRun namespace; adding the correct using fixed compilation. No product change.
- The first external reflection invocation passed a PowerShell wrapper instead of a string and failed before Word startup. The typed argument correction is a runner fix.
- current-r2 ran all eight conversions. Its four OLE lifecycle cases passed. All four LaTeX production validators passed, but the outer fixture assertion expected one final paragraph mark while the saved source already contained two. Original results and source XML remain. The fixture expectation was corrected; current-basic-r3 then completed all eight lifecycle cases.
- Two further narrowly selected scenarios did not complete: display with Track Changes enabled and a supplementary U+20000 source. Each external PowerShell reflection host terminated with StackOverflowException after writing the source document/XML and before a result record. They do not count as passes, reproduced screenshot errors, or proven product failures. No matching exception stack was retained by the inspected Application log. They were not retried after this failure.
- Exact owned process cleanup verified PID, full UTC process creation time, executable command and seed path before stopping only the two orphaned test Word processes: 32416 / 2026-10-04T14:08:23.7590519Z and 45948 / 2026-10-04T14:08:28.0277091Z.
- Final audit confirms original Word processes 88900, 93828 and 95356 with unchanged creation identities, unchanged HKCU Word/PPT add-in values and unchanged installed DLL hashes.
- No original affected document, source formula or historical build version was provided. Table/numbered/complex source topologies, those two inconclusive scenarios, exact historical reproduction, broad matrix and visual output remain unverified.

## Outcome and next required input

No production conversion patch was justified by the verified basic cases. Their remaining code is unchanged from the CP266 candidate. This does not assert that the historical report is invalid or already fixed.

The smallest useful next input is the affected Word document, or a minimal copy retaining the problem equation and surrounding paragraph/table structure, ideally together with the old VisualTeX version. Then the two exact operations can be reproduced against current and matching prior code without speculative expansion.

Only focused acceptance code, artifact scripts and this report/history were added. Existing CP262–266 work, CP255 acceptance limits and the frozen R4 resolver are preserved. No install, registry write, runtime deployment, Git mutation, package change or full-release acceptance is claimed.
