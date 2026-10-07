# Installed acceptance feedback repair — 2026-10-05

## Latest delivery — CP301

Repair package: `artifacts/nsis-acceptance-cp300-20261005/VisualTeX_1.2.8_x64_CP300_Acceptance_20261005.exe`;376946517bytes; SHA256 `66A950AE8730198F986F53676BEFE74AF5194FA8BAF57A849D1475BDEADAADDC`. Build and actual reverse extraction passed; not installed. Embedded x64 Word/PPT match the final frozenR3 candidate recorded in CP300. Both Office MSIs (19files each), Office frontend23, NativeRender83 and OCR74 files checked. Companion differs from pre-bundle build only by the expected unique3byte UNK→NSS marker. See `artifacts/nsis-acceptance-cp300-20261005/{build-r2/build-state.json,verified-embedded/report.json,environment-final.json}`.

Six LoadBehavior locations, installed DLLs and original process identities remain unchanged; no new protected process. Original Downloads document SHA remains5726E65007EF18AB0BCACAD44B4FD7E6CE2CF2C8C69A89AA6A3C7D7736C81C91. No deployment/commit/push. Test results and limits are in the CP300 table below: existing CP255 offline failure remains; no physical Ribbon click or full six-way visual claim. The previous indefinite stall was not reproduced, so the measured200 completion does not prove every possible hang resolved. User should retest this exact package against the reported document and entry points.

## Scope and evidence

User reports 16 LaTeX→MathType redraw failures in latex_100_inline_100_display, VisualTeX→MathType appearing stuck for minutes without progress, and native PowerPoint OMML failing to edit by both double-click and Ribbon Edit. Word operation has ended or Word was closed. Required outcome: ordinary formulas convert correctly, redraw/conversion progress is visible, large-document stalls are diagnosed and fixed, and PowerPoint editing retains pure native OMML without persistent extra structures.

Installed Word `9673B5DC13EE854E11BCB627DAEB4B799E52D2D78E480DDBB2044AE4E148E282` and PPT `8500D4BC18CB6761221A89A5581B0ABC43C9ECD4CF1918CF711E45A7F635F492` match CP295. Screenshot, vsto-client log, original process snapshot and installed hashes saved under `artifacts/acceptance-feedback-20261005/`. All prior modifications/failures preserved; HEAD66a8c709. No deployment, registry/security changes or user-document manipulation authorized by this diagnostic workflow. Only owned isolated Office instances may be tested; heldR4/frozenR3 remains in force.

## Initial findings; not yet fixes

- The redraw rejection occurs in `PrepareStandaloneMathTypeOleData` before source deletion. Screenshot shows hbar `mi` style changing to normal and ket fence structure losing its native boundary. Need distinguish writer/readback loss from unsupported representation with raw MTEF and native rendering. Do not delete the content-protection check or normalize away style/structure merely to pass.
- Word `ConvertFormulaFormatAsync` captures the full source synchronously before any progress status. It reports only coarse StatusBar text; UI-thread work can delay that text. Actual200 timing must be measured by capture/render/native preparation/source bind/write/numbering; previous20/40 outside-process measurements are not installed200 runtime acceptance.
- PPT `ReadFormulaAtScreenPoint` explicitly ignores native equations without metadata and `OnNativeDoubleClick` repeats that gate. Pure native OMML intentionally has no stored VisualTeX metadata. This proves the double-click route is excluded. The Ribbon route allows a transient native reference and must be traced separately before assigning the same cause.

Next: create bounded codec/native and PPT entry reproductions, obtain exact200 source corpus without touching the open document, then implement shared fixes and progress reporting with independent post-write checks. No new repair or acceptance result yet.

## CP297 progress

Exact saved corpus reproduced 16/200 failures with installed DLL. Candidate physics-r1 fixed hbar style/identity and direct mfenced bra/ket; 192/200 offline passed. Remaining8 are actual MathJax explicit rows blocked by the shared template-support gate, now corrected pending rebuild. Shared code and new tests: MathTypeMtefCodec.cs, MathTypeDiracTemplateTests.cs, MathTypeLetterlikeNativeStyleTests.cs. No comparator relaxed. New tests10/10; MathType regression919/920 with independent CHAR-decoration invariant failure A/default/tilde, baseline comparison pending. All intermediate failures retained. Candidate Word hash B905C186538D7090A73D7D388C6F5B06CD128682FD09356B68A603531D9026B9; frozenR3 build. Evidence codec-before, codec-after-r1, tests and build-physics-r1.log. New test font assertion was corrected from a nonexistent MathML attribute to actual MTEF font/encoding bytes; no production output requirement removed.

Native evidence authority: https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later (tmDIRAC30, left/right variations and slots); installed AFM MT Extra code104 planckover2pi. Real native rendering pending. External NativePhysicsOracle failed independently without useful results; its failure is not a production failure.

Owned PPT simple standalone native edit callback created correct session in1001ms, cancelled safely and retained pure native XML. Evidence ppt-entry-before/native-entry.json. This does not validate physical clicks, visible editor pixels, commit or the reported button failure. Double-click exclusion is proven in code; text-inserted equations and active native text selection remain to be tested. User exited existing PPT voluntarily. Own Companion PID81460/port43136 remains active for following tests; own PPT34320 was closed. Progress/performance fixes and actual200 Word measurements not yet done.

## CP298 actual200 + progress

Original200 offline 200/200; frozenR3 Word ADC6D6D193EA7581C56F646B84840C1072344DC89A2CFCB4BC99E24D907DBB2A. Owned Word direct redraw65.384s,200 converted/0 skipped,400 post-write/reopen checks passed (word200-direct-r1). Owned original200 VT seed redraw66.349s; actual VT→MT production operation124.703s,600 persisted source/content/display checks passed including saved reopen (word200-vt-mt-r1). No all-symbol/pixel claim. Native preview2.197s for188 unique templates; conversion inclusive write66.738s,preflight29.495s,capture13.506s. Numbering31.781s is nested and cannot be added to these times. This run did not reproduce an indefinite stall.

Independent STA progress window now receives stage/count/time without pumping Office events; six-way rendering reports bounded20 completion. Four bulk/redraw entry families show before capture. Per-item capture/preflight/write trace verified; physical progress-window visual verification pending. Source protection retained.

PPT candidate uses transient native zone coordinates and package order, replaces the native equation's exact range, and native double-click hit-routing; no persistent metadata/extra shape. Actual updated commit runner first failed on missing svg/svgBase64/MathML namespace test payload, not yet passing product result. Builds/attempts preserved through cp298-ppt-r5 and ppt-standalone-r1/r2/r3. Test scope includes standalone/mixed text, editor callback, native commit, hit-routing and reopen; pending results. External accent installed-DLL probe failed and hung separately; owned78756/child93032 cleanup attempt OS denied/errored, recorded in accent-probe-owned-cleanup.json. No security changes.

## CP299 native PPT verification

PPT root repair additionally reuses an existing native MathZone after replacing its text; reinserting an equation toggled native context and caused post-insertion selection failure. Both standalone and mixed text passed edit callback, actual session commit z^3+9, same shape ID/count, surrounding prose/color, native hit routing and saved reopen (ppt-standalone-r5, ppt-mixed-r1). Physical mouse double-click via real production hook also passed in owned PPT (ppt-physical-r1), no direct callback substitution for that test. Physical Ribbon click still untested. No persistent metadata/OLE/extra shapes added.

Candidate cp299-feedback-r1 frozenR3 x64 Word0C247C8F0590015826C36861A132C995C1089A66E488268B41AC77126FB10A8F; PPT7DDD3DAAB78B0F00064C9AF5ED6E28002D116FB063788957FD70CF81A2CEB74A. LoadBehavior6-path before/after unchanged. Existing200 timings are CP298 evidence, not a fresh full run of this candidate. Full offline2423/2426; two SVG sandbox access failures independently resolved by running same test assembly with authorized temp writes (VectorExport6/6). Remaining accent invariant reproduced exactly with read-only HEAD codec in separate output (A/default/tilde226→241), knownCP255 not this patch; no assertion weakened. Baseline runtime output issue and original external probe failure retained. Broader PPT regression/package inspection/progress visual check pending.

## CP300 final bounded regression and packaging

Final frozenR3 candidate `build-cp299-feedback-r3`: Word `096EF484DFEDE92B6906C048429FB660CC102615F2BAE11195740E745B192358`; PPT `7DDD3DAAB78B0F00064C9AF5ED6E28002D116FB063788957FD70CF81A2CEB74A`. Final progress changes own the visible window to Word and prevent an old operation's disposal clearing a newer operation. `progress-window-r1/progress-window.json` independently reads visible native controls: while the Office STA is deliberately blocked for1300ms, stage,37/200 and00:01 refresh; disposal callback passes. This is not a pixel screenshot or rerun of the original200 on the final binary.

| Check | Result and limits | Evidence below artifacts/acceptance-feedback-20261005 |
|---|---|---|
| Original200 installed rejection |16/200 reproduced; corrected shared codec200/200; strict signature retained|codec-before, codec-after-r2|
| Actual Word LaTeX→MT200 |65.384s;200 converted;400 write/reopen checks pass|word200-direct-r1|
| Actual Word VT→MT200 |124.703s;200 converted;600 seed/write/reopen checks pass; no indefinite stall reproduced|word200-vt-mt-r1|
| Progress under blocked Office STA |Visible stage/count/time refresh pass|progress-window-r1|
| Native PPT edit |Standalone/mixed native source, commit, same shape, surrounding text, save/reopen; physical double-click passes; Ribbon callback tested, physical Ribbon not tested|ppt-standalone-r5, ppt-mixed-r1, ppt-physical-r1|
| Independent saved PPTX |3/3: native oMath/sSup, no OLE/VisualTeX identity/customXml; mixed prose preserved|ppt-package-independent.json|
| PPT basics |10/10; precise mixed-zone edit and whole-object conversion rejection both checked|ppt-basic-regression-r2|
| PPT styles |7/7 bounded color/host font/uniform size/semantic style/reopen; arbitrary run fonts/mixed sizes precondition not established|ppt-style-regression-r2/native-styles.json|
| Full offline |2425/2426; pre-existing CP255 wide-accent failure reproduced with HEAD codec|tests/full-regression-authorized.trx, tests/head-accent-baseline-r3.trx|

Original failed attempts remain. The first basic run9/10 expected mixed native reading to fail; that obsolete feature expectation was replaced by stronger positive edit and negative whole-object conversion checks. The first style run was rejected by the owned-instance guard while the previous owned PPT was exiting; no tests executed until the precondition was restored. No comparator or content-freshness boundary was relaxed. Owned Companion81460 and failed external probe processes are now gone; no remaining Word/PPT owned instances (`owned-runtime-cleanup.json`).

Changed production areas: shared `MathTypeMtefCodec.cs` handles native glyph/style authorities and DIRAC templates; Word `WordOperationProgress.cs`, `ThisAddIn.cs`, `ThisAddIn.FormatConversion.cs` and the four operation cores report stage/count/time and render in bounded20 chunks; PPT `PowerPointFormulaService{,.NativeOmml}.cs`, `PowerPointNativeOmmlPackage.cs`, `PowerPointNativeOmmlKernel.cs`, `ThisAddIn.cs` identify native zones transiently and replace only the selected native range. No persistent identity structure was added to OMML. New acceptance runners and codec tests are in the corresponding test projects.

New NSIS build: `artifacts/nsis-acceptance-cp300-20261005`. Initial sandbox CIM read failed before compilation and was retained; authorized reviewed build proceeds into separate `build-r2`, with frozenR3 and exact x64 candidate SHA checks. x64/x86 Office, dependency/Ribbon probes and frontend builds pass; Rust/NSIS/extraction still pending at this checkpoint. No deployment, registry/security modification or commit. Pending scope remains physical Ribbon, multiple native equations in one text box, arbitrary character fonts/sizes, old CP255/F02/F09 and broad six-way visual fidelity. Source-capture13.506s/preflight29.495s/inclusive-write66.738s explain this measured200 run; they do not prove the precise cause of the user's previous stall.
