# PowerPoint basic-function repair — CP262 — 2026-10-04

CP271范围调整（2026-10-05）：用户明确“不用纠结一次undo”。PPT单次Undo不再阻塞本轮根因修复；现有多步原生撤销与失败恢复保护保留，不把CP270两个单次Undo失败改成通过。专用原生命令探针四种组合均复现记录拆分，产品DLL未变，环境前后一致，详见docs/architecture-root-remediation-20261005.md及artifacts/architecture-root-20261005/cp271-undo-capability-r1。停止扩大撤销研究，继续其余保真问题。

## CP263 — 2026-10-04 — PowerPoint native styles and Ribbon continuation

Final bounded result, 10:24 UTC: native styles 7/7 supported checks pass, with one explicitly unestablished custom-font/mixed-size fixture group; CP262 basics remain 10/10. This is direct service/host acceptance and exported-slide/package evidence. Physical Ribbon clicks remain 0 and are blocked by the current UI/candidate-loading setup described below. It is not complete font/color or Ribbon acceptance.

Final frozen-R3 build: artifacts/omml-edit-audit-20261001/build-cp263-native-final-r2. PowerPoint DLL SHA256 5857E55597B9D5352120F8EA80EFE7294EF5CBEBF651DAC237211493EA459BC3; Word DLL FB031CF2A98E8C2266C90BF20730606985956220A7610A4CC88EC26763885EF2 (unchanged from CP262). Compiled converter contract, basic source contract, double-click safety and git diff --check pass. The pre-existing native static missing-source-string failure recorded in CP262 is not changed or claimed green.

Repairs are confined to PowerPointNativeOmmlKernel.cs and new PowerPointNativeOmmlFormatting.cs. The PPT-boundary adapter maps explicit opaque DrawingML RGB into the shared Word color contract without changing raw source snapshots. Word linearization now retains exact UnicodeMath text and styled color spans instead of discarding color and applying compatibility normalization. Exact inserted spans receive color before native conversion. Upright plain Latin/Greek intent is transferred after the native zone exists and before Professional: the helper requires one indexed math zone, validates the entire text-element sequence, then applies only exact mapped letters in reverse order. Mapping failure throws rather than silently accepting known loss. No broad post-conversion text-frame font write, quoted ordinary-text substitution, shared Word converter change or registry change was used.

Final evidence: artifacts/powerpoint-basic-functions-20261004/styles-final-r2/native-styles.json and its PPTX/OMML/MathML/LaTeX/PNG files; basic-final-cp263-r2/basic-functions.json; contract-final-cp263-r2/converter-contract.json; cp263-final-state.json. The seven supported assertions cover red/blue insertion, native source recovery, color roundtrip, actual host Cambria Math/uniform32pt, mathematical bold-italic scalar preservation, native upright x, and save/reopen. PNG red/blue counts: insertion46/58; native-authored130/227; reinserted128/225; reopened128/225. Baseline insertion/reinsertion were0/0. PNGs are Office-rendered slide exports with pixel-count evidence, not desktop screenshots or a claim of visual CUA inspection.

The arbitrary-font/mixed-size test attempted Arial/Georgia and18/32pt, but the native host stored Cambria Math and uniform32pt before any source recovery. The baseline's false preconditions remain in retained artifacts; the final test reports this group as not established, rather than a plugin failure or pass. A naive baseline expectation of a literal bold-italic MathML attribute was replaced with semantic validation of the actual mathematical bold-italic scalar U+1D49A. The genuine upright-x failure was retained as a failing test until the production repair passed it. No pre-existing regression expectation was weakened.

Remaining limits: arbitrary/custom/mixed per-run faces and sizes, inherited/theme/transformed/transparent colors, colored structural ink (rules/radicals/accents), combining/supplementary upright variants, full native editing fidelity, physical Ribbon/Companion lifecycle and x86 Office are not accepted. The existing source/editor font contract is formula-wide, not an arbitrary per-run typography model. Do not generalize the7/7 to those cases. Shared Word/frozen six-way work and CP255 are unchanged.

Process audit10:24: owned style PowerPoint107984 (start10:22:42.492Z) and basic PowerPoint82776 (start10:23:25.261Z) exited after normal close/quit; no PowerPoint or acceptance process remains. Original Word88900/93828/95356 retain their original start times. No user document, unknown process, installation, registry value, commit or push was changed.

Opening-stage evidence after the user's 09:55 request. HEAD remains 67e57fef62a4abcfa9577fac8849c53817f83893. Baseline owned-host evidence at artifacts/powerpoint-basic-functions-20261004/styles-baseline-r3: colored MathML insertion loses red/blue (saved native XML lacks color; exported slide PNG has 0/0 colored pixels). Native authored red/blue has 130/227 pixels, but recovered MathML and reinserted output lose both (0/0 pixels). Two separate production loss points are raw DrawingML a:rPr handed to the Word reverse transform, and forward Word UnicodeMath materialization reduced to plain NFKC string before PPT insertion. Baseline test also demonstrates test preconditions matter: native host ignored requested Arial/Georgia and normalized attempted mixed 18/32 pt to Cambria Math 32 pt. Those two requested formats were not established, so their failing expectations do not independently prove plugin loss. Mathematical upright/bold-italic source read and write need separate treatment. Baseline r3 retained seven failed assertions and exact source/PNG artifacts; no results hidden or old assertions relaxed.

Actual Ribbon clicks remain unrun: current Windows connector has no native mouse/screenshot surface; existing scripts use UIA patterns or unscoped global input. Registered product path is C:\Program Files\VisualTeX\WindowsOffice\VSTO\VisualTeX.PowerPointVsto.dll (AFE19183CBCB0F3E9EE418F95C45E704078EF7AD801AE67892DC01B314192505), not CP262. HKCU existing LoadBehavior=0 / acceptance auto-load-suppression description, HKLM=3. No registry/installation change or foreground workaround performed. Any physical candidate Ribbon acceptance requires an authorized supported UI and candidate loading mechanism.

## Final bounded outcome, 09:43 UTC

The final frozen-R3 candidate passes 10/10 owned PowerPoint host groups, the compiled converter contract, the new source contracts, double-click safety, and diff hygiene. This is basic service/host acceptance, not physical Ribbon/Companion UI, full native formatting fidelity, all plugin features, or release acceptance. No install, deployment, push or second commit occurred.

User-created source checkpoint: 67e57fef62a4abcfa9577fac8849c53817f83893 (194 files). Existing CP255 candidate, held R4 and unrelated dirty/generated material remain unchanged. The task used the existing Windows Devspace checkout. The initial connector could inspect but not perform Git mutation; the user completed the checkpoint before repairs began.

## Proven common failures and changes

1. PowerPoint's reflection call expected five parameters, but shared Word ReplaceWithPreparedOmmlDirect has six. Exact CP257 product DLLs reproduced MissingMethodException before host mutation. PowerPoint now resolves the six-type signature and passes false for private scratch-range inline promotion. No shared Word converter implementation was changed.

2. Native OMML had no editable transient source or exact native lookup. Native objects intentionally persist no VisualTeX metadata. Passive discovery now captures only document/SlideID/Shape.Id text-shape candidates. Explicit edit/convert captures the exact native branch from a temporary PPTX copy and creates transient source metadata; native references never fall back to other FormulaIds or shape names. AlternateContent fallback pictures with the same ID are excluded. Whole-object replacement requires exactly one equation-bearing paragraph without adjacent text, fields, breaks or other paragraphs.

3. Current selection is authoritative. Previous-formula fallback was removed because blank-slide deselection and Ribbon transitions can both appear as ppSelectionNone. Source hydration occurs after acquiring the operation gate, so repeated Edit cannot invalidate the active session's source snapshot. The shared replacement boundary checks the captured native source before deletion.

4. The explicit font operation previously used a collection/selection span that could include prose. Actual host evidence shows MathZones default Count1/Length22 for a mixed textbox whose entire text has length22. The font operation now uses one indexed zone, only after native command-state confirmation. Read-only is checked before mutation. Passive selection uses no MathZones detection, clipboard copy or package export. Existing Count guards remain; unsafe attempts to restore raw-range/whole-frame writes and Length guards were rejected and were not applied.

5. Replacement preserves user-authored alternative text and title; Tags retain the managed metadata for SVG/OLE. Later Configure calls preserve non-metadata alternative text instead of overwriting it. Native output remains free of VisualTeX identity tags.

## Exact candidate and evidence

Build: artifacts/omml-edit-audit-20261001/build-cp262-ppt-final
- PowerPoint DLL SHA256: 09CF38476F0957592085389B35B2A7D608B8D6A1A728976711EC6328D0275A01
- Word DLL SHA256: FB031CF2A98E8C2266C90BF20730606985956220A7610A4CC88EC26763885EF2
- Held R4 SHA256: 6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC

Final evidence: artifacts/powerpoint-basic-functions-20261004/final-state.json, host-final/basic-functions.json and acceptance.log, final-contract/converter-contract.json. Final-state records source hashes, exact scope and process inventory.

The 10 passed groups are SVG insert/select/edit; SVG font/delete; OLE insert/edit/SVG conversion; native insertion; native explicit source recovery; native font/edit/OMML↔SVG; prose and blank-deselection authority; changed-source/repeated-Edit ownership; mixed-text rejection and neighbor-font preservation; native save/reopen. The native↔SVG lifecycle also preserves user alternative text/title. Fixtures are synthetic SVG/PNG plus a simple native fraction/radical, not visual-fidelity exemplars.

OLE initially returned E_ACCESSDENIED because the newly written test placed preview files outside the server's controlled temp directory. Both PNG and SVG/EMF now use existing production materialization paths. Final OLE passes; no server allowlist/security setting was changed and that fixture failure is not a repaired product bug.

The existing powerpoint_native_omml_regression.mjs still fails its old source-string requirement for IsTextInsertionReference(selection.ObjectId). The same required string is absent in the unchanged base checkpoint; no old assertion was edited. The new basic source regression independently passes the retained nonmutating-detection contract and new ownership rules. An attempted in-memory whole-script baseline wrapper had a quoting error and is not counted as a baseline test run; direct HEAD source verification establishes the missing string.

## Intermediate evidence and corrections

- before-contract-r2 records exact pre-fix MissingMethodException with CP257 Word9D94D119…CC68AF7 / PPTBE6AB48D…84A013. before-contract contains an initial PowerShell reserved-variable harness error and is not product evidence.
- host-before-selection and host-native-source-r2 are under %TEMP%/VisualTeX/acceptance/vsto-flow/artifacts/powerpoint-basic-functions-20261004, because the established harness anchors relative destinations there. Later host runs use explicit checkout-absolute destinations.
- Original native shape-selection failure, duplicate AlternateContent branch failure, collection-span/prose font failure, OLE fixture-path failures and the new caller's empty-forbidden-ID false positive are retained. The latter was corrected by checking the actual reopened transient ID; no legacy assertion was weakened.
- One candidate compile failed on an ambiguous Office Shapes type; explicit PowerPoint alias corrected it.
- A few connector edits returned UNKNOWN; disk inspection established absence before retry.
- The stage briefly explored MathZones-based detection, then replaced it with passive descriptors plus on-command package capture. That candidate was not deployed.

## Remaining limits

Mixed-text or multiple-equation boxes are not accepted for whole-object edit/conversion; doing so would require a range-safe edit/layout contract. Single-zone font changes are separately tested with adjacent prose preserved. There was no physical Ribbon click/Companion editor lifecycle, MathType creation/conversion, screenshot/visual check, x86 Office run, broad matrix or performance test.

Native PPTX can express DrawingML a:rPr formatting that the reused Word reverse transform does not fully interpret. Advanced color, arbitrary/custom faces and mixed per-run size/style fidelity are not accepted by this bounded corpus. Editing such native shapes can lose formatting that is carried only by DrawingML a:rPr; the simple structural fixture does not establish safety for those variants. This is a remaining source-recovery limitation, not a full formatting-fidelity fix. The prior six-way fidelity work remains frozen.

Final process audit contains no PowerPoint or acceptance process and only original Word PIDs88900/93828/95356 at their original start times. Test ownership is verified before creating a presentation; cleanup quits only the owned process when it has no other presentations. No registry add-in suppression or unknown-process kill is used.

## 2026-10-05 CP269–CP270 根因修复回归

最终候选build-cp270-root-final-r2（frozen R3）PPT SHA25630957E9D7BC8CBD710C9A9BCE825408504D83087B7CD13ACAC3E506B0886D406。被动选择不再写标签/名称/纵横比或查重；复制身份修复归于显式编辑，读写及延迟几何恢复使用精确slide/shape引用。文本创建不再复用取消后的选区，提交检查真实作者内容/格式，纯转换先准备，失败仅撤销本次可证明的Undo前缀并检查完整恢复。兼容a14正文与fallback分支按结构区分，未知/多choice拒绝。

同一最终候选基本操作10/10、限定样式7/7通过；新增文本安全/身份15项及原生Undo能力前提1项通过。inline/block的单次Undo两项仍失败（退回线性文本），Redo未运行，不能宣称此操作单次Undo/Redo已验收。任意逐字符字体/混字号前提未建立、实际Ribbon点击仍0。详见artifacts/architecture-root-20261005/cp270-final-r2-{text,basic,styles}/及docs/architecture-root-remediation-20261005.md。所有用户环境核对不变，旧失败保留，未安装/部署/提交。
