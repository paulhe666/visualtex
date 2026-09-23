# CP07.5 — mathematical alphabets, upright functions and structural style editing

2026-09-13. Independent engine only. Product checkout remains `reference-3d207d7` at `00b48e17b736c7e57a0a72e19e271df82906f89d`. This is a non-Git source checkpoint: no engine Git add/commit/push, application install, production editor switch, Office document operation or release.

## Implemented slice

A real `Style` math node now owns authored wrapper spelling, source trivia, an explicit math variant and an editable row body. Supported product spellings are `\\mathnormal`, `\\mathrm`, `\\mathbf`, `\\mathit`, `\\mathbfit`, `\\boldsymbol`, `\\bm`, `\\mathsf`, `\\mathtt`, `\\mathbb`, `\\mathcal`, `\\mathscr` and `\\mathfrak`.

Layout does not use CSS synthetic bold/italic or replace glyphs after layout. The requested mathematical alphabet participates in the layout cache key and maps supported Latin/digit/Greek atoms to actual Unicode Mathematical Alphanumeric / legacy exception code points in the injected fixed OpenType math font. The surrounding unstyled formula remains on its own variant and cannot inherit a cached styled glyph accidentally.

Common named functions including `\\sin`, `\\cos`, `\\tan`, `\\log`, `\\ln`, `\\exp`, `\\det`, `\\gcd`, `\\arg`, `\\ker`, `\\dim`, `\\hom` are upright Operator atoms with their original command spelling. `\\operatorname{...}` is an editable upright mathematical text operator rendered by the math font, not ordinary `\\text`; scripts remain structural.

`SetStyle` is a mathematical transaction. It wraps the current selection, creates an editable empty body at a collapsed caret, participates in the shared undo/redo timeline, preserves selected child identities and allocates a fresh wrapper identity. Source patch reconciliation, projection, TreeIndex navigation, copying/normalization, scene accessibility paths and layout fingerprinting all understand Style explicitly.

The isolated sandbox exposes style buttons through the same `EngineEditorHost`; no second state store or callback-only test editor was introduced.

## Verification executed

- Strict verification receipt: `engine/target/verification/2026-09-13T08-55-12-854Z/report.json`.
- All engine verification steps before the production-protection guard passed. `build-and-native-wasm` reports 268 Rust tests passed, 0 failed.
- Native/WASM lifecycle parity: 2,151 cases.
- Source-session parity: 222 cases.
- Source-edit parity: 440 cases.
- Editor/history parity: 473 cases including 100 complete history states.
- Style-specific Rust regression: 8/8 passed, covering authored spelling, empty/incomplete source, genuine font glyphs, common functions, `operatorname`, SetStyle history, source-edit identity retention and cache isolation.
- Layout parity after rebuilding both targets: 2,974 requests, 2,030 valid scenes, maximum numeric difference `0`, exact SVG parity under the fixed font resources.
- Real owned Edge browser style regression: PASS. Fourteen recorded browser events were all trusted key/pointer events. It verifies Shift+arrow selection plus style button, bold undo/redo, empty `\\mathrm{}` then typing, selected `\\mathbb{R}`, mixed authored styles/functions/`operatorname`, vector export, incomplete-source refusal and source undo recovery. Report: `engine/target/style-browser/report.json`; screenshot: `engine/target/style-browser/styles.png`.
- Existing root, fence, accent, operator, sandbox, input-fault and SVG geometry browser regressions all passed in the same strict verification run.

Current rebuilt native CLI SHA256: `61a9c06b71705f9fb5fb1ac8f51cd89af17c54d08e6da775e2a270a1fa2cdfe8`.
Current Node-target WASM SHA256: `d6a2aea8b9906809422d721efcfc893bd92e4028365cb00c0eabd5fb0ec331dd`.
Style-inclusive layout transcript SHA256: `332f587ff53d4d133ffd55016a18812e304839e4e1ee0bb878928ea968690513`.
Strict verification source SHA256: `3c5d8f430fa54e15e8b938c958526d9f175440a9b9e2621142bcdfda7aa49194`.

## Production protection alert

The final strict protection step is **FAILED because other concurrent work changed production/Office working-copy files after P00**. At the current observation the external dirty set includes `RibbonVectorIconRenderer.cs`, `PreparedWordBulkFormula.cs`, `ThisAddIn.cs`, `WordFormulaFormatConversion.cs`, `WordFormulaService.FormatConversion.cs`, `WordFormulaService.cs`, `WordInlineObjectGeometry.cs`, `WordRibbonLocalization.cs`, `target-word.json`, plus untracked `WordFormulaService.InlineBaselineRepair.cs`. The tracked diff observed outside this engine lane was 477 insertions / 24 deletions across 9 tracked files.

Those changes were not created, modified, reverted, staged or cleaned by this engine task. Therefore this checkpoint status is **engine/component verification PASS; global Office protection gate ALERT/FAILED**. P07 isolated core work can continue, but no Office integration/P10+ acceptance may claim the P00 guard is green until the concurrent production changes are separately reconciled.

## Gate

P07.5 style/function slice may advance to spacing/phantom/paint structures inside `engine/`. P07 overall is still incomplete. Matrices/cases/aligned, macro/command completion, production main-app switch, Office output, interactive Office window, hidden converter, package and MathLive/MathJax removal remain later gates.
