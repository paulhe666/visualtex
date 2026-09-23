# P15 — Retired formula-runtime removal

Status: complete for the verified Windows production scope on 2026-09-16. `P16` may begin only after `engine/tools/p15-verify.mjs` passes. P16 installer/package mutation remains separately permission-gated.

## Boundary

P15 removes MathLive and MathJax from the formal Windows production runtime after P14. The legacy MathLive editor and renderer-specific regression helpers may remain in source only as isolated development oracles; they must be unreachable from every production entry and absent from built desktop/Office assets. This stage does not change Word/PowerPoint VSTO, Office Session/server contracts, numbering/reference algorithms, private MathType runtime or macOS Office integration.

## Production editor and previews

`EditorWorkspace` is core-only. The migration selector/fallback is removed from the production path. Unsupported or incomplete source remains owned by the core editor and is preserved fail-closed until repaired; reopening no longer selects MathLive. Outer multi-formula editing is also core-owned: Enter creates another outer formula, empty Backspace removes an outer row, core undo is arbitrated before document-structure undo, redo restores rows and independent owners never cross-write.

`MathPreview` now renders from the autonomous core. Toolbar placeholders use core `EmptySlot` geometry plus host-only preview decoration; no MathLive HTML macro enters the AST, Office export or persistence. The input-behaviour menu uses the same core preview.

## Source validation and package compatibility

`LatexCopyService` no longer imports `mathlive/ssr`. VisualTeX performs its own synchronous draft checks for empty source, balanced groups/environments, trailing scripts, required command arguments and known commands. Registered custom symbols and maintained compatibility aliases remain accepted; complete unknown commands retain the historical `unknown-command` error.

Package compatibility data is renderer-independent. Supported physics/bm/siunitx shorthands are expanded through bounded VisualTeX normalization into standard LaTeX primitives before core parsing. Existing source spelling remains in persisted/copy source. The verified browser export set includes high-order derivatives, vector/operator aliases, physics matrix quantities, siunitx options/ranges/units and the previously supported common forms.

A complete outer `equation`, `equation*` or `displaymath` wrapper is removed only from the export projection because display mode is already supplied by the export API; nested mathematical structures such as `aligned` remain intact. Embedded/non-covering environments are not stripped.

## SVG, MathML and Office output

`src/export/runtime.ts` is core-backed and asynchronous. SVG path geometry and Presentation MathML come from the same autonomous runtime/font resources used by editing. The public SVG result retains padding, natural width/height/baseline, UTF-8 base64, transparent PowerPoint hit target and optional white background semantics. The MathML result strips only core-private Office geometry attributes.

The P15 browser export acceptance covers fraction, root, matrix, multiline, equation+aligned, upright constants, accents, physics, siunitx, Chinese MathML and fail-closed placeholder/unknown/empty source. No returned SVG contains MathJax/MJX renderer material. Office headless conversion remains the P13 core provider for native OLE, OMML, MathType and cross-platform-picture modes including equation tags.

## Dependency and build audit

`apps/windows/scripts/p15_retired_runtime_audit.mjs` walks the import graph from:

- `src/main.tsx`
- `src/desktop/main.tsx`
- `src/office/dialog/main.tsx`

The final audit reaches 139 production source files and reports zero `mathlive`/`mathjax-full` package imports and zero legacy editor/runtime modules. All production CSS reports zero `.ML__*`, MathLive suggestion-popover or retired cloned-popover selectors.

`mathlive` and `mathjax-full` are development-oracle dependencies only. `npm ls --omit=dev mathlive mathjax-full` is empty. Standard and no-OCR frontend contracts use the same autonomous core; no-OCR changes OCR payload capabilities, not formula-engine resources.

Final production resources are:

- desktop JS `main-I4TNfDC1.js`, SHA256 `5a72556a09ee2cce1093ed88055f9107b01fae790d3ad5b2c334fdb7c4e10914`;
- Office JS `dialog-yV0v4sBO.js`, SHA256 `50aa671ed660b891eda140f6938f6174dadc7ab73d16e344b48c4d503929f634`;
- core WASM `visualtex_engine_wasm_bg-DLwqUPYp.wasm`, SHA256 `6ee676b2c25a5b3b6cfe16241a208808110e9b7561514c94e4eddca343954ff7`;
- bundled math font `latinmodern-math-CUFlBgaJ.otf`, SHA256 `6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f`.

Desktop and Office carry byte-identical WASM/font resources. No KaTeX/MathLive font files or MathJax runtime asset is present in either final bundle.

## License/data boundary

The bundled Latin Modern Math 1.959 resource is accompanied in source by `README-Latin-Modern-Math.txt`, `GUST-FONT-LICENSE.txt` and `latinmodern-math.manifest.json`; the README states distribution under the GUST Font License, an instance of LPPL. Its byte identity is the fixed SHA above.

The existing rare-integral vector-data path retains its STIX Two Math OFL distribution requirement and `integral_distribution_regression.mts` passes. MathLive and MathJax remain installed only as explicit development oracles and retain their upstream license files under `node_modules`; the P15 production audit proves they are not formal runtime/package dependencies.

## Hardening and performance revalidation

P14 hardening gates were rerun after the removal: resource corruption recovery, fatal WASM-runtime rebuild, multi-window/Session isolation, double Apply, close-during-export, OCR stale/cancel handling, real browser clipboard/accessibility, 1000-formula virtualization and physical Windows IME all remain PASS.

The current 187-node real-browser local-edit event-to-frame p95 is about 13.98 ms; cursor event-to-frame p95 is about 4.12 ms. One hundred cursor moves add zero layout calls and zero document revisions. P13 production headless performance after P15 completes 100 mixed formulas with zero failures and 1000 OMML conversions at about 1.859 ms/formula, also with zero failures.

## Final real Windows WebView2 acceptance

The final Office production bundle was mapped into the workspace debug companion through the normal Tauri resource build; no installed `%LOCALAPPDATA%` Office UI was overwritten. A dedicated owned Word process PID 44812 and owned debug companion PID 39660 were used only for this acceptance and were both cleaned up afterward.

The real VisualTeX Ribbon opened WebView2 with `dialog-yV0v4sBO.js`, one core editor/input and zero MathLive fields. It loaded the final core WASM/font assets. A read-only Cambria Math system-font probe still returned a real MATH-table face through Tauri's postMessage fallback.

Using OS SendKeys into the foreground WebView2 core input, create source `abc` produced exactly one `VisualTeX.Formula.1` inline OLE at 10.5 pt with Position 0. Selecting that OLE and invoking the real `编辑所选公式` Ribbon restored `abc`; appending `d` and applying updated the same object to `abcd`. Final metadata still reported exactly one VisualTeX OLE, 10.5 pt source/render size, Position 0 and core-derived geometry/baseline. No duplicate object was inserted. The owned Word was closed without saving and the owned companion was stopped; no user Word document was touched.

Machine-readable acceptance is `engine/specs/P15-windows-acceptance.json`.

## P16 gate

P16 may perform only non-mutating/pre-package checks after the P15 verifier passes. Building installers, installing/overwriting Office components, publishing or release mutation still requires explicit user permission under the active workspace constraints. No P15 result is presented as installed-package acceptance.
