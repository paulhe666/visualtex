# CP15 retired runtime complete 1

Status: `complete`.

`p16Permitted=true`, but installer/package/install/release mutation still requires new explicit user approval.

This checkpoint freezes P15: the Windows production desktop and Office entry graphs are core-only, MathLive/MathJax are isolated development oracles, the formal SVG/Presentation MathML export path uses the autonomous core, source validation and package compatibility are VisualTeX-owned, and the final candidate passed browser, performance, fault-recovery and real Word/WebView2 create/edit acceptance.

## Final verification

Final receipt:

`engine/target/p15-verification/2026-09-16T06-09-08-451Z/report.json`

Receipt SHA256:

`e80bba65d1670bb70811bbc77e1cf49da820f5f3ce9ffb5a23583d18ba3228fb`

Verified P15 source digest:

`ec8a234d64884fc4b14d2451a744c13e995d39bac5b828aa73eafd73aa215c76`

The verifier passed 30/30 gates with `sourceUnchanged=true`, `frozenOfficeWorkingTreeUnchanged=true` and `p16Permitted=true`.

## Production dependency removal

`apps/windows/scripts/p15_retired_runtime_audit.mjs` walks all production imports from desktop/main/Office entrypoints. The accepted graph contains 139 reachable source files with zero MathLive package imports, zero MathJax package imports, zero reachable legacy editor/runtime files and zero retired `.ML__*`/MathLive popover CSS selectors.

`mathlive` and `mathjax-full` are dev-only oracle dependencies. They are absent from the production dependency set and absent from both final production bundles. No KaTeX/MathLive font asset or MathJax runtime asset is present in the accepted desktop or Office outputs.

Final desktop script:

`main-I4TNfDC1.js`

SHA256:

`5a72556a09ee2cce1093ed88055f9107b01fae790d3ad5b2c334fdb7c4e10914`

Final Office script:

`dialog-yV0v4sBO.js`

SHA256:

`50aa671ed660b891eda140f6938f6174dadc7ab73d16e344b48c4d503929f634`

The final desktop and Office bundles share the same core WASM SHA256:

`6ee676b2c25a5b3b6cfe16241a208808110e9b7561514c94e4eddca343954ff7`

and the same Latin Modern Math SHA256:

`6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f`.

## Core-only product behaviour

The production workspace no longer selects a legacy editable engine. Unsupported/incomplete source remains core-owned and fail-closed, survives reload, and can be repaired from source view. Outer multi-formula rows use independent core owners with Enter creation, empty-Backspace deletion, core-before-document undo arbitration and document redo.

MathPreview and input-behaviour formula previews use core SVG. Placeholder tiles use core `EmptySlot` layout geometry plus host-only decoration. LatexCopyService no longer calls MathLive SSR validation. Package aliases are renderer-independent compatibility data and are normalized into supported LaTeX primitives before core parsing.

The formal `latexToSvg`/`latexToMathMl` runtime is asynchronous and core-backed. The P15 real-browser export gate covers ordinary structures, multiline/equation+aligned, upright constants, accents, physics aliases, high-order derivatives, vector/operator aliases, matrix quantities, siunitx units/options/ranges, Chinese MathML and fail-closed unknown/placeholder/empty source.

## Real Windows acceptance

Machine-readable record:

`engine/specs/P15-windows-acceptance.json`

SHA256:

`544d67df45ba0407dd74e14b0057bf567df262b7d248480c4b9b9ec85b66b9d8`

A final normal Tauri debug resource build mapped the exact accepted Office production assets into the owned debug companion without touching installed/AppData Office UI. The real Word VisualTeX Ribbon opened WebView2 with `dialog-yV0v4sBO.js`, one core editor, one core input and zero MathLive fields.

On an owned blank Word process, OS keyboard input created one inline VisualTeX OLE from `abc`. Selecting that exact object and invoking `编辑所选公式` restored `abc`; appending `d` and product Ctrl+S updated the same OLE to `abcd`. The final document still had one VisualTeX OLE, 10.5pt source/render font size and Word Position 0, with no duplicate object. The owned Word process was closed without saving and the owned companion was stopped.

## Hardening/performance revalidation

P14 resource corruption recovery, fatal runtime rebuild, Office race isolation, stale OCR ownership, browser input/accessibility, structured clipboard/zoom, 1000-formula virtualization and actual Windows IME all remain passing after P15 removal.

The accepted 187-node browser fixture reports local-edit event-to-next-frame p95 about 13.98ms and cursor event-to-next-frame p95 about 4.12ms. One hundred cursor moves add zero layout calls and zero content revisions. P13 headless conversion still reports zero failures for 100 mixed formulas and 1000 OMML formulas, with the 1000 OMML batch averaging about 1.859ms/formula.

## License/data boundary

The production mathematical font is Latin Modern Math 1.959 with its local GUST Font License/provenance files. Existing rare-integral data retains the STIX Two Math OFL distribution requirement; its distribution regression passes. Retired MathLive/MathJax packages remain development-only and keep their upstream license files outside production runtime assets.

## Frozen Office boundary and repository safety

P15 did not modify the protected WordVsto, PowerPointVsto, Windows Office Contracts, FormulaOleServer/VstoShared/WindowsOleBridge protocol components, Tauri Office server/session/state/backend/pipe files, `src/office/shared`, `src/office/api`, or `src/office/dialog/useOfficeSession.ts`. The verifier confirmed the protected tracked boundary was empty before and after all 30 gates.

Branch remains `reference-3d207d7`; HEAD remains `d1f0aff8db9275c90dc18c7318dc2a4d7bbbd760`.

No reset, clean, stash, branch switch, pull, rebase, merge, commit, push, installer build/install/overwrite or release action was performed. Existing dirty/untracked files were preserved.

## P16 boundary

P16 may now perform read-only/pre-package provenance checks. Building installers, replacing fixed Office subpackages, installing/overwriting Office components, final installed-package Ribbon acceptance, publishing or release mutation is not authorized by this checkpoint and requires explicit user approval.
