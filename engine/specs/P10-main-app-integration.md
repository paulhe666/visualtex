# P10 main-app integration contract

Status: browser/main-app candidate implemented; **P10 is not complete** until real Tauri close/relaunch acceptance can be run without disturbing an existing user VisualTeX process.

## Ownership

- The new core owns mathematical source, structural selection, composition and mathematical undo/redo for a selected core document session.
- Zustand remains a product/UI projection and persists the exact source published by core; it does not renormalize accepted core source.
- CodeMirror is a source view over the same core history when a core session is active. It must not keep a second undo stack.
- Legacy MathLive offsets and core node/selection identities are different token families. A `visualtex-core-selection-v1` token is owner-, document-, revision- and source-bound and must fail closed outside that exact owner state.
- Office wire DTOs, Office Session ownership and Word/PowerPoint document algorithms are outside P10 and remain untouched.

## Engine selection

- Engine selection occurs before either editable surface mounts for each explicit desktop editor/document generation.
- The selection preflight evaluates the complete formula owner, immutable custom-symbol snapshot and verified font layout. It may select legacy at the session boundary during migration.
- A mounted core session is locked to core until that editor generation ends. A later layout/resource error may disable the current core view but must never swap the editable tree to MathLive.
- Multi-line product documents are not yet admitted by the first P10 core slice; they select legacy before mount.
- Default production builds do not contain the core editor/WASM/font candidate unless `VITE_VISUALTEX_CORE_EDITOR=1` is explicitly set.

## Natural input behavior

P10 preserves user intent, not MathLive implementation details.

- Superscript/subscript/accent/style auto-exit is driven by the core structural selection path plus explicit `exitStructure`; no DOM/model-offset heuristic is copied from MathLive.
- `End` keeps structural boundary semantics. Leaving a completed structure is an explicit structural navigation action, e.g. ArrowRight where appropriate.
- Auto-paired delimiters use the core delimiter transaction. When disabled, the typed delimiter is an ordinary math input.
- Inline auto-escape uses bounded `replaceRecentMath`; it can consume only recent literal ASCII math input and merges with the same core typing history group so one user typing action remains one undo item.
- Command candidates use the product command data/usage/preferences but the core `CommandInputState` owns query, selection, cancellation and confirmation.
- User custom symbols are frozen as immutable definitions at session creation; editing the library later cannot mutate the current formula asset.

## Fonts and view state

- Font preferences are resolved to real verified font bytes. The bundled Latin Modern Math asset is SHA-bound; Windows system fonts are read only through the bounded `read_system_font_asset` command and re-hashed before registration in core.
- Font/theme/zoom/input-behavior changes are view/layout context changes, not source rewrites and not mathematical history items.
- If a new font preference cannot be resolved during an already selected core session, the session keeps its previous verified layout and surfaces a notice. It does not change engines or source.

## Static preview

- In an explicit core candidate build, safe `MathPreview(staticLayout)` formulas may use a cached, non-interactive core SVG path.
- A cache miss creates one short-lived document owner, exports validated SVG, destroys that owner synchronously and asserts that runtime active-session count returns to the pre-render value.
- No `BrowserInputBridge`, command state or editable input is created for a static preview.
- Placeholders or formulas that cannot be statically rendered by the core remain on the existing preview renderer. Renderer identity is part of the current request; stale SVG from an earlier font/source request must not be displayed.
- Default legacy builds resolve the core preview module to a resource-free stub and must contain no core WASM/font assets.

## Desktop product behavior covered in this checkpoint

- Real Edge product page using the built desktop frontend: keyboard input, structural toolbar action, command input, custom symbol insertion, source projection, CodeMirror/core history, header undo/redo, document-level undo/redo generation changes, settings hot-update, static preview, file save/download and file-input open.
- Deferred OCR uses the existing product recognizer but insertion is scoped by the core one-shot OCR target. P10 does not start or import OCR model/runtime code.
- Windows system-font asset acceptance verifies real Cambria Math and Microsoft YaHei bytes, SHA-256 and face index when present.
- Existing legacy editor, input-behavior, history and application data/configuration regressions remain mandatory.

## Explicitly not covered yet

- Real Tauri process close then second launch for this core candidate. During this checkpoint an installed user VisualTeX process was already running, so dev/single-instance launch must not be used to steal or close it.
- Installer-level no-OCR size/install verification. No no-OCR installer is produced as part of this checkpoint; source/configuration isolation is verified only.
- Office editing or Office document output. Those begin in later phases.

## Stop conditions

- Any mid-session core→MathLive swap.
- Two undo stacks affecting one core formula.
- Browser font-family fallback presented as verified core typography.
- Source mutation caused only by a view/font/theme setting.
- Static previews retaining core editor sessions or input bridges.
- Any P10 validation that requires closing the user's running VisualTeX/Word processes or modifying Office/VSTO/evidence paths.
