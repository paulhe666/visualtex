# P09 legacy archive read-compatibility foundation

Status: isolated read/round-trip compatibility implemented. This is not complete product configuration activation and does not mark P09 complete.

## Actual Windows formats inspected

`apps/windows/src/App.tsx` saves JSON with the `.visualtex.json` extension. `apps/windows/src/stores/editorStore.ts::toDocument` emits FormulaDocument version 3 with formula blocks, macros and settings. Its `visualtex-editor` persisted state contains lines, activeLineId, settings, usage and snippet history; an older single `latex` field remains a supported input to its merge function.

`apps/windows/src/runtime/applicationConfiguration.ts` defines `.vtxconfig`, schema `visualtex-user-configuration`, version 1. It contains editorSettings, storage string entries, personalization/history, Word preferences and window metadata. `visualtex-formula-hotkeys-v1` contains the persisted hotkey store envelope (current store version 2); `visualtex.custom-symbols.v1` contains the existing symbol-library envelope. `historyTypes.ts` defines read-only historical DocumentCheckpoint records and old MathLive selection ranges.

No separate binary `.vtx` format is inferred from its name. The codec recognizes actual JSON envelopes by content, independent of a filename. Unknown envelopes are rejected explicitly while the caller retains the input.

## Read and round-trip boundary

`web/host/legacy-archive.ts::LegacyArchive` keeps the complete original envelope privately. Unchanged export returns the exact original string, including BOM and JSON formatting. Public snapshots/projections are deep copies. Formula IDs, display mode, exact source, timestamps, unknown fields, macro strings, settings, history, shortcut templates and custom-symbol data are retained. Unknown TeX and incomplete drafts are not normalized or repaired.

A compare-and-swap `replaceFormulaSource(id, expectedSource, replacement)` updates only the designated LaTeX field. It refuses ambiguous/missing IDs or stale sources. Other fields, including timestamps and document version, remain unchanged. The older single-LaTeX persisted envelope remains single-LaTeX when edited; reading it does not force a schema upgrade. Historical checkpoint records are read-only.

The codec parses bounded JSON only: at most 96,000,000 UTF-8 bytes, depth 48, one million JSON values, 10,000 formula rows and 1,048,576 source UTF-16 units per row. Nested shortcut/library strings are checked when projected. Exceeded limits and invalid data fail explicitly rather than truncating records. Structural core/session budgets still apply when a formula is opened.

Font preference projection preserves the requested string, including unknown future preferences. Explicit document settings take precedence on file-open; separately stored values take precedence on workspace hydration; imported configuration storage takes precedence over current storage. These choices follow the inspected product paths. Reading a font preference does not resolve or activate a browser font. All actual layout still requires registered verified resources.

Old numeric MathLive selections remain archival data only. They are never copied into a core NodeId or selection. Restoring source into a fresh EngineEditorSessionOwner creates fresh core identity and a fresh history; subsequent source/visual changes use the existing single core history. The legacy checkpoint is not replayed as a second engine undo stack.

No persistent storage, live application store, credentials, Word state, OCR installer or operating-system configuration is read or written by the codec. Storage fields are retained data, not commands. Macro/shortcut template strings are not executed by this layer.

## Acceptance

`fixtures/legacy-archive-samples.json` contains synthetic fixtures derived from the actual source schemas, not private user files. `tools/legacy-archive-contract-review.mjs` covers nine contracts, including a thousand-formula round trip, pure config projections, invalid inputs and seven actual WASM formula-session lifecycles.

`tools/legacy-archive-browser-review.mjs` uses owned Edge, trusted load/typing/Ctrl+Z/Y, accepted and unsupported/incomplete formula hosts, unchanged independent content, source-only save, re-open and unchanged browser storage. Unknown/incomplete-source layout errors are expected diagnostics with preserved source, not passed-off rendering successes.

The isolated P09 adapters now cover the remaining pre-integration contracts. `product-compat.ts` resolves retained font preferences only through explicit registered font identities, carries verified custom-text bindings, keeps unknown preferences unresolved, converts persisted hotkeys to core `EditCommand` actions, rejects conflicting chords, and converts only the historical inert `#?` slot marker into a core placeholder before bounded `pasteLatex`. `source-editor-bridge.ts` binds source-view edits to document/revision/source identity, uses opaque core-origin objects to prevent update echo, sends incremental UTF-16 edits through the core source transaction, and delegates undo/redo to the same core history. `codemirror-source-bridge-review.mjs` exercises the existing project `@codemirror/state` transaction model; `p09-product-bridge-browser-review.mjs` exercises verified resources and a trusted persisted shortcut in owned Edge.

Live Zustand/localStorage application, the production React `LatexSourceEditor` wiring, actual file dialogs and the default application entry point remain P10 integration work. P09 does not write those stores or switch the application/Office path.
