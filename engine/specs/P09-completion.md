# P09 completion

Status: complete in the isolated self-owned engine/sandbox. This phase-complete claim does not switch the production VisualTeX editor, Office/VSTO path, OCR recognizer, installer, storage backend or file-dialog entry points; those remain later integration gates.

P09 closes the source/language/input compatibility layer over the frozen P08 multiline/grid foundation:

- `CommandInputState` owns transient `\\` command input, stale-target refusal and candidate confirmation without writing source until one core action is accepted. `\\partial` symbol insertion and the partial-derivative template are distinct actions.
- The fixed built-in macro set, matrix-family environments, one-column cases, `array{cc|c}`, `\\middle`, rare integrals and the existing Windows command registry capability inventory remain source-preserving and structurally editable.
- Structural clipboard schema v2 carries identity-free mathematical structure, LaTeX/text fallbacks and immutable custom-symbol definitions, with digest/schema/budget validation and fresh identities on paste.
- Custom symbols are immutable content-addressed assets with native role/metrics/limits behavior. Path arcs, rounded rectangles, affine fill/stroke, clip rectangles, asset-local ordered erase and explicitly verified text-outline binding render through the self-owned scene/SVG path. No raw authored HTML/SVG/CSS or browser/system font-family fallback is used.
- Deferred OCR returns only LaTeX and inserts through a strict one-shot target bound to the original session handle, document ID, revision, source ID, selection epoch and mathematical selection. Selection changes, edits/undo, IME begin/cancel, parking, disposal, superseding requests and cross-owner/runtime delivery retire the old target instead of retargeting to the current editor.
- Legacy FormulaDocument/editor-state/configuration/checkpoint JSON remains readable without upgrading schema, importing MathLive offsets as NodeIds, executing macro/shortcut strings or rewriting unknown fields. Historical formula source updates use compare-and-swap semantics.
- Persisted font preferences are retained as requested strings and may affect layout only through explicit host mappings to registered content-addressed font identities. Unknown preferences remain unresolved. Custom text artwork likewise requires an exact registered font binding.
- Persisted formula hotkeys compile to core `EditCommand` actions. Known descriptor IDs preserve structural actions; retained custom templates can only become bounded `pasteLatex` transactions after the historical inert `#?` slot marker is converted to `\\placeholder{}`. Duplicate chord ownership is rejected.
- `SourceEditorBridge` is the CodeMirror-facing source boundary. It uses document/revision/source identity, opaque core-origin objects and incremental UTF-16 source edits. Core-to-view synchronization is not echoed as a new edit; source and visual edits share the single core history owner.

## Required acceptance gate

`tools/p09-verify.mjs` is the stable-source gate. It must start from and finish with the same source SHA and pass the mandatory P08 regression plus command inventory/state/browser, macro parity/browser, matrix variants, `\\middle`, input bridge, structural clipboard, custom-symbol Native/WASM and owned-Edge gates, OCR Native/WASM and owned-Edge gates, legacy archive contract/browser, actual CodeMirror transaction bridge and product compatibility browser review.

Only tests actually executed by that gate count as P09-complete evidence. Generated reports live below `engine/target/` and are not source inputs.

## Boundary to P10

P10 may wire these adapters into the production React/CodeMirror editor and settings/storage surfaces. P09 itself does not edit `apps/windows` production code, does not change the default engine, does not alter Office wire DTOs or Word documents, and does not treat sandbox success as production acceptance.
