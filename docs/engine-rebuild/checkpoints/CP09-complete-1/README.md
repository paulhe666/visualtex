# CP09 complete 1

Stable non-Git checkpoint for completion of P09 in the isolated self-owned VisualTeX engine and sandbox.

Stable source SHA256: `1d36238eae7323c283d7cc48e767638c48aafbaec85f3def65b6ee36c1cc2571`

Verification receipt: `engine/target/p09-verification/2026-09-14T09-27-13-056Z/report.json`

The stable-source P09 gate passed 20/20 steps with `sourceUnchanged=true`. It includes the full mandatory P08 regression plus command registry/input state, bounded macros, matrix-family variants, `\middle`, browser input bridge, structural clipboard, custom-symbol Native/WASM and owned-Edge acceptance, strict deferred OCR target parity/browser acceptance, legacy archive contract/browser compatibility, actual CodeMirror transaction bridging, and the verified font/hotkey product compatibility browser gate.

P09 now owns the source/language/input compatibility boundary. Command confirmation produces core structural transactions; `\partial` symbol and partial-derivative template remain separate. Structural clipboard v2 is identity-free and digest-bound. Custom-symbol artwork supports bounded vector paths/arcs, rounded rectangles, affine fill/stroke, layered `clipRect`, asset-local ordered erase, and legacy text only through an explicit registered content-addressed font binding resolved to real Rust-shaped outlines. No browser `font-family`, raw HTML/SVG/CSS injection, or cross-formula erase is permitted.

Deferred OCR insertion is strict and one-shot: the original session handle, document/revision/source identity, selection epoch and mathematical selection are captured before recognition. Selection changes, edits followed by undo, IME begin/cancel, superseding requests, park/destroy/reopen, another owner/runtime, or a disposed/uncommitted host cannot silently retarget a late result to the current formula.

Legacy FormulaDocument/editor-state/configuration/checkpoint JSON remains readable without schema-forcing upgrades. Persisted font preferences can affect layout only through explicit registered font identities; unknown preferences remain unresolved. Persisted shortcut targets compile to core `EditCommand` actions or a bounded `pasteLatex` fallback for retained custom templates. `SourceEditorBridge` binds CodeMirror-facing updates to document/revision/source identity, tags core-to-view synchronization with opaque origin identity, and delegates source undo/redo to the same core history as visual editing.

This checkpoint is **P09 complete**, not production integration complete. No default VisualTeX application entry point was switched, no production React/CodeMirror store was mutated, and no Office/VSTO/Word path or user Word document was touched. P00 protection intentionally remains red because of independent external Office/VSTO/evidence work. P10 is the next integration phase.

No Git commit, push, pull, merge, rebase, reset, clean, stash or branch switch belongs to this checkpoint.
