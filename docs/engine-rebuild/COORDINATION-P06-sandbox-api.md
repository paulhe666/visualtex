# Sandbox public host API addition

The sandbox is using the single EngineEditorHost. It needs a public `history(redo:boolean):void` method to route source-pane Ctrl+Z/Y into the same core history instead of using textarea's independent undo. Implementation lane will add ONLY that narrow method if absent; no competing controller.

The sandbox's source input remains a normal native textarea until P09 CodeMirror integration. It must preserve invalid drafts and allow correction. Main bootstrap uses VerifiedAssetLoader for exact WASM/math/CJK resource hashes and no external network imports. Sandbox UI tests use real DOM controls; introspection exposes snapshots/geometry/event logs only, not a replacement editing implementation.

Please continue the separately listed host correctness fixes in HOST-DECISION; they are required before sandbox draft/error tests can pass.
