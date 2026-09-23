# P06 implementation scope coordination

P05 checkpoint `CP05-font-layout-parity` is archived (SHA256 d74235a464b9e9f38aeb056c7109587af05bac1e225bea293ce0abf84b6261b9). Combined native Rust tests 189 pass; layout parity 183 requests/106 scenes with Chinese, independent review 519 requests/360 scenes, all passed on identical binary hashes in the CP05 report. This is not real UI/IME acceptance.

This implementation lane now owns:
- `engine/crates/core/src/scene_queries.rs`, corresponding Runtime `queryScene` and scene-token retention in protocol.rs (to avoid re-layout for pointer/selection events).
- `engine/web/host/editor-controller.ts`, `input-bridge.ts` and isolated sandbox assets under `engine/web/sandbox/`.
- browser test/server utilities named `sandbox-*` under engine/tools.

The existing `svg-view.ts` will be reused as the renderer, not replaced. Review lanes should use their own filenames and record issues rather than racing these implementation files. Exact current source will be read before any targeted change. Core geometry/navigation may still receive independent fixes in their named modules; preserve API behavior and tests.

No production application, Office backend/document/installation, macOS/Web project, Git commit or remote mutation is authorized by this work. Real Windows IME is separate from CDP composition replay and must not be falsely labeled passed. All archives remain non-Git.
