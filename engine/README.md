# VisualTeX independent math engine

This directory is an isolated Rust workspace. It is not wired into the production editor or any Office session. Follow `../docs/engine-rebuild/VISUALTEX_ENGINE_PROJECT_PLAN.md` and the immutable checkpoint reports before extending it.

## P03 checkpoint scope

Implemented: single-owner sessions, monotonic handles/node/source identities, revision guards, bounded protocol, Native/WASM bindings, K0 lossless source and mathematical trees, recursive parsing, canonical serialization, explicit incomplete drafts and unsupported syntax, context-checked subtree source edits, and identity maps for repeated subexpressions. See `specs/K0-capabilities.md` for the precise language boundary.

Not yet implemented at CP03: mathematical caret/selection transactions and undo, font layout, SVG, real system IME or Office integration. Parse/source-edit tests must not be reported as visual-editor or Office acceptance. CP02's immutable archive remains under the checkpoint directory; later files do not retroactively extend that checkpoint.

## Reproducible Windows build

Run from this directory. Toolchain is pinned locally, not in the production application's root.

```text
rustup target add --toolchain 1.96.0 wasm32-unknown-unknown
cargo install wasm-bindgen-cli --version 0.2.105 --locked --root .tools --target-dir target/bindgen-tool
node tools/build.mjs
```

The build driver tests Rust, builds the native CLI and WASM, generates Node/browser bindings and TypeScript declarations, then actually executes the same request corpus on both targets. Build outputs stay under `engine/target`, `engine/web/pkg` and `engine/web/pkg-node`. Do not copy them over production assets.

Pinned dependencies are in Cargo.lock. Updating dependencies requires rerunning parity and protection checks. CLI build dependencies are separate from core runtime dependencies.

```text
node ../docs/engine-rebuild/tools/inspect-baseline.mjs --check docs/engine-rebuild/checkpoints/20260913-P00-provenance.json --git-transition docs/engine-rebuild/checkpoints/20260913-P00-git-transition.json
node ../docs/engine-rebuild/tools/contract-snapshot.mjs --check
```

The guard resolves paths from its own file, not the current directory.

## Internal boundary

`web/host/protocol.ts` describes internal v1 requests and replies; generated `visualtex_engine_wasm.d.ts` describes the runtime class. Both are unrelated to Office ProtocolVersion 2 or storage schema 1. A host owns one runtime and must free it when disposed. It sends complete copied requests and uses monotonically allocated handles, never pointers into WASM linear memory.

No source, metadata, saved document, registry entry, running application or global Node/Cargo configuration is modified by these tests.
