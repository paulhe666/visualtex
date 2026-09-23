# Implementation lane — coordination checkpoint

This is the active implementation writer using workspace ws_7e5a0b84a3. The review-lane note and review tests were discovered after the first K0 cargo test. Preserve the other lane's independently named files; no cleanup, reset, or commit.

Current HEAD 00b48e17b736c7e57a0a72e19e271df82906f89d is an observed external pre-engine checkpoint. CP00 original bytes are still guarded with an explicit Git-transition manifest. The implementation lane did not execute that commit.

Implementation-owned files: engine/crates/core/src/**, crates/wasm/src/lib.rs, crates/test-cli/src/main.rs, engine/web/host/protocol.ts, engine/tools/build.mjs and protocol/source parity drivers, engine/crates/core/tests/parser.rs. The next edits extend session.rs/protocol.rs to persist actual parsed source and expose inspectDocument/replaceSource. Then positions/transactions are built on the same single writer. Please avoid racing these files.

Review-owned files observed: tests/source_contract_review.rs, tests/model_contract_review.rs, tests/parser_contract_review.rs, fixtures/k0-review-corpus.json, and the review coordination note. Their current tests have been executed and retained. The observed model serializer changes preventing command concatenation (including scripted bases/lossless mode) and retaining placeholder labels are accepted; do not revert them to the implementation writer's first version.

The public core now exports source/model/parser/symbols modules. Review integration tests can use the real crate instead of #[path] inclusion with stub identity types; this would remove duplicated source-module unit runs and dead-code warnings. This note does not itself edit review files.

Latest cargo test: all executed cases passed (10 core unit + 15 implementation parser + 11 model-review + 9 parser-review + 12 source-review + 2 CLI). Some source unit tests run again through path-included review modules, so this is not 59 unique independent scenarios. K0 source/model are not yet wired into the actual WASM request route; do not report P03 complete until native/WASM parser parity runs.

CP02 source archive and report are immutable under docs/engine-rebuild/checkpoints/P02-source and 20260913-P02-report.md. Future checkpoints must use a different phase directory. Production app, Office code, user documents and installed resources remain untouched. No engine commit/push or package installation is authorized by this note.
