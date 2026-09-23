# Concurrent workspace observation — review lane

Observed 2026-09-13 12:10–12:11 +08:00 by the continuation/review assistant.

The user requests continued engine implementation, periodic non-Git checkpoints, and no commit/push of engine changes. Checkout remains reference-3d207d7 at 00b48e17b736c7e57a0a72e19e271df82906f89d.

Before this assistant performed any file writes, engine/Cargo.toml and engine/crates/core/Cargo.toml acquired unicode-segmentation and engine/crates/core/src/source.rs appeared at 12:10:49. These are concurrent changes from another writer, not this assistant's implementation. P00 protection check and the P02 2151-case Native/WASM suite passed at 12:10.

To avoid overwriting active work, this assistant initially owns only independently named regression/specification files, starting with engine/crates/core/tests/source_contract_review.rs and engine/fixtures/k0-review-corpus.json. Do not count these as completed parser/transaction implementation. Review tests intentionally verify public behavior and retain failures. Core src/model/parser/session/protocol files are left to the active implementation writer until current state is understood.

## Review lane update, 12:24 +08:00

Implementation coordination note has now been read. Core src/session/protocol/positions and implementation parity drivers remain implementation-owned; review will avoid racing these files. The earlier additive AnalyzeSource API in protocol.rs and protocol.ts is now tested by tools/syntax-parity-review.mjs (305 full Native/WASM equality cases); this was added before reading the implementation ownership note. It is a stateless diagnostics API, not an alternative source owner, and analysis IDs must never become edit anchors. The implementation writer's replaceSource/inspectDocument routes are retained.

All review integration tests now import the public core directly. Current review suites: source 9, model 8, parser 6, session 8. The session suite has passed against replace_source and the real Runtime. Source-session parity (222 cases) has also been executed successfully. No P04 stable local-edit/history claim is made.

Immutable review source checkpoint CP03-review-1 and audit.json exist. Archive SHA256: 0d33cdfdd429b3ad86e496707ae3fb8c71dacb884edd3be23bcff4c013e6e0e2. That checkpoint precedes the API parity work and must not be overwritten. Review will make another directory for subsequent checkpoints.

Next review-owned additions are limited to unowned host infrastructure files (engine/web/host/coordinates.ts, resources.ts, and independent host regression tools), plus independently named core tests. These implement viewport coordinate validation and hash-verified resource loading without changing core API, production assets, or the main app. They are preparatory units for later stages, not P05/P06 acceptance. Please leave those files to this lane; core layout, transactions and bindings stay with implementation.

## Review update, 12:38 +08:00

Read the implementation lane's 20260913-P03-report.md. Independent source_edit_contract_review.rs now adds 8 passing tests, including 2048 deterministic incremental/full-parse differential cases, source-map set invariants, UTF-16/UTF-8 checks, unchanged sibling identities and removal of duplicate formulas. tools/source-edit-parity-review.mjs passed 668 actual Native/WASM requests, including 320 sequential random patches; no deleted NodeId was reused.

Host resources/coordinates units now pass 23 tests. Two newly reproduced edge cases were fixed: an immediate retry must not subscribe to the previous cancelled request generation, and double-encoded asset traversal is rejected. Their tests remain. Manifest budgets have hard upper caps; no missing crypto or MIME is silently bypassed.

New review-owned tools/browser-host-review.mjs has executed in an independently spawned/closed headless Edge 152.0.4191.62 with a temporary profile, loopback-only whitelisted fixture server and CSP. Actual browser WASM started through the hash-verified loader, patchSource replaced Chinese/astral text while retaining denominator identity; stale revision/MIME failures rejected. Browser DPR 1/1.25/1.5/2 produced identical natural coordinates. Real read-only Latin Modern font bytes verified by hash but NOT parsed/shaped by this test. Browser WASM hash 63e1e115ed1f3ea70f6c447333adcce06ba7038c341e0482ef412e39a45b574b. No existing browser/Office processes attached.

Next review focus: independent P04 transaction/selection tests when the public API appears, and browser composition event ordering in the owned browser. CDP composition traces will explicitly remain protocol-level tests, never reported as real Windows system-IME acceptance. Core P04 and layout implementation files remain implementation-owned. Host files and browser test harness stay review-owned to avoid concurrent rewrites.

This note is coordination evidence, not permission to modify production/Office files, stop verification, commit/push, replace packages, close documents, or waive phase gates. No full P04/P05/P06 acceptance is asserted here.
