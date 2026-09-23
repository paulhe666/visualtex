# P04 transaction review — three failures reproduced and fixed

2026-09-13 ~12:54 +08. Implementation-owned transactions.rs was not edited by review. Please resolve these before core/session integration. Run `cargo test --locked -p visualtex-engine-core --test transactions_contract_review -- --nocapture`.

Initial independent execution ran 12 tests: 9 passed, 3 failed. Resolution is recorded below; retain the original failure details:

1. `a_pasted_trailing_comment_cannot_comment_out_existing_math`: paste `x% keep` before `a+b`. Actual projected source `x% keepa+b`; the tree canonicalizes as `xa+b`, but reopening that source gives only `x`. Expected safe preserved source `x% keep\na+b`.
2. `a_comment_only_paste_cannot_erase_following_source_semantics`: paste `% keep` before `a+b`. Actual `% keepa+b` reopens as an empty formula. Expected `% keep\na+b`.
3. `empty_paste_does_not_allocate_an_unused_root_identity`: PasteLatex with empty string returns changed=false but advances allocator next_unused from 3 to 4. No-op must not consume an invisible temporary root identity.

Likely fixes: preserve the line-comment terminator whenever trailing pasted trivia is transferred in front of an existing row entry, and commit the working allocator only for a changed tree (or early-return empty paste). Do not discard comments or canonicalize the whole document to bypass the issue. model.rs now contains the reviewed end-of-line comment boundary helper for newly wrapped rows; coordinate shared logic rather than introducing divergent % escaping rules.

Passing tests construct the full minimum K0 formula through structural commands (not PasteLatex), two-step fraction boundary deletion, grapheme-safe text deletion, text-edge shift-selection, independent duplication/editing, literal math metacharacters, text split/merge, explicit unsupported capability and atomic failures.

## Resolution at 12:55 +08

Review applied two narrow fixes after the three tests reproduced the failures. In the shared model.rs Row serializer, copied leading trivia ending in an unterminated line comment now receives a newline before the next live math child. This fixes both paste cases and any other relocation of such trivia without dropping comments. In transactions.rs prepare, the speculative allocator is committed only when prepared.changed is true. Other transaction/session logic was left intact.

All 12 transaction-review tests and all 11 projection-review tests now PASS. Transactions and position_map have been publicly exported, so both review test files were changed to import the public core APIs, removing their temporary exact-source includes. Do not count helper-source duplicates as separate implementation coverage.

## Performance observation for later interactive reply design

`node tools/source-performance-review.mjs` ran actual debug Node/WASM. Only 4 nodes reparsed in every case. Formula sizes 72/520/4104 nodes had median patch round trips 1.2795/7.137/52.111 ms; full JSON responses 21122/152161/1233368 bytes. No-op at 4104 nodes was 39.6716 ms. This is not production typing/paint or an optimized-release benchmark. Global lex/reindex/validate, canonical construction and full snapshots remain costly. Do not treat the full inspect-style snapshot response as an accepted high-frequency typing/selection path; use explicit compact change replies/cache invalidation when wiring P04/P05/P06, with separate end-to-end measurements. No performance gate has been weakened.
