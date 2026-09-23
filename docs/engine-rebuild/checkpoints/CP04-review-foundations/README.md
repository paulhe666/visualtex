# CP04 review foundations — non-Git checkpoint

Date: 2026-09-13. Baseline remains reference-3d207d7 at 00b48e17b736c7e57a0a72e19e271df82906f89d. This checkpoint records isolated P04 foundations and review evidence. It does not assert that all P04 transactions/history/composition, P05 layout or P06 visual editing have passed.

## Native review results

The latest full Rust run passed 117 test functions. Position review: 11 against exported positions API. Position-map review: 8 against the exact in-progress source module (not yet connected to public core/WASM at checkpoint preparation). Projection review: 11 against the exported core projection API.

Projection review first reproduced three grammar errors when a preserved bare Row was moved under a new fraction/script parent: missing grouping around a+b, a second script on an already scripted base, and command concatenation as `\fracab`. All three are fixed without globally canonicalizing old input. A fourth sequential-edit failure involving an end-of-line comment swallowing a newly introduced closing brace was then reproduced and fixed. Tests retain these cases, source-span reindexing, exact unchanged sibling text and NodeId checks.

See REVIEW-P04-PROJECTION-BLOCKERS.md for exact test names, failure evidence and resolution. Position-map source tests cover affinity at insertion, text split/merge, combining-mark and ZWJ joins, selection direction and deletion fallback.

## Host/browser evidence

23 host resource/coordinate unit tests passed. Resources require build/protocol agreement, hash/MIME/decoded-length verification and fixed resource URL confinement; bounded loading and cancellation/retry/disposal are exercised. These are not font-layout tests.

An independently spawned headless Edge 152.0.4191.62 executed actual web-target WASM behind the verified asset loader. It created a real core session, replaced Chinese/astral text via a subtree source patch, retained denominator identity, rejected stale mutation/MIME mismatch and freed the session. Four emulated DPRs produced identical natural coordinates. Existing browser/Word processes were not attached or closed.

The browser test also captured CDP composition event sequences. It committed Chinese text and then an ordinary character (`速度x`) and separately cancelled preedit to an empty textarea. These are CDP Input.imeSetComposition/Input.insertText traces, NOT physical Windows IME acceptance; compositionend was observed with isTrusted=false. No assertion of real OS IME, mathematical visual editor, font shaping, WebView2 embedding or Office UI is made.

## Source checkpoint

The archive in this directory excludes font files, production binaries, user documents, credentials and build/tool caches. Its hash and source-stability check are recorded separately in audit.json. Do not replace an earlier checkpoint, reset the original worktree, commit, push, install or switch the main editor merely because this checkpoint exists.
