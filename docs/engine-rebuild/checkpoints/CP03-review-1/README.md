# CP03 review checkpoint 1

Purpose: immutable non-Git source checkpoint for the user's MathLive replacement work. Does not replace CP02 or assert completion of P03/P04/P05/P06.

Baseline remains local 00b48e17b736c7e57a0a72e19e271df82906f89d. No engine commit/push, installer overwrite, Word save/close, or production edit is authorized or performed by this checkpoint.

Independent review contribution: source_contract_review.rs (9 tests), model_contract_review.rs (8 tests), parser_contract_review.rs (6 tests), and k0-review-corpus.json (41 cases: 25 supported K0, 6 incomplete, 3 invalid, 7 unsupported). Includes 256 deterministic mixed-Unicode lexer samples and 512 recursive K0 syntax samples.

Three reproduced regressions were fixed in model.rs: preserved control-word plus inserted letter merging into a different command, the same boundary with a scripted base, and canonical serialization dropping EmptySlot.label. Tests first failed on the old implementation and remain in the suite.

Core/parser implementation was concurrently produced by another writer; this review does not claim authorship of that source. Tests now use exported public core modules rather than duplicate source inclusion. See COORDINATION-20260913-review.md.

An npm source archive, its SHA256, source-stability result, and final test/protection results will be recorded in the checkpoint audit report. The archive excludes target/.tools/production binaries and font files. The file list is controlled by engine/package.json.
