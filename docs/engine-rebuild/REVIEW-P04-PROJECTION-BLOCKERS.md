# P04 projection regressions — reproduced and fixed, 2026-09-13

Initial review did not overwrite implementation-owned model.rs/projection.rs. At 12:49 the reviewed model.rs origin-copy block is still unchanged since 12:40; the review lane is applying a narrowly scoped fix to that block and matching source-range offsets now, without changing transaction/session files. Preserve the fixed block when adding later serializers. Exact reproduction: `cargo test --locked -p visualtex-engine-core --test projection_contract_review -- --nocapture`.

Nine tests ran, six passed and three failed. Existing source nodes reused under a new structural parent must honor the new role even when their origin can be copied losslessly:

1. `preserved_unbraced_multi_atom_base_is_grouped_before_attaching_a_script`: a parsed bare Row `a+b` becomes the base of a new Scripts node. Expected `{a+b}^{2}`, actual `a+b^{2}` (changes meaning).
2. `preserved_scripted_base_is_grouped_before_a_second_superscript`: a parsed bare Row `x^2` becomes the base of new superscript 3. Expected `{x^2}^{3}`, actual `x^2^{3}` (invalid double superscript).
3. `moving_an_existing_bare_row_into_fraction_argument_adds_required_braces`: existing bare Row `ab` moved into a new fraction numerator. Expected `\frac{ab}{c}`, actual `\fracab{c}` (changes command token).

Likely shared cause: model::emit immediately copies origin and returns before applying the argument/base-required grouping. Safe preservation depends on both unchanged source content and the new structural context. Adjust the copied-row wrapper and all descendant projected source ranges together, retaining unchanged descendant IDs and spelling. Old untouched formulas such as `\frac12` must continue to round-trip byte-for-byte; do not solve by canonicalizing the entire old document.

Tests additionally verify synthesized control-word separator coordinates, stable identities, exact untouched comments/source, stale ancestor-origin rejection, unsupported opaque projection rejection and duplicate identity rejection. Preserve these tests. The report is a gate failure, not authorization to alter Office code or production entrypoints.

## Resolution at 12:50 +08

The review lane applied a narrow fix in model::emit's existing-origin path: a bare row reused as a fraction argument or grouped script base is wrapped when its new grammar requires it. A safe single original symbol retains its spelling; this keeps `\\frac12` and untouched neighboring unbraced arguments intact. Projected descendant spans shift past the opening brace, and the row itself receives the braced full range.

Added trailing-comment handling so a newly introduced closing brace is not consumed by an old end-of-line comment. A second structural edit initially reproduced the same swallowed brace after the origin shortcut was invalidated; the regular Row serializer now applies the same newline boundary rule. Escaped percent is not treated as a comment.

All 11 projection-review tests now PASS. Full `cargo test --workspace --locked` passed 117 test functions, including 11 position reviews and 8 position-map reviews. Position-map tests currently include the exact source file until its public module is connected; this is native source validation, not a WASM position-map acceptance claim. model.rs SHA256 after review fixes: 0f3668908c123bb6f767228a960eb1a6a8ab2996cc4469364b42ce6dda065396.

The positions review now executes against the public positions API and all 11 tests pass with the implementation writer's fixes. No failing selection or projection tests remain in those suites. P04 transaction/history/IME gates remain incomplete.
