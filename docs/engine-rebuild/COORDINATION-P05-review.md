# P05 continuation/review lane — 2026-09-13

The user requested uninterrupted continuation, periodic non-Git checkpoints and no commit/push of the rebuild. Current product baseline is 00b48e1; production apps/windows has no diff against it.

This lane observed concurrent writes after its first read: math_layout.rs appeared, and layout/structures.rs changed its fraction command matching from dfrac/tfrac to \\dfrac/\\tfrac before an exact-text edit could apply. The attempted edit did NOT apply. Do not attribute that fix to this lane. An actual full build returned a failing display_and_text_fractions_obey_distinct_math_styles test (23.0704 != 36.16) before that concurrent correction. It is not a tool outage.

To avoid collisions, this lane owns separately named P05 review tests (layout_contract_review.rs, scene_contract_review.rs, fonts_contract_review.rs), review runners and review checkpoints. It will not overwrite fonts.rs/layout.rs/scene.rs/protocol.rs or production files while implementation is active. Discovered failures will be preserved and specified here with reproducible commands; targeted fixes require re-reading current code first.

P04-resumed source checkpoint and tests are present; P05/P06 acceptance has not yet been established. Do not count core tests as actual OS IME, Word UI, or installation acceptance.

## Confirmed public-API regressions (actual execution)
`cargo +1.96.0 test --locked -p visualtex-engine-core --test layout_contract_review -- --nocapture` returned 5 passed, 2 failed:
- Empty source produces `export_ready=true` despite `analysis.exportable=false`; SVG export must reject empty or comment-only source. Fix should use the authoritative parse/exportability rule, not a special empty-string exception.
- `scene.vertical(Position::row(numerator.id,0), true, None)` in `\\frac{a+b}{c+d}` returns root row node 1 offset 0 instead of denominator. Pure geometric nearest-neighbor candidates include enclosing rows. Navigation needs structure-aware argument transitions with geometry selecting the horizontal landing point.
The dfrac/tfrac override now passes all four ambient styles, including nested script/fraction cases. This lane did NOT apply that fix. Preserve all 7 review cases.

This lane is next adding scene validation and font resource regression tests under its independent filenames, leaving the core fixes to the active implementation lane initially.
