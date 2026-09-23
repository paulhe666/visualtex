# P04 transient split-position blocker

Independent randomized transaction regression reproduced an invalid returned PositionMap. Transactions already use the public exported core APIs in these tests.

`cargo test --locked -p visualtex-engine-core --test transactions_contract_review -- --nocapture`

Minimized regression: `replacing_part_of_text_maps_old_carets_past_deleted_temporary_split_nodes`.

Start with `\text{abcd}`, select Text graphemes 1..4 (also 1..3), and PasteLatex `\text{X}`. The result tree/source is valid, but applying the returned PositionMap to every old caret fails InvalidPosition. The command's chosen new selection alone is valid, so checking only that selection misses the bug.

Likely mechanism: ensure_row_range splits the old Text node; TextSplit override maps old internal positions to a newly allocated right/middle text ID. That temporary ID is then removed when the selection is replaced. PositionMap::between computes removed_ids only from old/new trees, so the ephemeral node never receives a Removed fallback. The override therefore returns an ID absent from the final tree.

Need explicit composition of intermediate split/splice maps or a validated fallback for ephemeral split targets. A fallback must resolve to a real final position near the replacement, preserving unrelated suffix positions and affinity. Do not clear all positions to document zero or hide the returned-map error. Every original position, not just the final live selection, is validated by the independent regression.

Original random trace: seed=504020260913a17e, sequence=1, action=67. Source `\text{速度速度2😀度12😀}\text{速度}b\}`, reversed Text selection on Node81 from grapheme10 to1, PasteLatex `\text{速度}`. Failure occurs in valid_after while mapping old_index.carets, after command success.

The randomized test attempts 20 sequences of 80 actions (1600 total), but currently stops at this failure; do NOT report all 1600 as passing until repaired and rerun. Earlier paste comment/no-op defects were fixed and their 12 deterministic tests passed before this new regression.
