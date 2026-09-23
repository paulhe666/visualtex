//! Differential source-edit tests: incremental syntax must agree with a fresh
//! full parse, while unrelated nodes retain their actual editing identities.
use std::collections::BTreeSet;
use std::sync::Arc;
use serde_json::Value;
use visualtex_engine_core::model::{IdAllocator, LatexMode, Node, NodeKind, latex, validate_tree};
use visualtex_engine_core::parser::{self, ParseResult};
use visualtex_engine_core::session::NodeId;
use visualtex_engine_core::source::SourceText;
use visualtex_engine_core::source_edit::{self, ParseScope, PreparedSourceEdit, SourceEdit};
use visualtex_engine_core::ErrorCode;

fn seed(text: &str) -> (Arc<ParseResult>, IdAllocator) {
    let mut ids = IdAllocator::new(2);
    let parsed = Arc::new(parser::parse(SourceText::new(1, text).unwrap(), NodeId(1), &mut ids).unwrap());
    (parsed, ids)
}
fn all_ids(node: &Node) -> BTreeSet<NodeId> {
    let mut ids = BTreeSet::new(); node.visit(&mut |n| { ids.insert(n.id); }); ids
}
fn without_ids(value: &mut Value) {
    match value {
        Value::Object(map) => { map.remove("id"); for child in map.values_mut() { without_ids(child); } },
        Value::Array(items) => for item in items { without_ids(item); },
        _ => {},
    }
}
fn compare(previous: &Arc<ParseResult>, prepared: &PreparedSourceEdit) {
    let source = prepared.document.source.clone();
    let full = parser::parse(source.clone(), previous.tree.id, &mut IdAllocator::new(2)).unwrap();
    let mut incremental_json = serde_json::to_value(&prepared.document).unwrap();
    let mut full_json = serde_json::to_value(&full).unwrap();
    without_ids(&mut incremental_json); without_ids(&mut full_json);
    assert_eq!(incremental_json, full_json, "incremental parse differs from full parse for {:?}", source.text());
    assert_eq!(latex(&prepared.document.tree, LatexMode::Lossless), source.text());
    assert_eq!(validate_tree(&prepared.document.tree).unwrap(), prepared.document.facts);
    let old = all_ids(&previous.tree);
    let new = all_ids(&prepared.document.tree);
    let retained: BTreeSet<_> = prepared.mapping.retained_ids.iter().copied().collect();
    let removed: BTreeSet<_> = prepared.mapping.removed_ids.iter().copied().collect();
    let inserted: BTreeSet<_> = prepared.mapping.inserted_ids.iter().copied().collect();
    assert_eq!(retained.len(), prepared.mapping.retained_ids.len());
    assert_eq!(removed.len(), prepared.mapping.removed_ids.len());
    assert_eq!(inserted.len(), prepared.mapping.inserted_ids.len());
    assert_eq!(retained, old.intersection(&new).copied().collect());
    assert_eq!(removed, old.difference(&new).copied().collect());
    assert_eq!(inserted, new.difference(&old).copied().collect());
    for changed in &prepared.mapping.changed_ids { assert!(new.contains(changed)); }
}
fn edit(previous: &Arc<ParseResult>, ids: &mut IdAllocator, start: usize, end: usize, replacement: &str) -> PreparedSourceEdit {
    let request = SourceEdit { start_utf16: previous.source.byte_to_utf16(start).unwrap(),
        end_utf16: previous.source.byte_to_utf16(end).unwrap(), replacement: replacement.into() };
    source_edit::prepare(previous, &request, previous.source.id() + 1, ids).unwrap()
}

#[test]
fn local_fraction_edit_preserves_unrelated_fraction_and_both_boundaries() {
    let text = "p+\\frac{x_1+\\text{速度}}{y^2}+\\frac{x_1}{y^2}+q";
    let (previous, mut ids) = seed(text);
    let NodeKind::Row { entries, .. } = &previous.tree.kind else { panic!() };
    let untouched: BTreeSet<_> = entries.iter().filter(|e| matches!(e.node.kind, NodeKind::Fraction { .. }))
        .nth(1).map(|e| all_ids(&e.node)).unwrap();
    let start = text.find("x_1").unwrap();
    let changed = edit(&previous, &mut ids, start, start + 1, "z+2");
    compare(&previous, &changed);
    assert!(matches!(changed.mapping.parse_scope, ParseScope::Subtree { .. }));
    let retained: BTreeSet<_> = changed.mapping.retained_ids.iter().copied().collect();
    assert!(untouched.is_subset(&retained));
    assert!(retained.contains(&entries.first().unwrap().node.id));
    assert!(retained.contains(&entries.last().unwrap().node.id));
    assert_eq!(latex(&previous.tree, LatexMode::Lossless), text, "pre-edit snapshot is immutable");
}

#[test]
fn comments_braces_control_words_and_script_markers_cannot_escape_reparse_checks() {
    let original = "a+\\frac{x+{b}}{y_2}+c";
    let at = original.find("x+").unwrap();
    for replacement in ["%", "%\n", "}", "{", "}{", "\\", "\\frac", "\\unknown{a}", "x^", "x_1^2", "😀", "e\u{301}"] {
        let (previous, mut ids) = seed(original);
        let result = edit(&previous, &mut ids, at, at + 1, replacement);
        compare(&previous, &result);
    }
}

#[test]
fn edits_to_unknown_and_incomplete_source_still_match_a_fresh_parse() {
    for original in ["\\unknown[a{b}]{c}+x", "a+\\frac{x}{", "\\", "x^1^2", "a%comment\r\nb"] {
        let boundaries: Vec<_> = original.char_indices().map(|(b, _)| b).chain(std::iter::once(original.len())).collect();
        for &at in &boundaries {
            let (previous, mut ids) = seed(original);
            let result = edit(&previous, &mut ids, at, at, "z");
            compare(&previous, &result);
        }
    }
}

#[test]
fn no_op_edit_uses_the_original_tree_and_consumes_no_identity() {
    let original = "\\frac{速度}{x_1}";
    let (previous, mut ids) = seed(original);
    let next_id = ids.next_unused();
    let start = original.find("速度").unwrap();
    let result = edit(&previous, &mut ids, start, start + "速度".len(), "速度");
    assert!(!result.changed);
    assert!(Arc::ptr_eq(&result.document, &previous));
    assert_eq!(ids.next_unused(), next_id);
    assert!(matches!(result.mapping.parse_scope, ParseScope::None));
    assert_eq!(result.mapping.lexed_tokens, 0);
    assert_eq!(result.mapping.old_range.source_id, result.mapping.new_range.source_id);
}

#[test]
fn source_edit_rejects_split_surrogates_and_reversed_ranges() {
    let (previous, mut ids) = seed("a😀中");
    for (start_utf16, end_utf16) in [(2, 2), (1, 2), (3, 1), (100, 100)] {
        let error = source_edit::prepare(&previous, &SourceEdit { start_utf16, end_utf16, replacement: "x".into() }, 2, &mut ids).err().unwrap();
        assert_eq!(error.code, ErrorCode::InvalidPosition);
    }
    assert_eq!(previous.source.text(), "a😀中");
}

#[test]
fn unicode_replacement_maps_utf16_and_utf8_independently() {
    let original = "\\frac{a😀b}{速度}+q";
    let (previous, mut ids) = seed(original);
    let start = original.find('😀').unwrap();
    let result = edit(&previous, &mut ids, start, start + '😀'.len_utf8(), "👩‍🔬e\u{301}");
    compare(&previous, &result);
    assert_eq!(result.mapping.old_range.byte_end - result.mapping.old_range.byte_start, 4);
    assert_eq!(result.mapping.old_range.utf16_end - result.mapping.old_range.utf16_start, 2);
    assert_eq!(result.mapping.new_range.byte_end - result.mapping.new_range.byte_start, "👩‍🔬e\u{301}".len() as u32);
    assert_eq!(result.mapping.new_range.utf16_end - result.mapping.new_range.utf16_start, "👩‍🔬e\u{301}".encode_utf16().count() as u32);
}

#[test]
fn removing_a_middle_duplicate_does_not_borrow_the_survivors_identity() {
    let original = "\\frac{x}{y}+\\frac{x}{y}+\\frac{x}{y}";
    let (previous, mut ids) = seed(original);
    let NodeKind::Row { entries, .. } = &previous.tree.kind else { panic!() };
    let first = all_ids(&entries[0].node);
    let middle = all_ids(&entries[2].node);
    let last = all_ids(&entries[4].node);
    let start = original.find('+').unwrap();
    let end = start + "+\\frac{x}{y}".len();
    let result = edit(&previous, &mut ids, start, end, "");
    compare(&previous, &result);
    let retained: BTreeSet<_> = result.mapping.retained_ids.iter().copied().collect();
    let removed: BTreeSet<_> = result.mapping.removed_ids.iter().copied().collect();
    assert!(first.is_subset(&retained));
    assert!(last.is_subset(&retained));
    assert!(middle.is_subset(&removed));
}

#[test]
fn seeded_source_edits_match_full_parse_at_every_result() {
    let seeds = ["x+1", "\\frac{x_1+2}{y^2}+z", "\\alpha x", "{{a+b}}", "\\text{速度 😀 e\u{301}}+x",
        "x^{a_b}+y", "\\unknown[foo]{a}+b", "\\frac{}{}", "a%keep\r\nb", "\\frac{x}{"];
    let replacement = ["", "x", "中", "😀", "\u{301}", "\\", "{", "}", "%", "\r\n", "^", "_", "\\frac{a}{b}", "\\text{速度}"];
    let mut random = 0x4b30_5345_5054_2026u64;
    let mut next = || { random ^= random << 13; random ^= random >> 7; random ^= random << 17; random as usize };
    for case in 0..2048 {
        let original = seeds[next() % seeds.len()];
        let boundaries: Vec<_> = original.char_indices().map(|(byte, _)| byte).chain(std::iter::once(original.len())).collect();
        let a = boundaries[next() % boundaries.len()];
        let b = boundaries[next() % boundaries.len()];
        let replace = replacement[next() % replacement.len()];
        let (previous, mut ids) = seed(original);
        let request = SourceEdit { start_utf16: previous.source.byte_to_utf16(a.min(b)).unwrap(),
            end_utf16: previous.source.byte_to_utf16(a.max(b)).unwrap(), replacement: replace.into() };
        let result = source_edit::prepare(&previous, &request, 2, &mut ids)
            .unwrap_or_else(|error| panic!("case={case}, source={original:?}, edit={request:?}: {error}"));
        compare(&previous, &result);
    }
}
