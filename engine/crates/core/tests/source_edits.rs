use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;
use serde_json::Value;
use visualtex_engine_core::ErrorCode;
use visualtex_engine_core::model::{self, IdAllocator, LatexMode, Node, NodeKind};
use visualtex_engine_core::parser;
use visualtex_engine_core::session::{DocumentSnapshot, EngineSession, NodeId, Revision};
use visualtex_engine_core::source::SourceText;
use visualtex_engine_core::source_edit::{ParseScope, SourceEdit};

fn setup(source: &str) -> EngineSession {
    let mut session = EngineSession::new("source-edit-test".into()).unwrap();
    session.replace_source(Revision(0), source.into()).unwrap();
    session
}
fn edit(session: &mut EngineSession, start: usize, end: usize, replacement: &str) -> (DocumentSnapshot, visualtex_engine_core::source_edit::SourceChangeMap) {
    let source = SourceText::new(1, session.source()).unwrap();
    session.patch_source(session.revision(), SourceEdit {
        start_utf16: source.byte_to_utf16(start).unwrap(), end_utf16: source.byte_to_utf16(end).unwrap(), replacement: replacement.into(),
    }).unwrap()
}
fn ids(root: &Node) -> BTreeSet<NodeId> {
    let mut result = BTreeSet::new(); root.visit(&mut |node| { assert!(result.insert(node.id)); }); result
}
fn symbols(root: &Node) -> Vec<(String, NodeId, u32)> {
    let mut result = Vec::new();
    root.visit(&mut |node| if let NodeKind::Symbol { value, .. } = &node.kind {
        result.push((value.clone(), node.id, node.origin.as_ref().unwrap().span.byte_start));
    }); result
}
fn semantic(node: &Node) -> Value {
    let mut value = serde_json::to_value(node).unwrap();
    fn strip(value: &mut Value) {
        match value {
            Value::Object(map) => { map.remove("id"); map.remove("origin"); for value in map.values_mut() { strip(value); } }
            Value::Array(array) => for value in array { strip(value); },
            _ => {}
        }
    }
    strip(&mut value); value
}
fn verify_against_full(snapshot: &DocumentSnapshot) {
    let source = SourceText::new(snapshot.source_id, snapshot.source.as_str()).unwrap();
    let full = parser::parse(Arc::clone(&source), NodeId(1), &mut IdAllocator::new(2)).unwrap();
    assert_eq!(semantic(&snapshot.analysis.tree), semantic(&full.tree), "{}", snapshot.source);
    assert_eq!(serde_json::to_value(&snapshot.analysis.diagnostics).unwrap(), serde_json::to_value(&full.diagnostics).unwrap());
    assert_eq!(snapshot.analysis.facts, full.facts);
    assert_eq!((snapshot.analysis.complete, snapshot.analysis.supported, snapshot.analysis.exportable), (full.complete, full.supported, full.exportable));
    assert_eq!(model::latex(&snapshot.analysis.tree, LatexMode::Lossless), snapshot.source);
    snapshot.analysis.tree.visit(&mut |node| {
        let origin = node.origin.as_ref().unwrap();
        assert_eq!(origin.span.source_id, snapshot.source_id);
        assert_eq!(origin.span.utf16_start, source.byte_to_utf16(origin.span.byte_start as usize).unwrap());
        assert_eq!(origin.span.utf16_end, source.byte_to_utf16(origin.span.byte_end as usize).unwrap());
        assert_eq!(origin.text(), &snapshot.source[origin.span.byte_start as usize..origin.span.byte_end as usize]);
    });
    assert_eq!(ids(&snapshot.analysis.tree).len(), snapshot.analysis.facts.nodes);
}

#[test]
fn only_closed_numerator_is_reparsed_and_duplicate_siblings_keep_identity() {
    let mut session = setup(r"\frac{x+x}{y+y}");
    let before = session.document_snapshot();
    let before_symbols = symbols(&before.analysis.tree);
    let (after, mapping) = edit(&mut session, 6, 7, "a");
    assert_eq!(after.source, r"\frac{a+x}{y+y}");
    let ParseScope::Subtree { node_id, parsed_nodes } = mapping.parse_scope else { panic!("must parse the closed numerator only") };
    assert!(parsed_nodes < before.analysis.facts.nodes);
    assert!(mapping.retained_ids.contains(&node_id));
    assert_eq!(mapping.removed_ids, vec![before_symbols[0].1]);
    let after_symbols = symbols(&after.analysis.tree);
    assert_ne!(after_symbols[0].1, before_symbols[0].1);
    for index in 1..before_symbols.len() { assert_eq!(after_symbols[index].1, before_symbols[index].1); }
    assert!(mapping.lexed_tokens > parsed_nodes);
    verify_against_full(&after);
}

#[test]
fn insertion_shifts_source_ranges_without_reassigning_identical_neighbors() {
    let mut session = setup("x+x+x");
    let old = symbols(&session.document_snapshot().analysis.tree);
    let (after, map) = edit(&mut session, 2, 2, "y+");
    assert!(matches!(map.parse_scope, ParseScope::Formula { .. }));
    let new = symbols(&after.analysis.tree);
    assert_eq!(new[0].1, old[0].1); assert_eq!(new[1].1, old[1].1);
    assert_eq!(new[4].1, old[2].1); assert_eq!(new[6].1, old[4].1);
    assert_eq!(new[4].2, old[2].2 + 2);
    assert_eq!(map.inserted_ids.len(), 2); assert!(map.removed_ids.is_empty());
    verify_against_full(&after);
}

#[test]
fn empty_argument_retains_its_row_while_a_new_atom_is_inserted() {
    let mut session = setup(r"\frac{}{y}");
    let before = session.document_snapshot();
    let mut empty_row = None;
    before.analysis.tree.visit(&mut |node| if node.is_empty_row() { empty_row = Some(node.id); });
    let (after, map) = edit(&mut session, 6, 6, "x");
    let ParseScope::Subtree { node_id, .. } = map.parse_scope else { panic!() };
    assert_eq!(Some(node_id), empty_row);
    assert!(map.removed_ids.is_empty()); assert_eq!(map.inserted_ids.len(), 1);
    assert!(after.analysis.exportable);
    verify_against_full(&after);
}

#[test]
fn unicode_text_edit_has_scalar_boundaries_and_preserves_text_node_identity() {
    let mut session = setup("\\frac{\\text{速度😀e\u{301}}}{y}");
    let before = session.document_snapshot();
    let mut text_id = NodeId(0);
    before.analysis.tree.visit(&mut |node| if matches!(node.kind, NodeKind::Text { .. }) { text_id = node.id; });
    let start = session.source().find('😀').unwrap();
    let (after, map) = edit(&mut session, start, start + '😀'.len_utf8(), "🌍");
    assert!(map.retained_ids.contains(&text_id));
    assert!(matches!(map.parse_scope, ParseScope::Subtree { .. }));
    verify_against_full(&after);
    let scalar_start = after.source[..start].encode_utf16().count() as u32;
    let before_failure = serde_json::to_value(session.document_snapshot()).unwrap();
    let failure = session.patch_source(session.revision(), SourceEdit { start_utf16: scalar_start + 1, end_utf16: scalar_start + 1, replacement: "x".into() });
    assert_eq!(failure.unwrap_err().code, ErrorCode::InvalidPosition);
    assert_eq!(serde_json::to_value(session.document_snapshot()).unwrap(), before_failure);
}

#[test]
fn lexical_or_brace_boundary_changes_use_full_formula_fallback() {
    for replacement in ["%", "\\", "}{", "{"] {
        let mut session = setup(r"\frac{x}{y}+z");
        let (after, map) = edit(&mut session, 7, 7, replacement);
        assert!(matches!(map.parse_scope, ParseScope::Formula { .. }), "{replacement}");
        verify_against_full(&after);
    }
    let mut session = setup(r"\frac{x}{y}");
    let (after, map) = edit(&mut session, 5, 6, "");
    assert!(matches!(map.parse_scope, ParseScope::Formula { .. }));
    verify_against_full(&after);
}

#[test]
fn source_draft_recovery_and_opaque_neighbors_preserve_input() {
    let mut session = setup(r"\frac{x}{");
    let end = session.source().len();
    let (after, map) = edit(&mut session, end, end, "y}");
    assert!(after.analysis.exportable); assert!(matches!(map.parse_scope, ParseScope::Formula { .. }));
    verify_against_full(&after);
    let mut session = setup(r"\future[z]{opaque}+\frac{x}{y}+\alpha");
    let old = session.document_snapshot();
    let unknown_id = old.analysis.tree.children()[0].id;
    let start = session.source().rfind("{x}").unwrap() + 1;
    let (after, map) = edit(&mut session, start, start + 1, "xx");
    assert!(!after.analysis.supported); assert!(map.retained_ids.contains(&unknown_id));
    verify_against_full(&after);
}

#[test]
fn no_op_and_all_validation_failures_leave_session_unchanged() {
    let mut session = setup(r"\frac{x}{y}");
    let previous_ids = ids(&session.document_snapshot().analysis.tree);
    let initial = serde_json::to_value(session.document_snapshot()).unwrap();
    let revision = session.revision();
    let (after, map) = edit(&mut session, 6, 7, "x");
    assert!(matches!(map.parse_scope, ParseScope::None)); assert_eq!(map.lexed_tokens, 0);
    assert_eq!(after.revision, revision); assert_eq!(serde_json::to_value(after).unwrap(), initial);
    for bad in [
        SourceEdit { start_utf16: 7, end_utf16: 6, replacement: "".into() },
        SourceEdit { start_utf16: 999, end_utf16: 999, replacement: "".into() },
        SourceEdit { start_utf16: 6, end_utf16: 7, replacement: "x".repeat(9000) },
        SourceEdit { start_utf16: 6, end_utf16: 7, replacement: " ".repeat(262145) },
    ] {
        assert!(session.patch_source(revision, bad).is_err());
        assert_eq!(serde_json::to_value(session.document_snapshot()).unwrap(), initial);
    }
    assert_eq!(session.patch_source(Revision(0), SourceEdit { start_utf16: 6, end_utf16: 7, replacement: "q".into() }).unwrap_err().code, ErrorCode::StaleRevision);
    let (changed, map) = edit(&mut session, 6, 7, "q");
    assert!(map.inserted_ids.iter().all(|id| id.0 > previous_ids.iter().map(|id| id.0).max().unwrap()));
    verify_against_full(&changed);
}

#[test]
fn deterministic_random_edits_match_full_parse_with_unique_live_ids() {
    let pieces = ["x", "+", "_2", "^", "{", "}", "\\frac", "\\alpha ", " ", "%c\n", "中", "\\future", "\\", ""];
    let mut seed = 0x9a34_7281_u64;
    let mut session = setup(r"\frac{x+x}{y^2}+\text{速度}");
    let mut historical_max = ids(&session.document_snapshot().analysis.tree).last().unwrap().0;
    let mut old_ids = ids(&session.document_snapshot().analysis.tree);
    for _ in 0..350 {
        seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
        let text = session.source().to_owned();
        let mut boundaries: Vec<_> = text.char_indices().map(|(i, _)| i).collect();
        boundaries.push(text.len());
        let start_index = (seed >> 32) as usize % boundaries.len();
        seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
        let end_index = (start_index + (seed >> 32) as usize % 3).min(boundaries.len() - 1);
        let replacement = pieces[seed as usize % pieces.len()];
        let source = SourceText::new(1, text.as_str()).unwrap();
        let result = session.patch_source(session.revision(), SourceEdit {
            start_utf16: source.byte_to_utf16(boundaries[start_index]).unwrap(),
            end_utf16: source.byte_to_utf16(boundaries[end_index]).unwrap(), replacement: replacement.into(),
        });
        match result {
            Ok((snapshot, map)) => {
                let new_ids = ids(&snapshot.analysis.tree);
                assert_eq!(map.retained_ids.iter().copied().collect::<BTreeSet<_>>(), old_ids.intersection(&new_ids).copied().collect());
                assert_eq!(map.removed_ids.iter().copied().collect::<BTreeSet<_>>(), old_ids.difference(&new_ids).copied().collect());
                for id in &map.inserted_ids { assert!(id.0 > historical_max); }
                historical_max = historical_max.max(new_ids.last().unwrap().0);
                old_ids = new_ids;
                verify_against_full(&snapshot);
            }
            Err(error) => { assert_eq!(error.code, ErrorCode::ResourceLimit); assert_eq!(session.source(), text); }
        }
    }
}

#[test]
fn reusable_source_edit_api_preserves_allocator_on_failure() {
    let initial = setup(r"\frac{x}{y}").document_snapshot().analysis;
    let request = SourceEdit { start_utf16: 6, end_utf16: 7, replacement: "qq".into() };
    let mut ids = IdAllocator::new(u32::MAX - 1);
    let result = visualtex_engine_core::source_edit::prepare(&initial, &request, 3, &mut ids);
    assert!(matches!(result, Err(ref e) if e.code == ErrorCode::ResourceLimit));
    assert_eq!(ids.next_unused(), u32::MAX - 1);
    assert!(visualtex_engine_core::source_edit::prepare(&initial, &request, initial.source.id(), &mut ids).is_err());
    assert_eq!(ids.next_unused(), u32::MAX - 1);
}

#[test]
fn independent_numerators_never_exchange_identity_when_repeated() {
    let mut session = setup(r"\frac{x}{y}+\frac{x}{y}");
    let before = session.document_snapshot();
    let before_by_position: BTreeMap<_, _> = symbols(&before.analysis.tree).into_iter().map(|(s, id, pos)| ((s, pos), id)).collect();
    let (after, _) = edit(&mut session, 6, 7, "q");
    for (symbol, id, position) in symbols(&after.analysis.tree) {
        if position != 6 { assert_eq!(Some(&id), before_by_position.get(&(symbol, position))); }
    }
    verify_against_full(&after);
}
