//! P04 position mapping independent invariants, including Unicode cluster joins.
use visualtex_engine_core::{model, positions, session, position_map};
use std::sync::Arc;
use model::{AtomClass, Grouping, Node, NodeKind, RowEntry};
use positions::{Affinity, Position, Selection, TreeIndex};
use position_map::{PositionMap, PositionStep, merge_text_map};
use session::NodeId;

fn sym(id: u32) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::Symbol {
        value: "x".into(), class: AtomClass::Ordinary, command: None,
    }})
}
fn row(id: u32, children: Vec<Arc<Node>>) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::Row { grouping: Grouping::Bare,
        entries: children.into_iter().map(|node| RowEntry { leading: String::new(), node }).collect(), trailing: String::new() } })
}
fn text(id: u32, value: &str) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::Text {
        command: "\\text".into(), before_body: String::new(), text: value.into(),
    }})
}
fn rp(id: u32, offset: u32, affinity: Affinity) -> Position { Position::Row { node_id: NodeId(id), offset, affinity } }
fn tp(id: u32, grapheme: u32, affinity: Affinity) -> Position { Position::Text { node_id: NodeId(id), grapheme, affinity } }

#[test]
fn insertion_maps_before_and_after_affinity_to_opposite_edges() {
    let old = row(1, vec![sym(2), sym(3)]);
    let new = row(1, vec![sym(2), sym(4), sym(5), sym(3)]);
    let mapping = PositionMap::between(&old, &new, vec![]).unwrap();
    let index = TreeIndex::new(&new).unwrap();
    assert_eq!(mapping.apply_position(rp(1, 1, Affinity::Before), &index).unwrap(), rp(1, 1, Affinity::Before));
    assert_eq!(mapping.apply_position(rp(1, 1, Affinity::After), &index).unwrap(), rp(1, 3, Affinity::After));
    for affinity in [Affinity::Before, Affinity::After] {
        assert_eq!(mapping.apply_position(rp(1, 0, affinity), &index).unwrap(), rp(1, 0, affinity));
        assert_eq!(mapping.apply_position(rp(1, 2, affinity), &index).unwrap(), rp(1, 4, affinity));
    }
    assert_eq!(mapping.inserted_ids, vec![NodeId(4), NodeId(5)]);
    assert!(mapping.removed_ids.is_empty());
}

#[test]
fn deleting_a_structure_maps_all_its_internal_carets_to_the_deletion_site() {
    let fraction = Arc::new(Node { id: NodeId(3), origin: None, kind: NodeKind::Fraction {
        command: "\\frac".into(), before_numerator: String::new(), numerator: row(4, vec![text(5, "速度")]),
        before_denominator: String::new(), denominator: row(6, vec![sym(7)]),
    }});
    let old = row(1, vec![sym(2), fraction, sym(8)]);
    let new = row(1, vec![sym(2), sym(8)]);
    let mapping = PositionMap::between(&old, &new, vec![]).unwrap();
    let old_index = TreeIndex::new(&old).unwrap();
    let new_index = TreeIndex::new(&new).unwrap();
    for caret in &old_index.carets {
        let mapped = mapping.apply_position(*caret, &new_index).unwrap();
        if matches!(caret.node_id().0, 4 | 5 | 6) {
            assert!(mapped.same_location(Position::row(NodeId(1), 1)));
        }
    }
    let selected = mapping.apply_selection(&Selection::Node { node_id: NodeId(3) }, &new_index).unwrap();
    assert!(selected.is_collapsed());
    assert!(selected.focus().unwrap().same_location(Position::row(NodeId(1), 1)));
}

#[test]
fn joining_combining_mark_snaps_to_new_grapheme_edges_by_affinity() {
    let old = row(1, vec![text(2, "eX")]);
    let new = row(1, vec![text(2, "e\u{301}X")]);
    let map = PositionMap::between(&old, &new, vec![]).unwrap();
    let index = TreeIndex::new(&new).unwrap();
    assert_eq!(map.apply_position(tp(2, 1, Affinity::Before), &index).unwrap(), tp(2, 0, Affinity::Before));
    assert_eq!(map.apply_position(tp(2, 1, Affinity::After), &index).unwrap(), tp(2, 1, Affinity::After));
    assert_eq!(map.apply_position(tp(2, 2, Affinity::After), &index).unwrap(), tp(2, 2, Affinity::After));
}

#[test]
fn joining_emoji_zwj_sequence_never_returns_a_middle_scalar_position() {
    let old = row(1, vec![text(2, "👩🚀x")]);
    let new = row(1, vec![text(2, "👩‍🚀x")]);
    let map = PositionMap::between(&old, &new, vec![]).unwrap();
    let index = TreeIndex::new(&new).unwrap();
    assert_eq!(map.apply_position(tp(2, 1, Affinity::Before), &index).unwrap(), tp(2, 0, Affinity::Before));
    assert_eq!(map.apply_position(tp(2, 1, Affinity::After), &index).unwrap(), tp(2, 1, Affinity::After));
    for grapheme in 0..=3 { for affinity in [Affinity::Before, Affinity::After] {
        index.validate_position(map.apply_position(tp(2, grapheme, affinity), &index).unwrap()).unwrap();
    }}
}

#[test]
fn merging_adjacent_text_retains_correct_relative_grapheme_positions() {
    let old = row(1, vec![text(2, "ab"), text(3, "中😀")]);
    let new = row(1, vec![text(2, "ab中😀")]);
    let map = PositionMap::between(&old, &new, vec![
        merge_text_map(NodeId(2), NodeId(2), "ab", "ab中😀", 0).unwrap(),
        merge_text_map(NodeId(3), NodeId(2), "中😀", "ab中😀", 2).unwrap(),
    ]).unwrap();
    let index = TreeIndex::new(&new).unwrap();
    for affinity in [Affinity::Before, Affinity::After] { for grapheme in 0..=2 {
        assert_eq!(map.apply_position(tp(2, grapheme, affinity), &index).unwrap(), tp(2, grapheme, affinity));
        assert_eq!(map.apply_position(tp(3, grapheme, affinity), &index).unwrap(), tp(2, grapheme + 2, affinity));
    }}
    assert!(map.apply_selection(&Selection::Node { node_id: NodeId(3) }, &index).unwrap().is_collapsed());
}

#[test]
fn split_text_boundary_respects_affinity_and_retains_the_original_left_identity() {
    let old = row(1, vec![text(2, "ab中😀")]);
    let new = row(1, vec![text(2, "ab"), text(3, "中😀")]);
    let map = PositionMap::between(&old, &new, vec![PositionStep::TextSplit { node_id: NodeId(2), right_id: NodeId(3), at: 2 }]).unwrap();
    let index = TreeIndex::new(&new).unwrap();
    assert_eq!(map.apply_position(tp(2, 2, Affinity::Before), &index).unwrap(), tp(2, 2, Affinity::Before));
    assert_eq!(map.apply_position(tp(2, 2, Affinity::After), &index).unwrap(), tp(3, 0, Affinity::After));
    assert_eq!(map.apply_position(tp(2, 4, Affinity::After), &index).unwrap(), tp(3, 2, Affinity::After));
}

#[test]
fn invalid_merge_ranges_are_rejected_without_slicing_inside_unicode() {
    for (old, merged, offset) in [("中", "a中b", 2), ("x", "abc", 1), ("", "abc", 4), ("a", "abc", usize::MAX)] {
        assert!(merge_text_map(NodeId(2), NodeId(3), old, merged, offset).is_err());
    }
}

#[test]
fn unaffected_positions_and_selection_direction_are_preserved() {
    let old = row(1, vec![text(2, "速度"), sym(3), sym(4)]);
    let new = row(1, vec![text(2, "速度"), sym(3), sym(5), sym(4)]);
    let map = PositionMap::between(&old, &new, vec![]).unwrap();
    let index = TreeIndex::new(&new).unwrap();
    let backwards = Selection::Range { anchor: tp(2, 2, Affinity::After), focus: tp(2, 0, Affinity::Before) };
    assert_eq!(map.apply_selection(&backwards, &index).unwrap(), backwards);
    assert_eq!(map.apply_selection(&Selection::Node { node_id: NodeId(3) }, &index).unwrap(), Selection::Node { node_id: NodeId(3) });
}
