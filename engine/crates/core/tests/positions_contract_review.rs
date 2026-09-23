//! P04 review against the public core position/model APIs.
pub use visualtex_engine_core::{model, session, positions};
use std::sync::Arc;
use model::{Grouping, IdAllocator, Node, NodeKind, RowEntry};
use positions::{Affinity, Horizontal, NormalizedRange, Position, Selection, TreeIndex, grapheme_byte};
use session::NodeId;
use visualtex_engine_core::parser;
use visualtex_engine_core::source::SourceText;
use visualtex_engine_core::ErrorCode;

fn tree(source: &str) -> Arc<Node> {
    parser::parse(SourceText::new(1, source).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap().tree
}
fn text_node(tree: &Node) -> NodeId {
    let mut id = None;
    tree.visit(&mut |node| if matches!(node.kind, NodeKind::Text { .. }) { id = Some(node.id); });
    id.unwrap()
}
fn assert_text_range(range: NormalizedRange, node_id: NodeId, start: u32, end: u32) {
    assert_eq!(range, NormalizedRange::Text { node_id, start, end });
}

#[test]
fn shift_right_from_before_text_selects_one_grapheme_not_the_entire_run() {
    let root = tree("\\text{abc}+x");
    let id = text_node(&root);
    let index = TreeIndex::new(&root).unwrap();
    let selection = index.move_horizontal(&Selection::caret(Position::row(root.id, 0)), Horizontal::Right, true).unwrap();
    assert_text_range(index.normalize_range(&selection).unwrap(), id, 0, 1);
}

#[test]
fn shift_left_from_after_text_selects_one_grapheme_not_the_entire_run() {
    let root = tree("\\text{abc}+x");
    let id = text_node(&root);
    let index = TreeIndex::new(&root).unwrap();
    let selection = index.move_horizontal(&Selection::caret(Position::row(root.id, 1)), Horizontal::Left, true).unwrap();
    assert_text_range(index.normalize_range(&selection).unwrap(), id, 2, 3);
}

#[test]
fn visually_equivalent_text_and_row_edges_never_normalize_to_deleting_text() {
    let root = tree("\\text{abc}+x");
    let id = text_node(&root);
    let index = TreeIndex::new(&root).unwrap();
    for (row_offset, text_offset) in [(0, 0), (1, 3)] {
        for (anchor, focus) in [(Position::row(root.id, row_offset), Position::text(id, text_offset)),
            (Position::text(id, text_offset), Position::row(root.id, row_offset))] {
            let normalized = index.normalize_range(&Selection::Range { anchor, focus }).unwrap();
            match normalized {
                NormalizedRange::GridCells{..}=>panic!("A text boundary cannot become a cell rectangle"),
                NormalizedRange::Row { start, end, .. } | NormalizedRange::Text { start, end, .. } => {
                    assert_eq!(start, end, "equivalent caret edges became a destructive range");
                }
            }
        }
    }
}

#[test]
fn selection_ending_at_text_start_excludes_the_unselected_text_run() {
    let root = tree("a+\\text{bc}+d");
    let id = text_node(&root);
    let index = TreeIndex::new(&root).unwrap();
    for (anchor, focus) in [(Position::row(root.id, 0), Position::text(id, 0)),
        (Position::text(id, 0), Position::row(root.id, 0))] {
        assert!(matches!(index.normalize_range(&Selection::Range { anchor, focus }).unwrap(),
            NormalizedRange::Row { node_id, start: 0, end: 2, .. } if node_id == root.id));
    }
}

#[test]
fn selection_starting_at_text_end_excludes_the_unselected_text_run() {
    let root = tree("a+\\text{bc}+d");
    let id = text_node(&root);
    let index = TreeIndex::new(&root).unwrap();
    for (anchor, focus) in [(Position::row(root.id, 5), Position::text(id, 2)),
        (Position::text(id, 2), Position::row(root.id, 5))] {
        assert!(matches!(index.normalize_range(&Selection::Range { anchor, focus }).unwrap(),
            NormalizedRange::Row { node_id, start: 3, end: 5, .. } if node_id == root.id));
    }
}

#[test]
fn a_real_cross_fraction_range_explicitly_promotes_to_the_whole_structure() {
    let root = tree("a+\\frac{x+y}{z}+b");
    let NodeKind::Row { entries, .. } = &root.kind else { panic!() };
    let NodeKind::Fraction { numerator, denominator, .. } = &entries[2].node.kind else { panic!() };
    let index = TreeIndex::new(&root).unwrap();
    let selection = Selection::Range { anchor: Position::row(numerator.id, 1), focus: Position::row(denominator.id, 1) };
    assert_eq!(index.normalize_range(&selection).unwrap(), NormalizedRange::Row {
        node_id: root.id, start: 2, end: 3, promoted: true });
}

#[test]
fn node_selection_promotion_cannot_be_reported_as_an_exact_child_selection() {
    let root = tree("a+\\frac{x}{y}+b");
    let NodeKind::Row { entries, .. } = &root.kind else { panic!() };
    let fraction = &entries[2].node;
    let NodeKind::Fraction { numerator, .. } = &fraction.kind else { panic!() };
    let index = TreeIndex::new(&root).unwrap();
    assert_eq!(index.normalize_range(&Selection::Node { node_id: fraction.id }).unwrap(), NormalizedRange::Row {
        node_id: root.id, start: 2, end: 3, promoted: false });
    match index.normalize_range(&Selection::Node { node_id: numerator.id }) {
        Ok(range) => assert_eq!(range, NormalizedRange::Row { node_id: root.id, start: 2, end: 3, promoted: true }),
        Err(error) => assert_eq!(error.code, ErrorCode::InvalidPosition),
    }
}

#[test]
fn text_positions_count_graphemes_not_utf16_units_or_scalars() {
    let parts = ["e\u{301}", "👩🏽‍🚀", "🇨🇳", "中"];
    let text = parts.concat();
    let root = tree(&format!("\\text{{{text}}}"));
    let id = text_node(&root);
    let index = TreeIndex::new(&root).unwrap();
    let mut byte = 0;
    for (i, part) in parts.iter().enumerate() {
        assert_eq!(grapheme_byte(&text, i).unwrap(), byte);
        index.validate_position(Position::text(id, i)).unwrap();
        byte += part.len();
    }
    assert_eq!(grapheme_byte(&text, 4).unwrap(), byte);
    assert!(grapheme_byte(&text, 5).is_err());
    assert!(index.validate_position(Position::text(id, 5)).is_err());
    assert!(index.validate_position(Position::row(id, 0)).is_err());
}

#[test]
fn affinity_does_not_turn_a_collapsed_caret_into_a_range() {
    let a = Position::Row { node_id: NodeId(1), offset: 0, affinity: Affinity::Before };
    let b = Position::Row { node_id: NodeId(1), offset: 0, affinity: Affinity::After };
    assert!(a.same_location(b));
    assert!(Selection::Range { anchor: a, focus: b }.is_collapsed());
}

#[test]
#[cfg(target_pointer_width = "64")]
fn an_overflowing_constructor_offset_must_not_alias_zero() {
    let root = tree("x");
    let index = TreeIndex::new(&root).unwrap();
    let wrapped = u32::MAX as usize + 1;
    assert!(index.validate_position(Position::row(root.id, wrapped)).is_err());
}

#[test]
fn indexing_rejects_overbudget_tree_depth_before_navigation() {
    let mut root = Node::empty_row(NodeId(100), Grouping::Bare);
    for id in (1..80).rev() {
        root = Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::Row { grouping: Grouping::Braced,
            entries: vec![RowEntry { leading: String::new(), node: root }], trailing: String::new() } });
    }
    assert!(matches!(TreeIndex::new(&root), Err(error) if error.code == ErrorCode::ResourceLimit));
}
