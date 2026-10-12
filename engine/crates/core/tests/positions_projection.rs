use std::sync::Arc;
use visualtex_engine_core::model::{self, AtomClass, IdAllocator, LatexMode, Node, NodeKind, RowEntry};
use visualtex_engine_core::parser;
use visualtex_engine_core::positions::{Affinity, Horizontal, NormalizedRange, Position, Selection, TreeIndex};
use visualtex_engine_core::projection;
use visualtex_engine_core::session::NodeId;
use visualtex_engine_core::source::SourceText;

fn tree(text: &str) -> Arc<Node> {
    parser::parse(SourceText::new(1, text).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap().tree
}
fn child(root: &Node, index: usize) -> &Node {
    let NodeKind::Row { entries, .. } = &root.kind else { panic!() };
    &entries[index].node
}
fn math_symbol(id: u32, value: &str) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None,
        kind: NodeKind::Symbol { value: value.into(), class: AtomClass::Ordinary, command: None } })
}

#[test]
fn horizontal_navigation_enters_fraction_arguments_in_structural_order() {
    let root = tree(r"\frac{x}{y}");
    let index = TreeIndex::new(&root).unwrap();
    let NodeKind::Fraction { numerator, denominator, .. } = &child(&root, 0).kind else { panic!() };
    let mut selection = Selection::caret(Position::row(root.id, 0));
    for expected in [Position::row(numerator.id, 0), Position::row(numerator.id, 1),
        Position::row(denominator.id, 0), Position::row(denominator.id, 1), Position::row(root.id, 1)] {
        selection = index.move_horizontal(&selection, Horizontal::Right, false).unwrap();
        assert_eq!(selection, Selection::caret(expected));
    }
    assert_eq!(index.move_horizontal(&selection, Horizontal::Right, false).unwrap(), selection);
    assert_eq!(index.move_horizontal(&selection, Horizontal::Left, false).unwrap(), Selection::caret(Position::row(denominator.id, 1)));
}

#[test]
fn scripts_have_a_real_editable_base_row() {
    let root = tree(r"x_1^2");
    let index = TreeIndex::new(&root).unwrap();
    let NodeKind::Scripts { base, slots } = &child(&root, 0).kind else { panic!() };
    assert!(base.is_row());
    let mut selection = Selection::caret(Position::row(root.id, 0));
    for expected in [Position::row(base.id, 0), Position::row(base.id, 1), Position::row(slots[0].argument.id, 0)] {
        selection = index.move_horizontal(&selection, Horizontal::Right, false).unwrap();
        assert_eq!(selection, Selection::caret(expected));
    }
    assert_eq!(model::latex(&root, LatexMode::Lossless), r"x_1^2");
    assert_eq!(model::latex(&root, LatexMode::Canonical), r"x_{1}^{2}");
}

#[test]
fn text_movement_is_by_grapheme_and_skips_equivalent_outer_boundaries() {
    let root = tree("\\text{A😀e\u{301}}");
    let index = TreeIndex::new(&root).unwrap();
    let text = child(&root, 0);
    let mut selection = Selection::caret(Position::row(root.id, 0));
    for grapheme in [1, 2, 3] {
        selection = index.move_horizontal(&selection, Horizontal::Right, false).unwrap();
        assert_eq!(selection, Selection::caret(Position::text(text.id, grapheme)));
    }
    assert!(index.validate_position(Position::text(text.id, 4)).is_err());
    assert!(index.validate_position(Position::row(text.id, 0)).is_err());
    assert!(index.validate_position(Position::text(NodeId(99999), 0)).is_err());
}

#[test]
fn reversed_cross_structure_range_promotes_to_whole_fraction_without_cutting_it() {
    let root = tree(r"a+\frac{x}{y}+z");
    let index = TreeIndex::new(&root).unwrap();
    let fraction = child(&root, 2);
    let NodeKind::Fraction { numerator, denominator, .. } = &fraction.kind else { panic!() };
    let selection = Selection::Range { anchor: Position::row(denominator.id, 1), focus: Position::row(numerator.id, 0) };
    assert_eq!(index.normalize_range(&selection).unwrap(), NormalizedRange::Row { node_id: root.id, start: 2, end: 3, promoted: true });
    let forward = Selection::Range { anchor: Position::row(numerator.id, 0), focus: Position::row(denominator.id, 1) };
    assert_eq!(index.normalize_range(&selection).unwrap(), index.normalize_range(&forward).unwrap());
    // Collapsing lands at the edges of the highlighted (whole-fraction) span.
    assert_eq!(index.move_horizontal(&selection, Horizontal::Left, false).unwrap(), Selection::caret(Position::row(root.id, 2)));
    assert_eq!(index.move_horizontal(&selection, Horizontal::Right, false).unwrap(), Selection::caret(Position::row(root.id, 3)));
}

#[test]
fn extending_selection_preserves_the_original_anchor_and_affinity() {
    let root = tree("x+y");
    let index = TreeIndex::new(&root).unwrap();
    let anchor = Position::Row { node_id: root.id, offset: 2, affinity: Affinity::Before };
    let selection = index.move_horizontal(&Selection::caret(anchor), Horizontal::Left, true).unwrap();
    let Selection::Range { anchor: preserved, focus } = selection else { panic!() };
    assert_eq!(anchor, preserved); assert_eq!(focus.offset(), 1);
}

#[test]
fn projected_source_has_complete_new_ranges_without_reparsing_tree() {
    let mut root = tree(" \\frac  {x} %den\n { y } + \\alpha z ");
    let root_mut = Arc::make_mut(&mut root); root_mut.origin = None;
    let NodeKind::Row { entries, .. } = &mut root_mut.kind else { panic!() };
    let fraction = Arc::make_mut(&mut entries[0].node); fraction.origin = None;
    let NodeKind::Fraction { numerator, .. } = &mut fraction.kind else { panic!() };
    let numerator = Arc::make_mut(numerator); numerator.origin = None;
    let NodeKind::Row { entries, .. } = &mut numerator.kind else { panic!() };
    entries.push(RowEntry { leading: String::new(), node: math_symbol(100, "w") });
    let projected = projection::project_tree(root, 2).unwrap();
    assert_eq!(projected.source.text(), " \\frac  {xw} %den\n { y } + \\alpha z ");
    projected.tree.visit(&mut |node| {
        let origin = node.origin.as_ref().unwrap();
        assert_eq!(origin.span.source_id, 2);
        assert_eq!(origin.text(), &projected.source.text()[origin.span.byte_start as usize..origin.span.byte_end as usize]);
    });
    assert_eq!(model::latex(&projected.tree, LatexMode::Lossless), projected.source.text());
}

#[test]
fn source_projection_maps_inserted_command_separator_to_correct_child_ranges() {
    let mut root = tree(r"\alpha");
    let r = Arc::make_mut(&mut root); r.origin = None;
    let NodeKind::Row { entries, .. } = &mut r.kind else { panic!() };
    entries.push(RowEntry { leading: String::new(), node: math_symbol(100, "x") });
    let projected = projection::project_tree(root, 2).unwrap();
    assert_eq!(projected.source.text(), r"\alpha x");
    let x = projected.tree.find(NodeId(100)).unwrap();
    assert_eq!(x.origin.as_ref().unwrap().text(), "x");
    assert_eq!(x.origin.as_ref().unwrap().span.byte_start, 7);
}

#[test]
fn editing_script_base_to_multiple_atoms_adds_grouping_and_reifies_it() {
    let mut root = tree("x^2");
    let r = Arc::make_mut(&mut root); r.origin = None;
    let NodeKind::Row { entries, .. } = &mut r.kind else { panic!() };
    let scripts = Arc::make_mut(&mut entries[0].node); scripts.origin = None;
    let NodeKind::Scripts { base, .. } = &mut scripts.kind else { panic!() };
    let base = Arc::make_mut(base); base.origin = None;
    let NodeKind::Row { entries, .. } = &mut base.kind else { panic!() };
    entries.push(RowEntry { leading: String::new(), node: math_symbol(100, "y") });
    let projected = projection::project_tree(root, 2).unwrap();
    assert_eq!(projected.source.text(), "{xy}^2");
    assert_eq!(model::latex(&projected.tree, LatexMode::Canonical), "{xy}^{2}");
    let reparse = parser::parse(Arc::clone(&projected.source), NodeId(1), &mut IdAllocator::new(2)).unwrap();
    assert_eq!(projected.facts, reparse.facts);
    assert_eq!(model::latex(&reparse.tree, LatexMode::Canonical), model::latex(&projected.tree, LatexMode::Canonical));
}
