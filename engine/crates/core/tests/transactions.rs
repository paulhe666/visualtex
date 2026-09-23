use std::collections::BTreeSet;
use std::sync::Arc;
use visualtex_engine_core::{ErrorCode, model, parser};
use visualtex_engine_core::model::{IdAllocator, LatexMode, Node, NodeKind};
use visualtex_engine_core::parser::ParseResult;
use visualtex_engine_core::positions::{Affinity, Horizontal, Position, Selection, TreeIndex};
use visualtex_engine_core::session::NodeId;
use visualtex_engine_core::source::SourceText;
use visualtex_engine_core::transactions::{self, EditCommand, InputMode, PreparedEdit, StructureKind};

struct State { doc: Arc<ParseResult>, selection: Selection, ids: IdAllocator, next_source: u32 }
impl State {
    fn new(source: &str) -> Self {
        let mut ids = IdAllocator::new(2);
        let doc = Arc::new(parser::parse(SourceText::new(1, source).unwrap(), NodeId(1), &mut ids).unwrap());
        Self { doc, selection: Selection::caret(Position::row(NodeId(1), 0)), ids, next_source: 2 }
    }
    fn apply(&mut self, command: EditCommand) -> PreparedEdit {
        let before = self.doc.clone();
        let result = transactions::prepare(&self.doc, &self.selection, &command, self.next_source, &mut self.ids).unwrap();
        self.verify(&before, &result);
        if result.changed { self.next_source += 1; }
        self.doc = result.document.clone(); self.selection = result.selection.clone(); result
    }
    fn verify(&self, old: &ParseResult, result: &PreparedEdit) {
        let index = TreeIndex::new(&result.document.tree).unwrap();
        index.validate_selection(&result.selection).unwrap();
        let old_index = TreeIndex::new(&old.tree).unwrap();
        for position in &old_index.carets {
            result.position_map.apply_position(*position, &index).unwrap();
            let before_affinity = match *position {
                Position::Row { node_id, offset, .. } => Position::Row { node_id, offset, affinity: Affinity::Before },
                Position::Text { node_id, grapheme, .. } => Position::Text { node_id, grapheme, affinity: Affinity::Before },
            };
            result.position_map.apply_position(before_affinity, &index).unwrap();
        }
        assert_eq!(model::latex(&result.document.tree, LatexMode::Lossless), result.document.source.text());
        let fresh = parser::parse(result.document.source.clone(), NodeId(1), &mut IdAllocator::new(2)).unwrap();
        assert!(fresh.complete && fresh.supported, "{}", result.document.source.text());
        assert_eq!(fresh.facts, result.document.facts, "{}", result.document.source.text());
        assert_eq!(model::latex(&fresh.tree, LatexMode::Canonical), model::latex(&result.document.tree, LatexMode::Canonical));
        result.document.tree.visit(&mut |node| {
            let origin = node.origin.as_ref().unwrap();
            assert_eq!(origin.span.source_id, result.document.source.id());
            assert_eq!(origin.text(), &result.document.source.text()[origin.span.byte_start as usize..origin.span.byte_end as usize]);
        });
    }
    fn input(&mut self, text: &str) { self.apply(EditCommand::InsertText { text: text.into(), mode: InputMode::Math }); }
    fn text(&mut self, text: &str) { self.apply(EditCommand::InsertText { text: text.into(), mode: InputMode::Text }); }
    fn source(&self) -> &str { self.doc.source.text() }
}
fn entry(root: &Node, i: usize) -> &Node {
    let NodeKind::Row { entries, .. } = &root.kind else { panic!() }; &entries[i].node
}
fn ids(node: &Node) -> BTreeSet<NodeId> {
    let mut ids = BTreeSet::new(); node.visit(&mut |node| { assert!(ids.insert(node.id)); }); ids
}

#[test]
fn actual_fraction_construction_edits_slots_without_parsing_typing() {
    let mut s = State::new("");
    s.apply(EditCommand::InsertStructure { structure: StructureKind::Fraction });
    assert_eq!(s.source(), r"\frac{}{}");
    s.input("x+1"); assert_eq!(s.source(), r"\frac{x+1}{}");
    let NodeKind::Fraction { denominator, .. } = &entry(&s.doc.tree,0).kind else { panic!() };
    s.selection = Selection::caret(Position::row(denominator.id,0));
    s.input("y"); assert_eq!(s.source(), r"\frac{x+1}{y}"); assert!(s.doc.exportable);
}

#[test]
fn superscript_and_subscript_reuse_the_existing_structure_and_editable_base() {
    let mut s = State::new(""); s.input("x");
    let x_id = entry(&s.doc.tree, 0).id;
    s.apply(EditCommand::InsertStructure { structure: StructureKind::Superscript }); s.input("2");
    assert_eq!(s.source(), "x^{2}");
    s.selection = Selection::caret(Position::row(NodeId(1),1));
    let scripts_id = entry(&s.doc.tree,0).id;
    s.apply(EditCommand::InsertStructure { structure: StructureKind::Subscript }); s.input("i");
    assert_eq!(s.source(), "x^{2}_{i}"); assert_eq!(entry(&s.doc.tree,0).id, scripts_id);
    assert!(s.doc.tree.find(x_id).is_some());
    s.selection = Selection::caret(Position::row(NodeId(1),1));
    let old = s.doc.clone();
    let result = s.apply(EditCommand::InsertStructure { structure: StructureKind::Superscript });
    assert!(!result.changed); assert!(Arc::ptr_eq(&old,&s.doc));
    let NodeKind::Scripts { base, .. } = &entry(&s.doc.tree,0).kind else { panic!() };
    s.selection = Selection::caret(Position::row(base.id,1)); s.input("y");
    assert_eq!(s.source(), "{xy}^{2}_{i}");
}

#[test]
fn selected_script_is_wrapped_as_a_group_instead_of_creating_duplicate_script_syntax() {
    let mut s = State::new("x^2");
    s.selection = Selection::Node { node_id: entry(&s.doc.tree,0).id };
    s.apply(EditCommand::InsertStructure { structure: StructureKind::Superscript }); s.input("3");
    assert_eq!(s.source(), "{x^2}^{3}");
}

#[test]
fn wrapping_a_range_preserves_selected_node_ids_and_unselected_trivia() {
    let mut s = State::new("  x + y %tail\n +z");
    let a = entry(&s.doc.tree,0).id; let b = entry(&s.doc.tree,2).id;
    s.selection = Selection::Range { anchor: Position::row(NodeId(1),3), focus: Position::row(NodeId(1),0) };
    s.apply(EditCommand::InsertStructure { structure: StructureKind::Fraction }); s.input("2");
    assert_eq!(s.source(), "\\frac{  x + y}{2} %tail\n +z");
    assert!(s.doc.tree.find(a).is_some() && s.doc.tree.find(b).is_some());
}

#[test]
fn one_delete_selects_structure_and_the_next_delete_removes_it() {
    let mut s = State::new(r"a+\frac{x}{y}+z");
    s.selection = Selection::caret(Position::row(NodeId(1),3));
    let original = s.source().to_owned();
    let result = s.apply(EditCommand::DeleteBackward);
    assert!(!result.changed); assert_eq!(s.source(),original); assert!(matches!(s.selection,Selection::Node{..}));
    let result = s.apply(EditCommand::DeleteBackward);
    assert!(result.changed); assert_eq!(s.source(),"a++z"); assert_eq!(s.selection,Selection::caret(Position::row(NodeId(1),2)));
}

#[test]
fn empty_argument_backspace_navigates_out_without_flattening_the_denominator() {
    let mut s = State::new(r"\frac{}{y}");
    let NodeKind::Fraction { numerator, .. } = &entry(&s.doc.tree,0).kind else { panic!() };
    s.selection = Selection::caret(Position::row(numerator.id,0));
    let result = s.apply(EditCommand::DeleteBackward);
    assert!(!result.changed); assert_eq!(s.source(),r"\frac{}{y}");
    assert_eq!(s.selection,Selection::caret(Position::row(NodeId(1),0)));
}

#[test]
fn text_edits_and_deletion_obey_extended_graphemes() {
    let mut s = State::new(""); s.text("A😀e\u{301}");
    let id = entry(&s.doc.tree,0).id;
    s.apply(EditCommand::DeleteBackward); assert_eq!(s.source(),"\\text{A😀}");
    s.apply(EditCommand::DeleteBackward); assert_eq!(s.source(),r"\text{A}");
    s.text("\u{301}"); assert_eq!(s.source(),"\\text{A\u{301}}");
    assert_eq!(s.selection,Selection::caret(Position::text(id,1)));
    let map = s.apply(EditCommand::InsertText { text:"x".into(),mode:InputMode::Text }).position_map;
    let index = TreeIndex::new(&s.doc.tree).unwrap();
    assert_eq!(map.apply_position(Position::text(id,0),&index).unwrap(),Position::text(id,0));
    s.selection = Selection::Range { anchor:Position::row(NodeId(1),0),focus:Position::text(id,1) };
    s.apply(EditCommand::DeleteBackward); assert_eq!(s.source(),r"\text{x}");
}

#[test]
fn splitting_and_merging_text_maps_both_old_runs_and_preserves_unicode() {
    let mut s = State::new("\\text{A😀e\u{301}}");
    let id = entry(&s.doc.tree,0).id;
    s.selection = Selection::caret(Position::text(id,1));
    s.apply(EditCommand::SplitText); assert_eq!(s.source(),"\\text{A}\\text{😀e\u{301}}");
    assert_eq!(entry(&s.doc.tree,0).id,id);
    s.apply(EditCommand::MergeText { direction:Horizontal::Left });
    assert_eq!(s.source(),"\\text{A😀e\u{301}}"); assert_eq!(entry(&s.doc.tree,0).id,id);
    let mut s = State::new("\\text{A}\\text{\u{301}B}");
    s.selection = Selection::caret(Position::row(NodeId(1),1));
    s.apply(EditCommand::MergeText { direction:Horizontal::Left });
    assert_eq!(s.source(),"\\text{A\u{301}B}");
}

#[test]
fn a_structure_can_replace_only_selected_text_without_losing_its_neighbors() {
    let mut s = State::new(r"\text{abcde}");
    let old_id = entry(&s.doc.tree,0).id;
    s.selection = Selection::Range { anchor:Position::text(old_id,1),focus:Position::text(old_id,4) };
    s.apply(EditCommand::InsertStructure { structure:StructureKind::Fraction });
    s.input("2");
    assert_eq!(s.source(),r"\text{a}\frac{\text{bcd}}{2}\text{e}");
    assert_eq!(entry(&s.doc.tree,0).id,old_id);
}

#[test]
fn copying_a_formula_allocates_completely_disjoint_subtree_ids() {
    let mut s = State::new(r"\frac{x_1}{y}");
    let original = ids(entry(&s.doc.tree,0));
    s.selection = Selection::Node { node_id:entry(&s.doc.tree,0).id };
    s.apply(EditCommand::DuplicateSelection);
    let copied = ids(entry(&s.doc.tree,1));
    assert!(original.is_disjoint(&copied)); assert_eq!(ids(entry(&s.doc.tree,0)),original);
    assert_eq!(s.source(),r"\frac{x_1}{y}\frac{x_1}{y}");
    let NodeKind::Fraction { numerator, .. } = &entry(&s.doc.tree,1).kind else { panic!() };
    s.selection = Selection::caret(Position::row(numerator.id,1)); s.input("+z");
    assert_eq!(s.source(),r"\frac{x_1}{y}\frac{x_1+z}{y}");
}

#[test]
fn explicit_latex_paste_is_one_structural_insertion_with_comments_intact() {
    let mut s = State::new("z");
    s.apply(EditCommand::PasteLatex { latex:"\\frac{x}{y} %end\n".into() });
    assert_eq!(s.source(),"\\frac{x}{y} %end\nz");
    assert_eq!(s.doc.facts.fractions,1);
}

#[test]
fn math_typing_combining_marks_does_not_disagree_with_the_source_graphemes() {
    let mut s = State::new(""); s.input("A"); s.input("\u{301}");
    assert_eq!(s.source(),"A\u{301}");
    assert_eq!(s.doc.facts.nodes,2);
    s.apply(EditCommand::DeleteBackward); assert_eq!(s.source(),"");
}

#[test]
fn failed_operations_leave_document_selection_and_allocator_unchanged() {
    let mut s = State::new("x");
    for command in [EditCommand::PasteLatex { latex:r"\frac{".into() },
        EditCommand::PasteLatex { latex:r"\unknown{x}".into() },
        EditCommand::InsertText { text:"\\frac".into(),mode:InputMode::Math },
        EditCommand::InsertText { text:"x".repeat(9000),mode:InputMode::Math },
        EditCommand::SetStyle { style:"bold".into() }, EditCommand::SplitText,
    ] {
        let before = serde_json::to_value(&s.doc).unwrap(); let next = s.ids.next_unused();
        let result = transactions::prepare(&s.doc,&s.selection,&command,s.next_source,&mut s.ids);
        assert!(result.is_err()); assert_eq!(serde_json::to_value(&s.doc).unwrap(),before); assert_eq!(s.ids.next_unused(),next);
    }
    let unknown = State::new(r"\unknown{x}");
    assert!(matches!(transactions::prepare(&unknown.doc,&unknown.selection,&EditCommand::DeleteForward,2,&mut s.ids),Err(e) if e.code == ErrorCode::UnsupportedSyntax));
}
