//! Independent tests for the user's one-frame Shift+Enter mathematical command.
//! Native state tests are not substitutes for the separate two-frame browser test.
use std::collections::BTreeSet;
use serde_json::Value;
use visualtex_engine_core::{
    editing::{EditorAction, EditorActionResult},
    model::{Node, NodeKind},
    positions::{Affinity, Horizontal, Position, Selection, TreeIndex},
    session::{EngineSession, NodeId},
    source_edit::SourceEdit,
    transactions::{EditCommand, InputMode},
};

struct Owned { session: EngineSession, sequence: u32 }
impl Owned {
    fn new(source: &str) -> Self {
        Self { session: EngineSession::from_source("contained-line-review".into(), source.into()).unwrap(), sequence: 0 }
    }
    fn action(&mut self, action: EditorAction) -> EditorActionResult {
        self.sequence += 1;
        let before = self.session.document_snapshot();
        let result = self.session.apply_action(self.session.revision(), self.sequence, action).unwrap();
        let index = TreeIndex::new(&result.editor.document.analysis.tree).unwrap();
        index.validate_selection(&result.editor.selection).unwrap();
        if let Some(map) = &result.position_map {
            for position in TreeIndex::new(&before.analysis.tree).unwrap().carets {
                index.validate_position(map.apply_position(position, &index).unwrap()).unwrap();
            }
        }
        result
    }
    fn edit(&mut self, command: EditCommand) -> EditorActionResult {
        self.action(EditorAction::Edit { command, group_id: None })
    }
    fn select(&mut self, position: Position) {
        self.action(EditorAction::SetSelection { selection: Selection::caret(position) });
    }
    fn type_text(&mut self, text: &str) {
        self.edit(EditCommand::InsertText { text: text.into(), mode: InputMode::Math });
    }
    fn snapshot(&self) -> Value { serde_json::to_value(self.session.editor_snapshot().unwrap()).unwrap() }
}
fn grid(root: &Node) -> &Node {
    if matches!(root.kind, NodeKind::Grid { .. }) { return root; }
    for child in root.children() {
        let mut found = false; child.visit(&mut |n| found |= matches!(n.kind, NodeKind::Grid { .. }));
        if found { return grid(child); }
    }
    panic!("Expected a contained Grid");
}
fn cell(root: &Node, row_index: usize) -> &Node {
    let NodeKind::Grid { rows, .. } = &grid(root).kind else { unreachable!() };
    let NodeKind::GridRow { cells, .. } = &rows[row_index].kind else { unreachable!() };
    &cells[0]
}
fn ids(root: &Node) -> BTreeSet<u32> {
    let mut all = BTreeSet::new();
    root.visit(&mut |n| {
        assert!(all.insert(n.id.0));
        if let NodeKind::Grid { columns, .. } = &n.kind { for column in columns { assert!(all.insert(column.0)); } }
    });
    all
}
fn contents(s: &Value) -> Value {
    serde_json::json!({"source":s["document"]["source"],"tree":s["document"]["analysis"]["tree"],
        "selection":s["selection"],"documentId":s["document"]["documentId"]})
}

#[test]
fn promotion_is_one_history_entry_and_undo_redo_restore_exact_internal_identity() {
    let mut t = Owned::new("x+y=1"); let root = t.session.document_snapshot().analysis.tree.id;
    t.select(Position::row(root, 5)); let before = t.snapshot();
    let promoted = t.edit(EditCommand::SplitLine);
    assert_eq!(promoted.editor.undo_depth as u64, before["undoDepth"].as_u64().unwrap() + 1);
    assert_eq!(promoted.editor.document.document_id, "contained-line-review");
    assert_eq!(promoted.editor.document.analysis.tree.id, root);
    assert_eq!(promoted.editor.document.analysis.facts.grids, 1);
    assert_eq!(promoted.editor.document.analysis.facts.grid_rows, 2);
    let after = t.snapshot();
    t.action(EditorAction::Undo); assert_eq!(contents(&t.snapshot()), contents(&before));
    t.action(EditorAction::Redo); assert_eq!(contents(&t.snapshot()), contents(&after));
}

#[test]
fn every_root_boundary_splits_nodes_without_changing_their_identity() {
    for offset in 0..=3 {
        let mut t = Owned::new("abc"); let old = t.session.document_snapshot();
        let NodeKind::Row { entries, .. } = &old.analysis.tree.kind else { unreachable!() };
        let expected: Vec<_> = entries.iter().map(|e| e.node.id).collect();
        t.select(Position::row(old.analysis.tree.id, offset));
        let result = t.edit(EditCommand::SplitLine);
        let tree = &result.editor.document.analysis.tree;
        let NodeKind::Row { entries: left, .. } = &cell(tree, 0).kind else { unreachable!() };
        let NodeKind::Row { entries: right, .. } = &cell(tree, 1).kind else { unreachable!() };
        assert_eq!(left.iter().map(|e|e.node.id).collect::<Vec<_>>(), expected[..offset]);
        assert_eq!(right.iter().map(|e|e.node.id).collect::<Vec<_>>(), expected[offset..]);
        assert_eq!(result.editor.selection, Selection::caret(Position::row(cell(tree, 1).id, 0)));
    }
}

#[test]
fn repeated_breaks_extend_one_environment_instead_of_nesting_new_environments() {
    let mut t = Owned::new("x"); t.select(Position::row(NodeId(1), 1));
    for count in 2..=5 {
        t.edit(EditCommand::SplitLine); t.type_text("y");
        let state = t.session.document_snapshot();
        assert_eq!(state.analysis.facts.grids, 1);
        assert_eq!(state.analysis.facts.grid_rows, count);
        assert_eq!(t.session.source().matches("\\begin{gathered}").count(), 1);
    }
}

#[test]
fn splitting_an_existing_row_updates_both_affinities_without_invalid_anchors() {
    let mut t = Owned::new(r"\begin{gathered}abc\\z\end{gathered}");
    let old = t.session.document_snapshot(); let row = cell(&old.analysis.tree, 0).id;
    t.select(Position::row(row, 1));
    let result = t.edit(EditCommand::SplitLine);
    let tree = &result.editor.document.analysis.tree; let index = TreeIndex::new(tree).unwrap();
    let map = result.position_map.as_ref().unwrap();
    let before = Position::Row { node_id: row, offset: 1, affinity: Affinity::Before };
    let after = Position::Row { node_id: row, offset: 1, affinity: Affinity::After };
    let left = map.apply_position(before, &index).unwrap(); let right = map.apply_position(after, &index).unwrap();
    assert!(left.same_location(Position::row(cell(tree, 0).id, 1)));
    assert!(right.same_location(Position::row(cell(tree, 1).id, 0)));
    assert_eq!(result.editor.document.analysis.facts.grids, 1);
    assert_eq!(result.editor.document.analysis.facts.grid_rows, 3);
}

#[test]
fn unsupported_fraction_and_text_interior_breaks_do_not_mutate_any_state() {
    for source in [r"\frac{x}{y}", r"\text{abc}"] {
        let mut t = Owned::new(source); let snapshot = t.session.document_snapshot();
        let mut position = None;
        snapshot.analysis.tree.visit(&mut |node| {
            if let NodeKind::Fraction { numerator, .. } = &node.kind { position = Some(Position::row(numerator.id, 1)); }
            if matches!(node.kind, NodeKind::Text { .. }) { position = Some(Position::text(node.id, 1)); }
        });
        t.select(position.unwrap()); let before = t.snapshot(); t.sequence += 1;
        let failure = t.session.apply_action(t.session.revision(), t.sequence,
            EditorAction::Edit { command: EditCommand::SplitLine, group_id: None });
        assert!(failure.is_err(), "Interior structure unexpectedly cut: {source}");
        assert_eq!(t.snapshot(), before);
    }
}

#[test]
fn noncollapsed_root_selection_is_not_deleted_to_create_a_line_break() {
    let mut t = Owned::new("abc");
    t.action(EditorAction::SetSelection { selection: Selection::Range {
        anchor: Position::row(NodeId(1), 0), focus: Position::row(NodeId(1), 2),
    } });
    let before = t.snapshot(); t.sequence += 1;
    assert!(t.session.apply_action(t.session.revision(), t.sequence,
        EditorAction::Edit { command: EditCommand::SplitLine, group_id: None }).is_err());
    assert_eq!(t.snapshot(), before);
}

#[test]
fn pending_composition_blocks_break_without_committing_or_losing_preedit() {
    let mut t = Owned::new("x"); t.select(Position::row(NodeId(1), 1));
    let started = t.action(EditorAction::BeginComposition { mode: InputMode::Text });
    let id = started.editor.composition.unwrap().id;
    t.action(EditorAction::UpdateComposition { id, preedit: "sudu".into() });
    let before = t.snapshot(); t.sequence += 1;
    assert!(t.session.apply_action(t.session.revision(), t.sequence,
        EditorAction::Edit { command: EditCommand::SplitLine, group_id: None }).is_err());
    assert_eq!(t.snapshot(), before);
}

#[test]
fn duplicate_entire_multiline_allocates_disjoint_nodes_and_column_ids() {
    let mut t = Owned::new(r"\begin{gathered}a\\b\end{gathered}");
    let before = t.session.document_snapshot(); let original = ids(grid(&before.analysis.tree));
    t.action(EditorAction::SetSelection { selection: Selection::Node { node_id: grid(&before.analysis.tree).id } });
    let result = t.edit(EditCommand::DuplicateSelection);
    let NodeKind::Row { entries, .. } = &result.editor.document.analysis.tree.kind else { unreachable!() };
    assert_eq!(entries.len(), 2);
    assert_eq!(ids(&entries[0].node), original);
    assert!(original.is_disjoint(&ids(&entries[1].node)), "Copied multiline shares cell or column identity");
}

#[test]
fn deletion_at_row_boundary_merges_only_internal_rows_and_preserves_source_order() {
    let mut t = Owned::new(r"\begin{gathered}a\\b\end{gathered}");
    let before = t.session.document_snapshot();
    let a = ids(cell(&before.analysis.tree, 0)); let b = ids(cell(&before.analysis.tree, 1));
    t.select(Position::row(cell(&before.analysis.tree, 1).id, 0));
    let snapshot = t.snapshot(); let merged = t.edit(EditCommand::DeleteBackward);
    assert_eq!(merged.editor.document.analysis.facts.grids, 1);
    assert_eq!(merged.editor.document.analysis.facts.grid_rows, 1);
    assert_eq!(t.session.source(), r"\begin{gathered}ab\end{gathered}");
    let remaining = ids(cell(&merged.editor.document.analysis.tree, 0));
    assert!(a.iter().filter(|id| **id != cell(&before.analysis.tree, 0).id.0).all(|id| remaining.contains(id)));
    assert!(b.iter().filter(|id| **id != cell(&before.analysis.tree, 1).id.0).all(|id| remaining.contains(id)));
    t.action(EditorAction::Undo); assert_eq!(contents(&t.snapshot()), contents(&snapshot));
}

#[test]
fn source_patch_in_second_cell_changes_content_and_preserves_first_cell_and_column() {
    let mut t = Owned::new(r"\begin{gathered}a\\b\end{gathered}");
    let before = t.session.document_snapshot(); let first = ids(cell(&before.analysis.tree, 0));
    let NodeKind::Grid { columns: before_columns, .. } = &grid(&before.analysis.tree).kind else { unreachable!() };
    let at = t.session.source().find("\\\\b").unwrap() + 2;
    let (after, _) = t.session.patch_source(t.session.revision(), SourceEdit {
        start_utf16: at as u32, end_utf16: (at+1) as u32, replacement: "z".into(),
    }).unwrap();
    assert_eq!(ids(cell(&after.analysis.tree, 0)), first);
    let NodeKind::Grid { columns, .. } = &grid(&after.analysis.tree).kind else { unreachable!() };
    assert_eq!(columns, before_columns);
    let NodeKind::Row { entries, .. } = &cell(&after.analysis.tree, 1).kind else { unreachable!() };
    assert!(matches!(&entries[0].node.kind, NodeKind::Symbol { value, .. } if value == "z"), "Source changed but cell still has old content");
}

#[test]
fn duplicate_event_sequence_cannot_insert_a_second_environment_row() {
    let mut t = Owned::new("x"); t.select(Position::row(NodeId(1), 1));
    let revision = t.session.revision(); t.sequence += 1;
    let action = EditorAction::Edit { command: EditCommand::SplitLine, group_id: None };
    let first = t.session.apply_action(revision, t.sequence, action.clone()).unwrap();
    let duplicate = t.session.apply_action(revision, t.sequence, action).unwrap();
    assert!(duplicate.replayed);
    assert_eq!(duplicate.editor.document.analysis.facts.grid_rows, 2);
    assert_eq!(duplicate.editor.document.revision, first.editor.document.revision);
}

#[test]
fn outermost_row_merge_boundary_never_removes_the_formula_frame() {
    let mut t = Owned::new(r"\begin{gathered}a\\b\end{gathered}");
    let snapshot = t.session.document_snapshot(); t.select(Position::row(cell(&snapshot.analysis.tree, 0).id, 0));
    let before = t.snapshot(); let result = t.edit(EditCommand::MergeLine { direction: Horizontal::Left });
    assert!(!result.content_changed); assert_eq!(t.snapshot(), before);
}
