use std::collections::BTreeSet;
use serde_json::{Value, json};
use visualtex_engine_core::{EngineResult, ErrorCode};
use visualtex_engine_core::editing::{EditorAction, EditorActionResult, Vertical};
use visualtex_engine_core::history::{MAX_HISTORY_ENTRIES, MAX_HISTORY_WEIGHT};
use visualtex_engine_core::model::{Node, NodeKind};
use visualtex_engine_core::positions::{Affinity, Horizontal, Position, Selection, TreeIndex};
use visualtex_engine_core::session::{EngineSession, NodeId, Revision};
use visualtex_engine_core::source_edit::SourceEdit;
use visualtex_engine_core::transactions::{EditCommand, InputMode, StructureKind};

struct TestSession { session: EngineSession, sequence: u32 }
impl TestSession {
    fn new(source: &str) -> Self { Self { session: EngineSession::from_source("owned-history-tests".into(), source.into()).unwrap(), sequence: 0 } }
    fn try_action(&mut self, action: EditorAction) -> EngineResult<EditorActionResult> {
        self.sequence += 1;
        let before = self.session.document_snapshot();
        let result = self.session.apply_action(self.session.revision(), self.sequence, action)?;
        let index = TreeIndex::new(&result.editor.document.analysis.tree).unwrap();
        index.validate_selection(&result.editor.selection).unwrap();
        if let Some(map) = &result.position_map {
            for position in TreeIndex::new(&before.analysis.tree).unwrap().carets {
                map.apply_position(position, &index).unwrap();
                let position = match position {
                    Position::Row { node_id, offset, .. } => Position::Row { node_id, offset, affinity: Affinity::Before },
                    Position::Text { node_id, grapheme, .. } => Position::Text { node_id, grapheme, affinity: Affinity::Before },
                };
                map.apply_position(position, &index).unwrap();
            }
        }
        Ok(result)
    }
    fn action(&mut self, action: EditorAction) -> EditorActionResult { self.try_action(action).unwrap() }
    fn edit(&mut self, command: EditCommand) -> EditorActionResult { self.action(EditorAction::Edit { command, group_id: None }) }
    fn input(&mut self, text: &str, group: Option<&str>) -> EditorActionResult { self.action(EditorAction::Edit {
        command: EditCommand::InsertText { text: text.into(), mode: InputMode::Math }, group_id: group.map(str::to_owned) }) }
    fn select(&mut self, selection: Selection) { self.action(EditorAction::SetSelection { selection }); }
    fn undo(&mut self) -> EditorActionResult { self.action(EditorAction::Undo) }
    fn redo(&mut self) -> EditorActionResult { self.action(EditorAction::Redo) }
    fn frame(&self) -> Value {
        let s = self.session.editor_snapshot().unwrap();
        json!({"source":s.document.source,"sourceId":s.document.source_id,"analysis":s.document.analysis,
            "retainedTree":s.document.retained_tree,"canonicalLatex":s.document.canonical_latex,"selection":s.selection})
    }
    fn source(&self) -> &str { self.session.source() }
    fn editor_json(&self) -> Value { serde_json::to_value(self.session.editor_snapshot().unwrap()).unwrap() }
}
fn entry(tree: &Node, index: usize) -> &Node {
    let NodeKind::Row { entries, .. } = &tree.kind else { panic!() }; &entries[index].node
}
fn ids(node: &Node) -> BTreeSet<NodeId> {
    let mut ids = BTreeSet::new(); node.visit(&mut |n| { assert!(ids.insert(n.id)); }); ids
}

#[test]
fn undo_redo_restore_structure_identity_source_and_selection_not_revision_numbers() {
    let mut t = TestSession::new(r"\frac{x_1}{y}");
    assert_eq!(t.session.editor_snapshot().unwrap().undo_depth, 0);
    let original = t.frame();
    t.input("+z", None); let edited = t.frame();
    assert_eq!(t.session.revision(), Revision(1));
    let undone = t.undo(); assert!(undone.content_changed); assert_eq!(t.frame(), original); assert_eq!(t.session.revision(), Revision(2));
    t.redo(); assert_eq!(t.frame(), edited); assert_eq!(t.session.revision(), Revision(3));
    t.undo(); assert_eq!(t.frame(), original);
}

#[test]
fn selection_and_navigation_are_not_content_history() {
    let mut t = TestSession::new("xyz");
    let original = t.session.document_snapshot();
    for direction in [Horizontal::Left, Horizontal::Left, Horizontal::Right] {
        let result = t.action(EditorAction::MoveHorizontal { direction, extend: false });
        assert!(!result.content_changed); assert_eq!(result.editor.undo_depth, 0); assert_eq!(result.editor.document.revision, Revision(0));
    }
    assert_eq!(serde_json::to_value(t.session.document_snapshot()).unwrap(), serde_json::to_value(original).unwrap());
    assert!(!t.undo().content_changed);
}

#[test]
fn typing_groups_merge_only_adjacent_same_mode_inputs_and_movement_breaks_grouping() {
    let mut t = TestSession::new("");
    t.input("x",Some("typing-a")); t.input("y",Some("typing-a")); t.input("z",Some("typing-a"));
    assert_eq!(t.session.editor_snapshot().unwrap().undo_depth,1);
    let end = t.session.editor_snapshot().unwrap().selection;
    t.action(EditorAction::MoveHorizontal { direction:Horizontal::Left,extend:false });
    t.select(end); t.input("w",Some("typing-a"));
    assert_eq!(t.session.editor_snapshot().unwrap().undo_depth,2);
    t.undo(); assert_eq!(t.source(),"xyz"); t.undo(); assert_eq!(t.source(),"");
    t.redo(); t.redo(); assert_eq!(t.source(),"xyzw");
    t.edit(EditCommand::PasteLatex { latex:"+q".into() });
    t.input("r",Some("typing-a")); assert_eq!(t.session.editor_snapshot().unwrap().undo_depth,4);
}

#[test]
fn inline_shortcut_replacement_merges_into_one_math_typing_history_item() {
    let mut t=TestSession::new("");
    for ch in ["a","l","p","h","a"] { t.input(ch,Some("typing-shortcut")); }
    assert_eq!(t.source(),"alpha");assert_eq!(t.session.editor_snapshot().unwrap().undo_depth,1);
    t.action(EditorAction::Edit{command:EditCommand::ReplaceRecentMath{expected:"alpha".into(),latex:r"\alpha".into()},group_id:Some("typing-shortcut".into())});
    assert_eq!(t.source(),r"\alpha");assert_eq!(t.session.editor_snapshot().unwrap().undo_depth,1);
    t.undo();assert_eq!(t.source(),"");t.redo();assert_eq!(t.source(),r"\alpha");
}

#[test]
fn inline_shortcut_replacement_cannot_cross_structure_or_source_trivia() {
    for source in [r"x\frac{a}{b}","x % keep\n y"] {
        let mut t=TestSession::new(source);let before=t.editor_json();
        let result=t.try_action(EditorAction::Edit{command:EditCommand::ReplaceRecentMath{expected:"xy".into(),latex:r"\alpha".into()},group_id:Some("typing-shortcut".into())});
        assert!(matches!(result,Err(e) if matches!(e.code,ErrorCode::Conflict|ErrorCode::InvalidPosition)));
        assert_eq!(t.editor_json(),before);
    }
}

#[test]
fn source_and_visual_edits_share_history_including_invalid_drafts() {
    let mut t = TestSession::new(r"\frac{x}{y}");
    let first = t.frame();
    t.session.patch_source(Revision(0),SourceEdit { start_utf16:6,end_utf16:7,replacement:"z".into() }).unwrap();
    let second = t.frame();
    t.input("+q",None); let third = t.frame();
    t.session.replace_source(t.session.revision(),"\\frac{".into()).unwrap();
    let draft = t.frame(); assert!(!t.session.document_snapshot().analysis.complete);
    let unchanged = t.editor_json();
    assert!(matches!(t.try_action(EditorAction::Edit { command:EditCommand::DeleteBackward,group_id:None }),Err(e) if e.code == ErrorCode::IncompleteSource));
    assert_eq!(t.editor_json(),unchanged);
    t.undo(); assert_eq!(t.frame(),third); t.undo(); assert_eq!(t.frame(),second); t.undo(); assert_eq!(t.frame(),first);
    t.redo(); assert_eq!(t.frame(),second); t.redo(); assert_eq!(t.frame(),third); t.redo(); assert_eq!(t.frame(),draft);
}

#[test]
fn editing_after_undo_discards_redo_and_never_reuses_object_or_source_identities() {
    let mut t = TestSession::new("x");
    t.input("+y",None);
    let old = t.session.document_snapshot(); let old_ids = ids(&old.analysis.tree); let old_source_id = old.source_id;
    t.undo(); t.input("+z",None);
    let new = t.session.document_snapshot();
    assert!(new.source_id > old_source_id);
    let original = TestSession::new("x").session.document_snapshot();
    let original_ids = ids(&original.analysis.tree);
    let new_ids = ids(&new.analysis.tree);
    assert!(old_ids.difference(&original_ids).all(|id| !new_ids.contains(id)));
    assert_eq!(t.session.editor_snapshot().unwrap().redo_depth,0); assert!(!t.redo().content_changed);
}

#[test]
fn identical_duplicate_event_is_acknowledged_without_applying_it_twice() {
    let mut t = TestSession::new("");
    let action = EditorAction::Edit { command:EditCommand::InsertText { text:"x".into(),mode:InputMode::Math },group_id:None };
    t.action(action.clone());
    let before = t.editor_json();
    let replay = t.session.apply_action(Revision(0),1,action.clone()).unwrap();
    assert!(replay.replayed && !replay.content_changed); assert_eq!(t.editor_json(),before); assert_eq!(t.source(),"x");
    assert!(matches!(t.session.apply_action(Revision(1),1,action),Err(e) if e.code == ErrorCode::Conflict));
    assert!(matches!(t.session.apply_action(Revision(0),1,EditorAction::Undo),Err(e) if e.code == ErrorCode::Conflict));
    t.input("y",None);
    let replay = t.session.apply_action(Revision(0),1,EditorAction::Edit {
        command:EditCommand::InsertText { text:"x".into(),mode:InputMode::Math },group_id:None }).unwrap();
    assert!(replay.replayed); assert_eq!(t.source(),"xy");
}

#[test]
fn retired_event_sequences_cannot_be_reexecuted_after_receipt_eviction() {
    let mut t = TestSession::new("");
    let initial = EditorAction::MoveHorizontal { direction:Horizontal::Right,extend:false };
    t.action(initial.clone());
    for _ in 0..70 { t.action(initial.clone()); }
    assert_eq!(t.session.revision(),Revision(0));
    assert!(matches!(t.session.apply_action(Revision(0),1,initial),Err(e) if e.code == ErrorCode::Conflict));
    let before = t.editor_json();
    assert!(matches!(t.session.apply_action(Revision(9),99,EditorAction::Undo),Err(e) if e.code == ErrorCode::StaleRevision));
    assert_eq!(t.editor_json(),before);
}

#[test]
fn repeated_delete_event_does_not_turn_selection_into_an_unintended_deletion() {
    let mut t = TestSession::new(r"\frac{x}{y}");
    let action = EditorAction::Edit { command:EditCommand::DeleteBackward,group_id:None };
    let select = t.action(action.clone()); assert!(!select.content_changed);
    let selected = t.frame();
    assert!(matches!(select.editor.selection,Selection::Node { .. }));
    let replay = t.session.apply_action(Revision(0),1,action.clone()).unwrap(); assert!(replay.replayed);
    assert_eq!(t.frame(),selected);
    t.action(action); assert_eq!(t.source(),""); t.undo(); assert_eq!(t.frame(),selected);
}

#[test]
fn composition_preedit_never_mutates_source_and_commit_is_one_undoable_operation() {
    let mut t = TestSession::new("x");
    let original = t.frame();
    let begin = t.action(EditorAction::BeginComposition { mode:InputMode::Text });
    let id = begin.editor.composition.unwrap().id;
    for preedit in ["s","su","速","速度😀"] {
        let result = t.action(EditorAction::UpdateComposition { id,preedit:preedit.into() });
        assert_eq!(result.editor.document.revision,Revision(0)); assert_eq!(result.editor.undo_depth,0); assert_eq!(t.frame(),original);
    }
    let event = t.sequence+1;
    let committed = t.action(EditorAction::CommitComposition { id,text:None });
    assert_eq!(t.source(),"x\\text{速度😀}"); assert_eq!(committed.editor.undo_depth,1); assert!(committed.editor.composition.is_none());
    let result = t.session.apply_action(Revision(0),event,EditorAction::CommitComposition { id,text:None }).unwrap();
    assert!(result.replayed); assert_eq!(t.source(),"x\\text{速度😀}");
    assert!(matches!(t.try_action(EditorAction::CommitComposition { id,text:None }),Err(e) if e.code == ErrorCode::Conflict));
    t.undo(); assert_eq!(t.frame(),original); t.redo(); assert_eq!(t.source(),"x\\text{速度😀}");
}

#[test]
fn cancel_and_empty_composition_commit_preserve_selected_existing_formula() {
    let mut t = TestSession::new(r"\frac{x}{y}");
    let formula = entry(&t.session.document_snapshot().analysis.tree,0).id;
    t.select(Selection::Node { node_id:formula }); let original = t.frame();
    let id = t.action(EditorAction::BeginComposition { mode:InputMode::Text }).editor.composition.unwrap().id;
    t.action(EditorAction::UpdateComposition { id,preedit:"候选".into() });
    t.action(EditorAction::CancelComposition { id }); assert_eq!(t.frame(),original);
    let next = t.action(EditorAction::BeginComposition { mode:InputMode::Text }).editor.composition.unwrap().id;
    assert!(next > id);
    t.action(EditorAction::CommitComposition { id:next,text:Some(String::new()) });
    assert_eq!(t.frame(),original); assert_eq!(t.session.editor_snapshot().unwrap().undo_depth,0);
}

#[test]
fn composition_target_epoch_prevents_moving_away_and_back_then_committing_stale_text() {
    let mut t = TestSession::new("xy");
    let original = t.session.editor_snapshot().unwrap().selection;
    let id = t.action(EditorAction::BeginComposition { mode:InputMode::Text }).editor.composition.unwrap().id;
    t.action(EditorAction::UpdateComposition { id,preedit:"速度".into() });
    let before = t.editor_json();
    assert!(matches!(t.session.replace_source(Revision(0),"z".into()),Err(e) if e.code == ErrorCode::Conflict));
    assert_eq!(t.editor_json(),before);
    assert!(matches!(t.try_action(EditorAction::Edit { command:EditCommand::DeleteBackward,group_id:None }),Err(e) if e.code == ErrorCode::Conflict));
    t.action(EditorAction::MoveHorizontal { direction:Horizontal::Left,extend:false }); t.select(original);
    assert!(matches!(t.try_action(EditorAction::CommitComposition { id,text:None }),Err(e) if e.code == ErrorCode::Conflict));
    assert_eq!(t.source(),"xy");
    t.action(EditorAction::CancelComposition { id }); assert!(t.session.editor_snapshot().unwrap().composition.is_none());
}

#[test]
fn failed_composition_commit_retains_preedit_and_can_be_corrected() {
    let mut t = TestSession::new("x");
    let id = t.action(EditorAction::BeginComposition { mode:InputMode::Math }).editor.composition.unwrap().id;
    t.action(EditorAction::UpdateComposition { id,preedit:"\\frac".into() });
    let before = t.editor_json();
    assert!(matches!(t.try_action(EditorAction::CommitComposition { id,text:None }),Err(e) if e.code == ErrorCode::UnsupportedSyntax));
    assert_eq!(t.editor_json(),before);
    t.action(EditorAction::UpdateComposition { id,preedit:"y".into() });
    t.action(EditorAction::CommitComposition { id,text:None }); assert_eq!(t.source(),"xy");
}

#[test]
fn history_is_bounded_and_undo_redo_do_not_double_its_accounting_weight() {
    let mut t = TestSession::new("");
    for _ in 0..(MAX_HISTORY_ENTRIES+5) { t.input("x",None); }
    let snapshot = t.session.editor_snapshot().unwrap();
    assert_eq!(snapshot.undo_depth,MAX_HISTORY_ENTRIES); assert!(snapshot.history_weight <= MAX_HISTORY_WEIGHT);
    let weight = snapshot.history_weight;
    for _ in 0..MAX_HISTORY_ENTRIES { t.undo(); }
    assert_eq!(t.source(),"xxxxx"); assert_eq!(t.session.editor_snapshot().unwrap().history_weight,weight);
    assert!(!t.undo().content_changed);
    for _ in 0..MAX_HISTORY_ENTRIES { t.redo(); }
    assert_eq!(t.source().len(),MAX_HISTORY_ENTRIES+5); assert_eq!(t.session.editor_snapshot().unwrap().history_weight,weight);
}

#[test]
fn a_vertical_request_is_not_a_fake_source_based_cursor_movement() {
    let mut t = TestSession::new(r"\frac{x}{y}");
    let before = t.editor_json();
    let result = t.action(EditorAction::MoveVertical { direction:Vertical::Up,desired_x:Some(12.25) });
    assert_eq!(result.vertical_move.unwrap().desired_x,Some(12.25)); assert_eq!(t.editor_json(),before);
    assert!(matches!(t.try_action(EditorAction::MoveVertical { direction:Vertical::Down,desired_x:Some(f64::NAN) }),Err(e) if e.code == ErrorCode::InvalidInput));
    assert_eq!(t.editor_json(),before);
}

#[test]
fn seeded_edit_history_roundtrips_every_recorded_state_with_duplicate_formulas() {
    let mut t = TestSession::new("");
    let mut before = Vec::new(); let mut after = Vec::new(); let mut seed = 971_u64;
    for index in 0..60 {
        let count = match &t.session.document_snapshot().analysis.tree.kind { NodeKind::Row { entries, .. } => entries.len(), _=>0 };
        t.select(Selection::caret(Position::row(NodeId(1),count)));
        seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
        let command = match (seed >> 32) % 5 {
            0 => EditCommand::PasteLatex { latex:format!("\\frac{{x_{{{index}}}}}{{y}}") },
            1 => EditCommand::InsertText { text:"速度😀e\u{301}".into(),mode:InputMode::Text },
            2 => EditCommand::InsertText { text:"+x".into(),mode:InputMode::Math },
            3 => {
                if count > 0 { t.select(Selection::Range { anchor:Position::row(NodeId(1),count-1),focus:Position::row(NodeId(1),count) }); }
                EditCommand::DuplicateSelection
            }
            _ => EditCommand::InsertStructure { structure:StructureKind::Fraction },
        };
        let first = t.frame(); let result = t.edit(command);
        if result.content_changed { before.push(first); after.push(t.frame()); }
    }
    for frame in before.iter().rev() { t.undo(); assert_eq!(t.frame(),*frame); }
    for frame in &after { t.redo(); assert_eq!(t.frame(),*frame); }
}
