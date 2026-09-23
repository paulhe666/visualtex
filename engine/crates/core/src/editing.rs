//! The editing API of EngineSession. All content routes—including source edits
//! in session.rs—commit through the same history/selection owner here.
use std::collections::VecDeque;
use std::sync::Arc;
use serde::{Deserialize, Serialize};
use crate::composition::Composition;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::history::{History, StoredState};
use crate::model::{IdAllocator, NodeKind};
use crate::parser::ParseResult;
use crate::position_map::PositionMap;
use crate::positions::{Horizontal, NormalizedRange, Position, Selection, TreeIndex};
use crate::session::{DocumentSnapshot, EngineSession, NodeId, Revision};
use crate::transactions::{self, EditCommand, InputMode};

const MAX_EVENT_RECEIPTS: usize = 64;
const MAX_RECEIPT_BYTES: usize = 1_048_576;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Vertical { Up, Down }

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase", deny_unknown_fields)]
pub enum EditorAction {
    Edit { command: EditCommand, group_id: Option<String> },
    ApplyOcrLatex { target: crate::ocr::OcrInsertionTarget, latex: String },
    SetSelection { selection: Selection },
    PointerSelection { anchor: Position, focus: Position },
    ExtendGridSelection { direction: crate::grid_selection::GridDirection },
    MoveHorizontal { direction: Horizontal, extend: bool },
    MoveVertical { direction: Vertical, desired_x: Option<f64> },
    Undo,
    Redo,
    BeginComposition { mode: InputMode },
    UpdateComposition { id: u32, preedit: String },
    CommitComposition { id: u32, text: Option<String> },
    CancelComposition { id: u32 },
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct VerticalMoveRequest {
    pub selection: Selection,
    pub direction: Vertical,
    pub desired_x: Option<f64>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditorSnapshot {
    pub document: DocumentSnapshot,
    pub selection: Selection,
    pub normalized_selection: NormalizedRange,
    pub selection_epoch: u32,
    pub undo_depth: usize,
    pub redo_depth: usize,
    pub history_weight: usize,
    pub composition: Option<Composition>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditorActionResult {
    pub editor: EditorSnapshot,
    pub content_changed: bool,
    pub replayed: bool,
    pub position_map: Option<PositionMap>,
    pub vertical_move: Option<VerticalMoveRequest>,
}

struct EventReceipt {
    sequence: u32,
    expected_revision: Revision,
    action: String,
}

pub(crate) struct EditorState {
    pub selection: Selection,
    pub selection_epoch: u32,
    pub history: History,
    pub composition: Option<Composition>,
    pub(crate) ocr: crate::ocr::OcrState,
    next_composition_id: u32,
    last_event_sequence: u32,
    receipts: VecDeque<EventReceipt>,
    receipt_bytes: usize,
    can_group: bool,
    previous_input_mode: Option<InputMode>,
}
impl EditorState {
    pub fn new(root_id: NodeId, end_offset: usize) -> Self {
        Self { selection: Selection::caret(Position::row(root_id, end_offset)), selection_epoch: 0,
            history: History::default(), composition: None, ocr: crate::ocr::OcrState::default(), next_composition_id: 1,
            last_event_sequence: 0, receipts: VecDeque::new(), receipt_bytes: 0,
            can_group: false, previous_input_mode: None }
    }
    fn break_group(&mut self) { self.can_group = false; self.previous_input_mode = None; }
}

#[derive(Default)]
struct Outcome { content_changed: bool, position_map: Option<PositionMap>, vertical_move: Option<VerticalMoveRequest> }

impl EngineSession {
    pub fn editor_snapshot(&self) -> EngineResult<EditorSnapshot> {
        let index = TreeIndex::new(&self.document.tree)?;
        Ok(EditorSnapshot { document: self.document_snapshot(), selection: self.editor.selection.clone(),
            normalized_selection: index.normalize_range(&self.editor.selection)?, selection_epoch: self.editor.selection_epoch,
            undo_depth: self.editor.history.undo_depth(), redo_depth: self.editor.history.redo_depth(),
            history_weight: self.editor.history.weight(), composition: self.editor.composition.clone() })
    }

    pub(crate) fn ensure_no_composition(&self) -> EngineResult<()> {
        if self.editor.composition.is_some() {
            Err(EngineError::new(ErrorCode::Conflict, "Finish or cancel the current composition before another content operation"))
        } else { Ok(()) }
    }

    /// Event sequence numbers are scoped to this engine session, not an Office
    /// wire session. A retired sequence is never executed again after eviction.
    pub fn apply_action(&mut self, expected: Revision, event_sequence: u32, action: EditorAction) -> EngineResult<EditorActionResult> {
        if event_sequence == 0 {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Editor event sequence must be positive"));
        }
        let serialized = serde_json::to_string(&action)
            .map_err(|_| EngineError::new(ErrorCode::InvalidInput, "Editor action cannot be serialized"))?;
        if serialized.len() > MAX_RECEIPT_BYTES {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Editor event exceeds the retained replay budget"));
        }
        if event_sequence <= self.editor.last_event_sequence {
            if let Some(receipt) = self.editor.receipts.iter().find(|receipt| receipt.sequence == event_sequence) {
                if receipt.expected_revision == expected && receipt.action == serialized {
                    return Ok(EditorActionResult { editor: self.editor_snapshot()?, content_changed: false,
                        replayed: true, position_map: None, vertical_move: None });
                }
            }
            return Err(EngineError::new(ErrorCode::Conflict, "Editor event sequence was already used or retired"));
        }
        self.check_revision(expected)?;
        let outcome = match action {
            EditorAction::Edit { command, group_id } => {
                self.ensure_no_composition()?;
                self.apply_math_command(command, group_id, false)?
            }
            EditorAction::ApplyOcrLatex { target, latex } => {
                self.validate_ocr_target(&target)?;
                if latex.trim().is_empty(){return Err(EngineError::new(ErrorCode::InvalidInput,"OCR result is empty; no content was inserted"));}
                let outcome=self.apply_math_command(EditCommand::PasteLatex{latex},None,false)?;
                self.invalidate_ocr_target();outcome
            }
            EditorAction::SetSelection { selection } => {
                self.set_editor_selection(selection)?; self.editor.break_group(); Outcome::default()
            }
            EditorAction::PointerSelection { anchor, focus } => {
                self.ensure_no_composition()?;
                let next=TreeIndex::new(&self.document.tree)?.pointer_selection(anchor,focus)?;
                self.set_editor_selection(next)?;self.editor.break_group();Outcome::default()
            }
            EditorAction::ExtendGridSelection { direction } => {
                self.ensure_no_composition()?;
                let next=TreeIndex::new(&self.document.tree)?.extend_grid_selection(&self.editor.selection,direction)?;
                self.set_editor_selection(next)?;self.editor.break_group();Outcome::default()
            }
            EditorAction::MoveHorizontal { direction, extend } => {
                let next = TreeIndex::new(&self.document.tree)?.move_horizontal(&self.editor.selection, direction, extend)?;
                self.set_editor_selection(next)?; self.editor.break_group(); Outcome::default()
            }
            EditorAction::MoveVertical { direction, desired_x } => {
                if desired_x.is_some_and(|x| !x.is_finite()) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Desired horizontal coordinate must be finite"));
                }
                self.editor.break_group();
                Outcome { vertical_move: Some(VerticalMoveRequest { selection: self.editor.selection.clone(), direction, desired_x }), ..Outcome::default() }
            }
            EditorAction::Undo => { self.ensure_no_composition()?; self.replay_history(false)? }
            EditorAction::Redo => { self.ensure_no_composition()?; self.replay_history(true)? }
            EditorAction::BeginComposition { mode } => { self.begin_composition(mode)?; self.invalidate_ocr_target(); Outcome::default() }
            EditorAction::UpdateComposition { id, preedit } => {
                let composition = self.require_composition(id)?;
                composition.validate_target(self.revision, self.editor.selection_epoch, &self.editor.selection)?;
                Composition::validate_text(&preedit)?;
                self.editor.composition.as_mut().unwrap().preedit = preedit;
                Outcome::default()
            }
            EditorAction::CommitComposition { id, text } => {
                let composition = self.require_composition(id)?.clone();
                composition.validate_target(self.revision, self.editor.selection_epoch, &self.editor.selection)?;
                let text = text.unwrap_or(composition.preedit);
                Composition::validate_text(&text)?;
                let outcome = if text.is_empty() { Outcome::default() }
                    else { self.apply_math_command(EditCommand::InsertText { text, mode: composition.mode }, None, true)? };
                self.editor.composition = None; self.editor.break_group(); outcome
            }
            EditorAction::CancelComposition { id } => {
                self.require_composition(id)?;
                self.editor.composition = None; self.editor.break_group(); Outcome::default()
            }
        };
        let editor = self.editor_snapshot()?;
        self.editor.last_event_sequence = event_sequence;
        while self.editor.receipts.len() >= MAX_EVENT_RECEIPTS
            || self.editor.receipt_bytes + serialized.len() > MAX_RECEIPT_BYTES
        {
            if let Some(receipt) = self.editor.receipts.pop_front() { self.editor.receipt_bytes -= receipt.action.len(); }
            else { break; }
        }
        self.editor.receipt_bytes += serialized.len();
        self.editor.receipts.push_back(EventReceipt { sequence: event_sequence, expected_revision: expected, action: serialized });
        Ok(EditorActionResult { editor, content_changed: outcome.content_changed, replayed: false,
            position_map: outcome.position_map, vertical_move: outcome.vertical_move })
    }

    fn set_editor_selection(&mut self, selection: Selection) -> EngineResult<()> {
        TreeIndex::new(&self.document.tree)?.validate_selection(&selection)?;
        if selection != self.editor.selection {
            let epoch = self.editor.selection_epoch.checked_add(1)
                .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Selection epoch exhausted"))?;
            self.editor.selection = selection; self.editor.selection_epoch = epoch;
        }
        Ok(())
    }

    fn apply_math_command(&mut self, command: EditCommand, group_id: Option<String>, from_composition: bool) -> EngineResult<Outcome> {
        if group_id.as_ref().is_some_and(|g| g.is_empty() || g.len() > 128 || g.chars().any(char::is_control)) {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid input history group"));
        }
        let mut ids = IdAllocator::new(self.next_node_id);
        let prepared = transactions::prepare(&self.document, &self.editor.selection, &command, self.next_source_id, &mut ids)?;
        let input_mode = match command {
            EditCommand::InsertText { mode, .. } => Some(mode),
            EditCommand::ReplaceRecentMath { .. } => Some(InputMode::Math),
            _ => None,
        };
        if prepared.changed {
            let allow_group = !from_composition && self.editor.can_group && input_mode.is_some()
                && input_mode == self.editor.previous_input_mode;
            let group_id = if input_mode.is_some() && !from_composition { group_id } else { None };
            let map = self.commit_document(prepared.document, ids.next_unused(), Some(prepared.selection),
                Some(prepared.position_map), group_id, allow_group, from_composition)?;
            self.editor.can_group = input_mode.is_some() && !from_composition;
            self.editor.previous_input_mode = input_mode;
            Ok(Outcome { content_changed: true, position_map: Some(map), vertical_move: None })
        } else {
            self.set_editor_selection(prepared.selection)?;
            if input_mode.is_none() { self.editor.break_group(); }
            Ok(Outcome { position_map: Some(prepared.position_map), ..Outcome::default() })
        }
    }

    /// All fallible identity, geometry-independent position and history-budget
    /// checks happen before any committed document/counter is replaced.
    pub(crate) fn commit_document(&mut self, document: Arc<ParseResult>, next_node_id: u32,
        selection: Option<Selection>, prepared_map: Option<PositionMap>, group_id: Option<String>,
        allow_group: bool, from_composition: bool) -> EngineResult<PositionMap>
    {
        if !from_composition { self.ensure_no_composition()?; }
        if document.source.id() != self.next_source_id || document.tree.id != self.root_id || next_node_id < self.next_node_id {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Content commit identities do not match the session owner"));
        }
        let next_source_id = self.next_source_id.checked_add(1)
            .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Source identity space exhausted"))?;
        let revision = self.revision.0.checked_add(1)
            .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Revision space exhausted"))?;
        let map = match prepared_map { Some(map) => map, None => PositionMap::between(&self.document.tree, &document.tree, Vec::new())? };
        let new_index = TreeIndex::new(&document.tree)?;
        let selection = match selection { Some(selection) => selection, None => map.apply_selection(&self.editor.selection, &new_index)? };
        new_index.validate_selection(&selection)?;
        let selection_epoch = if selection != self.editor.selection { self.editor.selection_epoch.checked_add(1)
            .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Selection epoch exhausted"))? } else { self.editor.selection_epoch };
        let last_complete = if document.complete { Some(Arc::clone(&document.tree)) } else { self.last_complete.clone() };
        let before = StoredState { document: Arc::clone(&self.document), selection: self.editor.selection.clone(), last_complete: self.last_complete.clone() };
        let after = StoredState { document: Arc::clone(&document), selection: selection.clone(), last_complete: last_complete.clone() };
        self.editor.history.record(before, after, group_id, allow_group)?;
        self.document = document; self.last_complete = last_complete;
        self.next_node_id = next_node_id; self.next_source_id = next_source_id; self.revision = Revision(revision);
        self.editor.selection = selection; self.editor.selection_epoch = selection_epoch;
        if !allow_group { self.editor.break_group(); }
        Ok(map)
    }

    fn replay_history(&mut self, redo: bool) -> EngineResult<Outcome> {
        let entry = if redo { self.editor.history.peek_redo() } else { self.editor.history.peek_undo() }.cloned();
        let Some(entry) = entry else { self.editor.break_group(); return Ok(Outcome::default()); };
        let state = if redo { entry.after } else { entry.before };
        let map = PositionMap::between(&self.document.tree, &state.document.tree, Vec::new())?;
        TreeIndex::new(&state.document.tree)?.validate_selection(&state.selection)?;
        let revision = self.revision.0.checked_add(1)
            .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Revision space exhausted"))?;
        let selection_epoch = self.editor.selection_epoch.checked_add(1)
            .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Selection epoch exhausted"))?;
        if redo { self.editor.history.accept_redo()?; } else { self.editor.history.accept_undo()?; }
        self.document = state.document; self.last_complete = state.last_complete; self.editor.selection = state.selection;
        self.revision = Revision(revision); self.editor.selection_epoch = selection_epoch;
        // next_node_id, next_source_id and event/composition sequence identities
        // are intentionally not rewound by document history.
        self.editor.break_group();
        Ok(Outcome { content_changed: true, position_map: Some(map), vertical_move: None })
    }

    fn begin_composition(&mut self, mode: InputMode) -> EngineResult<()> {
        self.ensure_no_composition()?;
        if !self.document.complete { return Err(EngineError::new(ErrorCode::IncompleteSource, "Resolve source draft before visual composition")); }
        if !self.document.supported { return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Visual composition requires supported structure")); }
        let id = self.editor.next_composition_id;
        let next = id.checked_add(1).ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Composition identity space exhausted"))?;
        self.editor.composition = Some(Composition { id, revision: self.revision, selection_epoch: self.editor.selection_epoch,
            selection: self.editor.selection.clone(), mode, preedit: String::new() });
        self.editor.next_composition_id = next; self.editor.break_group(); Ok(())
    }
    fn require_composition(&self, id: u32) -> EngineResult<&Composition> {
        self.editor.composition.as_ref().filter(|c| c.id == id)
            .ok_or_else(|| EngineError::new(ErrorCode::Conflict, "Composition is absent, retired or belongs to another input"))
    }

    /// Initial loading is not an edit: no artificial empty-document undo entry.
    /// Source and node IDs still belong to the same newly created owner.
    pub fn from_source(document_id: String, source: String) -> EngineResult<Self> {
        Self::from_source_with_symbols(document_id,source,Arc::default())
    }
    pub fn from_source_with_symbols(document_id: String, source: String, symbols:Arc<crate::custom_symbols::SymbolCatalog>) -> EngineResult<Self> {
        let mut session = Self::new(document_id)?;
        let mut ids = IdAllocator::new(2);
        let document = Arc::new(crate::parser::parse_with_symbols(crate::source::SourceText::new(1, source)?, session.root_id, &mut ids, symbols)?);
        let end = match &document.tree.kind { NodeKind::Row { entries, .. } => entries.len(), _ => 0 };
        let editor = EditorState::new(session.root_id, end);
        TreeIndex::new(&document.tree)?.validate_selection(&editor.selection)?;
        session.next_node_id = ids.next_unused();
        session.last_complete = document.complete.then(|| Arc::clone(&document.tree));
        session.document = document; session.editor = editor;
        Ok(session)
    }
}
