//! Bounded structural history. Entries own immutable document/tree snapshots and
//! selections; they do not infer operations or cursor locations from LaTeX diffs.
use std::collections::VecDeque;
use std::sync::Arc;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::Node;
use crate::parser::ParseResult;
use crate::positions::Selection;

pub const MAX_HISTORY_ENTRIES: usize = 128;
pub const MAX_HISTORY_WEIGHT: usize = 32 * 1024 * 1024;

#[derive(Clone)]
pub(crate) struct StoredState {
    pub document: Arc<ParseResult>,
    pub selection: Selection,
    pub last_complete: Option<Arc<Node>>,
}
impl StoredState {
    fn weight(&self) -> usize {
        // Conservative accounting units for source/index/token/tree ownership.
        // This is a budget weight, not a claim to measure allocator/RSS bytes.
        let retained_nodes = if self.document.complete { 0 } else {
            let mut count = 0;
            if let Some(tree) = &self.last_complete { tree.visit(&mut |_| count += 1); }
            count
        };
        self.document.source.len().saturating_mul(24)
            .saturating_add(self.document.tokens.len().saturating_mul(80))
            .saturating_add((self.document.facts.nodes + retained_nodes).saturating_mul(512))
            .saturating_add(self.document.symbol_assets.retained_bytes())
    }
}

#[derive(Clone)]
pub(crate) struct HistoryEntry {
    pub before: StoredState,
    pub after: StoredState,
    group: Option<String>,
    weight: usize,
}

#[derive(Default)]
pub(crate) struct History {
    undo: VecDeque<HistoryEntry>,
    redo: Vec<HistoryEntry>,
    weight: usize,
}
impl History {
    pub fn undo_depth(&self) -> usize { self.undo.len() }
    pub fn redo_depth(&self) -> usize { self.redo.len() }
    pub fn weight(&self) -> usize { self.weight }
    pub fn peek_undo(&self) -> Option<&HistoryEntry> { self.undo.back() }
    pub fn peek_redo(&self) -> Option<&HistoryEntry> { self.redo.last() }
    pub fn record(&mut self, before: StoredState, after: StoredState, group: Option<String>, allow_group: bool) -> EngineResult<()> {
        if group.as_ref().is_some_and(|g| g.is_empty() || g.len() > 128 || g.chars().any(char::is_control)) {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid input history group"));
        }
        let merge = allow_group && group.is_some() && self.redo.is_empty() && self.undo.back().is_some_and(|entry|
            entry.group == group && Arc::ptr_eq(&entry.after.document, &before.document) && entry.after.selection == before.selection);
        let before = if merge { self.undo.back().unwrap().before.clone() } else { before };
        let weight = before.weight().saturating_add(after.weight());
        if weight > MAX_HISTORY_WEIGHT {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "A single edit exceeds the bounded history weight budget"));
        }
        let entry = HistoryEntry { before, after, group, weight };
        // All fallible validation completed. Only now change stacks/ownership.
        for old in self.redo.drain(..) { self.weight -= old.weight; }
        if merge { self.weight -= self.undo.pop_back().unwrap().weight; }
        while self.undo.len() >= MAX_HISTORY_ENTRIES || self.weight + weight > MAX_HISTORY_WEIGHT {
            if let Some(old) = self.undo.pop_front() { self.weight -= old.weight; }
            else { break; }
        }
        self.weight += weight;
        self.undo.push_back(entry);
        Ok(())
    }
    pub fn accept_undo(&mut self) -> EngineResult<()> {
        let entry = self.undo.pop_back().ok_or_else(|| EngineError::new(ErrorCode::InvalidInput, "Undo history is empty"))?;
        self.redo.push(entry); Ok(())
    }
    pub fn accept_redo(&mut self) -> EngineResult<()> {
        let entry = self.redo.pop().ok_or_else(|| EngineError::new(ErrorCode::InvalidInput, "Redo history is empty"))?;
        self.undo.push_back(entry); Ok(())
    }
}
