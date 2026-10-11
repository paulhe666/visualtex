//! Position transforms shared by visual edits, source edits and history replay.
use std::collections::BTreeSet;
use serde::Serialize;
use unicode_segmentation::UnicodeSegmentation;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{Node, NodeKind};
use crate::positions::{Affinity, Position, Selection, TreeIndex};
use crate::session::NodeId;

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
pub enum PositionStep {
    RowSplice { node_id: NodeId, start: u32, removed: u32, inserted: u32 },
    RowSplit { node_id: NodeId, left_id: NodeId, right_id: NodeId, at: u32, right_offset: u32 },
    RowOffset { node_id: NodeId, target_id: NodeId, offset: u32 },
    TextRemap { node_id: NodeId, target_id: NodeId, before: Vec<u32>, after: Vec<u32> },
    TextSplit { node_id: NodeId, right_id: NodeId, at: u32 },
    Removed { node_id: NodeId, fallback: Position },
}

impl PositionStep {
    fn source_id(&self) -> NodeId {
        match self {
            Self::RowSplice { node_id, .. } | Self::TextRemap { node_id, .. }
                | Self::TextSplit { node_id, .. } | Self::Removed { node_id, .. }
                | Self::RowSplit { node_id, .. } | Self::RowOffset { node_id, .. } => *node_id,
        }
    }
    fn apply(&self, position: Position) -> EngineResult<Position> {
        match (self, position) {
            (Self::RowSplice { node_id, start, removed, inserted }, Position::Row { node_id: id, offset, affinity }) if *node_id == id => {
                let end = start.checked_add(*removed).ok_or_else(invalid_map)?;
                let offset = if offset < *start { offset }
                    else if offset > end { offset.checked_sub(*removed).and_then(|v| v.checked_add(*inserted)).ok_or_else(invalid_map)? }
                    else if affinity == Affinity::Before { *start }
                    else { start.checked_add(*inserted).ok_or_else(invalid_map)? };
                Ok(Position::Row { node_id: id, offset, affinity })
            }
            (Self::RowSplit { node_id, left_id, right_id, at, right_offset }, Position::Row { node_id:id,offset,affinity }) if *node_id==id => {
                let (node_id,offset)=if offset>*at || (offset==*at && affinity==Affinity::After){(*right_id,(offset-at).checked_add(*right_offset).ok_or_else(invalid_map)?)}else{(*left_id,offset)};
                Ok(Position::Row{node_id,offset,affinity})
            }
            (Self::RowOffset {node_id,target_id,offset:delta},Position::Row{node_id:id,offset,affinity}) if *node_id==id => {
                Ok(Position::Row{node_id:*target_id,offset:offset.checked_add(*delta).ok_or_else(invalid_map)?,affinity})
            }
            (Self::TextRemap { node_id, target_id, before, after }, Position::Text { node_id: id, grapheme, affinity }) if *node_id == id => {
                let target = if affinity == Affinity::Before { before } else { after };
                let grapheme = *target.get(grapheme as usize).ok_or_else(invalid_map)?;
                Ok(Position::Text { node_id: *target_id, grapheme, affinity })
            }
            (Self::TextSplit { node_id, right_id, at }, Position::Text { node_id: id, grapheme, affinity }) if *node_id == id => {
                if grapheme > *at || (grapheme == *at && affinity == Affinity::After) {
                    Ok(Position::Text { node_id: *right_id, grapheme: grapheme - at, affinity })
                } else { Ok(position) }
            }
            (Self::Removed { node_id, fallback }, _) if *node_id == position.node_id() => Ok(*fallback),
            _ => Ok(position),
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PositionMap {
    pub retained_ids: Vec<NodeId>,
    pub removed_ids: Vec<NodeId>,
    pub inserted_ids: Vec<NodeId>,
    pub steps: Vec<PositionStep>,
}

impl PositionMap {
    pub fn between(old: &Node, new: &Node, overrides: Vec<PositionStep>) -> EngineResult<Self> {
        Self::between_indexed(&TreeIndex::new(old)?, &TreeIndex::new(new)?, overrides)
    }
    /// `between` for callers that already hold both trees' indexes.
    pub fn between_indexed(old_index: &TreeIndex<'_>, new_index: &TreeIndex<'_>, overrides: Vec<PositionStep>) -> EngineResult<Self> {
        let new = new_index.root;
        let old_ids: BTreeSet<_> = old_index.nodes.keys().copied().collect();
        let new_ids: BTreeSet<_> = new_index.nodes.keys().copied().collect();
        let retained_ids: Vec<_> = old_ids.intersection(&new_ids).copied().collect();
        let removed_ids: Vec<_> = old_ids.difference(&new_ids).copied().collect();
        let inserted_ids = new_ids.difference(&old_ids).copied().collect();
        let override_ids: BTreeSet<_> = overrides.iter().map(PositionStep::source_id).collect();
        let mut steps = Vec::new();
        for id in &retained_ids {
            if override_ids.contains(id) { continue; }
            let old_node = old_index.node(*id)?;
            let new_node = new_index.node(*id)?;
            match (&old_node.kind, &new_node.kind) {
                (NodeKind::Row { entries: a, .. }, NodeKind::Row { entries: b, .. }) => {
                    let a: Vec<_> = a.iter().map(|e| e.node.id).collect();
                    let b: Vec<_> = b.iter().map(|e| e.node.id).collect();
                    if a != b {
                        let prefix = a.iter().zip(&b).take_while(|(a,b)| a == b).count();
                        let suffix = a[prefix..].iter().rev().zip(b[prefix..].iter().rev()).take_while(|(a,b)| a == b).count();
                        steps.push(PositionStep::RowSplice { node_id: *id, start: prefix as u32,
                            removed: (a.len() - prefix - suffix) as u32, inserted: (b.len() - prefix - suffix) as u32 });
                    }
                }
                (NodeKind::Text { text: a, .. }, NodeKind::Text { text: b, .. }) if a != b => {
                    steps.push(text_change_map(*id, a, b));
                }
                _ => {}
            }
        }
        steps.extend(overrides);
        let mut fallback_steps = Vec::new();
        for id in &removed_ids {
            if override_ids.contains(id) { continue; }
            let mut current = *id;
            let mut fallback = Position::row(new.id, 0);
            while let Ok((row, offset)) = old_index.row_owner(current) {
                if new_ids.contains(&row) && new_index.node(row)?.is_row() {
                    fallback = Position::Row { node_id: row, offset: offset as u32, affinity: Affinity::Before };
                    for step in &steps { fallback = step.apply(fallback)?; }
                    break;
                }
                current = row;
            }
            new_index.validate_position(fallback)?;
            fallback_steps.push(PositionStep::Removed { node_id: *id, fallback });
        }
        steps.extend(fallback_steps);
        Ok(Self { retained_ids, removed_ids, inserted_ids, steps })
    }
    pub fn apply_position(&self, position: Position, result: &TreeIndex<'_>) -> EngineResult<Position> {
        let mut position = position;
        for step in &self.steps { position = step.apply(position)?; }
        result.validate_position(position)?;
        Ok(position)
    }
    pub fn apply_selection(&self, selection: &Selection, result: &TreeIndex<'_>) -> EngineResult<Selection> {
        let selection = match selection {
            Selection::GridCells{grid_id,anchor_cell,focus_cell} => {
                if result.grid_rectangle(*grid_id,*anchor_cell,*focus_cell).is_ok() {selection.clone()}
                else {Selection::caret(self.apply_position(Position::row(*focus_cell,0),result)?)}
            }
            Selection::Range { anchor, focus } => Selection::Range {
                anchor: self.apply_position(*anchor, result)?, focus: self.apply_position(*focus, result)?,
            },
            Selection::Node { node_id } if result.nodes.contains_key(node_id) => selection.clone(),
            Selection::Node { node_id } => {
                let fallback = self.steps.iter().find_map(|step| match step {
                    PositionStep::Removed { node_id: id, fallback } if id == node_id => Some(*fallback),
                    PositionStep::RowOffset { node_id:id,target_id,offset } if id==node_id => Some(Position::row(*target_id,*offset as usize)),
                    PositionStep::RowSplit { node_id:id,left_id,.. } if id==node_id => Some(Position::row(*left_id,0)),
                    PositionStep::TextRemap { node_id: id, target_id, before, .. } if id == node_id =>
                        before.first().map(|g| Position::text(*target_id, *g as usize)),
                    _ => None,
                }).ok_or_else(invalid_map)?;
                Selection::caret(fallback)
            }
        };
        result.validate_selection(&selection)?;
        Ok(selection)
    }
}

fn invalid_map() -> EngineError { EngineError::new(ErrorCode::InvalidPosition, "Position map does not cover a valid input position") }

fn grapheme_boundaries(text: &str) -> Vec<usize> {
    let mut boundaries: Vec<_> = text.grapheme_indices(true).map(|(byte, _)| byte).collect();
    boundaries.push(text.len()); boundaries
}
fn adjacent_graphemes(boundaries: &[usize], byte: usize) -> (u32, u32) {
    match boundaries.binary_search(&byte) {
        Ok(index) => (index as u32, index as u32),
        Err(index) => (index.saturating_sub(1) as u32, index.min(boundaries.len()-1) as u32),
    }
}

/// A text edit may join combining marks or a ZWJ sequence across the insertion.
/// Transform byte positions first, then snap to valid extended-grapheme edges.
fn text_change_map(id: NodeId, old: &str, new: &str) -> PositionStep {
    let prefix = old.chars().zip(new.chars()).take_while(|(a,b)| a == b).map(|(c,_)| c.len_utf8()).sum::<usize>();
    let suffix = old[prefix..].chars().rev().zip(new[prefix..].chars().rev())
        .take_while(|(a,b)| a == b).map(|(c,_)| c.len_utf8()).sum::<usize>();
    let old_end = old.len() - suffix;
    let new_end = new.len() - suffix;
    let new_boundaries = grapheme_boundaries(new);
    let mut before = Vec::new(); let mut after = Vec::new();
    for byte in grapheme_boundaries(old) {
        let pair = if byte < prefix { (byte, byte) }
            else if byte > old_end { (byte - old_end + new_end, byte - old_end + new_end) }
            else { (prefix, new_end) };
        before.push(adjacent_graphemes(&new_boundaries, pair.0).0);
        after.push(adjacent_graphemes(&new_boundaries, pair.1).1);
    }
    PositionStep::TextRemap { node_id: id, target_id: id, before, after }
}

pub fn merge_text_map(old_id: NodeId, target_id: NodeId, old: &str, merged: &str, byte_offset: usize) -> EngineResult<PositionStep> {
    if !merged.is_char_boundary(byte_offset) || byte_offset + old.len() > merged.len()
        || &merged[byte_offset..byte_offset+old.len()] != old
    { return Err(invalid_map()); }
    let new_boundaries = grapheme_boundaries(merged);
    let mut before = Vec::new(); let mut after = Vec::new();
    for byte in grapheme_boundaries(old) {
        let pair = adjacent_graphemes(&new_boundaries, byte_offset + byte);
        before.push(pair.0); after.push(pair.1);
    }
    Ok(PositionStep::TextRemap { node_id: old_id, target_id, before, after })
}
