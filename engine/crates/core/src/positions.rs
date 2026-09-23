//! Mathematical positions and structural navigation. Source offsets and
//! third-party editor offsets are deliberately absent from these types.
use std::collections::BTreeMap;
use serde::{Deserialize, Serialize};
use unicode_segmentation::UnicodeSegmentation;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{Node, NodeKind};
use crate::session::NodeId;

const MAX_CARETS: usize = 65_536;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Affinity { Before, After }

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase", deny_unknown_fields)]
pub enum Position {
    Row { node_id: NodeId, offset: u32, affinity: Affinity },
    Text { node_id: NodeId, grapheme: u32, affinity: Affinity },
}

impl Position {
    pub fn row(node_id: NodeId, offset: usize) -> Self {
        Self::Row { node_id, offset: u32::try_from(offset).unwrap_or(u32::MAX), affinity: Affinity::After }
    }
    pub fn text(node_id: NodeId, grapheme: usize) -> Self {
        Self::Text { node_id, grapheme: u32::try_from(grapheme).unwrap_or(u32::MAX), affinity: Affinity::After }
    }
    pub fn node_id(self) -> NodeId { match self { Self::Row { node_id, .. } | Self::Text { node_id, .. } => node_id } }
    pub fn offset(self) -> usize { match self { Self::Row { offset, .. } => offset as usize, Self::Text { grapheme, .. } => grapheme as usize } }
    pub fn affinity(self) -> Affinity { match self { Self::Row { affinity, .. } | Self::Text { affinity, .. } => affinity } }
    pub fn same_location(self, other: Self) -> bool {
        match (self, other) {
            (Self::Row { node_id: a, offset: x, .. }, Self::Row { node_id: b, offset: y, .. }) => a == b && x == y,
            (Self::Text { node_id: a, grapheme: x, .. }, Self::Text { node_id: b, grapheme: y, .. }) => a == b && x == y,
            _ => false,
        }
    }
    fn key(self) -> (u8, NodeId, u32) {
        match self {
            Self::Row { node_id, offset, .. } => (0, node_id, offset),
            Self::Text { node_id, grapheme, .. } => (1, node_id, grapheme),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase", deny_unknown_fields)]
pub enum Selection {
    Range { anchor: Position, focus: Position },
    Node { node_id: NodeId },
    GridCells { grid_id: NodeId, anchor_cell: NodeId, focus_cell: NodeId },
}
impl Selection {
    pub fn caret(position: Position) -> Self { Self::Range { anchor: position, focus: position } }
    pub fn is_collapsed(&self) -> bool {
        matches!(self, Self::Range { anchor, focus } if anchor.same_location(*focus))
    }
    pub fn focus(&self) -> Option<Position> { match self { Self::Range { focus, .. } => Some(*focus), _ => None } }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Horizontal { Left, Right }

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
pub enum NormalizedRange {
    Row { node_id: NodeId, start: u32, end: u32, promoted: bool },
    Text { node_id: NodeId, start: u32, end: u32 },
    GridCells { rectangle: crate::grid_selection::GridRectangle },
}

pub struct IndexEntry<'a> {
    pub node: &'a Node,
    pub parent: Option<NodeId>,
    pub index_in_row: Option<usize>,
}

pub struct TreeIndex<'a> {
    pub root: &'a Node,
    pub nodes: BTreeMap<NodeId, IndexEntry<'a>>,
    pub carets: Vec<Position>,
    ranks: BTreeMap<(u8, NodeId, u32), usize>,
}

impl<'a> TreeIndex<'a> {
    pub fn new(root: &'a Node) -> EngineResult<Self> {
        crate::model::validate_tree(root)?;
        let mut result = Self { root, nodes: BTreeMap::new(), carets: Vec::new(), ranks: BTreeMap::new() };
        result.collect(root, None, None)?;
        Ok(result)
    }
    fn push(&mut self, position: Position) -> EngineResult<()> {
        if self.carets.len() >= MAX_CARETS {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Mathematical caret budget exhausted"));
        }
        self.ranks.insert(position.key(), self.carets.len());
        self.carets.push(position);
        Ok(())
    }
    fn collect(&mut self, node: &'a Node, parent: Option<NodeId>, index_in_row: Option<usize>) -> EngineResult<()> {
        if self.nodes.insert(node.id, IndexEntry { node, parent, index_in_row }).is_some() {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Duplicate node identity in position index"));
        }
        match &node.kind {
            NodeKind::Row { entries, .. } => {
                self.push(Position::row(node.id, 0))?;
                for (index, entry) in entries.iter().enumerate() {
                    self.collect(&entry.node, Some(node.id), Some(index))?;
                    self.push(Position::row(node.id, index + 1))?;
                }
            }
            NodeKind::Text { text, .. } => {
                for offset in 0..=text.graphemes(true).count() { self.push(Position::text(node.id, offset))?; }
            }
            NodeKind::Grid { rows, .. } => for row in rows { self.collect(row,Some(node.id),None)?; },
            NodeKind::GridRow { cells, .. } => for cell in cells { self.collect(cell,Some(node.id),None)?; },
            NodeKind::Fraction { numerator, denominator, .. } => {
                self.collect(numerator, Some(node.id), None)?;
                self.collect(denominator, Some(node.id), None)?;
            }
            NodeKind::Fence { body, .. } | NodeKind::Accent { body, .. } | NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => self.collect(body, Some(node.id), None)?,
            NodeKind::Annotation { annotation, base, .. } => { self.collect(annotation, Some(node.id), None)?; self.collect(base, Some(node.id), None)?; }
            NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => for slot in crate::operators::navigation_slots(parts){self.collect(&slot.argument,Some(node.id),None)?;},
            NodeKind::GeneralOperator { body, parts, .. } => {
                for slot in crate::operators::navigation_slots(parts){self.collect(&slot.argument,Some(node.id),None)?;}
                self.collect(body,Some(node.id),None)?;
            }
            NodeKind::MacroCall { arguments, .. } => for argument in arguments {self.collect(argument,Some(node.id),None)?;},
            NodeKind::Radical { degree, radicand, .. } => {
                if let Some(degree) = degree { self.collect(degree, Some(node.id), None)?; }
                self.collect(radicand, Some(node.id), None)?;
            }
            NodeKind::Scripts { base, slots } => {
                self.collect(base, Some(node.id), None)?;
                for slot in slots { self.collect(&slot.argument, Some(node.id), None)?; }
            }
            _ => {}
        }
        Ok(())
    }
    pub fn node(&self, id: NodeId) -> EngineResult<&'a Node> {
        self.nodes.get(&id).map(|entry| entry.node).ok_or_else(invalid_position)
    }
    pub fn validate_position(&self, position: Position) -> EngineResult<()> {
        if self.ranks.contains_key(&position.key()) { Ok(()) } else { Err(invalid_position()) }
    }
    pub fn validate_selection(&self, selection: &Selection) -> EngineResult<()> {
        match selection {
            Selection::GridCells { grid_id, anchor_cell, focus_cell } => self.grid_rectangle(*grid_id,*anchor_cell,*focus_cell).map(|_|()),
            Selection::Range { anchor, focus } => { self.validate_position(*anchor)?; self.validate_position(*focus) }
            Selection::Node { node_id } => {
                if *node_id == self.root.id { return Err(invalid_position()); }
                self.node(*node_id).map(|_| ())
            }
        }
    }
    /// Raw cross-cell text ranges remain ambiguous and must not promote into
    /// deleting the whole grid. Identified cell rectangles use dedicated edits.
    pub fn validate_matrix_fragment(&self, selection: &Selection) -> EngineResult<()> {
        let owner = |mut id:NodeId| -> Option<(NodeId,NodeId)> {
            while let Some(parent)=self.nodes.get(&id).and_then(|entry|entry.parent) {
                if matches!(self.nodes.get(&parent)?.node.kind,NodeKind::GridRow{..}) {
                    let grid=self.nodes.get(&parent)?.parent?;
                    if matches!(self.nodes.get(&grid)?.node.kind,NodeKind::Grid{environment,..} if environment!=crate::grids::GridKind::Gathered) {
                        return Some((grid,id));
                    }
                }
                id=parent;
            }
            None
        };
        let unsupported=match selection {
            Selection::GridCells{..}=>false,
            Selection::Range{anchor,focus}=>owner(anchor.node_id())!=owner(focus.node_id()),
            Selection::Node{node_id}=>{
                let entry=self.nodes.get(node_id).ok_or_else(invalid_position)?;
                (matches!(entry.node.kind,NodeKind::GridRow{..})
                    && entry.parent.and_then(|id|self.nodes.get(&id)).is_some_and(|e|matches!(e.node.kind,
                        NodeKind::Grid{environment,..} if environment!=crate::grids::GridKind::Gathered)))
                    || (entry.parent.and_then(|id|self.nodes.get(&id)).is_some_and(|e|matches!(e.node.kind,NodeKind::GridRow{..}))
                        && owner(*node_id).is_some())
            }
        };
        if unsupported {return Err(EngineError::new(ErrorCode::InvalidPosition,
            "Cross-cell matrix fragments require rectangular selection; select one cell's content or the whole matrix explicitly"));}
        Ok(())
    }
    pub fn rank(&self, position: Position) -> EngineResult<usize> {
        self.ranks.get(&position.key()).copied().ok_or_else(invalid_position)
    }
    pub fn row_owner(&self, mut id: NodeId) -> EngineResult<(NodeId, usize)> {
        loop {
            let entry = self.nodes.get(&id).ok_or_else(invalid_position)?;
            if let (Some(parent), Some(index)) = (entry.parent, entry.index_in_row) { return Ok((parent, index)); }
            id = entry.parent.ok_or_else(invalid_position)?;
        }
    }
    pub fn ancestor_rows(&self, mut id: NodeId) -> Vec<NodeId> {
        let mut result = Vec::new();
        while let Some(entry) = self.nodes.get(&id) {
            if entry.node.is_row() { result.push(id); }
            match entry.parent { Some(parent) => id = parent, None => break }
        }
        result
    }
    fn boundary_equivalent(&self, a: Position, b: Position) -> bool {
        if a.same_location(b) { return true; }
        let (text_position, row_position) = match (a, b) {
            (Position::Text { .. }, Position::Row { .. }) => (a, b),
            (Position::Row { .. }, Position::Text { .. }) => (b, a),
            _ => return false,
        };
        let Some(entry) = self.nodes.get(&text_position.node_id()) else { return false };
        if entry.parent != Some(row_position.node_id()) { return false; }
        let Some(index) = entry.index_in_row else { return false };
        let NodeKind::Text { text, .. } = &entry.node.kind else { return false };
        (text_position.offset() == 0 && row_position.offset() == index)
            || (text_position.offset() == text.graphemes(true).count() && row_position.offset() == index + 1)
    }
    /// A completed accent chain is one navigation boundary. Do not cross a
    /// sibling, an unfinished script slot, or a fraction/annotation boundary.
    fn completed_accent_exit(&self, position: Position) -> EngineResult<Option<Position>> {
        let node=self.node(position.node_id())?;
        let at_end=match &node.kind {
            NodeKind::Row{entries,..}=>position.offset()==entries.len(),
            NodeKind::Text{text,..}=>position.offset()==text.graphemes(true).count(),
            _=>false,
        };
        if !at_end {return Ok(None);}
        let mut current=node.id;let mut target=None;
        while let Some(parent)=self.nodes.get(&current).and_then(|entry|entry.parent) {
            match &self.node(parent)?.kind {
                NodeKind::Row{entries,..} if entries.last().is_some_and(|entry|entry.node.id==current)=>current=parent,
                NodeKind::Scripts{slots,..} if slots.last().is_some_and(|slot|slot.argument.id==current)
                    && !self.node(parent)?.has_unfilled_slots()=>current=parent,
                NodeKind::Accent{body,..} if body.id==current && body.has_formula_content() && !body.has_unfilled_slots()=>{
                    let (row,offset)=self.row_owner(parent)?;
                    target=Some(Position::row(row,offset+1));current=parent;
                }
                _=>break,
            }
        }
        Ok(target)
    }
    pub fn move_horizontal(&self, selection: &Selection, direction: Horizontal, extend: bool) -> EngineResult<Selection> {
        self.validate_selection(selection)?;
        if let Selection::GridCells{grid_id,anchor_cell,focus_cell}=selection {
            if extend {return self.extend_grid_selection(selection,if direction==Horizontal::Left {
                crate::grid_selection::GridDirection::Left
            }else{crate::grid_selection::GridDirection::Right});}
            let rect=self.grid_rectangle(*grid_id,*anchor_cell,*focus_cell)?;
            let cell=if direction==Horizontal::Left {self.cell_at(*grid_id,rect.row_start,rect.column_start)?}
                else{self.cell_at(*grid_id,rect.row_end,rect.column_end)?};
            let NodeKind::Row{entries,..}=&self.node(cell)?.kind else{return Err(invalid_position())};
            return Ok(Selection::caret(Position::row(cell,if direction==Horizontal::Left{0}else{entries.len()})));
        }
        let (anchor, focus) = match selection {
            Selection::GridCells{..}=>unreachable!(),
            Selection::Node { node_id } => {
                let (row, index) = self.row_owner(*node_id)?;
                let left = Position::row(row, index); let right = Position::row(row, index + 1);
                if !extend { return Ok(Selection::caret(if direction == Horizontal::Left { left } else { right })); }
                (left, right)
            }
            Selection::Range { anchor, focus } => {
                if !extend && !selection.is_collapsed() {
                    let left = self.rank(*anchor)? < self.rank(*focus)?;
                    let target = if (direction == Horizontal::Left) == left { *anchor } else { *focus };
                    return Ok(Selection::caret(target));
                }
                (*anchor, *focus)
            }
        };
        if !extend && direction==Horizontal::Right {
            if let Some(target)=self.completed_accent_exit(focus)? {return Ok(Selection::caret(target));}
        }
        let rank = self.rank(focus)?;
        let mut candidate = rank;
        loop {
            let next = match direction {
                Horizontal::Left => candidate.checked_sub(1),
                Horizontal::Right => (candidate + 1 < self.carets.len()).then_some(candidate + 1),
            };
            let Some(next) = next else { break };
            candidate = next;
            if !self.boundary_equivalent(focus, self.carets[candidate]) { break; }
        }
        let next = self.carets[candidate];
        Ok(if extend { Selection::Range { anchor, focus: next } } else { Selection::caret(next) })
    }
    pub fn normalize_range(&self, selection: &Selection) -> EngineResult<NormalizedRange> {
        self.validate_selection(selection)?;
        if let Selection::GridCells{grid_id,anchor_cell,focus_cell}=selection {
            return Ok(NormalizedRange::GridCells{rectangle:self.grid_rectangle(*grid_id,*anchor_cell,*focus_cell)?});
        }
        if let Selection::Node { node_id } = selection {
            let (row, index) = self.row_owner(*node_id)?;
            let promoted = self.nodes.get(node_id).is_none_or(|entry| entry.parent != Some(row));
            return Ok(NormalizedRange::Row { node_id: row, start: index as u32, end: index as u32 + 1, promoted });
        }
        let Selection::Range { anchor, focus } = selection else { unreachable!() };
        let (mut anchor, mut focus) = (*anchor, *focus);
        // A text edge and its enclosing row boundary are the same location.
        // Prefer a common Text coordinate when the selection remains in that
        // run; otherwise lift only exact text endpoints to their outer row.
        if let Position::Text { node_id, .. } = anchor {
            if let Some(mapped) = self.position_in_text(focus, node_id) { focus = mapped; }
        }
        if let Position::Text { node_id, .. } = focus {
            if let Some(mapped) = self.position_in_text(anchor, node_id) { anchor = mapped; }
        }
        if anchor.node_id() != focus.node_id() {
            anchor = self.outer_text_boundary(anchor);
            focus = self.outer_text_boundary(focus);
        }
        if anchor.node_id() == focus.node_id() {
            let start = anchor.offset().min(focus.offset()) as u32;
            let end = anchor.offset().max(focus.offset()) as u32;
            return Ok(match anchor {
                Position::Row { node_id, .. } => NormalizedRange::Row { node_id, start, end, promoted: false },
                Position::Text { node_id, .. } => NormalizedRange::Text { node_id, start, end },
            });
        }
        let (left, right) = if self.rank(anchor)? <= self.rank(focus)? { (anchor, focus) } else { (focus, anchor) };
        let left_rows = self.ancestor_rows(left.node_id());
        let right_rows = self.ancestor_rows(right.node_id());
        let row = left_rows.into_iter().find(|id| right_rows.contains(id)).ok_or_else(invalid_position)?;
        let start = self.boundary_in_ancestor(left, row, false)?;
        let end = self.boundary_in_ancestor(right, row, true)?;
        Ok(NormalizedRange::Row { node_id: row, start: start as u32, end: end as u32, promoted: true })
    }
    fn position_in_text(&self, position: Position, text_id: NodeId) -> Option<Position> {
        if position.node_id() == text_id { return Some(position); }
        let entry = self.nodes.get(&text_id)?;
        let NodeKind::Text { text, .. } = &entry.node.kind else { return None };
        let index = entry.index_in_row?;
        if entry.parent != Some(position.node_id()) { return None; }
        let grapheme = if position.offset() == index { 0 }
            else if position.offset() == index + 1 { text.graphemes(true).count() as u32 }
            else { return None; };
        Some(Position::Text { node_id: text_id, grapheme, affinity: position.affinity() })
    }
    fn outer_text_boundary(&self, position: Position) -> Position {
        let Position::Text { node_id, grapheme, affinity } = position else { return position };
        let Some(entry) = self.nodes.get(&node_id) else { return position };
        let (Some(parent), Some(index)) = (entry.parent, entry.index_in_row) else { return position };
        let NodeKind::Text { text, .. } = &entry.node.kind else { return position };
        let offset = if grapheme == 0 { index }
            else if grapheme as usize == text.graphemes(true).count() { index + 1 }
            else { return position; };
        Position::Row { node_id: parent, offset: offset as u32, affinity }
    }
    fn boundary_in_ancestor(&self, position: Position, row: NodeId, at_end: bool) -> EngineResult<usize> {
        if position.node_id() == row { return Ok(position.offset()); }
        let mut id = position.node_id();
        loop {
            let entry = self.nodes.get(&id).ok_or_else(invalid_position)?;
            if entry.parent == Some(row) {
                return Ok(entry.index_in_row.ok_or_else(invalid_position)? + usize::from(at_end));
            }
            id = entry.parent.ok_or_else(invalid_position)?;
        }
    }
}

pub fn invalid_position() -> EngineError {
    EngineError::new(ErrorCode::InvalidPosition, "Mathematical position does not belong to the current tree")
}

pub fn grapheme_byte(text: &str, offset: usize) -> EngineResult<usize> {
    if offset == text.graphemes(true).count() { return Ok(text.len()); }
    text.grapheme_indices(true).nth(offset).map(|(byte, _)| byte).ok_or_else(invalid_position)
}
