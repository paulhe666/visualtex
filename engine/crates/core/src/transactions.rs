//! Mathematical tree operations. Source parsing is used only for explicit
//! PasteLatex; ordinary typing, navigation and deletion never reparse LaTeX.
mod lines;
mod matrix;
mod aligned;
mod grid_cells;
use std::collections::BTreeSet;
use std::sync::Arc;
use serde::{Deserialize, Serialize};
use unicode_segmentation::UnicodeSegmentation;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{Grouping, IdAllocator, Node, NodeKind, RowEntry, ScriptKind, ScriptSlot};
use crate::parser::{self, ParseResult};
use crate::position_map::{PositionMap, PositionStep, merge_text_map};
use crate::positions::{Horizontal, NormalizedRange, Position, Selection, TreeIndex, grapheme_byte, invalid_position};
use crate::projection::project_tree_with_symbols;
use crate::session::NodeId;
use crate::source::{MAX_SOURCE_BYTES, SourceText};
use crate::symbols::{literal_class, named_symbol};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum InputMode { Math, Text }

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StructureKind { Fraction, Binomial, SkewedFraction, Superscript, Subscript, Text, SquareRoot, IndexedRoot }

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase", deny_unknown_fields)]
pub enum EditCommand {
    InsertText { text: String, mode: InputMode },
    ReplaceRange { text: String, mode: InputMode },
    ReplaceRecentMath { expected: String, latex: String },
    InsertStructure { structure: StructureKind },
    InsertSymbol { command: String },
    InsertCustomSymbol { command: String },
    InsertFence { left: String, right: String },
    TypeDelimiter { text: String },
    InsertOperator { command: String, with_lower: bool, with_upper: bool },
    InsertGeneralOperator { with_lower: bool, with_upper: bool },
    SplitLine,
    MergeLine { direction: Horizontal },
    InsertMatrix { rows: u32, columns: u32 },
    InsertCases { rows: u32 },
    InsertAligned { rows: u32, pairs: u32 },
    InsertGridRow { after: bool },
    DeleteGridRow,
    InsertGridColumn { after: bool },
    DeleteGridColumn,
    MoveGridCell { direction: Horizontal },
    ClearGridCells,
    PasteGridCells { latex: String },
    PasteStructured { latex: String, fragment: crate::clipboard::StructuredClipboardFragment },
    SetOperatorLimits { mode: crate::operators::LimitMode },
    InsertAccent { command: String },
    InsertAnnotation { command: String },
    ExitStructure,
    DeleteBackward,
    DeleteForward,
    PasteLatex { latex: String },
    DuplicateSelection,
    SplitText,
    MergeText { direction: Horizontal },
    SetStyle { style: String },
    InsertSpace { command: String },
    InsertPhantom { command: String },
    InsertBoxed,
    SetColor { color: crate::paint::RgbColor },
}

pub struct PreparedEdit {
    pub document: Arc<ParseResult>,
    pub selection: Selection,
    pub position_map: PositionMap,
    pub changed: bool,
}

pub fn prepare(document: &Arc<ParseResult>, selection: &Selection, command: &EditCommand,
    source_id: u32, ids: &mut IdAllocator) -> EngineResult<PreparedEdit>
{
    if !document.complete {
        return Err(EngineError::new(ErrorCode::IncompleteSource, "Resolve the current source draft before a visual edit"));
    }
    if !document.supported {
        return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Unsupported source remains preserved; resolve it explicitly before a visual edit"));
    }
    let index=TreeIndex::new(&document.tree)?;
    index.validate_selection(selection)?;
    index.validate_matrix_fragment(selection)?;
    if matches!(selection,Selection::GridCells{..}) && !matches!(command,
        EditCommand::ClearGridCells|EditCommand::DeleteBackward|EditCommand::DeleteForward|
        EditCommand::PasteGridCells{..}|EditCommand::PasteLatex{..}|EditCommand::PasteStructured{..}) {
        return Err(EngineError::new(ErrorCode::InvalidPosition,"Use cell clear/paste or collapse the rectangular selection before editing"));
    }
    let mut working_ids = ids.clone();
    let mut builder = Builder { tree: Arc::clone(&document.tree), selection: selection.clone(),
        ids: &mut working_ids, overrides: Vec::new(), changed: false, symbols:Arc::clone(&document.symbol_assets) };
    match command {
        EditCommand::InsertText { text, mode } | EditCommand::ReplaceRange { text, mode } => builder.insert_text(text, *mode)?,
        EditCommand::ReplaceRecentMath { expected, latex } => builder.replace_recent_math(expected,latex,source_id)?,
        EditCommand::InsertStructure { structure } => builder.insert_structure(*structure)?,
        EditCommand::InsertSymbol { command } => builder.insert_symbol(command)?,
        EditCommand::InsertCustomSymbol { command } => builder.insert_custom_symbol(command)?,
        EditCommand::InsertFence { left, right } => builder.insert_fence(left, right)?,
        EditCommand::TypeDelimiter { text } => builder.type_delimiter(text)?,
        EditCommand::InsertOperator { command, with_lower, with_upper } => builder.insert_operator(command,*with_lower,*with_upper)?,
        EditCommand::InsertGeneralOperator { with_lower, with_upper } => builder.insert_general_operator(*with_lower,*with_upper)?,
        EditCommand::SplitLine => builder.split_line()?,
        EditCommand::MergeLine { direction } => { builder.merge_line(*direction)?; },
        EditCommand::InsertMatrix { rows, columns } => builder.insert_matrix(*rows,*columns)?,
        EditCommand::InsertCases { rows } => builder.insert_cases(*rows)?,
        EditCommand::InsertAligned { rows, pairs } => builder.insert_aligned(*rows,*pairs)?,
        EditCommand::InsertGridRow { after } => builder.insert_grid_row(*after)?,
        EditCommand::DeleteGridRow => builder.delete_grid_row()?,
        EditCommand::InsertGridColumn { after } => builder.insert_grid_column(*after)?,
        EditCommand::DeleteGridColumn => builder.delete_grid_column()?,
        EditCommand::MoveGridCell { direction } => builder.move_grid_cell(*direction)?,
        EditCommand::ClearGridCells => builder.clear_grid_cells()?,
        EditCommand::PasteGridCells { latex } => builder.paste_grid_cells(latex,source_id)?,
        EditCommand::PasteStructured { latex, fragment } => builder.paste_structured(latex,fragment,source_id)?,
        EditCommand::SetOperatorLimits { mode } => builder.set_operator_limits(*mode)?,
        EditCommand::InsertAccent { command } => builder.insert_accent(command)?,
        EditCommand::InsertAnnotation { command } => builder.insert_annotation(command)?,
        EditCommand::ExitStructure => builder.exit_structure()?,
        EditCommand::DeleteBackward => builder.delete(Horizontal::Left)?,
        EditCommand::DeleteForward => builder.delete(Horizontal::Right)?,
        EditCommand::PasteLatex { latex } => builder.paste(latex, source_id)?,
        EditCommand::DuplicateSelection => builder.duplicate()?,
        EditCommand::SplitText => builder.split_command()?,
        EditCommand::MergeText { direction } => builder.merge(*direction)?,
        EditCommand::SetStyle { style } => builder.set_style(style)?,
        EditCommand::InsertSpace { command } => builder.insert_space(command)?,
        EditCommand::InsertPhantom { command } => builder.insert_phantom(command)?,
        EditCommand::InsertBoxed => builder.insert_boxed()?,
        EditCommand::SetColor { color } => builder.set_color(*color)?,
    }
    if builder.changed {
        crate::model::validate_tree(&builder.tree)?;
        let normalized = normalize_literal_clusters(&builder.tree);
        if !Arc::ptr_eq(&normalized, &builder.tree) {
            let normalization_map = PositionMap::between(&builder.tree, &normalized, Vec::new())?;
            builder.selection = normalization_map.apply_selection(&builder.selection, &TreeIndex::new(&normalized)?)?;
            builder.overrides.extend(normalization_map.steps);
            builder.tree = normalized;
        }
    }
    let result = if builder.changed {
        if source_id <= document.source.id() {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Visual edits require a fresh source identity"));
        }
        project_tree_with_symbols(builder.tree, source_id, builder.symbols)?
    } else { Arc::clone(document) };
    let new_index = TreeIndex::new(&result.tree)?;
    new_index.validate_selection(&builder.selection)?;
    let position_map = PositionMap::between(&document.tree, &result.tree, builder.overrides)?;
    let prepared = PreparedEdit { document: result, selection: builder.selection, position_map, changed: builder.changed };
    // Speculative paste roots or other temporary nodes are not document edits.
    // A no-op must leave the monotonic identity allocator unchanged as well.
    if prepared.changed { *ids = working_ids; }
    Ok(prepared)
}

struct Builder<'a> {
    tree: Arc<Node>,
    selection: Selection,
    ids: &'a mut IdAllocator,
    overrides: Vec<PositionStep>,
    changed: bool,
    symbols:Arc<crate::custom_symbols::SymbolCatalog>,
}

impl Builder<'_> {
    fn node(&mut self, kind: NodeKind) -> EngineResult<Arc<Node>> {
        Ok(Arc::new(Node { id: self.ids.allocate()?, origin: None, kind }))
    }
    fn row(&mut self, grouping: Grouping, entries: Vec<RowEntry>) -> EngineResult<Arc<Node>> {
        self.node(NodeKind::Row { grouping, entries, trailing: String::new() })
    }
    fn normalized(&self) -> EngineResult<NormalizedRange> {
        TreeIndex::new(&self.tree)?.normalize_range(&self.selection)
    }
    fn update(&mut self, id: NodeId, f: impl FnOnce(&mut Node) -> EngineResult<()>) -> EngineResult<()> {
        self.tree = update_node(&self.tree, id, f)?;
        self.changed = true;
        Ok(())
    }
    fn insert_custom_symbol(&mut self, command:&str)->EngineResult<()> {
        let asset=self.symbols.get(command).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Custom symbol is absent from this document's frozen resource snapshot"))?;
        asset.ensure_supported()?;
        let (row,start,end)=self.ensure_row_range()?;
        let node=self.node(NodeKind::CustomSymbol{command:format!("\\{}",asset.command()),asset_sha256:asset.sha256().into(),asset,parts:Vec::new()})?;
        self.splice_row(row,start,end,vec![RowEntry{leading:String::new(),node}],true)?;
        self.selection=Selection::caret(Position::row(row,start+1));Ok(())
    }
    fn insert_symbol(&mut self, command: &str) -> EngineResult<()> {
        let (value, class) = named_symbol(command).ok_or_else(||
            EngineError::new(ErrorCode::UnsupportedSyntax, "Command is not a registered mathematical atom"))?;
        // Never turn \\partial into a derivative template or interpret an arbitrary
        // command string as LaTeX. Structural commands have their own transactions.
        let (row, start, end) = self.ensure_row_range()?;
        let node = self.node(NodeKind::Symbol { value: value.into(), class, command: Some(command.into()) })?;
        self.splice_row(row, start, end, vec![RowEntry { leading: String::new(), node }], true)?;
        self.selection = Selection::caret(Position::row(row, start + 1));
        Ok(())
    }
    fn row_entries(&self, id: NodeId) -> EngineResult<Vec<RowEntry>> {
        let node = self.tree.find(id).ok_or_else(invalid_position)?;
        match &node.kind { NodeKind::Row { entries, .. } => Ok(entries.clone()), _ => Err(invalid_position()) }
    }
    fn splice_row(&mut self, id: NodeId, start: usize, end: usize, mut replacement: Vec<RowEntry>, preserve_trivia: bool) -> EngineResult<()> {
        let old_entries = self.row_entries(id)?;
        if start > end || end > old_entries.len() { return Err(invalid_position()); }
        let mut surviving = BTreeSet::new();
        for entry in &replacement { entry.node.visit(&mut |node| { surviving.insert(node.id); }); }
        let mut removed = BTreeSet::new();
        for entry in &old_entries[start..end] {
            entry.node.visit(&mut |node| { if !surviving.contains(&node.id) { removed.insert(node.id); } });
        }
        self.update(id, move |node| {
            let NodeKind::Row { entries, trailing, .. } = &mut node.kind else { return Err(invalid_position()); };
            if start > end || end > entries.len() { return Err(invalid_position()); }
            let trivia = if preserve_trivia { entries[start..end].iter().map(|e| e.leading.as_str()).collect::<String>() } else { String::new() };
            if !trivia.is_empty() {
                if let Some(first) = replacement.first_mut() { first.leading.insert_str(0, &trivia); }
                else if let Some(next) = entries.get_mut(end) { next.leading.insert_str(0, &trivia); }
                else { trailing.insert_str(0, &trivia); }
            }
            entries.splice(start..end, replacement);
            Ok(())
        })?;
        // Include transient split nodes, not just IDs from the original tree.
        // Their source anchors may have been introduced by an earlier step in
        // this same transaction before that fragment was replaced/deleted.
        for node_id in removed {
            self.overrides.push(PositionStep::Removed { node_id, fallback: Position::row(id, start) });
        }
        Ok(())
    }
    fn replace_text(&mut self, id: NodeId, start: usize, end: usize, replacement: &str) -> EngineResult<()> {
        if replacement.len() > MAX_SOURCE_BYTES { return Err(EngineError::new(ErrorCode::ResourceLimit, "Text insertion exceeds the source budget")); }
        let old = self.tree.find(id).ok_or_else(invalid_position)?;
        let NodeKind::Text { text, .. } = &old.kind else { return Err(invalid_position()); };
        let byte_start = grapheme_byte(text, start)?;
        let byte_end = grapheme_byte(text, end)?;
        if byte_start > byte_end { return Err(invalid_position()); }
        if &text[byte_start..byte_end] == replacement {
            self.selection = Selection::caret(Position::text(id, end));
            return Ok(());
        }
        let mut new_text = text.clone();
        new_text.replace_range(byte_start..byte_end, replacement);
        let caret_byte = byte_start + replacement.len();
        let mut boundaries: Vec<_> = new_text.grapheme_indices(true).map(|(b,_)| b).collect();
        boundaries.push(new_text.len());
        let caret = match boundaries.binary_search(&caret_byte) { Ok(i) | Err(i) => i };
        self.update(id, move |node| {
            let NodeKind::Text { text, .. } = &mut node.kind else { return Err(invalid_position()); };
            *text = new_text; Ok(())
        })?;
        self.selection = Selection::caret(Position::text(id, caret));
        Ok(())
    }
    fn math_entries(&mut self, text: &str) -> EngineResult<Vec<RowEntry>> {
        if text.len() > MAX_SOURCE_BYTES { return Err(EngineError::new(ErrorCode::ResourceLimit, "Math insertion exceeds the source budget")); }
        let mut entries = Vec::new();
        for grapheme in text.graphemes(true) {
            if entries.len() >= crate::model::MAX_NODES {
                return Err(EngineError::new(ErrorCode::ResourceLimit, "Inserted math atoms exceed the node budget"));
            }
            if grapheme.chars().all(char::is_whitespace) { continue; }
            if crate::operators::spec(grapheme).is_some() {
                entries.push(RowEntry{leading:String::new(),node:self.node(NodeKind::LargeOperator{command:grapheme.into(),parts:Vec::new()})?});
                continue;
            }
            if grapheme.chars().any(char::is_control) || matches!(grapheme, "\\" | "^" | "_" | "~") {
                return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Use a structural command or explicit LaTeX paste for this math input"));
            }
            let escaped = match grapheme { "{" => Some("\\{"), "}" => Some("\\}"), "%" => Some("\\%"),
                "&" => Some("\\&"), "#" => Some("\\#"), "$" => Some("\\$"), _ => None };
            let (value, class, command) = if let Some(command) = escaped {
                let (value, class) = named_symbol(command).unwrap(); (value.to_owned(), class, Some(command.into()))
            } else {
                if grapheme == "'" { return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Prime notation is outside the K0 structural capability")); }
                (grapheme.to_owned(), literal_class(grapheme), None)
            };
            entries.push(RowEntry { leading: String::new(), node: self.node(NodeKind::Symbol { value, class, command })? });
        }
        Ok(entries)
    }
    fn insert_text(&mut self, text: &str, mode: InputMode) -> EngineResult<()> {
        let range = self.normalized()?;
        if let NormalizedRange::Text { node_id, start, end } = range {
            return self.replace_text(node_id, start as usize, end as usize, text);
        }
        let NormalizedRange::Row { node_id, start, end, .. } = range else { unreachable!() };
        if text.is_empty() {
            if start != end { self.splice_row(node_id, start as usize, end as usize, Vec::new(), true)?; }
            self.selection = Selection::caret(Position::row(node_id, start as usize));
            return Ok(());
        }
        if mode == InputMode::Text && start == end {
            let entries = self.row_entries(node_id)?;
            if start > 0 {
                if let NodeKind::Text { text: previous, .. } = &entries[start as usize - 1].node.kind {
                    return self.replace_text(entries[start as usize - 1].node.id, previous.graphemes(true).count(), previous.graphemes(true).count(), text);
                }
            }
            if let Some(entry) = entries.get(start as usize) {
                if matches!(entry.node.kind, NodeKind::Text { .. }) { return self.replace_text(entry.node.id, 0, 0, text); }
            }
        }
        let entries = match mode {
            InputMode::Math => self.math_entries(text)?,
            InputMode::Text => vec![RowEntry { leading: String::new(), node: self.node(NodeKind::Text {
                command: "\\text".into(), before_body: String::new(), text: text.into() })? }],
        };
        if entries.is_empty() && start == end { return Ok(()); }
        let count = entries.len();
        let text_target = if mode == InputMode::Text { Some(entries[0].node.id) } else { None };
        self.splice_row(node_id, start as usize, end as usize, entries, true)?;
        self.selection = Selection::caret(match text_target {
            Some(id) => Position::text(id, text.graphemes(true).count()),
            None => Position::row(node_id, start as usize + count),
        });
        Ok(())
    }
    fn split_text(&mut self, id: NodeId, at: usize) -> EngineResult<(NodeId, NodeId, usize)> {
        let index = TreeIndex::new(&self.tree)?;
        let (row, offset) = index.row_owner(id)?;
        let old = index.node(id)?;
        let NodeKind::Text { command, before_body, text } = &old.kind else { return Err(invalid_position()); };
        let byte = grapheme_byte(text, at)?;
        let mut left = old.clone(); left.origin = None;
        if let NodeKind::Text { text: value, .. } = &mut left.kind { *value = text[..byte].into(); }
        let right_kind = NodeKind::Text { command: command.clone(), before_body: String::new(), text: text[byte..].into() };
        let _ = before_body;
        let right = self.node(right_kind)?;
        let right_id = right.id;
        self.splice_row(row, offset, offset + 1, vec![
            RowEntry { leading: String::new(), node: Arc::new(left) }, RowEntry { leading: String::new(), node: right },
        ], true)?;
        self.overrides.push(PositionStep::TextSplit { node_id: id, right_id, at: at as u32 });
        self.selection = Selection::caret(Position::text(right_id, 0));
        Ok((row, right_id, offset + 1))
    }
    /// Structural insertion inside text first splits only that TextRun. The
    /// unselected prefix/suffix remain editable text with explicit position maps.
    fn ensure_row_range(&mut self) -> EngineResult<(NodeId, usize, usize)> {
        match self.normalized()? {
            NormalizedRange::GridCells{..}=>Err(invalid_position()),
            NormalizedRange::Row { node_id, start, end, .. } => Ok((node_id, start as usize, end as usize)),
            NormalizedRange::Text { node_id, start, end } => {
                let index = TreeIndex::new(&self.tree)?;
                let (row, offset) = index.row_owner(node_id)?;
                let NodeKind::Text { text, .. } = &index.node(node_id)?.kind else { return Err(invalid_position()); };
                let len = text.graphemes(true).count();
                let (start, end) = (start as usize, end as usize);
                if start == end {
                    if start == 0 { return Ok((row, offset, offset)); }
                    if start == len { return Ok((row, offset + 1, offset + 1)); }
                    self.split_text(node_id, start)?;
                    return Ok((row, offset + 1, offset + 1));
                }
                if end < len { self.split_text(node_id, end)?; }
                if start > 0 {
                    self.split_text(node_id, start)?;
                    Ok((row, offset + 1, offset + 2))
                } else { Ok((row, offset, offset + 1)) }
            }
        }
    }
    fn insert_operator(&mut self, command:&str, with_lower:bool, with_upper:bool)->EngineResult<()> {
        use crate::operators::OperatorPart;
        if crate::operators::spec(command).is_none(){return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"Unsupported large operator"));}
        let (row,start,end)=self.ensure_row_range()?;let entries=self.row_entries(row)?;
        let mut parts=Vec::new();let mut focus=None;
        for (enabled,script) in [(with_lower,ScriptKind::Subscript),(with_upper,ScriptKind::Superscript)] {
            if enabled {
                let argument=self.row(Grouping::Braced,Vec::new())?;
                focus=focus.or(Some(Position::row(argument.id,0)));
                parts.push(OperatorPart::Script{slot:ScriptSlot{script,before_marker:String::new(),before_argument:String::new(),argument}});
            }
        }
        let operator=self.node(NodeKind::LargeOperator{command:command.into(),parts})?;
        let mut replacement=vec![RowEntry{leading:String::new(),node:operator}];
        // The selected expression is an operand, not an upper/lower bound.
        replacement.extend_from_slice(&entries[start..end]);
        self.splice_row(row,start,end,replacement,false)?;
        self.selection=Selection::caret(focus.unwrap_or(Position::row(row,start+1)));Ok(())
    }
    fn insert_general_operator(&mut self,with_lower:bool,with_upper:bool)->EngineResult<()> {
        use crate::operators::OperatorPart;
        let (row,start,end)=self.ensure_row_range()?;let entries=self.row_entries(row)?;
        let body=self.row(Grouping::Braced,entries[start..end].to_vec())?;
        let body_end=Position::row(body.id,end-start);let mut focus=None;let mut parts=Vec::new();
        for (enabled,script) in [(with_lower,ScriptKind::Subscript),(with_upper,ScriptKind::Superscript)] {
            if enabled {
                let argument=self.row(Grouping::Braced,Vec::new())?;focus=focus.or(Some(Position::row(argument.id,0)));
                parts.push(OperatorPart::Script{slot:ScriptSlot{script,before_marker:String::new(),before_argument:String::new(),argument}});
            }
        }
        let operator=self.node(NodeKind::GeneralOperator{command:"\\mathop".into(),before_body:String::new(),body,parts})?;
        self.splice_row(row,start,end,vec![RowEntry{leading:String::new(),node:operator}],false)?;
        self.selection=Selection::caret(focus.unwrap_or(body_end));Ok(())
    }
    fn active_operator(&self)->EngineResult<NodeId>{
        let index=TreeIndex::new(&self.tree)?;
        let mut current=match &self.selection {
            Selection::Node{node_id}=>*node_id,
            Selection::Range{anchor,focus} if anchor.same_location(*focus)=>{
                if let Position::Row{node_id,offset,..}=focus {
                    if let NodeKind::Row{entries,..}=&index.node(*node_id)?.kind {
                        if let Some(entry)=offset.checked_sub(1).and_then(|i|entries.get(i as usize)) {
                            if crate::operators::parts(&entry.node).is_some(){return Ok(entry.node.id);}
                        }
                    }
                }
                focus.node_id()
            }
            _=>return Err(EngineError::new(ErrorCode::InvalidPosition,"Select an operator or place the caret in its bounds")),
        };
        loop {
            if crate::operators::parts(index.node(current)?).is_some(){return Ok(current);}
            current=index.nodes.get(&current).and_then(|e|e.parent).ok_or_else(invalid_position)?;
        }
    }
    fn set_operator_limits(&mut self,mode:crate::operators::LimitMode)->EngineResult<()> {
        use crate::operators::{self,OperatorPart,LimitMode};
        let target=self.active_operator()?;
        let parts=operators::parts(self.tree.find(target).ok_or_else(invalid_position)?).ok_or_else(invalid_position)?;
        if operators::effective_mode(parts)==mode{return Ok(());}
        let mut parts=parts.to_vec();let mut leftover=String::new();
        if mode==LimitMode::Auto {
            let mut next=Vec::new();
            for part in parts {
                match part {
                    OperatorPart::Limits{before,..}=>leftover.push_str(&before),
                    OperatorPart::Script{mut slot}=>{
                        slot.before_marker.insert_str(0,&leftover);leftover.clear();next.push(OperatorPart::Script{slot});
                    }
                }
            }
            parts=next;
        } else {
            if parts.len()>=operators::MAX_OPERATOR_PARTS{return Err(EngineError::new(ErrorCode::ResourceLimit,"Operator directive budget exhausted"));}
            parts.push(OperatorPart::Limits{before:String::new(),command:mode.command().unwrap().into()});
        }
        let owner=if leftover.is_empty(){None}else{Some(TreeIndex::new(&self.tree)?.row_owner(target)?)};
        self.update(target,move |node|{*operators::parts_mut(node).ok_or_else(invalid_position)?=parts;Ok(())})?;
        if let Some((row,offset))=owner {
            self.update(row,move |node|{let NodeKind::Row{entries,trailing,..}=&mut node.kind else{return Err(invalid_position());};
                if let Some(next)=entries.get_mut(offset+1){next.leading.insert_str(0,&leftover);}else{trailing.insert_str(0,&leftover);}Ok(())})?;
        }
        Ok(())
    }
    fn insert_space(&mut self, command: &str) -> EngineResult<()> {
        let width_mu=crate::spacing::space_mu(command).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Unsupported explicit math space"))?;
        let (row,start,end)=self.ensure_row_range()?;
        let node=self.node(NodeKind::Space{command:command.into(),width_mu})?;
        self.splice_row(row,start,end,vec![RowEntry{leading:String::new(),node}],true)?;
        self.selection=Selection::caret(Position::row(row,start+1));Ok(())
    }
    fn insert_phantom(&mut self, command: &str) -> EngineResult<()> {
        let axis=crate::spacing::phantom_axis(command).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Unsupported phantom command"))?;
        let (row,start,end)=self.ensure_row_range()?;let entries=self.row_entries(row)?;
        let body=self.row(Grouping::Braced,entries[start..end].to_vec())?;
        let caret=Position::row(body.id,end-start);
        let node=self.node(NodeKind::Phantom{command:command.into(),before_body:String::new(),axis,body})?;
        self.splice_row(row,start,end,vec![RowEntry{leading:String::new(),node}],false)?;
        self.selection=Selection::caret(caret);Ok(())
    }
    fn set_color(&mut self, color:crate::paint::RgbColor) -> EngineResult<()> {
        let (row,start,end)=self.ensure_row_range()?;let entries=self.row_entries(row)?;
        let body=self.row(Grouping::Braced,entries[start..end].to_vec())?;
        let caret=Position::row(body.id,end-start);
        let node=self.node(NodeKind::Color{spec:crate::paint::ColorSpec::html(color),before_body:String::new(),body})?;
        self.splice_row(row,start,end,vec![RowEntry{leading:String::new(),node}],false)?;
        self.selection=Selection::caret(caret);Ok(())
    }
    fn insert_boxed(&mut self) -> EngineResult<()> {
        let (row,start,end)=self.ensure_row_range()?;let entries=self.row_entries(row)?;
        let body=self.row(Grouping::Braced,entries[start..end].to_vec())?;
        let caret=Position::row(body.id,end-start);
        let node=self.node(NodeKind::Boxed{before_body:String::new(),body})?;
        self.splice_row(row,start,end,vec![RowEntry{leading:String::new(),node}],false)?;
        self.selection=Selection::caret(caret);Ok(())
    }
    fn insert_accent(&mut self, command: &str) -> EngineResult<()> {
        if crate::accents::spec(command).is_none() { return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Unsupported accent command")); }
        let (row, start, end) = self.ensure_row_range()?;
        let entries = self.row_entries(row)?;
        let body = self.row(Grouping::Braced, entries[start..end].to_vec())?;
        let caret = Position::row(body.id, end-start);
        let node = self.node(NodeKind::Accent { command: command.into(), before_body: String::new(), body })?;
        self.splice_row(row, start, end, vec![RowEntry { leading: String::new(), node }], false)?;
        self.selection = Selection::caret(caret);Ok(())
    }
    fn insert_annotation(&mut self, command: &str) -> EngineResult<()> {
        let above = crate::accents::annotation_above(command).ok_or_else(|| EngineError::new(ErrorCode::UnsupportedSyntax, "Unsupported annotation command"))?;
        let (row, start, end) = self.ensure_row_range()?;let entries = self.row_entries(row)?;
        let base = self.row(Grouping::Braced, entries[start..end].to_vec())?;
        let annotation = self.row(Grouping::Braced, Vec::new())?;
        let caret = Position::row(annotation.id, 0);
        let node = self.node(NodeKind::Annotation { command: command.into(), above, before_annotation: String::new(),
            annotation, before_base: String::new(), base })?;
        self.splice_row(row, start, end, vec![RowEntry { leading: String::new(), node }], false)?;
        self.selection = Selection::caret(caret);Ok(())
    }
    fn set_style(&mut self, command: &str) -> EngineResult<()> {
        let variant=crate::styles::command_variant(command)
            .ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Unsupported mathematical style command"))?;
        let (row,start,end)=self.ensure_row_range()?;let entries=self.row_entries(row)?;
        let body=self.row(Grouping::Braced,entries[start..end].to_vec())?;
        let caret=Position::row(body.id,end-start);
        let node=self.node(NodeKind::Style {command:command.into(),before_body:String::new(),variant,body})?;
        self.splice_row(row,start,end,vec![RowEntry{leading:String::new(),node}],false)?;
        self.selection=Selection::caret(caret);Ok(())
    }
    fn exit_structure(&mut self) -> EngineResult<()> {
        let Selection::Range { anchor, focus } = self.selection else { return Ok(()); };
        if !anchor.same_location(focus) { return Ok(()); }
        let index = TreeIndex::new(&self.tree)?;let node = index.node(focus.node_id())?;
        let end = match (&node.kind, focus) {
            (NodeKind::Row { entries, .. }, Position::Row { offset, .. }) => offset as usize == entries.len(),
            (NodeKind::Text { text, .. }, Position::Text { grapheme, .. }) => grapheme as usize == text.graphemes(true).count(),
            _ => false,
        };
        if !end { return Ok(()); }
        let mut current = node.id;let mut target = None;
        while let Some(parent) = index.nodes.get(&current).and_then(|n| n.parent) {
            let n = index.node(parent)?;
            match &n.kind {
                NodeKind::Row { entries, .. } => {
                    if entries.last().is_none_or(|e| e.node.id != current) { break; }
                    current = parent;
                }
                NodeKind::Accent { body, .. } if body.id == current => {
                    let (row,offset) = index.row_owner(parent)?;target = Some(Position::row(row,offset+1));current=parent;
                }
                _ => {
                    if target.is_none() { let (row,offset)=index.row_owner(parent)?;target=Some(Position::row(row,offset+1)); }
                    break;
                }
            }
        }
        if let Some(target) = target { self.selection = Selection::caret(target); }
        Ok(())
    }

    fn insert_fence(&mut self, left: &str, right: &str) -> EngineResult<()> {
        if crate::delimiters::glyph(left).is_none() || crate::delimiters::glyph(right).is_none() {
            return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Unsupported fence delimiter"));
        }
        let (row, start, end) = self.ensure_row_range()?;
        let entries = self.row_entries(row)?;
        let body = self.row(Grouping::Bare, entries[start..end].to_vec())?;
        let focus = Position::row(body.id, end - start);
        let node = self.node(NodeKind::Fence { before_left: String::new(), left: left.into(), body,
            before_right: String::new(), right: right.into() })?;
        self.splice_row(row, start, end, vec![RowEntry { leading: String::new(), node }], false)?;
        self.selection = Selection::caret(focus);Ok(())
    }

    fn type_delimiter(&mut self, text: &str) -> EngineResult<()> {
        if let Some((left, right)) = crate::delimiters::matching_pair(text) { return self.insert_fence(left, right); }
        if !matches!(text, ")" | "]" | "}") {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Expected a typed bracket delimiter"));
        }
        let close = crate::delimiters::glyph(crate::delimiters::literal_close_source(text));
        let index = TreeIndex::new(&self.tree)?;
        if let Selection::Range { anchor, focus } = &self.selection {
            if anchor.same_location(*focus) {
                let node = index.node(focus.node_id())?;
                let at_end = match (&node.kind, focus) {
                    (NodeKind::Row { entries, .. }, Position::Row { offset, .. }) => *offset as usize == entries.len(),
                    (NodeKind::Text { text, .. }, Position::Text { grapheme, .. }) => *grapheme as usize == text.graphemes(true).count(),
                    _ => false,
                };
                if at_end {
                    let mut child = node.id;
                    while let Some(parent) = index.nodes.get(&child).and_then(|e| e.parent) {
                        let parent_node = index.node(parent)?;
                        if let NodeKind::Fence { body, right, .. } = &parent_node.kind {
                            if body.id == child && close == crate::delimiters::glyph(right) {
                                let (row, offset) = index.row_owner(parent)?;
                                self.selection = Selection::caret(Position::row(row, offset + 1));
                                return Ok(());
                            }
                            // A mismatch must not jump through an inner pair
                            // merely because an outer pair happens to match.
                            break;
                        }
                        if parent_node.children().last().is_none_or(|last| last.id != child) { break; }
                        child = parent;
                    }
                }
            }
        }
        // Keep unmatched closers as literal mathematical content. Never delete
        // an existing closer or trigger a second speculative edit.
        self.insert_text(text, InputMode::Math)
    }

    fn insert_structure(&mut self, structure: StructureKind) -> EngineResult<()> {
        let (row, mut start, end) = self.ensure_row_range()?;
        let mut entries = self.row_entries(row)?;
        if matches!(structure, StructureKind::Superscript | StructureKind::Subscript) && start == end && start > 0 {
            let script = if structure == StructureKind::Superscript { ScriptKind::Superscript } else { ScriptKind::Subscript };
            if let Some(parts) = crate::operators::parts(&entries[start-1].node) {
                if let Some(slot)=crate::operators::slots(parts).find(|slot|slot.script==script) {
                    self.selection=Selection::caret(Position::row(slot.argument.id,0));return Ok(());
                }
                let target=entries[start-1].node.id;let argument=self.row(Grouping::Braced,Vec::new())?;let arg_id=argument.id;
                self.update(target,move |node|{let parts=crate::operators::parts_mut(node).ok_or_else(invalid_position)?;
                    parts.push(crate::operators::OperatorPart::Script{slot:ScriptSlot{script,before_marker:String::new(),before_argument:String::new(),argument}});Ok(())})?;
                self.selection=Selection::caret(Position::row(arg_id,0));return Ok(());
            }
            if let NodeKind::Scripts { slots, .. } = &entries[start - 1].node.kind {
                if let Some(slot) = slots.iter().find(|slot| slot.script == script) {
                    self.selection = Selection::caret(Position::row(slot.argument.id, 0)); return Ok(());
                }
                let target = entries[start - 1].node.id;
                let argument = self.row(Grouping::Braced, Vec::new())?;
                let arg_id = argument.id;
                self.update(target, move |node| {
                    let NodeKind::Scripts { slots, .. } = &mut node.kind else { return Err(invalid_position()); };
                    slots.push(ScriptSlot { script, before_marker: String::new(), before_argument: String::new(), argument }); Ok(())
                })?;
                self.selection = Selection::caret(Position::row(arg_id, 0)); return Ok(());
            }
            start -= 1;
        }
        let selected = entries.drain(start..end).collect::<Vec<_>>();
        let (node, focus) = match structure {
            StructureKind::Fraction | StructureKind::Binomial | StructureKind::SkewedFraction => {
                let command = match structure {
                    StructureKind::Binomial => "\\binom",
                    StructureKind::SkewedFraction => "\\nicefrac",
                    _ => "\\frac",
                };
                let numerator = self.row(Grouping::Braced, selected)?;
                let denominator = self.row(Grouping::Braced, Vec::new())?;
                let focus = if numerator.is_empty_row() { numerator.id } else { denominator.id };
                let node = self.node(NodeKind::Fraction { command: command.into(), before_numerator: String::new(),
                    numerator, before_denominator: String::new(), denominator })?;
                (node, Position::row(focus, 0))
            }
            StructureKind::SquareRoot | StructureKind::IndexedRoot => {
                let radicand = self.row(Grouping::Braced, selected)?;
                let degree = if structure == StructureKind::IndexedRoot { Some(self.row(Grouping::Bare, Vec::new())?) } else { None };
                let focus = degree.as_ref().map_or(radicand.id, |d| d.id);
                let node = self.node(NodeKind::Radical { command: "\\sqrt".into(), before_degree: String::new(),
                    degree, before_radicand: String::new(), radicand })?;
                (node, Position::row(focus, 0))
            }
            StructureKind::Superscript | StructureKind::Subscript => {
                let grouping = if selected.len() == 1 && !matches!(selected[0].node.kind, NodeKind::Scripts { .. } | NodeKind::LargeOperator { .. } | NodeKind::GeneralOperator { .. } | NodeKind::CustomSymbol { .. }) { Grouping::Bare } else { Grouping::Braced };
                let base = self.row(grouping, selected)?;
                let argument = self.row(Grouping::Braced, Vec::new())?;
                let focus = if base.is_empty_row() { base.id } else { argument.id };
                let script = if structure == StructureKind::Superscript { ScriptKind::Superscript } else { ScriptKind::Subscript };
                let node = self.node(NodeKind::Scripts { base, slots: vec![ScriptSlot { script, before_marker: String::new(),
                    before_argument: String::new(), argument }] })?;
                (node, Position::row(focus, 0))
            }
            StructureKind::Text => {
                let mut text = String::new();
                for entry in selected { text.push_str(&crate::model::latex(&entry.node, crate::model::LatexMode::Lossless)); }
                let node = self.node(NodeKind::Text { command: "\\text".into(), before_body: String::new(), text })?;
                let focus = Position::text(node.id, 0); (node, focus)
            }
        };
        self.splice_row(row, start, end, vec![RowEntry { leading: String::new(), node }], false)?;
        self.selection = Selection::caret(focus); Ok(())
    }
    fn delete(&mut self, direction: Horizontal) -> EngineResult<()> {
        let range = self.normalized()?;
        match range {
            NormalizedRange::GridCells{..}=>self.clear_grid_cells(),
            NormalizedRange::Text { node_id, start, end } => {
                if start != end { return self.replace_text(node_id, start as usize, end as usize, ""); }
                let NodeKind::Text { text, .. } = &self.tree.find(node_id).unwrap().kind else { unreachable!() };
                let count = text.graphemes(true).count() as u32;
                if direction == Horizontal::Left && start > 0 { return self.replace_text(node_id, start as usize - 1, start as usize, ""); }
                if direction == Horizontal::Right && start < count { return self.replace_text(node_id, start as usize, start as usize + 1, ""); }
                self.selection = TreeIndex::new(&self.tree)?.move_horizontal(&self.selection, direction, false)?;
                Ok(())
            }
            NormalizedRange::Row { node_id, start, end, .. } => {
                if start != end {
                    self.splice_row(node_id, start as usize, end as usize, Vec::new(), true)?;
                    self.selection = Selection::caret(Position::row(node_id, start as usize)); return Ok(());
                }
                let entries = self.row_entries(node_id)?;
                let offset = if direction == Horizontal::Left { start.checked_sub(1).map(|v| v as usize) }
                    else { ((start as usize) < entries.len()).then_some(start as usize) };
                let Some(offset) = offset else {
                    if self.merge_line(direction)? { return Ok(()); }
                    self.selection = TreeIndex::new(&self.tree)?.move_horizontal(&self.selection, direction, false)?;
                    return Ok(());
                };
                let target = &entries[offset].node;
                match &target.kind {
                    NodeKind::Symbol { .. } | NodeKind::Space { .. } | NodeKind::ColorChange { .. } | NodeKind::EmptySlot { .. } => {
                        self.splice_row(node_id, offset, offset+1, Vec::new(), true)?;
                        self.selection = Selection::caret(Position::row(node_id, offset));
                    }
                    NodeKind::Text { text, .. } => {
                        let count = text.graphemes(true).count();
                        if count == 0 {
                            self.splice_row(node_id, offset, offset+1, Vec::new(), true)?;
                            self.selection = Selection::caret(Position::row(node_id, offset));
                        } else if direction == Horizontal::Left { self.replace_text(target.id, count-1, count, "")?; }
                        else { self.replace_text(target.id, 0, 1, "")?; }
                    }
                    _ => { self.selection = Selection::Node { node_id: target.id }; }
                }
                Ok(())
            }
        }
    }
    fn replace_recent_math(&mut self,expected:&str,latex:&str,source_id:u32)->EngineResult<()> {
        let expected_graphemes:Vec<&str>=expected.graphemes(true).collect();
        if expected_graphemes.is_empty()||expected_graphemes.len()>64||latex.is_empty()
            ||expected_graphemes.iter().any(|g|g.chars().any(|ch|!ch.is_ascii()||ch.is_control())) {
            return Err(EngineError::new(ErrorCode::InvalidInput,"Inline shortcut replacement exceeds its bounded input contract"));
        }
        let count=expected_graphemes.len();
        let NormalizedRange::Row{node_id,start,end,..}=self.normalized()? else {
            return Err(EngineError::new(ErrorCode::InvalidPosition,"Inline shortcut replacement requires a mathematical row caret"));
        };
        if start!=end||(start as usize)<count{return Err(EngineError::new(ErrorCode::InvalidPosition,"Inline shortcut replacement target is stale"));}
        let entries=self.row_entries(node_id)?;let from=start as usize-count;let to=start as usize;
        for (entry,wanted) in entries[from..to].iter().zip(expected_graphemes) {
            if !entry.leading.is_empty(){return Err(EngineError::new(ErrorCode::Conflict,"Inline shortcut replacement cannot cross source trivia"));}
            let NodeKind::Symbol{value,command:None,..}=&entry.node.kind else {
                return Err(EngineError::new(ErrorCode::Conflict,"Inline shortcut replacement can only consume recent literal math input"));
            };
            if value!=wanted {
                return Err(EngineError::new(ErrorCode::Conflict,"Inline shortcut replacement target no longer matches the expected literal input"));
            }
        }
        self.selection=Selection::Range{anchor:Position::row(node_id,from),focus:Position::row(node_id,to)};
        self.paste(latex,source_id)
    }
    fn paste_structured(&mut self,latex:&str,fragment:&crate::clipboard::StructuredClipboardFragment,source_id:u32)->EngineResult<()> {
        self.symbols=crate::clipboard::validate_structural_fragment_with_symbols(latex,fragment,Arc::clone(&self.symbols))?;
        self.paste(latex,source_id)
    }
    fn paste(&mut self, latex: &str, source_id: u32) -> EngineResult<()> {
        if matches!(self.selection,Selection::GridCells{..}) {return self.paste_grid_cells(latex,source_id);}
        let root_id = self.ids.allocate()?;
        let parsed = parser::parse_with_symbols(SourceText::new(source_id, latex)?, root_id, self.ids, Arc::clone(&self.symbols))?;
        if !parsed.complete { return Err(EngineError::new(ErrorCode::IncompleteSource, "Pasted source is incomplete; original document is unchanged")); }
        if !parsed.supported { return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Pasted source exceeds the current visual capability; original document is unchanged")); }
        let NodeKind::Row { entries, trailing, .. } = &parsed.tree.kind else { unreachable!() };
        let mut replacement = entries.clone();
        // Preserve final pasted trivia by attaching it to the insertion's next
        // neighbor (or row tail); do not silently discard source comments.
        let (row, start, end) = self.ensure_row_range()?;
        let count = replacement.len();
        if count == 0 && start == end && trailing.is_empty() { return Ok(()); }
        self.splice_row(row, start, end, std::mem::take(&mut replacement), true)?;
        if !trailing.is_empty() {
            let trailing = trailing.clone();
            self.update(row, move |node| {
                let NodeKind::Row { entries, trailing: tail, .. } = &mut node.kind else { return Err(invalid_position()); };
                if let Some(next) = entries.get_mut(start+count) { next.leading.insert_str(0, &trailing); }
                else { tail.insert_str(0, &trailing); }
                Ok(())
            })?;
        }
        self.selection = Selection::caret(Position::row(row, start+count)); Ok(())
    }
    fn duplicate(&mut self) -> EngineResult<()> {
        match self.normalized()? {
            NormalizedRange::GridCells{..}=>Err(invalid_position()),
            NormalizedRange::Text { node_id, start, end } => {
                if start == end { return Ok(()); }
                let NodeKind::Text { text, .. } = &self.tree.find(node_id).unwrap().kind else { unreachable!() };
                let copy = text[grapheme_byte(text, start as usize)?..grapheme_byte(text, end as usize)?].to_owned();
                self.replace_text(node_id, end as usize, end as usize, &copy)
            }
            NormalizedRange::Row { node_id, start, end, .. } => {
                if start == end { return Ok(()); }
                let entries = self.row_entries(node_id)?;
                let mut copies = Vec::new();
                for entry in &entries[start as usize..end as usize] {
                    copies.push(RowEntry { leading: entry.leading.clone(), node: clone_with_fresh_ids(&entry.node, self.ids)? });
                }
                self.splice_row(node_id, end as usize, end as usize, copies, true)?;
                self.selection = Selection::Range { anchor: Position::row(node_id, end as usize),
                    focus: Position::row(node_id, (end + end - start) as usize) };
                Ok(())
            }
        }
    }
    fn split_command(&mut self) -> EngineResult<()> {
        let NormalizedRange::Text { node_id, start, end } = self.normalized()? else {
            return Err(EngineError::new(ErrorCode::InvalidPosition, "Text split requires a text caret")); };
        if start != end { return Err(EngineError::new(ErrorCode::InvalidPosition, "Text split requires a collapsed caret")); }
        self.split_text(node_id, start as usize).map(|_| ())
    }
    fn merge(&mut self, direction: Horizontal) -> EngineResult<()> {
        let range = self.normalized()?;
        let (row, boundary) = match range {
            NormalizedRange::Row { node_id, start, end, .. } if start == end => (node_id, start as usize),
            NormalizedRange::Text { node_id, start, end } if start == end => {
                let index = TreeIndex::new(&self.tree)?;
                let (row, offset) = index.row_owner(node_id)?;
                let NodeKind::Text { text, .. } = &index.node(node_id)?.kind else { return Err(invalid_position()); };
                let at_boundary = if direction == Horizontal::Left { start == 0 }
                    else { start as usize == text.graphemes(true).count() };
                if !at_boundary {
                    return Err(EngineError::new(ErrorCode::InvalidPosition, "Text merge requires the matching start or end boundary"));
                }
                (row, if direction == Horizontal::Left { offset } else { offset+1 })
            }
            _ => return Err(EngineError::new(ErrorCode::InvalidPosition, "Text merge requires a caret next to adjacent text runs")),
        };
        let entries = self.row_entries(row)?;
        if boundary == 0 || boundary >= entries.len() { return Err(invalid_position()); }
        let left = &entries[boundary-1].node; let right = &entries[boundary].node;
        let (NodeKind::Text { text: a, command: ac, .. }, NodeKind::Text { text: b, command: bc, .. }) = (&left.kind, &right.kind)
            else { return Err(EngineError::new(ErrorCode::InvalidPosition, "Only adjacent matching text runs can merge")); };
        if ac != bc { return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Text styles differ")); }
        let merged = format!("{a}{b}");
        let left_map = merge_text_map(left.id, left.id, a, &merged, 0)?;
        let right_map = merge_text_map(right.id, left.id, b, &merged, a.len())?;
        let focus = Position::text(left.id, a.graphemes(true).count());
        let mut combined = left.as_ref().clone(); combined.origin = None;
        if let NodeKind::Text { text, .. } = &mut combined.kind { *text = merged; }
        // Preserve text-relative anchors before the splice records a fallback
        // for the removed right node; otherwise its positions collapse too early.
        self.overrides.extend([left_map, right_map]);
        self.splice_row(row, boundary-1, boundary+1, vec![RowEntry { leading: String::new(), node: Arc::new(combined) }], true)?;
        self.selection = Selection::caret(focus); Ok(())
    }
}

/// A Unicode grapheme can span two separately delivered input events. Only
/// literal adjacent math atoms with no source trivia between them may join.
/// Untouched subtrees are never normalized speculatively.
fn normalize_literal_clusters(node: &Arc<Node>) -> Arc<Node> {
    if node.origin.is_some() { return Arc::clone(node); }
    let mut result = node.as_ref().clone();
    let mut changed = false;
    let mut normalize_child = |child: &mut Arc<Node>| {
        let next = normalize_literal_clusters(child);
        if !Arc::ptr_eq(child, &next) { *child = next; changed = true; }
    };
    match &mut result.kind {
        NodeKind::Row { entries, .. } => {
            for entry in entries.iter_mut() { normalize_child(&mut entry.node); }
            let mut index = 0;
            while index + 1 < entries.len() {
                let combined = match (&entries[index].node.kind, &entries[index+1].node.kind) {
                    (NodeKind::Symbol { value: a, command: None, .. }, NodeKind::Symbol { value: b, command: None, .. })
                        if entries[index+1].leading.is_empty() => {
                            let text = format!("{a}{b}");
                            (text.graphemes(true).count() == 1).then_some(text)
                        }
                    _ => None,
                };
                if let Some(text) = combined {
                    let left = Arc::make_mut(&mut entries[index].node);
                    left.origin = None;
                    left.kind = NodeKind::Symbol { class: literal_class(&text), value: text, command: None };
                    entries.remove(index+1); changed = true;
                } else { index += 1; }
            }
        }
        NodeKind::Grid { rows, .. } => for row in rows { normalize_child(row); },
        NodeKind::GridRow { cells, .. } => for cell in cells { normalize_child(cell); },
        NodeKind::Fraction { numerator, denominator, .. } => { normalize_child(numerator); normalize_child(denominator); }
        NodeKind::Fence { body, .. } | NodeKind::Accent { body, .. } | NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => normalize_child(body),
        NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => for slot in crate::operators::slots_mut(parts){normalize_child(&mut slot.argument);},
        NodeKind::GeneralOperator { body, parts, .. } => {
            normalize_child(body);for slot in crate::operators::slots_mut(parts){normalize_child(&mut slot.argument);}
        }
        NodeKind::MacroCall { arguments, .. } => for argument in arguments {normalize_child(argument);},
        NodeKind::Annotation { annotation, base, .. } => { normalize_child(annotation); normalize_child(base); }
        NodeKind::Radical { degree, radicand, .. } => {
            if let Some(degree) = degree { normalize_child(degree); }
            normalize_child(radicand);
        }
        NodeKind::Scripts { base, slots } => { normalize_child(base); for slot in slots { normalize_child(&mut slot.argument); } }
        _ => {}
    }
    if changed { Arc::new(result) } else { Arc::clone(node) }
}

pub fn clone_with_fresh_ids(node: &Arc<Node>, ids: &mut IdAllocator) -> EngineResult<Arc<Node>> {
    let mut result = node.as_ref().clone();
    result.id = ids.allocate()?;
    match &mut result.kind {
        NodeKind::Row { entries, .. } => for entry in entries { entry.node = clone_with_fresh_ids(&entry.node, ids)?; },
        NodeKind::Grid { columns, rows, .. } => {
            for column in columns { *column=crate::grids::GridColumnId::allocate(ids)?; }
            for row in rows { *row=clone_with_fresh_ids(row,ids)?; }
        }
        NodeKind::GridRow { cells, .. } => for cell in cells { *cell=clone_with_fresh_ids(cell,ids)?; },
        NodeKind::Fraction { numerator, denominator, .. } => {
            *numerator = clone_with_fresh_ids(numerator, ids)?; *denominator = clone_with_fresh_ids(denominator, ids)?;
        }
        NodeKind::Fence { body, .. } | NodeKind::Accent { body, .. } | NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => *body = clone_with_fresh_ids(body, ids)?,
        NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => for slot in crate::operators::slots_mut(parts){slot.argument=clone_with_fresh_ids(&slot.argument,ids)?;},
        NodeKind::GeneralOperator { body, parts, .. } => {
            *body=clone_with_fresh_ids(body,ids)?;
            for slot in crate::operators::slots_mut(parts){slot.argument=clone_with_fresh_ids(&slot.argument,ids)?;}
        }
        NodeKind::MacroCall { arguments, .. } => for argument in arguments { *argument=clone_with_fresh_ids(argument,ids)?; },
        NodeKind::Annotation { annotation, base, .. } => { *annotation=clone_with_fresh_ids(annotation,ids)?; *base=clone_with_fresh_ids(base,ids)?; }
        NodeKind::Radical { degree, radicand, .. } => {
            if let Some(degree) = degree { *degree = clone_with_fresh_ids(degree, ids)?; }
            *radicand = clone_with_fresh_ids(radicand, ids)?;
        }
        NodeKind::Scripts { base, slots } => {
            *base = clone_with_fresh_ids(base, ids)?;
            for slot in slots { slot.argument = clone_with_fresh_ids(&slot.argument, ids)?; }
        }
        _ => {}
    }
    Ok(Arc::new(result))
}

fn update_node(root: &Arc<Node>, target: NodeId, f: impl FnOnce(&mut Node) -> EngineResult<()>) -> EngineResult<Arc<Node>> {
    let index = TreeIndex::new(root)?;
    let mut path = Vec::new(); let mut current = target;
    loop {
        path.push(current);
        let entry = index.nodes.get(&current).ok_or_else(invalid_position)?;
        match entry.parent { Some(parent) => current = parent, None => break }
    }
    path.reverse();
    fn apply<F: FnOnce(&mut Node) -> EngineResult<()>>(node: &Arc<Node>, path: &[NodeId], f: &mut Option<F>) -> EngineResult<Arc<Node>> {
        if path.first() != Some(&node.id) { return Err(invalid_position()); }
        let mut changed = node.as_ref().clone(); changed.origin = None;
        if path.len() == 1 { f.take().ok_or_else(invalid_position)?(&mut changed)?; }
        else {
            let next = path[1]; let mut found = false;
            let mut visit = |child: &mut Arc<Node>| -> EngineResult<()> {
                if child.id == next { *child = apply(child, &path[1..], f)?; found = true; } Ok(())
            };
            match &mut changed.kind {
                NodeKind::Row { entries, .. } => for entry in entries { visit(&mut entry.node)?; },
                NodeKind::Grid { rows, .. } => for row in rows { visit(row)?; },
                NodeKind::GridRow { cells, .. } => for cell in cells { visit(cell)?; },
                NodeKind::Fraction { numerator, denominator, .. } => { visit(numerator)?; visit(denominator)?; }
                NodeKind::Fence { body, .. } | NodeKind::Accent { body, .. } | NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => visit(body)?,
                NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => for slot in crate::operators::slots_mut(parts){visit(&mut slot.argument)?;},
                NodeKind::GeneralOperator { body, parts, .. } => {
                    visit(body)?;for slot in crate::operators::slots_mut(parts){visit(&mut slot.argument)?;}
                }
                NodeKind::MacroCall { arguments, .. } => for argument in arguments {visit(argument)?;},
                NodeKind::Annotation { annotation, base, .. } => { visit(annotation)?; visit(base)?; }
                NodeKind::Radical { degree, radicand, .. } => { if let Some(degree) = degree { visit(degree)?; } visit(radicand)?; }
                NodeKind::Scripts { base, slots } => { visit(base)?; for slot in slots { visit(&mut slot.argument)?; } }
                _ => {}
            }
            if !found { return Err(invalid_position()); }
        }
        Ok(Arc::new(changed))
    }
    apply(root, &path, &mut Some(f))
}
