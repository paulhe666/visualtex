//! Context-safe source editing. The entire source is lexed so comments/control
//! words cannot invalidate an assumed group boundary. Only a proven closed
//! braced row is reparsed; otherwise the formula is reparsed. Identity mapping
//! uses source ranges and structural roles, never content hashes.
use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;
use serde::{Deserialize, Serialize};
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{Grouping, IdAllocator, Node, NodeKind, ScriptKind};
use crate::parser::{self, Diagnostic, ParseResult};
use crate::session::NodeId;
use crate::source::{MAX_SOURCE_BYTES, SourceSpan, SourceText, tokenize};

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SourceEdit {
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub replacement: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
pub enum ParseScope {
    None,
    Subtree { node_id: NodeId, parsed_nodes: usize },
    Formula { parsed_nodes: usize, reason: String },
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SourceChangeMap {
    pub old_range: SourceSpan,
    pub new_range: SourceSpan,
    pub retained_ids: Vec<NodeId>,
    pub removed_ids: Vec<NodeId>,
    pub inserted_ids: Vec<NodeId>,
    pub changed_ids: Vec<NodeId>,
    pub parse_scope: ParseScope,
    /// Lexing and final validation are still whole-formula operations. This
    /// counter makes that cost visible rather than claiming wholly local work.
    pub lexed_tokens: usize,
}

pub struct PreparedSourceEdit {
    pub document: Arc<ParseResult>,
    pub mapping: SourceChangeMap,
    pub changed: bool,
}

#[derive(Clone, Copy)]
struct PatchRange { start: usize, end: usize, replacement_len: usize }

impl PatchRange {
    fn shifted(&self, value: usize) -> usize {
        // The final source has already passed its byte budget; callers use this
        // only for ends at/after the removed range, so subtraction is safe.
        value - (self.end - self.start) + self.replacement_len
    }
    fn map_span(&self, start: usize, end: usize, keep_enclosing: bool) -> Option<(usize, usize)> {
        if end <= self.start { return Some((start, end)); }
        if start >= self.end { return Some((self.shifted(start), self.shifted(end))); }
        if keep_enclosing && start <= self.start && end >= self.end
            && (start < self.start || end > self.end)
        {
            return Some((start, self.shifted(end)));
        }
        None
    }
}

pub fn prepare(previous: &Arc<ParseResult>, edit: &SourceEdit, next_source_id: u32,
    ids: &mut IdAllocator) -> EngineResult<PreparedSourceEdit>
{
    let mut working_ids = ids.clone();
    let prepared = prepare_working(previous, edit, next_source_id, &mut working_ids)?;
    *ids = working_ids;
    Ok(prepared)
}

fn prepare_working(previous: &Arc<ParseResult>, edit: &SourceEdit, next_source_id: u32,
    ids: &mut IdAllocator) -> EngineResult<PreparedSourceEdit>
{
    let start = previous.source.utf16_to_byte(edit.start_utf16)?;
    let end = previous.source.utf16_to_byte(edit.end_utf16)?;
    if start > end {
        return Err(EngineError::new(ErrorCode::InvalidPosition, "Source edit range is reversed"));
    }
    let old_range = previous.source.span(start, end)?;
    if previous.source.text()[start..end] == edit.replacement {
        let retained_ids = node_map(&previous.tree).into_keys().collect();
        return Ok(PreparedSourceEdit { document: Arc::clone(previous), changed: false,
            mapping: SourceChangeMap { old_range, new_range: old_range, retained_ids,
                removed_ids: Vec::new(), inserted_ids: Vec::new(), changed_ids: Vec::new(),
                parse_scope: ParseScope::None, lexed_tokens: 0 } });
    }
    if next_source_id <= previous.source.id() {
        return Err(EngineError::new(ErrorCode::InvalidInput, "Changed source requires a fresh monotonic source identity"));
    }
    let final_len = (previous.source.len() - (end - start)).saturating_add(edit.replacement.len());
    if final_len > MAX_SOURCE_BYTES {
        return Err(EngineError::new(ErrorCode::ResourceLimit, "Edited source exceeds the byte budget"));
    }
    let mut text = String::with_capacity(final_len);
    text.push_str(&previous.source.text()[..start]);
    text.push_str(&edit.replacement);
    text.push_str(&previous.source.text()[end..]);
    let source = SourceText::new(next_source_id, text)?;
    let patch = PatchRange { start, end, replacement_len: edit.replacement.len() };
    let tokens = tokenize(&source)?;
    let lexed_tokens = tokens.len();
    let mut all_current_origins = true;
    previous.tree.visit(&mut |node| {
        if !node.origin.as_ref().is_some_and(|o| o.span.source_id == previous.source.id()) {
            all_current_origins = false;
        }
    });
    let mut reason = if !previous.complete { "previousSourceIncomplete" }
        else if !all_current_origins { "nonCurrentOrigins" } else { "noClosedContainingRow" };
    let mut document = None;
    let mut scope = None;
    if previous.complete && all_current_origins {
        if let Some(candidate) = smallest_closed_row(&previous.tree, patch) {
            let span = candidate.origin.as_ref().unwrap().span;
            let new_end = patch.shifted(span.byte_end as usize);
            if let Some(mut fragment) = parser::parse_braced_fragment(Arc::clone(&source), &tokens,
                span.byte_start as usize, new_end, Arc::clone(&previous.symbol_assets))?
            {
                let parsed_nodes = fragment.facts.nodes;
                reconcile(candidate, &mut fragment.tree, patch, ids)?;
                let tree = rebase_and_replace(&previous.tree, candidate.id, &fragment.tree, &source, patch)?;
                let mut diagnostics = Vec::new();
                for diagnostic in &previous.diagnostics {
                    if diagnostic.span.byte_start >= span.byte_start && diagnostic.span.byte_end <= span.byte_end {
                        continue;
                    }
                    diagnostics.push(rebase_diagnostic(diagnostic, &source, patch)?);
                }
                diagnostics.extend(fragment.diagnostics);
                document = Some(parser::assemble_with_symbols(Arc::clone(&source), tree, tokens.clone(), diagnostics, Arc::clone(&previous.symbol_assets))?);
                scope = Some(ParseScope::Subtree { node_id: candidate.id, parsed_nodes });
            } else { reason = "fragmentBoundaryChanged"; }
        }
    }
    let document = match document {
        Some(document) => document,
        None => {
            let mut parsed = parser::parse_tokens_with_symbols(Arc::clone(&source), tokens, previous.tree.id, &mut IdAllocator::new(2), Arc::clone(&previous.symbol_assets))?;
            let parsed_nodes = parsed.facts.nodes;
            reconcile(&previous.tree, &mut parsed.tree, patch, ids)?;
            crate::model::validate_tree(&parsed.tree)?;
            scope = Some(ParseScope::Formula { parsed_nodes, reason: reason.into() });
            parsed
        }
    };
    let old_nodes = node_map(&previous.tree);
    let new_nodes = node_map(&document.tree);
    let old_ids: BTreeSet<NodeId> = old_nodes.keys().copied().collect();
    let new_ids: BTreeSet<NodeId> = new_nodes.keys().copied().collect();
    let retained_ids: Vec<_> = old_ids.intersection(&new_ids).copied().collect();
    let removed_ids = old_ids.difference(&new_ids).copied().collect();
    let inserted_ids: Vec<_> = new_ids.difference(&old_ids).copied().collect();
    let mut changed_ids = inserted_ids.clone();
    for id in &retained_ids {
        let old_text = old_nodes[id].origin.as_ref().map(|o| o.text());
        let new_text = new_nodes[id].origin.as_ref().map(|o| o.text());
        if old_text != new_text { changed_ids.push(*id); }
    }
    changed_ids.sort();
    let mapping = SourceChangeMap { old_range, new_range: source.span(start, start + edit.replacement.len())?,
        retained_ids, removed_ids, inserted_ids, changed_ids, parse_scope: scope.unwrap(), lexed_tokens };
    Ok(PreparedSourceEdit { document: Arc::new(document), mapping, changed: true })
}

fn smallest_closed_row(node: &Node, patch: PatchRange) -> Option<&Node> {
    let mut best = None;
    if let (NodeKind::Row { grouping: Grouping::Braced, .. }, Some(origin)) = (&node.kind, &node.origin) {
        let start = origin.span.byte_start as usize;
        let end = origin.span.byte_end as usize;
        if end >= start + 2 && origin.text().starts_with('{') && origin.text().ends_with('}')
            && start < patch.start && patch.end < end
        { best = Some(node); }
    }
    for child in node.children() {
        if let Some(candidate) = smallest_closed_row(child, patch) {
            let len = |n: &Node| n.origin.as_ref().unwrap().span.byte_end - n.origin.as_ref().unwrap().span.byte_start;
            if best.is_none_or(|current| len(candidate) < len(current)) { best = Some(candidate); }
        }
    }
    best
}

fn node_map(root: &Node) -> BTreeMap<NodeId, &Node> {
    fn walk<'a>(node: &'a Node, map: &mut BTreeMap<NodeId, &'a Node>) {
        map.insert(node.id, node);
        for child in node.children() { walk(child, map); }
    }
    let mut result = BTreeMap::new();
    walk(root, &mut result);
    result
}

fn kind(node: &Node) -> u8 {
    match &node.kind {
        NodeKind::Row { grouping: Grouping::Bare, .. } => 0,
        NodeKind::Row { grouping: Grouping::Braced, .. } => 1,
        NodeKind::Symbol { .. } => 2, NodeKind::Text { .. } => 3,
        NodeKind::Fraction { .. } => 4, NodeKind::Scripts { .. } => 5,
        NodeKind::EmptySlot { .. } => 6, NodeKind::UnparsedSource { .. } => 7,
        NodeKind::Radical { .. } => 8, NodeKind::Fence { .. } => 9,
        NodeKind::Accent { .. } => 10, NodeKind::Annotation { .. } => 11,
        NodeKind::LargeOperator { .. } => 12,
        NodeKind::Style { .. } => 13,
        NodeKind::Space { .. } => 14,
        NodeKind::Phantom { .. } => 15,
        NodeKind::Boxed { .. } => 16,
        NodeKind::Color { .. } => 17,
        NodeKind::ColorChange { .. } => 18,
        NodeKind::GeneralOperator { .. } => 19,
        NodeKind::Grid { .. } => 20,
        NodeKind::GridRow { .. } => 21,
        NodeKind::MacroCall { .. } => 22,
        NodeKind::MiddleDelimiter { .. } => 23,
        NodeKind::CustomSymbol { .. } => 24,
        NodeKind::StyleChange { .. } => 25,
    }
}

type IdentityKey = (usize, usize, u8, u8);
fn visit_roles<'a>(node: &'a Node, role: u8, f: &mut impl FnMut(&'a Node, u8)) {
    f(node, role);
    match &node.kind {
        NodeKind::Row { entries, .. } => for entry in entries { visit_roles(&entry.node, 1, f); },
        NodeKind::Fraction { numerator, denominator, .. } => {
            visit_roles(numerator, 2, f); visit_roles(denominator, 3, f);
        }
        NodeKind::Fence { body, .. } => visit_roles(body, 9, f),
        NodeKind::Accent { body, .. } => visit_roles(body, 10, f),
        NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => for slot in crate::operators::slots(parts) {
            visit_roles(&slot.argument,if slot.script==ScriptKind::Superscript{13}else{14},f);
        },
        NodeKind::GeneralOperator { body, parts, .. } => {
            visit_roles(body,19,f);
            for slot in crate::operators::slots(parts){visit_roles(&slot.argument,if slot.script==ScriptKind::Superscript{13}else{14},f);}
        }
        NodeKind::MacroCall { arguments, .. } => for (index,argument) in arguments.iter().enumerate() {
            visit_roles(argument,30+index as u8,f);
        },
        NodeKind::Grid { rows, .. } => for row in rows { visit_roles(row,20,f); },
        NodeKind::GridRow { cells, .. } => for cell in cells { visit_roles(cell,21,f); },
        NodeKind::Style { body, .. } => visit_roles(body, 15, f),
        NodeKind::Phantom { body, .. } => visit_roles(body, 16, f),
        NodeKind::Boxed { body, .. } => visit_roles(body, 17, f),
        NodeKind::Color { body, .. } => visit_roles(body, 18, f),
        NodeKind::Annotation { annotation, base, .. } => { visit_roles(annotation, 11, f); visit_roles(base, 12, f); }
        NodeKind::Radical { degree, radicand, .. } => {
            if let Some(degree) = degree { visit_roles(degree, 7, f); }
            visit_roles(radicand, 8, f);
        }
        NodeKind::Scripts { base, slots } => {
            visit_roles(base, 4, f);
            for slot in slots { visit_roles(&slot.argument, if slot.script == ScriptKind::Superscript { 5 } else { 6 }, f); }
        }
        _ => {}
    }
}
fn mutate_roles(node: &mut Arc<Node>, role: u8, f: &mut impl FnMut(&mut Node, u8) -> EngineResult<()>) -> EngineResult<()> {
    let node = Arc::make_mut(node);
    f(node, role)?;
    match &mut node.kind {
        NodeKind::Row { entries, .. } => for entry in entries { mutate_roles(&mut entry.node, 1, f)?; },
        NodeKind::Fraction { numerator, denominator, .. } => {
            mutate_roles(numerator, 2, f)?; mutate_roles(denominator, 3, f)?;
        }
        NodeKind::Fence { body, .. } => mutate_roles(body, 9, f)?,
        NodeKind::Accent { body, .. } => mutate_roles(body, 10, f)?,
        NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => for slot in crate::operators::slots_mut(parts) {
            mutate_roles(&mut slot.argument,if slot.script==ScriptKind::Superscript{13}else{14},f)?;
        },
        NodeKind::GeneralOperator { body, parts, .. } => {
            mutate_roles(body,19,f)?;
            for slot in crate::operators::slots_mut(parts){mutate_roles(&mut slot.argument,if slot.script==ScriptKind::Superscript{13}else{14},f)?;}
        }
        NodeKind::MacroCall { arguments, .. } => for (index,argument) in arguments.iter_mut().enumerate() {
            mutate_roles(argument,30+index as u8,f)?;
        },
        NodeKind::Grid { rows, .. } => for row in rows { mutate_roles(row,20,f)?; },
        NodeKind::GridRow { cells, .. } => for cell in cells { mutate_roles(cell,21,f)?; },
        NodeKind::Style { body, .. } => mutate_roles(body, 15, f)?,
        NodeKind::Phantom { body, .. } => mutate_roles(body, 16, f)?,
        NodeKind::Boxed { body, .. } => mutate_roles(body, 17, f)?,
        NodeKind::Color { body, .. } => mutate_roles(body, 18, f)?,
        NodeKind::Annotation { annotation, base, .. } => { mutate_roles(annotation, 11, f)?; mutate_roles(base, 12, f)?; }
        NodeKind::Radical { degree, radicand, .. } => {
            if let Some(degree) = degree { mutate_roles(degree, 7, f)?; }
            mutate_roles(radicand, 8, f)?;
        }
        NodeKind::Scripts { base, slots } => {
            mutate_roles(base, 4, f)?;
            for slot in slots { mutate_roles(&mut slot.argument, if slot.script == ScriptKind::Superscript { 5 } else { 6 }, f)?; }
        }
        _ => {}
    }
    Ok(())
}

fn reconcile(old: &Node, new: &mut Arc<Node>, patch: PatchRange, ids: &mut IdAllocator) -> EngineResult<()> {
    let mut candidates: BTreeMap<IdentityKey, Vec<&Node>> = BTreeMap::new();
    visit_roles(old, 0, &mut |node, role| {
        if role == 0 { return; }
        if let Some(origin) = &node.origin {
            if let Some((start, end)) = patch.map_span(origin.span.byte_start as usize, origin.span.byte_end as usize, true) {
                candidates.entry((start, end, role, kind(node))).or_default().push(node);
            }
        }
    });
    let mut new_counts = BTreeMap::new();
    visit_roles(new, 0, &mut |node, role| {
        if let Some(origin) = &node.origin {
            *new_counts.entry((origin.span.byte_start as usize, origin.span.byte_end as usize, role, kind(node))).or_insert(0usize) += 1;
        }
    });
    let mut assigned = BTreeSet::new();
    mutate_roles(new, 0, &mut |node, role| {
        if role == 0 { node.id = old.id; assigned.insert(old.id); return Ok(()); }
        let existing = node.origin.as_ref().and_then(|origin| {
            let key = (origin.span.byte_start as usize, origin.span.byte_end as usize, role, kind(node));
            let list = candidates.get(&key)?;
            if list.len() != 1 || new_counts.get(&key) != Some(&1) { return None; }
            let previous = list[0];
            let previous_span = previous.origin.as_ref()?.span;
            let outside = previous_span.byte_end as usize <= patch.start || previous_span.byte_start as usize >= patch.end;
            if outside && previous.origin.as_ref()?.text() != origin.text() { return None; }
            (!assigned.contains(&previous.id)).then_some(previous.id)
        });
        node.id = match existing { Some(id) => id, None => ids.allocate()? };
        if !assigned.insert(node.id) {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Source identity reconciliation produced a duplicate"));
        }
        Ok(())
    })?;
    // Column identities are metadata, not caret/source nodes. Match retained
    // grids before assigning fresh column IDs; never expose parser-local IDs.
    let old_nodes=node_map(old);
    mutate_roles(new,0,&mut |node,_|{
        if let NodeKind::Grid { columns, .. }=&mut node.kind {
            let prior=old_nodes.get(&node.id).and_then(|n|match &n.kind{NodeKind::Grid{columns,..}=>Some(columns),_=>None});
            for (i,column) in columns.iter_mut().enumerate(){
                *column=match prior.and_then(|c|c.get(i)){Some(id)=>*id,None=>crate::grids::GridColumnId::allocate(ids)?};
            }
        }
        Ok(())
    })
}

fn rebase_and_replace(old: &Arc<Node>, target: NodeId, replacement: &Arc<Node>, source: &Arc<SourceText>, patch: PatchRange) -> EngineResult<Arc<Node>> {
    if old.id == target { return Ok(Arc::clone(replacement)); }
    let mut node = old.as_ref().clone();
    if let Some(origin) = &node.origin {
        let (start, end) = patch.map_span(origin.span.byte_start as usize, origin.span.byte_end as usize, true)
            .ok_or_else(|| EngineError::new(ErrorCode::InvalidPosition, "Closed subtree edit unexpectedly crossed a sibling boundary"))?;
        node.origin = Some(source.slice(start, end)?);
    }
    match &mut node.kind {
        NodeKind::Row { entries, .. } => for entry in entries {
            entry.node = rebase_and_replace(&entry.node, target, replacement, source, patch)?;
        },
        NodeKind::Grid { rows, .. } => for row in rows { *row=rebase_and_replace(row,target,replacement,source,patch)?; },
        NodeKind::GridRow { cells, .. } => for cell in cells { *cell=rebase_and_replace(cell,target,replacement,source,patch)?; },
        NodeKind::Fraction { numerator, denominator, .. } => {
            *numerator = rebase_and_replace(numerator, target, replacement, source, patch)?;
            *denominator = rebase_and_replace(denominator, target, replacement, source, patch)?;
        }
        NodeKind::Fence { body, .. } | NodeKind::Accent { body, .. } | NodeKind::Style { body, .. }
        | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => *body = rebase_and_replace(body, target, replacement, source, patch)?,
        NodeKind::Annotation { annotation, base, .. } => {
            *annotation = rebase_and_replace(annotation, target, replacement, source, patch)?;
            *base = rebase_and_replace(base, target, replacement, source, patch)?;
        }
        NodeKind::Radical { degree, radicand, .. } => {
            if let Some(degree) = degree { *degree = rebase_and_replace(degree, target, replacement, source, patch)?; }
            *radicand = rebase_and_replace(radicand, target, replacement, source, patch)?;
        }
        NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => for slot in crate::operators::slots_mut(parts) {
            slot.argument=rebase_and_replace(&slot.argument,target,replacement,source,patch)?;
        },
        NodeKind::GeneralOperator { body, parts, .. } => {
            *body=rebase_and_replace(body,target,replacement,source,patch)?;
            for slot in crate::operators::slots_mut(parts){slot.argument=rebase_and_replace(&slot.argument,target,replacement,source,patch)?;}
        }
        NodeKind::MacroCall { arguments, .. } => for argument in arguments {
            *argument=rebase_and_replace(argument,target,replacement,source,patch)?;
        },
        NodeKind::Scripts { base, slots } => {
            *base = rebase_and_replace(base, target, replacement, source, patch)?;
            for slot in slots { slot.argument = rebase_and_replace(&slot.argument, target, replacement, source, patch)?; }
        }
        _ => {}
    }
    Ok(Arc::new(node))
}

fn rebase_diagnostic(old: &Diagnostic, source: &SourceText, patch: PatchRange) -> EngineResult<Diagnostic> {
    let (start, end) = patch.map_span(old.span.byte_start as usize, old.span.byte_end as usize, true)
        .ok_or_else(|| EngineError::new(ErrorCode::InvalidPosition, "Diagnostic crosses a replaced source range"))?;
    let mut diagnostic = old.clone();
    diagnostic.span = source.span(start, end)?;
    Ok(diagnostic)
}
