//! Lossless source projection of a tree changed by mathematical transactions.
//! This serializes and reindexes known structure; it does NOT parse that source
//! to rediscover the operation. Source editor parsing remains a separate route.
use std::collections::BTreeMap;
use std::sync::Arc;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{self, Node, NodeKind, ProjectionRange};
use crate::parser::{self, ParseResult};
use crate::session::NodeId;
use crate::source::{SourceText, tokenize};

pub fn project_tree(tree: Arc<Node>, source_id: u32) -> EngineResult<Arc<ParseResult>> {
    project_tree_with_symbols(tree,source_id,Arc::default())
}
pub fn project_tree_with_symbols(tree: Arc<Node>, source_id: u32, symbols:Arc<crate::custom_symbols::SymbolCatalog>) -> EngineResult<Arc<ParseResult>> {
    let (text, ranges) = model::projected_latex(&tree)?;
    let source = SourceText::new(source_id, text)?;
    let tree = reindex(&tree, &source, &ranges)?;
    let tokens = tokenize(&source)?;
    Ok(Arc::new(parser::assemble_with_symbols(source, tree, tokens, Vec::new(), symbols)?))
}

fn reindex(node: &Arc<Node>, source: &Arc<SourceText>, ranges: &BTreeMap<NodeId, ProjectionRange>) -> EngineResult<Arc<Node>> {
    let range = ranges.get(&node.id).ok_or_else(|| EngineError::new(ErrorCode::InvalidInput, "Missing node source projection"))?;
    let mut result = node.as_ref().clone();
    result.origin = Some(source.slice(range.start, range.end)?);
    match &mut result.kind {
        NodeKind::Row { grouping, entries, .. } => {
            *grouping = range.row_grouping.ok_or_else(|| EngineError::new(ErrorCode::InvalidInput, "Missing row grouping projection"))?;
            for entry in entries { entry.node = reindex(&entry.node, source, ranges)?; }
        }
        NodeKind::Grid { rows, .. } => for row in rows { *row=reindex(row,source,ranges)?; },
        NodeKind::GridRow { cells, .. } => for cell in cells { *cell=reindex(cell,source,ranges)?; },
        NodeKind::Fraction { numerator, denominator, .. } => {
            *numerator = reindex(numerator, source, ranges)?;
            *denominator = reindex(denominator, source, ranges)?;
        }
        NodeKind::Fence { body, .. } | NodeKind::Accent { body, .. } | NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => *body = reindex(body, source, ranges)?,
        NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => for slot in crate::operators::slots_mut(parts) {
            slot.argument=reindex(&slot.argument,source,ranges)?;
        },
        NodeKind::GeneralOperator { body, parts, .. } => {
            *body=reindex(body,source,ranges)?;
            for slot in crate::operators::slots_mut(parts){slot.argument=reindex(&slot.argument,source,ranges)?;}
        }
        NodeKind::MacroCall { arguments, .. } => for argument in arguments { *argument=reindex(argument,source,ranges)?; },
        NodeKind::Annotation { annotation, base, .. } => { *annotation = reindex(annotation, source, ranges)?; *base = reindex(base, source, ranges)?; }
        NodeKind::Radical { degree, radicand, .. } => {
            if let Some(degree) = degree { *degree = reindex(degree, source, ranges)?; }
            *radicand = reindex(radicand, source, ranges)?;
        }
        NodeKind::Scripts { base, slots } => {
            *base = reindex(base, source, ranges)?;
            for slot in slots { slot.argument = reindex(&slot.argument, source, ranges)?; }
        }
        NodeKind::UnparsedSource { .. } => {
            return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Opaque source requires explicit source editing before visual mutation"));
        }
        _ => {}
    }
    Ok(Arc::new(result))
}
