use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;
use serde::Serialize;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::session::NodeId;
use crate::source::SourceSlice;

pub const MAX_NODES: usize = 8192;
pub const MAX_DEPTH: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum AtomClass { Ordinary, Number, Binary, Relation, Open, Close, Punctuation, Operator, Inner }

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum Grouping { Bare, Braced }

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ScriptKind { Superscript, Subscript }

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RowEntry {
    pub leading: String,
    pub node: Arc<Node>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ScriptSlot {
    pub script: ScriptKind,
    pub before_marker: String,
    pub before_argument: String,
    pub argument: Arc<Node>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
pub enum NodeKind {
    Row { grouping: Grouping, entries: Vec<RowEntry>, trailing: String },
    Grid { environment: crate::grids::GridKind, columns: Vec<crate::grids::GridColumnId>,
        array_spec: Option<crate::grids::ArrayColumnSpec>, before_begin_name: String, before_end_name: String, rows: Vec<Arc<Node>> },
    GridRow { cells: Vec<Arc<Node>>, separators: Vec<String>, row_break: String },
    Symbol { value: String, class: AtomClass, command: Option<String> },
    Text { command: String, before_body: String, text: String },
    Fraction {
        command: String,
        before_numerator: String,
        numerator: Arc<Node>,
        before_denominator: String,
        denominator: Arc<Node>,
    },
    Radical {
        command: String,
        before_degree: String,
        degree: Option<Arc<Node>>,
        before_radicand: String,
        radicand: Arc<Node>,
    },
    Fence { before_left: String, left: String, body: Arc<Node>, before_right: String, right: String },
    MiddleDelimiter { before_delimiter: String, delimiter: String },
    Accent { command: String, before_body: String, body: Arc<Node> },
    Annotation { command: String, above: bool, before_annotation: String,
        annotation: Arc<Node>, before_base: String, base: Arc<Node> },
    LargeOperator { command: String, parts: Vec<crate::operators::OperatorPart> },
    GeneralOperator { command: String, before_body: String, body: Arc<Node>, parts: Vec<crate::operators::OperatorPart> },
    CustomSymbol { command: String, asset_sha256: String,
        #[serde(skip)] asset: Arc<crate::custom_symbols::SymbolAsset>, parts: Vec<crate::operators::OperatorPart> },
    MacroCall { command: String, macro_kind: crate::macros::BuiltinMacroKind,
        before_arguments: Vec<String>, arguments: Vec<Arc<Node>> },
    Style { command: String, before_body: String, variant: crate::styles::MathVariant, body: Arc<Node> },
    Space { command: String, width_mu: i16 },
    Boxed { before_body: String, body: Arc<Node> },
    Color { spec: crate::paint::ColorSpec, before_body: String, body: Arc<Node> },
    ColorChange { spec: crate::paint::ColorSpec },
    /// \displaystyle-like declaration: changes the math style of the
    /// following siblings in its row only, like a color declaration.
    StyleChange { command: String, style: crate::layout::MathStyle },
    Phantom { command: String, before_body: String, axis: crate::spacing::PhantomAxis, body: Arc<Node> },
    Scripts { base: Arc<Node>, slots: Vec<ScriptSlot> },
    EmptySlot { label: String },
    UnparsedSource { raw: String },
}

#[derive(Debug, Clone, Serialize)]
pub struct Node {
    pub id: NodeId,
    pub origin: Option<SourceSlice>,
    #[serde(flatten)]
    pub kind: NodeKind,
}

impl Node {
    pub fn empty_row(id: NodeId, grouping: Grouping) -> Arc<Self> {
        Arc::new(Self { id, origin: None,
            kind: NodeKind::Row { grouping, entries: Vec::new(), trailing: String::new() } })
    }

    pub fn children(&self) -> Vec<&Node> {
        match &self.kind {
            NodeKind::Row { entries, .. } => entries.iter().map(|entry| entry.node.as_ref()).collect(),
            NodeKind::Grid { rows, .. } => rows.iter().map(Arc::as_ref).collect(),
            NodeKind::GridRow { cells, .. } => cells.iter().map(Arc::as_ref).collect(),
            NodeKind::Fraction { numerator, denominator, .. } => vec![numerator, denominator],
            NodeKind::Fence { body, .. } | NodeKind::Accent { body, .. } => vec![body],
            NodeKind::Annotation { annotation, base, .. } => vec![annotation, base],
            NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => crate::operators::slots(parts).map(|slot|slot.argument.as_ref()).collect(),
            NodeKind::GeneralOperator { body, parts, .. } => std::iter::once(body.as_ref())
                .chain(crate::operators::slots(parts).map(|slot|slot.argument.as_ref())).collect(),
            NodeKind::MacroCall { arguments, .. } => arguments.iter().map(Arc::as_ref).collect(),
            NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => vec![body],
            NodeKind::Radical { degree, radicand, .. } => degree.iter().map(Arc::as_ref)
                .chain(std::iter::once(radicand.as_ref())).collect(),
            NodeKind::Scripts { base, slots } => {
                let mut nodes = vec![base.as_ref()];
                nodes.extend(slots.iter().map(|slot| slot.argument.as_ref()));
                nodes
            }
            _ => Vec::new(),
        }
    }

    pub fn find(&self, id: NodeId) -> Option<&Node> {
        if self.id == id { return Some(self); }
        self.children().into_iter().find_map(|child| child.find(id))
    }

    pub fn visit(&self, f: &mut impl FnMut(&Node)) {
        f(self);
        for child in self.children() { child.visit(f); }
    }

    pub fn is_row(&self) -> bool { matches!(self.kind, NodeKind::Row { .. }) }

    pub fn is_empty_row(&self) -> bool {
        matches!(&self.kind, NodeKind::Row { entries, .. } if entries.is_empty())
    }

    /// Empty grouping is valid source but does not fill a required argument.
    /// Keep is_empty_row structural: navigation needs to distinguish a row
    /// with an editable empty group from a physically empty row.
    pub fn has_formula_content(&self) -> bool {
        match &self.kind {
            NodeKind::Symbol { value, .. } => !value.is_empty(),
            NodeKind::LargeOperator { .. } | NodeKind::CustomSymbol { .. } | NodeKind::Space { .. } | NodeKind::MiddleDelimiter { .. } => true,
            NodeKind::MacroCall { arguments, .. } => arguments.iter().any(|argument| argument.has_formula_content()),
            NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => body.has_formula_content(),
            NodeKind::ColorChange { .. } | NodeKind::StyleChange { .. } => false,
            NodeKind::Text { text, .. } => !text.is_empty(),
            NodeKind::UnparsedSource { raw } => !raw.is_empty(),
            NodeKind::EmptySlot { .. } => false,
            _ => self.children().into_iter().any(Node::has_formula_content),
        }
    }

    pub fn has_unfilled_slots(&self) -> bool {
        match &self.kind {
            NodeKind::EmptySlot { .. } => true,
            NodeKind::Fraction { numerator, denominator, .. } => {
                !numerator.has_formula_content() || !denominator.has_formula_content()
                    || numerator.has_unfilled_slots() || denominator.has_unfilled_slots()
            }
            NodeKind::Radical { degree, radicand, .. } => !radicand.has_formula_content()
                || radicand.has_unfilled_slots()
                || degree.as_ref().is_some_and(|d| !d.has_formula_content() || d.has_unfilled_slots()),
            NodeKind::Fence { body, .. } | NodeKind::Accent { body, .. } => !body.has_formula_content() || body.has_unfilled_slots(),
            NodeKind::Annotation { annotation, base, .. } => !base.has_formula_content() || !annotation.has_formula_content()
                || base.has_unfilled_slots() || annotation.has_unfilled_slots(),
            NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => crate::operators::slots(parts)
                .any(|slot| !slot.argument.has_formula_content() || slot.argument.has_unfilled_slots()),
            NodeKind::GeneralOperator { body, parts, .. } => !body.has_formula_content() || body.has_unfilled_slots()
                || crate::operators::slots(parts).any(|slot|!slot.argument.has_formula_content() || slot.argument.has_unfilled_slots()),
            NodeKind::MacroCall { arguments, .. } => arguments.iter().any(|argument| !argument.has_formula_content() || argument.has_unfilled_slots()),
            NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } | NodeKind::Color { body, .. } => !body.has_formula_content() || body.has_unfilled_slots(),
            NodeKind::Scripts { base, slots } => !base.has_formula_content() || base.has_unfilled_slots()
                || slots.iter().any(|slot| !slot.argument.has_formula_content() || slot.argument.has_unfilled_slots()),
            _ => self.children().iter().any(|child| child.has_unfilled_slots()),
        }
    }
}

#[derive(Debug, Clone)]
pub struct IdAllocator { next: u32 }

impl IdAllocator {
    pub fn new(next: u32) -> Self { Self { next } }
    pub fn next_unused(&self) -> u32 { self.next }
    pub fn allocate(&mut self) -> EngineResult<NodeId> {
        if self.next == 0 {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Invalid node identity allocator"));
        }
        let next = self.next.checked_add(1)
            .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Node identity space exhausted"))?;
        let id = NodeId(self.next);
        self.next = next;
        Ok(id)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LatexMode { Lossless, Canonical }

pub fn latex(node: &Node, mode: LatexMode) -> String {
    let mut output = String::new();
    emit(node, mode, false, &mut output, &mut None);
    output
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct ProjectionRange {
    pub start: usize,
    pub end: usize,
    pub row_grouping: Option<Grouping>,
}

pub(crate) fn projected_latex(node: &Node) -> EngineResult<(String, BTreeMap<NodeId, ProjectionRange>)> {
    let facts = validate_tree(node)?;
    let mut output = String::new();
    let mut ranges = Some(BTreeMap::new());
    emit(node, LatexMode::Lossless, false, &mut output, &mut ranges);
    let ranges = ranges.unwrap();
    if ranges.len() != facts.nodes {
        return Err(EngineError::new(ErrorCode::InvalidInput, "Stale source origin prevents a complete mathematical source projection"));
    }
    Ok((output, ranges))
}

fn emit(node: &Node, mode: LatexMode, argument: bool, output: &mut String,
    ranges: &mut Option<BTreeMap<NodeId, ProjectionRange>>) {
    let start = output.len();
    if mode == LatexMode::Lossless {
        if let Some(origin) = &node.origin {
            // Source can be unchanged while its grammatical role has changed.
            // Moving a bare multi-atom row into a fraction/script needs grouping;
            // copying its old origin verbatim would silently change the formula.
            // A single original TeX symbol remains unbraced when lexically safe,
            // preserving untouched arguments in spellings such as \\frac12.
            let wrap = argument && match &node.kind {
                NodeKind::Row { grouping: Grouping::Bare, entries, .. } => {
                    let one_original_symbol = entries.len() == 1
                        && matches!(entries[0].node.kind, NodeKind::Symbol { .. })
                        && entries[0].node.origin.as_ref().is_some_and(|child| child.text() == origin.text());
                    !one_original_symbol || (ends_control_word(output)
                        && origin.text().as_bytes().first().is_some_and(u8::is_ascii_alphabetic))
                }
                _ => false,
            };
            if wrap { output.push('{'); }
            let content_start = output.len();
            output.push_str(origin.text());
            if wrap {
                // A source fragment ending in a comment must not swallow the
                // newly introduced closing brace or its following denominator.
                if ends_in_line_comment(origin.text()) { output.push('\n'); }
                output.push('}');
            }
            if let Some(ranges) = ranges {
                node.visit(&mut |child| {
                    if let Some(child_origin) = &child.origin {
                        if child_origin.span.source_id == origin.span.source_id
                            && child_origin.span.byte_start >= origin.span.byte_start
                            && child_origin.span.byte_end <= origin.span.byte_end
                        {
                            let row_grouping = match child.kind { NodeKind::Row { grouping, .. } => Some(grouping), _ => None };
                            ranges.insert(child.id, ProjectionRange {
                                start: content_start + (child_origin.span.byte_start - origin.span.byte_start) as usize,
                                end: content_start + (child_origin.span.byte_end - origin.span.byte_start) as usize, row_grouping });
                        }
                    }
                });
                if wrap {
                    ranges.insert(node.id, ProjectionRange { start, end: output.len(), row_grouping: Some(Grouping::Braced) });
                }
            }
            return;
        }
    }
    match &node.kind {
        NodeKind::Row { grouping, entries, trailing } => {
            let braced = *grouping == Grouping::Braced || argument;
            if braced { output.push('{'); }
            let mut previous_ends_in_control_word = false;
            for entry in entries {
                if mode == LatexMode::Lossless && !entry.leading.is_empty() {
                    output.push_str(&entry.leading);
                    // Pasted trailing trivia may become a new sibling's prefix.
                    // Keep its comment, but terminate it before emitting live math.
                    if ends_in_line_comment(&entry.leading) { output.push('\n'); }
                    // Trivia belongs to the source, not the formula token.
                    previous_ends_in_control_word = false;
                }
                let mut boundary = output.len();
                emit(&entry.node, mode, false, output, ranges);
                // Inspect emitted bytes, not just the immediate node kind: a
                // scripted base or nested bare row may also start with a letter.
                // This is required after structural edits in lossless mode too;
                // otherwise preserving "\\alpha" then inserting x creates the
                // unrelated command "\\alphax". Untouched origins remain exact.
                if previous_ends_in_control_word
                    && output.as_bytes().get(boundary).is_some_and(u8::is_ascii_alphabetic)
                {
                    output.insert(boundary, ' ');
                    if let Some(ranges) = ranges {
                        entry.node.visit(&mut |child| if let Some(range) = ranges.get_mut(&child.id) {
                            range.start += 1; range.end += 1;
                        });
                    }
                    boundary += 1;
                }
                if output.len() > boundary {
                    // Only examine this fragment, avoiding quadratic rescans of
                    // long rows of ordinary letters when serializing each atom.
                    previous_ends_in_control_word = ends_control_word(&output[boundary..]);
                }
            }
            if mode == LatexMode::Lossless {
                output.push_str(trailing);
                // Reprojecting an edited row must retain the safe boundary
                // introduced when a formerly bare comment-tailed row was wrapped.
                if braced && ends_in_line_comment(trailing) { output.push('\n'); }
            }
            if braced { output.push('}'); }
        }
        NodeKind::Grid { environment: crate::grids::GridKind::Substack, before_begin_name, rows, .. } => {
            output.push_str("\\substack");
            if mode == LatexMode::Lossless { output.push_str(before_begin_name); }
            output.push('{');
            for row in rows { emit(row, mode, false, output, ranges); }
            if mode == LatexMode::Lossless && ends_in_line_comment(output) { output.push('\n'); }
            output.push('}');
        }
        NodeKind::Grid { environment, array_spec, before_begin_name, before_end_name, rows, .. } => {
            output.push_str("\\begin");
            if mode == LatexMode::Lossless { output.push_str(before_begin_name); }
            output.push('{'); output.push_str(environment.name()); output.push('}');
            if let Some(spec)=array_spec {
                if mode==LatexMode::Lossless {output.push_str(&spec.before_spec);output.push('{');output.push_str(&spec.source);output.push('}');}
                else {output.push('{');output.push_str(&spec.canonical_source());output.push('}');}
            }
            for row in rows { emit(row, mode, false, output, ranges); }
            if mode == LatexMode::Lossless && ends_in_line_comment(output) { output.push('\n'); }
            output.push_str("\\end");
            if mode == LatexMode::Lossless { output.push_str(before_end_name); }
            output.push('{'); output.push_str(environment.name()); output.push('}');
        }
        NodeKind::GridRow { cells, separators, row_break } => {
            for (i,cell) in cells.iter().enumerate() {
                if i > 0 {
                    if mode == LatexMode::Lossless && ends_in_line_comment(output) { output.push('\n'); }
                    output.push_str(&separators[i-1]);
                }
                emit(cell, mode, false, output, ranges);
            }
            if !row_break.is_empty() && mode == LatexMode::Lossless && ends_in_line_comment(output) { output.push('\n'); }
            output.push_str(row_break);
        }
        NodeKind::Symbol { value, command, .. } => output.push_str(command.as_deref().unwrap_or(value)),
        NodeKind::Text { command, before_body, text } => {
            output.push_str(command);
            if mode == LatexMode::Lossless { output.push_str(before_body); }
            output.push('{');
            escape_text(text, output);
            output.push('}');
        }
        NodeKind::Fraction { command, before_numerator, numerator, before_denominator, denominator } => {
            output.push_str(command);
            if mode == LatexMode::Lossless { output.push_str(before_numerator); }
            emit(numerator, mode, true, output, ranges);
            if mode == LatexMode::Lossless { output.push_str(before_denominator); }
            emit(denominator, mode, true, output, ranges);
        }
        NodeKind::Radical { command, before_degree, degree, before_radicand, radicand } => {
            output.push_str(command);
            if let Some(degree) = degree {
                if mode == LatexMode::Lossless { output.push_str(before_degree); }
                output.push('[');
                // A literal right bracket inside an edited optional index must
                // not prematurely close TeX's optional argument. Group only
                // this index, retaining the existing editable row identity.
                if optional_needs_group(degree, mode) {
                    let mut grouped = degree.as_ref().clone();
                    grouped.origin = None;
                    if let NodeKind::Row { grouping, .. } = &mut grouped.kind { *grouping = Grouping::Braced; }
                    emit(&grouped, mode, false, output, ranges);
                } else { emit(degree, mode, false, output, ranges); }
                if mode == LatexMode::Lossless && ends_in_line_comment(output) { output.push('\n'); }
                output.push(']');
            }
            if mode == LatexMode::Lossless { output.push_str(before_radicand); }
            emit(radicand, mode, true, output, ranges);
        }
        NodeKind::Accent { command, before_body, body } => {
            output.push_str(command);
            if mode == LatexMode::Lossless { output.push_str(before_body); }
            emit(body, mode, true, output, ranges);
        }
        NodeKind::Annotation { command, before_annotation, annotation, before_base, base, .. } => {
            output.push_str(command);
            if mode == LatexMode::Lossless { output.push_str(before_annotation); }
            emit(annotation, mode, true, output, ranges);
            if mode == LatexMode::Lossless { output.push_str(before_base); }
            emit(base, mode, true, output, ranges);
        }
        NodeKind::Fence { before_left, left, body, before_right, right } => {
            output.push_str("\\left");
            if mode == LatexMode::Lossless { output.push_str(before_left); }
            output.push_str(left);
            let needs_boundary = ends_control_word(output);
            let boundary = output.len();
            emit(body, mode, false, output, ranges);
            if needs_boundary && output.as_bytes().get(boundary).is_some_and(u8::is_ascii_alphabetic) {
                output.insert(boundary, ' ');
                if let Some(ranges) = ranges {
                    body.visit(&mut |node| { if let Some(r) = ranges.get_mut(&node.id) { r.start += 1; r.end += 1; } });
                }
            }
            if mode == LatexMode::Lossless && ends_in_line_comment(&output[boundary..]) { output.push('\n'); }
            output.push_str("\\right");
            if mode == LatexMode::Lossless { output.push_str(before_right); }
            output.push_str(right);
        }
        NodeKind::MiddleDelimiter { before_delimiter, delimiter } => {
            output.push_str("\\middle");
            if mode == LatexMode::Lossless { output.push_str(before_delimiter); }
            output.push_str(delimiter);
        }
        NodeKind::GeneralOperator { command, before_body, body, parts } => {
            output.push_str(command);
            if mode == LatexMode::Lossless { output.push_str(before_body); }
            emit(body, mode, true, output, ranges);
            emit_operator_parts(parts, mode, output, ranges);
        }
        NodeKind::MacroCall { command, before_arguments, arguments, .. } => {
            output.push_str(command);
            for (index, argument_node) in arguments.iter().enumerate() {
                if mode == LatexMode::Lossless { output.push_str(&before_arguments[index]); }
                emit(argument_node, mode, true, output, ranges);
            }
        }
        NodeKind::LargeOperator { command, parts } | NodeKind::CustomSymbol { command, parts, .. } => {
            output.push_str(command);
            emit_operator_parts(parts, mode, output, ranges);
        }
        NodeKind::Style { command, before_body, body, .. } => {
            output.push_str(command);
            if mode == LatexMode::Lossless { output.push_str(before_body); }
            emit(body, mode, true, output, ranges);
        }
        NodeKind::Color { spec, before_body, body } => {
            output.push_str("\\textcolor");spec.emit(mode==LatexMode::Lossless,output);
            if mode==LatexMode::Lossless {output.push_str(before_body);}
            emit(body,mode,true,output,ranges);
        }
        NodeKind::ColorChange { spec } => {output.push_str("\\color");spec.emit(mode==LatexMode::Lossless,output);}
        NodeKind::StyleChange { command, .. } => output.push_str(command),
        NodeKind::Space { command, .. } => output.push_str(command),
        NodeKind::Boxed { before_body, body } => {
            output.push_str("\\boxed");
            if mode == LatexMode::Lossless { output.push_str(before_body); }
            emit(body, mode, true, output, ranges);
        }
        NodeKind::Phantom { command, before_body, body, .. } => {
            output.push_str(command);
            if mode == LatexMode::Lossless { output.push_str(before_body); }
            emit(body, mode, true, output, ranges);
        }
        NodeKind::Scripts { base, slots } => {
            let base_requires_group = matches!(&base.kind, NodeKind::Row { entries, trailing, grouping: Grouping::Bare }
                if entries.len() != 1 || entries.first().is_some_and(|e| matches!(e.node.kind, NodeKind::Scripts { .. } | NodeKind::LargeOperator { .. } | NodeKind::GeneralOperator { .. } | NodeKind::CustomSymbol { .. }))
                    // A pasted comment at a bare base's end must not swallow
                    // the following script marker or its enclosing group.
                    || (mode == LatexMode::Lossless && ends_in_line_comment(trailing)));
            emit(base, mode, base_requires_group, output, ranges);
            for slot in slots {
                if mode == LatexMode::Lossless { output.push_str(&slot.before_marker); }
                output.push(match slot.script { ScriptKind::Superscript => '^', ScriptKind::Subscript => '_' });
                if mode == LatexMode::Lossless { output.push_str(&slot.before_argument); }
                emit(&slot.argument, mode, true, output, ranges);
            }
        }
        NodeKind::EmptySlot { label } => {
            output.push_str("\\placeholder{");
            escape_text(label, output);
            output.push('}');
        }
        NodeKind::UnparsedSource { raw } => output.push_str(raw),
    }
    if let Some(ranges) = ranges {
        let row_grouping = match node.kind {
            NodeKind::Row { grouping, .. } => Some(if argument { Grouping::Braced } else { grouping }), _ => None,
        };
        ranges.insert(node.id, ProjectionRange { start, end: output.len(), row_grouping });
    }
}

fn emit_operator_parts(parts: &[crate::operators::OperatorPart], mode: LatexMode,
    output: &mut String, ranges: &mut Option<BTreeMap<NodeId, ProjectionRange>>) {
    for part in parts {
        match part {
            crate::operators::OperatorPart::Limits { before, command } => {
                if mode == LatexMode::Lossless { output.push_str(before); }
                output.push_str(command);
            }
            crate::operators::OperatorPart::Script { slot } => {
                if mode == LatexMode::Lossless { output.push_str(&slot.before_marker); }
                output.push(if slot.script == ScriptKind::Superscript { '^' } else { '_' });
                if mode == LatexMode::Lossless { output.push_str(&slot.before_argument); }
                emit(&slot.argument, mode, true, output, ranges);
            }
        }
    }
}

fn optional_needs_group(node: &Node, mode: LatexMode) -> bool {
    // TeX scans an optional argument before expanding commands. Only braces
    // protect an inner closing bracket; a nested sqrt's own [] does not.
    let exposed_argument = |arg: &Node| mode == LatexMode::Lossless
        && arg.origin.as_ref().is_some_and(|origin| {
            matches!(&arg.kind, NodeKind::Row { grouping: Grouping::Bare, entries, .. }
                if entries.len() == 1 && entries[0].node.origin.as_ref().is_some_and(|o| o.text() == origin.text())
                    && matches!(&entries[0].node.kind, NodeKind::Symbol { value, command: None, .. } if value.contains(']')))
        });
    match &node.kind {
        NodeKind::Row { grouping: Grouping::Bare, entries, .. } => entries.iter().any(|e| optional_needs_group(&e.node, mode)),
        NodeKind::Symbol { value, command: None, .. } => value.contains(']'),
        NodeKind::Scripts { base, slots } => optional_needs_group(base, mode)
            || slots.iter().any(|s| exposed_argument(&s.argument)),
        NodeKind::Fraction { numerator, denominator, .. } => exposed_argument(numerator) || exposed_argument(denominator),
        NodeKind::Radical { degree, radicand, .. } => degree.is_some() || exposed_argument(radicand),
        NodeKind::Fence { left, right, body, .. } => left == "]" || right == "]" || optional_needs_group(body, mode),
        NodeKind::LargeOperator { parts, .. } | NodeKind::CustomSymbol { parts, .. } => crate::operators::slots(parts).any(|slot|exposed_argument(&slot.argument)),
        NodeKind::GeneralOperator { body, parts, .. } => exposed_argument(body)
            || crate::operators::slots(parts).any(|slot|exposed_argument(&slot.argument)),
        NodeKind::MacroCall { arguments, .. } => arguments.iter().any(|argument| exposed_argument(argument)),
        NodeKind::Color { spec,body,.. } => spec.model.is_some()||exposed_argument(body),
        NodeKind::ColorChange { spec } => spec.model.is_some(),
        NodeKind::Style { body, .. } | NodeKind::Phantom { body, .. } | NodeKind::Boxed { body, .. } => exposed_argument(body),
        NodeKind::Accent { body, .. } => exposed_argument(body),
        NodeKind::Annotation { annotation, base, .. } => exposed_argument(annotation) || exposed_argument(base),
        _ => false,
    }
}

pub(crate) fn ends_in_line_comment(text: &str) -> bool {
    let mut comment = false;
    let mut chars = text.chars();
    while let Some(ch) = chars.next() {
        if ch == '\n' || ch == '\r' { comment = false; }
        else if !comment {
            match ch {
                '%' => comment = true,
                '\\' => { chars.next(); }, // escaped percent/backslash is not a comment start
                _ => {},
            }
        }
    }
    comment
}

fn ends_control_word(output: &str) -> bool {
    let bytes = output.as_bytes();
    let mut index = bytes.len();
    while index > 0 && bytes[index - 1].is_ascii_alphabetic() { index -= 1; }
    if index == bytes.len() { return false; }
    let word_start = index;
    while index > 0 && bytes[index - 1] == b'\\' { index -= 1; }
    // A doubled backslash is a control symbol, not the start of a control word.
    (word_start - index) % 2 == 1
}

fn escape_text(text: &str, output: &mut String) {
    for ch in text.chars() {
        match ch {
            '\\' => output.push_str("\\textbackslash{}"),
            '^' => output.push_str("\\textasciicircum{}"),
            '~' => output.push_str("\\textasciitilde{}"),
            '{' | '}' | '$' | '&' | '#' | '_' | '%' => { output.push('\\'); output.push(ch); }
            _ => output.push(ch),
        }
    }
}

#[derive(Debug, Default, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TreeFacts {
    pub nodes: usize,
    pub rows: usize,
    pub fractions: usize,
    pub radicals: usize,
    pub fences: usize,
    pub accents: usize,
    pub annotations: usize,
    pub large_operators: usize,
    pub general_operators: usize,
    pub custom_symbols: usize,
    pub grids: usize,
    pub grid_rows: usize,
    pub grid_columns: usize,
    pub styles: usize,
    pub spaces: usize,
    pub boxes: usize,
    pub colors: usize,
    pub color_changes: usize,
    pub phantoms: usize,
    pub scripts: usize,
    pub text_runs: usize,
    pub empty_slots: usize,
    pub unparsed_sources: usize,
}

pub fn validate_tree(root: &Node) -> EngineResult<TreeFacts> {
    fn visit(node: &Node, depth: usize, seen: &mut BTreeSet<NodeId>, facts: &mut TreeFacts, grid_child: bool) -> EngineResult<()> {
        if depth > MAX_DEPTH || facts.nodes + facts.grid_columns >= MAX_NODES {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Math tree depth or node budget exhausted"));
        }
        if node.id.0 == 0 || !seen.insert(node.id) {
            return Err(EngineError::new(ErrorCode::InvalidInput, "Math tree contains duplicate or zero node identities"));
        }
        facts.nodes += 1;
        match &node.kind {
            NodeKind::Row { .. } => facts.rows += 1,
            NodeKind::Grid { environment, columns, array_spec, rows, .. } => {
                facts.grids += 1;
                if rows.is_empty() || rows.len() > crate::grids::MAX_GRID_ROWS
                    || columns.len() > crate::grids::MAX_GRID_COLUMNS || !environment.accepts_columns(columns.len()) {
                    return Err(EngineError::new(ErrorCode::ResourceLimit, "Invalid or oversized contained grid shape"));
                }
                match (environment,array_spec) {
                    (crate::grids::GridKind::Array,Some(spec))=>{spec.validate()?;if spec.alignments.len()!=columns.len(){return Err(EngineError::new(ErrorCode::InvalidInput,"Array column specification does not match its grid width"));}},
                    (crate::grids::GridKind::Array,None)=>return Err(EngineError::new(ErrorCode::InvalidInput,"Array grid requires an explicit column specification")),
                    (_,Some(_))=>return Err(EngineError::new(ErrorCode::InvalidInput,"Only array grids may carry an array column specification")),
                    (_,None)=>{},
                }
                for column in columns {
                    if column.0 == 0 || !seen.insert(NodeId(column.0)) {
                        return Err(EngineError::new(ErrorCode::InvalidInput, "Duplicate or zero grid column identity"));
                    }
                    facts.grid_columns += 1;
                }
                for (i,row) in rows.iter().enumerate() {
                    let NodeKind::GridRow { cells, row_break, .. } = &row.kind else {
                        return Err(EngineError::new(ErrorCode::InvalidInput, "Grid children must be mathematical rows"));
                    };
                    if cells.len() != columns.len() || (i+1 < rows.len()) != !row_break.is_empty() {
                        return Err(EngineError::new(ErrorCode::InvalidInput, "Grid cell count or row separator is inconsistent"));
                    }
                }
            }
            NodeKind::GridRow { cells, separators, row_break } => {
                if !grid_child { return Err(EngineError::new(ErrorCode::InvalidInput, "Mathematical grid row has no grid owner")); }
                facts.grid_rows += 1;
                if cells.is_empty() || cells.len() > crate::grids::MAX_GRID_COLUMNS || separators.len()+1 != cells.len()
                    || cells.iter().any(|c|!matches!(c.kind,NodeKind::Row{grouping:Grouping::Bare,..}))
                    || separators.iter().any(|s|s!="&") || (!row_break.is_empty() && row_break!="\\\\") {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Malformed mathematical grid row"));
                }
            }
            NodeKind::Fraction { command, numerator, denominator, .. } => {
                facts.fractions += 1;
                if crate::fractions::spec(command).is_none() {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Unknown fraction presentation command"));
                }
                if !numerator.is_row() || !denominator.is_row() {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Fraction arguments must be editable rows"));
                }
            }
            NodeKind::Radical { degree, radicand, .. } => {
                facts.radicals += 1;
                if !radicand.is_row() || degree.as_ref().is_some_and(|d| !d.is_row()) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Radical arguments must be editable rows"));
                }
            }
            NodeKind::Fence { body, .. } => {
                facts.fences += 1;
                if !body.is_row() { return Err(EngineError::new(ErrorCode::InvalidInput, "Fence body must be an editable row")); }
            }
            NodeKind::MiddleDelimiter { delimiter, .. } => {
                if crate::delimiters::glyph(delimiter).is_none() {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Unknown middle delimiter"));
                }
            }
            NodeKind::Accent { body, command, .. } => {
                facts.accents += 1;
                if !body.is_row() || crate::accents::spec(command).is_none() { return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid accent command or body")); }
            }
            NodeKind::Annotation { annotation, base, above, command, .. } => {
                facts.annotations += 1;
                if !base.is_row() || !annotation.is_row() || crate::accents::annotation_above(command) != Some(*above) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid annotation structure"));
                }
            }
            NodeKind::Scripts { slots, .. } => {
                facts.scripts += 1;
                if slots.iter().any(|slot| !slot.argument.is_row()) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Script arguments must be editable rows"));
                }
            }
            NodeKind::LargeOperator { command, parts } => {
                facts.large_operators+=1;
                if crate::operators::spec(command).is_none() || parts.len()>crate::operators::MAX_OPERATOR_PARTS {
                    return Err(EngineError::new(ErrorCode::InvalidInput,"Invalid large operator command or tail"));
                }
                for part in parts {match part {
                    crate::operators::OperatorPart::Limits{command,..} if crate::operators::LimitMode::from_command(command).is_none()=>
                        return Err(EngineError::new(ErrorCode::InvalidInput,"Invalid operator placement directive")),
                    crate::operators::OperatorPart::Script{slot} if !slot.argument.is_row()=>
                        return Err(EngineError::new(ErrorCode::InvalidInput,"Operator scripts require editable rows")),
                    _=>{},
                }}
            }
            NodeKind::CustomSymbol { command, asset_sha256, asset, parts } => {
                facts.custom_symbols += 1;
                if command != &format!("\\{}",asset.command()) || asset_sha256 != asset.sha256()
                    || !crate::operators::valid_parts(parts) || (asset.role()!=AtomClass::Operator && !parts.is_empty()) {
                    return Err(EngineError::new(ErrorCode::InvalidInput,"Custom symbol binding/role differs from its immutable resource"));
                }
            }
            NodeKind::GeneralOperator { command, body, parts, .. } => {
                facts.general_operators += 1;
                if command != "\\mathop" || !body.is_row() || !crate::operators::valid_parts(parts) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid general operator body or tail"));
                }
            }
            NodeKind::MacroCall { command, macro_kind, before_arguments, arguments } => {
                let Some(spec)=crate::macros::spec(command) else {
                    return Err(EngineError::new(ErrorCode::InvalidInput,"Unknown built-in macro command"));
                };
                if spec.kind!=*macro_kind || arguments.len()!=spec.arity || before_arguments.len()!=spec.arity
                    || arguments.iter().any(|argument|!argument.is_row()) {
                    return Err(EngineError::new(ErrorCode::InvalidInput,"Malformed built-in macro call"));
                }
            }
            NodeKind::Style { command, variant, body, .. } => {
                facts.styles += 1;
                if !body.is_row() || crate::styles::command_variant(command) != Some(*variant) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid mathematical style structure"));
                }
            }
            NodeKind::Color { spec,body,.. } => {
                facts.colors+=1;spec.validate()?;
                if !body.is_row(){return Err(EngineError::new(ErrorCode::InvalidInput,"Color body must be an editable row"));}
            }
            NodeKind::ColorChange { spec } => {facts.color_changes+=1;spec.validate()?;}
            NodeKind::StyleChange { command, style } => {
                if crate::styles::style_change(command) != Some(*style) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid math style declaration"));
                }
            }
            NodeKind::Boxed { body, .. } => {
                facts.boxes += 1;
                if !body.is_row() {return Err(EngineError::new(ErrorCode::InvalidInput,"Boxed body must be an editable row"));}
            }
            NodeKind::Space { command, width_mu } => {
                facts.spaces += 1;
                if crate::spacing::space_mu(command) != Some(*width_mu) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid explicit math spacing command"));
                }
            }
            NodeKind::Phantom { command, axis, body, .. } => {
                facts.phantoms += 1;
                if !body.is_row() || crate::spacing::phantom_axis(command) != Some(*axis) {
                    return Err(EngineError::new(ErrorCode::InvalidInput, "Invalid phantom structure"));
                }
            }
            NodeKind::Text { .. } => facts.text_runs += 1,
            NodeKind::EmptySlot { .. } => facts.empty_slots += 1,
            NodeKind::UnparsedSource { .. } => facts.unparsed_sources += 1,
            _ => {}
        }
        for child in node.children() { visit(child, depth + 1, seen, facts, matches!(node.kind,NodeKind::Grid{..}))?; }
        Ok(())
    }
    let mut facts = TreeFacts::default();
    visit(root, 0, &mut BTreeSet::new(), &mut facts, false)?;
    Ok(facts)
}
