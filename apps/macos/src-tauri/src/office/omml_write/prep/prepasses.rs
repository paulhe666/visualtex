//! The MathML pre-passes of WordOmmlConverter.TransformMathMlToOmml, one
//! function per C# method, in chain order:
//!
//! ```text
//! ValidateMathMlForOmml
//! s01 NormalizeMathTypeBinomialPiles        s07 NormalizeMathMlAccents
//! s02 NormalizeMathJaxBinomialDelimiters    s08 NormalizeNaryArguments
//! s03 NormalizePlainBracketPairs            s09 NormalizeMathMlTokenStyleAuthority
//! s04 NormalizeFencedMathMlTables           s10 MathMlColorContract.MaterializeInheritedColors
//! s05 NormalizeNestedEmptyBaseScripts       s11 ProjectMathMlSpacesForOmml
//! s06 NormalizeMathMlPrescripts             s12 ReplaceExtendedIntegralsWithOfficePlaceholders
//! IsBlockMathMl
//! ```
//!
//! Every pass re-reads its input and writes it back with XLinq's
//! `ToString(SaveOptions.DisableFormatting)` (see xlinq.rs), so whitespace
//! text nodes and CRLF newlines carry from pass to pass exactly as on Windows.
//! Copies (`new XElement(other)`, adding a node that already has a parent)
//! are real deep copies here as well: later edits of the original never reach
//! the copy, which several passes observably depend on.

use std::collections::HashMap;

use crate::office::omml_write::prep::contracts::{
    describe_omml_space, effective_color, has_tex_class, is_projectable_endpoint, resolve_effective_variant,
    try_get_row, MathSpaceOmmlCapability,
};
use crate::office::omml_write::prep::dom::{Doc, Name, NodeId, MATHML_NS};
use crate::office::omml_write::prep::netstr::{
    contains_ignore_case, contains_unicode_letter, eq_ignore_case, is_null_or_empty, is_white_space, len16, trim,
};
use crate::office::omml_write::prep::xlinq::{
    add_after_self, add_node, element_to_string, elements_after_self, formula_reader, has_elements,
    is_namespace_declaration, new_element_with, nodes, parse, remove, replace_nodes, replace_with, BLOCK_READER,
    PREPASS_READER, XDOCUMENT_PARSE_PRESERVE,
};

/// WordOmmlConverter.NaryCharacters.
pub const NARY_CHARACTERS: &str =
    "\u{2211}\u{220F}\u{2210}\u{222B}\u{222C}\u{222D}\u{222E}\u{222F}\u{2230}\u{2231}\u{2232}\u{2233}\u{2A0B}\u{2A0C}\u{2A0D}\u{2A0E}\u{2A0F}\u{2A10}\u{2A11}\u{2A12}\u{2A13}\u{2A14}\u{2A15}\u{2A16}\u{2A17}\u{2A18}\u{2A19}\u{2A1A}\u{2A1B}\u{2A1C}\u{22C0}\u{22C1}\u{22C2}\u{22C3}";

/// WordOmmlConverter.ExtendedIntegralCharacters.
pub const EXTENDED_INTEGRAL_CHARACTERS: &str =
    "\u{222F}\u{2230}\u{2231}\u{2232}\u{2233}\u{2A0B}\u{2A0C}\u{2A0D}\u{2A0E}\u{2A0F}\u{2A10}\u{2A11}\u{2A12}\u{2A13}\u{2A14}\u{2A15}\u{2A16}\u{2A17}\u{2A18}\u{2A19}\u{2A1A}\u{2A1B}\u{2A1C}";

/// WordOmmlConverter.OmmlSpaceTextAttribute.
pub const OMML_SPACE_TEXT_ATTRIBUTE: &str = "data-visualtex-omml-space";

/// WordOmmlConverter.OmmlRunRoleAttribute.
pub const OMML_RUN_ROLE_ATTRIBUTE: &str = "data-visualtex-omml-run-role";

fn is_nary_character(c: char) -> bool {
    NARY_CHARACTERS.contains(c)
}

/// `!IsNullOrEmpty(value) && value.All(c => NaryCharacters.IndexOf(c) >= 0)`.
/// A surrogate unit is never in the set, so per char is the same test.
fn is_nary_text(value: &str) -> bool {
    !value.is_empty() && value.chars().all(is_nary_character)
}

fn is_mathml(doc: &Doc, element: NodeId, local: &str) -> bool {
    doc.is(element, MATHML_NS, local)
}

fn is_mathml_any(doc: &Doc, element: NodeId, locals: &[&str]) -> bool {
    doc.name(element).is_some_and(|name| name.ns == MATHML_NS && locals.contains(&name.local.as_str()))
}

fn new_mathml(doc: &mut Doc, local: &str) -> NodeId {
    doc.new_element(Name::new(MATHML_NS, local))
}

/// `new XElement(mml + "mfenced", new XAttribute("open", open),
/// new XAttribute("close", close), new XAttribute("separators", ""), operand)`.
fn new_mfenced(doc: &mut Doc, open: &str, close: &str, operand: NodeId) -> NodeId {
    let fenced = new_mathml(doc, "mfenced");
    doc.set_attr(fenced, "open", open);
    doc.set_attr(fenced, "close", close);
    doc.set_attr(fenced, "separators", "");
    add_node(doc, fenced, operand);
    fenced
}

/// `new XElement(mml + "mrow", new XElement(mml + "mspace", new XAttribute("width", "0em")))`.
fn new_empty_operand(doc: &mut Doc) -> NodeId {
    let space = new_mathml(doc, "mspace");
    doc.set_attr(space, "width", "0em");
    new_element_with(doc, MATHML_NS, "mrow", &[space])
}

fn text_is_significant(doc: &Doc, node: NodeId) -> bool {
    doc.text(node).is_some_and(|text| !is_white_space(text))
}

// ---- ValidateMathMlForOmml ------------------------------------------------------

/// WordOmmlConverter.ValidateMathMlForOmml.
pub fn validate_mathml_for_omml(mathml: &str) -> Result<(), String> {
    let document = parse(mathml, &formula_reader("MathML"))?;
    if !document.root().is_some_and(|root| is_mathml(&document.doc, root, "math")) {
        return Err("The OMML source must have a Presentation MathML math root.".to_string());
    }
    if !document.descendants_named(MATHML_NS, "merror").is_empty() {
        return Err("MathML contains an error node and cannot be inserted as OMML.".to_string());
    }
    Ok(())
}

// ---- s01 NormalizeMathTypeBinomialPiles -------------------------------------------

/// WordOmmlConverter.NormalizeMathTypeBinomialPiles.
pub fn normalize_math_type_binomial_piles(mathml: &str) -> Result<String, String> {
    if is_white_space(mathml) || !contains_ignore_case(mathml, "data-mtef-pile") {
        return Ok(mathml.to_string());
    }
    let mut document = parse(mathml, &XDOCUMENT_PARSE_PRESERVE)?;
    for table in document.descendants_named(MATHML_NS, "mtable") {
        let doc = &mut document.doc;
        if !eq_ignore_case(doc.attr(table, "data-mtef-pile"), "true") {
            continue;
        }
        let rows = doc.elements_named(table, MATHML_NS, "mtr");
        if rows.len() != 2 {
            continue;
        }
        let cells: Vec<Vec<NodeId>> = rows.iter().map(|&row| doc.elements_named(row, MATHML_NS, "mtd")).collect();
        if cells.iter().any(|row| row.len() != 1) {
            continue;
        }
        // CopyCellAsRow(cell) => new XElement(mml + "mrow", new XElement(cell).Nodes())
        let copy_cell_as_row = |doc: &mut Doc, cell: NodeId| {
            let copy = doc.deep_clone(cell);
            let content = nodes(doc, copy);
            new_element_with(doc, MATHML_NS, "mrow", &content)
        };
        let numerator = copy_cell_as_row(doc, cells[0][0]);
        let denominator = copy_cell_as_row(doc, cells[1][0]);
        let fraction = new_mathml(doc, "mfrac");
        doc.set_attr(fraction, "linethickness", "0");
        add_node(doc, fraction, numerator);
        add_node(doc, fraction, denominator);
        replace_with(doc, table, fraction)?;
    }
    document.to_string()
}

// ---- s02 NormalizeMathJaxBinomialDelimiters -----------------------------------------

/// WordOmmlConverter.NormalizeMathJaxBinomialDelimiters.
pub fn normalize_mathjax_binomial_delimiters(mathml: &str) -> Result<String, String> {
    if is_white_space(mathml) || !mathml.contains("linethickness") {
        return Ok(mathml.to_string());
    }
    let mut document = parse(mathml, &XDOCUMENT_PARSE_PRESERVE)?;

    // TryGetWrappedFence
    fn wrapped_fence(doc: &Doc, element: NodeId, role: &str) -> Option<String> {
        if !is_mathml(doc, element, "mrow") || !has_tex_class(doc, element, role) {
            return None;
        }
        let inner = doc.elements(element);
        if inner.len() != 1
            || !is_mathml(doc, inner[0], "mo")
            || has_elements(doc, inner[0])
            || is_white_space(&doc.value(inner[0]))
        {
            return None;
        }
        Some(trim(&doc.value(inner[0])).to_string())
    }

    let rows: Vec<NodeId> = document.descendants_named(MATHML_NS, "mrow").into_iter().rev().collect();
    let doc = &mut document.doc;
    for row in rows {
        let children = doc.elements(row);
        if children.len() != 3 {
            continue;
        }
        let Some(open) = wrapped_fence(doc, children[0], "OPEN") else { continue };
        let Some(close) = wrapped_fence(doc, children[2], "CLOSE") else { continue };
        let fraction = children[1];
        if !is_mathml(doc, fraction, "mfrac") {
            continue;
        }
        let thickness = trim(doc.attr(fraction, "linethickness").unwrap_or(""));
        if !thickness.starts_with('0') || doc.elements(fraction).len() != 2 {
            continue;
        }
        let copy = doc.deep_clone(fraction);
        let fenced = new_mfenced(doc, &open, &close, copy);
        replace_nodes(doc, row, &[fenced]);
    }
    document.to_string()
}

// ---- s03 NormalizePlainBracketPairs ---------------------------------------------------

/// WordOmmlConverter.AdaptiveBracketPairs.
const ADAPTIVE_BRACKET_PAIRS: &[(&str, &str)] = &[
    ("(", ")"),
    ("[", "]"),
    ("{", "}"),
    ("\u{27E8}", "\u{27E9}"),
    ("\u{230A}", "\u{230B}"),
    ("\u{2308}", "\u{2309}"),
];

/// A recorded bracket pair of NormalizePlainBracketPairs.Rebuild.
struct BracketPair {
    close: usize,
    open: String,
    close_text: String,
    in_script: bool,
}

/// NormalizePlainBracketPairs.Token.
fn bracket_token(doc: &Doc, element: NodeId, role: &str) -> Option<String> {
    let mut token = element;
    if is_mathml(doc, element, "mrow") {
        let inner = doc.elements(element);
        if inner.len() != 1 {
            return None;
        }
        let wrapper_class = doc.attr(element, "data-mjx-texclass");
        if !is_null_or_empty(wrapper_class) && !eq_ignore_case(wrapper_class, role) {
            return None;
        }
        token = inner[0];
    }
    if !is_mathml(doc, token, "mo") || has_elements(doc, token) {
        return None;
    }
    if doc.attr(token, "mathvariant").is_some() {
        return None;
    }
    let token_class = doc.attr(token, "data-mjx-texclass");
    if !is_null_or_empty(token_class) && !eq_ignore_case(token_class, role) {
        return None;
    }
    Some(trim(&doc.value(token)).to_string())
}

/// NormalizePlainBracketPairs.ScriptBaseClose (script names match by local name).
fn script_base_close(doc: &Doc, element: NodeId) -> Option<String> {
    if !matches!(doc.local(element), "msup" | "msub" | "msubsup") {
        return None;
    }
    let first = doc.elements(element).first().copied()?;
    bracket_token(doc, first, "CLOSE")
}

fn brackets_match(open: &str, close: &str) -> bool {
    ADAPTIVE_BRACKET_PAIRS.iter().any(|&(o, c)| o == open && c == close)
}

fn is_open_bracket(value: Option<&str>) -> bool {
    value.is_some_and(|value| ADAPTIVE_BRACKET_PAIRS.iter().any(|&(open, _)| open == value))
}

/// NormalizePlainBracketPairs.Rebuild; None when no pair was found (the C#
/// returns the very same list, which the caller detects with ReferenceEquals).
fn rebuild_brackets(doc: &mut Doc, items: &[NodeId]) -> Result<Option<Vec<NodeId>>, String> {
    let mut pairs: HashMap<usize, BracketPair> = HashMap::new();
    let mut stack: Vec<(usize, String)> = Vec::new();
    for (index, &item) in items.iter().enumerate() {
        let open = bracket_token(doc, item, "OPEN");
        if is_open_bracket(open.as_deref()) {
            stack.push((index, open.unwrap()));
            continue;
        }
        let script_close = script_base_close(doc, item);
        let in_script = script_close.is_some();
        let close = match script_close {
            Some(close) => Some(close),
            None => bracket_token(doc, item, "CLOSE"),
        };
        let Some(close) = close else { continue };
        let Some((top_index, top_open)) = stack.last() else { continue };
        if !brackets_match(top_open, &close) {
            continue;
        }
        let (top_index, top_open) = (*top_index, top_open.clone());
        stack.pop();
        // An empty group has no operand to build; leave it as text.
        if !in_script && index == top_index + 1 {
            continue;
        }
        if in_script && index == top_index + 1 {
            continue;
        }
        pairs.insert(top_index, BracketPair { close: index, open: top_open, close_text: close, in_script });
    }
    if pairs.is_empty() {
        return Ok(None);
    }
    Ok(Some(build_brackets(doc, items, &pairs, 0, items.len())?))
}

/// NormalizePlainBracketPairs.Rebuild.Build.
fn build_brackets(
    doc: &mut Doc,
    items: &[NodeId],
    pairs: &HashMap<usize, BracketPair>,
    start: usize,
    end: usize,
) -> Result<Vec<NodeId>, String> {
    let mut result = Vec::new();
    let mut index = start;
    while index < end {
        let Some(pair) = pairs.get(&index) else {
            result.push(items[index]);
            index += 1;
            continue;
        };
        let body = build_brackets(doc, items, pairs, index + 1, pair.close)?;
        let operand = if body.len() == 1 && is_mathml(doc, body[0], "mtable") {
            doc.deep_clone(body[0])
        } else {
            let copies: Vec<NodeId> = body.iter().map(|&node| doc.deep_clone(node)).collect();
            new_element_with(doc, MATHML_NS, "mrow", &copies)
        };
        let fenced = new_mfenced(doc, &pair.open, &pair.close_text, operand);
        if pair.in_script {
            let script = doc.deep_clone(items[pair.close]);
            let first = doc.elements(script)[0];
            replace_with(doc, first, fenced)?;
            result.push(script);
        } else {
            result.push(fenced);
        }
        index = pair.close + 1;
    }
    Ok(result)
}

/// WordOmmlConverter.NormalizePlainBracketPairs.
pub fn normalize_plain_bracket_pairs(mathml: &str) -> Result<String, String> {
    if is_white_space(mathml) {
        return Ok(mathml.to_string());
    }
    // MathType read-back (data-mtef-*) keeps its literal brackets.
    if mathml.contains("data-mtef-") {
        return Ok(mathml.to_string());
    }
    let mut document = parse(mathml, &XDOCUMENT_PARSE_PRESERVE)?;
    const CONTAINER_NAMES: &[&str] = &["math", "mrow", "mstyle", "mtd", "msqrt", "mpadded", "mphantom"];
    let containers: Vec<NodeId> = document
        .descendants()
        .into_iter()
        .filter(|&element| is_mathml_any(&document.doc, element, CONTAINER_NAMES))
        .rev()
        .collect();
    let doc = &mut document.doc;
    let mut changed = false;
    for container in containers {
        let children = doc.elements(container);
        if children.len() < 2 {
            continue;
        }
        let last = children[children.len() - 1];
        // Explicit \left...\right rows are already one native group; only
        // their body can hold further bare brackets.
        let explicit_row = try_get_row(doc, container).is_some()
            || (has_tex_class(doc, container, "INNER")
                && has_tex_class(doc, children[0], "OPEN")
                && has_tex_class(doc, last, "CLOSE"));
        let (head, middle, tail): (Vec<NodeId>, Vec<NodeId>, Vec<NodeId>) = if explicit_row {
            (vec![children[0]], children[1..children.len() - 1].to_vec(), vec![last])
        } else {
            (Vec::new(), children.clone(), Vec::new())
        };
        let Some(rebuilt) = rebuild_brackets(doc, &middle)? else { continue };
        let content: Vec<NodeId> = head.into_iter().chain(rebuilt).chain(tail).collect();
        replace_nodes(doc, container, &content);
        changed = true;
    }
    if changed { document.to_string() } else { Ok(mathml.to_string()) }
}

// ---- s04 NormalizeFencedMathMlTables ----------------------------------------------------

const FALLBACK_FENCE_CHARACTERS: &[&str] = &[
    "(", ")", "[", "]", "{", "}", "|", "\u{2016}", "\u{2308}", "\u{2309}", "\u{230A}", "\u{230B}", "\u{27E8}",
    "\u{27E9}", "/", "\\", "\u{2191}", "\u{2193}", "\u{2195}", "\u{21D1}", "\u{21D3}", "\u{21D5}",
];

/// NormalizeFencedMathMlTables.IsLegacyTableFence.
fn is_legacy_table_fence(doc: &Doc, element: NodeId, expected_tex_class: &str) -> bool {
    if !is_mathml(doc, element, "mo") || !is_projectable_endpoint(doc, element) {
        return false;
    }
    if has_tex_class(doc, element, expected_tex_class) {
        return true;
    }
    if eq_ignore_case(doc.attr(element, "fence"), "true") || eq_ignore_case(doc.attr(element, "stretchy"), "true") {
        return true;
    }
    FALLBACK_FENCE_CHARACTERS.contains(&doc.value(element).as_str())
}

/// WordOmmlConverter.NormalizeFencedMathMlTables.
pub fn normalize_fenced_mathml_tables(mathml: &str) -> Result<String, String> {
    if is_white_space(mathml) || !mathml.contains("mrow") {
        return Ok(mathml.to_string());
    }
    let mut document = parse(mathml, &XDOCUMENT_PARSE_PRESERVE)?;
    let rows: Vec<NodeId> = document.descendants_named(MATHML_NS, "mrow").into_iter().rev().collect();
    let doc = &mut document.doc;
    for row in rows {
        let children = doc.elements(row);
        if children.len() < 2 {
            continue;
        }
        let first = children[0];
        let last = children[children.len() - 1];
        let explicit_fence = try_get_row(doc, row).is_some();
        let table_body = children.len() == 3 && is_mathml(doc, children[1], "mtable");
        let legacy_table_fence =
            table_body && is_legacy_table_fence(doc, first, "OPEN") && is_legacy_table_fence(doc, last, "CLOSE");
        if !explicit_fence && !legacy_table_fence {
            continue;
        }
        let open = doc.value(first);
        let close = doc.value(last);
        if open.is_empty() && close.is_empty() {
            continue;
        }
        let body_nodes: Vec<NodeId> = nodes(doc, row).into_iter().filter(|&node| node != first && node != last).collect();
        let operand = if table_body {
            doc.deep_clone(children[1])
        } else {
            new_element_with(doc, MATHML_NS, "mrow", &body_nodes)
        };
        let fenced = new_mfenced(doc, &open, &close, operand);
        replace_nodes(doc, row, &[fenced]);
    }
    document.to_string()
}

// ---- s05 NormalizeNestedEmptyBaseScripts ---------------------------------------------------

/// NormalizeNestedEmptyBaseScripts.IsEmptyMathNode.
fn is_empty_math_node(doc: &Doc, element: NodeId) -> bool {
    if is_mathml(doc, element, "mspace") {
        return false;
    }
    if doc.child_nodes(element).iter().any(|&node| text_is_significant(doc, node)) {
        return false;
    }
    doc.elements(element).iter().all(|&child| is_empty_math_node(doc, child))
}

/// NormalizeNestedEmptyBaseScripts.IsOnlyContentOfOuterScriptArgument.
fn is_only_content_of_outer_script_argument(doc: &Doc, candidate: NodeId) -> bool {
    const TRANSPARENT_WRAPPERS: &[&str] = &["mrow", "mstyle", "mpadded", "mphantom", "semantics"];
    let mut current = candidate;
    while let Some(parent) = doc.parent(current) {
        if is_mathml_any(doc, parent, TRANSPARENT_WRAPPERS) {
            if doc.elements(parent).iter().any(|&sibling| sibling != current && !is_empty_math_node(doc, sibling)) {
                return false;
            }
            if doc.child_nodes(parent).iter().any(|&node| text_is_significant(doc, node)) {
                return false;
            }
            current = parent;
            continue;
        }
        if !is_mathml_any(doc, parent, &["msub", "msup", "msubsup"]) {
            return false;
        }
        let position = doc.elements(parent).iter().position(|&child| child == current);
        return position.is_some_and(|position| position >= 1);
    }
    false
}

/// WordOmmlConverter.NormalizeNestedEmptyBaseScripts.
pub fn normalize_nested_empty_base_scripts(mathml: &str) -> Result<String, String> {
    if is_white_space(mathml) {
        return Ok(mathml.to_string());
    }
    let mut document = parse(mathml, &PREPASS_READER)?;
    let scripts: Vec<NodeId> = document
        .descendants()
        .into_iter()
        .filter(|&element| is_mathml_any(&document.doc, element, &["msub", "msup"]))
        .rev()
        .collect();
    let doc = &mut document.doc;
    for script in scripts {
        let children = doc.elements(script);
        if children.len() < 2
            || !is_empty_math_node(doc, children[0])
            || !is_only_content_of_outer_script_argument(doc, script)
        {
            continue;
        }
        let copy = doc.deep_clone(children[1]);
        replace_with(doc, script, copy)?;
    }
    document.to_string()
}

// ---- s06 NormalizeMathMlPrescripts ---------------------------------------------------------

/// NormalizeMathMlPrescripts.IsEmptyBase.
fn is_empty_base(doc: &Doc, element: NodeId) -> bool {
    is_mathml_any(doc, element, &["mrow", "mi", "mn", "mtext"])
        && doc.value(element).is_empty()
        && doc.elements(element).iter().all(|&child| is_empty_base(doc, child))
}

/// WordOmmlConverter.NormalizeMathMlPrescripts.
pub fn normalize_mathml_prescripts(mathml: &str) -> Result<String, String> {
    const SEQUENCE_NAMES: &[&str] = &["math", "mrow", "mstyle", "mtd", "msqrt", "mphantom", "mpadded", "menclose"];
    const BASE_NAMES: &[&str] = &[
        "mi", "mn", "mrow", "mstyle", "mfrac", "msqrt", "mroot", "msub", "msup", "msubsup", "mmultiscripts", "mover",
        "munder", "munderover", "mfenced",
    ];
    let mut document = parse(mathml, &XDOCUMENT_PARSE_PRESERVE)?;
    let scripts: Vec<NodeId> = document
        .descendants()
        .into_iter()
        .filter(|&element| is_mathml_any(&document.doc, element, &["msub", "msup", "msubsup"]))
        .rev()
        .collect();
    let doc = &mut document.doc;
    for script in scripts {
        let Some(parent) = doc.parent(script) else { continue };
        if !is_mathml_any(doc, parent, SEQUENCE_NAMES) {
            continue;
        }
        let local = doc.local(script).to_string();
        let arguments = doc.elements(script);
        let expected = if local == "msubsup" { 3 } else { 2 };
        if arguments.len() != expected || !is_empty_base(doc, arguments[0]) {
            continue;
        }
        let Some(basis) = elements_after_self(doc, script).first().copied() else { continue };
        if !is_mathml_any(doc, basis, BASE_NAMES) || is_white_space(&doc.value(basis)) {
            continue;
        }
        let sub = if local == "msup" { new_mathml(doc, "none") } else { doc.deep_clone(arguments[1]) };
        let sup = if local == "msub" {
            new_mathml(doc, "none")
        } else {
            doc.deep_clone(arguments[if local == "msubsup" { 2 } else { 1 }])
        };
        if is_white_space(&doc.value(sub)) && is_white_space(&doc.value(sup)) {
            continue;
        }
        let native = new_mathml(doc, "mmultiscripts");
        let attributes: Vec<(Name, String)> =
            doc.attrs(script).iter().filter(|(name, _)| !is_namespace_declaration(name)).cloned().collect();
        for (name, value) in attributes {
            doc.set_attr_ns(native, &name.ns, &name.local, Some(&value));
        }
        let basis_copy = doc.deep_clone(basis);
        let prescripts = new_mathml(doc, "mprescripts");
        for node in [basis_copy, prescripts, sub, sup] {
            add_node(doc, native, node);
        }
        replace_with(doc, script, native)?;
        remove(doc, basis)?;
    }
    document.to_string()
}

// ---- s07 NormalizeMathMlAccents -----------------------------------------------------------

/// NormalizeMathMlAccents.accentCharacters.
fn accent_character(source: &str) -> Option<&'static str> {
    Some(match source {
        "^" => "\u{0302}",
        "~" => "\u{0303}",
        "\u{2192}" => "\u{20D7}",
        "\u{2190}" => "\u{20D6}",
        "\u{2194}" => "\u{20E1}",
        "\u{00AF}" | "\u{203E}" | "\u{2015}" | "\u{02C9}" => "\u{0305}",
        "\u{02D9}" => "\u{0307}",
        "\u{00A8}" => "\u{0308}",
        "\u{02C7}" => "\u{030C}",
        "\u{02D8}" => "\u{0306}",
        "\u{00B4}" => "\u{0301}",
        "`" => "\u{0300}",
        "\u{02DA}" => "\u{030A}",
        _ => return None,
    })
}

/// WordOmmlConverter.NormalizeMathMlAccents.
pub fn normalize_mathml_accents(mathml: &str) -> Result<String, String> {
    if is_white_space(mathml) {
        return Ok(mathml.to_string());
    }
    let mut document = parse(mathml, &PREPASS_READER)?;
    let movers = document.descendants_named(MATHML_NS, "mover");
    let doc = &mut document.doc;
    for mover in movers {
        let children = doc.elements(mover);
        if children.len() != 2 {
            continue;
        }
        let mark = children[1];
        if !is_mathml(doc, mark, "mo") {
            continue;
        }
        if eq_ignore_case(doc.attr(mark, "accent"), "false") || eq_ignore_case(doc.attr(mover, "accent"), "false") {
            continue;
        }
        let Some(combining) = accent_character(&doc.value(mark)) else { continue };
        doc.set_attr(mover, "accent", "true");
        doc.set_attr(mark, "accent", "true");
        doc.remove_attr(mark, "stretchy");
        doc.remove_attr(mark, "data-mjx-pseudoscript");
        doc.set_value(mark, combining);
    }
    document.to_string()
}

// ---- s08 NormalizeNaryArguments ------------------------------------------------------------

const LIMIT_NAMES: &[&str] = &["munder", "mover", "munderover", "msub", "msup", "msubsup"];

/// NormalizeNaryArguments.IsNaryLimit.
fn is_nary_limit(doc: &Doc, element: NodeId) -> bool {
    if !is_mathml_any(doc, element, LIMIT_NAMES) {
        return false;
    }
    doc.elements(element)
        .first()
        .is_some_and(|&op| is_mathml(doc, op, "mo") && is_nary_text(&doc.value(op)))
}

/// WordOmmlConverter.NormalizeNaryArguments.
pub fn normalize_nary_arguments(mathml: &str) -> Result<String, String> {
    if is_white_space(mathml) {
        return Ok(mathml.to_string());
    }
    let mut document = parse(mathml, &PREPASS_READER)?;

    // Bare n-ary operators become a synthetic limit structure.
    let operators: Vec<NodeId> = document
        .descendants_named(MATHML_NS, "mo")
        .into_iter()
        .filter(|&element| {
            let doc = &document.doc;
            is_nary_text(&doc.value(element))
                && doc.parent(element).is_none_or(|parent| !is_mathml_any(doc, parent, LIMIT_NAMES))
        })
        .collect();
    let doc = &mut document.doc;
    for op in operators {
        let has_argument = !elements_after_self(doc, op).is_empty();
        let copy = doc.deep_clone(op);
        let empty_row = new_mathml(doc, "mrow");
        let synthetic_limit = new_element_with(doc, MATHML_NS, "msub", &[copy, empty_row]);
        replace_with(doc, op, synthetic_limit)?;
        if !has_argument {
            let operand = new_empty_operand(doc);
            add_after_self(doc, synthetic_limit, operand)?;
        }
    }

    // Fold every consecutive n-ary chain from right to left.
    let parents: Vec<NodeId> = match doc.root {
        Some(root) => doc.descendants_and_self(root).into_iter().filter(|&element| has_elements(doc, element)).collect(),
        None => Vec::new(),
    };
    for parent in parents {
        let children = doc.elements(parent);
        let mut chains: Vec<(usize, usize, Option<NodeId>)> = Vec::new();
        let mut start = 0;
        while start < children.len() {
            if !is_nary_limit(doc, children[start]) {
                start += 1;
                continue;
            }
            let mut end = start;
            while end + 1 < children.len() && is_nary_limit(doc, children[end + 1]) {
                end += 1;
            }
            if end > start {
                chains.push((start, end, children.get(end + 1).copied()));
            }
            start = end + 1;
        }
        chains.sort_by(|a, b| b.0.cmp(&a.0));
        for (chain_start, chain_end, chain_operand) in chains {
            let mut operand = match chain_operand {
                None => {
                    let operand = new_empty_operand(doc);
                    add_after_self(doc, children[chain_end], operand)?;
                    operand
                }
                Some(existing) if is_mathml_any(doc, existing, &["mrow", "mstyle"]) => existing,
                Some(existing) => {
                    let operand = new_mathml(doc, "mrow");
                    replace_with(doc, existing, operand)?;
                    add_node(doc, operand, existing);
                    operand
                }
            };
            for index in (chain_start + 1..=chain_end).rev() {
                let inner = children[index];
                let outer = children[index - 1];
                remove(doc, inner)?;
                remove(doc, operand)?;
                let nested_operand = new_element_with(doc, MATHML_NS, "mrow", &[inner, operand]);
                add_after_self(doc, outer, nested_operand)?;
                operand = nested_operand;
            }
        }
    }

    // A flat operand after a limit structure gets its own mrow.
    let limits: Vec<NodeId> = match doc.root {
        Some(root) => {
            doc.descendants_and_self(root).into_iter().filter(|&element| is_mathml_any(doc, element, LIMIT_NAMES)).collect()
        }
        None => Vec::new(),
    };
    for limit in limits {
        let Some(op) = doc.elements(limit).first().copied() else { continue };
        if !is_mathml(doc, op, "mo") || !is_nary_text(&doc.value(op)) {
            continue;
        }
        let Some(argument) = elements_after_self(doc, limit).first().copied() else { continue };
        if is_mathml_any(doc, argument, &["mrow", "mstyle"]) {
            continue;
        }
        let wrapper = new_element_with(doc, MATHML_NS, "mrow", &[argument]);
        replace_with(doc, argument, wrapper)?;
    }
    document.to_string()
}

// ---- s09 NormalizeMathMlTokenStyleAuthority ------------------------------------------------

/// WordOmmlConverter.NormalizeMathMlTokenStyleAuthority.
pub fn normalize_mathml_token_style_authority(mathml: &str) -> Result<String, String> {
    if is_white_space(mathml) {
        return Ok(mathml.to_string());
    }
    let mut document = parse(mathml, &XDOCUMENT_PARSE_PRESERVE)?;
    let tokens: Vec<NodeId> = document
        .descendants()
        .into_iter()
        .filter(|&element| is_mathml_any(&document.doc, element, &["mi", "mo", "mn", "mtext", "ms"]))
        .collect();
    let doc = &mut document.doc;
    for token in tokens {
        let Some(variant) = resolve_effective_variant(doc, token) else { continue };
        doc.set_attr(token, "mathvariant", &variant);
        let local = doc.local(token).to_string();
        let role = if local == "mo" && contains_unicode_letter(&doc.value(token)) {
            "named-operator".to_string()
        } else {
            local
        };
        doc.set_attr(token, OMML_RUN_ROLE_ATTRIBUTE, &role);
    }
    match doc.root {
        Some(root) => element_to_string(doc, root),
        None => Ok(mathml.to_string()),
    }
}

// ---- s10 MathMlColorContract.MaterializeInheritedColors -------------------------------------

/// MathMlColorContract.MaterializeInheritedColors.
pub fn materialize_inherited_colors(mathml: &str) -> Result<String, String> {
    let mut document = parse(mathml, &XDOCUMENT_PARSE_PRESERVE)?;
    let elements = document.descendants();
    let doc = &mut document.doc;
    for element in elements {
        if let Some(color) = effective_color(doc, element) {
            doc.set_attr(element, "mathcolor", &color);
        }
    }
    document.to_string()
}

// ---- s11 ProjectMathMlSpacesForOmml ---------------------------------------------------------

/// WordOmmlConverter.ProjectMathMlSpacesForOmml: (MathML, capabilities).
pub fn project_mathml_spaces_for_omml(mathml: &str) -> Result<(String, Vec<MathSpaceOmmlCapability>), String> {
    if !mathml.contains("mspace") {
        return Ok((mathml.to_string(), Vec::new()));
    }
    let mut document = parse(mathml, &XDOCUMENT_PARSE_PRESERVE)?;
    let spaces = document.descendants_named(MATHML_NS, "mspace");
    let doc = &mut document.doc;
    let mut capabilities = Vec::new();
    for space in spaces {
        // Never trust a caller-supplied projection attribute.
        doc.remove_attr(space, OMML_SPACE_TEXT_ATTRIBUTE);
        let capability = describe_omml_space(
            doc.attr(space, "width"),
            doc.attr(space, "height"),
            doc.attr(space, "depth"),
            doc.attr(space, "linebreak"),
        );
        if capability.can_project() {
            // SetAttributeValue(name, null) would remove; Text is set whenever
            // the capability can project.
            doc.set_attr_ns(space, "", OMML_SPACE_TEXT_ATTRIBUTE, capability.text.as_deref());
        }
        capabilities.push(capability);
    }
    Ok((document.to_string()?, capabilities))
}

// ---- s12 ReplaceExtendedIntegralsWithOfficePlaceholders -------------------------------------

/// WordOmmlConverter.ReplaceExtendedIntegralsWithOfficePlaceholders:
/// (MathML, the n-ary characters in document order).
pub fn replace_extended_integrals_with_office_placeholders(mathml: &str) -> Result<(String, Vec<String>), String> {
    if is_white_space(mathml) {
        return Ok((mathml.to_string(), Vec::new()));
    }
    let mut document = parse(mathml, &PREPASS_READER)?;
    let operators: Vec<NodeId> = document
        .descendants_named(MATHML_NS, "mo")
        .into_iter()
        .filter(|&element| {
            let value = document.doc.value(element);
            len16(&value) == 1 && value.chars().next().is_some_and(is_nary_character)
        })
        .collect();
    let doc = &mut document.doc;
    let characters: Vec<String> = operators.iter().map(|&op| doc.value(op)).collect();
    for op in operators {
        let value = doc.value(op);
        if value.chars().next().is_some_and(|c| EXTENDED_INTEGRAL_CHARACTERS.contains(c)) {
            doc.set_value(op, "\u{222B}");
        }
    }
    Ok((document.to_string()?, characters))
}

// ---- IsBlockMathMl ------------------------------------------------------------------------------

/// WordOmmlConverter.IsBlockMathMl: display="block" on the root
/// (OrdinalIgnoreCase); false whenever the MathML cannot be read.
pub fn is_block_mathml(mathml: &str) -> bool {
    match parse(mathml, &BLOCK_READER) {
        Ok(document) => document.root().is_some_and(|root| eq_ignore_case(document.doc.attr(root, "display"), "block")),
        Err(_) => false,
    }
}

#[cfg(test)]
mod tests {
    //! Paths the Windows corpus does not reach (it has no MathType piles, no
    //! extended integrals, no named colors, no prefixed MathML); expected
    //! values follow the C# code paths step by step.
    use super::*;

    const M: &str = "http://www.w3.org/1998/Math/MathML";

    #[test]
    fn mathtype_piles_become_bar_less_fractions() {
        let input = format!(
            "<math xmlns=\"{M}\"><mrow><mo>(</mo><mtable data-mtef-pile=\"TRUE\" columnalign=\"left\"><mtr><mtd><mi>n</mi></mtd></mtr>\
             <mtr><mtd></mtd></mtr></mtable><mo>)</mo></mrow></math>"
        );
        assert_eq!(
            normalize_math_type_binomial_piles(&input).unwrap(),
            format!(
                "<math xmlns=\"{M}\"><mrow><mo>(</mo><mfrac linethickness=\"0\"><mrow><mi>n</mi></mrow><mrow /></mfrac><mo>)</mo></mrow></math>"
            )
        );
        // A genuine matrix (no private pile marker) and a 2x2 pile stay.
        let matrix = format!("<math xmlns=\"{M}\"><mtable><mtr><mtd><mi>a</mi></mtd></mtr><mtr><mtd><mi>b</mi></mtd></mtr></mtable><!--x--></math>");
        assert_eq!(normalize_math_type_binomial_piles(&matrix).unwrap(), matrix);
        // The marker makes the pass re-serialize even when no pile qualifies.
        let pile = format!("<math xmlns=\"{M}\"><mtable data-mtef-pile=\"true\"><mtr><mtd/><mtd/></mtr><mtr><mtd/></mtr></mtable></math>");
        assert_eq!(
            normalize_math_type_binomial_piles(&pile).unwrap(),
            format!("<math xmlns=\"{M}\"><mtable data-mtef-pile=\"true\"><mtr><mtd /><mtd /></mtr><mtr><mtd /></mtr></mtable></math>")
        );
    }

    #[test]
    fn extended_integrals_use_office_placeholders() {
        let input = format!("<math xmlns=\"{M}\">\n<mo>\u{222F}</mo><mi>x</mi><mo>\u{2211}</mo><mo>\u{222B}\u{222B}</mo><mo>\u{2A0C}</mo></math>");
        let (mathml, characters) = replace_extended_integrals_with_office_placeholders(&input).unwrap();
        assert_eq!(
            mathml,
            format!("<math xmlns=\"{M}\">\r\n<mo>\u{222B}</mo><mi>x</mi><mo>\u{2211}</mo><mo>\u{222B}\u{222B}</mo><mo>\u{222B}</mo></math>")
        );
        assert_eq!(characters, ["\u{222F}", "\u{2211}", "\u{2A0C}"]);
        assert_eq!(replace_extended_integrals_with_office_placeholders(" ").unwrap(), (" ".to_string(), vec![]));
    }

    #[test]
    fn spaces_project_only_exact_widths() {
        let input = format!(
            "<math xmlns=\"{M}\"><mspace data-visualtex-omml-space=\"x\" width=\"1.0em\" height=\"0ex\"/>\
             <mspace width=\"0.167em\" data-visualtex-omml-space=\"y\"/><mspace/><mspace width=\"2em\" linebreak=\"newline\"/></math>"
        );
        let (mathml, capabilities) = project_mathml_spaces_for_omml(&input).unwrap();
        assert_eq!(
            mathml,
            format!(
                "<math xmlns=\"{M}\"><mspace width=\"1.0em\" height=\"0ex\" data-visualtex-omml-space=\"\u{2003}\" />\
                 <mspace width=\"0.167em\" /><mspace data-visualtex-omml-space=\"\" /><mspace width=\"2em\" linebreak=\"newline\" /></math>"
            )
        );
        assert_eq!(capabilities.iter().map(|c| c.can_project()).collect::<Vec<_>>(), [true, false, true, false]);
        assert_eq!(project_mathml_spaces_for_omml("").unwrap().0, "");
    }

    #[test]
    fn colors_are_inherited_and_normalized() {
        let input = format!(
            "<math xmlns=\"{M}\" mathcolor=\"red\"><mi color=\"#0a0\">x</mi><mn mathcolor=\"foo\">1</mn>\
             <mrow mathcolor=\" Transparent \"><mo>+</mo></mrow></math>"
        );
        assert_eq!(
            materialize_inherited_colors(&input).unwrap(),
            format!(
                "<math xmlns=\"{M}\" mathcolor=\"#FF0000\"><mi color=\"#0a0\" mathcolor=\"#00AA00\">x</mi>\
                 <mn mathcolor=\"#FF0000\">1</mn><mrow mathcolor=\"#FF0000\"><mo mathcolor=\"#FF0000\">+</mo></mrow></math>"
            )
        );
    }

    #[test]
    fn prefixed_mathml_keeps_its_prefix() {
        let input = format!(
            "<m:math xmlns:m=\"{M}\"><m:mrow data-mjx-texclass=\"INNER\"><m:mo data-mjx-texclass=\"OPEN\">(</m:mo>\
             <m:mi>x</m:mi><m:mo data-mjx-texclass=\"CLOSE\">)</m:mo></m:mrow></m:math>"
        );
        validate_mathml_for_omml(&input).unwrap();
        assert_eq!(
            normalize_fenced_mathml_tables(&input).unwrap(),
            format!(
                "<m:math xmlns:m=\"{M}\"><m:mrow data-mjx-texclass=\"INNER\"><m:mfenced open=\"(\" close=\")\" separators=\"\">\
                 <m:mrow><m:mi>x</m:mi></m:mrow></m:mfenced></m:mrow></m:math>"
            )
        );
        assert!(is_block_mathml(&format!("<m:math xmlns:m=\"{M}\" display=\" BLOCK\"/>")) == false);
        assert!(is_block_mathml(&format!("<m:math xmlns:m=\"{M}\" display=\"Block\"/>")));
        assert!(!is_block_mathml("<math"));
    }

    #[test]
    fn token_style_writes_the_root_element_only() {
        let input = format!("\n<math xmlns=\"{M}\"><mi mathvariant=\" Bold \">x</mi><mo>sin</mo><mi>\u{2032}</mi><mi>ab</mi><mtext><mi>y</mi></mtext></math>\n");
        assert_eq!(
            normalize_mathml_token_style_authority(&input).unwrap(),
            format!(
                "<math xmlns=\"{M}\"><mi mathvariant=\"bold\" data-visualtex-omml-run-role=\"mi\">x</mi>\
                 <mo mathvariant=\"normal\" data-visualtex-omml-run-role=\"named-operator\">sin</mo>\
                 <mi mathvariant=\"normal\" data-visualtex-omml-run-role=\"mi\">\u{2032}</mi>\
                 <mi mathvariant=\"normal\" data-visualtex-omml-run-role=\"mi\">ab</mi>\
                 <mtext><mi mathvariant=\"italic\" data-visualtex-omml-run-role=\"mi\">y</mi></mtext></math>"
            )
        );
        // XDocument.ToString keeps the top-level whitespace (as CRLF).
        assert_eq!(materialize_inherited_colors(&input).unwrap(), format!("\r\n<math xmlns=\"{M}\"><mi mathvariant=\" Bold \">x</mi><mo>sin</mo><mi>\u{2032}</mi><mi>ab</mi><mtext><mi>y</mi></mtext></math>\r\n"));
    }

    #[test]
    fn validation_and_chain_errors() {
        assert!(validate_mathml_for_omml("<math/>").is_err());
        assert!(validate_mathml_for_omml(&format!("<math xmlns=\"{M}\"><merror/></math>")).is_err());
        assert!(validate_mathml_for_omml(&format!("<!DOCTYPE math><math xmlns=\"{M}\"/>")).is_err());
        // FormulaXml rejects an element at XmlReader.Depth 256.
        let deep = format!("<math xmlns=\"{M}\">{}{}</math>", "<mrow>".repeat(256), "</mrow>".repeat(256));
        assert!(validate_mathml_for_omml(&deep).is_err());
        let deep = format!("<math xmlns=\"{M}\">{}{}</math>", "<mrow>".repeat(255), "</mrow>".repeat(255));
        assert!(validate_mathml_for_omml(&deep).is_ok());
        assert!(crate::office::omml_write::prep::prepare_mathml_for_omml(" \n").is_err());
        assert!(normalize_mathml_prescripts("").is_err());
        // A bare n-ary root cannot get a sibling operand (XDocument throws).
        assert!(normalize_nary_arguments(&format!("<mo xmlns=\"{M}\">\u{2211}</mo>")).is_err());
    }

    #[test]
    fn whole_chain_on_an_extended_integral() {
        let prepared = crate::office::omml_write::prep::prepare_mathml_for_omml(&format!(
            "<math xmlns=\"{M}\" display=\"block\"><msubsup><mo>\u{2233}</mo><mi>a</mi><mi>b</mi></msubsup><mi>f</mi></math>"
        ))
        .unwrap();
        assert_eq!(prepared.nary_characters, ["\u{2233}"]);
        assert!(prepared.is_block);
        assert_eq!(
            prepared.mathml,
            format!(
                "<math xmlns=\"{M}\" display=\"block\"><msubsup><mo mathvariant=\"normal\" data-visualtex-omml-run-role=\"mo\">\u{222B}</mo>\
                 <mi mathvariant=\"italic\" data-visualtex-omml-run-role=\"mi\">a</mi><mi mathvariant=\"italic\" data-visualtex-omml-run-role=\"mi\">b</mi></msubsup>\
                 <mrow><mi mathvariant=\"italic\" data-visualtex-omml-run-role=\"mi\">f</mi></mrow></math>"
            )
        );
    }
}
