//! Port of WordOmmlConverter.TransformOmmlToMathMl and the recovery passes it
//! runs over the projection (WordOmmlConverter.cs).

use super::contracts::{
    contains_ignore_case, dotnet_trim, eq_ignore_case, find_omml_space_text, is_dotnet_digit, is_math_symbol,
    is_null_or_white_space, maximum_omml_space_text_length, normalize_color, resolve_effective_variant,
    to_lower_invariant, try_parse_int,
};
use super::dom::{Doc, Name, NodeId, MATHML_NS, OMML_NS, WORD_NS};
use super::projection;
use super::xml;
use regex::Regex;
use std::collections::HashSet;
use std::sync::LazyLock;

const M: &str = OMML_NS;
const W: &str = WORD_NS;
const ML: &str = MATHML_NS;

const EQUATION_NUMBER_PLACEHOLDER_PREFIX: &str = "981730";
const NATIVE_NUMBER_BOOKMARK_PREFIX: &str = "VTEqNum_";

fn invalid_data(message: &str) -> String {
    format!("InvalidDataException: {message}")
}

fn named(doc: &Doc, nodes: Vec<NodeId>, ns: &str, local: &str) -> Vec<NodeId> {
    nodes.into_iter().filter(|&node| doc.is(node, ns, local)).collect()
}

fn has_elements(doc: &Doc, node: NodeId) -> bool {
    doc.child_nodes(node).iter().any(|&child| doc.is_element(child))
}

fn child_path(doc: &Doc, node: Option<NodeId>, path: &[(&str, &str)]) -> Option<NodeId> {
    let mut current = node?;
    for &(ns, local) in path {
        current = doc.element(current, ns, local)?;
    }
    Some(current)
}

fn new_element(doc: &mut Doc, ns: &str, local: &str) -> NodeId {
    doc.new_element(Name::new(ns, local))
}

/// WordOmmlConverter.TransformOmmlToMathMl(wordOpenXml, display).
pub fn omml_to_mathml(word_open_xml: &str, display: bool) -> Result<String, String> {
    let (mut source, equation) = strip_visual_tex_native_equation_number(word_open_xml)?;
    // Remove only the empty zero-extent representation of a missing native
    // prescript half before reverse conversion; real phantom content is retained.
    remove_empty_prescript_phantoms(&mut source, equation);

    // Project(omml) reads the serialized tree again with ReaderSettings().
    let mut projection_source = Doc::new();
    let projection_root = projection_source.import(&source, equation);
    projection_source.root = Some(projection_root);
    xml::drop_insignificant_whitespace(&mut projection_source, projection_root)?;
    let (mut target, root) = projection::project_tree(&projection_source, Some(projection_root));

    canonicalize_mathml_element(&mut target, root);
    restore_merged_numeric_punctuation_tokens(&mut target, root);
    restore_math_symbol_text_tokens(&mut target, root);
    restore_word_canonical_asterisk_operators(&mut target, root);
    restore_omml_no_bar_fraction_semantics(&source, equation, &mut target, root);
    restore_omml_matrix_column_alignment(&source, equation, &mut target, root);
    restore_omml_accent_semantics(&source, equation, &mut target, root);
    restore_omml_limit_base_semantics(&source, equation, &mut target, root);
    restore_omml_script_base_run_semantics(&source, equation, &mut target, root);
    restore_omml_function_application_runs(&source, equation, &mut target, root);
    restore_omml_explicit_space_tokens(&mut target, root);
    target.set_attr(root, "display", if display { "block" } else { "inline" });
    Ok(xml::to_default_namespace_xml(&target, root, ML))
}

// ---- the sPre phantom cleanup ------------------------------------------------

fn remove_empty_prescript_phantoms(doc: &mut Doc, equation: NodeId) {
    let scripts = named(doc, doc.descendants_and_self(equation), M, "sPre");
    for script in scripts {
        for slot in doc.elements(script) {
            if !(doc.is(slot, M, "sub") || doc.is(slot, M, "sup")) {
                continue;
            }
            let slot_children = doc.elements(slot);
            if slot_children.len() != 1 {
                continue;
            }
            let phantom = slot_children[0];
            if !is_hidden_phantom(doc, phantom) || has_visible_math_content(doc, Some(phantom)) {
                continue;
            }
            let properties = doc.element(phantom, M, "phantPr");
            if ["zeroAsc", "zeroDesc", "zeroWid"].iter().all(|name| is_math_boolean_enabled(doc, properties, M, name)) {
                doc.clear_children(slot);
            }
        }
    }
}

fn is_hidden_phantom(doc: &Doc, element: NodeId) -> bool {
    let show = child_path(doc, Some(element), &[(M, "phantPr"), (M, "show")]);
    doc.is(element, M, "phant") && show.is_some_and(|show| !read_math_boolean_value(doc, show))
}

fn has_visible_math_content(doc: &Doc, slot: Option<NodeId>) -> bool {
    let Some(slot) = slot else { return false };
    doc.descendants_and_self(slot)
        .into_iter()
        .any(|element| (doc.is(element, M, "t") || doc.is(element, W, "t")) && !doc.value(element).is_empty())
}

fn is_math_boolean_enabled(doc: &Doc, properties: Option<NodeId>, ns: &str, local: &str) -> bool {
    match properties.and_then(|properties| doc.element(properties, ns, local)) {
        None => false,
        Some(property) => read_math_boolean_value(doc, property),
    }
}

fn read_math_boolean_value(doc: &Doc, property: NodeId) -> bool {
    match doc.attr_ns(property, M, "val") {
        None => true,
        Some(value) => !matches!(to_lower_invariant(value).as_str(), "0" | "false" | "off"),
    }
}

// ---- equation-number stripping -------------------------------------------------

/// ExtractSingleOMath: the one m:oMath of the input, as a standalone tree.
fn extract_single_omath(omml: &str) -> Result<(Doc, NodeId), String> {
    if is_null_or_white_space(omml) {
        return Err(invalid_data("Office produced an empty OMML transformation."));
    }
    let document = xml::parse(omml, xml::FORMULA_READER).map_err(|error| format!("OMML: {error}"))?;
    let root = document.root.expect("parsed document has a root");
    let equations: Vec<NodeId> = named(&document, document.descendants_and_self(root), M, "oMath")
        .into_iter()
        .take(2)
        .collect();
    if equations.is_empty() {
        return Err(invalid_data("Office MathML conversion did not produce an m:oMath node."));
    }
    if equations.len() != 1 {
        return Err(invalid_data("Expected exactly one OMML equation; the input contains multiple equations."));
    }
    let mut equation_doc = Doc::new();
    let equation = equation_doc.import(&document, equations[0]);
    equation_doc.root = Some(equation);
    xml::normalize_text_line_ends(&mut equation_doc, equation);
    Ok((equation_doc, equation))
}

/// StripVisualTeXNativeEquationNumber (allowUnboundDirectSequence: false).
pub fn strip_visual_tex_native_equation_number(word_open_xml: &str) -> Result<(Doc, NodeId), String> {
    let (mut doc, mut equation) = extract_single_omath(word_open_xml)?;
    let mut removed = 0;
    while let Some((body, separator_index)) = try_resolve_visual_tex_native_equation_number(&doc, equation) {
        equation = strip_equation_number_wrapper_body(
            &mut doc,
            body,
            separator_index,
            "The generated VisualTeX equation-number wrapper contains no formula body.",
        )?;
        doc.root = Some(equation);
        removed += 1;
        if removed > 8 {
            return Err(invalid_data("The VisualTeX native equation-number wrapper is recursively malformed."));
        }
    }
    Ok((doc, equation))
}

fn is_current_omml_content(doc: &Doc, element: NodeId) -> bool {
    let mut chain = vec![element];
    chain.extend(doc.ancestors(element));
    !chain.into_iter().any(|ancestor| {
        let name = doc.name(ancestor).unwrap();
        doc.is(ancestor, W, "del")
            || doc.is(ancestor, W, "moveFrom")
            || (name.ns == W && name.local.ends_with("PrChange"))
    })
}

fn current_omml_text_elements(doc: &Doc, container: NodeId, ns: &str, local: &str) -> Vec<NodeId> {
    named(doc, doc.descendants(container), ns, local)
        .into_iter()
        .filter(|&element| is_current_omml_content(doc, element))
        .collect()
}

fn current_math_run_text_elements(doc: &Doc, run: NodeId) -> Vec<NodeId> {
    current_omml_text_elements(doc, run, M, "t")
        .into_iter()
        .filter(|&text| doc.ancestors(text).into_iter().find(|&a| doc.is(a, M, "r")) == Some(run))
        .collect()
}

fn current_math_run_text(doc: &Doc, run: NodeId) -> String {
    current_math_run_text_elements(doc, run).into_iter().map(|text| doc.value(text)).collect()
}

fn strip_equation_number_wrapper_body(
    doc: &mut Doc,
    body: NodeId,
    separator_index: usize,
    empty_body_message: &str,
) -> Result<NodeId, String> {
    let body_elements = doc.elements(body);
    if separator_index >= body_elements.len() {
        return Err(invalid_data("The equation-number wrapper separator is outside its equation body."));
    }
    let mut formula_elements: Vec<NodeId> =
        body_elements[..separator_index].iter().map(|&element| doc.deep_clone(element)).collect();
    let separator = body_elements[separator_index];
    let separator_text = current_math_run_text(doc, separator);
    if separator_text.encode_utf16().count() > 1 && separator_text.ends_with('#') {
        let merged_tail = doc.deep_clone(separator);
        let tail_texts = current_math_run_text_elements(doc, merged_tail);
        if let Some(&last_text) = tail_texts.last() {
            let value = doc.value(last_text);
            if let Some(stripped) = value.strip_suffix('#') {
                let stripped = stripped.to_string();
                doc.set_value(last_text, &stripped);
            }
            if tail_texts.iter().any(|&text| !doc.value(text).is_empty()) {
                formula_elements.push(merged_tail);
            }
        }
    }
    if formula_elements.is_empty() {
        return Err(invalid_data(empty_body_message));
    }
    let equation = new_element(doc, M, "oMath");
    for element in formula_elements {
        doc.append(equation, element);
    }
    Ok(equation)
}

fn try_resolve_word_native_equation_number_host(doc: &Doc, equation: NodeId) -> Option<(NodeId, usize, NodeId)> {
    let root_children: Vec<NodeId> =
        doc.elements(equation).into_iter().filter(|&element| !doc.is(element, M, "ctrlPr")).collect();
    if root_children.len() != 1 || !doc.is(root_children[0], M, "eqArr") {
        return None;
    }
    let entries = doc.elements_named(root_children[0], M, "e");
    if entries.len() != 1 {
        return None;
    }
    let candidate_body = entries[0];
    let children = doc.elements(candidate_body);
    let mut found = None;
    for index in 0..children.len().saturating_sub(1) {
        let separator = children[index];
        if !doc.is(separator, M, "r") {
            continue;
        }
        if !current_math_run_text(doc, separator).ends_with('#') {
            continue;
        }
        let delimiter = children[index + 1];
        if !doc.is(delimiter, M, "d") || !is_word_native_equation_number_delimiter(doc, delimiter) {
            continue;
        }
        if children[index + 2..].iter().any(|&element| !doc.is(element, M, "ctrlPr")) {
            continue;
        }
        if found.is_some() {
            return None;
        }
        found = Some((candidate_body, index, delimiter));
    }
    found
}

fn is_word_native_equation_number_delimiter(doc: &Doc, delimiter: NodeId) -> bool {
    if doc.elements_named(delimiter, M, "e").len() != 1 {
        return false;
    }
    let properties = doc.element(delimiter, M, "dPr");
    let begin = child_path(doc, properties, &[(M, "begChr")]).and_then(|e| doc.attr_ns(e, M, "val"));
    let end = child_path(doc, properties, &[(M, "endChr")]).and_then(|e| doc.attr_ns(e, M, "val"));
    if begin.is_some_and(|b| !b.is_empty() && b != "(") {
        return false;
    }
    if end.is_some_and(|e| !e.is_empty() && e != ")") {
        return false;
    }
    true
}

fn try_resolve_visual_tex_native_equation_number(doc: &Doc, equation: NodeId) -> Option<(NodeId, usize)> {
    let equation_arrays = doc.elements_named(equation, M, "eqArr");
    if equation_arrays.len() != 1 {
        return None;
    }
    let entries = doc.elements_named(equation_arrays[0], M, "e");
    if entries.len() != 1 {
        return None;
    }
    let candidate_body = entries[0];
    let children = doc.elements(candidate_body);
    for index in 0..children.len().saturating_sub(1) {
        let candidate_separator = children[index];
        if !doc.is(candidate_separator, M, "r") {
            continue;
        }
        if !current_math_run_text(doc, candidate_separator).ends_with('#') {
            continue;
        }
        let candidate_number = children[index + 1];
        if !doc.is(candidate_number, M, "d") {
            continue;
        }
        let number_text: String =
            named(doc, doc.descendants(candidate_number), M, "t").into_iter().map(|t| doc.value(t)).collect();
        let generated_placeholder = number_text.contains(EQUATION_NUMBER_PLACEHOLDER_PREFIX);
        let field_code: String = named(doc, doc.descendants(candidate_number), W, "instrText")
            .into_iter()
            .map(|t| doc.value(t))
            .collect();
        let reference = format!("REF {NATIVE_NUMBER_BOOKMARK_PREFIX}");
        let generated_reference =
            contains_ignore_case(&number_text, &reference) || contains_ignore_case(&field_code, &reference);
        let generated_direct_sequence = match try_resolve_word_native_equation_number_host(doc, equation) {
            Some((_, native_index, native_number)) => {
                native_index == index
                    && native_number == candidate_number
                    && is_visual_tex_direct_sequence_delimiter(doc, candidate_number)
            }
            None => false,
        };
        if !generated_placeholder && !generated_reference && !generated_direct_sequence {
            continue;
        }
        return Some((candidate_body, index));
    }
    None
}

fn is_visual_tex_direct_sequence_delimiter(doc: &Doc, delimiter: NodeId) -> bool {
    // RegexOptions.IgnoreCase | CultureInvariant: only ASCII letters fold to these.
    static SEQ: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r#"\b[Ss][Ee][Qq]\s+(?:"[^"]+"|[^\s\\]+)"#).unwrap());
    static REF: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"\b[Rr][Ee][Ff]\s+[Vv][Tt][Ee][Qq][Nn][Uu][Mm]_").unwrap());
    let mut field_code: String = current_omml_text_elements(doc, delimiter, W, "instrText")
        .into_iter()
        .map(|t| doc.value(t))
        .collect();
    field_code.extend(current_omml_text_elements(doc, delimiter, M, "t").into_iter().map(|t| doc.value(t)));
    if !SEQ.is_match(&field_code) {
        return false;
    }
    !REF.is_match(&field_code)
}

// ---- recovery passes -----------------------------------------------------------

fn canonicalize_mathml_element(doc: &mut Doc, root: NodeId) {
    for element in doc.descendants_and_self(root) {
        let local = doc.local(element).to_string();
        doc.set_name(element, Name::new(ML, &local));
    }
}

fn restore_merged_numeric_punctuation_tokens(doc: &mut Doc, root: NodeId) {
    for number in named(doc, doc.descendants_and_self(root), ML, "mn") {
        let value = doc.value(number);
        if is_null_or_white_space(&value) || (!value.contains(',') && !value.contains(';')) {
            continue;
        }
        if !value.chars().all(|c| is_dotnet_digit(c) || c == ',' || c == ';') {
            continue;
        }
        let mut tokens = Vec::new();
        let mut current = String::new();
        for c in value.chars() {
            if c == ',' || c == ';' {
                if !current.is_empty() {
                    tokens.push(("mn", std::mem::take(&mut current)));
                }
                tokens.push(("mo", c.to_string()));
            } else {
                current.push(c);
            }
        }
        if !current.is_empty() {
            tokens.push(("mn", current));
        }
        if tokens.len() <= 1 {
            continue;
        }
        let token_nodes: Vec<NodeId> = tokens
            .into_iter()
            .map(|(name, text)| {
                let id = new_element(doc, ML, name);
                doc.append_text(id, &text);
                id
            })
            .collect();
        let parent_is_row = doc.parent(number).is_some_and(|parent| doc.is(parent, ML, "mrow"));
        if parent_is_row {
            for token in token_nodes {
                doc.insert_before(number, token);
            }
            doc.detach(number);
        } else {
            let row = new_element(doc, ML, "mrow");
            for token in token_nodes {
                doc.append(row, token);
            }
            doc.replace(number, row);
        }
    }
}

fn contains_only_math_symbols(value: &str) -> bool {
    !value.is_empty() && value.chars().all(is_math_symbol)
}

fn restore_math_symbol_text_tokens(doc: &mut Doc, root: NodeId) {
    for text in named(doc, doc.descendants_and_self(root), ML, "mtext") {
        let value = doc.value(text);
        if value.is_empty() || !contains_only_math_symbols(&value) {
            continue;
        }
        doc.set_name(text, Name::new(ML, "mo"));
    }
}

fn restore_word_canonical_asterisk_operators(doc: &mut Doc, root: NodeId) {
    let tokens: Vec<NodeId> = doc
        .descendants_and_self(root)
        .into_iter()
        .filter(|&e| doc.is(e, ML, "mi") || doc.is(e, ML, "mo"))
        .filter(|&e| !has_elements(doc, e) && doc.value(e) == "*")
        .collect();
    for token in tokens {
        doc.set_name(token, Name::new(ML, "mo"));
        let value = doc.value(token).replace('*', "\u{2217}");
        doc.set_value(token, &value);
    }
}

fn restore_omml_no_bar_fraction_semantics(source: &Doc, equation: NodeId, doc: &mut Doc, root: NodeId) {
    let source_fractions = named(source, source.descendants_and_self(equation), M, "f");
    let target_fractions = named(doc, doc.descendants_and_self(root), ML, "mfrac");
    if source_fractions.is_empty() || source_fractions.len() != target_fractions.len() {
        return;
    }
    for (index, &fraction) in source_fractions.iter().enumerate() {
        let fraction_type = child_path(source, Some(fraction), &[(M, "fPr"), (M, "type")])
            .and_then(|e| source.attr_ns(e, M, "val"));
        if !fraction_type.is_some_and(|t| eq_ignore_case(t, "noBar")) {
            continue;
        }
        doc.set_attr(target_fractions[index], "linethickness", "0");
    }
}

// The stylesheet keeps only the first column's justification of an m:m.
// Record every column so the LaTeX reader can tell `aligned` from a matrix.
fn restore_omml_matrix_column_alignment(source: &Doc, equation: NodeId, doc: &mut Doc, root: NodeId) {
    let source_tables: Vec<NodeId> = source
        .descendants_and_self(equation)
        .into_iter()
        .filter(|&e| source.is(e, M, "m") || source.is(e, M, "eqArr"))
        .collect();
    let target_tables = named(doc, doc.descendants_and_self(root), ML, "mtable");
    if source_tables.is_empty() || source_tables.len() != target_tables.len() {
        return;
    }
    for (index, &table) in source_tables.iter().enumerate() {
        if !source.is(table, M, "m") {
            continue;
        }
        let mut columns: Vec<String> = Vec::new();
        let groups = match source.element(table, M, "mPr") {
            Some(properties) => source.elements_named(properties, M, "mcs"),
            None => Vec::new(),
        };
        for group in groups {
            for column in source.elements_named(group, M, "mc") {
                let properties = source.element(column, M, "mcPr");
                let count = match try_parse_int(
                    child_path(source, properties, &[(M, "count")]).and_then(|e| source.attr_ns(e, M, "val")),
                ) {
                    Some(parsed) if parsed > 0 => parsed.min(64) as usize,
                    _ => 1,
                };
                let justification = to_lower_invariant(dotnet_trim(
                    child_path(source, properties, &[(M, "mcJc")])
                        .and_then(|e| source.attr_ns(e, M, "val"))
                        .unwrap_or("center"),
                ));
                for _ in 0..count {
                    columns.push(justification.clone());
                }
            }
        }
        if columns.len() > 1 {
            doc.set_attr(target_tables[index], "data-vt-columnalign", &columns.join(" "));
        }
    }
}

fn canonical_mathml_accent_character(office_character: Option<&str>) -> String {
    match office_character {
        None | Some("") => "^",
        Some("\u{0302}") => "^",
        Some("\u{0303}") => "~",
        Some("\u{20D7}") => "→",
        Some("\u{20D6}") => "←",
        Some("\u{20E1}") => "↔",
        Some("\u{0305}") => "¯",
        Some("\u{0307}") => "˙",
        Some("\u{0308}") => "¨",
        Some("\u{030C}") => "ˇ",
        Some("\u{0306}") => "˘",
        Some("\u{0301}") => "´",
        Some("\u{0300}") => "`",
        Some("\u{030A}") => "˚",
        Some(other) => other,
    }
    .to_string()
}

fn restore_omml_accent_semantics(source: &Doc, equation: NodeId, doc: &mut Doc, root: NodeId) {
    let source_accents = named(source, source.descendants_and_self(equation), M, "acc");
    if source_accents.is_empty() {
        return;
    }
    let target_accents: Vec<NodeId> = named(doc, doc.descendants_and_self(root), ML, "mover")
        .into_iter()
        .filter(|&e| doc.attr(e, "accent").is_some_and(|a| eq_ignore_case(a, "true")))
        .collect();
    let count = source_accents.len().min(target_accents.len());
    for index in 0..count {
        let source_accent = source_accents[index];
        let office_character = child_path(source, Some(source_accent), &[(M, "accPr"), (M, "chr")])
            .and_then(|e| source.attr_ns(e, M, "val"));
        let canonical_character = canonical_mathml_accent_character(office_character);
        let target_accent = target_accents[index];
        let children = doc.elements(target_accent);
        if children.len() < 2 {
            continue;
        }
        let restored_mark = new_element(doc, ML, "mo");
        doc.set_attr(restored_mark, "accent", "true");
        doc.append_text(restored_mark, &canonical_character);
        let control_run = child_path(source, Some(source_accent), &[(M, "accPr"), (M, "ctrlPr"), (W, "rPr")]);
        let bold = child_path(source, control_run, &[(W, "b")]);
        let italic = child_path(source, control_run, &[(W, "i")]);
        if bold.is_some() || italic.is_some() {
            let enabled = |property: Option<NodeId>| {
                property.is_some_and(|p| !matches!(source.attr_ns(p, W, "val"), Some("0" | "false" | "off")))
            };
            let variant = if enabled(bold) {
                if enabled(italic) { "bold-italic" } else { "bold" }
            } else if enabled(italic) {
                "italic"
            } else {
                "normal"
            };
            doc.set_attr(restored_mark, "mathvariant", variant);
        }
        let color_value = child_path(source, control_run, &[(W, "color")]).and_then(|e| source.attr_ns(e, W, "val"));
        let color = normalize_color(Some(&format!("#{}", color_value.unwrap_or(""))));
        if let Some(color) = color {
            doc.set_attr(restored_mark, "mathcolor", &color);
        }
        doc.replace(children[1], restored_mark);
    }
}

fn is_standard_tex_operator_name(value: &str) -> bool {
    matches!(
        value,
        "arccos" | "arcsin" | "arctan" | "arg" | "cos" | "cosh" | "cot" | "coth" | "csc" | "deg" | "det" | "dim"
            | "exp" | "gcd" | "hom" | "inf" | "ker" | "lg" | "lim" | "liminf" | "limsup" | "ln" | "log" | "max"
            | "min" | "Pr" | "sec" | "sin" | "sinh" | "sup" | "tan" | "tanh"
    )
}

fn flatten_mathml_token_text(doc: &Doc, element: NodeId) -> String {
    let name = doc.name(element).unwrap();
    if name.ns != ML {
        return String::new();
    }
    if matches!(name.local.as_str(), "mi" | "mn" | "mo" | "mtext") {
        return doc.value(element);
    }
    if !matches!(name.local.as_str(), "mrow" | "mstyle" | "mpadded" | "semantics") {
        return String::new();
    }
    let mut builder = String::new();
    for child in doc.elements(element) {
        let value = flatten_mathml_token_text(doc, child);
        if value.is_empty() && has_elements(doc, child) {
            return String::new();
        }
        builder.push_str(&value);
    }
    builder
}

fn direct_text(doc: &Doc, run: NodeId) -> Vec<NodeId> {
    doc.elements_named(run, M, "t")
}

fn restore_omml_limit_base_semantics(source: &Doc, equation: NodeId, doc: &mut Doc, root: NodeId) {
    let limits: Vec<NodeId> = source
        .descendants_and_self(equation)
        .into_iter()
        .filter(|&e| source.is(e, M, "limLow") || source.is(e, M, "limUpp"))
        .collect();
    if limits.is_empty() {
        return;
    }
    let mut used_targets: HashSet<NodeId> = HashSet::new();
    for limit in limits {
        let Some(base_slot) = source.element(limit, M, "e") else { continue };
        let visible_runs: Vec<NodeId> = named(source, source.descendants(base_slot), M, "r")
            .into_iter()
            .filter(|&run| direct_text(source, run).into_iter().any(|t| !is_null_or_white_space(&source.value(t))))
            .collect();
        if visible_runs.is_empty() {
            continue;
        }
        let base_text: String = visible_runs
            .iter()
            .flat_map(|&run| direct_text(source, run))
            .map(|t| source.value(t))
            .collect();
        if is_null_or_white_space(&base_text) {
            continue;
        }
        let explicitly_normal = visible_runs
            .iter()
            .all(|&run| child_path(source, Some(run), &[(M, "rPr"), (M, "nor")]).is_some());
        if !explicitly_normal && !is_standard_tex_operator_name(&base_text) {
            continue;
        }
        let lower = source.is(limit, M, "limLow");
        let target_names: [&str; 2] = if lower { ["msub", "munder"] } else { ["msup", "mover"] };
        let target = doc.descendants_and_self(root).into_iter().find(|&element| {
            target_names.iter().any(|&n| doc.is(element, ML, n)) && !used_targets.contains(&element) && {
                let children = doc.elements(element);
                children.len() >= 2 && flatten_mathml_token_text(doc, children[0]) == base_text
            }
        });
        let Some(target) = target else { continue };
        let target_children = doc.elements(target);
        let replacement = new_element(doc, ML, "mi");
        doc.set_attr(replacement, "mathvariant", "normal");
        doc.append_text(replacement, &base_text);
        doc.replace(target_children[0], replacement);
        used_targets.insert(target);
    }
}

fn restore_omml_script_base_run_semantics(source: &Doc, equation: NodeId, doc: &mut Doc, root: NodeId) {
    let scripts: Vec<NodeId> = source
        .descendants_and_self(equation)
        .into_iter()
        .filter(|&e| source.is(e, M, "sSub") || source.is(e, M, "sSup") || source.is(e, M, "sSubSup"))
        .collect();
    if scripts.is_empty() {
        return;
    }
    let mut used_targets: HashSet<NodeId> = HashSet::new();
    for script in scripts {
        let Some(base_slot) = source.element(script, M, "e") else { continue };
        let base_text: String =
            named(source, source.descendants(base_slot), M, "t").into_iter().map(|t| source.value(t)).collect();
        if is_null_or_white_space(&base_text) {
            continue;
        }
        let target_name = if source.is(script, M, "sSub") {
            "msub"
        } else if source.is(script, M, "sSup") {
            "msup"
        } else {
            "msubsup"
        };
        let target = named(doc, doc.descendants_and_self(root), ML, target_name).into_iter().find(|&element| {
            !used_targets.contains(&element) && {
                let children = doc.elements(element);
                children.len() >= 2 && flatten_mathml_token_text(doc, children[0]) == base_text
            }
        });
        let Some(target) = target else { continue };
        // Claim the corresponding base even when no repair is required.
        used_targets.insert(target);

        let base_contents: Vec<NodeId> =
            source.elements(base_slot).into_iter().filter(|&e| !source.is(e, M, "ePr")).collect();
        if base_contents.len() != 1 || !source.is(base_contents[0], M, "r") {
            continue;
        }
        let properties = source.element(base_contents[0], M, "rPr");
        let projected_base = doc.elements(target)[0];
        if is_math_boolean_enabled(source, properties, M, "nor") {
            for token in doc.descendants_and_self(projected_base) {
                if doc.is(token, ML, "mi") || doc.is(token, ML, "mo") || doc.is(token, ML, "mn") {
                    doc.set_name(token, Name::new(ML, "mtext"));
                }
            }
            continue;
        }
        let style = child_path(source, properties, &[(M, "sty")]).and_then(|e| source.attr_ns(e, M, "val"));
        if !style.is_some_and(|s| eq_ignore_case(s, "p")) {
            continue;
        }
        if doc
            .descendants_and_self(projected_base)
            .into_iter()
            .any(|e| !doc.is(e, ML, "mrow") && !doc.is(e, ML, "mi"))
        {
            continue;
        }
        let identifiers = named(doc, doc.descendants_and_self(projected_base), ML, "mi");
        if identifiers.is_empty()
            || identifiers.iter().any(|&identifier| resolve_effective_variant(doc, identifier).as_deref() != Some("normal"))
        {
            continue;
        }
        let replacement = new_element(doc, ML, "mi");
        doc.set_attr(replacement, "mathvariant", "normal");
        doc.append_text(replacement, &base_text);
        doc.replace(projected_base, replacement);
    }
}

fn restore_omml_function_application_runs(source: &Doc, equation: NodeId, doc: &mut Doc, root: NodeId) {
    let source_runs = named(source, source.descendants_and_self(equation), M, "r");
    let mut index = 0;
    while index + 1 < source_runs.len() {
        let run = source_runs[index];
        index += 1;
        let properties = source.element(run, M, "rPr");
        let style = child_path(source, properties, &[(M, "sty")]).and_then(|e| source.attr_ns(e, M, "val"));
        let explicit_normal = child_path(source, properties, &[(M, "nor")]).is_some()
            || style.is_some_and(|s| eq_ignore_case(s, "p"));
        if !explicit_normal {
            continue;
        }
        let function_name: String = direct_text(source, run).into_iter().map(|t| source.value(t)).collect();
        if is_null_or_white_space(&function_name) {
            continue;
        }
        let next_run = source_runs[index];
        let next_text: String = direct_text(source, next_run).into_iter().map(|t| source.value(t)).collect();
        if next_text.chars().next() != Some('\u{2061}') {
            continue;
        }
        collapse_mathml_upright_sequence(doc, root, &function_name);
    }
}

fn collapse_mathml_upright_sequence(doc: &mut Doc, root: NodeId, expected_text: &str) -> bool {
    for parent in doc.descendants_and_self(root) {
        let children = doc.elements(parent);
        for start in 0..children.len() {
            let mut matched: Vec<NodeId> = Vec::new();
            let mut builder = String::new();
            for &child in &children[start..] {
                if !doc.is(child, ML, "mi") {
                    break;
                }
                let variant = doc.attr(child, "mathvariant").unwrap_or("");
                if !contains_ignore_case(variant, "normal") && !contains_ignore_case(variant, "upright") {
                    break;
                }
                matched.push(child);
                builder.push_str(&doc.value(child));
                if !expected_text.starts_with(builder.as_str()) {
                    break;
                }
                if builder != expected_text {
                    continue;
                }
                let replacement = new_element(doc, ML, "mi");
                doc.set_attr(replacement, "mathvariant", "normal");
                doc.append_text(replacement, expected_text);
                doc.insert_before(matched[0], replacement);
                for element in matched {
                    doc.detach(element);
                }
                return true;
            }
        }
    }
    false
}

fn restore_omml_explicit_space_tokens(doc: &mut Doc, root: NodeId) {
    let maximum = maximum_omml_space_text_length();
    for parent in doc.descendants_and_self(root) {
        let children = doc.elements(parent);
        let mut index = 0;
        while index < children.len() {
            let first = children[index];
            if !doc.is(first, ML, "mo") || has_elements(doc, first) || find_omml_space_text(&doc.value(first)).is_none()
            {
                index += 1;
                continue;
            }
            let mut text = doc.value(first);
            let mut identity = find_omml_space_text(&text).unwrap();
            let mut last = index;
            for next in index + 1..children.len() {
                let candidate = children[next];
                let candidate_value = doc.value(candidate);
                if !doc.is(candidate, ML, "mo")
                    || has_elements(doc, candidate)
                    || doc.attrs(first) != doc.attrs(candidate)
                    || find_omml_space_text(&candidate_value).is_none()
                    || text.encode_utf16().count() + candidate_value.encode_utf16().count() > maximum
                {
                    break;
                }
                let combined = format!("{text}{candidate_value}");
                let Some(combined_identity) = find_omml_space_text(&combined) else { break };
                text = combined;
                identity = combined_identity;
                last = next;
            }
            let space = new_element(doc, ML, "mspace");
            let attributes: Vec<(Name, String)> =
                doc.attrs(first).iter().filter(|(name, _)| !(name.ns.is_empty() && name.local == "width")).cloned().collect();
            for (name, value) in attributes {
                doc.set_attr_ns(space, &name.ns, &name.local, Some(&value));
            }
            doc.set_attr(space, "width", identity.mathml_width);
            doc.replace(first, space);
            for &removed in &children[index + 1..=last] {
                doc.detach(removed);
            }
            index = last + 1;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = r#"xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

    #[test]
    fn multiple_equations_are_rejected() {
        let omml = format!("<m:oMathPara {NS}><m:oMath/><m:oMath/></m:oMathPara>");
        let error = omml_to_mathml(&omml, true).unwrap_err();
        assert!(error.contains("Expected exactly one OMML equation"), "{error}");
    }

    #[test]
    fn display_attribute_and_default_namespace() {
        let omml = format!("<m:oMath {NS}><m:r><m:t>x</m:t></m:r></m:oMath>");
        assert_eq!(
            omml_to_mathml(&omml, false).unwrap(),
            r#"<math display="inline" xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>"#
        );
    }

    #[test]
    fn merged_numeric_punctuation_is_split() {
        let omml = format!("<m:oMath {NS}><m:r><m:t>1,2</m:t></m:r></m:oMath>");
        assert_eq!(
            omml_to_mathml(&omml, true).unwrap(),
            r#"<math display="block" xmlns="http://www.w3.org/1998/Math/MathML"><mrow><mn>1</mn><mo>,</mo><mn>2</mn></mrow></math>"#
        );
    }

    #[test]
    fn explicit_em_spaces_become_mspace() {
        let omml = format!("<m:oMath {NS}><m:r><m:t>a\u{2003}\u{2003}b</m:t></m:r></m:oMath>");
        let mathml = omml_to_mathml(&omml, true).unwrap();
        assert!(mathml.contains(r#"<mspace stretchy="false" width="2em" />"#), "{mathml}");
    }

    #[test]
    fn visual_tex_equation_number_is_stripped() {
        let omml = format!(
            "<m:oMath {NS}><m:eqArr><m:e><m:r><m:t>x#</m:t></m:r><m:d><m:e><m:r><m:t>981730</m:t></m:r></m:e></m:d></m:e></m:eqArr></m:oMath>"
        );
        let (doc, equation) = strip_visual_tex_native_equation_number(&omml).unwrap();
        assert_eq!(doc.canonical(equation), format!("<{{{M}}}oMath><{{{M}}}r><{{{M}}}t>[x]</t></r></oMath>"));
    }
}
