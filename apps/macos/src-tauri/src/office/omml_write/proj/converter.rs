//! Port of the WordOmmlConverter.cs passes that run after the MathML
//! preparation (s01..s12) in TransformMathMlToOmml, plus
//! ReplaceExtendedIntegralsWithOfficePlaceholders (whose n-ary list feeds
//! RestoreExtendedIntegralCharacters), IsBlockMathMl and
//! BuildWordNativeNumberedOmml.
//!
//! Each pass parses its input like the C# does (XDocument.Parse(...,
//! LoadOptions.PreserveWhitespace) unless noted) and serializes with
//! `Root.ToString(SaveOptions.DisableFormatting)`.

use crate::office::omml_write::proj::dom::{Doc, Name, NodeId, MATHML_NS, OMML_NS, WORD_NS, XML_NS};
use crate::office::omml_write::proj::netstr::{
    contains_ordinal_ignore_case, contains_unicode_letter, equals_ordinal_ignore_case, first_letter_run,
    is_null_or_white_space, is_white_space, len16, parse_combining_characters, substring, substring_from,
    to_lower_invariant, trim,
};
use crate::office::omml_write::proj::xml_reader::{self, ReadOptions, CONVERTER_READER, IGNORE_WHITESPACE_READER, PRESERVE_WHITESPACE};
use crate::office::omml_write::proj::xml_writer::{document_to_string_disable_formatting, to_string_disable_formatting};

const M: &str = OMML_NS;
const W: &str = WORD_NS;
const MML: &str = MATHML_NS;

/// WordOmmlConverter.NaryCharacters.
pub const NARY_CHARACTERS: &str = "∑∏∐∫∬∭∮∯∰∱∲∳⨋⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜⋀⋁⋂⋃";
/// WordOmmlConverter.ExtendedIntegralCharacters.
pub const EXTENDED_INTEGRAL_CHARACTERS: &str = "∯∰∱∲∳⨋⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜";
const NATIVE_NUMBER_BOOKMARK_PREFIX: &str = "VTEqNum_";

fn invalid_data(message: &str) -> String {
    format!("InvalidDataException: {message}")
}

fn argument_error(message: &str, parameter: &str) -> String {
    format!("ArgumentException: {message} (Parameter '{parameter}')")
}

/// A one-UTF-16-unit string whose unit is in `set` (`s.Length == 1 &&
/// set.IndexOf(s[0]) >= 0`; every character of the sets is in the BMP).
fn single_char_in(value: &str, set: &str) -> bool {
    let mut chars = value.chars();
    match (chars.next(), chars.next()) {
        (Some(c), None) => (c as u32) < 0x10000 && set.contains(c),
        _ => false,
    }
}

fn parse(xml: &str, options: ReadOptions) -> Result<(Doc, NodeId), String> {
    let doc = xml_reader::parse(xml, options)?;
    let root = doc.root.ok_or_else(|| "XmlException: Root element is missing.".to_string())?;
    Ok((doc, root))
}

/// XDocument.Parse(xml, LoadOptions.PreserveWhitespace).
fn parse_preserving(xml: &str) -> Result<(Doc, NodeId), String> {
    parse(xml, PRESERVE_WHITESPACE)
}

fn root_to_string(doc: &Doc, root: NodeId) -> Result<String, String> {
    to_string_disable_formatting(doc, root)
}

fn new_element(doc: &mut Doc, ns: &str, local: &str) -> NodeId {
    doc.new_element(Name::new(ns, local))
}

fn element_with_val(doc: &mut Doc, ns: &str, local: &str, value: &str) -> NodeId {
    let element = new_element(doc, ns, local);
    doc.set_attr_ns(element, M, "val", Some(value));
    element
}

/// Descendants(name) of the document (root included).
fn document_descendants(doc: &Doc, root: NodeId, ns: &str, local: &str) -> Vec<NodeId> {
    doc.descendants_and_self(root).into_iter().filter(|&id| doc.is(id, ns, local)).collect()
}

// ---- ReplaceExtendedIntegralsWithOfficePlaceholders / IsBlockMathMl ---------------

/// ReplaceExtendedIntegralsWithOfficePlaceholders(mathMl): the MathML with
/// extended integrals replaced by U+222B, and the value of every MathML mo
/// whose value is one n-ary character, in document order.
pub fn replace_extended_integrals_with_office_placeholders(mathml: &str) -> Result<(String, Vec<String>), String> {
    if is_null_or_white_space(mathml) {
        return Ok((mathml.to_string(), Vec::new()));
    }
    let document = xml_reader::parse_document(mathml, CONVERTER_READER)?;
    let (before_root, after_root) = (document.before_root, document.after_root);
    let mut doc = document.doc;
    let root = doc.root.ok_or_else(|| "XmlException: Root element is missing.".to_string())?;
    let operators: Vec<NodeId> = document_descendants(&doc, root, MML, "mo")
        .into_iter()
        .filter(|&element| single_char_in(&doc.value(element), NARY_CHARACTERS))
        .collect();
    let characters: Vec<String> = operators.iter().map(|&element| doc.value(element)).collect();
    for &op in &operators {
        if single_char_in(&doc.value(op), EXTENDED_INTEGRAL_CHARACTERS) {
            // Office's MML2OMML transform knows how to attach limits and the
            // following operand to a standard integral; the exact extended
            // character is restored in OMML right after the transform.
            doc.set_value(op, "∫");
        }
    }
    // document.ToString(SaveOptions.DisableFormatting)
    Ok((document_to_string_disable_formatting(&doc, root, &before_root, &after_root)?, characters))
}

/// IsBlockMathMl(mathMl): display="block" (OrdinalIgnoreCase) on the root;
/// false when the MathML cannot be read.
pub fn is_block_mathml(mathml: &str) -> bool {
    match parse(mathml, IGNORE_WHITESPACE_READER) {
        Ok((doc, root)) => equals_ordinal_ignore_case(doc.attr(root, "display"), "block"),
        Err(_) => false,
    }
}

// ---- o00 ExtractSingleOMath ------------------------------------------------------------

/// ExtractSingleOMath(omml): the one m:oMath of the input, serialized alone.
pub fn extract_single_omath(omml: &str) -> Result<String, String> {
    if is_null_or_white_space(omml) {
        return Err(invalid_data("Office produced an empty OMML transformation."));
    }
    let (doc, root) = parse(omml, xml_reader::formula_reader("OMML"))?;
    // This API promises one complete equation, not the first equation in a
    // larger range. Counting also detects an illegally nested oMath root.
    let equations: Vec<NodeId> = document_descendants(&doc, root, M, "oMath").into_iter().take(2).collect();
    if equations.is_empty() {
        return Err(invalid_data("Office MathML conversion did not produce an m:oMath node."));
    }
    if equations.len() != 1 {
        return Err(invalid_data("Expected exactly one OMML equation; the input contains multiple equations."));
    }
    to_string_disable_formatting(&doc, equations[0])
}

// ---- o01 RestoreExtendedIntegralCharacters -------------------------------------------------

pub fn restore_extended_integral_characters(omml: &str, source_nary_characters: &[String]) -> Result<String, String> {
    if source_nary_characters.is_empty()
        || !source_nary_characters.iter().any(|character| single_char_in(character, EXTENDED_INTEGRAL_CHARACTERS))
    {
        return Ok(omml.to_string());
    }
    let (mut doc, root) = parse_preserving(omml)?;
    let naries = document_descendants(&doc, root, M, "nary");
    if naries.len() != source_nary_characters.len() {
        return Err(invalid_data(&format!(
            "Office changed the number of n-ary operators while converting extended integrals. MathML={}; OMML={}.",
            source_nary_characters.len(),
            naries.len()
        )));
    }
    for (index, source_character) in source_nary_characters.iter().enumerate() {
        if !single_char_in(source_character, EXTENDED_INTEGRAL_CHARACTERS) {
            continue;
        }
        let properties = match doc.element(naries[index], M, "naryPr") {
            Some(properties) => properties,
            None => {
                let properties = new_element(&mut doc, M, "naryPr");
                doc.add_first(naries[index], properties);
                properties
            }
        };
        let character = match doc.element(properties, M, "chr") {
            Some(character) => character,
            None => {
                let character = new_element(&mut doc, M, "chr");
                doc.add_first(properties, character);
                character
            }
        };
        doc.set_attr_ns(character, M, "val", Some(source_character));
    }
    root_to_string(&doc, root)
}

// ---- o02 NormalizeExplicitUprightRuns --------------------------------------------------------

pub fn normalize_explicit_upright_runs(omml: &str, mathml: &str) -> Result<String, String> {
    if is_null_or_white_space(omml) || is_null_or_white_space(mathml) {
        return Ok(omml.to_string());
    }
    let (mut doc, root) = parse_preserving(omml)?;
    for run in document_descendants(&doc, root, M, "r") {
        let properties = doc.element(run, M, "rPr");
        // The source token projection now provides the style to Office.
        // Only an actual plain mathematical run permits upright cleanup.
        let style = properties.and_then(|p| doc.element(p, M, "sty")).and_then(|sty| doc.attr_ns(sty, M, "val"));
        if style != Some("p") {
            continue;
        }
        let properties = properties.unwrap();
        if let Some(normal) = doc.element(properties, M, "nor") {
            doc.detach(normal);
        }
        let text: String = doc.elements_named(run, M, "t").into_iter().map(|t| doc.value(t)).collect();
        if !contains_unicode_letter(&text) {
            continue;
        }
        let word_properties = match doc.element(run, W, "rPr") {
            Some(word_properties) => word_properties,
            None => {
                let word_properties = new_element(&mut doc, W, "rPr");
                doc.insert_after(properties, word_properties);
                word_properties
            }
        };
        if doc.element(word_properties, W, "noProof").is_none() {
            let no_proof = new_element(&mut doc, W, "noProof");
            doc.append(word_properties, no_proof);
        }
    }
    root_to_string(&doc, root)
}

// ---- o03 NormalizeAppliedFunctionStructures ---------------------------------------------------

struct Application {
    function_token: String,
    open: Option<&'static str>,
    close: Option<&'static str>,
    argument_text_elements: usize,
}

fn run_text(doc: &Doc, run: NodeId) -> String {
    doc.elements_named(run, M, "t").into_iter().map(|t| doc.value(t)).collect()
}

/// CloneRunWithText(sourceRun, math, text).
fn clone_run_with_text(doc: &mut Doc, source_run: NodeId, text: &str) -> NodeId {
    let clone = doc.deep_clone(source_run);
    for t in doc.elements_named(clone, M, "t") {
        doc.detach(t);
    }
    let text_element = new_element(doc, M, "t");
    doc.add_string(text_element, text);
    let first = text.chars().next();
    let last = text.chars().last();
    if !text.is_empty() && (first.is_some_and(is_white_space) || last.is_some_and(is_white_space)) {
        doc.set_attr_ns(text_element, XML_NS, "space", Some("preserve"));
    }
    doc.append(clone, text_element);
    clone
}

/// TextElementEnd(value, count): UTF-16 end of the first `count` text elements.
fn text_element_end(value: &str, count: usize) -> usize {
    if value.is_empty() || count == 0 {
        return 0;
    }
    let starts = parse_combining_characters(value);
    if starts.len() <= count {
        return len16(value);
    }
    starts[count]
}

/// FindMatchingDelimiterEnd: UTF-16 index after the matching close, or None.
fn find_matching_delimiter_end(value: &str, open: &str, close: &str, depth: &mut i32, started: &mut bool) -> Option<usize> {
    // Compares one UTF-16 unit at a time; open/close are ASCII, so a
    // surrogate unit never matches and scalar iteration is equivalent.
    let mut index16 = 0usize;
    for c in value.chars() {
        let units = c.len_utf16();
        let is_open = units == 1 && open.len() == 1 && open.starts_with(c);
        let is_close = units == 1 && close.len() == 1 && close.starts_with(c);
        if is_open {
            *depth += 1;
            *started = true;
        } else if *started && is_close {
            *depth -= 1;
            if *depth == 0 {
                return Some(index16 + 1);
            }
        }
        index16 += units;
    }
    None
}

pub fn normalize_applied_function_structures(omml: &str, mathml: &str) -> Result<String, String> {
    if is_null_or_white_space(omml)
        || is_null_or_white_space(mathml)
        || (!mathml.contains('\u{2061}') && !contains_ordinal_ignore_case(mathml, "2061"))
    {
        return Ok(omml.to_string());
    }

    let (source, source_root) = parse_preserving(mathml)?;
    let mut applications = Vec::new();
    for marker in document_descendants(&source, source_root, MML, "mo") {
        if source.value(marker) != "\u{2061}" {
            continue;
        }
        let Some(parent) = source.parent(marker) else { continue };
        let siblings = source.elements(parent);
        let marker_index = siblings.iter().position(|&sibling| sibling == marker).unwrap();
        if marker_index == 0 || marker_index + 1 >= siblings.len() {
            continue;
        }

        let function_source = siblings[marker_index - 1];
        let function_token = source
            .descendants_and_self(function_source)
            .into_iter()
            .filter(|&element| source.is(element, MML, "mi") || source.is(element, MML, "mo"))
            .find_map(|element| first_letter_run(&source.value(element)))
            .unwrap_or_default();
        if is_null_or_white_space(&function_token) {
            continue;
        }

        let argument_source = siblings[marker_index + 1];
        let (mut open, mut close) = (None, None);
        if source.is(argument_source, MML, "mo") {
            (open, close) = match source.value(argument_source).as_str() {
                "(" => (Some("("), Some(")")),
                "[" => (Some("["), Some("]")),
                "{" => (Some("{"), Some("}")),
                _ => (None, None),
            };
        }

        let simple_argument = source.is(argument_source, MML, "mi")
            || source.is(argument_source, MML, "mn")
            || source.is(argument_source, MML, "mo")
            || source.is(argument_source, MML, "mtext");
        let argument_text_elements = if simple_argument && open.is_none() {
            parse_combining_characters(&source.value(argument_source)).len()
        } else {
            0
        };
        applications.push(Application { function_token, open, close, argument_text_elements });
    }
    if applications.is_empty() {
        return Ok(omml.to_string());
    }

    let (mut target, root) = parse_preserving(omml)?;
    for application in &applications {
        let marker_text = document_descendants(&target, root, M, "t")
            .into_iter()
            .find(|&text| target.value(text).contains('\u{2061}'));
        let marker_run = marker_text.and_then(|text| target.parent(text));
        let Some(marker_run) = marker_run.filter(|&run| target.is(run, M, "r")) else { continue };
        if target.parent(marker_run).is_none() {
            continue;
        }
        let Some(function_element) = target.previous_element(marker_run) else { continue };
        let function_visible_text: String =
            target.descendants_named(function_element, M, "t").into_iter().map(|t| target.value(t)).collect();
        if !contains_ordinal_ignore_case(&function_visible_text, &application.function_token) {
            continue;
        }

        let marker_run_text = run_text(&target, marker_run);
        let Some(marker_index) = marker_run_text.find('\u{2061}') else { continue };
        // NormalizeExplicitUprightRuns deliberately split the function name
        // into its own m:r, so ApplyFunction must begin the following run.
        // Anything else is an unfamiliar Office-XSL shape and is left intact.
        if marker_index != 0 {
            continue;
        }
        let after_marker = substring_from(&marker_run_text, 1)?;
        let mut argument_elements: Vec<NodeId> = Vec::new();
        let mut consumed = vec![marker_run];
        let mut tail: Option<NodeId> = None;
        let mut completed = false;

        if let (Some(open), Some(close)) = (application.open, application.close) {
            let mut depth = 0;
            let mut started = false;
            let end = find_matching_delimiter_end(&after_marker, open, close, &mut depth, &mut started);
            if let Some(end) = end {
                let head = substring(&after_marker, 0, end)?;
                argument_elements.push(clone_run_with_text(&mut target, marker_run, &head));
                if end < len16(&after_marker) {
                    let rest = substring_from(&after_marker, end)?;
                    tail = Some(clone_run_with_text(&mut target, marker_run, &rest));
                }
                completed = true;
            } else if started {
                if !after_marker.is_empty() {
                    argument_elements.push(clone_run_with_text(&mut target, marker_run, &after_marker));
                }
                for sibling in target.elements_after_self(marker_run) {
                    if target.is(sibling, M, "r") {
                        let sibling_text = run_text(&target, sibling);
                        let end = find_matching_delimiter_end(&sibling_text, open, close, &mut depth, &mut started);
                        consumed.push(sibling);
                        if let Some(end) = end {
                            if end > 0 {
                                let head = substring(&sibling_text, 0, end)?;
                                argument_elements.push(clone_run_with_text(&mut target, sibling, &head));
                            }
                            if end < len16(&sibling_text) {
                                let rest = substring_from(&sibling_text, end)?;
                                tail = Some(clone_run_with_text(&mut target, sibling, &rest));
                            }
                            completed = true;
                            break;
                        }
                        argument_elements.push(target.deep_clone(sibling));
                        continue;
                    }

                    let visible: String =
                        target.descendants_named(sibling, M, "t").into_iter().map(|t| target.value(t)).collect();
                    let _ = find_matching_delimiter_end(&visible, open, close, &mut depth, &mut started);
                    consumed.push(sibling);
                    argument_elements.push(target.deep_clone(sibling));
                    if started && depth == 0 {
                        completed = true;
                        break;
                    }
                }
            }
        } else {
            let source_text_elements = application.argument_text_elements;
            if !after_marker.is_empty() {
                let take = if source_text_elements > 0 {
                    text_element_end(&after_marker, source_text_elements)
                } else {
                    text_element_end(&after_marker, 1)
                };
                if take > 0 {
                    let head = substring(&after_marker, 0, take)?;
                    argument_elements.push(clone_run_with_text(&mut target, marker_run, &head));
                    if take < len16(&after_marker) {
                        let rest = substring_from(&after_marker, take)?;
                        tail = Some(clone_run_with_text(&mut target, marker_run, &rest));
                    }
                    completed = true;
                }
            } else if let Some(sibling) = target.next_element(marker_run) {
                consumed.push(sibling);
                if target.is(sibling, M, "r") && source_text_elements > 0 {
                    let sibling_text = run_text(&target, sibling);
                    let take = text_element_end(&sibling_text, source_text_elements);
                    if take > 0 {
                        let head = substring(&sibling_text, 0, take)?;
                        argument_elements.push(clone_run_with_text(&mut target, sibling, &head));
                        if take < len16(&sibling_text) {
                            let rest = substring_from(&sibling_text, take)?;
                            tail = Some(clone_run_with_text(&mut target, sibling, &rest));
                        }
                        completed = true;
                    }
                } else {
                    argument_elements.push(target.deep_clone(sibling));
                    completed = true;
                }
            }
        }

        if !completed || argument_elements.is_empty() {
            continue;
        }

        let function_copy = target.deep_clone(function_element);
        let function_name = new_element(&mut target, M, "fName");
        target.append(function_name, function_copy);
        let argument = new_element(&mut target, M, "e");
        for element in argument_elements {
            target.append(argument, element);
        }
        let function = new_element(&mut target, M, "func");
        target.append(function, function_name);
        target.append(function, argument);
        target.replace(function_element, function);
        for element in consumed {
            if target.parent(element).is_none() {
                return Err("InvalidOperationException: The parent is missing.".to_string());
            }
            target.detach(element);
        }
        if let Some(tail) = tail {
            target.insert_after(function, tail);
        }
    }

    root_to_string(&target, root)
}

// ---- o04 NormalizeExplicitTableColumnAlignment ------------------------------------------------

struct SourceTable {
    row_count: usize,
    column_count: usize,
    alignments: Vec<&'static str>,
}

/// NormalizeMathMlColumnAlignment(value).
fn normalize_mathml_column_alignment(value: &str) -> &'static str {
    match to_lower_invariant(trim(value)).as_str() {
        "left" => "left",
        "right" => "right",
        _ => "center",
    }
}

pub fn normalize_explicit_table_column_alignment(omml: &str, mathml: &str) -> Result<String, String> {
    if is_null_or_white_space(omml) || is_null_or_white_space(mathml) {
        return Ok(omml.to_string());
    }
    let (source, source_root) = parse_preserving(mathml)?;
    let (mut target, root) = parse_preserving(omml)?;
    let mut source_tables = Vec::new();
    for table in document_descendants(&source, source_root, MML, "mtable") {
        let Some(columnalign) = source.attr(table, "columnalign") else { continue };
        if is_null_or_white_space(columnalign) {
            continue;
        }
        let rows = source.elements_named(table, MML, "mtr");
        let column_count =
            rows.iter().map(|&row| source.elements_named(row, MML, "mtd").len()).max().unwrap_or(0);
        let raw: Vec<&str> = columnalign.split(is_white_space).filter(|part| !part.is_empty()).collect();
        if column_count == 0 || raw.is_empty() {
            continue;
        }
        let alignments = (0..column_count)
            .map(|index| normalize_mathml_column_alignment(raw[index.min(raw.len() - 1)]))
            .collect();
        source_tables.push(SourceTable { row_count: rows.len(), column_count, alignments });
    }
    if source_tables.is_empty() {
        return Ok(omml.to_string());
    }

    let matrices = document_descendants(&target, root, M, "m");
    let mut used: Vec<NodeId> = Vec::new();
    for source_table in &source_tables {
        let found = matrices.iter().copied().find(|&matrix| {
            if used.contains(&matrix) {
                return false;
            }
            let rows = target.elements_named(matrix, M, "mr");
            if rows.len() != source_table.row_count {
                return false;
            }
            rows.iter().all(|&row| target.elements_named(row, M, "e").len() == source_table.column_count)
        });
        let Some(found) = found else { continue };
        used.push(found);

        let properties = match target.element(found, M, "mPr") {
            Some(properties) => properties,
            None => {
                let properties = new_element(&mut target, M, "mPr");
                target.add_first(found, properties);
                properties
            }
        };
        let columns = new_element(&mut target, M, "mcs");
        for &alignment in &source_table.alignments {
            let count = element_with_val(&mut target, M, "count", "1");
            let justification = element_with_val(&mut target, M, "mcJc", alignment);
            let column_properties = new_element(&mut target, M, "mcPr");
            target.append(column_properties, count);
            target.append(column_properties, justification);
            let column = new_element(&mut target, M, "mc");
            target.append(column, column_properties);
            target.append(columns, column);
        }
        match target.element(properties, M, "mcs") {
            None => target.append(properties, columns),
            Some(existing) => target.replace(existing, columns),
        }
    }

    root_to_string(&target, root)
}

// ---- o05 NormalizeOmmlPlaceholderVisibility ---------------------------------------------------

pub fn normalize_omml_placeholder_visibility(omml: &str) -> Result<String, String> {
    if is_null_or_white_space(omml) {
        return Ok(omml.to_string());
    }
    let (mut doc, root) = parse_preserving(omml)?;
    for matrix in document_descendants(&doc, root, M, "m") {
        let properties = match doc.element(matrix, M, "mPr") {
            Some(properties) => properties,
            None => {
                let properties = new_element(&mut doc, M, "mPr");
                doc.add_first(matrix, properties);
                properties
            }
        };
        let placeholder_visibility = match doc.element(properties, M, "plcHide") {
            Some(element) => element,
            None => {
                let element = new_element(&mut doc, M, "plcHide");
                doc.add_first(properties, element);
                element
            }
        };
        doc.set_attr_ns(placeholder_visibility, M, "val", Some("1"));
    }
    root_to_string(&doc, root)
}

// ---- o06 NormalizeOptionalPrescriptSlots -------------------------------------------------------

/// HasVisibleMathContent(slot).
fn has_visible_math_content(doc: &Doc, slot: Option<NodeId>) -> bool {
    let Some(slot) = slot else { return false };
    doc.descendants_and_self(slot)
        .into_iter()
        .any(|element| (doc.is(element, M, "t") || doc.is(element, W, "t")) && !doc.value(element).is_empty())
}

pub fn normalize_optional_prescript_slots(omml: &str) -> Result<String, String> {
    let (mut doc, root) = parse_preserving(omml)?;
    for script in document_descendants(&doc, root, M, "sPre") {
        if !has_visible_math_content(&doc, doc.element(script, M, "e")) {
            continue;
        }
        for name in ["sub", "sup"] {
            let slot = doc.element(script, M, name);
            let other = doc.element(script, M, if name == "sub" { "sup" } else { "sub" });
            let Some(slot) = slot else { continue };
            if !has_visible_math_content(&doc, other) || has_visible_math_content(&doc, Some(slot)) {
                continue;
            }
            // Preserve unfamiliar controls/extensions rather than replacing them.
            if doc
                .descendants(slot)
                .into_iter()
                .any(|element| !doc.is(element, M, "r") && !doc.is(element, M, "t") && !doc.is(element, M, "rPr"))
            {
                continue;
            }
            // CT_SPre has a pair of slots but no subHide/supHide: represent the
            // explicitly absent half as an empty native phantom with zero extent.
            let show = element_with_val(&mut doc, M, "show", "0");
            let zero_width = element_with_val(&mut doc, M, "zeroWid", "1");
            let zero_ascent = element_with_val(&mut doc, M, "zeroAsc", "1");
            let zero_descent = element_with_val(&mut doc, M, "zeroDesc", "1");
            let properties = new_element(&mut doc, M, "phantPr");
            for child in [show, zero_width, zero_ascent, zero_descent] {
                doc.append(properties, child);
            }
            let body = new_element(&mut doc, M, "e");
            let phantom = new_element(&mut doc, M, "phant");
            doc.append(phantom, properties);
            doc.append(phantom, body);
            // slot.ReplaceNodes(...)
            doc.clear_children(slot);
            doc.append(slot, phantom);
        }
    }
    root_to_string(&doc, root)
}

// ---- o07 NormalizeDisplayNaryOmml ----------------------------------------------------------------

/// HasNaryLimitContent(limit): any element named "t" (any namespace) with
/// non-whitespace text.
fn has_nary_limit_content(doc: &Doc, limit: Option<NodeId>) -> bool {
    let Some(limit) = limit else { return false };
    doc.descendants_and_self(limit)
        .into_iter()
        .filter(|&element| doc.local(element) == "t")
        .any(|element| !is_null_or_white_space(&doc.value(element)))
}

/// SetNaryLimitVisibility(properties, propertyName, hidden).
fn set_nary_limit_visibility(doc: &mut Doc, properties: NodeId, local: &str, hidden: bool) {
    let property = doc.element(properties, M, local);
    if !hidden {
        if let Some(property) = property {
            doc.detach(property);
        }
        return;
    }
    let property = match property {
        Some(property) => property,
        None => {
            let property = new_element(doc, M, local);
            doc.append(properties, property);
            property
        }
    };
    doc.set_attr_ns(property, M, "val", Some("1"));
}

pub fn normalize_display_nary_omml(omml: &str, display: bool) -> Result<String, String> {
    let (mut doc, root) = parse_preserving(omml)?;
    for nary in document_descendants(&doc, root, M, "nary") {
        let properties = match doc.element(nary, M, "naryPr") {
            Some(properties) => properties,
            None => {
                let properties = new_element(&mut doc, M, "naryPr");
                doc.add_first(nary, properties);
                properties
            }
        };

        // Only display equations force a growing integral/sum; empty limits
        // are hidden for both inline and display OMML.
        let grow = doc.element(properties, M, "grow");
        if display {
            let grow = match grow {
                Some(grow) => grow,
                None => {
                    let grow = new_element(&mut doc, M, "grow");
                    doc.append(properties, grow);
                    grow
                }
            };
            doc.set_attr_ns(grow, M, "val", Some("1"));
        } else if let Some(grow) = grow {
            doc.detach(grow);
        }

        let sub_hidden = !has_nary_limit_content(&doc, doc.element(nary, M, "sub"));
        set_nary_limit_visibility(&mut doc, properties, "subHide", sub_hidden);
        let sup_hidden = !has_nary_limit_content(&doc, doc.element(nary, M, "sup"));
        set_nary_limit_visibility(&mut doc, properties, "supHide", sup_hidden);
    }
    root_to_string(&doc, root)
}

// ---- o08 NormalizeGeneratedOmmlArgumentOrder -------------------------------------------------------

/// RequiredOmmlArgumentNames(ownerName).
fn required_omml_argument_names(owner_name: &str) -> &'static [&'static str] {
    match owner_name {
        "f" => &["num", "den"],
        "rad" => &["deg", "e"],
        "nary" => &["sub", "sup", "e"],
        "sSub" => &["e", "sub"],
        "sSup" => &["e", "sup"],
        "sSubSup" => &["e", "sub", "sup"],
        "sPre" => &["sub", "sup", "e"],
        _ => &[],
    }
}

pub fn normalize_generated_omml_argument_order(omml: &str) -> Result<String, String> {
    let (mut doc, root) = parse_preserving(omml)?;
    let owners: Vec<NodeId> = doc
        .descendants_and_self(root)
        .into_iter()
        .filter(|&element| doc.name(element).is_some_and(|name| name.ns == M))
        .rev()
        .collect();
    for owner in owners {
        let owner_local = doc.local(owner).to_string();
        let roles = required_omml_argument_names(&owner_local);
        if roles.is_empty() {
            continue;
        }
        let property_name = format!("{owner_local}Pr");
        let children = doc.elements(owner);
        let is_property = |doc: &Doc, child: NodeId| doc.is(child, M, &property_name);
        let is_role = |doc: &Doc, child: NodeId, role: &str| doc.is(child, M, role);
        // Normalize the producer's known argument grammar, never reorder
        // operands within an argument or guess through extensions/duplicates.
        if roles.iter().any(|role| children.iter().filter(|&&child| is_role(&doc, child, role)).count() != 1)
            || children.iter().filter(|&&child| is_property(&doc, child)).count() > 1
            || children
                .iter()
                .any(|&child| !is_property(&doc, child) && !roles.iter().any(|role| is_role(&doc, child, role)))
            || doc
                .child_nodes(owner)
                .iter()
                .any(|&node| doc.text(node).is_some_and(|text| !is_null_or_white_space(text)))
        {
            continue;
        }
        let mut ordered: Vec<NodeId> = children.iter().copied().filter(|&child| is_property(&doc, child)).collect();
        for role in roles {
            ordered.push(children.iter().copied().find(|&child| is_role(&doc, child, role)).unwrap());
        }
        if children == ordered {
            continue;
        }
        // Office MML2OMML emits sPre as e/sub/sup, but CT_SPre requires
        // sub/sup/e. owner.ReplaceNodes(ordered.Select(child => new XElement(child))).
        let copies: Vec<NodeId> = ordered.iter().map(|&child| doc.deep_clone(child)).collect();
        doc.clear_children(owner);
        for copy in copies {
            doc.append(owner, copy);
        }
    }
    root_to_string(&doc, root)
}

// ---- TransformMathMlToOmml (o00..o08) --------------------------------------------------------------

/// The o-stages of TransformMathMlToOmml after the projection.
pub fn finish_omml(projection: &str, prepared_mathml: &str, nary_characters: &[String], is_block: bool) -> Result<String, String> {
    let omml = extract_single_omath(projection)?;
    let omml = restore_extended_integral_characters(&omml, nary_characters)?;
    let omml = normalize_explicit_upright_runs(&omml, prepared_mathml)?;
    let omml = normalize_applied_function_structures(&omml, prepared_mathml)?;
    let omml = normalize_explicit_table_column_alignment(&omml, prepared_mathml)?;
    let omml = normalize_omml_placeholder_visibility(&omml)?;
    let omml = normalize_optional_prescript_slots(&omml)?;
    let omml = normalize_display_nary_omml(&omml, is_block)?;
    normalize_generated_omml_argument_order(&omml)
}

// ---- BuildWordNativeNumberedOmml ---------------------------------------------------------------------

/// Guid.TryParseExact(identifier, "N", out _) for a 32-character identifier:
/// after Trim, exactly 32 characters in four 8-character hex groups, each
/// optionally starting with '+' and then "0x" (Guid.TryParseHex).
fn guid_try_parse_exact_n(identifier: &str) -> bool {
    let trimmed = trim(identifier);
    let units: Vec<u16> = trimmed.encode_utf16().collect();
    if units.len() != 32 {
        return false;
    }
    units.chunks(8).all(|chunk| {
        let mut group: &[u16] = chunk;
        if group.first() == Some(&(b'+' as u16)) {
            group = &group[1..];
        }
        if group.len() > 1 && group[0] == b'0' as u16 && (group[1] | 0x20) == b'x' as u16 {
            group = &group[2..];
        }
        group.iter().all(|&unit| unit < 128 && (unit as u8).is_ascii_hexdigit())
    })
}

/// ValidateManagedEquationBookmarkName(bookmarkName, expectedPrefix, parameterName).
fn validate_managed_equation_bookmark_name(bookmark_name: &str, expected_prefix: &str, parameter: &str) -> Result<(), String> {
    let message = format!("The VisualTeX bookmark must use the {expected_prefix}<32-hex-FormulaId> form.");
    if !bookmark_name.starts_with(expected_prefix) || len16(bookmark_name) != len16(expected_prefix) + 32 {
        return Err(argument_error(&message, parameter));
    }
    let identifier = &bookmark_name[expected_prefix.len()..];
    if !guid_try_parse_exact_n(identifier) {
        return Err(argument_error(&message, parameter));
    }
    Ok(())
}

fn has_field_control_character(value: &str) -> bool {
    value.chars().any(|c| matches!(c, '\r' | '\n' | '\u{0013}' | '\u{0014}' | '\u{0015}'))
}

/// BuildWordNativeNumberedOmml: the Word-native `#(SEQ)` equation-number
/// host around a semantic OMML equation.
#[allow(clippy::too_many_arguments)]
pub fn build_word_native_numbered_omml(
    semantic_omml: &str,
    sequence_name: &str,
    number_bookmark_name: Option<&str>,
    restart_heading_level: u32,
    heading_separator: &str,
    initial_sequence_result: &str,
    literal_heading_prefix: Option<&str>,
    restart_sequence_value: Option<u32>,
) -> Result<String, String> {
    if is_null_or_white_space(semantic_omml) {
        return Err(argument_error("The semantic OMML payload must not be empty.", "semanticOmml"));
    }
    if is_null_or_white_space(sequence_name) || len16(sequence_name) > 80 || has_field_control_character(sequence_name) {
        return Err(argument_error("The Word SEQ identifier is invalid.", "sequenceName"));
    }
    if sequence_name.contains('"') || sequence_name.contains('\\') {
        return Err(argument_error("The Word SEQ identifier contains field-syntax characters.", "sequenceName"));
    }
    let sequence_token = if sequence_name.chars().any(is_white_space) {
        format!("\"{sequence_name}\"")
    } else {
        sequence_name.to_string()
    };
    if restart_heading_level > 9 {
        return Err("ArgumentOutOfRangeException: restartHeadingLevel".to_string());
    }
    if is_null_or_white_space(initial_sequence_result) || has_field_control_character(initial_sequence_result) {
        return Err(argument_error("The initial SEQ result is invalid.", "initialSequenceResult"));
    }
    let number_bookmark_name = number_bookmark_name.filter(|name| !is_null_or_white_space(name));
    if let Some(name) = number_bookmark_name {
        validate_managed_equation_bookmark_name(name, NATIVE_NUMBER_BOOKMARK_PREFIX, "numberBookmarkName")?;
    }
    if has_field_control_character(heading_separator) {
        return Err(argument_error(
            "The heading separator contains an invalid Word field-control character.",
            "headingSeparator",
        ));
    }
    if let Some(prefix) = literal_heading_prefix {
        if restart_heading_level == 0
            || is_null_or_white_space(prefix)
            || len16(prefix) > 80
            || prefix.encode_utf16().any(|unit| unit < 0x20 || (0x7F..=0x9F).contains(&unit))
        {
            return Err(argument_error(
                "A literal chapter prefix requires a valid heading scope and plain text.",
                "literalHeadingPrefix",
            ));
        }
    }

    let (mut doc, equation) = parse_preserving(&extract_single_omath(semantic_omml)?)?;
    let formula_nodes: Vec<NodeId> =
        doc.elements(equation).into_iter().map(|element| doc.deep_clone(element)).collect();
    if formula_nodes.is_empty() {
        return Err(invalid_data("A numbered Word equation requires a nonempty mathematical body."));
    }

    let word_run_properties = doc.descendants_named(equation, W, "rPr");
    let representative_run_properties = word_run_properties
        .iter()
        .copied()
        .find(|&properties| doc.element(properties, W, "sz").is_some())
        .or_else(|| word_run_properties.first().copied());

    let mut builder = NumberedBuilder { doc: &mut doc, representative: representative_run_properties };

    let mut number_nodes: Vec<NodeId> = Vec::new();
    if let Some(name) = number_bookmark_name {
        let start = builder.word_element("bookmarkStart");
        builder.doc.set_attr_ns(start, W, "id", Some("31901"));
        builder.doc.set_attr_ns(start, W, "name", Some(name));
        number_nodes.push(start);
    }

    if restart_heading_level > 0 {
        // STYLEREF reads an automatic paragraph number, not a number typed
        // into heading text. Preserve a literal heading as native text.
        match literal_heading_prefix {
            None => {
                let instruction = format!(" STYLEREF {restart_heading_level} \\w \\* MERGEFORMAT ");
                number_nodes.extend(builder.native_field(&instruction, "1"));
            }
            Some(prefix) => {
                let prefix_run = builder.text_run(prefix);
                let normal = builder.math_element("nor");
                let properties = builder.math_element("rPr");
                builder.doc.append(properties, normal);
                builder.doc.add_first(prefix_run, properties);
                number_nodes.push(prefix_run);
            }
        }
        if !heading_separator.is_empty() {
            // A numbering delimiter is literal text, not a mathematical operator.
            let separator_run = builder.text_run(heading_separator);
            let normal = builder.math_element("nor");
            let properties = builder.math_element("rPr");
            builder.doc.append(properties, normal);
            builder.doc.add_first(separator_run, properties);
            number_nodes.push(separator_run);
        }
    }

    if restart_sequence_value.is_some_and(|value| value < 1) {
        return Err("ArgumentOutOfRangeException: restartSequenceValue".to_string());
    }
    let mut sequence_switch = String::new();
    if let Some(value) = restart_sequence_value {
        sequence_switch.push_str(&format!(" \\r {value}"));
    }
    if restart_heading_level > 0 {
        sequence_switch.push_str(&format!(" \\s {restart_heading_level}"));
    }
    let instruction = format!(" SEQ {sequence_token}{sequence_switch} \\* ARABIC \\* MERGEFORMAT ");
    number_nodes.extend(builder.native_field(&instruction, initial_sequence_result));
    if number_bookmark_name.is_some() {
        let end = builder.word_element("bookmarkEnd");
        builder.doc.set_attr_ns(end, W, "id", Some("31901"));
        number_nodes.push(end);
    }

    let control = builder.control_properties();
    let delimiter_properties = builder.math_element("dPr");
    builder.doc.append(delimiter_properties, control);
    let delimiter_body = builder.math_element("e");
    for node in number_nodes {
        builder.doc.append(delimiter_body, node);
    }
    let delimiter = builder.math_element("d");
    builder.doc.append(delimiter, delimiter_properties);
    builder.doc.append(delimiter, delimiter_body);

    let equation_body = builder.math_element("e");
    for node in formula_nodes {
        builder.doc.append(equation_body, node);
    }
    let hash = builder.text_run("#");
    builder.doc.append(equation_body, hash);
    builder.doc.append(equation_body, delimiter);

    let max_dist = builder.math_element("maxDist");
    builder.doc.set_attr_ns(max_dist, M, "val", Some("1"));
    let control = builder.control_properties();
    let array_properties = builder.math_element("eqArrPr");
    builder.doc.append(array_properties, max_dist);
    builder.doc.append(array_properties, control);
    let array = builder.math_element("eqArr");
    builder.doc.append(array, array_properties);
    builder.doc.append(array, equation_body);
    let result = builder.math_element("oMath");
    builder.doc.append(result, array);
    to_string_disable_formatting(&doc, result)
}

struct NumberedBuilder<'a> {
    doc: &'a mut Doc,
    representative: Option<NodeId>,
}

impl NumberedBuilder<'_> {
    fn math_element(&mut self, local: &str) -> NodeId {
        new_element(self.doc, M, local)
    }

    fn word_element(&mut self, local: &str) -> NodeId {
        new_element(self.doc, W, local)
    }

    // RunProperties(noProof, italic)
    fn run_properties(&mut self, no_proof: bool, italic: bool) -> NodeId {
        let properties = self.word_element("rPr");
        for name in ["rFonts", "sz", "szCs"] {
            if let Some(property) = self.representative.and_then(|r| self.doc.element(r, W, name)) {
                let copy = self.doc.deep_clone(property);
                self.doc.append(properties, copy);
            }
        }
        if italic {
            let element = self.word_element("i");
            self.doc.append(properties, element);
        }
        if no_proof {
            let element = self.word_element("noProof");
            self.doc.append(properties, element);
        }
        properties
    }

    // ControlProperties()
    fn control_properties(&mut self) -> NodeId {
        let properties = self.run_properties(false, true);
        let control = self.math_element("ctrlPr");
        self.doc.append(control, properties);
        control
    }

    fn run_with(&mut self, properties: NodeId, content: NodeId) -> NodeId {
        let run = self.math_element("r");
        self.doc.append(run, properties);
        self.doc.append(run, content);
        run
    }

    // BoundaryRun(content)
    fn boundary_run(&mut self, content: NodeId) -> NodeId {
        let properties = self.run_properties(false, true);
        self.run_with(properties, content)
    }

    // InstructionRun(instruction)
    fn instruction_run(&mut self, instruction: &str) -> NodeId {
        let properties = self.run_properties(false, false);
        let text = self.word_element("instrText");
        self.doc.set_attr_ns(text, XML_NS, "space", Some("preserve"));
        self.doc.add_string(text, instruction);
        self.run_with(properties, text)
    }

    // ResultRun(result)
    fn result_run(&mut self, result: &str) -> NodeId {
        let properties = self.run_properties(true, false);
        let text = self.math_element("t");
        self.doc.add_string(text, result);
        self.run_with(properties, text)
    }

    // TextRun(text)
    fn text_run(&mut self, value: &str) -> NodeId {
        let properties = self.run_properties(false, false);
        let text = self.math_element("t");
        self.doc.add_string(text, value);
        self.run_with(properties, text)
    }

    fn field_char(&mut self, kind: &str) -> NodeId {
        let element = self.word_element("fldChar");
        self.doc.set_attr_ns(element, W, "fldCharType", Some(kind));
        element
    }

    // NativeField(instruction, result)
    fn native_field(&mut self, instruction: &str, result: &str) -> Vec<NodeId> {
        let begin = self.field_char("begin");
        let begin_run = self.boundary_run(begin);
        let instruction_run = self.instruction_run(instruction);
        let separate = self.field_char("separate");
        let separate_run = self.boundary_run(separate);
        let result_run = self.result_run(result);
        let end = self.field_char("end");
        let end_run = self.boundary_run(end);
        vec![begin_run, instruction_run, separate_run, result_run, end_run]
    }
}

// ---- ApplyExplicitTransferMathFont / NormalizeMathFontName -------------------

/// WordOmmlConverter.NormalizeMathFontName.
pub fn normalize_math_font_name(math_font_name: &str) -> String {
    if is_null_or_white_space(math_font_name) {
        "Cambria Math".to_string()
    } else {
        trim(math_font_name).to_string()
    }
}

/// WordOmmlConverter.ApplyExplicitTransferMathFont(omml, mathFontName): every
/// mathematical run (m:nor text runs keep their text font) and every m:ctrlPr
/// names the document's OpenType math font in w:rFonts ascii/hAnsi. Windows
/// imports exactly this into the document.
pub fn apply_explicit_transfer_math_font(omml: &str, math_font_name: &str) -> Result<String, String> {
    if is_null_or_white_space(omml) {
        return Ok(omml.to_string());
    }
    let font = normalize_math_font_name(math_font_name);
    let (mut doc, root) = parse_preserving(omml)?;

    fn set_run_fonts(doc: &mut Doc, properties: NodeId, font: &str) {
        let fonts = match doc.element(properties, W, "rFonts") {
            Some(fonts) => fonts,
            None => {
                let fonts = new_element(doc, W, "rFonts");
                doc.add_first(properties, fonts);
                fonts
            }
        };
        doc.set_attr_ns(fonts, W, "ascii", Some(font));
        doc.set_attr_ns(fonts, W, "hAnsi", Some(font));
    }

    for run in document_descendants(&doc, root, M, "r") {
        let math_properties = doc.element(run, M, "rPr");
        if let Some(properties) = math_properties {
            if doc.element(properties, M, "nor").is_some() {
                continue;
            }
        }
        let word_properties = match doc.element(run, W, "rPr") {
            Some(properties) => properties,
            None => {
                let properties = new_element(&mut doc, W, "rPr");
                match math_properties {
                    Some(math) => doc.insert_after(math, properties),
                    None => doc.add_first(run, properties),
                }
                properties
            }
        };
        set_run_fonts(&mut doc, word_properties, &font);
    }
    for control in document_descendants(&doc, root, M, "ctrlPr") {
        let word_properties = match doc.element(control, W, "rPr") {
            Some(properties) => properties,
            None => {
                let properties = new_element(&mut doc, W, "rPr");
                doc.append(control, properties);
                properties
            }
        };
        set_run_fonts(&mut doc, word_properties, &font);
    }
    root_to_string(&doc, root)
}
