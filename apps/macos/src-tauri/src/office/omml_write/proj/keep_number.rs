//! A numbered Word equation edited where its number is not rebuilt (the WPS
//! add-in, office::wps): the edited formula goes back into the equation array
//! in front of the original "#" and (number), which keep their fields,
//! bookmarks and prefix text.
//!
//! Word writes a field instruction inside an equation as math text, and WPS
//! turns every "*" of math text into U+2217 when it reads a document, so the
//! "\* ARABIC \* MERGEFORMAT" switches of a SEQ or STYLEREF come back from WPS
//! as "\<U+2217> ...", which Word reads as a bookmark name ("Error! Bookmark not
//! defined."). The kept number's instructions are written as normal text
//! (m:nor), which WPS leaves alone, with the switch character restored.

use super::converter::extract_single_omath;
use crate::office::dotnet::dom::{Doc, Name, NodeId, OMML_NS, WORD_NS, XMLNS_NS};
use super::xml_reader::{self, PRESERVE_WHITESPACE};
use super::xml_writer::to_string_disable_formatting;

const M: &str = OMML_NS;
const W: &str = WORD_NS;

fn parse_equation(xml: &str) -> Result<(Doc, NodeId), String> {
    let doc = xml_reader::parse(&extract_single_omath(xml)?, PRESERVE_WHITESPACE)?;
    let root = doc.root.ok_or_else(|| "XmlException: Root element is missing.".to_string())?;
    Ok((doc, root))
}

fn run_text(doc: &Doc, run: NodeId) -> String {
    doc.elements_named(run, M, "t").into_iter().map(|text| doc.value(text)).collect()
}

/// The row of a numbered equation and the index of its "#" run: the equation
/// is one equation array with one row that ends in a run ending with "#", the
/// (number) delimiter and control properties (the shape the OMML reader
/// strips before editing, try_resolve_word_native_equation_number_host).
fn number_slot(doc: &Doc, equation: NodeId) -> Option<(NodeId, usize)> {
    let children: Vec<NodeId> =
        doc.elements(equation).into_iter().filter(|&element| !doc.is(element, M, "ctrlPr")).collect();
    if children.len() != 1 || !doc.is(children[0], M, "eqArr") {
        return None;
    }
    let rows = doc.elements_named(children[0], M, "e");
    if rows.len() != 1 {
        return None;
    }
    let items = doc.elements(rows[0]);
    (0..items.len().saturating_sub(1))
        .find(|&index| {
            doc.is(items[index], M, "r")
                && run_text(doc, items[index]).ends_with('#')
                && doc.is(items[index + 1], M, "d")
                && items[index + 2..].iter().all(|&element| doc.is(element, M, "ctrlPr") || is_marker(doc, element))
        })
        .map(|index| (rows[0], index))
}

/// A bookmark or proofing mark, which an editor may leave anywhere in a row.
fn is_marker(doc: &Doc, element: NodeId) -> bool {
    ["bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd"]
        .iter()
        .any(|local| doc.is(element, W, local))
}

/// Every run that holds field instruction text (between a field's begin and
/// separate characters) becomes a normal-text run with "*" switches.
fn normalize_field_instructions(doc: &mut Doc, number: NodeId) {
    // One entry per open field: true while its instruction is being read.
    let mut fields: Vec<bool> = Vec::new();
    for run in doc.descendants_named(number, M, "r") {
        let mut instruction = false;
        for child in doc.elements(run) {
            if doc.is(child, W, "fldChar") {
                match doc.attr_ns(child, W, "fldCharType") {
                    Some("begin") => fields.push(true),
                    Some("separate") => {
                        if let Some(reading) = fields.last_mut() {
                            *reading = false;
                        }
                    }
                    Some("end") => {
                        fields.pop();
                    }
                    _ => {}
                }
            } else if fields.last() == Some(&true) && (doc.is(child, M, "t") || doc.is(child, W, "instrText")) {
                instruction = true;
                let text = doc.value(child).replace('\u{2217}', "*");
                if doc.is(child, W, "instrText") {
                    doc.set_name(child, Name::new(M, "t"));
                    // A namespace declaration for its old name (the shared
                    // writer declares w as the default namespace there) would
                    // clash with the new one.
                    let declarations: Vec<Name> = doc
                        .attrs(child)
                        .iter()
                        .map(|(name, _)| name.clone())
                        .filter(Doc::is_namespace_declaration)
                        .collect();
                    for name in declarations {
                        doc.set_attr_ns(child, &name.ns, &name.local, None);
                    }
                }
                doc.set_value(child, &text);
            }
        }
        if instruction {
            let properties = match doc.element(run, M, "rPr") {
                Some(properties) => properties,
                None => {
                    let properties = doc.new_element(Name::new(M, "rPr"));
                    doc.add_first(run, properties);
                    properties
                }
            };
            // m:nor and m:scr/m:sty are alternatives in a math run's properties.
            for style in doc.elements(properties) {
                if doc.is(style, M, "scr") || doc.is(style, M, "sty") {
                    doc.detach(style);
                }
            }
            if doc.element(properties, M, "nor").is_none() {
                let normal = doc.new_element(Name::new(M, "nor"));
                match doc.element(properties, M, "lit") {
                    Some(literal) => doc.insert_after(literal, normal),
                    None => doc.add_first(properties, normal),
                }
            }
        }
    }
}

/// The equation spelled with the m: and w: prefixes declared once on its
/// root. The shared writer builds a numbered equation from new elements, so it
/// comes out with default-namespace declarations (as Windows writes it), which
/// Word reads but WPS does not take for an equation.
fn use_m_and_w_prefixes(doc: &mut Doc, equation: NodeId) {
    for element in doc.descendants_and_self(equation) {
        let declarations: Vec<Name> = doc
            .attrs(element)
            .iter()
            .map(|(name, _)| name.clone())
            .filter(Doc::is_namespace_declaration)
            .collect();
        for name in declarations {
            doc.set_attr_ns(element, &name.ns, &name.local, None);
        }
    }
    doc.set_attr_ns(equation, XMLNS_NS, "m", Some(M));
    doc.set_attr_ns(equation, XMLNS_NS, "w", Some(W));
}

/// An equation whose field instructions WPS keeps intact: each instruction run
/// written as normal-text math (see the module comment), in m:/w: spelling.
pub fn instructions_as_normal_text(omml: &str) -> Result<String, String> {
    let (mut doc, equation) = parse_equation(omml)?;
    normalize_field_instructions(&mut doc, equation);
    use_m_and_w_prefixes(&mut doc, equation);
    to_string_disable_formatting(&doc, equation)
}

/// The formula of a numbered equation (Word's export of it, or its m:oMath)
/// without its "#" and (number), with its runs as they are: what a number in
/// another style is built around when WPS renumbers the equation.
pub fn equation_body(numbered_xml: &str) -> Result<String, String> {
    let (mut doc, equation) = parse_equation(numbered_xml)?;
    let (row, hash) = number_slot(&doc, equation)
        .ok_or_else(|| "InvalidDataException: The Word equation has no equation number.".to_string())?;
    let array = doc
        .parent(row)
        .ok_or_else(|| "InvalidDataException: The Word equation number has no equation array.".to_string())?;
    let items = doc.elements(row);
    let separator = items[hash];
    // A "#" merged into the formula's last run ("+1=0#") leaves that run's text.
    let texts = doc.elements_named(separator, M, "t");
    if let Some(&last) = texts.last() {
        let value = doc.value(last);
        let formula_text = value.strip_suffix('#').unwrap_or(&value).to_string();
        doc.set_value(last, &formula_text);
    }
    let keep_separator = texts.iter().any(|&text| !doc.value(text).is_empty());
    for &item in &items[..hash] {
        doc.detach(item);
        doc.insert_before(array, item);
    }
    if keep_separator {
        doc.detach(separator);
        doc.insert_before(array, separator);
    }
    doc.detach(array);
    if doc.elements(equation).iter().all(|&element| doc.is(element, M, "ctrlPr")) {
        return Err("InvalidDataException: The numbered Word equation has no formula.".to_string());
    }
    to_string_disable_formatting(&doc, equation)
}

/// The edited equation (one m:oMath) with the number of the original numbered
/// equation (Word's export of it, or its m:oMath).
pub fn keep_equation_number(original_xml: &str, edited_omml: &str) -> Result<String, String> {
    let (mut doc, equation) = parse_equation(original_xml)?;
    let (row, hash) = number_slot(&doc, equation)
        .ok_or_else(|| "InvalidDataException: The original Word equation has no equation number to keep.".to_string())?;
    let (edited, edited_root) = parse_equation(edited_omml)?;
    let formula = edited.elements(edited_root);
    if formula.is_empty() {
        return Err("InvalidDataException: The edited Word equation is empty.".to_string());
    }

    let items = doc.elements(row);
    let separator = items[hash];
    let number = items[hash + 1];
    for &item in &items[..hash] {
        doc.detach(item);
    }
    // A "#" merged into the old formula's last run ("+1=0#") keeps only "#".
    let texts = doc.elements_named(separator, M, "t");
    if run_text(&doc, separator) != "#" {
        for (index, &text) in texts.iter().enumerate() {
            if index + 1 == texts.len() {
                doc.set_value(text, "#");
            } else {
                doc.detach(text);
            }
        }
    }
    for &element in &formula {
        let copy = doc.import(&edited, element);
        doc.insert_before(separator, copy);
    }
    normalize_field_instructions(&mut doc, number);
    to_string_disable_formatting(&doc, equation)
}

#[cfg(test)]
mod tests {
    use super::keep_equation_number;

    const NS: &str = r#"xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

    /// A VisualTeX numbered equation as WPS exports it after reading it: the
    /// SEQ switches turned into U+2217, a bookmark around the number.
    fn wps_numbered() -> String {
        format!(
            concat!(
                r#"<w:p {ns}><m:oMathPara><m:oMath><m:eqArr><m:eqArrPr><m:maxDist m:val="1"/></m:eqArrPr><m:e>"#,
                r#"<m:r><m:t>a+b</m:t></m:r><m:r><m:t>#</m:t></m:r>"#,
                r#"<m:d><m:dPr><m:ctrlPr><w:rPr><w:i/></w:rPr></m:ctrlPr></m:dPr><m:e>"#,
                r#"<w:bookmarkStart w:id="7" w:name="_Ref123"/>"#,
                r#"<m:r><m:rPr><m:nor/><m:sty m:val="p"/></m:rPr><m:t>0.0-</m:t></m:r>"#,
                r#"<m:r><w:rPr><w:i/></w:rPr><w:fldChar w:fldCharType="begin"/></m:r>"#,
                r#"<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t xml:space="preserve"> SEQ Equation \s 2 \∗ ARABIC \∗ MERGEFORMAT </m:t></m:r>"#,
                r#"<m:r><w:fldChar w:fldCharType="separate"/></m:r><m:r><m:t>4</m:t></m:r>"#,
                r#"<m:r><w:fldChar w:fldCharType="end"/></m:r><w:bookmarkEnd w:id="7"/>"#,
                r#"</m:e></m:d></m:e></m:eqArr></m:oMath></m:oMathPara></w:p>"#,
            ),
            ns = NS
        )
    }

    #[test]
    fn the_edited_formula_keeps_the_number_and_its_fields() {
        let edited = format!(r#"<m:oMath {NS}><m:r><m:t>x</m:t></m:r><m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>"#);
        let kept = keep_equation_number(&wps_numbered(), &edited).unwrap();
        assert!(!kept.contains("a+b"), "{kept}");
        assert!(kept.contains("<m:t>x</m:t>") && kept.contains("<m:sSup>"), "{kept}");
        // The number: "#", bookmark, prefix, field and its result, in order.
        let order = ["<m:t>y</m:t>", "<m:t>#</m:t>", "_Ref123", "0.0-", "begin", " SEQ Equation", "separate", "<m:t>4</m:t>", "end"];
        let mut from = 0;
        for part in order {
            let at = kept[from..].find(part).unwrap_or_else(|| panic!("{part} missing after {from}: {kept}")) + from;
            from = at;
        }
        // The instruction is normal text with ASCII switches.
        assert!(kept.contains(r#"<m:r><m:rPr><m:nor /></m:rPr><m:t xml:space="preserve"> SEQ Equation \s 2 \* ARABIC \* MERGEFORMAT </m:t></m:r>"#)
            || kept.contains(r#"<m:r><m:rPr><m:nor/></m:rPr><m:t xml:space="preserve"> SEQ Equation \s 2 \* ARABIC \* MERGEFORMAT </m:t></m:r>"#), "{kept}");
        assert!(!kept.contains('\u{2217}'), "{kept}");
    }

    #[test]
    fn the_body_of_a_numbered_equation_drops_the_number() {
        let body = super::equation_body(&wps_numbered()).unwrap();
        assert!(body.starts_with("<m:oMath"), "{body}");
        assert!(body.contains("<m:t>a+b</m:t>"), "{body}");
        for gone in ["eqArr", "#", "SEQ", "0.0-", "_Ref123", "<m:d>"] {
            assert!(!body.contains(gone), "{gone} in {body}");
        }
        // A "#" merged into the formula's last run leaves that run's text.
        let merged = wps_numbered().replace("<m:r><m:t>a+b</m:t></m:r><m:r><m:t>#</m:t></m:r>", "<m:r><m:t>a+b#</m:t></m:r>");
        let body = super::equation_body(&merged).unwrap();
        assert!(body.contains("<m:t>a+b</m:t>") && !body.contains('#'), "{body}");
        let plain = format!(r#"<m:oMath {NS}><m:r><m:t>z</m:t></m:r></m:oMath>"#);
        assert!(super::equation_body(&plain).is_err());
        // A bookmark an editor left after the number does not hide it.
        let marked = wps_numbered().replace("</m:e></m:d></m:e></m:eqArr>", r#"</m:e></m:d><w:bookmarkEnd w:id="9"/></m:e></m:eqArr>"#);
        let body = super::equation_body(&marked).unwrap();
        assert!(body.contains("<m:t>a+b</m:t>") && !body.contains("SEQ"), "{body}");
    }

    #[test]
    fn a_hash_merged_into_the_old_formula_is_split_off() {
        let original = wps_numbered().replace("<m:r><m:t>a+b</m:t></m:r><m:r><m:t>#</m:t></m:r>", "<m:r><m:t>a+b#</m:t></m:r>");
        let edited = format!(r#"<m:oMath {NS}><m:r><m:t>z</m:t></m:r></m:oMath>"#);
        let kept = keep_equation_number(&original, &edited).unwrap();
        assert!(kept.contains("<m:r><m:t>z</m:t></m:r><m:r><m:t>#</m:t></m:r><m:d>"), "{kept}");
        assert!(!kept.contains("a+b"), "{kept}");
    }

    #[test]
    fn an_instruction_in_instr_text_becomes_math_text() {
        let original = wps_numbered().replace(
            r#"<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t xml:space="preserve"> SEQ Equation \s 2 \∗ ARABIC \∗ MERGEFORMAT </m:t></m:r>"#,
            r#"<m:r><w:rPr/><w:instrText xml:space="preserve"> SEQ Equation \* ARABIC </w:instrText></m:r>"#,
        );
        let edited = format!(r#"<m:oMath {NS}><m:r><m:t>z</m:t></m:r></m:oMath>"#);
        let kept = keep_equation_number(&original, &edited).unwrap();
        assert!(!kept.contains("instrText"), "{kept}");
        assert!(kept.contains(r#" SEQ Equation \* ARABIC </m:t>"#), "{kept}");
    }

    #[test]
    fn instruction_runs_become_normal_text() {
        let numbered = format!(
            concat!(
                r#"<m:oMath {ns}><m:r><m:t>x#</m:t></m:r><m:d><m:e>"#,
                r#"<m:r><w:rPr/><w:fldChar w:fldCharType="begin"/></m:r>"#,
                r#"<m:r><w:rPr/><w:instrText xml:space="preserve"> SEQ Equation \* ARABIC \* MERGEFORMAT </w:instrText></m:r>"#,
                r#"<m:r><w:fldChar w:fldCharType="separate"/></m:r><m:r><m:t>1</m:t></m:r><m:r><w:fldChar w:fldCharType="end"/></m:r>"#,
                r#"</m:e></m:d></m:oMath>"#,
            ),
            ns = NS
        );
        let normal = super::instructions_as_normal_text(&numbered).unwrap();
        assert!(normal.contains("<m:nor"), "{normal}");
        assert!(normal.contains(r#" SEQ Equation \* ARABIC \* MERGEFORMAT </m:t>"#), "{normal}");
        assert!(!normal.contains("instrText"), "{normal}");
        // The result digit and the formula stay ordinary math text.
        assert!(normal.contains("<m:r><m:t>1</m:t></m:r>"), "{normal}");
    }

    #[test]
    fn a_number_from_the_shared_writer_becomes_wps_safe() {
        let body = format!(r#"<m:oMath {NS}><m:r><m:t>x</m:t></m:r></m:oMath>"#);
        let numbered = crate::office::omml_batch::numbered_equation(&body, "Equation", 2, "-", Some("0.0"), Some(1), 1, "Cambria Math").unwrap();
        let normal = super::instructions_as_normal_text(&numbered).unwrap();
        // Spelled with m: and w: prefixes (WPS does not read the shared
        // writer's default-namespace spelling as an equation).
        assert!(normal.starts_with(r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#), "{normal}");
        assert!(!normal.contains("xmlns=") && !normal.contains("xmlns:p"), "{normal}");
        assert!(normal.contains(r#"<m:r><m:rPr><m:nor /></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" /></w:rPr><m:t xml:space="preserve"> SEQ Equation \r 1 \s 2 \* ARABIC \* MERGEFORMAT </m:t></m:r>"#), "{normal}");
        assert!(normal.contains(r#"<w:fldChar w:fldCharType="begin" />"#), "{normal}");
        assert!(!normal.contains("instrText"), "{normal}");
    }

    #[test]
    fn an_unnumbered_equation_is_refused() {
        let plain = format!(r#"<m:oMath {NS}><m:r><m:t>a#</m:t></m:r><m:d><m:e><m:r><m:t>1</m:t></m:r></m:e></m:d></m:oMath>"#);
        let edited = format!(r#"<m:oMath {NS}><m:r><m:t>z</m:t></m:r></m:oMath>"#);
        assert!(keep_equation_number(&plain, &edited).is_err());
    }
}
