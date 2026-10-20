//! Word equations (MathML -> OMML), ported from VisualTeX for Windows' C# Word
//! writer so both platforms write identical Word equations
//! (docs/equation-format-v2.md):
//! - `prep`: WordOmmlConverter's MathML passes before the projection
//!   (ValidateMathMlForOmml ... ReplaceExtendedIntegralsWithOfficePlaceholders);
//! - `proj`: MathMlOmmlProjection.Project, the OMML passes after it
//!   (ExtractSingleOMath ... NormalizeGeneratedOmmlArgumentOrder) and
//!   BuildWordNativeNumberedOmml, Word's native #(SEQ) equation number.
//!
//! The port is held to the C# output by the Windows forward corpus
//! (ForwardDump in the Windows tests): `fixtures/windows_forward.json` holds a
//! reference subset checked by every test run, and the whole corpus runs as an
//! ignored test.

use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};

#[allow(dead_code, deprecated, unused_imports)]
mod prep;
#[allow(dead_code, deprecated, unused_imports)]
mod proj;

/// WordOmmlConverter.TransformMathMlToOmml: Presentation MathML (as Windows'
/// Office dialog produces it) to one m:oMath.
pub(crate) fn transform_mathml_to_omml(mathml: &str) -> Result<String, String> {
    let prepared = prep::prepare_mathml_for_omml(mathml)?;
    let projection = proj::project(&prepared.mathml)?;
    proj::finish_omml(
        &projection,
        &prepared.mathml,
        &prepared.nary_characters,
        prepared.is_block,
    )
}

/// WordOmmlConverter.BuildWordNativeNumberedOmml.
#[allow(clippy::too_many_arguments)]
pub(crate) fn build_word_native_numbered_omml(
    semantic_omml: &str,
    sequence_name: &str,
    number_bookmark_name: Option<&str>,
    restart_heading_level: u32,
    heading_separator: &str,
    initial_sequence_result: &str,
    literal_heading_prefix: Option<&str>,
    restart_sequence_value: Option<u32>,
) -> Result<String, String> {
    proj::build_word_native_numbered_omml(
        semantic_omml,
        sequence_name,
        number_bookmark_name,
        restart_heading_level,
        heading_separator,
        initial_sequence_result,
        literal_heading_prefix,
        restart_sequence_value,
    )
}

/// WordOmmlConverter.ExtractSingleOMath: the one m:oMath of Word's export of
/// an equation, with the namespaces it uses.
pub(crate) fn extract_single_omath(xml: &str) -> Result<String, String> {
    proj::extract_single_omath(xml)
}

/// What Windows imports into the document for an equation:
/// ApplyExplicitTransferMathFont(ExtractSingleOMath(omml), document math font).
/// Every mathematical run (not m:nor text) and every control property names the
/// document's OpenType math font explicitly.
pub(crate) fn insertion_omml(omml: &str, math_font_name: &str) -> Result<String, String> {
    proj::apply_explicit_transfer_math_font(&proj::extract_single_omath(omml)?, math_font_name)
}

/// A one-equation-per-paragraph package for Word's InsertFile/Documents.Open
/// (Word for Mac cannot import OMML through Range.InsertXML). Its settings name
/// the document's math font, as Windows' staging packages do.
pub(crate) fn word_equation_docx(equations: &[String], math_font_name: &str) -> Result<Vec<u8>, String> {
    if equations.is_empty() || equations.len() > 1000 {
        return Err("A Word equation package must contain 1 to 1000 equations".to_string());
    }
    let font = proj::normalize_math_font_name(math_font_name);
    let mut document = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" \
         xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body>",
    );
    for equation in equations {
        if equation.contains("<!DOCTYPE") || equation.contains("<!ENTITY") {
            return Err("A Word equation contains a document type declaration".to_string());
        }
        document.push_str("<w:p>");
        document.push_str(equation);
        document.push_str("</w:p>");
    }
    document.push_str("<w:sectPr/></w:body></w:document>");
    let settings = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" \
         xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">\
         <m:mathPr><m:mathFont m:val=\"{}\"/></m:mathPr></w:settings>",
        font.replace('&', "&amp;").replace('<', "&lt;").replace('"', "&quot;")
    );
    super::macos_offline::build_stored_zip(&[
        (
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
             <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
             <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
             <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\
             <Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>\
             </Types>"
                .as_bytes(),
        ),
        (
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
             <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\
             </Relationships>"
                .as_bytes(),
        ),
        (
            "word/_rels/document.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
             <Relationship Id=\"rIdSettings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>\
             </Relationships>"
                .as_bytes(),
        ),
        ("word/document.xml", document.as_bytes()),
        ("word/settings.xml", settings.as_bytes()),
    ])
}

/// Everything the Word commit needs for one equation from the editor's MathML:
/// the semantic OMML (Base64URL, kept for later conversions), the equation as
/// it is imported, and the one-paragraph package the add-in inserts.
pub(crate) struct WordEquationArtifacts {
    pub omml_base64: String,
    pub insertion_omml: String,
    pub docx: Vec<u8>,
}

pub(crate) fn word_equation_artifacts(
    mathml: &str,
    math_font_name: &str,
) -> Result<WordEquationArtifacts, String> {
    let omml = transform_mathml_to_omml(mathml)?;
    let insertion = insertion_omml(&omml, math_font_name)?;
    let docx = word_equation_docx(std::slice::from_ref(&insertion), math_font_name)?;
    Ok(WordEquationArtifacts {
        omml_base64: URL_SAFE_NO_PAD.encode(omml.as_bytes()),
        insertion_omml: insertion,
        docx,
    })
}

#[cfg(test)]
mod tests {
    use std::fmt::Write as _;

    fn first_difference(expected: &str, actual: &str) -> String {
        let position = expected
            .char_indices()
            .zip(actual.chars())
            .find(|((_, e), a)| e != a)
            .map(|((index, _), _)| index)
            .unwrap_or_else(|| expected.len().min(actual.len()));
        let from = expected[..position].char_indices().rev().nth(40).map(|(i, _)| i).unwrap_or(0);
        format!(
            "at {position}: expected …{}… got …{}…",
            &expected[from..expected.len().min(position + 80)],
            &actual[from.min(actual.len())..actual.len().min(position + 80)]
        )
    }

    /// Windows' MathML -> OMML for every row: TransformMathMlToOmml and, where
    /// present, BuildWordNativeNumberedOmml in the five number styles.
    fn check_rows(rows: &[serde_json::Value]) -> (usize, Vec<String>) {
        const STYLES: [(&str, &str, u32, &str, &str, Option<&str>, Option<u32>); 5] = [
            ("continuous", "Equation", 0, "", "1", None, None),
            ("h1dot-literal", "\u{516c}\u{5f0f}", 1, ".", "3", Some("2"), None),
            ("h1dash-auto-restart", "Equation", 1, "-", "1", None, Some(1)),
            ("h2dot-literal-spaced", "My Eq", 2, ".", "12", Some("2.3"), None),
            ("h2dash-auto-restart", "\u{516c}\u{5f0f}", 2, "-", "1", None, Some(1)),
        ];
        let mut failures = Vec::new();
        let mut checked = 0;
        for row in rows {
            let source = row["source"].as_str().unwrap_or("?");
            let (Some(mathml), Some(expected)) = (row["mathml"].as_str(), row["omml"].as_str()) else {
                continue;
            };
            checked += 1;
            match super::transform_mathml_to_omml(mathml) {
                Ok(actual) if actual == expected => {}
                Ok(actual) => failures.push(format!("{source}: OMML {}", first_difference(expected, &actual))),
                Err(error) => failures.push(format!("{source}: {error}")),
            }
            if let Some(numbered) = row.get("numbered") {
                for (name, label, level, separator, initial, literal, restart) in STYLES {
                    let Some(expected) = numbered[name].as_str() else { continue };
                    match super::build_word_native_numbered_omml(
                        row["omml"].as_str().unwrap(), label, None, level, separator, initial, literal, restart,
                    ) {
                        Ok(actual) if actual == expected => {}
                        Ok(actual) => failures.push(format!("{source} {name}: {}", first_difference(expected, &actual))),
                        Err(error) => failures.push(format!("{source} {name}: {error}")),
                    }
                }
            }
            if let Some(insertions) = row.get("numberedInsertion") {
                for (name, label, level, separator, initial, literal, restart) in STYLES {
                    let Some(expected) = insertions[name].as_str() else { continue };
                    let actual = super::build_word_native_numbered_omml(
                        row["omml"].as_str().unwrap(), label, None, level, separator, initial, literal, restart,
                    )
                    .and_then(|numbered| super::insertion_omml(&numbered, "Cambria Math"));
                    match actual {
                        Ok(actual) if actual == expected => {}
                        Ok(actual) => failures.push(format!("{source} {name} insertion: {}", first_difference(expected, &actual))),
                        Err(error) => failures.push(format!("{source} {name} insertion: {error}")),
                    }
                }
            }
            if let Some(expected) = row["insertion"].as_str() {
                match super::insertion_omml(row["omml"].as_str().unwrap(), "Cambria Math") {
                    Ok(actual) if actual == expected => {}
                    Ok(actual) => failures.push(format!("{source} insertion: {}", first_difference(expected, &actual))),
                    Err(error) => failures.push(format!("{source} insertion: {error}")),
                }
            }
        }
        (checked, failures)
    }

    #[test]
    fn reproduces_the_windows_word_writer() {
        let rows: Vec<serde_json::Value> =
            serde_json::from_str(include_str!("fixtures/windows_forward.json")).unwrap();
        let (checked, failures) = check_rows(&rows);
        assert!(checked > 100, "only {checked} reference rows");
        assert!(failures.is_empty(), "{} of {checked} differ:\n{}", failures.len(), failures.join("\n"));
    }

    /// The whole Windows forward corpus (ForwardDump in the Windows tests):
    /// `VT_OMML_FORWARD_CORPUS=<forward.json> cargo test --lib -- --ignored full_windows_forward`.
    #[test]
    #[ignore]
    fn reproduces_the_full_windows_forward_corpus() {
        let path = std::env::var("VT_OMML_FORWARD_CORPUS").expect("set VT_OMML_FORWARD_CORPUS");
        let rows: Vec<serde_json::Value> =
            serde_json::from_str(&std::fs::read_to_string(path).unwrap()).unwrap();
        let (checked, failures) = check_rows(&rows);
        let mut report = String::new();
        for failure in failures.iter().take(20) {
            let _ = writeln!(report, "{failure}");
        }
        assert!(failures.is_empty(), "{} of {checked} differ:\n{report}", failures.len());
    }

    #[test]
    fn a_numbered_equation_package_holds_the_shared_number() {
        let omml = super::transform_mathml_to_omml(
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></math>",
        )
        .unwrap();
        let numbered = super::build_word_native_numbered_omml(&omml, "Equation", None, 1, "-", "1", None, Some(1)).unwrap();
        let insertion = super::insertion_omml(&numbered, "Cambria Math").unwrap();
        assert!(insertion.contains(" STYLEREF 1 \\w \\* MERGEFORMAT "));
        assert!(insertion.contains(" SEQ Equation \\r 1 \\s 1 \\* ARABIC \\* MERGEFORMAT "));
        assert!(insertion.contains("Cambria Math"));
        let docx = super::word_equation_docx(&[insertion], "Cambria Math").unwrap();
        assert!(docx.starts_with(b"PK\x03\x04"));
    }
}
