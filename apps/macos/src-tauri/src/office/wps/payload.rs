//! The document state the WPS add-in writes for a VisualTeX formula.
//!
//! It is exactly what the Word-for-Mac add-in stores (VTWordImages.bas, VTWordFormulaData.bas): the
//! picture's Title and AlternativeText, the owner bookmark and the formula's
//! document variables. A document can therefore move between WPS and Word and
//! stay editable in both. The names and formats are kept here, in one place,
//! so the add-in's JavaScript only writes what it is given.

use serde::Serialize;

pub(crate) const FORMULA_REF_PREFIX: &str = "visualtex:formula-ref:v1:";
const PAYLOAD_CHUNK_SIZE: usize = 20_000;
const PAYLOAD_MAX_CHUNKS: usize = 128;
const LATEX_PREFIX: &str = "VT_Latex_";
const OMML_PREFIX: &str = "VT_OMML_";
const METADATA_PREFIX: &str = "VT_Metadata_";
const FORMAT_PREFIX: &str = "VT_Format_";
const IMAGE_SCALE_PREFIX: &str = "VT_ImageScale_";
const IMAGE_INK_CENTER_PREFIX: &str = "VT_ImageInkCenter_";
const IMAGE_LETTER_FONT_PREFIX: &str = "VT_ImageLetterFont_";
const OWNER_BOOKMARK_PREFIX: &str = "VT_I_";
const KNOWN_LETTER_FONTS: [&str; 6] = ["katex", "times", "cambria", "stix", "palatino", "helvetica"];

/// What an image formula's document state is built from.
pub(crate) struct WordImageState<'a> {
    pub formula_id: &'a str,
    pub display_mode: &'a str,
    pub numbered: bool,
    pub latex_base64: &'a str,
    pub omml_base64: &'a str,
    pub metadata: &'a str,
    pub font_size_pt: f64,
    pub reference_width_pt: f64,
    pub reference_height_pt: f64,
    pub reference_baseline_pt: f64,
    pub ink_center_y_ratio: Option<f64>,
    pub letter_font: &'a str,
}

/// Document variables to set, and the chunked payloads whose chunks beyond
/// the new count must be deleted (a shorter formula needs fewer chunks).
#[derive(Debug, Clone, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct WordVariables {
    pub set: Vec<(String, String)>,
    pub chunked: Vec<ChunkedPayload>,
}

#[derive(Debug, Clone, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ChunkedPayload {
    pub stem: String,
    pub count: usize,
    pub max_chunks: usize,
}

/// `visualtex:formula-ref:v1:<id>:<inline|block>:<0|1>` (VTFormulaReference).
pub(crate) fn formula_reference(formula_id: &str, display_mode: &str, numbered: bool) -> String {
    format!(
        "{FORMULA_REF_PREFIX}{formula_id}:{display_mode}:{}",
        if numbered { "1" } else { "0" }
    )
}

/// `VT_I_` plus the formula id without dashes (VTWordImageOwnerBookmarkName).
pub(crate) fn owner_bookmark_name(formula_id: &str) -> String {
    format!("{OWNER_BOOKMARK_PREFIX}{}", formula_id.replace('-', ""))
}

fn stem(prefix: &str, formula_id: &str) -> String {
    format!("{prefix}{}", formula_id.replace('-', "_"))
}

/// A number as the VBA reader (Val) accepts it: plain decimal, no exponent.
fn number_text(value: f64) -> String {
    let text = format!("{value:.6}");
    let text = text.trim_end_matches('0').trim_end_matches('.');
    if text.is_empty() || text == "-0" { "0".to_string() } else { text.to_string() }
}

fn push_chunked(
    variables: &mut WordVariables,
    prefix: &str,
    formula_id: &str,
    payload: &str,
    label: &str,
) -> Result<(), String> {
    if payload.is_empty() {
        return Err(format!("The {label} payload is empty"));
    }
    let chunks: Vec<&str> = payload
        .as_bytes()
        .chunks(PAYLOAD_CHUNK_SIZE)
        .map(|chunk| std::str::from_utf8(chunk).map_err(|_| format!("The {label} payload is not ASCII")))
        .collect::<Result<_, _>>()?;
    if chunks.len() > PAYLOAD_MAX_CHUNKS {
        return Err(format!("The {label} payload needs more than {PAYLOAD_MAX_CHUNKS} document variables"));
    }
    let stem = stem(prefix, formula_id);
    for (index, chunk) in chunks.iter().enumerate() {
        variables.set.push((format!("{stem}_{:03}", index + 1), (*chunk).to_string()));
    }
    variables.set.push((format!("{stem}_Count"), chunks.len().to_string()));
    variables.chunked.push(ChunkedPayload { stem, count: chunks.len(), max_chunks: PAYLOAD_MAX_CHUNKS });
    Ok(())
}

/// The variables the Word-for-Mac add-in writes when it commits an image
/// formula (VTSetWordLatexPayload … VTSetWordImageLetterFont).
pub(crate) fn word_image_variables(state: &WordImageState) -> Result<WordVariables, String> {
    if !matches!(state.display_mode, "inline" | "block") {
        return Err("The formula display mode must be inline or block".to_string());
    }
    if state.numbered && state.display_mode != "block" {
        return Err("Only display formulas can be numbered".to_string());
    }
    for (value, label) in [
        (state.latex_base64, "LaTeX"),
        (state.omml_base64, "OMML"),
    ] {
        if !value.bytes().all(|byte| byte.is_ascii_alphanumeric() || byte == b'-' || byte == b'_') {
            return Err(format!("The {label} payload is not Base64URL"));
        }
    }
    if !state.metadata.is_ascii() {
        return Err("The formula metadata envelope is not ASCII".to_string());
    }
    if !(state.reference_width_pt > 0.0 && state.reference_height_pt > 0.0)
        || !(-256.0..=0.0).contains(&state.reference_baseline_pt)
        || !(state.font_size_pt.is_finite() && state.font_size_pt > 0.0)
    {
        return Err("The formula image geometry is invalid".to_string());
    }
    let mut variables = WordVariables { set: Vec::new(), chunked: Vec::new() };
    push_chunked(&mut variables, LATEX_PREFIX, state.formula_id, state.latex_base64, "LaTeX")?;
    push_chunked(&mut variables, OMML_PREFIX, state.formula_id, state.omml_base64, "OMML")?;
    push_chunked(&mut variables, METADATA_PREFIX, state.formula_id, state.metadata, "metadata")?;
    variables.set.push((
        stem(FORMAT_PREFIX, state.formula_id),
        format!("{}|{}", state.display_mode, if state.numbered { "1" } else { "0" }),
    ));
    variables.set.push((
        stem(IMAGE_SCALE_PREFIX, state.formula_id),
        [
            state.font_size_pt,
            state.reference_width_pt,
            state.reference_height_pt,
            state.reference_baseline_pt,
            state.font_size_pt,
        ]
        .iter()
        .map(|value| number_text(*value))
        .collect::<Vec<_>>()
        .join("|"),
    ));
    if let Some(ratio) = state.ink_center_y_ratio.filter(|ratio| (0.0..=1.0).contains(ratio)) {
        variables.set.push((stem(IMAGE_INK_CENTER_PREFIX, state.formula_id), number_text(ratio)));
    }
    if KNOWN_LETTER_FONTS.contains(&state.letter_font) {
        variables.set.push((stem(IMAGE_LETTER_FONT_PREFIX, state.formula_id), state.letter_font.to_string()));
    }
    Ok(variables)
}

#[cfg(test)]
mod tests {
    use super::*;

    const ID: &str = "5028ce4b-71c7-4d06-a165-61151fd64183";

    fn state<'a>(latex: &'a str, metadata: &'a str) -> WordImageState<'a> {
        WordImageState {
            formula_id: ID,
            display_mode: "inline",
            numbered: false,
            latex_base64: latex,
            omml_base64: "PG06b01hdGg-",
            metadata,
            font_size_pt: 12.0,
            reference_width_pt: 70.5,
            reference_height_pt: 18.0,
            reference_baseline_pt: -4.25,
            ink_center_y_ratio: Some(0.5),
            letter_font: "times",
        }
    }

    #[test]
    fn names_match_the_word_for_mac_add_in() {
        assert_eq!(
            formula_reference(ID, "block", true),
            "visualtex:formula-ref:v1:5028ce4b-71c7-4d06-a165-61151fd64183:block:1"
        );
        assert_eq!(owner_bookmark_name(ID), "VT_I_5028ce4b71c74d06a16561151fd64183");
        assert!(owner_bookmark_name(ID).len() <= 40);
        let variables = word_image_variables(&state("TGFUZVg", "visualtex:v1:deflate:abc")).unwrap();
        let names: Vec<&str> = variables.set.iter().map(|(name, _)| name.as_str()).collect();
        assert_eq!(
            names,
            [
                "VT_Latex_5028ce4b_71c7_4d06_a165_61151fd64183_001",
                "VT_Latex_5028ce4b_71c7_4d06_a165_61151fd64183_Count",
                "VT_OMML_5028ce4b_71c7_4d06_a165_61151fd64183_001",
                "VT_OMML_5028ce4b_71c7_4d06_a165_61151fd64183_Count",
                "VT_Metadata_5028ce4b_71c7_4d06_a165_61151fd64183_001",
                "VT_Metadata_5028ce4b_71c7_4d06_a165_61151fd64183_Count",
                "VT_Format_5028ce4b_71c7_4d06_a165_61151fd64183",
                "VT_ImageScale_5028ce4b_71c7_4d06_a165_61151fd64183",
                "VT_ImageInkCenter_5028ce4b_71c7_4d06_a165_61151fd64183",
                "VT_ImageLetterFont_5028ce4b_71c7_4d06_a165_61151fd64183",
            ]
        );
        let value = |name: &str| variables.set.iter().find(|(key, _)| key.starts_with(name)).unwrap().1.clone();
        assert_eq!(value("VT_Format_"), "inline|0");
        assert_eq!(value("VT_ImageScale_"), "12|70.5|18|-4.25|12");
        assert_eq!(value("VT_ImageInkCenter_"), "0.5");
        assert_eq!(value("VT_ImageLetterFont_"), "times");
    }

    #[test]
    fn long_payloads_are_split_into_twenty_thousand_character_chunks() {
        let latex = "A".repeat(45_001);
        let variables = word_image_variables(&state(&latex, "visualtex:v1:deflate:abc")).unwrap();
        let latex_chunks: Vec<&(String, String)> = variables
            .set
            .iter()
            .filter(|(name, _)| name.starts_with("VT_Latex_") && !name.ends_with("_Count"))
            .collect();
        assert_eq!(latex_chunks.len(), 3);
        assert_eq!(latex_chunks[0].1.len(), 20_000);
        assert_eq!(latex_chunks[2].1.len(), 5_001);
        assert!(latex_chunks[2].0.ends_with("_003"));
        assert_eq!(variables.chunked[0].count, 3);
    }

    #[test]
    fn invalid_state_is_rejected() {
        let mut numbered_inline = state("TGFUZVg", "visualtex:v1:deflate:abc");
        numbered_inline.numbered = true;
        assert!(word_image_variables(&numbered_inline).is_err());
        assert!(word_image_variables(&state("not base64!", "visualtex:v1:deflate:abc")).is_err());
        let mut unknown_font = state("TGFUZVg", "visualtex:v1:deflate:abc");
        unknown_font.letter_font = "comic";
        let variables = word_image_variables(&unknown_font).unwrap();
        assert!(!variables.set.iter().any(|(name, _)| name.starts_with("VT_ImageLetterFont_")));
    }
}
