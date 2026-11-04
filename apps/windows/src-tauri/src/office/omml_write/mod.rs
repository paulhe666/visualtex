//! Word equations (MathML -> OMML): the shared Word writer, ported from VisualTeX for Windows' C# Word writer
//! (WordOmmlConverter) to Rust on the Mac, used here by the document export (document_export.rs) so both apps
//! write identical Word equations (docs/equation-format-v2.md).
//! - `prep`: WordOmmlConverter's MathML passes before the projection;
//! - `proj`: MathMlOmmlProjection.Project, the OMML passes after it and BuildWordNativeNumberedOmml.

#[allow(dead_code, deprecated, unused_imports)]
mod prep;
#[allow(dead_code, deprecated, unused_imports)]
mod proj;

/// WordOmmlConverter.TransformMathMlToOmml: Presentation MathML to one m:oMath.
pub(crate) fn transform_mathml_to_omml(mathml: &str) -> Result<String, String> {
    let prepared = prep::prepare_mathml_for_omml(mathml)?;
    let projection = proj::project(&prepared.mathml)?;
    proj::finish_omml(&projection, &prepared.mathml, &prepared.nary_characters, prepared.is_block)
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

/// ApplyExplicitTransferMathFont(ExtractSingleOMath(omml), math font): what Word imports for an equation.
pub(crate) fn insertion_omml(omml: &str, math_font_name: &str) -> Result<String, String> {
    proj::apply_explicit_transfer_math_font(&proj::extract_single_omath(omml)?, math_font_name)
}
