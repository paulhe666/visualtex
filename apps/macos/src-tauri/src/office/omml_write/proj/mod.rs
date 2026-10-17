//! MathML -> Word OMML writer, ported byte-exact from the Windows add-in's
//! WordOmmlConverter.TransformMathMlToOmml (everything after the MathML
//! preparation passes s01..s12) and WordOmmlConverter.BuildWordNativeNumberedOmml:
//!
//! ```text
//! projection = MathMlOmmlProjection.Project(m)          -> project
//! o00 ExtractSingleOMath(projection)                     -> extract_single_omath
//! o01 RestoreExtendedIntegralCharacters(o, naryChars)    -> restore_extended_integral_characters
//! o02 NormalizeExplicitUprightRuns(o, m)                 -> normalize_explicit_upright_runs
//! o03 NormalizeAppliedFunctionStructures(o, m)           -> normalize_applied_function_structures
//! o04 NormalizeExplicitTableColumnAlignment(o, m)        -> normalize_explicit_table_column_alignment
//! o05 NormalizeOmmlPlaceholderVisibility(o)              -> normalize_omml_placeholder_visibility
//! o06 NormalizeOptionalPrescriptSlots(o)                 -> normalize_optional_prescript_slots
//! o07 NormalizeDisplayNaryOmml(o, display)               -> normalize_display_nary_omml
//! o08 NormalizeGeneratedOmmlArgumentOrder(o)             -> normalize_generated_omml_argument_order
//! ```
//!
//! `m` is the prepared MathML (the s12 output), `naryChars` the n-ary list of
//! ReplaceExtendedIntegralsWithOfficePlaceholders (run on the s11 output) and
//! `display` is IsBlockMathMl(m).
//!
//! Byte-exactness rests on three emulations: `xml_reader` builds the trees
//! System.Xml.Linq builds (namespace declarations as attributes, `<a></a>` as
//! empty-string content), `xml_writer` reproduces XLinq's ElementWriter on
//! .NET's XmlWellFormedWriter (prefix lookup, generated `p{N}` prefixes,
//! default-namespace declarations, escaping, "\r\n" new lines) and `netstr`
//! gives .NET 8's UTF-16 string semantics and Unicode 15.0 data.
//!
//! Errors are the C# exceptions as text ("InvalidDataException: ...").
//!
//! Known representation limits (none occurs in the Windows reference corpus):
//! XML comments and processing instructions are dropped when read (XLinq keeps
//! them unless the reader ignores them), and CDATA sections become text.

// The shared arena DOM (also used by the other half of the port); not every
// helper is needed here, and its reference-port Doc::parse uses a deprecated
// quick-xml call that this crate never reaches (it reads with xml_reader).
#[allow(dead_code, deprecated)]
pub mod dom;
mod converter;
mod netstr;
mod projection;
mod unicode_tables;
mod xml_reader;
mod xml_writer;

/// MathMlOmmlProjection.Project(preparedMathMl).
pub fn project(prepared_mathml: &str) -> Result<String, String> {
    projection::project(prepared_mathml)
}

/// o00: WordOmmlConverter.ExtractSingleOMath(xml).
pub fn extract_single_omath(xml: &str) -> Result<String, String> {
    converter::extract_single_omath(xml)
}

/// o01: WordOmmlConverter.RestoreExtendedIntegralCharacters(omml, naryCharacters).
pub fn restore_extended_integral_characters(omml: &str, nary_characters: &[String]) -> Result<String, String> {
    converter::restore_extended_integral_characters(omml, nary_characters)
}

/// o02: WordOmmlConverter.NormalizeExplicitUprightRuns(omml, mathMl).
pub fn normalize_explicit_upright_runs(omml: &str, prepared_mathml: &str) -> Result<String, String> {
    converter::normalize_explicit_upright_runs(omml, prepared_mathml)
}

/// o03: WordOmmlConverter.NormalizeAppliedFunctionStructures(omml, mathMl).
pub fn normalize_applied_function_structures(omml: &str, prepared_mathml: &str) -> Result<String, String> {
    converter::normalize_applied_function_structures(omml, prepared_mathml)
}

/// o04: WordOmmlConverter.NormalizeExplicitTableColumnAlignment(omml, mathMl).
pub fn normalize_explicit_table_column_alignment(omml: &str, prepared_mathml: &str) -> Result<String, String> {
    converter::normalize_explicit_table_column_alignment(omml, prepared_mathml)
}

/// o05: WordOmmlConverter.NormalizeOmmlPlaceholderVisibility(omml).
pub fn normalize_omml_placeholder_visibility(omml: &str) -> Result<String, String> {
    converter::normalize_omml_placeholder_visibility(omml)
}

/// o06: WordOmmlConverter.NormalizeOptionalPrescriptSlots(omml).
pub fn normalize_optional_prescript_slots(omml: &str) -> Result<String, String> {
    converter::normalize_optional_prescript_slots(omml)
}

/// o07: WordOmmlConverter.NormalizeDisplayNaryOmml(omml, display).
pub fn normalize_display_nary_omml(omml: &str, display: bool) -> Result<String, String> {
    converter::normalize_display_nary_omml(omml, display)
}

/// o08: WordOmmlConverter.NormalizeGeneratedOmmlArgumentOrder(omml).
pub fn normalize_generated_omml_argument_order(omml: &str) -> Result<String, String> {
    converter::normalize_generated_omml_argument_order(omml)
}

/// o00..o08 over a projection: the rest of TransformMathMlToOmml.
pub fn finish_omml(
    projection: &str,
    prepared_mathml: &str,
    nary_characters: &[String],
    is_block: bool,
) -> Result<String, String> {
    converter::finish_omml(projection, prepared_mathml, nary_characters, is_block)
}

/// WordOmmlConverter.ReplaceExtendedIntegralsWithOfficePlaceholders(mathMl):
/// (MathMl, NaryCharacters). Run it on the s11 output; its MathML is s12.
pub fn replace_extended_integrals_with_office_placeholders(mathml: &str) -> Result<(String, Vec<String>), String> {
    converter::replace_extended_integrals_with_office_placeholders(mathml)
}

/// WordOmmlConverter.IsBlockMathMl(mathMl).
pub fn is_block_mathml(mathml: &str) -> bool {
    converter::is_block_mathml(mathml)
}

/// TransformMathMlToOmml from the s11 output on: placeholders, display,
/// projection and o00..o08.
pub fn transform_prepared_mathml_to_omml(s11_mathml: &str) -> Result<String, String> {
    let (prepared, nary_characters) = replace_extended_integrals_with_office_placeholders(s11_mathml)?;
    let display = is_block_mathml(&prepared);
    let projection = project(&prepared)?;
    finish_omml(&projection, &prepared, &nary_characters, display)
}

/// WordOmmlConverter.BuildWordNativeNumberedOmml(semanticOmml, sequenceName,
/// numberBookmarkName, restartHeadingLevel, headingSeparator,
/// initialSequenceResult, literalHeadingPrefix, restartSequenceValue).
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
    converter::build_word_native_numbered_omml(
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

/// WordOmmlConverter.ApplyExplicitTransferMathFont(omml, mathFontName).
pub fn apply_explicit_transfer_math_font(omml: &str, math_font_name: &str) -> Result<String, String> {
    converter::apply_explicit_transfer_math_font(omml, math_font_name)
}

#[cfg(test)]
mod behaviour_tests;

/// WordOmmlConverter.NormalizeMathFontName.
pub fn normalize_math_font_name(math_font_name: &str) -> String {
    converter::normalize_math_font_name(math_font_name)
}
