//! The MathML pre-passes of the Windows add-in's
//! `WordOmmlConverter.TransformMathMlToOmml(string mathMl)`, ported so that
//! VisualTeX for Mac prepares MathML for the OMML projection byte for byte
//! like Windows (.NET 8):
//!
//! ```text
//! ValidateMathMlForOmml(mathMl)
//! s01 NormalizeMathTypeBinomialPiles         -> normalize_math_type_binomial_piles
//! s02 NormalizeMathJaxBinomialDelimiters     -> normalize_mathjax_binomial_delimiters
//! s03 NormalizePlainBracketPairs             -> normalize_plain_bracket_pairs
//! s04 NormalizeFencedMathMlTables            -> normalize_fenced_mathml_tables
//! s05 NormalizeNestedEmptyBaseScripts        -> normalize_nested_empty_base_scripts
//! s06 NormalizeMathMlPrescripts              -> normalize_mathml_prescripts
//! s07 NormalizeMathMlAccents                 -> normalize_mathml_accents
//! s08 NormalizeNaryArguments                 -> normalize_nary_arguments
//! s09 NormalizeMathMlTokenStyleAuthority     -> normalize_mathml_token_style_authority
//! s10 MathMlColorContract.MaterializeInheritedColors -> materialize_inherited_colors
//! s11 ProjectMathMlSpacesForOmml(m).MathMl   -> project_mathml_spaces_for_omml
//! s12 ReplaceExtendedIntegralsWithOfficePlaceholders -> replace_extended_integrals_with_office_placeholders
//! IsBlockMathMl                              -> is_block_mathml
//! ```
//!
//! The projection (MathMlOmmlProjection.Project) and the OMML post-passes are
//! ported separately and consume [`PreparedMathMl`].

#[allow(dead_code)]
pub mod dom;
pub mod contracts;
pub mod netstr;
mod prepasses;
mod unicode_letters;
pub mod xlinq;

pub use contracts::MathSpaceOmmlCapability;
pub use prepasses::{
    is_block_mathml, materialize_inherited_colors, normalize_fenced_mathml_tables, normalize_math_type_binomial_piles,
    normalize_mathjax_binomial_delimiters, normalize_mathml_accents, normalize_mathml_prescripts,
    normalize_mathml_token_style_authority, normalize_nary_arguments, normalize_nested_empty_base_scripts,
    normalize_plain_bracket_pairs, project_mathml_spaces_for_omml, replace_extended_integrals_with_office_placeholders,
    validate_mathml_for_omml, EXTENDED_INTEGRAL_CHARACTERS, NARY_CHARACTERS, OMML_RUN_ROLE_ATTRIBUTE,
    OMML_SPACE_TEXT_ATTRIBUTE,
};

/// The MathML that TransformMathMlToOmml hands to MathMlOmmlProjection.Project,
/// with what the post-passes need from the pre-passes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PreparedMathMl {
    /// The s12 output.
    pub mathml: String,
    /// ReplaceExtendedIntegralsWithOfficePlaceholders(...).NaryCharacters, for
    /// RestoreExtendedIntegralCharacters.
    pub nary_characters: Vec<String>,
    /// IsBlockMathMl of the s12 output (NormalizeDisplayNaryOmml's `display`).
    pub is_block: bool,
}

/// The pre-pass part of TransformMathMlToOmml: the empty-input check,
/// ValidateMathMlForOmml, s01..s12 and IsBlockMathMl.
pub fn prepare_mathml_for_omml(mathml: &str) -> Result<PreparedMathMl, String> {
    if netstr::is_white_space(mathml) {
        return Err("VisualTeX did not provide MathML for the Word OMML formula.".to_string());
    }
    validate_mathml_for_omml(mathml)?;
    let mathml = normalize_math_type_binomial_piles(mathml)?;
    let mathml = normalize_mathjax_binomial_delimiters(&mathml)?;
    let mathml = normalize_plain_bracket_pairs(&mathml)?;
    let mathml = normalize_fenced_mathml_tables(&mathml)?;
    let mathml = normalize_nested_empty_base_scripts(&mathml)?;
    let mathml = normalize_mathml_prescripts(&mathml)?;
    let mathml = normalize_mathml_accents(&mathml)?;
    let mathml = normalize_nary_arguments(&mathml)?;
    let mathml = normalize_mathml_token_style_authority(&mathml)?;
    let mathml = materialize_inherited_colors(&mathml)?;
    let (mathml, _capabilities) = project_mathml_spaces_for_omml(&mathml)?;
    let (mathml, nary_characters) = replace_extended_integrals_with_office_placeholders(&mathml)?;
    let is_block = is_block_mathml(&mathml);
    Ok(PreparedMathMl { mathml, nary_characters, is_block })
}
