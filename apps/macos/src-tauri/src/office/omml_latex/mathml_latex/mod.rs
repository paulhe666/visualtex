//! MathML -> LaTeX: MathMlToLatexConverter.Convert from the Windows add-in,
//! with the token-style, color, accent-geometry and space contracts it uses.
#![allow(dead_code)]

use super::dom;

mod contracts;
mod converter;
mod formula_xml;
mod netstr;
mod tables;
mod unicode_tables;

#[cfg(test)]
mod unit_tests;

/// Every case the C# converter rejects with an exception is an `Err`.
/// One character as the converter writes it (\alpha for α).
pub(super) fn token_latex(value: &str) -> String {
    converter::convert_token(value)
}

pub(super) fn mathml_to_latex(mathml: &str) -> Result<String, String> {
    converter::convert(mathml)
}
