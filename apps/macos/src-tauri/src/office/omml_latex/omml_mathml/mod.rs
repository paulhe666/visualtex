//! Word OMML -> MathML: OmmlMathMlProjection.Project and
//! WordOmmlConverter.TransformOmmlToMathMl from the Windows add-in.
#![allow(dead_code)]

use super::dom;

mod contracts;
mod projection;
mod transform;
mod xml;

pub(super) fn omml_to_mathml(omml: &str, display: bool) -> Result<String, String> {
    transform::omml_to_mathml(omml, display)
}
