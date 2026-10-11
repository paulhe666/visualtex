//! Bounded two-argument fraction syntax. Every form keeps the same editable
//! numerator/denominator roles; presentation is derived from an authored command.
//! No arbitrary TeX expansion or old renderer is used by this registry.
use crate::layout::MathStyle;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FractionForm { Barred, Binomial, Bevelled }
#[derive(Debug, Clone, Copy)]
pub struct FractionSpec {
    pub form: FractionForm,
    pub forced_style: Option<MathStyle>,
    /// \cfrac keeps both arguments in display style (continued fractions).
    pub display_arguments: bool,
}
pub fn spec(command: &str) -> Option<FractionSpec> {
    use FractionForm::*;
    let (form, forced_style) = match command {
        "\\frac" => (Barred, None),
        "\\dfrac" => (Barred, Some(MathStyle::Display)),
        "\\tfrac" => (Barred, Some(MathStyle::Text)),
        "\\binom" => (Binomial, None),
        "\\dbinom" => (Binomial, Some(MathStyle::Display)),
        "\\tbinom" => (Binomial, Some(MathStyle::Text)),
        "\\nicefrac" => (Bevelled, None),
        "\\cfrac" => (Barred, Some(MathStyle::Display)),
        _ => return None,
    };
    Some(FractionSpec { form, forced_style, display_arguments: command == "\\cfrac" })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn only_bounded_fraction_spellings_have_presentation_rules() {
        assert_eq!(spec("\\binom").unwrap().form, FractionForm::Binomial);
        assert_eq!(spec("\\dbinom").unwrap().forced_style, Some(MathStyle::Display));
        assert_eq!(spec("\\tbinom").unwrap().forced_style, Some(MathStyle::Text));
        assert_eq!(spec("\\nicefrac").unwrap().form, FractionForm::Bevelled);
        for command in ["\\genfrac", "\\over", "\\atop", "\\choose", "\\futurefrac"] {
            assert!(spec(command).is_none(), "unimplemented syntax cannot acquire a rule");
        }
    }
}
