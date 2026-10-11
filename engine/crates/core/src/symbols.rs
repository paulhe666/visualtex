use crate::model::AtomClass;

/// Deliberately bounded K0 symbol vocabulary. This is not a copied math-engine
/// command registry; unsupported structural commands remain opaque source.
pub fn named_symbol(command: &str) -> Option<(&'static str, AtomClass)> {
    use AtomClass::*;
    Some(match command {
        "\\alpha" => ("α", Ordinary), "\\beta" => ("β", Ordinary),
        "\\gamma" => ("γ", Ordinary), "\\delta" => ("δ", Ordinary),
        "\\epsilon" => ("ϵ", Ordinary), "\\varepsilon" => ("ε", Ordinary),
        "\\zeta" => ("ζ", Ordinary), "\\eta" => ("η", Ordinary),
        "\\theta" => ("θ", Ordinary), "\\vartheta" => ("ϑ", Ordinary),
        "\\iota" => ("ι", Ordinary), "\\kappa" => ("κ", Ordinary),
        "\\lambda" => ("λ", Ordinary), "\\mu" => ("μ", Ordinary),
        "\\nu" => ("ν", Ordinary), "\\xi" => ("ξ", Ordinary),
        "\\pi" => ("π", Ordinary), "\\varpi" => ("ϖ", Ordinary),
        "\\rho" => ("ρ", Ordinary), "\\varrho" => ("ϱ", Ordinary),
        "\\sigma" => ("σ", Ordinary), "\\varsigma" => ("ς", Ordinary),
        "\\tau" => ("τ", Ordinary), "\\upsilon" => ("υ", Ordinary),
        "\\phi" => ("ϕ", Ordinary), "\\varphi" => ("φ", Ordinary),
        "\\chi" => ("χ", Ordinary), "\\psi" => ("ψ", Ordinary), "\\omega" => ("ω", Ordinary),
        "\\Gamma" => ("Γ", Ordinary), "\\Delta" => ("Δ", Ordinary),
        "\\Theta" => ("Θ", Ordinary), "\\Lambda" => ("Λ", Ordinary),
        "\\Xi" => ("Ξ", Ordinary), "\\Pi" => ("Π", Ordinary),
        "\\Sigma" => ("Σ", Ordinary), "\\Upsilon" => ("Υ", Ordinary),
        "\\Phi" => ("Φ", Ordinary), "\\Psi" => ("Ψ", Ordinary), "\\Omega" => ("Ω", Ordinary),
        "\\infty" => ("∞", Ordinary), "\\partial" => ("∂", Ordinary), "\\nabla" => ("∇", Ordinary),
        "\\sin" => ("sin", Operator), "\\cos" => ("cos", Operator), "\\tan" => ("tan", Operator),
        "\\cot" => ("cot", Operator), "\\sec" => ("sec", Operator), "\\csc" => ("csc", Operator),
        "\\sinh" => ("sinh", Operator), "\\cosh" => ("cosh", Operator), "\\tanh" => ("tanh", Operator), "\\coth" => ("coth", Operator),
        "\\arcsin" => ("arcsin", Operator), "\\arccos" => ("arccos", Operator), "\\arctan" => ("arctan", Operator),
        "\\log" => ("log", Operator), "\\ln" => ("ln", Operator), "\\exp" => ("exp", Operator),
        "\\det" => ("det", Operator), "\\gcd" => ("gcd", Operator), "\\arg" => ("arg", Operator), "\\deg" => ("deg", Operator),
        "\\ker" => ("ker", Operator), "\\dim" => ("dim", Operator), "\\hom" => ("hom", Operator),
        "\\times" => ("×", Binary), "\\cdot" => ("⋅", Binary), "\\div" => ("÷", Binary),
        "\\pm" => ("±", Binary), "\\mp" => ("∓", Binary),
        "\\le" | "\\leq" => ("≤", Relation), "\\ge" | "\\geq" => ("≥", Relation),
        "\\ne" | "\\neq" => ("≠", Relation), "\\approx" => ("≈", Relation),
        "\\equiv" => ("≡", Relation),
        // amsmath \bmod is the binary operator word "mod".
        "\\bmod" => ("mod", Binary), "\\to" | "\\rightarrow" => ("→", Relation),
        "\\leftarrow" => ("←", Relation), "\\in" => ("∈", Relation), "\\notin" => ("∉", Relation),
        "\\{" => ("{", Open), "\\}" => ("}", Close),
        "\\%" => ("%", Ordinary), "\\_" => ("_", Ordinary), "\\#" => ("#", Ordinary),
        "\\&" => ("&", Ordinary), "\\$" => ("$", Ordinary),
        _ => return crate::symbol_catalog::lookup(command),
    })
}

pub fn literal_class(value: &str) -> AtomClass {
    use AtomClass::*;
    match value {
        "+" | "-" | "−" | "*" | "×" | "÷" | "±" | "∓" | "⋅" => Binary,
        "=" | "<" | ">" | "≤" | "≥" | "≠" | "≈" => Relation,
        "(" | "[" => Open,
        ")" | "]" => Close,
        "," | ";" | ":" => Punctuation,
        _ if value.chars().all(|c| c.is_ascii_digit()) => Number,
        _ => crate::symbol_catalog::literal_class(value).unwrap_or(Ordinary),
    }
}
