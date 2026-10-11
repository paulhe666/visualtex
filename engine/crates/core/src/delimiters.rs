//! Recognized TeX delimiter spellings. The '.' delimiter has no glyph/advance.
//! This is language data, not a dependency on a third-party editing engine.
pub fn glyph(source: &str) -> Option<Option<char>> {
    Some(match source {
        "." => None,
        "(" | "\\lparen" => Some('('), ")" | "\\rparen" => Some(')'),
        "[" | "\\lbrack" => Some('['), "]" | "\\rbrack" => Some(']'),
        "\\{" | "\\lbrace" => Some('{'), "\\}" | "\\rbrace" => Some('}'),
        "|" | "\\vert" | "\\lvert" | "\\rvert" => Some('|'),
        "\\|" | "\\Vert" | "\\lVert" | "\\rVert" | "‖" => Some('‖'),
        "<" | "\\langle" | "⟨" => Some('⟨'), ">" | "\\rangle" | "⟩" => Some('⟩'),
        "\\lfloor" | "⌊" => Some('⌊'), "\\rfloor" | "⌋" => Some('⌋'),
        "\\lceil" | "⌈" => Some('⌈'), "\\rceil" | "⌉" => Some('⌉'),
        "/" => Some('/'), "\\backslash" => Some('\\'),
        _ => return None,
    })
}
/// Visual typing uses these literal characters; source editing stays literal TeX.
pub fn matching_pair(open: &str) -> Option<(&'static str, &'static str)> {
    match open { "(" => Some(("(", ")")), "[" => Some(("[", "]")),
        "{" => Some(("\\{", "\\}")), _ => None }
}
pub fn literal_close_source(close: &str) -> &str {
    match close { "}" => "\\}", _ => close }
}

/// Fixed-size delimiter commands (\big( ... \Biggr]). The size is the
/// delimiter's total height in em, matching the common TeX/KaTeX ladder.
pub fn big_command(command: &str) -> Option<(f64, crate::model::AtomClass)> {
    use crate::model::AtomClass;
    let (base, suffix) = ["\\Bigg", "\\bigg", "\\Big", "\\big"].iter()
        .find_map(|base| command.strip_prefix(base).map(|rest| (*base, rest)))?;
    let size = match base { "\\big" => 1.2, "\\Big" => 1.8, "\\bigg" => 2.4, _ => 3.0 };
    let class = match suffix { "" => AtomClass::Ordinary, "l" => AtomClass::Open, "r" => AtomClass::Close, "m" => AtomClass::Relation, _ => return None };
    Some((size, class))
}

/// The control word that starts a stored fixed-size delimiter spelling.
pub fn big_spelling(spelling: &str) -> Option<(f64, crate::model::AtomClass)> {
    let rest = spelling.strip_prefix('\\')?;
    let end = rest.find(|c: char| !c.is_ascii_alphabetic()).unwrap_or(rest.len());
    big_command(&spelling[..end + 1])
}

/// Precomposed negations for \not followed by a relation.
pub fn negated_relation(value: &str) -> Option<&'static str> {
    Some(match value {
        "=" => "≠", "<" => "≮", ">" => "≯", "∈" => "∉", "∋" => "∌", "≡" => "≢", "∼" => "≁",
        "≤" => "≰", "≥" => "≱", "⊂" => "⊄", "⊃" => "⊅", "⊆" => "⊈", "⊇" => "⊉", "≈" => "≉",
        "≅" => "≇", "≃" => "≄", "∥" => "∦", "∣" => "∤", "≺" => "⊀", "≻" => "⊁", "→" => "↛",
        "←" => "↚", "↔" => "↮", "⇒" => "⇏", "⇐" => "⇍", "⇔" => "⇎", "⊢" => "⊬", "⊨" => "⊭",
        "≍" => "≭", "⊑" => "⋢", "⊒" => "⋣", "⊲" => "⋪", "⊳" => "⋫", "⊴" => "⋬", "⊵" => "⋭",
        _ => return None,
    })
}
