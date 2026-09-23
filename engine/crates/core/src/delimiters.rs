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
