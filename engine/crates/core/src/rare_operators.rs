//! P07 Unicode-defined integral operators, kept distinct from ESINT-specific
//! outline aliases. Product command spellings are inventoried from
//! apps/windows/src/math/extendedIntegralCompatibility.ts (read-only reference).
//! Font lookup is still authoritative: no substitution with a similar integral.
#[derive(Debug, Clone, Copy)]
pub struct RareOperator { pub command: &'static str, pub glyph: char }
pub const OPERATORS: &[RareOperator] = &[
    RareOperator { command: "\\intclockwise", glyph: '∱' },
    RareOperator { command: "\\varointclockwise", glyph: '∲' },
    RareOperator { command: "\\ointctrclockwise", glyph: '∳' },
    RareOperator { command: "\\varointctrclockwise", glyph: '∳' },
    RareOperator { command: "\\sumint", glyph: '⨋' },
    RareOperator { command: "\\iiiint", glyph: '⨌' },
    RareOperator { command: "\\intbar", glyph: '⨍' },
    RareOperator { command: "\\dashint", glyph: '⨍' },
    RareOperator { command: "\\intBar", glyph: '⨎' },
    RareOperator { command: "\\ddashint", glyph: '⨎' },
    RareOperator { command: "\\fint", glyph: '⨏' },
    RareOperator { command: "\\cirfnint", glyph: '⨐' },
    RareOperator { command: "\\awint", glyph: '⨑' },
    RareOperator { command: "\\intctrclockwise", glyph: '⨑' },
    RareOperator { command: "\\rppolint", glyph: '⨒' },
    RareOperator { command: "\\scpolint", glyph: '⨓' },
    RareOperator { command: "\\npolint", glyph: '⨔' },
    RareOperator { command: "\\pointint", glyph: '⨕' },
    RareOperator { command: "\\quatint", glyph: '⨖' },
    RareOperator { command: "\\intlarhk", glyph: '⨗' },
    RareOperator { command: "\\intx", glyph: '⨘' },
    RareOperator { command: "\\intcap", glyph: '⨙' },
    RareOperator { command: "\\intcup", glyph: '⨚' },
    RareOperator { command: "\\upint", glyph: '⨛' },
    RareOperator { command: "\\lowint", glyph: '⨜' },
];
pub fn glyph(command: &str) -> Option<char> {
    // A literal Unicode integral uses the same structure and mode as its name.
    // Keep original spelling on the node; do not normalize unrelated source.
    if command.starts_with('\\') {
        return OPERATORS.iter().find(|entry| entry.command == command).map(|entry| entry.glyph);
    }
    let mut chars = command.chars();
    let ch = chars.next()?;
    if chars.next().is_some() { return None; }
    OPERATORS.iter().any(|entry| entry.glyph == ch).then_some(ch)
}
