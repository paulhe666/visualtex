//! Mathematical alphabet/style commands. The source spelling remains owned by
//! the Style node; layout maps supported mathematical characters to the Unicode
//! Mathematical Alphanumeric Symbols encoded by the injected math font.
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum MathVariant {
    Default,
    Upright,
    Bold,
    Italic,
    BoldItalic,
    Sans,
    SansBold,
    SansItalic,
    SansBoldItalic,
    Monospace,
    DoubleStruck,
    Script,
    BoldScript,
    Fraktur,
    BoldFraktur,
}

/// Unicode-math spellings already exposed by the product. Keep their authored
/// source; use real alphabet variants instead of expanding old engine macros.
pub const STYLE_ALIASES: &[(&str, MathVariant)] = &[
    ("\\symup", MathVariant::Upright), ("\\symit", MathVariant::Italic),
    ("\\symbf", MathVariant::Bold), ("\\symbfup", MathVariant::Bold),
    ("\\symbfit", MathVariant::BoldItalic), ("\\symbb", MathVariant::DoubleStruck),
    ("\\symcal", MathVariant::Script), ("\\symscr", MathVariant::Script),
    ("\\symbfcal", MathVariant::BoldScript), ("\\symbfscr", MathVariant::BoldScript),
    ("\\symfrak", MathVariant::Fraktur), ("\\symbffrak", MathVariant::BoldFraktur),
    ("\\symsfup", MathVariant::Sans), ("\\symsfit", MathVariant::SansItalic),
    ("\\symbfsfup", MathVariant::SansBold), ("\\symbfsfit", MathVariant::SansBoldItalic),
    ("\\symtt", MathVariant::Monospace),
];

/// Declarations that switch the math style of the rest of their group.
pub fn style_change(command: &str) -> Option<crate::layout::MathStyle> {
    use crate::layout::MathStyle;
    Some(match command {
        "\\displaystyle" => MathStyle::Display,
        "\\textstyle" => MathStyle::Text,
        "\\scriptstyle" => MathStyle::Script,
        "\\scriptscriptstyle" => MathStyle::ScriptScript,
        _ => return None,
    })
}

pub fn command_variant(command: &str) -> Option<MathVariant> {
    Some(match command {
        "\\mathnormal" => MathVariant::Default,
        "\\mathrm" => MathVariant::Upright,
        "\\mathbf" => MathVariant::Bold,
        "\\mathit" => MathVariant::Italic,
        "\\mathbfit" | "\\boldsymbol" | "\\bm" => MathVariant::BoldItalic,
        "\\mathsf" => MathVariant::Sans,
        "\\mathtt" => MathVariant::Monospace,
        "\\mathbb" => MathVariant::DoubleStruck,
        "\\mathcal" | "\\mathscr" => MathVariant::Script,
        "\\mathfrak" => MathVariant::Fraktur,
        _ => return STYLE_ALIASES.iter().find(|(name, _)| *name == command).map(|(_, variant)| *variant),
    })
}

pub fn canonical_command(variant: MathVariant) -> &'static str {
    match variant {
        MathVariant::Default => "\\mathnormal",
        MathVariant::Upright => "\\mathrm",
        MathVariant::Bold => "\\mathbf",
        MathVariant::Italic => "\\mathit",
        MathVariant::BoldItalic => "\\mathbfit",
        MathVariant::Sans => "\\mathsf",
        MathVariant::SansBold => "\\symbfsfup",
        MathVariant::SansItalic => "\\symsfit",
        MathVariant::SansBoldItalic => "\\symbfsfit",
        MathVariant::Monospace => "\\mathtt",
        MathVariant::DoubleStruck => "\\mathbb",
        MathVariant::Script => "\\mathcal",
        MathVariant::BoldScript => "\\symbfcal",
        MathVariant::Fraktur => "\\mathfrak",
        MathVariant::BoldFraktur => "\\symbffrak",
    }
}

/// Map a mathematical atom character into the explicitly requested alphabet.
/// Characters without a defined variant are retained; callers still validate
/// actual glyph availability in the injected font and never synthesize outlines.
pub fn map_character(ch: char, variant: MathVariant) -> char {
    use MathVariant::*;
    if variant == Default { return default_math_character(ch); }
    if variant == Upright { return ch; }

    if let Some(mapped) = map_ascii(ch, variant) { return mapped; }
    if let Some(mapped) = map_greek(ch, variant) { return mapped; }
    ch
}

pub fn default_math_character(ch: char) -> char {
    match ch {
        'A'..='Z' => scalar(0x1D434 + u32::from(ch) - u32::from('A'), ch),
        'h' => 'ℎ',
        'a'..='z' => scalar(0x1D44E + u32::from(ch) - u32::from('a'), ch),
        _ => ch,
    }
}

fn map_ascii(ch: char, variant: MathVariant) -> Option<char> {
    use MathVariant::*;
    let upper = ch.is_ascii_uppercase();
    let lower = ch.is_ascii_lowercase();
    let digit = ch.is_ascii_digit();
    let index = if upper { u32::from(ch) - u32::from('A') }
        else if lower { u32::from(ch) - u32::from('a') }
        else if digit { u32::from(ch) - u32::from('0') }
        else { return None };

    let cp = match variant {
        Bold if upper => 0x1D400 + index,
        Bold if lower => 0x1D41A + index,
        Bold if digit => 0x1D7CE + index,
        Italic if upper => 0x1D434 + index,
        Italic if lower => if ch == 'h' { return Some('ℎ'); } else { 0x1D44E + index },
        BoldItalic if upper => 0x1D468 + index,
        BoldItalic if lower => 0x1D482 + index,
        Sans if upper => 0x1D5A0 + index,
        Sans if lower => 0x1D5BA + index,
        Sans if digit => 0x1D7E2 + index,
        SansBold if upper => 0x1D5D4 + index,
        SansBold if lower => 0x1D5EE + index,
        SansBold if digit => 0x1D7EC + index,
        SansItalic if upper => 0x1D608 + index,
        SansItalic if lower => 0x1D622 + index,
        SansBoldItalic if upper => 0x1D63C + index,
        SansBoldItalic if lower => 0x1D656 + index,
        Monospace if upper => 0x1D670 + index,
        Monospace if lower => 0x1D68A + index,
        Monospace if digit => 0x1D7F6 + index,
        BoldScript if upper => 0x1D4D0 + index,
        BoldScript if lower => 0x1D4EA + index,
        BoldFraktur if upper => 0x1D56C + index,
        BoldFraktur if lower => 0x1D586 + index,
        Script => return script_ascii(ch),
        Fraktur => return fraktur_ascii(ch),
        DoubleStruck => return double_struck_ascii(ch),
        _ => return None,
    };
    char::from_u32(cp)
}

fn script_ascii(ch: char) -> Option<char> {
    const UPPER: [u32; 26] = [
        0x1D49C,0x212C,0x1D49E,0x1D49F,0x2130,0x2131,0x1D4A2,0x210B,0x2110,0x1D4A5,0x1D4A6,0x2112,0x2133,
        0x1D4A9,0x1D4AA,0x1D4AB,0x1D4AC,0x211B,0x1D4AE,0x1D4AF,0x1D4B0,0x1D4B1,0x1D4B2,0x1D4B3,0x1D4B4,0x1D4B5,
    ];
    const LOWER: [u32; 26] = [
        0x1D4B6,0x1D4B7,0x1D4B8,0x1D4B9,0x212F,0x1D4BB,0x210A,0x1D4BD,0x1D4BE,0x1D4BF,0x1D4C0,0x1D4C1,0x1D4C2,
        0x1D4C3,0x2134,0x1D4C5,0x1D4C6,0x1D4C7,0x1D4C8,0x1D4C9,0x1D4CA,0x1D4CB,0x1D4CC,0x1D4CD,0x1D4CE,0x1D4CF,
    ];
    ascii_table(ch, &UPPER, &LOWER, None)
}

fn fraktur_ascii(ch: char) -> Option<char> {
    const UPPER: [u32; 26] = [
        0x1D504,0x1D505,0x212D,0x1D507,0x1D508,0x1D509,0x1D50A,0x210C,0x2111,0x1D50D,0x1D50E,0x1D50F,0x1D510,
        0x1D511,0x1D512,0x1D513,0x1D514,0x211C,0x1D516,0x1D517,0x1D518,0x1D519,0x1D51A,0x1D51B,0x1D51C,0x2128,
    ];
    const LOWER: [u32; 26] = [
        0x1D51E,0x1D51F,0x1D520,0x1D521,0x1D522,0x1D523,0x1D524,0x1D525,0x1D526,0x1D527,0x1D528,0x1D529,0x1D52A,
        0x1D52B,0x1D52C,0x1D52D,0x1D52E,0x1D52F,0x1D530,0x1D531,0x1D532,0x1D533,0x1D534,0x1D535,0x1D536,0x1D537,
    ];
    ascii_table(ch, &UPPER, &LOWER, None)
}

fn double_struck_ascii(ch: char) -> Option<char> {
    const UPPER: [u32; 26] = [
        0x1D538,0x1D539,0x2102,0x1D53B,0x1D53C,0x1D53D,0x1D53E,0x210D,0x1D540,0x1D541,0x1D542,0x1D543,0x1D544,
        0x2115,0x1D546,0x2119,0x211A,0x211D,0x1D54A,0x1D54B,0x1D54C,0x1D54D,0x1D54E,0x1D54F,0x1D550,0x2124,
    ];
    const LOWER: [u32; 26] = [
        0x1D552,0x1D553,0x1D554,0x1D555,0x1D556,0x1D557,0x1D558,0x1D559,0x1D55A,0x1D55B,0x1D55C,0x1D55D,0x1D55E,
        0x1D55F,0x1D560,0x1D561,0x1D562,0x1D563,0x1D564,0x1D565,0x1D566,0x1D567,0x1D568,0x1D569,0x1D56A,0x1D56B,
    ];
    const DIGITS: [u32; 10] = [0x1D7D8,0x1D7D9,0x1D7DA,0x1D7DB,0x1D7DC,0x1D7DD,0x1D7DE,0x1D7DF,0x1D7E0,0x1D7E1];
    ascii_table(ch, &UPPER, &LOWER, Some(&DIGITS))
}

fn ascii_table(ch: char, upper: &[u32;26], lower: &[u32;26], digits: Option<&[u32;10]>) -> Option<char> {
    let cp = if ch.is_ascii_uppercase() { upper[(u32::from(ch)-u32::from('A')) as usize] }
        else if ch.is_ascii_lowercase() { lower[(u32::from(ch)-u32::from('a')) as usize] }
        else if ch.is_ascii_digit() { digits?[(u32::from(ch)-u32::from('0')) as usize] }
        else { return None };
    char::from_u32(cp)
}

fn map_greek(ch: char, variant: MathVariant) -> Option<char> {
    use MathVariant::*;
    let base = match variant { Bold=>0x1D6A8, Italic=>0x1D6E2, BoldItalic=>0x1D71C,
        SansBold=>0x1D756, SansBoldItalic=>0x1D790, _=>return None };
    let offset = match ch {
        'Α'=>0,'Β'=>1,'Γ'=>2,'Δ'=>3,'Ε'=>4,'Ζ'=>5,'Η'=>6,'Θ'=>7,'Ι'=>8,'Κ'=>9,'Λ'=>10,'Μ'=>11,'Ν'=>12,'Ξ'=>13,'Ο'=>14,'Π'=>15,'Ρ'=>16,
        'Σ'=>18,'Τ'=>19,'Υ'=>20,'Φ'=>21,'Χ'=>22,'Ψ'=>23,'Ω'=>24,
        'α'=>26,'β'=>27,'γ'=>28,'δ'=>29,'ε'=>30,'ζ'=>31,'η'=>32,'θ'=>33,'ι'=>34,'κ'=>35,'λ'=>36,'μ'=>37,'ν'=>38,'ξ'=>39,'ο'=>40,'π'=>41,'ρ'=>42,'ς'=>43,'σ'=>44,'τ'=>45,'υ'=>46,'φ'=>47,'χ'=>48,'ψ'=>49,'ω'=>50,
        'ϴ'=>17,'∇'=>25,'∂'=>51,'ϵ'=>52,'ϑ'=>53,'ϰ'=>54,'ϕ'=>55,'ϱ'=>56,'ϖ'=>57,
        _=>return None,
    };
    char::from_u32(base + offset)
}

fn scalar(cp: u32, fallback: char) -> char { char::from_u32(cp).unwrap_or(fallback) }

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn core_alphabets_map_to_unicode_math_characters() {
        assert_eq!(map_character('A', MathVariant::Bold), '𝐀');
        assert_eq!(map_character('x', MathVariant::Italic), '𝑥');
        assert_eq!(map_character('h', MathVariant::Italic), 'ℎ');
        assert_eq!(map_character('R', MathVariant::DoubleStruck), 'ℝ');
        assert_eq!(map_character('C', MathVariant::DoubleStruck), 'ℂ');
        assert_eq!(map_character('A', MathVariant::Script), '𝒜');
        assert_eq!(map_character('H', MathVariant::Fraktur), 'ℌ');
        assert_eq!(map_character('α', MathVariant::BoldItalic), '𝜶');
        assert_eq!(map_character('Γ', MathVariant::Bold), '𝚪');
        assert_eq!(map_character('中', MathVariant::Bold), '中');
    }
    #[test]
    fn source_commands_have_explicit_variants() {
        assert_eq!(command_variant("\\mathbf"), Some(MathVariant::Bold));
        assert_eq!(command_variant("\\mathrm"), Some(MathVariant::Upright));
        assert_eq!(command_variant("\\bm"), Some(MathVariant::BoldItalic));
        assert_eq!(command_variant("\\unknown"), None);
    }
}
