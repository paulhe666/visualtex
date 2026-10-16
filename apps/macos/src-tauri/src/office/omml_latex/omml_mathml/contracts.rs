//! Small shared VstoShared helpers reached by TransformOmmlToMathMl:
//! MathMlColorContract.Normalize, MathMlTokenStyleContract.ResolveEffectiveVariant,
//! MathSpaceContract.FindOmmlSpaceText, and the .NET string primitives they use.

use super::dom::{Doc, NodeId, MATHML_NS};
use regex::Regex;
use std::sync::LazyLock;

// ---- .NET string primitives ------------------------------------------------

/// string.IsNullOrWhiteSpace (char.IsWhiteSpace == Unicode White_Space).
pub fn is_null_or_white_space(value: &str) -> bool {
    value.chars().all(char::is_whitespace)
}

/// String.Trim() (char.IsWhiteSpace).
pub fn dotnet_trim(value: &str) -> &str {
    value.trim_matches(char::is_whitespace)
}

/// ToLowerInvariant: simple (one-to-one) case mapping per character.
pub fn to_lower_invariant(value: &str) -> String {
    value
        .chars()
        .map(|c| {
            if c == '\u{130}' {
                return 'i';
            }
            let mut lower = c.to_lowercase();
            match (lower.next(), lower.next()) {
                (Some(single), None) => single,
                _ => c,
            }
        })
        .collect()
}

/// Ordinal ignore-case equality for the ASCII literals this code compares with.
pub fn eq_ignore_case(value: &str, literal: &str) -> bool {
    value.eq_ignore_ascii_case(literal)
}

/// IndexOf(literal, StringComparison.OrdinalIgnoreCase) >= 0 for ASCII literals.
pub fn contains_ignore_case(value: &str, literal: &str) -> bool {
    value.to_ascii_lowercase().contains(&literal.to_ascii_lowercase())
}

/// int.TryParse(string) with NumberStyles.Integer and invariant signs.
pub fn try_parse_int(value: Option<&str>) -> Option<i32> {
    let value = value?;
    let trimmed = value.trim_matches(|c: char| matches!(c, '\t' | '\n' | '\u{B}' | '\u{C}' | '\r' | ' '));
    let (negative, digits) = match trimmed.as_bytes().first() {
        Some(b'-') => (true, &trimmed[1..]),
        Some(b'+') => (false, &trimmed[1..]),
        _ => (false, trimmed),
    };
    if digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let mut result: i64 = 0;
    for b in digits.bytes() {
        result = result * 10 + i64::from(b - b'0');
        if result > i64::from(i32::MAX) + 1 {
            return None;
        }
    }
    let result = if negative { -result } else { result };
    i32::try_from(result).ok()
}

/// char.IsDigit on one UTF-16 code unit (astral characters are surrogates).
pub fn is_dotnet_digit(c: char) -> bool {
    static DIGIT: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"\A\p{Nd}\z").unwrap());
    if c.is_ascii_digit() {
        return true;
    }
    if c.len_utf16() != 1 {
        return false;
    }
    let mut buffer = [0u8; 4];
    DIGIT.is_match(c.encode_utf8(&mut buffer))
}

/// CharUnicodeInfo.GetUnicodeCategory(...) == UnicodeCategory.MathSymbol.
pub fn is_math_symbol(c: char) -> bool {
    static MATH: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"\A\p{Sm}\z").unwrap());
    let mut buffer = [0u8; 4];
    MATH.is_match(c.encode_utf8(&mut buffer))
}

// ---- MathMlColorContract ---------------------------------------------------

/// MathMlColorContract.Normalize. Callers only pass "#" + value, which
/// System.Drawing.Color.FromName never knows, so only hex forms survive.
pub fn normalize_color(value: Option<&str>) -> Option<String> {
    static SIX: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"^#[0-9a-fA-F]{6}\n?$").unwrap());
    static THREE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"^#[0-9a-fA-F]{3}\n?$").unwrap());
    let color = dotnet_trim(value?);
    if color.is_empty() {
        return None;
    }
    if SIX.is_match(color) {
        return Some(color.to_ascii_uppercase());
    }
    if THREE.is_match(color) {
        let expanded: String = color[1..].chars().flat_map(|c| [c, c]).collect();
        return Some(format!("#{}", expanded.to_ascii_uppercase()));
    }
    known_color(color)
}

/// Color.FromName for the non-system known colors (alpha 255). Names are
/// matched case-insensitively, as KnownColor lookup does.
fn known_color(name: &str) -> Option<String> {
    let lower = name.to_ascii_lowercase();
    KNOWN_COLORS
        .iter()
        .find(|(known, _)| known.to_ascii_lowercase() == lower)
        .map(|(_, rgb)| format!("#{rgb:06X}"))
}

// System.Drawing web colors (Transparent has alpha 0 and is excluded).
const KNOWN_COLORS: &[(&str, u32)] = &[
    ("AliceBlue", 0xF0F8FF), ("AntiqueWhite", 0xFAEBD7), ("Aqua", 0x00FFFF), ("Aquamarine", 0x7FFFD4),
    ("Azure", 0xF0FFFF), ("Beige", 0xF5F5DC), ("Bisque", 0xFFE4C4), ("Black", 0x000000),
    ("BlanchedAlmond", 0xFFEBCD), ("Blue", 0x0000FF), ("BlueViolet", 0x8A2BE2), ("Brown", 0xA52A2A),
    ("BurlyWood", 0xDEB887), ("CadetBlue", 0x5F9EA0), ("Chartreuse", 0x7FFF00), ("Chocolate", 0xD2691E),
    ("Coral", 0xFF7F50), ("CornflowerBlue", 0x6495ED), ("Cornsilk", 0xFFF8DC), ("Crimson", 0xDC143C),
    ("Cyan", 0x00FFFF), ("DarkBlue", 0x00008B), ("DarkCyan", 0x008B8B), ("DarkGoldenrod", 0xB8860B),
    ("DarkGray", 0xA9A9A9), ("DarkGreen", 0x006400), ("DarkKhaki", 0xBDB76B), ("DarkMagenta", 0x8B008B),
    ("DarkOliveGreen", 0x556B2F), ("DarkOrange", 0xFF8C00), ("DarkOrchid", 0x9932CC), ("DarkRed", 0x8B0000),
    ("DarkSalmon", 0xE9967A), ("DarkSeaGreen", 0x8FBC8B), ("DarkSlateBlue", 0x483D8B),
    ("DarkSlateGray", 0x2F4F4F), ("DarkTurquoise", 0x00CED1), ("DarkViolet", 0x9400D3),
    ("DeepPink", 0xFF1493), ("DeepSkyBlue", 0x00BFFF), ("DimGray", 0x696969), ("DodgerBlue", 0x1E90FF),
    ("Firebrick", 0xB22222), ("FloralWhite", 0xFFFAF0), ("ForestGreen", 0x228B22), ("Fuchsia", 0xFF00FF),
    ("Gainsboro", 0xDCDCDC), ("GhostWhite", 0xF8F8FF), ("Gold", 0xFFD700), ("Goldenrod", 0xDAA520),
    ("Gray", 0x808080), ("Green", 0x008000), ("GreenYellow", 0xADFF2F), ("Honeydew", 0xF0FFF0),
    ("HotPink", 0xFF69B4), ("IndianRed", 0xCD5C5C), ("Indigo", 0x4B0082), ("Ivory", 0xFFFFF0),
    ("Khaki", 0xF0E68C), ("Lavender", 0xE6E6FA), ("LavenderBlush", 0xFFF0F5), ("LawnGreen", 0x7CFC00),
    ("LemonChiffon", 0xFFFACD), ("LightBlue", 0xADD8E6), ("LightCoral", 0xF08080), ("LightCyan", 0xE0FFFF),
    ("LightGoldenrodYellow", 0xFAFAD2), ("LightGray", 0xD3D3D3), ("LightGreen", 0x90EE90),
    ("LightPink", 0xFFB6C1), ("LightSalmon", 0xFFA07A), ("LightSeaGreen", 0x20B2AA),
    ("LightSkyBlue", 0x87CEFA), ("LightSlateGray", 0x778899), ("LightSteelBlue", 0xB0C4DE),
    ("LightYellow", 0xFFFFE0), ("Lime", 0x00FF00), ("LimeGreen", 0x32CD32), ("Linen", 0xFAF0E6),
    ("Magenta", 0xFF00FF), ("Maroon", 0x800000), ("MediumAquamarine", 0x66CDAA), ("MediumBlue", 0x0000CD),
    ("MediumOrchid", 0xBA55D3), ("MediumPurple", 0x9370DB), ("MediumSeaGreen", 0x3CB371),
    ("MediumSlateBlue", 0x7B68EE), ("MediumSpringGreen", 0x00FA9A), ("MediumTurquoise", 0x48D1CC),
    ("MediumVioletRed", 0xC71585), ("MidnightBlue", 0x191970), ("MintCream", 0xF5FFFA),
    ("MistyRose", 0xFFE4E1), ("Moccasin", 0xFFE4B5), ("NavajoWhite", 0xFFDEAD), ("Navy", 0x000080),
    ("OldLace", 0xFDF5E6), ("Olive", 0x808000), ("OliveDrab", 0x6B8E23), ("Orange", 0xFFA500),
    ("OrangeRed", 0xFF4500), ("Orchid", 0xDA70D6), ("PaleGoldenrod", 0xEEE8AA), ("PaleGreen", 0x98FB98),
    ("PaleTurquoise", 0xAFEEEE), ("PaleVioletRed", 0xDB7093), ("PapayaWhip", 0xFFEFD5),
    ("PeachPuff", 0xFFDAB9), ("Peru", 0xCD853F), ("Pink", 0xFFC0CB), ("Plum", 0xDDA0DD),
    ("PowderBlue", 0xB0E0E6), ("Purple", 0x800080), ("RebeccaPurple", 0x663399), ("Red", 0xFF0000),
    ("RosyBrown", 0xBC8F8F), ("RoyalBlue", 0x4169E1), ("SaddleBrown", 0x8B4513), ("Salmon", 0xFA8072),
    ("SandyBrown", 0xF4A460), ("SeaGreen", 0x2E8B57), ("SeaShell", 0xFFF5EE), ("Sienna", 0xA0522D),
    ("Silver", 0xC0C0C0), ("SkyBlue", 0x87CEEB), ("SlateBlue", 0x6A5ACD), ("SlateGray", 0x708090),
    ("Snow", 0xFFFAFA), ("SpringGreen", 0x00FF7F), ("SteelBlue", 0x4682B4), ("Tan", 0xD2B48C),
    ("Teal", 0x008080), ("Thistle", 0xD8BFD8), ("Tomato", 0xFF6347), ("Turquoise", 0x40E0D0),
    ("Violet", 0xEE82EE), ("Wheat", 0xF5DEB3), ("White", 0xFFFFFF), ("WhiteSmoke", 0xF5F5F5),
    ("Yellow", 0xFFFF00), ("YellowGreen", 0x9ACD32),
];

// ---- MathMlTokenStyleContract ----------------------------------------------

fn is_mathml_element(doc: &Doc, element: NodeId) -> bool {
    doc.name(element).is_some_and(|name| name.ns.is_empty() || name.ns == MATHML_NS)
}

pub fn resolve_declared_variant(doc: &Doc, element: NodeId) -> Option<String> {
    let mut chain = vec![element];
    chain.extend(doc.ancestors(element));
    for current in chain {
        if !is_mathml_element(doc, current) {
            break;
        }
        if let Some(declared) = doc.attr(current, "mathvariant") {
            let declared = dotnet_trim(declared);
            if !declared.is_empty() {
                return Some(to_lower_invariant(declared));
            }
        }
        if doc.local(current) == "math" {
            break;
        }
    }
    None
}

pub fn resolve_effective_variant(doc: &Doc, token: NodeId) -> Option<String> {
    let declared = resolve_declared_variant(doc, token);
    if !is_mathml_element(doc, token) || !matches!(doc.local(token), "mi" | "mo" | "mn" | "mtext" | "ms") {
        return None;
    }
    if let Some(declared) = declared.as_deref() {
        if !is_null_or_white_space(declared) {
            return Some(to_lower_invariant(dotnet_trim(declared)));
        }
    }
    if !doc.elements(token).is_empty() {
        return None;
    }
    let value = doc.value(token);
    let trimmed = dotnet_trim(&value);
    let italic = doc.local(token) == "mi"
        && is_single_mathml_character(trimmed)
        && !matches!(trimmed, "′" | "″" | "‴" | "⁗" | "‵" | "‶" | "‷");
    Some(if italic { "italic" } else { "normal" }.to_string())
}

pub fn is_single_mathml_character(text: &str) -> bool {
    static SINGLE: LazyLock<Regex> = LazyLock::new(|| {
        Regex::new(
            r"\A[^\r\n\u{2028}\u{2029}][\u{300}-\u{36F}\u{1AB0}-\u{1ABE}\u{1DC0}-\u{1DFF}\u{20D0}-\u{20EF}]*\z",
        )
        .unwrap()
    });
    SINGLE.is_match(text)
}

// ---- MathSpaceContract -----------------------------------------------------

pub struct OmmlSpace {
    pub mathml_width: &'static str,
    pub text: &'static str,
}

pub const OMML_SPACES: &[OmmlSpace] = &[
    OmmlSpace { mathml_width: "1em", text: "\u{2003}" },
    OmmlSpace { mathml_width: "2em", text: "\u{2003}\u{2003}" },
];

pub fn maximum_omml_space_text_length() -> usize {
    OMML_SPACES.iter().map(|space| space.text.encode_utf16().count()).max().unwrap_or(0)
}

pub fn find_omml_space_text(text: &str) -> Option<&'static OmmlSpace> {
    OMML_SPACES.iter().find(|space| space.text == text)
}
