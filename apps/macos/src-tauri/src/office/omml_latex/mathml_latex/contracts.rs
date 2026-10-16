//! Ports of the shared MathML contracts the converter reaches:
//! MathMlTokenStyleContract, MathMlColorContract, MathMlAccentGeometry and
//! MathSpaceContract.ParseWidth.

use std::sync::LazyLock;

use regex::Regex;

use super::dom::{Doc, NodeId, MATHML_NS};
use super::netstr::{eq_ignore_case, is_null_or_white_space, to_lower_invariant, to_upper_invariant};

// ---- MathMlTokenStyleContract -------------------------------------------------

fn is_mathml_element(doc: &Doc, element: NodeId) -> bool {
    doc.name(element).is_some_and(|name| name.ns.is_empty() || name.ns == MATHML_NS)
}

fn ancestors_and_self(doc: &Doc, element: NodeId) -> Vec<NodeId> {
    let mut out = vec![element];
    out.extend(doc.ancestors(element));
    out
}

pub fn resolve_declared_variant(doc: &Doc, element: NodeId) -> Option<String> {
    for current in ancestors_and_self(doc, element) {
        if !is_mathml_element(doc, current) {
            break;
        }
        if let Some(declared) = doc.attr(current, "mathvariant").map(str::trim) {
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

pub fn resolve_effective_variant_auto(doc: &Doc, token: NodeId) -> Option<String> {
    let declared = resolve_declared_variant(doc, token);
    resolve_effective_variant(doc, token, declared.as_deref())
}

pub fn resolve_effective_variant(doc: &Doc, token: NodeId, declared: Option<&str>) -> Option<String> {
    if !is_mathml_element(doc, token) || !matches!(doc.local(token), "mi" | "mo" | "mn" | "mtext" | "ms") {
        return None;
    }
    if !is_null_or_white_space(declared) {
        return Some(to_lower_invariant(declared.unwrap().trim()));
    }
    if !doc.elements(token).is_empty() {
        return None;
    }
    let value = doc.value(token);
    let value = value.trim();
    let italic = doc.local(token) == "mi"
        && is_single_mathml_character(value)
        && !matches!(
            value,
            "\u{2032}" | "\u{2033}" | "\u{2034}" | "\u{2057}" | "\u{2035}" | "\u{2036}" | "\u{2037}"
        );
    Some(if italic { "italic" } else { "normal" }.to_string())
}

/// `\A(?:[\uD800-\uDBFF][\uDC00-\uDFFF]|[^\uD800-\uDFFF\r\n  ])
///  [̀-ͯ᪰-᪾᷀-᷿⃐-⃯]*\z`
pub fn is_single_mathml_character(text: &str) -> bool {
    let mut chars = text.chars();
    match chars.next() {
        None => false,
        Some('\r' | '\n' | '\u{2028}' | '\u{2029}') => false,
        Some(_) => chars.all(|c| {
            matches!(c as u32, 0x0300..=0x036F | 0x1AB0..=0x1ABE | 0x1DC0..=0x1DFF | 0x20D0..=0x20EF)
        }),
    }
}

// ---- MathMlColorContract -----------------------------------------------------

static HEX6: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"^#[0-9a-fA-F]{6}$").unwrap());
static HEX3: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"^#[0-9a-fA-F]{3}$").unwrap());

/// System.Drawing known, non-system, opaque colors (.NET 8: includes RebeccaPurple).
static NAMED_COLORS: &[(&str, u32)] = &[
    ("aliceblue", 0xF0F8FF), ("antiquewhite", 0xFAEBD7), ("aqua", 0x00FFFF), ("aquamarine", 0x7FFFD4),
    ("azure", 0xF0FFFF), ("beige", 0xF5F5DC), ("bisque", 0xFFE4C4), ("black", 0x000000),
    ("blanchedalmond", 0xFFEBCD), ("blue", 0x0000FF), ("blueviolet", 0x8A2BE2), ("brown", 0xA52A2A),
    ("burlywood", 0xDEB887), ("cadetblue", 0x5F9EA0), ("chartreuse", 0x7FFF00), ("chocolate", 0xD2691E),
    ("coral", 0xFF7F50), ("cornflowerblue", 0x6495ED), ("cornsilk", 0xFFF8DC), ("crimson", 0xDC143C),
    ("cyan", 0x00FFFF), ("darkblue", 0x00008B), ("darkcyan", 0x008B8B), ("darkgoldenrod", 0xB8860B),
    ("darkgray", 0xA9A9A9), ("darkgreen", 0x006400), ("darkkhaki", 0xBDB76B), ("darkmagenta", 0x8B008B),
    ("darkolivegreen", 0x556B2F), ("darkorange", 0xFF8C00), ("darkorchid", 0x9932CC), ("darkred", 0x8B0000),
    ("darksalmon", 0xE9967A), ("darkseagreen", 0x8FBC8B), ("darkslateblue", 0x483D8B), ("darkslategray", 0x2F4F4F),
    ("darkturquoise", 0x00CED1), ("darkviolet", 0x9400D3), ("deeppink", 0xFF1493), ("deepskyblue", 0x00BFFF),
    ("dimgray", 0x696969), ("dodgerblue", 0x1E90FF), ("firebrick", 0xB22222), ("floralwhite", 0xFFFAF0),
    ("forestgreen", 0x228B22), ("fuchsia", 0xFF00FF), ("gainsboro", 0xDCDCDC), ("ghostwhite", 0xF8F8FF),
    ("gold", 0xFFD700), ("goldenrod", 0xDAA520), ("gray", 0x808080), ("green", 0x008000),
    ("greenyellow", 0xADFF2F), ("honeydew", 0xF0FFF0), ("hotpink", 0xFF69B4), ("indianred", 0xCD5C5C),
    ("indigo", 0x4B0082), ("ivory", 0xFFFFF0), ("khaki", 0xF0E68C), ("lavender", 0xE6E6FA),
    ("lavenderblush", 0xFFF0F5), ("lawngreen", 0x7CFC00), ("lemonchiffon", 0xFFFACD), ("lightblue", 0xADD8E6),
    ("lightcoral", 0xF08080), ("lightcyan", 0xE0FFFF), ("lightgoldenrodyellow", 0xFAFAD2), ("lightgray", 0xD3D3D3),
    ("lightgreen", 0x90EE90), ("lightpink", 0xFFB6C1), ("lightsalmon", 0xFFA07A), ("lightseagreen", 0x20B2AA),
    ("lightskyblue", 0x87CEFA), ("lightslategray", 0x778899), ("lightsteelblue", 0xB0C4DE), ("lightyellow", 0xFFFFE0),
    ("lime", 0x00FF00), ("limegreen", 0x32CD32), ("linen", 0xFAF0E6), ("magenta", 0xFF00FF),
    ("maroon", 0x800000), ("mediumaquamarine", 0x66CDAA), ("mediumblue", 0x0000CD), ("mediumorchid", 0xBA55D3),
    ("mediumpurple", 0x9370DB), ("mediumseagreen", 0x3CB371), ("mediumslateblue", 0x7B68EE), ("mediumspringgreen", 0x00FA9A),
    ("mediumturquoise", 0x48D1CC), ("mediumvioletred", 0xC71585), ("midnightblue", 0x191970), ("mintcream", 0xF5FFFA),
    ("mistyrose", 0xFFE4E1), ("moccasin", 0xFFE4B5), ("navajowhite", 0xFFDEAD), ("navy", 0x000080),
    ("oldlace", 0xFDF5E6), ("olive", 0x808000), ("olivedrab", 0x6B8E23), ("orange", 0xFFA500),
    ("orangered", 0xFF4500), ("orchid", 0xDA70D6), ("palegoldenrod", 0xEEE8AA), ("palegreen", 0x98FB98),
    ("paleturquoise", 0xAFEEEE), ("palevioletred", 0xDB7093), ("papayawhip", 0xFFEFD5), ("peachpuff", 0xFFDAB9),
    ("peru", 0xCD853F), ("pink", 0xFFC0CB), ("plum", 0xDDA0DD), ("powderblue", 0xB0E0E6),
    ("purple", 0x800080), ("red", 0xFF0000), ("rosybrown", 0xBC8F8F), ("royalblue", 0x4169E1),
    ("saddlebrown", 0x8B4513), ("salmon", 0xFA8072), ("sandybrown", 0xF4A460), ("seagreen", 0x2E8B57),
    ("seashell", 0xFFF5EE), ("sienna", 0xA0522D), ("silver", 0xC0C0C0), ("skyblue", 0x87CEEB),
    ("slateblue", 0x6A5ACD), ("slategray", 0x708090), ("snow", 0xFFFAFA), ("springgreen", 0x00FF7F),
    ("steelblue", 0x4682B4), ("tan", 0xD2B48C), ("teal", 0x008080), ("thistle", 0xD8BFD8),
    ("tomato", 0xFF6347), ("turquoise", 0x40E0D0), ("violet", 0xEE82EE), ("wheat", 0xF5DEB3),
    ("white", 0xFFFFFF), ("whitesmoke", 0xF5F5F5), ("yellow", 0xFFFF00), ("yellowgreen", 0x9ACD32),
    ("rebeccapurple", 0x663399),
];

pub fn normalize_color(value: Option<&str>) -> Option<String> {
    let color = value?.trim();
    if color.is_empty() {
        return None;
    }
    if HEX6.is_match(color) {
        return Some(to_upper_invariant(color));
    }
    if HEX3.is_match(color) {
        let doubled: String = color[1..].chars().flat_map(|c| [c, c]).collect();
        return Some(to_upper_invariant(&format!("#{doubled}")));
    }
    NAMED_COLORS
        .iter()
        .find(|(name, _)| eq_ignore_case(Some(color), name))
        .map(|(_, rgb)| format!("#{rgb:06X}"))
}

pub fn explicit_color(doc: &Doc, element: NodeId) -> Option<String> {
    let raw = doc
        .attr(element, "mathcolor")
        .or_else(|| doc.attr_ns(element, MATHML_NS, "mathcolor"))
        .or_else(|| doc.attr(element, "color"))
        .or_else(|| doc.attr_ns(element, MATHML_NS, "color"));
    normalize_color(raw)
}

pub fn effective_color(doc: &Doc, element: NodeId) -> Option<String> {
    ancestors_and_self(doc, element).into_iter().find_map(|id| explicit_color(doc, id))
}

pub fn wrap_color(source: String, color: Option<&str>) -> String {
    match color {
        Some(color) if !is_null_or_white_space(Some(&source)) => {
            format!("\\textcolor{{{color}}}{{{source}}}")
        }
        _ => source,
    }
}

// ---- MathMlAccentGeometry ----------------------------------------------------

pub fn accent_is_stretchy(doc: &Doc, mark: NodeId) -> bool {
    let declared = doc.attr(mark, "stretchy").map(str::trim);
    if eq_ignore_case(declared, "false") || declared == Some("0") {
        return false;
    }
    if eq_ignore_case(declared, "true") || declared == Some("1") {
        return true;
    }
    let value = doc.value(mark);
    matches!(
        value.trim(),
        "^" | "\u{02C6}" | "\u{0302}" | "~" | "\u{02DC}" | "\u{0303}"
            | "\u{2192}" | "\u{2190}" | "\u{2194}" | "\u{20D7}" | "\u{20D6}" | "\u{20E1}"
            | "\u{00AF}" | "\u{203E}" | "\u{2015}" | "\u{02C9}" | "\u{2322}" | "\u{23DC}"
    )
}

// ---- MathSpaceContract.ParseWidth (LatexSource only) ---------------------------

static DIMENSION: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"^(?<number>[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+))\s*(?<unit>em|ex|pt|pc|px|in|cm|mm)$").unwrap()
});

static NATIVE_SPACES: &[(&str, &str)] = &[
    ("1em", "\\quad "),
    ("2em", "\\qquad "),
    ("0.167em", "\\,"),
    ("0.222em", "\\:"),
    ("0.278em", "\\;"),
    ("-0.167em", "\\!"),
];

/// `MathSpaceContract.ParseWidth(width).LatexSource`.
pub fn space_latex_source(width: Option<&str>) -> Option<String> {
    let mut source = width.unwrap_or("").trim();
    if source.is_empty() {
        source = "0em";
    }
    let captures = DIMENSION.captures(source)?;
    let number = &captures["number"];
    let unit = &captures["unit"];
    let canonical = format!("{}{}", normalize_decimal(number), unit);
    if let Some((_, latex)) = NATIVE_SPACES.iter().find(|(width, _)| *width == canonical) {
        return Some(latex.to_string());
    }
    Some(format!("\\hspace{{{number}{unit}}}"))
}

fn normalize_decimal(value: &str) -> String {
    let negative = value.starts_with('-');
    let value = if negative || value.starts_with('+') { &value[1..] } else { value };
    let dot = value.find('.');
    let whole = match dot {
        None => value,
        Some(dot) => &value[..dot],
    }
    .trim_start_matches('0');
    let fraction = match dot {
        None => "",
        Some(dot) => value[dot + 1..].trim_end_matches('0'),
    };
    let whole = if whole.is_empty() { "0" } else { whole };
    let result = if fraction.is_empty() { whole.to_string() } else { format!("{whole}.{fraction}") };
    if negative && result != "0" {
        format!("-{result}")
    } else {
        result
    }
}
