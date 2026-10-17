//! The VstoShared contracts the MathML pre-passes call into:
//! MathMlExplicitFenceContract, MathMlTokenStyleContract, MathMlColorContract
//! and MathSpaceContract (C# names are kept in the doc comments).

use crate::office::omml_write::prep::dom::{Doc, NodeId, MATHML_NS};
use crate::office::omml_write::prep::netstr::{eq_ignore_case, is_null_or_white_space, to_lower_invariant, to_upper_invariant, trim};
use crate::office::omml_write::prep::xlinq::{has_elements, is_namespace_declaration};

// ---- MathMlExplicitFenceContract ---------------------------------------------

/// MathMlExplicitFenceContract.HasTexClass: data-mjx-texclass equals
/// `expected` (OrdinalIgnoreCase).
pub fn has_tex_class(doc: &Doc, element: NodeId, expected: &str) -> bool {
    eq_ignore_case(doc.attr(element, "data-mjx-texclass"), expected)
}

fn is_mathml_or_none(ns: &str) -> bool {
    ns == MATHML_NS || ns.is_empty()
}

/// MathMlExplicitFenceContract.IsProjectableEndpoint.
pub fn is_projectable_endpoint(doc: &Doc, element: NodeId) -> bool {
    let Some(name) = doc.name(element) else { return false };
    if name.local != "mo" || !is_mathml_or_none(&name.ns) || has_elements(doc, element) {
        return false;
    }
    doc.attrs(element).iter().all(|(attribute, value)| {
        is_namespace_declaration(attribute)
            || (attribute.ns.is_empty() && attribute.local == "data-mjx-texclass")
            || (attribute.ns.is_empty()
                && matches!(attribute.local.as_str(), "fence" | "stretchy" | "symmetric")
                && eq_ignore_case(Some(value), "true"))
    })
}

/// MathMlExplicitFenceContract.TryGetRow; returns (open, close, body).
pub fn try_get_row(doc: &Doc, row: NodeId) -> Option<(String, String, Vec<NodeId>)> {
    let name = doc.name(row)?;
    if name.local != "mrow" || !is_mathml_or_none(&name.ns) || !has_tex_class(doc, row, "INNER") {
        return None;
    }
    let elements = doc.elements(row);
    if elements.len() < 2 {
        return None;
    }
    let first = elements[0];
    let last = elements[elements.len() - 1];
    if doc.name(first)?.ns != name.ns
        || doc.name(last)?.ns != name.ns
        || !has_tex_class(doc, first, "OPEN")
        || !has_tex_class(doc, last, "CLOSE")
        || !is_projectable_endpoint(doc, first)
        || !is_projectable_endpoint(doc, last)
    {
        return None;
    }
    let open = doc.value(first);
    let close = doc.value(last);
    if open.is_empty() && close.is_empty() {
        return None;
    }
    Some((open, close, elements[1..elements.len() - 1].to_vec()))
}

// ---- MathMlTokenStyleContract -------------------------------------------------

/// MathMlTokenStyleContract.IsMathMlElement.
fn is_mathml_element(doc: &Doc, element: NodeId) -> bool {
    doc.name(element).is_some_and(|name| is_mathml_or_none(&name.ns))
}

/// MathMlTokenStyleContract.ResolveDeclaredVariant.
pub fn resolve_declared_variant(doc: &Doc, element: NodeId) -> Option<String> {
    let mut current = Some(element);
    while let Some(node) = current {
        if !is_mathml_element(doc, node) {
            break;
        }
        if let Some(declared) = doc.attr(node, "mathvariant").map(trim) {
            if !declared.is_empty() {
                return Some(to_lower_invariant(declared));
            }
        }
        if doc.local(node) == "math" {
            break;
        }
        current = doc.parent(node);
    }
    None
}

/// MathMlTokenStyleContract.ResolveEffectiveVariant(token).
pub fn resolve_effective_variant(doc: &Doc, token: NodeId) -> Option<String> {
    let declared = resolve_declared_variant(doc, token);
    resolve_effective_variant_with(doc, token, declared.as_deref())
}

/// MathMlTokenStyleContract.ResolveEffectiveVariant(token, declaredVariant).
pub fn resolve_effective_variant_with(doc: &Doc, token: NodeId, declared: Option<&str>) -> Option<String> {
    if !is_mathml_element(doc, token) || !matches!(doc.local(token), "mi" | "mo" | "mn" | "mtext" | "ms") {
        return None;
    }
    if !is_null_or_white_space(declared) {
        return Some(to_lower_invariant(trim(declared.unwrap())));
    }
    if has_elements(doc, token) {
        return None;
    }
    let value = doc.value(token);
    let text = trim(&value);
    let italic = doc.local(token) == "mi"
        && is_single_mathml_character(text)
        && !matches!(text, "\u{2032}" | "\u{2033}" | "\u{2034}" | "\u{2057}" | "\u{2035}" | "\u{2036}" | "\u{2037}");
    Some(if italic { "italic" } else { "normal" }.to_string())
}

/// MathMlTokenStyleContract.IsSingleMathMlCharacter:
/// `\A(?:[\uD800-\uDBFF][\uDC00-\uDFFF]|[^\uD800-\uDFFF\r\n  ])`
/// `[̀-ͯ᪰-᪾᷀-᷿⃐-⃯]*\z`.
pub fn is_single_mathml_character(text: &str) -> bool {
    let mut chars = text.chars();
    match chars.next() {
        None => false,
        Some('\r' | '\n' | '\u{2028}' | '\u{2029}') => false,
        Some(_) => chars.all(|c| {
            matches!(c, '\u{0300}'..='\u{036F}' | '\u{1AB0}'..='\u{1ABE}' | '\u{1DC0}'..='\u{1DFF}' | '\u{20D0}'..='\u{20EF}')
        }),
    }
}

// ---- MathMlColorContract ------------------------------------------------------

/// System.Drawing known colors that are neither system colors nor
/// transparent (Color.FromName(name).IsKnownColor && !IsSystemColor && A == 255),
/// keyed by the Color property name (looked up OrdinalIgnoreCase).
static KNOWN_COLORS: &[(&str, u32)] = &[
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

fn is_hex_color(color: &str, digits: usize) -> bool {
    color.len() == digits + 1 && color.starts_with('#') && color[1..].bytes().all(|b| b.is_ascii_hexdigit())
}

/// MathMlColorContract.Normalize.
pub fn normalize_color(value: Option<&str>) -> Option<String> {
    let color = trim(value?);
    if color.is_empty() {
        return None;
    }
    // Regex "^#[0-9a-fA-F]{6}$": after Trim no trailing "\n" is left for "$".
    if is_hex_color(color, 6) {
        return Some(to_upper_invariant(color));
    }
    if is_hex_color(color, 3) {
        let doubled: String = color[1..].chars().flat_map(|c| [c, c]).collect();
        return Some(format!("#{}", to_upper_invariant(&doubled)));
    }
    // Color.FromName: known names are matched OrdinalIgnoreCase; system
    // colors and Transparent are rejected by the caller's checks.
    KNOWN_COLORS
        .iter()
        .find(|(name, _)| eq_ignore_case(Some(color), name))
        .map(|(_, rgb)| format!("#{:02X}{:02X}{:02X}", rgb >> 16, (rgb >> 8) & 0xFF, rgb & 0xFF))
}

/// MathMlColorContract.Explicit.
pub fn explicit_color(doc: &Doc, element: NodeId) -> Option<String> {
    let declared = doc
        .attr(element, "mathcolor")
        .or_else(|| doc.attr_ns(element, MATHML_NS, "mathcolor"))
        .or_else(|| doc.attr(element, "color"))
        .or_else(|| doc.attr_ns(element, MATHML_NS, "color"));
    normalize_color(declared)
}

/// MathMlColorContract.Effective: the first explicit color on the element
/// or its ancestors.
pub fn effective_color(doc: &Doc, element: NodeId) -> Option<String> {
    let mut current = Some(element);
    while let Some(node) = current {
        if let Some(color) = explicit_color(doc, node) {
            return Some(color);
        }
        current = doc.parent(node);
    }
    None
}

// ---- MathSpaceContract ----------------------------------------------------------

/// MathSpaceContract.NativeSpaces: (MathML width, LaTeX source).
const NATIVE_SPACES: &[(&str, &str)] = &[
    ("1em", "\\quad "),
    ("2em", "\\qquad "),
    ("0.167em", "\\,"),
    ("0.222em", "\\:"),
    ("0.278em", "\\;"),
    ("-0.167em", "\\!"),
];

/// MathSpaceContract.OmmlSpaces: (MathML width, OMML text).
const OMML_SPACES: &[(&str, &str)] = &[("1em", "\u{2003}"), ("2em", "\u{2003}\u{2003}")];

/// MathSpaceSourceSupport.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MathSpaceSourceSupport {
    ExactDimension,
    SymbolicMathSpace,
    NeedsContext,
    NotRepresented,
}

/// MathSpaceDescriptor (the parts the OMML projection reads).
#[derive(Clone, Debug)]
pub struct MathSpaceDescriptor {
    pub source_width: String,
    pub canonical_width: Option<String>,
    pub latex_source: Option<String>,
    pub source_support: MathSpaceSourceSupport,
}

impl MathSpaceDescriptor {
    pub fn source_supported(&self) -> bool {
        self.source_support == MathSpaceSourceSupport::ExactDimension
    }
}

/// MathSpaceOmmlCapability.
#[derive(Clone, Debug)]
pub struct MathSpaceOmmlCapability {
    pub width: MathSpaceDescriptor,
    pub text: Option<String>,
    pub height_supported: bool,
    pub depth_supported: bool,
    pub linebreak_supported: bool,
    pub audit_limits: Vec<String>,
}

impl MathSpaceOmmlCapability {
    pub fn width_supported(&self) -> bool {
        self.text.is_some()
    }
    pub fn can_project(&self) -> bool {
        self.audit_limits.is_empty()
    }
}

/// MathSpaceContract.DimensionPattern:
/// `\A(?<number>[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+))\s*(?<unit>em|ex|pt|pc|px|in|cm|mm)\z`
/// (.NET `\s` is char.IsWhiteSpace). Returns (number, unit).
fn match_dimension(source: &str) -> Option<(&str, &str)> {
    let bytes = source.as_bytes();
    let mut index = 0;
    if matches!(bytes.first(), Some(b'+' | b'-')) {
        index += 1;
    }
    let integer_start = index;
    while index < bytes.len() && bytes[index].is_ascii_digit() {
        index += 1;
    }
    let has_integer = index > integer_start;
    if index < bytes.len() && bytes[index] == b'.' {
        index += 1;
        let fraction_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if !has_integer && index == fraction_start {
            return None;
        }
    } else if !has_integer {
        return None;
    }
    let number = &source[..index];
    let rest = source[index..].trim_start_matches(char::is_whitespace);
    match rest {
        "em" | "ex" | "pt" | "pc" | "px" | "in" | "cm" | "mm" => Some((number, rest)),
        _ => None,
    }
}

/// MathSpaceContract.NormalizeDecimal (lexical, no floating point).
fn normalize_decimal(value: &str) -> String {
    let negative = value.starts_with('-');
    let value = if negative || value.starts_with('+') { &value[1..] } else { value };
    let dot = value.find('.');
    let whole = match dot {
        None => value,
        Some(index) => &value[..index],
    }
    .trim_start_matches('0');
    let fraction = match dot {
        None => "",
        Some(index) => value[index + 1..].trim_end_matches('0'),
    };
    let whole = if whole.is_empty() { "0" } else { whole };
    let result = if fraction.is_empty() { whole.to_string() } else { format!("{whole}.{fraction}") };
    if negative && result != "0" {
        format!("-{result}")
    } else {
        result
    }
}

/// MathSpaceContract.ParseWidth.
pub fn parse_width(width: Option<&str>) -> MathSpaceDescriptor {
    let mut source = trim(width.unwrap_or("")).to_string();
    if source.is_empty() {
        source = "0em".to_string();
    }
    if let Some((number, unit)) = match_dimension(&source) {
        let canonical = normalize_decimal(number) + unit;
        let native = NATIVE_SPACES.iter().find(|(width, _)| *width == canonical);
        let latex = match native {
            Some((_, latex)) => latex.to_string(),
            None => format!("\\hspace{{{number}{unit}}}"),
        };
        return MathSpaceDescriptor {
            source_width: source.clone(),
            canonical_width: Some(canonical),
            latex_source: Some(latex),
            source_support: MathSpaceSourceSupport::ExactDimension,
        };
    }
    let symbolic = source.strip_prefix("negative").unwrap_or(&source);
    if matches!(
        symbolic,
        "veryverythinmathspace"
            | "verythinmathspace"
            | "thinmathspace"
            | "mediummathspace"
            | "thickmathspace"
            | "verythickmathspace"
            | "veryverythickmathspace"
    ) {
        return MathSpaceDescriptor {
            canonical_width: Some(source.clone()),
            source_width: source,
            latex_source: None,
            source_support: MathSpaceSourceSupport::SymbolicMathSpace,
        };
    }
    let support = if source.ends_with('%') {
        MathSpaceSourceSupport::NeedsContext
    } else {
        MathSpaceSourceSupport::NotRepresented
    };
    MathSpaceDescriptor { source_width: source, canonical_width: None, latex_source: None, source_support: support }
}

/// MathSpaceContract.HaveSameWidth.
pub fn have_same_width(first: Option<&str>, second: Option<&str>) -> bool {
    let left = parse_width(first);
    let right = parse_width(second);
    let zero = |descriptor: &MathSpaceDescriptor| {
        let canonical = descriptor.canonical_width.as_deref().unwrap();
        &canonical[..canonical.len() - 2] == "0"
    };
    if left.source_supported() && right.source_supported() && zero(&left) && zero(&right) {
        return true;
    }
    match (&left.canonical_width, &right.canonical_width) {
        (Some(a), Some(b)) => a == b,
        _ => false,
    }
}

/// MathSpaceContract.DescribeOmmlSpace(width, height, depth, linebreak).
pub fn describe_omml_space(
    width: Option<&str>,
    height: Option<&str>,
    depth: Option<&str>,
    linebreak: Option<&str>,
) -> MathSpaceOmmlCapability {
    let descriptor = parse_width(width);
    let identity = OMML_SPACES
        .iter()
        .find(|(mathml_width, _)| descriptor.canonical_width.as_deref() == Some(*mathml_width));
    let zero_width = have_same_width(width, Some("0em"));
    let height_supported = have_same_width(height, Some("0em"));
    let depth_supported = have_same_width(depth, Some("0em"));
    let linebreak_supported = is_null_or_white_space(linebreak) || trim(linebreak.unwrap()) == "auto";
    let mut limits = Vec::new();
    if identity.is_none() && !zero_width {
        limits.push(format!(
            "OMML has no verified exact mapping for width '{}'; the width is not rounded to another Unicode space.",
            descriptor.source_width
        ));
    }
    if !height_supported {
        limits.push(format!("OMML explicit-space height is not represented: '{}'.", height.unwrap_or("")));
    }
    if !depth_supported {
        limits.push(format!("OMML explicit-space depth is not represented: '{}'.", depth.unwrap_or("")));
    }
    if !linebreak_supported {
        limits.push(format!("OMML explicit-space linebreak is not represented: '{}'.", linebreak.unwrap_or("")));
    }
    let text = match identity {
        Some((_, text)) => Some(text.to_string()),
        None => zero_width.then(String::new),
    };
    MathSpaceOmmlCapability {
        width: descriptor,
        text,
        height_supported,
        depth_supported,
        linebreak_supported,
        audit_limits: limits,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn colors_normalize_like_system_drawing() {
        assert_eq!(normalize_color(Some(" #a1b2c3 ")).as_deref(), Some("#A1B2C3"));
        assert_eq!(normalize_color(Some("#0a0")).as_deref(), Some("#00AA00"));
        assert_eq!(normalize_color(Some("red")).as_deref(), Some("#FF0000"));
        assert_eq!(normalize_color(Some("DARKSEAGREEN")).as_deref(), Some("#8FBC8B"));
        assert_eq!(normalize_color(Some("rebeccapurple")).as_deref(), Some("#663399"));
        assert_eq!(normalize_color(Some("transparent")), None);
        assert_eq!(normalize_color(Some("Control")), None);
        assert_eq!(normalize_color(Some("grey")), None);
        assert_eq!(normalize_color(Some("#12345")), None);
        assert_eq!(normalize_color(Some("  ")), None);
    }

    #[test]
    fn widths_are_lexical() {
        assert_eq!(parse_width(Some("1.0em")).canonical_width.as_deref(), Some("1em"));
        assert_eq!(parse_width(Some("+.50 em")).canonical_width.as_deref(), Some("0.5em"));
        assert_eq!(parse_width(Some("-0.000mm")).canonical_width.as_deref(), Some("0mm"));
        assert_eq!(parse_width(Some("007.em")).canonical_width.as_deref(), Some("7em"));
        assert_eq!(parse_width(Some("007.")).canonical_width, None);
        assert_eq!(parse_width(Some("thinmathspace")).canonical_width.as_deref(), Some("thinmathspace"));
        assert!(have_same_width(Some("0pt"), Some("0em")));
        assert!(!have_same_width(Some("0.1pt"), Some("0em")));
        let capability = describe_omml_space(Some(" 1.00em"), None, Some("0ex"), Some(" auto "));
        assert!(capability.can_project());
        assert_eq!(capability.text.as_deref(), Some("\u{2003}"));
        assert!(!describe_omml_space(Some("0.167em"), None, None, None).can_project());
        assert_eq!(describe_omml_space(None, None, None, None).text.as_deref(), Some(""));
    }
}
