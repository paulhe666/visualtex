//! .NET 8 string and char semantics the ports of the Windows converters rely on.
//!
//! * `string.Length`, `IndexOf`, indices and `Substring` count UTF-16 code units.
//! * `char.IsWhiteSpace` / `string.Trim()` / `string.IsNullOrWhiteSpace` / `Split(null, ...)` use Unicode
//!   White_Space (Zs, Zl, Zp, U+0009..U+000D, U+0085), which is exactly Rust's `char::is_whitespace` (no astral
//!   character is White_Space, and a surrogate code unit never is). .NET Regex `\s` is the same set.
//! * `char` predicates work per UTF-16 code unit: a surrogate unit is never a letter, digit, upper or lower, so
//!   astral characters fail them. `CharUnicodeInfo.GetUnicodeCategory(string, index)` reads whole scalars; .NET 8
//!   ships Unicode 15.0 data (`unicode_tables.rs`).
//! * .NET's Regex works on UTF-16 code units: `\p{L}` never matches either half of a surrogate pair, so astral
//!   letters are not letters to it.
//! * `StringInfo.ParseCombiningCharacters` is .NET 8's TextSegmentationUtility (UAX#29 rev. 35 extended grapheme
//!   clusters, no GB9c) over Unicode 15.0 Grapheme_Cluster_Break data.
//! * Invariant casing (ToUpperInvariant / ToLowerInvariant, ICU `u_toupper` / `u_tolower`) is simple, one-to-one
//!   case mapping, except that U+0131 stays U+0131 on upper-casing and U+0130 stays U+0130 on lower-casing.
//!   OrdinalIgnoreCase (`OrdinalCasing`) compares the same upper-case forms, except that it keeps U+017F (long s).
//!   Checked against .NET 8 on Windows for every code point; case pairs newer than the Unicode data of Windows'
//!   ICU are the only differences.

use super::unicode_tables::{
    Gcb, DIGIT_LOW, GRAPHEME_BREAK, GREEK_LOWER, GREEK_UPPER, LETTER, LETTER_LOW, LETTER_NUMBER, MATH_SYMBOL,
};
use regex::Regex;
use std::sync::LazyLock;

fn in_ranges(table: &[(u32, u32)], value: u32) -> bool {
    table
        .binary_search_by(|&(lo, hi)| {
            if hi < value {
                std::cmp::Ordering::Less
            } else if lo > value {
                std::cmp::Ordering::Greater
            } else {
                std::cmp::Ordering::Equal
            }
        })
        .is_ok()
}

// ---- length, code units, substrings -----------------------------------------

/// `string.Length`: the number of UTF-16 code units.
pub fn len16(value: &str) -> usize {
    value.chars().map(char::len_utf16).sum()
}

/// The string's UTF-16 code units (`string[index]`).
pub fn units(value: &str) -> Vec<u16> {
    value.encode_utf16().collect()
}

/// `value.Substring(start, length)` in UTF-16 units.
pub fn substring(value: &str, start: usize, length: usize) -> Result<String, String> {
    let units: Vec<u16> = value.encode_utf16().collect();
    if start > units.len() || start + length > units.len() {
        return Err("ArgumentOutOfRangeException: Index and length must refer to a location within the string."
            .to_string());
    }
    String::from_utf16(&units[start..start + length])
        .map_err(|_| "ArgumentException: The substring splits a surrogate pair (invalid XML character).".to_string())
}

/// `value.Substring(start)` in UTF-16 units. Splitting a surrogate pair
/// would leave a lone surrogate, which the XML writer rejects; report it.
pub fn substring_from(value: &str, start: usize) -> Result<String, String> {
    substring(value, start, len16(value).saturating_sub(start))
}

// ---- white space ------------------------------------------------------------

/// `char.IsWhiteSpace`.
pub fn is_white_space(c: char) -> bool {
    c.is_whitespace()
}

/// `string.Trim()`.
pub fn trim(value: &str) -> &str {
    value.trim_matches(is_white_space)
}

/// `string.IsNullOrWhiteSpace(value)`, for a string or a nullable one.
pub fn is_null_or_white_space<'a>(value: impl Into<Option<&'a str>>) -> bool {
    value.into().is_none_or(|text| text.chars().all(is_white_space))
}

/// `string.IsNullOrEmpty(value)`.
pub fn is_null_or_empty(value: Option<&str>) -> bool {
    value.is_none_or(str::is_empty)
}

// ---- casing -----------------------------------------------------------------

/// Simple upper-case mapping (`u_toupper`) as used by ToUpperInvariant. Rust exposes the full mapping; where that
/// is a single character it is the simple mapping. Where it expands, the simple mapping is the character itself
/// except for the Greek letters with ypogegrammeni, which map to their prosgegrammeni (title-case) forms.
fn upper_char(c: char) -> char {
    if c == '\u{0131}' {
        return c;
    }
    let code = c as u32;
    let greek = match code {
        0x1F80..=0x1F87 | 0x1F90..=0x1F97 | 0x1FA0..=0x1FA7 => Some(code + 8),
        0x1FB3 => Some(0x1FBC),
        0x1FC3 => Some(0x1FCC),
        0x1FF3 => Some(0x1FFC),
        _ => None,
    };
    if let Some(mapped) = greek.and_then(char::from_u32) {
        return mapped;
    }
    let mut mapped = c.to_uppercase();
    match (mapped.next(), mapped.next()) {
        (Some(single), None) => single,
        _ => c,
    }
}

/// Simple lower-case mapping (`u_tolower`) as used by ToLowerInvariant. The
/// full lower-case mapping is single-character for everything but U+0130,
/// which .NET keeps.
fn lower_char(c: char) -> char {
    if c == '\u{0130}' {
        return c;
    }
    let mut mapped = c.to_lowercase();
    match (mapped.next(), mapped.next()) {
        (Some(single), None) => single,
        _ => c,
    }
}

/// The upper-case mapping OrdinalIgnoreCase compares with (.NET's OrdinalCasing): ToUpperInvariant's, except that
/// U+017F (long s) stays itself.
fn ordinal_upper_char(c: char) -> char {
    if c == '\u{017F}' {
        return c;
    }
    upper_char(c)
}

fn ordinal_upper(value: &str) -> String {
    value.chars().map(ordinal_upper_char).collect()
}

/// `string.ToUpperInvariant()`.
pub fn to_upper_invariant(value: &str) -> String {
    value.chars().map(upper_char).collect()
}

/// `string.ToLowerInvariant()`.
pub fn to_lower_invariant(value: &str) -> String {
    value.chars().map(lower_char).collect()
}

/// `string.Equals(a, b, StringComparison.OrdinalIgnoreCase)`; a null `a` is
/// never equal to a literal.
pub fn eq_ignore_case<'a>(a: impl Into<Option<&'a str>>, b: &str) -> bool {
    match a.into() {
        None => false,
        Some(a) => a.chars().map(ordinal_upper_char).eq(b.chars().map(ordinal_upper_char)),
    }
}

/// `value.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0`.
/// Simple case mapping never changes the UTF-16 length, so comparing the
/// upper-cased strings is the ordinal ignore-case search.
pub fn contains_ignore_case(value: &str, needle: &str) -> bool {
    ordinal_upper(value).contains(&ordinal_upper(needle))
}

// ---- numbers ----------------------------------------------------------------

/// `int.TryParse(s, out var n)` with NumberStyles.Integer (leading/trailing
/// ASCII white space, optional sign, ASCII digits, trailing NULs, overflow fails).
pub fn int_try_parse(value: Option<&str>) -> Option<i32> {
    let value = value?;
    let white = |c: char| matches!(c, '\t' | '\n' | '\u{000B}' | '\u{000C}' | '\r' | ' ');
    let mut text = value.trim_start_matches(white);
    text = text.trim_end_matches('\0');
    text = text.trim_end_matches(white);
    let (negative, digits) = if let Some(rest) = text.strip_prefix('-') {
        (true, rest)
    } else if let Some(rest) = text.strip_prefix('+') {
        (false, rest)
    } else {
        (false, text)
    };
    if digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let mut result: i64 = 0;
    for b in digits.bytes() {
        result = result * 10 + (b - b'0') as i64;
        if result > i32::MAX as i64 + 1 {
            return None;
        }
    }
    let result = if negative { -result } else { result };
    i32::try_from(result).ok()
}

// ---- character classes ------------------------------------------------------

/// `char.IsLetter` for a code unit at or below U+024F (the only range the
/// MathML -> LaTeX converter can observe a letter result in).
pub fn is_letter_low(unit: u16) -> bool {
    unit <= 0x24F && in_ranges(LETTER_LOW, unit as u32)
}

/// `char.IsLetterOrDigit` for a code unit at or below U+024F.
pub fn is_letter_or_digit_low(unit: u16) -> bool {
    unit <= 0x24F && (in_ranges(LETTER_LOW, unit as u32) || in_ranges(DIGIT_LOW, unit as u32))
}

/// `char.IsLower` for a Greek-range code unit.
pub fn is_lower_greek(unit: u16) -> bool {
    in_ranges(GREEK_LOWER, unit as u32)
}

/// `char.IsUpper` for a Greek-range code unit.
pub fn is_upper_greek(unit: u16) -> bool {
    in_ranges(GREEK_UPPER, unit as u32)
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

/// `CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.MathSymbol`.
pub fn is_math_symbol(c: char) -> bool {
    in_ranges(MATH_SYMBOL, c as u32)
}

/// `char.GetUnicodeCategory(s, 0) == UnicodeCategory.MathSymbol`.
pub fn first_is_math_symbol(value: &str) -> bool {
    value.chars().next().is_some_and(is_math_symbol)
}

/// WordOmmlConverter.ContainsUnicodeLetter: any code point (surrogate pairs
/// read as one, CharUnicodeInfo.GetUnicodeCategory(string, int)) whose
/// category is Lu, Ll, Lt, Lm, Lo or Nl.
pub fn contains_unicode_letter(value: &str) -> bool {
    value.chars().any(|c| in_ranges(LETTER, c as u32) || in_ranges(LETTER_NUMBER, c as u32))
}

/// `Regex.Match(value, @"\p{L}+")` (.NET, UTF-16 code units): the first
/// maximal run of BMP letters, or None.
pub fn first_letter_run(value: &str) -> Option<String> {
    let mut run = String::new();
    for c in value.chars() {
        // A surrogate code unit is UnicodeCategory.Surrogate for the regex.
        let letter = (c as u32) < 0x10000 && in_ranges(LETTER, c as u32);
        if letter {
            run.push(c);
        } else if !run.is_empty() {
            break;
        }
    }
    if run.is_empty() {
        None
    } else {
        Some(run)
    }
}

// ---- text elements ----------------------------------------------------------

fn grapheme_break(c: Option<char>) -> Gcb {
    // EOF (and ill-formed data, decoded as U+FFFD) is Other.
    let Some(c) = c else { return Gcb::Other };
    let value = c as u32;
    match GRAPHEME_BREAK.binary_search_by(|&(lo, hi, _)| {
        if hi < value {
            std::cmp::Ordering::Less
        } else if lo > value {
            std::cmp::Ordering::Greater
        } else {
            std::cmp::Ordering::Equal
        }
    }) {
        Ok(index) => GRAPHEME_BREAK[index].2,
        Err(_) => Gcb::Other,
    }
}

/// TextSegmentationUtility.GetLengthOfFirstExtendedGraphemeCluster (.NET 8),
/// in scalars of `chars`.
fn first_cluster_scalars(chars: &[char]) -> usize {
    // Processor: `offset` is the index of the current scalar; the type of the
    // position past the end is Other.
    let mut offset = 0usize;
    let type_at = |index: usize| grapheme_break(chars.get(index).copied());
    let mut current = type_at(offset);

    while current == Gcb::Prepend {
        offset += 1;
        current = type_at(offset);
    }
    if offset > 0 && matches!(current, Gcb::Control | Gcb::CR | Gcb::LF) {
        return offset;
    }

    let previous = current;
    offset = (offset + 1).min(chars.len());
    current = type_at(offset);
    let advance = |offset: &mut usize, current: &mut Gcb| {
        *offset = (*offset + 1).min(chars.len());
        *current = type_at(*offset);
    };

    #[derive(Clone, Copy)]
    enum State {
        Cr,
        ControlOrLf,
        L,
        LvOrV,
        LvtOrT,
        ExtPict,
        Ri,
        Default,
    }
    let mut state = match previous {
        Gcb::CR => State::Cr,
        Gcb::Control | Gcb::LF => State::ControlOrLf,
        Gcb::L => State::L,
        Gcb::LV | Gcb::V => State::LvOrV,
        Gcb::LVT | Gcb::T => State::LvtOrT,
        Gcb::ExtPict => State::ExtPict,
        Gcb::RegionalIndicator => State::Ri,
        _ => State::Default,
    };
    loop {
        match state {
            State::Cr => {
                if current != Gcb::LF {
                    return offset;
                }
                advance(&mut offset, &mut current);
                state = State::ControlOrLf;
            }
            State::ControlOrLf => return offset,
            State::L => match current {
                Gcb::L => {
                    advance(&mut offset, &mut current);
                    state = State::L;
                }
                Gcb::V | Gcb::LV => {
                    advance(&mut offset, &mut current);
                    state = State::LvOrV;
                }
                Gcb::LVT => {
                    advance(&mut offset, &mut current);
                    state = State::LvtOrT;
                }
                _ => break,
            },
            State::LvOrV => match current {
                Gcb::V => {
                    advance(&mut offset, &mut current);
                    state = State::LvOrV;
                }
                Gcb::T => {
                    advance(&mut offset, &mut current);
                    state = State::LvtOrT;
                }
                _ => break,
            },
            State::LvtOrT => {
                if current == Gcb::T {
                    advance(&mut offset, &mut current);
                    state = State::LvtOrT;
                } else {
                    break;
                }
            }
            State::ExtPict => {
                while current == Gcb::Extend {
                    advance(&mut offset, &mut current);
                }
                if current != Gcb::ZWJ {
                    break;
                }
                advance(&mut offset, &mut current);
                if current != Gcb::ExtPict {
                    break;
                }
                advance(&mut offset, &mut current);
                state = State::ExtPict;
            }
            State::Ri => {
                if current == Gcb::RegionalIndicator {
                    advance(&mut offset, &mut current);
                }
                break;
            }
            State::Default => break,
        }
    }
    // rules GB9, GB9a
    while matches!(current, Gcb::Extend | Gcb::ZWJ | Gcb::SpacingMark) {
        advance(&mut offset, &mut current);
    }
    offset
}

/// `StringInfo.ParseCombiningCharacters(value)`: UTF-16 start offsets of the
/// text elements.
pub fn parse_combining_characters(value: &str) -> Vec<usize> {
    let chars: Vec<char> = value.chars().collect();
    let mut starts = Vec::new();
    let mut index = 0usize;
    let mut offset16 = 0usize;
    while index < chars.len() {
        starts.push(offset16);
        let length = first_cluster_scalars(&chars[index..]).max(1);
        for c in &chars[index..index + length] {
            offset16 += c.len_utf16();
        }
        index += length;
    }
    starts
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dotnet_string_primitives() {
        assert_eq!(len16("a\u{1D44E}"), 3);
        assert_eq!(trim("\u{A0} x \u{3000}"), "x");
        assert!(is_null_or_white_space(Some("\u{2028}\t")));
        assert!(is_null_or_white_space("\u{2028}\t"));
        assert!(is_null_or_white_space(None));
        assert!(!is_null_or_white_space(Some("\u{200B}")));
        assert!(eq_ignore_case(Some("Open"), "OPEN"));
        assert!(eq_ignore_case("Open", "OPEN"));
        assert!(!eq_ignore_case(None, "OPEN"));
        assert!(!eq_ignore_case(Some("\u{0131}nner"), "INNER"));
        assert_eq!(to_lower_invariant("BOLD-Italic\u{0130}"), "bold-italic\u{0130}");
        assert_eq!(to_upper_invariant("#a1b2c3"), "#A1B2C3");
        assert!(contains_ignore_case("<mtable DATA-MTEF-PILE=\"true\">", "data-mtef-pile"));
        assert_eq!(int_try_parse(Some(" -12 ")), Some(-12));
        assert_eq!(int_try_parse(Some("12\0")), Some(12));
        assert_eq!(int_try_parse(Some("2147483648")), None);
        assert!(contains_unicode_letter("sin"));
        assert!(contains_unicode_letter("\u{2160}"));
        assert!(contains_unicode_letter("\u{1D44E}"));
        assert!(!contains_unicode_letter("\u{2211}+\u{2061}"));
        assert!(is_math_symbol('\u{2211}') && first_is_math_symbol("+x") && !is_math_symbol('x'));
        assert!(is_dotnet_digit('\u{0663}') && !is_dotnet_digit('\u{1D7CE}'));
    }

    #[test]
    fn casing_follows_dotnet() {
        assert_eq!(to_upper_invariant("\u{1F80}\u{1FB3}\u{017F}\u{0131}"), "\u{1F88}\u{1FBC}S\u{0131}");
        assert!(eq_ignore_case(Some("\u{1F80}"), "\u{1F88}"));
        assert!(!eq_ignore_case(Some("\u{017F}"), "S"));
        assert!(!contains_ignore_case("x\u{017F}y", "s"));
        assert_eq!(to_lower_invariant("\u{130}BOLD"), "\u{130}bold");
    }

    #[test]
    fn text_elements() {
        assert_eq!(parse_combining_characters("abc"), vec![0, 1, 2]);
        assert_eq!(parse_combining_characters("e\u{301}x"), vec![0, 2]);
        assert_eq!(parse_combining_characters("\r\nx"), vec![0, 2]);
        assert_eq!(parse_combining_characters("\u{1D465}y"), vec![0, 2]);
        assert_eq!(parse_combining_characters("\u{1F468}\u{200D}\u{1F469}a"), vec![0, 5]);
        assert_eq!(parse_combining_characters("\u{1F1FA}\u{1F1F8}\u{1F1FA}"), vec![0, 4]);
        assert_eq!(parse_combining_characters(""), Vec::<usize>::new());
    }

    #[test]
    fn letters() {
        assert_eq!(first_letter_run("2sin x").as_deref(), Some("sin"));
        assert_eq!(first_letter_run("\u{1D465}ab").as_deref(), Some("ab"));
        assert!(contains_unicode_letter("\u{1D465}"));
        assert!(contains_unicode_letter("\u{2160}"));
        assert!(!contains_unicode_letter("+2"));
    }
}
