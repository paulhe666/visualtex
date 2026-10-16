//! .NET string/char semantics the converter relies on.
//!
//! * `string.Trim()` trims `char.IsWhiteSpace`, which is exactly Unicode
//!   White_Space, i.e. Rust's `str::trim`.
//! * `char` predicates work per UTF-16 code unit; a surrogate unit is never a
//!   letter/digit/upper/lower, so astral characters fail those predicates.
//! * Invariant casing is simple (1:1) case mapping, except that .NET keeps
//!   U+0131 (dotless i) on ToUpper and U+0130 (dotted I) on ToLower.

use super::unicode_tables::{DIGIT_LOW, GREEK_LOWER, GREEK_UPPER, LETTER_LOW, MATH_SYMBOL};

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

/// UTF-16 length (`string.Length`).
pub fn len16(value: &str) -> usize {
    value.encode_utf16().count()
}

pub fn units(value: &str) -> Vec<u16> {
    value.encode_utf16().collect()
}

/// `char.IsLetter` for a code unit at or below U+024F (the only range the
/// converter can observe a letter result in).
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

/// `char.GetUnicodeCategory(s, 0) == UnicodeCategory.MathSymbol`.
pub fn first_is_math_symbol(value: &str) -> bool {
    value.chars().next().is_some_and(|c| in_ranges(MATH_SYMBOL, c as u32))
}

fn upper_char(c: char) -> char {
    if c == '\u{0131}' {
        return c;
    }
    let mut mapped = c.to_uppercase();
    match (mapped.next(), mapped.next()) {
        (Some(single), None) => single,
        _ => c,
    }
}

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

pub fn to_upper_invariant(value: &str) -> String {
    value.chars().map(upper_char).collect()
}

pub fn to_lower_invariant(value: &str) -> String {
    value.chars().map(lower_char).collect()
}

/// `string.Equals(a, b, StringComparison.OrdinalIgnoreCase)` (null-safe).
pub fn eq_ignore_case(a: Option<&str>, b: &str) -> bool {
    match a {
        None => false,
        Some(a) => a.chars().map(upper_char).eq(b.chars().map(upper_char)),
    }
}

/// `value.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0`.
pub fn contains_ignore_case(value: &str, needle: &str) -> bool {
    to_upper_invariant(value).contains(&to_upper_invariant(needle))
}

/// `string.IsNullOrWhiteSpace`.
pub fn is_null_or_white_space(value: Option<&str>) -> bool {
    value.is_none_or(|v| v.chars().all(char::is_whitespace))
}

/// `int.TryParse(s, out var n)` with NumberStyles.Integer (leading/trailing
/// ASCII white space, optional sign, ASCII digits, overflow fails).
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
