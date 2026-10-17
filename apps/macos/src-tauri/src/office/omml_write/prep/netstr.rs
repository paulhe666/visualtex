//! .NET string/char semantics the pre-passes rely on.
//!
//! * `string.Length`, `IndexOf`, `Substring` count UTF-16 code units.
//! * `string.Trim()` and `string.IsNullOrWhiteSpace` use `char.IsWhiteSpace`,
//!   which is exactly the Unicode White_Space set (Zs, Zl, Zp, U+0009..U+000D,
//!   U+0085), i.e. Rust's `char::is_whitespace`. .NET Regex `\s` is the same
//!   set (`[\f\n\r\t\v\x85\p{Z}]`).
//! * Invariant casing (ICU `ChangeCaseInvariant`, .NET 5+ on Windows) and
//!   OrdinalIgnoreCase (`OrdinalCasing`, ICU `u_toupper`) are simple, one-to-one
//!   case mappings, except that U+0131 stays U+0131 on upper-casing and U+0130
//!   stays U+0130 on lower-casing.

use crate::office::omml_write::prep::unicode_letters::LETTER_OR_LETTER_NUMBER;

/// `string.Length`: the number of UTF-16 code units.
pub fn len16(value: &str) -> usize {
    value.chars().map(char::len_utf16).sum()
}

/// `string.Trim()`.
pub fn trim(value: &str) -> &str {
    value.trim_matches(char::is_whitespace)
}

/// `string.IsNullOrWhiteSpace(value)`.
pub fn is_null_or_white_space(value: Option<&str>) -> bool {
    value.is_none_or(|text| text.chars().all(char::is_whitespace))
}

/// `string.IsNullOrWhiteSpace` for a non-null string.
pub fn is_white_space(value: &str) -> bool {
    value.chars().all(char::is_whitespace)
}

/// `string.IsNullOrEmpty(value)`.
pub fn is_null_or_empty(value: Option<&str>) -> bool {
    value.is_none_or(str::is_empty)
}

/// Simple upper-case mapping (`u_toupper`) as used by OrdinalIgnoreCase and
/// ToUpperInvariant. Rust exposes the full mapping; where that is a single
/// character it is the simple mapping. Where it expands, the simple mapping is
/// the character itself except for the Greek letters with ypogegrammeni,
/// which map to their prosgegrammeni (title-case) forms.
pub fn upper_char(c: char) -> char {
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
pub fn lower_char(c: char) -> char {
    if c == '\u{0130}' {
        return c;
    }
    let mut mapped = c.to_lowercase();
    match (mapped.next(), mapped.next()) {
        (Some(single), None) => single,
        _ => c,
    }
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
pub fn eq_ignore_case(a: Option<&str>, b: &str) -> bool {
    match a {
        None => false,
        Some(a) => a.chars().map(upper_char).eq(b.chars().map(upper_char)),
    }
}

/// `value.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0`.
/// Simple case mapping never changes the UTF-16 length, so comparing the
/// upper-cased strings is the ordinal ignore-case search.
pub fn contains_ignore_case(value: &str, needle: &str) -> bool {
    to_upper_invariant(value).contains(&to_upper_invariant(needle))
}

fn in_ranges(table: &[(u32, u32)], value: u32) -> bool {
    table
        .binary_search_by(|&(low, high)| {
            if high < value {
                std::cmp::Ordering::Less
            } else if low > value {
                std::cmp::Ordering::Greater
            } else {
                std::cmp::Ordering::Equal
            }
        })
        .is_ok()
}

/// WordOmmlConverter.ContainsUnicodeLetter: any code point (surrogate pairs
/// read as one, CharUnicodeInfo.GetUnicodeCategory(string, int)) whose
/// category is Lu, Ll, Lt, Lm, Lo or Nl.
pub fn contains_unicode_letter(value: &str) -> bool {
    value.chars().any(|c| in_ranges(LETTER_OR_LETTER_NUMBER, c as u32))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dotnet_string_primitives() {
        assert_eq!(len16("a\u{1D44E}"), 3);
        assert_eq!(trim("\u{A0} x \u{3000}"), "x");
        assert!(is_null_or_white_space(Some("\u{2028}\t")));
        assert!(!is_null_or_white_space(Some("\u{200B}")));
        assert!(eq_ignore_case(Some("Open"), "OPEN"));
        assert!(!eq_ignore_case(Some("\u{0131}nner"), "INNER"));
        assert!(eq_ignore_case(Some("\u{017F}"), "S"));
        assert_eq!(to_lower_invariant("BOLD-Italic\u{0130}"), "bold-italic\u{0130}");
        assert_eq!(to_upper_invariant("#a1b2c3"), "#A1B2C3");
        assert!(contains_ignore_case("<mtable DATA-MTEF-PILE=\"true\">", "data-mtef-pile"));
        assert!(contains_unicode_letter("sin"));
        assert!(contains_unicode_letter("\u{2160}"));
        assert!(contains_unicode_letter("\u{1D44E}"));
        assert!(!contains_unicode_letter("\u{2211}+\u{2061}"));
    }
}
