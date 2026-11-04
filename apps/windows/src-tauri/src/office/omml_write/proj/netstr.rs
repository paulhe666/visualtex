//! .NET 8 string/char semantics the writer passes rely on.
//!
//! * `string.Length`, indices and `Substring` count UTF-16 code units.
//! * `char.IsWhiteSpace` / `string.Trim()` / `Split(null, ...)` use Unicode
//!   White_Space, which is exactly Rust's `char::is_whitespace` (no astral
//!   character is White_Space, and a surrogate code unit never is).
//! * `CharUnicodeInfo.GetUnicodeCategory(string, index)` reads whole scalars;
//!   .NET 8 ships Unicode 15.0 data (`unicode_tables.rs`).
//! * .NET's Regex works on UTF-16 code units: `\p{L}` never matches either
//!   half of a surrogate pair, so astral letters are not letters to it.
//! * `StringInfo.ParseCombiningCharacters` is .NET 8's
//!   TextSegmentationUtility (UAX#29 rev. 35 extended grapheme clusters, no
//!   GB9c) over Unicode 15.0 Grapheme_Cluster_Break data.
//! * Invariant casing is simple (1:1) case mapping, except that .NET keeps
//!   U+0131 on ToUpper and U+0130 on ToLower (as in the reference port).

use crate::office::omml_write::proj::unicode_tables::{Gcb, GRAPHEME_BREAK, LETTER, LETTER_NUMBER};

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

/// `string.Length` (UTF-16 code units).
pub fn len16(value: &str) -> usize {
    value.chars().map(char::len_utf16).sum()
}

/// `char.IsWhiteSpace`.
pub fn is_white_space(c: char) -> bool {
    c.is_whitespace()
}

/// `string.IsNullOrWhiteSpace`.
pub fn is_null_or_white_space(value: &str) -> bool {
    value.chars().all(is_white_space)
}

/// `string.Trim()`.
pub fn trim(value: &str) -> &str {
    value.trim_matches(is_white_space)
}

/// WordOmmlConverter.ContainsUnicodeLetter: any scalar whose category is a
/// letter (Lu, Ll, Lt, Lm, Lo) or LetterNumber (Nl).
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

/// `value.Substring(start)` in UTF-16 units. Splitting a surrogate pair
/// would leave a lone surrogate, which the XML writer rejects; report it.
pub fn substring_from(value: &str, start: usize) -> Result<String, String> {
    substring(value, start, len16(value).saturating_sub(start))
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

/// `ToUpperInvariant()`.
pub fn to_upper_invariant(value: &str) -> String {
    value.chars().map(upper_char).collect()
}

/// `ToLowerInvariant()`.
pub fn to_lower_invariant(value: &str) -> String {
    value.chars().map(lower_char).collect()
}

/// `string.Equals(a, b, StringComparison.OrdinalIgnoreCase)` (null-safe).
pub fn equals_ordinal_ignore_case(a: Option<&str>, b: &str) -> bool {
    match a {
        None => false,
        Some(a) => a.chars().map(upper_char).eq(b.chars().map(upper_char)),
    }
}

/// `value.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0`.
pub fn contains_ordinal_ignore_case(value: &str, needle: &str) -> bool {
    to_upper_invariant(value).contains(&to_upper_invariant(needle))
}

#[cfg(test)]
mod tests {
    use super::*;

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
