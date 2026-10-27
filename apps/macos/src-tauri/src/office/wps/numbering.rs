//! Equation numbers for formulas the WPS add-in creates (office::wps).
//!
//! The Word add-in decides a new number from the document (VTWordAdapter.bas:
//! VTV2PlanInsertion, VTV2HeadingScope, VTEquationStartsNumberScope, ...,
//! themselves the port of Windows' WordEquationNumbering). WPS has no such
//! code, so the add-in reads the same facts from the WPS document (headings,
//! numbered equations and pictures) and these functions apply the same rules,
//! so a number made in WPS is the number Word would have made, and Word keeps
//! it when it renumbers the document.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;

pub(crate) const MODE_SEQUENCE: &str = "sequence";
pub(crate) const MODE_CHAPTER: &str = "chapter";
pub(crate) const MODE_SECTION: &str = "section";
/// SEQ label of numbered image formulas (VT_WORD_IMAGE_SEQUENCE_NAME).
pub(crate) const IMAGE_SEQUENCE_NAME: &str = "VisualTeXEquation";
/// Word's built-in Equation caption label as Word for Mac in English names it;
/// the label of a document's existing numbers wins.
pub(crate) const DEFAULT_EQUATION_LABEL: &str = "Equation";

/// One heading paragraph (outline level 1..2, not in a table or frame).
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct HeadingFact {
    pub start: i64,
    pub level: u32,
    /// The paragraph's list label (ListFormat.ListString), "" without one.
    #[serde(default)]
    pub list_string: String,
    /// The paragraph carries Word list numbering (ListType is not none).
    #[serde(default)]
    pub list_numbered: bool,
    #[serde(default)]
    pub text: String,
}

/// One numbered display equation: #( number ) with its SEQ field.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct EquationFact {
    pub start: i64,
    /// The SEQ field's instruction.
    pub instruction: String,
    /// The visible number inside the parentheses ("2.1-3").
    pub number_text: String,
}

/// One numbered image formula (Title ...:block:1).
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ImageFact {
    pub start: i64,
    /// The instruction of its hidden increment SEQ field.
    #[serde(default)]
    pub increment_instruction: String,
}

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct NumberingFacts {
    /// Where the new formula goes.
    pub position: i64,
    #[serde(default)]
    pub headings: Vec<HeadingFact>,
    #[serde(default)]
    pub equations: Vec<EquationFact>,
    #[serde(default)]
    pub images: Vec<ImageFact>,
    /// Earlier versions' document variables VT_EquationNumberingMode and
    /// VT_EquationNumberingSeparator, when the document has them.
    #[serde(default)]
    pub mode_variable: Option<String>,
    #[serde(default)]
    pub separator_variable: Option<String>,
}

/// The document's numbering style: mode and separator ("." or "-").
#[derive(Debug, Clone, PartialEq, Serialize)]
pub(crate) struct NumberingStyle {
    pub mode: String,
    pub separator: String,
}

impl NumberingStyle {
    fn sequence() -> Self {
        Self { mode: MODE_SEQUENCE.to_string(), separator: ".".to_string() }
    }

    /// VTV2DocumentFormat: heading level 0 (continuous), 1 (chapter) or 2
    /// (section) and the number separator ("" when continuous).
    pub(crate) fn level_and_separator(&self) -> (u32, String) {
        let level = match self.mode.as_str() {
            MODE_CHAPTER => 1,
            MODE_SECTION => 2,
            _ => 0,
        };
        let separator = if level == 0 {
            String::new()
        } else if self.separator == "." {
            ".".to_string()
        } else {
            "-".to_string()
        };
        (level, separator)
    }
}

/// The number of a new numbered equation (omml_write's numbered builder input).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct EquationPlan {
    pub sequence_name: String,
    pub heading_level: u32,
    pub separator: String,
    /// Written as text before the separator; None with no heading level or
    /// when a STYLEREF field reads the heading's list number.
    pub literal_prefix: Option<String>,
    pub automatic: bool,
    /// The \r value, when the new number starts its series.
    pub restart: Option<u32>,
    pub initial: u32,
    /// The start of the equation whose \r switch moves to the new number.
    #[serde(default)]
    pub transfer_from: Option<i64>,
}

/// The number of a new numbered image formula.
#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ImagePlan {
    /// Heading prefix and display separator ("0.0‐"), "" when continuous.
    pub prefix_text: String,
    /// The hidden increment SEQ restarts at 1 (first in its chapter/section).
    pub restart_at_one: bool,
    /// The next numbered image formula whose restart changes, with the value it
    /// needs (VTReconcileNumberingAfterEdit).
    pub next_restart: Option<NextImageRestart>,
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct NextImageRestart {
    pub start: i64,
    pub restart_at_one: bool,
}

// ---- text helpers (VTV2IsSpace, VTV2IsDigit, ...) ----------------------------

/// .NET char.IsWhiteSpace for the characters headings use (VTV2IsSpace).
fn is_space(character: char) -> bool {
    matches!(character as u32, 9..=13 | 32 | 133 | 160 | 5760 | 8192..=8202 | 8232 | 8233 | 8239 | 8287 | 12288)
}

/// ASCII and fullwidth digits (.NET \d in these labels, VTV2IsDigit).
fn is_digit(character: char) -> bool {
    character.is_ascii_digit() || ('\u{FF10}'..='\u{FF19}').contains(&character)
}

fn ascii_digit(character: char) -> char {
    if ('\u{FF10}'..='\u{FF19}').contains(&character) {
        char::from_u32(character as u32 - 0xFF10 + '0' as u32).unwrap_or(character)
    } else {
        character
    }
}

fn trim_spaces(value: &str) -> &str {
    value.trim_matches(is_space)
}

fn is_word_character(character: Option<char>) -> bool {
    match character {
        None => false,
        Some(c) => c.is_ascii_alphanumeric() || c == '_' || c as u32 > 127,
    }
}

/// VTV2RemoveWholeWord: every case-insensitive whole-word occurrence removed.
fn remove_whole_word(text: &str, word: &str) -> String {
    let mut chars: Vec<char> = text.chars().collect();
    let target: Vec<char> = word.chars().collect();
    'outer: loop {
        let lowered: Vec<char> = chars.iter().map(|c| c.to_ascii_lowercase()).collect();
        let mut index = 0;
        while index + target.len() <= lowered.len() {
            if lowered[index..index + target.len()] == target[..] {
                let before = if index > 0 { Some(chars[index - 1]) } else { None };
                let after = chars.get(index + target.len()).copied();
                if !is_word_character(before) && !is_word_character(after) {
                    chars.drain(index..index + target.len());
                    continue 'outer;
                }
            }
            index += 1;
        }
        break;
    }
    chars.into_iter().collect()
}

/// The CJK chapter words and the punctuation a heading label may carry
/// (VTV2HeadingDecorations).
const HEADING_DECORATIONS: &str = "\u{7B2C}\u{7AE0}\u{8282}\u{7BC7}\u{90E8}\u{5377}.\u{FF0E}\u{3001},:\u{FF1A};\u{FF1B}-\u{2013}\u{2014}()\u{FF08}\u{FF09}[]\u{3010}\u{3011}{}";
/// Separators between the components of a typed heading number.
const HEADING_NUMBER_SEPARATORS: &str = ".\u{FF0E}-\u{2013}\u{2014}";
/// What may follow a typed heading number (VTV2HeadingNumberTerminators).
const HEADING_NUMBER_TERMINATORS: &str = "\u{7AE0}\u{8282}\u{7BC7}\u{90E8}.\u{FF0E}\u{3001}:\u{FF1A})\u{FF09}-\u{2013}\u{2014}";

fn max_components(outline_level: u32) -> usize {
    outline_level.max(1) as usize
}

/// VTV2HeadingListNumber (Windows NormalizeHeadingNumberText): a list label
/// such as "3.2", "Chapter 3" or the CJK chapter form becomes "3" / "3.2";
/// any other text gives "".
pub(crate) fn heading_list_number(value: &str, outline_level: u32) -> String {
    let cleaned: String = value.chars().filter(|c| !matches!(c, '\t' | '\r' | '\n' | '\u{7}')).collect();
    let text = trim_spaces(&cleaned);
    if text.is_empty() {
        return String::new();
    }
    let mut numbers = String::new();
    let mut current = String::new();
    let mut remainder = String::new();
    let mut count = 0;
    for character in text.chars() {
        if is_digit(character) {
            current.push(ascii_digit(character));
        } else {
            if !current.is_empty() {
                if !numbers.is_empty() {
                    numbers.push('.');
                }
                numbers.push_str(&current);
                count += 1;
                current.clear();
            }
            remainder.push(character);
        }
    }
    if !current.is_empty() {
        if !numbers.is_empty() {
            numbers.push('.');
        }
        numbers.push_str(&current);
        count += 1;
    }
    if count == 0 || count > max_components(outline_level) {
        return String::new();
    }
    for word in ["chapter", "section", "part"] {
        remainder = remove_whole_word(&remainder, word);
    }
    if remainder.chars().any(|c| !HEADING_DECORATIONS.contains(c) && !is_space(c)) {
        return String::new();
    }
    numbers
}

/// VTV2HeadingTextNumber (Windows ParseHeadingNumberFromText): typed headings
/// such as "8. Title", "8-1 Title" or the CJK "chapter 8" form.
pub(crate) fn heading_text_number(value: &str, outline_level: u32) -> String {
    let cleaned: String = value.chars().filter(|c| !matches!(c, '\r' | '\n' | '\u{7}')).collect();
    let text: Vec<char> = trim_spaces(&cleaned).chars().collect();
    if text.is_empty() {
        return String::new();
    }
    let skip_spaces = |mut index: usize| {
        while index < text.len() && is_space(text[index]) {
            index += 1;
        }
        index
    };
    let digits_end = |mut index: usize| {
        while index < text.len() && is_digit(text[index]) {
            index += 1;
        }
        index
    };
    let mut position = 0;
    if text[0] == '\u{7B2C}' {
        position = skip_spaces(1);
    }
    let end = digits_end(position);
    if end == position {
        return String::new();
    }
    let mut number: String = text[position..end].iter().map(|c| ascii_digit(*c)).collect();
    let mut boundaries = vec![end];
    let mut numbers = vec![number.clone()];
    let mut probe = end;
    loop {
        let mut after = skip_spaces(probe);
        if after >= text.len() || !HEADING_NUMBER_SEPARATORS.contains(text[after]) {
            break;
        }
        after = skip_spaces(after + 1);
        let end = digits_end(after);
        if end == after {
            break;
        }
        number.push('.');
        number.extend(text[after..end].iter().map(|c| ascii_digit(*c)));
        boundaries.push(end);
        numbers.push(number.clone());
        probe = end;
    }
    for index in (0..boundaries.len()).rev() {
        let after = skip_spaces(boundaries[index]);
        let spaces = after - boundaries[index];
        let terminated = after < text.len() && HEADING_NUMBER_TERMINATORS.contains(text[after]);
        if terminated || spaces > 0 {
            let count = numbers[index].split('.').count();
            if count == 0 || count > max_components(outline_level) {
                return String::new();
            }
            return numbers[index].clone();
        }
    }
    String::new()
}

/// A heading as the scope rules read it: start, level, explicit number, list
/// numbering (VTV2CollectHeadings item).
#[derive(Debug, Clone)]
struct Heading {
    start: i64,
    level: u32,
    explicit: String,
    list_numbered: bool,
}

fn headings_up_to(facts: &[HeadingFact], target_level: u32) -> Vec<Heading> {
    let mut headings: Vec<Heading> = facts
        .iter()
        .filter(|heading| heading.level >= 1 && heading.level <= target_level)
        .map(|heading| {
            let mut explicit = heading_list_number(&heading.list_string, heading.level);
            if explicit.is_empty() {
                explicit = heading_text_number(&heading.text, heading.level);
            }
            Heading { start: heading.start, level: heading.level, explicit, list_numbered: heading.list_numbered }
        })
        .collect();
    headings.sort_by_key(|heading| heading.start);
    headings
}

fn pad_to_level(number: &str, level: u32, target_level: u32) -> String {
    let mut text = number.to_string();
    for _ in level..target_level {
        text.push_str(".0");
    }
    text
}

/// VTV2HeadingScope (Windows ResolveHeadingScopeAtPosition): the chapter or
/// section number at position ("3", "3.2", or "0" / "0.0" before any heading),
/// the start of the heading that gives it (None when none does) and whether
/// that heading has Word list numbering (read live with STYLEREF).
fn heading_scope(headings: &[Heading], position: i64, target_level: u32) -> (String, Option<i64>, bool) {
    if target_level < 1 {
        return (String::new(), None, false);
    }
    // The last numbered heading of each level up to the formula, each deeper
    // level only inside the scope chosen for the level above it.
    let mut chosen: Option<usize> = None;
    let mut number_text = String::new();
    for level in 1..=target_level {
        for index in (0..headings.len()).rev() {
            let heading = &headings[index];
            if heading.level != level || heading.start > position {
                continue;
            }
            if let Some(chosen) = chosen {
                if heading.start <= headings[chosen].start {
                    break;
                }
            }
            if !heading.explicit.is_empty() {
                chosen = Some(index);
                number_text = pad_to_level(&heading.explicit, level, target_level);
                break;
            }
        }
    }
    if let Some(chosen) = chosen {
        return (number_text, Some(headings[chosen].start), headings[chosen].list_numbered);
    }
    // Count unnumbered headings only in a document without explicit numbers.
    let uses_explicit = headings.iter().any(|heading| !heading.explicit.is_empty());
    let mut counters = [0u32; 10];
    let mut anchor: Option<(usize, String)> = None;
    for (index, heading) in headings.iter().enumerate() {
        let mut text = heading.explicit.clone();
        if text.is_empty() {
            if uses_explicit {
                continue;
            }
            let level = heading.level as usize;
            counters[level] += 1;
            for deeper in counters.iter_mut().skip(level + 1) {
                *deeper = 0;
            }
            text = (1..=level)
                .map(|deeper| counters[deeper].max(1).to_string())
                .collect::<Vec<_>>()
                .join(".");
        }
        text = pad_to_level(&text, heading.level, target_level);
        if heading.start > position {
            break;
        }
        anchor = Some((index, text));
    }
    match anchor {
        Some((index, text)) => (text, Some(headings[index].start), headings[index].list_numbered),
        None => ((if target_level == 1 { "0" } else { "0.0" }).to_string(), None, false),
    }
}

// ---- equation numbers -----------------------------------------------------------

/// VTV2ParseNumberLabel: "2.1-3" -> ("2.1", "-", 3); "3" -> ("", "", 3).
pub(crate) fn parse_number_label(text: &str) -> Option<(String, String, u32)> {
    let chars: Vec<char> = text.chars().collect();
    let mut index = chars.len();
    while index > 0 && chars[index - 1].is_ascii_digit() {
        index -= 1;
    }
    let ordinal_text: String = chars[index..].iter().collect();
    if ordinal_text.is_empty() || ordinal_text.len() > 9 {
        return None;
    }
    if ordinal_text.len() > 1 && ordinal_text.starts_with('0') {
        return None;
    }
    let ordinal: u32 = ordinal_text.parse().ok()?;
    if ordinal < 1 {
        return None;
    }
    if index == 0 {
        return Some((String::new(), String::new(), ordinal));
    }
    let separator = chars[index - 1];
    if !".-\u{2010}\u{2011}\u{2012}\u{2013}\u{2212}".contains(separator) {
        return None;
    }
    let candidate: String = chars[..index - 1].iter().collect();
    if candidate.is_empty() {
        return None;
    }
    let mut parts = candidate.split('.');
    let first = parts.next().unwrap_or_default();
    if first.is_empty() {
        return None;
    }
    if !first.chars().all(|c| c.is_ascii_digit()) && !first.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    for part in parts {
        if part.is_empty() || !part.chars().all(|c| c.is_ascii_digit()) {
            return None;
        }
    }
    Some((candidate, separator.to_string(), ordinal))
}

/// VTV2SequenceName: "SEQ Equation \r 1 ..." -> Equation; SEQ "My Eq" -> My Eq.
pub(crate) fn sequence_name(instruction: &str) -> String {
    let text = trim_spaces(instruction);
    if text.len() < 4 || !text[..4].eq_ignore_ascii_case("SEQ ") {
        return String::new();
    }
    let text = trim_spaces(&text[4..]);
    if let Some(rest) = text.strip_prefix('"') {
        return match rest.find('"') {
            Some(closing) if closing > 0 => rest[..closing].to_string(),
            _ => String::new(),
        };
    }
    text.chars().take_while(|c| !is_space(*c) && *c != '\\').collect()
}

/// VTV2RestartValue: the N of a "\r N" switch, or 0.
pub(crate) fn restart_value(instruction: &str) -> u32 {
    let lowered = instruction.to_ascii_lowercase();
    let Some(position) = lowered.find("\\r ") else { return 0 };
    let digits: String = instruction[position + 3..]
        .trim_start_matches(' ')
        .chars()
        .take_while(|c| c.is_ascii_digit())
        .collect();
    if !digits.is_empty() && digits.len() < 9 {
        digits.parse().unwrap_or(0)
    } else {
        0
    }
}

/// VTParseEquationNumberPrefix: the style a visible number prefix shows: none
/// -> sequence, "2." -> chapter, "2.1-" -> section with a dash.
fn style_from_prefix(prefix: &str) -> Option<NumberingStyle> {
    let cleaned: String = prefix
        .chars()
        .filter_map(|c| match c {
            '0'..='9' | '.' => Some(c),
            '-' | '\u{2010}' | '\u{2011}' | '\u{2013}' | '\u{2212}' => Some('-'),
            _ => None,
        })
        .collect();
    if cleaned.is_empty() {
        return Some(NumberingStyle::sequence());
    }
    let separator = cleaned.chars().last()?;
    if separator != '.' && separator != '-' {
        return None;
    }
    let body = &cleaned[..cleaned.len() - 1];
    if body.is_empty() || body.contains('-') {
        return None;
    }
    let mode = if body.contains('.') { MODE_SECTION } else { MODE_CHAPTER };
    Some(NumberingStyle { mode: mode.to_string(), separator: separator.to_string() })
}

/// The visible prefix of a number: the text before its ordinal digits.
fn number_prefix(number_text: &str) -> &str {
    number_text.trim_end_matches(|c: char| c.is_ascii_digit())
}

/// VTResolveEquationNumberingUncached for a WPS document: the earlier
/// versions' document variables, else the style the document's first numbered
/// equation shows (Word reads image numbers through their VT_R_ bookmark, whose
/// text never ends in a digit, so pictures do not decide it there either),
/// else the user's default ("mode|separator", VisualTeXNumberingPreference.txt),
/// else continuous numbering.
pub(crate) fn numbering_style(facts: &NumberingFacts, preference: Option<&str>) -> NumberingStyle {
    if let Some(mode) = facts.mode_variable.as_deref().map(|mode| mode.trim().to_ascii_lowercase()) {
        if [MODE_SEQUENCE, MODE_CHAPTER, MODE_SECTION].contains(&mode.as_str()) {
            let separator = if mode != MODE_SEQUENCE
                && facts.separator_variable.as_deref().map(str::trim) == Some("-")
            {
                "-"
            } else {
                "."
            };
            return NumberingStyle { mode, separator: separator.to_string() };
        }
    }
    let first = facts
        .equations
        .iter()
        .filter(|equation| equation.number_text.trim().ends_with(|c: char| c.is_ascii_digit()))
        .min_by_key(|equation| equation.start)
        .map(|equation| number_prefix(equation.number_text.trim()));
    if let Some(style) = first.and_then(style_from_prefix) {
        return style;
    }
    if let Some(line) = preference {
        let mut fields = line.trim().split('|');
        if let (Some(mode), Some(separator), None) = (fields.next(), fields.next(), fields.next()) {
            let mode = mode.trim().to_ascii_lowercase();
            let separator = separator.trim();
            match mode.as_str() {
                MODE_SEQUENCE => return NumberingStyle::sequence(),
                MODE_CHAPTER | MODE_SECTION if separator == "." || separator == "-" => {
                    return NumberingStyle { mode, separator: separator.to_string() };
                }
                _ => {}
            }
        }
    }
    NumberingStyle::sequence()
}

/// A numbered equation as the plan reads it (VTV2NumberedEquationEntry).
struct Entry {
    start: i64,
    sequence_name: String,
    prefix: String,
    ordinal: u32,
    restart: u32,
}

fn equation_entries(facts: &NumberingFacts) -> Vec<Entry> {
    let mut entries: Vec<Entry> = facts
        .equations
        .iter()
        .filter_map(|equation| {
            let (prefix, _, ordinal) = parse_number_label(equation.number_text.trim())?;
            Some(Entry {
                start: equation.start,
                sequence_name: sequence_name(&equation.instruction),
                prefix,
                ordinal,
                restart: restart_value(&equation.instruction),
            })
        })
        .collect();
    entries.sort_by_key(|entry| entry.start);
    entries
}

/// VTV2PlanInsertion: the number a new display equation at the position takes.
pub(crate) fn plan_equation(facts: &NumberingFacts, style: &NumberingStyle) -> EquationPlan {
    let position = facts.position;
    let (heading_level, separator) = style.level_and_separator();
    let entries = equation_entries(facts);
    let mut plan = EquationPlan {
        sequence_name: DEFAULT_EQUATION_LABEL.to_string(),
        heading_level,
        separator,
        literal_prefix: None,
        automatic: false,
        restart: None,
        initial: 1,
        transfer_from: None,
    };
    if !entries.is_empty() {
        let nearest = entries.iter().rposition(|entry| entry.start < position).unwrap_or(0);
        let mut first = nearest;
        while first > 0
            && entries[first - 1].prefix == entries[nearest].prefix
            && entries[first - 1].sequence_name.eq_ignore_ascii_case(&entries[nearest].sequence_name)
        {
            first -= 1;
        }
        plan.sequence_name = entries[first].sequence_name.clone();
        if position <= entries[first].start {
            plan.restart = Some(entries[first].ordinal).filter(|value| *value > 0);
            if entries[first].restart > 0 {
                plan.transfer_from = Some(entries[first].start);
            }
            plan.initial = entries[first].ordinal;
        } else {
            plan.initial = entries[first].ordinal + 1;
        }
    }
    if heading_level > 0 {
        let headings = headings_up_to(&facts.headings, heading_level);
        let (number_text, scope_start, automatic) = heading_scope(&headings, position, heading_level);
        plan.automatic = automatic && scope_start.is_some();
        if !plan.automatic {
            plan.literal_prefix = Some(number_text);
        }
    }
    plan
}

/// The display separator of image numbers (VTEquationNumberingDisplaySeparator):
/// U+2010 HYPHEN for a dash, "." otherwise.
fn image_display_separator(style: &NumberingStyle) -> &'static str {
    if style.separator == "-" {
        "\u{2010}"
    } else {
        "."
    }
}

/// VTEquationStartsNumberScope for an image at `position`: in chapter or
/// section numbering, true when no numbered image comes before it or a
/// heading of that level lies between the previous one and it.
fn image_starts_scope(previous: Option<i64>, position: i64, headings: &[Heading], target_level: u32) -> bool {
    if target_level < 1 {
        return false;
    }
    match previous {
        None => true,
        Some(previous) => headings
            .iter()
            .any(|heading| heading.level <= target_level && heading.start > previous && heading.start < position),
    }
}

/// The number of a new numbered image formula at the position
/// (VTWriteSingleParagraphImageNumber and VTReconcileNumberingAfterEdit).
pub(crate) fn plan_image(facts: &NumberingFacts, style: &NumberingStyle) -> ImagePlan {
    let position = facts.position;
    let (level, _) = style.level_and_separator();
    let headings = headings_up_to(&facts.headings, level.max(1));
    let mut prefix_text = String::new();
    if level > 0 {
        let (number_text, _, _) = heading_scope(&headings, position, level);
        prefix_text = format!("{number_text}{}", image_display_separator(style));
    }
    let mut images: Vec<&ImageFact> = facts.images.iter().collect();
    images.sort_by_key(|image| image.start);
    let previous = images.iter().filter(|image| image.start < position).map(|image| image.start).last();
    let restart_at_one = image_starts_scope(previous, position, &headings, level);
    // The next numbered image now follows the new one: its restart may change.
    let next_restart = if level > 0 {
        images.iter().find(|image| image.start >= position).and_then(|next| {
            let wanted = image_starts_scope(Some(position), next.start, &headings, level);
            let has = restart_value(&next.increment_instruction) > 0;
            (wanted != has).then_some(NextImageRestart { start: next.start, restart_at_one: wanted })
        })
    } else {
        None
    };
    ImagePlan { prefix_text, restart_at_one, next_restart }
}

/// VTImageSequenceInstruction: the hidden increment ("SEQ VisualTeXEquation
/// [\r 1] \h") and the visible current value ("SEQ VisualTeXEquation \c \*
/// ARABIC") of an image number.
pub(crate) fn image_increment_instruction(restart_at_one: bool) -> String {
    if restart_at_one {
        format!("SEQ {IMAGE_SEQUENCE_NAME} \\r 1 \\h")
    } else {
        format!("SEQ {IMAGE_SEQUENCE_NAME} \\h")
    }
}

pub(crate) fn image_current_instruction() -> String {
    format!("SEQ {IMAGE_SEQUENCE_NAME} \\c \\* ARABIC")
}

// ---- every number in the document's style (VTUpdateEquationNumbersCore) ------
//
// "更新公式编号" and "编号格式" rewrite each number that is not already in the
// document's numbering style: VTV2RestyleNumbers for numbered equations and
// VTUpgradeWordNumberFormats (VTBuildEquationHeadingPrefixes,
// VTImagePlaceRefIsCanonical) for numbered pictures.

/// A numbered display equation as VTV2NumberedEquationEntry reads it.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct NumberedEquationFact {
    pub start: i64,
    /// The SEQ field's instruction, with "*" where WPS shows U+2217.
    pub instruction: String,
    /// The visible number ("2.1-3"); a number WPS damaged shows a field error.
    pub number_text: String,
    /// The instructions of the STYLEREF fields in the number.
    #[serde(default)]
    pub style_refs: Vec<String>,
    /// Fields in the number other than SEQ and STYLEREF.
    #[serde(default)]
    pub other_fields: u32,
    /// A display equation "批量编号" numbers now: it has no number yet.
    #[serde(default)]
    pub unnumbered: bool,
    /// An older VisualTeX number: a REF to a hidden SEQ elsewhere
    /// (VTNativeEquationArrayReferenceField). It is written anew in the
    /// document's style (VTWritePureNativeNumber).
    #[serde(default)]
    pub legacy: bool,
}

/// A numbered picture's number as VTImagePlaceRefIsCanonical reads it.
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct NumberedImageFact {
    pub start: i64,
    #[serde(default)]
    pub increment_instruction: String,
    #[serde(default)]
    pub current_instruction: String,
    /// The text between "(" and the current-value SEQ field.
    #[serde(default)]
    pub prefix_text: String,
    /// One hidden increment and one current-value SEQ, no other field, and
    /// ")" closing the MACROBUTTON's text.
    #[serde(default)]
    pub well_formed: bool,
    /// A display picture "批量编号" numbers now.
    #[serde(default)]
    pub unnumbered: bool,
}

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RestyleFacts {
    #[serde(default)]
    pub headings: Vec<HeadingFact>,
    #[serde(default)]
    pub equations: Vec<NumberedEquationFact>,
    #[serde(default)]
    pub images: Vec<NumberedImageFact>,
    #[serde(default)]
    pub mode_variable: Option<String>,
    #[serde(default)]
    pub separator_variable: Option<String>,
}

impl RestyleFacts {
    /// The facts numbering_style reads.
    pub(crate) fn style_facts(&self) -> NumberingFacts {
        NumberingFacts {
            position: 0,
            headings: self.headings.clone(),
            equations: self
                .equations
                .iter()
                .map(|equation| EquationFact {
                    start: equation.start,
                    instruction: equation.instruction.clone(),
                    number_text: equation.number_text.clone(),
                })
                .collect(),
            images: Vec::new(),
            mode_variable: self.mode_variable.clone(),
            separator_variable: self.separator_variable.clone(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct EquationRestyle {
    pub start: i64,
    /// The number is not in the document's style and is written again.
    pub changed: bool,
    pub plan: EquationPlan,
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ImageRestyle {
    pub start: i64,
    pub changed: bool,
    pub prefix_text: String,
    pub restart_at_one: bool,
    pub increment: String,
    pub current: String,
}

fn contains_ignore_case(text: &str, needle: &str) -> bool {
    text.to_ascii_lowercase().contains(&needle.to_ascii_lowercase())
}

/// VTV2SequenceMatches: the SEQ instruction has the label, restart and \s
/// level asked for, with Word's ARABIC and MERGEFORMAT switches.
fn sequence_matches(instruction: &str, name: &str, restart: u32, heading_level: u32) -> bool {
    if !sequence_name(instruction).eq_ignore_ascii_case(name) || restart_value(instruction) != restart {
        return false;
    }
    let lowered = instruction.to_ascii_lowercase();
    match lowered.find("\\s ") {
        Some(position) => {
            let level: String = instruction[position + 3..].chars().take(2).collect();
            if heading_level == 0 || !level.trim().starts_with(&heading_level.to_string()) {
                return false;
            }
        }
        None if heading_level > 0 => return false,
        None => {}
    }
    contains_ignore_case(instruction, "\\* ARABIC") && contains_ignore_case(instruction, "\\* MERGEFORMAT")
}

/// VTV2NumberHasStyle: the number already has exactly the requested form.
fn number_has_style(
    equation: &NumberedEquationFact,
    name: &str,
    heading_level: u32,
    separator: &str,
    literal_prefix: &str,
    automatic: bool,
    restart: u32,
) -> bool {
    if !sequence_matches(&equation.instruction, name, restart, heading_level) {
        return false;
    }
    let style_ref = format!("STYLEREF {heading_level} ");
    if equation.style_refs.iter().any(|code| !contains_ignore_case(code, &style_ref)) || equation.other_fields > 0 {
        return false;
    }
    let (prefix, number_separator, _) = parse_number_label(equation.number_text.trim()).unwrap_or_default();
    let style_refs = equation.style_refs.len();
    if heading_level == 0 {
        style_refs == 0 && prefix.is_empty()
    } else if automatic {
        style_refs == 1 && number_separator == separator
    } else {
        style_refs == 0 && prefix == literal_prefix && number_separator == separator
    }
}

/// VTV2RestyleNumbers: the number each numbered equation takes in the style
/// (its own label, the chapter or section prefix, \r 1 on the first number of
/// each chapter/section or of the document) and whether it must be written.
pub(crate) fn restyle_equations(facts: &RestyleFacts, style: &NumberingStyle) -> Vec<EquationRestyle> {
    let (heading_level, separator) = style.level_and_separator();
    let headings = headings_up_to(&facts.headings, heading_level.max(1));
    let mut equations: Vec<&NumberedEquationFact> = facts.equations.iter().collect();
    equations.sort_by_key(|equation| equation.start);
    // A new number takes the label the document's numbers already use.
    let document_label = equations
        .iter()
        .filter(|equation| !equation.unnumbered && !equation.legacy)
        .map(|equation| sequence_name(&equation.instruction))
        .find(|name| !name.is_empty())
        .unwrap_or_else(|| DEFAULT_EQUATION_LABEL.to_string());
    let mut counters: HashMap<String, u32> = HashMap::new();
    let mut result = Vec::new();
    for equation in equations {
        let name = if equation.unnumbered || equation.legacy {
            document_label.clone()
        } else {
            sequence_name(&equation.instruction)
        };
        if name.is_empty() {
            continue;
        }
        let mut literal_prefix = String::new();
        let mut automatic = false;
        let mut scope = "continuous".to_string();
        if heading_level > 0 {
            let (number_text, scope_start, list_numbered) = heading_scope(&headings, equation.start, heading_level);
            automatic = list_numbered && scope_start.is_some();
            if !automatic {
                literal_prefix = number_text;
            }
            scope = scope_start.unwrap_or(-1).to_string();
        }
        let counter = counters.entry(format!("{}|{scope}", name.to_lowercase())).or_insert(0);
        *counter += 1;
        let ordinal = *counter;
        let restart = u32::from(ordinal == 1);
        let changed = equation.unnumbered
            || equation.legacy
            || !number_has_style(equation, &name, heading_level, &separator, &literal_prefix, automatic, restart);
        let plan = EquationPlan {
            sequence_name: name,
            heading_level,
            separator: separator.clone(),
            literal_prefix: (heading_level > 0 && !automatic).then_some(literal_prefix),
            automatic,
            restart: (restart > 0).then_some(1),
            initial: ordinal,
            transfer_from: None,
        };
        result.push(EquationRestyle { start: equation.start, changed, plan });
    }
    result
}

fn has_switch(instruction: &str, letter: char) -> bool {
    let switch = format!("\\{letter}");
    instruction.split(is_space).any(|token| token.eq_ignore_ascii_case(&switch))
}

/// VTEquationSequenceFieldIsCompatible: the switches that change what a
/// picture number's SEQ field counts are the ones asked for.
fn image_sequence_compatible(instruction: &str, hidden: bool, current: bool, restart_at_one: bool) -> bool {
    if !sequence_name(instruction).eq_ignore_ascii_case(IMAGE_SEQUENCE_NAME)
        || has_switch(instruction, 'h') != hidden
        || has_switch(instruction, 'c') != current
        || has_switch(instruction, 's')
    {
        return false;
    }
    let explicit_restart = has_switch(instruction, 'r');
    if current {
        !explicit_restart
    } else if restart_at_one {
        explicit_restart && restart_value(instruction) == 1
    } else {
        !explicit_restart
    }
}

/// VTUpgradeWordNumberFormats for pictures: each number's chapter or section
/// prefix and restart (VTBuildEquationHeadingPrefixes: the running number
/// starts again after every heading of the numbering level), and whether its
/// place reference must be written again (VTImagePlaceRefIsCanonical).
pub(crate) fn restyle_images(facts: &RestyleFacts, style: &NumberingStyle) -> Vec<ImageRestyle> {
    let (level, _) = style.level_and_separator();
    let headings = headings_up_to(&facts.headings, level.max(1));
    let mut images: Vec<&NumberedImageFact> = facts.images.iter().collect();
    images.sort_by_key(|image| image.start);
    let mut previous: Option<i64> = None;
    let mut result = Vec::new();
    for image in images {
        let (prefix_text, restart_at_one) = if level > 0 {
            let (number_text, _, _) = heading_scope(&headings, image.start, level);
            (
                format!("{number_text}{}", image_display_separator(style)),
                image_starts_scope(previous, image.start, &headings, level),
            )
        } else {
            (String::new(), false)
        };
        previous = Some(image.start);
        let canonical = !image.unnumbered
            && image.well_formed
            && image_sequence_compatible(&image.increment_instruction, true, false, restart_at_one)
            && image_sequence_compatible(&image.current_instruction, false, true, false)
            && image.prefix_text == prefix_text;
        result.push(ImageRestyle {
            start: image.start,
            changed: !canonical,
            increment: image_increment_instruction(restart_at_one),
            current: image_current_instruction(),
            prefix_text,
            restart_at_one,
        });
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    fn heading(start: i64, level: u32, list: &str, numbered: bool, text: &str) -> HeadingFact {
        HeadingFact { start, level, list_string: list.to_string(), list_numbered: numbered, text: text.to_string() }
    }

    fn equation(start: i64, instruction: &str, number: &str) -> EquationFact {
        EquationFact { start, instruction: instruction.to_string(), number_text: number.to_string() }
    }

    #[test]
    fn number_labels_parse_like_word() {
        assert_eq!(parse_number_label("3"), Some((String::new(), String::new(), 3)));
        assert_eq!(parse_number_label("2.1-3"), Some(("2.1".into(), "-".into(), 3)));
        assert_eq!(parse_number_label("0.0\u{2010}1"), Some(("0.0".into(), "\u{2010}".into(), 1)));
        assert_eq!(parse_number_label("A.2"), Some(("A".into(), ".".into(), 2)));
        assert_eq!(parse_number_label("03"), None);
        assert_eq!(parse_number_label("0"), None);
        assert_eq!(parse_number_label("2.x-3"), None);
        assert_eq!(parse_number_label("Error! Bookmark not defined."), None);
    }

    #[test]
    fn sequence_names_and_restarts_come_from_the_instruction() {
        assert_eq!(sequence_name(" SEQ Equation \\r 1 \\s 2 \\* ARABIC \\* MERGEFORMAT "), "Equation");
        assert_eq!(sequence_name("SEQ \"My Eq\" \\* ARABIC"), "My Eq");
        assert_eq!(restart_value(" SEQ Equation \\r 1 \\s 2 \\* ARABIC "), 1);
        assert_eq!(restart_value(" SEQ Equation \\s 2 \\* ARABIC "), 0);
        // WPS spells the ARABIC switch with U+2217; nothing here depends on it.
        assert_eq!(sequence_name(" SEQ Equation \\s 2 \\\u{2217} ARABIC "), "Equation");
    }

    #[test]
    fn heading_numbers_come_from_list_labels_and_typed_text() {
        assert_eq!(heading_list_number("3.2", 2), "3.2");
        assert_eq!(heading_list_number("Chapter 3", 1), "3");
        assert_eq!(heading_list_number("\u{7B2C}3\u{7AE0}", 1), "3");
        assert_eq!(heading_list_number("3.2.1", 2), "");
        assert_eq!(heading_list_number("Appendix 3", 1), "");
        assert_eq!(heading_text_number("8. Title", 1), "8");
        assert_eq!(heading_text_number("8-1 Title", 2), "8.1");
        assert_eq!(heading_text_number("\u{7B2C}8\u{7AE0} Title", 1), "8");
        assert_eq!(heading_text_number("Title 8", 1), "");
        assert_eq!(heading_text_number("2020 results", 1), "2020");
    }

    #[test]
    fn the_style_is_read_from_the_first_number_then_the_preference() {
        let mut facts = NumberingFacts { position: 0, ..Default::default() };
        assert_eq!(numbering_style(&facts, None), NumberingStyle::sequence());
        assert_eq!(
            numbering_style(&facts, Some("section|-")),
            NumberingStyle { mode: MODE_SECTION.into(), separator: "-".into() }
        );
        // Numbered pictures do not decide the style (Word reads none from them).
        facts.images.push(ImageFact { start: 10, increment_instruction: image_increment_instruction(true) });
        assert_eq!(numbering_style(&facts, Some("chapter|.")).mode, MODE_CHAPTER);
        facts.equations.push(equation(50, " SEQ Equation \\s 1 \\* ARABIC ", "2.4"));
        assert_eq!(
            numbering_style(&facts, Some("section|-")),
            NumberingStyle { mode: MODE_CHAPTER.into(), separator: ".".into() }
        );
        facts.mode_variable = Some("sequence".into());
        assert_eq!(numbering_style(&facts, None), NumberingStyle::sequence());
    }

    #[test]
    fn a_first_equation_in_a_continuous_document() {
        let facts = NumberingFacts { position: 100, ..Default::default() };
        let plan = plan_equation(&facts, &NumberingStyle::sequence());
        assert_eq!(plan, EquationPlan {
            sequence_name: "Equation".into(),
            heading_level: 0,
            separator: String::new(),
            literal_prefix: None,
            automatic: false,
            restart: None,
            initial: 1,
            transfer_from: None,
        });
    }

    #[test]
    fn an_equation_before_the_first_number_takes_over_its_restart() {
        let facts = NumberingFacts {
            position: 10,
            equations: vec![
                equation(100, " SEQ Equation \\r 1 \\* ARABIC \\* MERGEFORMAT ", "1"),
                equation(200, " SEQ Equation \\* ARABIC \\* MERGEFORMAT ", "2"),
            ],
            ..Default::default()
        };
        let plan = plan_equation(&facts, &NumberingStyle::sequence());
        assert_eq!(plan.restart, Some(1));
        assert_eq!(plan.transfer_from, Some(100));
        assert_eq!(plan.initial, 1);
        let later = NumberingFacts { position: 150, ..facts };
        let plan = plan_equation(&later, &NumberingStyle::sequence());
        assert_eq!((plan.restart, plan.transfer_from, plan.initial), (None, None, 2));
    }

    #[test]
    fn section_numbers_use_the_heading_scope() {
        let style = NumberingStyle { mode: MODE_SECTION.into(), separator: "-".into() };
        // Before any heading: the literal "0.0".
        let facts = NumberingFacts { position: 5, headings: vec![heading(50, 1, "", false, "Intro")], ..Default::default() };
        let plan = plan_equation(&facts, &style);
        assert_eq!((plan.heading_level, plan.separator.as_str(), plan.literal_prefix.as_deref()), (2, "-", Some("0.0")));
        // Unnumbered headings are counted: chapter 2, no section -> "2.0".
        let facts = NumberingFacts {
            position: 300,
            headings: vec![heading(50, 1, "", false, "Intro"), heading(200, 1, "", false, "Methods")],
            ..Default::default()
        };
        assert_eq!(plan_equation(&facts, &style).literal_prefix.as_deref(), Some("2.0"));
        // A list-numbered heading is read live (STYLEREF).
        let facts = NumberingFacts {
            position: 300,
            headings: vec![heading(50, 1, "1", true, "Intro"), heading(200, 2, "1.1", true, "Part")],
            ..Default::default()
        };
        let plan = plan_equation(&facts, &style);
        assert!(plan.automatic);
        assert_eq!(plan.literal_prefix, None);
        // Typed heading numbers are literal.
        let facts = NumberingFacts {
            position: 300,
            headings: vec![heading(50, 1, "", false, "3 Results"), heading(200, 2, "", false, "3.4 Data")],
            ..Default::default()
        };
        assert_eq!(plan_equation(&facts, &style).literal_prefix.as_deref(), Some("3.4"));
    }

    #[test]
    fn image_numbers_restart_per_chapter_and_fix_the_next_one() {
        let style = NumberingStyle { mode: MODE_CHAPTER.into(), separator: "-".into() };
        let image = |start: i64, restart: bool| ImageFact { start, increment_instruction: image_increment_instruction(restart) };
        // First image of the document: restart, prefix "0" + U+2010.
        let facts = NumberingFacts { position: 10, ..Default::default() };
        let plan = plan_image(&facts, &style);
        assert_eq!(plan, ImagePlan { prefix_text: "0\u{2010}".into(), restart_at_one: true, next_restart: None });
        // Before the first image of a chapter: it restarts, the old first stops.
        let facts = NumberingFacts {
            position: 120,
            headings: vec![heading(100, 1, "", false, "One")],
            images: vec![image(150, true)],
            ..Default::default()
        };
        let plan = plan_image(&facts, &style);
        assert_eq!(plan.prefix_text, "1\u{2010}");
        assert!(plan.restart_at_one);
        assert_eq!(plan.next_restart, Some(NextImageRestart { start: 150, restart_at_one: false }));
        // Continuous numbering never restarts.
        let plan = plan_image(&facts, &NumberingStyle::sequence());
        assert_eq!(plan, ImagePlan { prefix_text: String::new(), restart_at_one: false, next_restart: None });
        assert_eq!(image_current_instruction(), "SEQ VisualTeXEquation \\c \\* ARABIC");
    }

    fn numbered(start: i64, instruction: &str, number: &str) -> NumberedEquationFact {
        NumberedEquationFact {
            start,
            instruction: instruction.to_string(),
            number_text: number.to_string(),
            style_refs: Vec::new(),
            other_fields: 0,
            unnumbered: false,
            legacy: false,
        }
    }

    #[test]
    fn restyle_writes_older_reference_numbers_anew_with_the_document_label() {
        let chapter = NumberingStyle { mode: MODE_CHAPTER.into(), separator: "-".into() };
        let mut legacy = numbered(300, "", "(1)");
        legacy.legacy = true;
        let facts = RestyleFacts {
            equations: vec![numbered(100, "SEQ Equation \\r 1 \\* ARABIC", "0-1"), legacy],
            ..Default::default()
        };
        let result = restyle_equations(&facts, &chapter);
        assert_eq!(result.len(), 2);
        assert!(result[1].changed);
        assert_eq!(result[1].plan.sequence_name, "Equation");
        assert_eq!(result[1].plan.initial, 2);
    }

    #[test]
    fn restyle_keeps_numbers_in_style_and_rewrites_the_others() {
        let section = NumberingStyle { mode: MODE_SECTION.into(), separator: "-".into() };
        let facts = RestyleFacts {
            headings: vec![heading(100, 1, "", false, "One"), heading(200, 2, "", false, "Sub")],
            equations: vec![
                // In style: first of chapter 1 before its section, \r 1 \s 2.
                numbered(150, " SEQ Equation \\r 1 \\s 2 \\* ARABIC \\* MERGEFORMAT ", "1.0-1"),
                // First of section 1.1 but no \r: rewritten with \r 1.
                numbered(250, " SEQ Equation \\s 2 \\* ARABIC \\* MERGEFORMAT ", "1.1-1"),
                // Second of the section, damaged by WPS (field error): rewritten.
                numbered(300, " SEQ Equation \\s 2 \\* ARABIC \\* MERGEFORMAT ", "Error! Bookmark not defined."),
            ],
            ..Default::default()
        };
        let result = restyle_equations(&facts, &section);
        assert_eq!(result.iter().map(|item| item.changed).collect::<Vec<_>>(), vec![false, true, true]);
        assert_eq!(result[1].plan.literal_prefix.as_deref(), Some("1.1"));
        assert_eq!(result[1].plan.restart, Some(1));
        assert_eq!(result[2].plan.initial, 2);
        assert_eq!(result[2].plan.restart, None);
        // Changing to chapter numbers with a dot rewrites all three.
        let chapter = NumberingStyle { mode: MODE_CHAPTER.into(), separator: ".".into() };
        let result = restyle_equations(&facts, &chapter);
        assert!(result.iter().all(|item| item.changed));
        assert_eq!(
            result.iter().map(|item| (item.plan.literal_prefix.clone(), item.plan.initial)).collect::<Vec<_>>(),
            vec![(Some("1".into()), 1), (Some("1".into()), 2), (Some("1".into()), 3)]
        );
        // Continuous numbers: no prefix, no \s, \r 1 only on the first.
        let result = restyle_equations(&facts, &NumberingStyle::sequence());
        assert_eq!(result[0].plan.literal_prefix, None);
        assert_eq!(result[0].plan.restart, Some(1));
        assert_eq!(result[1].plan.restart, None);
        let in_style = RestyleFacts {
            equations: vec![numbered(10, " SEQ Equation \\r 1 \\* ARABIC \\* MERGEFORMAT ", "1")],
            ..Default::default()
        };
        assert!(!restyle_equations(&in_style, &NumberingStyle::sequence())[0].changed);
    }

    #[test]
    fn batch_numbering_plans_new_numbers_among_the_existing_ones() {
        let section = NumberingStyle { mode: MODE_SECTION.into(), separator: "-".into() };
        let mut new_first = numbered(120, "", "");
        new_first.unnumbered = true;
        let mut new_second = numbered(300, "", "");
        new_second.unnumbered = true;
        let facts = RestyleFacts {
            headings: vec![heading(100, 1, "", false, "One")],
            equations: vec![
                new_first,
                numbered(150, " SEQ Formula \\r 1 \\s 2 \\* ARABIC \\* MERGEFORMAT ", "1.0-1"),
                new_second,
            ],
            ..Default::default()
        };
        let result = restyle_equations(&facts, &section);
        // The new first number takes the restart and the document's label;
        // the old first loses its restart; the new last follows on.
        assert_eq!(result[0].plan.sequence_name, "Formula");
        assert_eq!((result[0].changed, result[0].plan.restart, result[0].plan.initial), (true, Some(1), 1));
        assert_eq!((result[1].changed, result[1].plan.restart, result[1].plan.initial), (true, None, 2));
        assert_eq!((result[2].changed, result[2].plan.initial), (true, 3));
        let image = NumberedImageFact { start: 200, unnumbered: true, ..Default::default() };
        let facts = RestyleFacts { images: vec![image], ..Default::default() };
        let result = restyle_images(&facts, &section);
        assert!(result[0].changed && result[0].restart_at_one);
        assert_eq!(result[0].prefix_text, "0.0\u{2010}");
    }

    #[test]
    fn restyle_reads_automatic_chapter_numbers_from_styleref() {
        let chapter = NumberingStyle { mode: MODE_CHAPTER.into(), separator: "-".into() };
        let mut equation = numbered(150, " SEQ Equation \\r 1 \\s 1 \\* ARABIC \\* MERGEFORMAT ", "2-1");
        equation.style_refs = vec![" STYLEREF 1 \\s ".into()];
        let facts = RestyleFacts {
            headings: vec![heading(100, 1, "2", true, "Two")],
            equations: vec![equation.clone()],
            ..Default::default()
        };
        let result = restyle_equations(&facts, &chapter);
        assert!(!result[0].changed);
        assert!(result[0].plan.automatic);
        assert_eq!(result[0].plan.literal_prefix, None);
        // A STYLEREF of the wrong level is rewritten.
        equation.style_refs = vec![" STYLEREF 2 \\s ".into()];
        let facts = RestyleFacts { equations: vec![equation], ..facts };
        assert!(restyle_equations(&facts, &chapter)[0].changed);
    }

    #[test]
    fn restyle_images_restart_after_headings_and_check_the_place_reference() {
        let style = NumberingStyle { mode: MODE_CHAPTER.into(), separator: "-".into() };
        let image = |start: i64, restart: bool, prefix: &str| NumberedImageFact {
            start,
            increment_instruction: image_increment_instruction(restart),
            current_instruction: image_current_instruction(),
            prefix_text: prefix.to_string(),
            well_formed: true,
            unnumbered: false,
        };
        let facts = RestyleFacts {
            headings: vec![heading(100, 1, "", false, "One"), heading(300, 1, "", false, "Two")],
            images: vec![
                image(50, true, "0\u{2010}"),
                image(150, true, "1\u{2010}"),
                image(200, true, "1\u{2010}"),
                image(350, false, "2\u{2010}"),
            ],
            ..Default::default()
        };
        let result = restyle_images(&facts, &style);
        assert_eq!(result.iter().map(|item| item.restart_at_one).collect::<Vec<_>>(), vec![true, true, false, true]);
        assert_eq!(result.iter().map(|item| item.changed).collect::<Vec<_>>(), vec![false, false, true, true]);
        assert_eq!(result[2].increment, "SEQ VisualTeXEquation \\h");
        assert_eq!(result[3].prefix_text, "2\u{2010}");
        // A section dot style changes every prefix.
        let section = NumberingStyle { mode: MODE_SECTION.into(), separator: ".".into() };
        let result = restyle_images(&facts, &section);
        assert_eq!(result[1].prefix_text, "1.0.");
        assert!(result.iter().all(|item| item.changed));
        // Continuous numbering: no prefix and no restart anywhere.
        let result = restyle_images(&facts, &NumberingStyle::sequence());
        assert!(result.iter().all(|item| item.prefix_text.is_empty() && !item.restart_at_one));
        // A number that is not one MACROBUTTON with its two SEQ fields is rebuilt.
        let mut broken = image(50, true, "0\u{2010}");
        broken.well_formed = false;
        let facts = RestyleFacts { images: vec![broken], ..Default::default() };
        assert!(restyle_images(&facts, &style)[0].changed);
    }
}
