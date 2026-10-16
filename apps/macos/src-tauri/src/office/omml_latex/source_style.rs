//! LaTeX read back from a Word or PowerPoint equation, written the way people
//! type it. Office equations stay pure Office objects with no VisualTeX source
//! stored in them, so this is all the user gets back.
//!
//! Port of MathMlToLatexConverter.WordEquationSourceStyle (Windows); the two
//! must produce the same text, which the reference corpus test checks.

const TEXT_COMMANDS: [&str; 10] = [
    "\\text", "\\textrm", "\\textit", "\\textbf", "\\textsf", "\\texttt",
    "\\textup", "\\textnormal", "\\mbox", "\\operatorname",
];

fn matrix_fence(token: &str) -> Option<(&'static str, &'static str)> {
    match token {
        "(" => Some(("pmatrix", ")")),
        "[" => Some(("bmatrix", "]")),
        "|" => Some(("vmatrix", "|")),
        "\\|" => Some(("Vmatrix", "\\|")),
        "\\{" => Some(("Bmatrix", "\\}")),
        _ => None,
    }
}

/// 𝑎 -> a (the mathematical alphanumerics PowerPoint stores), x^{2} -> x^2,
/// {x}^2 -> x^2, f^{\prime} -> f', \left(\begin{matrix}..\end{matrix}\right)
/// -> pmatrix (bmatrix, vmatrix, Vmatrix, Bmatrix), \left\{ matrix \right.
/// -> cases, and drops the spaces TeX ignores. Text groups are left alone.
pub(crate) fn word_equation_source_style(latex: &str) -> String {
    if latex.is_empty() {
        return String::new();
    }
    let tokens = tokens(latex);
    let tokens = alphanumerics(&tokens);
    let tokens = matrices(&tokens);
    let tokens = scripts(&tokens);
    spacing(&tokens)
}

fn is_letter(value: char) -> bool {
    value.is_ascii_alphabetic()
}

fn is_letter_or_digit(value: char) -> bool {
    value.is_ascii_alphanumeric()
}

fn is_space(value: char) -> bool {
    matches!(value, ' ' | '\t' | '\r' | '\n')
}

fn first_char(token: &str) -> char {
    token.chars().next().unwrap_or('\0')
}

/// Control words, control symbols, runs of ASCII letters, runs of spaces (as
/// one " "), and every other character on its own.
fn tokens(latex: &str) -> Vec<String> {
    let chars: Vec<char> = latex.chars().collect();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < chars.len() {
        let start = index;
        let current = chars[index];
        if current == '\\' && index + 1 < chars.len() {
            index += 1;
            if is_letter(chars[index]) {
                while index < chars.len() && is_letter(chars[index]) {
                    index += 1;
                }
            } else {
                index += 1;
            }
            tokens.push(chars[start..index].iter().collect());
        } else if is_space(current) {
            while index < chars.len() && is_space(chars[index]) {
                index += 1;
            }
            tokens.push(" ".to_string());
        } else if is_letter(current) {
            while index < chars.len() && is_letter(chars[index]) {
                index += 1;
            }
            tokens.push(chars[start..index].iter().collect());
        } else {
            index += 1;
            tokens.push(current.to_string());
        }
    }
    tokens
}

/// Unicode mathematical alphanumerics (U+1D400-U+1D7FF), which PowerPoint
/// stores for math letters (𝑎 for an italic a), as LaTeX tokens.
fn math_alphanumeric(value: char) -> Option<Vec<String>> {
    const LATIN: [(u32, &str); 13] = [
        (0x1D400, "\\mathbf"), (0x1D434, ""), (0x1D468, "\\boldsymbol"), (0x1D49C, "\\mathcal"),
        (0x1D4D0, "\\boldsymbol\\mathcal"), (0x1D504, "\\mathfrak"), (0x1D538, "\\mathbb"),
        (0x1D56C, "\\boldsymbol\\mathfrak"), (0x1D5A0, "\\mathsf"), (0x1D5D4, "\\boldsymbol\\mathsf"),
        (0x1D608, "\\mathsf"), (0x1D63C, "\\boldsymbol\\mathsf"), (0x1D670, "\\mathtt"),
    ];
    const GREEK_BLOCKS: [(u32, &str); 5] = [
        (0x1D6A8, "\\boldsymbol"), (0x1D6E2, ""), (0x1D71C, "\\boldsymbol"),
        (0x1D756, "\\boldsymbol"), (0x1D790, "\\boldsymbol"),
    ];
    const GREEK: [char; 58] = [
        'Α', 'Β', 'Γ', 'Δ', 'Ε', 'Ζ', 'Η', 'Θ', 'Ι', 'Κ', 'Λ', 'Μ', 'Ν', 'Ξ', 'Ο', 'Π', 'Ρ', 'ϴ', 'Σ',
        'Τ', 'Υ', 'Φ', 'Χ', 'Ψ', 'Ω', '∇', 'α', 'β', 'γ', 'δ', 'ε', 'ζ', 'η', 'θ', 'ι', 'κ', 'λ', 'μ',
        'ν', 'ξ', 'ο', 'π', 'ρ', 'ς', 'σ', 'τ', 'υ', 'φ', 'χ', 'ψ', 'ω', '∂', 'ϵ', 'ϑ', 'ϰ', 'ϕ', 'ϱ', 'ϖ',
    ];
    const DIGITS: [(u32, &str); 5] = [
        (0x1D7CE, "\\mathbf"), (0x1D7D8, "\\mathbb"), (0x1D7E2, "\\mathsf"),
        (0x1D7EC, "\\boldsymbol\\mathsf"), (0x1D7F6, "\\mathtt"),
    ];
    fn styled(style: &str, inner: Vec<String>) -> Vec<String> {
        let commands: Vec<&str> = style.split('\\').filter(|part| !part.is_empty()).collect();
        let mut output = Vec::new();
        for command in &commands {
            output.push(format!("\\{command}"));
            output.push("{".to_string());
        }
        output.extend(inner);
        output.extend(commands.iter().map(|_| "}".to_string()));
        output
    }
    let code = value as u32;
    if value == '\u{210E}' {
        return Some(vec!["h".to_string()]);
    }
    if code == 0x1D6A4 || code == 0x1D6A5 {
        return Some(vec![if code == 0x1D6A4 { "\\imath" } else { "\\jmath" }.to_string()]);
    }
    for (start, style) in LATIN {
        if (start..start + 52).contains(&code) {
            let offset = (code - start) as u8;
            let letter = if offset < 26 { (b'A' + offset) as char } else { (b'a' + offset - 26) as char };
            return Some(styled(style, vec![letter.to_string()]));
        }
    }
    for (start, style) in GREEK_BLOCKS {
        if (start..start + 58).contains(&code) {
            let base = GREEK[(code - start) as usize].to_string();
            let inner = tokens(super::mathml_latex::token_latex(&base).trim())
                .into_iter()
                .filter(|token| token != " ")
                .collect();
            return Some(styled(style, inner));
        }
    }
    for (start, style) in DIGITS {
        if (start..start + 10).contains(&code) {
            return Some(styled(style, vec![((b'0' + (code - start) as u8) as char).to_string()]));
        }
    }
    None
}

fn alphanumerics(tokens: &[String]) -> Vec<String> {
    let mut output = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        if let Some(text_end) = text_group_end(tokens, index) {
            output.extend_from_slice(&tokens[index..=text_end]);
            index = text_end + 1;
            continue;
        }
        let mut chars = tokens[index].chars();
        if let (Some(value), None) = (chars.next(), chars.next()) {
            if let Some(mapped) = math_alphanumeric(value) {
                output.extend(mapped);
                index += 1;
                continue;
            }
        }
        output.push(tokens[index].clone());
        index += 1;
    }
    output
}

fn is_control_word(token: &str) -> bool {
    let mut chars = token.chars();
    chars.next() == Some('\\') && chars.next().is_some_and(is_letter)
}

fn skip_spaces(tokens: &[String], mut index: usize) -> usize {
    while index < tokens.len() && tokens[index] == " " {
        index += 1;
    }
    index
}

fn matching_brace(tokens: &[String], open: usize) -> Option<usize> {
    let mut depth = 0usize;
    for (index, token) in tokens.iter().enumerate().skip(open) {
        if token == "{" {
            depth += 1;
        } else if token == "}" {
            depth = depth.saturating_sub(1);
            if depth == 0 {
                return Some(index);
            }
        }
    }
    None
}

/// The last token of a text-mode group (\text{..}) starting at index.
fn text_group_end(tokens: &[String], index: usize) -> Option<usize> {
    if !TEXT_COMMANDS.contains(&tokens[index].as_str()) {
        return None;
    }
    let mut open = skip_spaces(tokens, index + 1);
    if open < tokens.len() && tokens[open] == "*" {
        open = skip_spaces(tokens, open + 1);
    }
    if open < tokens.len() && tokens[open] == "{" {
        matching_brace(tokens, open)
    } else {
        None
    }
}

fn is_environment(tokens: &[String], index: usize, command: &str, name: &str) -> bool {
    index + 3 < tokens.len()
        && tokens[index] == command
        && tokens[index + 1] == "{"
        && tokens[index + 2] == name
        && tokens[index + 3] == "}"
}

fn matrices(tokens: &[String]) -> Vec<String> {
    let mut output = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        if let Some(text_end) = text_group_end(tokens, index) {
            output.extend_from_slice(&tokens[index..=text_end]);
            index = text_end + 1;
            continue;
        }
        if tokens[index] == "\\left" {
            if let Some((named, last)) = named_matrix(tokens, index) {
                output.extend(named);
                index = last + 1;
                continue;
            }
        }
        output.push(tokens[index].clone());
        index += 1;
    }
    output
}

fn named_matrix(tokens: &[String], left: usize) -> Option<(Vec<String>, usize)> {
    let fence = skip_spaces(tokens, left + 1);
    let (kind_environment, kind_close) = matrix_fence(tokens.get(fence)?)?;
    let begin = skip_spaces(tokens, fence + 1);
    if !is_environment(tokens, begin, "\\begin", "matrix") {
        return None;
    }
    let mut depth = 0usize;
    let mut end = None;
    for index in begin..tokens.len() {
        if is_environment(tokens, index, "\\begin", "matrix") {
            depth += 1;
        } else if is_environment(tokens, index, "\\end", "matrix") {
            depth = depth.saturating_sub(1);
            if depth == 0 {
                end = Some(index);
                break;
            }
        }
    }
    let end = end?;
    let right = skip_spaces(tokens, end + 4);
    if tokens.get(right).map(String::as_str) != Some("\\right") {
        return None;
    }
    let close = skip_spaces(tokens, right + 1);
    let close_token = tokens.get(close)?;
    let environment = if close_token == kind_close {
        kind_environment
    } else if tokens[fence] == "\\{" && close_token == "." {
        "cases"
    } else {
        return None;
    };
    let mut named: Vec<String> = ["\\begin", "{", environment, "}"].iter().map(|item| item.to_string()).collect();
    named.extend(matrices(&tokens[begin + 4..end]));
    named.extend(["\\end", "{", environment, "}"].iter().map(|item| item.to_string()));
    Some((named, close))
}

fn scripts(tokens: &[String]) -> Vec<String> {
    let mut output: Vec<String> = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        if let Some(text_end) = text_group_end(tokens, index) {
            output.extend_from_slice(&tokens[index..=text_end]);
            index = text_end + 1;
            continue;
        }
        let token = &tokens[index];
        // {x}^2 -> x^2 (PowerPoint letters arrive braced); a command's own
        // argument (\mathrm{x}^2, \frac{a}{x}^2) keeps its braces.
        if token == "{"
            && index + 3 < tokens.len()
            && tokens[index + 2] == "}"
            && (tokens[index + 3] == "^" || tokens[index + 3] == "_")
            && tokens[index + 1].chars().count() == 1
            && is_letter_or_digit(first_char(&tokens[index + 1]))
        {
            let previous = output.iter().rev().find(|item| *item != " ");
            if previous.map_or(true, |item| !is_control_word(item) && !matches!(item.as_str(), "}" | "]" | "^" | "_")) {
                output.push(tokens[index + 1].clone());
                index += 3;
                continue;
            }
        }
        if (token == "^" || token == "_") && index + 1 < tokens.len() && tokens[index + 1] == "{" {
            let close = matching_brace(tokens, index + 1);
            let inner: Vec<&String> = match close {
                Some(close) => tokens[index + 2..close].iter().filter(|item| *item != " ").collect(),
                None => Vec::new(),
            };
            if let Some(close) = close {
                // x^{2} -> x^2, keeping a space before a following letter or digit.
                if inner.len() == 1 && inner[0].chars().count() == 1 && is_letter_or_digit(first_char(inner[0])) {
                    output.push(token.clone());
                    output.push(inner[0].clone());
                    if close + 1 < tokens.len() && is_letter_or_digit(first_char(&tokens[close + 1])) {
                        output.push(" ".to_string());
                    }
                    index = close + 1;
                    continue;
                }
                // f^{\prime} -> f', the same thing in TeX.
                let previous = output.iter().rev().find(|item| *item != " ");
                if token == "^"
                    && !inner.is_empty()
                    && inner.iter().all(|item| *item == "\\prime")
                    && previous.is_some_and(|item| !matches!(item.as_str(), "{" | "&" | "^" | "_" | "\\\\"))
                {
                    output.push("'".repeat(inner.len()));
                    index = close + 1;
                    continue;
                }
            }
        }
        output.push(token.clone());
        index += 1;
    }
    output
}

fn needs_space(previous: &str, before_previous: Option<&str>, next: &str) -> bool {
    (is_control_word(previous) && is_letter(first_char(next)))
        || (matches!(before_previous, Some("^") | Some("_"))
            && previous.chars().count() == 1
            && is_letter_or_digit(first_char(previous))
            && is_letter_or_digit(first_char(next)))
}

fn spacing(tokens: &[String]) -> String {
    let mut text = String::new();
    let mut previous: Option<String> = None;
    let mut before_previous: Option<String> = None;
    let mut index = 0;
    while index < tokens.len() {
        if let Some(text_end) = text_group_end(tokens, index) {
            for item in &tokens[index..=text_end] {
                text.push_str(item);
            }
            before_previous = previous.take();
            previous = Some("}".to_string());
            index = text_end + 1;
            continue;
        }
        let token = &tokens[index];
        if token != " " {
            // A control word or a bare script needs a space before a letter
            // that a mapped PowerPoint letter put right after it (\sin x).
            if let Some(previous) = previous.as_deref() {
                if !text.ends_with(' ') && needs_space(previous, before_previous.as_deref(), token) {
                    text.push(' ');
                }
            }
            text.push_str(token);
            before_previous = previous.take();
            previous = Some(token.clone());
            index += 1;
            continue;
        }
        let next = tokens.get(index + 1);
        index += 1;
        let (Some(previous), Some(next)) = (previous.as_deref(), next) else { continue };
        // A space ends a control word before a letter (\sin x) and a bare
        // script before one (\int_0^1 x); around & and \\ it keeps matrices
        // and aligned rows readable.
        let keep = (is_control_word(previous) && (is_letter_or_digit(first_char(next)) || next == "*"))
            || (matches!(before_previous.as_deref(), Some("^") | Some("_"))
                && previous.chars().count() == 1
                && is_letter_or_digit(first_char(previous))
                && is_letter_or_digit(first_char(next)))
            || matches!(previous, "&" | "\\\\")
            || matches!(next.as_str(), "&" | "\\\\");
        if keep {
            text.push(' ');
        }
    }
    text
}
