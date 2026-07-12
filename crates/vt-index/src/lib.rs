use std::path::Path;

use regex::{Regex, RegexBuilder};
use vt_protocol::{
    ProjectIndex, ProjectSearchMatch, ProjectSearchRequest, ProjectSymbol, SymbolKind,
};

#[derive(Debug, thiserror::Error)]
pub enum IndexError {
    #[error("search query must not be empty")]
    EmptyQuery,
    #[error(transparent)]
    Regex(#[from] regex::Error),
}

pub fn index_file(path: impl AsRef<Path>, source: &str) -> Result<ProjectIndex, IndexError> {
    let path = path.as_ref().to_path_buf();
    let visible = mask_latex_comments(source);
    let mut symbols = Vec::new();

    if path.extension().and_then(|value| value.to_str()) == Some("bib") {
        index_bibliography(&path, source, &visible, &mut symbols)?;
    } else {
        index_latex_commands(&path, source, &visible, &mut symbols)?;
    }
    symbols.sort_by_key(|symbol| (symbol.start_byte, symbol.end_byte));
    Ok(ProjectIndex { symbols })
}

pub fn search_file(
    path: impl AsRef<Path>,
    source: &str,
    request: &ProjectSearchRequest,
) -> Result<Vec<ProjectSearchMatch>, IndexError> {
    if request.query.is_empty() {
        return Err(IndexError::EmptyQuery);
    }
    let pattern = RegexBuilder::new(&regex::escape(&request.query))
        .case_insensitive(!request.case_sensitive)
        .unicode(true)
        .build()?;
    let limit = request.max_results.max(1);
    let path = path.as_ref().to_path_buf();
    let mut matches = Vec::new();
    for occurrence in pattern.find_iter(source) {
        if request.whole_word && !is_whole_word(source, occurrence.start(), occurrence.end()) {
            continue;
        }
        let (line, column) = byte_position(source, occurrence.start());
        matches.push(ProjectSearchMatch {
            file: path.clone(),
            start_byte: occurrence.start(),
            end_byte: occurrence.end(),
            line,
            column,
            preview: line_preview(source, occurrence.start(), occurrence.end(), 180),
        });
        if matches.len() >= limit {
            break;
        }
    }
    Ok(matches)
}

fn index_latex_commands(
    path: &Path,
    source: &str,
    visible: &str,
    output: &mut Vec<ProjectSymbol>,
) -> Result<(), IndexError> {
    let command_pattern = Regex::new(
        r"\\(?P<command>label|ref|eqref|autoref|cref|Cref|cite|citep|citet|parencite|textcite)\s*(?:\[[^\]]*\]\s*)*\{(?P<keys>[^{}]*)\}",
    )?;
    for captures in command_pattern.captures_iter(visible) {
        let Some(command) = captures.name("command") else {
            continue;
        };
        let Some(keys) = captures.name("keys") else {
            continue;
        };
        let kind = match command.as_str() {
            "label" => SymbolKind::LabelDefinition,
            "ref" | "eqref" | "autoref" | "cref" | "Cref" => SymbolKind::Reference,
            _ => SymbolKind::Citation,
        };
        for (relative_start, key) in comma_separated_keys(keys.as_str()) {
            let start = keys.start() + relative_start;
            let end = start + key.len();
            push_symbol(
                output,
                path,
                source,
                kind,
                key,
                start,
                end,
                Some(format!("\\{}", command.as_str())),
            );
        }
    }

    let package_pattern = Regex::new(r"\\usepackage\s*(?:\[[^\]]*\]\s*)?\{(?P<packages>[^{}]+)\}")?;
    for captures in package_pattern.captures_iter(visible) {
        let Some(packages) = captures.name("packages") else {
            continue;
        };
        for (relative_start, package) in comma_separated_keys(packages.as_str()) {
            let start = packages.start() + relative_start;
            push_symbol(
                output,
                path,
                source,
                SymbolKind::Package,
                package,
                start,
                start + package.len(),
                Some("\\usepackage".to_owned()),
            );
        }
    }

    let macro_pattern = Regex::new(
        r"\\(?:newcommand|renewcommand|providecommand|DeclareRobustCommand)\*?\s*(?:\{\s*)?\\(?P<name>[A-Za-z@]+)",
    )?;
    for captures in macro_pattern.captures_iter(visible) {
        let Some(name) = captures.name("name") else {
            continue;
        };
        push_symbol(
            output,
            path,
            source,
            SymbolKind::MacroDefinition,
            format!("\\{}", name.as_str()),
            name.start().saturating_sub(1),
            name.end(),
            Some("macro".to_owned()),
        );
    }
    Ok(())
}

fn index_bibliography(
    path: &Path,
    source: &str,
    visible: &str,
    output: &mut Vec<ProjectSymbol>,
) -> Result<(), IndexError> {
    let entry_pattern =
        Regex::new(r"(?m)^\s*@(?P<entry_type>[A-Za-z]+)\s*\{\s*(?P<key>[^,\s{}]+)\s*,")?;
    for captures in entry_pattern.captures_iter(visible) {
        let Some(entry_type) = captures.name("entry_type") else {
            continue;
        };
        let Some(key) = captures.name("key") else {
            continue;
        };
        push_symbol(
            output,
            path,
            source,
            SymbolKind::BibliographyEntry,
            key.as_str(),
            key.start(),
            key.end(),
            Some(format!("@{}", entry_type.as_str().to_ascii_lowercase())),
        );
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_symbol(
    output: &mut Vec<ProjectSymbol>,
    path: &Path,
    source: &str,
    kind: SymbolKind,
    key: impl Into<String>,
    start_byte: usize,
    end_byte: usize,
    detail: Option<String>,
) {
    let (line, column) = byte_position(source, start_byte);
    output.push(ProjectSymbol {
        kind,
        key: key.into(),
        file: path.to_path_buf(),
        start_byte,
        end_byte,
        line,
        column,
        detail,
    });
}

fn comma_separated_keys(value: &str) -> Vec<(usize, &str)> {
    let mut keys = Vec::new();
    let mut offset = 0usize;
    for part in value.split(',') {
        let leading = part.len().saturating_sub(part.trim_start().len());
        let key = part.trim();
        if !key.is_empty() {
            keys.push((offset + leading, key));
        }
        offset += part.len() + 1;
    }
    keys
}

fn mask_latex_comments(source: &str) -> String {
    let mut bytes = source.as_bytes().to_vec();
    let mut line_start = 0usize;
    while line_start < bytes.len() {
        let line_end = bytes[line_start..]
            .iter()
            .position(|value| *value == b'\n')
            .map_or(bytes.len(), |relative| line_start + relative);
        let mut cursor = line_start;
        while cursor < line_end {
            if bytes[cursor] == b'%' && !is_escaped(&bytes, cursor) {
                for value in &mut bytes[cursor..line_end] {
                    *value = b' ';
                }
                break;
            }
            cursor += 1;
        }
        line_start = line_end.saturating_add(1);
    }
    String::from_utf8(bytes).expect("comment masking preserves UTF-8")
}

fn is_escaped(bytes: &[u8], index: usize) -> bool {
    let mut backslashes = 0usize;
    let mut cursor = index;
    while cursor > 0 && bytes[cursor - 1] == b'\\' {
        backslashes += 1;
        cursor -= 1;
    }
    backslashes % 2 == 1
}

fn byte_position(source: &str, byte: usize) -> (u32, u32) {
    let mut safe = byte.min(source.len());
    while safe > 0 && !source.is_char_boundary(safe) {
        safe -= 1;
    }
    let prefix = &source[..safe];
    let line = prefix.bytes().filter(|value| *value == b'\n').count() as u32 + 1;
    let line_start = prefix.rfind('\n').map_or(0, |index| index + 1);
    let column = source[line_start..safe].chars().count() as u32 + 1;
    (line, column)
}

fn is_whole_word(source: &str, start: usize, end: usize) -> bool {
    let before = source[..start].chars().next_back();
    let after = source[end..].chars().next();
    before.is_none_or(|value| !is_word_character(value))
        && after.is_none_or(|value| !is_word_character(value))
}

fn is_word_character(value: char) -> bool {
    value.is_alphanumeric() || value == '_'
}

fn line_preview(source: &str, start: usize, end: usize, max_chars: usize) -> String {
    let line_start = source[..start].rfind('\n').map_or(0, |index| index + 1);
    let line_end = source[end..]
        .find('\n')
        .map_or(source.len(), |relative| end + relative);
    let line = source[line_start..line_end].trim();
    if line.chars().count() <= max_chars {
        return line.to_owned();
    }
    let match_offset = source[line_start..start].chars().count();
    let half = max_chars / 2;
    let skip = match_offset.saturating_sub(half);
    let preview = line.chars().skip(skip).take(max_chars).collect::<String>();
    if skip > 0 {
        format!("…{preview}")
    } else {
        format!("{preview}…")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use proptest::prelude::*;

    #[test]
    fn indexes_latex_and_bibliography_without_comments() {
        let latex = r#"\usepackage{amsmath, cleveref}
\newcommand{\vect}[1]{\mathbf{#1}}
\label{sec:中文}
See \cref{sec:中文,fig:one} and \citep{einstein1905, knuth84}.
% \label{ignored}
"#;
        let index = index_file("main.tex", latex).unwrap();
        assert!(index.symbols.iter().any(|symbol| {
            symbol.kind == SymbolKind::LabelDefinition && symbol.key == "sec:中文"
        }));
        assert!(
            index
                .symbols
                .iter()
                .any(|symbol| { symbol.kind == SymbolKind::Reference && symbol.key == "fig:one" })
        );
        assert!(
            index
                .symbols
                .iter()
                .any(|symbol| { symbol.kind == SymbolKind::Citation && symbol.key == "knuth84" })
        );
        assert!(index.symbols.iter().any(|symbol| {
            symbol.kind == SymbolKind::MacroDefinition && symbol.key == "\\vect"
        }));
        assert!(!index.symbols.iter().any(|symbol| symbol.key == "ignored"));

        let bibliography = r#"@article{einstein1905,
  title={Zur Elektrodynamik}
}
@book{knuth84, title={The TeXbook}}
"#;
        let bib_index = index_file("refs.bib", bibliography).unwrap();
        assert_eq!(
            bib_index
                .symbols
                .iter()
                .filter(|symbol| symbol.kind == SymbolKind::BibliographyEntry)
                .count(),
            2
        );
    }

    #[test]
    fn unicode_search_reports_character_column_and_whole_words() {
        let source = "第一行\n中文 abc abc2 ABC\n";
        let request = ProjectSearchRequest {
            query: "abc".into(),
            case_sensitive: false,
            whole_word: true,
            max_results: 10,
        };
        let matches = search_file("main.tex", source, &request).unwrap();
        assert_eq!(matches.len(), 2);
        assert_eq!((matches[0].line, matches[0].column), (2, 4));
    }

    proptest! {
        #[test]
        fn arbitrary_unicode_search_never_panics(source in ".{0,300}", query in ".{1,12}") {
            let request = ProjectSearchRequest {
                query,
                case_sensitive: false,
                whole_word: false,
                max_results: 20,
            };
            let _ = search_file(std::path::PathBuf::from("f.tex"), &source, &request);
        }
    }
}
