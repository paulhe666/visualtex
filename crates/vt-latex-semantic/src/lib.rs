use std::collections::HashMap;

use uuid::Uuid;
use vt_protocol::{FileId, NodeAttributes, NodeId, NodeKind, SourceSpan, SupportLevel, VisualNode};

#[derive(Clone, Debug, Default)]
pub struct SemanticDocument {
    pub nodes: Vec<VisualNode>,
}

impl SemanticDocument {
    pub fn parse(file_id: FileId, source: &str) -> Self {
        Parser::new(file_id, source).parse()
    }

    pub fn node_at_byte(&self, byte: usize) -> Option<&VisualNode> {
        self.nodes
            .iter()
            .filter(|node| node.source.start_byte <= byte && byte <= node.source.end_byte)
            .min_by_key(|node| node.source.end_byte.saturating_sub(node.source.start_byte))
    }
}

struct Parser<'a> {
    file_id: FileId,
    source: &'a str,
    nodes: Vec<VisualNode>,
    occurrence: HashMap<String, usize>,
}

impl<'a> Parser<'a> {
    fn new(file_id: FileId, source: &'a str) -> Self {
        Self {
            file_id,
            source,
            nodes: Vec::new(),
            occurrence: HashMap::new(),
        }
    }

    fn parse(mut self) -> SemanticDocument {
        self.parse_commands();
        self.parse_math();
        self.parse_environments();
        self.parse_paragraphs();
        self.nodes
            .sort_by_key(|node| (node.source.start_byte, node.source.end_byte));

        let children = self.nodes.iter().map(|node| node.id).collect::<Vec<_>>();
        let root = VisualNode {
            id: self.stable_id(NodeKind::Document, "document", 0),
            kind: NodeKind::Document,
            support: SupportLevel::Native,
            source: SourceSpan {
                file_id: self.file_id,
                start_byte: 0,
                end_byte: self.source.len(),
            },
            children,
            text: None,
            command: None,
            attributes: NodeAttributes::default(),
        };
        self.nodes.insert(0, root);
        SemanticDocument { nodes: self.nodes }
    }

    fn parse_commands(&mut self) {
        const COMMANDS: &[(&str, NodeKind, SupportLevel)] = &[
            ("title", NodeKind::Title, SupportLevel::Native),
            ("author", NodeKind::Author, SupportLevel::Native),
            ("section", NodeKind::Section, SupportLevel::Native),
            ("section*", NodeKind::Section, SupportLevel::Native),
            ("subsection", NodeKind::Subsection, SupportLevel::Native),
            ("subsection*", NodeKind::Subsection, SupportLevel::Native),
            ("cite", NodeKind::Citation, SupportLevel::Partial),
            ("citep", NodeKind::Citation, SupportLevel::Partial),
            ("citet", NodeKind::Citation, SupportLevel::Partial),
            ("ref", NodeKind::Reference, SupportLevel::Partial),
            ("eqref", NodeKind::Reference, SupportLevel::Partial),
            ("footnote", NodeKind::Footnote, SupportLevel::Native),
            (
                "bibliography",
                NodeKind::Bibliography,
                SupportLevel::Partial,
            ),
        ];

        let bytes = self.source.as_bytes();
        let mut cursor = 0;
        while cursor < bytes.len() {
            if bytes[cursor] != b'\\' || self.is_commented(cursor) {
                cursor += 1;
                continue;
            }
            let name_start = cursor + 1;
            let mut name_end = name_start;
            while name_end < bytes.len()
                && (bytes[name_end].is_ascii_alphabetic() || bytes[name_end] == b'@')
            {
                name_end += 1;
            }
            if name_end < bytes.len() && bytes[name_end] == b'*' {
                name_end += 1;
            }
            let command = &self.source[name_start..name_end];
            let Some((_, kind, support)) = COMMANDS.iter().find(|(name, _, _)| *name == command)
            else {
                cursor = name_end.max(cursor + 1);
                continue;
            };
            let mut brace = name_end;
            while brace < bytes.len() && bytes[brace].is_ascii_whitespace() {
                brace += 1;
            }
            if brace >= bytes.len() || bytes[brace] != b'{' {
                self.push_node(
                    kind.clone(),
                    SupportLevel::Unstable,
                    cursor,
                    name_end,
                    Some(command.to_owned()),
                    None,
                );
                cursor = name_end;
                continue;
            }
            match balanced_group(self.source, brace) {
                Some(end) => {
                    let content = self.source[brace + 1..end - 1].to_owned();
                    self.push_node(
                        kind.clone(),
                        *support,
                        cursor,
                        end,
                        Some(command.to_owned()),
                        Some(content),
                    );
                    cursor = end;
                }
                None => {
                    self.push_node(
                        kind.clone(),
                        SupportLevel::Unstable,
                        cursor,
                        self.source.len(),
                        Some(command.to_owned()),
                        Some(self.source[brace + 1..].to_owned()),
                    );
                    break;
                }
            }
        }
    }

    fn parse_math(&mut self) {
        let bytes = self.source.as_bytes();
        let mut cursor = 0;
        while cursor < bytes.len() {
            if self.is_commented(cursor) {
                cursor += 1;
                continue;
            }
            if bytes[cursor] == b'\\' && cursor + 1 < bytes.len() && bytes[cursor + 1] == b'[' {
                if let Some(relative) = self.source[cursor + 2..].find("\\]") {
                    let end = cursor + 2 + relative + 2;
                    self.push_if_uncovered(
                        NodeKind::DisplayMath,
                        SupportLevel::Native,
                        cursor,
                        end,
                        Some("\\[".into()),
                        Some(self.source[cursor + 2..end - 2].to_owned()),
                    );
                    cursor = end;
                    continue;
                }
                self.push_if_uncovered(
                    NodeKind::DisplayMath,
                    SupportLevel::Unstable,
                    cursor,
                    self.source.len(),
                    Some("\\[".into()),
                    Some(self.source[cursor + 2..].to_owned()),
                );
                break;
            }
            if bytes[cursor] == b'$' && !is_escaped(bytes, cursor) {
                let display = cursor + 1 < bytes.len() && bytes[cursor + 1] == b'$';
                let delimiter = if display { "$$" } else { "$" };
                let content_start = cursor + delimiter.len();
                if let Some(end_start) = find_unescaped(self.source, delimiter, content_start) {
                    let end = end_start + delimiter.len();
                    self.push_if_uncovered(
                        if display {
                            NodeKind::DisplayMath
                        } else {
                            NodeKind::InlineMath
                        },
                        SupportLevel::Native,
                        cursor,
                        end,
                        Some(delimiter.into()),
                        Some(self.source[content_start..end_start].to_owned()),
                    );
                    cursor = end;
                    continue;
                }
                self.push_if_uncovered(
                    if display {
                        NodeKind::DisplayMath
                    } else {
                        NodeKind::InlineMath
                    },
                    SupportLevel::Unstable,
                    cursor,
                    self.source.len(),
                    Some(delimiter.into()),
                    Some(self.source[content_start..].to_owned()),
                );
                break;
            }
            cursor += 1;
        }
    }

    fn parse_environments(&mut self) {
        let mut cursor = 0;
        while let Some(relative) = self.source[cursor..].find("\\begin{") {
            let start = cursor + relative;
            if self.is_commented(start) {
                cursor = start + 7;
                continue;
            }
            let name_start = start + "\\begin{".len();
            let Some(name_end_rel) = self.source[name_start..].find('}') else {
                self.push_if_uncovered(
                    NodeKind::RawLatex,
                    SupportLevel::Unstable,
                    start,
                    self.source.len(),
                    Some("begin".into()),
                    Some(self.source[start..].to_owned()),
                );
                break;
            };
            let name_end = name_start + name_end_rel;
            let name = &self.source[name_start..name_end];
            let begin_end = name_end + 1;
            if name == "document" {
                cursor = begin_end;
                continue;
            }
            let close = format!("\\end{{{name}}}");
            let (end, support) = match self.source[begin_end..].find(&close) {
                Some(relative_end) => (
                    begin_end + relative_end + close.len(),
                    support_for_environment(name),
                ),
                None => (self.source.len(), SupportLevel::Unstable),
            };
            let kind = kind_for_environment(name);
            self.push_if_uncovered(
                kind,
                support,
                start,
                end,
                Some(name.to_owned()),
                Some(self.source[begin_end..end.saturating_sub(close.len())].to_owned()),
            );
            if end == self.source.len() && support == SupportLevel::Unstable {
                break;
            }
            cursor = end;
        }
    }

    fn parse_paragraphs(&mut self) {
        let document_begin = "\\begin{document}";
        let document_end = "\\end{document}";
        let body_start = self
            .source
            .find(document_begin)
            .map_or(0, |index| index + document_begin.len());
        let body_end = self
            .source
            .rfind(document_end)
            .filter(|index| *index >= body_start)
            .unwrap_or(self.source.len());

        let body = &self.source[body_start..body_end];
        let mut block_start = body_start;
        for block in body.split_inclusive("\n\n") {
            let block_end = block_start + block.len();
            self.push_paragraph_block(block_start, block_end);
            block_start = block_end;
        }
    }

    fn push_paragraph_block(&mut self, start: usize, end: usize) {
        let (mut start, end) = trim_source_range(self.source, start, end);
        if start >= end {
            return;
        }

        loop {
            let structural_end = self
                .nodes
                .iter()
                .filter(|node| node.source.start_byte == start)
                .filter(|node| {
                    matches!(
                        node.kind,
                        NodeKind::Section
                            | NodeKind::Subsection
                            | NodeKind::Title
                            | NodeKind::Author
                    )
                })
                .map(|node| node.source.end_byte)
                .max();
            if let Some(structural_end) = structural_end {
                (start, _) = trim_source_range(self.source, structural_end, end);
                if start >= end {
                    return;
                }
                continue;
            }

            let fragment = &self.source[start..end];
            let removable_command = ["\\maketitle", "\\tableofcontents"]
                .into_iter()
                .find(|command| fragment.starts_with(command));
            if let Some(command) = removable_command {
                let after_command = start + command.len();
                start = self.source[after_command..end]
                    .find('\n')
                    .map_or(end, |newline| after_command + newline + 1);
                (start, _) = trim_source_range(self.source, start, end);
                if start >= end {
                    return;
                }
                continue;
            }
            break;
        }

        let fragment = &self.source[start..end];
        if fragment.starts_with('\\') || self.covered_by_structural_node(start, end) {
            return;
        }
        let support = if self.nodes.iter().any(|node| {
            node.source.start_byte >= start
                && node.source.end_byte <= end
                && !matches!(node.kind, NodeKind::Paragraph | NodeKind::Text)
        }) {
            SupportLevel::Partial
        } else {
            SupportLevel::Native
        };
        self.push_node(
            NodeKind::Paragraph,
            support,
            start,
            end,
            None,
            Some(fragment.to_owned()),
        );
    }

    fn push_if_uncovered(
        &mut self,
        kind: NodeKind,
        support: SupportLevel,
        start: usize,
        end: usize,
        command: Option<String>,
        text: Option<String>,
    ) {
        if !self.covered_by_same_span(start, end) {
            self.push_node(kind, support, start, end, command, text);
        }
    }

    fn push_node(
        &mut self,
        kind: NodeKind,
        support: SupportLevel,
        start: usize,
        end: usize,
        command: Option<String>,
        text: Option<String>,
    ) {
        let identity = format!(
            "{:?}|{}|{}",
            kind,
            command.as_deref().unwrap_or(""),
            text.as_deref().unwrap_or("")
        );
        let occurrence = {
            let value = self.occurrence.entry(identity.clone()).or_insert(0);
            let occurrence = *value;
            *value += 1;
            occurrence
        };
        let id = self.stable_id(kind.clone(), &identity, occurrence);
        let attributes = parse_node_attributes(&kind, &self.source[start..end]);
        self.nodes.push(VisualNode {
            id,
            kind,
            support,
            source: SourceSpan {
                file_id: self.file_id,
                start_byte: start,
                end_byte: end,
            },
            children: Vec::new(),
            text,
            command,
            attributes,
        });
    }

    fn stable_id(&self, kind: NodeKind, identity: &str, occurrence: usize) -> NodeId {
        let name = format!("{:?}|{identity}|{occurrence}", kind);
        NodeId(Uuid::new_v5(&self.file_id.0, name.as_bytes()))
    }

    fn covered_by_same_span(&self, start: usize, end: usize) -> bool {
        self.nodes
            .iter()
            .any(|node| node.source.start_byte == start && node.source.end_byte == end)
    }

    fn covered_by_structural_node(&self, start: usize, end: usize) -> bool {
        self.nodes.iter().any(|node| {
            node.source.start_byte <= start
                && node.source.end_byte >= end
                && !matches!(
                    node.kind,
                    NodeKind::InlineMath | NodeKind::Citation | NodeKind::Reference
                )
        })
    }

    fn is_commented(&self, byte: usize) -> bool {
        let bytes = self.source.as_bytes();
        let byte = byte.min(bytes.len());
        let line_start = bytes[..byte]
            .iter()
            .rposition(|value| *value == b'\n')
            .map_or(0, |index| index + 1);
        (line_start..byte).any(|index| bytes[index] == b'%' && !is_escaped(bytes, index))
    }
}

fn trim_source_range(source: &str, mut start: usize, mut end: usize) -> (usize, usize) {
    while start < end {
        let character = source[start..end].chars().next().expect("non-empty range");
        if !character.is_whitespace() {
            break;
        }
        start += character.len_utf8();
    }
    while start < end {
        let character = source[start..end]
            .chars()
            .next_back()
            .expect("non-empty range");
        if !character.is_whitespace() {
            break;
        }
        end -= character.len_utf8();
    }
    (start, end)
}

fn balanced_group(source: &str, open: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    let mut depth = 0usize;
    for (offset, value) in bytes[open..].iter().enumerate() {
        let index = open + offset;
        if is_escaped(bytes, index) {
            continue;
        }
        match value {
            b'{' => depth += 1,
            b'}' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Some(index + 1);
                }
            }
            _ => {}
        }
    }
    None
}

fn find_unescaped(source: &str, needle: &str, from: usize) -> Option<usize> {
    let mut cursor = from;
    while let Some(relative) = source[cursor..].find(needle) {
        let index = cursor + relative;
        if !is_escaped(source.as_bytes(), index) {
            return Some(index);
        }
        cursor = index + needle.len();
    }
    None
}

fn is_escaped(bytes: &[u8], index: usize) -> bool {
    let mut backslashes = 0;
    let mut cursor = index;
    while cursor > 0 && bytes[cursor - 1] == b'\\' {
        backslashes += 1;
        cursor -= 1;
    }
    backslashes % 2 == 1
}

fn parse_node_attributes(kind: &NodeKind, fragment: &str) -> NodeAttributes {
    let mut attributes = NodeAttributes::default();
    if !matches!(kind, NodeKind::Figure | NodeKind::Table) {
        return attributes;
    }

    attributes.placement = environment_placement(fragment);
    attributes.caption = command_group_argument(fragment, "caption");
    attributes.label = command_group_argument(fragment, "label");

    if *kind == NodeKind::Figure {
        if let Some((options, path)) = includegraphics_arguments(fragment) {
            attributes.image_path = Some(path);
            attributes.image_width = options.as_deref().and_then(option_width);
        }
    } else if let Some((column_spec, body)) = tabular_parts(fragment) {
        attributes.column_spec = Some(column_spec);
        attributes.table_rows = parse_simple_table_rows(&body);
    }
    attributes
}

fn environment_placement(fragment: &str) -> Option<String> {
    let begin = fragment.find("\\begin{")?;
    let name_open = begin + "\\begin".len();
    let name_end = balanced_group(fragment, name_open)?;
    let mut cursor = name_end;
    while fragment
        .as_bytes()
        .get(cursor)
        .is_some_and(u8::is_ascii_whitespace)
    {
        cursor += 1;
    }
    parse_bracket_group(fragment, cursor).map(|(value, _)| value)
}

fn command_group_argument(fragment: &str, command: &str) -> Option<String> {
    let marker = format!("\\{command}");
    let command_start = fragment.find(&marker)?;
    let mut cursor = command_start + marker.len();
    while fragment
        .as_bytes()
        .get(cursor)
        .is_some_and(u8::is_ascii_whitespace)
    {
        cursor += 1;
    }
    if let Some((_, end)) = parse_bracket_group(fragment, cursor) {
        cursor = end;
        while fragment
            .as_bytes()
            .get(cursor)
            .is_some_and(u8::is_ascii_whitespace)
        {
            cursor += 1;
        }
    }
    if fragment.as_bytes().get(cursor) != Some(&b'{') {
        return None;
    }
    let end = balanced_group(fragment, cursor)?;
    Some(fragment[cursor + 1..end - 1].to_owned())
}

fn includegraphics_arguments(fragment: &str) -> Option<(Option<String>, String)> {
    let marker = "\\includegraphics";
    let start = fragment.find(marker)?;
    let mut cursor = start + marker.len();
    while fragment
        .as_bytes()
        .get(cursor)
        .is_some_and(u8::is_ascii_whitespace)
    {
        cursor += 1;
    }
    let options = parse_bracket_group(fragment, cursor);
    if let Some((_, end)) = &options {
        cursor = *end;
        while fragment
            .as_bytes()
            .get(cursor)
            .is_some_and(u8::is_ascii_whitespace)
        {
            cursor += 1;
        }
    }
    if fragment.as_bytes().get(cursor) != Some(&b'{') {
        return None;
    }
    let end = balanced_group(fragment, cursor)?;
    Some((
        options.map(|(value, _)| value),
        fragment[cursor + 1..end - 1].to_owned(),
    ))
}

fn option_width(options: &str) -> Option<String> {
    options.split(',').find_map(|option| {
        let (key, value) = option.split_once('=')?;
        (key.trim() == "width").then(|| value.trim().to_owned())
    })
}

fn tabular_parts(fragment: &str) -> Option<(String, String)> {
    let marker = "\\begin{tabular}";
    let start = fragment.find(marker)?;
    let mut cursor = start + marker.len();
    while fragment
        .as_bytes()
        .get(cursor)
        .is_some_and(u8::is_ascii_whitespace)
    {
        cursor += 1;
    }
    if fragment.as_bytes().get(cursor) != Some(&b'{') {
        return None;
    }
    let spec_end = balanced_group(fragment, cursor)?;
    let close = fragment[spec_end..].find("\\end{tabular}")? + spec_end;
    Some((
        fragment[cursor + 1..spec_end - 1].to_owned(),
        fragment[spec_end..close].trim().to_owned(),
    ))
}

fn parse_simple_table_rows(body: &str) -> Vec<Vec<String>> {
    split_unescaped_sequence(body, "\\\\")
        .into_iter()
        .map(str::trim)
        .filter(|row| !row.is_empty() && !row.starts_with("\\hline"))
        .map(|row| {
            split_unescaped_sequence(row.trim_start_matches("\\hline").trim(), "&")
                .into_iter()
                .map(|cell| cell.trim().to_owned())
                .collect()
        })
        .collect()
}

fn parse_bracket_group(source: &str, open: usize) -> Option<(String, usize)> {
    if source.as_bytes().get(open) != Some(&b'[') {
        return None;
    }
    let bytes = source.as_bytes();
    let mut depth = 0usize;
    for index in open..bytes.len() {
        if is_escaped(bytes, index) {
            continue;
        }
        match bytes[index] {
            b'[' => depth += 1,
            b']' => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    return Some((source[open + 1..index].to_owned(), index + 1));
                }
            }
            _ => {}
        }
    }
    None
}

fn split_unescaped_sequence<'a>(source: &'a str, delimiter: &str) -> Vec<&'a str> {
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut cursor = 0usize;
    while cursor + delimiter.len() <= source.len() {
        if source[cursor..].starts_with(delimiter)
            && (cursor == 0 || !is_escaped(source.as_bytes(), cursor))
        {
            parts.push(&source[start..cursor]);
            cursor += delimiter.len();
            start = cursor;
        } else {
            cursor += source[cursor..].chars().next().map_or(1, char::len_utf8);
        }
    }
    parts.push(&source[start..]);
    parts
}

fn kind_for_environment(name: &str) -> NodeKind {
    match name {
        "abstract" => NodeKind::Abstract,
        "figure" | "figure*" => NodeKind::Figure,
        "table" | "table*" | "tabular" | "longtable" => NodeKind::Table,
        "itemize" | "enumerate" | "description" => NodeKind::List,
        "theorem" | "lemma" | "proof" | "definition" | "proposition" => NodeKind::Theorem,
        "equation" | "equation*" | "align" | "align*" | "gather" | "multline" => {
            NodeKind::DisplayMath
        }
        _ => NodeKind::RawLatex,
    }
}

fn support_for_environment(name: &str) -> SupportLevel {
    match kind_for_environment(name) {
        NodeKind::Abstract | NodeKind::List | NodeKind::DisplayMath => SupportLevel::Native,
        NodeKind::Figure | NodeKind::Table | NodeKind::Theorem => SupportLevel::Partial,
        NodeKind::RawLatex => SupportLevel::Opaque,
        _ => SupportLevel::Opaque,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use proptest::prelude::*;

    #[test]
    fn parses_common_paper_nodes_and_preserves_unknown_environment() {
        let file_id = FileId::new();
        let source = r#"\title{你好}
\author{Paul}
\begin{document}
\section{Introduction}
Text with $E=mc^2$ and \cite{einstein}.
\begin{mystery}
\custom{unchanged}
\end{mystery}
\end{document}"#;
        let document = SemanticDocument::parse(file_id, source);
        assert!(
            document
                .nodes
                .iter()
                .any(|node| node.kind == NodeKind::Title)
        );
        assert!(
            document
                .nodes
                .iter()
                .any(|node| node.kind == NodeKind::Section)
        );
        assert!(
            document
                .nodes
                .iter()
                .any(|node| node.kind == NodeKind::InlineMath)
        );
        assert!(document.nodes.iter().any(|node| {
            node.kind == NodeKind::RawLatex && node.support == SupportLevel::Opaque
        }));
        assert!(
            !document
                .nodes
                .iter()
                .any(|node| node.command.as_deref() == Some("document"))
        );
    }

    #[test]
    fn extracts_figure_and_simple_table_attributes() {
        let source = r#"\documentclass{article}
\usepackage{graphicx}
\begin{document}
\begin{figure}[htbp]
\custombefore
\includegraphics[keepaspectratio,width=0.72\linewidth]{figures/sample image.pdf}
\caption[Short]{A detailed caption}
\label{fig:sample}
\end{figure}
\begin{table}[t]
\caption{Measurements}
\label{tab:data}
\begin{tabular}{lc}
Name & Value \\
Alpha & 1 \\
\end{tabular}
\end{table}
\end{document}"#;
        let document = SemanticDocument::parse(FileId::new(), source);
        let figure = document
            .nodes
            .iter()
            .find(|node| node.kind == NodeKind::Figure)
            .unwrap();
        assert_eq!(figure.attributes.placement.as_deref(), Some("htbp"));
        assert_eq!(
            figure.attributes.image_path.as_deref(),
            Some("figures/sample image.pdf")
        );
        assert_eq!(
            figure.attributes.image_width.as_deref(),
            Some("0.72\\linewidth")
        );
        assert_eq!(
            figure.attributes.caption.as_deref(),
            Some("A detailed caption")
        );
        assert_eq!(figure.attributes.label.as_deref(), Some("fig:sample"));

        let table = document
            .nodes
            .iter()
            .find(|node| node.kind == NodeKind::Table && node.command.as_deref() == Some("table"))
            .unwrap();
        assert_eq!(table.attributes.placement.as_deref(), Some("t"));
        assert_eq!(table.attributes.caption.as_deref(), Some("Measurements"));
        assert_eq!(table.attributes.column_spec.as_deref(), Some("lc"));
        assert_eq!(
            table.attributes.table_rows,
            vec![
                vec!["Name".to_owned(), "Value".to_owned()],
                vec!["Alpha".to_owned(), "1".to_owned()]
            ]
        );
    }

    #[test]
    fn document_wrapper_does_not_hide_body_nodes() {
        let source = r#"\documentclass{article}
\begin{document}
\section{Body}
Plain text.

\begin{equation}
x^2=1
\end{equation}
\end{document}"#;
        let document = SemanticDocument::parse(FileId::new(), source);
        assert!(
            document
                .nodes
                .iter()
                .any(|node| node.kind == NodeKind::Paragraph
                    && node.text.as_deref() == Some("Plain text."))
        );
        assert!(
            document
                .nodes
                .iter()
                .any(|node| node.kind == NodeKind::DisplayMath
                    && node.command.as_deref() == Some("equation"))
        );
        assert!(!document.nodes.iter().any(|node| {
            node.kind == NodeKind::Paragraph
                && node
                    .text
                    .as_deref()
                    .is_some_and(|text| text.contains("documentclass"))
        }));
    }

    #[test]
    fn incomplete_group_becomes_unstable() {
        let document = SemanticDocument::parse(FileId::new(), "\\section{unfinished");
        assert!(
            document.nodes.iter().any(
                |node| node.kind == NodeKind::Section && node.support == SupportLevel::Unstable
            )
        );
    }

    #[test]
    fn identical_parse_produces_identical_ids() {
        let file_id = FileId::new();
        let first = SemanticDocument::parse(file_id, "\\section{A}\nText");
        let second = SemanticDocument::parse(file_id, "\\section{A}\nText");
        assert_eq!(
            first.nodes.iter().map(|node| node.id).collect::<Vec<_>>(),
            second.nodes.iter().map(|node| node.id).collect::<Vec<_>>()
        );
    }

    proptest! {
        #[test]
        fn arbitrary_unicode_and_incomplete_latex_never_panics(
            source_chars in proptest::collection::vec(any::<char>(), 0..300),
        ) {
            let source = source_chars.into_iter().collect::<String>();
            let parsed = SemanticDocument::parse(FileId::new(), &source);
            prop_assert!(!parsed.nodes.is_empty());
            prop_assert_eq!(parsed.nodes[0].kind.clone(), NodeKind::Document);
            prop_assert_eq!(parsed.nodes[0].source.end_byte, source.len());
        }
    }
}
