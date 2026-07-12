use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::path::{Component, Path, PathBuf};

use serde::{Deserialize, Serialize};
use tree_sitter::{InputEdit, Parser, Point, Range, Tree};
use vt_protocol::{DependencyKind, ProjectDependencyEdge, ProjectDependencyGraph};

#[derive(Debug, thiserror::Error)]
pub enum SyntaxError {
    #[error("failed to load the Tree-sitter LaTeX grammar: {0}")]
    Language(String),
    #[error("Tree-sitter did not return a syntax tree")]
    ParseCancelled,
    #[error("syntax edit range {start_byte}..{old_end_byte} is invalid for {old_len} bytes")]
    InvalidEditRange {
        start_byte: usize,
        old_end_byte: usize,
        old_len: usize,
    },
    #[error("syntax edit does not produce the supplied new source length")]
    InvalidNewLength,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SyntaxRange {
    pub start_byte: usize,
    pub end_byte: usize,
    pub start_row: usize,
    pub start_column: usize,
    pub end_row: usize,
    pub end_column: usize,
}

impl From<Range> for SyntaxRange {
    fn from(range: Range) -> Self {
        Self {
            start_byte: range.start_byte,
            end_byte: range.end_byte,
            start_row: range.start_point.row,
            start_column: range.start_point.column,
            end_row: range.end_point.row,
            end_column: range.end_point.column,
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SyntaxIssue {
    pub kind: String,
    pub start_byte: usize,
    pub end_byte: usize,
    pub missing: bool,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DependencyDirective {
    pub kind: DependencyKind,
    pub raw_path: String,
    pub start_byte: usize,
    pub end_byte: usize,
}

#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct SyntaxUpdate {
    pub changed_ranges: Vec<SyntaxRange>,
    pub issues: Vec<SyntaxIssue>,
    pub dependencies_changed: bool,
}

pub struct SyntaxDocument {
    parser: Parser,
    tree: Tree,
    source_len: usize,
    dependencies: Vec<DependencyDirective>,
    issues: Vec<SyntaxIssue>,
    protected_ranges: Vec<SyntaxRange>,
}

impl fmt::Debug for SyntaxDocument {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SyntaxDocument")
            .field("source_len", &self.source_len)
            .field("dependencies", &self.dependencies)
            .field("issues", &self.issues)
            .field("protected_ranges", &self.protected_ranges)
            .finish_non_exhaustive()
    }
}

impl SyntaxDocument {
    pub fn parse(source: &str) -> Result<Self, SyntaxError> {
        let mut parser = configured_parser()?;
        let tree = parser
            .parse(source, None)
            .ok_or(SyntaxError::ParseCancelled)?;
        let dependencies = collect_dependencies(&tree, source);
        let issues = collect_issues(&tree);
        let protected_ranges = collect_protected_ranges(&tree);
        Ok(Self {
            parser,
            tree,
            source_len: source.len(),
            dependencies,
            issues,
            protected_ranges,
        })
    }

    pub fn apply_edit(
        &mut self,
        old_source: &str,
        new_source: &str,
        start_byte: usize,
        old_end_byte: usize,
        new_end_byte: usize,
    ) -> Result<SyntaxUpdate, SyntaxError> {
        if old_source.len() != self.source_len
            || start_byte > old_end_byte
            || old_end_byte > old_source.len()
        {
            return Err(SyntaxError::InvalidEditRange {
                start_byte,
                old_end_byte,
                old_len: old_source.len(),
            });
        }
        if new_end_byte < start_byte {
            return Err(SyntaxError::InvalidNewLength);
        }
        let expected_new_len =
            old_source.len() - (old_end_byte - start_byte) + (new_end_byte - start_byte);
        if expected_new_len != new_source.len() {
            return Err(SyntaxError::InvalidNewLength);
        }

        let edit = InputEdit {
            start_byte,
            old_end_byte,
            new_end_byte,
            start_position: point_at(old_source, start_byte),
            old_end_position: point_at(old_source, old_end_byte),
            new_end_position: point_at(new_source, new_end_byte),
        };
        let mut edited_tree = self.tree.clone();
        edited_tree.edit(&edit);
        let new_tree = self
            .parser
            .parse(new_source, Some(&edited_tree))
            .ok_or(SyntaxError::ParseCancelled)?;
        let mut changed_ranges = edited_tree
            .changed_ranges(&new_tree)
            .map(SyntaxRange::from)
            .collect::<Vec<_>>();
        if changed_ranges.is_empty() {
            changed_ranges.push(SyntaxRange {
                start_byte,
                end_byte: new_end_byte,
                start_row: edit.start_position.row,
                start_column: edit.start_position.column,
                end_row: edit.new_end_position.row,
                end_column: edit.new_end_position.column,
            });
        }
        let dependencies = collect_dependencies(&new_tree, new_source);
        let issues = collect_issues(&new_tree);
        let protected_ranges = collect_protected_ranges(&new_tree);
        let dependencies_changed = dependencies != self.dependencies;

        self.tree = new_tree;
        self.source_len = new_source.len();
        self.dependencies = dependencies;
        self.issues = issues.clone();
        self.protected_ranges = protected_ranges;

        Ok(SyntaxUpdate {
            changed_ranges,
            issues,
            dependencies_changed,
        })
    }

    pub fn source_len(&self) -> usize {
        self.source_len
    }

    pub fn dependencies(&self) -> &[DependencyDirective] {
        &self.dependencies
    }

    pub fn issues(&self) -> &[SyntaxIssue] {
        &self.issues
    }

    pub fn protected_ranges(&self) -> &[SyntaxRange] {
        &self.protected_ranges
    }

    pub fn root_sexp(&self) -> String {
        self.tree.root_node().to_sexp()
    }
}

pub fn build_dependency_graph<'a>(
    documents: impl IntoIterator<Item = (PathBuf, &'a SyntaxDocument)>,
    known_files: &BTreeSet<PathBuf>,
) -> ProjectDependencyGraph {
    let mut edges = Vec::new();
    for (source_file, document) in documents {
        for dependency in document.dependencies() {
            let target_file = resolve_dependency_path(&source_file, &dependency.raw_path)
                .unwrap_or_else(|| PathBuf::from(&dependency.raw_path));
            edges.push(ProjectDependencyEdge {
                source_file: source_file.clone(),
                resolved: known_files.contains(&target_file),
                target_file,
                raw_path: dependency.raw_path.clone(),
                kind: dependency.kind,
                start_byte: dependency.start_byte,
                end_byte: dependency.end_byte,
            });
        }
    }
    edges.sort_by(|left, right| {
        (
            &left.source_file,
            left.start_byte,
            left.end_byte,
            &left.target_file,
        )
            .cmp(&(
                &right.source_file,
                right.start_byte,
                right.end_byte,
                &right.target_file,
            ))
    });
    let cycles = dependency_cycles(&edges);
    ProjectDependencyGraph { edges, cycles }
}

pub fn resolve_dependency_path(source_file: &Path, raw_path: &str) -> Option<PathBuf> {
    let raw_path = raw_path.trim();
    if raw_path.is_empty()
        || raw_path.contains(['\\', '#', '$', '{', '}'])
        || Path::new(raw_path).is_absolute()
    {
        return None;
    }

    let mut components = source_file
        .parent()
        .unwrap_or_else(|| Path::new(""))
        .components()
        .filter_map(|component| match component {
            Component::Normal(value) => Some(value.to_owned()),
            Component::CurDir => None,
            Component::ParentDir | Component::RootDir | Component::Prefix(_) => None,
        })
        .collect::<Vec<_>>();

    for component in Path::new(raw_path).components() {
        match component {
            Component::Normal(value) => components.push(value.to_owned()),
            Component::CurDir => {}
            Component::ParentDir => {
                components.pop()?;
            }
            Component::RootDir | Component::Prefix(_) => return None,
        }
    }

    let mut target = components.into_iter().collect::<PathBuf>();
    if target.extension().is_none() {
        target.set_extension("tex");
    }
    Some(target)
}

fn configured_parser() -> Result<Parser, SyntaxError> {
    let mut parser = Parser::new();
    parser
        .set_language(&codebook_tree_sitter_latex::LANGUAGE.into())
        .map_err(|error| SyntaxError::Language(error.to_string()))?;
    Ok(parser)
}

fn point_at(source: &str, byte: usize) -> Point {
    let prefix = &source.as_bytes()[..byte];
    let row = prefix.iter().filter(|value| **value == b'\n').count();
    let column = prefix
        .iter()
        .rposition(|value| *value == b'\n')
        .map_or(prefix.len(), |newline| prefix.len() - newline - 1);
    Point::new(row, column)
}

fn collect_dependencies(tree: &Tree, source: &str) -> Vec<DependencyDirective> {
    let mut dependencies = Vec::new();
    visit_named_nodes(tree, |node| {
        if node.kind() != "latex_include" {
            return;
        }
        let Some(command) = node.child_by_field_name("command") else {
            return;
        };
        let Some(path) = node.child_by_field_name("path") else {
            return;
        };
        let Ok(command) = command.utf8_text(source.as_bytes()) else {
            return;
        };
        let Ok(path_text) = path.utf8_text(source.as_bytes()) else {
            return;
        };
        let Some(kind) = dependency_kind(command) else {
            return;
        };
        let raw_path = strip_group(path_text);
        if raw_path.is_empty() {
            return;
        }
        dependencies.push(DependencyDirective {
            kind,
            raw_path: raw_path.to_owned(),
            start_byte: node.start_byte(),
            end_byte: node.end_byte(),
        });
    });
    dependencies.sort_by_key(|dependency| (dependency.start_byte, dependency.end_byte));
    dependencies
}

fn dependency_kind(command: &str) -> Option<DependencyKind> {
    match command.trim() {
        "\\input" => Some(DependencyKind::Input),
        "\\include" => Some(DependencyKind::Include),
        "\\subfile" => Some(DependencyKind::Subfile),
        "\\subfileinclude" => Some(DependencyKind::SubfileInclude),
        _ => None,
    }
}

fn strip_group(value: &str) -> &str {
    let value = value.trim();
    value
        .strip_prefix('{')
        .and_then(|value| value.strip_suffix('}'))
        .unwrap_or(value)
        .trim()
}

fn collect_issues(tree: &Tree) -> Vec<SyntaxIssue> {
    let mut issues = Vec::new();
    visit_all_nodes(tree, |node| {
        if node.is_error() || node.is_missing() {
            issues.push(SyntaxIssue {
                kind: node.kind().to_owned(),
                start_byte: node.start_byte(),
                end_byte: node.end_byte(),
                missing: node.is_missing(),
            });
        }
    });
    issues.sort_by_key(|issue| (issue.start_byte, issue.end_byte, issue.kind.clone()));
    issues
}

fn collect_protected_ranges(tree: &Tree) -> Vec<SyntaxRange> {
    const PROTECTED_KINDS: &[&str] = &[
        "verbatim_environment",
        "minted_environment",
        "listing_environment",
        "luacode_environment",
        "pycode_environment",
        "sageblock_environment",
        "sagesilent_environment",
        "asy_environment",
        "asydef_environment",
    ];
    let mut ranges = Vec::new();
    visit_named_nodes(tree, |node| {
        if PROTECTED_KINDS.contains(&node.kind()) {
            ranges.push(SyntaxRange::from(node.range()));
        }
    });
    ranges.sort_by_key(|range| (range.start_byte, range.end_byte));
    ranges
}

fn visit_named_nodes(tree: &Tree, mut visitor: impl FnMut(tree_sitter::Node<'_>)) {
    let mut cursor = tree.walk();
    let mut reached_root = false;
    while !reached_root {
        let node = cursor.node();
        if node.is_named() {
            visitor(node);
        }
        if cursor.goto_first_child() {
            continue;
        }
        if cursor.goto_next_sibling() {
            continue;
        }
        loop {
            if !cursor.goto_parent() {
                reached_root = true;
                break;
            }
            if cursor.goto_next_sibling() {
                break;
            }
        }
    }
}

fn visit_all_nodes(tree: &Tree, mut visitor: impl FnMut(tree_sitter::Node<'_>)) {
    let mut cursor = tree.walk();
    let mut reached_root = false;
    while !reached_root {
        visitor(cursor.node());
        if cursor.goto_first_child() {
            continue;
        }
        if cursor.goto_next_sibling() {
            continue;
        }
        loop {
            if !cursor.goto_parent() {
                reached_root = true;
                break;
            }
            if cursor.goto_next_sibling() {
                break;
            }
        }
    }
}

fn dependency_cycles(edges: &[ProjectDependencyEdge]) -> Vec<Vec<PathBuf>> {
    let mut adjacency = BTreeMap::<PathBuf, Vec<PathBuf>>::new();
    for edge in edges.iter().filter(|edge| edge.resolved) {
        adjacency
            .entry(edge.source_file.clone())
            .or_default()
            .push(edge.target_file.clone());
        adjacency.entry(edge.target_file.clone()).or_default();
    }
    for targets in adjacency.values_mut() {
        targets.sort();
        targets.dedup();
    }

    let mut state = BTreeMap::<PathBuf, u8>::new();
    let mut stack = Vec::<PathBuf>::new();
    let mut cycles = BTreeSet::<Vec<PathBuf>>::new();
    for node in adjacency.keys() {
        detect_cycles(node, &adjacency, &mut state, &mut stack, &mut cycles);
    }
    cycles.into_iter().collect()
}

fn detect_cycles(
    node: &PathBuf,
    adjacency: &BTreeMap<PathBuf, Vec<PathBuf>>,
    state: &mut BTreeMap<PathBuf, u8>,
    stack: &mut Vec<PathBuf>,
    cycles: &mut BTreeSet<Vec<PathBuf>>,
) {
    match state.get(node).copied().unwrap_or(0) {
        2 => return,
        1 => {
            if let Some(start) = stack.iter().position(|value| value == node) {
                cycles.insert(canonical_cycle(&stack[start..]));
            }
            return;
        }
        _ => {}
    }

    state.insert(node.clone(), 1);
    stack.push(node.clone());
    if let Some(targets) = adjacency.get(node) {
        for target in targets {
            detect_cycles(target, adjacency, state, stack, cycles);
        }
    }
    stack.pop();
    state.insert(node.clone(), 2);
}

fn canonical_cycle(cycle: &[PathBuf]) -> Vec<PathBuf> {
    if cycle.is_empty() {
        return Vec::new();
    }
    let start = cycle
        .iter()
        .enumerate()
        .min_by(|(_, left), (_, right)| left.cmp(right))
        .map_or(0, |(index, _)| index);
    let mut normalized = cycle[start..]
        .iter()
        .chain(cycle[..start].iter())
        .cloned()
        .collect::<Vec<_>>();
    normalized.push(normalized[0].clone());
    normalized
}

#[cfg(test)]
mod tests {
    use super::*;
    use proptest::prelude::*;

    #[test]
    fn incremental_tree_matches_full_parse_and_reports_local_change() {
        let old =
            "\\documentclass{article}\n\\begin{document}\n\\section{Old}\nText\n\\end{document}\n";
        let start = old.find("Old").unwrap();
        let old_end = start + "Old".len();
        let replacement = "新标题";
        let mut new = old.to_owned();
        new.replace_range(start..old_end, replacement);

        let mut incremental = SyntaxDocument::parse(old).unwrap();
        let update = incremental
            .apply_edit(old, &new, start, old_end, start + replacement.len())
            .unwrap();
        let full = SyntaxDocument::parse(&new).unwrap();

        assert_eq!(incremental.root_sexp(), full.root_sexp());
        assert!(!update.changed_ranges.is_empty());
        assert!(
            update
                .changed_ranges
                .iter()
                .all(|range| range.start_byte <= start + replacement.len())
        );
    }

    #[test]
    fn extracts_include_commands_but_not_comments() {
        let source = r#"% \input{ignored}
\input{chapters/intro}
\include{chapters/results.tex}
\subfile{appendix/a}
\subfileinclude{appendix/b}
"#;
        let document = SyntaxDocument::parse(source).unwrap();
        assert_eq!(document.dependencies().len(), 4);
        assert_eq!(document.dependencies()[0].kind, DependencyKind::Input);
        assert_eq!(document.dependencies()[0].raw_path, "chapters/intro");
        assert_eq!(
            resolve_dependency_path(Path::new("main.tex"), "chapters/intro"),
            Some(PathBuf::from("chapters/intro.tex"))
        );
        assert_eq!(
            resolve_dependency_path(Path::new("chapters/one.tex"), "../shared"),
            Some(PathBuf::from("shared.tex"))
        );
    }

    #[test]
    fn recognizes_verbatim_and_minted_as_protected_ranges() {
        let source = r#"\begin{document}
\begin{verbatim}
\section{not syntax}
\end{verbatim}
\begin{minted}{rust}
fn main() { println!("hi"); }
\end{minted}
\end{document}"#;
        let document = SyntaxDocument::parse(source).unwrap();
        assert_eq!(document.protected_ranges().len(), 2);
        for range in document.protected_ranges() {
            let fragment = &source[range.start_byte..range.end_byte];
            assert!(fragment.contains("begin"));
        }
    }

    #[test]
    fn dependency_graph_marks_missing_targets_and_detects_cycles() {
        let main = SyntaxDocument::parse("\\input{chapter}\n\\input{missing}").unwrap();
        let chapter = SyntaxDocument::parse("\\input{main}").unwrap();
        let known = BTreeSet::from([PathBuf::from("main.tex"), PathBuf::from("chapter.tex")]);
        let graph = build_dependency_graph(
            [
                (PathBuf::from("main.tex"), &main),
                (PathBuf::from("chapter.tex"), &chapter),
            ],
            &known,
        );
        assert_eq!(graph.edges.len(), 3);
        assert_eq!(graph.edges.iter().filter(|edge| edge.resolved).count(), 2);
        assert_eq!(graph.cycles.len(), 1);
        assert_eq!(graph.cycles[0].first(), graph.cycles[0].last());
    }

    #[test]
    fn malformed_latex_returns_error_nodes_without_failing_parse() {
        let document = SyntaxDocument::parse("\\begin{document}\n\\section{unfinished").unwrap();
        assert!(!document.issues().is_empty());
    }

    proptest! {
        #[test]
        fn arbitrary_unicode_full_and_incremental_parse_never_panic(
            characters in proptest::collection::vec(any::<char>(), 0..200),
            suffix in proptest::collection::vec(any::<char>(), 0..30),
        ) {
            let source = characters.into_iter().collect::<String>();
            let suffix = suffix.into_iter().collect::<String>();
            let mut updated = source.clone();
            updated.push_str(&suffix);
            let mut document = SyntaxDocument::parse(&source).unwrap();
            document
                .apply_edit(
                    &source,
                    &updated,
                    source.len(),
                    source.len(),
                    updated.len(),
                )
                .unwrap();
            let full = SyntaxDocument::parse(&updated).unwrap();
            prop_assert_eq!(document.root_sexp(), full.root_sexp());
        }
    }
}
