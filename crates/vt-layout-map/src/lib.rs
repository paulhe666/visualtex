use std::collections::HashMap;
use std::fs;
use std::path::{Component, Path, PathBuf};
use std::time::Duration;

use regex::Regex;
use vt_compiler::{CompileRequest, CompilerError};
use vt_pdf::{PdfError, PdfService};
use vt_protocol::{
    BuildId, CompileStatus, LayoutBox, LayoutMapArtifact, MappingConfidence, MappingMethod,
    NodeKind, PdfPageInfo, PdfPixelDiffReport, PdfPoint, PdfRect, ProjectConfig, Revision,
    SupportLevel, VisualNode,
};
use walkdir::WalkDir;

const SHADOW_PREAMBLE: &str = concat!(
    "\\usepackage{zref-savepos,zref-abspage}",
    "\\makeatletter\\zref@addprop{savepos}{abspage}\\makeatother"
);
const PIXEL_COMPARE_WIDTH: u32 = 1_440;
const PIXEL_COMPARE_TOLERANCE: u8 = 0;
const EXACT_PIXEL_RATIO: f64 = 0.000_001;
const ENABLE_EXPERIMENTAL_SAVE_POS_MARKERS: bool = false;

#[derive(Debug, thiserror::Error)]
pub enum LayoutMapError {
    #[error("project root is invalid: {0}")]
    InvalidProjectRoot(PathBuf),
    #[error("source file escapes project root: {0}")]
    SourceEscape(PathBuf),
    #[error("source contains no documentclass command")]
    MissingDocumentClass,
    #[error("source span is not a valid UTF-8 boundary: {start}..{end}")]
    InvalidSourceSpan { start: usize, end: usize },
    #[error("shadow copying rejects symbolic links: {0}")]
    UnsafeSymlink(PathBuf),
    #[error(transparent)]
    Compiler(#[from] CompilerError),
    #[error(transparent)]
    Pdf(#[from] PdfError),
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Regex(#[from] regex::Error),
}

#[derive(Clone, Debug)]
pub struct LayoutMapRequest {
    pub project_root: PathBuf,
    pub config: ProjectConfig,
    pub source_revision: Revision,
    pub source_file: PathBuf,
    pub source_text: String,
    pub nodes: Vec<VisualNode>,
    pub authoritative_pdf: PathBuf,
}

#[derive(Clone, Debug)]
struct InstrumentedNode {
    node: VisualNode,
    start_key: Option<String>,
    end_key: Option<String>,
}

#[derive(Clone, Copy, Debug)]
struct RawMarker {
    x_sp: i64,
    y_sp: i64,
    page: u32,
}

pub async fn build_layout_map(
    request: LayoutMapRequest,
) -> Result<LayoutMapArtifact, LayoutMapError> {
    let project_root = canonical_project_root(&request.project_root)?;
    let source_file = normalized_relative(&request.source_file)?;
    let source_on_disk = project_root.join(&source_file);
    if !source_on_disk.is_file() {
        return Err(LayoutMapError::SourceEscape(source_file));
    }

    let build_id = BuildId::new();
    let shadow_root = project_root
        .join(".visualtex/shadow")
        .join(build_id.0.simple().to_string());
    copy_project(&project_root, &shadow_root)?;

    let (instrumented_source, instrumented_nodes) =
        instrument_source(&request.source_text, &request.nodes)?;
    let shadow_source = shadow_root.join(&source_file);
    if let Some(parent) = shadow_source.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(&shadow_source, instrumented_source)?;

    let mut shadow_config = request.config.clone();
    shadow_config.output_directory = PathBuf::from(".visualtex/build");
    let compile = vt_compiler::compile(CompileRequest {
        project_root: shadow_root.clone(),
        config: shadow_config.clone(),
        source_revision: request.source_revision,
        timeout: Duration::from_secs(120),
    })
    .await?;

    let Some(shadow_pdf_path) = compile.pdf_path.clone() else {
        return Ok(LayoutMapArtifact {
            build_id,
            source_revision: request.source_revision,
            shadow_root,
            shadow_pdf_path: None,
            compile_status: compile.status,
            diagnostics: compile.diagnostics,
            boxes: request.nodes.iter().map(unmapped_layout_box).collect(),
            pixel_diff: None,
        });
    };

    if compile.status != CompileStatus::Succeeded {
        return Ok(LayoutMapArtifact {
            build_id,
            source_revision: request.source_revision,
            shadow_root,
            shadow_pdf_path: Some(shadow_pdf_path),
            compile_status: compile.status,
            diagnostics: compile.diagnostics,
            boxes: request.nodes.iter().map(unmapped_layout_box).collect(),
            pixel_diff: None,
        });
    }

    let pdf_service = PdfService::new(project_root.join(".visualtex/cache/layout-map"));
    let shadow_pdf_info = pdf_service.document_info(&shadow_pdf_path)?;
    let marker_file = shadow_aux_path(&shadow_root, &shadow_config);
    let raw_markers = if marker_file.is_file() {
        parse_marker_aux(&fs::read_to_string(marker_file)?)?
    } else {
        HashMap::new()
    };
    let pixel_diff = Some(pdf_service.compare_documents(
        &request.authoritative_pdf,
        &shadow_pdf_path,
        PIXEL_COMPARE_WIDTH,
        PIXEL_COMPARE_TOLERANCE,
    )?);
    let pixel_exact = pixel_diff.as_ref().is_some_and(pixel_diff_is_exact);

    let instrumented_by_id = instrumented_nodes
        .into_iter()
        .map(|item| (item.node.id, item))
        .collect::<HashMap<_, _>>();
    let mut boxes = Vec::with_capacity(request.nodes.len());
    for node in &request.nodes {
        let Some(instrumented) = instrumented_by_id.get(&node.id) else {
            boxes.push(unmapped_layout_box(node));
            continue;
        };
        let start_marker = instrumented
            .start_key
            .as_ref()
            .and_then(|key| raw_markers.get(key))
            .and_then(|marker| marker_to_pdf_point(*marker, &shadow_pdf_info.pages));
        let end_marker = instrumented
            .end_key
            .as_ref()
            .and_then(|key| raw_markers.get(key))
            .and_then(|marker| marker_to_pdf_point(*marker, &shadow_pdf_info.pages));
        let rects = collect_synctex_rects(
            &shadow_root,
            &source_file,
            &request.source_text,
            node,
            &shadow_pdf_path,
        )
        .await;
        let (confidence, method) = mapping_quality(
            node,
            start_marker.as_ref(),
            end_marker.as_ref(),
            &rects,
            pixel_exact,
        );
        boxes.push(LayoutBox {
            node_id: node.id,
            source: node.source.clone(),
            rects,
            start_marker,
            end_marker,
            confidence,
            method,
        });
    }

    Ok(LayoutMapArtifact {
        build_id,
        source_revision: request.source_revision,
        shadow_root,
        shadow_pdf_path: Some(shadow_pdf_path),
        compile_status: compile.status,
        diagnostics: compile.diagnostics,
        boxes,
        pixel_diff,
    })
}

fn instrument_source(
    source: &str,
    nodes: &[VisualNode],
) -> Result<(String, Vec<InstrumentedNode>), LayoutMapError> {
    let documentclass = source
        .find("\\documentclass")
        .ok_or(LayoutMapError::MissingDocumentClass)?;
    let preamble_insertion = source[documentclass..]
        .find('\n')
        .map(|offset| documentclass + offset)
        .unwrap_or(source.len());
    let document_body_start = source
        .find("\\begin{document}")
        .map(|offset| offset + "\\begin{document}".len())
        .unwrap_or(source.len());

    let mut insertions = vec![(preamble_insertion, 0_u8, SHADOW_PREAMBLE.to_owned())];
    let mut instrumented_nodes = Vec::new();
    for node in nodes {
        if !is_mappable(node, document_body_start) {
            continue;
        }
        let start = node.source.start_byte;
        let end = node.source.end_byte;
        if start > end
            || end > source.len()
            || !source.is_char_boundary(start)
            || !source.is_char_boundary(end)
        {
            return Err(LayoutMapError::InvalidSourceSpan { start, end });
        }
        let compact_id = node.id.0.simple().to_string();
        let (start_key, end_key) = if supports_zero_layout_markers(node) {
            let start_key = format!("vtS{compact_id}");
            let end_key = format!("vtE{compact_id}");
            insertions.push((start, 1, start_marker_latex(node, &start_key)));
            insertions.push((end, 2, end_marker_latex(&end_key)));
            (Some(start_key), Some(end_key))
        } else {
            (None, None)
        };
        instrumented_nodes.push(InstrumentedNode {
            node: node.clone(),
            start_key,
            end_key,
        });
    }

    insertions.sort_by(|left, right| right.0.cmp(&left.0).then_with(|| right.1.cmp(&left.1)));
    let original_line_count = source.lines().count();
    let mut instrumented = source.to_owned();
    for (byte, _, text) in insertions {
        instrumented.insert_str(byte, &text);
    }
    debug_assert_eq!(original_line_count, instrumented.lines().count());
    Ok((instrumented, instrumented_nodes))
}

fn is_mappable(node: &VisualNode, document_body_start: usize) -> bool {
    node.source.start_byte >= document_body_start
        && node.source.end_byte >= node.source.start_byte
        && matches!(node.support, SupportLevel::Native | SupportLevel::Partial)
        && !matches!(
            node.kind,
            NodeKind::Document | NodeKind::Preamble | NodeKind::RawLatex
        )
}

fn supports_zero_layout_markers(node: &VisualNode) -> bool {
    ENABLE_EXPERIMENTAL_SAVE_POS_MARKERS
        && matches!(
            node.kind,
            NodeKind::Paragraph
                | NodeKind::Text
                | NodeKind::InlineMath
                | NodeKind::Citation
                | NodeKind::Reference
                | NodeKind::Footnote
        )
}

fn start_marker_latex(node: &VisualNode, key: &str) -> String {
    if node.kind == NodeKind::Paragraph {
        format!("\\leavevmode\\vadjust pre{{\\zsavepos{{{key}}}}}")
    } else {
        format!("\\vadjust pre{{\\zsavepos{{{key}}}}}")
    }
}

fn end_marker_latex(key: &str) -> String {
    format!("\\vadjust{{\\zsavepos{{{key}}}}}")
}

fn canonical_project_root(path: &Path) -> Result<PathBuf, LayoutMapError> {
    if !path.is_dir() {
        return Err(LayoutMapError::InvalidProjectRoot(path.to_path_buf()));
    }
    Ok(path.canonicalize()?)
}

fn normalized_relative(path: &Path) -> Result<PathBuf, LayoutMapError> {
    if path.is_absolute()
        || path.components().any(|component| {
            matches!(
                component,
                Component::ParentDir | Component::RootDir | Component::Prefix(_)
            )
        })
    {
        return Err(LayoutMapError::SourceEscape(path.to_path_buf()));
    }
    Ok(path.to_path_buf())
}

fn copy_project(source_root: &Path, shadow_root: &Path) -> Result<(), LayoutMapError> {
    if shadow_root.exists() {
        fs::remove_dir_all(shadow_root)?;
    }
    fs::create_dir_all(shadow_root)?;
    let entries = WalkDir::new(source_root)
        .follow_links(false)
        .into_iter()
        .filter_entry(|entry| !is_excluded_entry(source_root, entry.path()));
    for entry in entries {
        let entry = entry.map_err(|error| {
            LayoutMapError::Io(error.into_io_error().unwrap_or_else(|| {
                std::io::Error::other("failed to traverse project for shadow build")
            }))
        })?;
        if entry.path() == source_root {
            continue;
        }
        let relative = entry
            .path()
            .strip_prefix(source_root)
            .map_err(|_| LayoutMapError::SourceEscape(entry.path().to_path_buf()))?;
        let destination = shadow_root.join(relative);
        if entry.file_type().is_symlink() {
            return Err(LayoutMapError::UnsafeSymlink(relative.to_path_buf()));
        }
        if entry.file_type().is_dir() {
            fs::create_dir_all(&destination)?;
        } else if entry.file_type().is_file() {
            if let Some(parent) = destination.parent() {
                fs::create_dir_all(parent)?;
            }
            fs::copy(entry.path(), &destination)?;
            fs::set_permissions(&destination, fs::metadata(entry.path())?.permissions())?;
        }
    }
    Ok(())
}

fn is_excluded_entry(root: &Path, path: &Path) -> bool {
    if path == root {
        return false;
    }
    let Ok(relative) = path.strip_prefix(root) else {
        return true;
    };
    matches!(
        relative.components().next(),
        Some(Component::Normal(value))
            if value == ".visualtex" || value == ".git" || value == "target" || value == "node_modules"
    )
}

fn shadow_aux_path(shadow_root: &Path, config: &ProjectConfig) -> PathBuf {
    let stem = config
        .root_file
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("main");
    shadow_root
        .join(&config.output_directory)
        .join(format!("{stem}.aux"))
}

fn parse_marker_aux(contents: &str) -> Result<HashMap<String, RawMarker>, LayoutMapError> {
    let pattern = Regex::new(concat!(
        r"\\zref@newlabel\{(?P<key>vt[SE][0-9a-f]+)\}\{",
        r"\\posx\{(?P<x>-?[0-9]+)\}",
        r"\\posy\{(?P<y>-?[0-9]+)\}",
        r"\\abspage\{(?P<page>[0-9]+)\}\}"
    ))?;
    let mut markers = HashMap::new();
    for captures in pattern.captures_iter(contents) {
        let Some(key) = captures.name("key").map(|value| value.as_str().to_owned()) else {
            continue;
        };
        let Some(x_sp) = captures
            .name("x")
            .and_then(|value| value.as_str().parse().ok())
        else {
            continue;
        };
        let Some(y_sp) = captures
            .name("y")
            .and_then(|value| value.as_str().parse().ok())
        else {
            continue;
        };
        let Some(page) = captures
            .name("page")
            .and_then(|value| value.as_str().parse().ok())
        else {
            continue;
        };
        markers.insert(key, RawMarker { x_sp, y_sp, page });
    }
    Ok(markers)
}

fn marker_to_pdf_point(marker: RawMarker, pages: &[PdfPageInfo]) -> Option<PdfPoint> {
    let page = marker.page.checked_sub(1)? as usize;
    let page_info = pages.get(page)?;
    let x = marker.x_sp as f32 / 65_536.0;
    let y_from_bottom = marker.y_sp as f32 / 65_536.0;
    Some(PdfPoint {
        page: marker.page,
        x,
        y: page_info.height_points - y_from_bottom,
    })
}

async fn collect_synctex_rects(
    shadow_root: &Path,
    source_file: &Path,
    source: &str,
    node: &VisualNode,
    shadow_pdf: &Path,
) -> Vec<PdfRect> {
    let start = byte_position(source, node.source.start_byte);
    let end_byte = node.source.end_byte.saturating_sub(1);
    let end = byte_position(source, end_byte);
    let samples = sample_source_positions(start, end, 256);

    let mut boxes = Vec::new();
    for (line, column) in samples {
        if let Ok(result) =
            vt_synctex::forward_search(shadow_root, source_file, line, column, shadow_pdf).await
        {
            for rect in result.boxes {
                if !boxes
                    .iter()
                    .any(|existing| approximately_same_rect(existing, &rect))
                {
                    boxes.push(rect);
                }
            }
        }
    }
    boxes.sort_by(|left, right| {
        left.page
            .cmp(&right.page)
            .then_with(|| left.y.total_cmp(&right.y))
            .then_with(|| left.x.total_cmp(&right.x))
    });
    boxes
}

fn sample_source_positions(start: (u32, u32), end: (u32, u32), limit: usize) -> Vec<(u32, u32)> {
    if start.0 == end.0 {
        if start.1 >= end.1 {
            return if start == end {
                vec![start]
            } else {
                vec![start, end]
            };
        }
        let column_count = (end.1 - start.1 + 1) as usize;
        let stride = column_count.div_ceil(limit.max(2)).max(1) as u32;
        let mut samples = Vec::with_capacity(column_count.min(limit) + 2);
        samples.push(start);
        let mut column = start.1.saturating_add(stride);
        while column < end.1 {
            samples.push((start.0, column));
            column = column.saturating_add(stride);
        }
        samples.push(end);
        samples.sort_unstable();
        samples.dedup();
        return samples;
    }
    if start.0 > end.0 {
        return vec![start, end];
    }
    let line_count = (end.0 - start.0 + 1) as usize;
    let stride = line_count.div_ceil(limit.max(2));
    let mut samples = Vec::with_capacity(line_count.min(limit) + 2);
    samples.push(start);
    let mut line = start.0 + 1;
    while line < end.0 {
        samples.push((line, 1));
        line = line.saturating_add(stride.max(1) as u32);
    }
    samples.push(end);
    samples.sort_unstable();
    samples.dedup();
    samples
}

fn byte_position(source: &str, byte: usize) -> (u32, u32) {
    let mut safe_byte = byte.min(source.len());
    while safe_byte > 0 && !source.is_char_boundary(safe_byte) {
        safe_byte -= 1;
    }
    let prefix = &source[..safe_byte];
    let line = prefix.bytes().filter(|value| *value == b'\n').count() as u32 + 1;
    let line_start = prefix.rfind('\n').map_or(0, |index| index + 1);
    let column = source[line_start..safe_byte].chars().count() as u32 + 1;
    (line, column)
}

fn approximately_same_rect(left: &PdfRect, right: &PdfRect) -> bool {
    left.page == right.page
        && (left.x - right.x).abs() < 0.05
        && (left.y - right.y).abs() < 0.05
        && (left.width - right.width).abs() < 0.05
        && (left.height - right.height).abs() < 0.05
}

fn mapping_quality(
    node: &VisualNode,
    start: Option<&PdfPoint>,
    end: Option<&PdfPoint>,
    rects: &[PdfRect],
    pixel_exact: bool,
) -> (MappingConfidence, MappingMethod) {
    match (start.is_some() && end.is_some(), rects.is_empty()) {
        (true, false) if pixel_exact && node.support == SupportLevel::Native => (
            MappingConfidence::Exact,
            MappingMethod::ShadowMarkerAndSyncTex,
        ),
        (true, false) => (
            MappingConfidence::High,
            MappingMethod::ShadowMarkerAndSyncTex,
        ),
        (true, true) => (MappingConfidence::Low, MappingMethod::ShadowMarker),
        (false, false)
            if pixel_exact
                && supports_high_confidence_sync_tex(node)
                && sync_tex_is_node_specific(node) =>
        {
            (MappingConfidence::High, MappingMethod::SyncTex)
        }
        (false, false) => (MappingConfidence::Medium, MappingMethod::SyncTex),
        (false, true) => (MappingConfidence::Unmapped, MappingMethod::None),
    }
}

fn supports_high_confidence_sync_tex(node: &VisualNode) -> bool {
    node.support == SupportLevel::Native
        || (node.support == SupportLevel::Partial
            && matches!(
                node.kind,
                NodeKind::Paragraph | NodeKind::Figure | NodeKind::Table
            ))
}

fn sync_tex_is_node_specific(node: &VisualNode) -> bool {
    matches!(
        node.kind,
        NodeKind::Section
            | NodeKind::Subsection
            | NodeKind::Paragraph
            | NodeKind::InlineMath
            | NodeKind::DisplayMath
            | NodeKind::Figure
            | NodeKind::Table
            | NodeKind::List
            | NodeKind::Theorem
    )
}

fn pixel_diff_is_exact(report: &PdfPixelDiffReport) -> bool {
    report.page_count_matches && report.maximum_changed_ratio <= EXACT_PIXEL_RATIO
}

fn unmapped_layout_box(node: &VisualNode) -> LayoutBox {
    LayoutBox {
        node_id: node.id,
        source: node.source.clone(),
        rects: Vec::new(),
        start_marker: None,
        end_marker: None,
        confidence: MappingConfidence::Unmapped,
        method: MappingMethod::None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vt_protocol::{FileId, NodeId, SourceSpan};

    fn paragraph_node(source: &str, start: usize, end: usize) -> VisualNode {
        VisualNode {
            id: NodeId::new(),
            kind: NodeKind::Paragraph,
            support: SupportLevel::Native,
            source: SourceSpan {
                file_id: FileId::new(),
                start_byte: start,
                end_byte: end,
            },
            children: Vec::new(),
            text: Some(source[start..end].to_owned()),
            command: None,
            attributes: vt_protocol::NodeAttributes::default(),
        }
    }

    #[test]
    fn safe_shadow_source_preserves_lines_and_user_text() {
        let source =
            "\\documentclass{article}\n\\begin{document}\n中文 paragraph.\n\\end{document}\n";
        let start = source.find("中文").unwrap();
        let end = start + "中文 paragraph.".len();
        let node = paragraph_node(source, start, end);
        let (instrumented, nodes) = instrument_source(source, &[node]).unwrap();
        assert_eq!(source.lines().count(), instrumented.lines().count());
        assert!(instrumented.contains(SHADOW_PREAMBLE));
        assert!(!instrumented.contains("\\zsavepos{vtS"));
        assert!(instrumented.contains("中文 paragraph."));
        assert_eq!(nodes.len(), 1);
        assert!(nodes[0].start_key.is_none());
        assert!(nodes[0].end_key.is_none());
    }

    #[test]
    fn preamble_nodes_are_not_instrumented() {
        let source =
            "\\documentclass{article}\n\\title{A}\n\\begin{document}\nBody\n\\end{document}\n";
        let start = source.find("\\title").unwrap();
        let end = start + "\\title{A}".len();
        let mut node = paragraph_node(source, start, end);
        node.kind = NodeKind::Title;
        let (_, nodes) = instrument_source(source, &[node]).unwrap();
        assert!(nodes.is_empty());
    }

    #[test]
    fn parses_marker_positions_with_absolute_page() {
        let aux = concat!(
            "\\zref@newlabel{vtSabc}{\\posx{65536}\\posy{131072}\\abspage{2}}\n",
            "\\zref@newlabel{vtEabc}{\\posx{196608}\\posy{262144}\\abspage{2}}\n"
        );
        let parsed = parse_marker_aux(aux).unwrap();
        assert_eq!(parsed["vtSabc"].x_sp, 65_536);
        assert_eq!(parsed["vtEabc"].page, 2);
    }

    #[test]
    fn multi_line_sampling_covers_each_line_within_limit() {
        assert_eq!(
            sample_source_positions((3, 4), (6, 8), 16),
            vec![(3, 4), (4, 1), (5, 1), (6, 8)]
        );
        let sampled = sample_source_positions((1, 1), (1_000, 1), 10);
        assert!(sampled.len() <= 12);
        assert_eq!(sampled.first(), Some(&(1, 1)));
        assert_eq!(sampled.last(), Some(&(1_000, 1)));
    }

    #[test]
    fn single_line_sampling_includes_interior_columns() {
        assert_eq!(
            sample_source_positions((4, 2), (4, 12), 4),
            vec![(4, 2), (4, 5), (4, 8), (4, 11), (4, 12)]
        );
        let sampled = sample_source_positions((2, 1), (2, 1_000), 10);
        assert!(sampled.len() <= 12);
        assert_eq!(sampled.first(), Some(&(2, 1)));
        assert_eq!(sampled.last(), Some(&(2, 1_000)));
    }

    #[test]
    fn partial_paragraph_with_inline_latex_can_use_high_confidence_sync_tex() {
        let source = "正文含有 $E=mc^2$ 以及普通文字。";
        let mut node = paragraph_node(source, 0, source.len());
        node.support = SupportLevel::Partial;
        let rect = PdfRect {
            page: 1,
            x: 10.0,
            y: 20.0,
            width: 120.0,
            height: 14.0,
        };
        assert_eq!(
            mapping_quality(&node, None, None, &[rect], true),
            (MappingConfidence::High, MappingMethod::SyncTex)
        );
    }

    #[test]
    fn inline_math_sync_tex_mapping_is_high_confidence_when_layout_is_unchanged() {
        let source = "$E=mc^2$";
        let mut node = paragraph_node(source, 0, source.len());
        node.kind = NodeKind::InlineMath;
        let rect = PdfRect {
            page: 1,
            x: 72.0,
            y: 120.0,
            width: 48.0,
            height: 12.0,
        };
        assert_eq!(
            mapping_quality(&node, None, None, &[rect], true),
            (MappingConfidence::High, MappingMethod::SyncTex)
        );
    }

    #[test]
    fn unicode_byte_positions_return_character_columns() {
        let source = "第一行\nA中B。";
        let byte = source.find('B').unwrap();
        assert_eq!(byte_position(source, byte), (2, 3));
        let punctuation = source.find('。').unwrap();
        assert_eq!(byte_position(source, punctuation + 1), (2, 4));
    }
}
