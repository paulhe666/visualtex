use std::path::PathBuf;

use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use uuid::Uuid;

pub const PROTOCOL_VERSION: u32 = 1;

macro_rules! id_type {
    ($name:ident) => {
        #[derive(
            Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd, Serialize, Deserialize,
        )]
        #[serde(transparent)]
        pub struct $name(pub Uuid);

        impl $name {
            pub fn new() -> Self {
                Self(Uuid::new_v4())
            }
        }

        impl Default for $name {
            fn default() -> Self {
                Self::new()
            }
        }
    };
}

id_type!(FileId);
id_type!(OperationId);
id_type!(NodeId);
id_type!(BuildId);
id_type!(ProjectId);

#[derive(Clone, Copy, Debug, Default, Eq, Ord, PartialEq, PartialOrd, Serialize, Deserialize)]
#[serde(transparent)]
pub struct Revision(pub u64);

impl Revision {
    pub fn next(self) -> Self {
        Self(self.0.saturating_add(1))
    }
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum EditOrigin {
    SourceEditor,
    VisualEditor,
    MathEditor,
    Ocr,
    ExternalFileChange,
    Formatter,
    Refactor,
    UndoRedo,
    Plugin,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextEdit {
    pub operation_id: OperationId,
    pub origin: EditOrigin,
    pub file_id: FileId,
    pub base_revision: Revision,
    pub start_byte: usize,
    pub end_byte: usize,
    pub replacement: String,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AppliedEdit {
    pub operation_id: OperationId,
    pub file_id: FileId,
    pub old_revision: Revision,
    pub new_revision: Revision,
    pub start_byte: usize,
    pub old_end_byte: usize,
    pub new_end_byte: usize,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SourceSpan {
    pub file_id: FileId,
    pub start_byte: usize,
    pub end_byte: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SupportLevel {
    Native,
    Partial,
    Opaque,
    Unstable,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum NodeKind {
    Document,
    Preamble,
    Title,
    Author,
    Abstract,
    Section,
    Subsection,
    Paragraph,
    Text,
    InlineMath,
    DisplayMath,
    Figure,
    Table,
    List,
    Theorem,
    Citation,
    Reference,
    Footnote,
    Bibliography,
    RawLatex,
}

#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NodeAttributes {
    pub placement: Option<String>,
    pub caption: Option<String>,
    pub label: Option<String>,
    pub image_path: Option<String>,
    pub image_width: Option<String>,
    pub column_spec: Option<String>,
    #[serde(default)]
    pub table_rows: Vec<Vec<String>>,
}

#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NodeAttributesPatch {
    pub placement: Option<String>,
    pub caption: Option<String>,
    pub label: Option<String>,
    pub image_path: Option<String>,
    pub image_width: Option<String>,
    pub column_spec: Option<String>,
    pub table_rows: Option<Vec<Vec<String>>>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct VisualNode {
    pub id: NodeId,
    pub kind: NodeKind,
    pub support: SupportLevel,
    pub source: SourceSpan,
    pub children: Vec<NodeId>,
    pub text: Option<String>,
    pub command: Option<String>,
    #[serde(default)]
    pub attributes: NodeAttributes,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum VisualPatch {
    Reset {
        revision: Revision,
        nodes: Vec<VisualNode>,
    },
    Replace {
        revision: Revision,
        removed: Vec<NodeId>,
        upserted: Vec<VisualNode>,
    },
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum DiagnosticSeverity {
    Error,
    Warning,
    Information,
    Hint,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Diagnostic {
    pub severity: DiagnosticSeverity,
    pub message: String,
    pub file: Option<PathBuf>,
    pub line: Option<u32>,
    pub column: Option<u32>,
    pub code: Option<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SymbolKind {
    LabelDefinition,
    Reference,
    Citation,
    BibliographyEntry,
    MacroDefinition,
    Package,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectSymbol {
    pub kind: SymbolKind,
    pub key: String,
    pub file: PathBuf,
    pub start_byte: usize,
    pub end_byte: usize,
    pub line: u32,
    pub column: u32,
    pub detail: Option<String>,
}

#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectIndex {
    pub symbols: Vec<ProjectSymbol>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DependencyKind {
    Input,
    Include,
    Subfile,
    SubfileInclude,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectDependencyEdge {
    pub source_file: PathBuf,
    pub target_file: PathBuf,
    pub raw_path: String,
    pub kind: DependencyKind,
    pub start_byte: usize,
    pub end_byte: usize,
    pub resolved: bool,
}

#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectDependencyGraph {
    pub edges: Vec<ProjectDependencyEdge>,
    pub cycles: Vec<Vec<PathBuf>>,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectSearchRequest {
    pub query: String,
    pub case_sensitive: bool,
    pub whole_word: bool,
    pub max_results: usize,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectSearchMatch {
    pub file: PathBuf,
    pub start_byte: usize,
    pub end_byte: usize,
    pub line: u32,
    pub column: u32,
    pub preview: String,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectReplaceRequest {
    pub query: String,
    pub replacement: String,
    pub case_sensitive: bool,
    pub whole_word: bool,
    pub max_replacements: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SymbolRenameKind {
    Label,
    Citation,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SymbolRenameRequest {
    pub kind: SymbolRenameKind,
    pub old_key: String,
    pub new_key: String,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectTextReplacement {
    pub start_byte: usize,
    pub end_byte: usize,
    pub expected: String,
    pub replacement: String,
    pub line: u32,
    pub column: u32,
    pub preview: String,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectReplaceFilePlan {
    pub file: PathBuf,
    pub expected_sha256: String,
    pub replacements: Vec<ProjectTextReplacement>,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectReplacePlan {
    pub plan_id: OperationId,
    pub description: String,
    pub files: Vec<ProjectReplaceFilePlan>,
    pub total_replacements: usize,
    pub truncated: bool,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectReplaceOutcome {
    pub plan_id: OperationId,
    pub changed_files: Vec<PathBuf>,
    pub total_replacements: usize,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CompileStatus {
    Queued,
    Running,
    Succeeded,
    Failed,
    Cancelled,
    TimedOut,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CompileArtifact {
    pub build_id: BuildId,
    pub source_revision: Revision,
    pub pdf_path: Option<PathBuf>,
    pub synctex_path: Option<PathBuf>,
    pub diagnostics: Vec<Diagnostic>,
    pub status: CompileStatus,
    pub started_at: DateTime<Utc>,
    pub finished_at: Option<DateTime<Utc>>,
    pub stdout: String,
    pub stderr: String,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectTemplateSummary {
    pub id: String,
    pub name: String,
    pub description: String,
    pub engine: String,
    pub root_file: PathBuf,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectConfig {
    pub root_file: PathBuf,
    pub engine: TexEngine,
    pub builder: TexBuilder,
    pub output_directory: PathBuf,
    pub shell_escape: bool,
    pub restricted_mode: bool,
}

impl Default for ProjectConfig {
    fn default() -> Self {
        Self {
            root_file: PathBuf::from("main.tex"),
            engine: TexEngine::XeLatex,
            builder: TexBuilder::Latexmk,
            output_directory: PathBuf::from(".visualtex/build"),
            shell_escape: false,
            restricted_mode: true,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum TexEngine {
    PdfLatex,
    XeLatex,
    LuaLatex,
}

impl TexEngine {
    pub fn executable(self) -> &'static str {
        match self {
            Self::PdfLatex => "pdflatex",
            Self::XeLatex => "xelatex",
            Self::LuaLatex => "lualatex",
        }
    }

    pub fn latexmk_flag(self) -> &'static str {
        match self {
            Self::PdfLatex => "-pdf",
            Self::XeLatex => "-xelatex",
            Self::LuaLatex => "-lualatex",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum TexBuilder {
    Latexmk,
    Tectonic,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfRect {
    pub page: u32,
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfPageInfo {
    pub index: u32,
    pub width_points: f32,
    pub height_points: f32,
    pub rotation_degrees: i16,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfDocumentInfo {
    pub pdf_path: PathBuf,
    pub fingerprint: String,
    pub byte_len: u64,
    pub pages: Vec<PdfPageInfo>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfPixelRect {
    pub x: u32,
    pub y: u32,
    pub width: u32,
    pub height: u32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfRenderRequest {
    pub pdf_path: PathBuf,
    pub page_index: u32,
    pub target_width_pixels: u32,
    pub tile: Option<PdfPixelRect>,
    pub grayscale: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfRenderedImage {
    pub pdf_fingerprint: String,
    pub page_index: u32,
    pub page_width_pixels: u32,
    pub page_height_pixels: u32,
    pub image_width_pixels: u32,
    pub image_height_pixels: u32,
    pub tile: Option<PdfPixelRect>,
    pub cache_path: PathBuf,
    pub cache_hit: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfPixelDiffPage {
    pub page_index: u32,
    pub width_pixels: u32,
    pub height_pixels: u32,
    pub changed_pixels: u64,
    pub total_pixels: u64,
    pub changed_ratio: f64,
    pub maximum_channel_delta: u8,
    pub mean_absolute_channel_delta: f64,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfPixelDiffReport {
    pub page_count_matches: bool,
    pub left_page_count: u32,
    pub right_page_count: u32,
    pub tolerance: u8,
    pub pages: Vec<PdfPixelDiffPage>,
    pub maximum_changed_ratio: f64,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum MappingConfidence {
    Exact,
    High,
    Medium,
    Low,
    Unmapped,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum MappingMethod {
    ShadowMarkerAndSyncTex,
    ShadowMarker,
    SyncTex,
    None,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfPoint {
    pub page: u32,
    pub x: f32,
    pub y: f32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutBox {
    pub node_id: NodeId,
    pub source: SourceSpan,
    pub rects: Vec<PdfRect>,
    pub start_marker: Option<PdfPoint>,
    pub end_marker: Option<PdfPoint>,
    pub confidence: MappingConfidence,
    pub method: MappingMethod,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutMapArtifact {
    pub build_id: BuildId,
    pub source_revision: Revision,
    pub shadow_root: PathBuf,
    pub shadow_pdf_path: Option<PathBuf>,
    pub compile_status: CompileStatus,
    pub diagnostics: Vec<Diagnostic>,
    pub boxes: Vec<LayoutBox>,
    pub pixel_diff: Option<PdfPixelDiffReport>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ForwardSearchResult {
    pub pdf_path: PathBuf,
    pub boxes: Vec<PdfRect>,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct InverseSearchResult {
    pub source_path: PathBuf,
    pub line: u32,
    pub column: Option<u32>,
    pub offset: Option<i64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OcrCandidate {
    pub latex: String,
    pub confidence: f32,
    pub backend: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaOcrResult {
    pub candidates: Vec<OcrCandidate>,
    pub model_version: Option<String>,
    pub warnings: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OcrRegion {
    pub kind: String,
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub text: Option<String>,
    pub latex: Option<String>,
    pub confidence: f32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentOcrResult {
    #[serde(default)]
    pub image_path: Option<PathBuf>,
    #[serde(default)]
    pub page_width: u32,
    #[serde(default)]
    pub page_height: u32,
    pub regions: Vec<OcrRegion>,
    pub reading_order: Vec<usize>,
    pub model_version: Option<String>,
    pub warnings: Vec<String>,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OcrWorkerHealth {
    pub available: bool,
    pub backend: String,
    pub model_version: Option<String>,
    pub detail: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ToolInfo {
    pub name: String,
    pub path: Option<PathBuf>,
    pub version: Option<String>,
    pub available: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ExternalChangeKind {
    Modified,
    Deleted,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalFileChange {
    pub file_id: FileId,
    pub path: PathBuf,
    pub kind: ExternalChangeKind,
    pub buffer_dirty: bool,
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalChangeReport {
    pub reloaded: Vec<DocumentSnapshot>,
    pub conflicts: Vec<ExternalFileChange>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ExternalConflictResolution {
    ReloadDisk,
    KeepBuffer,
    SaveCopyAndReload,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentSnapshot {
    pub file_id: FileId,
    pub path: PathBuf,
    pub revision: Revision,
    pub text: String,
    pub dirty: bool,
    pub nodes: Vec<VisualNode>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalConflictOutcome {
    pub snapshot: DocumentSnapshot,
    pub conflict_copy_path: Option<PathBuf>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RpcRequest {
    pub jsonrpc: String,
    pub id: serde_json::Value,
    pub method: String,
    #[serde(default)]
    pub params: serde_json::Value,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RpcResponse {
    pub jsonrpc: String,
    pub id: serde_json::Value,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub result: Option<serde_json::Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<RpcError>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RpcError {
    pub code: i32,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data: Option<serde_json::Value>,
}

#[derive(Debug, thiserror::Error)]
pub enum ProtocolError {
    #[error("protocol version mismatch: expected {expected}, received {received}")]
    VersionMismatch { expected: u32, received: u32 },
}
