//! The offline Office integration on macOS: the add-ins talk to the app through files in their runtime folders
//! and visualtex: URLs, and the app calls back through VBA. mod.rs keeps the constants, the request and state types,
//! paths and atomic writes and the Tauri commands; each area is a child module, the tests are in tests.rs.

use crate::office::server::metadata_from_session;
use crate::office::sessions::{
    valid_uuid, CreateOfficeSessionInput, FormulaLine, OfficeFormulaSession, OfficeHost,
    OfficeSessionMode, OfficeSessionStatus, SessionError, VisualTeXFormulaMetadata,
};
use crate::office::state::OfficeCompanionState;
use base64::{
    engine::general_purpose::{STANDARD as BASE64_STANDARD, URL_SAFE_NO_PAD},
    Engine as _,
};
use flate2::{read::{DeflateDecoder, ZlibDecoder}, write::DeflateEncoder, Compression};
use serde::{Deserialize, Serialize};
#[cfg(target_os = "macos")]
use objc2::{rc::Retained, AnyThread, MainThreadMarker};
#[cfg(target_os = "macos")]
use objc2_app_kit::{
    NSApplication, NSEvent, NSScreen, NSWindowCollectionBehavior,
};
#[cfg(target_os = "macos")]
use objc2_core_services::{
    kAnyTransactionID, kAutoGenerateReturnID, keyErrorNumber, keyErrorString,
};
#[cfg(target_os = "macos")]
use objc2_foundation::{
    NSAppleEventDescriptor, NSAppleEventSendOptions, NSAppleScript, NSString,
};
use serde_json::{json, Value};
#[cfg(target_os = "macos")]
use std::cell::RefCell;
use std::fs::{self, OpenOptions};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::{
    atomic::{AtomicBool, AtomicU64, Ordering},
    mpsc, Mutex, OnceLock,
};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tauri::utils::config::BackgroundThrottlingPolicy;
use tauri::{
    AppHandle, Emitter, Manager, Monitor, WebviewUrl, WebviewWindow, WebviewWindowBuilder,
};
use uuid::Uuid;

mod presentation_redraw;

mod artifacts;
mod document_import;
mod editor_window;
mod fast_open;
mod formula_restore;
mod geometry;
mod health;
mod open_url;
mod powerpoint;
mod requests;
mod sessions;
mod vba;
mod word;

pub(crate) use artifacts::*;
pub(crate) use document_import::*;
pub(crate) use editor_window::*;
pub(crate) use fast_open::*;
pub(crate) use formula_restore::*;
use geometry::*;
pub(crate) use health::*;
pub(crate) use open_url::*;
use powerpoint::*;
pub(crate) use requests::*;
use sessions::*;
pub(crate) use vba::*;
use word::*;

const OFFLINE_PROTOCOL_VERSION: u32 = 1;
const REQUEST_FILE: &str = "request.json";
const DISPATCH_FILE: &str = "dispatch.txt";
const RESULT_PNG_FILE: &str = "formula.png";
const RESULT_SVG_FILE: &str = "formula.svg";
const RESULT_WORD_SVG_DOCX_FILE: &str = "formula-svg.docx";
const DOCUMENT_IMPORT_MANIFEST_FILE: &str = "document-import.txt";
const DOCUMENT_IMPORT_PROGRESS_FILE: &str = "document-import-progress.txt";
const LATEX_REDRAW_VECTOR_BATCH_FILE: &str = "latex-redraw-vectors.docx";
const DOCUMENT_NATIVE_BATCH_FILE: &str = "document-native-batch.docx";
const LATEX_REDRAW_SOURCE_FILE: &str = "latex-redraw-source.txt";
const LATEX_REDRAW_PREFLIGHT_MANIFEST_FILE: &str = "latex-redraw-preflight.txt";
const LATEX_REDRAW_FONT_SIZES_FILE: &str = "latex-redraw-font-sizes.txt";
const FORMULA_RESTORE_SOURCE_FILE: &str = "formula-restore-source.txt";
const WORD_PERFORMANCE_TRACE_SENTINEL: &str = "word-performance-trace.enabled";
const MAX_LATEX_REDRAW_SOURCE_BYTES: u64 = 5 * 1024 * 1024;
const EDITOR_READY_FILE: &str = "editor-ready.json";
const EDITOR_PERFORMANCE_FILE: &str = "editor-performance.jsonl";
const EDITOR_OPEN_ERROR_FILE: &str = "editor-open-error.txt";
const OFFICE_EDITOR_ACTIVATE_EVENT: &str = "visualtex-office-editor-activate";
const OFFICE_EDITOR_CLEAR_EVENT: &str = "visualtex-office-editor-clear";
const OFFICE_EDITOR_WINDOW_SIZE_FILE: &str = "editor-window-size.json";
// Measured from the user's current Office editor: 843 × 568 logical pixels on
// a 1470 × 956 logical-pixel display. Persist and restore the proportion rather
// than the absolute pixels so Retina scaling and different displays keep the
// same window density and screen coverage.
const DEFAULT_OFFICE_EDITOR_WIDTH_RATIO: f64 = 843.0 / 1470.0;
const DEFAULT_OFFICE_EDITOR_HEIGHT_RATIO: f64 = 568.0 / 956.0;
const DEFAULT_OFFICE_EDITOR_FALLBACK_WIDTH: f64 = 843.0;
const DEFAULT_OFFICE_EDITOR_FALLBACK_HEIGHT: f64 = 568.0;
const MIN_OFFICE_EDITOR_WIDTH: f64 = 500.0;
const MIN_OFFICE_EDITOR_HEIGHT: f64 = 300.0;
const MAX_OFFICE_EDITOR_WIDTH: f64 = 2600.0;
const MAX_OFFICE_EDITOR_HEIGHT: f64 = 1800.0;
const OFFICE_EDITOR_TITLE_BAR_ALLOWANCE: f64 = 28.0;
const WORD_POINTER_FILE: &str = "word-active-session.txt";
const IMAGE_INK_CENTER_CLI_ARGUMENT: &str = "--office-image-ink-center";
const POWERPOINT_POINTER_FILE: &str = "powerpoint-active-session.txt";
const WORD_RUNTIME_SUFFIX: &str = "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime";
const POWERPOINT_RUNTIME_SUFFIX: &str =
    "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime";
const METADATA_PREFIX: &str = "visualtex:v1:deflate:";
const PENDING_PREFIX: &str = "visualtex:pending:v1:";
const MAX_REQUEST_BYTES: u64 = 256 * 1024;
const WORD_FAST_OPEN_INBOX_SUFFIX: &str = "Library/Containers/com.microsoft.Word/Data/Library/Application Support/VisualTeX/FastOpen/word";
const POWERPOINT_FAST_OPEN_INBOX_SUFFIX: &str = "Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/VisualTeX/FastOpen/powerpoint";
const FAST_OPEN_MAX_AGE: Duration = Duration::from_secs(10);
const FAST_OPEN_MIN_STABLE_AGE: Duration = Duration::from_millis(20);
const FAST_OPEN_FUTURE_TOLERANCE: Duration = Duration::from_secs(2);
const FAST_OPEN_POLL_INTERVAL: Duration = Duration::from_millis(25);
const FAST_OPEN_READY_FILE: &str = "resident-ready";
const FAST_OPEN_READY_HEARTBEAT_INTERVAL: Duration = Duration::from_secs(1);
const MAX_METADATA_BYTES: usize = 2 * 1024 * 1024;
const MAX_OMML_BYTES: usize = 4 * 1024 * 1024;
const MAX_DOCUMENT_IMPORT_MANIFEST_BYTES: usize = 16 * 1024 * 1024;
const MAX_IDENTITY_CHARS: usize = 2048;
const MAX_SHAPE_NAME_CHARS: usize = 128;
const MAX_WORD_WIDTH_PT: f64 = 500.0;
const WORD_REFERENCE_FONT_SIZE_PT: f64 = 14.0;
// MathJax's built-in TeX paths render about 9.5% narrower than Word's Cambria
// Math at the same nominal point size, so the historical KaTeX/Computer Modern
// path keeps its 1.1 Word-image calibration. Times is different after VisualTeX
// replaces the letter glyphs with real Times New Roman SVG <text>: two same-
// LaTeX Word fixtures measured 36 px vs 35 px and 62 px vs 60 px horizontally,
// but 12 px vs 11 px vertically. Keep the small horizontal compensation needed
// by MathJax's TeX layout positions while removing the erroneous 10% vertical
// enlargement that made Times image formulas taller and shifted their baseline.
const WORD_TEX_IMAGE_VISUAL_SCALE: f64 = 1.1;
const WORD_TIMES_IMAGE_WIDTH_SCALE: f64 = 1.067;
const WORD_TIMES_IMAGE_HEIGHT_SCALE: f64 = 1.0;

const MIN_WORD_FONT_SIZE_PT: f64 = 1.0;
const MAX_WORD_FONT_SIZE_PT: f64 = 512.0;
const POWERPOINT_REFERENCE_FONT_SIZE_PT: f64 = 14.0;
const DEFAULT_POWERPOINT_FONT_SIZE_PT: f64 = 18.0;
const MIN_POWERPOINT_FONT_SIZE_PT: f64 = 1.0;
const MAX_POWERPOINT_FONT_SIZE_PT: f64 = 512.0;
static WORD_DISPATCH_LOCK: Mutex<()> = Mutex::new(());
static POWERPOINT_DISPATCH_LOCK: Mutex<()> = Mutex::new(());
static OFFICE_EDITOR_SIZE_WRITE_GENERATION: AtomicU64 = AtomicU64::new(0);
static FAST_OPEN_ACCESS_BACKOFF_UNTIL_MS: AtomicU64 = AtomicU64::new(0);
static FAST_OPEN_WATCHER_STARTED: OnceLock<()> = OnceLock::new();
static FAST_OPEN_KICK_IN_FLIGHT: AtomicBool = AtomicBool::new(false);
const FAST_OPEN_ACCESS_BACKOFF_MS: u64 = 2_000;

#[derive(Debug, Clone, Copy)]
struct OfficeEditorWindowSize {
    width: f64,
    height: f64,
}

#[derive(Debug, Clone, Copy, Default, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct OfficeEditorWindowSizePreference {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    width_ratio: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    height_ratio: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    fallback_width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    fallback_height: Option<f64>,
    // Compatibility with the absolute-size file written by early 1.2.4 builds.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    height: Option<f64>,
}

#[derive(Debug, Clone, Copy)]
struct OfficeEditorMonitorGeometry {
    screen_width: f64,
    screen_height: f64,
    maximum_inner_width: f64,
    maximum_inner_height: f64,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflinePowerPointRequest {
    #[serde(default)]
    native_text_target: Option<MacOfflineNativeTextTarget>,
    presentation_identity: String,
    slide_index: u32,
    slide_id: u32,
    shape_index: u32,
    shape_id: u32,
    shape_name: String,
    left: f64,
    top: f64,
    width: f64,
    height: f64,
    rotation: f64,
    z_order: u32,
    #[serde(default)]
    font_size_pt: Option<f64>,
    #[serde(default)]
    reference_width_pt: Option<f64>,
    #[serde(default)]
    reference_height_pt: Option<f64>,
    /// WPS Presentation: the Office language id of the text at the caret, for
    /// a native equation's runs (wps::presentation::language_tag).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    language_id: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflineNativeTextTarget {
    range_start: u32,
    range_length: u32,
    leading_paragraph: bool,
    trailing_paragraph: bool,
    // The text is a table cell's: PowerPoint takes the equation there only
    // through its linear format (VTWriteCellEquation).
    #[serde(default)]
    table_cell: bool,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflineDocumentImportRequest {
    bookmark_name: String,
    default_font_size_pt: f64,
    // Word's global conversion sizes for inline and display formulas; absent
    // when formulas of that kind follow the text.
    #[serde(default)]
    inline_font_size_pt: Option<f64>,
    #[serde(default)]
    display_font_size_pt: Option<f64>,
    #[serde(default)]
    redraw_scope: Option<String>,
    #[serde(default)]
    output_kind: Option<String>,
    #[serde(default)]
    source_kind: Option<String>,
    /// Update Equation Format's font step: a picture-to-picture formula restore that draws the formulas again in the
    /// fonts of the Settings, each at its own size (office::formula_fonts).
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    font_refresh: bool,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflineSessionRequest {
    protocol_version: u32,
    session_id: String,
    host: String,
    mode: String,
    #[serde(default)]
    operation: Option<String>,
    #[serde(default)]
    fork_copied_formula: bool,
    formula_id: Option<String>,
    display_mode: String,
    numbered: bool,
    #[serde(default)]
    native_equation: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    native_edit_xml: Option<String>,
    source_document_id: Option<String>,
    source_object_id: Option<String>,
    encoded_metadata: Option<String>,
    pending_marker: Option<String>,
    #[serde(default)]
    font_size_pt: Option<f64>,
    #[serde(default)]
    reference_width_pt: Option<f64>,
    #[serde(default)]
    reference_height_pt: Option<f64>,
    #[serde(default)]
    formula_letter_font: Option<String>,
    #[serde(default)]
    word_math_font_name: Option<String>,
    /// Equation format v2: the document's math font (Word's
    /// Document.OMathFontName). Word equations are written in it, as
    /// VisualTeX for Windows writes them.
    #[serde(default)]
    word_document_math_font_name: Option<String>,
    power_point: Option<MacOfflinePowerPointRequest>,
    #[serde(default)]
    document_import: Option<MacOfflineDocumentImportRequest>,
    /// "wps" when a WPS add-in opened the Session (office::wps); the result
    /// then goes to that add-in instead of a Word or PowerPoint macro.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    client: Option<String>,
    /// An AI agent's formula (open_agent_session): drawn in the hidden editor
    /// and applied without being shown.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    agent: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineFormulaRestoreTarget {
    source_start: usize,
    source_end: usize,
    source_text: String,
    display_mode: String,
    font_size_pt: f64,
    source_kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    formula_id: Option<String>,
    #[serde(default)]
    numbered: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    metadata: Option<VisualTeXFormulaMetadata>,
    #[serde(skip_serializing_if = "Option::is_none")]
    math_ml: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    latex: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineDocumentImportPublicRequest {
    protocol_version: u32,
    session_id: String,
    host: String,
    source_document_id: String,
    bookmark_name: String,
    default_font_size_pt: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    inline_font_size_pt: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    display_font_size_pt: Option<f64>,
    operation: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    redraw_scope: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    output_kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    restore_targets: Option<Vec<MacOfflineFormulaRestoreTarget>>,
    /// The restore draws picture formulas again in the current fonts (Update Equation Format).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    font_refresh: bool,
    /// "wps" when the WPS add-in started the operation (the windows name WPS).
    #[serde(skip_serializing_if = "Option::is_none")]
    client: Option<String>,
    /// A presentation's text boxes in `source`: each is searched for LaTeX on
    /// its own (presentation_redraw).
    #[serde(skip_serializing_if = "Option::is_none")]
    segments: Option<Vec<presentation_redraw::PresentationSegment>>,
    /// An AI agent's import: {text, sourceFormat, outputKind, numberDisplayFormulas}
    /// (open_agent_import); the window inserts it without waiting for the user.
    #[serde(skip_serializing_if = "Option::is_none")]
    agent: Option<Value>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineDocumentImportProgress {
    current: usize,
    total: usize,
    stage: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MacOfflineDocumentImportCommitInput {
    output_kind: String,
    items: Vec<MacOfflineDocumentImportCommitItem>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MacOfflineLatexRedrawFontRangeInput {
    source_start: usize,
    source_end: usize,
    source_text: String,
    display_mode: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MacOfflineLatexRedrawFontQueryInput {
    ranges: Vec<MacOfflineLatexRedrawFontRangeInput>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "lowercase",
    rename_all_fields = "camelCase"
)]
pub enum MacOfflineDocumentImportCommitItem {
    Text {
        text: String,
        #[serde(default)]
        source_start: Option<usize>,
        #[serde(default)]
        source_end: Option<usize>,
        #[serde(default)]
        source_text: Option<String>,
        #[serde(default)]
        paragraph_id: Option<String>,
        #[serde(default)]
        paragraph_style: Option<String>,
        #[serde(default)]
        paragraph_alignment: Option<String>,
        #[serde(default)]
        list_kind: Option<String>,
        #[serde(default)]
        list_level: Option<u32>,
        #[serde(default)]
        paragraph_start: bool,
        #[serde(default)]
        paragraph_end: bool,
    },
    Formula {
        formula_id: String,
        latex: String,
        display_mode: String,
        #[serde(default)]
        numbered: bool,
        font_size_pt: f64,
        metadata: VisualTeXFormulaMetadata,
        #[serde(default)]
        omml_base64: String,
        #[serde(default)]
        omml_docx_base64: String,
        /// MathML for the shared Word OMML writer; replaces the two above.
        #[serde(default)]
        word_math_ml: Option<String>,
        #[serde(default)]
        svg_base64: Option<String>,
        #[serde(default)]
        png_base64: Option<String>,
        #[serde(default)]
        width: Option<f64>,
        #[serde(default)]
        height: Option<f64>,
        #[serde(default)]
        baseline: Option<f64>,
        #[serde(default)]
        ink_center_y_ratio: Option<f64>,
        #[serde(default)]
        source_start: Option<usize>,
        #[serde(default)]
        source_end: Option<usize>,
        #[serde(default)]
        source_text: Option<String>,
        #[serde(default)]
        paragraph_id: Option<String>,
        #[serde(default)]
        paragraph_style: Option<String>,
        #[serde(default)]
        paragraph_alignment: Option<String>,
        #[serde(default)]
        list_kind: Option<String>,
        #[serde(default)]
        list_level: Option<u32>,
        #[serde(default)]
        paragraph_start: bool,
        #[serde(default)]
        paragraph_end: bool,
    },
}

#[derive(Debug, Clone)]
struct DocumentParagraphTransfer {
    id: String,
    style: String,
    alignment: String,
    list_kind: String,
    list_level: u32,
    start: bool,
    end: bool,
}

#[derive(Debug, Clone, Copy)]
struct WordGeometry {
    width: f64,
    height: f64,
    baseline: i32,
    font_size_pt: f64,
    reference_width_pt: f64,
    reference_height_pt: f64,
    reference_baseline_pt: f64,
}

#[derive(Debug, Clone, Copy)]
struct PowerPointGeometry {
    left: f64,
    top: f64,
    width: f64,
    height: f64,
    font_size_pt: f64,
    reference_width_pt: f64,
    reference_height_pt: f64,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflinePluginHealth {
    loaded: bool,
    plugin_version: Option<String>,
    source_revision: Option<String>,
    host: String,
    timestamp: Option<String>,
    status_path: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineOfficeEditorActivation {
    session_id: String,
    host: OfficeHost,
    generation: u64,
    received_epoch_ms: u64,
}

#[derive(Debug, Clone)]
struct ActiveOfficeEditorSession {
    activation: MacOfflineOfficeEditorActivation,
    received_at: Instant,
    ready: bool,
}

#[derive(Debug, Default)]
struct OfficeEditorRuntime {
    next_generation: u64,
    word: Option<ActiveOfficeEditorSession>,
    powerpoint: Option<ActiveOfficeEditorSession>,
}

impl OfficeEditorRuntime {
    fn active(&self, host: OfficeHost) -> Option<&ActiveOfficeEditorSession> {
        match host {
            OfficeHost::Word => self.word.as_ref(),
            OfficeHost::Powerpoint => self.powerpoint.as_ref(),
        }
    }

    fn active_mut(&mut self, host: OfficeHost) -> &mut Option<ActiveOfficeEditorSession> {
        match host {
            OfficeHost::Word => &mut self.word,
            OfficeHost::Powerpoint => &mut self.powerpoint,
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct OfficeEditorPerformanceRecord {
    schema: &'static str,
    session_id: String,
    host: OfficeHost,
    stage: String,
    epoch_ms: u64,
    elapsed_ms: f64,
    generation: Option<u64>,
    details: Value,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MacOfflineOfficeEditorReadyInput {
    session_id: String,
    generation: u64,
    frontend_epoch_ms: u64,
    hydrate_ms: f64,
    editor_mounted_ms: f64,
    content_ready_ms: f64,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MacOfflineOfficeEditorPrewarmDiagnosticInput {
    stage: String,
    editor_ready: bool,
    mathfield_hosts: u32,
    elapsed_ms: f64,
}

#[derive(Debug, Clone, Copy, Default)]
struct ResidentEditorFocusState {
    app_active: bool,
    window_can_become_key: bool,
    window_is_key: bool,
    window_is_main: bool,
    window_is_visible: bool,
    window_on_active_space: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct MacOfflineOfficeEditorReadyMarker {
    schema: &'static str,
    session_id: String,
    host: OfficeHost,
    generation: u64,
    epoch_ms: u64,
    url_received_epoch_ms: u64,
    frontend_epoch_ms: u64,
    hydrate_ms: f64,
    editor_mounted_ms: f64,
    content_ready_ms: f64,
    show_focus_ms: f64,
    window_focused: bool,
    window_visible: bool,
    app_active: bool,
    window_can_become_key: bool,
    window_is_main: bool,
}

fn office_editor_runtime() -> &'static Mutex<OfficeEditorRuntime> {
    static RUNTIME: OnceLock<Mutex<OfficeEditorRuntime>> = OnceLock::new();
    RUNTIME.get_or_init(|| Mutex::new(OfficeEditorRuntime::default()))
}

fn epoch_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis() as u64)
        .unwrap_or_default()
}

fn performance_logger() -> Option<&'static mpsc::Sender<OfficeEditorPerformanceRecord>> {
    static LOGGER: OnceLock<Option<mpsc::Sender<OfficeEditorPerformanceRecord>>> = OnceLock::new();
    LOGGER
        .get_or_init(|| {
            let (sender, receiver) = mpsc::channel::<OfficeEditorPerformanceRecord>();
            let spawned = std::thread::Builder::new()
                .name("visualtex-office-performance".to_string())
                .spawn(move || {
                    while let Ok(record) = receiver.recv() {
                        let Ok(directory) = session_directory(record.host, &record.session_id) else {
                            continue;
                        };
                        if fs::create_dir_all(&directory).is_err() {
                            continue;
                        }
                        let path = directory.join(EDITOR_PERFORMANCE_FILE);
                        let Ok(mut line) = serde_json::to_vec(&record) else {
                            continue;
                        };
                        line.push(b'\n');
                        if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(&path) {
                            if file.write_all(&line).is_ok() {
                                let _ = set_mode(&path, 0o600);
                            }
                        }
                    }
                });
            if spawned.is_ok() {
                Some(sender)
            } else {
                None
            }
        })
        .as_ref()
}

fn queue_editor_performance(
    host: OfficeHost,
    session_id: &str,
    stage: impl Into<String>,
    elapsed_ms: f64,
    generation: Option<u64>,
    details: Value,
) {
    let _ = performance_logger().and_then(|logger| {
        logger
            .send(OfficeEditorPerformanceRecord {
                schema: "visualtex-office-editor-performance-v1",
                session_id: session_id.to_string(),
                host,
                stage: stage.into(),
                epoch_ms: epoch_ms(),
                elapsed_ms,
                generation,
                details,
            })
            .ok()
    });
}

fn queue_editor_performance_at_epoch(
    host: OfficeHost,
    session_id: &str,
    stage: impl Into<String>,
    record_epoch_ms: u64,
    details: Value,
) {
    let _ = performance_logger().and_then(|logger| {
        logger
            .send(OfficeEditorPerformanceRecord {
                schema: "visualtex-office-editor-performance-v1",
                session_id: session_id.to_string(),
                host,
                stage: stage.into(),
                epoch_ms: record_epoch_ms,
                elapsed_ms: 0.0,
                generation: None,
                details,
            })
            .ok()
    });
}

fn user_home() -> Result<PathBuf, String> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .filter(|path| path.is_absolute())
        .ok_or_else(|| "Unable to resolve the current user's home directory".to_string())
}

pub(crate) fn runtime_root(host: OfficeHost) -> Result<PathBuf, String> {
    let suffix = match host {
        OfficeHost::Word => WORD_RUNTIME_SUFFIX,
        OfficeHost::Powerpoint => POWERPOINT_RUNTIME_SUFFIX,
    };
    Ok(user_home()?.join(suffix))
}

fn word_performance_trace_enabled() -> bool {
    runtime_root(OfficeHost::Word)
        .map(|root| root.join(WORD_PERFORMANCE_TRACE_SENTINEL).is_file())
        .unwrap_or(false)
}

fn host_from_request_name(value: &str) -> Result<OfficeHost, String> {
    match value {
        "word" => Ok(OfficeHost::Word),
        "powerpoint" => Ok(OfficeHost::Powerpoint),
        _ => Err("Offline Office request host must be word or powerpoint".to_string()),
    }
}

fn sessions_root(host: OfficeHost) -> Result<PathBuf, String> {
    Ok(runtime_root(host)?.join("OfficeSessions"))
}

fn ensure_runtime_root(host: OfficeHost) -> Result<PathBuf, String> {
    let root = runtime_root(host)?;
    let sessions = root.join("OfficeSessions");
    fs::create_dir_all(&sessions)
        .map_err(|error| format!("Unable to create {}: {error}", sessions.display()))?;
    set_mode(&root, 0o700)?;
    set_mode(&sessions, 0o700)?;
    Ok(root)
}

fn session_directory(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    validate_uuid(session_id, "Session id")?;
    Ok(sessions_root(host)?.join(session_id))
}

fn request_path(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(host, session_id)?.join(REQUEST_FILE))
}

fn dispatch_path(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(host, session_id)?.join(DISPATCH_FILE))
}

fn document_import_manifest_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?.join(DOCUMENT_IMPORT_MANIFEST_FILE))
}

fn latex_redraw_vector_batch_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?.join(LATEX_REDRAW_VECTOR_BATCH_FILE))
}

fn document_native_batch_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?.join(DOCUMENT_NATIVE_BATCH_FILE))
}

fn latex_redraw_preflight_manifest_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?
        .join(LATEX_REDRAW_PREFLIGHT_MANIFEST_FILE))
}

fn latex_redraw_font_sizes_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?.join(LATEX_REDRAW_FONT_SIZES_FILE))
}

fn formula_restore_source_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?.join(FORMULA_RESTORE_SOURCE_FILE))
}

fn result_png_path(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(host, session_id)?.join(RESULT_PNG_FILE))
}

fn result_svg_path(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(host, session_id)?.join(RESULT_SVG_FILE))
}

fn result_word_svg_docx_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?.join(RESULT_WORD_SVG_DOCX_FILE))
}

fn native_word_document_path(formula_id: &str) -> Result<PathBuf, String> {
    validate_uuid(formula_id, "Formula id")?;
    let directory = runtime_root(OfficeHost::Word)?.join("NativeDocuments");
    fs::create_dir_all(&directory)
        .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
    set_mode(&directory, 0o700)?;
    Ok(directory.join(format!("{formula_id}.docx")))
}

fn word_image_cache_paths(formula_id: &str) -> Result<(PathBuf, PathBuf, PathBuf), String> {
    validate_uuid(formula_id, "Formula id")?;
    let directory = runtime_root(OfficeHost::Word)?.join("ImageDocuments");
    fs::create_dir_all(&directory)
        .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
    set_mode(&directory, 0o700)?;
    Ok((
        directory.join(format!("{formula_id}.svg")),
        directory.join(format!("{formula_id}.docx")),
        directory.join(format!("{formula_id}.png")),
    ))
}

fn cleanup_session_files_at(
    directory: &Path,
    remove_document_formula_files: bool,
) -> Result<(), String> {
    for name in [
        REQUEST_FILE,
        DISPATCH_FILE,
        RESULT_PNG_FILE,
        RESULT_SVG_FILE,
        RESULT_WORD_SVG_DOCX_FILE,
        DOCUMENT_IMPORT_MANIFEST_FILE,
        LATEX_REDRAW_VECTOR_BATCH_FILE,
        LATEX_REDRAW_SOURCE_FILE,
        FORMULA_RESTORE_SOURCE_FILE,
        LATEX_REDRAW_PREFLIGHT_MANIFEST_FILE,
        LATEX_REDRAW_FONT_SIZES_FILE,
        EDITOR_OPEN_ERROR_FILE,
        "native-math.bin",
        "native-paragraph-context.bin",
        "native-context-status.txt",
        "native-context-paragraph.txt",
        "native-result.txt",
        "native-edit-copy-status.txt", "native-edit-current.bin",
        "native-edit-current.mathml", "native-edit-original.omml",
        "formula.docx",
    ] {
        let path = directory.join(name);
        match fs::remove_file(&path) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => return Err(format!("Unable to remove {}: {error}", path.display())),
        }
    }
    if remove_document_formula_files {
        if let Ok(entries) = fs::read_dir(directory) {
            for entry in entries.flatten() {
                let name = entry.file_name();
                let name = name.to_string_lossy();
                if name.starts_with("document-formula-") {
                    let path = entry.path();
                    if path.is_file() {
                        let _ = fs::remove_file(path);
                    }
                }
            }
        }
    }
    match fs::remove_dir(directory) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::DirectoryNotEmpty => Ok(()),
        Err(error) => Err(format!(
            "Unable to remove offline Office Session directory {}: {error}",
            directory.display()
        )),
    }
}

fn cleanup_session_files(host: OfficeHost, session_id: &str) -> Result<(), String> {
    cleanup_session_files_at(
        &session_directory(host, session_id)?,
        host == OfficeHost::Word,
    )
}

fn pointer_path(host: OfficeHost) -> Result<PathBuf, String> {
    Ok(sessions_root(host)?.join(match host {
        OfficeHost::Word => WORD_POINTER_FILE,
        OfficeHost::Powerpoint => POWERPOINT_POINTER_FILE,
    }))
}

#[tauri::command]
pub fn report_macos_offline_office_editor_prewarm_diagnostic(
    window: WebviewWindow,
    input: MacOfflineOfficeEditorPrewarmDiagnosticInput,
) -> Result<(), String> {
    let host = editor_window_host(window.label()).ok_or_else(|| {
        "Only a VisualTeX Office formula editor can report prewarm diagnostics".to_string()
    })?;
    if input.stage.is_empty()
        || input.stage.len() > 64
        || !input
            .stage
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-')
    {
        return Err("Office editor prewarm diagnostic stage is invalid".to_string());
    }
    if !input.elapsed_ms.is_finite() || !(0.0..=120_000.0).contains(&input.elapsed_ms) {
        return Err("Office editor prewarm diagnostic elapsedMs is invalid".to_string());
    }
    let diagnostic = serde_json::to_vec(&json!({
        "schema": "visualtex-office-editor-prewarm-diagnostic-v1",
        "host": office_host_name(host),
        "windowLabel": window.label(),
        "stage": input.stage,
        "editorReady": input.editor_ready,
        "mathfieldHosts": input.mathfield_hosts,
        "elapsedMs": input.elapsed_ms,
        "epochMs": epoch_ms(),
        "processId": std::process::id(),
    }))
    .map_err(|error| format!("Unable to encode Office editor prewarm diagnostic: {error}"))?;
    let diagnostic_path = editor_prewarm_diagnostic_path(window.app_handle(), host)?;
    atomic_write_runtime(&diagnostic_path, &diagnostic, 0o600)
}

#[tauri::command]
pub fn report_macos_offline_office_editor_prewarmed(window: WebviewWindow) -> Result<(), String> {
    let host = editor_window_host(window.label()).ok_or_else(|| {
        "Only a VisualTeX Office formula editor can report prewarming".to_string()
    })?;
    // Keep the prewarmed WKWebView continuously mounted. Only its native alpha
    // and mouse handling change, matching eb2fcf2a and avoiding a delayed wake.
    if active_editor_session(host).is_none() {
        set_resident_editor_parked(&window, true)?;
    }
    let marker = serde_json::to_vec(&json!({
        "schema": "visualtex-office-editor-prewarmed-v1",
        "host": office_host_name(host),
        "windowLabel": window.label(),
        "epochMs": epoch_ms(),
        "processId": std::process::id(),
    }))
    .map_err(|error| format!("Unable to encode Office editor prewarm marker: {error}"))?;
    let marker_path = editor_prewarmed_marker_path(window.app_handle(), host)?;
    atomic_write_runtime(&marker_path, &marker, 0o600)?;
    Ok(())
}

#[tauri::command]
pub fn get_macos_offline_office_editor_activation(
    window: WebviewWindow,
) -> Result<Option<MacOfflineOfficeEditorActivation>, String> {
    let host = editor_window_host(window.label())
        .ok_or_else(|| "Only a VisualTeX Office formula editor can query activation".to_string())?;
    Ok(active_editor_session(host).map(|active| active.activation))
}

#[tauri::command]
pub fn report_macos_offline_office_editor_ready(
    _app: AppHandle,
    window: WebviewWindow,
    input: MacOfflineOfficeEditorReadyInput,
) -> Result<(), String> {
    validate_uuid(&input.session_id, "Office editor Session id")?;
    for (value, label) in [
        (input.hydrate_ms, "hydrateMs"),
        (input.editor_mounted_ms, "editorMountedMs"),
        (input.content_ready_ms, "contentReadyMs"),
    ] {
        if !value.is_finite() || !(0.0..=120_000.0).contains(&value) {
            return Err(format!("Office editor {label} is invalid"));
        }
    }
    if input.hydrate_ms > input.editor_mounted_ms
        || input.editor_mounted_ms > input.content_ready_ms
    {
        return Err("Office editor readiness stages are out of order".to_string());
    }

    let host = editor_window_host(window.label())
        .ok_or_else(|| "Only a VisualTeX Office formula editor can report readiness".to_string())?;
    let mut runtime = office_editor_runtime()
        .lock()
        .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
    let active = runtime
        .active_mut(host)
        .as_mut()
        .ok_or_else(|| "The Office editor has no active Session".to_string())?;
    if active.activation.session_id != input.session_id
        || active.activation.generation != input.generation
    {
        return Err("Ignoring stale Office editor readiness".to_string());
    }
    active.ready = true;
    let active = active.clone();
    let silent = {
        let request = read_request(&input.session_id)?;
        request.agent || matches!(request.operation.as_deref(), Some("nativeToImage" | "imageToNative"))
    };
    let report_received_ms = active.received_at.elapsed().as_secs_f64() * 1000.0;
    let frontend_origin_ms = (report_received_ms - input.content_ready_ms).max(0.0);

    queue_editor_performance(
        host,
        &input.session_id,
        "frontend-hydrated",
        frontend_origin_ms + input.hydrate_ms,
        Some(input.generation),
        json!({ "durationMs": input.hydrate_ms }),
    );
    queue_editor_performance(
        host,
        &input.session_id,
        "frontend-editor-mounted",
        frontend_origin_ms + input.editor_mounted_ms,
        Some(input.generation),
        json!({ "durationMs": input.editor_mounted_ms }),
    );
    queue_editor_performance(
        host,
        &input.session_id,
        "frontend-content-ready",
        frontend_origin_ms + input.content_ready_ms,
        Some(input.generation),
        json!({
            "durationMs": input.content_ready_ms,
            "frontendEpochMs": input.frontend_epoch_ms,
        }),
    );

    let (focus, window_focused, window_visible) = if silent {
        set_resident_editor_content_visible(&window, false)?;
        set_resident_editor_parked(&window, true)?;
        (inspect_resident_editor_focus(&window)?, false, false)
    } else {
        set_resident_editor_content_visible(&window, true)?;
        configure_resident_editor_space_behavior(&window)?;
        center_resident_editor_on_pointer_screen(&window)?;
        order_resident_editor_to_front(&window)?;
        let focus = inspect_resident_editor_focus(&window)?;
        let window_focused = focus.window_is_key;
        let window_visible = focus.window_is_visible && focus.window_on_active_space;
        (focus, window_focused, window_visible)
    };
    let show_focus_ms = active.received_at.elapsed().as_secs_f64() * 1000.0;
    drop(runtime);
    let ready_epoch_ms = epoch_ms();
    queue_editor_performance(
        host,
        &input.session_id,
        "window-show-focus",
        show_focus_ms,
        Some(input.generation),
        json!({
            "silent": silent,
            "appActive": focus.app_active,
            "canBecomeKey": focus.window_can_become_key,
            "isKey": focus.window_is_key,
            "isMain": focus.window_is_main,
            "isVisible": focus.window_is_visible,
            "onActiveSpace": focus.window_on_active_space,
        }),
    );

    let marker = MacOfflineOfficeEditorReadyMarker {
        schema: "visualtex-office-editor-ready-v1",
        session_id: input.session_id.clone(),
        host,
        generation: input.generation,
        epoch_ms: ready_epoch_ms,
        url_received_epoch_ms: active.activation.received_epoch_ms,
        frontend_epoch_ms: input.frontend_epoch_ms,
        hydrate_ms: input.hydrate_ms,
        editor_mounted_ms: input.editor_mounted_ms,
        content_ready_ms: input.content_ready_ms,
        show_focus_ms,
        window_focused,
        window_visible,
        app_active: focus.app_active,
        window_can_become_key: focus.window_can_become_key,
        window_is_main: focus.window_is_main,
    };
    crate::spawn_background_task("visualtex-office-editor-ready", move || {
        let Ok(path) = session_directory(host, &input.session_id)
            .map(|directory| directory.join(EDITOR_READY_FILE))
        else {
            return;
        };
        let Ok(bytes) = serde_json::to_vec_pretty(&marker) else {
            return;
        };
        let _ = atomic_write(&path, &bytes, 0o600);
    });
    Ok(())
}

#[tauri::command]
pub fn present_macos_offline_office_editor_window(
    window: WebviewWindow,
    session_id: String,
    generation: u64,
) -> Result<(), String> {
    validate_uuid(&session_id, "Office editor Session id")?;
    let host = editor_window_host(window.label())
        .ok_or_else(|| "Only a VisualTeX Office formula editor can reveal itself".to_string())?;
    let active = active_editor_session(host)
        .ok_or_else(|| "The Office editor has no active Session".to_string())?;
    if active.activation.session_id != session_id || active.activation.generation != generation {
        return Err("The Office editor Session is no longer active".to_string());
    }
    set_resident_editor_content_visible(&window, true)?;
    present_resident_editor_window(window.app_handle(), &window)
}

#[tauri::command]
pub fn close_macos_offline_office_editor_window(
    window: WebviewWindow,
    session_id: Option<String>,
    generation: Option<u64>,
) -> Result<(), String> {
    let app = window.app_handle().clone();
    let Some(host) = editor_window_host(window.label()) else {
        if window.label().starts_with("office-native-document-") {
            window.destroy().map_err(|error| {
                format!("Unable to close the VisualTeX document importer: {error}")
            })?;
            #[cfg(target_os = "macos")]
            {
                let main_visible = app
                    .get_webview_window("main")
                    .and_then(|main| main.is_visible().ok())
                    .unwrap_or(false);
                if !has_open_office_editor(&app)
                    && !main_visible
                    && crate::office::background::is_background_mode()
                {
                    app.set_activation_policy(tauri::ActivationPolicy::Accessory)
                        .map_err(|error| {
                            format!("Unable to return VisualTeX to Office background mode: {error}")
                        })?;
                }
            }
            return Ok(());
        }
        return Err("Only a VisualTeX Office formula editor can close itself".to_string());
    };
    let session_id = session_id.ok_or_else(|| {
        "The Office formula editor close request is missing sessionId".to_string()
    })?;
    let generation = generation.ok_or_else(|| {
        "The Office formula editor close request is missing generation".to_string()
    })?;
    {
        // Keep activation validation and hiding atomic with respect to a new
        // URL activation. A late close from generation N must never hide the
        // already-hydrating generation N+1.
        let mut runtime = office_editor_runtime()
            .lock()
            .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
        runtime
            .active(host)
            .filter(|active| {
                active.activation.session_id == session_id
                    && active.activation.generation == generation
            })
            .ok_or_else(|| "The Office editor Session is no longer active".to_string())?;
        set_resident_editor_content_visible(&window, false)?;
        set_resident_editor_parked(&window, true)
            .map_err(|error| format!("Unable to close the VisualTeX Office editor: {error}"))?;
        #[cfg(target_os = "macos")]
        order_resident_editor_behind(&window)
            .map_err(|error| format!("Unable to close the VisualTeX Office editor: {error}"))?;
        *runtime.active_mut(host) = None;
    }
    let _ = window.emit(
        OFFICE_EDITOR_CLEAR_EVENT,
        json!({ "sessionId": session_id, "generation": generation }),
    );

    let parked_focus = inspect_resident_editor_focus(&window).ok();
    // Formula editing never changes application activation policy or Office
    // activation. Closing the overlay therefore only parks its own native window.
    queue_editor_performance(
        host,
        &session_id,
        "editor-visible-complete",
        0.0,
        Some(generation),
        json!({
            "parked": true,
            "officeFocusRequested": false,
            "isKeyAfterParking": parked_focus.map(|focus| focus.window_is_key),
            "isVisibleAfterParking": parked_focus.map(|focus| focus.window_is_visible),
            "onActiveSpaceAfterParking": parked_focus.map(|focus| focus.window_on_active_space),
        }),
    );
    Ok(())
}

/// A document window's message for the agent waiting for its session: true when an agent waits (the
/// window then closes instead of showing it), false when the user started the operation.
#[tauri::command]
pub fn note_macos_agent_session(session_id: String, note: String) -> bool {
    let watched = agent_sessions().lock().is_ok_and(|sessions| matches!(sessions.get(&session_id), Some(None)));
    if watched {
        if let Ok(mut notes) = agent_notes().lock() {
            notes.entry(session_id).or_default().push(note.chars().take(4000).collect());
        }
    }
    watched
}

/// What an agent asked the import window to insert (open_agent_import).
const AGENT_IMPORT_FILE: &str = "agent-import.json";

fn decode_png(value: &str) -> Result<Vec<u8>, String> {
    let payload = value
        .split_once(',')
        .filter(|(prefix, _)| prefix.starts_with("data:image/png;base64"))
        .map(|(_, payload)| payload)
        .unwrap_or(value);
    let bytes = BASE64_STANDARD
        .decode(payload.trim())
        .map_err(|error| format!("Unable to decode the Office PNG export: {error}"))?;
    if bytes.len() < 8 || &bytes[..8] != b"\x89PNG\r\n\x1a\n" {
        return Err("Office formula export is not a valid PNG image".to_string());
    }
    Ok(bytes)
}

#[cfg(unix)]
fn set_mode(path: &Path, mode: u32) -> Result<(), String> {
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(path, fs::Permissions::from_mode(mode))
        .map_err(|error| format!("Unable to set permissions on {}: {error}", path.display()))
}

#[cfg(not(unix))]
fn set_mode(_path: &Path, _mode: u32) -> Result<(), String> {
    Ok(())
}

fn atomic_write_with_durability(
    path: &Path,
    bytes: &[u8],
    mode: u32,
    durable: bool,
) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("Path has no parent: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    set_mode(parent, 0o700)?;
    let temporary = parent.join(format!(
        ".{}.{}.tmp",
        path.file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("visualtex"),
        Uuid::new_v4()
    ));
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&temporary)
        .map_err(|error| format!("Unable to create {}: {error}", temporary.display()))?;
    file.write_all(bytes)
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    if durable {
        file.sync_all()
            .map_err(|error| format!("Unable to sync {}: {error}", temporary.display()))?;
    }
    set_mode(&temporary, mode)?;
    fs::rename(&temporary, path).map_err(|error| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to replace {}: {error}", path.display())
    })?;
    set_mode(path, mode)
}

fn atomic_write(path: &Path, bytes: &[u8], mode: u32) -> Result<(), String> {
    atomic_write_with_durability(path, bytes, mode, true)
}

pub(crate) fn atomic_write_runtime(path: &Path, bytes: &[u8], mode: u32) -> Result<(), String> {
    // Session artifacts are consumed only after the atomic rename completes.
    // They do not need crash-durable fsync before a synchronous Office callback,
    // and skipping it removes several full storage barriers from every Apply.
    atomic_write_with_durability(path, bytes, mode, false)
}

#[cfg(target_os = "macos")]
thread_local! {
    // NSAppleScript is main-thread-bound in practice. Keep compiled callback
    // programs in a main-thread local cache instead of paying ~0.5 s to compile
    // the same Word/PowerPoint Apply script on every formula edit.
    static OFFICE_VBA_APPLESCRIPT_CACHE: RefCell<Vec<(String, Retained<NSAppleScript>)>> =
        const { RefCell::new(Vec::new()) };

}

const CALLBACK_ERROR_FILE: &str = "callback-error.txt";

#[tauri::command]
pub fn get_macos_offline_document_import_request(
    session_id: String,
) -> Result<MacOfflineDocumentImportPublicRequest, String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    document_import_request_data(&request)
}

#[tauri::command]
pub fn report_macos_offline_latex_redraw_stage(
    session_id: String,
    stage: String,
    elapsed_ms: f64,
    item_count: usize,
) -> Result<(), String> {
    validate_uuid(&session_id, "Session id")?;
    validate_bounded_text(&stage, 96, "LaTeX redraw timing stage")?;
    if !stage.starts_with("latex-redraw-")
        || !elapsed_ms.is_finite()
        || elapsed_ms < 0.0
        || elapsed_ms > 3_600_000.0
        || item_count > 1000
    {
        return Err("LaTeX redraw timing report is invalid".to_string());
    }
    queue_editor_performance(
        OfficeHost::Word,
        &session_id,
        &stage,
        elapsed_ms,
        None,
        json!({ "itemCount": item_count }),
    );
    Ok(())
}

#[tauri::command]
pub async fn resolve_macos_offline_latex_redraw_font_sizes(
    session_id: String,
    input: MacOfflineLatexRedrawFontQueryInput,
) -> Result<Vec<f64>, String> {
    tokio::task::spawn_blocking(move || {
        resolve_latex_redraw_font_sizes_blocking(session_id, input)
    })
    .await
    .map_err(|error| format!("Word LaTeX redraw font preflight task failed: {error}"))?
}

#[tauri::command]
pub fn focus_macos_offline_document_import_target(
    window: WebviewWindow,
    operation: String,
) -> Result<(), String> {
    if !window.label().starts_with("office-native-document-") {
        return Err("Only the VisualTeX document importer can focus Word".to_string());
    }
    if !matches!(
        operation.as_str(),
        "documentImport" | "latexRedraw" | "formulaRestore"
    ) {
        return Err("Word document operation is invalid".to_string());
    }
    // Telling Word would start Word for a WPS user: WPS comes forward instead,
    // after the import window (a redraw's window is never shown).
    if document_window_is_wps(&window) {
        let shown = window.is_visible().unwrap_or(false);
        window
            .hide()
            .map_err(|error| format!("Unable to hide the VisualTeX document importer: {error}"))?;
        if shown {
            super::wps::installer::activate_running_wps();
        }
        return Ok(());
    }
    if document_window_is_powerpoint(&window) {
        window
            .hide()
            .map_err(|error| format!("Unable to hide the VisualTeX document importer: {error}"))?;
        let _ = Command::new("/usr/bin/osascript")
            .args(["-e", "tell application \"Microsoft PowerPoint\" to activate"])
            .output();
        return Ok(());
    }
    set_word_document_operation_preparing_status(&operation)?;
    window
        .hide()
        .map_err(|error| format!("Unable to hide the VisualTeX document importer: {error}"))
}

#[tauri::command]
pub fn restore_macos_offline_document_import_window(window: WebviewWindow) -> Result<(), String> {
    if !window.label().starts_with("office-native-document-") {
        return Err("Only the VisualTeX document importer can restore itself".to_string());
    }
    if !document_window_is_wps(&window) && !document_window_is_powerpoint(&window) {
        clear_word_document_import_status();
    }
    let app = window.app_handle().clone();
    crate::office::background::activate_foreground_app(&app)?;
    window.show().map_err(|error| error.to_string())?;
    window.unminimize().map_err(|error| error.to_string())?;
    window.set_focus().map_err(|error| error.to_string())
}

#[tauri::command]
pub fn get_macos_offline_document_import_progress(
    session_id: String,
) -> Result<MacOfflineDocumentImportProgress, String> {
    validate_uuid(&session_id, "Session id")?;
    let host = read_request(&session_id)
        .ok()
        .and_then(|request| host_from_request_name(&request.host).ok())
        .unwrap_or(OfficeHost::Word);
    let path = session_directory(host, &session_id)?.join(DOCUMENT_IMPORT_PROGRESS_FILE);
    let source = match fs::read_to_string(path) {
        Ok(value) => value,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            return Ok(MacOfflineDocumentImportProgress {
                current: 0,
                total: 0,
                stage: "preparing".to_string(),
            });
        }
        Err(error) => return Err(format!("Unable to read document import progress: {error}")),
    };
    let mut current = None;
    let mut total = None;
    let mut stage = None;
    for line in source.lines() {
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        match key {
            "current" => current = value.parse::<usize>().ok(),
            "total" => total = value.parse::<usize>().ok(),
            "stage" => stage = Some(value.to_string()),
            _ => {}
        }
    }
    let current =
        current.ok_or_else(|| "Document import progress is missing current".to_string())?;
    let total = total.ok_or_else(|| "Document import progress is missing total".to_string())?;
    let stage = stage.ok_or_else(|| "Document import progress is missing stage".to_string())?;
    if total > 2048 || current > total || stage.len() > 32 || stage.chars().any(char::is_control) {
        return Err("Document import progress is invalid".to_string());
    }
    Ok(MacOfflineDocumentImportProgress {
        current,
        total,
        stage,
    })
}

#[tauri::command]
pub async fn commit_macos_offline_document_import(
    session_id: String,
    input: MacOfflineDocumentImportCommitInput,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<(), String> {
    let state = state.inner().clone();
    let id = session_id.clone();
    let committed = tokio::task::spawn_blocking(move || commit_document_import_blocking(state, session_id, input))
        .await
        .map_err(|error| format!("Offline document import task failed: {error}"))
        .and_then(|result| result);
    finish_agent_session(&id, committed.clone());
    committed
}

#[tauri::command]
pub async fn cancel_macos_offline_document_import(session_id: String) -> Result<(), String> {
    let notes = take_agent_session_notes(&session_id);
    let reason = if notes.is_empty() { "The operation was cancelled".to_string() } else { notes.join(" ") };
    finish_agent_session(&session_id, Err(reason));
    tokio::task::spawn_blocking(move || cancel_document_import_blocking(session_id))
        .await
        .map_err(|error| format!("Offline document import cancel task failed: {error}"))?
}

#[tauri::command]
pub fn get_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn update_macos_offline_office_session(
    session_id: String,
    patch: Value,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    state
        .session_store
        .patch(&session_id, patch)
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn delete_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<(), String> {
    validate_uuid(&session_id, "Session id")?;
    let session = state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())?;
    state
        .session_store
        .delete(&session_id)
        .map_err(|error| error.to_string())?;
    cleanup_session_files(session.host, &session_id)
}

#[tauri::command]
pub async fn commit_macos_offline_office_session(
    session_id: String,
    patch: Option<Value>,
    apply_started_epoch_ms: Option<u64>,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || {
        validate_uuid(&session_id, "Session id")?;
        if let Some(started_epoch_ms) = apply_started_epoch_ms {
            let now = epoch_ms();
            if started_epoch_ms > 0
                && started_epoch_ms <= now
                && now - started_epoch_ms <= 60_000
            {
                let session = state
                    .session_store
                    .get(&session_id)
                    .map_err(|error| error.to_string())?;
                queue_editor_performance_at_epoch(
                    session.host,
                    &session_id,
                    "apply-ui-start",
                    started_epoch_ms,
                    json!({}),
                );
            }
        }
        if let Some(patch) = patch {
            state
                .session_store
                .patch(&session_id, patch)
                .map_err(|error| error.to_string())?;
        }
        commit_session_blocking(state, session_id)
    })
    .await
    .map_err(|error| format!("Offline Office commit task failed: {error}"))?
}

#[tauri::command]
pub async fn cancel_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || cancel_session_blocking(state, session_id))
        .await
        .map_err(|error| format!("Offline Office cancel task failed: {error}"))?
}

#[tauri::command]
pub fn get_macos_offline_plugin_health() -> Result<Vec<MacOfflinePluginHealth>, String> {
    let _ = refresh_health_signal("word");
    let _ = refresh_health_signal("powerpoint");
    Ok(vec![read_health("word")?, read_health("powerpoint")?])
}

#[cfg(test)]
mod tests;
