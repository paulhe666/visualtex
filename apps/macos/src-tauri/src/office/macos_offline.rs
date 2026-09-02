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
use objc2_app_kit::NSApplication;
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
    atomic::{AtomicU64, Ordering},
    mpsc, Mutex, OnceLock,
};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tauri::utils::config::BackgroundThrottlingPolicy;
use tauri::{
    AppHandle, Emitter, Manager, Monitor, WebviewUrl, WebviewWindow, WebviewWindowBuilder,
};
use uuid::Uuid;

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
const WORD_TEX_SHALLOW_DESCENT_FLOOR_PT: f64 = 1.91;
const WORD_TIMES_IMAGE_WIDTH_SCALE: f64 = 1.067;
const WORD_TIMES_IMAGE_HEIGHT_SCALE: f64 = 1.0;

fn word_image_visual_scales(formula_letter_font: Option<&str>) -> (f64, f64) {
    match formula_letter_font {
        Some("times") => (WORD_TIMES_IMAGE_WIDTH_SCALE, WORD_TIMES_IMAGE_HEIGHT_SCALE),
        _ => (WORD_TEX_IMAGE_VISUAL_SCALE, WORD_TEX_IMAGE_VISUAL_SCALE),
    }
}
const MIN_WORD_FONT_SIZE_PT: f64 = 1.0;
const MAX_WORD_FONT_SIZE_PT: f64 = 512.0;
const POWERPOINT_REFERENCE_FONT_SIZE_PT: f64 = 14.0;
const DEFAULT_POWERPOINT_FONT_SIZE_PT: f64 = 18.0;
const MIN_POWERPOINT_FONT_SIZE_PT: f64 = 1.0;
const MAX_POWERPOINT_FONT_SIZE_PT: f64 = 512.0;
static WORD_DISPATCH_LOCK: Mutex<()> = Mutex::new(());
static POWERPOINT_DISPATCH_LOCK: Mutex<()> = Mutex::new(());
static OFFICE_EDITOR_SIZE_WRITE_GENERATION: AtomicU64 = AtomicU64::new(0);
static FAST_OPEN_WATCHER_STARTED: OnceLock<()> = OnceLock::new();

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
    presentation_identity: String,
    slide_index: u32,
    slide_id: u32,
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
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflineDocumentImportRequest {
    bookmark_name: String,
    default_font_size_pt: f64,
    #[serde(default)]
    redraw_scope: Option<String>,
    #[serde(default)]
    output_kind: Option<String>,
    #[serde(default)]
    source_kind: Option<String>,
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
    formula_id: Option<String>,
    display_mode: String,
    numbered: bool,
    #[serde(default)]
    native_equation: bool,
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
    power_point: Option<MacOfflinePowerPointRequest>,
    #[serde(default)]
    document_import: Option<MacOfflineDocumentImportRequest>,
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
        omml_base64: String,
        omml_docx_base64: String,
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

fn performance_logger() -> &'static mpsc::Sender<OfficeEditorPerformanceRecord> {
    static LOGGER: OnceLock<mpsc::Sender<OfficeEditorPerformanceRecord>> = OnceLock::new();
    LOGGER.get_or_init(|| {
        let (sender, receiver) = mpsc::channel::<OfficeEditorPerformanceRecord>();
        std::thread::Builder::new()
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
            })
            .expect("VisualTeX Office performance logger thread must start");
        sender
    })
}

fn queue_editor_performance(
    host: OfficeHost,
    session_id: &str,
    stage: impl Into<String>,
    elapsed_ms: f64,
    generation: Option<u64>,
    details: Value,
) {
    let _ = performance_logger().send(OfficeEditorPerformanceRecord {
        schema: "visualtex-office-editor-performance-v1",
        session_id: session_id.to_string(),
        host,
        stage: stage.into(),
        epoch_ms: epoch_ms(),
        elapsed_ms,
        generation,
        details,
    });
}

fn queue_editor_performance_at_epoch(
    host: OfficeHost,
    session_id: &str,
    stage: impl Into<String>,
    record_epoch_ms: u64,
    details: Value,
) {
    let _ = performance_logger().send(OfficeEditorPerformanceRecord {
        schema: "visualtex-office-editor-performance-v1",
        session_id: session_id.to_string(),
        host,
        stage: stage.into(),
        epoch_ms: record_epoch_ms,
        elapsed_ms: 0.0,
        generation: None,
        details,
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

fn fast_open_inbox_roots_for_home(home: &Path) -> Vec<(OfficeHost, PathBuf)> {
    vec![
        (
            OfficeHost::Word,
            home.join(WORD_FAST_OPEN_INBOX_SUFFIX),
        ),
        (
            OfficeHost::Powerpoint,
            home.join(POWERPOINT_FAST_OPEN_INBOX_SUFFIX),
        ),
    ]
}

fn fast_open_inbox_roots() -> Result<Vec<(OfficeHost, PathBuf)>, String> {
    Ok(fast_open_inbox_roots_for_home(&user_home()?))
}

fn fast_open_session_id(path: &Path) -> Option<String> {
    let file_name = path.file_name()?.to_str()?;
    let session_id = file_name.strip_suffix(".json")?;
    if validate_uuid(session_id, "Fast-open Session id").is_ok() {
        Some(session_id.to_string())
    } else {
        None
    }
}

fn fast_open_modified_is_recent(modified: SystemTime, now: SystemTime) -> bool {
    match now.duration_since(modified) {
        Ok(age) => (FAST_OPEN_MIN_STABLE_AGE..=FAST_OPEN_MAX_AGE).contains(&age),
        Err(error) => error.duration() <= FAST_OPEN_FUTURE_TOLERANCE,
    }
}

fn persist_fast_open_claim(
    expected_host: OfficeHost,
    session_id: &str,
    claim_path: &Path,
) -> Result<(), String> {
    let metadata = fs::symlink_metadata(claim_path).map_err(|error| {
        format!(
            "Unable to inspect claimed Office fast-open request {}: {error}",
            claim_path.display()
        )
    })?;
    if metadata.file_type().is_symlink()
        || !metadata.is_file()
        || metadata.len() == 0
        || metadata.len() > MAX_REQUEST_BYTES
    {
        return Err("Office fast-open request has an invalid file type or size".to_string());
    }
    let bytes = fs::read(claim_path).map_err(|error| {
        format!(
            "Unable to read claimed Office fast-open request {}: {error}",
            claim_path.display()
        )
    })?;
    let request: MacOfflineSessionRequest = serde_json::from_slice(&bytes)
        .map_err(|error| format!("Office fast-open request contains invalid JSON: {error}"))?;
    validate_request(&request, session_id)?;
    if host_from_request_name(&request.host)? != expected_host {
        return Err("Office fast-open request host does not match its sandbox inbox".to_string());
    }
    if request.operation.as_deref().unwrap_or("formula") != "formula" {
        return Err("Office fast-open accepts only ordinary formula requests".to_string());
    }

    ensure_runtime_root(expected_host)?;
    atomic_write_runtime(&request_path(expected_host, session_id)?, &bytes, 0o600)
}

pub(crate) fn consume_fast_open_request(app: &AppHandle) -> Result<bool, String> {
    let now = SystemTime::now();
    let mut candidates = Vec::new();
    'host_roots: for (host, root) in fast_open_inbox_roots()? {
        let root_metadata = loop {
            match fs::symlink_metadata(&root) {
                Ok(metadata) => break metadata,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue 'host_roots,
                Err(error) => {
                    return Err(format!(
                        "Unable to inspect Office fast-open inbox {}: {error}",
                        root.display()
                    ))
                }
            }
        };
        if root_metadata.file_type().is_symlink() || !root_metadata.is_dir() {
            return Err(format!(
                "Office fast-open inbox is not a real directory: {}",
                root.display()
            ));
        }
        let entries = loop {
            match fs::read_dir(&root) {
                Ok(entries) => break entries,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) => {
                    return Err(format!(
                        "Unable to enumerate Office fast-open inbox {}: {error}",
                        root.display()
                    ))
                }
            }
        };
        'entries: for entry in entries {
            let entry = match entry {
                Ok(entry) => entry,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) => {
                    return Err(format!(
                        "Unable to enumerate Office fast-open inbox {}: {error}",
                        root.display()
                    ))
                }
            };
            let path = entry.path();
            let Some(session_id) = fast_open_session_id(&path) else {
                continue;
            };
            let metadata = loop {
                match fs::symlink_metadata(&path) {
                    Ok(metadata) => break metadata,
                    Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                    Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue 'entries,
                    Err(error) => {
                        return Err(format!(
                            "Unable to inspect Office fast-open request {}: {error}",
                            path.display()
                        ))
                    }
                }
            };
            if metadata.file_type().is_symlink()
                || !metadata.is_file()
                || metadata.len() == 0
                || metadata.len() > MAX_REQUEST_BYTES
            {
                continue;
            }
            let modified = metadata.modified().map_err(|error| {
                format!(
                    "Unable to read Office fast-open request timestamp {}: {error}",
                    path.display()
                )
            })?;
            if !fast_open_modified_is_recent(modified, now) {
                continue;
            }
            candidates.push((modified, host, path, session_id));
        }
    }

    candidates.sort_by(|left, right| right.0.cmp(&left.0));
    for (_modified, host, path, session_id) in candidates {
        let parent = path
            .parent()
            .ok_or_else(|| "Office fast-open request has no inbox parent".to_string())?;
        let claim_path = parent.join(format!(
            ".{session_id}.{}.claim",
            Uuid::new_v4()
        ));
        match fs::rename(&path, &claim_path) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue,
            Err(error) => {
                return Err(format!(
                    "Unable to claim Office fast-open request {}: {error}",
                    path.display()
                ))
            }
        }

        let persist_result = persist_fast_open_claim(host, &session_id, &claim_path);
        let _ = fs::remove_file(&claim_path);
        persist_result?;
        let url = format!("visualtex://office/open?session={session_id}");
        handle_open_url(app, &url)?;
        return Ok(true);
    }
    Ok(false)
}

fn refresh_fast_open_ready_markers() -> Result<(), String> {
    let heartbeat = format!("visualtex-fast-open-ready-v1\n{}\n", epoch_ms());
    for (_host, root) in fast_open_inbox_roots()? {
        fs::create_dir_all(&root).map_err(|error| {
            format!(
                "Unable to create Office fast-open inbox {}: {error}",
                root.display()
            )
        })?;
        set_mode(&root, 0o700)?;
        atomic_write_runtime(&root.join(FAST_OPEN_READY_FILE), heartbeat.as_bytes(), 0o600)?;
    }
    Ok(())
}

pub(crate) fn start_fast_open_inbox_watcher(app: AppHandle) {
    if FAST_OPEN_WATCHER_STARTED.set(()).is_err() {
        return;
    }
    std::thread::spawn(move || {
        let mut last_heartbeat = Instant::now()
            .checked_sub(FAST_OPEN_READY_HEARTBEAT_INTERVAL)
            .unwrap_or_else(Instant::now);
        loop {
            if last_heartbeat.elapsed() >= FAST_OPEN_READY_HEARTBEAT_INTERVAL {
                if let Err(error) = refresh_fast_open_ready_markers() {
                    eprintln!("Unable to refresh VisualTeX Office fast-open readiness: {error}");
                }
                last_heartbeat = Instant::now();
            }
            if let Err(error) = consume_fast_open_request(&app) {
                eprintln!("Unable to consume VisualTeX Office fast-open request: {error}");
                std::thread::sleep(Duration::from_millis(250));
                continue;
            }
            std::thread::sleep(FAST_OPEN_POLL_INTERVAL);
        }
    });
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

fn paeth_predictor(left: u8, up: u8, upper_left: u8) -> u8 {
    let left = i32::from(left);
    let up = i32::from(up);
    let upper_left = i32::from(upper_left);
    let estimate = left + up - upper_left;
    let left_distance = (estimate - left).abs();
    let up_distance = (estimate - up).abs();
    let upper_left_distance = (estimate - upper_left).abs();
    if left_distance <= up_distance && left_distance <= upper_left_distance {
        left as u8
    } else if up_distance <= upper_left_distance {
        up as u8
    } else {
        upper_left as u8
    }
}

fn png_ink_center_y_ratio_from_bytes(bytes: &[u8]) -> Result<f64, String> {
    const SIGNATURE: &[u8; 8] = b"\x89PNG\r\n\x1a\n";
    if bytes.len() < SIGNATURE.len() || &bytes[..8] != SIGNATURE {
        return Err("Cached formula image is not a PNG file".to_string());
    }

    let mut offset = 8usize;
    let mut width = 0usize;
    let mut height = 0usize;
    let mut bit_depth = 0u8;
    let mut color_type = 0u8;
    let mut interlace = 0u8;
    let mut compressed = Vec::new();
    while offset + 12 <= bytes.len() {
        let length = u32::from_be_bytes(bytes[offset..offset + 4].try_into().unwrap()) as usize;
        let chunk_start = offset + 8;
        let chunk_end = chunk_start
            .checked_add(length)
            .ok_or_else(|| "PNG chunk length overflowed".to_string())?;
        let crc_end = chunk_end
            .checked_add(4)
            .ok_or_else(|| "PNG chunk boundary overflowed".to_string())?;
        if crc_end > bytes.len() {
            return Err("Cached formula PNG has a truncated chunk".to_string());
        }
        let chunk_type = &bytes[offset + 4..offset + 8];
        match chunk_type {
            b"IHDR" => {
                if length != 13 {
                    return Err("Cached formula PNG has an invalid IHDR".to_string());
                }
                width = u32::from_be_bytes(bytes[chunk_start..chunk_start + 4].try_into().unwrap()) as usize;
                height = u32::from_be_bytes(bytes[chunk_start + 4..chunk_start + 8].try_into().unwrap()) as usize;
                bit_depth = bytes[chunk_start + 8];
                color_type = bytes[chunk_start + 9];
                if bytes[chunk_start + 10] != 0 || bytes[chunk_start + 11] != 0 {
                    return Err("Cached formula PNG uses unsupported compression/filter methods".to_string());
                }
                interlace = bytes[chunk_start + 12];
            }
            b"IDAT" => compressed.extend_from_slice(&bytes[chunk_start..chunk_end]),
            b"IEND" => break,
            _ => {}
        }
        offset = crc_end;
    }

    if width == 0 || height == 0 || width > 200_000 || height > 200_000 {
        return Err("Cached formula PNG has invalid dimensions".to_string());
    }
    if bit_depth != 8 || !matches!(color_type, 4 | 6) || interlace != 0 {
        return Err("Cached formula PNG is not a supported 8-bit alpha image".to_string());
    }
    let bytes_per_pixel = if color_type == 6 { 4usize } else { 2usize };
    let stride = width
        .checked_mul(bytes_per_pixel)
        .ok_or_else(|| "Cached formula PNG row is too large".to_string())?;
    let expected = height
        .checked_mul(stride + 1)
        .ok_or_else(|| "Cached formula PNG is too large".to_string())?;
    if expected > 512 * 1024 * 1024 {
        return Err("Cached formula PNG exceeds the decode limit".to_string());
    }

    let mut inflated = Vec::with_capacity(expected);
    ZlibDecoder::new(compressed.as_slice())
        .read_to_end(&mut inflated)
        .map_err(|error| format!("Unable to inflate cached formula PNG: {error}"))?;
    if inflated.len() != expected {
        return Err("Cached formula PNG has an unexpected decoded size".to_string());
    }

    let mut previous = vec![0u8; stride];
    let mut current = vec![0u8; stride];
    let mut minimum_y = None::<usize>;
    let mut maximum_y = 0usize;
    for y in 0..height {
        let row_start = y * (stride + 1);
        let filter = inflated[row_start];
        let source = &inflated[row_start + 1..row_start + 1 + stride];
        for index in 0..stride {
            let left = if index >= bytes_per_pixel {
                current[index - bytes_per_pixel]
            } else {
                0
            };
            let up = previous[index];
            let upper_left = if index >= bytes_per_pixel {
                previous[index - bytes_per_pixel]
            } else {
                0
            };
            current[index] = match filter {
                0 => source[index],
                1 => source[index].wrapping_add(left),
                2 => source[index].wrapping_add(up),
                3 => source[index].wrapping_add(((u16::from(left) + u16::from(up)) / 2) as u8),
                4 => source[index].wrapping_add(paeth_predictor(left, up, upper_left)),
                _ => return Err("Cached formula PNG uses an unsupported row filter".to_string()),
            };
        }
        let alpha_offset = bytes_per_pixel - 1;
        let row_has_ink = current
            .chunks_exact(bytes_per_pixel)
            .any(|pixel| pixel[alpha_offset] >= 16);
        if row_has_ink {
            minimum_y.get_or_insert(y);
            maximum_y = y;
        }
        std::mem::swap(&mut previous, &mut current);
        current.fill(0);
    }

    let minimum_y = minimum_y.ok_or_else(|| "Cached formula PNG contains no visible ink".to_string())?;
    let painted_center = (minimum_y + maximum_y + 1) as f64 / 2.0;
    let ratio = painted_center / height as f64;
    if !ratio.is_finite() || !(0.0..=1.0).contains(&ratio) {
        return Err("Cached formula PNG produced an invalid ink center".to_string());
    }
    Ok(ratio)
}

fn cached_formula_ink_center_y_ratio(formula_id: &str) -> Result<f64, String> {
    let (_, _, png_path) = word_image_cache_paths(formula_id)?;
    let metadata = fs::symlink_metadata(&png_path)
        .map_err(|error| format!("Unable to inspect cached formula PNG {}: {error}", png_path.display()))?;
    if metadata.file_type().is_symlink() || !metadata.is_file() || metadata.len() == 0 || metadata.len() > 128 * 1024 * 1024 {
        return Err("Cached formula PNG has an invalid file type or size".to_string());
    }
    let bytes = fs::read(&png_path)
        .map_err(|error| format!("Unable to read cached formula PNG {}: {error}", png_path.display()))?;
    png_ink_center_y_ratio_from_bytes(&bytes)
}

pub fn run_image_ink_center_cli_if_requested() -> Option<i32> {
    let arguments = std::env::args().collect::<Vec<_>>();
    let Some(index) = arguments.iter().position(|argument| argument == IMAGE_INK_CENTER_CLI_ARGUMENT) else {
        return None;
    };
    let formula_id = arguments.get(index + 1).map(String::as_str).unwrap_or("");
    match cached_formula_ink_center_y_ratio(formula_id) {
        Ok(ratio) => {
            println!("{ratio:.9}");
            Some(0)
        }
        Err(error) => {
            eprintln!("{error}");
            Some(2)
        }
    }
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

fn validate_uuid(value: &str, label: &str) -> Result<(), String> {
    if valid_uuid(value) {
        Ok(())
    } else {
        Err(format!("{label} must be a canonical UUID"))
    }
}

fn validate_bounded_text(value: &str, maximum: usize, label: &str) -> Result<(), String> {
    if value.chars().count() > maximum || value.chars().any(char::is_control) {
        return Err(format!(
            "{label} contains unsupported characters or is too long"
        ));
    }
    Ok(())
}

fn validate_finite_geometry(value: f64, label: &str) -> Result<(), String> {
    if !value.is_finite() || value.abs() > 10_000_000.0 {
        Err(format!("PowerPoint {label} is invalid"))
    } else {
        Ok(())
    }
}

fn validate_latex_redraw_source_size(byte_len: u64) -> Result<(), String> {
    if byte_len == 0 || byte_len > MAX_LATEX_REDRAW_SOURCE_BYTES {
        Err("LaTeX redraw source must contain 1 byte to 5 MB".to_string())
    } else {
        Ok(())
    }
}

fn validate_request(request: &MacOfflineSessionRequest, session_id: &str) -> Result<(), String> {
    if request.protocol_version != OFFLINE_PROTOCOL_VERSION {
        return Err("Unsupported VisualTeX macOS offline protocol version".to_string());
    }
    validate_uuid(&request.session_id, "Request Session id")?;
    if request.session_id != session_id {
        return Err("Request Session id does not match the custom URL".to_string());
    }
    if !matches!(request.host.as_str(), "word" | "powerpoint") {
        return Err("Offline Office request host must be word or powerpoint".to_string());
    }
    if !matches!(request.mode.as_str(), "create" | "edit") {
        return Err("Offline Office request mode must be create or edit".to_string());
    }
    let operation = request.operation.as_deref().unwrap_or("formula");
    if matches!(operation, "documentImport" | "latexRedraw" | "formulaRestore") {
        if request.host != "word" || request.mode != "create" {
            return Err("Document import is supported only as a new Word operation".to_string());
        }
        let document_import = request
            .document_import
            .as_ref()
            .ok_or_else(|| "Document import request is missing its insertion anchor".to_string())?;
        let source_document_id = request.source_document_id.as_deref().ok_or_else(|| {
            "Document import request is missing the Word document identity".to_string()
        })?;
        validate_bounded_text(source_document_id, MAX_IDENTITY_CHARS, "sourceDocumentId")?;
        validate_bounded_text(
            &document_import.bookmark_name,
            40,
            "documentImport bookmarkName",
        )?;
        if !document_import.bookmark_name.starts_with("VT_D_") {
            return Err("Document import bookmark name is invalid".to_string());
        }
        if !document_import.default_font_size_pt.is_finite()
            || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT)
                .contains(&document_import.default_font_size_pt)
        {
            return Err(
                "Document import default font size is outside the supported range".to_string(),
            );
        }
        if operation == "latexRedraw" {
            if !matches!(
                document_import.redraw_scope.as_deref(),
                Some("selection" | "document")
            ) {
                return Err("LaTeX redraw scope must be selection or document".to_string());
            }
            if !matches!(
                document_import.output_kind.as_deref(),
                Some("omml" | "image")
            ) {
                return Err("LaTeX redraw output kind must be omml or image".to_string());
            }
            if document_import.source_kind.is_some() {
                return Err("LaTeX redraw request contains a formula restore source kind".to_string());
            }
            let source_path =
                session_directory(OfficeHost::Word, session_id)?.join(LATEX_REDRAW_SOURCE_FILE);
            let source_metadata = fs::metadata(&source_path)
                .map_err(|error| format!("Unable to read {}: {error}", source_path.display()))?;
            validate_latex_redraw_source_size(source_metadata.len())?;
        } else if operation == "formulaRestore" {
            if !matches!(
                document_import.redraw_scope.as_deref(),
                Some("selection" | "document")
            ) {
                return Err("Formula restore scope must be selection or document".to_string());
            }
            if !matches!(document_import.output_kind.as_deref(), Some("latex" | "image")) {
                return Err("Formula restore output kind must be latex or image".to_string());
            }
            if !matches!(document_import.source_kind.as_deref(), Some("omml" | "image")) {
                return Err("Formula restore source kind must be omml or image".to_string());
            }
            if document_import.output_kind.as_deref() == Some("image")
                && document_import.source_kind.as_deref() != Some("omml")
            {
                return Err("Only native OMML can be restored to a VisualTeX image".to_string());
            }
            let source_metadata = fs::metadata(formula_restore_source_path(session_id)?)
                .map_err(|error| format!("Unable to inspect formula restore source: {error}"))?;
            if source_metadata.file_type().is_symlink()
                || !source_metadata.is_file()
                || source_metadata.len() == 0
                || source_metadata.len() as usize > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES
            {
                return Err("Formula restore source has an invalid size".to_string());
            }
        } else if document_import.redraw_scope.is_some()
            || document_import.output_kind.is_some()
            || document_import.source_kind.is_some()
        {
            return Err("Document import request contains formula transform fields".to_string());
        }
        if request.formula_id.is_some()
            || request.encoded_metadata.is_some()
            || request.pending_marker.is_some()
            || request.power_point.is_some()
            || request.document_import.is_none()
            || request.numbered
            || request.native_equation
        {
            return Err("Document import request contains formula-only fields".to_string());
        }
        return Ok(());
    }
    if !matches!(operation, "formula" | "nativeToImage" | "imageToNative")
        || request.document_import.is_some()
    {
        return Err("Unsupported offline Office operation".to_string());
    }
    if matches!(operation, "nativeToImage" | "imageToNative") {
        let output_matches_operation = if operation == "nativeToImage" {
            !request.native_equation
        } else {
            request.native_equation
        };
        if request.host != "word" || request.mode != "edit" || !output_matches_operation {
            return Err(format!(
                "{operation} conversion requires a matching Word edit output request"
            ));
        }
        if request.formula_id.is_none()
            || request.source_document_id.is_none()
            || request.source_object_id.is_none()
            || request.encoded_metadata.is_none()
            || request.pending_marker.is_some()
            || request.power_point.is_some()
        {
            return Err(format!(
                "{operation} conversion is missing its existing Word formula identity"
            ));
        }
    }
    if !matches!(request.display_mode.as_str(), "inline" | "block") {
        return Err("Offline Office displayMode must be inline or block".to_string());
    }
    if request.numbered && (request.host != "word" || request.display_mode != "block") {
        return Err("Only Word display formulas can be numbered".to_string());
    }
    if request.native_equation && request.host != "word" {
        return Err("Native equations are supported only by Word requests".to_string());
    }
    if let Some(formula_id) = request.formula_id.as_deref() {
        validate_uuid(formula_id, "Formula id")?;
    }
    for (value, label) in [
        (request.source_document_id.as_deref(), "sourceDocumentId"),
        (request.source_object_id.as_deref(), "sourceObjectId"),
        (request.pending_marker.as_deref(), "pendingMarker"),
    ] {
        if let Some(value) = value {
            validate_bounded_text(value, MAX_IDENTITY_CHARS, label)?;
        }
    }
    if let Some(marker) = request.pending_marker.as_deref() {
        if !marker.starts_with(PENDING_PREFIX) {
            return Err("Offline Office pending marker is invalid".to_string());
        }
    }
    if let Some(encoded) = request.encoded_metadata.as_deref() {
        if encoded.len() > MAX_METADATA_BYTES || !encoded.starts_with(METADATA_PREFIX) {
            return Err("Offline Office metadata envelope is invalid".to_string());
        }
    }
    for (value, label) in [
        (request.font_size_pt, "fontSizePt"),
        (request.reference_width_pt, "referenceWidthPt"),
        (request.reference_height_pt, "referenceHeightPt"),
    ] {
        if let Some(value) = value {
            if !value.is_finite() || value <= 0.0 {
                return Err(format!(
                    "Offline Office {label} must be a positive finite number"
                ));
            }
        }
    }
    if let Some(font_size) = request.font_size_pt {
        if !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&font_size) {
            return Err(
                "Offline Office Word fontSizePt is outside the supported range".to_string(),
            );
        }
    }
    if request.host != "word"
        && (request.font_size_pt.is_some()
            || request.reference_width_pt.is_some()
            || request.reference_height_pt.is_some())
    {
        return Err("PowerPoint requests must not contain Word font-size metadata".to_string());
    }

    match (request.host.as_str(), request.power_point.as_ref()) {
        ("word", None) => {}
        ("word", Some(_)) => {
            return Err("Word request must not contain PowerPoint geometry".to_string())
        }
        ("powerpoint", None) => return Err("PowerPoint request requires geometry".to_string()),
        ("powerpoint", Some(powerpoint)) => {
            validate_bounded_text(
                &powerpoint.presentation_identity,
                MAX_IDENTITY_CHARS,
                "PowerPoint presentation identity",
            )?;
            validate_bounded_text(
                &powerpoint.shape_name,
                MAX_SHAPE_NAME_CHARS,
                "PowerPoint shape name",
            )?;
            if powerpoint.slide_index == 0 || powerpoint.slide_id == 0 || powerpoint.z_order == 0 {
                return Err("PowerPoint slide and z-order references must be positive".to_string());
            }
            for (value, label) in [
                (powerpoint.left, "left"),
                (powerpoint.top, "top"),
                (powerpoint.width, "width"),
                (powerpoint.height, "height"),
                (powerpoint.rotation, "rotation"),
            ] {
                validate_finite_geometry(value, label)?;
            }
            if powerpoint.width <= 0.0 || powerpoint.height <= 0.0 {
                return Err("PowerPoint formula geometry must have positive dimensions".to_string());
            }
            for (value, label) in [
                (powerpoint.font_size_pt, "fontSizePt"),
                (powerpoint.reference_width_pt, "referenceWidthPt"),
                (powerpoint.reference_height_pt, "referenceHeightPt"),
            ] {
                if let Some(value) = value {
                    if !value.is_finite() || value <= 0.0 {
                        return Err(format!(
                            "PowerPoint formula {label} must be a positive finite number"
                        ));
                    }
                }
            }
            if let Some(font_size) = powerpoint.font_size_pt {
                if !(MIN_POWERPOINT_FONT_SIZE_PT..=MAX_POWERPOINT_FONT_SIZE_PT).contains(&font_size)
                {
                    return Err(
                        "PowerPoint formula fontSizePt is outside the supported range".to_string(),
                    );
                }
            }
        }
        _ => unreachable!(),
    }
    Ok(())
}

fn read_request(session_id: &str) -> Result<MacOfflineSessionRequest, String> {
    validate_uuid(session_id, "Session id")?;
    let mut candidates = Vec::new();
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        let path = request_path(host, session_id)?;
        match fs::symlink_metadata(&path) {
            Ok(metadata) => candidates.push((host, path, metadata)),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => {
                return Err(format!(
                    "Unable to inspect offline Office request metadata at {}: {error}",
                    path.display()
                ))
            }
        }
    }
    let (expected_host, path, metadata) = match candidates.len() {
        1 => candidates.remove(0),
        0 => {
            return Err(
                "Offline Office request was not found in either host runtime directory".to_string(),
            )
        }
        _ => {
            return Err(
                "The same Offline Office Session exists in both host runtime directories"
                    .to_string(),
            )
        }
    };
    if metadata.file_type().is_symlink()
        || !metadata.is_file()
        || metadata.len() == 0
        || metadata.len() > MAX_REQUEST_BYTES
    {
        return Err("Offline Office request has an invalid size".to_string());
    }
    let bytes = fs::read(&path)
        .map_err(|error| format!("Unable to read offline Office request: {error}"))?;
    let request: MacOfflineSessionRequest = serde_json::from_slice(&bytes)
        .map_err(|error| format!("Offline Office request contains invalid JSON: {error}"))?;
    validate_request(&request, session_id)?;
    if host_from_request_name(&request.host)? != expected_host {
        return Err(
            "Offline Office request host does not match its Application Scripts runtime directory"
                .to_string(),
        );
    }
    Ok(request)
}

fn decode_metadata(encoded: &str) -> Result<VisualTeXFormulaMetadata, String> {
    let payload = encoded
        .strip_prefix(METADATA_PREFIX)
        .ok_or_else(|| "VisualTeX formula metadata prefix is invalid".to_string())?;
    if payload.is_empty() || payload.len() > MAX_METADATA_BYTES {
        return Err("VisualTeX formula metadata payload is invalid".to_string());
    }
    let compressed = URL_SAFE_NO_PAD
        .decode(payload)
        .map_err(|error| format!("Unable to decode VisualTeX formula metadata: {error}"))?;
    let decoder = DeflateDecoder::new(compressed.as_slice());
    let mut json = Vec::new();
    decoder
        .take((MAX_METADATA_BYTES + 1) as u64)
        .read_to_end(&mut json)
        .map_err(|error| format!("Unable to inflate VisualTeX formula metadata: {error}"))?;
    if json.len() > MAX_METADATA_BYTES {
        return Err("VisualTeX formula metadata expands beyond the allowed size".to_string());
    }
    let metadata: VisualTeXFormulaMetadata = serde_json::from_slice(&json)
        .map_err(|error| format!("VisualTeX formula metadata JSON is invalid: {error}"))?;
    validate_metadata(&metadata)?;
    Ok(metadata)
}

fn validate_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<(), String> {
    if metadata.schema != "visualtex-formula" || metadata.schema_version != 1 {
        return Err("Unsupported VisualTeX formula metadata schema".to_string());
    }
    validate_uuid(&metadata.formula_id, "Metadata formulaId")?;
    if metadata.lines.is_empty() || metadata.lines.len() > 512 {
        return Err("VisualTeX formula metadata must contain 1 to 512 lines".to_string());
    }
    for line in &metadata.lines {
        validate_uuid(&line.id, "Metadata line id")?;
        if line.latex.len() > 1_000_000 {
            return Err("A VisualTeX formula line exceeds the 1 MB limit".to_string());
        }
    }
    if !matches!(metadata.display_mode.as_str(), "inline" | "block") {
        return Err("VisualTeX metadata displayMode is invalid".to_string());
    }
    for (value, label) in [
        (metadata.render_width_px, "renderWidthPx"),
        (metadata.render_height_px, "renderHeightPx"),
        (metadata.reference_width_pt, "referenceWidthPt"),
        (metadata.reference_height_pt, "referenceHeightPt"),
    ] {
        if let Some(value) = value {
            if !value.is_finite() || value <= 0.0 {
                return Err(format!(
                    "VisualTeX metadata {label} must be positive and finite"
                ));
            }
        }
    }
    if let Some(value) = metadata.font_size_pt {
        if !value.is_finite() || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&value) {
            return Err("VisualTeX metadata fontSizePt is outside the Office range".to_string());
        }
    }
    if let Some(value) = metadata.reference_baseline_pt {
        if !value.is_finite() || !(-256.0..=0.0).contains(&value) {
            return Err("VisualTeX metadata referenceBaselinePt is invalid".to_string());
        }
    }
    if let Some(value) = metadata.image_ink_center_y_ratio {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err("VisualTeX metadata imageInkCenterYRatio is invalid".to_string());
        }
    }
    Ok(())
}

fn replace_mathlive_latex_command(source: &str, command: &str, replacement: &str) -> String {
    let pattern = format!("\\{command}");
    let mut output = String::with_capacity(source.len());
    let mut cursor = 0;
    while let Some(relative) = source[cursor..].find(&pattern) {
        let start = cursor + relative;
        let end = start + pattern.len();
        output.push_str(&source[cursor..start]);
        let followed_by_command_letter = source[end..]
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_alphabetic());
        if followed_by_command_letter {
            output.push_str(&pattern);
        } else {
            output.push_str(replacement);
        }
        cursor = end;
    }
    output.push_str(&source[cursor..]);
    output
}

fn normalize_mathlive_upright_commands(source: &str) -> String {
    let mut normalized = source.to_string();
    for (command, replacement) in [
        ("capitalDifferentialD", "\\mathrm{D}"),
        ("differentialD", "\\mathrm{d}"),
        ("exponentialE", "\\mathrm{e}"),
        ("imaginaryI", "\\mathrm{i}"),
        ("imaginaryJ", "\\mathrm{j}"),
    ] {
        normalized = replace_mathlive_latex_command(&normalized, command, replacement);
    }
    for prefix in ["\\mathrm{d", "\\textrm{d"] {
        let mut output = String::with_capacity(normalized.len());
        let mut cursor = 0;
        while let Some(relative) = normalized[cursor..].find(prefix) {
            let start = cursor + relative;
            let variable_start = start + prefix.len();
            output.push_str(&normalized[cursor..start]);
            let Some(variable) = normalized[variable_start..].chars().next() else {
                output.push_str(prefix);
                cursor = variable_start;
                continue;
            };
            let variable_end = variable_start + variable.len_utf8();
            if variable.is_ascii_alphabetic() && normalized[variable_end..].starts_with('}') {
                output.push_str("\\mathrm{d}");
                output.push(variable);
                cursor = variable_end + 1;
            } else {
                output.push_str(prefix);
                cursor = variable_start;
            }
        }
        output.push_str(&normalized[cursor..]);
        normalized = output;
    }
    normalized
}

fn latex_character_is_escaped(source: &str, index: usize) -> bool {
    let bytes = source.as_bytes();
    let mut cursor = index;
    let mut slash_count = 0;
    while cursor > 0 && bytes[cursor - 1] == b'\\' {
        slash_count += 1;
        cursor -= 1;
    }
    slash_count % 2 == 1
}

fn read_latex_environment_token(source: &str, index: usize) -> Option<(bool, String, usize)> {
    let rest = source.get(index..)?;
    let (is_begin, name_start) = if rest.starts_with("\\begin{") {
        (true, index + "\\begin{".len())
    } else if rest.starts_with("\\end{") {
        (false, index + "\\end{".len())
    } else {
        return None;
    };
    let name_end = source[name_start..].find('}')? + name_start;
    let name = &source[name_start..name_end];
    if name.is_empty()
        || !name.chars().enumerate().all(|(position, character)| {
            character.is_ascii_alphabetic()
                || (character == '*' && position == name.chars().count() - 1)
        })
    {
        return None;
    }
    Some((is_begin, name.to_string(), name_end + 1))
}

fn update_latex_environment_stack(environments: &mut Vec<String>, is_begin: bool, name: String) {
    if is_begin {
        environments.push(name);
    } else if let Some(index) = environments.iter().rposition(|value| value == &name) {
        environments.remove(index);
    }
}

fn has_top_level_alignment_marker(source: &str) -> bool {
    let mut brace_depth = 0_u32;
    let mut environments = Vec::new();
    let mut index = 0;
    while index < source.len() {
        if let Some((is_begin, name, end)) = read_latex_environment_token(source, index) {
            update_latex_environment_stack(&mut environments, is_begin, name);
            index = end;
            continue;
        }
        let character = source[index..].chars().next().expect("valid UTF-8");
        if character == '{' && !latex_character_is_escaped(source, index) {
            brace_depth += 1;
        } else if character == '}' && !latex_character_is_escaped(source, index) {
            brace_depth = brace_depth.saturating_sub(1);
        } else if character == '&'
            && !latex_character_is_escaped(source, index)
            && brace_depth == 0
            && environments.is_empty()
        {
            return true;
        }
        index += character.len_utf8();
    }
    false
}

fn top_level_relation_index(source: &str) -> Option<usize> {
    const RELATION_COMMANDS: &[&str] = &[
        "\\Longleftrightarrow",
        "\\Longrightarrow",
        "\\Leftrightarrow",
        "\\Rightarrow",
        "\\leftrightarrow",
        "\\rightarrow",
        "\\leftarrow",
        "\\subseteq",
        "\\supseteq",
        "\\notin",
        "\\approx",
        "\\equiv",
        "\\simeq",
        "\\propto",
        "\\mapsto",
        "\\subset",
        "\\supset",
        "\\cong",
        "\\neq",
        "\\leq",
        "\\geq",
        "\\sim",
        "\\to",
        "\\ne",
        "\\le",
        "\\ge",
        "\\in",
    ];
    let mut brace_depth = 0_u32;
    let mut environments = Vec::new();
    let mut index = 0;
    while index < source.len() {
        if let Some((is_begin, name, end)) = read_latex_environment_token(source, index) {
            update_latex_environment_stack(&mut environments, is_begin, name);
            index = end;
            continue;
        }
        let character = source[index..].chars().next().expect("valid UTF-8");
        if character == '{' && !latex_character_is_escaped(source, index) {
            brace_depth += 1;
            index += 1;
            continue;
        }
        if character == '}' && !latex_character_is_escaped(source, index) {
            brace_depth = brace_depth.saturating_sub(1);
            index += 1;
            continue;
        }
        if brace_depth == 0 && environments.is_empty() {
            if matches!(character, '=' | '<' | '>') {
                return Some(index);
            }
            if character == '\\' {
                for command in RELATION_COMMANDS {
                    if !source[index..].starts_with(command) {
                        continue;
                    }
                    let next = source[index + command.len()..].chars().next();
                    if next.is_some_and(|value| value.is_ascii_alphabetic()) {
                        continue;
                    }
                    return Some(index);
                }
            }
        }
        index += character.len_utf8();
    }
    None
}

fn add_latex_alignment_marker(source: &str) -> String {
    if source.is_empty() || has_top_level_alignment_marker(source) {
        return source.to_string();
    }
    let Some(index) = top_level_relation_index(source) else {
        return source.to_string();
    };
    format!("{}&{}", &source[..index], &source[index..])
}

fn wrap_latex_environment(name: &str, body: &str) -> String {
    format!("\\begin{{{name}}}\n{body}\n\\end{{{name}}}")
}

fn format_document_formula_rows(lines: &[String], align_relations: bool) -> String {
    lines
        .iter()
        .enumerate()
        .map(|(index, line)| {
            let content = if align_relations {
                add_latex_alignment_marker(line)
            } else {
                line.clone()
            };
            if index + 1 < lines.len() {
                format!("{content} \\\\")
            } else {
                content
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn canonical_document_formula_latex(metadata: &VisualTeXFormulaMetadata) -> Result<String, String> {
    // `metadata.lines` stores logical editor rows. A single logical formula may
    // itself contain source-formatting newlines, especially inside equation or
    // equation* environments. Keep those internal newlines inside the same row;
    // splitting them here would rebuild one equation as several equations and
    // make the Rust validator disagree with the TypeScript serializer.
    let mut lines = metadata
        .lines
        .iter()
        .map(|line| {
            normalize_mathlive_upright_commands(
                &line.latex.replace("\r\n", "\n").replace('\r', "\n"),
            )
            .trim()
            .to_string()
        })
        .filter(|line| !line.is_empty())
        .collect::<Vec<_>>();
    if lines.is_empty() {
        lines.push(String::new());
    }
    let joined = lines.join("\n");
    let formatted = match metadata.code_format.as_str() {
        "raw" => joined,
        "inline-dollar" => lines
            .iter()
            .map(|line| format!("${line}$"))
            .collect::<Vec<_>>()
            .join("\n"),
        "inline-paren" => lines
            .iter()
            .map(|line| format!("\\({line}\\)"))
            .collect::<Vec<_>>()
            .join("\n"),
        "display-dollar" => lines
            .iter()
            .map(|line| format!("$$\n{line}\n$$"))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "display-bracket" => lines
            .iter()
            .map(|line| format!("\\[\n{line}\n\\]"))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "equation" => lines
            .iter()
            .map(|line| wrap_latex_environment("equation", line))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "equation-star" => lines
            .iter()
            .map(|line| wrap_latex_environment("equation*", line))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "align" => wrap_latex_environment("align", &format_document_formula_rows(&lines, true)),
        "align-star" => {
            wrap_latex_environment("align*", &format_document_formula_rows(&lines, true))
        }
        "aligned" => format!(
            "\\[\n{}\n\\]",
            wrap_latex_environment("aligned", &format_document_formula_rows(&lines, true),)
        ),
        "gather" => wrap_latex_environment("gather", &format_document_formula_rows(&lines, false)),
        "gather-star" => {
            wrap_latex_environment("gather*", &format_document_formula_rows(&lines, false))
        }
        "multline" => {
            wrap_latex_environment("multline", &format_document_formula_rows(&lines, false))
        }
        "multline-star" => {
            wrap_latex_environment("multline*", &format_document_formula_rows(&lines, false))
        }
        "equation-split" => wrap_latex_environment(
            "equation",
            &wrap_latex_environment("split", &format_document_formula_rows(&lines, true)),
        ),
        "equation-star-split" => wrap_latex_environment(
            "equation*",
            &wrap_latex_environment("split", &format_document_formula_rows(&lines, true)),
        ),
        _ => {
            return Err(format!(
                "Document formula metadata codeFormat is unsupported: {}",
                metadata.code_format
            ))
        }
    };
    Ok(formatted)
}

fn normalized_serialized_latex(source: &str) -> String {
    source
        .replace("\r\n", "\n")
        .replace('\r', "\n")
        .trim()
        .to_string()
}

fn validate_document_formula_metadata_match(
    metadata: &VisualTeXFormulaMetadata,
    formula_id: &str,
    latex: &str,
    display_mode: &str,
    numbered: bool,
) -> Result<String, String> {
    validate_metadata(metadata)?;
    if metadata.formula_id != formula_id
        || metadata.display_mode != display_mode
        || metadata.numbered != numbered
    {
        return Err(
            "Document formula metadata identity does not match its formula block".to_string(),
        );
    }
    // The document-import frontend builds metadata, SVG, PNG and OMML from one
    // normalized editor document. Treat that submitted serialization as the
    // source of truth instead of rebuilding it again in Rust. Re-serializing
    // here creates false mismatches for harmless differences such as CRLF/LF,
    // environment whitespace, alignat arguments and internal equation
    // newlines, even though every rendered artifact belongs to the same
    // formula. Rust still validates the full metadata schema and identity above
    // and requires the formula block and metadata to carry the same non-empty
    // normalized source.
    let formula_latex = normalized_serialized_latex(latex);
    let metadata_latex = normalized_serialized_latex(&metadata.latex);
    if formula_latex.is_empty() || metadata_latex != formula_latex {
        return Err("Document formula metadata LaTeX does not match its formula block".to_string());
    }
    Ok(formula_latex)
}

fn encode_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<String, String> {
    validate_metadata(metadata)?;
    let json = serde_json::to_vec(metadata)
        .map_err(|error| format!("Unable to encode VisualTeX formula metadata: {error}"))?;
    let mut encoder = DeflateEncoder::new(Vec::new(), Compression::best());
    encoder
        .write_all(&json)
        .map_err(|error| format!("Unable to compress VisualTeX formula metadata: {error}"))?;
    let compressed = encoder
        .finish()
        .map_err(|error| format!("Unable to finish VisualTeX formula metadata: {error}"))?;
    Ok(format!(
        "{METADATA_PREFIX}{}",
        URL_SAFE_NO_PAD.encode(compressed)
    ))
}

fn hex_encode(value: &str) -> String {
    value
        .as_bytes()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn import_request(
    state: &OfficeCompanionState,
    request: MacOfflineSessionRequest,
) -> Result<OfficeFormulaSession, String> {
    match state.session_store.get(&request.session_id) {
        Ok(existing) => return Ok(existing),
        Err(SessionError::NotFound) => {}
        Err(error) => return Err(error.to_string()),
    }

    let original_metadata = request
        .encoded_metadata
        .as_deref()
        .map(decode_metadata)
        .transpose()?;
    let metadata_formula_id = original_metadata
        .as_ref()
        .map(|value| value.formula_id.clone());
    let formula_id = match (request.formula_id.clone(), metadata_formula_id) {
        (Some(request_id), Some(metadata_id)) if request_id != metadata_id => {
            return Err("Request formulaId does not match encoded metadata".to_string())
        }
        (Some(request_id), _) => request_id,
        (None, Some(metadata_id)) => metadata_id,
        (None, None) if request.mode == "create" => Uuid::new_v4().to_string(),
        (None, None) => return Err("Edit request does not contain a formulaId".to_string()),
    };
    validate_uuid(&formula_id, "Imported formula id")?;

    let host = match request.host.as_str() {
        "word" => OfficeHost::Word,
        "powerpoint" => OfficeHost::Powerpoint,
        _ => unreachable!(),
    };
    let mode = match request.mode.as_str() {
        "create" => OfficeSessionMode::Create,
        "edit" => OfficeSessionMode::Edit,
        _ => unreachable!(),
    };
    let lines = original_metadata
        .as_ref()
        .map(|metadata| {
            metadata
                .lines
                .iter()
                .map(|line| FormulaLine {
                    id: line.id.clone(),
                    latex: line.latex.clone(),
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_else(|| {
            vec![FormulaLine {
                id: Uuid::new_v4().to_string(),
                latex: String::new(),
            }]
        });
    let source_document_id = match host {
        OfficeHost::Word => request.source_document_id.clone(),
        OfficeHost::Powerpoint => request.power_point.as_ref().map(|powerpoint| {
            format!(
                "visualtex-ppt-native-presentation:{}",
                powerpoint.presentation_identity
            )
        }),
    };
    let source_object_id = match host {
        OfficeHost::Word => match mode {
            OfficeSessionMode::Create => request.pending_marker.clone(),
            OfficeSessionMode::Edit => request
                .source_object_id
                .clone()
                .or_else(|| request.encoded_metadata.clone()),
        },
        OfficeHost::Powerpoint => request.power_point.as_ref().map(|powerpoint| {
            format!(
                "visualtex-ppt-native-edit:{}:{}",
                powerpoint.slide_index,
                hex_encode(&powerpoint.shape_name)
            )
        }),
    };
    let title = original_metadata
        .as_ref()
        .map(|metadata| metadata.title.clone())
        .unwrap_or_else(|| match host {
            OfficeHost::Word => "Word Formula".to_string(),
            OfficeHost::Powerpoint => "PowerPoint Formula".to_string(),
        });
    let code_format = original_metadata
        .as_ref()
        .map(|metadata| metadata.code_format.clone())
        .unwrap_or_else(|| "latex".to_string());
    let font_size_pt = request
        .font_size_pt
        .or_else(|| {
            request
                .power_point
                .as_ref()
                .and_then(|powerpoint| powerpoint.font_size_pt)
        })
        .or_else(|| {
            original_metadata
                .as_ref()
                .and_then(|metadata| metadata.font_size_pt)
        });

    let session_id = request.session_id.clone();
    let session_operation = request
        .operation
        .as_deref()
        .filter(|operation| *operation != "formula")
        .map(str::to_string);
    match state.session_store.create_external(
        session_id.clone(),
        CreateOfficeSessionInput {
            mode,
            host,
            operation: session_operation,
            formula_id: Some(formula_id),
            source_document_id,
            source_object_id,
            title: Some(title),
            lines: Some(lines),
            active_line_id: None,
            code_format: Some(code_format),
            display_mode: Some(request.display_mode),
            numbered: Some(request.numbered),
            font_size_pt,
            formula_letter_font: original_metadata
                .as_ref()
                .and_then(|metadata| metadata.formula_letter_font.clone()),
            formula_chinese_font: original_metadata
                .as_ref()
                .and_then(|metadata| metadata.formula_chinese_font.clone()),
            export_width: None,
            export_height: None,
            original_metadata,
            auto_commit_on_close: Some(true),
        },
    ) {
        Ok(session) => Ok(session),
        Err(SessionError::Conflict(_)) => state
            .session_store
            .get(&session_id)
            .map_err(|error| error.to_string()),
        Err(error) => Err(error.to_string()),
    }
}

pub(crate) fn parse_office_url(value: &str) -> Result<String, String> {
    const PREFIX: &str = "visualtex://office/open?session=";
    let session_id = value
        .strip_prefix(PREFIX)
        .ok_or_else(|| "VisualTeX URL must use visualtex://office/open".to_string())?;
    if session_id.contains(['&', '#', '?', '/', '%']) {
        return Err("VisualTeX URL contains unsupported query data".to_string());
    }
    validate_uuid(session_id, "VisualTeX URL Session id")?;
    Ok(session_id.to_string())
}

fn office_host_name(host: OfficeHost) -> &'static str {
    match host {
        OfficeHost::Word => "word",
        OfficeHost::Powerpoint => "powerpoint",
    }
}

#[cfg(target_os = "macos")]
fn restore_office_host_focus(host: OfficeHost) {
    let bundle_identifier = match host {
        OfficeHost::Word => "com.microsoft.Word",
        OfficeHost::Powerpoint => "com.microsoft.Powerpoint",
    };
    if !crate::office::background::activate_application_by_bundle_identifier(bundle_identifier) {
        eprintln!(
            "Unable to return focus to {} after closing the VisualTeX formula editor",
            office_host_name(host)
        );
    }
}

#[cfg(not(target_os = "macos"))]
fn restore_office_host_focus(_host: OfficeHost) {}

fn editor_window_label(host: OfficeHost) -> &'static str {
    match host {
        OfficeHost::Word => "office-native-word-editor",
        OfficeHost::Powerpoint => "office-native-powerpoint-editor",
    }
}

fn editor_prewarmed_marker_path(app: &AppHandle, host: OfficeHost) -> Result<PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data
        .join("office")
        .join(format!("{}-editor-prewarmed.json", office_host_name(host))))
}

fn editor_prewarm_diagnostic_path(app: &AppHandle, host: OfficeHost) -> Result<PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data
        .join("office")
        .join(format!("{}-editor-prewarm-diagnostic.json", office_host_name(host))))
}

fn editor_window_host(label: &str) -> Option<OfficeHost> {
    match label {
        "office-native-word-editor" => Some(OfficeHost::Word),
        "office-native-powerpoint-editor" => Some(OfficeHost::Powerpoint),
        _ => None,
    }
}

fn office_editor_monitor_geometry(monitor: &Monitor) -> Option<OfficeEditorMonitorGeometry> {
    let scale_factor = monitor.scale_factor();
    if !scale_factor.is_finite() || scale_factor <= 0.0 {
        return None;
    }
    let screen_width = monitor.size().width as f64 / scale_factor;
    let screen_height = monitor.size().height as f64 / scale_factor;
    let work_width = monitor.work_area().size.width as f64 / scale_factor;
    let work_height = monitor.work_area().size.height as f64 / scale_factor;
    if screen_width <= 0.0 || screen_height <= 0.0 {
        return None;
    }
    Some(OfficeEditorMonitorGeometry {
        screen_width,
        screen_height,
        maximum_inner_width: work_width.max(MIN_OFFICE_EDITOR_WIDTH),
        maximum_inner_height: (work_height - OFFICE_EDITOR_TITLE_BAR_ALLOWANCE)
            .max(MIN_OFFICE_EDITOR_HEIGHT),
    })
}

fn primary_office_editor_monitor_geometry(
    app: &AppHandle,
) -> Option<OfficeEditorMonitorGeometry> {
    app.primary_monitor()
        .ok()
        .flatten()
        .as_ref()
        .and_then(office_editor_monitor_geometry)
}

fn current_office_editor_monitor_geometry(
    app: &AppHandle,
    window: &WebviewWindow,
) -> Option<OfficeEditorMonitorGeometry> {
    window
        .current_monitor()
        .ok()
        .flatten()
        .as_ref()
        .and_then(office_editor_monitor_geometry)
        .or_else(|| primary_office_editor_monitor_geometry(app))
}

fn normalize_office_editor_ratio(value: Option<f64>, fallback: f64) -> f64 {
    value
        .filter(|value| value.is_finite() && *value >= 0.1 && *value <= 1.0)
        .unwrap_or(fallback)
}

fn normalize_office_editor_window_size(
    width: f64,
    height: f64,
    monitor: Option<OfficeEditorMonitorGeometry>,
) -> OfficeEditorWindowSize {
    let width = if width.is_finite() {
        width
    } else {
        DEFAULT_OFFICE_EDITOR_FALLBACK_WIDTH
    };
    let height = if height.is_finite() {
        height
    } else {
        DEFAULT_OFFICE_EDITOR_FALLBACK_HEIGHT
    };
    let maximum_width = monitor
        .map(|geometry| geometry.maximum_inner_width)
        .unwrap_or(MAX_OFFICE_EDITOR_WIDTH)
        .min(MAX_OFFICE_EDITOR_WIDTH)
        .max(MIN_OFFICE_EDITOR_WIDTH);
    let maximum_height = monitor
        .map(|geometry| geometry.maximum_inner_height)
        .unwrap_or(MAX_OFFICE_EDITOR_HEIGHT)
        .min(MAX_OFFICE_EDITOR_HEIGHT)
        .max(MIN_OFFICE_EDITOR_HEIGHT);
    OfficeEditorWindowSize {
        width: width.clamp(MIN_OFFICE_EDITOR_WIDTH, maximum_width),
        height: height.clamp(MIN_OFFICE_EDITOR_HEIGHT, maximum_height),
    }
}

fn resolve_office_editor_window_size(
    preference: Option<OfficeEditorWindowSizePreference>,
    monitor: Option<OfficeEditorMonitorGeometry>,
) -> OfficeEditorWindowSize {
    let preference = preference.unwrap_or_default();
    let fallback_width = preference
        .fallback_width
        .or(preference.width)
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(DEFAULT_OFFICE_EDITOR_FALLBACK_WIDTH);
    let fallback_height = preference
        .fallback_height
        .or(preference.height)
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(DEFAULT_OFFICE_EDITOR_FALLBACK_HEIGHT);

    let (width, height) = if let Some(geometry) = monitor {
        let migrated_width_ratio = preference
            .width
            .filter(|value| value.is_finite() && *value > 0.0)
            .map(|value| value / geometry.screen_width);
        let migrated_height_ratio = preference
            .height
            .filter(|value| value.is_finite() && *value > 0.0)
            .map(|value| value / geometry.screen_height);
        let width_ratio = normalize_office_editor_ratio(
            preference.width_ratio.or(migrated_width_ratio),
            DEFAULT_OFFICE_EDITOR_WIDTH_RATIO,
        );
        let height_ratio = normalize_office_editor_ratio(
            preference.height_ratio.or(migrated_height_ratio),
            DEFAULT_OFFICE_EDITOR_HEIGHT_RATIO,
        );
        (
            geometry.screen_width * width_ratio,
            geometry.screen_height * height_ratio,
        )
    } else {
        (fallback_width, fallback_height)
    };
    normalize_office_editor_window_size(width, height, monitor)
}

fn office_editor_window_size_path(app: &AppHandle) -> Result<PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data.join("office").join(OFFICE_EDITOR_WINDOW_SIZE_FILE))
}

fn read_office_editor_window_size_preference(
    app: &AppHandle,
) -> Option<OfficeEditorWindowSizePreference> {
    let path = office_editor_window_size_path(app).ok()?;
    let bytes = fs::read(path).ok()?;
    serde_json::from_slice(&bytes).ok()
}

fn load_office_editor_window_size(app: &AppHandle) -> OfficeEditorWindowSize {
    resolve_office_editor_window_size(
        read_office_editor_window_size_preference(app),
        primary_office_editor_monitor_geometry(app),
    )
}

pub(crate) fn configuration_office_editor_window_size(
    app: &AppHandle,
) -> Option<(f64, f64)> {
    let size = load_office_editor_window_size(app);
    Some((size.width, size.height))
}

pub(crate) fn apply_configuration_office_editor_window_size(
    app: &AppHandle,
    width: f64,
    height: f64,
) -> Result<(), String> {
    let monitor = primary_office_editor_monitor_geometry(app);
    let size = normalize_office_editor_window_size(width, height, monitor);
    let preference = OfficeEditorWindowSizePreference {
        width_ratio: monitor.map(|geometry| {
            (size.width / geometry.screen_width).clamp(0.1, 1.0)
        }),
        height_ratio: monitor.map(|geometry| {
            (size.height / geometry.screen_height).clamp(0.1, 1.0)
        }),
        fallback_width: Some(size.width),
        fallback_height: Some(size.height),
        width: None,
        height: None,
    };
    let path = office_editor_window_size_path(app)?;
    let bytes = serde_json::to_vec_pretty(&preference).map_err(|error| error.to_string())?;
    atomic_write(&path, &bytes, 0o600)?;

    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        if let Some(window) = app.get_webview_window(editor_window_label(host)) {
            window
                .set_size(tauri::LogicalSize::new(size.width, size.height))
                .map_err(|error| error.to_string())?;
        }
    }
    Ok(())
}

pub(crate) fn schedule_persist_office_editor_window_size(
    app: &AppHandle,
    label: &str,
    physical_width: u32,
    physical_height: u32,
) {
    if editor_window_host(label).is_none() || physical_width == 0 || physical_height == 0 {
        return;
    }
    let Some(window) = app.get_webview_window(label) else {
        return;
    };
    let scale_factor = window.scale_factor().unwrap_or(1.0).max(0.1);
    let fallback_width = physical_width as f64 / scale_factor;
    let fallback_height = physical_height as f64 / scale_factor;
    let monitor = current_office_editor_monitor_geometry(app, &window);
    let preference = OfficeEditorWindowSizePreference {
        width_ratio: monitor.map(|geometry| {
            (fallback_width / geometry.screen_width).clamp(0.1, 1.0)
        }),
        height_ratio: monitor.map(|geometry| {
            (fallback_height / geometry.screen_height).clamp(0.1, 1.0)
        }),
        fallback_width: Some(fallback_width),
        fallback_height: Some(fallback_height),
        width: None,
        height: None,
    };
    let generation = OFFICE_EDITOR_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst) + 1;
    let app = app.clone();
    std::thread::spawn(move || {
        std::thread::sleep(Duration::from_millis(250));
        if OFFICE_EDITOR_SIZE_WRITE_GENERATION.load(Ordering::SeqCst) != generation {
            return;
        }
        let Ok(path) = office_editor_window_size_path(&app) else {
            return;
        };
        let Ok(bytes) = serde_json::to_vec_pretty(&preference) else {
            return;
        };
        if let Err(error) = atomic_write(&path, &bytes, 0o600) {
            eprintln!("Unable to persist the Office editor window size: {error}");
        }
    });
}

fn document_import_window_label(session_id: &str) -> String {
    format!("office-native-document-{}", session_id.replace('-', ""))
}

fn active_editor_session(host: OfficeHost) -> Option<ActiveOfficeEditorSession> {
    office_editor_runtime()
        .lock()
        .ok()
        .and_then(|runtime| runtime.active(host).cloned())
}

fn clear_editor_session(host: OfficeHost, session_id: &str, generation: u64) -> Result<(), String> {
    let mut runtime = office_editor_runtime()
        .lock()
        .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
    let matches = runtime.active(host).is_some_and(|active| {
        active.activation.session_id == session_id && active.activation.generation == generation
    });
    if !matches {
        return Err("The Office editor Session is no longer active".to_string());
    }
    *runtime.active_mut(host) = None;
    Ok(())
}

fn clear_any_editor_session(host: OfficeHost) -> Option<MacOfflineOfficeEditorActivation> {
    let mut runtime = office_editor_runtime().lock().ok()?;
    runtime
        .active_mut(host)
        .take()
        .map(|active| active.activation)
}

#[cfg(target_os = "macos")]
fn set_resident_editor_parked(window: &WebviewWindow, parked: bool) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            // Keep the resident WKWebView continuously alive, exactly as in
            // eb2fcf2a. orderOut()/hide() suspends WebKit and causes the observed
            // multi-second transparent wake-up before the editor can paint.
            // 1% native opacity is still visible on bright/high-contrast desktop
            // backgrounds, so park at a much smaller non-zero alpha instead.
            native_window.setAlphaValue(if parked { 0.001 } else { 1.0 });
            native_window.setIgnoresMouseEvents(parked);
            // A ready Office editor must remain visually above Word/PowerPoint
            // even when macOS refuses a cross-application activation request.
            // Only the dedicated editor is promoted; the desktop main window
            // stays at/below the normal level and is never revealed here.
            native_window.setLevel(if parked {
                objc2_app_kit::NSNormalWindowLevel
            } else {
                objc2_app_kit::NSFloatingWindowLevel
            });
        })
        .map_err(|error| format!("Unable to update the resident Office editor window: {error}"))
}

#[cfg(not(target_os = "macos"))]
fn set_resident_editor_parked(window: &WebviewWindow, parked: bool) -> Result<(), String> {
    if parked {
        window.hide()
    } else {
        window.show()
    }
    .map_err(|error| error.to_string())
}

#[cfg(target_os = "macos")]
fn wake_resident_editor_for_hydration(window: &WebviewWindow) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            native_window.setAlphaValue(1.0);
            native_window.setIgnoresMouseEvents(true);
            native_window.setLevel(objc2_app_kit::NSNormalWindowLevel);
        })
        .map_err(|error| format!("Unable to wake the resident Office editor window: {error}"))
}

#[cfg(not(target_os = "macos"))]
fn wake_resident_editor_for_hydration(window: &WebviewWindow) -> Result<(), String> {
    window.show().map_err(|error| error.to_string())
}

#[cfg(target_os = "macos")]
fn order_main_window_behind_office_editor(app: &AppHandle) -> Result<(), String> {
    let Some(main_window) = app.get_webview_window("main") else {
        return Ok(());
    };
    if !main_window.is_visible().unwrap_or(false) {
        return Ok(());
    }
    main_window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            // ActivateAllWindows raises every normal-level VisualTeX window as
            // one application group. Put only the desktop workspace one level
            // below normal before activation so it stays behind Word/PowerPoint
            // without being hidden, moved, minimized, or resized.
            native_window.setLevel(objc2_app_kit::NSNormalWindowLevel - 1);
            native_window.orderBack(None);
        })
        .map_err(|error| format!("Unable to keep the VisualTeX main window behind Office: {error}"))
}

#[cfg(not(target_os = "macos"))]
fn order_main_window_behind_office_editor(_app: &AppHandle) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
fn restore_main_window_level_after_office_editor(app: &AppHandle) -> Result<(), String> {
    let Some(main_window) = app.get_webview_window("main") else {
        return Ok(());
    };
    main_window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            native_window.setLevel(objc2_app_kit::NSNormalWindowLevel);
            // The Office host is active again when this runs. Keep the restored
            // desktop workspace behind it until the user explicitly clicks it.
            native_window.orderBack(None);
        })
        .map_err(|error| format!("Unable to restore the VisualTeX main window level: {error}"))
}

#[cfg(not(target_os = "macos"))]
fn restore_main_window_level_after_office_editor(_app: &AppHandle) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
fn inspect_resident_editor_focus(
    window: &WebviewWindow,
) -> Result<ResidentEditorFocusState, String> {
    let (sender, receiver) = mpsc::sync_channel(1);
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            let state = MainThreadMarker::new()
                .map(|main_thread| {
                    let application = NSApplication::sharedApplication(main_thread);
                    ResidentEditorFocusState {
                        app_active: application.isActive(),
                        window_can_become_key: native_window.canBecomeKeyWindow(),
                        window_is_key: native_window.isKeyWindow(),
                        window_is_main: native_window.isMainWindow(),
                    }
                })
                .unwrap_or_default();
            let _ = sender.send(state);
        })
        .map_err(|error| format!("Unable to inspect the Office editor focus state: {error}"))?;
    receiver
        .recv_timeout(Duration::from_millis(250))
        .map_err(|error| format!("Timed out inspecting the Office editor focus state: {error}"))
}

#[cfg(target_os = "macos")]
fn make_resident_editor_key(window: &WebviewWindow) -> Result<bool, String> {
    let (sender, receiver) = mpsc::sync_channel(1);
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            native_window.orderFrontRegardless();
            native_window.makeKeyAndOrderFront(None);
            native_window.makeKeyWindow();
            let _ = sender.send(native_window.isKeyWindow());
        })
        .map_err(|error| format!("Unable to make the Office editor key: {error}"))?;

    let initially_key = receiver
        .recv_timeout(Duration::from_millis(250))
        .map_err(|error| format!("Timed out making the Office editor key: {error}"))?;
    window.set_focus().map_err(|error| error.to_string())?;
    if initially_key {
        return Ok(true);
    }

    // Both AppKit activation and Tauri's focus message can complete one run-loop
    // turn after their API call returns. Poll the actual NSWindow on the main
    // thread for a short bounded interval instead of reading either state early.
    for attempt in 0..12 {
        let focus = inspect_resident_editor_focus(window)?;
        if focus.window_is_key || window.is_focused().unwrap_or(false) {
            return Ok(true);
        }
        if attempt < 11 {
            std::thread::sleep(Duration::from_millis(5));
        }
    }
    Ok(false)
}

#[cfg(not(target_os = "macos"))]
fn inspect_resident_editor_focus(
    window: &WebviewWindow,
) -> Result<ResidentEditorFocusState, String> {
    Ok(ResidentEditorFocusState {
        app_active: window.is_focused().unwrap_or(false),
        window_can_become_key: true,
        window_is_key: window.is_focused().unwrap_or(false),
        window_is_main: window.is_focused().unwrap_or(false),
    })
}

#[cfg(not(target_os = "macos"))]
fn make_resident_editor_key(window: &WebviewWindow) -> Result<bool, String> {
    window.set_focus().map_err(|error| error.to_string())?;
    Ok(window.is_focused().unwrap_or(false))
}

#[cfg(target_os = "macos")]
fn present_resident_editor_window(app: &AppHandle, window: &WebviewWindow) -> Result<(), String> {
    // Keep the already-validated foreground sequence unchanged. Window-size
    // restoration happens earlier while the resident editor is still parked.
    set_resident_editor_parked(window, false)?;
    window.show().map_err(|error| error.to_string())?;
    window.unminimize().map_err(|error| error.to_string())?;
    order_main_window_behind_office_editor(app)?;
    crate::office::background::activate_foreground_app(app)?;
    let _ = make_resident_editor_key(window)?;
    window.set_focus().map_err(|error| error.to_string())
}

#[cfg(not(target_os = "macos"))]
fn present_resident_editor_window(_app: &AppHandle, window: &WebviewWindow) -> Result<(), String> {
    window.show().map_err(|error| error.to_string())?;
    window.set_focus().map_err(|error| error.to_string())
}

fn set_resident_editor_content_visible(
    window: &WebviewWindow,
    visible: bool,
) -> Result<(), String> {
    let opacity = if visible { "1" } else { "0" };
    window
        .eval(format!(
            "if (document.body) {{ document.body.style.opacity = '{opacity}'; }}"
        ))
        .map_err(|error| format!("Unable to update resident Office editor content: {error}"))
}

fn create_editor_window(app: &AppHandle, host: OfficeHost) -> Result<WebviewWindow, String> {
    let label = editor_window_label(host);
    if let Some(window) = app.get_webview_window(label) {
        return Ok(window);
    }

    let theme = crate::persisted_app_theme(app);
    let saved_size = load_office_editor_window_size(app);
    let path = format!(
        "office-native-dialog.html?transport=tauri&officeHost={}&theme={theme}",
        office_host_name(host),
    );
    let window = WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title("VisualTeX Office Formula")
        .inner_size(saved_size.width, saved_size.height)
        .min_inner_size(MIN_OFFICE_EDITOR_WIDTH, MIN_OFFICE_EDITOR_HEIGHT)
        .focused(false)
        .skip_taskbar(true)
        .visible(false)
        .background_throttling(BackgroundThrottlingPolicy::Disabled)
        .build()
        .map_err(|error| format!("Unable to initialize the VisualTeX Office editor: {error}"))?;
    // Keep the resident editor continuously ordered at an imperceptible alpha,
    // matching eb2fcf2a. A truly hidden WKWebView is suspended before React can
    // mount and later needs seconds to wake and repaint.
    set_resident_editor_parked(&window, true)?;
    window
        .show()
        .map_err(|error| format!("Unable to prewarm the VisualTeX Office editor: {error}"))?;
    Ok(window)
}

pub(crate) fn prewarm_office_editor_windows(app: &AppHandle) -> Result<(), String> {
    // Both WebKit processes start once with the resident app. Formula
    // interactions only switch a Session in an already initialized WebView.
    crate::office::background::install_application_icon(app)?;
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        create_editor_window(app, host)?;
    }
    Ok(())
}

fn open_editor_window(
    app: &AppHandle,
    host: OfficeHost,
    session_id: &str,
    received_epoch_ms: u64,
    received_at: Instant,
    silent: bool,
) -> Result<(), String> {
    let label = editor_window_label(host);
    let reused = app.get_webview_window(label).is_some();
    let window = create_editor_window(app, host)?;
    set_resident_editor_content_visible(&window, false)?;
    let activation = {
        let mut runtime = office_editor_runtime()
            .lock()
            .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
        // Match eb2fcf2a: park only by alpha, never by orderOut/hide, so the
        // resident WebView remains mounted and can hydrate immediately.
        set_resident_editor_parked(&window, true)?;
        runtime.next_generation = runtime.next_generation.saturating_add(1).max(1);
        let activation = MacOfflineOfficeEditorActivation {
            session_id: session_id.to_string(),
            host,
            generation: runtime.next_generation,
            received_epoch_ms,
        };
        *runtime.active_mut(host) = Some(ActiveOfficeEditorSession {
            activation: activation.clone(),
            received_at,
            ready: false,
        });
        activation
    };
    let elapsed_ms = received_at.elapsed().as_secs_f64() * 1000.0;
    queue_editor_performance(
        host,
        session_id,
        if reused {
            "window-reused"
        } else {
            "window-created"
        },
        elapsed_ms,
        Some(activation.generation),
        json!({ "windowLabel": label }),
    );
    if let Err(error) = window.emit(OFFICE_EDITOR_ACTIVATE_EVENT, activation.clone()) {
        let _ = clear_editor_session(host, session_id, activation.generation);
        return Err(format!("Unable to activate the VisualTeX Office editor: {error}"));
    }
    queue_editor_performance(
        host,
        session_id,
        "activation-event-sent",
        received_at.elapsed().as_secs_f64() * 1000.0,
        Some(activation.generation),
        json!({}),
    );
    // Ordinary editing hydrates at full native alpha before foreground
    // promotion. A direct native-to-image conversion uses the same resident
    // renderer, but remains parked and mouse-inert for the whole automatic
    // commit so Word never flashes the formula editor.
    if !silent {
        wake_resident_editor_for_hydration(&window)?;
        window.center().map_err(|error| error.to_string())?;
        window.show().map_err(|error| error.to_string())?;
        window.unminimize().map_err(|error| error.to_string())?;
    }
    Ok(())
}

fn set_word_document_operation_preparing_status(operation: &str) -> Result<(), String> {
    let status = if operation == "latexRedraw" {
        "VisualTeX is rendering LaTeX formulas..."
    } else if operation == "formulaRestore" {
        "VisualTeX is restoring Word formulas..."
    } else {
        "VisualTeX is preparing the Word document import..."
    };
    let script = format!(
        "tell application \"Microsoft Word\" to set status bar to {:?}",
        status
    );
    let output = Command::new("/usr/bin/osascript")
        .args(["-e", "tell application \"Microsoft Word\" to activate", "-e", &script])
        .output()
        .map_err(|error| format!("Unable to activate Microsoft Word: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        Err(if detail.is_empty() {
            "Unable to activate Microsoft Word".to_string()
        } else {
            format!("Unable to activate Microsoft Word: {detail}")
        })
    }
}

fn clear_word_document_import_status() {
    let _ = Command::new("/usr/bin/osascript")
        .args([
            "-e",
            "tell application \"Microsoft Word\" to set status bar to \"\"",
        ])
        .output();
}

fn open_document_import_window(app: &AppHandle, session_id: &str) -> Result<(), String> {
    crate::office::background::activate_foreground_app(app)?;
    crate::office::background::install_application_icon(app)?;

    let label = document_import_window_label(session_id);
    if let Some(window) = app.get_webview_window(&label) {
        crate::office::background::activate_foreground_app(app)?;
        window.show().map_err(|error| error.to_string())?;
        window.unminimize().map_err(|error| error.to_string())?;
        crate::office::background::activate_foreground_app(app)?;
        window.set_focus().map_err(|error| error.to_string())?;
        order_main_window_behind_office_editor(app)?;
        return Ok(());
    }
    let theme = crate::persisted_app_theme(app);
    let path = format!(
        "index.html?view=office-document-import&sessionId={session_id}&transport=tauri&theme={theme}"
    );
    let window = WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title("VisualTeX Word 文档批量导入")
        .inner_size(1260.0, 840.0)
        .min_inner_size(860.0, 620.0)
        .center()
        .build()
        .map_err(|error| format!("Unable to open the VisualTeX document importer: {error}"))?;
    window.show().map_err(|error| error.to_string())?;
    window.unminimize().map_err(|error| error.to_string())?;
    crate::office::background::activate_foreground_app(app)?;
    window.set_focus().map_err(|error| error.to_string())?;
    order_main_window_behind_office_editor(app)?;
    Ok(())
}

fn open_word_latex_redraw_window(app: &AppHandle, session_id: &str) -> Result<(), String> {
    crate::office::background::install_application_icon(app)?;
    let label = document_import_window_label(session_id);
    if let Some(window) = app.get_webview_window(&label) {
        window.hide().map_err(|error| error.to_string())?;
        return Ok(());
    }
    let theme = crate::persisted_app_theme(app);
    let path = format!(
        "index.html?view=office-word-latex-redraw&sessionId={session_id}&transport=tauri&theme={theme}"
    );
    WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title("VisualTeX Word LaTeX Redraw")
        .inner_size(560.0, 300.0)
        .min_inner_size(420.0, 220.0)
        .center()
        .focused(false)
        .skip_taskbar(true)
        .visible(false)
        .background_throttling(BackgroundThrottlingPolicy::Disabled)
        .build()
        .map_err(|error| format!("Unable to start the VisualTeX Word redraw renderer: {error}"))?;
    Ok(())
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
    app: AppHandle,
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
    let silent = matches!(
        read_request(&input.session_id)?.operation.as_deref(),
        Some("nativeToImage" | "imageToNative")
    );
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
        set_resident_editor_parked(&window, false)?;
        window.center().map_err(|error| error.to_string())?;
        window.show().map_err(|error| error.to_string())?;
        window.unminimize().map_err(|error| error.to_string())?;
        order_main_window_behind_office_editor(&app)?;
        crate::office::background::activate_foreground_app(&app)?;
        if !inspect_resident_editor_focus(&window)?.app_active {
            crate::office::background::activate_foreground_app_via_launch_services(&app)?;
        }
        let native_window_key = make_resident_editor_key(&window)?;
        let focus = inspect_resident_editor_focus(&window)?;
        let window_focused =
            native_window_key || focus.window_is_key || window.is_focused().unwrap_or(false);
        let window_visible = window.is_visible().unwrap_or(false);
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
        json!({ "silent": silent }),
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
    std::thread::spawn(move || {
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
        *runtime.active_mut(host) = None;
    }
    let _ = window.emit(
        OFFICE_EDITOR_CLEAR_EVENT,
        json!({ "sessionId": session_id, "generation": generation }),
    );

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
    // Applying or cancelling a formula ends the temporary VisualTeX editing
    // interaction. Explicitly return the foreground application to the Office
    // host without changing the visibility of the user's main workspace.
    restore_office_host_focus(host);
    queue_editor_performance(
        host,
        &session_id,
        "editor-visible-complete",
        0.0,
        Some(generation),
        json!({ "parked": true, "officeFocusRequested": true }),
    );
    if !has_open_office_editor(&app) {
        restore_main_window_level_after_office_editor(&app)?;
    }
    Ok(())
}

pub(crate) fn has_open_office_editor(app: &AppHandle) -> bool {
    [OfficeHost::Word, OfficeHost::Powerpoint]
        .into_iter()
        .any(|host| {
            active_editor_session(host).is_some()
                && app.get_webview_window(editor_window_label(host)).is_some()
        })
}

fn has_recent_office_editor_request_in_roots(
    now: SystemTime,
    max_age: Duration,
    roots: impl IntoIterator<Item = PathBuf>,
) -> bool {
    roots.into_iter().any(|root| {
        let Ok(entries) = fs::read_dir(&root) else {
            return false;
        };
        entries.flatten().any(|entry| {
            let request = entry.path().join(REQUEST_FILE);
            let Ok(modified) = fs::metadata(request).and_then(|metadata| metadata.modified())
            else {
                return false;
            };
            now.duration_since(modified)
                .map(|age| age <= max_age)
                .unwrap_or(true)
        })
    })
}

pub(crate) fn has_recent_office_editor_request(max_age: Duration) -> bool {
    has_recent_office_editor_request_in_roots(
        SystemTime::now(),
        max_age,
        [OfficeHost::Word, OfficeHost::Powerpoint]
            .into_iter()
            .filter_map(|host| sessions_root(host).ok()),
    )
}

pub(crate) fn focus_open_office_editor(app: &AppHandle) -> bool {
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        if let Some(active) = active_editor_session(host) {
            let Some(window) = app.get_webview_window(editor_window_label(host)) else {
                continue;
            };
            // A parked active window is still hydrating. Treat it as owned so
            // a second native double-click route cannot launch a duplicate.
            // LaunchServices Reopen may still raise the desktop main window;
            // demote only that window without focusing transparent editor content.
            if !active.ready {
                let _ = order_main_window_behind_office_editor(app);
                return true;
            }
            let _ = window.show();
            let _ = present_resident_editor_window(app, &window);
            return true;
        }
    }
    false
}

pub(crate) fn handle_open_url(app: &AppHandle, value: &str) -> Result<(), String> {
    let received_at = Instant::now();
    let received_epoch_ms = epoch_ms();
    let session_id = parse_office_url(value)?;
    let state = app
        .try_state::<OfficeCompanionState>()
        .ok_or_else(|| "VisualTeX Office state is not initialized".to_string())?;
    // The host is not trusted until request validation completes, so queue the
    // first stage immediately after read_request resolves it below.
    let request = read_request(&session_id)?;
    let host = host_from_request_name(&request.host)?;
    queue_editor_performance(
        host,
        &session_id,
        "url-received",
        0.0,
        None,
        json!({ "receivedEpochMs": received_epoch_ms }),
    );
    queue_editor_performance(
        host,
        &session_id,
        "request-read",
        received_at.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    ensure_runtime_root(host)?;

    if matches!(
        request.operation.as_deref(),
        Some("documentImport" | "latexRedraw" | "formulaRestore")
    ) {
        for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
            if let Some(window) = app.get_webview_window(editor_window_label(host)) {
                let _ = set_resident_editor_content_visible(&window, false);
                if let Some(active) = clear_any_editor_session(host) {
                    let _ = window.emit(
                        OFFICE_EDITOR_CLEAR_EVENT,
                        json!({
                            "sessionId": active.session_id,
                            "generation": active.generation,
                        }),
                    );
                }
                let _ = set_resident_editor_parked(&window, true);
            }
        }
        for (label, window) in app.webview_windows() {
            if label.starts_with("office-native-document-")
                && label != document_import_window_label(&session_id)
            {
                let _ = window.destroy();
            }
        }
        return if matches!(
            request.operation.as_deref(),
            Some("latexRedraw" | "formulaRestore")
        ) {
            open_word_latex_redraw_window(app, &session_id)
        } else {
            open_document_import_window(app, &session_id)
        };
    }

    let silent = matches!(
        request.operation.as_deref(),
        Some("nativeToImage" | "imageToNative")
    );
    import_request(state.inner(), request)?;
    queue_editor_performance(
        host,
        &session_id,
        "request-imported",
        received_at.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    open_editor_window(
        app,
        host,
        &session_id,
        received_epoch_ms,
        received_at,
        silent,
    )
}

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

fn atomic_write_runtime(path: &Path, bytes: &[u8], mode: u32) -> Result<(), String> {
    // Session artifacts are consumed only after the atomic rename completes.
    // They do not need crash-durable fsync before a synchronous Office callback,
    // and skipping it removes several full storage barriers from every Apply.
    atomic_write_with_durability(path, bytes, mode, false)
}

fn sanitize_dispatch_value(value: &str, label: &str) -> Result<String, String> {
    if value.contains(['\r', '\n', '\0']) {
        return Err(format!("{label} contains unsupported control characters"));
    }
    Ok(value.to_string())
}

fn dispatch_text(entries: &[(&str, String)]) -> Result<String, String> {
    let dynamic = entries
        .iter()
        .map(|(key, value)| ((*key).to_string(), value.clone()))
        .collect::<Vec<_>>();
    dynamic_dispatch_text(&dynamic)
}

fn dynamic_dispatch_text(entries: &[(String, String)]) -> Result<String, String> {
    let mut seen = std::collections::HashSet::new();
    let mut output = String::new();
    for (key, value) in entries {
        if !seen.insert(key.as_str())
            || key.is_empty()
            || !key.bytes().all(|byte| byte.is_ascii_alphanumeric())
        {
            return Err("VisualTeX dispatch contains an invalid key".to_string());
        }
        output.push_str(key);
        output.push('=');
        output.push_str(&sanitize_dispatch_value(value, key)?);
        output.push('\n');
    }
    Ok(output)
}

fn vba_callback_script(host: OfficeHost) -> &'static str {
    match host {
        OfficeHost::Word => {
            r#"with timeout of 1800 seconds
tell application "Microsoft Word"
run VB macro macro name "VisualTeX_ApplyPendingResult"
end tell
end timeout"#
        }
        OfficeHost::Powerpoint => {
            r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "Microsoft PowerPoint has no active presentation"
run VB macro macro name "VisualTeX_ApplyPendingResult" list of parameters {}
end tell"#
        }
    }
}

fn run_vba_callback(host: OfficeHost) -> Result<(), String> {
    run_office_vba_script(vba_callback_script(host), "Office VBA callback")
}

fn run_vba_callback_on_main_thread(
    app: Option<&AppHandle>,
    host: OfficeHost,
) -> Result<(), String> {
    let Some(app) = app else {
        return run_vba_callback(host);
    };
    if host == OfficeHost::Word {
        let (sender, receiver) = mpsc::sync_channel(1);
        app.run_on_main_thread(move || {
            let _ = sender.send(execute_word_vba_apple_event());
        })
        .map_err(|error| format!("Unable to schedule the Office VBA callback: {error}"))?;
        return receiver
            .recv_timeout(Duration::from_secs(1_800))
            .map_err(|error| {
                format!("The Office VBA callback main-thread callback did not finish: {error}")
            })?;
    }
    run_office_vba_script_on_main_thread(
        app,
        vba_callback_script(host),
        "Office VBA callback",
    )
}

pub(crate) fn run_double_click_edit_macro(host: OfficeHost) -> Result<(), String> {
    let script = match host {
        OfficeHost::Word => {
            r#"tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "VisualTeX_DoubleClickEditSelected"
end tell"#
        }
        OfficeHost::Powerpoint => {
            r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "Microsoft PowerPoint has no active presentation"
run VB macro macro name "VisualTeX_DoubleClickEditSelected" list of parameters {}
end tell"#
        }
    };
    run_office_vba_script(script, "Office double-click edit macro")
}

pub(crate) fn run_word_image_double_click_edit_macro() -> Result<(), String> {
    run_office_vba_script(
        r#"tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "VisualTeX_EditSelectedImageFromNativeMonitor"
end tell"#,
        "Word image double-click edit macro",
    )
}

fn run_office_vba_script_subprocess(script: &str, label: &str) -> Result<(), String> {
    let output = Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(script)
        .output()
        .map_err(|error| format!("Unable to launch the {label}: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        Err(if detail.is_empty() {
            format!("The {label} failed")
        } else {
            format!("The {label} failed: {detail}")
        })
    }
}

#[cfg(target_os = "macos")]
thread_local! {
    // NSAppleScript is main-thread-bound in practice. Keep compiled callback
    // programs in a main-thread local cache instead of paying ~0.5 s to compile
    // the same Word/PowerPoint Apply script on every formula edit.
    static OFFICE_VBA_APPLESCRIPT_CACHE: RefCell<Vec<(String, Retained<NSAppleScript>)>> =
        const { RefCell::new(Vec::new()) };
}

#[cfg(target_os = "macos")]
fn execute_office_vba_apple_script(script: &str, label: &str) -> Result<(), String> {
    OFFICE_VBA_APPLESCRIPT_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let compiled_index = if let Some(index) = cache
            .iter()
            .position(|(cached_source, _)| cached_source == script)
        {
            index
        } else {
            let source = NSString::from_str(script);
            let Some(compiled_script) =
                NSAppleScript::initWithSource(NSAppleScript::alloc(), &source)
            else {
                return Err(format!("Unable to initialize the {label}"));
            };
            let mut compilation_error = None;
            if !unsafe { compiled_script.compileAndReturnError(Some(&mut compilation_error)) } {
                return Err(format!(
                    "Unable to compile the {label}: {compilation_error:?}"
                ));
            }
            cache.push((script.to_string(), compiled_script));
            cache.len() - 1
        };

        let compiled_script = &cache[compiled_index].1;
        let mut execution_error = None;
        let _result = unsafe { compiled_script.executeAndReturnError(Some(&mut execution_error)) };
        if let Some(error) = execution_error {
            Err(format!("The {label} failed: {error:?}"))
        } else {
            Ok(())
        }
    })
}

#[cfg(target_os = "macos")]
const fn apple_event_code(code: &[u8; 4]) -> u32 {
    u32::from_be_bytes(*code)
}

#[cfg(target_os = "macos")]
fn execute_word_vba_apple_event() -> Result<(), String> {
    // Word publishes this command as sWRD/1149 with the macro-name parameter
    // 5112 in Word.sdef. Sending that event directly avoids the fixed
    // NSAppleScript execution overhead while invoking the identical VBA entry.
    let bundle_id = NSString::from_str("com.microsoft.Word");
    let target = NSAppleEventDescriptor::descriptorWithBundleIdentifier(&bundle_id);
    let event = NSAppleEventDescriptor::appleEventWithEventClass_eventID_targetDescriptor_returnID_transactionID(
        apple_event_code(b"sWRD"),
        apple_event_code(b"1149"),
        Some(&target),
        kAutoGenerateReturnID as i16,
        kAnyTransactionID,
    );
    let macro_name = NSString::from_str("VisualTeX_ApplyPendingResult");
    let macro_descriptor = NSAppleEventDescriptor::descriptorWithString(&macro_name);
    event.setParamDescriptor_forKeyword(&macro_descriptor, apple_event_code(b"5112"));
    let reply = event
        .sendEventWithOptions_timeout_error(NSAppleEventSendOptions::DefaultOptions, 1_800.0)
        .map_err(|error| format!("The Office VBA callback failed: {error}"))?;
    if let Some(error_number) = reply.paramDescriptorForKeyword(keyErrorNumber) {
        let code = error_number.int32Value();
        if code != 0 {
            let detail = reply
                .paramDescriptorForKeyword(keyErrorString)
                .and_then(|descriptor| descriptor.stringValue())
                .map(|value| value.to_string())
                .unwrap_or_else(|| "Microsoft Word returned an Apple Event error".to_string());
            return Err(format!("The Office VBA callback failed ({code}): {detail}"));
        }
    }
    Ok(())
}

#[cfg(target_os = "macos")]
fn run_office_vba_script_on_main_thread(
    app: &AppHandle,
    script: &str,
    label: &str,
) -> Result<(), String> {
    // NSAppleScript may wait indefinitely when instantiated on a Tauri worker
    // thread. Dispatch its complete compile/execute lifetime to AppKit's main
    // thread, then return only the owned Result through a bounded channel.
    let script = script.to_string();
    let label = label.to_string();
    let callback_label = label.clone();
    let (sender, receiver) = mpsc::sync_channel(1);
    app.run_on_main_thread(move || {
        let result = execute_office_vba_apple_script(&script, &callback_label);
        let _ = sender.send(result);
    })
    .map_err(|error| format!("Unable to schedule the {label}: {error}"))?;
    receiver
        .recv_timeout(Duration::from_secs(1_800))
        .map_err(|error| format!("The {label} main-thread callback did not finish: {error}"))?
}

fn run_office_vba_script(script: &str, label: &str) -> Result<(), String> {
    run_office_vba_script_subprocess(script, label)
}

#[cfg(not(target_os = "macos"))]
fn run_office_vba_script_on_main_thread(
    _app: &AppHandle,
    script: &str,
    label: &str,
) -> Result<(), String> {
    run_office_vba_script_subprocess(script, label)
}

fn with_dispatch_pointer<T>(
    host: OfficeHost,
    session_id: &str,
    operation: impl FnOnce() -> Result<T, String>,
) -> Result<T, String> {
    let lock = match host {
        OfficeHost::Word => &WORD_DISPATCH_LOCK,
        OfficeHost::Powerpoint => &POWERPOINT_DISPATCH_LOCK,
    };
    let _guard = lock
        .lock()
        .map_err(|_| "VisualTeX Office dispatch lock is unavailable".to_string())?;
    let pointer = pointer_path(host)?;
    atomic_write_runtime(&pointer, session_id.as_bytes(), 0o600)?;
    let result = operation();
    let _ = fs::remove_file(pointer);
    result
}

fn scale_word_reference_geometry(
    reference_width_pt: f64,
    reference_height_pt: f64,
    reference_baseline_pt: f64,
    font_size_pt: f64,
) -> Result<WordGeometry, String> {
    if !reference_width_pt.is_finite()
        || !reference_height_pt.is_finite()
        || reference_width_pt <= 0.0
        || reference_height_pt <= 0.0
        || !reference_baseline_pt.is_finite()
        || !(-256.0..=0.0).contains(&reference_baseline_pt)
        || !font_size_pt.is_finite()
        || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&font_size_pt)
    {
        return Err("Word formula point-size reference geometry is invalid".to_string());
    }
    let point_scale = font_size_pt / WORD_REFERENCE_FONT_SIZE_PT;
    let width = reference_width_pt * point_scale;
    let height = reference_height_pt * point_scale;
    if !width.is_finite()
        || !height.is_finite()
        || width <= 0.0
        || height <= 0.0
        || width > 10_000.0
        || height > 10_000.0
    {
        return Err("Word formula point-size geometry is invalid".to_string());
    }
    // Word exposes Font.Position only as whole points. SVG and imported
    // picture dimensions are quantized independently, so a mathematical half
    // point can arrive just below the boundary (the real Times fixture reports
    // -1.491 pt instead of -1.5 pt). Use a small rounding tolerance for negative
    // descents; unlike a fixed offset, this changes only boundary cases and
    // preserves the distinct fraction/integral/sum/root baseline classes.
    let raw_baseline = reference_baseline_pt * point_scale;
    let baseline = if raw_baseline < 0.0 {
        (raw_baseline - 0.01).round()
    } else {
        0.0
    }
    .clamp(-256.0, 0.0) as i32;
    Ok(WordGeometry {
        width,
        height,
        baseline,
        font_size_pt,
        reference_width_pt,
        reference_height_pt,
        reference_baseline_pt,
    })
}

fn calculate_word_svg_geometry_with_scale(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
    width_scale: f64,
    height_scale: f64,
) -> Result<WordGeometry, String> {
    if !width.is_finite()
        || !height.is_finite()
        || !baseline.is_finite()
        || width <= 0.0
        || height <= 0.0
        || baseline < 0.0
        || baseline > height
        || !width_scale.is_finite()
        || !(0.5..=2.0).contains(&width_scale)
        || !height_scale.is_finite()
        || !(0.5..=2.0).contains(&height_scale)
    {
        return Err("Word formula SVG geometry is invalid".to_string());
    }
    let natural_width = width * 0.75 * width_scale;
    let natural_height = height * 0.75 * height_scale;
    let reference_scale = f64::min(1.0, MAX_WORD_WIDTH_PT / natural_width);
    let reference_width_pt = natural_width * reference_scale;
    let reference_height_pt = natural_height * reference_scale;
    let descent_ratio = (height - baseline) / height;
    // Preserve the fractional descent at the canonical 14 pt size. Word's
    // Font.Position is integral, so rounding this reference and rounding again
    // after point-size scaling visibly under-corrects short subscript formulas
    // such as L_z beside L^2.
    let reference_baseline_pt = (-(reference_height_pt * descent_ratio).max(0.0))
        .clamp(-256.0, 0.0);
    scale_word_reference_geometry(
        reference_width_pt,
        reference_height_pt,
        reference_baseline_pt,
        font_size_pt,
    )
}

#[cfg(test)]
fn calculate_word_svg_geometry(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
) -> Result<WordGeometry, String> {
    calculate_word_svg_geometry_with_scale(
        width,
        height,
        baseline,
        font_size_pt,
        WORD_TEX_IMAGE_VISUAL_SCALE,
        WORD_TEX_IMAGE_VISUAL_SCALE,
    )
}

fn calculate_word_svg_geometry_for_font(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
    formula_letter_font: Option<&str>,
) -> Result<WordGeometry, String> {
    let (width_scale, height_scale) = word_image_visual_scales(formula_letter_font);
    let geometry = calculate_word_svg_geometry_with_scale(
        width,
        height,
        baseline,
        font_size_pt,
        width_scale,
        height_scale,
    )?;
    // KaTeX and the non-Times letter-font replacements share MathJax's very
    // shallow SVG descent for superscript-only expressions. Word's integral
    // Font.Position then leaves x^2 visibly above adjacent OMML on screen and
    // by 0.5-0.67 pt in PDF ink measurements. Apply a canonical-size descent
    // floor only to that shallow class; fractions, integrals, sums and roots
    // already exceed it and keep their formula-specific baseline. Times uses
    // real text glyph metrics and retains its independent calibration.
    if formula_letter_font != Some("times")
        && geometry.reference_baseline_pt < 0.0
        && geometry.reference_baseline_pt > -WORD_TEX_SHALLOW_DESCENT_FLOOR_PT
    {
        scale_word_reference_geometry(
            geometry.reference_width_pt,
            geometry.reference_height_pt,
            -WORD_TEX_SHALLOW_DESCENT_FLOOR_PT,
            font_size_pt,
        )
    } else {
        Ok(geometry)
    }
}

fn calculate_word_geometry(
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
) -> Result<WordGeometry, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    if !export.width.is_finite()
        || !export.height.is_finite()
        || export.width <= 0.0
        || export.height <= 0.0
    {
        return Err("Word formula export has invalid dimensions".to_string());
    }

    // The editor persists the user's current selection on the Session. That
    // value must win over the immutable launch request; otherwise changing the
    // size in the editor is silently replaced by the size observed when the
    // Office command was first opened.
    let font_size_pt = session
        .font_size_pt
        .filter(|value| {
            value.is_finite() && *value >= MIN_WORD_FONT_SIZE_PT && *value <= MAX_WORD_FONT_SIZE_PT
        })
        .or_else(|| {
            request.font_size_pt.filter(|value| {
                value.is_finite()
                    && *value >= MIN_WORD_FONT_SIZE_PT
                    && *value <= MAX_WORD_FONT_SIZE_PT
            })
        })
        .or_else(|| {
            session
                .original_metadata
                .as_ref()
                .and_then(|metadata| metadata.font_size_pt)
                .filter(|value| {
                    value.is_finite()
                        && *value >= MIN_WORD_FONT_SIZE_PT
                        && *value <= MAX_WORD_FONT_SIZE_PT
                })
        })
        .unwrap_or(WORD_REFERENCE_FONT_SIZE_PT);
    let baseline = export
        .baseline
        .filter(|value| value.is_finite() && *value >= 0.0 && *value <= export.height)
        .ok_or_else(|| {
            "Word formula export is missing a valid mathematical baseline".to_string()
        })?;
    // Never silently fall back to the image bottom for a missing baseline. That
    // makes simple formulas look acceptable while subscripts, fractions and
    // tall delimiters are inserted at visibly different vertical positions on
    // Word versions that round InlineShape.Font.Position differently. Every
    // current renderer provides a mathematical baseline, so absence means the
    // artifact is incomplete and must not be committed with corrupt geometry.
    // Document import and edit replacement share this exact 14 pt reference
    // geometry path, including maximum width, descent and baseline rounding.
    calculate_word_svg_geometry_for_font(
        export.width,
        export.height,
        baseline,
        font_size_pt,
        Some(session.formula_letter_font.as_str()),
    )
}

fn calculate_powerpoint_geometry(
    request: &MacOfflinePowerPointRequest,
    session: &OfficeFormulaSession,
) -> Result<PowerPointGeometry, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "PowerPoint Session has no formula export".to_string())?;
    if !export.width.is_finite()
        || !export.height.is_finite()
        || export.width <= 0.0
        || export.height <= 0.0
    {
        return Err("PowerPoint formula export has invalid dimensions".to_string());
    }

    // MathJax exports at a stable 14 pt reference size. The imported SVG keeps
    // its vector paths, so a PowerPoint point size is represented by uniformly
    // scaling the natural SVG bounds rather than rasterizing or stretching it.
    let reference_width_pt = export.width * 0.75;
    let reference_height_pt = export.height * 0.75;
    if !reference_width_pt.is_finite()
        || !reference_height_pt.is_finite()
        || reference_width_pt <= 0.0
        || reference_height_pt <= 0.0
    {
        return Err("PowerPoint SVG reference geometry is invalid".to_string());
    }

    let original = session.original_metadata.as_ref();
    let previous_reference_height = request
        .reference_height_pt
        .filter(|value| value.is_finite() && *value > 0.0)
        .or_else(|| {
            original
                .and_then(|metadata| metadata.reference_height_pt)
                .filter(|value| value.is_finite() && *value > 0.0)
        })
        .or_else(|| {
            original
                .and_then(|metadata| metadata.render_height_px)
                .filter(|value| value.is_finite() && *value > 0.0)
                .map(|value| value * 0.75)
        });

    let committed_font_size = session.font_size_pt.filter(|value| {
        value.is_finite()
            && *value >= MIN_POWERPOINT_FONT_SIZE_PT
            && *value <= MAX_POWERPOINT_FONT_SIZE_PT
    });
    let declared_font_size = request
        .font_size_pt
        .filter(|value| {
            value.is_finite()
                && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                && *value <= MAX_POWERPOINT_FONT_SIZE_PT
        })
        .or_else(|| {
            original
                .and_then(|metadata| metadata.font_size_pt)
                .filter(|value| {
                    value.is_finite()
                        && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                        && *value <= MAX_POWERPOINT_FONT_SIZE_PT
                })
        });

    // The actual selected shape height wins when a user resized an existing SVG
    // manually. This converts that physical height back to an equivalent point
    // size and prevents the next edit from jumping to a stale stored value.
    let observed_font_size = previous_reference_height
        .map(|height| request.height / height * POWERPOINT_REFERENCE_FONT_SIZE_PT)
        .filter(|value| {
            value.is_finite()
                && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                && *value <= MAX_POWERPOINT_FONT_SIZE_PT
        });
    // The VBA adapter already converts a manually resized existing shape into
    // the Session's initial fontSizePt. After the editor opens, however, the
    // Session value is the user's explicit choice and must not be overwritten
    // by re-observing the old shape bounds during commit.
    let font_size_pt = committed_font_size
        .or(observed_font_size)
        .or(declared_font_size)
        .unwrap_or(DEFAULT_POWERPOINT_FONT_SIZE_PT);
    let point_scale = font_size_pt / POWERPOINT_REFERENCE_FONT_SIZE_PT;
    let target_width = reference_width_pt * point_scale;
    let target_height = reference_height_pt * point_scale;
    let center_x = request.left + request.width / 2.0;
    let center_y = request.top + request.height / 2.0;
    let left = center_x - target_width / 2.0;
    let top = center_y - target_height / 2.0;
    for (value, label) in [
        (left, "target left"),
        (top, "target top"),
        (target_width, "target width"),
        (target_height, "target height"),
    ] {
        validate_finite_geometry(value, label)?;
    }
    if target_width <= 0.0
        || target_height <= 0.0
        || target_width > 10_000.0
        || target_height > 10_000.0
    {
        return Err("PowerPoint target formula dimensions are invalid".to_string());
    }
    Ok(PowerPointGeometry {
        left,
        top,
        width: target_width,
        height: target_height,
        font_size_pt,
        reference_width_pt,
        reference_height_pt,
    })
}

fn materialize_result_png(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Office Session has no formula export".to_string())?;
    let png = export
        .png_base64
        .as_deref()
        .ok_or_else(|| "Offline Office Session requires a PNG export".to_string())
        .and_then(decode_png)?;
    let path = result_png_path(session.host, &session.id)?;
    atomic_write_runtime(&path, &png, 0o600)?;
    Ok(path)
}

fn decode_svg(value: &str) -> Result<Vec<u8>, String> {
    let bytes = BASE64_STANDARD
        .decode(value.trim())
        .map_err(|error| format!("Unable to decode the Office SVG export: {error}"))?;
    if bytes.is_empty() || bytes.len() > MAX_METADATA_BYTES * 4 {
        return Err("Office formula SVG export is empty or too large".to_string());
    }
    let svg = std::str::from_utf8(&bytes)
        .map_err(|_| "Office formula SVG export is not UTF-8".to_string())?;
    let normalized = svg.trim_start();
    if !normalized.starts_with("<svg")
        && !(normalized.starts_with("<?xml") && normalized.contains("<svg"))
    {
        return Err("Office formula export is not a valid SVG document".to_string());
    }
    let lower = normalized.to_ascii_lowercase();
    for forbidden in [
        "<!doctype",
        "<!entity",
        "<script",
        "<foreignobject",
        "href=\"http:",
        "href=\"https:",
        "href=\"//",
        "xlink:href=\"http:",
        "xlink:href=\"https:",
        "xlink:href=\"//",
    ] {
        if lower.contains(forbidden) {
            return Err("Office formula SVG export contains unsafe external content".to_string());
        }
    }
    Ok(bytes)
}

fn materialize_result_svg(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Office Session has no formula export".to_string())?;
    let svg = decode_svg(&export.svg_base64)?;
    let path = result_svg_path(session.host, &session.id)?;
    atomic_write_runtime(&path, &svg, 0o600)?;
    Ok(path)
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xffff_ffff_u32;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            let mask = 0_u32.wrapping_sub(crc & 1);
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    !crc
}

fn push_zip_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_zip_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

struct StoredZipEntry {
    name: Vec<u8>,
    crc: u32,
    size: u32,
    offset: u32,
}

fn build_stored_zip<N, C>(entries: &[(N, C)]) -> Result<Vec<u8>, String>
where
    N: AsRef<str>,
    C: AsRef<[u8]>,
{
    let entry_count = u16::try_from(entries.len())
        .map_err(|_| "Word SVG staging package has too many ZIP entries".to_string())?;
    let mut output = Vec::new();
    let mut records = Vec::with_capacity(entries.len());

    for (name, contents) in entries {
        let name = name.as_ref();
        let contents = contents.as_ref();
        let name_bytes = name.as_bytes();
        let name_length = u16::try_from(name_bytes.len())
            .map_err(|_| "Word SVG staging package contains an overlong ZIP path".to_string())?;
        let size = u32::try_from(contents.len())
            .map_err(|_| "Word SVG staging package entry is too large".to_string())?;
        let offset = u32::try_from(output.len())
            .map_err(|_| "Word SVG staging package is too large".to_string())?;
        let checksum = crc32(contents);

        push_zip_u32(&mut output, 0x0403_4b50);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 0x0800);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 33);
        push_zip_u32(&mut output, checksum);
        push_zip_u32(&mut output, size);
        push_zip_u32(&mut output, size);
        push_zip_u16(&mut output, name_length);
        push_zip_u16(&mut output, 0);
        output.extend_from_slice(name_bytes);
        output.extend_from_slice(contents);

        records.push(StoredZipEntry {
            name: name_bytes.to_vec(),
            crc: checksum,
            size,
            offset,
        });
    }

    let central_offset = u32::try_from(output.len())
        .map_err(|_| "Word SVG staging package is too large".to_string())?;
    for record in &records {
        let name_length = u16::try_from(record.name.len())
            .map_err(|_| "Word SVG staging package contains an overlong ZIP path".to_string())?;
        push_zip_u32(&mut output, 0x0201_4b50);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 0x0800);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 33);
        push_zip_u32(&mut output, record.crc);
        push_zip_u32(&mut output, record.size);
        push_zip_u32(&mut output, record.size);
        push_zip_u16(&mut output, name_length);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u32(&mut output, 0);
        push_zip_u32(&mut output, record.offset);
        output.extend_from_slice(&record.name);
    }
    let central_size = u32::try_from(output.len())
        .map_err(|_| "Word SVG staging package is too large".to_string())?
        .checked_sub(central_offset)
        .ok_or_else(|| "Word SVG staging package central directory is invalid".to_string())?;

    push_zip_u32(&mut output, 0x0605_4b50);
    push_zip_u16(&mut output, 0);
    push_zip_u16(&mut output, 0);
    push_zip_u16(&mut output, entry_count);
    push_zip_u16(&mut output, entry_count);
    push_zip_u32(&mut output, central_size);
    push_zip_u32(&mut output, central_offset);
    push_zip_u16(&mut output, 0);
    Ok(output)
}

fn build_word_svg_docx(
    svg: &[u8],
    png: &[u8],
    width_points: f64,
    height_points: f64,
) -> Result<Vec<u8>, String> {
    let width_emu = (width_points * 12_700.0).round();
    let height_emu = (height_points * 12_700.0).round();
    if !width_emu.is_finite()
        || !height_emu.is_finite()
        || width_emu <= 0.0
        || height_emu <= 0.0
        || width_emu > i64::MAX as f64
        || height_emu > i64::MAX as f64
    {
        return Err("Word SVG staging package dimensions are invalid".to_string());
    }
    let width_emu = width_emu as i64;
    let height_emu = height_emu as i64;

    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let package_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let document_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.png"/>
  <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.svg"/>
</Relationships>"#;
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="{width_emu}" cy="{height_emu}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="1" name="VisualTeX Formula" descr="VisualTeX SVG formula"/>
            <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="0" name="formula.svg"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rIdPng" cstate="print">
                      <a:extLst>
                        <a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                      </a:extLst>
                    </a:blip>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:noFill/><a:ln><a:noFill/></a:ln>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>"#
    );

    build_stored_zip(&[
        ("[Content_Types].xml", content_types.as_slice()),
        ("_rels/.rels", package_relationships.as_slice()),
        ("word/document.xml", document.as_bytes()),
        (
            "word/_rels/document.xml.rels",
            document_relationships.as_slice(),
        ),
        ("word/media/formula.png", png),
        ("word/media/formula.svg", svg),
    ])
}

struct WordSvgBatchEntry {
    svg: Vec<u8>,
    png: Vec<u8>,
    width_points: f64,
    height_points: f64,
}

fn build_word_svg_batch_docx(entries: &[WordSvgBatchEntry]) -> Result<Vec<u8>, String> {
    if entries.is_empty() || entries.len() > 1000 {
        return Err("Word SVG redraw batch must contain 1 to 1000 drawings".to_string());
    }
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let package_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let mut document_relationships = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
    );
    let mut document = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\">\n  <w:body>\n",
    );
    let mut zip_entries = vec![
        ("[Content_Types].xml".to_string(), content_types.to_vec()),
        ("_rels/.rels".to_string(), package_relationships.to_vec()),
    ];

    for (index, entry) in entries.iter().enumerate() {
        let width_emu = (entry.width_points * 12_700.0).round();
        let height_emu = (entry.height_points * 12_700.0).round();
        if !width_emu.is_finite()
            || !height_emu.is_finite()
            || width_emu <= 0.0
            || height_emu <= 0.0
            || width_emu > i64::MAX as f64
            || height_emu > i64::MAX as f64
        {
            return Err("Word SVG redraw batch dimensions are invalid".to_string());
        }
        let item_number = index + 1;
        let png_rel = format!("rIdPng{item_number}");
        let svg_rel = format!("rIdSvg{item_number}");
        let png_name = format!("formula-{item_number}.png");
        let svg_name = format!("formula-{item_number}.svg");
        document_relationships.push_str(&format!(
            "  <Relationship Id=\"{png_rel}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{png_name}\"/>\n  <Relationship Id=\"{svg_rel}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{svg_name}\"/>\n",
        ));
        document.push_str(&format!(
            r#"    <w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{}" cy="{}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{}" name="VisualTeX Formula {}" descr="VisualTeX SVG formula"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="{}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{}" cstate="print"><a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip r:embed="{}"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
"#,
            width_emu as i64,
            height_emu as i64,
            item_number,
            item_number,
            svg_name,
            png_rel,
            svg_rel,
            width_emu as i64,
            height_emu as i64,
        ));
        zip_entries.push((
            format!("word/media/{png_name}"),
            entry.png.clone(),
        ));
        zip_entries.push((
            format!("word/media/{svg_name}"),
            entry.svg.clone(),
        ));
    }
    document_relationships.push_str("</Relationships>");
    document.push_str(
        "    <w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>\n  </w:body>\n</w:document>",
    );
    zip_entries.push((
        "word/document.xml".to_string(),
        document.into_bytes(),
    ));
    zip_entries.push((
        "word/_rels/document.xml.rels".to_string(),
        document_relationships.into_bytes(),
    ));
    build_stored_zip(&zip_entries)
}

fn build_word_omml_batch_docx(omml_fragments: &[String]) -> Result<Vec<u8>, String> {
    if omml_fragments.is_empty() || omml_fragments.len() > 1000 {
        return Err("Word OMML batch must contain 1 to 1000 equations".to_string());
    }
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let package_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let mut document = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body>\n",
    );
    for fragment in omml_fragments {
        let normalized = fragment.trim();
        if !normalized.starts_with("<m:oMath")
            || !normalized.ends_with("</m:oMath>")
            || normalized.contains("<!DOCTYPE")
            || normalized.contains("<!ENTITY")
        {
            return Err("Word OMML batch contains an invalid equation fragment".to_string());
        }
        document.push_str("<w:p>");
        document.push_str(normalized);
        document.push_str("</w:p>\n");
    }
    document.push_str("<w:sectPr/></w:body></w:document>");
    build_stored_zip(&[
        ("[Content_Types].xml", content_types.as_slice()),
        ("_rels/.rels", package_relationships.as_slice()),
        ("word/document.xml", document.as_bytes()),
    ])
}

fn materialize_word_svg_package(
    session: &OfficeFormulaSession,
    geometry: WordGeometry,
) -> Result<(PathBuf, PathBuf, PathBuf), String> {
    if session.host != OfficeHost::Word {
        return Err("Word SVG package materialization requires a Word Session".to_string());
    }
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    let svg = decode_svg(&export.svg_base64)?;
    let png = export
        .png_base64
        .as_deref()
        .ok_or_else(|| "Word SVG staging requires a PNG compatibility preview".to_string())
        .and_then(decode_png)?;
    let svg_path = result_svg_path(OfficeHost::Word, &session.id)?;
    let png_path = result_png_path(OfficeHost::Word, &session.id)?;
    let document_path = result_word_svg_docx_path(&session.id)?;
    atomic_write_runtime(&svg_path, &svg, 0o600)?;
    atomic_write_runtime(&png_path, &png, 0o600)?;
    let package = build_word_svg_docx(&svg, &png, geometry.width, geometry.height)?;
    atomic_write_runtime(&document_path, &package, 0o600)?;

    // Persist a formula-scoped image package even when the current Word output
    // is OMML. A later OMML-to-image conversion can then stay entirely inside
    // Word instead of waking the renderer and rebuilding the same SVG.
    let (cached_svg_path, cached_document_path, cached_png_path) =
        word_image_cache_paths(&session.formula_id)?;
    atomic_write_runtime(&cached_svg_path, &svg, 0o600)?;
    atomic_write_runtime(&cached_document_path, &package, 0o600)?;
    atomic_write_runtime(&cached_png_path, &png, 0o600)?;
    Ok((svg_path, document_path, png_path))
}

fn materialize_powerpoint_svg(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    if session.host != OfficeHost::Powerpoint {
        return Err("PowerPoint SVG materialization requires a PowerPoint Session".to_string());
    }
    materialize_result_svg(session)
}

fn commit_word(
    app: Option<&AppHandle>,
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    metadata: &str,
    canonical_latex: &str,
    geometry: WordGeometry,
) -> Result<(), String> {
    let commit_started = Instant::now();
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    let omml_base64 = export
        .omml_base64
        .as_deref()
        .ok_or_else(|| "Word formula export has no OMML payload".to_string())?;
    let omml_bytes = URL_SAFE_NO_PAD
        .decode(omml_base64)
        .map_err(|_| "Word formula OMML payload is not valid Base64URL".to_string())?;
    if omml_bytes.is_empty() || omml_bytes.len() > MAX_OMML_BYTES {
        return Err("Word formula OMML payload is empty or too large".to_string());
    }
    let omml = std::str::from_utf8(&omml_bytes)
        .map_err(|_| "Word formula OMML payload is not UTF-8".to_string())?;
    if !omml.trim_start().starts_with("<m:oMath")
        || !omml.contains("http://schemas.openxmlformats.org/officeDocument/2006/math")
        || omml.contains("<!DOCTYPE")
        || omml.contains("<!ENTITY")
    {
        return Err("Word formula OMML payload is not a safe Office Math fragment".to_string());
    }
    let native_document_path = if request.native_equation {
        let omml_docx_base64 = export
            .omml_docx_base64
            .as_deref()
            .ok_or_else(|| "Word formula export has no native DOCX payload".to_string())?;
        let omml_docx = URL_SAFE_NO_PAD
            .decode(omml_docx_base64)
            .map_err(|_| "Word formula native DOCX payload is not valid Base64URL".to_string())?;
        if omml_docx.len() < 128
            || omml_docx.len() > MAX_OMML_BYTES * 8
            || !omml_docx.starts_with(b"PK\x03\x04")
        {
            return Err("Word formula native DOCX payload is invalid or too large".to_string());
        }
        let path = native_word_document_path(&session.formula_id)?;
        atomic_write(&path, &omml_docx, 0o600)?;
        Some(path)
    } else {
        None
    };

    // Every successful Word export refreshes the durable image cache. Native
    // OMML commits do not pass the Session-scoped image paths to VBA, but the
    // formula-scoped cache makes the later OMML-to-image command a sub-second
    // in-Word transaction with no renderer wake-up.
    let prepared_image_artifacts = materialize_word_svg_package(session, geometry)?;
    let image_artifacts = if request.native_equation {
        None
    } else {
        Some(prepared_image_artifacts)
    };
    queue_editor_performance(
        OfficeHost::Word,
        &session.id,
        "apply-word-artifacts-ready",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({ "nativeEquation": request.native_equation }),
    );
    let source_marker = request
        .source_object_id
        .clone()
        .or_else(|| request.encoded_metadata.clone())
        .unwrap_or_default();
    let pending_marker = request.pending_marker.clone().unwrap_or_default();
    let latex = canonical_latex.trim();
    if latex.is_empty() {
        return Err("Word native-equation conversion requires non-empty LaTeX".to_string());
    }
    let latex_base64 = URL_SAFE_NO_PAD.encode(latex.as_bytes());
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session.id.clone()),
        ("action", "commit".to_string()),
        ("host", "word".to_string()),
        (
            "operation",
            request
                .operation
                .clone()
                .unwrap_or_else(|| "formula".to_string()),
        ),
        (
            "performanceTrace",
            if word_performance_trace_enabled() { "1" } else { "0" }.to_string(),
        ),
        ("mode", request.mode.clone()),
        ("formulaId", session.formula_id.clone()),
        ("displayMode", session.display_mode.clone()),
        (
            "numbered",
            if session.numbered { "1" } else { "0" }.to_string(),
        ),
        (
            "nativeEquation",
            if request.native_equation { "1" } else { "0" }.to_string(),
        ),
        (
            "imagePath",
            image_artifacts
                .as_ref()
                .map(|artifacts| artifacts.0.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        (
            "vectorDocumentPath",
            image_artifacts
                .as_ref()
                .map(|artifacts| artifacts.1.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        (
            "fallbackImagePath",
            image_artifacts
                .as_ref()
                .map(|artifacts| artifacts.2.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        ("metadata", metadata.to_string()),
        ("latexBase64", latex_base64),
        ("ommlBase64", omml_base64.to_string()),
        (
            "nativeDocumentPath",
            native_document_path
                .as_ref()
                .map(|path| path.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        ("pendingMarker", pending_marker),
        ("sourceMarker", source_marker),
        (
            "sourceDocumentId",
            request.source_document_id.clone().unwrap_or_default(),
        ),
        ("widthPoints", format!("{:.6}", geometry.width)),
        ("heightPoints", format!("{:.6}", geometry.height)),
        ("baseline", geometry.baseline.to_string()),
        ("fontSizePt", format!("{:.6}", geometry.font_size_pt)),
        (
            "referenceWidthPt",
            format!("{:.6}", geometry.reference_width_pt),
        ),
        (
            "referenceHeightPt",
            format!("{:.6}", geometry.reference_height_pt),
        ),
        (
            "referenceBaselinePt",
            format!("{:.6}", geometry.reference_baseline_pt),
        ),
        (
            "inkCenterYRatio",
            session
                .export_result
                .as_ref()
                .and_then(|value| value.ink_center_y_ratio)
                .filter(|value| value.is_finite() && (0.0..=1.0).contains(value))
                .map(|value| format!("{value:.8}"))
                .unwrap_or_default(),
        ),
    ])?;
    atomic_write_runtime(
        &dispatch_path(OfficeHost::Word, &session.id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    queue_editor_performance(
        OfficeHost::Word,
        &session.id,
        "apply-word-dispatch-ready",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    with_dispatch_pointer(OfficeHost::Word, &session.id, || {
        run_vba_callback_on_main_thread(app, OfficeHost::Word)
    })?;
    queue_editor_performance(
        OfficeHost::Word,
        &session.id,
        "apply-word-vba-complete",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    Ok(())
}

fn commit_powerpoint(
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    metadata: &str,
    geometry: PowerPointGeometry,
) -> Result<(), String> {
    let powerpoint = request
        .power_point
        .as_ref()
        .ok_or_else(|| "PowerPoint request geometry is missing".to_string())?;
    let image_path = materialize_powerpoint_svg(session)?;
    let fallback_image_path = materialize_result_png(session).ok();
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session.id.clone()),
        ("action", "commit".to_string()),
        ("host", "powerpoint".to_string()),
        ("mode", request.mode.clone()),
        ("formulaId", session.formula_id.clone()),
        ("displayMode", "block".to_string()),
        ("numbered", "0".to_string()),
        ("imagePath", image_path.to_string_lossy().to_string()),
        (
            "fallbackImagePath",
            fallback_image_path
                .as_ref()
                .map(|path| path.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        ("metadata", metadata.to_string()),
        (
            "pendingMarker",
            request.pending_marker.clone().unwrap_or_default(),
        ),
        (
            "sourceMarker",
            request.encoded_metadata.clone().unwrap_or_default(),
        ),
        ("sourceShapeName", powerpoint.shape_name.clone()),
        ("shapeName", format!("VisualTeX_{}", session.formula_id)),
        ("targetLeft", format!("{:.6}", geometry.left)),
        ("targetTop", format!("{:.6}", geometry.top)),
        ("targetWidth", format!("{:.6}", geometry.width)),
        ("targetHeight", format!("{:.6}", geometry.height)),
        ("fontSizePt", format!("{:.6}", geometry.font_size_pt)),
        (
            "referenceWidthPt",
            format!("{:.6}", geometry.reference_width_pt),
        ),
        (
            "referenceHeightPt",
            format!("{:.6}", geometry.reference_height_pt),
        ),
        ("rotation", format!("{:.6}", powerpoint.rotation)),
        ("zOrder", powerpoint.z_order.to_string()),
        (
            "presentationIdentity",
            powerpoint.presentation_identity.clone(),
        ),
        ("slideIndex", powerpoint.slide_index.to_string()),
        ("slideId", powerpoint.slide_id.to_string()),
    ])?;
    atomic_write_runtime(
        &dispatch_path(OfficeHost::Powerpoint, &session.id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Powerpoint, &session.id, || {
        run_vba_callback(OfficeHost::Powerpoint)
    })?;
    Ok(())
}

fn cancel_host(request: &MacOfflineSessionRequest) -> Result<(), String> {
    let host = if request.host == "word" {
        OfficeHost::Word
    } else {
        OfficeHost::Powerpoint
    };
    let entries = vec![
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", request.session_id.clone()),
        ("action", "cancel".to_string()),
        ("host", request.host.clone()),
        ("mode", request.mode.clone()),
        (
            "pendingMarker",
            request.pending_marker.clone().unwrap_or_default(),
        ),
        (
            "sourceDocumentId",
            request.source_document_id.clone().unwrap_or_default(),
        ),
    ];
    let dispatch = dispatch_text(&entries)?;
    atomic_write(
        &dispatch_path(host, &request.session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    if request.mode == "create" {
        with_dispatch_pointer(host, &request.session_id, || run_vba_callback(host))?;
    }
    Ok(())
}

fn parse_formula_restore_manifest(
    request: &MacOfflineSessionRequest,
) -> Result<(String, Vec<MacOfflineFormulaRestoreTarget>), String> {
    let path = formula_restore_source_path(&request.session_id)?;
    let metadata = fs::metadata(&path)
        .map_err(|error| format!("Unable to inspect {}: {error}", path.display()))?;
    if metadata.file_type().is_symlink()
        || !metadata.is_file()
        || metadata.len() == 0
        || metadata.len() as usize > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES
    {
        return Err("Word formula restore source has an invalid size".to_string());
    }
    let text = fs::read_to_string(&path)
        .map_err(|error| format!("Unable to read {}: {error}", path.display()))?;
    let mut values = std::collections::HashMap::<String, String>::new();
    for line in text.lines() {
        let (key, value) = line
            .split_once('=')
            .ok_or_else(|| "Word formula restore source contains an invalid line".to_string())?;
        if key.is_empty()
            || !key.bytes().all(|byte| byte.is_ascii_alphanumeric())
            || values.insert(key.to_string(), value.to_string()).is_some()
        {
            return Err("Word formula restore source contains an invalid key".to_string());
        }
    }
    let required = |key: &str| {
        values
            .get(key)
            .map(String::as_str)
            .ok_or_else(|| format!("Word formula restore source is missing {key}"))
    };
    if required("protocolVersion")? != OFFLINE_PROTOCOL_VERSION.to_string()
        || required("sessionId")? != request.session_id
    {
        return Err("Word formula restore source identity is invalid".to_string());
    }
    let item_count = required("itemCount")?
        .parse::<usize>()
        .map_err(|_| "Word formula restore item count is invalid".to_string())?;
    if item_count == 0 || item_count > 1000 {
        return Err("Word formula restore supports 1 to 1000 formulas".to_string());
    }
    let target_text = String::from_utf8(
        URL_SAFE_NO_PAD
            .decode(required("targetTextBase64")?)
            .map_err(|error| format!("Unable to decode Word restore target text: {error}"))?,
    )
    .map_err(|_| "Word formula restore target text is not UTF-8".to_string())?;
    if target_text.len() > MAX_LATEX_REDRAW_SOURCE_BYTES as usize {
        return Err("Word formula restore target exceeds 5 MB".to_string());
    }
    let source_kind = required("sourceKind")?;
    if !matches!(source_kind, "omml" | "image") {
        return Err("Word formula restore source kind is invalid".to_string());
    }

    let mut targets = Vec::with_capacity(item_count);
    let mut previous_end = 0usize;
    for index in 0..item_count {
        let prefix = format!("item{index}");
        let start = required(&format!("{prefix}sourceStart"))?
            .parse::<usize>()
            .map_err(|_| "Word formula restore sourceStart is invalid".to_string())?;
        let end = required(&format!("{prefix}sourceEnd"))?
            .parse::<usize>()
            .map_err(|_| "Word formula restore sourceEnd is invalid".to_string())?;
        let source_text = String::from_utf8(
            URL_SAFE_NO_PAD
                .decode(required(&format!("{prefix}sourceTextBase64"))?)
                .map_err(|error| format!("Unable to decode Word formula source text: {error}"))?,
        )
        .map_err(|_| "Word formula source text is not UTF-8".to_string())?;
        validate_formula_restore_range(
            &mut previous_end,
            start,
            end,
            &source_text,
        )?;
        let display_mode = required(&format!("{prefix}displayMode"))?.to_string();
        if !matches!(display_mode.as_str(), "inline" | "block") {
            return Err("Word formula restore display mode is invalid".to_string());
        }
        let font_size_pt = required(&format!("{prefix}fontSizePt"))?
            .parse::<f64>()
            .map_err(|_| "Word formula restore font size is invalid".to_string())?;
        if !font_size_pt.is_finite()
            || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&font_size_pt)
        {
            return Err("Word formula restore font size is outside the supported range".to_string());
        }
        let payload = String::from_utf8(
            URL_SAFE_NO_PAD
                .decode(required(&format!("{prefix}payloadBase64"))?)
                .map_err(|error| format!("Unable to decode Word formula restore payload: {error}"))?,
        )
        .map_err(|_| "Word formula restore payload is not UTF-8".to_string())?;
        if payload.is_empty() || payload.len() > MAX_OMML_BYTES.max(MAX_METADATA_BYTES) {
            return Err("Word formula restore payload is invalid or excessive".to_string());
        }
        let (math_ml, latex) = if source_kind == "omml" {
            (Some(word_omml_to_mathml(&payload)?), None)
        } else {
            let metadata = decode_metadata(&payload)?;
            let latex = canonical_document_formula_latex(&metadata)?;
            if metadata.display_mode != display_mode {
                return Err("Word image formula display metadata changed".to_string());
            }
            (None, Some(latex))
        };
        targets.push(MacOfflineFormulaRestoreTarget {
            source_start: start,
            source_end: end,
            source_text,
            display_mode,
            font_size_pt,
            source_kind: source_kind.to_string(),
            math_ml,
            latex,
        });
    }
    Ok((target_text, targets))
}

fn extract_omath_fragment(word_open_xml: &str) -> Result<String, String> {
    if word_open_xml.is_empty()
        || word_open_xml.len() > MAX_OMML_BYTES
        || word_open_xml.contains("<!DOCTYPE")
        || word_open_xml.contains("<!ENTITY")
    {
        return Err("Word OMML XML is invalid or excessive".to_string());
    }
    let starts = ["<m:oMath>", "<m:oMath "];
    let start = starts
        .iter()
        .filter_map(|marker| word_open_xml.find(marker))
        .min()
        .ok_or_else(|| "Word range XML does not contain one native OMath".to_string())?;
    let relative_end = word_open_xml[start..]
        .find("</m:oMath>")
        .ok_or_else(|| "Word native OMath XML is incomplete".to_string())?;
    let end = start + relative_end + "</m:oMath>".len();
    if word_open_xml[end..].contains("<m:oMath>")
        || word_open_xml[end..].contains("<m:oMath ")
    {
        return Err("Word formula restore item contains multiple OMath objects".to_string());
    }
    let mut fragment = word_open_xml[start..end].to_string();
    let tag_end = fragment
        .find('>')
        .ok_or_else(|| "Word native OMath start tag is incomplete".to_string())?;
    let mut declarations = String::new();
    let mut cursor = 0usize;
    while let Some(offset) = word_open_xml[cursor..].find("xmlns") {
        let attribute_start = cursor + offset;
        let Some(equal_offset) = word_open_xml[attribute_start..].find('=') else {
            break;
        };
        let equal = attribute_start + equal_offset;
        let name = word_open_xml[attribute_start..equal].trim();
        if !name.bytes().all(|byte| {
            byte.is_ascii_alphanumeric() || matches!(byte, b':' | b'-' | b'_')
        }) {
            cursor = equal + 1;
            continue;
        }
        let Some(quote) = word_open_xml.as_bytes().get(equal + 1).copied() else {
            break;
        };
        if quote != b'\'' && quote != b'"' {
            cursor = equal + 1;
            continue;
        }
        let value_start = equal + 2;
        let Some(value_end_offset) = word_open_xml[value_start..].find(quote as char) else {
            break;
        };
        let value_end = value_start + value_end_offset;
        if !fragment[..tag_end].contains(&format!("{name}="))
            && !declarations.contains(&format!(" {name}="))
        {
            declarations.push(' ');
            declarations.push_str(name);
            declarations.push('=');
            declarations.push(quote as char);
            declarations.push_str(&word_open_xml[value_start..value_end]);
            declarations.push(quote as char);
        }
        cursor = value_end + 1;
    }
    if !fragment[..tag_end].contains("xmlns:m=") && !declarations.contains(" xmlns:m=") {
        declarations.push_str(
            " xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"",
        );
    }
    if !fragment[..tag_end].contains("xmlns:w=") && !declarations.contains(" xmlns:w=") {
        declarations.push_str(
            " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
        );
    }
    fragment.insert_str(tag_end, &declarations);
    Ok(fragment)
}

fn find_xml_start_tag(value: &str, tag: &str, start: usize) -> Option<usize> {
    let exact = format!("<{tag}>");
    let attributed = format!("<{tag} ");
    [exact.as_str(), attributed.as_str()]
        .iter()
        .filter_map(|marker| value[start..].find(marker).map(|offset| start + offset))
        .min()
}

fn strip_visualtex_numbered_equation_array(fragment: &str) -> Option<String> {
    // VisualTeX display numbering is represented as one OMML Equation Array:
    //   <m:eqArr><m:e>FORMULA <m:r><m:t>#</m:t>...</m:r>
    //                    <m:d>REF VT_N_...</m:d></m:e></m:eqArr>
    // Word expands WordOpenXML requested from only the FORMULA Range back to
    // that complete OMath container, so VBA cannot crop the XML by Range alone.
    // Strip only this exact VisualTeX-owned signature before the standard Word
    // OMML -> MathML transform. Ordinary Word Equation Arrays remain untouched.
    let eq_array_start = find_xml_start_tag(fragment, "m:eqArr", 0)?;
    let eq_array_end_offset = fragment[eq_array_start..].find("</m:eqArr>")?;
    let eq_array_end = eq_array_start + eq_array_end_offset;
    let eq_array = &fragment[eq_array_start..eq_array_end];
    if !eq_array.contains("REF VT_N_") {
        return None;
    }

    let marker_relative = eq_array.rfind("<m:t>#</m:t>")?;
    let marker = eq_array_start + marker_relative;
    if !fragment[marker..eq_array_end].contains("REF VT_N_") {
        return None;
    }
    let marker_run_start = ["<m:r>", "<m:r "]
        .iter()
        .filter_map(|tag| fragment[..marker].rfind(tag))
        .max()?;
    if marker_run_start < eq_array_start {
        return None;
    }
    let marker_run_end_offset = fragment[marker..eq_array_end].find("</m:r>")?;
    let marker_run_end = marker + marker_run_end_offset + "</m:r>".len();
    if marker_run_end >= eq_array_end {
        return None;
    }

    let properties_end = fragment[eq_array_start..marker_run_start]
        .find("</m:eqArrPr>")
        .map(|offset| eq_array_start + offset + "</m:eqArrPr>".len())
        .unwrap_or(eq_array_start);
    let entry_start = find_xml_start_tag(fragment, "m:e", properties_end)?;
    if entry_start >= marker_run_start {
        return None;
    }
    let entry_tag_end = fragment[entry_start..marker_run_start].find('>')? + entry_start;
    let body_start = entry_tag_end + 1;
    if body_start >= marker_run_start {
        return None;
    }
    let body = fragment[body_start..marker_run_start].trim();
    if body.is_empty() {
        return None;
    }

    let omath_tag_end = fragment.find('>')?;
    let omath_start = &fragment[..=omath_tag_end];
    Some(format!("{omath_start}{body}</m:oMath>"))
}

fn decode_xslt_output(bytes: &[u8]) -> Result<String, String> {
    if bytes.starts_with(&[0xff, 0xfe]) || bytes.starts_with(&[0xfe, 0xff]) {
        let little_endian = bytes.starts_with(&[0xff, 0xfe]);
        let body = &bytes[2..];
        if body.len() % 2 != 0 {
            return Err("Word MathML output contains invalid UTF-16".to_string());
        }
        let units = body
            .chunks_exact(2)
            .map(|pair| {
                if little_endian {
                    u16::from_le_bytes([pair[0], pair[1]])
                } else {
                    u16::from_be_bytes([pair[0], pair[1]])
                }
            })
            .collect::<Vec<_>>();
        return String::from_utf16(&units)
            .map_err(|_| "Word MathML output is not valid UTF-16".to_string());
    }
    String::from_utf8(bytes.to_vec())
        .map_err(|_| "Word MathML output is not valid UTF-8".to_string())
}

pub(crate) fn word_omml_to_mathml(word_open_xml: &str) -> Result<String, String> {
    let fragment = extract_omath_fragment(word_open_xml)?;
    let fragment = strip_visualtex_numbered_equation_array(&fragment).unwrap_or(fragment);
    let stylesheet = Path::new(
        "/Applications/Microsoft Word.app/Contents/Resources/omml2mathml.xsl",
    );
    if !stylesheet.is_file() {
        return Err("Microsoft Word OMML conversion stylesheet is missing".to_string());
    }
    let mut child = Command::new("/usr/bin/xsltproc")
        .args(["--nonet"])
        .arg(stylesheet)
        .arg("-")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|error| format!("Unable to start the Word OMML converter: {error}"))?;
    child
        .stdin
        .take()
        .ok_or_else(|| "Word OMML converter stdin is unavailable".to_string())?
        .write_all(fragment.as_bytes())
        .map_err(|error| format!("Unable to send OMML to the Word converter: {error}"))?;
    let output = child
        .wait_with_output()
        .map_err(|error| format!("Unable to wait for the Word OMML converter: {error}"))?;
    if !output.status.success() {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        return Err(if detail.is_empty() {
            "Microsoft Word could not convert the native formula to MathML".to_string()
        } else {
            format!("Microsoft Word could not convert the native formula to MathML: {detail}")
        });
    }
    let math_ml = decode_xslt_output(&output.stdout)?;
    if math_ml.trim().is_empty() || math_ml.len() > MAX_OMML_BYTES {
        return Err("Microsoft Word returned empty or excessive MathML".to_string());
    }
    Ok(math_ml)
}

fn document_import_request_data(
    request: &MacOfflineSessionRequest,
) -> Result<MacOfflineDocumentImportPublicRequest, String> {
    let operation = request.operation.as_deref().unwrap_or("");
    if !matches!(operation, "documentImport" | "latexRedraw" | "formulaRestore")
        || request.host != "word"
    {
        return Err("Offline Office request is not a Word document operation".to_string());
    }
    let document_import = request
        .document_import
        .as_ref()
        .ok_or_else(|| "Document import request is missing insertion information".to_string())?;
    let (source, restore_targets) = if operation == "latexRedraw" {
        let source_path = session_directory(OfficeHost::Word, &request.session_id)?
            .join(LATEX_REDRAW_SOURCE_FILE);
        let bytes = fs::read(&source_path)
            .map_err(|error| format!("Unable to read {}: {error}", source_path.display()))?;
        if bytes.is_empty() || bytes.len() as u64 > MAX_LATEX_REDRAW_SOURCE_BYTES {
            return Err("LaTeX redraw source must contain 1 byte to 5 MB".to_string());
        }
        (
            Some(
                String::from_utf8(bytes)
                    .map_err(|_| "LaTeX redraw source is not valid UTF-8".to_string())?,
            ),
            None,
        )
    } else if operation == "formulaRestore" {
        let (source, targets) = parse_formula_restore_manifest(request)?;
        (Some(source), Some(targets))
    } else {
        (None, None)
    };
    Ok(MacOfflineDocumentImportPublicRequest {
        protocol_version: request.protocol_version,
        session_id: request.session_id.clone(),
        host: request.host.clone(),
        source_document_id: request.source_document_id.clone().ok_or_else(|| {
            "Document import request is missing the Word document identity".to_string()
        })?,
        bookmark_name: document_import.bookmark_name.clone(),
        default_font_size_pt: document_import.default_font_size_pt,
        operation: operation.to_string(),
        redraw_scope: document_import.redraw_scope.clone(),
        output_kind: document_import.output_kind.clone(),
        source_kind: document_import.source_kind.clone(),
        source,
        restore_targets,
    })
}

fn document_formula_file_path(
    session_id: &str,
    formula_id: &str,
    extension: &str,
) -> Result<PathBuf, String> {
    validate_uuid(formula_id, "Document formula id")?;
    if !extension.bytes().all(|byte| byte.is_ascii_alphanumeric()) {
        return Err("Document formula file extension is invalid".to_string());
    }
    Ok(session_directory(OfficeHost::Word, session_id)?
        .join(format!("document-formula-{formula_id}.{extension}")))
}

fn validate_document_omml_payload(value: &str) -> Result<(), String> {
    let bytes = URL_SAFE_NO_PAD
        .decode(value)
        .map_err(|_| "Document formula OMML payload is not valid Base64URL".to_string())?;
    if bytes.is_empty() || bytes.len() > MAX_OMML_BYTES {
        return Err("Document formula OMML payload is empty or too large".to_string());
    }
    let omml = std::str::from_utf8(&bytes)
        .map_err(|_| "Document formula OMML payload is not UTF-8".to_string())?;
    if !omml.trim_start().starts_with("<m:oMath")
        || !omml.contains("http://schemas.openxmlformats.org/officeDocument/2006/math")
        || omml.contains("<!DOCTYPE")
        || omml.contains("<!ENTITY")
    {
        return Err("Document formula OMML payload is not a safe Office Math fragment".to_string());
    }
    Ok(())
}

fn decode_document_native_docx(value: &str) -> Result<Vec<u8>, String> {
    let bytes = URL_SAFE_NO_PAD
        .decode(value)
        .map_err(|_| "Document formula native DOCX is not valid Base64URL".to_string())?;
    if bytes.len() < 128 || bytes.len() > MAX_OMML_BYTES * 8 || !bytes.starts_with(b"PK\x03\x04") {
        return Err("Document formula native DOCX payload is invalid or too large".to_string());
    }
    Ok(bytes)
}

fn decode_document_image_fallback_png(value: Option<&str>) -> Result<Vec<u8>, String> {
    let value = value.ok_or_else(|| {
        "Image document formula is missing its PNG compatibility preview".to_string()
    })?;
    let bytes = decode_png(value)?;
    if bytes.len() < 24 {
        return Err("Image document formula PNG compatibility preview is truncated".to_string());
    }
    let width = u32::from_be_bytes(bytes[16..20].try_into().map_err(|_| {
        "Image document formula PNG compatibility preview has an invalid IHDR".to_string()
    })?);
    let height = u32::from_be_bytes(bytes[20..24].try_into().map_err(|_| {
        "Image document formula PNG compatibility preview has an invalid IHDR".to_string()
    })?);
    if width <= 1 || height <= 1 {
        return Err(
            "Image document formula PNG compatibility preview must contain rendered formula pixels"
                .to_string(),
        );
    }
    Ok(bytes)
}

fn calculate_document_image_geometry(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
    formula_letter_font: Option<&str>,
) -> Result<WordGeometry, String> {
    if !width.is_finite()
        || !height.is_finite()
        || !baseline.is_finite()
        || width <= 0.0
        || height <= 0.0
        || baseline < 0.0
        || baseline > height
    {
        return Err("Document formula SVG geometry is invalid".to_string());
    }
    calculate_word_svg_geometry_for_font(
        width,
        height,
        baseline,
        font_size_pt,
        formula_letter_font,
    )
}

fn resolve_document_paragraph_transfer(
    paragraph_id: &Option<String>,
    paragraph_style: &Option<String>,
    paragraph_alignment: &Option<String>,
    list_kind: &Option<String>,
    list_level: Option<u32>,
    paragraph_start: bool,
    paragraph_end: bool,
) -> Result<Option<DocumentParagraphTransfer>, String> {
    let Some(id) = paragraph_id.as_deref() else {
        if paragraph_style.is_some()
            || paragraph_alignment.is_some()
            || list_kind.is_some()
            || list_level.is_some()
            || paragraph_start
            || paragraph_end
        {
            return Err("Document paragraph metadata is missing its paragraph id".to_string());
        }
        return Ok(None);
    };
    validate_uuid(id, "Document paragraph id")?;
    let style = paragraph_style.as_deref().unwrap_or("normal");
    if !matches!(
        style,
        "normal" | "heading1" | "heading2" | "heading3" | "heading4" | "quote" | "code"
    ) {
        return Err("Document paragraph style is invalid".to_string());
    }
    let alignment = paragraph_alignment.as_deref().unwrap_or("left");
    if !matches!(alignment, "left" | "center" | "right" | "justify") {
        return Err("Document paragraph alignment is invalid".to_string());
    }
    let resolved_list_kind = list_kind.as_deref().unwrap_or("none");
    if !matches!(resolved_list_kind, "none" | "bullet" | "number") {
        return Err("Document paragraph list kind is invalid".to_string());
    }
    let resolved_list_level = list_level.unwrap_or(0);
    if (resolved_list_kind == "none" && resolved_list_level != 0)
        || (resolved_list_kind != "none" && !(1..=9).contains(&resolved_list_level))
    {
        return Err("Document paragraph list level is invalid".to_string());
    }
    Ok(Some(DocumentParagraphTransfer {
        id: id.to_string(),
        style: style.to_string(),
        alignment: alignment.to_string(),
        list_kind: resolved_list_kind.to_string(),
        list_level: resolved_list_level,
        start: paragraph_start,
        end: paragraph_end,
    }))
}

fn append_document_paragraph_entries(
    entries: &mut Vec<(String, String)>,
    prefix: &str,
    paragraph: Option<&DocumentParagraphTransfer>,
) {
    let value = |field: &str| format!("{prefix}{field}");
    if let Some(paragraph) = paragraph {
        entries.push((value("paragraphId"), paragraph.id.clone()));
        entries.push((value("paragraphStyle"), paragraph.style.clone()));
        entries.push((value("paragraphAlignment"), paragraph.alignment.clone()));
        entries.push((value("listKind"), paragraph.list_kind.clone()));
        entries.push((value("listLevel"), paragraph.list_level.to_string()));
        entries.push((
            value("paragraphStart"),
            if paragraph.start { "1" } else { "0" }.to_string(),
        ));
        entries.push((
            value("paragraphEnd"),
            if paragraph.end { "1" } else { "0" }.to_string(),
        ));
    }
}

fn validate_latex_redraw_range(
    source_utf16: &[u16],
    previous_end: &mut usize,
    start: usize,
    end: usize,
    original: &str,
) -> Result<(), String> {
    let original_utf16 = original.encode_utf16().collect::<Vec<_>>();
    if start < *previous_end
        || end <= start
        || end > source_utf16.len()
        || source_utf16[start..end] != original_utf16
    {
        return Err(
            "LaTeX redraw formula range does not match the Word source snapshot".to_string(),
        );
    }
    *previous_end = end;
    Ok(())
}

fn validate_formula_restore_range(
    previous_end: &mut usize,
    start: usize,
    end: usize,
    original: &str,
) -> Result<(), String> {
    let span = end.saturating_sub(start);
    if start < *previous_end
        || end <= start
        || end > 50_000_000
        || span > 5_000_000
        || original.is_empty()
        || original.len() > MAX_LATEX_REDRAW_SOURCE_BYTES as usize
    {
        return Err("Word formula restore range is invalid or overlapping".to_string());
    }
    *previous_end = end;
    Ok(())
}

fn parse_latex_redraw_font_sizes(source: &str, expected_count: usize) -> Result<Vec<f64>, String> {
    let mut item_count = None;
    let mut values = vec![None; expected_count];
    for line in source.lines() {
        let Some((key, raw)) = line.split_once('=') else {
            continue;
        };
        if key == "itemCount" {
            item_count = raw.parse::<usize>().ok();
            continue;
        }
        let Some(index_text) = key
            .strip_prefix("item")
            .and_then(|value| value.strip_suffix("fontSizePt"))
        else {
            continue;
        };
        let index = index_text
            .parse::<usize>()
            .map_err(|_| "LaTeX redraw font-size result contains an invalid item index".to_string())?;
        if index >= expected_count || values[index].is_some() {
            return Err("LaTeX redraw font-size result contains an invalid or duplicate item".to_string());
        }
        let value = raw
            .parse::<f64>()
            .map_err(|_| "LaTeX redraw font-size result contains an invalid number".to_string())?;
        if !value.is_finite() || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&value) {
            return Err("LaTeX redraw font-size result is outside the supported range".to_string());
        }
        values[index] = Some(value);
    }
    if item_count != Some(expected_count) {
        return Err("LaTeX redraw font-size result count does not match the request".to_string());
    }
    values
        .into_iter()
        .map(|value| value.ok_or_else(|| "LaTeX redraw font-size result is incomplete".to_string()))
        .collect()
}

fn resolve_latex_redraw_font_sizes_blocking(
    session_id: String,
    input: MacOfflineLatexRedrawFontQueryInput,
) -> Result<Vec<f64>, String> {
    let preflight_started = Instant::now();
    validate_uuid(&session_id, "Session id")?;
    queue_editor_performance(
        OfficeHost::Word,
        &session_id,
        "latex-redraw-font-backend-start",
        0.0,
        None,
        json!({ "itemCount": input.ranges.len() }),
    );
    if input.ranges.is_empty() || input.ranges.len() > 1000 {
        return Err("LaTeX redraw font preflight must contain 1 to 1000 ranges".to_string());
    }
    let request = read_request(&session_id)?;
    let public_request = document_import_request_data(&request)?;
    if public_request.operation != "latexRedraw" {
        return Err("Only a Word LaTeX redraw session can resolve source font sizes".to_string());
    }
    let source = public_request
        .source
        .as_deref()
        .ok_or_else(|| "LaTeX redraw source snapshot is missing".to_string())?;
    let source_utf16 = source.encode_utf16().collect::<Vec<_>>();
    let mut previous_end = 0usize;
    for range in &input.ranges {
        validate_latex_redraw_range(
            &source_utf16,
            &mut previous_end,
            range.source_start,
            range.source_end,
            &range.source_text,
        )?;
        if !matches!(range.display_mode.as_str(), "inline" | "block") {
            return Err("LaTeX redraw font preflight contains an invalid display mode".to_string());
        }
    }

    let manifest_path = latex_redraw_preflight_manifest_path(&session_id)?;
    let result_path = latex_redraw_font_sizes_path(&session_id)?;
    let mut entries = vec![
        ("protocolVersion".to_string(), OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId".to_string(), session_id.clone()),
        ("operation".to_string(), "latexRedraw".to_string()),
        (
            "outputKind".to_string(),
            public_request
                .output_kind
                .clone()
                .ok_or_else(|| "LaTeX redraw output kind is missing".to_string())?,
        ),
        (
            "sourceDocumentId".to_string(),
            public_request.source_document_id.clone(),
        ),
        ("bookmarkName".to_string(), public_request.bookmark_name.clone()),
        ("itemCount".to_string(), input.ranges.len().to_string()),
    ];
    for (index, range) in input.ranges.iter().enumerate() {
        let prefix = format!("item{index}");
        entries.push((format!("{prefix}sourceStart"), range.source_start.to_string()));
        entries.push((format!("{prefix}sourceEnd"), range.source_end.to_string()));
        entries.push((
            format!("{prefix}sourceTextBase64"),
            URL_SAFE_NO_PAD.encode(range.source_text.as_bytes()),
        ));
        entries.push((
            format!("{prefix}displayMode"),
            range.display_mode.clone(),
        ));
    }
    let manifest = dynamic_dispatch_text(&entries)?;
    if manifest.len() > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES {
        return Err("LaTeX redraw font preflight exceeds the transfer limit".to_string());
    }
    atomic_write(&manifest_path, manifest.as_bytes(), 0o600)?;
    let _ = fs::remove_file(&result_path);

    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session_id.clone()),
        ("action", "latexRedrawPreflight".to_string()),
        ("host", "word".to_string()),
        (
            "sourceDocumentId",
            public_request.source_document_id.clone(),
        ),
        ("bookmarkName", public_request.bookmark_name.clone()),
        (
            "documentImportPath",
            manifest_path.to_string_lossy().to_string(),
        ),
        (
            "fontSizeResultPath",
            result_path.to_string_lossy().to_string(),
        ),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Word, &session_id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    let result = fs::read_to_string(&result_path)
        .map_err(|error| format!("Unable to read Word LaTeX redraw font sizes: {error}"))?;
    let values = parse_latex_redraw_font_sizes(&result, input.ranges.len())?;
    queue_editor_performance(
        OfficeHost::Word,
        &session_id,
        "latex-redraw-font-backend-complete",
        preflight_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({ "itemCount": values.len() }),
    );
    let _ = fs::remove_file(manifest_path);
    let _ = fs::remove_file(result_path);
    Ok(values)
}

fn commit_document_import_blocking(
    state: OfficeCompanionState,
    session_id: String,
    input: MacOfflineDocumentImportCommitInput,
) -> Result<(), String> {
    let commit_started = Instant::now();
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    let public_request = document_import_request_data(&request)?;
    let is_redraw = public_request.operation == "latexRedraw";
    let is_formula_restore = public_request.operation == "formulaRestore";
    let is_range_replace = is_redraw || is_formula_restore;
    if is_redraw {
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-backend-start",
            0.0,
            None,
            json!({ "itemCount": input.items.len(), "outputKind": input.output_kind }),
        );
    }
    let valid_output = if is_formula_restore {
        matches!(input.output_kind.as_str(), "latex" | "image")
    } else {
        matches!(input.output_kind.as_str(), "omml" | "image")
    };
    if !valid_output {
        return Err("Document formula output kind is invalid for this operation".to_string());
    }
    if is_range_replace
        && public_request.output_kind.as_deref() != Some(input.output_kind.as_str())
    {
        return Err("Word range replacement output kind changed after the request".to_string());
    }
    let maximum_items = if is_range_replace { 1000 } else { 2048 };
    if input.items.is_empty() || input.items.len() > maximum_items {
        return Err(format!(
            "Document operation must contain 1 to {maximum_items} blocks"
        ));
    }
    let range_source_utf16 = public_request
        .source
        .as_deref()
        .map(|source| source.encode_utf16().collect::<Vec<_>>());
    if is_range_replace {
        let source = range_source_utf16
            .as_deref()
            .ok_or_else(|| "Word range replacement source snapshot is missing".to_string())?;
        let mut preflight_end = 0usize;
        for item in &input.items {
            let (source_start, source_end, source_text, has_paragraph_metadata, item_kind) =
                match item {
                    MacOfflineDocumentImportCommitItem::Formula {
                        source_start,
                        source_end,
                        source_text,
                        paragraph_id,
                        paragraph_style,
                        paragraph_alignment,
                        list_kind,
                        list_level,
                        paragraph_start,
                        paragraph_end,
                        ..
                    } => (
                        *source_start,
                        *source_end,
                        source_text.as_deref(),
                        paragraph_id.is_some()
                            || paragraph_style.is_some()
                            || paragraph_alignment.is_some()
                            || list_kind.is_some()
                            || list_level.is_some()
                            || *paragraph_start
                            || *paragraph_end,
                        "formula",
                    ),
                    MacOfflineDocumentImportCommitItem::Text {
                        source_start,
                        source_end,
                        source_text,
                        paragraph_id,
                        paragraph_style,
                        paragraph_alignment,
                        list_kind,
                        list_level,
                        paragraph_start,
                        paragraph_end,
                        ..
                    } => (
                        *source_start,
                        *source_end,
                        source_text.as_deref(),
                        paragraph_id.is_some()
                            || paragraph_style.is_some()
                            || paragraph_alignment.is_some()
                            || list_kind.is_some()
                            || list_level.is_some()
                            || *paragraph_start
                            || *paragraph_end,
                        "text",
                    ),
                };
            if has_paragraph_metadata {
                return Err("Word range replacement items cannot carry paragraph metadata".to_string());
            }
            if (is_redraw && item_kind != "formula")
                || (is_formula_restore
                    && ((input.output_kind == "latex" && item_kind != "text")
                        || (input.output_kind == "image" && item_kind != "formula")))
            {
                return Err("Word range replacement contains an incompatible item kind".to_string());
            }
            let start = source_start
                .ok_or_else(|| "Word replacement item is missing sourceStart".to_string())?;
            let end = source_end
                .ok_or_else(|| "Word replacement item is missing sourceEnd".to_string())?;
            let original = source_text
                .ok_or_else(|| "Word replacement item is missing sourceText".to_string())?;
            if is_formula_restore {
                validate_formula_restore_range(
                    &mut preflight_end,
                    start,
                    end,
                    original,
                )?;
            } else {
                validate_latex_redraw_range(
                    source,
                    &mut preflight_end,
                    start,
                    end,
                    original,
                )?;
            }
        }
    }
    let mut previous_redraw_end = 0usize;

    let mut entries = vec![
        (
            "protocolVersion".to_string(),
            OFFLINE_PROTOCOL_VERSION.to_string(),
        ),
        ("sessionId".to_string(), session_id.clone()),
        ("operation".to_string(), public_request.operation.clone()),
        ("outputKind".to_string(), input.output_kind.clone()),
        (
            "sourceDocumentId".to_string(),
            public_request.source_document_id.clone(),
        ),
        (
            "bookmarkName".to_string(),
            public_request.bookmark_name.clone(),
        ),
        ("itemCount".to_string(), input.items.len().to_string()),
    ];
    let redraw_vector_batch_path = if is_range_replace && input.output_kind == "image" {
        Some(latex_redraw_vector_batch_path(&session_id)?)
    } else {
        None
    };
    let native_batch_document_path = if input.output_kind == "omml" {
        Some(document_native_batch_path(&session_id)?)
    } else {
        None
    };
    if let Some(batch_path) = native_batch_document_path.as_ref() {
        entries.push((
            "nativeBatchDocumentPath".to_string(),
            batch_path.to_string_lossy().to_string(),
        ));
    }
    let mut redraw_vector_entries = Vec::new();
    let mut native_batch_entries = Vec::<String>::new();
    let mut metadata_to_cache = Vec::new();
    let mut formula_count = 0usize;
    let mut text_bytes = 0usize;
    let mut active_paragraph_id: Option<String> = None;

    for (index, item) in input.items.iter().enumerate() {
        let prefix = format!("item{index}");
        match item {
            MacOfflineDocumentImportCommitItem::Text {
                text,
                source_start,
                source_end,
                source_text,
                paragraph_id,
                paragraph_style,
                paragraph_alignment,
                list_kind,
                list_level,
                paragraph_start,
                paragraph_end,
            } => {
                if is_redraw {
                    return Err("LaTeX redraw accepts formula items only".to_string());
                }
                if !is_formula_restore
                    && (source_start.is_some() || source_end.is_some() || source_text.is_some())
                {
                    return Err("Document import text contains range replacement coordinates".to_string());
                }
                let paragraph = resolve_document_paragraph_transfer(
                    paragraph_id,
                    paragraph_style,
                    paragraph_alignment,
                    list_kind,
                    *list_level,
                    *paragraph_start,
                    *paragraph_end,
                )?;
                if let Some(paragraph) = paragraph.as_ref() {
                    if paragraph.start {
                        if active_paragraph_id.is_some() {
                            return Err(
                                "Document paragraphs overlap in the transfer stream".to_string()
                            );
                        }
                        active_paragraph_id = Some(paragraph.id.clone());
                    } else if active_paragraph_id.as_deref() != Some(paragraph.id.as_str()) {
                        return Err(
                            "Document paragraph continuation has no matching start".to_string()
                        );
                    }
                } else if active_paragraph_id.is_some() {
                    return Err(
                        "Document paragraph content is missing paragraph metadata".to_string()
                    );
                }
                text_bytes = text_bytes.saturating_add(text.len());
                if text_bytes > 4 * 1024 * 1024 {
                    return Err("Document import text exceeds the 4 MB limit".to_string());
                }
                entries.push((format!("{prefix}kind"), "text".to_string()));
                entries.push((
                    format!("{prefix}textBase64"),
                    URL_SAFE_NO_PAD.encode(text.as_bytes()),
                ));
                if is_formula_restore {
                    entries.push((
                        format!("{prefix}sourceStart"),
                        source_start.expect("validated restore sourceStart").to_string(),
                    ));
                    entries.push((
                        format!("{prefix}sourceEnd"),
                        source_end.expect("validated restore sourceEnd").to_string(),
                    ));
                    entries.push((
                        format!("{prefix}sourceTextBase64"),
                        URL_SAFE_NO_PAD.encode(
                            source_text
                                .as_deref()
                                .expect("validated restore sourceText")
                                .as_bytes(),
                        ),
                    ));
                }
                append_document_paragraph_entries(&mut entries, &prefix, paragraph.as_ref());
                if paragraph.as_ref().is_some_and(|value| value.end) {
                    active_paragraph_id = None;
                }
            }
            MacOfflineDocumentImportCommitItem::Formula {
                formula_id,
                latex,
                display_mode,
                numbered,
                font_size_pt,
                metadata,
                omml_base64,
                omml_docx_base64,
                svg_base64,
                png_base64,
                width,
                height,
                baseline,
                ink_center_y_ratio,
                source_start,
                source_end,
                source_text,
                paragraph_id,
                paragraph_style,
                paragraph_alignment,
                list_kind,
                list_level,
                paragraph_start,
                paragraph_end,
            } => {
                let paragraph = resolve_document_paragraph_transfer(
                    paragraph_id,
                    paragraph_style,
                    paragraph_alignment,
                    list_kind,
                    *list_level,
                    *paragraph_start,
                    *paragraph_end,
                )?;
                if is_range_replace {
                    if paragraph.is_some() {
                        return Err(
                            "Word range replacement formulas cannot carry paragraph metadata".to_string()
                        );
                    }
                    let start = (*source_start)
                        .ok_or_else(|| "Word replacement formula is missing sourceStart".to_string())?;
                    let end = (*source_end)
                        .ok_or_else(|| "Word replacement formula is missing sourceEnd".to_string())?;
                    let original = source_text
                        .as_deref()
                        .ok_or_else(|| "Word replacement formula is missing sourceText".to_string())?;
                    let source = range_source_utf16
                        .as_ref()
                        .ok_or_else(|| "Word replacement source snapshot is missing".to_string())?;
                    if is_formula_restore {
                        validate_formula_restore_range(
                            &mut previous_redraw_end,
                            start,
                            end,
                            original,
                        )?;
                    } else {
                        validate_latex_redraw_range(
                            source,
                            &mut previous_redraw_end,
                            start,
                            end,
                            original,
                        )?;
                    }
                } else if source_start.is_some() || source_end.is_some() || source_text.is_some() {
                    return Err(
                        "Document import formula contains LaTeX redraw coordinates".to_string()
                    );
                }
                if display_mode == "block" && paragraph.is_some() {
                    return Err("Display formulas must own their Word paragraph".to_string());
                }
                if let Some(paragraph) = paragraph.as_ref() {
                    if paragraph.start {
                        if active_paragraph_id.is_some() {
                            return Err(
                                "Document paragraphs overlap in the transfer stream".to_string()
                            );
                        }
                        active_paragraph_id = Some(paragraph.id.clone());
                    } else if active_paragraph_id.as_deref() != Some(paragraph.id.as_str()) {
                        return Err(
                            "Document paragraph continuation has no matching start".to_string()
                        );
                    }
                } else if active_paragraph_id.is_some() {
                    return Err(
                        "Document paragraph content is missing paragraph metadata".to_string()
                    );
                }
                formula_count += 1;
                if formula_count > 1000 {
                    return Err("Document operations support at most 1000 formulas".to_string());
                }
                validate_uuid(formula_id, "Document formula id")?;
                if latex.trim().is_empty() || latex.len() > 1_000_000 || latex.contains('\0') {
                    return Err(
                        "A document formula contains invalid or excessive LaTeX".to_string()
                    );
                }
                if !matches!(display_mode.as_str(), "inline" | "block") {
                    return Err("Document formula display mode is invalid".to_string());
                }
                if *numbered && display_mode != "block" {
                    return Err("Only document display formulas can be numbered".to_string());
                }
                if !font_size_pt.is_finite()
                    || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(font_size_pt)
                {
                    return Err(
                        "Document formula font size is outside the supported range".to_string()
                    );
                }
                validate_document_omml_payload(omml_base64)?;
                let native_docx = decode_document_native_docx(omml_docx_base64)?;
                let native_document_path = native_word_document_path(formula_id)?;
                atomic_write(&native_document_path, &native_docx, 0o600)?;
                let native_batch_document_index = if native_batch_document_path.is_some() {
                    let omml_bytes = URL_SAFE_NO_PAD
                        .decode(omml_base64)
                        .map_err(|_| "Document formula OMML payload is not valid Base64URL".to_string())?;
                    let omml_fragment = String::from_utf8(omml_bytes)
                        .map_err(|_| "Document formula OMML payload is not UTF-8".to_string())?;
                    native_batch_entries.push(omml_fragment);
                    native_batch_entries.len()
                } else {
                    0
                };

                let mut resolved_metadata = metadata.clone();
                let canonical_latex = validate_document_formula_metadata_match(
                    &resolved_metadata,
                    formula_id,
                    latex,
                    display_mode,
                    *numbered,
                )?;
                resolved_metadata.latex = canonical_latex.clone();
                resolved_metadata.font_size_pt = Some(*font_size_pt);

                let mut image_path = String::new();
                let mut vector_document_path = String::new();
                let mut vector_document_index = 0usize;
                let mut fallback_image_path = String::new();
                let geometry = if input.output_kind == "image" {
                    let svg_value = svg_base64
                        .as_deref()
                        .ok_or_else(|| "Image document formula is missing SVG data".to_string())?;
                    let svg = decode_svg(svg_value)?;
                    let png = decode_document_image_fallback_png(png_base64.as_deref())?;
                    let width = width
                        .ok_or_else(|| "Image document formula width is missing".to_string())?;
                    let height = height
                        .ok_or_else(|| "Image document formula height is missing".to_string())?;
                    let baseline = baseline.ok_or_else(|| {
                        "Image document formula is missing its mathematical baseline".to_string()
                    })?;
                    let geometry = calculate_document_image_geometry(
                        width,
                        height,
                        baseline,
                        *font_size_pt,
                        resolved_metadata.formula_letter_font.as_deref(),
                    )?;
                    let svg_path = document_formula_file_path(&session_id, formula_id, "svg")?;
                    let png_path = document_formula_file_path(&session_id, formula_id, "png")?;
                    atomic_write(&svg_path, &svg, 0o600)?;
                    atomic_write(&png_path, &png, 0o600)?;
                    if let Some(batch_path) = redraw_vector_batch_path.as_ref() {
                        vector_document_index = redraw_vector_entries.len() + 1;
                        redraw_vector_entries.push(WordSvgBatchEntry {
                            svg,
                            png,
                            width_points: geometry.width,
                            height_points: geometry.height,
                        });
                        vector_document_path = batch_path.to_string_lossy().to_string();
                    } else {
                        let vector_path =
                            document_formula_file_path(&session_id, formula_id, "docx")?;
                        let package =
                            build_word_svg_docx(&svg, &png, geometry.width, geometry.height)?;
                        atomic_write(&vector_path, &package, 0o600)?;
                        vector_document_path = vector_path.to_string_lossy().to_string();
                    }
                    image_path = svg_path.to_string_lossy().to_string();
                    fallback_image_path = png_path.to_string_lossy().to_string();
                    resolved_metadata.render_width_px = Some(width);
                    resolved_metadata.render_height_px = Some(height);
                    resolved_metadata.reference_width_pt = Some(geometry.reference_width_pt);
                    resolved_metadata.reference_height_pt = Some(geometry.reference_height_pt);
                    resolved_metadata.reference_baseline_pt = Some(geometry.reference_baseline_pt);
                    resolved_metadata.image_ink_center_y_ratio = ink_center_y_ratio
                        .filter(|value| value.is_finite() && (0.0..=1.0).contains(value));
                    geometry
                } else {
                    resolved_metadata.reference_width_pt = None;
                    resolved_metadata.reference_height_pt = None;
                    resolved_metadata.reference_baseline_pt = None;
                    resolved_metadata.image_ink_center_y_ratio = None;
                    WordGeometry {
                        width: *font_size_pt,
                        height: (*font_size_pt * 1.8).max(18.0),
                        baseline: 0,
                        font_size_pt: *font_size_pt,
                        reference_width_pt: WORD_REFERENCE_FONT_SIZE_PT,
                        reference_height_pt: WORD_REFERENCE_FONT_SIZE_PT,
                        reference_baseline_pt: 0.0,
                    }
                };
                let encoded_metadata = encode_metadata(&resolved_metadata)?;
                metadata_to_cache.push(resolved_metadata);

                entries.push((format!("{prefix}kind"), "formula".to_string()));
                entries.push((format!("{prefix}formulaId"), formula_id.clone()));
                entries.push((
                    format!("{prefix}latexBase64"),
                    URL_SAFE_NO_PAD.encode(canonical_latex.as_bytes()),
                ));
                entries.push((format!("{prefix}displayMode"), display_mode.clone()));
                entries.push((
                    format!("{prefix}numbered"),
                    if *numbered { "1" } else { "0" }.to_string(),
                ));
                entries.push((
                    format!("{prefix}fontSizePt"),
                    format!("{:.6}", font_size_pt),
                ));
                entries.push((format!("{prefix}metadata"), encoded_metadata));
                entries.push((format!("{prefix}ommlBase64"), omml_base64.clone()));
                entries.push((
                    format!("{prefix}nativeDocumentPath"),
                    native_document_path.to_string_lossy().to_string(),
                ));
                if native_batch_document_index > 0 {
                    entries.push((
                        format!("{prefix}nativeBatchDocumentIndex"),
                        native_batch_document_index.to_string(),
                    ));
                }
                entries.push((format!("{prefix}imagePath"), image_path));
                entries.push((format!("{prefix}vectorDocumentPath"), vector_document_path));
                if vector_document_index > 0 {
                    entries.push((
                        format!("{prefix}vectorDocumentIndex"),
                        vector_document_index.to_string(),
                    ));
                }
                entries.push((format!("{prefix}fallbackImagePath"), fallback_image_path));
                entries.push((
                    format!("{prefix}widthPoints"),
                    format!("{:.6}", geometry.width),
                ));
                entries.push((
                    format!("{prefix}heightPoints"),
                    format!("{:.6}", geometry.height),
                ));
                entries.push((format!("{prefix}baseline"), geometry.baseline.to_string()));
                entries.push((
                    format!("{prefix}referenceWidthPt"),
                    format!("{:.6}", geometry.reference_width_pt),
                ));
                entries.push((
                    format!("{prefix}referenceHeightPt"),
                    format!("{:.6}", geometry.reference_height_pt),
                ));
                entries.push((
                    format!("{prefix}referenceBaselinePt"),
                    format!("{:.6}", geometry.reference_baseline_pt),
                ));
                if let Some(value) = ink_center_y_ratio
                    .filter(|value| value.is_finite() && (0.0..=1.0).contains(value))
                {
                    entries.push((
                        format!("{prefix}inkCenterYRatio"),
                        format!("{value:.8}"),
                    ));
                }
                if is_range_replace {
                    entries.push((
                        format!("{prefix}sourceStart"),
                        source_start
                            .expect("validated redraw sourceStart")
                            .to_string(),
                    ));
                    entries.push((
                        format!("{prefix}sourceEnd"),
                        source_end.expect("validated redraw sourceEnd").to_string(),
                    ));
                    entries.push((
                        format!("{prefix}sourceTextBase64"),
                        URL_SAFE_NO_PAD.encode(
                            source_text
                                .as_deref()
                                .expect("validated redraw sourceText")
                                .as_bytes(),
                        ),
                    ));
                }
                append_document_paragraph_entries(&mut entries, &prefix, paragraph.as_ref());
                if paragraph.as_ref().is_some_and(|value| value.end) {
                    active_paragraph_id = None;
                }
            }
        }
    }
    if active_paragraph_id.is_some() {
        return Err("Document paragraph transfer ended before its paragraph boundary".to_string());
    }
    if formula_count == 0 && text_bytes == 0 {
        return Err("Document import contains no visible content".to_string());
    }
    if let Some(batch_path) = redraw_vector_batch_path.as_ref() {
        let package = build_word_svg_batch_docx(&redraw_vector_entries)?;
        atomic_write(batch_path, &package, 0o600)?;
    }
    if let Some(batch_path) = native_batch_document_path.as_ref() {
        let package = build_word_omml_batch_docx(&native_batch_entries)?;
        atomic_write(batch_path, &package, 0o600)?;
        entries.push((
            "nativeBatchFormulaCount".to_string(),
            native_batch_entries.len().to_string(),
        ));
    }

    let manifest = dynamic_dispatch_text(&entries)?;
    if manifest.len() > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES {
        return Err(
            "Document import expands beyond the 16 MB Word transfer limit; split the source into smaller imports"
                .to_string(),
        );
    }
    let manifest_path = document_import_manifest_path(&session_id)?;
    atomic_write(&manifest_path, manifest.as_bytes(), 0o600)?;
    let progress_path =
        session_directory(OfficeHost::Word, &session_id)?.join(DOCUMENT_IMPORT_PROGRESS_FILE);
    atomic_write(
        &progress_path,
        format!("current=0\ntotal={}\nstage=preparing\n", input.items.len()).as_bytes(),
        0o600,
    )?;
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session_id.clone()),
        ("action", "documentCommit".to_string()),
        ("host", "word".to_string()),
        (
            "sourceDocumentId",
            public_request.source_document_id.clone(),
        ),
        ("bookmarkName", public_request.bookmark_name.clone()),
        (
            "documentImportPath",
            manifest_path.to_string_lossy().to_string(),
        ),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    if is_redraw {
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-backend-prepared",
            commit_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({ "itemCount": input.items.len() }),
        );
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-word-callback-start",
            commit_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({ "itemCount": input.items.len() }),
        );
    }
    with_dispatch_pointer(OfficeHost::Word, &session_id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    if is_redraw {
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-word-callback-complete",
            commit_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({ "itemCount": input.items.len() }),
        );
    }

    for metadata in metadata_to_cache {
        let formula_id = metadata.formula_id.clone();
        let _ = state.formula_cache.put(&formula_id, metadata);
    }
    let _ = cleanup_session_files(OfficeHost::Word, &session_id);
    Ok(())
}

fn cancel_document_import_blocking(session_id: String) -> Result<(), String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    let public_request = document_import_request_data(&request)?;
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session_id.clone()),
        ("action", "documentCancel".to_string()),
        ("host", "word".to_string()),
        ("sourceDocumentId", public_request.source_document_id),
        ("bookmarkName", public_request.bookmark_name),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Word, &session_id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    let _ = cleanup_session_files(OfficeHost::Word, &session_id);
    Ok(())
}

fn complete_session(
    state: &OfficeCompanionState,
    session_id: &str,
) -> Result<OfficeFormulaSession, String> {
    state
        .session_store
        .patch(session_id, json!({ "status": "completed", "error": null }))
        .map_err(|error| error.to_string())
}

fn fail_session(state: &OfficeCompanionState, session_id: &str, error: &str) {
    let _ = state
        .session_store
        .patch(session_id, json!({ "status": "failed", "error": error }));
}

fn commit_session_blocking(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    let commit_started = Instant::now();
    validate_uuid(&session_id, "Session id")?;
    let session = state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())?;
    queue_editor_performance(
        session.host,
        &session_id,
        "apply-backend-start",
        0.0,
        None,
        json!({ "mode": session.mode, "dirty": session.dirty }),
    );
    if session.status == OfficeSessionStatus::Completed {
        let _ = cleanup_session_files(session.host, &session_id);
        return Ok(session);
    }
    if session.status != OfficeSessionStatus::Committing {
        return Err("Offline Office Session is not ready to commit".to_string());
    }
    let request = read_request(&session_id)?;
    let source_is_native_word_equation = session.host == OfficeHost::Word
        && request
            .source_object_id
            .as_deref()
            .is_some_and(|value| value.starts_with("VT_F_"));
    let word_format_conversion_requested = session.host == OfficeHost::Word
        && source_is_native_word_equation != request.native_equation;
    if session.mode == OfficeSessionMode::Edit
        && !session.dirty
        && !word_format_conversion_requested
    {
        let completed = complete_session(&state, &session_id)?;
        queue_editor_performance(
            session.host,
            &session_id,
            "apply-noop-complete",
            commit_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({}),
        );
        let _ = cleanup_session_files(session.host, &session_id);
        return Ok(completed);
    }
    let mut metadata = metadata_from_session(&session);
    let result = match session.host {
        OfficeHost::Word => {
            let geometry = calculate_word_geometry(&request, &session)?;
            metadata.latex = canonical_document_formula_latex(&metadata)?;
            metadata.font_size_pt = Some(geometry.font_size_pt);
            metadata.reference_width_pt = Some(geometry.reference_width_pt);
            metadata.reference_height_pt = Some(geometry.reference_height_pt);
            metadata.reference_baseline_pt = Some(geometry.reference_baseline_pt);
            let encoded = encode_metadata(&metadata)?;
            commit_word(
                state.app.as_ref(),
                &request,
                &session,
                &encoded,
                &metadata.latex,
                geometry,
            )
        }
        OfficeHost::Powerpoint => {
            let powerpoint = request
                .power_point
                .as_ref()
                .ok_or_else(|| "PowerPoint request geometry is missing".to_string())?;
            let geometry = calculate_powerpoint_geometry(powerpoint, &session)?;
            metadata.font_size_pt = Some(geometry.font_size_pt);
            metadata.reference_width_pt = Some(geometry.reference_width_pt);
            metadata.reference_height_pt = Some(geometry.reference_height_pt);
            metadata.reference_baseline_pt = None;
            let encoded = encode_metadata(&metadata)?;
            commit_powerpoint(&request, &session, &encoded, geometry)
        }
    };
    if let Err(error) = result {
        fail_session(&state, &session_id, &error);
        return Err(error);
    }
    queue_editor_performance(
        session.host,
        &session_id,
        "apply-office-callback-complete",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    let completed = complete_session(&state, &session_id)?;
    queue_editor_performance(
        session.host,
        &session_id,
        "apply-backend-complete",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );

    // The Office document already contains the durable edit metadata. Cache
    // refresh and ephemeral Session cleanup must not extend the user's Apply
    // wait after the host callback has succeeded.
    let maintenance_state = state.clone();
    let maintenance_session_id = session_id.clone();
    let maintenance_host = session.host;
    let formula_id = session.formula_id.clone();
    std::thread::spawn(move || {
        if let Err(error) = maintenance_state.formula_cache.put(&formula_id, metadata) {
            eprintln!("Unable to refresh the VisualTeX formula cache after Apply: {error}");
        }
        let _ = cleanup_session_files(maintenance_host, &maintenance_session_id);
    });
    Ok(completed)
}

fn cancel_session_blocking(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    let host = host_from_request_name(&request.host)?;
    if let Err(error) = cancel_host(&request) {
        fail_session(&state, &session_id, &error);
        return Err(error);
    }
    let cancelled = state
        .session_store
        .patch(
            &session_id,
            json!({
                "status": "cancelled",
                "explicitCancel": true,
                "error": null
            }),
        )
        .map_err(|error| error.to_string())?;
    let _ = cleanup_session_files(host, &session_id);
    Ok(cancelled)
}

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
    clear_word_document_import_status();
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
    let path =
        session_directory(OfficeHost::Word, &session_id)?.join(DOCUMENT_IMPORT_PROGRESS_FILE);
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
    tokio::task::spawn_blocking(move || commit_document_import_blocking(state, session_id, input))
        .await
        .map_err(|error| format!("Offline document import task failed: {error}"))?
}

#[tauri::command]
pub async fn cancel_macos_offline_document_import(session_id: String) -> Result<(), String> {
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

#[cfg(target_os = "macos")]
pub(crate) fn refresh_health_signal(host: &str) -> bool {
    let (process_name, script) = match host {
        "word" => (
            "Microsoft Word",
            r#"tell application "Microsoft Word" to run VB macro macro name "AutoExec""#,
        ),
        "powerpoint" => (
            "Microsoft PowerPoint",
            r#"tell application "Microsoft PowerPoint" to run VB macro macro name "Auto_Open" list of parameters {}"#,
        ),
        _ => return false,
    };
    let running = Command::new("/usr/bin/pgrep")
        .args(["-x", process_name])
        .output()
        .is_ok_and(|output| output.status.success());
    if !running {
        return false;
    }
    Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(script)
        .output()
        .is_ok_and(|output| output.status.success())
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn refresh_health_signal(_host: &str) -> bool {
    false
}

pub(crate) fn health_path(host: &str) -> Result<PathBuf, String> {
    Ok(runtime_root(host_from_request_name(host)?)?
        .join("OfficePluginStatus")
        .join(format!("{host}.json")))
}

fn read_health(host: &str) -> Result<MacOfflinePluginHealth, String> {
    let path = health_path(host)?;
    let fallback = || MacOfflinePluginHealth {
        loaded: false,
        plugin_version: None,
        source_revision: None,
        host: host.to_string(),
        timestamp: None,
        status_path: path.display().to_string(),
    };
    let bytes = match fs::read(&path) {
        Ok(bytes) => bytes,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(fallback()),
        Err(error) => return Err(format!("Unable to read {} health: {error}", host)),
    };
    let value: Value = serde_json::from_slice(&bytes)
        .map_err(|error| format!("{host} health file contains invalid JSON: {error}"))?;
    let plugin_version = value
        .get("pluginVersion")
        .and_then(Value::as_str)
        .map(str::to_string);
    let source_revision = value
        .get("sourceRevision")
        .and_then(Value::as_str)
        .map(str::to_string);
    let reported_host = value
        .get("host")
        .and_then(Value::as_str)
        .unwrap_or(host)
        .to_string();
    let timestamp = value
        .get("timestamp")
        .and_then(Value::as_str)
        .filter(|value| {
            !value.is_empty() && value.len() <= 64 && !value.chars().any(char::is_control)
        })
        .map(str::to_string);
    let loaded = value
        .get("loaded")
        .and_then(Value::as_bool)
        .unwrap_or(false)
        && plugin_version.as_deref() == Some(env!("CARGO_PKG_VERSION"))
        && reported_host == host
        && timestamp.is_some();
    Ok(MacOfflinePluginHealth {
        loaded,
        plugin_version,
        source_revision,
        host: reported_host,
        timestamp,
        status_path: path.display().to_string(),
    })
}

#[tauri::command]
pub fn get_macos_offline_plugin_health() -> Result<Vec<MacOfflinePluginHealth>, String> {
    let _ = refresh_health_signal("word");
    let _ = refresh_health_signal("powerpoint");
    Ok(vec![read_health("word")?, read_health("powerpoint")?])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strips_only_visualtex_numbered_equation_array_wrapper() {
        let numbered = concat!(
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
            "<m:eqArr><m:eqArrPr/><m:e>",
            "<m:r><m:t>x</m:t></m:r>",
            "<m:r><m:t>#</m:t></m:r>",
            "<m:d><m:e><m:r><m:t> REF VT_N_12345678 </m:t></m:r></m:e></m:d>",
            "</m:e></m:eqArr></m:oMath>"
        );
        let stripped = strip_visualtex_numbered_equation_array(numbered)
            .expect("VisualTeX numbered Equation Array should be stripped");
        assert!(stripped.contains("<m:r><m:t>x</m:t></m:r>"));
        assert!(!stripped.contains("<m:eqArr"));
        assert!(!stripped.contains("<m:t>#</m:t>"));
        assert!(!stripped.contains("VT_N_"));
        assert!(stripped.starts_with("<m:oMath "));
        assert!(stripped.ends_with("</m:oMath>"));

        let ordinary = concat!(
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">",
            "<m:eqArr><m:eqArrPr/><m:e><m:r><m:t>x=y</m:t></m:r></m:e></m:eqArr>",
            "</m:oMath>"
        );
        assert!(strip_visualtex_numbered_equation_array(ordinary).is_none());

        let unrelated_hash = concat!(
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">",
            "<m:eqArr><m:eqArrPr/><m:e>",
            "<m:r><m:t>x</m:t></m:r><m:r><m:t>#</m:t></m:r>",
            "<m:r><m:t>ordinary text</m:t></m:r>",
            "</m:e></m:eqArr></m:oMath>"
        );
        assert!(strip_visualtex_numbered_equation_array(unrelated_hash).is_none());
    }

    #[cfg(target_os = "macos")]
    #[test]
    fn numbered_equation_array_omml_to_latex_excludes_visualtex_number_field() {
        let numbered = concat!(
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
            "<m:eqArr><m:eqArrPr><m:maxDist m:val=\"1\"/></m:eqArrPr><m:e>",
            "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e>",
            "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>",
            "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>#</m:t></m:r>",
            "<m:d><m:e><m:r><w:fldChar w:fldCharType=\"begin\"/></m:r>",
            "<m:r><m:t xml:space=\"preserve\"> REF VT_N_12345678 </m:t></m:r>",
            "<m:r><w:fldChar w:fldCharType=\"end\"/></m:r></m:e></m:d>",
            "</m:e></m:eqArr></m:oMath>"
        );
        let math_ml = word_omml_to_mathml(numbered)
            .expect("Word should convert the stripped formula body to MathML");
        assert!(!math_ml.contains("VT_N_"));
        assert!(!math_ml.contains("REF"));
        let latex = crate::office::omml_batch::mathml_to_latex(&math_ml)
            .expect("the stripped MathML should convert to LaTeX");
        assert!(latex.contains('x'));
        assert!(latex.contains('2'));
        assert!(!latex.contains("VT_N_"));
        assert!(!latex.contains("REF"));
        assert!(!latex.contains("\\#"));
    }

    #[test]
    fn formula_restore_manifest_from_environment_is_valid() {
        let Ok(session_id) = std::env::var("VISUALTEX_FORMULA_RESTORE_SESSION") else {
            return;
        };
        let request = read_request(&session_id)
            .expect("the requested real formula restore Session must be readable");
        let public = document_import_request_data(&request)
            .expect("the requested real formula restore manifest must validate");
        assert_eq!(public.operation, "formulaRestore");
        assert!(public.restore_targets.is_some());
    }

    #[test]
    fn office_formula_editors_have_one_stable_window_label_per_host() {
        assert_eq!(
            editor_window_label(OfficeHost::Word),
            "office-native-word-editor"
        );
        assert_eq!(
            editor_window_label(OfficeHost::Powerpoint),
            "office-native-powerpoint-editor"
        );
        assert_eq!(
            editor_window_host("office-native-word-editor"),
            Some(OfficeHost::Word)
        );
        assert_eq!(
            editor_window_host("office-native-powerpoint-editor"),
            Some(OfficeHost::Powerpoint)
        );
        assert_eq!(editor_window_host("office-native-document-session"), None);
    }

    #[test]
    fn fast_open_inboxes_are_host_sandbox_only_and_uuid_named() {
        let home = Path::new("/Users/visualtex-test");
        let roots = fast_open_inbox_roots_for_home(home);
        assert_eq!(roots.len(), 2);
        assert_eq!(
            roots[0],
            (
                OfficeHost::Word,
                home.join(WORD_FAST_OPEN_INBOX_SUFFIX)
            )
        );
        assert_eq!(
            roots[1],
            (
                OfficeHost::Powerpoint,
                home.join(POWERPOINT_FAST_OPEN_INBOX_SUFFIX)
            )
        );
        assert!(roots
            .iter()
            .all(|(_, root)| root.to_string_lossy().contains("/Library/Containers/com.microsoft.")));

        let session_id = "12345678-1234-4234-9234-123456789abc";
        assert_eq!(
            fast_open_session_id(Path::new(&format!("{session_id}.json"))).as_deref(),
            Some(session_id)
        );
        assert!(fast_open_session_id(Path::new(&format!(".{session_id}.tmp"))).is_none());
        assert!(fast_open_session_id(Path::new("not-a-uuid.json")).is_none());
    }

    #[test]
    fn fast_open_requests_have_a_short_replay_window() {
        let now = UNIX_EPOCH + Duration::from_secs(100);
        assert!(fast_open_modified_is_recent(
            now - Duration::from_secs(FAST_OPEN_MAX_AGE.as_secs() - 1),
            now,
        ));
        assert!(!fast_open_modified_is_recent(
            now - Duration::from_secs(FAST_OPEN_MAX_AGE.as_secs() + 1),
            now,
        ));
        assert!(!fast_open_modified_is_recent(
            now - Duration::from_millis(FAST_OPEN_MIN_STABLE_AGE.as_millis() as u64 / 2),
            now,
        ));
        assert!(fast_open_modified_is_recent(
            now + FAST_OPEN_FUTURE_TOLERANCE,
            now,
        ));
        assert!(!fast_open_modified_is_recent(
            now + FAST_OPEN_FUTURE_TOLERANCE + Duration::from_secs(1),
            now,
        ));
    }

    #[test]
    fn recent_office_request_suppresses_only_the_matching_reopen_race() {
        let temporary = tempfile::TempDir::new().expect("temporary request root should exist");
        let sessions = temporary.path().join("OfficeSessions");
        let session = sessions.join("12345678-1234-4234-9234-123456789abc");
        fs::create_dir_all(&session).expect("temporary Session directory should exist");
        let request = session.join(REQUEST_FILE);
        fs::write(&request, b"{}\n").expect("temporary request should be writable");
        let modified = fs::metadata(&request)
            .and_then(|metadata| metadata.modified())
            .expect("temporary request modification time should be readable");

        assert!(has_recent_office_editor_request_in_roots(
            modified + Duration::from_secs(1),
            Duration::from_secs(3),
            [sessions.clone()],
        ));
        assert!(!has_recent_office_editor_request_in_roots(
            modified + Duration::from_secs(4),
            Duration::from_secs(3),
            [sessions],
        ));
    }

    #[test]
    fn office_editor_window_size_uses_measured_screen_ratio_and_safe_bounds() {
        let current_monitor = OfficeEditorMonitorGeometry {
            screen_width: 1470.0,
            screen_height: 956.0,
            maximum_inner_width: 1470.0,
            maximum_inner_height: 900.0,
        };
        let default = resolve_office_editor_window_size(None, Some(current_monitor));
        assert!((default.width - 843.0).abs() < 0.001);
        assert!((default.height - 568.0).abs() < 0.001);

        let larger_monitor = OfficeEditorMonitorGeometry {
            screen_width: 1920.0,
            screen_height: 1080.0,
            maximum_inner_width: 1920.0,
            maximum_inner_height: 1020.0,
        };
        let scaled = resolve_office_editor_window_size(None, Some(larger_monitor));
        assert!((scaled.width / 1920.0 - DEFAULT_OFFICE_EDITOR_WIDTH_RATIO).abs() < 0.001);
        assert!((scaled.height / 1080.0 - DEFAULT_OFFICE_EDITOR_HEIGHT_RATIO).abs() < 0.001);

        let legacy = resolve_office_editor_window_size(
            Some(OfficeEditorWindowSizePreference {
                width: Some(930.0),
                height: Some(700.0),
                ..Default::default()
            }),
            Some(current_monitor),
        );
        assert!((legacy.width - 930.0).abs() < 0.001);
        assert!((legacy.height - 700.0).abs() < 0.001);

        let minimum = normalize_office_editor_window_size(1.0, 2.0, Some(current_monitor));
        assert_eq!(minimum.width, MIN_OFFICE_EDITOR_WIDTH);
        assert_eq!(minimum.height, MIN_OFFICE_EDITOR_HEIGHT);

        let maximum = normalize_office_editor_window_size(
            f64::INFINITY,
            99_999.0,
            Some(current_monitor),
        );
        assert_eq!(maximum.width, DEFAULT_OFFICE_EDITOR_FALLBACK_WIDTH);
        assert_eq!(maximum.height, current_monitor.maximum_inner_height);
    }

    #[test]
    fn editor_ready_marker_exposes_machine_readable_frontend_stages() {
        let marker = MacOfflineOfficeEditorReadyMarker {
            schema: "visualtex-office-editor-ready-v1",
            session_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            host: OfficeHost::Word,
            generation: 4,
            epoch_ms: 900,
            url_received_epoch_ms: 100,
            frontend_epoch_ms: 850,
            hydrate_ms: 120.5,
            editor_mounted_ms: 150.25,
            content_ready_ms: 166.75,
            show_focus_ms: 171.0,
            window_focused: true,
            window_visible: true,
            app_active: true,
            window_can_become_key: true,
            window_is_main: true,
        };
        let value = serde_json::to_value(marker).expect("ready marker should serialize");
        assert_eq!(value["sessionId"], "12345678-1234-4234-9234-123456789abc");
        assert_eq!(value["host"], "word");
        assert_eq!(value["hydrateMs"], 120.5);
        assert_eq!(value["contentReadyMs"], 166.75);
        assert_eq!(value["showFocusMs"], 171.0);
    }

    #[test]
    fn runtime_roots_use_each_office_hosts_application_scripts_directory() {
        let word = runtime_root(OfficeHost::Word).expect("Word runtime root should resolve");
        let powerpoint =
            runtime_root(OfficeHost::Powerpoint).expect("PowerPoint runtime root should resolve");
        assert!(word.ends_with(WORD_RUNTIME_SUFFIX));
        assert!(powerpoint.ends_with(POWERPOINT_RUNTIME_SUFFIX));
        assert_ne!(word, powerpoint);
        assert!(!word.to_string_lossy().contains("UBF8T346G9.Office"));
        assert!(!powerpoint.to_string_lossy().contains("UBF8T346G9.Office"));
        assert!(!word.starts_with("/private/tmp"));
        assert!(!powerpoint.starts_with("/private/tmp"));
    }

    #[test]
    fn word_image_geometry_scales_from_the_14_point_reference() {
        let small = scale_word_reference_geometry(140.0, 28.0, -4.0, 10.5)
            .expect("10.5 pt geometry should scale");
        let large = scale_word_reference_geometry(140.0, 28.0, -4.0, 18.0)
            .expect("18 pt geometry should scale");

        assert!((small.width - 105.0).abs() < 0.001);
        assert!((small.height - 21.0).abs() < 0.001);
        assert_eq!(small.baseline, -3);
        assert!((large.width - 180.0).abs() < 0.001);
        assert!((large.height - 36.0).abs() < 0.001);
        assert_eq!(large.baseline, -5);
        assert!((large.width / small.width - 18.0 / 10.5).abs() < 0.001);
    }

    #[test]
    fn word_inline_baseline_rounds_only_after_font_size_scaling() {
        // Real MathJax geometry from the reported L^2 / L_z regression at the
        // canonical 14 pt export size. Their fractional descents are about
        // 0.825 pt and 3.25512 pt respectively.
        let superscript = calculate_word_svg_geometry(
            22.86186666666666,
            17.56613333333333,
            16.56613333333333,
            11.0,
        )
        .expect("L^2 geometry should resolve");
        let subscript = calculate_word_svg_geometry(
            22.39893333333333,
            17.69493333333333,
            13.749333333333333,
            11.0,
        )
        .expect("L_z geometry should resolve");

        assert!((superscript.reference_baseline_pt + 0.825).abs() < 0.001);
        assert!((subscript.reference_baseline_pt + 3.25512).abs() < 0.001);
        assert_eq!(superscript.baseline, -1);
        assert_eq!(subscript.baseline, -3);
        assert_eq!(subscript.baseline - superscript.baseline, -2);
    }

    #[test]
    fn word_times_image_geometry_matches_native_visual_calibration() {
        // Real 11 pt Times New Roman a^2+b^2 geometry from the reported Word
        // document. The old uniform 1.1 calibration produced about 38.245 x
        // 12.379 pt; Word copy-as-picture measured its ink at 36 x 12 px while
        // the same OMML was 35 x 11 px. Times therefore keeps only the small
        // horizontal TeX-layout compensation and uses natural vertical scale.
        // Its -1.49 pt formula-specific descent floors to -2 pt so Word does
        // not render the image baseline above the adjacent native OMath.
        let times = calculate_word_svg_geometry_for_font(
            59.00053333333333,
            19.0968,
            16.566133333333333,
            11.0,
            Some("times"),
        )
        .expect("Times image geometry should resolve");
        let katex = calculate_word_svg_geometry_for_font(
            59.00053333333333,
            19.0968,
            16.566133333333333,
            11.0,
            Some("katex"),
        )
        .expect("KaTeX image geometry should retain its historical calibration");

        assert!((times.width - 37.09763891428571).abs() < 0.001);
        assert!((times.height - 11.25347142857143).abs() < 0.001);
        assert_eq!(times.baseline, -2);
        assert!((times.reference_baseline_pt + 1.898).abs() < 0.001);
        assert!((katex.width - 38.24498857142857).abs() < 0.001);
        assert!((katex.height - 12.378818571428573).abs() < 0.001);
        assert_eq!(katex.baseline, -2);
    }

    #[test]
    fn word_non_times_fonts_lower_only_shallow_superscript_descent() {
        for font in ["katex", "cambria", "stix", "palatino", "helvetica"] {
            let shallow = calculate_word_svg_geometry_for_font(
                20.83792,
                17.789648,
                16.5848,
                11.0,
                Some(font),
            )
            .expect("non-Times shallow geometry should resolve");
            let deep = calculate_word_svg_geometry_for_font(
                17.21108,
                21.760533,
                14.1756,
                11.0,
                Some(font),
            )
            .expect("non-Times fraction geometry should resolve");

            assert_eq!(shallow.baseline, -2, "font={font}");
            assert!((shallow.reference_baseline_pt + 1.91).abs() < 0.001);
            assert_eq!(deep.baseline, -5, "font={font}");
            assert!(deep.reference_baseline_pt < -6.0);
        }
    }

    #[test]
    fn latex_redraw_source_limits_and_utf16_ranges_are_strict() {
        assert!(validate_latex_redraw_source_size(0).is_err());
        assert!(validate_latex_redraw_source_size(1).is_ok());
        assert!(validate_latex_redraw_source_size(MAX_LATEX_REDRAW_SOURCE_BYTES).is_ok());
        assert!(validate_latex_redraw_source_size(MAX_LATEX_REDRAW_SOURCE_BYTES + 1).is_err());

        let source = "前😀 $x$ 后 $$y$$";
        let source_utf16 = source.encode_utf16().collect::<Vec<_>>();
        let first_start = "前😀 ".encode_utf16().count();
        let first_text = "$x$";
        let first_end = first_start + first_text.encode_utf16().count();
        let second_start = first_end + " 后 ".encode_utf16().count();
        let second_text = "$$y$$";
        let second_end = second_start + second_text.encode_utf16().count();
        let mut previous_end = 0usize;

        validate_latex_redraw_range(
            &source_utf16,
            &mut previous_end,
            first_start,
            first_end,
            first_text,
        )
        .expect("the first formula must use Word-compatible UTF-16 offsets");
        validate_latex_redraw_range(
            &source_utf16,
            &mut previous_end,
            second_start,
            second_end,
            second_text,
        )
        .expect("the second non-overlapping formula must validate");
        assert_eq!(previous_end, second_end);

        let mut overlap_end = first_end;
        assert!(validate_latex_redraw_range(
            &source_utf16,
            &mut overlap_end,
            first_start,
            first_end,
            first_text,
        )
        .is_err());
        let mut mismatch_end = 0usize;
        assert!(validate_latex_redraw_range(
            &source_utf16,
            &mut mismatch_end,
            first_start,
            first_end,
            "$z$",
        )
        .is_err());
    }

    #[test]
    fn latex_redraw_request_rejects_invalid_scope_and_output_before_file_access() {
        let session_id = "32345678-1234-4234-9234-123456789abc".to_string();
        let mut request = MacOfflineSessionRequest {
            protocol_version: OFFLINE_PROTOCOL_VERSION,
            session_id: session_id.clone(),
            host: "word".to_string(),
            mode: "create".to_string(),
            operation: Some("latexRedraw".to_string()),
            formula_id: None,
            display_mode: "inline".to_string(),
            numbered: false,
            native_equation: false,
            source_document_id: Some("Document".to_string()),
            source_object_id: None,
            encoded_metadata: None,
            pending_marker: None,
            font_size_pt: None,
            reference_width_pt: None,
            reference_height_pt: None,
            power_point: None,
            document_import: Some(MacOfflineDocumentImportRequest {
                bookmark_name: "VT_D_323456781234423492341234".to_string(),
                default_font_size_pt: 12.0,
                redraw_scope: Some("page".to_string()),
                output_kind: Some("image".to_string()),
                source_kind: None,
            }),
        };
        assert!(validate_request(&request, &session_id)
            .unwrap_err()
            .contains("scope"));

        let document_import = request.document_import.as_mut().unwrap();
        document_import.redraw_scope = Some("selection".to_string());
        document_import.output_kind = Some("ole".to_string());
        assert!(validate_request(&request, &session_id)
            .unwrap_err()
            .contains("output kind"));

        request.operation = Some("documentImport".to_string());
        let document_import = request.document_import.as_mut().unwrap();
        document_import.output_kind = Some("image".to_string());
        assert!(validate_request(&request, &session_id)
            .unwrap_err()
            .contains("formula transform fields"));
    }

    #[test]
    fn document_image_requires_a_real_png_compatibility_preview() {
        assert!(decode_document_image_fallback_png(None).is_err());

        let transparent_placeholder = BASE64_STANDARD.encode(
            BASE64_STANDARD
                .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==")
                .expect("transparent PNG fixture should decode"),
        );
        assert!(decode_document_image_fallback_png(Some(&transparent_placeholder)).is_err());

        let supplied = BASE64_STANDARD.encode(
            BASE64_STANDARD
                .decode("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGNkYGD4z8DAwMDEAAUADigBA0dwHFEAAAAASUVORK5CYII=")
                .expect("PNG fixture should decode"),
        );
        let decoded = decode_document_image_fallback_png(Some(&supplied))
            .expect("supplied PNG compatibility preview should decode");
        assert!(decoded.starts_with(b"\x89PNG\r\n\x1a\n"));
        assert_eq!(u32::from_be_bytes(decoded[16..20].try_into().unwrap()), 2);
        assert_eq!(u32::from_be_bytes(decoded[20..24].try_into().unwrap()), 2);
    }

    #[cfg(target_os = "macos")]
    #[test]
    fn word_svg_staging_docx_is_a_valid_ooxml_zip() {
        let svg = br#"<svg xmlns="http://www.w3.org/2000/svg" width="16" height="8"><path d="M0 0h16v8H0z"/></svg>"#;
        let png = BASE64_STANDARD
            .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
            .expect("PNG fixture should decode");
        let package =
            build_word_svg_docx(svg, &png, 16.0, 8.0).expect("Word SVG package should build");
        assert!(package.starts_with(b"PK\x03\x04"));
        assert!(package
            .windows(b"word/media/formula.svg".len())
            .any(|value| { value == b"word/media/formula.svg" }));
        assert!(package
            .windows(b"drawing/2016/SVG/main".len())
            .any(|value| { value == b"drawing/2016/SVG/main" }));

        let directory = tempfile::tempdir().expect("temporary directory should exist");
        let path = directory.path().join("formula-svg.docx");
        fs::write(&path, &package).expect("Word SVG package should be writable");
        let output = Command::new("/usr/bin/unzip")
            .args(["-tqq"])
            .arg(&path)
            .output()
            .expect("macOS unzip should validate the package");
        assert!(
            output.status.success(),
            "generated Word SVG DOCX is invalid: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }

    #[cfg(target_os = "macos")]
    #[test]
    #[ignore = "writes a real Word SVG probe DOCX to VISUALTEX_WORD_SVG_PROBE_PATH"]
    fn write_word_svg_probe_docx() {
        let path = std::env::var("VISUALTEX_WORD_SVG_PROBE_PATH")
            .expect("set VISUALTEX_WORD_SVG_PROBE_PATH to an absolute .docx path");
        let svg = br#"<svg xmlns="http://www.w3.org/2000/svg" width="160" height="80" viewBox="0 0 160 80"><rect width="160" height="80" fill="white"/><path d="M20 55L55 20L90 55L125 20" fill="none" stroke="black" stroke-width="8"/></svg>"#;
        let png = BASE64_STANDARD
            .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
            .expect("PNG fixture should decode");
        let package =
            build_word_svg_docx(svg, &png, 160.0, 80.0).expect("Word SVG package should build");
        fs::write(path, package).expect("Word SVG probe should be writable");
    }

    #[test]
    fn native_word_documents_are_formula_scoped_and_outlive_sessions() {
        let formula_id = "12345678-1234-4234-9234-123456789abc";
        let path = native_word_document_path(formula_id)
            .expect("native Word document path should resolve");
        let runtime = runtime_root(OfficeHost::Word).expect("Word runtime root should resolve");

        assert!(path.starts_with(&runtime));
        assert_eq!(
            path.parent()
                .and_then(|value| value.file_name())
                .and_then(|value| value.to_str()),
            Some("NativeDocuments")
        );
        assert_eq!(
            path.file_name().and_then(|value| value.to_str()),
            Some("12345678-1234-4234-9234-123456789abc.docx")
        );
        assert!(!path.to_string_lossy().contains("OfficeSessions"));
    }

    #[test]
    fn completed_session_cleanup_removes_only_known_ephemeral_files() {
        let directory =
            std::env::temp_dir().join(format!("visualtex-offline-cleanup-test-{}", Uuid::new_v4()));
        fs::create_dir_all(&directory).expect("test Session directory should be created");
        for name in [
            REQUEST_FILE,
            DISPATCH_FILE,
            RESULT_PNG_FILE,
            RESULT_SVG_FILE,
        ] {
            fs::write(directory.join(name), b"temporary").expect("temporary file should exist");
        }
        fs::write(directory.join("keep.txt"), b"keep").expect("unknown file should exist");

        cleanup_session_files_at(&directory, false).expect("known files should be cleaned");
        for name in [
            REQUEST_FILE,
            DISPATCH_FILE,
            RESULT_PNG_FILE,
            RESULT_SVG_FILE,
        ] {
            assert!(!directory.join(name).exists());
        }
        assert!(directory.join("keep.txt").is_file());
        assert!(directory.is_dir());

        fs::remove_file(directory.join("keep.txt")).unwrap();
        cleanup_session_files_at(&directory, false)
            .expect("empty Session directory should be removed");
        assert!(!directory.exists());
    }

    #[test]
    fn powerpoint_svg_decoder_accepts_internal_vector_references_only() {
        let safe = BASE64_STANDARD.encode(
            br##"<svg xmlns="http://www.w3.org/2000/svg"><defs><path id="g" d="M0 0h1v1z"/></defs><use href="#g"/></svg>"##,
        );
        let decoded = decode_svg(&safe).expect("generated SVG should be accepted");
        assert!(std::str::from_utf8(&decoded).unwrap().contains("<use"));

        let external = BASE64_STANDARD.encode(
            br#"<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.com/a.png"/></svg>"#,
        );
        assert!(decode_svg(&external).is_err());
        let scripted = BASE64_STANDARD
            .encode(br#"<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>"#);
        assert!(decode_svg(&scripted).is_err());
    }

    #[cfg(target_os = "macos")]
    #[test]
    #[ignore = "requires target/source PowerPoint Sessions and explicit environment variables"]
    fn live_powerpoint_svg_commit_uses_the_real_ppam_transaction() {
        let session_id = std::env::var("VISUALTEX_LIVE_PPT_SESSION")
            .expect("set VISUALTEX_LIVE_PPT_SESSION to an open PowerPoint create Session");
        let request = read_request(&session_id).expect("PowerPoint request should be readable");
        assert_eq!(request.host, "powerpoint");
        assert!(request.mode == "create" || request.mode == "edit");
        let formula_id = request
            .formula_id
            .clone()
            .expect("PowerPoint request should contain a formula id");
        let source_session_id = std::env::var("VISUALTEX_LIVE_PPT_EXPORT_SESSION").expect(
            "set VISUALTEX_LIVE_PPT_EXPORT_SESSION to a completed VisualTeX formula Session",
        );
        validate_uuid(&source_session_id, "Source Session id").unwrap();
        let home = std::env::var("HOME").expect("HOME should be set on macOS");
        let source_session_path = PathBuf::from(home)
            .join("Library/Application Support/com.visualtex.studio/office/sessions")
            .join(&source_session_id)
            .join("session.json");
        let mut session: OfficeFormulaSession = serde_json::from_slice(
            &fs::read(&source_session_path).expect("source VisualTeX Session should be readable"),
        )
        .expect("source VisualTeX Session should decode");
        assert_eq!(session.host, OfficeHost::Powerpoint);
        let source_export = session
            .export_result
            .as_ref()
            .expect("source VisualTeX Session must contain a real formula export");
        decode_svg(&source_export.svg_base64)
            .expect("source VisualTeX Session must contain a validated SVG export");
        session.id = session_id.clone();
        session.mode = if request.mode == "edit" {
            OfficeSessionMode::Edit
        } else {
            OfficeSessionMode::Create
        };
        session.formula_id = formula_id;
        session.source_document_id = request.source_document_id.clone();
        session.source_object_id = request.source_object_id.clone();
        session.original_metadata = request
            .encoded_metadata
            .as_deref()
            .map(decode_metadata)
            .transpose()
            .expect("target PowerPoint metadata should decode");
        session.dirty = true;
        session.status = OfficeSessionStatus::Committing;
        session.explicit_cancel = false;
        session.error = None;

        let root =
            std::env::temp_dir().join(format!("visualtex-live-powerpoint-svg-{}", Uuid::new_v4()));
        let paths = crate::office::state::OfficePaths {
            certificate: root.join("localhost-cert.pem"),
            private_key: root.join("localhost-key.pem"),
            certificate_metadata: root.join("certificate.json"),
            install: root.join("install.json"),
            sessions: root.join("sessions"),
            recovery: root.join("recovery"),
            formula_cache: root.join("formulas"),
            root: root.clone(),
        };
        let session_store = crate::office::sessions::SessionStore::new(&paths)
            .expect("live Session store should initialize");
        let formula_cache = crate::office::formula_cache::FormulaMetadataCache::new(&paths)
            .expect("live formula cache should initialize");
        let state = OfficeCompanionState::new(
            None,
            crate::OcrState::default(),
            paths,
            "a".repeat(64),
            session_store,
            formula_cache,
            true,
        );
        let metadata =
            encode_metadata(&metadata_from_session(&session)).expect("live metadata should encode");
        let powerpoint = request
            .power_point
            .as_ref()
            .expect("live PowerPoint request should contain geometry");
        let geometry = calculate_powerpoint_geometry(powerpoint, &session)
            .expect("live PowerPoint geometry should resolve");
        commit_powerpoint(&request, &session, &metadata, geometry)
            .expect("real PowerPoint PPAM SVG transaction should succeed");
        let svg_path = result_svg_path(OfficeHost::Powerpoint, &session_id)
            .expect("SVG result path should resolve");
        assert_eq!(
            fs::read_to_string(svg_path).unwrap(),
            session.export_result.as_ref().unwrap().svg
        );
        fs::remove_dir_all(root).ok();
    }

    #[test]
    fn office_url_accepts_only_the_exact_canonical_form() {
        let id = "12345678-1234-4234-9234-123456789abc";
        assert_eq!(
            parse_office_url(&format!("visualtex://office/open?session={id}")),
            Ok(id.to_string())
        );
        assert!(parse_office_url(&format!("https://office/open?session={id}")).is_err());
        assert!(parse_office_url(&format!("visualtex://office/open?session={id}&x=1")).is_err());
        assert!(parse_office_url("visualtex://office/open?session=not-a-uuid").is_err());
    }

    #[test]
    fn offline_request_json_accepts_utf8_office_identities() {
        let session_id = "32345678-1234-4234-9234-123456789abc".to_string();
        let request = MacOfflineSessionRequest {
            protocol_version: OFFLINE_PROTOCOL_VERSION,
            session_id: session_id.clone(),
            host: "word".to_string(),
            mode: "create".to_string(),
            operation: None,
            formula_id: Some("12345678-1234-4234-9234-123456789abc".to_string()),
            display_mode: "inline".to_string(),
            numbered: false,
            native_equation: false,
            source_document_id: Some("/Users/测试/公式😀.docx".to_string()),
            source_object_id: Some("书签-公式".to_string()),
            encoded_metadata: None,
            pending_marker: Some(
                "visualtex:pending:v1:32345678-1234-4234-9234-123456789abc:12345678-1234-4234-9234-123456789abc"
                    .to_string(),
            ),
            font_size_pt: Some(10.5),
            reference_width_pt: Some(60.0),
            reference_height_pt: Some(15.0),
            power_point: None,
            document_import: None,
        };
        let json = serde_json::to_vec(&request).expect("UTF-8 request should encode");
        let decoded: MacOfflineSessionRequest =
            serde_json::from_slice(&json).expect("UTF-8 request should decode");
        validate_request(&decoded, &session_id).expect("UTF-8 request should validate");
        assert_eq!(
            decoded.source_document_id.as_deref(),
            Some("/Users/测试/公式😀.docx")
        );
    }

    #[test]
    fn metadata_codec_round_trips_the_shared_schema() {
        let metadata = VisualTeXFormulaMetadata {
            schema: "visualtex-formula".to_string(),
            schema_version: 1,
            formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            title: "Formula".to_string(),
            latex: "x^2".to_string(),
            lines: vec![crate::office::sessions::MetadataLine {
                id: "22345678-1234-4234-9234-123456789abc".to_string(),
                latex: "x^2".to_string(),
            }],
            code_format: "latex".to_string(),
            display_mode: "inline".to_string(),
            numbered: false,
            render_width_px: Some(50.0),
            render_height_px: Some(20.0),
            font_size_pt: Some(10.5),
            formula_letter_font: None,
            formula_chinese_font: None,
            reference_width_pt: Some(37.5),
            reference_height_pt: Some(15.0),
            reference_baseline_pt: Some(-3.0),
            image_ink_center_y_ratio: Some(0.47826087),
            created_with_version: "1.1.0".to_string(),
            updated_with_version: "1.1.0".to_string(),
            created_at: "unix-ms:1".to_string(),
            updated_at: "unix-ms:1".to_string(),
        };
        let encoded = encode_metadata(&metadata).expect("metadata should encode");
        let decoded = decode_metadata(&encoded).expect("metadata should decode");
        assert_eq!(decoded.formula_id, metadata.formula_id);
        assert_eq!(decoded.lines[0].latex, "x^2");
        assert_eq!(decoded.font_size_pt, Some(10.5));
        assert_eq!(decoded.reference_width_pt, Some(37.5));
        assert_eq!(decoded.reference_height_pt, Some(15.0));
        assert_eq!(decoded.reference_baseline_pt, Some(-3.0));
        assert_eq!(decoded.image_ink_center_y_ratio, Some(0.47826087));
    }

    fn document_formula_metadata(
        code_format: &str,
        lines: &[&str],
        latex: &str,
        display_mode: &str,
        numbered: bool,
    ) -> VisualTeXFormulaMetadata {
        VisualTeXFormulaMetadata {
            schema: "visualtex-formula".to_string(),
            schema_version: 1,
            formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            title: "Imported formula".to_string(),
            latex: latex.to_string(),
            lines: lines
                .iter()
                .map(|latex| crate::office::sessions::MetadataLine {
                    id: Uuid::new_v4().to_string(),
                    latex: (*latex).to_string(),
                })
                .collect(),
            code_format: code_format.to_string(),
            display_mode: display_mode.to_string(),
            numbered,
            render_width_px: None,
            render_height_px: None,
            font_size_pt: Some(14.0),
            formula_letter_font: None,
            formula_chinese_font: None,
            reference_width_pt: None,
            reference_height_pt: None,
            reference_baseline_pt: None,
            image_ink_center_y_ratio: None,
            created_with_version: "1.2.5".to_string(),
            updated_with_version: "1.2.5".to_string(),
            created_at: "unix-ms:1".to_string(),
            updated_at: "unix-ms:1".to_string(),
        }
    }

    #[test]
    fn document_formula_metadata_rebuilds_canonical_multiline_environments() {
        let cases = [
            (
                "align",
                vec!["a = b + c", "d = e"],
                r"\begin{align}
a &= b + c \\
d &= e
\end{align}",
            ),
            (
                "align-star",
                vec!["x = y", "y = z"],
                r"\begin{align*}
x &= y \\
y &= z
\end{align*}",
            ),
            (
                "aligned",
                vec!["p = q", "r = s"],
                r"\[
\begin{aligned}
p &= q \\
r &= s
\end{aligned}
\]",
            ),
            (
                "gather",
                vec!["a=b", "c=d"],
                r"\begin{gather}
a=b \\
c=d
\end{gather}",
            ),
            (
                "multline-star",
                vec!["a+b+c", "=d+e"],
                r"\begin{multline*}
a+b+c \\
=d+e
\end{multline*}",
            ),
            (
                "equation-split",
                vec!["a = b", "c = d"],
                r"\begin{equation}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation}",
            ),
            (
                "equation-star-split",
                vec!["a = b", "c = d"],
                r"\begin{equation*}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation*}",
            ),
        ];

        for (code_format, lines, expected) in cases {
            let metadata = document_formula_metadata(code_format, &lines, expected, "block", false);
            assert_eq!(
                canonical_document_formula_latex(&metadata).unwrap(),
                expected,
                "{code_format} canonical source"
            );
            assert_eq!(
                validate_document_formula_metadata_match(
                    &metadata,
                    &metadata.formula_id,
                    expected,
                    "block",
                    false,
                )
                .unwrap(),
                expected,
                "{code_format} metadata match"
            );
        }
    }

    #[test]
    fn document_formula_metadata_preserves_internal_equation_newlines() {
        let body = r"u(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}c_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y,\qquad
f(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}d_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y.";
        let canonical = r"\begin{equation*}
u(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}c_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y,\qquad
f(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}d_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y.
\end{equation*}";
        let metadata =
            document_formula_metadata("equation-star", &[body], canonical, "block", false);

        assert_eq!(
            canonical_document_formula_latex(&metadata).unwrap(),
            canonical
        );
        assert_eq!(
            validate_document_formula_metadata_match(
                &metadata,
                &metadata.formula_id,
                canonical,
                "block",
                false,
            )
            .unwrap(),
            canonical,
        );

        let separate_rows = document_formula_metadata(
            "equation-star",
            &["a=b", "c=d"],
            r"\begin{equation*}
a=b
\end{equation*}

\begin{equation*}
c=d
\end{equation*}",
            "block",
            false,
        );
        assert_eq!(
            canonical_document_formula_latex(&separate_rows).unwrap(),
            separate_rows.latex,
            "separate logical rows must remain separate equation environments",
        );
    }

    #[test]
    fn document_formula_metadata_accepts_frontend_environment_spacing() {
        let frontend_source = r"\begin {equation}
E=mc^2
\end {equation}";
        let metadata =
            document_formula_metadata("equation", &["E=mc^2"], frontend_source, "block", true);

        assert_ne!(
            canonical_document_formula_latex(&metadata).unwrap(),
            frontend_source,
            "the Rust formatter intentionally normalizes environment spacing",
        );
        assert_eq!(
            validate_document_formula_metadata_match(
                &metadata,
                &metadata.formula_id,
                frontend_source,
                "block",
                true,
            )
            .unwrap(),
            frontend_source,
            "a valid frontend serialization must not be rejected because Rust formats it differently",
        );
    }

    #[test]
    fn document_formula_metadata_match_rejects_structural_drift() {
        let canonical = r"\begin{align}
a &= b \\
c &= d
\end{align}";
        let metadata =
            document_formula_metadata("align", &["a = b", "c = d"], canonical, "block", true);
        assert!(validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            canonical,
            "block",
            true,
        )
        .is_ok());
        assert!(validate_document_formula_metadata_match(
            &metadata,
            "22345678-1234-4234-9234-123456789abc",
            canonical,
            "block",
            true,
        )
        .is_err());
        assert!(validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            canonical,
            "inline",
            true,
        )
        .is_err());
        assert!(validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            canonical,
            "block",
            false,
        )
        .is_err());
        assert!(validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            r"\begin{align}
a &= b \\
c &= e
\end{align}",
            "block",
            true,
        )
        .is_err());

        let mut stale_metadata = metadata.clone();
        stale_metadata.latex = "a = b\nc = d".to_string();
        assert!(validate_document_formula_metadata_match(
            &stale_metadata,
            &stale_metadata.formula_id,
            canonical,
            "block",
            true,
        )
        .is_err());
    }

    #[test]
    fn dispatch_rejects_newlines_and_duplicate_keys() {
        assert!(dispatch_text(&[("sessionId", "a\nb".to_string())]).is_err());
        assert!(dispatch_text(&[
            ("sessionId", "a".to_string()),
            ("sessionId", "b".to_string())
        ])
        .is_err());
    }

    #[test]
    fn office_geometry_preserves_visual_point_size_and_powerpoint_center() {
        let word_geometry = calculate_word_svg_geometry(100.0, 20.0, 15.0, 14.0)
            .expect("Word image geometry should apply its visual calibration");
        assert!((word_geometry.width - 82.5).abs() < 0.001);
        assert!((word_geometry.height - 16.5).abs() < 0.001);
        assert_eq!(word_geometry.baseline, -4);
        assert!((word_geometry.reference_width_pt - 82.5).abs() < 0.001);
        assert!((word_geometry.reference_height_pt - 16.5).abs() < 0.001);

        let request = MacOfflinePowerPointRequest {
            presentation_identity: "Deck".to_string(),
            slide_index: 1,
            slide_id: 2,
            shape_name: "VisualTeX_12345678-1234-4234-9234-123456789abc".to_string(),
            left: 100.0,
            top: 200.0,
            width: 120.0,
            height: 40.0,
            rotation: 0.0,
            z_order: 2,
            font_size_pt: None,
            reference_width_pt: None,
            reference_height_pt: None,
        };
        let session = OfficeFormulaSession {
            id: "32345678-1234-4234-9234-123456789abc".to_string(),
            mode: OfficeSessionMode::Edit,
            host: OfficeHost::Powerpoint,
            operation: None,
            formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            source_document_id: None,
            source_object_id: None,
            title: "Formula".to_string(),
            lines: vec![],
            active_line_id: None,
            code_format: "latex".to_string(),
            display_mode: "block".to_string(),
            numbered: false,
            font_size_pt: None,
            formula_letter_font: "katex".to_string(),
            formula_chinese_font: "system".to_string(),
            export_width: 0.0,
            export_height: 0.0,
            export_result: Some(crate::office::sessions::OfficeExportResult {
                svg: "<svg/>".to_string(),
                svg_base64: String::new(),
                png_base64: None,
                omml_base64: None,
                omml_docx_base64: None,
                width: 300.0,
                height: 50.0,
                baseline: None,
                ink_top_ratio: None,
                ink_bottom_ratio: None,
                ink_center_y_ratio: None,
            }),
            original_metadata: Some(VisualTeXFormulaMetadata {
                schema: "visualtex-formula".to_string(),
                schema_version: 1,
                formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
                title: "Formula".to_string(),
                latex: String::new(),
                lines: vec![],
                code_format: "latex".to_string(),
                display_mode: "block".to_string(),
                numbered: false,
                render_width_px: Some(120.0),
                render_height_px: Some(40.0),
                font_size_pt: None,
                formula_letter_font: None,
                formula_chinese_font: None,
                reference_width_pt: None,
                reference_height_pt: None,
                reference_baseline_pt: None,
                image_ink_center_y_ratio: None,
                created_with_version: "1".to_string(),
                updated_with_version: "1".to_string(),
                created_at: "1".to_string(),
                updated_at: "1".to_string(),
            }),
            dirty: true,
            status: OfficeSessionStatus::Committing,
            auto_commit_on_close: true,
            explicit_cancel: false,
            error: None,
            created_at: 1,
            updated_at: 1,
            expires_at: 2,
        };
        let geometry =
            calculate_powerpoint_geometry(&request, &session).expect("geometry should scale");
        assert!((geometry.height - 50.0).abs() < 0.001);
        assert!((geometry.width - 300.0).abs() < 0.001);
        assert!((geometry.left + geometry.width / 2.0 - 160.0).abs() < 0.001);
        assert!((geometry.top + geometry.height / 2.0 - 220.0).abs() < 0.001);
        assert!((geometry.font_size_pt - 18.6666666667).abs() < 0.001);
        assert!((geometry.reference_width_pt - 225.0).abs() < 0.001);
        assert!((geometry.reference_height_pt - 37.5).abs() < 0.001);

        let mut create_request = request.clone();
        create_request.font_size_pt = Some(28.0);
        create_request.reference_width_pt = None;
        create_request.reference_height_pt = None;
        let mut create_session = session.clone();
        create_session.mode = OfficeSessionMode::Create;
        create_session.original_metadata = None;
        let create_geometry = calculate_powerpoint_geometry(&create_request, &create_session)
            .expect("declared PowerPoint point size should scale the SVG");
        assert!((create_geometry.font_size_pt - 28.0).abs() < 0.001);
        assert!((create_geometry.width - 450.0).abs() < 0.001);
        assert!((create_geometry.height - 75.0).abs() < 0.001);
        assert!((create_geometry.left + create_geometry.width / 2.0 - 160.0).abs() < 0.001);
        assert!((create_geometry.top + create_geometry.height / 2.0 - 220.0).abs() < 0.001);

        let mut edited_request = request.clone();
        edited_request.font_size_pt = Some(18.0);
        edited_request.reference_width_pt = Some(225.0);
        edited_request.reference_height_pt = Some(37.5);
        let mut edited_session = session.clone();
        edited_session.font_size_pt = Some(32.0);
        let edited_geometry = calculate_powerpoint_geometry(&edited_request, &edited_session)
            .expect("the editor-selected PowerPoint size should override the launch geometry");
        assert!((edited_geometry.font_size_pt - 32.0).abs() < 0.001);
        assert!(
            (edited_geometry.width / edited_geometry.reference_width_pt - 32.0 / 14.0).abs()
                < 0.001
        );
        assert!(
            (edited_geometry.height / edited_geometry.reference_height_pt - 32.0 / 14.0).abs()
                < 0.001
        );

        let word_request = MacOfflineSessionRequest {
            protocol_version: OFFLINE_PROTOCOL_VERSION,
            session_id: "42345678-1234-4234-9234-123456789abc".to_string(),
            host: "word".to_string(),
            mode: "edit".to_string(),
            operation: None,
            formula_id: Some("12345678-1234-4234-9234-123456789abc".to_string()),
            display_mode: "inline".to_string(),
            numbered: false,
            native_equation: false,
            source_document_id: Some("Document".to_string()),
            source_object_id: Some("VT_F_12345678-1234-4234-9234-123456789abc".to_string()),
            encoded_metadata: None,
            pending_marker: None,
            font_size_pt: Some(10.5),
            reference_width_pt: Some(60.0),
            reference_height_pt: Some(15.0),
            power_point: None,
            document_import: None,
        };
        let mut word_session = session.clone();
        word_session.host = OfficeHost::Word;
        word_session.font_size_pt = Some(24.0);
        if let Some(export) = word_session.export_result.as_mut() {
            export.baseline = Some(40.0);
        }
        if let Some(metadata) = word_session.original_metadata.as_mut() {
            metadata.font_size_pt = Some(10.5);
        }
        let edited_word_geometry = calculate_word_geometry(&word_request, &word_session)
            .expect("the editor-selected Word size should override the launch request");
        assert!((edited_word_geometry.font_size_pt - 24.0).abs() < 0.001);
        assert!(
            (edited_word_geometry.width / edited_word_geometry.reference_width_pt - 24.0 / 14.0)
                .abs()
                < 0.001
        );
        assert!(
            (edited_word_geometry.height / edited_word_geometry.reference_height_pt - 24.0 / 14.0)
                .abs()
                < 0.001
        );
    }
}
