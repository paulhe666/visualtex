//! Local HTTP channel between the WPS add-in and VisualTeX (127.0.0.1 only).
//!
//! WPS for Mac cannot be scripted from outside and its add-ins cannot start an
//! app, so the add-in talks to the running VisualTeX instead: it opens a
//! formula Session here, polls it, and writes the result the editor produced
//! into the document. Apply in the editor waits until the add-in reports the
//! document written, as a Word Apply waits for the Word macro.

use super::numbering::{
    image_current_instruction, image_increment_instruction, numbering_style, plan_equation, plan_image,
    restyle_equations, restyle_images, EquationPlan, NumberingFacts, NumberingStyle, RestyleFacts, MODE_CHAPTER,
    MODE_SECTION, MODE_SEQUENCE,
};
use crate::office::sessions::OfficeSessionStatus;
use crate::office::state::OfficeCompanionState;
use axum::extract::{DefaultBodyLimit, Path, Query, Request, State};
use axum::http::{header, HeaderValue, Method, StatusCode};
use axum::middleware::{self, Next};
use axum::response::{IntoResponse, Response};
use axum::routing::{get, post};
use axum::{Json, Router};
use serde::Deserialize;
use serde_json::{json, Value};
use std::collections::HashMap;
use std::io::Write;
use std::path::PathBuf;
use std::sync::{Arc, Condvar, Mutex, OnceLock};
use std::time::{Duration, Instant};
use tauri::{AppHandle, Manager};
use uuid::Uuid;

pub(crate) const WPS_BRIDGE_PORT: u16 = 43129;
/// How long Apply waits for the add-in to write the document.
const APPLY_TIMEOUT: Duration = Duration::from_secs(45);
const MAX_POLL_WAIT_SECONDS: u64 = 25;
const MAX_BODY_BYTES: usize = 4 * 1024 * 1024;
/// A document operation's source: up to 1000 equations' XML (the Word add-in's
/// 16 MB formula restore limit) as JSON.
const MAX_DOCUMENT_BODY_BYTES: usize = 24 * 1024 * 1024;
/// How long the renderer waits for the add-in to read the formulas' font sizes.
const PREFLIGHT_TIMEOUT: Duration = Duration::from_secs(60);
const MAX_LOG_BYTES: u64 = 512 * 1024;
const MODULE_JS: &str = include_str!("../../../../office/wps/module.js");
/// A Session the add-in never finished (WPS quit while editing) is forgotten
/// when a new one opens after this long.
const STALE_SESSION_AGE: Duration = Duration::from_secs(12 * 60 * 60);

#[derive(Clone, Debug, PartialEq)]
enum Stage {
    Editing,
    /// A LaTeX redraw asks the add-in for the font size of each formula.
    Preflight,
    Ready,
    Applied,
    Failed(String),
    Cancelled,
}

struct Entry {
    stage: Stage,
    payload: Option<Value>,
    files: HashMap<String, PathBuf>,
    opened: Instant,
    /// A document operation (import, redraw, restore): it has a window, no
    /// editor Session.
    document: bool,
    /// The ranges of a font-size preflight, and the add-in's answer.
    query: Option<Value>,
    answer: Option<Result<Vec<f64>, String>>,
}

struct Bridge {
    app: AppHandle,
    token: String,
    log_path: PathBuf,
    entries: Mutex<HashMap<String, Entry>>,
    changed: Condvar,
    status: Mutex<BridgeStatus>,
}

#[derive(Clone, Debug, Default)]
pub(crate) struct BridgeStatus {
    pub running: bool,
    pub error: Option<String>,
}

static BRIDGE: OnceLock<Arc<Bridge>> = OnceLock::new();

/// Starts the listener once; a second call does nothing.
pub(crate) fn start(app: AppHandle, token: String, log_path: PathBuf) {
    let bridge = Arc::new(Bridge {
        app,
        token,
        log_path,
        entries: Mutex::new(HashMap::new()),
        changed: Condvar::new(),
        status: Mutex::new(BridgeStatus::default()),
    });
    if BRIDGE.set(bridge.clone()).is_err() {
        return;
    }
    tauri::async_runtime::spawn(async move {
        let listener = match tokio::net::TcpListener::bind(("127.0.0.1", WPS_BRIDGE_PORT)).await {
            Ok(listener) => listener,
            Err(error) => {
                bridge.set_status(false, Some(format!(
                    "VisualTeX cannot listen for WPS on 127.0.0.1:{WPS_BRIDGE_PORT}: {error}"
                )));
                return;
            }
        };
        bridge.set_status(true, None);
        let result = axum::serve(listener, router(bridge.clone())).await;
        bridge.set_status(false, result.err().map(|error| format!("The WPS bridge stopped: {error}")));
    });
}

pub(crate) fn status() -> BridgeStatus {
    BRIDGE
        .get()
        .and_then(|bridge| bridge.status.lock().ok().map(|status| status.clone()))
        .unwrap_or_default()
}

impl Bridge {
    fn set_status(&self, running: bool, error: Option<String>) {
        if let Some(message) = &error {
            eprintln!("{message}");
        }
        if let Ok(mut status) = self.status.lock() {
            status.running = running;
            status.error = error;
        }
    }

    fn update(&self, session_id: &str, change: impl FnOnce(&mut Entry)) {
        if let Ok(mut entries) = self.entries.lock() {
            let entry = entries.entry(session_id.to_string()).or_insert_with(|| Entry {
                stage: Stage::Editing,
                payload: None,
                files: HashMap::new(),
                opened: Instant::now(),
                document: false,
                query: None,
                answer: None,
            });
            change(entry);
        }
        self.changed.notify_all();
    }

    fn log(&self, line: &str) {
        if let Some(parent) = self.log_path.parent() {
            let _ = std::fs::create_dir_all(parent);
        }
        if std::fs::metadata(&self.log_path).map(|metadata| metadata.len() > MAX_LOG_BYTES).unwrap_or(false) {
            let _ = std::fs::remove_file(&self.log_path);
        }
        if let Ok(mut file) = std::fs::OpenOptions::new().create(true).append(true).open(&self.log_path) {
            let _ = writeln!(file, "{} {line}", chrono_like_now());
        }
    }
}

fn chrono_like_now() -> String {
    let millis = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|value| value.as_millis())
        .unwrap_or_default();
    format!("epochMs={millis}")
}

/// Hands a committed formula to the add-in and waits until it has written the
/// document. `files` maps the names the add-in downloads to local artifacts.
pub(crate) fn deliver_commit(
    session_id: &str,
    payload: Value,
    files: Vec<(String, PathBuf)>,
) -> Result<(), String> {
    deliver(session_id, payload, files, APPLY_TIMEOUT)
}

/// A document operation's result: every item at once, so the add-in gets
/// time for each (WPS writes about a formula a second at worst).
pub(crate) fn deliver_document_commit(
    session_id: &str,
    payload: Value,
    files: Vec<(String, PathBuf)>,
    item_count: usize,
) -> Result<(), String> {
    let timeout = APPLY_TIMEOUT + Duration::from_secs(2 * item_count.min(2048) as u64);
    deliver(session_id, payload, files, timeout)
}

fn deliver(session_id: &str, payload: Value, files: Vec<(String, PathBuf)>, timeout: Duration) -> Result<(), String> {
    let bridge = BRIDGE.get().ok_or_else(|| "The WPS connection of VisualTeX is not running".to_string())?;
    bridge.log(&format!("session={session_id} result-ready"));
    bridge.update(session_id, |entry| {
        entry.stage = Stage::Ready;
        entry.payload = Some(payload);
        entry.files = files.into_iter().collect();
    });
    let deadline = Instant::now() + timeout;
    let mut entries = bridge
        .entries
        .lock()
        .map_err(|_| "The WPS session state is unavailable".to_string())?;
    loop {
        match entries.get(session_id).map(|entry| entry.stage.clone()) {
            Some(Stage::Applied) => return Ok(()),
            Some(Stage::Failed(error)) => return Err(error),
            Some(Stage::Cancelled) => {
                return Err("WPS cancelled this formula before it was written".to_string())
            }
            _ => {}
        }
        let now = Instant::now();
        if now >= deadline {
            if let Some(entry) = entries.get_mut(session_id) {
                entry.stage = Stage::Failed("timeout".to_string());
            }
            bridge.log(&format!("session={session_id} apply-timeout"));
            return Err(
                "WPS did not write the formula in time. Check that the document is still open in WPS and that the VisualTeX tab is shown, then try again."
                    .to_string(),
            );
        }
        entries = bridge
            .changed
            .wait_timeout(entries, deadline - now)
            .map_err(|_| "The WPS session state is unavailable".to_string())?
            .0;
    }
}

/// Asks the add-in for the font size of each LaTeX formula a redraw found
/// (VTResolveWordLatexRedrawFontsDispatch) and waits for its answer.
pub(crate) fn request_font_sizes(session_id: &str, ranges: Value) -> Result<Vec<f64>, String> {
    let bridge = BRIDGE.get().ok_or_else(|| "The WPS connection of VisualTeX is not running".to_string())?;
    bridge.log(&format!("session={session_id} font-preflight"));
    bridge.update(session_id, |entry| {
        entry.stage = Stage::Preflight;
        entry.query = Some(ranges);
        entry.answer = None;
    });
    let deadline = Instant::now() + PREFLIGHT_TIMEOUT;
    let mut entries = bridge
        .entries
        .lock()
        .map_err(|_| "The WPS session state is unavailable".to_string())?;
    loop {
        if let Some(entry) = entries.get_mut(session_id) {
            if let Some(answer) = entry.answer.take() {
                entry.query = None;
                return answer;
            }
            match entry.stage.clone() {
                Stage::Cancelled => return Err("WPS cancelled the LaTeX redraw".to_string()),
                Stage::Failed(error) => return Err(error),
                _ => {}
            }
        }
        let now = Instant::now();
        if now >= deadline {
            if let Some(entry) = entries.get_mut(session_id) {
                entry.stage = Stage::Editing;
                entry.query = None;
            }
            return Err("WPS did not report the formulas' font sizes in time. Check that the document is still open in WPS, then try again.".to_string());
        }
        entries = bridge
            .changed
            .wait_timeout(entries, deadline - now)
            .map_err(|_| "The WPS session state is unavailable".to_string())?
            .0;
    }
}

/// Tells the add-in the formula was cancelled, so it removes its placeholder.
pub(crate) fn deliver_cancel(session_id: &str) {
    if let Some(bridge) = BRIDGE.get() {
        bridge.log(&format!("session={session_id} cancelled"));
        bridge.update(session_id, |entry| entry.stage = Stage::Cancelled);
    }
}

fn router(bridge: Arc<Bridge>) -> Router {
    Router::new()
        .route("/wps/v1/hello", get(hello))
        .route("/wps/v1/module.js", get(module_js))
        .route("/wps/v1/sessions", post(create_session))
        .route("/wps/v1/sessions/{id}", get(get_session))
        .route("/wps/v1/sessions/{id}/files/{name}", get(get_file))
        .route("/wps/v1/sessions/{id}/applied", post(applied))
        .route("/wps/v1/sessions/{id}/numbering", post(numbering))
        .route("/wps/v1/log", post(log_line))
        .route("/wps/v1/open", post(open_app))
        .route("/wps/v1/numbers/restyle", post(restyle_numbers))
        .route("/wps/v1/numbers/equations", post(renumber_equations))
        .route("/wps/v1/numbering/preference", post(write_numbering_preference))
        .route("/wps/v1/images/numbered", post(numbered_image_metadata))
        .route("/wps/v1/images/letter-fonts", post(image_letter_fonts))
        .route("/wps/v1/formulas/latex", post(formula_metadata_latex))
        .route("/wps/v1/omml/latex", post(omml_latex))
        .route("/wps/v1/conversion-font-sizes", get(read_conversion_font_sizes).post(write_conversion_font_sizes))
        .route(
            "/wps/v1/documents",
            post(create_document).layer(DefaultBodyLimit::max(MAX_DOCUMENT_BODY_BYTES)),
        )
        .route("/wps/v1/sessions/{id}/font-sizes", post(font_sizes))
        .route("/wps/v1/sessions/{id}/progress", post(progress))
        .layer(DefaultBodyLimit::max(MAX_BODY_BYTES))
        .layer(middleware::from_fn_with_state(bridge.clone(), cors_and_token))
        .with_state(bridge)
}

fn constant_time_equal(left: &str, right: &str) -> bool {
    left.len() == right.len()
        && left
            .bytes()
            .zip(right.bytes())
            .fold(0u8, |difference, (a, b)| difference | (a ^ b))
            == 0
}

fn query_token(request: &Request) -> Option<String> {
    request.uri().query()?.split('&').find_map(|pair| {
        let (key, value) = pair.split_once('=')?;
        (key == "token").then(|| value.to_string())
    })
}

/// The add-in page is a file:// page (WPS's own browser), so every response
/// allows any origin; the per-install token is what authorizes a call.
async fn cors_and_token(State(bridge): State<Arc<Bridge>>, request: Request, next: Next) -> Response {
    let mut response = if request.method() == Method::OPTIONS {
        StatusCode::NO_CONTENT.into_response()
    } else if query_token(&request).is_some_and(|token| constant_time_equal(&token, &bridge.token)) {
        next.run(request).await
    } else {
        (StatusCode::UNAUTHORIZED, Json(json!({ "error": "VisualTeX rejected the WPS add-in's key. Repair the WPS add-in in VisualTeX Settings." }))).into_response()
    };
    let headers = response.headers_mut();
    headers.insert(header::ACCESS_CONTROL_ALLOW_ORIGIN, HeaderValue::from_static("*"));
    headers.insert(header::ACCESS_CONTROL_ALLOW_HEADERS, HeaderValue::from_static("Content-Type"));
    headers.insert(header::ACCESS_CONTROL_ALLOW_METHODS, HeaderValue::from_static("GET, POST, OPTIONS"));
    headers.insert("Access-Control-Allow-Private-Network", HeaderValue::from_static("true"));
    headers.insert(header::CACHE_CONTROL, HeaderValue::from_static("no-store"));
    response
}

/// Identifies the served module.js, so an add-in still running an older
/// one (VisualTeX was updated while WPS stayed open) reloads it.
fn module_version() -> &'static str {
    static VERSION: OnceLock<String> = OnceLock::new();
    VERSION.get_or_init(|| {
        use sha2::{Digest, Sha256};
        Sha256::digest(MODULE_JS.as_bytes())
            .iter()
            .take(8)
            .map(|byte| format!("{byte:02x}"))
            .collect()
    })
}

async fn hello() -> Json<Value> {
    Json(json!({
        "ok": true,
        "app": "VisualTeX",
        "version": env!("CARGO_PKG_VERSION"),
        "protocol": 1,
        "module": module_version(),
    }))
}

async fn module_js() -> Response {
    (
        [(header::CONTENT_TYPE, HeaderValue::from_static("text/javascript; charset=utf-8"))],
        MODULE_JS,
    )
        .into_response()
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct CreateSessionBody {
    session_id: String,
    mode: String,
    display_mode: String,
    #[serde(default)]
    native_equation: bool,
    /// An edited numbered formula (it keeps its number).
    #[serde(default)]
    numbered: bool,
    source_document_id: String,
    #[serde(default)]
    formula_id: Option<String>,
    #[serde(default)]
    source_object_id: Option<String>,
    #[serde(default)]
    encoded_metadata: Option<String>,
    #[serde(default)]
    native_edit_xml: Option<String>,
    #[serde(default)]
    font_size_pt: Option<f64>,
    /// "nativeToImage" or "imageToNative": a hidden editor converts the
    /// selected formula (the ribbon's 转为图片公式 / 转为 Word OMML).
    #[serde(default)]
    operation: Option<String>,
}

fn bad_request(message: impl Into<String>) -> Response {
    (StatusCode::BAD_REQUEST, Json(json!({ "error": message.into() }))).into_response()
}

/// The offline request the Word-for-Mac add-in would write, marked as WPS.
fn session_request(body: &CreateSessionBody) -> Result<Value, String> {
    let session_id = Uuid::parse_str(&body.session_id)
        .map_err(|_| "The WPS session id is not a UUID".to_string())?
        .to_string();
    if session_id != body.session_id {
        return Err("The WPS session id must be a lowercase UUID".to_string());
    }
    let operation = body.operation.as_deref().unwrap_or("formula");
    if !matches!(operation, "formula" | "nativeToImage" | "imageToNative") {
        return Err("Unknown WPS formula operation".to_string());
    }
    let mut request = json!({
        "protocolVersion": 1,
        "sessionId": session_id,
        "host": "word",
        "client": "wps",
        "mode": body.mode,
        "operation": operation,
        "displayMode": body.display_mode,
        "numbered": body.numbered,
        "nativeEquation": body.native_equation,
        "sourceDocumentId": body.source_document_id,
    });
    let object = request.as_object_mut().expect("request is an object");
    if body.mode == "create" {
        object.insert("pendingMarker".into(), json!(format!("visualtex:pending:v1:{session_id}")));
    }
    for (key, value) in [
        ("formulaId", &body.formula_id),
        ("sourceObjectId", &body.source_object_id),
        ("encodedMetadata", &body.encoded_metadata),
        ("nativeEditXml", &body.native_edit_xml),
    ] {
        if let Some(value) = value {
            object.insert(key.into(), json!(value));
        }
    }
    if let Some(size) = body.font_size_pt {
        object.insert("fontSizePt".into(), json!(size));
    }
    Ok(request)
}

fn open_on_main_thread(app: &AppHandle, request: Value) -> Result<(), String> {
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    let handle = app.clone();
    app.run_on_main_thread(move || {
        let _ = sender.send(crate::office::macos_offline::open_wps_session(&handle, request));
    })
    .map_err(|error| format!("VisualTeX could not open the formula editor: {error}"))?;
    receiver
        .recv_timeout(Duration::from_secs(20))
        .map_err(|_| "VisualTeX did not open the formula editor in time".to_string())?
}

async fn create_session(State(bridge): State<Arc<Bridge>>, Json(body): Json<CreateSessionBody>) -> Response {
    let request = match session_request(&body) {
        Ok(request) => request,
        Err(error) => return bad_request(error),
    };
    let session_id = body.session_id.clone();
    bridge.log(&format!("session={session_id} open mode={} display={} native={}", body.mode, body.display_mode, body.native_equation));
    // The add-in runs one Session at a time and stops polling a finished one.
    forget_finished(&bridge);
    bridge.update(&session_id, |entry| {
        entry.stage = Stage::Editing;
        entry.payload = None;
        entry.files.clear();
    });
    let app = bridge.app.clone();
    let opened = tokio::task::spawn_blocking(move || open_on_main_thread(&app, request))
        .await
        .map_err(|error| format!("VisualTeX could not open the formula editor: {error}"))
        .and_then(|result| result);
    match opened {
        Ok(()) => Json(json!({ "sessionId": session_id })).into_response(),
        Err(error) => {
            bridge.log(&format!("session={session_id} open-failed {error}"));
            bridge.update(&session_id, |entry| entry.stage = Stage::Failed(error.clone()));
            bad_request(error)
        }
    }
}

/// The add-in runs one Session at a time and stops polling a finished one.
fn forget_finished(bridge: &Bridge) {
    if let Ok(mut entries) = bridge.entries.lock() {
        entries.retain(|_, entry| {
            matches!(entry.stage, Stage::Editing | Stage::Preflight | Stage::Ready)
                && entry.opened.elapsed() < STALE_SESSION_AGE
        });
    }
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct CreateDocumentBody {
    session_id: String,
    /// "documentImport", "latexRedraw" or "formulaRestore".
    operation: String,
    source_document_id: String,
    /// The add-in's VT_D_ anchor on the target range (or insertion point).
    bookmark_name: String,
    default_font_size_pt: f64,
    #[serde(default)]
    redraw_scope: Option<String>,
    #[serde(default)]
    output_kind: Option<String>,
    #[serde(default)]
    source_kind: Option<String>,
    /// The redraw's text, or the formula restore manifest.
    #[serde(default)]
    source: Option<String>,
    #[serde(default)]
    math_font: Option<String>,
}

/// The request the Word-for-Mac add-in writes for a document operation
/// (VTDocumentImportRequestJson and its redraw and restore variants), marked
/// as WPS. An import takes the conversion font sizes (VTWithConversionFontSizes).
fn document_request(body: &CreateDocumentBody) -> Result<Value, String> {
    let session_id = Uuid::parse_str(&body.session_id)
        .map_err(|_| "The WPS session id is not a UUID".to_string())?
        .to_string();
    if session_id != body.session_id {
        return Err("The WPS session id must be a lowercase UUID".to_string());
    }
    if !matches!(body.operation.as_str(), "documentImport" | "latexRedraw" | "formulaRestore") {
        return Err("Unknown WPS document operation".to_string());
    }
    let mut document_import = json!({
        "bookmarkName": body.bookmark_name,
        "defaultFontSizePt": body.default_font_size_pt,
    });
    let fields = document_import.as_object_mut().expect("documentImport is an object");
    for (key, value) in [
        ("redrawScope", &body.redraw_scope),
        ("outputKind", &body.output_kind),
        ("sourceKind", &body.source_kind),
    ] {
        if let Some(value) = value {
            fields.insert(key.into(), json!(value));
        }
    }
    if body.operation == "documentImport" {
        let (inline, display) = conversion_font_sizes();
        if inline > 0.0 {
            fields.insert("inlineFontSizePt".into(), json!(inline));
        }
        if display > 0.0 {
            fields.insert("displayFontSizePt".into(), json!(display));
        }
    }
    let mut request = json!({
        "protocolVersion": 1,
        "sessionId": session_id,
        "host": "word",
        "client": "wps",
        "mode": "create",
        "operation": body.operation,
        "formulaId": null,
        "displayMode": "inline",
        "numbered": false,
        "nativeEquation": false,
        "sourceDocumentId": body.source_document_id,
        "sourceObjectId": null,
        "encodedMetadata": null,
        "pendingMarker": null,
        "powerPoint": null,
        "documentImport": document_import,
    });
    if let Some(font) = body.math_font.as_deref().filter(|font| !font.trim().is_empty()) {
        request["wordDocumentMathFontName"] = json!(font);
    }
    Ok(request)
}

fn open_document_on_main_thread(app: &AppHandle, request: Value, source: Option<String>) -> Result<(), String> {
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    let handle = app.clone();
    app.run_on_main_thread(move || {
        let _ = sender.send(crate::office::macos_offline::open_wps_document_session(&handle, request, source));
    })
    .map_err(|error| format!("VisualTeX could not start the document operation: {error}"))?;
    receiver
        .recv_timeout(Duration::from_secs(20))
        .map_err(|_| "VisualTeX did not start the document operation in time".to_string())?
}

/// "全文格式转换", "LaTeX 重绘" and "批量导入": the add-in sends what the Word
/// add-in writes to its Session folder, and VisualTeX opens the same renderer
/// or import window. The add-in then polls the Session like a formula's.
async fn create_document(State(bridge): State<Arc<Bridge>>, Json(body): Json<CreateDocumentBody>) -> Response {
    let request = match document_request(&body) {
        Ok(request) => request,
        Err(error) => return bad_request(error),
    };
    let session_id = body.session_id.clone();
    bridge.log(&format!(
        "session={session_id} document operation={} output={:?} source={:?}",
        body.operation, body.output_kind, body.source_kind
    ));
    forget_finished(&bridge);
    bridge.update(&session_id, |entry| {
        entry.stage = Stage::Editing;
        entry.payload = None;
        entry.files.clear();
        entry.document = true;
        entry.opened = Instant::now();
    });
    let app = bridge.app.clone();
    let source = body.source;
    let opened = tokio::task::spawn_blocking(move || open_document_on_main_thread(&app, request, source))
        .await
        .map_err(|error| format!("VisualTeX could not start the document operation: {error}"))
        .and_then(|result| result);
    match opened {
        Ok(()) => Json(json!({ "sessionId": session_id })).into_response(),
        Err(error) => {
            bridge.log(&format!("session={session_id} document-open-failed {error}"));
            bridge.update(&session_id, |entry| entry.stage = Stage::Failed(error.clone()));
            bad_request(error)
        }
    }
}

#[derive(Debug, Deserialize)]
struct FontSizesBody {
    #[serde(default)]
    sizes: Option<Vec<f64>>,
    #[serde(default)]
    error: Option<String>,
}

/// The add-in's answer to a font-size preflight.
async fn font_sizes(State(bridge): State<Arc<Bridge>>, Path(id): Path<String>, Json(body): Json<FontSizesBody>) -> Response {
    let answer = match (body.sizes, body.error) {
        (Some(sizes), None) => Ok(sizes),
        (_, Some(error)) => Err(error),
        (None, None) => Err("WPS sent no font sizes".to_string()),
    };
    bridge.log(&format!("session={id} font-preflight answered ok={}", answer.is_ok()));
    bridge.update(&id, |entry| {
        if entry.stage == Stage::Preflight {
            entry.stage = Stage::Editing;
        }
        entry.answer = Some(answer);
    });
    Json(json!({ "ok": true })).into_response()
}

#[derive(Debug, Deserialize)]
struct ProgressBody {
    current: usize,
    total: usize,
}

/// How far the add-in is in writing a document operation (the import
/// window shows it).
async fn progress(Path(id): Path<String>, Json(body): Json<ProgressBody>) -> Response {
    match crate::office::macos_offline::write_document_progress(&id, body.current, body.total, "inserting") {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => bad_request(error),
    }
}

#[derive(Debug, Deserialize)]
struct PollQuery {
    #[serde(default)]
    wait: Option<u64>,
}

fn document_window_open(bridge: &Bridge, session_id: &str) -> bool {
    let label = format!("office-native-document-{}", session_id.replace('-', ""));
    bridge.app.get_webview_window(&label).is_some()
}

/// The Session's state as the add-in sees it. `done` means the editor closed
/// without anything to write (an unchanged edit).
fn snapshot(bridge: &Bridge, session_id: &str) -> Option<Value> {
    let entries = bridge.entries.lock().ok()?;
    let entry = entries.get(session_id)?;
    match &entry.stage {
        Stage::Ready => Some(json!({ "status": "ready", "payload": entry.payload })),
        Stage::Preflight => Some(json!({ "status": "preflight", "ranges": entry.query })),
        Stage::Applied => Some(json!({ "status": "applied" })),
        // An import that WPS could not write can be tried again from its window.
        Stage::Failed(_) if entry.document && document_window_open(bridge, session_id) => None,
        Stage::Failed(error) => Some(json!({ "status": "failed", "error": error })),
        Stage::Cancelled => Some(json!({ "status": "cancelled" })),
        // A document operation lasts while its window is open.
        Stage::Editing if entry.document => {
            if document_window_open(bridge, session_id) || entry.opened.elapsed() < Duration::from_secs(5) {
                None
            } else {
                Some(json!({ "status": "cancelled" }))
            }
        }
        Stage::Editing => {
            let state = bridge.app.try_state::<OfficeCompanionState>()?;
            match state.session_store.get(session_id) {
                Ok(session) => match session.status {
                    OfficeSessionStatus::Completed => Some(json!({ "status": "done" })),
                    OfficeSessionStatus::Cancelled => Some(json!({ "status": "cancelled" })),
                    OfficeSessionStatus::Failed => Some(json!({ "status": "failed", "error": session.error })),
                    _ => None,
                },
                Err(_) => Some(json!({ "status": "cancelled", "error": "VisualTeX no longer has this formula session" })),
            }
        }
    }
}

async fn get_session(
    State(bridge): State<Arc<Bridge>>,
    Path(id): Path<String>,
    Query(query): Query<PollQuery>,
) -> Response {
    if Uuid::parse_str(&id).is_err() {
        return bad_request("Invalid session id");
    }
    let known = bridge.entries.lock().map(|entries| entries.contains_key(&id)).unwrap_or(false);
    if !known {
        return (StatusCode::NOT_FOUND, Json(json!({ "status": "unknown" }))).into_response();
    }
    let deadline = Instant::now() + Duration::from_secs(query.wait.unwrap_or(0).min(MAX_POLL_WAIT_SECONDS));
    loop {
        if let Some(value) = snapshot(&bridge, &id) {
            return Json(value).into_response();
        }
        if Instant::now() >= deadline {
            return Json(json!({ "status": "editing" })).into_response();
        }
        tokio::time::sleep(Duration::from_millis(120)).await;
    }
}

async fn get_file(State(bridge): State<Arc<Bridge>>, Path((id, name)): Path<(String, String)>) -> Response {
    let path = bridge
        .entries
        .lock()
        .ok()
        .and_then(|entries| entries.get(&id).and_then(|entry| entry.files.get(&name).cloned()));
    let Some(path) = path else {
        return (StatusCode::NOT_FOUND, "missing").into_response();
    };
    let content_type = match path.extension().and_then(|value| value.to_str()) {
        Some("svg") => "image/svg+xml",
        Some("png") => "image/png",
        Some("docx") => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        _ => "application/octet-stream",
    };
    match tokio::fs::read(&path).await {
        Ok(bytes) => ([(header::CONTENT_TYPE, HeaderValue::from_static(content_type))], bytes).into_response(),
        Err(error) => (StatusCode::NOT_FOUND, format!("{error}")).into_response(),
    }
}

#[derive(Debug, Deserialize)]
struct AppliedBody {
    ok: bool,
    #[serde(default)]
    error: Option<String>,
}

async fn applied(State(bridge): State<Arc<Bridge>>, Path(id): Path<String>, Json(body): Json<AppliedBody>) -> Response {
    let stage = if body.ok {
        Stage::Applied
    } else {
        Stage::Failed(body.error.unwrap_or_else(|| "WPS could not write the formula".to_string()))
    };
    bridge.log(&format!("session={id} applied={stage:?}"));
    bridge.update(&id, |entry| entry.stage = stage);
    Json(json!({ "ok": true })).into_response()
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct NumberingBody {
    /// "equation" (an OMML equation) or "image".
    kind: String,
    facts: NumberingFacts,
}

/// The user's default numbering style for documents without numbers, as the
/// Word add-in keeps it ("mode|separator").
fn numbering_preference() -> Option<String> {
    let bytes = std::fs::read(numbering_preference_path()?).ok()?;
    (bytes.len() <= 256).then(|| String::from_utf8_lossy(&bytes).into_owned())
}

/// The numbered equation for a planned number: the Session's equation body
/// numbered by the shared writer, with field instructions WPS keeps, in a
/// one-paragraph package the add-in inserts.
fn numbered_equation_file(session_id: &str, payload: &Value, plan: &EquationPlan) -> Result<(String, PathBuf), String> {
    use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
    let body = payload["ommlBase64"]
        .as_str()
        .and_then(|value| URL_SAFE_NO_PAD.decode(value).ok())
        .and_then(|bytes| String::from_utf8(bytes).ok())
        .ok_or_else(|| "This WPS formula has no equation to number".to_string())?;
    let math_font = payload["mathFont"].as_str().unwrap_or_default();
    let numbered = crate::office::omml_batch::numbered_equation(
        &body,
        &plan.sequence_name,
        plan.heading_level,
        &plan.separator,
        plan.literal_prefix.as_deref(),
        plan.restart,
        plan.initial,
        math_font,
    )?;
    let equation = crate::office::omml_write::wps_field_instructions(&numbered)?;
    let docx = crate::office::omml_write::wps_equation_docx(&equation, math_font)?;
    let directory = std::env::temp_dir().join("VisualTeX-WPS");
    std::fs::create_dir_all(&directory)
        .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
    let name = format!("{session_id}-numbered.docx");
    let path = directory.join(&name);
    std::fs::write(&path, docx).map_err(|error| format!("Unable to write {}: {error}", path.display()))?;
    Ok((name, path))
}

/// Numbers a new display formula the add-in is about to write: the add-in
/// sends what the document holds (headings, numbered equations and pictures)
/// and gets the number Word would give it (office::wps::numbering) and, for an
/// equation, the numbered equation to insert.
async fn numbering(State(bridge): State<Arc<Bridge>>, Path(id): Path<String>, Json(body): Json<NumberingBody>) -> Response {
    let payload = bridge
        .entries
        .lock()
        .ok()
        .and_then(|entries| entries.get(&id).and_then(|entry| entry.payload.clone()));
    let Some(payload) = payload else {
        return bad_request("This WPS formula has no result to number");
    };
    let style = numbering_style(&body.facts, numbering_preference().as_deref());
    match body.kind.as_str() {
        "equation" => {
            let plan = plan_equation(&body.facts, &style);
            match numbered_equation_file(&id, &payload, &plan) {
                Ok((name, path)) => {
                    bridge.log(&format!("session={id} numbered equation {plan:?}"));
                    bridge.update(&id, |entry| {
                        entry.files.insert(name.clone(), path);
                    });
                    Json(json!({ "style": style, "plan": plan, "file": name })).into_response()
                }
                Err(error) => bad_request(error),
            }
        }
        "image" => {
            let plan = plan_image(&body.facts, &style);
            bridge.log(&format!("session={id} numbered image {plan:?}"));
            let next_increment = plan
                .next_restart
                .as_ref()
                .map(|next| image_increment_instruction(next.restart_at_one));
            Json(json!({
                "style": style,
                "plan": plan,
                "increment": image_increment_instruction(plan.restart_at_one),
                "current": image_current_instruction(),
                "nextIncrement": next_increment,
            }))
            .into_response()
        }
        _ => bad_request("Unknown numbering kind"),
    }
}

#[derive(Debug, Deserialize)]
struct LogBody {
    message: String,
}

async fn log_line(State(bridge): State<Arc<Bridge>>, Json(body): Json<LogBody>) -> Response {
    let message: String = body
        .message
        .chars()
        .take(2_000)
        .map(|character| if character.is_control() { ' ' } else { character })
        .collect();
    bridge.log(&format!("addin {message}"));
    StatusCode::NO_CONTENT.into_response()
}

fn numbering_preference_path() -> Option<PathBuf> {
    let home = std::env::var_os("HOME")?;
    Some(PathBuf::from(home).join("Library/Application Scripts/com.microsoft.Word/VisualTeXNumberingPreference.txt"))
}

/// A numbering style the user chose ("编号格式"): mode and "." or "-".
fn chosen_style(mode: Option<&str>, separator: Option<&str>) -> Result<Option<NumberingStyle>, String> {
    let Some(mode) = mode.map(|mode| mode.trim().to_ascii_lowercase()) else { return Ok(None) };
    match mode.as_str() {
        MODE_SEQUENCE => Ok(Some(NumberingStyle { mode, separator: ".".to_string() })),
        MODE_CHAPTER | MODE_SECTION => match separator.map(str::trim) {
            Some(separator @ ("." | "-")) => Ok(Some(NumberingStyle { mode, separator: separator.to_string() })),
            _ => Err("The equation number separator must be . or -".to_string()),
        },
        _ => Err("Unknown equation numbering mode".to_string()),
    }
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RestyleBody {
    facts: RestyleFacts,
    /// The style chosen for this document in this WPS session, if any.
    #[serde(default)]
    mode: Option<String>,
    #[serde(default)]
    separator: Option<String>,
}

/// "更新公式编号" and "编号格式": the number every numbered equation and
/// picture takes in the document's style, and which must be written again.
async fn restyle_numbers(Json(body): Json<RestyleBody>) -> Response {
    let style = match chosen_style(body.mode.as_deref(), body.separator.as_deref()) {
        Ok(Some(style)) => style,
        Ok(None) => numbering_style(&body.facts.style_facts(), numbering_preference().as_deref()),
        Err(error) => return bad_request(error),
    };
    Json(json!({
        "style": style,
        "equations": restyle_equations(&body.facts, &style),
        "images": restyle_images(&body.facts, &style),
    }))
    .into_response()
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RenumberItem {
    /// WPS's export (WordOpenXML) of the numbered equation.
    xml: String,
    plan: EquationPlan,
    /// The equation has no number yet ("批量编号"): all of it is the formula.
    #[serde(default)]
    unnumbered: bool,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RenumberBody {
    batch_id: String,
    items: Vec<RenumberItem>,
    #[serde(default)]
    math_font: String,
}

/// Each numbered equation with its number written again for its plan (the
/// formula kept as it is), one package per equation for the add-in to put in
/// its place. Word copies a new number into the equation (VTV2RestyleNumbers);
/// WPS cannot copy part of an equation, so it replaces the whole equation.
async fn renumber_equations(State(bridge): State<Arc<Bridge>>, Json(body): Json<RenumberBody>) -> Response {
    if Uuid::parse_str(&body.batch_id).is_err() || body.items.is_empty() || body.items.len() > 1000 {
        return bad_request("Invalid equation renumbering request");
    }
    let directory = std::env::temp_dir().join("VisualTeX-WPS");
    if let Err(error) = std::fs::create_dir_all(&directory) {
        return bad_request(format!("Unable to create {}: {error}", directory.display()));
    }
    let mut files = Vec::new();
    let mut originals = Vec::new();
    let mut errors = Vec::new();
    for (index, item) in body.items.iter().enumerate() {
        // The equation as it was, for the add-in to put back if WPS does not
        // take the renumbered one after the old one is gone.
        let original = crate::office::omml_write::extract_single_omath(&item.xml)
            .and_then(|equation| crate::office::omml_write::wps_equation_docx(&equation, &body.math_font))
            .ok()
            .and_then(|docx| {
                let name = format!("{}-original-{index}.docx", body.batch_id);
                let path = directory.join(&name);
                std::fs::write(&path, docx).ok()?;
                bridge.update(&body.batch_id, |entry| {
                    entry.stage = Stage::Applied;
                    entry.files.insert(name.clone(), path);
                });
                Some(name)
            });
        originals.push(original.map(Value::String).unwrap_or(Value::Null));
        let plan = &item.plan;
        let formula = if item.unnumbered {
            crate::office::omml_write::extract_single_omath(&item.xml)
        } else {
            crate::office::omml_write::equation_body(&item.xml)
        };
        let built = formula
            .and_then(|formula| {
                crate::office::omml_write::build_word_native_numbered_omml(
                    &formula,
                    &plan.sequence_name,
                    None,
                    plan.heading_level,
                    &plan.separator,
                    &plan.initial.max(1).to_string(),
                    plan.literal_prefix.as_deref(),
                    plan.restart,
                )
            })
            .and_then(|numbered| crate::office::omml_write::wps_field_instructions(&numbered))
            .and_then(|equation| crate::office::omml_write::wps_equation_docx(&equation, &body.math_font));
        match built {
            Ok(docx) => {
                let name = format!("{}-renumbered-{index}.docx", body.batch_id);
                let path = directory.join(&name);
                match std::fs::write(&path, docx) {
                    Ok(()) => {
                        bridge.update(&body.batch_id, |entry| {
                            entry.stage = Stage::Applied;
                            entry.files.insert(name.clone(), path);
                        });
                        files.push(Value::String(name));
                        errors.push(Value::Null);
                    }
                    Err(error) => {
                        files.push(Value::Null);
                        errors.push(Value::String(format!("Unable to write {}: {error}", path.display())));
                    }
                }
            }
            Err(error) => {
                files.push(Value::Null);
                errors.push(Value::String(error));
            }
        }
    }
    bridge.log(&format!("renumbered batch={} items={}", body.batch_id, body.items.len()));
    Json(json!({ "files": files, "originals": originals, "errors": errors })).into_response()
}

#[derive(Debug, Deserialize)]
struct NumberedImagesBody {
    metadata: Vec<String>,
}

/// "批量编号" for picture formulas: each one's metadata marked numbered, and
/// the letter font its number takes (omml_batch's displayNumbering images).
async fn numbered_image_metadata(Json(body): Json<NumberedImagesBody>) -> Response {
    if body.metadata.len() > 1000 {
        return bad_request("Too many picture formulas in one request");
    }
    let items: Vec<Value> = body
        .metadata
        .iter()
        .map(|metadata| match crate::office::macos_offline::numbered_formula_metadata(metadata) {
            Ok(numbered) => json!({
                "metadata": numbered,
                "letterFont": crate::office::omml_batch::formula_letter_font(metadata),
            }),
            Err(error) => json!({ "error": error }),
        })
        .collect();
    Json(json!({ "items": items })).into_response()
}

/// The letter font each picture formula's metadata records, for the font of
/// its number (VTEnsureWordImageLetterFonts).
async fn image_letter_fonts(Json(body): Json<NumberedImagesBody>) -> Response {
    if body.metadata.len() > 1000 {
        return bad_request("Too many picture formulas in one request");
    }
    let fonts: Vec<Value> = body
        .metadata
        .iter()
        .map(|metadata| {
            crate::office::omml_batch::formula_letter_font(metadata)
                .map(Value::String)
                .unwrap_or(Value::Null)
        })
        .collect();
    Json(json!({ "fonts": fonts })).into_response()
}

/// The LaTeX a picture formula's metadata holds, for "将图片公式转为 LaTeX" when
/// the document variables no longer hold it.
async fn formula_metadata_latex(Json(body): Json<NumberedImagesBody>) -> Response {
    if body.metadata.len() > 1000 {
        return bad_request("Too many picture formulas in one request");
    }
    let items: Vec<Value> = body
        .metadata
        .iter()
        .map(|metadata| match crate::office::macos_offline::decode_metadata(metadata) {
            Ok(decoded) => json!({ "latex": decoded.latex }),
            Err(error) => json!({ "error": error }),
        })
        .collect();
    Json(json!({ "items": items })).into_response()
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OmmlLatexItem {
    /// WPS's export (WordOpenXML) of the equation.
    xml: String,
    display_mode: String,
}

#[derive(Debug, Deserialize)]
struct OmmlLatexBody {
    items: Vec<OmmlLatexItem>,
}

/// "将 OMML 公式转为 LaTeX": each equation read as LaTeX, "$...$" inline and
/// "$$...$$" on its own line (omml_batch's formula restore).
async fn omml_latex(Json(body): Json<OmmlLatexBody>) -> Response {
    if body.items.len() > 1000 {
        return bad_request("Too many equations in one request");
    }
    let items: Vec<Value> = body
        .items
        .iter()
        .map(|item| {
            let display = item.display_mode == "block";
            crate::office::macos_offline::word_omml_to_mathml(&item.xml, display)
                .and_then(|math_ml| crate::office::omml_batch::mathml_to_latex(&math_ml))
                .map(|latex| {
                    let text = if display { format!("$${latex}$$") } else { format!("${latex}$") };
                    json!({ "text": text })
                })
                .unwrap_or_else(|error| json!({ "error": error }))
        })
        .collect();
    Json(json!({ "items": items })).into_response()
}

/// The formula sizes "行内转换" and "行间转换" set, shared with the Word add-in
/// ("inline=12|display=14"; 0 follows the body text).
fn conversion_font_sizes_path() -> Option<PathBuf> {
    let home = std::env::var_os("HOME")?;
    Some(PathBuf::from(home).join("Library/Application Scripts/com.microsoft.Word/VisualTeXConversionFontSizes.txt"))
}

fn valid_formula_size(value: f64) -> f64 {
    if (1.0..=512.0).contains(&value) {
        value
    } else {
        0.0
    }
}

async fn read_conversion_font_sizes() -> Response {
    let (inline, display) = conversion_font_sizes();
    Json(json!({ "inline": inline, "display": display })).into_response()
}

/// The inline and display conversion sizes (0 follows the body text).
fn conversion_font_sizes() -> (f64, f64) {
    let (mut inline, mut display) = (0.0, 0.0);
    if let Some(text) = conversion_font_sizes_path()
        .and_then(|path| std::fs::read(path).ok())
        .filter(|bytes| bytes.len() <= 256)
        .map(|bytes| String::from_utf8_lossy(&bytes).into_owned())
    {
        for field in text.trim().split('|') {
            if let Some((name, value)) = field.split_once('=') {
                let value = valid_formula_size(value.trim().parse::<f64>().unwrap_or(0.0));
                match name.trim().to_ascii_lowercase().as_str() {
                    "inline" => inline = value,
                    "display" => display = value,
                    _ => {}
                }
            }
        }
    }
    (inline, display)
}

#[derive(Debug, Deserialize)]
struct ConversionFontSizesBody {
    inline: f64,
    display: f64,
}

async fn write_conversion_font_sizes(Json(body): Json<ConversionFontSizesBody>) -> Response {
    let Some(path) = conversion_font_sizes_path() else {
        return bad_request("Unable to resolve the conversion font size file");
    };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let text = format!(
        "inline={}|display={}",
        valid_formula_size(body.inline),
        valid_formula_size(body.display)
    );
    match std::fs::write(&path, text) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => bad_request(format!("Unable to write {}: {error}", path.display())),
    }
}

#[derive(Debug, Deserialize)]
struct PreferenceBody {
    mode: String,
    #[serde(default)]
    separator: Option<String>,
}

/// The format chosen in "编号格式" becomes the default for documents without
/// numbers, shared with the Word add-in (VTWriteEquationNumberingPreference).
async fn write_numbering_preference(Json(body): Json<PreferenceBody>) -> Response {
    let style = match chosen_style(Some(&body.mode), body.separator.as_deref()) {
        Ok(Some(style)) => style,
        Ok(None) => return bad_request("Unknown equation numbering mode"),
        Err(error) => return bad_request(error),
    };
    let Some(path) = numbering_preference_path() else {
        return bad_request("Unable to resolve the numbering preference file");
    };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    match std::fs::write(&path, format!("{}|{}", style.mode, style.separator)) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => bad_request(format!("Unable to write {}: {error}", path.display())),
    }
}

/// The ribbon's "Open VisualTeX" (VTWordRibbonOpen): shows the main window.
async fn open_app(State(bridge): State<Arc<Bridge>>) -> Response {
    match crate::office::background::reveal_main_window(&bridge.app) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => (StatusCode::INTERNAL_SERVER_ERROR, Json(json!({ "error": error }))).into_response(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn body(mode: &str) -> CreateSessionBody {
        CreateSessionBody {
            session_id: "0f8fad5b-d9cb-469f-a165-70867728950e".to_string(),
            mode: mode.to_string(),
            display_mode: "inline".to_string(),
            native_equation: false,
            numbered: false,
            source_document_id: "wps:/Users/a/doc.docx".to_string(),
            formula_id: None,
            source_object_id: None,
            encoded_metadata: None,
            native_edit_xml: None,
            font_size_pt: Some(12.0),
            operation: None,
        }
    }

    #[test]
    fn create_requests_carry_a_pending_marker_and_the_wps_client() {
        let request = session_request(&body("create")).unwrap();
        assert_eq!(request["client"], "wps");
        assert_eq!(request["host"], "word");
        assert_eq!(request["pendingMarker"], "visualtex:pending:v1:0f8fad5b-d9cb-469f-a165-70867728950e");
        assert_eq!(request["fontSizePt"], 12.0);
        assert!(request.get("encodedMetadata").is_none());
        let edit = session_request(&body("edit")).unwrap();
        assert!(edit.get("pendingMarker").is_none());
        assert_eq!(edit["numbered"], false);
        let mut numbered = body("edit");
        numbered.numbered = true;
        assert_eq!(session_request(&numbered).unwrap()["numbered"], true);
        assert_eq!(edit["operation"], "formula");
        let mut conversion = body("edit");
        conversion.operation = Some("nativeToImage".to_string());
        assert_eq!(session_request(&conversion).unwrap()["operation"], "nativeToImage");
        conversion.operation = Some("documentImport".to_string());
        assert!(session_request(&conversion).is_err());
    }

    #[test]
    fn session_ids_must_be_canonical() {
        let mut upper = body("create");
        upper.session_id = upper.session_id.to_uppercase();
        assert!(session_request(&upper).is_err());
        let mut junk = body("create");
        junk.session_id = "not-a-uuid".to_string();
        assert!(session_request(&junk).is_err());
    }

    #[test]
    fn module_version_is_a_short_stable_hash() {
        let version = module_version();
        assert_eq!(version.len(), 16);
        assert!(version.chars().all(|c| c.is_ascii_hexdigit()));
        assert_eq!(version, module_version());
    }

    #[test]
    fn tokens_compare_exactly() {
        assert!(constant_time_equal("abc", "abc"));
        assert!(!constant_time_equal("abc", "abd"));
        assert!(!constant_time_equal("abc", "abcd"));
    }
}
