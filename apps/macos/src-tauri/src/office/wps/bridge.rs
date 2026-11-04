//! Local HTTP channel between the WPS add-ins and VisualTeX (127.0.0.1 only).
//!
//! WPS for Mac cannot be scripted from outside and its add-ins cannot start an
//! app, so the add-ins talk to the running VisualTeX instead: an add-in opens a
//! formula Session here, polls it, and writes the result the editor produced
//! into the document. Apply in the editor waits until the add-in reports the
//! document written, as a Word Apply waits for the Word macro. The Writer
//! add-in sends Word requests, the Presentation add-in PowerPoint requests.

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
use std::collections::{HashMap, VecDeque};
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
/// The WPS Presentation add-in's logic (office/wps/module-wpp.js).
const MODULE_WPP_JS: &str = include_str!("../../../../office/wps/module-wpp.js");
/// How long the Presentation add-in may wait for an event before it asks again.
const MAX_EVENT_WAIT_SECONDS: u64 = 25;
/// A double-click the add-in hears about later than this is not acted on.
const EVENT_LIFETIME: Duration = Duration::from_secs(2);
const MAX_PENDING_EVENTS: usize = 16;
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
    /// An AI agent's formula or import (office::automation::writes): the app opens it a moment after the
    /// add-in registered it (macos_offline::open_agent_session, open_agent_import).
    agent: bool,
}

struct Bridge {
    app: AppHandle,
    token: String,
    log_path: PathBuf,
    entries: Mutex<HashMap<String, Entry>>,
    changed: Condvar,
    status: Mutex<BridgeStatus>,
    events: Mutex<EventLog>,
    agent: Mutex<AgentChannel>,
}

// AI agents (office::automation::hosts, docs/office-mcp/MACOS.md): WPS cannot be scripted from outside, so the
// add-ins ask here for an agent's commands, run them with WPS's own API and post the results.

/// How long an add-in waits for a command before it asks again.
const AGENT_POLL_SECONDS: u64 = 20;
/// An add-in that has not asked for this long is gone (WPS quit, or the add-in is not loaded).
const AGENT_HOST_LIFETIME: Duration = Duration::from_secs(AGENT_POLL_SECONDS + 25);
/// How long a command may wait for an add-in to take it.
const AGENT_TAKE_TIMEOUT: Duration = Duration::from_secs(AGENT_POLL_SECONDS + 15);
/// How long an agent's formula or import may wait for the app to open it.
const AGENT_OPEN_GRACE: Duration = Duration::from_secs(60);

struct AgentCommand {
    id: String,
    host: String,
    operation: String,
    arguments: Value,
}

#[derive(Default)]
struct AgentChannel {
    /// When each add-in ("writer", "presentation") last asked, and what it said about itself.
    seen: HashMap<String, (Instant, Value)>,
    queue: VecDeque<AgentCommand>,
    /// Commands an add-in has taken and not answered yet.
    taken: HashMap<String, Instant>,
    results: HashMap<String, Value>,
}

/// What the add-ins learn from VisualTeX rather than from WPS: WPS
/// Presentation raises no double-click event, so the Office double-click
/// monitor (office::powerpoint_native) reports double-clicks in WPS here.
#[derive(Default)]
struct EventLog {
    sequence: u64,
    recent: VecDeque<(u64, Instant, Value)>,
}

impl EventLog {
    fn push(&mut self, event: Value) {
        self.sequence += 1;
        self.recent.push_back((self.sequence, Instant::now(), event));
        while self.recent.len() > MAX_PENDING_EVENTS {
            self.recent.pop_front();
        }
    }

    /// The events after `after` that are recent enough to act on.
    fn since(&self, after: u64) -> Vec<Value> {
        self.recent
            .iter()
            .filter(|(sequence, at, _)| *sequence > after && at.elapsed() < EVENT_LIFETIME)
            .map(|(sequence, _, event)| {
                let mut event = event.clone();
                event["sequence"] = json!(sequence);
                event
            })
            .collect()
    }
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
        events: Mutex::new(EventLog::default()),
        agent: Mutex::new(AgentChannel::default()),
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
                agent: false,
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

/// A double-click in WPS, at a point in the screen coordinates Office uses
/// (top-left origin, points). The Presentation add-in decides whether it hit
/// a VisualTeX formula.
pub(crate) fn notify_double_click(x: f64, y: f64) {
    if !x.is_finite() || !y.is_finite() {
        return;
    }
    if let Some(bridge) = BRIDGE.get() {
        if let Ok(mut events) = bridge.events.lock() {
            events.push(json!({ "kind": "doubleClick", "x": x, "y": y }));
        }
    }
}

/// Tells the add-in the formula was cancelled, so it removes its placeholder.
pub(crate) fn deliver_cancel(session_id: &str) {
    if let Some(bridge) = BRIDGE.get() {
        bridge.log(&format!("session={session_id} cancelled"));
        bridge.update(session_id, |entry| entry.stage = Stage::Cancelled);
    }
}

/// The WPS add-ins that take agent commands now ("writer", "presentation"), with what each said about itself
/// (its WPS version).
pub(crate) fn agent_hosts() -> Vec<(String, Value)> {
    let Some(bridge) = BRIDGE.get() else { return Vec::new() };
    let Ok(channel) = bridge.agent.lock() else { return Vec::new() };
    channel
        .seen
        .iter()
        .filter(|(_, (at, _))| at.elapsed() < AGENT_HOST_LIFETIME)
        .map(|(host, (_, info))| (host.clone(), info.clone()))
        .collect()
}

/// Runs one agent operation in a WPS add-in and returns its data; the error is (code, message).
pub(crate) async fn agent_run(host: &str, operation: &str, arguments: Value, deadline: Duration)
    -> Result<Value, (String, String)> {
    let bridge = BRIDGE.get().ok_or_else(|| ("internal".to_string(), "The WPS connection of VisualTeX is not running".to_string()))?;
    let id = Uuid::new_v4().to_string();
    if let Ok(mut channel) = bridge.agent.lock() {
        channel.queue.push_back(AgentCommand {
            id: id.clone(),
            host: host.to_string(),
            operation: operation.to_string(),
            arguments,
        });
    }
    bridge.log(&format!("agent command={id} host={host} op={operation}"));
    let started = Instant::now();
    loop {
        let (result, waiting, taken) = match bridge.agent.lock() {
            Ok(mut channel) => {
                let result = channel.results.remove(&id);
                let waiting = channel.queue.iter().any(|command| command.id == id);
                let taken = channel.taken.contains_key(&id);
                (result, waiting, taken)
            }
            Err(_) => return Err(("internal".into(), "The WPS agent channel is unavailable".into())),
        };
        if let Some(result) = result {
            return agent_answer(result);
        }
        let late = (waiting && started.elapsed() > AGENT_TAKE_TIMEOUT.min(deadline)) || started.elapsed() > deadline;
        if late {
            if let Ok(mut channel) = bridge.agent.lock() {
                channel.queue.retain(|command| command.id != id);
                channel.taken.remove(&id);
            }
            bridge.log(&format!("agent command={id} timeout taken={taken}"));
            return Err(if taken {
                ("host_unresponsive".into(), "WPS did not finish in time (a dialog may be open in WPS)".into())
            } else {
                ("addin_not_loaded".into(), "The VisualTeX add-in in WPS does not answer; show the VisualTeX tab in WPS, or restart WPS".into())
            });
        }
        tokio::time::sleep(Duration::from_millis(40)).await;
    }
}

fn agent_answer(result: Value) -> Result<Value, (String, String)> {
    if result.get("ok").and_then(Value::as_bool) == Some(true) {
        return Ok(result.get("data").cloned().unwrap_or(Value::Null));
    }
    let error = result.get("error");
    let field = |name: &str, fallback: &str| {
        error.and_then(|error| error.get(name)).and_then(Value::as_str).unwrap_or(fallback).to_string()
    };
    Err((field("code", "operation_failed"), field("message", "WPS reported a failure")))
}

fn agent_enabled(bridge: &Bridge) -> bool {
    bridge
        .app
        .try_state::<OfficeCompanionState>()
        .is_some_and(|state| crate::office::automation::settings::settings(&state.paths.root).enabled)
}

#[derive(Debug, Deserialize)]
struct AgentPollQuery {
    /// "writer" or "presentation"
    host: String,
    #[serde(default)]
    wait: Option<u64>,
    /// what the add-in says about itself (the WPS version)
    #[serde(default)]
    version: Option<String>,
}

/// The add-in asks for its next command; it is told when agent access is off (and asks again later).
async fn agent_commands(State(bridge): State<Arc<Bridge>>, Query(query): Query<AgentPollQuery>) -> Response {
    if !matches!(query.host.as_str(), "writer" | "presentation") {
        return bad_request("Unknown WPS add-in");
    }
    let enabled = agent_enabled(&bridge);
    if let Ok(mut channel) = bridge.agent.lock() {
        let info = json!({ "version": query.version.clone().unwrap_or_default() });
        if enabled {
            channel.seen.insert(query.host.clone(), (Instant::now(), info));
        } else {
            channel.seen.remove(&query.host);
        }
    }
    if !enabled {
        return Json(json!({ "enabled": false, "commands": [] })).into_response();
    }
    let deadline = Instant::now() + Duration::from_secs(query.wait.unwrap_or(0).min(AGENT_POLL_SECONDS));
    loop {
        let next = bridge.agent.lock().ok().and_then(|mut channel| {
            let index = channel.queue.iter().position(|command| command.host == query.host)?;
            let command = channel.queue.remove(index)?;
            channel.taken.insert(command.id.clone(), Instant::now());
            channel.seen.insert(query.host.clone(), (Instant::now(), json!({ "version": query.version.clone().unwrap_or_default() })));
            Some(command)
        });
        if let Some(command) = next {
            bridge.log(&format!("agent command={} taken", command.id));
            return Json(json!({
                "enabled": true,
                "commands": [{ "id": command.id, "op": command.operation, "arguments": command.arguments }],
            }))
            .into_response();
        }
        if Instant::now() >= deadline {
            return Json(json!({ "enabled": true, "commands": [] })).into_response();
        }
        tokio::time::sleep(Duration::from_millis(40)).await;
    }
}

/// `{id, ok, data}` or `{id, ok: false, error: {code, message}}`.
async fn agent_result(State(bridge): State<Arc<Bridge>>, Json(body): Json<Value>) -> Response {
    let Some(id) = body.get("id").and_then(Value::as_str).map(str::to_string) else {
        return bad_request("The agent result has no id");
    };
    bridge.log(&format!("agent command={id} answered ok={}", body.get("ok").and_then(Value::as_bool).unwrap_or(false)));
    if let Ok(mut channel) = bridge.agent.lock() {
        if channel.taken.remove(&id).is_some() {
            channel.results.insert(id, body);
        }
    }
    Json(json!({ "ok": true })).into_response()
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
        .route("/wps/v1/events", get(events))
        .route("/wps/v1/presentation/equations", post(presentation_equations))
        .route("/wps/v1/agent/commands", get(agent_commands))
        .route("/wps/v1/agent/results", post(agent_result).layer(DefaultBodyLimit::max(MAX_DOCUMENT_BODY_BYTES)))
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

#[derive(Debug, Deserialize)]
struct ModuleQuery {
    /// "wpp" for the Presentation add-in; the Writer add-in sends none.
    #[serde(default)]
    host: Option<String>,
}

fn for_presentation(query: &ModuleQuery) -> bool {
    query.host.as_deref() == Some("wpp")
}

/// The logic an add-in loads: the Writer add-in's or the Presentation add-in's.
fn module_source(query: &ModuleQuery) -> &'static str {
    if for_presentation(query) {
        MODULE_WPP_JS
    } else {
        MODULE_JS
    }
}

fn short_hash(source: &str) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(source.as_bytes())
        .iter()
        .take(8)
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

/// Identifies the served module, so an add-in still running an older one
/// (VisualTeX was updated while WPS stayed open) reloads it.
fn module_version(query: &ModuleQuery) -> &'static str {
    static WRITER: OnceLock<String> = OnceLock::new();
    static PRESENTATION: OnceLock<String> = OnceLock::new();
    let version = if for_presentation(query) { &PRESENTATION } else { &WRITER };
    version.get_or_init(|| short_hash(module_source(query)))
}

async fn hello(Query(query): Query<ModuleQuery>) -> Json<Value> {
    Json(json!({
        "ok": true,
        "app": "VisualTeX",
        "version": env!("CARGO_PKG_VERSION"),
        "protocol": 1,
        "module": module_version(&query),
    }))
}

async fn module_js(Query(query): Query<ModuleQuery>) -> Response {
    (
        [(header::CONTENT_TYPE, HeaderValue::from_static("text/javascript; charset=utf-8"))],
        module_source(&query),
    )
        .into_response()
}

#[derive(Debug, Deserialize)]
struct EventsQuery {
    /// The last event the add-in has seen; none asks only for the sequence.
    #[serde(default)]
    after: Option<u64>,
    #[serde(default)]
    wait: Option<u64>,
}

/// The Presentation add-in waits here for double-clicks (notify_double_click).
async fn events(State(bridge): State<Arc<Bridge>>, Query(query): Query<EventsQuery>) -> Response {
    let read = |bridge: &Bridge| {
        bridge
            .events
            .lock()
            .map(|events| {
                // A sequence ahead of ours is from before VisualTeX restarted.
                let after = query.after.map(|after| if after > events.sequence { 0 } else { after });
                (events.sequence, after.map(|after| events.since(after)).unwrap_or_default())
            })
            .unwrap_or_default()
    };
    let deadline = Instant::now() + Duration::from_secs(query.wait.unwrap_or(0).min(MAX_EVENT_WAIT_SECONDS));
    loop {
        let (sequence, events) = read(&bridge);
        if !events.is_empty() || query.after.is_none() || Instant::now() >= deadline {
            return Json(json!({ "sequence": sequence, "events": events })).into_response();
        }
        tokio::time::sleep(Duration::from_millis(40)).await;
    }
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct CreateSessionBody {
    session_id: String,
    /// "powerpoint" from the Presentation add-in; the Writer add-in sends none.
    #[serde(default)]
    host: Option<String>,
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
    /// A Presentation formula's slide shape, as the PowerPoint add-in
    /// describes it (VTPowerPointGeometryJson).
    #[serde(default)]
    power_point: Option<Value>,
    /// A pasted copy of a Presentation formula gets its own identity
    /// (VTShouldForkPowerPointFormulaCopy).
    #[serde(default)]
    fork_copied_formula: bool,
    /// An AI agent's formula: the request goes back to the add-in, and the app draws the agent's LaTeX in
    /// the hidden editor (office::automation::writes) instead of opening the editor.
    #[serde(default)]
    agent: bool,
}

fn bad_request(message: impl Into<String>) -> Response {
    (StatusCode::BAD_REQUEST, Json(json!({ "error": message.into() }))).into_response()
}

/// The offline request the Word-for-Mac add-in would write, marked as WPS
/// (or the PowerPoint add-in's, for the Presentation add-in).
fn session_request(body: &CreateSessionBody) -> Result<Value, String> {
    let session_id = Uuid::parse_str(&body.session_id)
        .map_err(|_| "The WPS session id is not a UUID".to_string())?
        .to_string();
    if session_id != body.session_id {
        return Err("The WPS session id must be a lowercase UUID".to_string());
    }
    match body.host.as_deref().unwrap_or("word") {
        "word" => {}
        "powerpoint" => return presentation_session_request(body, &session_id),
        _ => return Err("Unknown WPS formula host".to_string()),
    }
    if body.power_point.is_some() || body.fork_copied_formula {
        return Err("A WPS Writer formula has no slide shape".to_string());
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

/// The request the PowerPoint-for-Mac add-in writes, marked as WPS: for a
/// picture formula (VisualTeX_NewFormula, VTPowerPointEditShape) a display
/// formula sized and placed by the slide shape it replaces; for a native
/// equation (VTNativeCreateRequest, VTNativeEditRequest) the text it goes into
/// or replaces, or the slide it gets a new text box on.
fn presentation_session_request(body: &CreateSessionBody, session_id: &str) -> Result<Value, String> {
    if body.operation.as_deref().unwrap_or("formula") != "formula" || body.numbered || body.font_size_pt.is_some() {
        return Err("WPS Presentation formulas are sized by their slide shape and have no number".to_string());
    }
    let power_point = body
        .power_point
        .as_ref()
        .filter(|value| value.is_object())
        .ok_or_else(|| "The WPS Presentation formula has no slide shape".to_string())?;
    if body.native_equation {
        return native_presentation_request(body, session_id, power_point);
    }
    if body.native_edit_xml.is_some() {
        return Err("A WPS Presentation picture formula has no native equation".to_string());
    }
    let mut request = json!({
        "protocolVersion": 1,
        "sessionId": session_id,
        "host": "powerpoint",
        "client": "wps",
        "mode": body.mode,
        "operation": "formula",
        "displayMode": "block",
        "numbered": false,
        "nativeEquation": false,
        "sourceDocumentId": body.source_document_id,
        "powerPoint": power_point,
        "forkCopiedFormula": body.fork_copied_formula,
    });
    let object = request.as_object_mut().expect("request is an object");
    if body.mode == "create" {
        // The placeholder shape carries the same marker (VTPendingMarker).
        let formula_id = body
            .formula_id
            .as_deref()
            .ok_or_else(|| "A new WPS Presentation formula needs its formula id".to_string())?;
        object.insert("pendingMarker".into(), json!(format!("visualtex:pending:v1:{session_id}:{formula_id}")));
    }
    for (key, value) in [
        ("formulaId", &body.formula_id),
        ("sourceObjectId", &body.source_object_id),
        ("encodedMetadata", &body.encoded_metadata),
    ] {
        if let Some(value) = value {
            object.insert(key.into(), json!(value));
        }
    }
    Ok(request)
}

fn native_presentation_request(body: &CreateSessionBody, session_id: &str, power_point: &Value) -> Result<Value, String> {
    let target = body.source_object_id.as_deref().unwrap_or_default();
    if !(target.starts_with("visualtex-ppt-native-text:") || target.starts_with("visualtex-ppt-native-slide:"))
        || body.encoded_metadata.is_some()
        || body.fork_copied_formula
        || !matches!(body.display_mode.as_str(), "inline" | "block")
    {
        return Err("A WPS Presentation native equation needs its text or slide".to_string());
    }
    // An edit carries the equation as WPS saved it (presentation::Zone).
    if (body.mode == "edit") != body.native_edit_xml.is_some() {
        return Err("A WPS Presentation native equation edit needs the equation it replaces".to_string());
    }
    let mut request = json!({
        "protocolVersion": 1,
        "sessionId": session_id,
        "host": "powerpoint",
        "client": "wps",
        "mode": body.mode,
        "operation": "formula",
        "displayMode": body.display_mode,
        "numbered": false,
        "nativeEquation": true,
        "sourceDocumentId": body.source_document_id,
        "sourceObjectId": target,
        "powerPoint": power_point,
    });
    if let Some(xml) = &body.native_edit_xml {
        request["nativeEditXml"] = json!(xml);
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
    bridge.log(&format!("session={session_id} open mode={} display={} native={} agent={}", body.mode, body.display_mode, body.native_equation, body.agent));
    // The add-in runs one Session at a time and stops polling a finished one.
    forget_finished(&bridge);
    bridge.update(&session_id, |entry| {
        entry.stage = Stage::Editing;
        entry.payload = None;
        entry.files.clear();
        entry.agent = body.agent;
        entry.opened = Instant::now();
    });
    if body.agent {
        return Json(json!({ "sessionId": session_id, "request": request })).into_response();
    }
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
    /// "powerpoint" from the Presentation add-in (a LaTeX redraw or a formula
    /// restore); the Writer add-in sends none.
    #[serde(default)]
    host: Option<String>,
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
    /// An AI agent's import: the request goes back to the add-in and the app opens the import window with the
    /// agent's text (macos_offline::open_agent_import).
    #[serde(default)]
    agent: bool,
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
    match body.host.as_deref().unwrap_or("word") {
        "word" => {}
        "powerpoint" => return presentation_document_request(body, &session_id),
        _ => return Err("Unknown WPS document host".to_string()),
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

/// A WPS Presentation LaTeX redraw or formula restore: the request the
/// PowerPoint add-in writes (office::macos_offline::presentation_redraw). Its
/// source is the text boxes' text, or the formulas to restore, as JSON.
fn presentation_document_request(body: &CreateDocumentBody, session_id: &str) -> Result<Value, String> {
    if !matches!(body.operation.as_str(), "latexRedraw" | "formulaRestore") || body.source.is_none() || body.math_font.is_some() {
        return Err("WPS Presentation redraws LaTeX and restores formulas only".to_string());
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
    Ok(json!({
        "protocolVersion": 1,
        "sessionId": session_id,
        "host": "powerpoint",
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
    }))
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
        entry.agent = body.agent;
    });
    if body.agent && body.operation == "documentImport" {
        return Json(json!({ "sessionId": session_id, "request": request })).into_response();
    }
    if body.agent {
        // the window tells the agent what it skipped instead of showing it (note_macos_agent_session)
        crate::office::macos_offline::watch_agent_session(&session_id);
    }
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
            let grace = if entry.agent { AGENT_OPEN_GRACE } else { Duration::from_secs(5) };
            if document_window_open(bridge, session_id) || entry.opened.elapsed() < grace {
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
                // an agent's formula until the app has opened it
                Err(_) if entry.agent && entry.opened.elapsed() < AGENT_OPEN_GRACE => None,
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

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ShapeText {
    slide_id: u32,
    shape_id: u32,
    /// The shape's text as WPS gives it now.
    text: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct EquationsBody {
    /// A copy of the presentation WPS saved (Presentation.SaveCopyAs).
    path: String,
    /// One shape, answered with its zones...
    #[serde(default)]
    slide_id: Option<u32>,
    #[serde(default)]
    shape_id: Option<u32>,
    #[serde(default)]
    text: Option<String>,
    /// ...or several (a formula restore), each answered on its own.
    #[serde(default)]
    shapes: Option<Vec<ShapeText>>,
}

/// Where the native equations of text boxes are, from a copy WPS saved:
/// WPS's object model has no equations (super::presentation). The copy is
/// read once and deleted.
async fn presentation_equations(Json(body): Json<EquationsBody>) -> Response {
    let found = tokio::task::spawn_blocking(move || {
        let home = std::env::var_os("HOME").map(PathBuf::from).ok_or("Unable to resolve the home folder")?;
        let path = super::presentation::saved_copy_path(&body.path, &home)?;
        let answer = match (&body.shapes, body.slide_id, body.shape_id, &body.text) {
            (Some(shapes), None, None, None) if !shapes.is_empty() && shapes.len() <= 4096 => {
                let list: Vec<(u32, u32, String)> =
                    shapes.iter().map(|shape| (shape.slide_id, shape.shape_id, shape.text.clone())).collect();
                super::presentation::saved_shapes_equations(&path, &list).map(|results| {
                    let shapes: Vec<Value> = list
                        .iter()
                        .zip(results)
                        .map(|((slide_id, shape_id, _), result)| match result {
                            Ok(zones) => json!({ "slideId": slide_id, "shapeId": shape_id, "zones": zones }),
                            Err(error) => json!({ "slideId": slide_id, "shapeId": shape_id, "error": error }),
                        })
                        .collect();
                    json!({ "shapes": shapes })
                })
            }
            (None, Some(slide_id), Some(shape_id), Some(text)) => {
                super::presentation::saved_shape_equations(&path, slide_id, shape_id, text).map(|zones| json!({ "zones": zones }))
            }
            _ => Err("Name one text box, or 1 to 4096 of them".to_string()),
        };
        let _ = std::fs::remove_file(&path);
        answer
    })
    .await
    .map_err(|error| format!("Reading WPS's saved copy failed: {error}"))
    .and_then(|result| result);
    match found {
        Ok(answer) => Json(answer).into_response(),
        Err(error) => bad_request(error),
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
            host: None,
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
            power_point: None,
            fork_copied_formula: false,
            agent: false,
        }
    }

    fn slide_shape() -> Value {
        json!({
            "presentationIdentity": "Deck.pptx", "slideIndex": 1, "slideId": 256, "shapeIndex": 2,
            "shapeId": 5, "shapeName": "VisualTeX_7c9e6679-7425-40de-944b-e07fc1f90ae7",
            "left": 100.0, "top": 80.0, "width": 180.0, "height": 42.0, "rotation": 0.0, "zOrder": 2,
            "fontSizePt": 18.0, "referenceWidthPt": null, "referenceHeightPt": null
        })
    }

    fn presentation_body(mode: &str) -> CreateSessionBody {
        let mut body = body(mode);
        body.host = Some("powerpoint".to_string());
        body.display_mode = "block".to_string();
        body.font_size_pt = None;
        body.formula_id = Some("7c9e6679-7425-40de-944b-e07fc1f90ae7".to_string());
        body.source_document_id = "Deck.pptx".to_string();
        body.power_point = Some(slide_shape());
        body
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
    fn presentation_requests_are_powerpoint_picture_formulas() {
        let create = session_request(&presentation_body("create")).unwrap();
        assert_eq!(create["host"], "powerpoint");
        assert_eq!(create["client"], "wps");
        assert_eq!(create["displayMode"], "block");
        assert_eq!(create["nativeEquation"], false);
        assert_eq!(create["powerPoint"]["slideId"], 256);
        assert_eq!(
            create["pendingMarker"],
            "visualtex:pending:v1:0f8fad5b-d9cb-469f-a165-70867728950e:7c9e6679-7425-40de-944b-e07fc1f90ae7"
        );
        assert!(create.get("fontSizePt").is_none());
        let mut edit = presentation_body("edit");
        edit.encoded_metadata = Some("visualtex:v1:deflate:abc".to_string());
        edit.fork_copied_formula = true;
        let edit = session_request(&edit).unwrap();
        assert!(edit.get("pendingMarker").is_none());
        assert_eq!(edit["encodedMetadata"], "visualtex:v1:deflate:abc");
        assert_eq!(edit["forkCopiedFormula"], true);

        let mut no_shape = presentation_body("create");
        no_shape.power_point = None;
        assert!(session_request(&no_shape).is_err());
        let mut no_id = presentation_body("create");
        no_id.formula_id = None;
        assert!(session_request(&no_id).is_err());
        let mut sized = presentation_body("create");
        sized.font_size_pt = Some(18.0);
        assert!(session_request(&sized).is_err());
        // A native equation names its text or slide, not a picture.
        let mut native = presentation_body("create");
        native.native_equation = true;
        assert!(session_request(&native).is_err());
        native.source_object_id = Some("visualtex-ppt-native-text:256:5:3:0".to_string());
        native.display_mode = "inline".to_string();
        let request = session_request(&native).unwrap();
        assert_eq!(request["nativeEquation"], true);
        assert_eq!(request["displayMode"], "inline");
        assert!(request.get("pendingMarker").is_none() && request.get("formulaId").is_none());
        let mut edit = native;
        edit.mode = "edit".to_string();
        assert!(session_request(&edit).is_err(), "an edit without its equation");
        edit.native_edit_xml = Some("<a14:m/>".to_string());
        assert_eq!(session_request(&edit).unwrap()["nativeEditXml"], "<a14:m/>");
        let mut picture_with_equation = presentation_body("edit");
        picture_with_equation.native_edit_xml = Some("<a14:m/>".to_string());
        assert!(session_request(&picture_with_equation).is_err());
        let mut writer_with_shape = body("create");
        writer_with_shape.power_point = Some(slide_shape());
        assert!(session_request(&writer_with_shape).is_err());
        let mut unknown = body("create");
        unknown.host = Some("excel".to_string());
        assert!(session_request(&unknown).is_err());
    }

    #[test]
    fn double_clicks_are_reported_once_and_only_while_fresh() {
        let mut log = EventLog::default();
        assert!(log.since(0).is_empty());
        log.push(json!({ "kind": "doubleClick", "x": 1.0, "y": 2.0 }));
        log.push(json!({ "kind": "doubleClick", "x": 3.0, "y": 4.0 }));
        let all = log.since(0);
        assert_eq!(all.len(), 2);
        assert_eq!(all[1]["sequence"], 2);
        assert_eq!(all[1]["x"], 3.0);
        assert_eq!(log.since(2).len(), 0);
        for index in 0..40 {
            log.push(json!({ "kind": "doubleClick", "x": index, "y": 0 }));
        }
        assert_eq!(log.recent.len(), MAX_PENDING_EVENTS);
        assert_eq!(log.since(0).len(), MAX_PENDING_EVENTS);
        if let Some(oldest) = log.recent.front_mut() {
            oldest.1 = Instant::now() - EVENT_LIFETIME;
        }
        assert_eq!(log.since(0).len(), MAX_PENDING_EVENTS - 1);
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
        let writer = ModuleQuery { host: None };
        let presentation = ModuleQuery { host: Some("wpp".to_string()) };
        let version = module_version(&writer);
        assert_eq!(version.len(), 16);
        assert!(version.chars().all(|c| c.is_ascii_hexdigit()));
        assert_eq!(version, module_version(&writer));
        assert_eq!(module_version(&presentation), short_hash(MODULE_WPP_JS));
        assert_ne!(module_version(&presentation), version);
        assert_eq!(module_source(&presentation), MODULE_WPP_JS);
        assert_eq!(module_source(&writer), MODULE_JS);
    }

    #[test]
    fn tokens_compare_exactly() {
        assert!(constant_time_equal("abc", "abc"));
        assert!(!constant_time_equal("abc", "abd"));
        assert!(!constant_time_equal("abc", "abcd"));
    }
}
