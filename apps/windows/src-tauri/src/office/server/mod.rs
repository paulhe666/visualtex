//! The Office companion server the add-ins and the Office editor talk to: its constants, router, start and stop,
//! security and page handlers. Each area of the API is a child module.

use crate::office::powerpoint_native::{
    self, PowerPointInteractionEvent, PowerPointNativeSelection,
};
use crate::office::sessions::{
    CreateOfficeSessionInput, MetadataLine, OfficeFormulaSession, OfficeHost, OfficeSessionMode,
    OfficeSessionStatus, SessionError, VisualTeXFormulaMetadata,
};
use crate::office::state::{
    append_office_log, OfficeCompanionState, MAX_OFFICE_REQUEST_BYTES, OFFICE_BIND_IP,
    OFFICE_PROTOCOL_VERSION, OFFICE_UI_VERSION,
};
use crate::office::word_native;
use axum::body::Bytes;
use axum::extract::{Path as AxumPath, Query, Request, State};
use axum::http::{header, HeaderMap, HeaderValue, StatusCode};
use axum::middleware::{self, Next};
use axum::response::{Html, IntoResponse, Response};
use axum::routing::{get, post};
use axum::{Json, Router};
use axum_server::tls_rustls::RustlsConfig;
use serde::{Deserialize, Serialize};
use std::collections::VecDeque;
use std::fs;
use std::net::{SocketAddr, TcpListener};
use std::path::Path;
use std::sync::{Mutex, OnceLock};
use tower_http::limit::RequestBodyLimitLayer;
use tower_http::services::ServeDir;
use uuid::Uuid;

#[cfg(target_os = "windows")]
use std::ffi::c_void;
#[cfg(target_os = "windows")]
use tauri::webview::{PageLoadEvent, PageLoadPayload};
#[cfg(target_os = "windows")]
use tauri::{Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder};

mod editor_window;
mod ocr;
mod powerpoint;
mod rendering;
mod sessions;

pub(crate) use editor_window::*;
use ocr::*;
use powerpoint::*;
pub(crate) use rendering::*;
use sessions::*;

#[cfg(target_os = "windows")]
const SW_HIDE: i32 = 0;
#[cfg(target_os = "windows")]
const SW_SHOWNOACTIVATE: i32 = 4;
#[cfg(target_os = "windows")]
const SW_RESTORE: i32 = 9;

#[cfg(target_os = "windows")]
#[link(name = "user32")]
unsafe extern "system" {
    fn ShowWindow(hwnd: *mut c_void, command: i32) -> i32;
    fn IsWindowVisible(hwnd: *mut c_void) -> i32;
    fn UpdateWindow(hwnd: *mut c_void) -> i32;
    fn BringWindowToTop(hwnd: *mut c_void) -> i32;
    fn SetForegroundWindow(hwnd: *mut c_void) -> i32;
}

const INSTALL_TOKEN_HEADER: &str = "x-visualtex-install-token";
const OFFICE_CSP: &str = "default-src 'none'; script-src 'self' 'unsafe-eval'; style-src 'self' 'unsafe-inline'; connect-src 'self'; img-src 'self' data: blob:; font-src 'self' data:; object-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'self' https://*.office.com https://*.officeapps.live.com";
static POWERPOINT_COMMIT_LOCK: Mutex<()> = Mutex::new(());
static OFFICE_BATCH_CONVERSION_QUEUE: OnceLock<Mutex<VecDeque<Vec<String>>>> =
    OnceLock::new();
#[cfg(target_os = "windows")]
static OFFICE_EDITOR_ACTIVE_SESSION: OnceLock<Mutex<Option<String>>> = OnceLock::new();
#[cfg(target_os = "windows")]
static OFFICE_EDITOR_WINDOW_TRANSITION_LOCK: OnceLock<Mutex<()>> = OnceLock::new();
#[cfg(target_os = "windows")]
static OFFICE_EDITOR_PAGE_READY: std::sync::atomic::AtomicBool =
    std::sync::atomic::AtomicBool::new(false);
#[cfg(target_os = "windows")]
static OFFICE_EDITOR_SIZE_WRITE_GENERATION: std::sync::atomic::AtomicU64 =
    std::sync::atomic::AtomicU64::new(0);
#[cfg(target_os = "windows")]
const OFFICE_EDITOR_WINDOW_SIZE_FILE: &str = "editor-window-size.json";
#[cfg(target_os = "windows")]
const DEFAULT_OFFICE_EDITOR_WIDTH: f64 = 852.0;
#[cfg(target_os = "windows")]
const DEFAULT_OFFICE_EDITOR_HEIGHT: f64 = 500.57142857142856;
#[cfg(target_os = "windows")]
const MIN_OFFICE_EDITOR_WIDTH: f64 = 640.0;
#[cfg(target_os = "windows")]
const MIN_OFFICE_EDITOR_HEIGHT: f64 = 480.0;
#[cfg(target_os = "windows")]
const MAX_OFFICE_EDITOR_WIDTH: f64 = 2560.0;
#[cfg(target_os = "windows")]
const MAX_OFFICE_EDITOR_HEIGHT: f64 = 1600.0;

#[derive(Clone)]
pub(crate) struct ServerContext {
    pub(crate) companion: OfficeCompanionState,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct HealthResponse {
    ok: bool,
    app_version: &'static str,
    office_ui_version: &'static str,
    protocol_version: u32,
    ocr_available: bool,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ThemeResponse {
    theme: String,
    editor_layout: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct OfficePreferencesResponse {
    powerpoint_default_font_size_pt: f64,
    editor_preferences: serde_json::Value,
}

const OFFICE_BATCH_CONVERSION_QUEUE_CHUNK_SIZE: usize = 256;

fn constant_time_eq(left: &[u8], right: &[u8]) -> bool {
    if left.len() != right.len() {
        return false;
    }
    left.iter()
        .zip(right)
        .fold(0_u8, |difference, (a, b)| difference | (a ^ b))
        == 0
}

fn valid_session_id(value: &str) -> bool {
    (16..=128).contains(&value.len())
        && value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_'))
}

/// The page metadata the Office dialog reads. The install token is not among it: the
/// HTML is served to anyone on this computer, so VisualTeX's own windows receive the
/// token through `office_install_token_script` (docs/office-mcp/DESIGN.md §9.2).
fn inject_dialog_meta(
    html: &str,
    theme: &str,
    editor_layout: &str,
) -> Result<String, StatusCode> {
    let marker = "</head>";
    let index = html.find(marker).ok_or(StatusCode::INTERNAL_SERVER_ERROR)?;
    let native_powerpoint_commit = if cfg!(target_os = "macos") {
        "true"
    } else {
        "false"
    };
    let meta = format!(
        "<meta name=\"visualtex-native-powerpoint-commit\" content=\"{native_powerpoint_commit}\" />\n<meta name=\"visualtex-theme\" content=\"{theme}\" />\n<meta name=\"visualtex-editor-layout\" content=\"{editor_layout}\" />\n"
    );
    Ok(format!("{}{}{}", &html[..index], meta, &html[index..]))
}

async fn read_office_html(
    ui_root: &Path,
    relative: &str,
    theme: &str,
    editor_layout: &str,
) -> Result<Html<String>, StatusCode> {
    let path = ui_root.join(relative);
    let html = tokio::fs::read_to_string(&path)
        .await
        .map_err(|_| StatusCode::NOT_FOUND)?;
    inject_dialog_meta(&html, theme, editor_layout).map(Html)
}

/// The install token for the Office pages VisualTeX itself opens: defined before the
/// page's scripts run, and only on the companion's own origin.
#[cfg(target_os = "windows")]
fn office_install_token_script(app: &tauri::AppHandle) -> String {
    let token = app
        .try_state::<OfficeCompanionState>()
        .map(|state| state.install_token.to_string())
        .unwrap_or_default();
    let origin = format!("https://127.0.0.1:{}", crate::office::state::office_port());
    format!(
        "(function(){{if(window.location.origin==={origin:?}){{Object.defineProperty(window,'__VISUALTEX_INSTALL_TOKEN__',{{value:{token:?},writable:false,configurable:false}});}}}})();"
    )
}

async fn health(State(context): State<ServerContext>) -> Json<HealthResponse> {
    Json(HealthResponse {
        ok: true,
        app_version: env!("CARGO_PKG_VERSION"),
        office_ui_version: OFFICE_UI_VERSION,
        protocol_version: OFFICE_PROTOCOL_VERSION,
        ocr_available: context.companion.ocr_available,
    })
}

#[derive(Debug, Default, Deserialize)]
struct DialogRuntimeQuery {
    runtime: Option<String>,
}

async fn dialog(
    AxumPath(session_id): AxumPath<String>,
    Query(query): Query<DialogRuntimeQuery>,
    State(context): State<ServerContext>,
) -> Result<Html<String>, StatusCode> {
    if !valid_session_id(&session_id) {
        return Err(StatusCode::NOT_FOUND);
    }
    let store = context.companion.session_store.clone();
    let lookup_id = session_id.clone();
    run_session_operation(move || store.get(&lookup_id))
        .await
        .map_err(|_| StatusCode::NOT_FOUND)?;
    let _ = query.runtime;
    let theme = context.companion.current_theme();
    let editor_layout = context.companion.current_editor_layout();
    read_office_html(
        &context.companion.paths.ui_root,
        "dialog/index.html",
        &theme,
        &editor_layout,
    )
    .await
}

async fn dialog_shell(
    State(context): State<ServerContext>,
) -> Result<Html<String>, StatusCode> {
    let theme = context.companion.current_theme();
    let editor_layout = context.companion.current_editor_layout();
    read_office_html(
        &context.companion.paths.ui_root,
        "dialog/index.html",
        &theme,
        &editor_layout,
    )
    .await
}

async fn api_theme(State(context): State<ServerContext>) -> Json<ThemeResponse> {
    Json(ThemeResponse {
        theme: context.companion.current_theme(),
        editor_layout: context.companion.current_editor_layout(),
    })
}

async fn api_preferences(
    State(context): State<ServerContext>,
) -> Json<OfficePreferencesResponse> {
    Json(OfficePreferencesResponse {
        powerpoint_default_font_size_pt: context
            .companion
            .powerpoint_default_font_size_pt(),
        editor_preferences: context.companion.current_editor_preferences(),
    })
}

async fn api_status(
    State(context): State<ServerContext>,
) -> Json<crate::office::state::OfficeCompanionStatus> {
    Json(context.companion.snapshot())
}

async fn get_office_platform_status(
    State(context): State<ServerContext>,
) -> Json<crate::office::platform::OfficePlatformStatus> {
    Json(context.companion.platform_backend.status())
}

#[cfg(target_os = "windows")]
const OFFICE_EDITOR_WINDOW_LABEL: &str = "office-session-editor";
#[cfg(target_os = "windows")]
const OFFICE_CONVERTER_WINDOW_LABEL: &str = "office-session-converter";

/// Seconds a render batch may wait for the converter window to take it.
const RENDER_BATCH_START_SECONDS: u64 = 20;

async fn api_not_found() -> impl IntoResponse {
    (
        StatusCode::NOT_FOUND,
        Json(serde_json::json!({ "error": "Unknown VisualTeX Office API route" })),
    )
}

async fn not_found() -> impl IntoResponse {
    StatusCode::NOT_FOUND
}

fn add_security_headers(mut response: Response, request_path: &str) -> Response {
    let headers = response.headers_mut();
    headers.insert(
        header::X_CONTENT_TYPE_OPTIONS,
        HeaderValue::from_static("nosniff"),
    );
    headers.insert(
        header::REFERRER_POLICY,
        HeaderValue::from_static("no-referrer"),
    );
    headers.insert(
        header::CONTENT_SECURITY_POLICY,
        HeaderValue::from_static(OFFICE_CSP),
    );
    headers.insert(
        header::HeaderName::from_static("permissions-policy"),
        HeaderValue::from_static("camera=(), microphone=(), geolocation=()"),
    );
    headers.insert(
        header::HeaderName::from_static("cross-origin-resource-policy"),
        HeaderValue::from_static("same-origin"),
    );
    if request_path == "/health"
        || request_path.starts_with("/api/")
        || request_path.ends_with(".html")
        || request_path == "/dialog"
        || request_path.starts_with("/dialog/")
    {
        headers.insert(header::CACHE_CONTROL, HeaderValue::from_static("no-store"));
    }
    response
}

async fn security_and_auth(
    State(context): State<ServerContext>,
    request: Request,
    next: Next,
) -> Response {
    let request_path = request.uri().path().to_string();
    if request_path.starts_with("/api/v1/") {
        let supplied = request
            .headers()
            .get(INSTALL_TOKEN_HEADER)
            .and_then(|value| value.to_str().ok())
            .unwrap_or_default();
        if !constant_time_eq(
            supplied.as_bytes(),
            context.companion.install_token.as_bytes(),
        ) {
            return add_security_headers(
                (
                    StatusCode::UNAUTHORIZED,
                    Json(serde_json::json!({ "error": "Invalid VisualTeX install token" })),
                )
                    .into_response(),
                &request_path,
            );
        }
    }

    add_security_headers(next.run(request).await, &request_path)
}

pub(crate) fn build_router(companion: OfficeCompanionState) -> Router {
    let context = ServerContext {
        companion: companion.clone(),
    };
    let ui_root = companion.paths.ui_root.clone();
    let api = Router::new()
        .route("/status", get(api_status))
        .route("/theme", get(api_theme))
        .route("/preferences", get(api_preferences))
        .route("/platform/status", get(get_office_platform_status))
        .route("/app/reveal", post(reveal_desktop_app))
        .route(
            "/app/sessions/{session_id}/open",
            post(open_desktop_session),
        )
        .route(
            "/app/editor/prewarm",
            post(prewarm_desktop_editor),
        )
        .route(
            "/app/converter/prewarm",
            post(prewarm_desktop_conversion),
        )
        .route(
            "/app/converter/convert-batch",
            post(open_desktop_batch_conversion),
        )
        .route(
            "/app/converter/next-batch",
            get(take_desktop_batch_conversion),
        )
        .route(
            "/app/sessions/{session_id}/convert",
            post(open_desktop_conversion),
        )
        .route(
            "/app/sessions/{session_id}/bulk-import",
            post(open_desktop_bulk_import),
        )
        .route(
            "/app/sessions/{session_id}/close",
            post(close_desktop_session),
        )
        .route("/app/converter/render-batches", post(create_render_batch))
        .route("/app/converter/next-render-batch", get(take_render_batch))
        .route("/app/converter/render-batches/{id}",
            get(get_render_batch).delete(delete_render_batch))
        .route("/app/converter/render-batches/{id}/complete", post(complete_render_batch))
        .route("/app/converter/render-batches/{id}/fail", post(fail_render_batch))
        .route("/app/converter/render-batches/{id}/result", get(wait_render_batch))
        .route("/sessions", post(create_session))
        .route("/sessions/batch-delete", post(delete_sessions_batch))
        .route(
            "/sessions/{session_id}",
            get(get_session).patch(patch_session).delete(delete_session),
        )
        .route(
            "/formulas/{formula_id}/metadata",
            get(get_formula_metadata).put(put_formula_metadata),
        )
        .route("/word/inline-baseline", post(apply_word_inline_baseline))
        .route(
            "/powerpoint/selection",
            get(get_powerpoint_native_selection),
        )
        .route(
            "/powerpoint/selection/mark",
            post(mark_powerpoint_native_selection),
        )
        .route(
            "/powerpoint/slide/snapshot",
            get(get_powerpoint_native_slide_snapshot),
        )
        .route(
            "/powerpoint/shape/mark-last",
            post(mark_last_powerpoint_native_formula),
        )
        .route(
            "/powerpoint/shape/replace-last",
            post(replace_last_powerpoint_native_formula),
        )
        .route(
            "/powerpoint/shape/delete",
            post(delete_powerpoint_native_shape),
        )
        .route(
            "/powerpoint/sessions/{session_id}/commit",
            post(commit_powerpoint_session),
        )
        .route(
            "/powerpoint/sessions/{session_id}/confirm",
            post(confirm_powerpoint_session),
        )
        .route("/powerpoint/events", get(get_powerpoint_events))
        .route("/ocr/status", get(get_ocr_status))
        .route(
            "/ocr/providers",
            get(get_ocr_provider_configuration).put(put_ocr_provider_configuration),
        )
        .route("/ocr/install", post(install_ocr))
        .route("/ocr/install/status", get(get_ocr_install_status))
        .route("/ocr/install/cancel", post(cancel_ocr_install))
        .route("/ocr/install/logs/open", post(open_ocr_install_logs))
        .route("/ocr/recognize", post(recognize_ocr))
        .route("/ocr/cancel", post(cancel_ocr))
        .route("/ocr/capture", post(capture_ocr_screenshot))
        .route("/ocr/restart", post(restart_ocr))
        .route("/ocr/warmup", post(warmup_ocr))
        .route("/ocr/reset", post(reset_ocr))
        .route("/ocr/events", get(get_ocr_events))
        .merge(crate::office::automation::api_routes())
        .fallback(api_not_found);

    Router::new()
        .route("/health", get(health))
        .route("/dialog", get(dialog_shell))
        .route("/dialog/{session_id}", get(dialog))
        .route(
            "/mcp",
            post(crate::office::automation::mcp::post)
                .get(crate::office::automation::mcp::get)
                .delete(crate::office::automation::mcp::delete),
        )
        .nest("/api/v1", api)
        .nest_service("/assets", ServeDir::new(ui_root.join("assets")))
        .fallback(not_found)
        .layer(RequestBodyLimitLayer::new(MAX_OFFICE_REQUEST_BYTES))
        // axum's extractors (Json, Bytes) keep their own 2 MB default unless told:
        // a converter's completion of 20 OLE formulas can exceed it
        .layer(axum::extract::DefaultBodyLimit::max(MAX_OFFICE_REQUEST_BYTES))
        .layer(middleware::from_fn_with_state(
            context.clone(),
            security_and_auth,
        ))
        .with_state(context)
}

/// Why a port could not be bound. Windows answers "in use" (WSAEADDRINUSE) or "access
/// denied" (WSAEACCES) both for a port another program listens on and for one inside a
/// reserved range, which no port listing shows.
fn describe_bind_error(error: &std::io::Error) -> String {
    let unavailable = matches!(
        error.kind(),
        std::io::ErrorKind::AddrInUse | std::io::ErrorKind::PermissionDenied
    );
    if unavailable {
        format!("in use by another program or reserved by Windows (WSL, Hyper-V): {error}")
    } else {
        error.to_string()
    }
}

/// Binds the first of `ports` that can be bound (0: any free port). Returns the
/// listener with the reasons the ports before it were skipped, or every reason when
/// none could be bound.
fn bind_first_available(
    ip: [u8; 4],
    ports: &[u16],
) -> Result<(TcpListener, Vec<String>), String> {
    let mut failures = Vec::new();
    for &port in ports {
        match TcpListener::bind(SocketAddr::from((ip, port))) {
            Ok(listener) => return Ok((listener, failures)),
            Err(error) => {
                let name = if port == 0 {
                    "any free port".to_string()
                } else {
                    format!("port {port}")
                };
                failures.push(format!("{name} is {}", describe_bind_error(&error)));
            }
        }
    }
    Err(format!(
        "VisualTeX Office companion cannot listen on {}: {}",
        std::net::Ipv4Addr::from(ip),
        failures.join("; ")
    ))
}

/// The port this process's companion had before a restart. Its stopped listener can
/// take a moment to close, so the port gets that moment before another one is chosen.
async fn rebind_own_port(ip: [u8; 4], port: u16, attempts: u32) -> Option<TcpListener> {
    for attempt in 0..attempts {
        if attempt > 0 {
            tokio::time::sleep(std::time::Duration::from_millis(100)).await;
        }
        if let Ok(listener) = TcpListener::bind(SocketAddr::from((ip, port))) {
            return Some(listener);
        }
    }
    None
}

/// Makes the bound port known: to this process (URLs of the Office editor pages and
/// their command permissions, which are per origin) and, on Windows, to the add-ins
/// and the MCP launcher through the registry.
fn announce_listening_port(companion: &OfficeCompanionState, port: u16) -> Result<(), String> {
    crate::office::state::set_active_office_port(port);
    companion.update_status(|status| status.port = port);
    if let Some(app) = companion.app.as_ref() {
        crate::office::state::allow_office_origin(app, port)?;
    }
    #[cfg(target_os = "windows")]
    if !crate::office::state::isolated_acceptance_runtime() {
        // Without the registry value the clients look at the standard port, which is
        // right unless this is another one: then they need it.
        if let Err(error) = crate::office::windows_backend::publish_companion_port(port) {
            append_office_log(
                &companion.paths,
                "companion.log",
                &format!("listening port {port} was not published: {error}"),
            );
        }
    }
    Ok(())
}

pub async fn run(companion: OfficeCompanionState) -> Result<(), String> {
    let candidates = crate::office::state::office_port_candidates();
    let executable = std::env::current_exe()
        .map(|path| path.display().to_string())
        .unwrap_or_else(|error| format!("<unresolved: {error}>"));
    append_office_log(
        &companion.paths,
        "companion.log",
        &format!(
            "server start pid={} executable={} data_root={} ui_root={} certificate={} private_key={} bind={}",
            std::process::id(),
            executable,
            companion.paths.root.display(),
            companion.paths.ui_root.display(),
            companion.paths.certificate.display(),
            companion.paths.private_key.display(),
            std::net::Ipv4Addr::from(OFFICE_BIND_IP)
        ),
    );

    let dialog_path = companion.paths.ui_root.join("dialog").join("index.html");
    if !dialog_path.is_file() {
        let error = format!(
            "Office UI resources are missing: {}",
            dialog_path.display()
        );
        append_office_log(&companion.paths, "companion.log", &error);
        return Err(error);
    }
    append_office_log(
        &companion.paths,
        "companion.log",
        &format!("Office UI resource validated: {}", dialog_path.display()),
    );

    append_office_log(
        &companion.paths,
        "companion.log",
        &format!("binding TCP listener, ports in order: {candidates:?}"),
    );
    let own_port = match crate::office::state::bound_office_port() {
        Some(port) => rebind_own_port(OFFICE_BIND_IP, port, 25).await,
        None => None,
    };
    let (listener, skipped) = match own_port {
        Some(listener) => (listener, Vec::new()),
        None => bind_first_available(OFFICE_BIND_IP, &candidates).map_err(|error| {
            append_office_log(&companion.paths, "companion.log", &format!("bind failed: {error}"));
            error
        })?,
    };
    for failure in &skipped {
        append_office_log(&companion.paths, "companion.log", &format!("bind skipped: {failure}"));
    }
    let address = listener.local_addr().map_err(|error| {
        let error = format!("Unable to read the Office companion listener address: {error}");
        append_office_log(&companion.paths, "companion.log", &error);
        error
    })?;
    append_office_log(
        &companion.paths,
        "companion.log",
        &format!("TCP listener bound at {address}"),
    );
    listener.set_nonblocking(true).map_err(|error| {
        let error = format!("Unable to configure Office companion listener: {error}");
        append_office_log(&companion.paths, "companion.log", &error);
        error
    })?;

    append_office_log(
        &companion.paths,
        "companion.log",
        "initializing TLS certificate and private key",
    );
    let tls = RustlsConfig::from_pem_file(
        &companion.paths.certificate,
        &companion.paths.private_key,
    )
    .await
    .map_err(|error| {
        let error = format!("Unable to load Office TLS certificate: {error}");
        append_office_log(&companion.paths, "companion.log", &error);
        error
    })?;
    append_office_log(
        &companion.paths,
        "companion.log",
        "TLS certificate and private key loaded successfully",
    );

    let handle = axum_server::Handle::new();
    {
        let mut stored = companion.server_handle.lock().map_err(|_| {
            let error = "Office companion server handle lock is unavailable".to_string();
            append_office_log(&companion.paths, "companion.log", &error);
            error
        })?;
        *stored = Some(handle.clone());
    }
    announce_listening_port(&companion, address.port()).map_err(|error| {
        append_office_log(&companion.paths, "companion.log", &error);
        error
    })?;
    companion.update_status(|status| {
        status.running = true;
        status.last_error = None;
    });
    companion.starting.store(false, std::sync::atomic::Ordering::SeqCst);
    append_office_log(
        &companion.paths,
        "companion.log",
        &format!(
            "HTTPS companion listening at https://{}; protocol={}; ui_version={}",
            address, OFFICE_PROTOCOL_VERSION, OFFICE_UI_VERSION
        ),
    );

    let server = axum_server::from_tcp_rustls(listener, tls)
        .map_err(|error| {
            let error = format!("Unable to create Office TLS server: {error}");
            append_office_log(&companion.paths, "companion.log", &error);
            error
        })?
        .handle(handle)
        .serve(build_router(companion.clone()).into_make_service());

    let result = server
        .await
        .map_err(|error| format!("VisualTeX Office companion server failed: {error}"));

    companion.update_status(|status| {
        status.running = false;
        if let Err(error) = &result {
            status.last_error = Some(error.clone());
        }
    });
    if let Ok(mut stored) = companion.server_handle.lock() {
        *stored = None;
    }
    match &result {
        Ok(()) => append_office_log(
            &companion.paths,
            "companion.log",
            "companion stopped normally",
        ),
        Err(error) => append_office_log(
            &companion.paths,
            "companion.log",
            &format!("companion stopped with error: {error}"),
        ),
    }
    result
}

pub fn stop(companion: &OfficeCompanionState) -> Result<(), String> {
    let handle = companion
        .server_handle
        .lock()
        .map_err(|_| "Office companion server handle lock is unavailable".to_string())?
        .clone();
    if let Some(handle) = handle {
        handle.graceful_shutdown(Some(std::time::Duration::from_secs(2)));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::office::sessions::SessionStore;
    use axum::body::Body;
    use axum::http::Request as HttpRequest;
    use http_body_util::BodyExt;
    use std::fs;
    use tempfile::TempDir;
    use tower::ServiceExt;

    #[test]
    fn batch_conversion_chunking_does_not_cap_total_sessions_at_256() {
        let session_ids = (0..1000)
            .map(|index| format!("session-{index:04}"))
            .collect::<Vec<_>>();
        let chunks = chunk_office_batch_conversion(session_ids.clone());
        assert_eq!(chunks.iter().map(Vec::len).collect::<Vec<_>>(), vec![256, 256, 256, 232]);
        assert_eq!(chunks.into_iter().flatten().collect::<Vec<_>>(), session_ids);
    }

    fn test_state(temp: &TempDir) -> OfficeCompanionState {
        let root = temp.path().join("office-data");
        let ui_root = temp.path().join("office-ui");
        fs::create_dir_all(ui_root.join("dialog")).unwrap();
        fs::create_dir_all(ui_root.join("assets")).unwrap();
        fs::write(
            ui_root.join("dialog").join("index.html"),
            concat!(
                "<html><head>",
                "<script type=\"module\" src=\"/assets/dialog-entry.js\"></script>",
                "<link rel=\"stylesheet\" href=\"/assets/dialog.css\">",
                "</head><body>dialog</body></html>"
            ),
        )
        .unwrap();
        fs::write(ui_root.join("assets").join("test.js"), "ok").unwrap();
        let paths = crate::office::state::OfficePaths {
            certificate: root.join("localhost-cert.pem"),
            private_key: root.join("localhost-key.pem"),
            certificate_metadata: root.join("certificate.json"),
            install: root.join("install.json"),
            sessions: root.join("sessions"),
            recovery: root.join("recovery"),
            formula_cache: root.join("formulas"),
            ui_root,
            root,
        };
        let session_store = SessionStore::new(&paths).expect("session store");
        let formula_cache = crate::office::formula_cache::FormulaMetadataCache::new(&paths)
            .expect("formula metadata cache");
        OfficeCompanionState::new(
            None,
            crate::ocr::OcrState::default(),
            paths,
            "a".repeat(64),
            session_store,
            formula_cache,
            true,
        )
    }

    fn close_policy_session(
        object_mode: &str,
        dirty: bool,
        math_ml: Option<&str>,
        png_base64: Option<&str>,
    ) -> OfficeFormulaSession {
        OfficeFormulaSession {
            id: Uuid::new_v4().to_string(),
            mode: OfficeSessionMode::Edit,
            host: OfficeHost::Word,
            formula_id: Uuid::new_v4().to_string(),
            source_document_id: Some(Uuid::new_v4().to_string()),
            source_object_id: Some("0:1".to_string()),
            title: "Formula".to_string(),
            lines: vec![crate::office::sessions::FormulaLine {
                id: Uuid::new_v4().to_string(),
                latex: "x+y".to_string(),
            }],
            active_line_id: None,
            code_format: "latex".to_string(),
            display_mode: "inline".to_string(),
            object_mode: object_mode.to_string(),
            numbered: false,
            math_type_number_position: "right".to_string(),
            font_size_pt: 11.0,
            export_width: 120.0,
            export_height: 30.0,
            export_result: Some(crate::office::sessions::OfficeExportResult {
                svg: "<svg viewBox=\"0 0 120 30\"></svg>".to_string(),
                svg_base64: "PHN2Zz48L3N2Zz4=".to_string(),
                math_ml: math_ml.map(str::to_string),
                png_base64: png_base64.map(str::to_string),
                width: 120.0,
                height: 30.0,
                baseline: Some(22.0),
                formula_letter_font: Some("times".to_string()),
                formula_chinese_font: Some("kaiti".to_string()),
                formula_faces: None,
            }),
            original_metadata: None,
            dirty,
            status: OfficeSessionStatus::Editing,
            auto_commit_on_close: true,
            explicit_cancel: false,
            error: None,
            created_at: 1,
            updated_at: 1,
            expires_at: u64::MAX,
        }
    }

    #[test]
    fn dirty_word_omml_close_commits_without_png() {
        let session = close_policy_session(
            "wordOmml",
            true,
            Some("<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi></math>"),
            None,
        );
        assert!(office_session_export_is_committable(&session));
        assert_eq!(
            close_patch_for_office_editor(&session)
                .and_then(|patch| patch.get("status").cloned()),
            Some(serde_json::json!("committing")),
        );
    }

    #[test]
    fn dirty_native_ole_close_without_png_cancels_instead_of_stranding() {
        let session = close_policy_session("nativeOle", true, None, None);
        assert!(!office_session_export_is_committable(&session));
        let patch = close_patch_for_office_editor(&session).expect("close patch");
        assert_eq!(patch.get("status"), Some(&serde_json::json!("cancelled")));
        assert_eq!(patch.get("explicitCancel"), Some(&serde_json::json!(true)));
    }

    #[test]
    fn unchanged_edit_close_commits_without_export() {
        let mut session = close_policy_session("wordOmml", false, None, None);
        session.export_result = None;
        assert!(office_session_export_is_committable(&session));
        assert_eq!(
            close_patch_for_office_editor(&session)
                .and_then(|patch| patch.get("status").cloned()),
            Some(serde_json::json!("committing")),
        );
    }

    #[test]
    fn powerpoint_mathtype_close_requires_mathml_not_an_image() {
        let mut session = close_policy_session("mathTypeOle", true, None, None);
        session.host = OfficeHost::Powerpoint;
        assert!(!office_session_export_is_committable(&session));
        session.export_result.as_mut().unwrap().math_ml = Some("<math><mi>x</mi></math>".into());
        assert!(office_session_export_is_committable(&session));
        session.export_result.as_mut().unwrap().svg.clear();
        session.export_result.as_mut().unwrap().svg_base64.clear();
        assert!(office_session_export_is_committable(&session));
    }

    #[test]
    fn dialog_shell_disables_html_caching() {
        let response = add_security_headers(StatusCode::OK.into_response(), "/dialog");
        assert_eq!(
            response.headers().get(header::CACHE_CONTROL),
            Some(&HeaderValue::from_static("no-store")),
        );
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn editor_page_load_url_resolves_only_valid_session_routes() {
        let session_id = "550e8400-e29b-41d4-a716-446655440000";
        let valid = tauri::Url::parse(&format!(
            "https://127.0.0.1:43127/dialog/{session_id}?runtime=vsto-desktop"
        ))
        .expect("valid editor URL");
        assert_eq!(
            office_editor_session_id_from_url(&valid).as_deref(),
            Some(session_id),
        );

        let query = tauri::Url::parse(&format!(
            "https://127.0.0.1:43127/dialog?runtime=vsto-desktop&sessionId={session_id}"
        ))
        .expect("valid query editor URL");
        assert_eq!(
            office_editor_session_id_from_url(&query).as_deref(),
            Some(session_id),
        );

        let shell = tauri::Url::parse("https://127.0.0.1:43127/dialog?runtime=vsto-desktop")
            .expect("valid shell URL");
        assert_eq!(office_editor_session_id_from_url(&shell), None);

        let invalid = tauri::Url::parse(
            "https://127.0.0.1:43127/dialog/not-a-session?runtime=vsto-desktop&sessionId=also-invalid",
        )
        .expect("valid URL with invalid Session id");
        assert_eq!(office_editor_session_id_from_url(&invalid), None);
    }

    #[test]
    fn session_ids_reject_path_traversal() {
        assert!(valid_session_id("550e8400-e29b-41d4-a716-446655440000"));
        assert!(!valid_session_id("../../etc/passwd"));
        assert!(!valid_session_id("short"));
        assert!(!valid_session_id("session%2Fescape"));
    }

    #[test]
    fn powerpoint_formula_minimum_size_preserves_svg_aspect_ratio() {
        let (width, height) = scale_powerpoint_formula_size(200.0, 10.0).expect("valid SVG bounds");
        assert!((height - 12.0).abs() < 0.000_001);
        assert!((width / height - 20.0).abs() < 0.000_001);
    }

    #[test]
    fn powerpoint_formula_maximum_size_preserves_svg_aspect_ratio() {
        let (width, height) =
            scale_powerpoint_formula_size(1600.0, 80.0).expect("valid SVG bounds");
        assert!((width - 600.0).abs() < 0.000_001);
        assert!((width / height - 20.0).abs() < 0.000_001);
    }

    #[test]
    fn powerpoint_formula_size_rejects_invalid_exports() {
        assert!(scale_powerpoint_formula_size(0.0, 10.0).is_err());
        assert!(scale_powerpoint_formula_size(f64::NAN, 10.0).is_err());
    }

    #[test]
    fn install_token_comparison_is_length_sensitive() {
        assert!(constant_time_eq(b"same-token", b"same-token"));
        assert!(!constant_time_eq(b"same-token", b"same-tokeN"));
        assert!(!constant_time_eq(b"same-token", b"same-token-longer"));
    }

    #[test]
    fn native_edit_reference_preserves_the_exact_selected_shape() {
        let decoded = decode_powerpoint_native_edit_reference(
            "visualtex-ppt-native-edit:12:47726170686963203131",
        )
        .expect("native edit reference");
        assert_eq!(decoded, (12, "Graphic 11".to_string()));
        assert!(decode_powerpoint_native_edit_reference(
            "visualtex-ppt-native-edit:0:47726170686963"
        )
        .is_none());
        assert!(
            decode_powerpoint_native_edit_reference("visualtex-ppt-native-edit:1:xyz").is_none()
        );
    }

    #[test]
    fn fixed_listener_is_loopback_only() {
        let address = OfficeCompanionState::socket_addr();
        assert!(address.ip().is_loopback());
        assert_eq!(address.port(), 43_127);
    }

    fn occupied_port() -> (TcpListener, u16) {
        let listener = TcpListener::bind(("127.0.0.1", 0)).expect("occupying listener");
        let port = listener.local_addr().expect("local address").port();
        (listener, port)
    }

    #[test]
    fn free_first_port_is_bound_without_trying_the_others() {
        let (first, free) = occupied_port();
        drop(first);
        let (listener, skipped) =
            bind_first_available(OFFICE_BIND_IP, &[free, 0]).expect("free port binds");
        assert_eq!(listener.local_addr().unwrap().port(), free);
        assert!(skipped.is_empty());
    }

    #[test]
    fn unavailable_ports_fall_back_to_the_next_candidate() {
        let (_first, first_port) = occupied_port();
        let (_second, second_port) = occupied_port();
        let (third, third_port) = occupied_port();
        drop(third);
        let (listener, skipped) =
            bind_first_available(OFFICE_BIND_IP, &[first_port, second_port, third_port, 0])
                .expect("third candidate binds");
        let address = listener.local_addr().unwrap();
        assert!(address.ip().is_loopback());
        assert_eq!(address.port(), third_port);
        assert_eq!(skipped.len(), 2);
        assert!(skipped[0].starts_with(&format!("port {first_port} is in use by another program or reserved by Windows")));
        assert!(skipped[1].starts_with(&format!("port {second_port} is ")));
    }

    #[test]
    fn last_candidate_takes_any_free_port() {
        let (_first, first_port) = occupied_port();
        let (listener, skipped) =
            bind_first_available(OFFICE_BIND_IP, &[first_port, 0]).expect("a free port binds");
        let port = listener.local_addr().unwrap().port();
        assert_ne!(port, 0);
        assert_ne!(port, first_port);
        assert_eq!(skipped.len(), 1);
    }

    #[tokio::test]
    async fn restart_waits_for_its_own_port_to_be_released() {
        let (occupying, port) = occupied_port();
        let release = tokio::spawn(async move {
            tokio::time::sleep(std::time::Duration::from_millis(250)).await;
            drop(occupying);
        });
        let listener = rebind_own_port(OFFICE_BIND_IP, port, 25).await.expect("port is released");
        assert_eq!(listener.local_addr().unwrap().port(), port);
        release.await.unwrap();

        // a port that stays taken is given up, so that another one can be chosen
        let (_held, held_port) = occupied_port();
        assert!(rebind_own_port(OFFICE_BIND_IP, held_port, 2).await.is_none());
    }

    #[test]
    fn no_bindable_port_names_every_port_and_reason() {
        let (_first, first_port) = occupied_port();
        let (_second, second_port) = occupied_port();
        let error = bind_first_available(OFFICE_BIND_IP, &[first_port, second_port])
            .expect_err("no candidate can be bound");
        assert!(error.starts_with("VisualTeX Office companion cannot listen on 127.0.0.1: "));
        assert!(error.contains(&format!("port {first_port} is in use by another program or reserved by Windows")));
        assert!(error.contains(&format!("port {second_port} is in use by another program or reserved by Windows")));
    }

    #[test]
    fn reserved_and_occupied_ports_read_the_same() {
        let mut errors = vec![
            std::io::Error::from(std::io::ErrorKind::AddrInUse),
            std::io::Error::from(std::io::ErrorKind::PermissionDenied),
        ];
        if cfg!(target_os = "windows") {
            // WSAEACCES and WSAEADDRINUSE, as bind returns them
            errors.push(std::io::Error::from_raw_os_error(10013));
            errors.push(std::io::Error::from_raw_os_error(10048));
        }
        for error in errors {
            let text = describe_bind_error(&error);
            assert!(text.starts_with("in use by another program or reserved by Windows (WSL, Hyper-V): "), "{text}");
        }
        let other = describe_bind_error(&std::io::Error::other("network is down"));
        assert_eq!(other, "network is down");
    }

    #[test]
    fn published_editor_preferences_are_kept_for_the_next_run() {
        let temp = TempDir::new().expect("temp dir");
        let state = test_state(&temp);
        assert_eq!(state.current_editor_preferences(), serde_json::json!({}));
        assert_eq!(state.published_editor_preferences(), serde_json::json!({}));
        let published = serde_json::json!({
            "settings": { "language": "en", "formulaFaces": { "letters": "times" }, "zoom": 2 },
            "commandUsage": { "alpha": 1 },
        });
        assert_eq!(state.set_current_editor_preferences(published.clone()), published);
        assert_eq!(state.published_editor_preferences(), published);

        // the next run serves them to the Office editor before the main window publishes
        let restarted = test_state(&temp);
        assert_eq!(
            restarted.current_editor_preferences(),
            serde_json::json!({ "settings": { "language": "en", "formulaFaces": { "letters": "times" } } }),
        );
        // but has sent nothing itself yet: its first publication is compared with nothing
        assert_eq!(restarted.published_editor_preferences(), serde_json::json!({}));
    }

    #[tokio::test]
    async fn router_serves_health_html_and_security_headers() {
        let temp = TempDir::new().expect("temp dir");
        let state = test_state(&temp);
        let router = build_router(state.clone());

        let health = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri("/health")
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(health.status(), StatusCode::OK);
        assert_eq!(
            health
                .headers()
                .get(header::X_CONTENT_TYPE_OPTIONS)
                .unwrap(),
            "nosniff"
        );
        assert!(health
            .headers()
            .contains_key(header::CONTENT_SECURITY_POLICY));
        let health_body = health.into_body().collect().await.unwrap().to_bytes();
        let health_json: serde_json::Value = serde_json::from_slice(&health_body).unwrap();
        assert_eq!(health_json["ok"], true);
        assert_eq!(health_json["protocolVersion"], OFFICE_PROTOCOL_VERSION);

        let bridge = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri("/bridge/index.html")
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(bridge.status(), StatusCode::NOT_FOUND);

        let invalid_dialog = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri("/dialog/..%2F..%2Fetc%2Fpasswd")
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(invalid_dialog.status(), StatusCode::NOT_FOUND);

        let unknown = router
            .oneshot(
                HttpRequest::builder()
                    .uri("/unknown-resource")
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(unknown.status(), StatusCode::NOT_FOUND);
    }

    #[tokio::test]
    async fn api_requires_install_token_and_rejects_oversized_requests() {
        let temp = TempDir::new().expect("temp dir");
        let state = test_state(&temp);
        let router = build_router(state.clone());

        let unauthorized = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri("/api/v1/status")
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(unauthorized.status(), StatusCode::UNAUTHORIZED);
        assert_eq!(
            unauthorized
                .headers()
                .get(header::X_CONTENT_TYPE_OPTIONS)
                .unwrap(),
            "nosniff"
        );

        let authorized = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri("/api/v1/status")
                    .header(INSTALL_TOKEN_HEADER, state.install_token.as_str())
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(authorized.status(), StatusCode::OK);

        let oversized = router
            .oneshot(
                HttpRequest::builder()
                    .method("POST")
                    .uri("/api/v1/status")
                    .header(INSTALL_TOKEN_HEADER, state.install_token.as_str())
                    .header(header::CONTENT_LENGTH, MAX_OFFICE_REQUEST_BYTES + 1)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(oversized.status(), StatusCode::PAYLOAD_TOO_LARGE);
    }

    async fn mcp_call(
        router: &Router,
        token: &str,
        session: Option<&str>,
        body: serde_json::Value,
    ) -> (StatusCode, Option<String>, serde_json::Value) {
        let mut request = HttpRequest::builder()
            .method("POST")
            .uri("/mcp")
            .header(header::AUTHORIZATION, format!("Bearer {token}"))
            .header(header::CONTENT_TYPE, "application/json");
        if let Some(session) = session {
            request = request.header("mcp-session-id", session);
        }
        let response = router
            .clone()
            .oneshot(request.body(Body::from(body.to_string())).unwrap())
            .await
            .unwrap();
        let status = response.status();
        let session = response
            .headers()
            .get("mcp-session-id")
            .map(|value| value.to_str().unwrap().to_string());
        let bytes = response.into_body().collect().await.unwrap().to_bytes();
        (status, session, serde_json::from_slice(&bytes).unwrap_or(serde_json::Value::Null))
    }

    #[tokio::test]
    async fn mcp_endpoint_is_off_by_default_and_needs_its_own_token() {
        use crate::office::automation::settings;
        let temp = TempDir::new().expect("temp dir");
        let state = test_state(&temp);
        let router = build_router(state.clone());
        let initialize = serde_json::json!({ "jsonrpc": "2.0", "id": 1, "method": "initialize",
            "params": { "protocolVersion": "2025-06-18", "capabilities": {},
                "clientInfo": { "name": "test-client", "version": "1" } } });

        // off: refused even with the install token
        let (status, _, _) = mcp_call(&router, state.install_token.as_str(), None, initialize.clone()).await;
        assert_eq!(status, StatusCode::FORBIDDEN);

        settings::update(&state.paths.root, settings::McpSettings {
            enabled: true,
            mode: settings::McpMode::Allow,
        })
        .unwrap();
        let token = fs::read_to_string(settings::token_path(&state.paths.root)).unwrap();
        let (status, _, _) = mcp_call(&router, state.install_token.as_str(), None, initialize.clone()).await;
        assert_eq!(status, StatusCode::UNAUTHORIZED);

        let (status, session, reply) = mcp_call(&router, &token, None, initialize).await;
        assert_eq!(status, StatusCode::OK);
        assert_eq!(reply["result"]["protocolVersion"], "2025-06-18");
        assert_eq!(reply["result"]["serverInfo"]["name"], "visualtex");
        let session = session.expect("session id");

        let initialized = serde_json::json!({ "jsonrpc": "2.0", "method": "notifications/initialized" });
        let (status, _, _) = mcp_call(&router, &token, Some(&session), initialized).await;
        assert_eq!(status, StatusCode::ACCEPTED);

        let list = serde_json::json!({ "jsonrpc": "2.0", "id": 2, "method": "tools/list" });
        let (status, _, reply) = mcp_call(&router, &token, Some(&session), list.clone()).await;
        assert_eq!(status, StatusCode::OK);
        assert_eq!(
            reply["result"]["tools"].as_array().unwrap().len(),
            8 + crate::office::automation::mcp::WRITE_TOOLS.len()
        );

        let call = serde_json::json!({ "jsonrpc": "2.0", "id": 3, "method": "tools/call",
            "params": { "name": "visualtex_status", "arguments": {} } });
        let (_, _, reply) = mcp_call(&router, &token, Some(&session), call).await;
        assert_eq!(reply["result"]["isError"], false);
        assert_eq!(reply["result"]["structuredContent"]["permission_mode"], "allow_changes");

        // writes need a connected Office host
        let call = serde_json::json!({ "jsonrpc": "2.0", "id": 6, "method": "tools/call",
            "params": { "name": "insert_formula", "arguments": { "document_id": "h0:x", "latex": "x" } } });
        let (_, _, reply) = mcp_call(&router, &token, Some(&session), call).await;
        assert_eq!(reply["result"]["isError"], true);
        assert!(reply["result"]["content"][0]["text"].as_str().unwrap().contains("document_not_found"));
        settings::update(&state.paths.root, settings::McpSettings {
            enabled: true,
            mode: settings::McpMode::ReadOnly,
        })
        .unwrap();
        let call = serde_json::json!({ "jsonrpc": "2.0", "id": 7, "method": "tools/call",
            "params": { "name": "delete_formula", "arguments": { "document_id": "h0:x", "formula_id": "f" } } });
        let (_, _, reply) = mcp_call(&router, &token, Some(&session), call).await;
        assert!(reply["result"]["content"][0]["text"].as_str().unwrap().contains("permission_denied"));
        settings::update(&state.paths.root, settings::McpSettings {
            enabled: true,
            mode: settings::McpMode::Allow,
        })
        .unwrap();

        let call = serde_json::json!({ "jsonrpc": "2.0", "id": 4, "method": "tools/call",
            "params": { "name": "read_formula", "arguments": { "document_id": "h0:x", "formula_id": "f" } } });
        let (_, _, reply) = mcp_call(&router, &token, Some(&session), call).await;
        assert_eq!(reply["result"]["isError"], true);
        assert!(reply["result"]["content"][0]["text"].as_str().unwrap().contains("document_not_found"));

        let unknown = serde_json::json!({ "jsonrpc": "2.0", "id": 5, "method": "tools/call",
            "params": { "name": "format_disk" } });
        let (_, _, reply) = mcp_call(&router, &token, Some(&session), unknown).await;
        assert_eq!(reply["error"]["code"], -32602);

        let (status, _, _) = mcp_call(&router, &token, Some("stale"), list.clone()).await;
        assert_eq!(status, StatusCode::NOT_FOUND);

        // a page in a browser cannot call it
        let response = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("POST")
                    .uri("/mcp")
                    .header(header::AUTHORIZATION, format!("Bearer {token}"))
                    .header(header::ORIGIN, "https://example.com")
                    .body(Body::from(list.to_string()))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(response.status(), StatusCode::FORBIDDEN);
    }

    #[tokio::test]
    async fn automation_channel_requires_install_token_and_host_secret() {
        let temp = TempDir::new().expect("temp dir");
        let state = test_state(&temp);
        let router = build_router(state.clone());
        let registration = serde_json::json!({ "app": "word", "vendor": "microsoft", "pid": 91_200 });
        let request = |token: &str| {
            HttpRequest::builder()
                .method("POST")
                .uri("/api/v1/automation/hosts")
                .header(INSTALL_TOKEN_HEADER, token)
                .header(header::CONTENT_TYPE, "application/json")
                .body(Body::from(registration.to_string()))
                .unwrap()
        };
        let refused = router.clone().oneshot(request("wrong")).await.unwrap();
        assert_eq!(refused.status(), StatusCode::UNAUTHORIZED);
        let created = router.clone().oneshot(request(state.install_token.as_str())).await.unwrap();
        assert_eq!(created.status(), StatusCode::CREATED);
        let body: serde_json::Value = serde_json::from_slice(
            &created.into_body().collect().await.unwrap().to_bytes(),
        )
        .unwrap();
        let host_id = body["hostId"].as_str().unwrap().to_string();
        let secret = body["secret"].as_str().unwrap().to_string();

        let poll = |secret: &str| {
            HttpRequest::builder()
                .uri(format!("/api/v1/automation/hosts/{host_id}/commands?wait=0"))
                .header(INSTALL_TOKEN_HEADER, state.install_token.as_str())
                .header("x-visualtex-automation-secret", secret)
                .body(Body::empty())
                .unwrap()
        };
        let wrong = router.clone().oneshot(poll("wrong")).await.unwrap();
        assert_eq!(wrong.status(), StatusCode::NOT_FOUND);
        let polled = router.clone().oneshot(poll(&secret)).await.unwrap();
        assert_eq!(polled.status(), StatusCode::OK);
        let body: serde_json::Value = serde_json::from_slice(
            &polled.into_body().collect().await.unwrap().to_bytes(),
        )
        .unwrap();
        assert_eq!(body["commands"], serde_json::json!([]));
    }

    #[tokio::test]
    async fn formula_metadata_cache_api_round_trips_and_validates_identity() {
        let temp = TempDir::new().expect("temp dir");
        let state = test_state(&temp);
        let token = state.install_token.to_string();
        let router = build_router(state);
        let formula_id = uuid::Uuid::new_v4().to_string();
        let line_id = uuid::Uuid::new_v4().to_string();
        let metadata = serde_json::json!({
            "schema": "visualtex-formula",
            "schemaVersion": 1,
            "formulaId": formula_id,
            "title": "Cached Formula",
            "latex": "a=b",
            "lines": [{ "id": line_id, "latex": "a=b" }],
            "codeFormat": "raw",
            "displayMode": "inline",
            "createdWithVersion": "1.0.6",
            "updatedWithVersion": "1.0.6",
            "createdAt": "2026-07-12T00:00:00Z",
            "updatedAt": "2026-07-12T00:00:00Z"
        });

        let put = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("PUT")
                    .uri(format!("/api/v1/formulas/{formula_id}/metadata"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/json")
                    .body(Body::from(metadata.to_string()))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(put.status(), StatusCode::OK);

        let get = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri(format!("/api/v1/formulas/{formula_id}/metadata"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(get.status(), StatusCode::OK);
        let body = get.into_body().collect().await.unwrap().to_bytes();
        let loaded: serde_json::Value = serde_json::from_slice(&body).unwrap();
        assert_eq!(loaded["formulaId"], formula_id);
        assert_eq!(loaded["latex"], "a=b");

        let mismatched_id = uuid::Uuid::new_v4().to_string();
        let conflict = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("PUT")
                    .uri(format!("/api/v1/formulas/{mismatched_id}/metadata"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/json")
                    .body(Body::from(metadata.to_string()))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(conflict.status(), StatusCode::CONFLICT);

        let unauthorized = router
            .oneshot(
                HttpRequest::builder()
                    .uri(format!("/api/v1/formulas/{formula_id}/metadata"))
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(unauthorized.status(), StatusCode::UNAUTHORIZED);
    }

    #[tokio::test]
    async fn session_api_supports_create_edit_cancel_and_delete() {
        let temp = TempDir::new().expect("temp dir");
        let state = test_state(&temp);
        let token = state.install_token.to_string();
        let router = build_router(state);

        let create = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("POST")
                    .uri("/api/v1/sessions")
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/json")
                    .body(Body::from(
                        serde_json::json!({
                            "mode": "create",
                            "host": "word",
                            "title": "API Formula",
                            "autoCommitOnClose": true
                        })
                        .to_string(),
                    ))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(create.status(), StatusCode::CREATED);
        let create_body = create.into_body().collect().await.unwrap().to_bytes();
        let created: serde_json::Value = serde_json::from_slice(&create_body).unwrap();
        let session_id = created["id"].as_str().unwrap().to_string();
        let line_id = created["lines"][0]["id"].as_str().unwrap().to_string();

        let dialog = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri(format!("/dialog/{session_id}"))
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(dialog.status(), StatusCode::OK);
        let dialog_body = dialog.into_body().collect().await.unwrap().to_bytes();
        let dialog_html = String::from_utf8(dialog_body.to_vec()).unwrap();
        assert!(!dialog_html.contains("office.js"));
        assert!(!dialog_html.contains("visualtex-install-token"));
        assert!(!dialog_html.contains(token.as_str()));

        let desktop_dialog = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri(format!(
                        "/dialog/{session_id}?runtime=vsto-desktop"
                    ))
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(desktop_dialog.status(), StatusCode::OK);
        let desktop_body = desktop_dialog
            .into_body()
            .collect()
            .await
            .unwrap()
            .to_bytes();
        let desktop_html = String::from_utf8(desktop_body.to_vec()).unwrap();
        assert!(!desktop_html.contains("office.js"));
        assert!(!desktop_html.contains(token.as_str()));

        let patch = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("PATCH")
                    .uri(format!("/api/v1/sessions/{session_id}"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/json")
                    .body(Body::from(
                        serde_json::json!({
                            "title": "Updated Formula",
                            "lines": [{ "id": line_id, "latex": "a=b" }],
                            "dirty": true,
                            "status": "editing"
                        })
                        .to_string(),
                    ))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(patch.status(), StatusCode::OK);

        let get = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri(format!("/api/v1/sessions/{session_id}"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(get.status(), StatusCode::OK);
        let get_body = get.into_body().collect().await.unwrap().to_bytes();
        let loaded: serde_json::Value = serde_json::from_slice(&get_body).unwrap();
        assert_eq!(loaded["title"], "Updated Formula");
        assert_eq!(loaded["lines"][0]["latex"], "a=b");

        let cancel = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("PATCH")
                    .uri(format!("/api/v1/sessions/{session_id}"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/json")
                    .body(Body::from(
                        serde_json::json!({
                            "status": "cancelled",
                            "explicitCancel": true
                        })
                        .to_string(),
                    ))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(cancel.status(), StatusCode::OK);

        let invalid_commit = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("PATCH")
                    .uri(format!("/api/v1/sessions/{session_id}"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/json")
                    .body(Body::from(
                        serde_json::json!({
                            "status": "committing",
                            "explicitCancel": false,
                            "exportResult": {
                                "svg": "<svg viewBox=\"0 0 10 10\"></svg>",
                                "svgBase64": "PHN2Zz48L3N2Zz4=",
                                "pngBase64": null,
                                "width": 10,
                                "height": 10,
                                "baseline": 8
                            }
                        })
                        .to_string(),
                    ))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(invalid_commit.status(), StatusCode::CONFLICT);

        let delete = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("DELETE")
                    .uri(format!("/api/v1/sessions/{session_id}"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(delete.status(), StatusCode::NO_CONTENT);

        let missing = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri(format!("/api/v1/sessions/{session_id}"))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(missing.status(), StatusCode::NOT_FOUND);

        let missing_dialog = router
            .oneshot(
                HttpRequest::builder()
                    .uri(format!("/dialog/{session_id}"))
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(missing_dialog.status(), StatusCode::NOT_FOUND);
    }

    #[tokio::test]
    async fn ocr_api_requires_token_and_validates_image_requests() {
        let temp = TempDir::new().expect("temp dir");
        let state = test_state(&temp);
        let token = state.install_token.to_string();
        let router = build_router(state);

        let unauthorized = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri("/api/v1/ocr/status")
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(unauthorized.status(), StatusCode::UNAUTHORIZED);

        let unavailable = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri("/api/v1/ocr/status?forceRefresh=true")
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(unavailable.status(), StatusCode::SERVICE_UNAVAILABLE);

        let events = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .uri("/api/v1/ocr/events?event=ocr-recognition-progress")
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(events.status(), StatusCode::OK);
        let events_body = events.into_body().collect().await.unwrap().to_bytes();
        let events_json: serde_json::Value = serde_json::from_slice(&events_body).unwrap();
        assert_eq!(events_json["cursor"], 0);
        assert_eq!(events_json["events"], serde_json::json!([]));

        let wrong_media = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("POST")
                    .uri("/api/v1/ocr/recognize")
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/json")
                    .body(Body::from("{}"))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(wrong_media.status(), StatusCode::UNSUPPORTED_MEDIA_TYPE);

        let empty = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("POST")
                    .uri("/api/v1/ocr/recognize")
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/octet-stream")
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(empty.status(), StatusCode::BAD_REQUEST);

        let oversized = router
            .clone()
            .oneshot(
                HttpRequest::builder()
                    .method("POST")
                    .uri("/api/v1/ocr/recognize")
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .header(header::CONTENT_TYPE, "application/octet-stream")
                    .body(Body::from(vec![0_u8; crate::ocr::MAX_IMAGE_BYTES + 1]))
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(oversized.status(), StatusCode::PAYLOAD_TOO_LARGE);

        let invalid_filter = router
            .oneshot(
                HttpRequest::builder()
                    .uri(format!("/api/v1/ocr/events?event={}", "x".repeat(65)))
                    .header(INSTALL_TOKEN_HEADER, &token)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(invalid_filter.status(), StatusCode::BAD_REQUEST);
    }
}
