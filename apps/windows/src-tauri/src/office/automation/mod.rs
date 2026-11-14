//! Agent access to Word, PowerPoint and WPS (docs/office-mcp/DESIGN.md): the
//! automation channel the add-ins poll, and the MCP endpoint agents call.
pub mod audit;
pub mod mcp;
pub mod registry;
pub mod settings;

use crate::office::server::ServerContext;
use axum::extract::{Path as AxumPath, Query, State};
use axum::http::{HeaderMap, StatusCode};
use axum::response::{IntoResponse, Response};
use axum::routing::{delete, get, post};
use axum::{Json, Router};
use serde::Deserialize;
use std::time::Duration;

const HOST_SECRET_HEADER: &str = "x-visualtex-automation-secret";

/// Routes under `/api/v1` (install token, like every add-in call).
pub(crate) fn api_routes() -> Router<ServerContext> {
    Router::new()
        .route("/automation/hosts", post(register_host))
        .route("/automation/hosts/{host_id}", delete(unregister_host))
        .route("/automation/hosts/{host_id}/commands", get(poll_commands))
        .route(
            "/automation/hosts/{host_id}/commands/{command_id}/result",
            post(complete_command),
        )
}

fn error(status: StatusCode, message: &str) -> Response {
    (status, Json(serde_json::json!({ "error": message }))).into_response()
}

fn host_secret(headers: &HeaderMap) -> &str {
    headers
        .get(HOST_SECRET_HEADER)
        .and_then(|value| value.to_str().ok())
        .unwrap_or_default()
}

async fn register_host(
    State(_context): State<ServerContext>,
    Json(registration): Json<registry::HostRegistration>,
) -> Response {
    match registry::register(registration) {
        Ok((host_id, secret)) => (
            StatusCode::CREATED,
            Json(serde_json::json!({
                "hostId": host_id,
                "secret": secret,
                "pollWaitSeconds": registry::MAX_POLL_WAIT.as_secs(),
            })),
        )
            .into_response(),
        Err(message) => error(StatusCode::BAD_REQUEST, &message),
    }
}

async fn unregister_host(AxumPath(host_id): AxumPath<String>, headers: HeaderMap) -> Response {
    registry::unregister(&host_id, host_secret(&headers));
    StatusCode::NO_CONTENT.into_response()
}

#[derive(Debug, Default, Deserialize)]
struct PollQuery {
    wait: Option<u64>,
}

async fn poll_commands(
    AxumPath(host_id): AxumPath<String>,
    Query(query): Query<PollQuery>,
    headers: HeaderMap,
) -> Response {
    let wait = Duration::from_secs(query.wait.unwrap_or(registry::MAX_POLL_WAIT.as_secs()));
    match registry::poll(&host_id, host_secret(&headers), wait).await {
        Some(commands) => Json(serde_json::json!({ "commands": commands })).into_response(),
        // the add-in registers again (the companion restarted or dropped it)
        None => error(StatusCode::NOT_FOUND, "Unknown automation host"),
    }
}

async fn complete_command(
    AxumPath((host_id, command_id)): AxumPath<(String, String)>,
    headers: HeaderMap,
    Json(result): Json<registry::CommandResult>,
) -> Response {
    if registry::complete(&host_id, host_secret(&headers), &command_id, result) {
        StatusCode::NO_CONTENT.into_response()
    } else {
        // the caller already gave up (deadline, cancellation): nothing to do
        error(StatusCode::NOT_FOUND, "Unknown or finished automation command")
    }
}

// Settings → AI agents (DESIGN.md §10).

#[derive(Debug, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct McpStatus {
    pub settings: settings::McpSettings,
    pub token_path: String,
    pub launcher_path: String,
    pub hosts: Vec<registry::HostInfo>,
    pub clients: Vec<mcp::ClientInfo>,
}

fn launcher_path() -> String {
    std::env::current_exe()
        .ok()
        .and_then(|path| path.parent().map(|parent| parent.join("visualtex-mcp.exe")))
        .map(|path| path.display().to_string())
        .unwrap_or_default()
}

fn status(state: &crate::office::state::OfficeCompanionState) -> McpStatus {
    McpStatus {
        settings: settings::settings(&state.paths.root),
        token_path: settings::token_path(&state.paths.root).display().to_string(),
        launcher_path: launcher_path(),
        hosts: registry::hosts(),
        clients: mcp::clients(),
    }
}

#[tauri::command]
pub fn get_mcp_status(
    state: tauri::State<'_, crate::office::state::OfficeCompanionState>,
) -> McpStatus {
    status(&state)
}

#[tauri::command]
pub fn set_mcp_settings(
    state: tauri::State<'_, crate::office::state::OfficeCompanionState>,
    settings: settings::McpSettings,
) -> Result<McpStatus, String> {
    let turned_off = !settings.enabled;
    settings::update(&state.paths.root, settings)?;
    if turned_off {
        mcp::disconnect_all();
    }
    Ok(status(&state))
}

#[tauri::command]
pub fn reset_mcp_access(
    state: tauri::State<'_, crate::office::state::OfficeCompanionState>,
) -> Result<McpStatus, String> {
    settings::reset_token(&state.paths.root)?;
    mcp::disconnect_all();
    Ok(status(&state))
}

#[tauri::command]
pub fn get_mcp_audit(
    state: tauri::State<'_, crate::office::state::OfficeCompanionState>,
    limit: Option<usize>,
) -> Vec<audit::AuditEntry> {
    audit::recent(&state.paths, limit.unwrap_or(100).min(1000))
}
