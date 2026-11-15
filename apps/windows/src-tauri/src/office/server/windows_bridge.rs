//! The Windows Office bridge: its requests and events, and committing a Word or WPS session.

use super::*;

#[derive(Debug, Default, Deserialize)]
pub(super) struct WindowsEventsQuery {
    pub(super) cursor: Option<u64>,
}

pub(super) fn windows_bridge_error(id: &str, error: String) -> Response {
    (
        StatusCode::SERVICE_UNAVAILABLE,
        Json(serde_json::json!({
            "protocolVersion": OFFICE_PROTOCOL_VERSION,
            "id": id,
            "ok": false,
            "error": {
                "code": "windows_office_bridge_unavailable",
                "message": error,
                "retryable": true
            }
        })),
    )
        .into_response()
}

pub(super) async fn get_office_platform_status(
    State(context): State<ServerContext>,
) -> Json<crate::office::platform::OfficePlatformStatus> {
    Json(context.companion.platform_backend.status())
}

pub(super) async fn windows_bridge_request(
    State(context): State<ServerContext>,
    Json(request): Json<serde_json::Value>,
) -> Response {
    let request_id = request
        .get("id")
        .and_then(serde_json::Value::as_str)
        .unwrap_or_default()
        .to_string();
    let backend = context.companion.platform_backend.clone();
    match tokio::task::spawn_blocking(move || backend.request(request)).await {
        Ok(Ok(result)) => Json(serde_json::json!({
            "protocolVersion": OFFICE_PROTOCOL_VERSION,
            "id": request_id,
            "ok": true,
            "result": result
        }))
        .into_response(),
        Ok(Err(error)) => windows_bridge_error(&request_id, error),
        Err(error) => windows_bridge_error(
            &request_id,
            format!("Windows Office bridge task failed: {error}"),
        ),
    }
}

pub(super) async fn get_windows_events(
    State(context): State<ServerContext>,
    Query(query): Query<WindowsEventsQuery>,
) -> Json<Vec<serde_json::Value>> {
    Json(
        context
            .companion
            .platform_backend
            .events_after(query.cursor.unwrap_or_default()),
    )
}

pub(super) fn decode_png_export(value: &str) -> Result<Vec<u8>, String> {
    let payload = value
        .split_once(',')
        .filter(|(prefix, _)| prefix.starts_with("data:image/png;base64"))
        .map(|(_, payload)| payload)
        .unwrap_or(value);
    let bytes = BASE64_STANDARD
        .decode(payload.trim())
        .map_err(|error| format!("Unable to decode Office PNG export: {error}"))?;
    if bytes.len() < 8 || &bytes[..8] != b"\x89PNG\r\n\x1a\n" {
        return Err("Office formula export is not a valid PNG image".to_string());
    }
    Ok(bytes)
}

pub(super) fn windows_office_temp_root() -> std::path::PathBuf {
    std::env::var_os("LOCALAPPDATA")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("office")
        .join("temp")
}

pub(super) fn commit_windows_session_blocking(
    companion: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    if !cfg!(target_os = "windows") {
        return Err("Windows Office session commits are available only on Windows".to_string());
    }
    let _commit_guard = WINDOWS_OFFICE_COMMIT_LOCK
        .lock()
        .map_err(|_| "Windows Office commit lock is unavailable".to_string())?;
    let session = companion
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())?;
    if session.status == OfficeSessionStatus::Completed {
        return Ok(session);
    }
    if session.status != OfficeSessionStatus::Committing {
        return Err("Windows Office Session is not ready to commit".to_string());
    }
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Windows Office Session has no exported formula image".to_string())?;
    let png = export
        .png_base64
        .as_deref()
        .ok_or_else(|| "Windows Office Session requires a PNG export".to_string())
        .and_then(decode_png_export)?;
    let temp_root = windows_office_temp_root();
    fs::create_dir_all(&temp_root)
        .map_err(|error| format!("Unable to create Windows Office temp directory: {error}"))?;
    let temporary = temp_root.join(format!("{session_id}.png"));
    fs::write(&temporary, png)
        .map_err(|error| format!("Unable to create Windows Office formula image: {error}"))?;

    let method = match (session.host, session.mode, session.display_mode.as_str()) {
        (OfficeHost::Powerpoint, OfficeSessionMode::Create, _) => "powerpoint.insertFormula",
        (OfficeHost::Powerpoint, OfficeSessionMode::Edit, _) => "powerpoint.replaceFormula",
        (OfficeHost::Word, OfficeSessionMode::Create, "inline") => "word.insertInlineFormula",
        (OfficeHost::Word, OfficeSessionMode::Create, _) => "word.insertDisplayFormula",
        (OfficeHost::Word, OfficeSessionMode::Edit, _) => "word.replaceFormula",
    };
    let metadata = metadata_from_session(&session);
    let request = serde_json::json!({
        "protocolVersion": OFFICE_PROTOCOL_VERSION,
        "id": Uuid::new_v4().to_string(),
        "method": method,
        "params": {
            "sessionId": &session.id,
            "formulaId": &session.formula_id,
            "imagePath": temporary.to_string_lossy(),
            "metadata": &metadata,
            "width": (export.width * 0.75).max(12.0),
            "height": (export.height * 0.75).max(12.0),
            "baseline": export.baseline.map(|value| (value * 0.75).max(0.0)),
            "sourceDocumentId": &session.source_document_id,
            "sourceObjectId": &session.source_object_id
        }
    });

    let result = companion.platform_backend.request(request);
    let _ = fs::remove_file(&temporary);
    result?;
    companion
        .formula_cache
        .put(&session.formula_id, metadata_from_session(&session))
        .map_err(|error| format!("Formula metadata could not be saved: {error}"))?;
    companion
        .session_store
        .patch(
            &session_id,
            serde_json::json!({ "status": "completed", "error": null }),
        )
        .map_err(|error| error.to_string())
}

pub(super) async fn commit_windows_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    if !valid_session_id(&session_id) {
        return StatusCode::NOT_FOUND.into_response();
    }
    let companion = context.companion.clone();
    let failure_store = context.companion.session_store.clone();
    let failure_id = session_id.clone();
    match tokio::task::spawn_blocking(move || {
        commit_windows_session_blocking(companion, session_id)
    })
    .await
    {
        Ok(Ok(session)) => Json(session).into_response(),
        Ok(Err(error)) => {
            let response_error = error.clone();
            let patch_id = failure_id.clone();
            let _ = run_session_operation(move || {
                failure_store.patch(
                    &patch_id,
                    serde_json::json!({ "status": "failed", "error": error }),
                )
            })
            .await;
            windows_bridge_error(&failure_id, response_error)
        }
        Err(error) => windows_bridge_error(
            &failure_id,
            format!("Windows Office commit task failed: {error}"),
        ),
    }
}
