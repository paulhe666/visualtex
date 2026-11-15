//! Rendering formulas for the add-ins: render batches, a single formula, and batch conversion.

use super::*;

#[derive(Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct BatchConversionRequest {
    pub(super) session_ids: Vec<String>,
}

pub(super) fn chunk_office_batch_conversion(session_ids: Vec<String>) -> Vec<Vec<String>> {
    session_ids
        .chunks(OFFICE_BATCH_CONVERSION_QUEUE_CHUNK_SIZE)
        .map(|chunk| chunk.to_vec())
        .collect()
}

pub(super) fn office_batch_conversion_queue() -> &'static Mutex<VecDeque<Vec<String>>> {
    OFFICE_BATCH_CONVERSION_QUEUE.get_or_init(|| Mutex::new(VecDeque::new()))
}

pub(super) async fn create_render_batch(
    State(context): State<ServerContext>,
    Json(request): Json<crate::office::render_batches::RenderBatchRequest>,
) -> Response {
    #[cfg(not(target_os = "windows"))]
    { let _ = (context, request); StatusCode::NOT_IMPLEMENTED.into_response() }
    #[cfg(target_os = "windows")]
    {
        let Some(app) = context.companion.app.clone() else {
            return desktop_app_unavailable();
        };
        let id = match crate::office::render_batches::create(request) {
            Ok(id) => id, Err(error) => return session_error_response(error),
        };
        let wake = tokio::task::spawn_blocking(move || {
            let window = ensure_desktop_conversion_window(&app)?;
            // The in-memory queue owns the work; the event is only a wake-up
            // (the window also polls, so a failed wake-up only costs time).
            if let Err(error) = window.eval(
                "window.dispatchEvent(new CustomEvent('visualtex-office-render-batch'));") {
                eprintln!("VisualTeX: render batch wake-up failed: {error}");
            }
            Ok::<(), String>(())
        }).await;
        let failure = match wake {
            Ok(Ok(())) => return (StatusCode::CREATED, Json(serde_json::json!({ "id": id }))).into_response(),
            Ok(Err(error)) => format!("Unable to open the VisualTeX conversion window: {error}"),
            Err(error) => format!("The task that opens the VisualTeX conversion window failed: {error}"),
        };
        let _ = crate::office::render_batches::release(&id);
        (StatusCode::INTERNAL_SERVER_ERROR, Json(serde_json::json!({ "error": failure }))).into_response()
    }
}

pub(super) async fn take_render_batch() -> Response {
    match crate::office::render_batches::take_next() {
        Ok(id) => Json(serde_json::json!({ "id": id })).into_response(),
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn get_render_batch(AxumPath(id): AxumPath<String>) -> Response {
    match crate::office::render_batches::snapshot(&id) {
        Ok(snapshot) => Json(snapshot).into_response(),
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn complete_render_batch(
    AxumPath(id): AxumPath<String>,
    Json(completion): Json<crate::office::render_batches::RenderBatchCompletion>,
) -> Response {
    match crate::office::render_batches::complete(&id, completion) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn wait_render_batch(AxumPath(id): AxumPath<String>) -> Response {
    let mut receiver = match crate::office::render_batches::subscribe(&id) {
        Ok(receiver) => receiver,
        Err(error) => return session_error_response(error),
    };
    // a batch the converter window never takes ends with a clear error
    // instead of the add-in's long timeout
    let watchdog_id = id.clone();
    let watchdog = tokio::spawn(async move {
        tokio::time::sleep(std::time::Duration::from_secs(RENDER_BATCH_START_SECONDS)).await;
        let _ = crate::office::render_batches::expire_untaken(&watchdog_id, RENDER_BATCH_START_SECONDS);
    });
    let wait = async {
        loop {
            if let Some(result) = receiver.borrow_and_update().clone() { return Ok(result); }
            receiver.changed().await.map_err(|_| SessionError::NotFound)?;
        }
    };
    let outcome = wait.await;
    watchdog.abort();
    match outcome {
        Ok(result) => Json(result).into_response(),
        Err(error) => session_error_response(error),
    }
}

#[derive(Debug, Deserialize)]
pub(super) struct RenderBatchFailure {
    #[serde(default)]
    pub(super) error: String,
}

/// The converter window could not render a batch (transport, decoding or a
/// rejected completion): the waiting add-in gets its reason at once.
pub(super) async fn fail_render_batch(
    AxumPath(id): AxumPath<String>,
    Json(failure): Json<RenderBatchFailure>,
) -> Response {
    match crate::office::render_batches::fail(&id, &failure.error) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn delete_render_batch(AxumPath(id): AxumPath<String>) -> Response {
    match crate::office::render_batches::release(&id) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => session_error_response(error),
    }
}

/// One formula through the hidden converter, for agents (`render_latex`, docs/office-mcp
/// DESIGN.md §4.5): the same render batch the add-ins use, created and awaited in-process.
/// Errors: (code, message); `latex_error` carries the renderer's message.
pub(crate) async fn render_single_formula(
    companion: &OfficeCompanionState,
    latex: &str,
    display: &str,
    font_size_pt: f64,
    limit: std::time::Duration,
) -> Result<crate::office::sessions::OfficeExportResult, (&'static str, String)> {
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (companion, latex, display, font_size_pt, limit);
        Err(("internal", "Rendering for agents is available on Windows only".to_string()))
    }
    #[cfg(target_os = "windows")]
    {
        use crate::office::render_batches::{self, RenderBatchInput, RenderBatchRequest};
        let Some(app) = companion.app.clone() else {
            return Err(("internal", "The VisualTeX converter is unavailable in this runtime".to_string()));
        };
        let input: CreateOfficeSessionInput = serde_json::from_value(serde_json::json!({
            "mode": "create", "host": "word", "codeFormat": "latex",
            "displayMode": display, "objectMode": "nativeOle", "fontSizePt": font_size_pt,
            "lines": [{ "id": Uuid::new_v4().to_string(), "latex": latex }],
        }))
        .map_err(|error| ("internal", error.to_string()))?;
        let id = Uuid::new_v4().to_string();
        render_batches::create(RenderBatchRequest {
            id: id.clone(),
            items: vec![RenderBatchInput { id: "formula".into(), request: input }],
        })
        .map_err(|error| ("latex_error", error.to_string()))?;
        let outcome = async {
            let mut receiver = render_batches::subscribe(&id).map_err(|error| ("internal", error.to_string()))?;
            tokio::task::spawn_blocking(move || {
                let window = ensure_desktop_conversion_window(&app)?;
                let _ = window.eval("window.dispatchEvent(new CustomEvent('visualtex-office-render-batch'));");
                Ok::<(), String>(())
            })
            .await
            .map_err(|error| ("internal", error.to_string()))?
            .map_err(|error| ("internal", format!("Unable to open the VisualTeX conversion window: {error}")))?;
            let watchdog_id = id.clone();
            let watchdog = tokio::spawn(async move {
                tokio::time::sleep(std::time::Duration::from_secs(RENDER_BATCH_START_SECONDS)).await;
                let _ = render_batches::expire_untaken(&watchdog_id, RENDER_BATCH_START_SECONDS);
            });
            let wait = async {
                loop {
                    if let Some(result) = receiver.borrow_and_update().clone() {
                        return Ok(result);
                    }
                    if receiver.changed().await.is_err() {
                        return Err(("internal", "The render batch ended without a result".to_string()));
                    }
                }
            };
            let result = tokio::time::timeout(limit, wait).await;
            watchdog.abort();
            result.map_err(|_| ("render_timeout", "VisualTeX did not finish rendering in time".to_string()))?
        }
        .await;
        let _ = render_batches::release(&id);
        let snapshot = outcome?;
        if let Some(error) = snapshot.error {
            return Err(("internal", error));
        }
        let item = snapshot
            .items
            .into_iter()
            .next()
            .ok_or(("internal", "The render batch returned no formula".to_string()))?;
        match (item.session.status, item.session.export_result) {
            (OfficeSessionStatus::Committing, Some(export)) => Ok(export),
            _ => Err((
                "latex_error",
                item.session.error.unwrap_or_else(|| "VisualTeX could not render this LaTeX".to_string()),
            )),
        }
    }
}

pub(super) async fn open_desktop_batch_conversion(
    State(context): State<ServerContext>,
    Json(request): Json<BatchConversionRequest>,
) -> Response {
    if request.session_ids.is_empty() {
        return (
            StatusCode::BAD_REQUEST,
            Json(serde_json::json!({
                "error": "Batch conversion requires at least one Session id"
            })),
        )
            .into_response();
    }
    if request
        .session_ids
        .iter()
        .any(|session_id| !valid_session_id(session_id))
    {
        return StatusCode::BAD_REQUEST.into_response();
    }

    let store = context.companion.session_store.clone();
    let lookup_ids = request.session_ids.clone();
    if let Err(error) = run_session_operation(move || {
        for session_id in lookup_ids {
            store.get(&session_id)?;
        }
        Ok(())
    })
    .await
    {
        return session_error_response(error);
    }

    #[cfg(not(target_os = "windows"))]
    {
        return (
            StatusCode::NOT_IMPLEMENTED,
            Json(serde_json::json!({
                "error": "Headless VSTO conversion is available only on Windows"
            })),
        )
            .into_response();
    }

    #[cfg(target_os = "windows")]
    {
        let Some(app) = context.companion.app.clone() else {
            return desktop_app_unavailable();
        };
        match tokio::task::spawn_blocking(move || {
            queue_desktop_batch_conversion(app, request.session_ids)
        })
        .await
        {
            Ok(Ok(())) => StatusCode::NO_CONTENT.into_response(),
            Ok(Err(error)) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({ "error": error })),
            )
                .into_response(),
            Err(error) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({
                    "error": format!("VisualTeX batch conversion task failed: {error}")
                })),
            )
                .into_response(),
        }
    }
}

pub(super) async fn take_desktop_batch_conversion() -> Response {
    match office_batch_conversion_queue().lock() {
        Ok(mut queue) => Json(BatchConversionRequest {
            session_ids: queue.pop_front().unwrap_or_default(),
        })
        .into_response(),
        Err(_) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(serde_json::json!({
                "error": "The VisualTeX batch conversion queue is unavailable"
            })),
        )
            .into_response(),
    }
}

pub(super) async fn open_desktop_conversion(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    if !valid_session_id(&session_id) {
        return StatusCode::NOT_FOUND.into_response();
    }
    let store = context.companion.session_store.clone();
    let lookup_id = session_id.clone();
    if let Err(error) = run_session_operation(move || store.get(&lookup_id)).await {
        return session_error_response(error);
    }

    #[cfg(not(target_os = "windows"))]
    {
        return (
            StatusCode::NOT_IMPLEMENTED,
            Json(serde_json::json!({
                "error": "Headless VSTO conversion is available only on Windows"
            })),
        )
            .into_response();
    }

    #[cfg(target_os = "windows")]
    {
        let Some(app) = context.companion.app.clone() else {
            return (
                StatusCode::SERVICE_UNAVAILABLE,
                Json(serde_json::json!({
                    "error": "The VisualTeX desktop application is unavailable"
                })),
            )
                .into_response();
        };
        match tokio::task::spawn_blocking(move || {
            open_desktop_conversion_window(app, session_id)
        })
        .await
        {
            Ok(Ok(())) => StatusCode::NO_CONTENT.into_response(),
            Ok(Err(error)) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({ "error": error })),
            )
                .into_response(),
            Err(error) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({
                    "error": format!("VisualTeX conversion window task failed: {error}")
                })),
            )
                .into_response(),
        }
    }
}
