//! OCR for the Office editor: the local runtime and the remote providers through the shared OCR state.

use super::*;

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct OcrStatusQuery {
    pub(super) force_refresh: Option<bool>,
}

#[derive(Debug, Default, Deserialize)]
pub(super) struct OcrEventsQuery {
    pub(super) cursor: Option<u64>,
    pub(super) event: Option<String>,
}

pub(super) fn ocr_error_response(error: String) -> Response {
    let status = if error.contains("OCR_CANCELLED") {
        StatusCode::CONFLICT
    } else if error.contains("not installed") || error.contains("Python executable is missing") {
        StatusCode::SERVICE_UNAVAILABLE
    } else {
        StatusCode::BAD_REQUEST
    };
    (status, Json(serde_json::json!({ "error": error }))).into_response()
}

pub(super) fn ocr_app(context: &ServerContext) -> Result<tauri::AppHandle, Box<Response>> {
    context.companion.app.clone().ok_or_else(|| {
        Box::new(
            (
                StatusCode::SERVICE_UNAVAILABLE,
                Json(serde_json::json!({
                    "error": "The VisualTeX OCR service is unavailable in this runtime"
                })),
            )
                .into_response(),
        )
    })
}

pub(super) async fn get_ocr_status(
    State(context): State<ServerContext>,
    Query(query): Query<OcrStatusQuery>,
) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context
        .companion
        .ocr
        .runtime_status(app, query.force_refresh.unwrap_or(false))
        .await
    {
        Ok(status) => Json(status).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn get_ocr_provider_configuration(
    State(context): State<ServerContext>,
) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.provider_configuration(&app) {
        Ok(configuration) => Json(configuration).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn put_ocr_provider_configuration(
    State(context): State<ServerContext>,
    Json(configuration): Json<crate::ocr::provider::OcrProviderConfigurationUpdate>,
) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context
        .companion
        .ocr
        .save_provider_configuration(&app, configuration)
    {
        Ok(configuration) => Json(configuration).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn install_ocr(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.install_runtime(app).await {
        Ok(status) => Json(status).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn get_ocr_install_status(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.install_status(app).await {
        Ok(status) => Json(status).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn cancel_ocr_install(State(context): State<ServerContext>) -> Response {
    match context.companion.ocr.cancel_install() {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn open_ocr_install_logs(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match crate::ocr::open_ocr_install_logs(app) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) fn ocr_header(headers: &HeaderMap, name: &'static str) -> Option<String> {
    headers
        .get(name)
        .and_then(|value| value.to_str().ok())
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(str::to_string)
}

pub(super) async fn recognize_ocr(
    State(context): State<ServerContext>,
    headers: HeaderMap,
    body: Bytes,
) -> Response {
    let content_type = headers
        .get(header::CONTENT_TYPE)
        .and_then(|value| value.to_str().ok())
        .unwrap_or_default();
    if !content_type
        .to_ascii_lowercase()
        .starts_with("application/octet-stream")
    {
        return (
            StatusCode::UNSUPPORTED_MEDIA_TYPE,
            Json(serde_json::json!({
                "error": "OCR recognition requires application/octet-stream"
            })),
        )
            .into_response();
    }
    if body.is_empty() {
        return (
            StatusCode::BAD_REQUEST,
            Json(serde_json::json!({ "error": "The OCR image is empty" })),
        )
            .into_response();
    }
    if body.len() > crate::ocr::MAX_IMAGE_BYTES {
        return (
            StatusCode::PAYLOAD_TOO_LARGE,
            Json(serde_json::json!({
                "error": "The OCR image is larger than the 20 MB limit"
            })),
        )
            .into_response();
    }

    let model = ocr_header(&headers, "x-visualtex-ocr-model")
        .unwrap_or_else(|| "PP-FormulaNet_plus-M".to_string());
    let extension =
        ocr_header(&headers, "x-visualtex-ocr-extension").unwrap_or_else(|| "png".to_string());
    if model.len() > 80 || extension.len() > 12 {
        return (
            StatusCode::BAD_REQUEST,
            Json(serde_json::json!({ "error": "Invalid OCR request headers" })),
        )
            .into_response();
    }

    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    let request = crate::ocr::OcrImageRequest {
        bytes: body.to_vec(),
        extension,
        model,
    };
    match context.companion.ocr.recognize(app, request).await {
        Ok(result) => Json(result).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct OcrCaptureRequest {
    #[serde(default)]
    pub(super) capture_mode: Option<String>,
}

/// Quick OCR from the Office editor: the editor window steps aside while the
/// user takes a screenshot (with the main app's chosen screenshot source),
/// then comes back to the front. 204 when the screenshot was cancelled.
#[cfg(target_os = "windows")]
pub(super) async fn capture_ocr_screenshot(
    State(context): State<ServerContext>,
    Json(request): Json<OcrCaptureRequest>,
) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    let requested = request
        .capture_mode
        .filter(|mode| !mode.trim().is_empty())
        .unwrap_or_else(crate::windows_silent_ocr_hotkey::configured_capture_mode);
    let capture_mode = match crate::windows_quick_ocr::normalize_capture_mode(&requested) {
        Ok(mode) => mode.to_string(),
        Err(error) => return ocr_error_response(error),
    };
    let editor = app.get_webview_window(OFFICE_EDITOR_WINDOW_LABEL);
    if let Some(window) = editor.as_ref() {
        let _ = window.minimize();
        // let WebView2/DWM leave the capture surface first (as the main app's quick OCR)
        tokio::time::sleep(std::time::Duration::from_millis(180)).await;
    }
    let capture = tauri::async_runtime::spawn_blocking(move || {
        crate::windows_quick_ocr::capture_windows_clipboard_image(&capture_mode, 60_000)
    })
    .await;
    tokio::time::sleep(std::time::Duration::from_millis(140)).await;
    if let Some(window) = editor.as_ref() {
        let _ = bring_session_window_to_front(window, None);
    }
    match capture {
        Err(error) => ocr_error_response(format!("Windows screenshot bridge task failed: {error}")),
        Ok(Err(error)) => ocr_error_response(error),
        Ok(Ok(None)) => StatusCode::NO_CONTENT.into_response(),
        Ok(Ok(Some(capture))) => Json(capture).into_response(),
    }
}

#[cfg(not(target_os = "windows"))]
pub(super) async fn capture_ocr_screenshot(
    State(_context): State<ServerContext>,
    Json(_request): Json<OcrCaptureRequest>,
) -> Response {
    (
        StatusCode::NOT_IMPLEMENTED,
        Json(serde_json::json!({ "error": "Quick OCR screenshots are available on Windows only" })),
    )
        .into_response()
}

pub(super) async fn cancel_ocr(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.cancel(&app) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn restart_ocr(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.restart(app).await {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn warmup_ocr(
    State(context): State<ServerContext>,
    headers: HeaderMap,
) -> Response {
    let model = ocr_header(&headers, "x-visualtex-ocr-model")
        .unwrap_or_else(|| "PP-FormulaNet_plus-M".to_string());
    if model.len() > 80 {
        return (
            StatusCode::BAD_REQUEST,
            Json(serde_json::json!({ "error": "Invalid OCR model" })),
        )
            .into_response();
    }
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.warmup_model(app, model).await {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn reset_ocr(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.reset_runtime(app).await {
        Ok(status) => Json(status).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

pub(super) async fn get_ocr_events(
    State(context): State<ServerContext>,
    Query(query): Query<OcrEventsQuery>,
) -> Response {
    if query.event.as_ref().is_some_and(|event| event.len() > 64) {
        return (
            StatusCode::BAD_REQUEST,
            Json(serde_json::json!({ "error": "Invalid OCR event filter" })),
        )
            .into_response();
    }
    Json(
        context
            .companion
            .ocr
            .poll_events(query.cursor.unwrap_or(u64::MAX), query.event.as_deref()),
    )
    .into_response()
}
