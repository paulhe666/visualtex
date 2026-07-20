use crate::office::powerpoint_native::{
    self, PowerPointInteractionEvent, PowerPointNativeSelection,
};
use crate::office::sessions::{
    CreateOfficeSessionInput, MetadataLine, OfficeFormulaSession, OfficeHost, OfficeSessionMode,
    OfficeSessionStatus, SessionError, VisualTeXFormulaMetadata,
};
use crate::office::state::{
    OfficeCompanionState, MAX_OFFICE_REQUEST_BYTES, OFFICE_PROTOCOL_VERSION, OFFICE_UI_VERSION,
};
use crate::office::word_native;
use axum::body::Bytes;
use axum::extract::{Path as AxumPath, Query, Request, State};
use axum::http::{header, HeaderMap, HeaderValue, StatusCode};
use axum::middleware::{self, Next};
use axum::response::{IntoResponse, Response};
use axum::routing::{get, post};
use axum::{Json, Router};
use axum_server::tls_rustls::RustlsConfig;
use serde::{Deserialize, Serialize};
use std::fs;
use std::net::TcpListener;
use std::sync::Mutex;
use tower_http::limit::RequestBodyLimitLayer;
use uuid::Uuid;

const INSTALL_TOKEN_HEADER: &str = "x-visualtex-install-token";
const OFFICE_CSP: &str = "default-src 'none'; connect-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'";
static POWERPOINT_COMMIT_LOCK: Mutex<()> = Mutex::new(());

#[derive(Clone)]
struct ServerContext {
    companion: OfficeCompanionState,
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
struct PreparedPowerPointCommit {
    session: OfficeFormulaSession,
    selection: PowerPointNativeSelection,
}

fn constant_time_eq(left: &[u8], right: &[u8]) -> bool {
    if left.len() != right.len() {
        return false;
    }
    left.iter()
        .zip(right)
        .fold(0_u8, |difference, (a, b)| difference | (a ^ b))
        == 0
}

/// Convert the MathJax SVG pixel bounds to PowerPoint points without ever
/// changing their aspect ratio.  Width and height used to be clamped to 12pt
/// independently; a short formula would therefore acquire an artificial
/// height while retaining its natural width, and PowerPoint compressed the
/// glyphs into that distorted box on every replacement.
fn scale_powerpoint_formula_size(width_px: f64, height_px: f64) -> Result<(f64, f64), String> {
    if !width_px.is_finite() || !height_px.is_finite() || width_px <= 0.0 || height_px <= 0.0 {
        return Err("PowerPoint formula export has invalid dimensions".to_string());
    }

    let natural_width = width_px * 0.75;
    let natural_height = height_px * 0.75;
    let maximum_scale = f64::min(600.0 / natural_width, 400.0 / natural_height);
    let preferred_scale = f64::min(1.0, maximum_scale);
    let minimum_scale = f64::max(12.0 / natural_width, 12.0 / natural_height);
    let scale = f64::min(maximum_scale, f64::max(preferred_scale, minimum_scale));

    Ok((natural_width * scale, natural_height * scale))
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

async fn api_status(
    State(context): State<ServerContext>,
) -> Json<crate::office::state::OfficeCompanionStatus> {
    Json(context.companion.snapshot())
}

async fn reveal_desktop_app(State(context): State<ServerContext>) -> Response {
    let Some(app) = context.companion.app.clone() else {
        return (
            StatusCode::SERVICE_UNAVAILABLE,
            Json(serde_json::json!({
                "error": "The VisualTeX desktop window is unavailable in this runtime"
            })),
        )
            .into_response();
    };

    match crate::office::background::reveal_main_window(&app) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => (
            StatusCode::INTERNAL_SERVER_ERROR,
            Json(serde_json::json!({ "error": error })),
        )
            .into_response(),
    }
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OcrStatusQuery {
    force_refresh: Option<bool>,
}

#[derive(Debug, Default, Deserialize)]
struct OcrEventsQuery {
    cursor: Option<u64>,
    event: Option<String>,
}

fn ocr_error_response(error: String) -> Response {
    let status = if error.contains("OCR_CANCELLED") {
        StatusCode::CONFLICT
    } else if error.contains("not installed") || error.contains("Python executable is missing") {
        StatusCode::SERVICE_UNAVAILABLE
    } else {
        StatusCode::BAD_REQUEST
    };
    (status, Json(serde_json::json!({ "error": error }))).into_response()
}

fn ocr_app(context: &ServerContext) -> Result<tauri::AppHandle, Box<Response>> {
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

async fn get_ocr_status(
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

async fn install_ocr(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.install_runtime(app).await {
        Ok(status) => Json(status).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

fn ocr_header(headers: &HeaderMap, name: &'static str) -> Option<String> {
    headers
        .get(name)
        .and_then(|value| value.to_str().ok())
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(str::to_string)
}

async fn recognize_ocr(
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
    if body.len() > crate::MAX_IMAGE_BYTES {
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
    let request = crate::OcrImageRequest {
        bytes: body.to_vec(),
        extension,
        model,
    };
    match context.companion.ocr.recognize(app, request).await {
        Ok(result) => Json(result).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

async fn prewarm_ocr(
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
    match context.companion.ocr.prewarm_model(app, model).await {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => ocr_error_response(error),
    }
}

async fn cancel_ocr(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.cancel(&app) {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => ocr_error_response(error),
    }
}

async fn restart_ocr(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.restart(app).await {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => ocr_error_response(error),
    }
}

async fn reset_ocr(State(context): State<ServerContext>) -> Response {
    let app = match ocr_app(&context) {
        Ok(app) => app,
        Err(response) => return *response,
    };
    match context.companion.ocr.reset_runtime(app).await {
        Ok(status) => Json(status).into_response(),
        Err(error) => ocr_error_response(error),
    }
}

async fn get_ocr_events(
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

async fn run_native_operation<T>(
    operation: impl FnOnce() -> Result<T, String> + Send + 'static,
) -> Result<T, String>
where
    T: Send + 'static,
{
    tokio::task::spawn_blocking(operation)
        .await
        .map_err(|error| format!("PowerPoint native task failed: {error}"))?
}

fn native_error_response(error: String) -> Response {
    (
        StatusCode::BAD_REQUEST,
        Json(serde_json::json!({ "error": error })),
    )
        .into_response()
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct MarkPowerPointFormulaRequest {
    formula_id: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ApplyWordInlineBaselineRequest {
    position: i32,
    formula_marker: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DeletePowerPointShapeRequest {
    slide_index: u32,
    shape_name: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct MarkLastPowerPointFormulaRequest {
    formula_id: String,
    previous_shape_names: Vec<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ReplaceLastPowerPointFormulaRequest {
    formula_id: String,
    previous_shape_names: Vec<String>,
    original_shape_name: String,
    left: f64,
    top: f64,
    width: f64,
    height: f64,
}

#[derive(Debug, Default, Deserialize)]
struct PowerPointEventsQuery {
    cursor: Option<u64>,
    host: Option<String>,
}

async fn get_powerpoint_native_selection() -> Response {
    match run_native_operation(powerpoint_native::selected_shape).await {
        Ok(selection) => Json(selection).into_response(),
        Err(error) => native_error_response(error),
    }
}

async fn apply_word_inline_baseline(
    Json(request): Json<ApplyWordInlineBaselineRequest>,
) -> Response {
    match run_native_operation(move || {
        word_native::apply_inline_baseline(request.position, &request.formula_marker)
    })
    .await
    {
        Ok(result) => Json(result).into_response(),
        Err(error) => native_error_response(error),
    }
}

async fn mark_powerpoint_native_selection(
    Json(request): Json<MarkPowerPointFormulaRequest>,
) -> Response {
    match run_native_operation(move || {
        powerpoint_native::mark_selected_formula(&request.formula_id)
    })
    .await
    {
        Ok(selection) => Json(selection).into_response(),
        Err(error) => native_error_response(error),
    }
}

async fn get_powerpoint_native_slide_snapshot() -> Response {
    match run_native_operation(powerpoint_native::active_slide_snapshot).await {
        Ok(snapshot) => Json(snapshot).into_response(),
        Err(error) => native_error_response(error),
    }
}

async fn mark_last_powerpoint_native_formula(
    Json(request): Json<MarkLastPowerPointFormulaRequest>,
) -> Response {
    match run_native_operation(move || {
        powerpoint_native::mark_last_inserted_formula(
            &request.formula_id,
            &request.previous_shape_names,
        )
    })
    .await
    {
        Ok(selection) => Json(selection).into_response(),
        Err(error) => native_error_response(error),
    }
}

async fn replace_last_powerpoint_native_formula(
    Json(request): Json<ReplaceLastPowerPointFormulaRequest>,
) -> Response {
    match run_native_operation(move || {
        powerpoint_native::replace_last_inserted_formula(
            &request.formula_id,
            &request.previous_shape_names,
            &request.original_shape_name,
            request.left,
            request.top,
            request.width,
            request.height,
        )
    })
    .await
    {
        Ok(selection) => Json(selection).into_response(),
        Err(error) => native_error_response(error),
    }
}

async fn delete_powerpoint_native_shape(
    Json(request): Json<DeletePowerPointShapeRequest>,
) -> Response {
    match run_native_operation(move || {
        powerpoint_native::delete_shape(request.slide_index, &request.shape_name)
    })
    .await
    {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => native_error_response(error),
    }
}

pub(crate) fn metadata_from_session(session: &OfficeFormulaSession) -> VisualTeXFormulaMetadata {
    let timestamp = format!("unix-ms:{}", session.updated_at);
    let original = session.original_metadata.as_ref();
    VisualTeXFormulaMetadata {
        schema: "visualtex-formula".to_string(),
        schema_version: 1,
        formula_id: session.formula_id.clone(),
        title: session.title.clone(),
        latex: session
            .lines
            .iter()
            .map(|line| line.latex.as_str())
            .collect::<Vec<_>>()
            .join("\n"),
        lines: session
            .lines
            .iter()
            .map(|line| MetadataLine {
                id: line.id.clone(),
                latex: line.latex.clone(),
            })
            .collect(),
        code_format: session.code_format.clone(),
        display_mode: session.display_mode.clone(),
        numbered: session.numbered,
        render_width_px: session
            .export_result
            .as_ref()
            .map(|value| value.width)
            .filter(|value| value.is_finite() && *value > 0.0),
        render_height_px: session
            .export_result
            .as_ref()
            .map(|value| value.height)
            .filter(|value| value.is_finite() && *value > 0.0),
        created_with_version: original
            .map(|value| value.created_with_version.clone())
            .unwrap_or_else(|| env!("CARGO_PKG_VERSION").to_string()),
        updated_with_version: env!("CARGO_PKG_VERSION").to_string(),
        created_at: original
            .map(|value| value.created_at.clone())
            .unwrap_or_else(|| timestamp.clone()),
        updated_at: timestamp,
    }
}

fn decode_powerpoint_native_edit_reference(value: &str) -> Option<(u32, String)> {
    let reference = value.strip_prefix("visualtex-ppt-native-edit:")?;
    let (slide_index, encoded_name) = reference.split_once(':')?;
    let slide_index = slide_index.parse::<u32>().ok().filter(|value| *value > 0)?;
    if encoded_name.is_empty() || encoded_name.len() > 512 || encoded_name.len() % 2 != 0 {
        return None;
    }
    let mut bytes = Vec::with_capacity(encoded_name.len() / 2);
    for pair in encoded_name.as_bytes().chunks_exact(2) {
        let digits = std::str::from_utf8(pair).ok()?;
        bytes.push(u8::from_str_radix(digits, 16).ok()?);
    }
    let shape_name = String::from_utf8(bytes).ok()?;
    (!shape_name.is_empty() && !shape_name.chars().any(char::is_control))
        .then_some((slide_index, shape_name))
}

fn prepare_powerpoint_session_blocking(
    companion: OfficeCompanionState,
    session_id: String,
    patch: serde_json::Value,
) -> Result<PreparedPowerPointCommit, String> {
    let _commit_guard = POWERPOINT_COMMIT_LOCK
        .lock()
        .map_err(|_| "PowerPoint commit lock is unavailable".to_string())?;
    let session = if patch.as_object().is_some_and(|value| !value.is_empty()) {
        companion
            .session_store
            .patch(&session_id, patch)
            .map_err(|error| error.to_string())?
    } else {
        companion
            .session_store
            .get(&session_id)
            .map_err(|error| error.to_string())?
    };
    if session.host != OfficeHost::Powerpoint || session.status != OfficeSessionStatus::Committing {
        return Err("PowerPoint Session is not ready to prepare".to_string());
    }

    // A command-page retry must decorate the already inserted shape, never
    // paste a second formula. The prepared selection remains immutable until
    // Office.js confirms that name, tags and accessibility metadata survived.
    if let Some(selection) = companion
        .prepared_powerpoint_commits
        .lock()
        .map_err(|_| "PowerPoint prepared-commit state is unavailable".to_string())?
        .get(&session_id)
        .cloned()
    {
        return Ok(PreparedPowerPointCommit { session, selection });
    }

    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "PowerPoint Session has no exported formula image".to_string())?;
    if export.svg.trim().is_empty() {
        return Err("PowerPoint Session contains an empty SVG export".to_string());
    }
    let temporary = std::env::temp_dir().join(format!(
        "visualtex-powerpoint-{}-{}.svg",
        session.formula_id,
        Uuid::new_v4()
    ));
    fs::write(&temporary, export.svg.as_bytes())
        .map_err(|error| format!("Unable to create temporary PowerPoint formula image: {error}"))?;
    let (target_width, target_height) = scale_powerpoint_formula_size(export.width, export.height)?;
    let target_slide_reference = session.source_object_id.as_deref().and_then(|value| {
        let reference = value.strip_prefix("visualtex-ppt-native-slide:")?;
        let mut fields = reference.split(':');
        let first = fields.next()?.parse::<u32>().ok()?;
        match fields.next() {
            Some(index) => Some((
                Some(first),
                index.parse::<u32>().ok().filter(|value| *value > 0),
            )),
            None => Some((None, (first > 0).then_some(first))),
        }
    });
    let expected_presentation_identity = session
        .source_document_id
        .as_deref()
        .and_then(|value| value.strip_prefix("visualtex-ppt-native-presentation:"));
    let edit_target = session
        .source_object_id
        .as_deref()
        .and_then(decode_powerpoint_native_edit_reference);
    let insertion = powerpoint_native::upsert_formula_picture_from_clipboard(
        &session.formula_id,
        &temporary.to_string_lossy(),
        target_width,
        target_height,
        export.height,
        session
            .original_metadata
            .as_ref()
            .and_then(|value| value.render_height_px),
        session.mode == OfficeSessionMode::Edit,
        edit_target.as_ref().map(|value| value.0),
        edit_target.as_ref().map(|value| value.1.as_str()),
        expected_presentation_identity,
        target_slide_reference.and_then(|value| value.0),
        target_slide_reference.and_then(|value| value.1),
    );
    let _ = fs::remove_file(&temporary);
    let selection = insertion?;
    companion
        .prepared_powerpoint_commits
        .lock()
        .map_err(|_| "PowerPoint prepared-commit state is unavailable".to_string())?
        .insert(session_id, selection.clone());
    Ok(PreparedPowerPointCommit { session, selection })
}

fn confirm_powerpoint_session_blocking(
    companion: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    let _commit_guard = POWERPOINT_COMMIT_LOCK
        .lock()
        .map_err(|_| "PowerPoint commit lock is unavailable".to_string())?;
    let session = companion
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())?;
    if session.host != OfficeHost::Powerpoint || session.status != OfficeSessionStatus::Committing {
        return Err("PowerPoint Session is not ready to confirm".to_string());
    }
    let has_prepared_shape = companion
        .prepared_powerpoint_commits
        .lock()
        .map_err(|_| "PowerPoint prepared-commit state is unavailable".to_string())?
        .contains_key(&session_id);
    if !has_prepared_shape {
        return Err("PowerPoint Session has no prepared formula shape".to_string());
    }
    companion
        .formula_cache
        .put(&session.formula_id, metadata_from_session(&session))
        .map_err(|error| format!("Formula metadata could not be saved: {error}"))?;
    companion
        .prepared_powerpoint_commits
        .lock()
        .map_err(|_| "PowerPoint prepared-commit state is unavailable".to_string())?
        .remove(&session_id);
    Ok(session)
}

async fn commit_powerpoint_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
    Json(patch): Json<serde_json::Value>,
) -> Response {
    let companion = context.companion.clone();
    let failure_store = context.companion.session_store.clone();
    let failure_id = session_id.clone();
    match run_native_operation(move || {
        prepare_powerpoint_session_blocking(companion, session_id, patch)
    })
    .await
    {
        Ok(prepared) => Json(prepared).into_response(),
        Err(error) => {
            let response_error = error.clone();
            let _ = run_session_operation(move || {
                failure_store.patch(
                    &failure_id,
                    serde_json::json!({ "status": "failed", "error": error }),
                )
            })
            .await;
            native_error_response(response_error)
        }
    }
}

async fn confirm_powerpoint_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    let companion = context.companion.clone();
    match run_native_operation(move || {
        confirm_powerpoint_session_blocking(companion, session_id)
    })
    .await
    {
        Ok(session) => Json(session).into_response(),
        Err(error) => native_error_response(error),
    }
}

async fn get_powerpoint_events(
    State(context): State<ServerContext>,
    Query(query): Query<PowerPointEventsQuery>,
) -> Json<Vec<PowerPointInteractionEvent>> {
    let host = match query.host.as_deref() {
        Some("word") => "word",
        Some("powerpoint") | None => "powerpoint",
        Some(_) => return Json(Vec::new()),
    };
    Json(
        context
            .companion
            .powerpoint_interactions
            .take_after(host, query.cursor.unwrap_or_default()),
    )
}

async fn run_session_operation<T>(
    operation: impl FnOnce() -> Result<T, SessionError> + Send + 'static,
) -> Result<T, SessionError>
where
    T: Send + 'static,
{
    tokio::task::spawn_blocking(operation)
        .await
        .map_err(|error| SessionError::Io(format!("Office Session task failed: {error}")))?
}

fn session_error_response(error: SessionError) -> Response {
    let status = match error {
        SessionError::Invalid(_) => StatusCode::BAD_REQUEST,
        SessionError::NotFound => StatusCode::NOT_FOUND,
        SessionError::Conflict(_) => StatusCode::CONFLICT,
        SessionError::Io(_) => StatusCode::INTERNAL_SERVER_ERROR,
    };
    (
        status,
        Json(serde_json::json!({ "error": error.to_string() })),
    )
        .into_response()
}

async fn create_session(
    State(context): State<ServerContext>,
    Json(input): Json<CreateOfficeSessionInput>,
) -> Response {
    let store = context.companion.session_store.clone();
    match run_session_operation(move || store.create(input)).await {
        Ok(session) => (StatusCode::CREATED, Json(session)).into_response(),
        Err(error) => session_error_response(error),
    }
}

async fn get_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    let store = context.companion.session_store.clone();
    match run_session_operation(move || store.get(&session_id)).await {
        Ok(session) => Json(session).into_response(),
        Err(error) => session_error_response(error),
    }
}

async fn patch_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
    Json(patch): Json<serde_json::Value>,
) -> Response {
    let store = context.companion.session_store.clone();
    match run_session_operation(move || store.patch(&session_id, patch)).await {
        Ok(session) => Json(session).into_response(),
        Err(error) => session_error_response(error),
    }
}

async fn get_formula_metadata(
    AxumPath(formula_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    let cache = context.companion.formula_cache.clone();
    match run_session_operation(move || cache.get(&formula_id)).await {
        Ok(metadata) => Json(metadata).into_response(),
        Err(error) => session_error_response(error),
    }
}

async fn put_formula_metadata(
    AxumPath(formula_id): AxumPath<String>,
    State(context): State<ServerContext>,
    Json(metadata): Json<VisualTeXFormulaMetadata>,
) -> Response {
    let cache = context.companion.formula_cache.clone();
    match run_session_operation(move || cache.put(&formula_id, metadata)).await {
        Ok(saved) => Json(saved).into_response(),
        Err(error) => session_error_response(error),
    }
}

async fn delete_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    let store = context.companion.session_store.clone();
    match run_session_operation(move || store.delete(&session_id)).await {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => session_error_response(error),
    }
}

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
    let api = Router::new()
        .route("/status", get(api_status))
        .route("/app/reveal", post(reveal_desktop_app))
        .route("/sessions", post(create_session))
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
        .route("/ocr/install", post(install_ocr))
        .route("/ocr/recognize", post(recognize_ocr))
        .route("/ocr/prewarm", post(prewarm_ocr))
        .route("/ocr/cancel", post(cancel_ocr))
        .route("/ocr/restart", post(restart_ocr))
        .route("/ocr/reset", post(reset_ocr))
        .route("/ocr/events", get(get_ocr_events))
        .fallback(api_not_found);

    Router::new()
        .route("/health", get(health))
        .nest("/api/v1", api)
        .fallback(not_found)
        .layer(RequestBodyLimitLayer::new(MAX_OFFICE_REQUEST_BYTES))
        .layer(middleware::from_fn_with_state(
            context.clone(),
            security_and_auth,
        ))
        .with_state(context)
}

fn bind_listener(address: std::net::SocketAddr) -> Result<TcpListener, String> {
    TcpListener::bind(address).map_err(|error| {
        format!(
            "VisualTeX Office companion cannot bind to {address}. The fixed port may already be in use: {error}"
        )
    })
}

pub async fn run(companion: OfficeCompanionState) -> Result<(), String> {
    let address = OfficeCompanionState::socket_addr();
    let listener = bind_listener(address)?;
    listener
        .set_nonblocking(true)
        .map_err(|error| format!("Unable to configure Office companion listener: {error}"))?;
    let tls =
        RustlsConfig::from_pem_file(&companion.paths.certificate, &companion.paths.private_key)
            .await
            .map_err(|error| format!("Unable to load Office TLS certificate: {error}"))?;
    let handle = axum_server::Handle::new();
    {
        let mut stored = companion
            .server_handle
            .lock()
            .map_err(|_| "Office companion server handle lock is unavailable".to_string())?;
        *stored = Some(handle.clone());
    }
    companion.update_status(|status| {
        status.running = true;
        status.last_error = None;
    });

    let result = axum_server::from_tcp_rustls(listener, tls)
        .map_err(|error| format!("Unable to create Office TLS server: {error}"))?
        .handle(handle)
        .serve(build_router(companion.clone()).into_make_service())
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

    fn test_state(temp: &TempDir) -> OfficeCompanionState {
        let root = temp.path().join("office-data");
        let paths = crate::office::state::OfficePaths {
            certificate: root.join("localhost-cert.pem"),
            private_key: root.join("localhost-key.pem"),
            certificate_metadata: root.join("certificate.json"),
            install: root.join("install.json"),
            sessions: root.join("sessions"),
            recovery: root.join("recovery"),
            formula_cache: root.join("formulas"),
            root,
        };
        let session_store = SessionStore::new(&paths).expect("session store");
        let formula_cache = crate::office::formula_cache::FormulaMetadataCache::new(&paths)
            .expect("formula metadata cache");
        OfficeCompanionState::new(
            None,
            crate::OcrState::default(),
            paths,
            "a".repeat(64),
            session_store,
            formula_cache,
            true,
        )
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

    #[test]
    fn occupied_port_returns_a_clear_error() {
        let first = TcpListener::bind(("127.0.0.1", 0)).expect("first listener");
        let address = first.local_addr().expect("local address");
        let error = bind_listener(address).expect_err("second bind must fail");
        assert!(error.contains("already be in use"));
        assert!(error.contains(&address.to_string()));
    }

    #[tokio::test]
    async fn router_serves_health_and_rejects_retired_html_routes() {
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
        assert_eq!(dialog.status(), StatusCode::NOT_FOUND);

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
                    .body(Body::from(vec![0_u8; crate::MAX_IMAGE_BYTES + 1]))
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
