//! PowerPoint: the native selection and slide snapshot, marking and replacing formulas, and committing a session.

use super::*;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct PreparedPowerPointCommit {
    pub(super) session: OfficeFormulaSession,
    pub(super) selection: PowerPointNativeSelection,
}

/// Convert the MathJax SVG pixel bounds to PowerPoint points without ever
/// changing their aspect ratio.  Width and height used to be clamped to 12pt
/// independently; a short formula would therefore acquire an artificial
/// height while retaining its natural width, and PowerPoint compressed the
/// glyphs into that distorted box on every replacement.
pub(super) fn scale_powerpoint_formula_size(width_px: f64, height_px: f64) -> Result<(f64, f64), String> {
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

pub(super) async fn run_native_operation<T>(
    operation: impl FnOnce() -> Result<T, String> + Send + 'static,
) -> Result<T, String>
where
    T: Send + 'static,
{
    tokio::task::spawn_blocking(operation)
        .await
        .map_err(|error| format!("PowerPoint native task failed: {error}"))?
}

pub(super) fn native_error_response(error: String) -> Response {
    (
        StatusCode::BAD_REQUEST,
        Json(serde_json::json!({ "error": error })),
    )
        .into_response()
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct MarkPowerPointFormulaRequest {
    pub(super) formula_id: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct ApplyWordInlineBaselineRequest {
    pub(super) position: i32,
    pub(super) formula_marker: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct DeletePowerPointShapeRequest {
    pub(super) slide_index: u32,
    pub(super) shape_name: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct MarkLastPowerPointFormulaRequest {
    pub(super) formula_id: String,
    pub(super) previous_shape_names: Vec<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct ReplaceLastPowerPointFormulaRequest {
    pub(super) formula_id: String,
    pub(super) previous_shape_names: Vec<String>,
    pub(super) original_shape_name: String,
    pub(super) left: f64,
    pub(super) top: f64,
    pub(super) width: f64,
    pub(super) height: f64,
}

#[derive(Debug, Default, Deserialize)]
pub(super) struct PowerPointEventsQuery {
    pub(super) cursor: Option<u64>,
    pub(super) host: Option<String>,
}

pub(super) async fn get_powerpoint_native_selection() -> Response {
    match run_native_operation(powerpoint_native::selected_shape).await {
        Ok(selection) => Json(selection).into_response(),
        Err(error) => native_error_response(error),
    }
}

pub(super) async fn apply_word_inline_baseline(
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

pub(super) async fn mark_powerpoint_native_selection(
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

pub(super) async fn get_powerpoint_native_slide_snapshot() -> Response {
    match run_native_operation(powerpoint_native::active_slide_snapshot).await {
        Ok(snapshot) => Json(snapshot).into_response(),
        Err(error) => native_error_response(error),
    }
}

pub(super) async fn mark_last_powerpoint_native_formula(
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

pub(super) async fn replace_last_powerpoint_native_formula(
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

pub(super) async fn delete_powerpoint_native_shape(
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

pub(super) fn metadata_from_session(session: &OfficeFormulaSession) -> VisualTeXFormulaMetadata {
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
        equation_tag: if session.display_mode == "block" {
            original.and_then(|metadata| metadata.equation_tag.clone())
        } else {
            None
        },
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
        baseline: session
            .export_result
            .as_ref()
            .and_then(|value| value.baseline)
            .filter(|value| value.is_finite() && *value >= 0.0),
        font_size_pt: Some(session.font_size_pt),
        render_font_size_pt: session
            .export_result
            .as_ref()
            .map(|_| session.font_size_pt)
            .or_else(|| original.and_then(|value| value.render_font_size_pt)),
        formula_letter_font: session
            .export_result
            .as_ref()
            .and_then(|value| value.formula_letter_font.clone())
            .or_else(|| original.and_then(|value| value.formula_letter_font.clone())),
        formula_chinese_font: session
            .export_result
            .as_ref()
            .and_then(|value| value.formula_chinese_font.clone())
            .or_else(|| original.and_then(|value| value.formula_chinese_font.clone())),
        formula_faces: session
            .export_result
            .as_ref()
            .and_then(|value| value.formula_faces.clone())
            .or_else(|| original.and_then(|value| value.formula_faces.clone())),
        native_omml_fingerprint: original
            .and_then(|value| value.native_omml_fingerprint.clone()),
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

pub(super) fn decode_powerpoint_native_edit_reference(value: &str) -> Option<(u32, String)> {
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

pub(super) fn prepare_powerpoint_session_blocking(
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

pub(super) fn confirm_powerpoint_session_blocking(
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

pub(super) async fn commit_powerpoint_session(
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

pub(super) async fn confirm_powerpoint_session(
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

pub(super) async fn get_powerpoint_events(
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
