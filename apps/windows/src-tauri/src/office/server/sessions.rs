//! The formula sessions the add-ins open: create, read, change and delete them, and formula metadata.

use super::*;

pub(super) async fn run_session_operation<T>(
    operation: impl FnOnce() -> Result<T, SessionError> + Send + 'static,
) -> Result<T, SessionError>
where
    T: Send + 'static,
{
    tokio::task::spawn_blocking(operation)
        .await
        .map_err(|error| SessionError::Io(format!("Office Session task failed: {error}")))?
}

pub(super) fn session_error_response(error: SessionError) -> Response {
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

pub(super) async fn create_session(
    State(context): State<ServerContext>,
    Json(input): Json<CreateOfficeSessionInput>,
) -> Response {
    let store = context.companion.session_store.clone();
    match run_session_operation(move || store.create(input)).await {
        Ok(session) => (StatusCode::CREATED, Json(session)).into_response(),
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn get_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    let store = context.companion.session_store.clone();
    match run_session_operation(move || store.get(&session_id)).await {
        Ok(session) => Json(session).into_response(),
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn patch_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
    Json(patch): Json<serde_json::Value>,
) -> Response {
    let store = context.companion.session_store.clone();
    let patch_session_id = session_id.clone();
    #[cfg(target_os = "windows")]
    let app = context.companion.app.clone();
    match run_session_operation(move || store.patch(&patch_session_id, patch)).await {
        Ok(session) => {
            #[cfg(target_os = "windows")]
            let hide_editor =
                matches!(
                    session.status,
                    OfficeSessionStatus::Completed
                        | OfficeSessionStatus::Cancelled
                        | OfficeSessionStatus::Failed
                )
                || (session.mode == OfficeSessionMode::Edit
                    && session.status == OfficeSessionStatus::Committing
                    && session.object_mode == "wordOmml");
            if hide_editor {
                if let Some(app) = app {
                    if let Err(error) =
                        hide_owned_office_editor_on_main_thread(app, session_id.clone()).await
                    {
                        crate::app_lifecycle::append_lifecycle_log(format!(
                            "edit-committing/terminal Session could not synchronously hide Office editor session={session_id}: {error}"
                        ));
                    }
                }
            }
            Json(session).into_response()
        }
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn get_formula_metadata(
    AxumPath(formula_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    let cache = context.companion.formula_cache.clone();
    match run_session_operation(move || cache.get(&formula_id)).await {
        Ok(metadata) => Json(metadata).into_response(),
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn put_formula_metadata(
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

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct DeleteSessionsBatchInput {
    pub(super) session_ids: Vec<String>,
}

pub(super) async fn delete_sessions_batch(
    State(context): State<ServerContext>,
    Json(input): Json<DeleteSessionsBatchInput>,
) -> Response {
    if input.session_ids.len() > 512 {
        return session_error_response(SessionError::Invalid(
            "Too many Office Sessions in one batch cleanup request".to_string(),
        ));
    }
    let store = context.companion.session_store.clone();
    match run_session_operation(move || store.delete_many(&input.session_ids)).await {
        Ok(_) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => session_error_response(error),
    }
}

pub(super) async fn delete_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    let store = context.companion.session_store.clone();
    match run_session_operation(move || store.delete(&session_id)).await {
        Ok(()) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => session_error_response(error),
    }
}
