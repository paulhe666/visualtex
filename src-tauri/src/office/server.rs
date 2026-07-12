use crate::office::sessions::{CreateOfficeSessionInput, SessionError, VisualTeXFormulaMetadata};
use crate::office::state::{
    OfficeCompanionState, MAX_OFFICE_REQUEST_BYTES, OFFICE_PROTOCOL_VERSION, OFFICE_UI_VERSION,
};
use axum::extract::{Path as AxumPath, Request, State};
use axum::http::{header, HeaderValue, StatusCode};
use axum::middleware::{self, Next};
use axum::response::{Html, IntoResponse, Response};
use axum::routing::{get, post};
use axum::{Json, Router};
use axum_server::tls_rustls::RustlsConfig;
use serde::Serialize;
use std::net::TcpListener;
use std::path::Path;
use tower_http::limit::RequestBodyLimitLayer;
use tower_http::services::ServeDir;

const INSTALL_TOKEN_HEADER: &str = "x-visualtex-install-token";
const OFFICE_CSP: &str = "default-src 'none'; script-src 'self' 'unsafe-eval'; style-src 'self' 'unsafe-inline'; connect-src 'self'; img-src 'self' data: blob:; font-src 'self' data:; object-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'self' https://*.office.com https://*.officeapps.live.com";

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

fn inject_install_token(html: &str, token: &str) -> Result<String, StatusCode> {
    let marker = "</head>";
    let index = html.find(marker).ok_or(StatusCode::INTERNAL_SERVER_ERROR)?;
    let meta = format!("<meta name=\"visualtex-install-token\" content=\"{token}\" />\n");
    Ok(format!("{}{}{}", &html[..index], meta, &html[index..]))
}

async fn read_office_html(
    ui_root: &Path,
    relative: &str,
    token: &str,
) -> Result<Html<String>, StatusCode> {
    let path = ui_root.join(relative);
    let html = tokio::fs::read_to_string(&path)
        .await
        .map_err(|_| StatusCode::NOT_FOUND)?;
    inject_install_token(&html, token).map(Html)
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

async fn bridge(State(context): State<ServerContext>) -> Result<Html<String>, StatusCode> {
    read_office_html(
        &context.companion.paths.ui_root,
        "bridge/index.html",
        &context.companion.install_token,
    )
    .await
}

async fn dialog(
    AxumPath(session_id): AxumPath<String>,
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
    read_office_html(
        &context.companion.paths.ui_root,
        "dialog/index.html",
        &context.companion.install_token,
    )
    .await
}

async fn api_status(
    State(context): State<ServerContext>,
) -> Json<crate::office::state::OfficeCompanionStatus> {
    Json(context.companion.snapshot())
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
    let ui_root = companion.paths.ui_root.clone();
    let api = Router::new()
        .route("/status", get(api_status))
        .route("/sessions", post(create_session))
        .route(
            "/sessions/{session_id}",
            get(get_session).patch(patch_session).delete(delete_session),
        )
        .route(
            "/formulas/{formula_id}/metadata",
            get(get_formula_metadata).put(put_formula_metadata),
        )
        .fallback(api_not_found);

    Router::new()
        .route("/health", get(health))
        .route("/bridge/index.html", get(bridge))
        .route("/dialog/{session_id}", get(dialog))
        .nest("/api/v1", api)
        .nest_service("/assets", ServeDir::new(ui_root.join("assets")))
        .nest_service(
            "/vendor/office-js",
            ServeDir::new(ui_root.join("vendor").join("office-js")),
        )
        .nest_service("/licenses", ServeDir::new(ui_root.join("licenses")))
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
        let ui_root = temp.path().join("office-ui");
        fs::create_dir_all(ui_root.join("bridge")).unwrap();
        fs::create_dir_all(ui_root.join("dialog")).unwrap();
        fs::create_dir_all(ui_root.join("assets")).unwrap();
        fs::create_dir_all(ui_root.join("vendor").join("office-js")).unwrap();
        fs::create_dir_all(ui_root.join("licenses")).unwrap();
        fs::write(
            ui_root.join("bridge").join("index.html"),
            "<html><head></head><body>bridge</body></html>",
        )
        .unwrap();
        fs::write(
            ui_root.join("dialog").join("index.html"),
            "<html><head></head><body>dialog</body></html>",
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
        OfficeCompanionState::new(paths, "a".repeat(64), session_store, formula_cache, true)
    }

    #[test]
    fn session_ids_reject_path_traversal() {
        assert!(valid_session_id("550e8400-e29b-41d4-a716-446655440000"));
        assert!(!valid_session_id("../../etc/passwd"));
        assert!(!valid_session_id("short"));
        assert!(!valid_session_id("session%2Fescape"));
    }

    #[test]
    fn install_token_comparison_is_length_sensitive() {
        assert!(constant_time_eq(b"same-token", b"same-token"));
        assert!(!constant_time_eq(b"same-token", b"same-tokeN"));
        assert!(!constant_time_eq(b"same-token", b"same-token-longer"));
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
        assert_eq!(bridge.status(), StatusCode::OK);
        let bridge_body = bridge.into_body().collect().await.unwrap().to_bytes();
        let bridge_html = String::from_utf8(bridge_body.to_vec()).unwrap();
        assert!(bridge_html.contains("visualtex-install-token"));
        assert!(bridge_html.contains(state.install_token.as_str()));

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
        assert_eq!(dialog.status(), StatusCode::OK);

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
}
