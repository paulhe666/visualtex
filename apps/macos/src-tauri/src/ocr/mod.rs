//! The local OCR runtime: the offline Python runtime and its installation, the PaddleOCR worker, model
//! warm-up and recognition, and the Tauri commands for all of it. Remote OCR providers are in provider.rs.
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::collections::VecDeque;
use std::fs::{self, OpenOptions};
use std::io::{BufRead, BufReader, BufWriter, Write};
use std::path::PathBuf;
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};
use std::sync::atomic::{AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, SystemTime, UNIX_EPOCH};
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Emitter, Manager, State};

pub(crate) mod provider;
#[path = "../../../../shared/visualtex-backend/src/ocr_common.rs"]
mod common;
mod paths;
mod python;
mod setup;
mod status;
mod worker;

pub(crate) use common::*;
use paths::*;
use python::*;
use setup::*;
use status::*;
use worker::*;

const OCR_RUNTIME_PROBE_CACHE_SCHEMA: u32 = 1;

const OCR_RUNTIME_PROBE_CACHE_FILE: &str = "runtime-probe-cache.json";

#[derive(Clone)]
struct RuntimePaths {
    root: PathBuf,
    python: PathBuf,
    input: PathBuf,
    processed: PathBuf,
    logs: PathBuf,
    cache: PathBuf,
    temp: PathBuf,
}

#[derive(Clone)]
pub(crate) struct OcrState {
    worker: Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: Arc<AtomicU32>,
    cancel_generation: Arc<AtomicU64>,
    runtime_status: Arc<Mutex<Option<OcrRuntimeStatus>>>,
    events: OcrEventBus,
    providers: provider::OcrProviderState,
}

impl Default for OcrState {
    fn default() -> Self {
        Self {
            worker: Arc::new(Mutex::new(None)),
            worker_pid: Arc::new(AtomicU32::new(0)),
            cancel_generation: Arc::new(AtomicU64::new(0)),
            runtime_status: Arc::new(Mutex::new(None)),
            events: OcrEventBus::default(),
            providers: provider::OcrProviderState::default(),
        }
    }
}

impl Drop for OcrState {
    fn drop(&mut self) {
        // OcrState is cloned into Tauri state and the Office companion server.
        // Destroying any temporary clone must not kill the shared OCR worker.
        // Only the final owner is allowed to terminate the process.
        if is_final_ocr_state_owner(&self.worker) {
            self.cancel_generation.fetch_add(1, Ordering::SeqCst);
            let _ = terminate_worker_process(&self.worker_pid);
        }
    }
}

impl OcrState {
    pub(crate) fn provider_configuration(
        &self,
        app: &AppHandle,
    ) -> Result<provider::OcrProviderConfigurationView, String> {
        self.providers.configuration(app)
    }

    pub(crate) fn save_provider_configuration(
        &self,
        app: &AppHandle,
        configuration: provider::OcrProviderConfigurationUpdate,
    ) -> Result<provider::OcrProviderConfigurationView, String> {
        self.providers.save_configuration(app, configuration)
    }

    pub(crate) fn active_provider(&self, app: &AppHandle) -> Result<String, String> {
        self.providers.active_provider(app)
    }

    pub(crate) async fn runtime_status(
        &self,
        app: AppHandle,
        force_refresh: bool,
    ) -> Result<OcrRuntimeStatus, String> {
        let runtime_status = self.runtime_status.clone();
        if !force_refresh {
            if let Some(status) = read_cached_runtime_status(&runtime_status)? {
                return Ok(status);
            }
        }

        tauri::async_runtime::spawn_blocking(move || {
            let status = get_runtime_status_inner(&app, force_refresh)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR runtime status task failed: {error}"))?
    }

    pub(crate) async fn install_runtime(&self, app: AppHandle) -> Result<OcrRuntimeStatus, String> {
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let status = install_runtime_inner(&app, &worker, &worker_pid)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR installer task failed: {error}"))?
    }

    pub(crate) async fn recognize(
        &self,
        app: AppHandle,
        request: OcrImageRequest,
    ) -> Result<OcrRecognitionResult, String> {
        let provider = self.providers.active_provider(&app)?;
        if provider == provider::LOCAL_PROVIDER {
            return self.recognize_local(app, request).await;
        }

        let expected_generation = self.cancel_generation.load(Ordering::SeqCst);
        let cancellation_generation = self.cancel_generation.clone();
        tokio::select! {
            result = self.providers.recognize_remote(&app, &request) => result,
            _ = wait_for_ocr_cancellation(cancellation_generation, expected_generation) => {
                Err(OCR_CANCELLED.to_string())
            }
        }
    }

    async fn recognize_local(
        &self,
        app: AppHandle,
        request: OcrImageRequest,
    ) -> Result<OcrRecognitionResult, String> {
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let cancel_generation = self.cancel_generation.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            run_recognition(
                &app,
                &worker,
                &worker_pid,
                &cancel_generation,
                &runtime_status,
                request,
            )
        })
        .await
        .map_err(|error| format!("OCR recognition task failed: {error}"))?
    }

    pub(crate) async fn prewarm_model(&self, app: AppHandle, model: String) -> Result<(), String> {
        if self.active_provider(&app)? != provider::LOCAL_PROVIDER {
            return Ok(());
        }
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            prewarm_model_inner(&app, &worker, &worker_pid, &runtime_status, &model)
        })
        .await
        .map_err(|error| format!("OCR model prewarm task failed: {error}"))?
    }

    pub(crate) fn cancel(&self, app: &AppHandle) -> Result<(), String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        if self
            .providers
            .active_provider(app)
            .unwrap_or_else(|_| provider::LOCAL_PROVIDER.to_string())
            != provider::LOCAL_PROVIDER
        {
            return Ok(());
        }
        terminate_worker_process(&self.worker_pid)?;
        cleanup_worker_temp(&runtime_paths(app)?)
    }

    pub(crate) async fn restart(&self, app: AppHandle) -> Result<(), String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        tauri::async_runtime::spawn_blocking(move || {
            stop_worker(&worker, &worker_pid)?;
            cleanup_worker_temp(&runtime_paths(&app)?)
        })
        .await
        .map_err(|error| format!("OCR restart task failed: {error}"))?
    }

    pub(crate) async fn reset_runtime(&self, app: AppHandle) -> Result<OcrRuntimeStatus, String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            write_cached_runtime_status(&runtime_status, None)?;
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            if paths.root.exists() {
                fs::remove_dir_all(&paths.root)
                    .map_err(|error| format!("Unable to remove OCR runtime: {error}"))?;
            }
            let status = get_runtime_status_inner(&app, false)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR reset task failed: {error}"))?
    }

    pub(crate) async fn install_optional_model(
        &self,
        app: AppHandle,
        package_path: PathBuf,
    ) -> Result<OcrRuntimeStatus, String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            crate::ocr_offline::install_optional_model_pack(&package_path, &paths.root)?;
            let status = get_runtime_status_inner(&app, false)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR model pack installation task failed: {error}"))?
    }

    pub(crate) async fn remove_optional_model(
        &self,
        app: AppHandle,
        model: String,
    ) -> Result<OcrRuntimeStatus, String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            crate::ocr_offline::remove_optional_model(&paths.root, &model)?;
            let status = get_runtime_status_inner(&app, false)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR model removal task failed: {error}"))?
    }

    pub(crate) fn poll_events(&self, cursor: u64, event: Option<&str>) -> OcrEventEnvelope {
        self.events.poll(cursor, event)
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrInstallProgress {
    stage: String,
    percent: u8,
    message: String,
    detail: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrRuntimeStatus {
    pub(crate) installed: bool,
    python_path: Option<String>,
    python_version: Option<String>,
    paddle_version: Option<String>,
    paddleocr_version: Option<String>,
    runtime_path: String,
    offline_bundle_available: bool,
    pub(crate) installed_models: Vec<String>,
    pub(crate) default_model: String,
    message: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct RuntimeProbe {
    python_version: String,
    paddle_version: String,
    paddleocr_version: String,
}

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
struct RuntimeProbeCache {
    schema_version: u32,
    python_path: String,
    python_size: u64,
    python_modified_ms: u64,
    probe: RuntimeProbe,
}

fn emit_progress(
    app: &AppHandle,
    stage: &str,
    percent: u8,
    message: impl Into<String>,
    detail: Option<String>,
) {
    let progress = OcrInstallProgress {
        stage: stage.to_string(),
        percent,
        message: message.into(),
        detail,
    };
    let _ = app.emit("ocr-install-progress", &progress);
    if let Some(state) = app.try_state::<OcrState>() {
        state.events.publish("ocr-install-progress", &progress);
    }
}

fn emit_recognition_progress(
    app: &AppHandle,
    request_id: &str,
    stage: &str,
    message: impl Into<String>,
    model: &str,
) {
    let progress = json!({
        "event": "progress",
        "id": request_id,
        "stage": stage,
        "message": message.into(),
        "model": model,
    });
    let _ = app.emit("ocr-recognition-progress", &progress);
    crate::quick_ocr::handle_ocr_progress(app, &progress);
    if let Some(state) = app.try_state::<OcrState>() {
        state.events.publish("ocr-recognition-progress", &progress);
    }
}

#[tauri::command]
pub(crate) async fn get_ocr_provider_configuration(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<provider::OcrProviderConfigurationView, String> {
    let state = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || state.provider_configuration(&app))
        .await
        .map_err(|error| format!("Unable to finish OCR provider configuration lookup: {error}"))?
}

#[tauri::command]
pub(crate) fn save_ocr_provider_configuration(
    app: AppHandle,
    state: State<'_, OcrState>,
    configuration: provider::OcrProviderConfigurationUpdate,
) -> Result<provider::OcrProviderConfigurationView, String> {
    let saved = state.save_provider_configuration(&app, configuration)?;
    let _ = app.emit("ocr-provider-configuration-changed", &saved);
    Ok(saved)
}

#[tauri::command]
pub(crate) async fn get_ocr_runtime_status(
    app: AppHandle,
    state: State<'_, OcrState>,
    force_refresh: Option<bool>,
) -> Result<OcrRuntimeStatus, String> {
    state
        .runtime_status(app, force_refresh.unwrap_or(false))
        .await
}

#[tauri::command]
pub(crate) async fn install_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    state.install_runtime(app).await
}

#[tauri::command]
pub(crate) async fn recognize_formula_image(
    app: AppHandle,
    state: State<'_, OcrState>,
    request: OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    state.recognize(app, request).await
}

#[tauri::command]
pub(crate) async fn prewarm_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: String,
) -> Result<(), String> {
    state.prewarm_model(app, model).await
}

#[tauri::command]
pub(crate) fn cancel_ocr_recognition(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.cancel(&app)
}

#[tauri::command]
pub(crate) async fn restart_ocr_worker(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.restart(app).await
}

#[tauri::command]
pub(crate) async fn reset_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    state.reset_runtime(app).await
}

#[tauri::command]
pub(crate) async fn install_optional_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    package_path: String,
) -> Result<OcrRuntimeStatus, String> {
    let package_path = package_path.trim();
    if package_path.is_empty() {
        return Err("No OCR model package was selected".to_string());
    }
    state
        .install_optional_model(app, PathBuf::from(package_path))
        .await
}

#[tauri::command]
pub(crate) async fn remove_optional_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: String,
) -> Result<OcrRuntimeStatus, String> {
    state.remove_optional_model(app, model).await
}

#[cfg(test)]
mod protocol_tests {
    use super::*;
    use std::io::{BufReader, Cursor};

    #[test]
    fn worker_protocol_skips_non_utf8_diagnostics_before_json() {
        let bytes = b"\xd5\xfd\xca\xbd\xc8\xd5\xd6\xbe\n{\"event\":\"ready\",\"ok\":true}\n";
        let mut reader = BufReader::new(Cursor::new(bytes));

        let value = read_worker_json(&mut reader, "closed", "test response")
            .expect("reader should skip non-protocol diagnostic bytes");

        assert_eq!(value.get("event").and_then(Value::as_str), Some("ready"));
    }

    #[test]
    fn worker_protocol_decodes_ascii_escaped_unicode() {
        let bytes = b"{\"message\":\"\\u6b63\\u5728\\u52a0\\u8f7d\"}\n";
        let mut reader = BufReader::new(Cursor::new(bytes));

        let value = read_worker_json(&mut reader, "closed", "test response")
            .expect("escaped Unicode JSON should parse");

        assert_eq!(
            value.get("message").and_then(Value::as_str),
            Some("正在加载")
        );
    }

    #[test]
    fn ocr_event_bus_supports_cursor_filtering_and_bounded_history() {
        let events = OcrEventBus::default();
        let baseline = events.poll(u64::MAX, None);
        assert_eq!(baseline.cursor, 0);
        assert!(baseline.events.is_empty());

        events.publish(
            "ocr-install-progress",
            &json!({ "stage": "python", "percent": 5 }),
        );
        events.publish(
            "ocr-recognition-progress",
            &json!({ "stage": "model", "model": "PP-FormulaNet_plus-M" }),
        );

        let install_only = events.poll(0, Some("ocr-install-progress"));
        assert_eq!(install_only.cursor, 2);
        assert_eq!(install_only.events.len(), 1);
        assert_eq!(install_only.events[0].event, "ocr-install-progress");

        let incremental = events.poll(1, None);
        assert_eq!(incremental.events.len(), 1);
        assert_eq!(incremental.events[0].id, 2);

        for index in 0..(MAX_OCR_EVENTS + 20) {
            events.publish("ocr-recognition-progress", &json!({ "index": index }));
        }
        let bounded = events.poll(0, None);
        assert_eq!(bounded.events.len(), MAX_OCR_EVENTS);
        assert_eq!(bounded.cursor, (MAX_OCR_EVENTS + 22) as u64);
        assert!(bounded.events.first().is_some_and(|event| event.id > 1));
    }

    #[test]
    fn temporary_ocr_state_clone_cannot_terminate_shared_worker() {
        let state = OcrState::default();
        let temporary = state.clone();
        assert!(!is_final_ocr_state_owner(&state.worker));
        drop(temporary);
        assert!(is_final_ocr_state_owner(&state.worker));
    }

    #[test]
    fn runtime_status_cache_round_trips_and_clears() {
        let cache = Arc::new(Mutex::new(None));
        let expected = OcrRuntimeStatus {
            installed: true,
            python_path: Some("/tmp/visualtex-python".to_string()),
            python_version: Some("3.13.0".to_string()),
            paddle_version: Some(PADDLE_VERSION.to_string()),
            paddleocr_version: Some(PADDLEOCR_VERSION.to_string()),
            runtime_path: "/tmp/visualtex-ocr".to_string(),
            offline_bundle_available: true,
            installed_models: vec![crate::ocr_offline::OFFLINE_DEFAULT_MODEL.to_string()],
            default_model: crate::ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: "ready".to_string(),
        };

        write_cached_runtime_status(&cache, Some(expected.clone()))
            .expect("cache write should succeed");
        let cached = read_cached_runtime_status(&cache)
            .expect("cache read should succeed")
            .expect("cached status should exist");
        assert_eq!(cached.python_version, expected.python_version);
        assert!(cached.installed);

        write_cached_runtime_status(&cache, None).expect("cache clear should succeed");
        assert!(read_cached_runtime_status(&cache)
            .expect("cache read after clear should succeed")
            .is_none());
    }
}

#[cfg(all(test, unix))]
mod tests {
    use super::*;

    #[test]
    fn active_worker_can_be_terminated_without_taking_the_worker_lock() {
        let mut child = Command::new("sleep")
            .arg("30")
            .spawn()
            .expect("failed to start test process");
        let pid = child.id();
        let worker_pid = AtomicU32::new(pid);

        assert!(terminate_worker_process(&worker_pid).expect("termination failed"));
        let status = child.wait().expect("failed to wait for test process");

        assert!(!status.success());
        assert_eq!(worker_pid.load(Ordering::SeqCst), 0);
    }
}
