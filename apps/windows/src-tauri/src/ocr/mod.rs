//! The local OCR runtime: the private Python environment and its installation, the PaddleOCR worker, model
//! warm-up and recognition, the storage location, and the Tauri commands for all of it. Remote OCR providers are in
//! provider.rs.
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::collections::VecDeque;
use std::env;
use std::fs::{self, OpenOptions};
use std::io::{BufRead, BufReader, BufWriter, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Mutex, TryLockError};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
#[cfg(windows)]
use std::os::windows::process::CommandExt;
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Emitter, Manager, State};

pub(crate) mod install;
#[cfg(windows)]
mod models;
pub(crate) mod provider;
pub(crate) mod storage;
#[cfg(windows)]
mod python_bundle;
mod paths;
mod python;
mod setup;
mod status;
mod worker;

use paths::*;
use python::*;
use setup::*;
use status::*;
use worker::*;

use install::{
    append_install_log, begin_install_log_session,
    cleanup_stale_process, decode_process_output,
    install_log_path, load_snapshot,
    run_logged_command, save_snapshot,
    CommandCapture, CommandLimits,
    InstallControl, InstallSnapshot, InstallState,
};

const PADDLE_VERSION: &str = "3.3.1";
const PADDLEOCR_VERSION: &str = "3.7.0";
const DEFAULT_OCR_MODEL: &str = "PP-FormulaNet_plus-M";
pub(crate) const MAX_IMAGE_BYTES: usize = 20 * 1024 * 1024;
const OCR_CANCELLED: &str = "OCR_CANCELLED";
pub(crate) const ALLOWED_MODELS: &[&str] = &[
    "PP-FormulaNet_plus-S",
    "PP-FormulaNet_plus-M",
    "PP-FormulaNet_plus-L",
    "PP-FormulaNet-S",
    "PP-FormulaNet-L",
];
const MAX_OCR_EVENTS: usize = 256;
const OCR_RUNTIME_STATUS_SCHEMA: u32 = 1;
const OCR_RUNTIME_STATUS_FILE: &str = "runtime-status.json";
const OCR_PREFERRED_MODEL_FILE: &str = "preferred-model.txt";
const PRIVATE_PYTHON_SITE_CUSTOMIZE: &str = include_str!("../../ocr/private_sitecustomize.py");
const OCR_WORKER_WAIT_POLL_INTERVAL: Duration = Duration::from_millis(100);
const OCR_WORKER_WAIT_NOTICE_INTERVAL: Duration = Duration::from_secs(2);
#[cfg(windows)]
const CREATE_NO_WINDOW: u32 = 0x0800_0000;

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrEventRecord {
    id: u64,
    event: String,
    payload: Value,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrEventEnvelope {
    cursor: u64,
    events: Vec<OcrEventRecord>,
}

#[derive(Clone, Default)]
pub(crate) struct OcrEventBus {
    next_id: Arc<AtomicU64>,
    events: Arc<Mutex<VecDeque<OcrEventRecord>>>,
}

impl OcrEventBus {
    fn publish<T: Serialize>(&self, event: &str, payload: &T) {
        let Ok(payload) = serde_json::to_value(payload) else {
            return;
        };
        let id = self.next_id.fetch_add(1, Ordering::SeqCst) + 1;
        if let Ok(mut events) = self.events.lock() {
            events.push_back(OcrEventRecord {
                id,
                event: event.to_string(),
                payload,
            });
            while events.len() > MAX_OCR_EVENTS {
                events.pop_front();
            }
        }
    }

    pub(crate) fn poll(&self, cursor: u64, event: Option<&str>) -> OcrEventEnvelope {
        let events = self
            .events
            .lock()
            .map(|events| {
                events
                    .iter()
                    .filter(|item| item.id > cursor)
                    .filter(|item| event.is_none_or(|name| item.event == name))
                    .cloned()
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        OcrEventEnvelope {
            cursor: self.next_id.load(Ordering::SeqCst),
            events,
        }
    }
}

#[derive(Clone)]
struct RuntimePaths {
    root: PathBuf,
    venv: PathBuf,
    python: PathBuf,
    input: PathBuf,
    processed: PathBuf,
    logs: PathBuf,
    cache: PathBuf,
    temp: PathBuf,
    storage_config_path: PathBuf,
    storage_source: String,
    storage_managed: bool,
}

#[derive(Clone)]
pub(crate) struct OcrState {
    worker: Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: Arc<AtomicU32>,
    cancel_generation: Arc<AtomicU64>,
    runtime_status: Arc<Mutex<Option<OcrRuntimeStatus>>>,
    storage_change_running: Arc<AtomicBool>,
    runtime_mutations: Arc<AtomicU32>,
    install_control: Arc<InstallControl>,
    #[cfg(windows)]
    model_download_control: Arc<models::ModelDownloadControl>,
    desired_warmup_model: Arc<Mutex<Option<String>>>,
    providers: provider::OcrProviderState,
    events: OcrEventBus,
}

impl Default for OcrState {
    fn default() -> Self {
        Self {
            worker: Arc::new(Mutex::new(None)),
            worker_pid: Arc::new(AtomicU32::new(0)),
            cancel_generation: Arc::new(AtomicU64::new(0)),
            runtime_status: Arc::new(Mutex::new(None)),
            storage_change_running: Arc::new(AtomicBool::new(false)),
            runtime_mutations: Arc::new(AtomicU32::new(0)),
            install_control: Arc::new(InstallControl::default()),
            #[cfg(windows)]
            model_download_control: Arc::new(models::ModelDownloadControl::default()),
            desired_warmup_model: Arc::new(Mutex::new(None)),
            providers: provider::OcrProviderState::default(),
            events: OcrEventBus::default(),
        }
    }
}

struct StorageChangeLease {
    running: Arc<AtomicBool>,
}

impl Drop for StorageChangeLease {
    fn drop(&mut self) {
        self.running.store(false, Ordering::SeqCst);
    }
}

struct RuntimeMutationLease {
    counter: Arc<AtomicU32>,
}

impl Drop for RuntimeMutationLease {
    fn drop(&mut self) {
        self.counter.store(0, Ordering::SeqCst);
    }
}

fn is_final_ocr_state_owner(worker: &Arc<Mutex<Option<OcrWorker>>>) -> bool {
    Arc::strong_count(worker) == 1
}

async fn wait_for_ocr_cancellation(generation: Arc<AtomicU64>, expected: u64) {
    loop {
        if generation.load(Ordering::SeqCst) != expected {
            return;
        }
        tokio::time::sleep(Duration::from_millis(80)).await;
    }
}

impl Drop for OcrState {
    fn drop(&mut self) {
        // OcrState is cloned into Tauri state and the Office companion server.
        // Destroying any temporary clone must not kill the shared OCR worker.
        // Only the final owner is allowed to terminate the process.
        if is_final_ocr_state_owner(&self.worker) {
            self.cancel_generation.fetch_add(1, Ordering::SeqCst);
            let _ = self.install_control.cancel();
            #[cfg(windows)]
            self.model_download_control.cancel();
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

    fn begin_runtime_mutation(&self, operation: &str) -> Result<RuntimeMutationLease, String> {
        if self.storage_change_running.load(Ordering::SeqCst) {
            return Err(format!(
                "The OCR storage location is currently being changed; wait for it to finish before {operation}"
            ));
        }
        self.runtime_mutations
            .compare_exchange(0, 1, Ordering::SeqCst, Ordering::SeqCst)
            .map_err(|_| {
                format!(
                    "Another OCR runtime or model operation is active; wait for it to finish before {operation}"
                )
            })?;
        if self.storage_change_running.load(Ordering::SeqCst) {
            self.runtime_mutations.store(0, Ordering::SeqCst);
            return Err(format!(
                "An OCR storage location change started before {operation}; retry after it finishes"
            ));
        }
        Ok(RuntimeMutationLease {
            counter: self.runtime_mutations.clone(),
        })
    }

    fn begin_storage_change(&self) -> Result<StorageChangeLease, String> {
        self.storage_change_running
            .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
            .map_err(|_| "An OCR storage location change is already running".to_string())?;
        let lease = StorageChangeLease {
            running: self.storage_change_running.clone(),
        };
        if self.runtime_mutations.load(Ordering::SeqCst) != 0 || self.install_control.is_running() {
            return Err("OCR runtime files are currently being modified; wait for the active operation to finish before changing storage".to_string());
        }
        #[cfg(windows)]
        if self.model_download_control.is_running() {
            return Err("An OCR model download is active; cancel or finish it before changing storage".to_string());
        }
        Ok(lease)
    }

    fn ensure_storage_not_changing(&self, operation: &str) -> Result<(), String> {
        if self.storage_change_running.load(Ordering::SeqCst) {
            return Err(format!(
                "The OCR storage location is currently being changed; wait for it to finish before {operation}"
            ));
        }
        Ok(())
    }

    pub(crate) async fn runtime_status(
        &self,
        app: AppHandle,
        force_refresh: bool,
    ) -> Result<OcrRuntimeStatus, String> {
        if self.storage_change_running.load(Ordering::SeqCst) {
            return Err("The OCR storage location is being changed; runtime status will refresh when the reset/switch operation finishes".to_string());
        }
        let runtime_status = self.runtime_status.clone();
        let install_control = self.install_control.clone();
        let installing = install_control.is_running();
        // A cached runtime status is only safe while installation is actively
        // mutating the environment. Outside installation, always resolve the
        // current storage pointer and verify the actual files on disk so a path
        // switch or reset cannot leave the UI on a stale "installed" state.
        if installing {
            if let Some(status) = read_cached_runtime_status(&runtime_status)? {
                return Ok(status);
            }
        }

        tauri::async_runtime::spawn_blocking(move || {
            // Do not run the full import-tokenizers probe while an install
            // task is active. That would mix a previous partial environment
            // error into the current installation UI.
            let mut status = if should_run_full_runtime_probe(force_refresh, installing) {
                get_runtime_status_inner(&app)?
            } else {
                get_runtime_status_fast(&app)?
            };
            if installing && !status.installed {
                status.message = install_control
                    .snapshot()?
                    .map(|snapshot| snapshot.message)
                    .unwrap_or_else(|| "OCR installation is in progress".to_string());
            }
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR runtime status task failed: {error}"))?
    }

    pub(crate) async fn install_runtime(&self, app: AppHandle) -> Result<OcrRuntimeStatus, String> {
        let mutation = self.begin_runtime_mutation("installing the OCR runtime")?;
        let lease = self.install_control.begin()?;
        let generation = lease.generation();
        let install_control = self.install_control.clone();
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        write_cached_runtime_status(&runtime_status, None)?;
        let install_app = app.clone();
        let status = tauri::async_runtime::spawn_blocking(move || {
            let _mutation = mutation;
            let _lease = lease;
            let status = install_runtime_inner(
                &install_app,
                &worker,
                &worker_pid,
                &install_control,
                generation,
            )?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok::<OcrRuntimeStatus, String>(status)
        })
        .await
        .map_err(|error| format!("OCR installer task failed: {error}"))??;

        Ok(status)
    }

    pub(crate) async fn install_status(&self, app: AppHandle) -> Result<InstallSnapshot, String> {
        self.ensure_storage_not_changing("reading OCR installation status")?;
        let control = self.install_control.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let paths = runtime_paths(&app)?;
            let mut snapshot = control.snapshot()?.unwrap_or_else(|| {
                load_snapshot(&paths.root)
                    .unwrap_or_else(|| InstallSnapshot::new(&install_log_path(&paths.root)))
            });
            if reconcile_interrupted_install_snapshot(&mut snapshot, control.is_running()) {
                save_snapshot(&paths.root, &snapshot)?;
            }
            control.set_snapshot(snapshot.clone())?;
            Ok(snapshot)
        })
        .await
        .map_err(|error| format!("OCR installation status task failed: {error}"))?
    }

    pub(crate) fn cancel_install(&self) -> Result<(), String> {
        self.install_control.cancel().map(|_| ())
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
        self.recognize_local_with_fallback(app, request, false).await
    }

    pub(crate) async fn recognize_with_local_fallback(
        &self,
        app: AppHandle,
        request: OcrImageRequest,
    ) -> Result<OcrRecognitionResult, String> {
        if self.providers.active_provider(&app)? != provider::LOCAL_PROVIDER {
            return self.recognize(app, request).await;
        }
        self.recognize_local_with_fallback(app, request, true).await
    }

    async fn recognize_local_with_fallback(
        &self,
        app: AppHandle,
        mut request: OcrImageRequest,
        allow_model_fallback: bool,
    ) -> Result<OcrRecognitionResult, String> {
        let mutation = self.begin_runtime_mutation("recognizing a formula")?;
        if !ALLOWED_MODELS.contains(&request.model.as_str()) {
            return Err(format!("Unsupported PP-FormulaNet model: {}", request.model));
        }
        let status = self.runtime_status(app.clone(), false).await?;
        if !status.installed {
            return Err(format!("OCR runtime is not installed: {}", status.message));
        }
        if !status.installed_models.iter().any(|model| model == &request.model) {
            if allow_model_fallback {
                request.model = if status
                    .installed_models
                    .iter()
                    .any(|model| model == &status.default_model)
                {
                    status.default_model.clone()
                } else {
                    status
                        .installed_models
                        .first()
                        .cloned()
                        .ok_or_else(|| "No OCR model is installed".to_string())?
                };
            } else {
                return Err(format!(
                    "OCR model {} is not installed. Import a verified .vtxocrmodel package or explicitly download it first.",
                    request.model
                ));
            }
        }
        {
            let mut desired = self
                .desired_warmup_model
                .lock()
                .map_err(|_| "OCR warmup selection is unavailable".to_string())?;
            *desired = Some(request.model.clone());
        }
        if let Ok(paths) = runtime_paths(&app) {
            if let Err(error) = write_preferred_ocr_model(&paths, &request.model) {
                eprintln!("Unable to persist the active OCR model: {error}");
            }
        }
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let cancel_generation = self.cancel_generation.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let _mutation = mutation;
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
        let terminate_result = terminate_worker_process(&self.worker_pid);

        // Killing the worker is the only reliable way to interrupt a Paddle
        // inference that is blocked inside native code. Rebuild and rewarm it
        // immediately in the background so cancelling one request does not
        // make the next request pay the full model startup cost again.
        schedule_worker_rewarm(
            app.clone(),
            self.worker.clone(),
            self.worker_pid.clone(),
            self.runtime_status.clone(),
            self.desired_warmup_model.clone(),
            self.storage_change_running.clone(),
            self.runtime_mutations.clone(),
        );

        terminate_result.map(|_| ())
    }

    pub(crate) fn shutdown(&self, app: &AppHandle) -> Result<(), String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let _ = self.install_control.cancel();
        stop_worker(&self.worker, &self.worker_pid)?;
        self.install_control.cancel_and_wait(Duration::from_secs(5))?;
        if let Ok(paths) = runtime_paths(app) {
            let _ = cleanup_stale_process(&paths.root);
        }
        Ok(())
    }

    pub(crate) async fn restart(&self, app: AppHandle) -> Result<(), String> {
        let mutation = self.begin_runtime_mutation("restarting the OCR worker")?;
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let _mutation = mutation;
            stop_worker(&worker, &worker_pid)?;
            cleanup_worker_temp(&runtime_paths(&app)?)
        })
        .await
        .map_err(|error| format!("OCR restart task failed: {error}"))?
    }

    pub(crate) fn schedule_startup_warmup(&self, app: AppHandle) {
        let state = self.clone();
        let _ = tauri::async_runtime::spawn(async move {
            // Give the desktop UI enough time to publish the model stored in
            // localStorage. If it already requested a warmup, that task owns
            // model selection and this background fallback must not enqueue a
            // stale second model. Pure background launches still prewarm from
            // the persisted preference after the delay.
            tokio::time::sleep(Duration::from_secs(15)).await;
            if state
                .active_provider(&app)
                .map(|provider| provider != provider::LOCAL_PROVIDER)
                .unwrap_or(false)
            {
                return;
            }
            let frontend_already_selected = state
                .desired_warmup_model
                .lock()
                .map(|desired| desired.is_some())
                .unwrap_or(false);
            if frontend_already_selected {
                return;
            }
            let status = state.runtime_status(app.clone(), false).await;
            if let Ok(status) = status {
                if !status.installed || status.installed_models.is_empty() {
                    return;
                }
                let preferred = runtime_paths(&app)
                    .ok()
                    .and_then(|paths| read_preferred_ocr_model(&paths));
                let model = preferred
                    .filter(|model| status.installed_models.iter().any(|installed| installed == model))
                    .unwrap_or_else(|| status.installed_models[0].clone());
                if let Err(error) = state.warmup_model(app, model).await {
                    eprintln!("Unable to prewarm the preferred OCR model: {error}");
                }
            }
        });
    }

    pub(crate) async fn warmup_model(
        &self,
        app: AppHandle,
        model: String,
    ) -> Result<(), String> {
        if self.active_provider(&app)? != provider::LOCAL_PROVIDER {
            return Ok(());
        }
        let mutation = self.begin_runtime_mutation("warming up an OCR model")?;
        if !ALLOWED_MODELS.contains(&model.as_str()) {
            return Err(format!("Unsupported PP-FormulaNet model: {model}"));
        }
        let status = self.runtime_status(app.clone(), false).await?;
        if !status.installed_models.iter().any(|installed| installed == &model) {
            return Err(format!(
                "OCR model {model} is not installed. Import a verified .vtxocrmodel package or explicitly download it first."
            ));
        }
        {
            let mut desired = self
                .desired_warmup_model
                .lock()
                .map_err(|_| "OCR warmup selection is unavailable".to_string())?;
            *desired = Some(model.clone());
        }
        if let Ok(paths) = runtime_paths(&app) {
            if let Err(error) = write_preferred_ocr_model(&paths, &model) {
                eprintln!("Unable to persist the preferred OCR model: {error}");
            }
        }
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        let desired_warmup_model = self.desired_warmup_model.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let _mutation = mutation;
            warmup_worker(
                &app,
                &worker,
                &worker_pid,
                &runtime_status,
                &desired_warmup_model,
                &model,
            )
        })
        .await
        .map_err(|error| format!("OCR warmup task failed: {error}"))?
    }

    pub(crate) async fn reset_runtime(&self, app: AppHandle) -> Result<OcrRuntimeStatus, String> {
        let mutation = self.begin_runtime_mutation("resetting the OCR runtime")?;
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let _ = self.install_control.cancel();
        let install_control = self.install_control.clone();
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        write_cached_runtime_status(&runtime_status, None)?;
        tauri::async_runtime::spawn_blocking(move || {
            let _mutation = mutation;
            install_control.cancel_and_wait(Duration::from_secs(15))?;
            write_cached_runtime_status(&runtime_status, None)?;
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            let _ = cleanup_stale_process(&paths.root);
            reset_runtime_contents(&paths.root)?;
            install_control.set_snapshot(InstallSnapshot::new(&install_log_path(&paths.root)))?;
            let status = get_runtime_status_fast(&app)?;
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
        let mutation = self.begin_runtime_mutation("installing an OCR model package")?;
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let _mutation = mutation;
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            #[cfg(windows)]
            models::install_model_pack(&package_path, &paths.root)?;
            #[cfg(not(windows))]
            crate::ocr_offline::install_optional_model_pack(&package_path, &paths.root)?;
            let status = get_runtime_status_fast(&app)?;
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
        let mutation = self.begin_runtime_mutation("removing an OCR model")?;
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let _mutation = mutation;
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            #[cfg(windows)]
            models::remove_model(&paths.root, &model)?;
            #[cfg(not(windows))]
            crate::ocr_offline::remove_optional_model(&paths.root, &model)?;
            let status = get_runtime_status_fast(&app)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR model removal task failed: {error}"))?
    }

    #[cfg(windows)]
    pub(crate) async fn model_catalog(
        &self,
        app: AppHandle,
    ) -> Result<models::ModelCatalog, String> {
        tauri::async_runtime::spawn_blocking(move || models::load_catalog(&app))
            .await
            .map_err(|error| format!("OCR model catalog task failed: {error}"))?
    }

    #[cfg(windows)]
    pub(crate) fn model_download_status(
        &self,
    ) -> Result<Option<models::ModelDownloadSnapshot>, String> {
        self.model_download_control.snapshot()
    }

    #[cfg(windows)]
    pub(crate) fn cancel_model_download(&self) -> bool {
        self.model_download_control.cancel()
    }

    #[cfg(windows)]
    pub(crate) async fn download_model(
        &self,
        app: AppHandle,
        model: String,
    ) -> Result<OcrRuntimeStatus, String> {
        if !models::KNOWN_MODELS.contains(&model.as_str()) {
            return Err(format!("Unsupported PP-FormulaNet model: {model}"));
        }
        let _mutation = self.begin_runtime_mutation("downloading and installing an OCR model")?;
        let lease = self.model_download_control.begin()?;
        let generation = lease.generation();
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let prepare_app = app.clone();
        let paths = tauri::async_runtime::spawn_blocking(move || {
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&prepare_app)?;
            probe_runtime_from_files(&paths).map_err(|error| {
                format!("Install and verify the OCR runtime before downloading a model: {error}")
            })?;
            Ok::<_, String>(paths)
        })
        .await
        .map_err(|error| format!("OCR model preparation task failed: {error}"))??;

        let events = self.events.clone();
        let download_app = app.clone();
        models::download_and_install_model(
            &download_app,
            &paths.root,
            &model,
            lease.control(),
            generation,
            |progress| {
                let _ = download_app.emit("ocr-model-download-progress", progress);
                events.publish("ocr-model-download-progress", progress);
            },
        )
        .await?;

        let runtime_status = self.runtime_status.clone();
        let status_app = app.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let status = get_runtime_status_fast(&status_app)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok::<OcrRuntimeStatus, String>(status)
        })
        .await
        .map_err(|error| format!("OCR model status refresh task failed: {error}"))?
    }

    pub(crate) async fn configure_storage(
        &self,
        app: AppHandle,
        selected_directory: PathBuf,
    ) -> Result<OcrRuntimeStatus, String> {
        let storage_lease = self.begin_storage_change()?;
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        let install_control = self.install_control.clone();
        write_cached_runtime_status(&runtime_status, None)?;
        tauri::async_runtime::spawn_blocking(move || {
            let _storage_lease = storage_lease;
            stop_worker(&worker, &worker_pid)?;
            write_cached_runtime_status(&runtime_status, None)?;
            let current = runtime_paths(&app)?;
            let change = storage::configure(
                &current.storage_config_path,
                &current.root,
                &selected_directory,
                &protected_ocr_storage_roots(&app),
            )?;

            let target_paths = runtime_paths_from_resolution(storage::StorageResolution {
                root: change.target_root.clone(),
                config_path: current.storage_config_path.clone(),
                source: "pending-switch".to_string(),
                managed: storage::marker_is_valid(&change.target_root),
            });
            let mut reused_target_environment = false;
            let target_probe = if change.adopted_existing {
                match probe_runtime(&target_paths) {
                    Ok(probe) => {
                        reused_target_environment = true;
                        Ok(probe)
                    }
                    Err(error) if storage::marker_is_valid(&change.target_root) => {
                        let _ = cleanup_stale_process(&change.target_root);
                        reset_runtime_contents(&change.target_root)?;
                        Err(format!(
                            "The selected VisualTeX OCR directory contained an incomplete environment and was reset before reinstalling: {error}"
                        ))
                    }
                    Err(error) => {
                        return Err(format!(
                            "The selected unmarked OCR environment could not be verified and was not modified: {}. Error: {}",
                            change.target_root.display(),
                            error
                        ));
                    }
                }
            } else {
                probe_runtime_from_files(&target_paths)
            };

            if change.reset_source {
                let _ = cleanup_stale_process(&current.root);
                reset_runtime_contents(&current.root)?;
            }

            if let Err(error) =
                storage::write_config(&current.storage_config_path, &change.target_root)
            {
                let _ = storage::write_config(
                    &current.storage_config_path,
                    &change.source_root,
                );
                return Err(format!(
                    "The OCR environment was reset, but the new storage pointer could not be written. Pointer: {}. Error: {}",
                    current.storage_config_path.display(),
                    error
                ));
            }

            let next = runtime_paths(&app)?;
            if next.root != change.target_root {
                let rollback = storage::write_config(
                    &current.storage_config_path,
                    &change.source_root,
                );
                return Err(format!(
                    "OCR storage pointer verification failed. Expected {}, resolved {}. Rollback status: {}",
                    change.target_root.display(),
                    next.root.display(),
                    rollback
                        .map(|_| "restored the previous pointer".to_string())
                        .unwrap_or_else(|error| format!("failed to restore the previous pointer: {error}"))
                ));
            }

            install_control.set_snapshot(InstallSnapshot::new(&install_log_path(&next.root)))?;
            let mut status = runtime_status_from_probe(&app, &next, target_probe);
            if reused_target_environment {
                status.message = format!(
                    "{} Existing OCR storage was adopted without reinstalling the environment or models.",
                    status.message
                );
            } else if change.reset_source {
                status.message = format!(
                    "{} The previous OCR environment was reset. Install the OCR runtime again at the new location.",
                    status.message
                );
            } else {
                status.message = format!(
                    "{} OCR storage location changed. Install the OCR runtime at this location.",
                    status.message
                );
            }
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR storage reset/switch task failed: {error}"))?
    }

    pub(crate) fn poll_events(&self, cursor: u64, event: Option<&str>) -> OcrEventEnvelope {
        self.events.poll(cursor, event)
    }
}

struct OcrWorker {
    child: Child,
    stdin: BufWriter<ChildStdin>,
    stdout: BufReader<ChildStdout>,
    pid_state: Arc<AtomicU32>,
    log_path: PathBuf,
    loaded_model: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrInstallProgress {
    stage: String,
    state: InstallState,
    percent: u8,
    message: String,
    detail: Option<String>,
    error: Option<String>,
    log_path: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrRuntimeStatus {
    installed: bool,
    python_path: Option<String>,
    python_version: Option<String>,
    paddle_version: Option<String>,
    paddleocr_version: Option<String>,
    runtime_path: String,
    storage_config_path: String,
    storage_source: String,
    storage_managed: bool,
    storage_available_bytes: Option<u64>,
    storage_persistent_across_uninstall: bool,
    runtime_bundle_available: bool,
    offline_bundle_available: bool,
    installed_models: Vec<String>,
    damaged_models: Vec<String>,
    model_catalog_available: bool,
    default_model: String,
    message: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrImageRequest {
    pub(crate) bytes: Vec<u8>,
    pub(crate) extension: String,
    pub(crate) model: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct OcrFormulaResult {
    pub(crate) latex: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrRecognitionResult {
    provider: String,
    model: String,
    elapsed_ms: u64,
    processed_width: u32,
    processed_height: u32,
    background_inverted: bool,
    background_luminance: f64,
    pub(crate) formulas: Vec<OcrFormulaResult>,
}

#[derive(Debug, Deserialize)]
struct WorkerRecognitionResponse {
    ok: bool,
    model: Option<String>,
    elapsed_ms: Option<u64>,
    processed_width: Option<u32>,
    processed_height: Option<u32>,
    background_inverted: Option<bool>,
    background_luminance: Option<f64>,
    formulas: Option<Vec<OcrFormulaResult>>,
    error: Option<String>,
    details: Option<String>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct RuntimeProbe {
    #[serde(rename = "pythonVersion", alias = "python_version")]
    python_version: String,
    #[serde(rename = "paddleVersion", alias = "paddle_version")]
    paddle_version: String,
    #[serde(rename = "paddleocrVersion", alias = "paddleocr_version")]
    paddleocr_version: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct RuntimeStatusManifest {
    #[serde(rename = "schemaVersion", alias = "schema_version")]
    schema_version: u32,
    #[serde(rename = "pythonPath", alias = "python_path")]
    python_path: String,
    #[serde(rename = "pythonVersion", alias = "python_version")]
    python_version: String,
    #[serde(rename = "paddleVersion", alias = "paddle_version")]
    paddle_version: String,
    #[serde(rename = "paddleocrVersion", alias = "paddleocr_version")]
    paddleocr_version: String,
}

#[derive(Debug, Deserialize)]
struct PythonProbe {
    version: String,
    major: u8,
    minor: u8,
    bits: u8,
    machine: String,
    executable: String,
}

fn emit_progress(
    app: &AppHandle,
    stage: &str,
    percent: u8,
    message: impl Into<String>,
    detail: Option<String>,
) {
    let state = match stage {
        "complete" => InstallState::Complete,
        "verify" => InstallState::Verifying,
        "dependencies-installed" => InstallState::DependenciesInstalled,
        _ => InstallState::Installing,
    };
    let progress = OcrInstallProgress {
        stage: stage.to_string(),
        state,
        percent,
        message: message.into(),
        detail,
        error: None,
        log_path: runtime_paths(app)
            .ok()
            .map(|paths| install_log_path(&paths.root).display().to_string()),
    };
    let _ = app.emit("ocr-install-progress", &progress);
    if let Some(state) = app.try_state::<OcrState>() {
        state.events.publish("ocr-install-progress", &progress);
    }
}

fn publish_install_snapshot(
    app: &AppHandle,
    control: &InstallControl,
    paths: &RuntimePaths,
    snapshot: &mut InstallSnapshot,
) -> Result<(), String> {
    snapshot.touch();
    save_snapshot(&paths.root, snapshot)?;
    control.set_snapshot(snapshot.clone())?;
    let progress = OcrInstallProgress {
        stage: snapshot
            .current_step
            .clone()
            .unwrap_or_else(|| "status".to_string()),
        state: snapshot.state.clone(),
        percent: snapshot.percent,
        message: snapshot.message.clone(),
        detail: snapshot.detail.clone(),
        error: snapshot.error.clone(),
        log_path: Some(snapshot.log_path.clone()),
    };
    let _ = app.emit("ocr-install-progress", &progress);
    if let Some(state) = app.try_state::<OcrState>() {
        state.events.publish("ocr-install-progress", &progress);
    }
    Ok(())
}

fn set_install_step(
    app: &AppHandle,
    control: &InstallControl,
    paths: &RuntimePaths,
    snapshot: &mut InstallSnapshot,
    state: InstallState,
    step: Option<&str>,
    percent: u8,
    message: impl Into<String>,
    detail: Option<String>,
    error: Option<String>,
) -> Result<(), String> {
    snapshot.state = state;
    snapshot.current_step = step.map(str::to_string);
    snapshot.percent = percent;
    snapshot.message = message.into();
    snapshot.detail = detail;
    snapshot.error = error;
    publish_install_snapshot(app, control, paths, snapshot)
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
    #[cfg(windows)]
    crate::windows_silent_ocr_hotkey::handle_ocr_progress(app, &progress);
    if let Some(state) = app.try_state::<OcrState>() {
        state.events.publish("ocr-recognition-progress", &progress);
    }
}

fn tail_text(value: &str, max_chars: usize) -> String {
    let total = value.chars().count();
    if total <= max_chars {
        return value.to_string();
    }
    value.chars().skip(total - max_chars).collect()
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct InstalledPackageProbe {
    distribution: String,
    version: String,
    module_path: String,
    executable: String,
    python_version: String,
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
pub(crate) async fn configure_ocr_storage_location(
    app: AppHandle,
    state: State<'_, OcrState>,
    selected_directory: String,
) -> Result<OcrRuntimeStatus, String> {
    let selected_directory = selected_directory.trim();
    if selected_directory.is_empty() {
        return Err("No OCR storage directory was selected".to_string());
    }
    state
        .configure_storage(app, PathBuf::from(selected_directory))
        .await
}

#[tauri::command]
pub(crate) fn open_ocr_storage_location(app: AppHandle) -> Result<(), String> {
    let paths = runtime_paths(&app)?;
    storage::ensure_marker(&paths.root)?;
    #[cfg(windows)]
    {
        let mut command = Command::new("explorer.exe");
        command.arg(&paths.root);
        hide_windows_console(&mut command);
        return command
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("Unable to open OCR storage location: {error}"));
    }
    #[cfg(not(windows))]
    {
        let _ = paths;
        Err("Opening the OCR storage location is currently available only on Windows".to_string())
    }
}

#[tauri::command]
pub(crate) async fn install_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    state.install_runtime(app).await
}

#[tauri::command]
pub(crate) async fn get_ocr_install_status(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<InstallSnapshot, String> {
    state.install_status(app).await
}

#[tauri::command]
pub(crate) fn cancel_ocr_install(state: State<'_, OcrState>) -> Result<(), String> {
    state.cancel_install()
}

#[tauri::command]
pub(crate) fn open_ocr_install_logs(app: AppHandle) -> Result<(), String> {
    let paths = runtime_paths(&app)?;
    fs::create_dir_all(&paths.logs)
        .map_err(|error| format!("Unable to create OCR installation log directory: {error}"))?;
    #[cfg(windows)]
    {
        let mut command = Command::new("explorer.exe");
        command.arg(&paths.logs);
        hide_windows_console(&mut command);
        return command
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("Unable to open OCR installation logs: {error}"));
    }
    #[cfg(not(windows))]
    {
        let _ = paths;
        Err("Opening OCR installation logs from the app is currently available only on Windows".to_string())
    }
}

#[tauri::command]
pub(crate) fn get_ocr_provider_configuration(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<provider::OcrProviderConfigurationView, String> {
    state.provider_configuration(&app)
}

#[tauri::command]
pub(crate) fn save_ocr_provider_configuration(
    app: AppHandle,
    state: State<'_, OcrState>,
    configuration: provider::OcrProviderConfigurationUpdate,
) -> Result<provider::OcrProviderConfigurationView, String> {
    state.save_provider_configuration(&app, configuration)
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
pub(crate) fn cancel_ocr_recognition(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.cancel(&app)
}

#[tauri::command]
pub(crate) async fn restart_ocr_worker(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.restart(app).await
}

#[tauri::command]
pub(crate) async fn warmup_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: String,
) -> Result<(), String> {
    state.warmup_model(app, model).await
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

#[cfg(windows)]
#[tauri::command]
pub(crate) async fn get_ocr_model_catalog(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<models::ModelCatalog, String> {
    state.model_catalog(app).await
}

#[cfg(windows)]
#[tauri::command]
pub(crate) fn get_ocr_model_download_status(
    state: State<'_, OcrState>,
) -> Result<Option<models::ModelDownloadSnapshot>, String> {
    state.model_download_status()
}

#[cfg(windows)]
#[tauri::command]
pub(crate) async fn download_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: String,
) -> Result<OcrRuntimeStatus, String> {
    state.download_model(app, model).await
}

#[cfg(windows)]
#[tauri::command]
pub(crate) fn cancel_ocr_model_download(state: State<'_, OcrState>) -> bool {
    state.cancel_model_download()
}

#[cfg(windows)]
#[tauri::command]
pub(crate) fn configure_silent_ocr(
    app: AppHandle,
    enabled: bool,
    model: String,
    copy_format: String,
    capture_mode: String,
) -> Result<String, String> {
    crate::windows_silent_ocr_hotkey::configure(
        &app,
        enabled,
        &model,
        &copy_format,
        &capture_mode,
    )
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
    fn desktop_and_office_ocr_state_clones_share_one_runtime_and_model_manager() {
        let desktop = OcrState::default();
        let office = desktop.clone();

        assert!(Arc::ptr_eq(&desktop.worker, &office.worker));
        assert!(Arc::ptr_eq(&desktop.worker_pid, &office.worker_pid));
        assert!(Arc::ptr_eq(
            &desktop.runtime_status,
            &office.runtime_status
        ));
        assert!(Arc::ptr_eq(
            &desktop.storage_change_running,
            &office.storage_change_running
        ));
        assert!(Arc::ptr_eq(
            &desktop.runtime_mutations,
            &office.runtime_mutations
        ));
        assert!(Arc::ptr_eq(
            &desktop.install_control,
            &office.install_control
        ));
        assert!(Arc::ptr_eq(
            &desktop.desired_warmup_model,
            &office.desired_warmup_model
        ));
        #[cfg(windows)]
        assert!(Arc::ptr_eq(
            &desktop.model_download_control,
            &office.model_download_control
        ));
    }

    #[test]
    fn storage_change_and_runtime_mutations_are_mutually_exclusive() {
        let state = OcrState::default();
        let mutation = state
            .begin_runtime_mutation("test mutation")
            .expect("first runtime mutation should start");
        assert!(state.begin_runtime_mutation("second mutation").is_err());
        assert!(state.begin_storage_change().is_err());
        drop(mutation);

        let storage = state
            .begin_storage_change()
            .expect("storage location change should start after the mutation ends");
        assert!(state.begin_runtime_mutation("mutation during storage change").is_err());
        assert!(state.begin_storage_change().is_err());
        drop(storage);
        assert!(state.begin_runtime_mutation("mutation after storage change").is_ok());
    }

    #[test]
    fn runtime_status_cache_round_trips_and_clears() {
        let cache = Arc::new(Mutex::new(None));
        let expected = OcrRuntimeStatus {
            installed: true,
            python_path: Some("/tmp/visualtex-python".to_string()),
            python_version: Some("3.12.10".to_string()),
            paddle_version: Some(PADDLE_VERSION.to_string()),
            paddleocr_version: Some(PADDLEOCR_VERSION.to_string()),
            runtime_path: "/tmp/visualtex-ocr".to_string(),
            storage_config_path: "/tmp/ocr-storage.json".to_string(),
            storage_source: "configured".to_string(),
            storage_managed: true,
            storage_available_bytes: Some(1024),
            storage_persistent_across_uninstall: true,
            runtime_bundle_available: true,
            offline_bundle_available: true,
            installed_models: vec![DEFAULT_OCR_MODEL.to_string()],
            damaged_models: Vec::new(),
            model_catalog_available: true,
            default_model: DEFAULT_OCR_MODEL.to_string(),
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

    fn test_runtime_paths(root: &Path) -> RuntimePaths {
        let venv = root.join("venv");
        RuntimePaths {
            root: root.to_path_buf(),
            python: venv.join("Scripts").join("python.exe"),
            venv,
            input: root.join("input"),
            processed: root.join("processed"),
            logs: root.join("logs"),
            cache: root.join("cache"),
            temp: root.join("tmp"),
            storage_config_path: root.join("ocr-storage.json"),
            storage_source: "test".to_string(),
            storage_managed: true,
        }
    }

    #[test]
    fn every_ocr_python_subprocess_cache_follows_runtime_storage() {
        let root = tempfile::tempdir().unwrap();
        let paths = test_runtime_paths(root.path());
        let mut command = Command::new("python.exe");
        configure_ocr_runtime_cache_environment(&mut command, &paths);

        let value = |name: &str| {
            command
                .get_envs()
                .find(|(key, _)| *key == std::ffi::OsStr::new(name))
                .and_then(|(_, value)| value)
                .map(PathBuf::from)
                .unwrap_or_else(|| panic!("missing OCR subprocess environment variable {name}"))
        };

        assert_eq!(value("PADDLE_PDX_CACHE_HOME"), paths.cache.join("paddlex"));
        assert_eq!(value("PADDLE_HOME"), paths.cache.join("paddle"));
        assert_eq!(value("MODELSCOPE_CACHE"), paths.cache.join("modelscope"));
        assert_eq!(value("HF_HOME"), paths.cache.join("huggingface"));
        assert_eq!(value("HF_HUB_CACHE"), paths.cache.join("huggingface").join("hub"));
        assert_eq!(
            value("HUGGINGFACE_HUB_CACHE"),
            paths.cache.join("huggingface").join("hub")
        );
        assert_eq!(value("XDG_CACHE_HOME"), paths.cache);
        assert_eq!(value("PIP_CACHE_DIR"), paths.cache.join("pip"));
        assert_eq!(value("TMPDIR"), paths.temp);
        assert_eq!(value("TMP"), paths.temp);
        assert_eq!(value("TEMP"), paths.temp);
        assert_eq!(
            command
                .get_envs()
                .find(|(key, _)| *key == std::ffi::OsStr::new("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"))
                .and_then(|(_, value)| value)
                .map(|value| value.to_string_lossy().into_owned())
                .as_deref(),
            Some("True")
        );
    }

    #[test]
    fn runtime_probe_accepts_camel_case_and_legacy_snake_case() {
        let camel: RuntimeProbe = serde_json::from_str(
            r#"{"pythonVersion":"3.12.10","paddleVersion":"3.3.1","paddleocrVersion":"3.7.0"}"#,
        )
        .expect("camelCase probe should parse");
        let snake: RuntimeProbe = serde_json::from_str(
            r#"{"python_version":"3.12.10","paddle_version":"3.3.1","paddleocr_version":"3.7.0"}"#,
        )
        .expect("legacy snake_case probe should parse");
        assert_eq!(camel.python_version, snake.python_version);
        assert_eq!(camel.paddle_version, PADDLE_VERSION);
        assert_eq!(snake.paddleocr_version, PADDLEOCR_VERSION);
    }

    #[test]
    fn invalid_runtime_probe_preserves_raw_stdout_and_stderr() {
        let error = parse_runtime_probe_output(
            "Paddle diagnostic\n{\"paddleVersion\":\"3.3.1\"}",
            "No ccache found (warning only)",
        )
        .expect_err("missing pythonVersion should fail");
        assert!(error.contains("Raw stdout"));
        assert!(error.contains("Paddle diagnostic"));
        assert!(error.contains("Raw stderr"));
        assert!(error.contains("No ccache found"));
    }

    #[test]
    fn bundled_python_replaces_incompatible_private_runtime() {
        assert!(should_replace_with_bundled_python(
            "The existing VisualTeX OCR environment uses a 32-bit interpreter"
        ));
        assert!(should_replace_with_bundled_python(
            "Unsupported VisualTeX OCR Python 3.13.5 (64-bit). It must be replaced by the bundled x64 Python 3.12 runtime."
        ));
        assert!(should_replace_with_bundled_python(
            "The existing VisualTeX OCR private Python is missing the app-local Microsoft OpenMP runtime vcomp140.dll"
        ));
    }

    #[test]
    fn fixed_wheelhouse_install_is_strictly_offline_and_hashed() {
        let root = tempfile::tempdir().unwrap();
        let assets = python_bundle::WindowsOfflineInstallAssets {
            wheelhouse: root.path().join("wheelhouse"),
            lockfile: root.path().join("requirements.lock"),
            manifest: python_bundle::WindowsPythonBundleManifest {
                schema_version: 2,
                platform: "windows".to_string(),
                architecture: "x64".to_string(),
                python_version: "3.12.10".to_string(),
                pip_version: "25.1.1".to_string(),
                archive: python_bundle::BundleFileRecord {
                    name: "python.zip".to_string(),
                    size: 1,
                    sha256: "0".repeat(64),
                },
                app_local_runtime: python_bundle::WindowsAppLocalRuntimeManifest {
                    open_mp: python_bundle::WindowsOpenMpRuntimeManifest {
                        file: python_bundle::BundleFileRecord {
                            name: "vcomp140.dll".to_string(),
                            size: 1,
                            sha256: "0".repeat(64),
                        },
                        version: "14.44.35211.0".to_string(),
                        source: "Microsoft Visual C++ Redistributable app-local OpenMP runtime"
                            .to_string(),
                    },
                },
                wheelhouse: python_bundle::WindowsWheelhouseManifest {
                    lock: python_bundle::BundleFileRecord {
                        name: "requirements.lock".to_string(),
                        size: 1,
                        sha256: "0".repeat(64),
                    },
                    files: Vec::new(),
                },
            },
        };
        let arguments = offline_pip_install_arguments(&assets);
        assert!(arguments.iter().any(|value| value == "--no-index"));
        assert!(arguments.iter().any(|value| value == "--find-links"));
        assert!(arguments.iter().any(|value| value == "--require-hashes"));
        assert!(arguments.iter().any(|value| value == "--only-binary=:all:"));
        assert!(arguments.iter().any(|value| value == "--no-deps"));
        assert!(arguments.iter().any(|value| value == "--no-compile"));
        assert!(arguments
            .iter()
            .any(|value| value == "--no-warn-script-location"));
        assert!(!arguments.iter().any(|value| value.contains("pypi.org")));
        assert!(!arguments.iter().any(|value| value == "--index-url"));
    }

    #[test]
    fn large_packages_have_long_total_timeout_and_separate_idle_timeout() {
        let paddle = command_limits_for_step("paddle");
        let paddleocr = command_limits_for_step("paddleocr");
        assert_eq!(paddle.step_timeout, Duration::from_secs(30 * 60));
        assert_eq!(paddleocr.step_timeout, Duration::from_secs(30 * 60));
        assert_eq!(paddle.idle_timeout, Duration::from_secs(6 * 60));
        assert!(paddle.idle_timeout < paddle.step_timeout);
    }

    #[test]
    fn failed_offline_install_resumes_that_step_instead_of_resetting_environment() {
        let root = tempfile::tempdir().unwrap();
        let mut snapshot = InstallSnapshot::new(&install_log_path(root.path()));
        snapshot.mark_step_complete("venv");
        snapshot.mark_step_complete("paddle");
        snapshot.mark_step_complete("paddleocr");
        snapshot.state = InstallState::InstallFailed;
        snapshot.current_step = Some("tokenizers".to_string());
        save_snapshot(root.path(), &snapshot).unwrap();
        let restored = load_snapshot(root.path()).expect("half-installed state should restore");
        assert!(restored.step_complete("paddle"));
        assert!(restored.step_complete("paddleocr"));
        assert!(!restored.step_complete("tokenizers"));
        assert_eq!(restored.current_step.as_deref(), Some("tokenizers"));
    }

    #[test]
    fn fast_runtime_probe_detects_missing_tokenizers_in_half_installed_environment() {
        let root = tempfile::tempdir().unwrap();
        let paths = test_runtime_paths(root.path());
        fs::create_dir_all(paths.python.parent().unwrap()).unwrap();
        fs::write(&paths.python, b"fake-python").unwrap();
        fs::write(paths.venv.join("pyvenv.cfg"), "version = 3.12.10\n").unwrap();
        let site_packages = paths.venv.join("Lib").join("site-packages");
        for module in ["paddle", "paddleocr", "ftfy", "wand"] {
            fs::create_dir_all(site_packages.join(module)).unwrap();
        }
        fs::create_dir_all(&site_packages).unwrap();
        fs::write(site_packages.join("imagesize.py"), b"# fake").unwrap();
        for dist_info in [
            format!("paddlepaddle-{PADDLE_VERSION}.dist-info"),
            format!("paddleocr-{PADDLEOCR_VERSION}.dist-info"),
            "imagesize-2.0.0.dist-info".to_string(),
            "ftfy-6.3.1.dist-info".to_string(),
            "Wand-0.7.2.dist-info".to_string(),
        ] {
            fs::create_dir_all(site_packages.join(dist_info)).unwrap();
        }

        let error = probe_runtime_from_files(&paths)
            .expect_err("missing tokenizers must not be reported as installed");
        assert!(error.contains("tokenizers"));

        fs::create_dir_all(site_packages.join("tokenizers")).unwrap();
        fs::create_dir_all(site_packages.join("tokenizers-0.19.1.dist-info")).unwrap();
        let probe = probe_runtime_from_files(&paths)
            .expect("complete dependency files should restore fast status");
        assert_eq!(probe.python_version, "3.12.10");
        assert_eq!(probe.paddle_version, PADDLE_VERSION);
        assert_eq!(probe.paddleocr_version, PADDLEOCR_VERSION);
    }

    #[test]
    fn package_probe_rejects_wrong_interpreter_and_external_module_path() {
        let root = tempfile::tempdir().unwrap();
        let paths = test_runtime_paths(root.path());
        let wrong_interpreter = InstalledPackageProbe {
            distribution: "tokenizers".to_string(),
            version: "0.19.1".to_string(),
            module_path: paths
                .venv
                .join("Lib/site-packages/tokenizers/__init__.py")
                .display()
                .to_string(),
            executable: "C:\\Python312\\python.exe".to_string(),
            python_version: "3.12.10".to_string(),
        };
        let error = validate_installed_package_probe(
            &paths,
            "tokenizers",
            Some("0.19.1"),
            &wrong_interpreter,
        )
        .expect_err("wrong interpreter should fail");
        assert!(error.contains("wrong interpreter"));

        let external_module = InstalledPackageProbe {
            executable: paths.python.display().to_string(),
            module_path: "C:\\Python312\\Lib\\site-packages\\tokenizers\\__init__.py"
                .to_string(),
            ..wrong_interpreter
        };
        let error = validate_installed_package_probe(
            &paths,
            "tokenizers",
            Some("0.19.1"),
            &external_module,
        )
        .expect_err("external module path should fail");
        assert!(error.contains("not installed inside"));
    }

    #[test]
    fn installation_blocks_full_runtime_probe_and_old_error_mixing() {
        assert!(should_run_full_runtime_probe(true, false));
        assert!(!should_run_full_runtime_probe(true, true));
        assert!(!should_run_full_runtime_probe(false, true));
    }

    #[test]
    fn interrupted_persisted_install_becomes_retryable_failure() {
        let root = tempfile::tempdir().unwrap();
        let mut snapshot = InstallSnapshot::new(&install_log_path(root.path()));
        snapshot.state = InstallState::Installing;
        snapshot.current_step = Some("tokenizers".to_string());
        snapshot.percent = 82;
        snapshot.mark_step_complete("paddle");
        snapshot.mark_step_complete("paddleocr");

        assert!(reconcile_interrupted_install_snapshot(&mut snapshot, false));
        assert_eq!(snapshot.state, InstallState::InstallFailed);
        assert_eq!(snapshot.current_step.as_deref(), Some("tokenizers"));
        assert_eq!(snapshot.percent, 82);
        assert!(snapshot.step_complete("paddle"));
        assert!(snapshot.error.as_deref().is_some_and(|value| value.contains("tokenizers")));

        let mut active = snapshot.clone();
        active.state = InstallState::Verifying;
        active.current_step = Some("verify".to_string());
        assert!(!reconcile_interrupted_install_snapshot(&mut active, true));
        assert_eq!(active.state, InstallState::Verifying);
    }

    #[test]
    fn damaged_venv_detection_excludes_incompatible_cancelled_and_timed_out_cases() {
        assert!(damaged_venv_python_error(
            "OCR virtual environment Python check failed with exit code: 1"
        ));
        assert!(damaged_venv_python_error(
            "OCR virtual environment returned invalid Python information"
        ));
        assert!(!damaged_venv_python_error(
            "The existing VisualTeX OCR environment uses Python 3.13.5"
        ));
        assert!(!damaged_venv_python_error(
            "OCR virtual environment Python check was cancelled"
        ));
        assert!(!damaged_venv_python_error(
            "OCR virtual environment Python check timed out"
        ));
    }

    #[test]
    fn runtime_status_manifest_round_trips_and_reads_legacy_fields() {
        let root = tempfile::tempdir().unwrap();
        let paths = test_runtime_paths(root.path());
        let expected = RuntimeProbe {
            python_version: "3.12.10".to_string(),
            paddle_version: PADDLE_VERSION.to_string(),
            paddleocr_version: PADDLEOCR_VERSION.to_string(),
        };
        write_runtime_status_manifest(&paths, &expected).unwrap();
        let current = read_runtime_status_manifest(&paths).expect("current manifest should load");
        assert_eq!(current.python_version, "3.12.10");

        let legacy = json!({
            "schema_version": OCR_RUNTIME_STATUS_SCHEMA,
            "python_path": paths.python.display().to_string(),
            "python_version": "3.12.9",
            "paddle_version": PADDLE_VERSION,
            "paddleocr_version": PADDLEOCR_VERSION,
        });
        fs::write(
            runtime_status_manifest_path(&paths),
            serde_json::to_vec_pretty(&legacy).unwrap(),
        )
        .unwrap();
        let legacy_probe =
            read_runtime_status_manifest(&paths).expect("legacy manifest should load");
        assert_eq!(legacy_probe.python_version, "3.12.9");
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
