use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::collections::VecDeque;
use std::fs::{self, OpenOptions};
use std::io::{BufRead, BufReader, BufWriter, Write};
use std::path::PathBuf;
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Emitter, Manager, State};

mod ocr_offline;
mod office;
mod quick_ocr;
mod system_math_glyphs;

const PADDLE_VERSION: &str = "3.3.1";
const PADDLEOCR_VERSION: &str = "3.7.0";
pub(crate) const MAX_IMAGE_BYTES: usize = 20 * 1024 * 1024;
const OCR_CANCELLED: &str = "OCR_CANCELLED";
const ALLOWED_MODELS: &[&str] = &[
    "PP-FormulaNet_plus-S",
    "PP-FormulaNet_plus-M",
    "PP-FormulaNet_plus-L",
    "PP-FormulaNet-S",
    "PP-FormulaNet-L",
];
const MAX_OCR_EVENTS: usize = 256;
const OCR_RUNTIME_PROBE_CACHE_SCHEMA: u32 = 1;
const OCR_RUNTIME_PROBE_CACHE_FILE: &str = "runtime-probe-cache.json";
const ACTIVE_THEME_FILE: &str = "active-theme.txt";
const THEME_CHANGED_EVENT: &str = "visualtex-theme-changed";
const MAIN_WINDOW_SIZE_FILE: &str = "main-window-size.json";
const MAIN_WINDOW_MODE_SIZES_FILE: &str = "main-window-mode-sizes-v1.json";
const DEFAULT_NORMAL_WINDOW_WIDTH: f64 = 1240.0;
const DEFAULT_NORMAL_WINDOW_HEIGHT: f64 = 820.0;
const DEFAULT_KEYPAD_WINDOW_WIDTH: f64 = 642.0;
const DEFAULT_KEYPAD_WINDOW_HEIGHT: f64 = 345.0;
const MIN_MAIN_WINDOW_WIDTH: f64 = 500.0;
const MIN_MAIN_WINDOW_HEIGHT: f64 = 300.0;
const MAX_MAIN_WINDOW_WIDTH: f64 = 4000.0;
const MAX_MAIN_WINDOW_HEIGHT: f64 = 3000.0;
static MAIN_WINDOW_SIZE_WRITE_GENERATION: AtomicU64 = AtomicU64::new(0);
static MAIN_WINDOW_KEYPAD_MODE: AtomicBool = AtomicBool::new(false);

fn normalize_app_theme(theme: &str) -> &'static str {
    match theme.trim() {
        "beige" => "beige",
        "dark" => "dark",
        "purple" => "purple",
        "green" => "green",
        "codex" => "codex",
        "notion" => "notion",
        "one" => "one",
        "proof" => "proof",
        "raycast" => "raycast",
        "rose-pine" => "rose-pine",
        "solarized" => "solarized",
        "vercel" => "vercel",
        "vscode-plus" => "vscode-plus",
        "xcode" => "xcode",
        "custom" => "custom",
        _ => "light",
    }
}

pub(crate) fn persisted_app_theme(app: &AppHandle) -> String {
    let Ok(app_data_dir) = app.path().app_data_dir() else {
        return "light".to_string();
    };
    fs::read_to_string(app_data_dir.join(ACTIVE_THEME_FILE))
        .map(|theme| normalize_app_theme(&theme).to_string())
        .unwrap_or_else(|_| "light".to_string())
}

#[tauri::command]
fn set_app_theme(app: AppHandle, theme: String) -> Result<String, String> {
    let normalized = normalize_app_theme(&theme).to_string();
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|error| error.to_string())?;
    fs::create_dir_all(&app_data_dir).map_err(|error| error.to_string())?;
    fs::write(app_data_dir.join(ACTIVE_THEME_FILE), &normalized)
        .map_err(|error| error.to_string())?;
    app.emit(THEME_CHANGED_EVENT, normalized.clone())
        .map_err(|error| error.to_string())?;
    Ok(normalized)
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ConfigurationWindowSize {
    width: f64,
    height: f64,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct AppWindowConfiguration {
    main: Option<ConfigurationWindowSize>,
    office_editor: Option<ConfigurationWindowSize>,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct MainWindowModeSizes {
    normal: ConfigurationWindowSize,
    keypad: ConfigurationWindowSize,
}

fn normalize_main_window_size(width: f64, height: f64) -> ConfigurationWindowSize {
    let width = if width.is_finite() {
        width
    } else {
        DEFAULT_NORMAL_WINDOW_WIDTH
    };
    let height = if height.is_finite() {
        height
    } else {
        DEFAULT_NORMAL_WINDOW_HEIGHT
    };
    ConfigurationWindowSize {
        width: width.clamp(MIN_MAIN_WINDOW_WIDTH, MAX_MAIN_WINDOW_WIDTH),
        height: height.clamp(MIN_MAIN_WINDOW_HEIGHT, MAX_MAIN_WINDOW_HEIGHT),
    }
}

fn main_window_size_path(app: &AppHandle) -> Result<PathBuf, String> {
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data_dir.join(MAIN_WINDOW_SIZE_FILE))
}

fn main_window_mode_sizes_path(app: &AppHandle) -> Result<PathBuf, String> {
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data_dir.join(MAIN_WINDOW_MODE_SIZES_FILE))
}

fn default_normal_window_size() -> ConfigurationWindowSize {
    normalize_main_window_size(DEFAULT_NORMAL_WINDOW_WIDTH, DEFAULT_NORMAL_WINDOW_HEIGHT)
}

fn default_keypad_window_size() -> ConfigurationWindowSize {
    normalize_main_window_size(DEFAULT_KEYPAD_WINDOW_WIDTH, DEFAULT_KEYPAD_WINDOW_HEIGHT)
}

fn read_main_window_size(app: &AppHandle) -> Option<ConfigurationWindowSize> {
    let path = main_window_size_path(app).ok()?;
    let bytes = fs::read(path).ok()?;
    let size: ConfigurationWindowSize = serde_json::from_slice(&bytes).ok()?;
    Some(normalize_main_window_size(size.width, size.height))
}

fn current_main_window_size(app: &AppHandle) -> Option<ConfigurationWindowSize> {
    let window = app.get_webview_window("main")?;
    let physical = window.inner_size().ok()?;
    let scale_factor = window.scale_factor().ok()?.max(0.1);
    Some(normalize_main_window_size(
        physical.width as f64 / scale_factor,
        physical.height as f64 / scale_factor,
    ))
}

fn read_main_window_mode_sizes(app: &AppHandle) -> MainWindowModeSizes {
    if let Ok(path) = main_window_mode_sizes_path(app) {
        if let Ok(bytes) = fs::read(path) {
            if let Ok(stored) = serde_json::from_slice::<MainWindowModeSizes>(&bytes) {
                return MainWindowModeSizes {
                    normal: normalize_main_window_size(stored.normal.width, stored.normal.height),
                    keypad: normalize_main_window_size(stored.keypad.width, stored.keypad.height),
                };
            }
        }
    }

    MainWindowModeSizes {
        normal: read_main_window_size(app).unwrap_or_else(default_normal_window_size),
        keypad: default_keypad_window_size(),
    }
}

fn write_main_window_mode_sizes(
    app: &AppHandle,
    sizes: MainWindowModeSizes,
) -> Result<(), String> {
    let path = main_window_mode_sizes_path(app)?;
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    let bytes = serde_json::to_vec_pretty(&sizes).map_err(|error| error.to_string())?;
    fs::write(path, bytes).map_err(|error| error.to_string())
}

fn write_main_window_profile_size(
    app: &AppHandle,
    keypad: bool,
    size: ConfigurationWindowSize,
) -> Result<(), String> {
    let normalized = normalize_main_window_size(size.width, size.height);
    let mut profiles = read_main_window_mode_sizes(app);
    if keypad {
        profiles.keypad = normalized;
    } else {
        profiles.normal = normalized;
        write_main_window_size(app, normalized)?;
    }
    write_main_window_mode_sizes(app, profiles)
}

fn write_main_window_size(app: &AppHandle, size: ConfigurationWindowSize) -> Result<(), String> {
    let path = main_window_size_path(app)?;
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    let bytes = serde_json::to_vec_pretty(&size).map_err(|error| error.to_string())?;
    fs::write(path, bytes).map_err(|error| error.to_string())
}

fn schedule_persist_main_window_size(app: &AppHandle, physical_width: u32, physical_height: u32) {
    if physical_width == 0 || physical_height == 0 {
        return;
    }
    let generation = MAIN_WINDOW_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst) + 1;
    let app = app.clone();
    std::thread::spawn(move || {
        std::thread::sleep(std::time::Duration::from_millis(250));
        if MAIN_WINDOW_SIZE_WRITE_GENERATION.load(Ordering::SeqCst) != generation {
            return;
        }
        let Some(size) = current_main_window_size(&app) else {
            return;
        };
        let keypad = MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst);
        if let Err(error) = write_main_window_profile_size(&app, keypad, size) {
            eprintln!("Unable to persist the VisualTeX main window size: {error}");
        }
    });
}

fn restore_main_window_size(app: &AppHandle) -> Result<(), String> {
    let profiles = read_main_window_mode_sizes(app);
    write_main_window_mode_sizes(app, profiles)?;
    write_main_window_size(app, profiles.normal)?;
    MAIN_WINDOW_KEYPAD_MODE.store(false, Ordering::SeqCst);
    let Some(window) = app.get_webview_window("main") else {
        return Ok(());
    };
    window
        .set_size(tauri::LogicalSize::new(
            profiles.normal.width,
            profiles.normal.height,
        ))
        .map_err(|error| error.to_string())?;
    window.center().map_err(|error| error.to_string())
}

#[tauri::command]
fn switch_main_window_mode(
    app: AppHandle,
    keypad: bool,
) -> Result<ConfigurationWindowSize, String> {
    let outgoing_keypad = MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst);
    if let Some(current) = current_main_window_size(&app) {
        write_main_window_profile_size(&app, outgoing_keypad, current)?;
    }

    let profiles = read_main_window_mode_sizes(&app);
    let target = if keypad {
        profiles.keypad
    } else {
        profiles.normal
    };

    MAIN_WINDOW_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst);
    MAIN_WINDOW_KEYPAD_MODE.store(keypad, Ordering::SeqCst);

    if let Some(window) = app.get_webview_window("main") {
        if window.is_maximized().unwrap_or(false) {
            window.unmaximize().map_err(|error| error.to_string())?;
        }
        window
            .set_size(tauri::LogicalSize::new(target.width, target.height))
            .map_err(|error| error.to_string())?;
    }

    Ok(target)
}

#[tauri::command]
fn get_app_window_configuration(app: AppHandle) -> Result<AppWindowConfiguration, String> {
    #[cfg(target_os = "macos")]
    let office_editor = office::macos_offline::configuration_office_editor_window_size(&app)
        .map(|(width, height)| ConfigurationWindowSize { width, height });
    #[cfg(not(target_os = "macos"))]
    let office_editor = None;

    let profiles = read_main_window_mode_sizes(&app);
    let main = if MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst) {
        Some(profiles.normal)
    } else {
        current_main_window_size(&app).or(Some(profiles.normal))
    };

    Ok(AppWindowConfiguration {
        main,
        office_editor,
    })
}

#[tauri::command]
fn apply_app_window_configuration(
    app: AppHandle,
    configuration: AppWindowConfiguration,
) -> Result<AppWindowConfiguration, String> {
    if let Some(requested) = configuration.main {
        let size = normalize_main_window_size(requested.width, requested.height);
        write_main_window_profile_size(&app, false, size)?;
        if !MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst) {
            if let Some(window) = app.get_webview_window("main") {
                window
                    .set_size(tauri::LogicalSize::new(size.width, size.height))
                    .map_err(|error| error.to_string())?;
                window.center().map_err(|error| error.to_string())?;
            }
        }
    }

    #[cfg(target_os = "macos")]
    if let Some(requested) = configuration.office_editor {
        office::macos_offline::apply_configuration_office_editor_window_size(
            &app,
            requested.width,
            requested.height,
        )?;
    }

    get_app_window_configuration(app)
}

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
}

impl Default for OcrState {
    fn default() -> Self {
        Self {
            worker: Arc::new(Mutex::new(None)),
            worker_pid: Arc::new(AtomicU32::new(0)),
            cancel_generation: Arc::new(AtomicU64::new(0)),
            runtime_status: Arc::new(Mutex::new(None)),
            events: OcrEventBus::default(),
        }
    }
}

fn is_final_ocr_state_owner(worker: &Arc<Mutex<Option<OcrWorker>>>) -> bool {
    Arc::strong_count(worker) == 1
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
            ocr_offline::install_optional_model_pack(&package_path, &paths.root)?;
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
            ocr_offline::remove_optional_model(&paths.root, &model)?;
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

struct OcrWorker {
    child: Child,
    stdin: BufWriter<ChildStdin>,
    stdout: BufReader<ChildStdout>,
    pid_state: Arc<AtomicU32>,
    log_path: PathBuf,
    loaded_model: Option<String>,
}

impl Drop for OcrWorker {
    fn drop(&mut self) {
        let pid = self.child.id();
        let _ = self.child.kill();
        let _ = self.child.wait();
        let _ = self
            .pid_state
            .compare_exchange(pid, 0, Ordering::SeqCst, Ordering::SeqCst);
    }
}

fn read_worker_json<R: BufRead>(
    reader: &mut R,
    closed_message: &str,
    response_name: &str,
) -> Result<Value, String> {
    loop {
        let mut bytes = Vec::new();
        let count = reader
            .read_until(b'\n', &mut bytes)
            .map_err(|error| format!("Unable to read {response_name}: {error}"))?;
        if count == 0 {
            return Err(closed_message.to_string());
        }

        while bytes
            .last()
            .is_some_and(|byte| matches!(*byte, b'\n' | b'\r'))
        {
            bytes.pop();
        }
        if bytes.iter().all(u8::is_ascii_whitespace) {
            continue;
        }

        let first_non_whitespace = bytes
            .iter()
            .copied()
            .find(|byte| !byte.is_ascii_whitespace());
        if first_non_whitespace != Some(b'{') {
            // Native dependencies occasionally write diagnostics to stdout.
            // Ignore those lines so they cannot corrupt the JSON protocol.
            continue;
        }

        match serde_json::from_slice(&bytes) {
            Ok(value) => return Ok(value),
            Err(error) => {
                // This fallback prevents a legacy Windows code page from
                // crashing the reader. New workers always emit ASCII-safe JSON.
                let output = String::from_utf8_lossy(&bytes);
                return serde_json::from_str(output.trim()).map_err(|lossy_error| {
                    format!(
                        "{response_name} returned invalid JSON: {error}; UTF-8-lossy parse: {lossy_error}; output={output:?}"
                    )
                });
            }
        }
    }
}

impl OcrWorker {
    fn worker_failure(&mut self, message: impl AsRef<str>) -> String {
        let status = match self.child.try_wait() {
            Ok(Some(status)) => status.to_string(),
            Ok(None) => "still running".to_string(),
            Err(error) => format!("status unavailable: {error}"),
        };
        format!(
            "{}; worker_status={status}; log={}",
            message.as_ref(),
            self.log_path.display()
        )
    }

    fn send(&mut self, app: &AppHandle, payload: &Value) -> Result<Value, String> {
        if let Some(status) = self.child.try_wait().map_err(|error| error.to_string())? {
            return Err(format!(
                "OCR worker exited unexpectedly: {status}; log={}",
                self.log_path.display()
            ));
        }

        serde_json::to_writer(&mut self.stdin, payload)
            .map_err(|error| format!("Unable to encode OCR request: {error}"))?;
        self.stdin
            .write_all(b"\n")
            .map_err(|error| self.worker_failure(format!("Unable to send OCR request: {error}")))?;
        self.stdin.flush().map_err(|error| {
            self.worker_failure(format!("Unable to flush OCR request: {error}"))
        })?;

        loop {
            let response = read_worker_json(
                &mut self.stdout,
                "OCR worker closed its output stream",
                "OCR response",
            )
            .map_err(|error| self.worker_failure(error))?;
            if response.get("event").and_then(Value::as_str) == Some("progress") {
                let _ = app.emit("ocr-recognition-progress", &response);
                quick_ocr::handle_ocr_progress(app, &response);
                if let Some(state) = app.try_state::<OcrState>() {
                    state.events.publish("ocr-recognition-progress", &response);
                }
                continue;
            }
            return Ok(response);
        }
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
    installed: bool,
    python_path: Option<String>,
    python_version: Option<String>,
    paddle_version: Option<String>,
    paddleocr_version: Option<String>,
    runtime_path: String,
    offline_bundle_available: bool,
    installed_models: Vec<String>,
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
struct OcrFormulaResult {
    latex: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrRecognitionResult {
    model: String,
    elapsed_ms: u64,
    processed_width: u32,
    processed_height: u32,
    background_inverted: bool,
    background_luminance: f64,
    formulas: Vec<OcrFormulaResult>,
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

fn runtime_paths(app: &AppHandle) -> Result<RuntimePaths, String> {
    let root = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve application data directory: {error}"))?
        .join("ocr-runtime");
    let offline_python = if cfg!(windows) {
        root.join("python").join("python.exe")
    } else {
        root.join("python").join("bin").join("python3")
    };
    let legacy_python = if cfg!(windows) {
        root.join("venv").join("Scripts").join("python.exe")
    } else {
        root.join("venv").join("bin").join("python")
    };
    let python = if offline_python.exists() || !legacy_python.exists() {
        offline_python
    } else {
        legacy_python
    };
    Ok(RuntimePaths {
        root: root.clone(),
        python,
        input: root.join("input"),
        processed: root.join("processed"),
        logs: root.join("logs"),
        cache: root.join("cache"),
        temp: root.join("tmp"),
    })
}

fn worker_script_path(app: &AppHandle) -> Result<PathBuf, String> {
    if let Ok(path) = app.path().resolve("ocr/worker.py", BaseDirectory::Resource) {
        if path.exists() {
            return Ok(path);
        }
    }

    #[cfg(debug_assertions)]
    {
        let development_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("ocr")
            .join("worker.py");
        if development_path.exists() {
            return Ok(development_path);
        }
    }

    Err("Unable to locate bundled OCR worker.py".to_string())
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

fn tail_text(value: &str, max_chars: usize) -> String {
    let total = value.chars().count();
    if total <= max_chars {
        return value.to_string();
    }
    value.chars().skip(total - max_chars).collect()
}

fn command_output(command: &mut Command, label: &str) -> Result<String, String> {
    let output = command
        .output()
        .map_err(|error| format!("Unable to start {label}: {error}"))?;
    if output.status.success() {
        return Ok(String::from_utf8_lossy(&output.stdout).trim().to_string());
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    let combined = format!("{stdout}\n{stderr}");
    let tail = tail_text(&combined, 6000);
    Err(format!(
        "{label} failed with {}:\n{}",
        output.status,
        tail.trim()
    ))
}

fn validate_runtime_probe(probe: &RuntimeProbe) -> Result<(), String> {
    if probe.paddle_version != PADDLE_VERSION || probe.paddleocr_version != PADDLEOCR_VERSION {
        return Err(format!(
            "OCR runtime version mismatch: PaddlePaddle {}, PaddleOCR {}; expected {} and {}",
            probe.paddle_version, probe.paddleocr_version, PADDLE_VERSION, PADDLEOCR_VERSION
        ));
    }
    Ok(())
}

fn python_runtime_signature(path: &PathBuf) -> Result<(u64, u64), String> {
    let metadata = fs::metadata(path)
        .map_err(|error| format!("Unable to inspect OCR Python runtime: {error}"))?;
    let modified_ms = metadata
        .modified()
        .ok()
        .and_then(|value| value.duration_since(UNIX_EPOCH).ok())
        .map(|duration| duration.as_millis() as u64)
        .unwrap_or_default();
    Ok((metadata.len(), modified_ms))
}

fn runtime_probe_cache_path(paths: &RuntimePaths) -> PathBuf {
    paths.root.join(OCR_RUNTIME_PROBE_CACHE_FILE)
}

fn read_persisted_runtime_probe(paths: &RuntimePaths) -> Option<RuntimeProbe> {
    let content = fs::read(runtime_probe_cache_path(paths)).ok()?;
    let cache: RuntimeProbeCache = serde_json::from_slice(&content).ok()?;
    if cache.schema_version != OCR_RUNTIME_PROBE_CACHE_SCHEMA
        || cache.python_path != paths.python.display().to_string()
    {
        return None;
    }
    let (python_size, python_modified_ms) = python_runtime_signature(&paths.python).ok()?;
    if cache.python_size != python_size || cache.python_modified_ms != python_modified_ms {
        return None;
    }
    validate_runtime_probe(&cache.probe).ok()?;
    Some(cache.probe)
}

fn write_persisted_runtime_probe(paths: &RuntimePaths, probe: &RuntimeProbe) -> Result<(), String> {
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    let (python_size, python_modified_ms) = python_runtime_signature(&paths.python)?;
    let cache = RuntimeProbeCache {
        schema_version: OCR_RUNTIME_PROBE_CACHE_SCHEMA,
        python_path: paths.python.display().to_string(),
        python_size,
        python_modified_ms,
        probe: probe.clone(),
    };
    let destination = runtime_probe_cache_path(paths);
    let temporary = paths.root.join(format!(
        ".{OCR_RUNTIME_PROBE_CACHE_FILE}.{}.tmp",
        std::process::id()
    ));
    fs::write(
        &temporary,
        serde_json::to_vec_pretty(&cache)
            .map_err(|error| format!("Unable to encode OCR runtime cache: {error}"))?,
    )
    .map_err(|error| format!("Unable to write OCR runtime cache: {error}"))?;
    if destination.exists() {
        fs::remove_file(&destination).ok();
    }
    fs::rename(&temporary, &destination)
        .map_err(|error| format!("Unable to activate OCR runtime cache: {error}"))
}

fn probe_runtime(paths: &RuntimePaths) -> Result<RuntimeProbe, String> {
    if !paths.python.exists() {
        return Err("OCR offline runtime is not installed".to_string());
    }
    let script = r#"import json, platform; import paddle; import paddleocr; import tokenizers, imagesize, ftfy, wand; from importlib.metadata import version; from paddleocr import FormulaRecognition; print(json.dumps({'python_version': platform.python_version(), 'paddle_version': paddle.__version__, 'paddleocr_version': version('paddleocr')}))"#;
    let output = command_output(
        Command::new(&paths.python).arg("-c").arg(script),
        "OCR runtime verification",
    )?;
    let probe: RuntimeProbe = serde_json::from_str(&output)
        .map_err(|error| format!("OCR runtime returned invalid version information: {error}"))?;
    validate_runtime_probe(&probe)?;
    if let Err(error) = write_persisted_runtime_probe(paths, &probe) {
        eprintln!("Unable to persist OCR runtime verification: {error}");
    }
    Ok(probe)
}

fn get_runtime_status_inner(
    app: &AppHandle,
    force_refresh: bool,
) -> Result<OcrRuntimeStatus, String> {
    let paths = runtime_paths(app)?;
    let offline_bundle_available = ocr_offline::bundle_available(app);
    let installed_models = ocr_offline::installed_models(&paths.root);
    let default_model_installed = installed_models
        .iter()
        .any(|model| model == ocr_offline::OFFLINE_DEFAULT_MODEL);
    let structural_error = if !paths.python.is_file() {
        Some("OCR offline runtime is not installed".to_string())
    } else if !default_model_installed {
        Some("The bundled PP-FormulaNet M model is not installed".to_string())
    } else {
        None
    };

    if structural_error.is_none() && !force_refresh {
        let cached_probe = read_persisted_runtime_probe(&paths);
        return Ok(OcrRuntimeStatus {
            installed: true,
            python_path: Some(paths.python.display().to_string()),
            python_version: cached_probe
                .as_ref()
                .map(|probe| probe.python_version.clone()),
            paddle_version: cached_probe
                .as_ref()
                .map(|probe| probe.paddle_version.clone()),
            paddleocr_version: cached_probe
                .as_ref()
                .map(|probe| probe.paddleocr_version.clone()),
            runtime_path: paths.root.display().to_string(),
            offline_bundle_available,
            installed_models,
            default_model: ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: if cached_probe.is_some() {
                "PaddleOCR formula runtime is ready; the default M model can be reused in memory"
                    .to_string()
            } else {
                "OCR runtime files are ready; the default M model will be verified during background preloading"
                    .to_string()
            },
        });
    }

    let probe_result = structural_error.map_or_else(|| probe_runtime(&paths), Err);
    match probe_result {
        Ok(probe) => Ok(OcrRuntimeStatus {
            installed: true,
            python_path: Some(paths.python.display().to_string()),
            python_version: Some(probe.python_version),
            paddle_version: Some(probe.paddle_version),
            paddleocr_version: Some(probe.paddleocr_version),
            runtime_path: paths.root.display().to_string(),
            offline_bundle_available,
            installed_models,
            default_model: ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: "PaddleOCR formula runtime is ready for offline recognition".to_string(),
        }),
        Err(error) => Ok(OcrRuntimeStatus {
            installed: false,
            python_path: paths
                .python
                .exists()
                .then(|| paths.python.display().to_string()),
            python_version: None,
            paddle_version: None,
            paddleocr_version: None,
            runtime_path: paths.root.display().to_string(),
            offline_bundle_available,
            installed_models,
            default_model: ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: if offline_bundle_available {
                format!("Offline OCR package is ready to install. Current runtime: {error}")
            } else {
                error
            },
        }),
    }
}

fn read_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
) -> Result<Option<OcrRuntimeStatus>, String> {
    cache
        .lock()
        .map(|status| status.clone())
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())
}

fn write_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    status: Option<OcrRuntimeStatus>,
) -> Result<(), String> {
    let mut guard = cache
        .lock()
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())?;
    *guard = status;
    Ok(())
}

fn cleanup_worker_temp(paths: &RuntimePaths) -> Result<(), String> {
    if paths.temp.exists() {
        fs::remove_dir_all(&paths.temp)
            .map_err(|error| format!("Unable to clean OCR temporary files: {error}"))?;
    }
    fs::create_dir_all(&paths.temp)
        .map_err(|error| format!("Unable to create OCR temporary directory: {error}"))
}

fn terminate_worker_process(worker_pid: &AtomicU32) -> Result<bool, String> {
    let pid = worker_pid.swap(0, Ordering::SeqCst);
    if pid == 0 {
        return Ok(false);
    }

    #[cfg(unix)]
    {
        let result = unsafe { libc::kill(pid as i32, libc::SIGKILL) };
        if result == 0 {
            return Ok(true);
        }
        let error = std::io::Error::last_os_error();
        if error.raw_os_error() == Some(libc::ESRCH) {
            return Ok(false);
        }
        return Err(format!("Unable to terminate OCR worker {pid}: {error}"));
    }

    #[cfg(windows)]
    {
        let status = Command::new("taskkill")
            .arg("/PID")
            .arg(pid.to_string())
            .arg("/T")
            .arg("/F")
            .status()
            .map_err(|error| format!("Unable to terminate OCR worker {pid}: {error}"))?;
        return Ok(status.success());
    }

    #[allow(unreachable_code)]
    Ok(false)
}

fn stop_worker(
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
) -> Result<(), String> {
    let terminate_result = terminate_worker_process(worker_pid);
    if let Ok(mut guard) = worker.lock() {
        guard.take();
    }
    terminate_result.map(|_| ())
}

fn install_runtime_inner(
    app: &AppHandle,
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
) -> Result<OcrRuntimeStatus, String> {
    stop_worker(worker, worker_pid)?;
    let paths = runtime_paths(app)?;
    ocr_offline::install_bundle(app, &paths.root, |stage, percent, message, detail| {
        emit_progress(app, stage, percent, message, detail);
    })?;

    emit_progress(app, "verify", 97, "正在验证离线 PP-FormulaNet 接口", None);
    let status = get_runtime_status_inner(app, true)?;
    if !status.installed {
        return Err(status.message);
    }
    if !status
        .installed_models
        .iter()
        .any(|model| model == ocr_offline::OFFLINE_DEFAULT_MODEL)
    {
        return Err("The bundled PP-FormulaNet M model was not installed".to_string());
    }
    emit_progress(
        app,
        "complete",
        100,
        "OCR 离线运行环境安装完成",
        Some("Python、PaddleOCR 与默认 M 模型均已内置，无需联网".to_string()),
    );
    Ok(status)
}

fn spawn_worker(
    app: &AppHandle,
    paths: &RuntimePaths,
    worker_pid: Arc<AtomicU32>,
) -> Result<OcrWorker, String> {
    let script = worker_script_path(app)?;
    fs::create_dir_all(&paths.logs)
        .map_err(|error| format!("Unable to create OCR log directory: {error}"))?;
    fs::create_dir_all(&paths.cache)
        .map_err(|error| format!("Unable to create OCR cache directory: {error}"))?;
    fs::create_dir_all(&paths.temp)
        .map_err(|error| format!("Unable to create OCR temporary directory: {error}"))?;
    let log_path = paths.logs.join("worker.log");
    let mut log_file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log_path)
        .map_err(|error| format!("Unable to open OCR worker log: {error}"))?;
    writeln!(
        log_file,
        "\n===== VisualTeX OCR worker start: pid pending, unix_ms={} =====",
        SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|duration| duration.as_millis())
            .unwrap_or_default()
    )
    .map_err(|error| format!("Unable to initialize OCR worker log: {error}"))?;
    let log_file_error = log_file
        .try_clone()
        .map_err(|error| format!("Unable to clone OCR log handle: {error}"))?;

    let mut child = Command::new(&paths.python)
        .arg(&script)
        .env("PYTHONUNBUFFERED", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .env("VISUALTEX_PARENT_PID", std::process::id().to_string())
        .env("VISUALTEX_OFFLINE_OCR", "1")
        .env("HF_HUB_OFFLINE", "1")
        .env("TRANSFORMERS_OFFLINE", "1")
        .env("MODELSCOPE_OFFLINE", "1")
        .env("PADDLE_PDX_MODEL_SOURCE", "BOS")
        .env("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
        .env("PADDLE_PDX_CACHE_HOME", paths.cache.join("paddlex"))
        .env("PADDLE_HOME", paths.cache.join("paddle"))
        .env("XDG_CACHE_HOME", &paths.cache)
        .env("TMPDIR", &paths.temp)
        .env("TMP", &paths.temp)
        .env("TEMP", &paths.temp)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::from(log_file_error))
        .spawn()
        .map_err(|error| format!("Unable to start OCR worker: {error}"))?;
    worker_pid.store(child.id(), Ordering::SeqCst);

    let stdin = child
        .stdin
        .take()
        .ok_or_else(|| "OCR worker stdin is unavailable".to_string())?;
    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "OCR worker stdout is unavailable".to_string())?;
    let mut worker = OcrWorker {
        child,
        stdin: BufWriter::new(stdin),
        stdout: BufReader::new(stdout),
        pid_state: worker_pid,
        log_path: log_path.clone(),
        loaded_model: None,
    };

    let ready = read_worker_json(
        &mut worker.stdout,
        "OCR worker closed before sending its ready signal",
        "OCR worker ready signal",
    )
    .map_err(|error| format!("{error}; log={}", log_path.display()))?;
    if ready.get("event").and_then(Value::as_str) != Some("ready") {
        return Err(format!("Unexpected OCR worker ready response: {ready}"));
    }
    Ok(worker)
}

fn prewarm_model_inner(
    app: &AppHandle,
    worker_state: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
    runtime_status: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    model: &str,
) -> Result<(), String> {
    if !ALLOWED_MODELS.contains(&model) {
        return Err(format!("Unsupported PP-FormulaNet model: {model}"));
    }

    let paths = runtime_paths(app)?;
    let installed_models = ocr_offline::installed_models(&paths.root);
    if !paths.python.is_file() || !installed_models.iter().any(|item| item == model) {
        return Err(format!(
            "The offline model pack for {model} is not installed"
        ));
    }

    let mut guard = worker_state
        .lock()
        .map_err(|_| "OCR worker lock is poisoned".to_string())?;
    if guard
        .as_ref()
        .and_then(|worker| worker.loaded_model.as_deref())
        == Some(model)
    {
        return Ok(());
    }

    if guard
        .as_ref()
        .and_then(|worker| worker.loaded_model.as_deref())
        .is_some_and(|loaded| loaded != model)
    {
        guard.take();
    }
    if guard.is_none() {
        *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
    }

    let request_id = format!(
        "prewarm-{}-{}",
        std::process::id(),
        SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|duration| duration.as_millis())
            .unwrap_or_default()
    );
    let payload = json!({
        "id": request_id,
        "action": "warmup",
        "model": model,
        "device": "cpu"
    });
    let response = match guard
        .as_mut()
        .ok_or_else(|| "OCR worker failed to start".to_string())?
        .send(app, &payload)
    {
        Ok(response) => response,
        Err(error) => {
            guard.take();
            return Err(error);
        }
    };
    if response.get("ok").and_then(Value::as_bool) != Some(true) {
        let error = response
            .get("error")
            .and_then(Value::as_str)
            .unwrap_or("OCR model preloading failed");
        guard.take();
        return Err(error.to_string());
    }

    let probe: RuntimeProbe = serde_json::from_value(response.clone())
        .map_err(|error| format!("OCR prewarm returned invalid runtime versions: {error}"))?;
    validate_runtime_probe(&probe)?;
    if let Err(error) = write_persisted_runtime_probe(&paths, &probe) {
        eprintln!("Unable to persist OCR prewarm verification: {error}");
    }
    if let Some(worker) = guard.as_mut() {
        worker.loaded_model = Some(model.to_string());
    }

    let status = OcrRuntimeStatus {
        installed: true,
        python_path: Some(paths.python.display().to_string()),
        python_version: Some(probe.python_version),
        paddle_version: Some(probe.paddle_version),
        paddleocr_version: Some(probe.paddleocr_version),
        runtime_path: paths.root.display().to_string(),
        offline_bundle_available: ocr_offline::bundle_available(app),
        installed_models,
        default_model: ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
        message: format!("PaddleOCR formula runtime and {model} are ready in memory"),
    };
    write_cached_runtime_status(runtime_status, Some(status))?;
    if let Some(elapsed_ms) = response.get("elapsed_ms").and_then(Value::as_u64) {
        eprintln!("VisualTeX OCR model {model} preloaded in {elapsed_ms} ms");
    }
    Ok(())
}

fn run_recognition(
    app: &AppHandle,
    worker_state: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
    cancel_generation: &Arc<AtomicU64>,
    runtime_status: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    request: OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    let request_generation = cancel_generation.load(Ordering::SeqCst);

    if request.bytes.is_empty() {
        return Err("The selected image is empty".to_string());
    }
    if request.bytes.len() > MAX_IMAGE_BYTES {
        return Err("The image is larger than the 20 MB limit".to_string());
    }
    if !ALLOWED_MODELS.contains(&request.model.as_str()) {
        return Err(format!(
            "Unsupported PP-FormulaNet model: {}",
            request.model
        ));
    }

    let paths = runtime_paths(app)?;
    if let Some(status) = read_cached_runtime_status(runtime_status)? {
        if !status.installed {
            return Err(format!("OCR runtime is not installed: {}", status.message));
        }
    }
    if !paths.python.exists() {
        return Err("OCR runtime is not installed: Python executable is missing".to_string());
    }
    fs::create_dir_all(&paths.input)
        .map_err(|error| format!("Unable to create OCR input directory: {error}"))?;
    fs::create_dir_all(&paths.processed)
        .map_err(|error| format!("Unable to create OCR processed directory: {error}"))?;

    let extension = request
        .extension
        .trim_start_matches('.')
        .to_ascii_lowercase();
    let allowed_extensions = ["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"];
    if !allowed_extensions.contains(&extension.as_str()) {
        return Err(format!("Unsupported image type: .{extension}"));
    }

    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|error| error.to_string())?
        .as_nanos();
    let request_id = format!("{}-{nonce}", std::process::id());
    let input_path = paths.input.join(format!("{request_id}.{extension}"));
    let processed_path = paths.processed.join(format!("{request_id}.png"));
    fs::write(&input_path, &request.bytes)
        .map_err(|error| format!("Unable to save OCR input image: {error}"))?;

    let payload = json!({
        "id": request_id,
        "action": "recognize",
        "image_path": input_path,
        "processed_path": processed_path,
        "model": request.model,
        "device": "cpu"
    });

    if cancel_generation.load(Ordering::SeqCst) != request_generation {
        let _ = fs::remove_file(&input_path);
        return Err(OCR_CANCELLED.to_string());
    }

    let response_result = (|| -> Result<Value, String> {
        let mut guard = worker_state
            .lock()
            .map_err(|_| "OCR worker lock is poisoned".to_string())?;

        let should_restart_for_model = guard
            .as_ref()
            .and_then(|worker| worker.loaded_model.as_deref())
            .is_some_and(|loaded_model| loaded_model != request.model);
        if should_restart_for_model {
            guard.take();
        }
        if guard.is_none() {
            *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
        }

        let first_result = guard
            .as_mut()
            .ok_or_else(|| "OCR worker failed to start".to_string())?
            .send(app, &payload);
        match first_result {
            Ok(response) => {
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }
                if response.get("ok").and_then(Value::as_bool) == Some(true) {
                    if let Some(worker) = guard.as_mut() {
                        worker.loaded_model = Some(request.model.clone());
                    }
                }
                Ok(response)
            }
            Err(first_error) => {
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }

                guard.take();
                *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
                let response = guard
                    .as_mut()
                    .ok_or_else(|| "OCR worker failed to restart".to_string())?
                    .send(app, &payload)
                    .map_err(|second_error| {
                        format!(
                            "OCR worker failed twice. First: {first_error}. Second: {second_error}"
                        )
                    })?;
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }
                if response.get("ok").and_then(Value::as_bool) == Some(true) {
                    if let Some(worker) = guard.as_mut() {
                        worker.loaded_model = Some(request.model.clone());
                    }
                }
                Ok(response)
            }
        }
    })();

    let _ = fs::remove_file(&input_path);
    let _ = fs::remove_file(&processed_path);
    let response_value = response_result?;

    let response: WorkerRecognitionResponse = serde_json::from_value(response_value)
        .map_err(|error| format!("Unable to decode OCR result: {error}"))?;
    if !response.ok {
        let mut message = response
            .error
            .unwrap_or_else(|| "PP-FormulaNet recognition failed".to_string());
        if let Some(details) = response.details {
            if !details.trim().is_empty() {
                let detail_tail = tail_text(&details, 3000);
                message.push('\n');
                message.push_str(detail_tail.trim());
            }
        }
        return Err(message);
    }

    let formulas = response.formulas.unwrap_or_default();
    if formulas.is_empty() {
        return Err("PP-FormulaNet returned no formulas".to_string());
    }

    Ok(OcrRecognitionResult {
        model: response.model.unwrap_or_else(|| request.model.clone()),
        elapsed_ms: response.elapsed_ms.unwrap_or_default(),
        processed_width: response.processed_width.unwrap_or_default(),
        processed_height: response.processed_height.unwrap_or_default(),
        background_inverted: response.background_inverted.unwrap_or(false),
        background_luminance: response.background_luminance.unwrap_or(255.0),
        formulas,
    })
}

#[tauri::command]
fn copy_png_to_clipboard(data_base64: String) -> Result<(), String> {
    let bytes = BASE64_STANDARD
        .decode(data_base64.trim())
        .map_err(|error| format!("Unable to decode PNG clipboard data: {error}"))?;
    if bytes.len() > 100 * 1024 * 1024 {
        return Err("PNG clipboard image is unexpectedly large".to_string());
    }
    const PNG_SIGNATURE: &[u8; 8] = b"\x89PNG\r\n\x1a\n";
    if !bytes.starts_with(PNG_SIGNATURE) {
        return Err("Clipboard image is not a valid PNG".to_string());
    }

    #[cfg(target_os = "macos")]
    {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|duration| duration.as_nanos())
            .unwrap_or_default();
        let temporary = std::env::temp_dir().join(format!(
            "visualtex-clipboard-{}-{nonce}.png",
            std::process::id()
        ));
        fs::write(&temporary, &bytes)
            .map_err(|error| format!("Unable to prepare PNG clipboard image: {error}"))?;

        let script = r#"on run argv
set pngPath to item 1 of argv
set pngFile to POSIX file pngPath
set the clipboard to (read pngFile as «class PNGf»)
end run"#;
        let output = Command::new("/usr/bin/osascript")
            .arg("-e")
            .arg(script)
            .arg(&temporary)
            .output();
        let _ = fs::remove_file(&temporary);
        let output = output.map_err(|error| format!("Unable to access the macOS clipboard: {error}"))?;
        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
            return Err(if stderr.is_empty() {
                "macOS rejected the PNG clipboard write".to_string()
            } else {
                format!("Unable to copy PNG to the macOS clipboard: {stderr}")
            });
        }
        return Ok(());
    }

    #[cfg(not(target_os = "macos"))]
    {
        Err("PNG clipboard export is currently supported on macOS only".to_string())
    }
}

#[tauri::command]
fn write_export_file(path: String, data_base64: String) -> Result<(), String> {
    let target = PathBuf::from(path.trim());
    if !target.is_absolute() {
        return Err("Export path must be absolute".to_string());
    }
    let extension = target
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    if !matches!(extension.as_str(), "md" | "svg" | "png" | "vtxconfig") {
        return Err("Unsupported export file extension".to_string());
    }
    let parent = target
        .parent()
        .ok_or_else(|| "Export path has no parent directory".to_string())?;
    if !parent.is_dir() {
        return Err("Export directory does not exist".to_string());
    }
    let bytes = BASE64_STANDARD
        .decode(data_base64.trim())
        .map_err(|error| format!("Unable to decode export data: {error}"))?;
    if bytes.len() > 100 * 1024 * 1024 {
        return Err("Export file is unexpectedly large".to_string());
    }

    let file_name = target
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| "Export filename is invalid".to_string())?;
    let temporary = parent.join(format!(".{file_name}.visualtex-{}.tmp", std::process::id()));
    let _ = fs::remove_file(&temporary);
    let write_result = (|| -> Result<(), String> {
        let mut file = OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&temporary)
            .map_err(|error| format!("Unable to create export file: {error}"))?;
        file.write_all(&bytes)
            .and_then(|_| file.sync_all())
            .map_err(|error| format!("Unable to write export file: {error}"))?;
        fs::rename(&temporary, &target)
            .map_err(|error| format!("Unable to finalize export file: {error}"))?;
        Ok(())
    })();
    if write_result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    write_result
}

#[tauri::command]
async fn get_ocr_runtime_status(
    app: AppHandle,
    state: State<'_, OcrState>,
    force_refresh: Option<bool>,
) -> Result<OcrRuntimeStatus, String> {
    state
        .runtime_status(app, force_refresh.unwrap_or(false))
        .await
}

#[tauri::command]
async fn install_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    state.install_runtime(app).await
}

#[tauri::command]
async fn recognize_formula_image(
    app: AppHandle,
    state: State<'_, OcrState>,
    request: OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    state.recognize(app, request).await
}

#[tauri::command]
async fn prewarm_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: String,
) -> Result<(), String> {
    state.prewarm_model(app, model).await
}

#[tauri::command]
fn cancel_ocr_recognition(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.cancel(&app)
}

#[tauri::command]
async fn restart_ocr_worker(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.restart(app).await
}

#[tauri::command]
async fn reset_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    state.reset_runtime(app).await
}

#[tauri::command]
async fn install_optional_ocr_model(
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
async fn remove_optional_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: String,
) -> Result<OcrRuntimeStatus, String> {
    state.remove_optional_model(app, model).await
}

#[cfg(all(target_os = "macos", not(debug_assertions)))]
fn claim_production_visualtex_url_handler() -> Result<(), String> {
    let executable = std::env::current_exe()
        .map_err(|error| format!("Unable to locate the VisualTeX executable: {error}"))?;
    let app_bundle = executable
        .parent()
        .and_then(|path| path.parent())
        .and_then(|path| path.parent())
        .filter(|path| path.extension().and_then(|value| value.to_str()) == Some("app"))
        .ok_or_else(|| {
            format!(
                "Unable to resolve the VisualTeX app bundle from {}",
                executable.display()
            )
        })?;

    let launch_services = "/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister";
    let register_status = Command::new(launch_services)
        .arg("-f")
        .arg(app_bundle)
        .status()
        .map_err(|error| format!("Unable to register VisualTeX with LaunchServices: {error}"))?;
    if !register_status.success() {
        return Err(format!(
            "LaunchServices registration failed with status {register_status}"
        ));
    }

    let script = r#"ObjC.import("CoreServices"); $.LSSetDefaultHandlerForURLScheme($("visualtex"), $("com.visualtex.studio"));"#;
    let handler_status = Command::new("/usr/bin/osascript")
        .args(["-l", "JavaScript", "-e", script])
        .status()
        .map_err(|error| format!("Unable to restore the VisualTeX URL handler: {error}"))?;
    if !handler_status.success() {
        return Err(format!(
            "VisualTeX URL handler restoration failed with status {handler_status}"
        ));
    }

    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    if let Some(status) = office::omml_batch::run_cli_if_requested() {
        std::process::exit(status);
    }
    if let Some(status) = office::macos_offline::run_image_ink_center_cli_if_requested() {
        std::process::exit(status);
    }
    let background_mode = office::background::is_background_mode();
    let application_started_at = std::time::Instant::now();
    let maintenance_install = std::env::args_os().any(|argument| {
        argument
            == std::ffi::OsStr::new(office::macos_offline_installer::MAINTENANCE_INSTALL_ARGUMENT)
    });
    let initial_office_url =
        std::env::args().find(|argument| argument.starts_with("visualtex://office/open?session="));
    let ocr_state = OcrState::default();
    let office_ocr_state = ocr_state.clone();
    let quick_ocr_state = quick_ocr::QuickOcrState::default();
    let app = tauri::Builder::default()
        .plugin(tauri_plugin_single_instance::init(
            |app, arguments, _cwd| {
                if arguments.iter().any(|argument| {
                    argument == office::background::BACKGROUND_ARGUMENT
                        || argument == office::macos_offline_installer::MAINTENANCE_INSTALL_ARGUMENT
                }) {
                    return;
                }
                #[cfg(target_os = "macos")]
                if let Some(url) = arguments
                    .iter()
                    .find(|argument| argument.starts_with("visualtex://office/open?session="))
                {
                    if let Err(error) = office::macos_offline::handle_open_url(app, url) {
                        eprintln!("Unable to open VisualTeX offline Office Session: {error}");
                    }
                    return;
                }
                #[cfg(target_os = "macos")]
                match office::macos_offline::consume_fast_open_request(app) {
                    Ok(true) => return,
                    Ok(false) => {}
                    Err(error) => {
                        eprintln!("Unable to consume VisualTeX Office fast-open request: {error}");
                        return;
                    }
                }
                let _ = office::background::reveal_main_window(app);
            },
        ))
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(ocr_state)
        .manage(quick_ocr_state)
        .setup(move |app| {
            if maintenance_install {
                match office::macos_offline_installer::install(app.handle()) {
                    Ok(status) => {
                        println!(
                            "{}",
                            serde_json::to_string(&status)
                                .unwrap_or_else(|_| "{\"ok\":true}".to_string())
                        );
                        std::process::exit(0);
                    }
                    Err(error) => {
                        eprintln!("VisualTeX native Office installation failed: {error}");
                        std::process::exit(1);
                    }
                }
            }

            #[cfg(not(debug_assertions))]
            {
                #[cfg(target_os = "macos")]
                std::thread::spawn(|| {
                    // LaunchServices repair is maintenance. Let a cold Office
                    // URL hydrate its resident editor before lsregister scans
                    // the application bundle.
                    std::thread::sleep(std::time::Duration::from_secs(2));
                    if let Err(error) = claim_production_visualtex_url_handler() {
                        eprintln!("Unable to claim the production visualtex:// handler: {error}");
                    }
                });
            }

            // Set the application icon while setup is still on AppKit's main
            // thread. Background-agent launches may stay accessory-only for a
            // long time, but the first later transition to a regular app must
            // already have the real VisualTeX Dock icon installed.
            office::background::install_application_icon(app.handle())
                .map_err(std::io::Error::other)?;
            restore_main_window_size(app.handle()).map_err(std::io::Error::other)?;
            if let Err(error) = quick_ocr::initialize(app.handle()) {
                eprintln!("VisualTeX quick OCR initialization warning: {error}");
            }
            let office_state = office::initialize(app.handle(), office_ocr_state.clone())
                .map_err(std::io::Error::other)?;
            if let Err(error) = office::powerpoint_native::start_double_click_monitor(
                app.handle().clone(),
                office_state.powerpoint_interactions.clone(),
            ) {
                eprintln!("Unable to start PowerPoint double-click monitor: {error}");
            }
            app.manage(office_state.clone());
            if initial_office_url.is_some() || background_mode {
                // Hide the desktop shell before WebKit prewarming. Otherwise a
                // cold Office URL can flash the main editor while the resident
                // formula WebViews initialize.
                office::background::hide_main_window(app.handle())
                    .map_err(std::io::Error::other)?;
            }
            if let Err(error) = office::macos_offline::prewarm_office_editor_windows(app.handle()) {
                // Prewarming is an optimization. handle_open_url still creates
                // the fixed host window lazily if WebKit was unavailable here.
                eprintln!("Unable to prewarm VisualTeX Office editors: {error}");
            }
            #[cfg(target_os = "macos")]
            office::macos_offline::start_fast_open_inbox_watcher(app.handle().clone());
            #[cfg(not(target_os = "macos"))]
            office::lifecycle::start(office_state);
            if let Some(url) = initial_office_url.as_deref() {
                office::macos_offline::handle_open_url(app.handle(), url)
                    .map_err(std::io::Error::other)?;
            } else if !background_mode {
                office::background::reveal_main_window(app.handle())
                    .map_err(std::io::Error::other)?;
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            write_export_file,
            copy_png_to_clipboard,
            quick_ocr::capture_quick_ocr_screenshot,
            quick_ocr::wait_for_quick_ocr_system_screenshot,
            quick_ocr::configure_silent_ocr,
            set_app_theme,
            get_app_window_configuration,
            apply_app_window_configuration,
            switch_main_window_mode,
            system_math_glyphs::probe_macos_math_fonts,
            system_math_glyphs::extract_macos_math_glyph,
            get_ocr_runtime_status,
            install_ocr_runtime,
            recognize_formula_image,
            prewarm_ocr_model,
            cancel_ocr_recognition,
            restart_ocr_worker,
            reset_ocr_runtime,
            install_optional_ocr_model,
            remove_optional_ocr_model,
            office::lifecycle::get_office_companion_status,
            office::lifecycle::start_office_companion,
            office::lifecycle::stop_office_companion,
            office::lifecycle::set_office_background_start,
            office::lifecycle::open_word,
            office::lifecycle::open_powerpoint,
            office::macos_offline::get_macos_offline_document_import_request,
            office::macos_offline::report_macos_offline_latex_redraw_stage,
            office::macos_offline::resolve_macos_offline_latex_redraw_font_sizes,
            office::macos_offline::focus_macos_offline_document_import_target,
            office::macos_offline::restore_macos_offline_document_import_window,
            office::macos_offline::get_macos_offline_document_import_progress,
            office::macos_offline::commit_macos_offline_document_import,
            office::macos_offline::cancel_macos_offline_document_import,
            office::macos_offline::get_macos_offline_office_session,
            office::macos_offline::report_macos_offline_office_editor_prewarm_diagnostic,
            office::macos_offline::report_macos_offline_office_editor_prewarmed,
            office::macos_offline::update_macos_offline_office_session,
            office::macos_offline::delete_macos_offline_office_session,
            office::macos_offline::commit_macos_offline_office_session,
            office::macos_offline::cancel_macos_offline_office_session,
            office::macos_offline::get_macos_offline_office_editor_activation,
            office::macos_offline::report_macos_offline_office_editor_ready,
            office::macos_offline::present_macos_offline_office_editor_window,
            office::macos_offline::close_macos_offline_office_editor_window,
            office::macos_offline::get_macos_offline_plugin_health,
            office::macos_offline_installer::get_macos_offline_office_install_status,
            office::macos_offline_installer::install_macos_offline_office_addins,
            office::macos_offline_installer::repair_macos_offline_office_addins,
            office::macos_offline_installer::uninstall_macos_offline_office_addins,
            office::macos_offline_installer::request_quit_macos_office_hosts_for_addin_update,
            office::macos_offline_installer::reveal_macos_powerpoint_addin,
            office::macos_offline_installer::open_macos_powerpoint_addin_tutorial
        ])
        .build(tauri::generate_context!())
        .expect("error while building VisualTeX");

    app.run(move |app, event| match event {
        #[cfg(target_os = "macos")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::Resized(size),
            ..
        } => {
            if label == "main" {
                schedule_persist_main_window_size(app, size.width, size.height);
            } else {
                office::macos_offline::schedule_persist_office_editor_window_size(
                    app,
                    &label,
                    size.width,
                    size.height,
                );
            }
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::CloseRequested { api, .. },
            ..
        } if label == "main" => {
            api.prevent_close();
            let _ = office::background::hide_main_window(app);
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::Opened { urls } => {
            for url in urls {
                if url.scheme() == "visualtex" {
                    if let Err(error) = office::macos_offline::handle_open_url(app, url.as_str()) {
                        eprintln!("Unable to open VisualTeX offline Office Session: {error}");
                    }
                }
            }
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::Reopen { .. } => {
            // Office fast-open writes a validated request into the host sandbox,
            // then asks LaunchServices to deliver this Reopen event to the resident
            // app. Consume that inbox before any startup/Dock reveal logic so the
            // editor can open without a Terminal prompt, custom URL, or /tmp socket.
            match office::macos_offline::consume_fast_open_request(app) {
                Ok(true) => return,
                Ok(false) => {}
                Err(error) => {
                    eprintln!("Unable to consume VisualTeX Office fast-open request on Reopen: {error}");
                    return;
                }
            }
            // LaunchServices can emit Reopen for the initial background launch.
            // Never reveal the desktop workspace during that startup window.
            if background_mode
                && application_started_at.elapsed() < std::time::Duration::from_secs(2)
            {
                return;
            }
            // A native Office launch starts a short-lived second VisualTeX
            // process. macOS can deliver Reopen before the single-instance URL
            // arguments establish the active formula Session. Defer the Dock
            // behavior briefly so that launch cannot reveal the desktop main
            // window over a hydrating Office editor.
            let app = app.clone();
            std::thread::spawn(move || {
                std::thread::sleep(std::time::Duration::from_millis(150));
                if office::macos_offline::focus_open_office_editor(&app) {
                    return;
                }
                // Word and PowerPoint write request.json before asking
                // LaunchServices to open VisualTeX. On a slower machine the
                // Reopen event can arrive before handle_open_url has registered
                // the active editor Session. Do not mistake that Office launch
                // for an explicit Dock click and reveal the desktop workspace.
                if office::macos_offline::has_recent_office_editor_request(
                    std::time::Duration::from_secs(3),
                ) {
                    return;
                }
                let _ = office::background::reveal_main_window(&app);
            });
        }
        tauri::RunEvent::ExitRequested { .. } => {
            #[cfg(target_os = "macos")]
            if let Err(error) = office::background::pause_launch_agent_for_quit() {
                eprintln!("Unable to pause VisualTeX Office background service: {error}");
            }
        }
        _ => {}
    });
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
    fn export_file_write_is_atomic_and_overwrites_existing_output() {
        let directory = tempfile::tempdir().expect("temporary export directory");
        let target = directory.path().join("formula.svg");
        let first = "<svg xmlns=\"http://www.w3.org/2000/svg\"></svg>";
        let second = "<svg xmlns=\"http://www.w3.org/2000/svg\"><path/></svg>";

        write_export_file(
            target.to_string_lossy().into_owned(),
            BASE64_STANDARD.encode(first.as_bytes()),
        )
        .expect("first export should succeed");
        assert_eq!(fs::read_to_string(&target).unwrap(), first);

        write_export_file(
            target.to_string_lossy().into_owned(),
            BASE64_STANDARD.encode(second.as_bytes()),
        )
        .expect("second export should replace the existing file");
        assert_eq!(fs::read_to_string(&target).unwrap(), second);
        assert!(fs::read_dir(directory.path()).unwrap().all(|entry| {
            !entry
                .unwrap()
                .file_name()
                .to_string_lossy()
                .contains("visualtex-")
        }));

        let unsupported = directory.path().join("formula.exe");
        assert!(write_export_file(
            unsupported.to_string_lossy().into_owned(),
            BASE64_STANDARD.encode(b"not allowed"),
        )
        .is_err());
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
            installed_models: vec![ocr_offline::OFFLINE_DEFAULT_MODEL.to_string()],
            default_model: ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
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
