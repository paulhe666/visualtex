use std::collections::BTreeSet;
use std::ffi::OsStr;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};
use serde::{Deserialize, Serialize};
use tauri::webview::PageLoadEvent;
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder};

pub const OFFICE_BACKGROUND_ARGUMENT: &str = "--office-background";
pub const OFFICE_BOOTSTRAP_ARGUMENT: &str = "--office-bootstrap";

const MAIN_WINDOW_SIZE_FILE: &str = "main-window-size.json";
const KEYPAD_WINDOW_SIZE_FILE: &str = "main-window-keypad-size.json";
const DEFAULT_MAIN_WINDOW_WIDTH: f64 = 1182.2857142857142;
const DEFAULT_MAIN_WINDOW_HEIGHT: f64 = 728.0;
const DEFAULT_KEYPAD_WINDOW_WIDTH: f64 = 760.0;
const DEFAULT_KEYPAD_WINDOW_HEIGHT: f64 = 260.0;
const LEGACY_DEFAULT_KEYPAD_WINDOW_HEIGHT: f64 = 520.0;
const MIN_MAIN_WINDOW_WIDTH: f64 = 640.0;
const MIN_MAIN_WINDOW_HEIGHT: f64 = 240.0;
const MAX_MAIN_WINDOW_WIDTH: f64 = 4000.0;
const MAX_MAIN_WINDOW_HEIGHT: f64 = 3000.0;
static MAIN_WINDOW_KEYPAD_MODE: AtomicBool = AtomicBool::new(false);
static MAIN_WINDOW_SIZE_WRITE_GENERATION: AtomicU64 = AtomicU64::new(0);

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct MainWindowSizePreference {
    width: f64,
    height: f64,
}

fn default_main_window_size(keypad_mode: bool) -> MainWindowSizePreference {
    if keypad_mode {
        MainWindowSizePreference {
            width: DEFAULT_KEYPAD_WINDOW_WIDTH,
            height: DEFAULT_KEYPAD_WINDOW_HEIGHT,
        }
    } else {
        MainWindowSizePreference {
            width: DEFAULT_MAIN_WINDOW_WIDTH,
            height: DEFAULT_MAIN_WINDOW_HEIGHT,
        }
    }
}

fn normalize_main_window_size_for_mode(
    width: f64,
    height: f64,
    keypad_mode: bool,
) -> MainWindowSizePreference {
    let defaults = default_main_window_size(keypad_mode);
    let width = if width.is_finite() { width } else { defaults.width };
    let height = if height.is_finite() { height } else { defaults.height };
    MainWindowSizePreference {
        width: width.clamp(MIN_MAIN_WINDOW_WIDTH, MAX_MAIN_WINDOW_WIDTH),
        height: height.clamp(MIN_MAIN_WINDOW_HEIGHT, MAX_MAIN_WINDOW_HEIGHT),
    }
}

fn normalize_main_window_size(width: f64, height: f64) -> MainWindowSizePreference {
    normalize_main_window_size_for_mode(width, height, false)
}

fn migrate_legacy_keypad_window_size(
    size: MainWindowSizePreference,
    keypad_mode: bool,
) -> MainWindowSizePreference {
    if keypad_mode
        && (size.width - DEFAULT_KEYPAD_WINDOW_WIDTH).abs() < 0.5
        && (size.height - LEGACY_DEFAULT_KEYPAD_WINDOW_HEIGHT).abs() < 0.5
    {
        return default_main_window_size(true);
    }
    size
}

fn main_window_size_path_for_mode(app: &AppHandle, keypad_mode: bool) -> Result<PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX app-data directory: {error}"))?;
    Ok(app_data.join(if keypad_mode {
        KEYPAD_WINDOW_SIZE_FILE
    } else {
        MAIN_WINDOW_SIZE_FILE
    }))
}

fn read_main_window_size_for_mode(
    app: &AppHandle,
    keypad_mode: bool,
) -> Option<MainWindowSizePreference> {
    let path = main_window_size_path_for_mode(app, keypad_mode).ok()?;
    let bytes = fs::read(&path).ok()?;
    let size = serde_json::from_slice::<MainWindowSizePreference>(&bytes).ok()?;
    let normalized = normalize_main_window_size_for_mode(size.width, size.height, keypad_mode);
    let migrated = migrate_legacy_keypad_window_size(normalized, keypad_mode);
    if keypad_mode
        && (migrated.width - normalized.width).abs() >= 0.5
        || keypad_mode && (migrated.height - normalized.height).abs() >= 0.5
    {
        if let Ok(bytes) = serde_json::to_vec_pretty(&migrated) {
            let _ = fs::write(&path, bytes);
        }
        append_lifecycle_log(format!(
            "migrated legacy keypad window size from {}x{} to {}x{}",
            normalized.width, normalized.height, migrated.width, migrated.height
        ));
    }
    Some(migrated)
}

fn read_main_window_size(app: &AppHandle) -> Option<MainWindowSizePreference> {
    read_main_window_size_for_mode(app, false)
}

fn write_main_window_size_for_mode(
    app: &AppHandle,
    size: MainWindowSizePreference,
    keypad_mode: bool,
) -> Result<(), String> {
    let path = main_window_size_path_for_mode(app, keypad_mode)?;
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    let bytes = serde_json::to_vec_pretty(&size).map_err(|error| error.to_string())?;
    fs::write(path, bytes).map_err(|error| error.to_string())
}

fn write_main_window_size(app: &AppHandle, size: MainWindowSizePreference) -> Result<(), String> {
    write_main_window_size_for_mode(app, size, false)
}

fn current_main_window_size(app: &AppHandle, keypad_mode: bool) -> Option<MainWindowSizePreference> {
    let window = app.get_webview_window("main")?;
    let physical = window.inner_size().ok()?;
    let scale_factor = window.scale_factor().ok()?.max(0.1);
    Some(normalize_main_window_size_for_mode(
        f64::from(physical.width) / scale_factor,
        f64::from(physical.height) / scale_factor,
        keypad_mode,
    ))
}

pub(crate) fn configuration_main_window_size(app: &AppHandle) -> Option<(f64, f64)> {
    if !MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst) {
        if let Some(size) = current_main_window_size(app, false) {
            return Some((size.width, size.height));
        }
    }
    read_main_window_size(app).map(|size| (size.width, size.height))
}

pub(crate) fn configuration_keypad_window_size(app: &AppHandle) -> Option<(f64, f64)> {
    if MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst) {
        if let Some(size) = current_main_window_size(app, true) {
            return Some((size.width, size.height));
        }
    }
    read_main_window_size_for_mode(app, true).map(|size| (size.width, size.height))
}

pub(crate) fn apply_configuration_main_window_size(
    app: &AppHandle,
    width: f64,
    height: f64,
) -> Result<(), String> {
    let size = normalize_main_window_size(width, height);
    write_main_window_size(app, size)?;
    if !MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst) {
        if let Some(window) = app.get_webview_window("main") {
            window
                .set_size(tauri::LogicalSize::new(size.width, size.height))
                .map_err(|error| error.to_string())?;
            window.center().map_err(|error| error.to_string())?;
        }
    }
    Ok(())
}

pub(crate) fn apply_configuration_keypad_window_size(
    app: &AppHandle,
    width: f64,
    height: f64,
) -> Result<(), String> {
    let size = normalize_main_window_size_for_mode(width, height, true);
    write_main_window_size_for_mode(app, size, true)?;
    if MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst) {
        if let Some(window) = app.get_webview_window("main") {
            window
                .set_size(tauri::LogicalSize::new(size.width, size.height))
                .map_err(|error| error.to_string())?;
            window.center().map_err(|error| error.to_string())?;
        }
    }
    Ok(())
}

pub(crate) fn set_main_window_keypad_mode(app: &AppHandle, enabled: bool) -> Result<(), String> {
    let previous_mode = MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst);
    MAIN_WINDOW_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst);

    if let Some(size) = current_main_window_size(app, previous_mode) {
        write_main_window_size_for_mode(app, size, previous_mode)?;
    }

    MAIN_WINDOW_KEYPAD_MODE.store(enabled, Ordering::SeqCst);
    let target_size = read_main_window_size_for_mode(app, enabled)
        .unwrap_or_else(|| default_main_window_size(enabled));

    if let Some(window) = app.get_webview_window("main") {
        if window.is_maximized().unwrap_or(false) {
            window.unmaximize().map_err(|error| error.to_string())?;
        }
        window
            .set_min_size(Some(tauri::LogicalSize::new(
                MIN_MAIN_WINDOW_WIDTH,
                MIN_MAIN_WINDOW_HEIGHT,
            )))
            .map_err(|error| error.to_string())?;
        window
            .set_size(tauri::LogicalSize::new(target_size.width, target_size.height))
            .map_err(|error| error.to_string())?;
    }

    append_lifecycle_log(format!(
        "main window switched to {} mode at {}x{}",
        if enabled { "keypad" } else { "normal" },
        target_size.width,
        target_size.height
    ));
    Ok(())
}

pub(crate) fn schedule_persist_main_window_size(
    app: &AppHandle,
    physical_width: u32,
    physical_height: u32,
) {
    if physical_width == 0 || physical_height == 0 {
        return;
    }
    let Some(window) = app.get_webview_window("main") else {
        return;
    };
    let keypad_mode = MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst);
    let scale_factor = window.scale_factor().unwrap_or(1.0).max(0.1);
    let size = normalize_main_window_size_for_mode(
        f64::from(physical_width) / scale_factor,
        f64::from(physical_height) / scale_factor,
        keypad_mode,
    );
    let generation = MAIN_WINDOW_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst) + 1;
    let app = app.clone();
    std::thread::spawn(move || {
        std::thread::sleep(std::time::Duration::from_millis(250));
        if MAIN_WINDOW_SIZE_WRITE_GENERATION.load(Ordering::SeqCst) != generation {
            return;
        }
        if let Err(error) = write_main_window_size_for_mode(&app, size, keypad_mode) {
            append_lifecycle_log(format!(
                "Unable to persist {} main window size: {error}",
                if keypad_mode { "keypad" } else { "normal" }
            ));
        }
    });
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AppRunMode {
    Desktop,
    OfficeBackground,
    OfficeBootstrap,
}

impl AppRunMode {
    pub fn current() -> Self {
        Self::from_arguments(std::env::args_os())
    }

    pub fn from_arguments(arguments: impl IntoIterator<Item = impl AsRef<OsStr>>) -> Self {
        let mut background = false;
        for argument in arguments {
            let argument = argument.as_ref();
            if argument == OsStr::new(OFFICE_BOOTSTRAP_ARGUMENT) {
                return Self::OfficeBootstrap;
            }
            if argument == OsStr::new(OFFICE_BACKGROUND_ARGUMENT) {
                background = true;
            }
        }
        if background {
            Self::OfficeBackground
        } else {
            Self::Desktop
        }
    }

    pub fn label(self) -> &'static str {
        match self {
            Self::Desktop => "desktop",
            Self::OfficeBackground => "office-background",
            Self::OfficeBootstrap => "office-bootstrap",
        }
    }

    pub fn creates_main_window(self) -> bool {
        matches!(self, Self::Desktop)
    }

    pub fn starts_companion(self) -> bool {
        !matches!(self, Self::OfficeBootstrap)
    }

    pub fn schedules_ocr_warmup(self) -> bool {
        !matches!(self, Self::OfficeBootstrap)
    }
}

pub fn arguments_request_desktop(arguments: &[String]) -> bool {
    !arguments.iter().any(|argument| {
        argument == OFFICE_BACKGROUND_ARGUMENT || argument == OFFICE_BOOTSTRAP_ARGUMENT
    })
}

fn lifecycle_log_path() -> PathBuf {
    if let Ok(Some(runtime)) = crate::office::state::acceptance_runtime() {
        return runtime.app_data_root.join("logs").join("app-lifecycle.log");
    }
    #[cfg(target_os = "windows")]
    let root = std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("logs");
    #[cfg(not(target_os = "windows"))]
    let root = std::env::temp_dir().join("VisualTeX").join("logs");
    root.join("app-lifecycle.log")
}

pub fn append_lifecycle_log(message: impl AsRef<str>) {
    let path = lifecycle_log_path();
    let Some(parent) = path.parent() else {
        return;
    };
    if fs::create_dir_all(parent).is_err() {
        return;
    }
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|value| value.as_secs())
        .unwrap_or_default();
    if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(
            file,
            "[{timestamp}] pid={} {}",
            std::process::id(),
            message.as_ref()
        );
        let _ = file.flush();
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MainAssetDiagnostic {
    pub index_present: bool,
    pub referenced_assets: Vec<String>,
    pub missing_assets: Vec<String>,
    pub detail: String,
}

fn extract_attribute_assets(source: &str, marker: &str, quote: char, output: &mut BTreeSet<String>) {
    let mut remainder = source;
    while let Some(index) = remainder.find(marker) {
        let value_start = index + marker.len();
        let tail = &remainder[value_start..];
        let Some(value_end) = tail.find(quote) else {
            break;
        };
        let raw = &tail[..value_end];
        let path = raw
            .split(['?', '#'])
            .next()
            .unwrap_or(raw)
            .trim_start_matches('/')
            .trim();
        if path.starts_with("assets/") && (path.ends_with(".js") || path.ends_with(".css")) {
            output.insert(path.to_string());
        }
        remainder = &tail[value_end + quote.len_utf8()..];
    }
}

pub fn referenced_frontend_assets(index_html: &str) -> Vec<String> {
    let mut assets = BTreeSet::new();
    extract_attribute_assets(index_html, "src=\"", '"', &mut assets);
    extract_attribute_assets(index_html, "href=\"", '"', &mut assets);
    extract_attribute_assets(index_html, "src='", '\'', &mut assets);
    extract_attribute_assets(index_html, "href='", '\'', &mut assets);
    assets.into_iter().collect()
}

pub fn diagnose_main_assets(
    index_html: Option<&[u8]>,
    mut asset_exists: impl FnMut(&str) -> bool,
) -> MainAssetDiagnostic {
    let Some(index_html) = index_html else {
        return MainAssetDiagnostic {
            index_present: false,
            referenced_assets: Vec::new(),
            missing_assets: vec!["index.html".to_string()],
            detail: "Tauri embedded main asset index.html is missing; Office companion resources are a separate installed resource tree and were not used for this check".to_string(),
        };
    };
    let Ok(index_html) = std::str::from_utf8(index_html) else {
        return MainAssetDiagnostic {
            index_present: true,
            referenced_assets: Vec::new(),
            missing_assets: vec!["index.html (invalid UTF-8)".to_string()],
            detail: "Tauri embedded index.html exists but is not valid UTF-8".to_string(),
        };
    };
    let referenced_assets = referenced_frontend_assets(index_html);
    let missing_assets = referenced_assets
        .iter()
        .filter(|path| !asset_exists(path))
        .cloned()
        .collect::<Vec<_>>();
    let detail = if missing_assets.is_empty() {
        format!(
            "Tauri embedded main assets verified: index.html plus {} referenced JS/CSS assets",
            referenced_assets.len()
        )
    } else {
        format!(
            "Tauri embedded index.html exists but referenced main assets are missing: {}",
            missing_assets.join(", ")
        )
    };
    MainAssetDiagnostic {
        index_present: true,
        referenced_assets,
        missing_assets,
        detail,
    }
}

pub fn validate_embedded_main_assets(app: &AppHandle) -> Result<MainAssetDiagnostic, String> {
    let resolver = app.asset_resolver();
    let index = resolver.get("index.html".to_string());
    let diagnostic = diagnose_main_assets(index.as_ref().map(|asset| asset.bytes.as_slice()), |path| {
        resolver.get(path.to_string()).is_some()
    });
    append_lifecycle_log(format!(
        "main-asset diagnostic index_present={} referenced={} missing={} detail={}",
        diagnostic.index_present,
        diagnostic.referenced_assets.len(),
        diagnostic.missing_assets.join("|"),
        diagnostic.detail
    ));
    if diagnostic.index_present && diagnostic.missing_assets.is_empty() {
        Ok(diagnostic)
    } else {
        Err(diagnostic.detail.clone())
    }
}

fn focus_existing_main_window(window: &WebviewWindow) -> Result<(), String> {
    if let Err(error) = window.show() {
        let message = format!("Unable to show existing main window: {error}");
        append_lifecycle_log(&message);
        return Err(message);
    }
    if let Err(error) = window.unminimize() {
        let message = format!("Unable to unminimize existing main window: {error}");
        append_lifecycle_log(&message);
        return Err(message);
    }
    if let Err(error) = window.set_focus() {
        let message = format!("Unable to focus existing main window: {error}");
        append_lifecycle_log(&message);
        return Err(message);
    }
    append_lifecycle_log("existing main window shown, unminimized and focused");
    Ok(())
}

pub fn ensure_main_window(app: &AppHandle) -> Result<WebviewWindow, String> {
    if let Some(window) = app.get_webview_window("main") {
        focus_existing_main_window(&window)?;
        return Ok(window);
    }

    // Every newly created desktop window starts in normal mode. Keypad mode is
    // an in-window compact workspace and restores its own geometry only after
    // the user explicitly enters it.
    MAIN_WINDOW_KEYPAD_MODE.store(false, Ordering::SeqCst);
    MAIN_WINDOW_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst);

    validate_embedded_main_assets(app).map_err(|error| {
        append_lifecycle_log(format!("main window creation blocked by embedded asset failure: {error}"));
        error
    })?;

    let saved_size = read_main_window_size(app)
        .unwrap_or_else(|| normalize_main_window_size(DEFAULT_MAIN_WINDOW_WIDTH, DEFAULT_MAIN_WINDOW_HEIGHT));
    append_lifecycle_log(format!(
        "creating main window from WebviewUrl::App(index.html) at {}x{}",
        saved_size.width, saved_size.height
    ));
    let window = WebviewWindowBuilder::new(
        app,
        "main",
        WebviewUrl::App("index.html".into()),
    )
    .title("VisualTeX")
    .inner_size(saved_size.width, saved_size.height)
    .min_inner_size(MIN_MAIN_WINDOW_WIDTH, MIN_MAIN_WINDOW_HEIGHT)
    .resizable(true)
    .center()
    .focused(true)
    .visible(true)
    .on_page_load(|_window, payload| {
        append_lifecycle_log(format!(
            "main window page-load event={:?} url={}",
            payload.event(),
            payload.url()
        ));
        if matches!(payload.event(), PageLoadEvent::Finished) {
            append_lifecycle_log("main window finished loading embedded application entry");
        }
    })
    .build()
    .map_err(|error| {
        let message = format!("Unable to create VisualTeX main window: {error}");
        append_lifecycle_log(&message);
        message
    })?;

    focus_existing_main_window(&window)?;
    append_lifecycle_log("new main window created successfully");
    Ok(window)
}

pub fn background_retention_enabled() -> bool {
    #[cfg(target_os = "windows")]
    {
        crate::office::windows_backend::background_start_enabled()
    }
    #[cfg(not(target_os = "windows"))]
    {
        crate::office::background::status().installed
    }
}

pub fn destroy_main_window_for_background(app: &AppHandle) -> Result<(), String> {
    let Some(window) = app.get_webview_window("main") else {
        append_lifecycle_log("main window close requested for background retention, but main no longer exists");
        return Ok(());
    };
    let keypad_mode = MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst);
    MAIN_WINDOW_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst);
    if let Some(size) = current_main_window_size(app, keypad_mode) {
        if let Err(error) = write_main_window_size_for_mode(app, size, keypad_mode) {
            append_lifecycle_log(format!(
                "Unable to persist {} main window size before destroy: {error}",
                if keypad_mode { "keypad" } else { "normal" }
            ));
        }
    }
    window.destroy().map_err(|error| {
        let message = format!("Unable to destroy main window while retaining background companion: {error}");
        append_lifecycle_log(&message);
        message
    })?;
    append_lifecycle_log("main window destroyed; background companion remains available for a future desktop launch");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normal_and_keypad_window_defaults_are_independent() {
        let normal = normalize_main_window_size_for_mode(f64::NAN, f64::NAN, false);
        let keypad = normalize_main_window_size_for_mode(f64::NAN, f64::NAN, true);
        assert_eq!(normal.width, DEFAULT_MAIN_WINDOW_WIDTH);
        assert_eq!(normal.height, DEFAULT_MAIN_WINDOW_HEIGHT);
        assert_eq!(keypad.width, DEFAULT_KEYPAD_WINDOW_WIDTH);
        assert_eq!(keypad.height, DEFAULT_KEYPAD_WINDOW_HEIGHT);
        assert!(keypad.width < normal.width);
        assert!(keypad.height < normal.height);
    }

    #[test]
    fn keypad_window_can_reach_main_window_minimum() {
        let keypad = normalize_main_window_size_for_mode(1.0, 1.0, true);
        assert_eq!(keypad.width, MIN_MAIN_WINDOW_WIDTH);
        assert_eq!(keypad.height, MIN_MAIN_WINDOW_HEIGHT);
    }

    #[test]
    fn legacy_keypad_default_migrates_to_compact_height() {
        let migrated = migrate_legacy_keypad_window_size(
            MainWindowSizePreference {
                width: DEFAULT_KEYPAD_WINDOW_WIDTH,
                height: LEGACY_DEFAULT_KEYPAD_WINDOW_HEIGHT,
            },
            true,
        );
        assert_eq!(migrated.width, DEFAULT_KEYPAD_WINDOW_WIDTH);
        assert_eq!(migrated.height, DEFAULT_KEYPAD_WINDOW_HEIGHT);

        let custom = migrate_legacy_keypad_window_size(
            MainWindowSizePreference {
                width: 900.0,
                height: 520.0,
            },
            true,
        );
        assert_eq!(custom.width, 900.0);
        assert_eq!(custom.height, 520.0);
    }

    #[test]
    fn run_modes_are_distinct_and_bootstrap_has_priority() {
        assert_eq!(
            AppRunMode::from_arguments(["visualtex.exe"]),
            AppRunMode::Desktop
        );
        assert_eq!(
            AppRunMode::from_arguments(["visualtex.exe", OFFICE_BACKGROUND_ARGUMENT]),
            AppRunMode::OfficeBackground
        );
        assert_eq!(
            AppRunMode::from_arguments([
                "visualtex.exe",
                OFFICE_BACKGROUND_ARGUMENT,
                OFFICE_BOOTSTRAP_ARGUMENT,
            ]),
            AppRunMode::OfficeBootstrap
        );
        assert!(AppRunMode::Desktop.creates_main_window());
        assert!(!AppRunMode::OfficeBackground.creates_main_window());
        assert!(!AppRunMode::OfficeBootstrap.starts_companion());
        assert!(!AppRunMode::OfficeBootstrap.schedules_ocr_warmup());
    }

    #[test]
    fn background_single_instance_notification_can_request_desktop() {
        assert!(arguments_request_desktop(&[
            "C:\\VisualTeX\\visualtex.exe".to_string()
        ]));
        assert!(!arguments_request_desktop(&[
            "visualtex.exe".to_string(),
            OFFICE_BACKGROUND_ARGUMENT.to_string(),
        ]));
        assert!(!arguments_request_desktop(&[
            "visualtex.exe".to_string(),
            OFFICE_BOOTSTRAP_ARGUMENT.to_string(),
        ]));
    }

    #[test]
    fn main_asset_diagnostic_isolated_from_office_resource_success() {
        let diagnostic = diagnose_main_assets(None, |_| true);
        assert!(!diagnostic.index_present);
        assert_eq!(diagnostic.missing_assets, ["index.html"]);
        assert!(diagnostic.detail.contains("Tauri embedded main asset"));
        assert!(diagnostic.detail.contains("Office companion resources are a separate"));
    }

    #[test]
    fn main_asset_diagnostic_reports_missing_referenced_bundle() {
        let html = br#"<link rel="stylesheet" href="/assets/index-a.css"><script type="module" src="/assets/index-b.js"></script>"#;
        let diagnostic = diagnose_main_assets(Some(html), |path| path.ends_with(".css"));
        assert!(diagnostic.index_present);
        assert_eq!(
            diagnostic.referenced_assets,
            ["assets/index-a.css", "assets/index-b.js"]
        );
        assert_eq!(diagnostic.missing_assets, ["assets/index-b.js"]);
    }

    #[test]
    fn main_asset_diagnostic_accepts_complete_embedded_bundle() {
        let html = br#"<link rel='stylesheet' href='/assets/index-a.css?x=1'><script src='/assets/index-b.js#v'></script>"#;
        let diagnostic = diagnose_main_assets(Some(html), |_| true);
        assert!(diagnostic.index_present);
        assert!(diagnostic.missing_assets.is_empty());
        assert_eq!(diagnostic.referenced_assets.len(), 2);
    }
}
