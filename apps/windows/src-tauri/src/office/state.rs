use crate::office::formula_cache::FormulaMetadataCache;
use crate::office::platform::{self, OfficePlatformBackend};
use crate::office::powerpoint_native::{
    PowerPointInteractionBus, PowerPointNativeSelection,
};
use crate::office::sessions::SessionStore;
use crate::OcrState;
use axum_server::Handle;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::net::SocketAddr;
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};
use std::sync::{Arc, Mutex, RwLock};
use tauri::AppHandle;

pub const OFFICE_BIND_IP: [u8; 4] = [127, 0, 0, 1];
pub const OFFICE_PORT: u16 = 43_127;
pub const OFFICE_PROTOCOL_VERSION: u32 = 1;
pub const OFFICE_UI_VERSION: &str = env!("CARGO_PKG_VERSION");
pub const MAX_OFFICE_REQUEST_BYTES: usize = 22 * 1024 * 1024;
pub const DEFAULT_POWERPOINT_FORMULA_FONT_SIZE_PT: f64 = 20.0;
pub const DEFAULT_MATHTYPE_DOUBLE_CLICK_EDIT_ENABLED: bool = true;

fn default_mathtype_double_click_edit_enabled() -> bool {
    DEFAULT_MATHTYPE_DOUBLE_CLICK_EDIT_ENABLED
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
struct OfficePreferencesFile {
    powerpoint_default_font_size_pt: f64,
    #[serde(default = "default_mathtype_double_click_edit_enabled")]
    mathtype_double_click_edit_enabled: bool,
}

fn normalize_formula_font_size_pt(value: f64) -> f64 {
    if !value.is_finite() {
        return DEFAULT_POWERPOINT_FORMULA_FONT_SIZE_PT;
    }
    (value.clamp(5.0, 200.0) * 2.0).round() / 2.0
}

fn office_preferences_path(paths: &OfficePaths) -> PathBuf {
    paths.root.join("office-preferences.json")
}

fn load_office_preferences(paths: &OfficePaths) -> OfficePreferencesFile {
    fs::read_to_string(office_preferences_path(paths))
        .ok()
        .and_then(|source| serde_json::from_str::<OfficePreferencesFile>(&source).ok())
        .map(|preferences| OfficePreferencesFile {
            powerpoint_default_font_size_pt: normalize_formula_font_size_pt(
                preferences.powerpoint_default_font_size_pt,
            ),
            mathtype_double_click_edit_enabled: preferences.mathtype_double_click_edit_enabled,
        })
        .unwrap_or(OfficePreferencesFile {
            powerpoint_default_font_size_pt: DEFAULT_POWERPOINT_FORMULA_FONT_SIZE_PT,
            mathtype_double_click_edit_enabled: DEFAULT_MATHTYPE_DOUBLE_CLICK_EDIT_ENABLED,
        })
}

fn persist_office_preferences(
    paths: &OfficePaths,
    preferences: &OfficePreferencesFile,
) -> Result<(), String> {
    fs::create_dir_all(&paths.root).map_err(|error| error.to_string())?;
    let target = office_preferences_path(paths);
    let temporary = target.with_extension("json.tmp");
    let payload = serde_json::to_vec_pretty(preferences)
        .map_err(|error| error.to_string())?;
    fs::write(&temporary, payload).map_err(|error| error.to_string())?;
    if target.exists() {
        fs::remove_file(&target).map_err(|error| error.to_string())?;
    }
    fs::rename(&temporary, &target).map_err(|error| error.to_string())
}

pub fn normalize_app_theme(theme: &str) -> &'static str {
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

pub fn normalize_app_editor_layout(layout: &str) -> &'static str {
    match layout.trim() {
        "standard" => "standard",
        _ => "classic",
    }
}

pub fn append_office_log(paths: &OfficePaths, file_name: &str, message: &str) {
    #[cfg(target_os = "windows")]
    let log_root = std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("office")
        .join("logs");
    #[cfg(not(target_os = "windows"))]
    let log_root = paths.root.join("logs");
    #[cfg(target_os = "windows")]
    let _ = paths;
    if fs::create_dir_all(&log_root).is_err() {
        return;
    }
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|value| value.as_secs())
        .unwrap_or_default();
    let path = log_root.join(file_name);
    if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(file, "[{timestamp}] {message}");
        let _ = file.flush();
    }
}

#[derive(Debug, Clone)]
pub struct OfficePaths {
    pub root: PathBuf,
    pub certificate: PathBuf,
    pub private_key: PathBuf,
    pub certificate_metadata: PathBuf,
    pub install: PathBuf,
    pub sessions: PathBuf,
    pub recovery: PathBuf,
    pub formula_cache: PathBuf,
    pub ui_root: PathBuf,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeCompanionStatus {
    pub running: bool,
    pub bind_address: String,
    pub port: u16,
    pub certificate_path: String,
    pub office_ui_version: String,
    pub protocol_version: u32,
    pub last_error: Option<String>,
}

impl OfficeCompanionStatus {
    pub fn stopped(paths: &OfficePaths) -> Self {
        Self {
            running: false,
            bind_address: "127.0.0.1".to_string(),
            port: OFFICE_PORT,
            certificate_path: paths.certificate.display().to_string(),
            office_ui_version: OFFICE_UI_VERSION.to_string(),
            protocol_version: OFFICE_PROTOCOL_VERSION,
            last_error: None,
        }
    }
}

#[derive(Clone)]
pub struct OfficeCompanionState {
    pub app: Option<AppHandle>,
    pub ocr: OcrState,
    pub paths: Arc<OfficePaths>,
    pub install_token: Arc<String>,
    pub status: Arc<RwLock<OfficeCompanionStatus>>,
    pub app_theme: Arc<RwLock<String>>,
    pub app_editor_layout: Arc<RwLock<String>>,
    pub app_editor_preferences: Arc<RwLock<serde_json::Value>>,
    pub powerpoint_default_font_size_pt: Arc<RwLock<f64>>,
    pub mathtype_double_click_edit_enabled: Arc<RwLock<bool>>,
    pub server_handle: Arc<Mutex<Option<Handle<SocketAddr>>>>,
    pub session_store: SessionStore,
    pub formula_cache: FormulaMetadataCache,
    pub platform_backend: Arc<dyn OfficePlatformBackend>,
    pub powerpoint_interactions: PowerPointInteractionBus,
    /// Native PowerPoint insertion is prepared before the Office.js command
    /// page writes durable tags and accessibility metadata. Keep the immutable
    /// prepared selection by Session id so retries never paste a second image.
    pub prepared_powerpoint_commits:
        Arc<Mutex<HashMap<String, PowerPointNativeSelection>>>,
    pub ocr_available: bool,
}

impl OfficeCompanionState {
    pub fn new(
        app: Option<AppHandle>,
        ocr: OcrState,
        paths: OfficePaths,
        install_token: String,
        session_store: SessionStore,
        formula_cache: FormulaMetadataCache,
        ocr_available: bool,
    ) -> Self {
        let status = OfficeCompanionStatus::stopped(&paths);
        let office_preferences = load_office_preferences(&paths);
        let powerpoint_default_font_size_pt = office_preferences.powerpoint_default_font_size_pt;
        let mathtype_double_click_edit_enabled =
            office_preferences.mathtype_double_click_edit_enabled;
        let platform_backend = platform::create_backend(app.as_ref(), &paths);
        Self {
            app,
            ocr,
            paths: Arc::new(paths),
            install_token: Arc::new(install_token),
            status: Arc::new(RwLock::new(status)),
            app_theme: Arc::new(RwLock::new("light".to_string())),
            app_editor_layout: Arc::new(RwLock::new("classic".to_string())),
            app_editor_preferences: Arc::new(RwLock::new(serde_json::json!({}))),
            powerpoint_default_font_size_pt: Arc::new(RwLock::new(
                powerpoint_default_font_size_pt,
            )),
            mathtype_double_click_edit_enabled: Arc::new(RwLock::new(
                mathtype_double_click_edit_enabled,
            )),
            server_handle: Arc::new(Mutex::new(None)),
            session_store,
            formula_cache,
            platform_backend,
            powerpoint_interactions: PowerPointInteractionBus::default(),
            prepared_powerpoint_commits: Arc::new(Mutex::new(HashMap::new())),
            ocr_available,
        }
    }

    pub fn socket_addr() -> SocketAddr {
        SocketAddr::from((OFFICE_BIND_IP, OFFICE_PORT))
    }

    pub fn snapshot(&self) -> OfficeCompanionStatus {
        self.status
            .read()
            .map(|value| value.clone())
            .unwrap_or_else(|_| OfficeCompanionStatus::stopped(&self.paths))
    }

    pub fn update_status(&self, mutate: impl FnOnce(&mut OfficeCompanionStatus)) {
        if let Ok(mut status) = self.status.write() {
            mutate(&mut status);
        }
    }

    pub fn current_theme(&self) -> String {
        self.app_theme
            .read()
            .map(|theme| theme.clone())
            .unwrap_or_else(|_| "light".to_string())
    }

    pub fn set_current_theme(&self, theme: &str) -> String {
        let normalized = normalize_app_theme(theme).to_string();
        if let Ok(mut current) = self.app_theme.write() {
            *current = normalized.clone();
        }
        normalized
    }

    pub fn current_editor_layout(&self) -> String {
        self.app_editor_layout
            .read()
            .map(|layout| layout.clone())
            .unwrap_or_else(|_| "classic".to_string())
    }

    pub fn set_current_editor_layout(&self, layout: &str) -> String {
        let normalized = normalize_app_editor_layout(layout).to_string();
        if let Ok(mut current) = self.app_editor_layout.write() {
            *current = normalized.clone();
        }
        normalized
    }

    pub fn current_editor_preferences(&self) -> serde_json::Value {
        self.app_editor_preferences
            .read()
            .map(|preferences| preferences.clone())
            .unwrap_or_else(|_| serde_json::json!({}))
    }

    pub fn set_current_editor_preferences(
        &self,
        preferences: serde_json::Value,
    ) -> serde_json::Value {
        let normalized = if preferences.is_object() {
            preferences
        } else {
            serde_json::json!({})
        };
        if let Ok(mut current) = self.app_editor_preferences.write() {
            *current = normalized.clone();
        }
        normalized
    }

    pub fn powerpoint_default_font_size_pt(&self) -> f64 {
        self.powerpoint_default_font_size_pt
            .read()
            .map(|value| *value)
            .unwrap_or(DEFAULT_POWERPOINT_FORMULA_FONT_SIZE_PT)
    }

    pub fn set_powerpoint_default_font_size_pt(
        &self,
        font_size_pt: f64,
    ) -> Result<f64, String> {
        let normalized = normalize_formula_font_size_pt(font_size_pt);
        let preferences = OfficePreferencesFile {
            powerpoint_default_font_size_pt: normalized,
            mathtype_double_click_edit_enabled: self.mathtype_double_click_edit_enabled(),
        };
        persist_office_preferences(&self.paths, &preferences)?;
        if let Ok(mut current) = self.powerpoint_default_font_size_pt.write() {
            *current = normalized;
        }
        Ok(normalized)
    }

    pub fn mathtype_double_click_edit_enabled(&self) -> bool {
        self.mathtype_double_click_edit_enabled
            .read()
            .map(|value| *value)
            .unwrap_or(DEFAULT_MATHTYPE_DOUBLE_CLICK_EDIT_ENABLED)
    }

    pub fn set_mathtype_double_click_edit_enabled(
        &self,
        enabled: bool,
    ) -> Result<bool, String> {
        let preferences = OfficePreferencesFile {
            powerpoint_default_font_size_pt: self.powerpoint_default_font_size_pt(),
            mathtype_double_click_edit_enabled: enabled,
        };
        persist_office_preferences(&self.paths, &preferences)?;
        if let Ok(mut current) = self.mathtype_double_click_edit_enabled.write() {
            *current = enabled;
        }
        Ok(enabled)
    }
}
