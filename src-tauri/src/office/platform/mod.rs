use crate::office::state::OfficePaths;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::sync::Arc;
use tauri::AppHandle;

#[cfg(target_os = "macos")]
pub mod macos;
#[cfg(target_os = "windows")]
pub mod windows;

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum OfficeIntegrationMode {
    Auto,
    Ole,
    Vsto,
}

impl Default for OfficeIntegrationMode {
    fn default() -> Self {
        Self::Auto
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficePlatformStatus {
    pub platform: String,
    pub mode: OfficeIntegrationMode,
    pub active_backend: String,
    pub ole_bridge_healthy: bool,
    pub vsto_word_healthy: bool,
    pub vsto_powerpoint_healthy: bool,
    pub office_catalog_registered: bool,
    pub current_user_certificate_trusted: bool,
    pub background_start_enabled: bool,
    pub last_error: Option<String>,
}

pub trait OfficePlatformBackend: Send + Sync {
    fn status(&self) -> OfficePlatformStatus;
    fn set_mode(&self, mode: OfficeIntegrationMode) -> Result<OfficePlatformStatus, String>;
    fn request(&self, request: Value) -> Result<Value, String>;
    fn events_after(&self, cursor: u64) -> Vec<Value>;
    fn shutdown(&self) -> Result<(), String>;
}

#[cfg(target_os = "macos")]
pub fn create_backend(
    _app: Option<&AppHandle>,
    paths: &OfficePaths,
) -> Arc<dyn OfficePlatformBackend> {
    Arc::new(macos::MacOfficePlatformBackend::new(paths.clone()))
}

#[cfg(target_os = "windows")]
pub fn create_backend(
    app: Option<&AppHandle>,
    paths: &OfficePaths,
) -> Arc<dyn OfficePlatformBackend> {
    Arc::new(windows::WindowsOfficePlatformBackend::new(app, paths.clone()))
}

#[cfg(not(any(target_os = "macos", target_os = "windows")))]
pub fn create_backend(
    _app: Option<&AppHandle>,
    _paths: &OfficePaths,
) -> Arc<dyn OfficePlatformBackend> {
    Arc::new(UnsupportedOfficePlatformBackend)
}

#[cfg(not(any(target_os = "macos", target_os = "windows")))]
struct UnsupportedOfficePlatformBackend;

#[cfg(not(any(target_os = "macos", target_os = "windows")))]
impl OfficePlatformBackend for UnsupportedOfficePlatformBackend {
    fn status(&self) -> OfficePlatformStatus {
        OfficePlatformStatus {
            platform: std::env::consts::OS.to_string(),
            mode: OfficeIntegrationMode::Auto,
            active_backend: "none".to_string(),
            ole_bridge_healthy: false,
            vsto_word_healthy: false,
            vsto_powerpoint_healthy: false,
            office_catalog_registered: false,
            current_user_certificate_trusted: false,
            background_start_enabled: false,
            last_error: Some("Office integration is not supported on this platform".to_string()),
        }
    }

    fn set_mode(&self, _mode: OfficeIntegrationMode) -> Result<OfficePlatformStatus, String> {
        Err("Office integration is not supported on this platform".to_string())
    }

    fn request(&self, _request: Value) -> Result<Value, String> {
        Err("Office integration is not supported on this platform".to_string())
    }

    fn events_after(&self, _cursor: u64) -> Vec<Value> {
        Vec::new()
    }

    fn shutdown(&self) -> Result<(), String> {
        Ok(())
    }
}
