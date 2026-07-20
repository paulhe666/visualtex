use crate::office::state::OfficePaths;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::sync::Arc;
use tauri::AppHandle;

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

pub fn create_backend(
    app: Option<&AppHandle>,
    paths: &OfficePaths,
) -> Arc<dyn OfficePlatformBackend> {
    Arc::new(windows::WindowsOfficePlatformBackend::new(app, paths.clone()))
}
