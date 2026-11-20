use crate::office::state::OfficePaths;
use serde::{Deserialize, Serialize};
use std::sync::Arc;

pub mod windows;

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum OfficeIntegrationMode {
    Auto,
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
    pub ole_local_server_healthy: bool,
    pub static_install_verified: bool,
    pub word_files_present: bool,
    pub word_registry_complete: bool,
    pub word_load_enabled: bool,
    pub powerpoint_files_present: bool,
    pub powerpoint_registry_complete: bool,
    pub powerpoint_load_enabled: bool,
    pub vsto_word_healthy: bool,
    pub vsto_powerpoint_healthy: bool,
    pub word_connected: bool,
    pub powerpoint_connected: bool,
    pub connection_verification_attempted: bool,
    pub companion_process_running: bool,
    pub companion_port_listening: bool,
    pub companion_https_healthy: bool,
    pub companion_certificate_matches: bool,
    pub companion_protocol_matches: bool,
    pub office_runtime_verified: bool,
    pub current_user_certificate_trusted: bool,
    pub background_start_enabled: bool,
    /// Word in this user session draws MathType's pictures with MathType's own
    /// fonts (MT Extra, the Euclid fonts); false when they are missing or an
    /// older copy such as Office's MT Extra takes their place.
    pub mathtype_fonts_healthy: bool,
    pub mathtype_fonts_problem: Option<String>,
    pub last_error: Option<String>,
}

pub trait OfficePlatformBackend: Send + Sync {
    fn status(&self) -> OfficePlatformStatus;
    fn set_mode(&self, mode: OfficeIntegrationMode) -> Result<OfficePlatformStatus, String>;
}

pub fn create_backend(paths: &OfficePaths) -> Arc<dyn OfficePlatformBackend> {
    Arc::new(windows::WindowsOfficePlatformBackend::new(paths.clone()))
}
