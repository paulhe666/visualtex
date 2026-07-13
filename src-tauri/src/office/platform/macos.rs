use super::{OfficeIntegrationMode, OfficePlatformBackend, OfficePlatformStatus};
use crate::office::state::OfficePaths;
use serde_json::Value;

pub struct MacOfficePlatformBackend {
    paths: OfficePaths,
}

impl MacOfficePlatformBackend {
    pub fn new(paths: OfficePaths) -> Self {
        Self { paths }
    }
}

impl OfficePlatformBackend for MacOfficePlatformBackend {
    fn status(&self) -> OfficePlatformStatus {
        OfficePlatformStatus {
            platform: "macos".to_string(),
            mode: OfficeIntegrationMode::Auto,
            active_backend: "officejs-applescript".to_string(),
            ole_bridge_healthy: false,
            vsto_word_healthy: false,
            vsto_powerpoint_healthy: false,
            office_catalog_registered: self.paths.install.exists(),
            current_user_certificate_trusted: self.paths.certificate.exists(),
            background_start_enabled: crate::office::background::status().installed,
            last_error: None,
        }
    }

    fn set_mode(&self, _mode: OfficeIntegrationMode) -> Result<OfficePlatformStatus, String> {
        Err("macOS Office integration mode is fixed to Office.js + AppleScript".to_string())
    }

    fn request(&self, _request: Value) -> Result<Value, String> {
        Err("Windows Office bridge requests are unavailable on macOS".to_string())
    }

    fn events_after(&self, _cursor: u64) -> Vec<Value> {
        Vec::new()
    }

    fn shutdown(&self) -> Result<(), String> {
        Ok(())
    }
}
