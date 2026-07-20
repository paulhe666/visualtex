use super::{OfficeIntegrationMode, OfficePlatformBackend, OfficePlatformStatus};
use crate::office::state::OfficePaths;
use crate::office::windows_backend::WindowsOfficeBackend;
use serde_json::Value;
use tauri::AppHandle;

pub struct WindowsOfficePlatformBackend {
    inner: WindowsOfficeBackend,
}

impl WindowsOfficePlatformBackend {
    pub fn new(app: Option<&AppHandle>, paths: OfficePaths) -> Self {
        Self {
            inner: WindowsOfficeBackend::new(app, paths),
        }
    }
}

impl OfficePlatformBackend for WindowsOfficePlatformBackend {
    fn status(&self) -> OfficePlatformStatus {
        self.inner.status()
    }

    fn set_mode(&self, mode: OfficeIntegrationMode) -> Result<OfficePlatformStatus, String> {
        self.inner.set_mode(mode)
    }

    fn request(&self, request: Value) -> Result<Value, String> {
        self.inner.request(request)
    }

    fn events_after(&self, cursor: u64) -> Vec<Value> {
        self.inner.events_after(cursor)
    }

    fn shutdown(&self) -> Result<(), String> {
        self.inner.shutdown()
    }
}
