use super::{OfficeIntegrationMode, OfficePlatformBackend, OfficePlatformStatus};
use crate::office::state::OfficePaths;
use crate::office::windows_backend::WindowsOfficeBackend;

pub struct WindowsOfficePlatformBackend {
    inner: WindowsOfficeBackend,
}

impl WindowsOfficePlatformBackend {
    pub fn new(paths: OfficePaths) -> Self {
        Self {
            inner: WindowsOfficeBackend::new(paths),
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
}
