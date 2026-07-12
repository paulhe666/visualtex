use crate::office::formula_cache::FormulaMetadataCache;
use crate::office::sessions::SessionStore;
use axum_server::Handle;
use serde::Serialize;
use std::net::SocketAddr;
use std::path::PathBuf;
use std::sync::{Arc, Mutex, RwLock};

pub const OFFICE_BIND_IP: [u8; 4] = [127, 0, 0, 1];
pub const OFFICE_PORT: u16 = 43_127;
pub const OFFICE_PROTOCOL_VERSION: u32 = 1;
pub const OFFICE_UI_VERSION: &str = env!("CARGO_PKG_VERSION");
pub const MAX_OFFICE_REQUEST_BYTES: usize = 22 * 1024 * 1024;

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
    pub paths: Arc<OfficePaths>,
    pub install_token: Arc<String>,
    pub status: Arc<RwLock<OfficeCompanionStatus>>,
    pub server_handle: Arc<Mutex<Option<Handle<SocketAddr>>>>,
    pub session_store: SessionStore,
    pub formula_cache: FormulaMetadataCache,
    pub ocr_available: bool,
}

impl OfficeCompanionState {
    pub fn new(
        paths: OfficePaths,
        install_token: String,
        session_store: SessionStore,
        formula_cache: FormulaMetadataCache,
        ocr_available: bool,
    ) -> Self {
        let status = OfficeCompanionStatus::stopped(&paths);
        Self {
            paths: Arc::new(paths),
            install_token: Arc::new(install_token),
            status: Arc::new(RwLock::new(status)),
            server_handle: Arc::new(Mutex::new(None)),
            session_store,
            formula_cache,
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
}
