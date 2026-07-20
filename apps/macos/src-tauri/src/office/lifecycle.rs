use crate::office::background;
use crate::office::certificate::ensure_companion_runtime;
use crate::office::formula_cache::FormulaMetadataCache;
use crate::office::server;
use crate::office::sessions::SessionStore;
use crate::office::state::{OfficeCompanionState, OfficeCompanionStatus, OfficePaths};
use crate::OcrState;
use std::path::PathBuf;
use std::process::Command;
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Manager};

fn ocr_worker_available(app: &AppHandle) -> bool {
    let bundled = app
        .path()
        .resolve("ocr/worker.py", BaseDirectory::Resource)
        .map(|path| path.is_file())
        .unwrap_or(false);
    #[cfg(debug_assertions)]
    {
        bundled
            || PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("ocr")
                .join("worker.py")
                .is_file()
    }
    #[cfg(not(debug_assertions))]
    bundled
}

pub fn initialize(app: &AppHandle, ocr: OcrState) -> Result<OfficeCompanionState, String> {
    let app_data = app.path().app_data_dir().map_err(|error| {
        format!("Unable to resolve VisualTeX application data directory: {error}")
    })?;
    let root = app_data.join("office");
    let paths = OfficePaths {
        certificate: root.join("localhost-cert.pem"),
        private_key: root.join("localhost-key.pem"),
        certificate_metadata: root.join("certificate.json"),
        install: root.join("install.json"),
        sessions: root.join("sessions"),
        recovery: root.join("recovery"),
        formula_cache: root.join("formulas"),
        root,
    };
    let install_token = ensure_companion_runtime(&paths)?;
    let session_store = SessionStore::new(&paths).map_err(|error| error.to_string())?;
    let formula_cache = FormulaMetadataCache::new(&paths).map_err(|error| error.to_string())?;
    Ok(OfficeCompanionState::new(
        Some(app.clone()),
        ocr,
        paths,
        install_token,
        session_store,
        formula_cache,
        ocr_worker_available(app),
    ))
}

pub fn start(state: OfficeCompanionState) {
    if state.snapshot().running {
        return;
    }
    let service = state.clone();
    tauri::async_runtime::spawn(async move {
        if let Err(error) = server::run(service.clone()).await {
            service.update_status(|status| {
                status.running = false;
                status.last_error = Some(error);
            });
        }
    });
}

#[tauri::command]
pub fn get_office_companion_status(
    state: tauri::State<'_, OfficeCompanionState>,
) -> OfficeCompanionStatus {
    state.snapshot()
}

#[tauri::command]
pub fn start_office_companion(
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeCompanionStatus, String> {
    if state.snapshot().running {
        return Ok(state.snapshot());
    }
    start(state.inner().clone());
    Ok(state.snapshot())
}

#[tauri::command]
pub fn stop_office_companion(
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeCompanionStatus, String> {
    server::stop(state.inner())?;
    state.update_status(|status| status.running = false);
    Ok(state.snapshot())
}

#[tauri::command]
pub fn set_office_background_start(enabled: bool) -> Result<(), String> {
    if enabled {
        background::install_launch_agent().map(|_| ())
    } else {
        background::uninstall_launch_agent().map(|_| ())
    }
}

fn open_office_application(name: &str) -> Result<(), String> {
    Command::new("/usr/bin/open")
        .args(["-a", name])
        .spawn()
        .map(|_| ())
        .map_err(|error| format!("Unable to launch {name}: {error}"))
}

#[tauri::command]
pub fn open_word() -> Result<(), String> {
    open_office_application("Microsoft Word")
}

#[tauri::command]
pub fn open_powerpoint() -> Result<(), String> {
    open_office_application("Microsoft PowerPoint")
}
