use crate::office::certificate::{ensure_office_install, regenerate_certificate};
use crate::office::server;
use crate::office::sessions::SessionStore;
use crate::office::state::{OfficeCompanionState, OfficeCompanionStatus, OfficePaths};
use std::path::PathBuf;
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Manager};

fn development_ui_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("dist-office")
}

fn resolve_ui_root(app: &AppHandle) -> Result<PathBuf, String> {
    if let Ok(resource) = app.path().resolve("office", BaseDirectory::Resource) {
        if resource.join("bridge").join("index.html").is_file()
            && resource.join("dialog").join("index.html").is_file()
        {
            return Ok(resource);
        }
    }
    let development = development_ui_root();
    if development.join("bridge").join("index.html").is_file()
        && development.join("dialog").join("index.html").is_file()
    {
        return Ok(development);
    }
    Err(
        "Office UI resources are missing. Run `npm run build:office` before starting VisualTeX."
            .to_string(),
    )
}

fn ocr_worker_available(app: &AppHandle) -> bool {
    app.path()
        .resolve("ocr/worker.py", BaseDirectory::Resource)
        .map(|path| path.is_file())
        .unwrap_or(false)
        || PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("ocr")
            .join("worker.py")
            .is_file()
}

pub fn initialize(app: &AppHandle) -> Result<OfficeCompanionState, String> {
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
        ui_root: resolve_ui_root(app)?,
        root,
    };
    let install_token = ensure_office_install(&paths)?;
    let session_store = SessionStore::new(&paths).map_err(|error| error.to_string())?;
    Ok(OfficeCompanionState::new(
        paths,
        install_token,
        session_store,
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
pub fn regenerate_office_certificate(
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeCompanionStatus, String> {
    if state.snapshot().running {
        return Err(
            "Stop the VisualTeX Office companion before regenerating its certificate.".to_string(),
        );
    }
    regenerate_certificate(&state.paths)?;
    Ok(state.snapshot())
}
