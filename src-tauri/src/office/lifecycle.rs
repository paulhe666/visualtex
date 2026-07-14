use crate::office::background;
use crate::office::certificate::{ensure_office_install, regenerate_certificate};
use crate::office::formula_cache::FormulaMetadataCache;
use crate::office::installer::{self, OfficeIntegrationStatus};
use crate::office::manifest::ManifestHost;
use crate::office::platform::{OfficeIntegrationMode, OfficePlatformStatus};
use crate::office::server;
use crate::office::sessions::SessionStore;
use crate::office::state::{OfficeCompanionState, OfficeCompanionStatus, OfficePaths};
use crate::OcrState;
#[cfg(target_os = "macos")]
use std::path::Path;
use std::path::PathBuf;
#[cfg(target_os = "windows")]
use std::process::Command;
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Manager};
use tokio::time::{sleep, Duration};

#[cfg(all(debug_assertions, target_os = "macos"))]
fn development_ui_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("dist-office-macos")
}

#[cfg(all(debug_assertions, target_os = "windows"))]
fn development_ui_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("dist-office-windows-ole")
}

#[cfg(all(
    debug_assertions,
    not(any(target_os = "macos", target_os = "windows"))
))]
fn development_ui_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("dist-office-macos")
}

fn resolve_ui_root(app: &AppHandle) -> Result<PathBuf, String> {
    if let Ok(resource) = app.path().resolve("office", BaseDirectory::Resource) {
        if resource.join("bridge").join("index.html").is_file()
            && resource.join("dialog").join("index.html").is_file()
        {
            return Ok(resource);
        }
    }
    #[cfg(debug_assertions)]
    {
        let development = development_ui_root();
        if development.join("bridge").join("index.html").is_file()
            && development.join("dialog").join("index.html").is_file()
        {
            return Ok(development);
        }
    }
    Err(
        "Office UI resources are missing. Run the platform Office build before starting VisualTeX."
            .to_string(),
    )
}

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
        ui_root: resolve_ui_root(app)?,
        root,
    };
    let install_token = ensure_office_install(&paths)?;
    // An installed (or temporarily paused) LaunchAgent means the user
    // explicitly installed Office integration. Reassert only VisualTeX's
    // GUID-prefixed manifests when the companion starts so a removed host
    // container is repaired before Office is opened again. Never clear or
    // rewrite the rest of Office's Wef cache.
    #[cfg(target_os = "macos")]
    {
        let background_status = background::status();
        let integration_configured = background_status.installed
            || (!background_status.plist_path.is_empty()
                && Path::new(&background_status.plist_path).is_file());
        if integration_configured {
            if let Err(error) = installer::install_available_manifests() {
                eprintln!("Unable to restore VisualTeX Office manifests: {error}");
            }
        }
    }
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

async fn run_blocking<T: Send + 'static>(
    operation: impl FnOnce() -> Result<T, String> + Send + 'static,
) -> Result<T, String> {
    tokio::task::spawn_blocking(operation)
        .await
        .map_err(|error| format!("Office integration task failed: {error}"))?
}

async fn wait_for_port_release() -> Result<(), String> {
    for _ in 0..30 {
        if tokio::net::TcpStream::connect(OfficeCompanionState::socket_addr())
            .await
            .is_err()
        {
            return Ok(());
        }
        sleep(Duration::from_millis(100)).await;
    }
    Err("VisualTeX Office companion did not release port 43127 in time".to_string())
}

async fn wait_for_trusted_health() -> Result<(), String> {
    let mut last_error = "VisualTeX Office companion did not become ready".to_string();
    for _ in 0..50 {
        match run_blocking(installer::verify_companion_health).await {
            Ok(()) => return Ok(()),
            Err(error) => last_error = error,
        }
        sleep(Duration::from_millis(100)).await;
    }
    Err(format!(
        "The VisualTeX Office HTTPS endpoint is not trusted or not reachable: {last_error}"
    ))
}

async fn status_for(state: &OfficeCompanionState) -> Result<OfficeIntegrationStatus, String> {
    let paths = (*state.paths).clone();
    let companion = state.snapshot();
    run_blocking(move || installer::integration_status(&paths, background::status(), companion))
        .await
}

#[tauri::command]
pub async fn get_office_integration_status(
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeIntegrationStatus, String> {
    status_for(state.inner()).await
}

#[tauri::command]
pub async fn install_office_integration(
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeIntegrationStatus, String> {
    let companion = state.inner().clone();
    let paths = (*companion.paths).clone();
    run_blocking(move || installer::trust_certificate(&paths)).await?;
    start(companion.clone());
    wait_for_trusted_health().await?;
    run_blocking(installer::install_available_manifests).await?;
    run_blocking(background::install_launch_agent).await?;
    status_for(&companion).await
}

#[tauri::command]
pub async fn repair_office_integration(
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeIntegrationStatus, String> {
    install_office_integration(state).await
}

#[tauri::command]
pub async fn uninstall_office_integration(
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeIntegrationStatus, String> {
    let companion = state.inner().clone();
    run_blocking(installer::uninstall_manifests).await?;
    run_blocking(background::uninstall_launch_agent).await?;
    status_for(&companion).await
}

#[tauri::command]
pub async fn regenerate_office_certificate(
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeIntegrationStatus, String> {
    let companion = state.inner().clone();
    server::stop(&companion)?;
    companion.update_status(|status| status.running = false);
    wait_for_port_release().await?;
    let paths = (*companion.paths).clone();
    run_blocking(move || {
        installer::remove_trusted_certificate(&paths)?;
        regenerate_certificate(&paths)?;
        installer::trust_certificate(&paths)
    })
    .await?;
    start(companion.clone());
    wait_for_trusted_health().await?;
    status_for(&companion).await
}

#[tauri::command]
pub fn get_office_platform_status(
    state: tauri::State<'_, OfficeCompanionState>,
) -> OfficePlatformStatus {
    state.platform_backend.status()
}

#[tauri::command]
pub fn set_office_background_start(enabled: bool) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        if enabled {
            background::install_launch_agent().map(|_| ())
        } else {
            background::uninstall_launch_agent().map(|_| ())
        }
    }
    #[cfg(target_os = "windows")]
    {
        crate::office::windows_backend::set_background_start_enabled(enabled)
    }
    #[cfg(not(any(target_os = "macos", target_os = "windows")))]
    {
        let _ = enabled;
        Err("Office background startup is supported only on macOS and Windows".to_string())
    }
}

#[tauri::command]
pub fn set_office_integration_mode(
    mode: OfficeIntegrationMode,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficePlatformStatus, String> {
    state.platform_backend.set_mode(mode)
}

#[cfg(target_os = "windows")]
fn run_windows_script(app: &AppHandle, script_name: &str, arguments: &[&str]) -> Result<(), String> {
    let script = app
        .path()
        .resolve(
            format!("scripts/{script_name}"),
            BaseDirectory::Resource,
        )
        .map_err(|error| format!("Unable to resolve Windows Office script: {error}"))?;
    if !script.is_file() {
        return Err(format!(
            "Windows Office script is missing: {}",
            script.display()
        ));
    }
    let status = Command::new("powershell.exe")
        .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-File"])
        .arg(&script)
        .args(arguments)
        .status()
        .map_err(|error| format!("Unable to start Windows Office script: {error}"))?;
    if status.success() {
        Ok(())
    } else {
        Err(format!(
            "Windows Office script {} failed with {status}",
            script.display()
        ))
    }
}

#[tauri::command]
pub fn install_windows_ole_integration(
    app: AppHandle,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficePlatformStatus, String> {
    #[cfg(target_os = "windows")]
    {
        run_windows_script(&app, "install_windows_ole.ps1", &[])?;
        return state
            .platform_backend
            .set_mode(OfficeIntegrationMode::Ole);
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (app, state);
        Err("Windows OLE integration can be installed only on Windows".to_string())
    }
}

#[tauri::command]
pub fn uninstall_windows_ole_integration(
    app: AppHandle,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficePlatformStatus, String> {
    #[cfg(target_os = "windows")]
    {
        run_windows_script(&app, "uninstall_windows_ole.ps1", &[])?;
        return state
            .platform_backend
            .set_mode(OfficeIntegrationMode::Auto);
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (app, state);
        Err("Windows OLE integration can be removed only on Windows".to_string())
    }
}

#[tauri::command]
pub fn repair_windows_office_integration(
    app: AppHandle,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficePlatformStatus, String> {
    #[cfg(target_os = "windows")]
    {
        run_windows_script(&app, "install_windows_ole.ps1", &[])?;
        state
            .platform_backend
            .set_mode(OfficeIntegrationMode::Ole)
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (app, state);
        Err("Windows Office repair is available only on Windows".to_string())
    }
}

#[tauri::command]
pub fn open_word() -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        return Command::new("cmd.exe")
            .args(["/C", "start", "", "winword.exe"])
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("Unable to launch Microsoft Word: {error}"));
    }
    #[cfg(not(target_os = "windows"))]
    installer::open_office_application(ManifestHost::Word)
}

#[tauri::command]
pub fn open_powerpoint() -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        return Command::new("cmd.exe")
            .args(["/C", "start", "", "powerpnt.exe"])
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("Unable to launch Microsoft PowerPoint: {error}"));
    }
    #[cfg(not(target_os = "windows"))]
    installer::open_office_application(ManifestHost::PowerPoint)
}
