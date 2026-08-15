use crate::office::background;
use crate::office::certificate::{
    certificate_sha1_thumbprint, ensure_office_install, regenerate_certificate,
};
use crate::office::formula_cache::FormulaMetadataCache;
use crate::office::installer::{self, OfficeIntegrationStatus};
#[cfg(not(target_os = "windows"))]
use crate::office::manifest::ManifestHost;
use crate::office::platform::{OfficeIntegrationMode, OfficePlatformStatus};
use crate::office::server;
use crate::office::sessions::SessionStore;
use crate::office::state::{
    append_office_log, OfficeCompanionState, OfficeCompanionStatus, OfficePaths, OFFICE_PORT,
    OFFICE_PROTOCOL_VERSION,
};
use crate::OcrState;
#[cfg(target_os = "macos")]
use std::path::Path;
use std::path::PathBuf;
#[cfg(target_os = "windows")]
use std::os::windows::process::CommandExt;
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
        .join("dist-office-windows-native")
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
        #[cfg(target_os = "windows")]
        if resource.join("dialog").join("index.html").is_file() {
            return Ok(resource);
        }
        #[cfg(not(target_os = "windows"))]
        if resource.join("bridge").join("index.html").is_file()
            && resource.join("dialog").join("index.html").is_file()
        {
            return Ok(resource);
        }
    }
    #[cfg(debug_assertions)]
    {
        let development = development_ui_root();
        #[cfg(target_os = "windows")]
        if development.join("dialog").join("index.html").is_file() {
            return Ok(development);
        }
        #[cfg(not(target_os = "windows"))]
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

struct PreparedOfficeConfiguration {
    paths: OfficePaths,
    install_token: String,
}

fn prepare_office_configuration(
    app: &AppHandle,
    phase: &str,
) -> Result<PreparedOfficeConfiguration, String> {
    let app_data = app.path().app_data_dir().map_err(|error| {
        format!("Unable to resolve VisualTeX application data directory: {error}")
    })?;
    let root = app_data.join("office");
    let mut paths = OfficePaths {
        certificate: root.join("localhost-cert.pem"),
        private_key: root.join("localhost-key.pem"),
        certificate_metadata: root.join("certificate.json"),
        install: root.join("install.json"),
        sessions: root.join("sessions"),
        recovery: root.join("recovery"),
        formula_cache: root.join("formulas"),
        ui_root: PathBuf::new(),
        root,
    };
    let executable = std::env::current_exe()
        .map_err(|error| format!("Unable to resolve the VisualTeX executable: {error}"))?;
    append_office_log(
        &paths,
        "startup.log",
        &format!(
            "{phase} begin pid={} executable={} app_data_root={} office_root={}",
            std::process::id(),
            executable.display(),
            app_data.display(),
            paths.root.display()
        ),
    );

    paths.ui_root = match resolve_ui_root(app) {
        Ok(path) => {
            append_office_log(
                &paths,
                "startup.log",
                &format!("{phase} Office UI resource root resolved: {}", path.display()),
            );
            path
        }
        Err(error) => {
            append_office_log(
                &paths,
                "startup.log",
                &format!("{phase} Office UI resource resolution failed: {error}"),
            );
            return Err(error);
        }
    };

    let install_token = ensure_office_install(&paths).map_err(|error| {
        append_office_log(
            &paths,
            "startup.log",
            &format!("{phase} Office data/certificate initialization failed: {error}"),
        );
        error
    })?;
    let certificate_thumbprint = certificate_sha1_thumbprint(&paths).map_err(|error| {
        append_office_log(
            &paths,
            "startup.log",
            &format!("{phase} certificate thumbprint calculation failed: {error}"),
        );
        error
    })?;

    #[cfg(target_os = "windows")]
    crate::office::windows_backend::write_shared_configuration(
        &executable,
        &app_data,
        &paths,
        &certificate_thumbprint,
        OFFICE_PORT,
        OFFICE_PROTOCOL_VERSION,
    )
    .map_err(|error| {
        append_office_log(
            &paths,
            "startup.log",
            &format!("{phase} OfficeIntegration registry write failed: {error}"),
        );
        error
    })?;

    append_office_log(
        &paths,
        "startup.log",
        &format!(
            "{phase} Office configuration ready certificate={} private_key={} thumbprint={} port={} protocol={} install_json={} office_dialog={}",
            paths.certificate.display(),
            paths.private_key.display(),
            certificate_thumbprint,
            OFFICE_PORT,
            OFFICE_PROTOCOL_VERSION,
            paths.install.display(),
            paths.ui_root.join("dialog").join("index.html").display()
        ),
    );

    Ok(PreparedOfficeConfiguration {
        paths,
        install_token,
    })
}

pub fn bootstrap_configuration(app: &AppHandle) -> Result<(), String> {
    let prepared = prepare_office_configuration(app, "office-bootstrap")?;
    let required = [
        &prepared.paths.certificate,
        &prepared.paths.private_key,
        &prepared.paths.certificate_metadata,
        &prepared.paths.install,
    ];
    let missing = required
        .iter()
        .filter(|path| !path.is_file())
        .map(|path| path.display().to_string())
        .collect::<Vec<_>>();
    if !missing.is_empty() || prepared.install_token.len() != 64 {
        let error = format!(
            "Office bootstrap verification failed: missing={} install_token_length={}",
            missing.join("|"),
            prepared.install_token.len()
        );
        append_office_log(&prepared.paths, "startup.log", &error);
        return Err(error);
    }
    append_office_log(
        &prepared.paths,
        "startup.log",
        "office-bootstrap completed without creating a WebView, starting the companion, or scheduling OCR/Office editor prewarm",
    );
    Ok(())
}

pub fn initialize(app: &AppHandle, ocr: OcrState) -> Result<OfficeCompanionState, String> {
    let prepared = prepare_office_configuration(app, "startup")?;
    let paths = prepared.paths;

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
    append_office_log(&paths, "startup.log", "startup initialization completed");
    Ok(OfficeCompanionState::new(
        Some(app.clone()),
        ocr,
        paths,
        prepared.install_token,
        session_store,
        formula_cache,
        ocr_worker_available(app),
    ))
}

pub fn start(state: OfficeCompanionState) {
    if state.snapshot().running {
        append_office_log(&state.paths, "companion.log", "start skipped: companion already marked running");
        return;
    }
    append_office_log(
        &state.paths,
        "companion.log",
        &format!(
            "companion task requested by pid={} (Office editor WebView remains on-demand)",
            std::process::id()
        ),
    );

    let service = state.clone();
    tauri::async_runtime::spawn(async move {
        if let Err(error) = server::run(service.clone()).await {
            append_office_log(
                &service.paths,
                "companion.log",
                &format!("companion terminated with error: {error}"),
            );
            service.update_status(|status| {
                status.running = false;
                status.last_error = Some(error);
            });
        }
    });
}

#[tauri::command]
pub fn set_app_theme(
    theme: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> String {
    state.set_current_theme(&theme)
}

#[tauri::command]
pub fn set_app_editor_layout(
    editor_layout: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> String {
    state.set_current_editor_layout(&editor_layout)
}

#[tauri::command]
pub fn set_app_editor_preferences(
    preferences: serde_json::Value,
    state: tauri::State<'_, OfficeCompanionState>,
) -> serde_json::Value {
    state.set_current_editor_preferences(preferences)
}

#[tauri::command]
pub fn set_powerpoint_default_font_size(
    font_size_pt: f64,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<f64, String> {
    state.set_powerpoint_default_font_size_pt(font_size_pt)
}

#[tauri::command]
pub fn get_powerpoint_default_font_size(
    state: tauri::State<'_, OfficeCompanionState>,
) -> f64 {
    state.powerpoint_default_font_size_pt()
}

#[tauri::command]
pub fn get_mathtype_double_click_edit_enabled(
    state: tauri::State<'_, OfficeCompanionState>,
) -> bool {
    state.mathtype_double_click_edit_enabled()
}

#[tauri::command]
pub fn set_mathtype_double_click_edit_enabled(
    enabled: bool,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<bool, String> {
    state.set_mathtype_double_click_edit_enabled(enabled)
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
const CREATE_NO_WINDOW: u32 = 0x0800_0000;

#[cfg(target_os = "windows")]
fn hidden_windows_command(program: impl AsRef<std::ffi::OsStr>) -> Command {
    let mut command = Command::new(program);
    command.creation_flags(CREATE_NO_WINDOW);
    command
}

#[cfg(target_os = "windows")]
fn powershell_compatible_path(path: &std::path::Path) -> PathBuf {
    let raw = path.as_os_str().to_string_lossy();
    if let Some(rest) = raw.strip_prefix(r"\\?\UNC\") {
        return PathBuf::from(format!(r"\\{rest}"));
    }
    if let Some(rest) = raw.strip_prefix(r"\\?\") {
        return PathBuf::from(rest);
    }
    path.to_path_buf()
}

#[cfg(target_os = "windows")]
fn windows_office_root() -> PathBuf {
    std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("office")
}

#[cfg(target_os = "windows")]
fn windows_office_log_root() -> PathBuf {
    windows_office_root().join("install-logs")
}

#[cfg(target_os = "windows")]
fn latest_bootstrap_log_tail() -> String {
    let root = windows_office_log_root();
    let Ok(entries) = std::fs::read_dir(&root) else {
        return format!("No bootstrap log directory exists at {}", root.display());
    };
    let latest = entries
        .filter_map(Result::ok)
        .filter(|entry| {
            entry
                .file_name()
                .to_string_lossy()
                .starts_with("vsto-bootstrap-")
                && entry.path().extension().is_some_and(|value| value == "log")
        })
        .filter_map(|entry| {
            let modified = entry.metadata().ok()?.modified().ok()?;
            Some((modified, entry.path()))
        })
        .max_by_key(|(modified, _)| *modified)
        .map(|(_, path)| path);
    let Some(path) = latest else {
        return format!("No vsto-bootstrap log exists at {}", root.display());
    };
    let Ok(text) = std::fs::read_to_string(&path) else {
        return format!("Unable to read latest bootstrap log: {}", path.display());
    };
    let lines = text.lines().collect::<Vec<_>>();
    let start = lines.len().saturating_sub(80);
    format!(
        "Latest bootstrap log: {}\n{}",
        path.display(),
        lines[start..].join("\n")
    )
}

#[cfg(target_os = "windows")]
fn current_visualtex_executable() -> Result<String, String> {
    std::env::current_exe()
        .map(|path| powershell_compatible_path(&path).to_string_lossy().to_string())
        .map_err(|error| format!("Unable to resolve the current VisualTeX.exe path: {error}"))
}

#[cfg(target_os = "windows")]
fn run_windows_script(
    app: &AppHandle,
    script_name: &str,
    arguments: &[String],
) -> Result<String, String> {
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
    let powershell_script = powershell_compatible_path(&script);
    let powershell = crate::ocr_install::windows_powershell_executable()?;
    let output = hidden_windows_command(&powershell)
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-WindowStyle",
            "Hidden",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
        ])
        .arg(&powershell_script)
        .args(arguments)
        .output()
        .map_err(|error| format!("Unable to start Windows Office script: {error}"))?;
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
    if output.status.success() {
        return Ok(stdout);
    }
    Err(format!(
        "Windows Office script failed: {}\nStatus: {}\n\nstdout:\n{}\n\nstderr:\n{}\n\n{}",
        script.display(),
        output.status,
        if stdout.is_empty() { "<empty>" } else { &stdout },
        if stderr.is_empty() { "<empty>" } else { &stderr },
        latest_bootstrap_log_tail()
    ))
}

#[tauri::command]
pub fn install_windows_ole_integration(
    app: AppHandle,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficePlatformStatus, String> {
    #[cfg(target_os = "windows")]
    {
        let executable = current_visualtex_executable()?;
        let arguments = vec!["-VisualTeXPath".to_string(), executable];
        run_windows_script(
            &app,
            "ensure_windows_office_certificate.ps1",
            &arguments,
        )?;
        run_windows_script(&app, "install_windows_vsto.ps1", &arguments)?;
        return state
            .platform_backend
            .set_mode(OfficeIntegrationMode::Vsto);
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (app, state);
        Err("Windows native Office integration can be installed only on Windows".to_string())
    }
}

#[tauri::command]
pub fn uninstall_windows_ole_integration(
    app: AppHandle,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficePlatformStatus, String> {
    #[cfg(target_os = "windows")]
    {
        run_windows_script(&app, "uninstall_windows_vsto.ps1", &[])?;
        return state
            .platform_backend
            .set_mode(OfficeIntegrationMode::Auto);
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (app, state);
        Err("Windows native Office integration can be removed only on Windows".to_string())
    }
}

#[tauri::command]
pub fn repair_windows_office_integration(
    app: AppHandle,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficePlatformStatus, String> {
    #[cfg(target_os = "windows")]
    {
        let executable = current_visualtex_executable()?;
        let arguments = vec!["-VisualTeXPath".to_string(), executable];
        run_windows_script(
            &app,
            "ensure_windows_office_certificate.ps1",
            &arguments,
        )?;
        run_windows_script(&app, "install_windows_vsto.ps1", &arguments)?;
        state
            .platform_backend
            .set_mode(OfficeIntegrationMode::Vsto)
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (app, state);
        Err("Windows Office repair is available only on Windows".to_string())
    }
}

#[tauri::command]
pub fn test_windows_office_runtime(
    app: AppHandle,
    state: tauri::State<'_, OfficeCompanionState>,
    force_close_office: bool,
) -> Result<OfficePlatformStatus, String> {
    #[cfg(target_os = "windows")]
    {
        let executable = current_visualtex_executable()?;
        let mut arguments = vec!["-VisualTeXPath".to_string(), executable];
        if force_close_office {
            arguments.push("-ForceCloseOffice".to_string());
        }
        run_windows_script(&app, "test_windows_office_runtime.ps1", &arguments)?;
        return Ok(state.platform_backend.status());
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (app, state, force_close_office);
        Err("Windows Office runtime verification is available only on Windows".to_string())
    }
}

#[tauri::command]
pub fn open_windows_office_logs() -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        let root = windows_office_root();
        std::fs::create_dir_all(&root).map_err(|error| {
            format!("Unable to create Windows Office log directory {}: {error}", root.display())
        })?;
        return hidden_windows_command("explorer.exe")
            .arg(&root)
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("Unable to open Windows Office log directory: {error}"));
    }
    #[cfg(not(target_os = "windows"))]
    {
        Err("Windows Office logs are available only on Windows".to_string())
    }
}

#[tauri::command]
pub fn open_word() -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        return hidden_windows_command("cmd.exe")
            .args(["/D", "/C", "start", "", "winword.exe"])
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
        return hidden_windows_command("cmd.exe")
            .args(["/D", "/C", "start", "", "powerpnt.exe"])
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("Unable to launch Microsoft PowerPoint: {error}"));
    }
    #[cfg(not(target_os = "windows"))]
    installer::open_office_application(ManifestHost::PowerPoint)
}
