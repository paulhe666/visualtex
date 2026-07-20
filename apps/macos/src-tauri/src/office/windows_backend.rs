#![cfg(target_os = "windows")]

use crate::office::platform::{OfficeIntegrationMode, OfficePlatformStatus};
use crate::office::state::OfficePaths;
use crate::office::windows_pipe::{locate_sidecar, WindowsPipeClient};
use serde_json::Value;
use std::fs;
use std::os::windows::process::CommandExt;
use std::path::PathBuf;
use std::process::Command;
use std::sync::Mutex;
use tauri::AppHandle;

const WORD_VSTO_KEY: &str =
    r"HKCU\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto";
const POWERPOINT_VSTO_KEY: &str =
    r"HKCU\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto";
const OLE_CATALOG_KEY: &str =
    r"HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\VisualTeX";
const WINDOWS_RUN_KEY: &str = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Run";
const WINDOWS_RUN_VALUE: &str = "VisualTeXOffice";
const CREATE_NO_WINDOW: u32 = 0x0800_0000;

fn hidden_command(program: &str) -> Command {
    let mut command = Command::new(program);
    command.creation_flags(CREATE_NO_WINDOW);
    command
}

pub struct WindowsOfficeBackend {
    paths: OfficePaths,
    mode: Mutex<OfficeIntegrationMode>,
    pipe: Option<WindowsPipeClient>,
    pipe_error: Option<String>,
}

impl WindowsOfficeBackend {
    pub fn new(app: Option<&AppHandle>, paths: OfficePaths) -> Self {
        let mode = read_mode(&paths).unwrap_or_default();
        let sidecar = locate_sidecar(app);
        let temp_root = windows_temp_root();
        let log_root = paths.root.join("logs");
        let (pipe, pipe_error) = match WindowsPipeClient::new(sidecar, temp_root, log_root) {
            Ok(client) => (Some(client), None),
            Err(error) => (None, Some(error)),
        };
        Self {
            paths,
            mode: Mutex::new(mode),
            pipe,
            pipe_error,
        }
    }

    pub fn status(&self) -> OfficePlatformStatus {
        let mode = self.mode.lock().map(|value| *value).unwrap_or_default();
        let vsto_word = registry_dword_equals(
            r"HKCU\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
            "LoadBehavior",
            3,
        );
        let vsto_powerpoint = registry_dword_equals(
            r"HKCU\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto",
            "LoadBehavior",
            3,
        );
        let ole_healthy = self.pipe.as_ref().is_some_and(WindowsPipeClient::is_healthy);
        let active_backend = match mode {
            OfficeIntegrationMode::Vsto if vsto_word && vsto_powerpoint => "vsto",
            OfficeIntegrationMode::Vsto => "unavailable-vsto",
            OfficeIntegrationMode::Ole => "ole",
            OfficeIntegrationMode::Auto if vsto_word && vsto_powerpoint => "vsto",
            OfficeIntegrationMode::Auto => "ole",
        };
        OfficePlatformStatus {
            platform: "windows".to_string(),
            mode,
            active_backend: active_backend.to_string(),
            ole_bridge_healthy: ole_healthy,
            vsto_word_healthy: vsto_word,
            vsto_powerpoint_healthy: vsto_powerpoint,
            office_catalog_registered: registry_key_exists(
                r"HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\VisualTeX",
            ),
            current_user_certificate_trusted: windows_certificate_trusted(&self.paths),
            background_start_enabled: registry_value_exists(WINDOWS_RUN_KEY, WINDOWS_RUN_VALUE),
            last_error: self.pipe_error.clone().or_else(|| {
                if mode == OfficeIntegrationMode::Vsto && !(vsto_word && vsto_powerpoint) {
                    Some("VSTO mode is selected but one or both VSTO add-ins are unavailable".to_string())
                } else {
                    None
                }
            }),
        }
    }

    pub fn set_mode(&self, mode: OfficeIntegrationMode) -> Result<OfficePlatformStatus, String> {
        apply_mode_selection(mode)?;
        write_mode_registry(mode)?;
        write_mode(&self.paths, mode)?;
        *self
            .mode
            .lock()
            .map_err(|_| "Windows Office mode lock is poisoned".to_string())? = mode;
        Ok(self.status())
    }

    pub fn request(&self, request: Value) -> Result<Value, String> {
        let status = self.status();
        if status.active_backend == "vsto" {
            return Err("The OLE Office.js manifest is disabled while VSTO mode is active".to_string());
        }
        self.pipe
            .as_ref()
            .ok_or_else(|| {
                self.pipe_error
                    .clone()
                    .unwrap_or_else(|| "Windows OLE bridge is unavailable".to_string())
            })?
            .request_value(request)
    }

    pub fn events_after(&self, cursor: u64) -> Vec<Value> {
        self.pipe
            .as_ref()
            .map(|pipe| pipe.events_after(cursor))
            .unwrap_or_default()
    }

    pub fn shutdown(&self) -> Result<(), String> {
        self.pipe.as_ref().map(WindowsPipeClient::shutdown).unwrap_or(Ok(()))
    }
}

pub fn set_background_start_enabled(enabled: bool) -> Result<(), String> {
    if !enabled {
        return registry_delete_value(WINDOWS_RUN_KEY, WINDOWS_RUN_VALUE);
    }
    let executable = std::env::current_exe()
        .map_err(|error| format!("Unable to resolve the VisualTeX executable: {error}"))?;
    let executable = executable
        .to_str()
        .ok_or_else(|| "The VisualTeX executable path is not valid UTF-8".to_string())?;
    if executable.contains('"') {
        return Err("The VisualTeX executable path contains an unsupported quote".to_string());
    }
    registry_set_string(
        WINDOWS_RUN_KEY,
        WINDOWS_RUN_VALUE,
        &format!("\"{executable}\" --office-background"),
    )
}

fn mode_path(paths: &OfficePaths) -> PathBuf {
    paths.root.join("windows-office-mode.json")
}

fn read_mode(paths: &OfficePaths) -> Option<OfficeIntegrationMode> {
    let bytes = fs::read(mode_path(paths)).ok()?;
    serde_json::from_slice::<OfficeIntegrationMode>(&bytes).ok()
}

fn write_mode(paths: &OfficePaths, mode: OfficeIntegrationMode) -> Result<(), String> {
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create Office settings directory: {error}"))?;
    let bytes = serde_json::to_vec_pretty(&mode)
        .map_err(|error| format!("Unable to serialize Office integration mode: {error}"))?;
    fs::write(mode_path(paths), bytes)
        .map_err(|error| format!("Unable to save Office integration mode: {error}"))
}

fn apply_mode_selection(mode: OfficeIntegrationMode) -> Result<(), String> {
    let vsto_installed = registry_key_exists(WORD_VSTO_KEY)
        && registry_key_exists(POWERPOINT_VSTO_KEY);
    let ole_available = ole_catalog_manifests_available();

    match mode {
        OfficeIntegrationMode::Vsto => {
            if !vsto_installed {
                return Err(
                    "Cannot enable VSTO mode until both Word and PowerPoint native add-ins are installed"
                        .to_string(),
                );
            }
            set_vsto_load_behavior(true)?;
            remove_ole_catalog()?;
        }
        OfficeIntegrationMode::Ole => {
            if !ole_available {
                return Err(
                    "Cannot enable OLE mode because the Windows Office manifests are not installed in the VisualTeX Office catalog"
                        .to_string(),
                );
            }
            set_vsto_load_behavior(false)?;
            register_ole_catalog()?;
        }
        OfficeIntegrationMode::Auto => {
            if vsto_installed {
                set_vsto_load_behavior(true)?;
                remove_ole_catalog()?;
            } else if ole_available {
                set_vsto_load_behavior(false)?;
                register_ole_catalog()?;
            } else {
                return Err(
                    "No usable Windows Office integration is installed. Install either the OLE manifests or both native add-ins first"
                        .to_string(),
                );
            }
        }
    }

    Ok(())
}

fn set_vsto_load_behavior(enabled: bool) -> Result<(), String> {
    let value = if enabled { 3 } else { 0 };
    for key in [WORD_VSTO_KEY, POWERPOINT_VSTO_KEY] {
        if registry_key_exists(key) {
            registry_set_dword(key, "LoadBehavior", value)?;
        }
    }
    Ok(())
}

fn office_catalog_path() -> PathBuf {
    std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("OfficeCatalog")
}

fn ole_catalog_manifests_available() -> bool {
    let catalog = office_catalog_path();
    catalog.join("VisualTeX.WindowsOle.Word.xml").is_file()
        && catalog
            .join("VisualTeX.WindowsOle.PowerPoint.xml")
            .is_file()
}

fn register_ole_catalog() -> Result<(), String> {
    let catalog = office_catalog_path();
    if !ole_catalog_manifests_available() {
        return Err(format!(
            "Windows Office manifests are missing from {}",
            catalog.display()
        ));
    }

    registry_set_string(OLE_CATALOG_KEY, "Id", "VisualTeX.WindowsOle")?;
    registry_set_string(OLE_CATALOG_KEY, "Url", &path_to_file_uri(&catalog))?;
    registry_set_dword(OLE_CATALOG_KEY, "Flags", 1)
}

fn remove_ole_catalog() -> Result<(), String> {
    if !registry_key_exists(OLE_CATALOG_KEY) {
        return Ok(());
    }
    let output = hidden_command("reg.exe")
        .args(["delete", OLE_CATALOG_KEY, "/f"])
        .output()
        .map_err(|error| format!("Unable to remove the Windows Office trusted catalog: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        Err(format!(
            "Unable to remove the Windows Office trusted catalog: {}",
            String::from_utf8_lossy(&output.stderr).trim()
        ))
    }
}

fn write_mode_registry(mode: OfficeIntegrationMode) -> Result<(), String> {
    let value = match mode {
        OfficeIntegrationMode::Auto => "auto",
        OfficeIntegrationMode::Ole => "ole",
        OfficeIntegrationMode::Vsto => "vsto",
    };
    registry_set_string(
        r"HKCU\Software\VisualTeX\OfficeIntegration",
        "Mode",
        value,
    )
}

fn registry_set_string(key: &str, name: &str, value: &str) -> Result<(), String> {
    registry_add(key, name, "REG_SZ", value)
}

fn registry_set_dword(key: &str, name: &str, value: u32) -> Result<(), String> {
    registry_add(key, name, "REG_DWORD", &value.to_string())
}

fn registry_delete_value(key: &str, name: &str) -> Result<(), String> {
    if !registry_value_exists(key, name) {
        return Ok(());
    }
    let output = hidden_command("reg.exe")
        .args(["delete", key, "/v", name, "/f"])
        .output()
        .map_err(|error| format!("Unable to update Windows startup state: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        Err(format!(
            "Unable to update Windows startup state: {}",
            String::from_utf8_lossy(&output.stderr).trim()
        ))
    }
}

fn registry_add(key: &str, name: &str, value_type: &str, value: &str) -> Result<(), String> {
    let output = hidden_command("reg.exe")
        .args(["add", key, "/v", name, "/t", value_type, "/d", value, "/f"])
        .output()
        .map_err(|error| format!("Unable to update Windows Office registry state: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        Err(format!(
            "Unable to update Windows Office registry state: {}",
            String::from_utf8_lossy(&output.stderr).trim()
        ))
    }
}

fn path_to_file_uri(path: &std::path::Path) -> String {
    let normalized = path.to_string_lossy().replace('\\', "/");
    let encoded = normalized
        .replace('%', "%25")
        .replace(' ', "%20")
        .replace('#', "%23");
    format!("file:///{encoded}")
}

fn windows_temp_root() -> PathBuf {
    std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("office")
        .join("temp")
}

fn windows_certificate_trusted(paths: &OfficePaths) -> bool {
    if !paths.certificate.is_file() {
        return false;
    }
    let Some(thumbprint) = registry_string_value(
        r"HKCU\Software\VisualTeX\OfficeIntegration",
        "CertificateThumbprint",
    ) else {
        return false;
    };
    hidden_command("certutil.exe")
        .args(["-user", "-store", "Root", &thumbprint])
        .output()
        .map(|output| {
            output.status.success()
                && String::from_utf8_lossy(&output.stdout)
                    .replace(' ', "")
                    .to_ascii_uppercase()
                    .contains(&thumbprint.replace(' ', "").to_ascii_uppercase())
        })
        .unwrap_or(false)
}

fn registry_string_value(key: &str, value: &str) -> Option<String> {
    let output = hidden_command("reg.exe")
        .args(["query", key, "/v", value])
        .output()
        .ok()?;
    if !output.status.success() {
        return None;
    }
    let text = String::from_utf8_lossy(&output.stdout);
    for line in text.lines() {
        if !line.contains(value) || !line.contains("REG_SZ") {
            continue;
        }
        let (_, remainder) = line.split_once("REG_SZ")?;
        let result = remainder.trim();
        if !result.is_empty() {
            return Some(result.to_string());
        }
    }
    None
}

fn registry_key_exists(key: &str) -> bool {
    hidden_command("reg.exe")
        .args(["query", key])
        .output()
        .map(|output| output.status.success())
        .unwrap_or(false)
}

fn registry_value_exists(key: &str, value: &str) -> bool {
    hidden_command("reg.exe")
        .args(["query", key, "/v", value])
        .output()
        .map(|output| output.status.success())
        .unwrap_or(false)
}

fn registry_dword_equals(key: &str, value: &str, expected: u32) -> bool {
    let output = match hidden_command("reg.exe")
        .args(["query", key, "/v", value])
        .output()
    {
        Ok(output) if output.status.success() => output,
        _ => return false,
    };
    let text = String::from_utf8_lossy(&output.stdout);
    text.split_whitespace().any(|token| {
        token
            .strip_prefix("0x")
            .and_then(|hex| u32::from_str_radix(hex, 16).ok())
            .is_some_and(|actual| actual == expected)
            || token.parse::<u32>().is_ok_and(|actual| actual == expected)
    })
}
