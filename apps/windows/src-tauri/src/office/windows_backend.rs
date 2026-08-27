#![cfg(target_os = "windows")]

use crate::office::platform::{OfficeIntegrationMode, OfficePlatformStatus};
use crate::office::state::OfficePaths;
use crate::office::windows_pipe::{locate_sidecar, WindowsPipeClient};
use serde_json::Value;
use std::fs;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::Mutex;
use tauri::AppHandle;

const WORD_VSTO_KEY: &str =
    r"HKLM\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto";
const POWERPOINT_VSTO_KEY: &str =
    r"HKLM\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto";
const OLE_LOCAL_SERVER_KEY: &str =
    r"HKLM\Software\Classes\CLSID\{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}\LocalServer32";
const OFFICE_MODE_KEY: &str = r"HKCU\Software\VisualTeX\OfficeIntegration";
const WORD_USER_PREFERENCES_KEY: &str = r"HKCU\Software\VisualTeX\Word";
const WORD_DEFAULT_NUMBERED_VALUE: &str = "DefaultDisplayEquationNumbered";
const WORD_DEFAULT_NUMBER_FORMAT_VALUE: &str = "DefaultEquationNumberFormat";
const WORD_VSTO_CLSID_KEY: &str =
    r"HKLM\Software\Classes\CLSID\{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}\InprocServer32";
const POWERPOINT_VSTO_CLSID_KEY: &str =
    r"HKLM\Software\Classes\CLSID\{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}\InprocServer32";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum RegistryView {
    Native,
    Registry32,
    Registry64,
}

impl RegistryView {
    fn reg_flag(self) -> Option<&'static str> {
        match self {
            Self::Native => None,
            Self::Registry32 => Some("/reg:32"),
            Self::Registry64 => Some("/reg:64"),
        }
    }
}
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
        let office_registry_view = resolve_office_registry_view();
        let word_load_enabled = registry_dword_equals_in_view(
            WORD_VSTO_KEY,
            "LoadBehavior",
            3,
            office_registry_view,
        );
        let powerpoint_load_enabled = registry_dword_equals_in_view(
            POWERPOINT_VSTO_KEY,
            "LoadBehavior",
            3,
            office_registry_view,
        );
        let word_files_present =
            registered_addin_file_exists(WORD_VSTO_KEY, office_registry_view);
        let powerpoint_files_present =
            registered_addin_file_exists(POWERPOINT_VSTO_KEY, office_registry_view);
        let word_registry_complete = addin_registry_complete(
            WORD_VSTO_KEY,
            WORD_VSTO_CLSID_KEY,
            "VisualTeX.WordVsto.ThisAddIn",
            office_registry_view,
        );
        let powerpoint_registry_complete = addin_registry_complete(
            POWERPOINT_VSTO_KEY,
            POWERPOINT_VSTO_CLSID_KEY,
            "VisualTeX.PowerPointVsto.ThisAddIn",
            office_registry_view,
        );
        let vsto_word = word_files_present && word_registry_complete && word_load_enabled;
        let vsto_powerpoint =
            powerpoint_files_present && powerpoint_registry_complete && powerpoint_load_enabled;
        let ole_local_server_healthy =
            native_ole_local_server_healthy(office_registry_view);
        let static_install_verified = registry_dword_equals(
            OFFICE_MODE_KEY,
            "FilesAndRegistryVerified",
            1,
        ) && vsto_word
            && vsto_powerpoint
            && ole_local_server_healthy;
        let word_connected = registry_dword_equals(OFFICE_MODE_KEY, "WordConnected", 1);
        let powerpoint_connected =
            registry_dword_equals(OFFICE_MODE_KEY, "PowerPointConnected", 1);
        let connection_verification_attempted = registry_dword_equals(
            OFFICE_MODE_KEY,
            "OfficeConnectionVerificationAttempted",
            1,
        );
        let companion_process_running =
            registry_dword_equals(OFFICE_MODE_KEY, "CompanionProcessRunning", 1);
        let companion_port_listening =
            registry_dword_equals(OFFICE_MODE_KEY, "CompanionPortListening", 1);
        let companion_https_healthy =
            registry_dword_equals(OFFICE_MODE_KEY, "CompanionHttpsHealthy", 1);
        let companion_certificate_matches =
            registry_dword_equals(OFFICE_MODE_KEY, "CompanionCertificateMatches", 1);
        let companion_protocol_matches =
            registry_dword_equals(OFFICE_MODE_KEY, "CompanionProtocolMatches", 1);
        let office_runtime_verified = registry_dword_equals(
            OFFICE_MODE_KEY,
            "OfficeRuntimeVerified",
            1,
        ) && companion_process_running
            && companion_port_listening
            && companion_https_healthy
            && companion_certificate_matches
            && companion_protocol_matches;
        let ole_bridge_healthy = self.pipe.as_ref().is_some_and(WindowsPipeClient::is_healthy);
        let active_backend = if static_install_verified {
            "vsto"
        } else {
            "unavailable-vsto"
        };

        let last_error = if !word_files_present || !powerpoint_files_present {
            Some("One or both native Office add-in assemblies are missing".to_string())
        } else if !word_registry_complete || !powerpoint_registry_complete {
            Some("One or both native Office add-in registry registrations are incomplete".to_string())
        } else if !word_load_enabled || !powerpoint_load_enabled {
            Some("One or both native Office add-ins are not allowed to load (LoadBehavior != 3)".to_string())
        } else if !ole_local_server_healthy {
            Some("The native Formula OLE LocalServer registration is missing or invalid".to_string())
        } else if !office_runtime_verified {
            registry_string_value(OFFICE_MODE_KEY, "LastRuntimeError")
                .filter(|value| !value.trim().is_empty())
                .or_else(|| {
                    Some("Static installation passed, but companion runtime validation is incomplete".to_string())
                })
        } else if !(word_connected && powerpoint_connected) {
            if connection_verification_attempted {
                registry_string_value(OFFICE_MODE_KEY, "LastRuntimeError")
                    .filter(|value| !value.trim().is_empty())
                    .or_else(|| {
                        Some("Word and PowerPoint connection verification did not complete successfully".to_string())
                    })
            } else {
                None
            }
        } else {
            self.pipe_error.clone()
        };

        OfficePlatformStatus {
            platform: "windows".to_string(),
            mode,
            active_backend: active_backend.to_string(),
            ole_bridge_healthy,
            ole_local_server_healthy,
            static_install_verified,
            word_files_present,
            word_registry_complete,
            word_load_enabled,
            powerpoint_files_present,
            powerpoint_registry_complete,
            powerpoint_load_enabled,
            vsto_word_healthy: vsto_word,
            vsto_powerpoint_healthy: vsto_powerpoint,
            word_connected,
            powerpoint_connected,
            connection_verification_attempted,
            companion_process_running,
            companion_port_listening,
            companion_https_healthy,
            companion_certificate_matches,
            companion_protocol_matches,
            office_runtime_verified,
            current_user_certificate_trusted: windows_certificate_trusted(&self.paths),
            background_start_enabled: registry_value_exists(WINDOWS_RUN_KEY, WINDOWS_RUN_VALUE),
            last_error,
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
        self.pipe
            .as_ref()
            .ok_or_else(|| {
                self.pipe_error
                    .clone()
                    .unwrap_or_else(|| "Windows Office bridge is unavailable".to_string())
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

pub fn write_shared_configuration(
    executable: &Path,
    app_data_root: &Path,
    paths: &OfficePaths,
    certificate_thumbprint: &str,
    companion_port: u16,
    protocol_version: u32,
) -> Result<(), String> {
    registry_set_string(
        OFFICE_MODE_KEY,
        "ExecutablePath",
        &executable.to_string_lossy(),
    )?;
    registry_set_string(
        OFFICE_MODE_KEY,
        "AppDataRoot",
        &app_data_root.to_string_lossy(),
    )?;
    registry_set_string(
        OFFICE_MODE_KEY,
        "CertificatePath",
        &paths.certificate.to_string_lossy(),
    )?;
    registry_set_string(
        OFFICE_MODE_KEY,
        "CertificateThumbprint",
        certificate_thumbprint,
    )?;
    registry_set_dword(OFFICE_MODE_KEY, "CompanionPort", u32::from(companion_port))?;
    registry_set_dword(OFFICE_MODE_KEY, "ProtocolVersion", protocol_version)
}

pub fn background_start_enabled() -> bool {
    registry_value_exists(WINDOWS_RUN_KEY, WINDOWS_RUN_VALUE)
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

fn resolve_office_registry_view() -> RegistryView {
    let configured = match registry_string_value(OFFICE_MODE_KEY, "OfficePlatform")
        .as_deref()
        .map(str::trim)
    {
        Some("x86") => Some(RegistryView::Registry32),
        Some("x64") => Some(RegistryView::Registry64),
        _ => None,
    };

    if configured.is_some_and(office_registration_exists_in_view) {
        return configured.unwrap_or(RegistryView::Native);
    }
    if office_registration_exists_in_view(RegistryView::Registry64) {
        return RegistryView::Registry64;
    }
    if office_registration_exists_in_view(RegistryView::Registry32) {
        return RegistryView::Registry32;
    }
    configured.unwrap_or(RegistryView::Native)
}

fn office_registration_exists_in_view(view: RegistryView) -> bool {
    registry_key_exists_in_view(WORD_VSTO_KEY, view)
        || registry_key_exists_in_view(POWERPOINT_VSTO_KEY, view)
        || registry_key_exists_in_view(OLE_LOCAL_SERVER_KEY, view)
}

fn apply_mode_selection(mode: OfficeIntegrationMode) -> Result<(), String> {
    let office_registry_view = resolve_office_registry_view();
    let vsto_installed = registry_key_exists_in_view(WORD_VSTO_KEY, office_registry_view)
        && registry_key_exists_in_view(POWERPOINT_VSTO_KEY, office_registry_view);

    match mode {
        OfficeIntegrationMode::Vsto => {
            if !vsto_installed {
                return Err(
                    "Cannot enable native Office mode until both Word and PowerPoint add-ins are installed"
                        .to_string(),
                );
            }
            set_vsto_load_behavior(true, office_registry_view)?;
        }
        OfficeIntegrationMode::Auto => {
            if vsto_installed {
                set_vsto_load_behavior(true, office_registry_view)?;
            }
        }
    }
    Ok(())
}

fn set_vsto_load_behavior(enabled: bool, view: RegistryView) -> Result<(), String> {
    let value = if enabled { 3 } else { 0 };
    for key in [WORD_VSTO_KEY, POWERPOINT_VSTO_KEY] {
        if registry_key_exists_in_view(key, view)
            && !registry_dword_equals_in_view(key, "LoadBehavior", value, view)
        {
            registry_set_dword_in_view(key, "LoadBehavior", value, view)?;
        }
    }
    Ok(())
}

fn addin_registry_complete(
    addin_key: &str,
    clsid_key: &str,
    expected_class: &str,
    view: RegistryView,
) -> bool {
    registry_string_value_in_view(addin_key, "FriendlyName", view)
        .is_some_and(|value| !value.trim().is_empty())
        && registry_string_value_in_view(addin_key, "Description", view)
            .is_some_and(|value| !value.trim().is_empty())
        && (registry_string_value_in_view(addin_key, "Manifest", view).is_some()
            || registry_string_value_in_view(addin_key, "CodeBase", view).is_some())
        && registry_default_string_value_in_view(clsid_key, view)
            .is_some_and(|value| value.eq_ignore_ascii_case("mscoree.dll"))
        && registry_string_value_in_view(clsid_key, "Class", view)
            .is_some_and(|value| value == expected_class)
        && registry_string_value_in_view(clsid_key, "Assembly", view).is_some()
        && registry_string_value_in_view(clsid_key, "RuntimeVersion", view).is_some()
        && registry_string_value_in_view(clsid_key, "CodeBase", view).is_some()
}

fn registered_addin_file_exists(addin_key: &str, view: RegistryView) -> bool {
    let value = registry_string_value_in_view(addin_key, "Manifest", view)
        .or_else(|| registry_string_value_in_view(addin_key, "CodeBase", view));
    let Some(value) = value else {
        return false;
    };
    let normalized = value
        .trim()
        .trim_end_matches("|vstolocal")
        .trim_end_matches('|');
    let path = normalized
        .strip_prefix("file:///")
        .or_else(|| normalized.strip_prefix("file://"))
        .unwrap_or(normalized)
        .replace('/', "\\")
        .replace("%20", " ");
    PathBuf::from(path).is_file()
}

fn native_ole_local_server_healthy(view: RegistryView) -> bool {
    let executable = registry_string_value_in_view(
        OLE_LOCAL_SERVER_KEY,
        "ServerExecutable",
        view,
    )
    .or_else(|| registry_default_string_value_in_view(OLE_LOCAL_SERVER_KEY, view));
    let Some(executable) = executable else {
        return false;
    };
    let executable = executable.trim().trim_matches('"');
    !executable.is_empty() && PathBuf::from(executable).is_file()
}

pub(crate) fn word_numbering_user_preferences() -> (bool, String) {
    let numbered = registry_dword_equals(
        WORD_USER_PREFERENCES_KEY,
        WORD_DEFAULT_NUMBERED_VALUE,
        1,
    );
    let format = registry_string_value(
        WORD_USER_PREFERENCES_KEY,
        WORD_DEFAULT_NUMBER_FORMAT_VALUE,
    )
    .map(|value| normalize_word_number_format(&value).to_string())
    .unwrap_or_else(|| "continuous".to_string());
    (numbered, format)
}

pub(crate) fn set_word_numbering_user_preferences(
    numbered: bool,
    format: &str,
) -> Result<(), String> {
    registry_set_dword(
        WORD_USER_PREFERENCES_KEY,
        WORD_DEFAULT_NUMBERED_VALUE,
        u32::from(numbered),
    )?;
    registry_set_string(
        WORD_USER_PREFERENCES_KEY,
        WORD_DEFAULT_NUMBER_FORMAT_VALUE,
        normalize_word_number_format(format),
    )
}

fn normalize_word_number_format(value: &str) -> &'static str {
    match value.trim() {
        "heading1-dot" => "heading1-dot",
        "heading1-dash" => "heading1-dash",
        "heading2-dot" => "heading2-dot",
        "heading2-dash" => "heading2-dash",
        _ => "continuous",
    }
}

fn write_mode_registry(mode: OfficeIntegrationMode) -> Result<(), String> {
    let value = match mode {
        OfficeIntegrationMode::Auto => "auto",
        OfficeIntegrationMode::Vsto => "vsto",
    };
    registry_set_string(OFFICE_MODE_KEY, "Mode", value)
}

fn registry_set_string(key: &str, name: &str, value: &str) -> Result<(), String> {
    registry_add_in_view(key, name, "REG_SZ", value, RegistryView::Native)
}

fn registry_set_dword(key: &str, name: &str, value: u32) -> Result<(), String> {
    registry_set_dword_in_view(key, name, value, RegistryView::Native)
}

fn registry_set_dword_in_view(
    key: &str,
    name: &str,
    value: u32,
    view: RegistryView,
) -> Result<(), String> {
    registry_add_in_view(key, name, "REG_DWORD", &value.to_string(), view)
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

fn registry_add_in_view(
    key: &str,
    name: &str,
    value_type: &str,
    value: &str,
    view: RegistryView,
) -> Result<(), String> {
    let mut command = hidden_command("reg.exe");
    command.args(["add", key, "/v", name, "/t", value_type, "/d", value, "/f"]);
    append_registry_view_flag(&mut command, view);
    let output = command
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

fn append_registry_view_flag(command: &mut Command, view: RegistryView) {
    if let Some(flag) = view.reg_flag() {
        command.arg(flag);
    }
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
    let Some(thumbprint) = registry_string_value(OFFICE_MODE_KEY, "CertificateThumbprint") else {
        return false;
    };
    let normalized_thumbprint = thumbprint.replace(' ', "").to_ascii_uppercase();
    if normalized_thumbprint.is_empty()
        || !normalized_thumbprint
            .chars()
            .all(|character| character.is_ascii_hexdigit())
    {
        return false;
    }
    let certificate_key = format!(
        r"HKCU\Software\Microsoft\SystemCertificates\Root\Certificates\{}",
        normalized_thumbprint
    );
    registry_key_exists_in_view(&certificate_key, RegistryView::Native)
}

fn registry_default_string_value_in_view(key: &str, view: RegistryView) -> Option<String> {
    let mut command = hidden_command("reg.exe");
    command.args(["query", key, "/ve"]);
    append_registry_view_flag(&mut command, view);
    let output = command.output().ok()?;
    parse_registry_string_output(&output.stdout, None)
}

fn registry_string_value(key: &str, value: &str) -> Option<String> {
    registry_string_value_in_view(key, value, RegistryView::Native)
}

fn registry_string_value_in_view(
    key: &str,
    value: &str,
    view: RegistryView,
) -> Option<String> {
    let mut command = hidden_command("reg.exe");
    command.args(["query", key, "/v", value]);
    append_registry_view_flag(&mut command, view);
    let output = command.output().ok()?;
    parse_registry_string_output(&output.stdout, Some(value))
}

fn parse_registry_string_output(bytes: &[u8], value: Option<&str>) -> Option<String> {
    let text = String::from_utf8_lossy(bytes);
    for line in text.lines() {
        if value.is_some_and(|value| !line.contains(value)) || !line.contains("REG_SZ") {
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

fn registry_key_exists_in_view(key: &str, view: RegistryView) -> bool {
    let mut command = hidden_command("reg.exe");
    command.args(["query", key]);
    append_registry_view_flag(&mut command, view);
    command
        .output()
        .map(|output| output.status.success())
        .unwrap_or(false)
}

fn registry_value_exists(key: &str, value: &str) -> bool {
    let mut command = hidden_command("reg.exe");
    command.args(["query", key, "/v", value]);
    command
        .output()
        .map(|output| output.status.success())
        .unwrap_or(false)
}

fn registry_dword_equals(key: &str, value: &str, expected: u32) -> bool {
    registry_dword_equals_in_view(key, value, expected, RegistryView::Native)
}

fn registry_dword_equals_in_view(
    key: &str,
    value: &str,
    expected: u32,
    view: RegistryView,
) -> bool {
    let mut command = hidden_command("reg.exe");
    command.args(["query", key, "/v", value]);
    append_registry_view_flag(&mut command, view);
    let output = match command.output() {
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
