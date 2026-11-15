#![cfg(target_os = "windows")]

//! What Settings > Office shows about WPS Office and the renderer service, so a user can see
//! why VisualTeX is missing in WPS without anyone reading the registry for them.
//!
//! WPS loads the COM add-ins Microsoft Office uses, from the Office registration in its own
//! registry view (a 32-bit WPS reads the 32-bit view). The add-ins note every host that loads
//! them (OfficeHostSelection.RecordLoad), and the installer's WPS step records its failure.
//! MathType equations in Word, PowerPoint and WPS are drawn by the VisualTeX.NativeRender
//! service; an installation stops it while it replaces files.

use crate::office::wps_mathtype_template::pe_machine;
use serde::Serialize;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};
use windows::core::{HSTRING, PCWSTR};
use windows::Win32::Foundation::ERROR_ACCESS_DENIED;
use windows::Win32::System::Registry::{
    RegGetValueW, HKEY, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, REG_ROUTINE_FLAGS, RRF_RT_ANY,
    RRF_RT_REG_DWORD, RRF_RT_REG_SZ, RRF_SUBKEY_WOW6432KEY, RRF_SUBKEY_WOW6464KEY,
};
use windows::Win32::System::Services::{
    CloseServiceHandle, OpenSCManagerW, OpenServiceW, QueryServiceStatus, StartServiceW,
    SC_HANDLE, SC_MANAGER_CONNECT, SERVICE_QUERY_STATUS, SERVICE_RUNNING, SERVICE_START,
    SERVICE_STATUS,
};

const OFFICE_MODE_KEY: &str = r"Software\VisualTeX\OfficeIntegration";
const KINGSOFT_COMMON_KEY: &str = r"Software\Kingsoft\Office\6.0\Common";
const NATIVE_RENDER_SERVICE: &str = "VisualTeX.NativeRender";
const WORD_CLSID: &str = "{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}";
const POWERPOINT_CLSID: &str = "{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}";
const FORMULA_OLE_CLSID: &str = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}";
const MACHINE_X86: u16 = 0x014c;
const MACHINE_X64: u16 = 0x8664;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Architecture {
    X86,
    X64,
}

impl Architecture {
    fn from_machine(machine: u16) -> Option<Self> {
        match machine {
            MACHINE_X86 => Some(Self::X86),
            MACHINE_X64 => Some(Self::X64),
            _ => None,
        }
    }

    fn name(self) -> &'static str {
        match self {
            Self::X86 => "x86",
            Self::X64 => "x64",
        }
    }

    fn view(self) -> REG_ROUTINE_FLAGS {
        match self {
            Self::X86 => RRF_SUBKEY_WOW6432KEY,
            Self::X64 => RRF_SUBKEY_WOW6464KEY,
        }
    }
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WpsInstallation {
    pub path: String,
    pub version: String,
    pub architecture: String,
    pub per_user: bool,
    /// The Word add-in is registered in this WPS's registry view, enabled, and built for it.
    pub writer_registered: bool,
    pub presentation_registered: bool,
    /// Listed in WPS's add-in whitelist (current user or this view's machine list).
    pub writer_whitelisted: bool,
    pub presentation_whitelisted: bool,
    pub formula_ole_registered: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HostLoad {
    pub at: String,
    pub version: String,
    pub architecture: String,
    /// False: WPS was left out at installation, so the add-in shows no ribbon there.
    pub enabled: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WpsIntegrationStatus {
    pub installations: Vec<WpsInstallation>,
    pub host_enabled: bool,
    pub writer_loaded: Option<HostLoad>,
    pub presentation_loaded: Option<HostLoad>,
    pub companion_error: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeRenderStatus {
    pub installed: bool,
    pub running: bool,
}

pub fn status() -> WpsIntegrationStatus {
    let installations = installations()
        .into_iter()
        .map(|(path, architecture)| describe(&path, architecture))
        .collect();
    WpsIntegrationStatus {
        installations,
        host_enabled: registry_dword(HKEY_CURRENT_USER, OFFICE_MODE_KEY, "Wps", REG_ROUTINE_FLAGS(0))
            .is_none_or(|value| value != 0),
        writer_loaded: host_load("WpsWriterLoaded"),
        presentation_loaded: host_load("WpsPresentationLoaded"),
        companion_error: registry_string(HKEY_CURRENT_USER, OFFICE_MODE_KEY, "WpsCompanionError", REG_ROUTINE_FLAGS(0))
            .filter(|value| !value.trim().is_empty()),
    }
}

fn describe(path: &Path, architecture: Architecture) -> WpsInstallation {
    let version = path
        .parent()
        .and_then(Path::parent)
        .and_then(Path::file_name)
        .map(|name| name.to_string_lossy().to_string())
        .unwrap_or_default();
    let per_user = std::env::var_os("LOCALAPPDATA").is_some_and(|root| {
        path.to_string_lossy()
            .to_lowercase()
            .starts_with(&root.to_string_lossy().to_lowercase())
    });
    WpsInstallation {
        path: path.to_string_lossy().to_string(),
        version,
        architecture: architecture.name().to_string(),
        per_user,
        writer_registered: addin_registered(architecture, "Word", "VisualTeX.WordVsto", WORD_CLSID),
        presentation_registered: addin_registered(
            architecture,
            "PowerPoint",
            "VisualTeX.PowerPointVsto",
            POWERPOINT_CLSID,
        ),
        writer_whitelisted: whitelisted(architecture, "WPS", "VisualTeX.WordVsto"),
        presentation_whitelisted: whitelisted(architecture, "WPP", "VisualTeX.PowerPointVsto"),
        formula_ole_registered: registry_string(
            HKEY_LOCAL_MACHINE,
            &format!(r"Software\Classes\CLSID\{FORMULA_OLE_CLSID}\LocalServer32"),
            "ServerExecutable",
            architecture.view(),
        )
        .is_some_and(|server| Path::new(server.trim().trim_matches('"')).is_file()),
    }
}

/// wps.exe of every WPS Office installation: the install roots WPS records (per user and per
/// machine, both registry views) and the usual folders.
fn installations() -> Vec<(PathBuf, Architecture)> {
    let mut candidates = Vec::new();
    for root in [HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE] {
        for view in [RRF_SUBKEY_WOW6464KEY, RRF_SUBKEY_WOW6432KEY] {
            if let Some(install) = registry_string(root, KINGSOFT_COMMON_KEY, "InstallRoot", view) {
                candidates.push(PathBuf::from(install.trim().trim_matches('"')).join("office6").join("wps.exe"));
            }
        }
    }
    for variable in ["LOCALAPPDATA", "ProgramW6432", "ProgramFiles(x86)"] {
        let Some(root) = std::env::var_os(variable) else {
            continue;
        };
        let parent = PathBuf::from(root).join("Kingsoft").join("WPS Office");
        let Ok(entries) = std::fs::read_dir(&parent) else {
            continue;
        };
        for entry in entries.filter_map(Result::ok) {
            candidates.push(entry.path().join("office6").join("wps.exe"));
        }
    }
    let mut seen = std::collections::HashSet::new();
    let mut found = Vec::new();
    for candidate in candidates {
        if !candidate.is_file() {
            continue;
        }
        let path = std::fs::canonicalize(&candidate)
            .map(|path| strip_verbatim(&path))
            .unwrap_or(candidate);
        if !seen.insert(path.to_string_lossy().to_lowercase()) {
            continue;
        }
        if let Some(architecture) = pe_machine(&path).and_then(Architecture::from_machine) {
            found.push((path, architecture));
        }
    }
    found
}

fn strip_verbatim(path: &Path) -> PathBuf {
    let text = path.to_string_lossy();
    PathBuf::from(text.strip_prefix(r"\\?\").unwrap_or(&text))
}

fn addin_registered(architecture: Architecture, host: &str, prog_id: &str, clsid: &str) -> bool {
    let view = architecture.view();
    let addin = format!(r"Software\Microsoft\Office\{host}\Addins\{prog_id}");
    let enabled = [HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER]
        .into_iter()
        .any(|root| registry_dword(root, &addin, "LoadBehavior", view) == Some(3));
    let server = format!(r"Software\Classes\CLSID\{clsid}\InprocServer32");
    let managed = registry_string(HKEY_LOCAL_MACHINE, &server, "", view)
        .is_some_and(|value| value.trim().eq_ignore_ascii_case("mscoree.dll"));
    let built_for_host = registry_string(HKEY_LOCAL_MACHINE, &server, "CodeBase", view)
        .and_then(|uri| path_from_file_uri(&uri))
        .filter(|path| path.is_file())
        .and_then(|path| pe_machine(&path))
        .and_then(Architecture::from_machine)
        == Some(architecture);
    enabled && managed && built_for_host
}

fn whitelisted(architecture: Architecture, list: &str, prog_id: &str) -> bool {
    let key = format!(r"Software\Kingsoft\Office\{list}\AddinsWL");
    registry_value_exists(HKEY_CURRENT_USER, &key, prog_id, REG_ROUTINE_FLAGS(0))
        || registry_value_exists(HKEY_LOCAL_MACHINE, &key, prog_id, architecture.view())
}

fn host_load(value: &str) -> Option<HostLoad> {
    let text = registry_string(HKEY_CURRENT_USER, OFFICE_MODE_KEY, value, REG_ROUTINE_FLAGS(0))?;
    let mut parts = text.split('|');
    Some(HostLoad {
        at: parts.next()?.trim().to_string(),
        version: parts.next().unwrap_or_default().trim().to_string(),
        architecture: parts.next().unwrap_or_default().trim().to_string(),
        enabled: parts.next().is_none_or(|state| state.trim() != "off"),
    })
}

/// A file:/// URI as .NET writes it (Uri.AbsoluteUri) back to a path.
fn path_from_file_uri(uri: &str) -> Option<PathBuf> {
    let rest = uri.trim().strip_prefix("file:///").or_else(|| uri.trim().strip_prefix("file://"))?;
    let bytes = rest.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' && index + 2 < bytes.len() {
            if let Ok(value) = u8::from_str_radix(&rest[index + 1..index + 3], 16) {
                decoded.push(value);
                index += 3;
                continue;
            }
        }
        decoded.push(bytes[index]);
        index += 1;
    }
    Some(PathBuf::from(String::from_utf8_lossy(&decoded).replace('/', "\\")))
}

fn registry_string(root: HKEY, key: &str, value: &str, view: REG_ROUTINE_FLAGS) -> Option<String> {
    let key = HSTRING::from(key);
    let value = HSTRING::from(value);
    let mut size = 0u32;
    let status = unsafe { RegGetValueW(root, &key, &value, RRF_RT_REG_SZ | view, None, None, Some(&mut size)) };
    if status.is_err() || size == 0 {
        return None;
    }
    let mut buffer = vec![0u16; size as usize / 2 + 1];
    let mut size = (buffer.len() * 2) as u32;
    let status = unsafe {
        RegGetValueW(root, &key, &value, RRF_RT_REG_SZ | view, None, Some(buffer.as_mut_ptr().cast()), Some(&mut size))
    };
    if status.is_err() {
        return None;
    }
    let length = buffer.iter().position(|unit| *unit == 0).unwrap_or(buffer.len());
    Some(String::from_utf16_lossy(&buffer[..length]))
}

fn registry_dword(root: HKEY, key: &str, value: &str, view: REG_ROUTINE_FLAGS) -> Option<u32> {
    let key = HSTRING::from(key);
    let value = HSTRING::from(value);
    let mut data = 0u32;
    let mut size = std::mem::size_of::<u32>() as u32;
    let status = unsafe {
        RegGetValueW(root, &key, &value, RRF_RT_REG_DWORD | view, None, Some((&mut data as *mut u32).cast()), Some(&mut size))
    };
    status.is_ok().then_some(data)
}

fn registry_value_exists(root: HKEY, key: &str, value: &str, view: REG_ROUTINE_FLAGS) -> bool {
    let key = HSTRING::from(key);
    let value = HSTRING::from(value);
    unsafe { RegGetValueW(root, &key, &value, RRF_RT_ANY | view, None, None, None) }.is_ok()
}

struct ServiceHandle(SC_HANDLE);

impl Drop for ServiceHandle {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseServiceHandle(self.0);
        }
    }
}

fn open_render_service(access: u32) -> windows::core::Result<Option<ServiceHandle>> {
    let manager = ServiceHandle(unsafe { OpenSCManagerW(PCWSTR::null(), PCWSTR::null(), SC_MANAGER_CONNECT) }?);
    let name = HSTRING::from(NATIVE_RENDER_SERVICE);
    match unsafe { OpenServiceW(manager.0, &name, access) } {
        Ok(service) => Ok(Some(ServiceHandle(service))),
        // 1060: ERROR_SERVICE_DOES_NOT_EXIST
        Err(error) if error.code() == windows::core::HRESULT::from_win32(1060) => Ok(None),
        Err(error) => Err(error),
    }
}

pub fn native_render_status() -> NativeRenderStatus {
    let Ok(Some(service)) = open_render_service(SERVICE_QUERY_STATUS) else {
        return NativeRenderStatus { installed: false, running: false };
    };
    let mut state = SERVICE_STATUS::default();
    let running = unsafe { QueryServiceStatus(service.0, &mut state) }.is_ok()
        && state.dwCurrentState == SERVICE_RUNNING;
    NativeRenderStatus { installed: true, running }
}

/// Starts the renderer service: directly where the user may (current installations), else
/// through one administrator prompt.
pub fn start_native_render() -> Result<NativeRenderStatus, String> {
    if native_render_status().running {
        return Ok(native_render_status());
    }
    let direct = match open_render_service(SERVICE_START | SERVICE_QUERY_STATUS) {
        Ok(Some(service)) => unsafe { StartServiceW(service.0, None) },
        Ok(None) => return Err("The VisualTeX renderer service is not installed. Repair the Office integration.".to_string()),
        Err(error) => Err(error),
    };
    match direct {
        Ok(()) => {}
        // 1056: ERROR_SERVICE_ALREADY_RUNNING
        Err(error) if error.code() == windows::core::HRESULT::from_win32(1056) => {}
        Err(error) if error.code() == ERROR_ACCESS_DENIED.to_hresult() => start_elevated()?,
        Err(error) => return Err(format!("Unable to start the VisualTeX renderer service: {error}")),
    }
    let deadline = Instant::now() + Duration::from_secs(20);
    while Instant::now() < deadline {
        let status = native_render_status();
        if status.running {
            return Ok(status);
        }
        std::thread::sleep(Duration::from_millis(250));
    }
    Err("The VisualTeX renderer service did not start within 20 seconds.".to_string())
}

fn start_elevated() -> Result<(), String> {
    use std::os::windows::process::CommandExt;
    let powershell = crate::ocr::install::windows_powershell_executable()?;
    let command = format!(
        "$p = Start-Process -FilePath (Join-Path $env:WINDIR 'System32\\sc.exe') -ArgumentList 'start','{NATIVE_RENDER_SERVICE}' -Verb RunAs -WindowStyle Hidden -Wait -PassThru; exit $p.ExitCode"
    );
    let output = std::process::Command::new(powershell)
        .creation_flags(0x0800_0000)
        .args(["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", &command])
        .output()
        .map_err(|error| format!("Unable to start the VisualTeX renderer service: {error}"))?;
    // 1056: already running
    match output.status.code() {
        Some(0) | Some(1056) => Ok(()),
        _ => Err(format!(
            "The VisualTeX renderer service was not started{}",
            crate::ocr::install::decode_process_output(&output.stderr)
                .lines()
                .find(|line| !line.trim().is_empty())
                .map(|line| format!(": {}", line.trim()))
                .unwrap_or_else(|| ".".to_string())
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn file_uris_become_paths() {
        assert_eq!(
            path_from_file_uri("file:///C:/Program%20Files%20(x86)/VisualTeX/WindowsOffice/VSTO-WPS/VisualTeX.WordVsto.dll"),
            Some(PathBuf::from(r"C:\Program Files (x86)\VisualTeX\WindowsOffice\VSTO-WPS\VisualTeX.WordVsto.dll"))
        );
        assert_eq!(
            path_from_file_uri("file:///C:/Users/%E5%BB%96/a.dll"),
            Some(PathBuf::from(r"C:\Users\廖\a.dll"))
        );
        assert_eq!(path_from_file_uri("C:/a.dll"), None);
    }

    #[test]
    fn reads_this_machines_wps_installations() {
        // Smoke test: never panics, and every installation found has a known bitness.
        for installation in status().installations {
            assert!(matches!(installation.architecture.as_str(), "x86" | "x64"));
        }
    }
}
