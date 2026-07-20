use serde::Serialize;
use std::ffi::OsStr;
use tauri::{AppHandle, Manager};

pub const BACKGROUND_ARGUMENT: &str = "--office-background";

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeBackgroundStatus {
    pub installed: bool,
    pub loaded: bool,
    pub running_in_background_mode: bool,
    pub plist_path: String,
    pub executable_path: String,
    pub last_error: Option<String>,
}

pub fn is_background_mode() -> bool {
    std::env::args_os().any(|argument| argument == OsStr::new(BACKGROUND_ARGUMENT))
}

pub fn status() -> OfficeBackgroundStatus {
    let installed = startup_registered();
    OfficeBackgroundStatus {
        installed,
        loaded: installed,
        running_in_background_mode: is_background_mode(),
        plist_path: String::new(),
        executable_path: std::env::current_exe()
            .map(|path| path.display().to_string())
            .unwrap_or_default(),
        last_error: None,
    }
}

pub fn install_launch_agent() -> Result<OfficeBackgroundStatus, String> {
    #[cfg(target_os = "windows")]
    {
        let executable = std::env::current_exe()
            .map_err(|error| format!("Unable to resolve VisualTeX executable: {error}"))?;
        let command = format!("\"{}\" {BACKGROUND_ARGUMENT}", executable.display());
        let output = std::process::Command::new("reg.exe")
            .args([
                "add",
                r"HKCU\Software\Microsoft\Windows\CurrentVersion\Run",
                "/v",
                "VisualTeXOffice",
                "/t",
                "REG_SZ",
                "/d",
                &command,
                "/f",
            ])
            .output()
            .map_err(|error| format!("Unable to register VisualTeX startup: {error}"))?;
        if !output.status.success() {
            return Err(String::from_utf8_lossy(&output.stderr).trim().to_string());
        }
    }
    Ok(status())
}

pub fn pause_launch_agent_for_quit() -> Result<(), String> {
    Ok(())
}

pub fn resume_installed_launch_agent() -> Result<(), String> {
    Ok(())
}

pub fn uninstall_launch_agent() -> Result<OfficeBackgroundStatus, String> {
    #[cfg(target_os = "windows")]
    {
        let _ = std::process::Command::new("reg.exe")
            .args([
                "delete",
                r"HKCU\Software\Microsoft\Windows\CurrentVersion\Run",
                "/v",
                "VisualTeXOffice",
                "/f",
            ])
            .output();
    }
    Ok(status())
}

pub fn reveal_main_window(app: &AppHandle) -> Result<(), String> {
    let window = app
        .get_webview_window("main")
        .ok_or_else(|| "VisualTeX main window is unavailable".to_string())?;
    window
        .show()
        .map_err(|error| format!("Unable to show VisualTeX main window: {error}"))?;
    let _ = window.unminimize();
    let _ = window.set_focus();
    Ok(())
}

pub fn hide_main_window(app: &AppHandle) -> Result<(), String> {
    let window = app
        .get_webview_window("main")
        .ok_or_else(|| "VisualTeX main window is unavailable".to_string())?;
    window
        .hide()
        .map_err(|error| format!("Unable to hide VisualTeX main window: {error}"))
}

#[cfg(target_os = "windows")]
fn startup_registered() -> bool {
    std::process::Command::new("reg.exe")
        .args([
            "query",
            r"HKCU\Software\Microsoft\Windows\CurrentVersion\Run",
            "/v",
            "VisualTeXOffice",
        ])
        .output()
        .map(|output| output.status.success())
        .unwrap_or(false)
}

#[cfg(not(target_os = "windows"))]
fn startup_registered() -> bool {
    false
}
