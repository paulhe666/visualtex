#![cfg(target_os = "windows")]

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::Serialize;
use std::ffi::c_void;
use std::os::windows::process::CommandExt;
use std::process::{Command, Stdio};
use std::thread;
use std::time::Duration;
use tauri::{AppHandle, Manager};

const CREATE_NO_WINDOW: u32 = 0x0800_0000;
const MIN_TIMEOUT_MS: u64 = 1_000;
const MAX_TIMEOUT_MS: u64 = 120_000;
const CF_UNICODETEXT: u32 = 13;
const GMEM_MOVEABLE: u32 = 0x0002;
const CLIPBOARD_OPEN_RETRIES: usize = 25;
const CLIPBOARD_OPEN_RETRY_DELAY_MS: u64 = 20;
const SW_MINIMIZE: i32 = 6;
const SW_RESTORE: i32 = 9;
const FOREGROUND_RESTORE_RETRIES: usize = 8;
const FOREGROUND_RESTORE_RETRY_DELAY_MS: u64 = 80;

#[link(name = "user32")]
unsafe extern "system" {
    fn OpenClipboard(hwnd: *mut c_void) -> i32;
    fn CloseClipboard() -> i32;
    fn EmptyClipboard() -> i32;
    fn SetClipboardData(format: u32, memory: *mut c_void) -> *mut c_void;
    fn FindWindowW(class_name: *const u16, window_name: *const u16) -> *mut c_void;
    fn ShowWindowAsync(hwnd: *mut c_void, command: i32) -> i32;
    fn SetForegroundWindow(hwnd: *mut c_void) -> i32;
    fn BringWindowToTop(hwnd: *mut c_void) -> i32;
    fn GetForegroundWindow() -> *mut c_void;
    fn GetWindowThreadProcessId(hwnd: *mut c_void, process_id: *mut u32) -> u32;
    fn AttachThreadInput(thread_id_attach: u32, thread_id_attach_to: u32, attach: i32) -> i32;
    fn SetFocus(hwnd: *mut c_void) -> *mut c_void;
}

#[link(name = "kernel32")]
unsafe extern "system" {
    fn GlobalAlloc(flags: u32, bytes: usize) -> *mut c_void;
    fn GlobalLock(memory: *mut c_void) -> *mut c_void;
    fn GlobalUnlock(memory: *mut c_void) -> i32;
    fn GlobalFree(memory: *mut c_void) -> *mut c_void;
    fn GetCurrentThreadId() -> u32;
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct WindowsQuickOcrCapture {
    data_base64: String,
    extension: String,
}

fn encode_powershell(script: &str) -> String {
    let bytes = script
        .encode_utf16()
        .flat_map(|unit| unit.to_le_bytes())
        .collect::<Vec<_>>();
    BASE64_STANDARD.encode(bytes)
}

pub(crate) fn normalize_capture_mode(value: &str) -> Result<&'static str, String> {
    match value.trim().to_ascii_lowercase().as_str() {
        "windows" | "immediate" => Ok("windows"),
        "pixpin" => Ok("pixpin"),
        "clipboard" | "system-screenshot" => Ok("clipboard"),
        _ => Err("Unsupported Windows OCR screenshot provider".to_string()),
    }
}

fn capture_script(capture_mode: &str, timeout_ms: u64) -> String {
    let launch = match capture_mode {
        "windows" => r#"
Start-Process -FilePath 'explorer.exe' -ArgumentList 'ms-screenclip:' -WindowStyle Hidden | Out-Null
"#,
        "pixpin" => r#"
$pixpinCandidates = New-Object System.Collections.Generic.List[string]
$runningPixPin = Get-Process -Name 'PixPin' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($null -ne $runningPixPin) {
    try {
        if (-not [string]::IsNullOrWhiteSpace($runningPixPin.Path)) {
            [void]$pixpinCandidates.Add($runningPixPin.Path)
        }
    } catch {}
}
try {
    $pixpinCommand = Get-Command 'PixPin.exe' -ErrorAction Stop
    if ($null -ne $pixpinCommand -and -not [string]::IsNullOrWhiteSpace($pixpinCommand.Source)) {
        [void]$pixpinCandidates.Add($pixpinCommand.Source)
    }
} catch {}
foreach ($candidateSpec in @(
    @{ Root = $env:LOCALAPPDATA; Relative = 'PixPin\PixPin.exe' },
    @{ Root = $env:LOCALAPPDATA; Relative = 'Programs\PixPin\PixPin.exe' },
    @{ Root = $env:ProgramFiles; Relative = 'PixPin\PixPin.exe' },
    @{ Root = ${env:ProgramFiles(x86)}; Relative = 'PixPin\PixPin.exe' }
)) {
    $root = [string]$candidateSpec.Root
    if ([string]::IsNullOrWhiteSpace($root)) {
        continue
    }
    $candidate = Join-Path $root ([string]$candidateSpec.Relative)
    [void]$pixpinCandidates.Add($candidate)
}
$pixpin = $pixpinCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($pixpin)) {
    throw 'PixPin.exe was not found. Install PixPin or start a portable PixPin instance, then retry.'
}
# PixPin's official documentation requires an already-running instance for the
# Windows -r scripting command. Cold-start it when needed, wait until the process
# exists, and only then send the interactive screenshot-and-copy script.
if ($null -eq $runningPixPin) {
    Start-Process -FilePath $pixpin | Out-Null
    $startupDeadline = [DateTime]::UtcNow.AddSeconds(8)
    do {
        Start-Sleep -Milliseconds 120
        $runningPixPin = Get-Process -Name 'PixPin' -ErrorAction SilentlyContinue | Select-Object -First 1
    } while ($null -eq $runningPixPin -and [DateTime]::UtcNow -lt $startupDeadline)
    if ($null -eq $runningPixPin) {
        throw 'PixPin was started but no running PixPin instance became available for the screenshot script.'
    }
}
Start-Process -FilePath $pixpin -ArgumentList @('-r', 'pixpin.screenShot(ShotAction.Copy)') -WindowStyle Hidden | Out-Null
"#,
        _ => "",
    };
    format!(
        r#"
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public static class VisualTeXClipboardNative {{
    [DllImport("user32.dll")]
    public static extern uint GetClipboardSequenceNumber();
}}
'@

$before = [VisualTeXClipboardNative]::GetClipboardSequenceNumber()
{launch}
$deadline = [DateTime]::UtcNow.AddMilliseconds({timeout_ms})
while ([DateTime]::UtcNow -lt $deadline) {{
    Start-Sleep -Milliseconds 120
    if ([VisualTeXClipboardNative]::GetClipboardSequenceNumber() -eq $before) {{
        continue
    }}
    try {{
        if (-not [System.Windows.Forms.Clipboard]::ContainsImage()) {{
            continue
        }}
        $image = [System.Windows.Forms.Clipboard]::GetImage()
        if ($null -eq $image) {{
            continue
        }}
        try {{
            $stream = New-Object System.IO.MemoryStream
            try {{
                $image.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
                [Console]::Out.Write([Convert]::ToBase64String($stream.ToArray()))
                exit 0
            }} finally {{
                $stream.Dispose()
            }}
        }} finally {{
            $image.Dispose()
        }}
    }} catch {{
        # The clipboard can be momentarily locked while Snipping Tool commits.
        # Keep polling until the timeout instead of surfacing a transient error.
    }}
}}
exit 2
"#,
    )
}

fn capture_windows_clipboard_image(
    capture_mode: &str,
    timeout_ms: u64,
) -> Result<Option<WindowsQuickOcrCapture>, String> {
    let capture_mode = normalize_capture_mode(capture_mode)?;
    let timeout_ms = timeout_ms.clamp(MIN_TIMEOUT_MS, MAX_TIMEOUT_MS);
    let script = capture_script(capture_mode, timeout_ms);
    let encoded = encode_powershell(&script);
    let output = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-STA",
            "-NonInteractive",
            "-EncodedCommand",
            encoded.as_str(),
        ])
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .creation_flags(CREATE_NO_WINDOW)
        .output()
        .map_err(|error| format!("Unable to start the Windows screenshot bridge: {error}"))?;

    if output.status.success() {
        let data_base64 = String::from_utf8(output.stdout)
            .map_err(|_| "Windows screenshot bridge returned invalid UTF-8 output".to_string())?
            .trim()
            .to_string();
        if data_base64.is_empty() {
            return Err("Windows screenshot bridge returned an empty image".to_string());
        }
        return Ok(Some(WindowsQuickOcrCapture {
            data_base64,
            extension: "png".to_string(),
        }));
    }

    if output.status.code() == Some(2) {
        return Ok(None);
    }

    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    let detail = if !stderr.is_empty() { stderr } else { stdout };
    Err(if detail.is_empty() {
        format!(
            "Windows screenshot bridge exited with status {:?}",
            output.status.code()
        )
    } else {
        format!("Windows screenshot bridge failed: {detail}")
    })
}

#[tauri::command]
pub(crate) fn minimize_visualtex_main_window(app: AppHandle) -> Result<(), String> {
    if let Some(window) = app.get_webview_window("main") {
        if window.minimize().is_ok() {
            thread::sleep(Duration::from_millis(45));
            if window.is_minimized().unwrap_or(false) {
                return Ok(());
            }
        }
    }

    let title = "VisualTeX"
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let hwnd = unsafe { FindWindowW(std::ptr::null(), title.as_ptr()) };
    if hwnd.is_null() {
        return Err("Unable to locate the VisualTeX main window to minimize".to_string());
    }
    if unsafe { ShowWindowAsync(hwnd, SW_MINIMIZE) } == 0 {
        return Err(format!(
            "Windows refused to minimize the VisualTeX main window: {}",
            std::io::Error::last_os_error()
        ));
    }
    Ok(())
}

fn restore_visualtex_foreground_window() -> Result<(), String> {
    let title = "VisualTeX"
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let hwnd = unsafe { FindWindowW(std::ptr::null(), title.as_ptr()) };
    if hwnd.is_null() {
        return Err("Unable to locate the VisualTeX main window after screenshot capture".to_string());
    }

    for _ in 0..FOREGROUND_RESTORE_RETRIES {
        unsafe {
            let _ = ShowWindowAsync(hwnd, SW_RESTORE);
        }
        let foreground = unsafe { GetForegroundWindow() };
        let current_thread = unsafe { GetCurrentThreadId() };
        let foreground_thread = if foreground.is_null() {
            0
        } else {
            unsafe { GetWindowThreadProcessId(foreground, std::ptr::null_mut()) }
        };
        let target_thread = unsafe { GetWindowThreadProcessId(hwnd, std::ptr::null_mut()) };

        let attached_foreground = foreground_thread != 0 && foreground_thread != current_thread;
        let attached_target = target_thread != 0
            && target_thread != current_thread
            && target_thread != foreground_thread;
        if attached_foreground {
            unsafe {
                let _ = AttachThreadInput(current_thread, foreground_thread, 1);
            }
        }
        if attached_target {
            unsafe {
                let _ = AttachThreadInput(current_thread, target_thread, 1);
            }
        }
        unsafe {
            let _ = BringWindowToTop(hwnd);
            let _ = SetForegroundWindow(hwnd);
            let _ = SetFocus(hwnd);
        }
        if attached_target {
            unsafe {
                let _ = AttachThreadInput(current_thread, target_thread, 0);
            }
        }
        if attached_foreground {
            unsafe {
                let _ = AttachThreadInput(current_thread, foreground_thread, 0);
            }
        }

        if unsafe { GetForegroundWindow() } == hwnd {
            return Ok(());
        }
        thread::sleep(Duration::from_millis(FOREGROUND_RESTORE_RETRY_DELAY_MS));
    }

    Err(format!(
        "Windows refused to activate the VisualTeX main window: {}",
        std::io::Error::last_os_error()
    ))
}

pub(crate) fn capture_windows_quick_ocr_bytes(
    capture_mode: &str,
    timeout_ms: u64,
) -> Result<Option<Vec<u8>>, String> {
    let Some(capture) = capture_windows_clipboard_image(capture_mode, timeout_ms)? else {
        return Ok(None);
    };
    BASE64_STANDARD
        .decode(capture.data_base64.as_bytes())
        .map(Some)
        .map_err(|error| format!("Windows screenshot bridge returned invalid image data: {error}"))
}

pub(crate) fn write_clipboard_text(text: &str) -> Result<(), String> {
    let mut opened = false;
    for _ in 0..CLIPBOARD_OPEN_RETRIES {
        if unsafe { OpenClipboard(std::ptr::null_mut()) } != 0 {
            opened = true;
            break;
        }
        thread::sleep(Duration::from_millis(CLIPBOARD_OPEN_RETRY_DELAY_MS));
    }
    if !opened {
        return Err(format!(
            "Unable to open the Windows clipboard: {}",
            std::io::Error::last_os_error()
        ));
    }

    let result = (|| {
        if unsafe { EmptyClipboard() } == 0 {
            return Err(format!(
                "Unable to clear the Windows clipboard: {}",
                std::io::Error::last_os_error()
            ));
        }

        let utf16 = text.encode_utf16().chain(std::iter::once(0)).collect::<Vec<_>>();
        let byte_len = utf16.len() * std::mem::size_of::<u16>();
        let memory = unsafe { GlobalAlloc(GMEM_MOVEABLE, byte_len) };
        if memory.is_null() {
            return Err(format!(
                "Unable to allocate Windows clipboard memory: {}",
                std::io::Error::last_os_error()
            ));
        }

        let locked = unsafe { GlobalLock(memory) };
        if locked.is_null() {
            unsafe {
                let _ = GlobalFree(memory);
            }
            return Err(format!(
                "Unable to lock Windows clipboard memory: {}",
                std::io::Error::last_os_error()
            ));
        }
        unsafe {
            std::ptr::copy_nonoverlapping(
                utf16.as_ptr(),
                locked.cast::<u16>(),
                utf16.len(),
            );
            let _ = GlobalUnlock(memory);
        }

        if unsafe { SetClipboardData(CF_UNICODETEXT, memory) }.is_null() {
            unsafe {
                let _ = GlobalFree(memory);
            }
            return Err(format!(
                "Unable to write text to the Windows clipboard: {}",
                std::io::Error::last_os_error()
            ));
        }

        // Ownership of memory transfers to the system after SetClipboardData.
        Ok(())
    })();

    unsafe {
        let _ = CloseClipboard();
    }
    result
}

#[tauri::command]
pub(crate) async fn capture_windows_quick_ocr(
    app: AppHandle,
    capture_mode: String,
    timeout_ms: u64,
) -> Result<Option<WindowsQuickOcrCapture>, String> {
    if let Some(window) = app.get_webview_window("main") {
        window
            .minimize()
            .map_err(|error| format!("Unable to minimize VisualTeX for quick OCR: {error}"))?;
        // Let WebView2/DWM fully leave the capture surface before opening the
        // system selector. This mirrors the macOS native quick-OCR flow and is
        // intentionally done in Rust so a frontend focus race cannot skip it.
        tokio::time::sleep(Duration::from_millis(180)).await;
    }
    let capture_result = tauri::async_runtime::spawn_blocking(move || {
        capture_windows_clipboard_image(&capture_mode, timeout_ms)
    })
    .await
    .map_err(|error| format!("Windows screenshot bridge task failed: {error}"))
    .and_then(|result| result);

    // Snipping Tool can still own the foreground for a short period after it
    // commits the clipboard image. Wait for that overlay to leave, restore the
    // Tauri window, then use the native Win32 foreground path with retries.
    tokio::time::sleep(Duration::from_millis(140)).await;
    let restore_result = crate::app_lifecycle::ensure_main_window(&app)
        .map(|_| ())
        .and_then(|_| restore_visualtex_foreground_window());
    match (capture_result, restore_result) {
        (Ok(capture), Ok(())) => Ok(capture),
        (Err(capture_error), Ok(())) => Err(capture_error),
        (Ok(_), Err(restore_error)) => Err(format!(
            "Unable to restore VisualTeX after quick OCR: {restore_error}"
        )),
        (Err(capture_error), Err(restore_error)) => Err(format!(
            "{capture_error}; additionally unable to restore VisualTeX: {restore_error}"
        )),
    }
}

#[tauri::command]
pub(crate) async fn write_windows_ocr_clipboard_text(text: String) -> Result<(), String> {
    tauri::async_runtime::spawn_blocking(move || write_clipboard_text(&text))
        .await
        .map_err(|error| format!("Windows clipboard task failed: {error}"))?
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn powershell_script_uses_native_screenclip_and_clipboard_sequence() {
        let script = capture_script("windows", 60_000);
        assert!(script.contains("ms-screenclip:"));
        assert!(script.contains("GetClipboardSequenceNumber"));
        assert!(script.contains("Clipboard]::ContainsImage"));
        assert!(script.contains("ImageFormat]::Png"));
    }

    #[test]
    fn pixpin_script_cold_starts_then_uses_official_copy_action() {
        let script = capture_script("pixpin", 60_000);
        assert!(script.contains("PixPin.exe"));
        assert!(script.contains("Get-Process -Name 'PixPin'"));
        assert!(script.contains("Start-Process -FilePath $pixpin | Out-Null"));
        assert!(script.contains("pixpin.screenShot(ShotAction.Copy)"));
        assert!(!script.contains("ms-screenclip:"));
    }

    #[test]
    fn wait_only_script_does_not_launch_a_capture_provider() {
        let script = capture_script("clipboard", 60_000);
        assert!(!script.contains("ms-screenclip:"));
        assert!(!script.contains("ShotAction.Copy"));
    }

    #[test]
    fn capture_mode_normalizes_legacy_names() {
        assert_eq!(normalize_capture_mode("immediate").unwrap(), "windows");
        assert_eq!(normalize_capture_mode("system-screenshot").unwrap(), "clipboard");
        assert!(normalize_capture_mode("unknown").is_err());
    }
}
