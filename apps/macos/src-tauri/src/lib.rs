use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::PathBuf;
use std::process::Command;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::{AppHandle, Emitter, Manager};
use tauri_plugin_dialog::{DialogExt, MessageDialogKind};

mod document_export;
#[path = "../../../shared/visualtex-backend/src/document_library.rs"]
mod document_library;
#[path = "../../../shared/visualtex-backend/src/ocr_offline.rs"]
mod ocr_offline;
mod ocr;
mod anywhere;
mod office;
mod quick_ocr;
mod svg_font_stabilizer;
mod system_math_glyphs;

const ACTIVE_THEME_FILE: &str = "active-theme.txt";
const THEME_CHANGED_EVENT: &str = "visualtex-theme-changed";
const MAIN_WINDOW_SIZE_FILE: &str = "main-window-size.json";
const MAIN_WINDOW_MODE_SIZES_FILE: &str = "main-window-mode-sizes-v1.json";
const DEFAULT_NORMAL_WINDOW_WIDTH: f64 = 1240.0;
const DEFAULT_NORMAL_WINDOW_HEIGHT: f64 = 820.0;
const DEFAULT_KEYPAD_WINDOW_WIDTH: f64 = 642.0;
const DEFAULT_KEYPAD_WINDOW_HEIGHT: f64 = 345.0;
const MIN_MAIN_WINDOW_WIDTH: f64 = 500.0;
const MIN_MAIN_WINDOW_HEIGHT: f64 = 300.0;
const MAX_MAIN_WINDOW_WIDTH: f64 = 4000.0;
const MAX_MAIN_WINDOW_HEIGHT: f64 = 3000.0;
static MAIN_WINDOW_SIZE_WRITE_GENERATION: AtomicU64 = AtomicU64::new(0);
static MAIN_WINDOW_KEYPAD_MODE: AtomicBool = AtomicBool::new(false);

#[cfg(target_os = "macos")]
const FAST_OPEN_MAIN_THREAD_WAIT: std::time::Duration = std::time::Duration::from_millis(300);

pub(crate) fn spawn_background_task(
    name: &'static str,
    task: impl FnOnce() + Send + 'static,
) -> bool {
    match std::thread::Builder::new()
        .name(name.to_string())
        .spawn(move || {
            if let Err(payload) =
                std::panic::catch_unwind(std::panic::AssertUnwindSafe(task))
            {
                eprintln!(
                    "VisualTeX background task {name} recovered from an internal error: {}",
                    panic_payload_detail(payload.as_ref())
                );
            }
        })
    {
        Ok(_) => true,
        Err(error) => {
            eprintln!("Unable to start VisualTeX background task {name}: {error}");
            false
        }
    }
}

pub(crate) fn panic_payload_detail(payload: &(dyn std::any::Any + Send)) -> String {
    if let Some(message) = payload.downcast_ref::<&str>() {
        (*message).to_string()
    } else if let Some(message) = payload.downcast_ref::<String>() {
        message.clone()
    } else {
        "unknown Rust panic".to_string()
    }
}

fn normalize_app_theme(theme: &str) -> &'static str {
    match theme.trim() {
        "beige" => "beige",
        "dark" => "dark",
        "purple" => "purple",
        "green" => "green",
        "codex" => "codex",
        "notion" => "notion",
        "one" => "one",
        "proof" => "proof",
        "raycast" => "raycast",
        "rose-pine" => "rose-pine",
        "solarized" => "solarized",
        "vercel" => "vercel",
        "vscode-plus" => "vscode-plus",
        "xcode" => "xcode",
        "custom" => "custom",
        _ => "light",
    }
}

pub(crate) fn persisted_app_theme(app: &AppHandle) -> String {
    let Ok(app_data_dir) = app.path().app_data_dir() else {
        return "light".to_string();
    };
    fs::read_to_string(app_data_dir.join(ACTIVE_THEME_FILE))
        .map(|theme| normalize_app_theme(&theme).to_string())
        .unwrap_or_else(|_| "light".to_string())
}

#[tauri::command]
fn set_app_theme(app: AppHandle, theme: String) -> Result<String, String> {
    let normalized = normalize_app_theme(&theme).to_string();
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|error| error.to_string())?;
    fs::create_dir_all(&app_data_dir).map_err(|error| error.to_string())?;
    fs::write(app_data_dir.join(ACTIVE_THEME_FILE), &normalized)
        .map_err(|error| error.to_string())?;
    app.emit(THEME_CHANGED_EVENT, normalized.clone())
        .map_err(|error| error.to_string())?;
    Ok(normalized)
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ConfigurationWindowSize {
    width: f64,
    height: f64,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct AppWindowConfiguration {
    main: Option<ConfigurationWindowSize>,
    office_editor: Option<ConfigurationWindowSize>,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct MainWindowModeSizes {
    normal: ConfigurationWindowSize,
    keypad: ConfigurationWindowSize,
}

fn normalize_main_window_size(width: f64, height: f64) -> ConfigurationWindowSize {
    let width = if width.is_finite() {
        width
    } else {
        DEFAULT_NORMAL_WINDOW_WIDTH
    };
    let height = if height.is_finite() {
        height
    } else {
        DEFAULT_NORMAL_WINDOW_HEIGHT
    };
    ConfigurationWindowSize {
        width: width.clamp(MIN_MAIN_WINDOW_WIDTH, MAX_MAIN_WINDOW_WIDTH),
        height: height.clamp(MIN_MAIN_WINDOW_HEIGHT, MAX_MAIN_WINDOW_HEIGHT),
    }
}

fn main_window_size_path(app: &AppHandle) -> Result<PathBuf, String> {
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data_dir.join(MAIN_WINDOW_SIZE_FILE))
}

fn main_window_mode_sizes_path(app: &AppHandle) -> Result<PathBuf, String> {
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data_dir.join(MAIN_WINDOW_MODE_SIZES_FILE))
}

fn default_normal_window_size() -> ConfigurationWindowSize {
    normalize_main_window_size(DEFAULT_NORMAL_WINDOW_WIDTH, DEFAULT_NORMAL_WINDOW_HEIGHT)
}

fn default_keypad_window_size() -> ConfigurationWindowSize {
    normalize_main_window_size(DEFAULT_KEYPAD_WINDOW_WIDTH, DEFAULT_KEYPAD_WINDOW_HEIGHT)
}

fn read_main_window_size(app: &AppHandle) -> Option<ConfigurationWindowSize> {
    let path = main_window_size_path(app).ok()?;
    let bytes = fs::read(path).ok()?;
    let size: ConfigurationWindowSize = serde_json::from_slice(&bytes).ok()?;
    Some(normalize_main_window_size(size.width, size.height))
}

fn current_main_window_size(app: &AppHandle) -> Option<ConfigurationWindowSize> {
    let window = app.get_webview_window("main")?;
    let physical = window.inner_size().ok()?;
    let scale_factor = window.scale_factor().ok()?.max(0.1);
    Some(normalize_main_window_size(
        physical.width as f64 / scale_factor,
        physical.height as f64 / scale_factor,
    ))
}

fn read_main_window_mode_sizes(app: &AppHandle) -> MainWindowModeSizes {
    if let Ok(path) = main_window_mode_sizes_path(app) {
        if let Ok(bytes) = fs::read(path) {
            if let Ok(stored) = serde_json::from_slice::<MainWindowModeSizes>(&bytes) {
                return MainWindowModeSizes {
                    normal: normalize_main_window_size(stored.normal.width, stored.normal.height),
                    keypad: normalize_main_window_size(stored.keypad.width, stored.keypad.height),
                };
            }
        }
    }

    MainWindowModeSizes {
        normal: read_main_window_size(app).unwrap_or_else(default_normal_window_size),
        keypad: default_keypad_window_size(),
    }
}

fn write_main_window_mode_sizes(
    app: &AppHandle,
    sizes: MainWindowModeSizes,
) -> Result<(), String> {
    let path = main_window_mode_sizes_path(app)?;
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    let bytes = serde_json::to_vec_pretty(&sizes).map_err(|error| error.to_string())?;
    fs::write(path, bytes).map_err(|error| error.to_string())
}

fn write_main_window_profile_size(
    app: &AppHandle,
    keypad: bool,
    size: ConfigurationWindowSize,
) -> Result<(), String> {
    let normalized = normalize_main_window_size(size.width, size.height);
    let mut profiles = read_main_window_mode_sizes(app);
    if keypad {
        profiles.keypad = normalized;
    } else {
        profiles.normal = normalized;
        write_main_window_size(app, normalized)?;
    }
    write_main_window_mode_sizes(app, profiles)
}

fn write_main_window_size(app: &AppHandle, size: ConfigurationWindowSize) -> Result<(), String> {
    let path = main_window_size_path(app)?;
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    let bytes = serde_json::to_vec_pretty(&size).map_err(|error| error.to_string())?;
    fs::write(path, bytes).map_err(|error| error.to_string())
}

fn schedule_persist_main_window_size(app: &AppHandle, physical_width: u32, physical_height: u32) {
    if physical_width == 0 || physical_height == 0 {
        return;
    }
    let generation = MAIN_WINDOW_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst) + 1;
    let app = app.clone();
    spawn_background_task("visualtex-main-window-size", move || {
        std::thread::sleep(std::time::Duration::from_millis(250));
        if MAIN_WINDOW_SIZE_WRITE_GENERATION.load(Ordering::SeqCst) != generation {
            return;
        }
        let Some(size) = current_main_window_size(&app) else {
            return;
        };
        let keypad = MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst);
        if let Err(error) = write_main_window_profile_size(&app, keypad, size) {
            eprintln!("Unable to persist the VisualTeX main window size: {error}");
        }
    });
}

fn restore_main_window_size(app: &AppHandle) -> Result<(), String> {
    let profiles = read_main_window_mode_sizes(app);
    write_main_window_mode_sizes(app, profiles)?;
    write_main_window_size(app, profiles.normal)?;
    MAIN_WINDOW_KEYPAD_MODE.store(false, Ordering::SeqCst);
    let Some(window) = app.get_webview_window("main") else {
        return Ok(());
    };
    window
        .set_size(tauri::LogicalSize::new(
            profiles.normal.width,
            profiles.normal.height,
        ))
        .map_err(|error| error.to_string())?;
    window.center().map_err(|error| error.to_string())
}

#[tauri::command]
fn switch_main_window_mode(
    app: AppHandle,
    keypad: bool,
) -> Result<ConfigurationWindowSize, String> {
    let outgoing_keypad = MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst);
    if let Some(current) = current_main_window_size(&app) {
        write_main_window_profile_size(&app, outgoing_keypad, current)?;
    }

    let profiles = read_main_window_mode_sizes(&app);
    let target = if keypad {
        profiles.keypad
    } else {
        profiles.normal
    };

    MAIN_WINDOW_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst);
    MAIN_WINDOW_KEYPAD_MODE.store(keypad, Ordering::SeqCst);

    if let Some(window) = app.get_webview_window("main") {
        if window.is_maximized().unwrap_or(false) {
            window.unmaximize().map_err(|error| error.to_string())?;
        }
        window
            .set_size(tauri::LogicalSize::new(target.width, target.height))
            .map_err(|error| error.to_string())?;
    }

    Ok(target)
}

#[tauri::command]
fn set_main_window_keypad_mode(app: AppHandle, enabled: bool) -> Result<(), String> {
    switch_main_window_mode(app, enabled).map(|_| ())
}

#[tauri::command]
fn minimize_visualtex_main_window(app: AppHandle) -> Result<(), String> {
    let window = app
        .get_webview_window("main")
        .ok_or_else(|| "VisualTeX main window is unavailable".to_string())?;
    window
        .minimize()
        .map_err(|error| format!("Unable to minimize VisualTeX: {error}"))
}

#[tauri::command]
fn get_app_window_configuration(app: AppHandle) -> Result<AppWindowConfiguration, String> {
    #[cfg(target_os = "macos")]
    let office_editor = office::macos_offline::configuration_office_editor_window_size(&app)
        .map(|(width, height)| ConfigurationWindowSize { width, height });
    #[cfg(not(target_os = "macos"))]
    let office_editor = None;

    let profiles = read_main_window_mode_sizes(&app);
    let main = if MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst) {
        Some(profiles.normal)
    } else {
        current_main_window_size(&app).or(Some(profiles.normal))
    };

    Ok(AppWindowConfiguration {
        main,
        office_editor,
    })
}

#[tauri::command]
fn apply_app_window_configuration(
    app: AppHandle,
    configuration: AppWindowConfiguration,
) -> Result<AppWindowConfiguration, String> {
    if let Some(requested) = configuration.main {
        let size = normalize_main_window_size(requested.width, requested.height);
        write_main_window_profile_size(&app, false, size)?;
        if !MAIN_WINDOW_KEYPAD_MODE.load(Ordering::SeqCst) {
            if let Some(window) = app.get_webview_window("main") {
                window
                    .set_size(tauri::LogicalSize::new(size.width, size.height))
                    .map_err(|error| error.to_string())?;
                window.center().map_err(|error| error.to_string())?;
            }
        }
    }

    #[cfg(target_os = "macos")]
    if let Some(requested) = configuration.office_editor {
        office::macos_offline::apply_configuration_office_editor_window_size(
            &app,
            requested.width,
            requested.height,
        )?;
    }

    get_app_window_configuration(app)
}

#[tauri::command]
fn copy_png_to_clipboard(data_base64: String) -> Result<(), String> {
    let bytes = BASE64_STANDARD
        .decode(data_base64.trim())
        .map_err(|error| format!("Unable to decode PNG clipboard data: {error}"))?;
    if bytes.len() > 100 * 1024 * 1024 {
        return Err("PNG clipboard image is unexpectedly large".to_string());
    }
    const PNG_SIGNATURE: &[u8; 8] = b"\x89PNG\r\n\x1a\n";
    if !bytes.starts_with(PNG_SIGNATURE) {
        return Err("Clipboard image is not a valid PNG".to_string());
    }

    #[cfg(target_os = "macos")]
    {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|duration| duration.as_nanos())
            .unwrap_or_default();
        let temporary = std::env::temp_dir().join(format!(
            "visualtex-clipboard-{}-{nonce}.png",
            std::process::id()
        ));
        fs::write(&temporary, &bytes)
            .map_err(|error| format!("Unable to prepare PNG clipboard image: {error}"))?;

        let script = r#"on run argv
set pngPath to item 1 of argv
set pngFile to POSIX file pngPath
set the clipboard to (read pngFile as «class PNGf»)
end run"#;
        let output = Command::new("/usr/bin/osascript")
            .arg("-e")
            .arg(script)
            .arg(&temporary)
            .output();
        let _ = fs::remove_file(&temporary);
        let output = output.map_err(|error| format!("Unable to access the macOS clipboard: {error}"))?;
        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
            return Err(if stderr.is_empty() {
                "macOS rejected the PNG clipboard write".to_string()
            } else {
                format!("Unable to copy PNG to the macOS clipboard: {stderr}")
            });
        }
        return Ok(());
    }

    #[cfg(not(target_os = "macos"))]
    {
        Err("PNG clipboard export is currently supported on macOS only".to_string())
    }
}

/// An image file dropped from the system on a document (Tauri's drag-drop event gives its path): its bytes, base64.
#[tauri::command]
fn read_image_file(path: String) -> Result<String, String> {
    let target = PathBuf::from(path.trim());
    let extension = target
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    if !target.is_absolute()
        || !matches!(extension.as_str(), "png" | "jpg" | "jpeg" | "gif" | "webp" | "bmp" | "svg" | "tif" | "tiff" | "heic" | "avif")
    {
        return Err("Not an image file".to_string());
    }
    let size = fs::metadata(&target).map_err(|error| format!("Unable to read the image: {error}"))?.len();
    if size > 50 * 1024 * 1024 {
        return Err("The image is too large".to_string());
    }
    let bytes = fs::read(&target).map_err(|error| format!("Unable to read the image: {error}"))?;
    Ok(BASE64_STANDARD.encode(bytes))
}

#[tauri::command]
fn write_export_file(path: String, data_base64: String) -> Result<(), String> {
    let target = PathBuf::from(path.trim());
    if !target.is_absolute() {
        return Err("Export path must be absolute".to_string());
    }
    let extension = target
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    // zip: a document's .tex or .md with its figures; docx: a document as Word
    if !matches!(extension.as_str(), "md" | "svg" | "png" | "vtxconfig" | "vtxdoc" | "tex" | "zip" | "docx") {
        return Err("Unsupported export file extension".to_string());
    }
    let parent = target
        .parent()
        .ok_or_else(|| "Export path has no parent directory".to_string())?;
    if !parent.is_dir() {
        return Err("Export directory does not exist".to_string());
    }
    let bytes = BASE64_STANDARD
        .decode(data_base64.trim())
        .map_err(|error| format!("Unable to decode export data: {error}"))?;
    if bytes.len() > 100 * 1024 * 1024 {
        return Err("Export file is unexpectedly large".to_string());
    }

    let file_name = target
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| "Export filename is invalid".to_string())?;
    let temporary = parent.join(format!(".{file_name}.visualtex-{}.tmp", std::process::id()));
    let _ = fs::remove_file(&temporary);
    let write_result = (|| -> Result<(), String> {
        let mut file = OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&temporary)
            .map_err(|error| format!("Unable to create export file: {error}"))?;
        file.write_all(&bytes)
            .and_then(|_| file.sync_all())
            .map_err(|error| format!("Unable to write export file: {error}"))?;
        fs::rename(&temporary, &target)
            .map_err(|error| format!("Unable to finalize export file: {error}"))?;
        Ok(())
    })();
    if write_result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    write_result
}

#[cfg(all(target_os = "macos", not(debug_assertions)))]
fn claim_production_visualtex_url_handler() -> Result<(), String> {
    let executable = std::env::current_exe()
        .map_err(|error| format!("Unable to locate the VisualTeX executable: {error}"))?;
    let app_bundle = executable
        .parent()
        .and_then(|path| path.parent())
        .and_then(|path| path.parent())
        .filter(|path| path.extension().and_then(|value| value.to_str()) == Some("app"))
        .ok_or_else(|| {
            format!(
                "Unable to resolve the VisualTeX app bundle from {}",
                executable.display()
            )
        })?;

    let launch_services = "/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister";
    let register_status = Command::new(launch_services)
        .arg("-f")
        .arg(app_bundle)
        .status()
        .map_err(|error| format!("Unable to register VisualTeX with LaunchServices: {error}"))?;
    if !register_status.success() {
        return Err(format!(
            "LaunchServices registration failed with status {register_status}"
        ));
    }

    let script = r#"ObjC.import("CoreServices"); $.LSSetDefaultHandlerForURLScheme($("visualtex"), $("com.visualtex.studio"));"#;
    let handler_status = Command::new("/usr/bin/osascript")
        .args(["-l", "JavaScript", "-e", script])
        .status()
        .map_err(|error| format!("Unable to restore the VisualTeX URL handler: {error}"))?;
    if !handler_status.success() {
        return Err(format!(
            "VisualTeX URL handler restoration failed with status {handler_status}"
        ));
    }

    Ok(())
}

fn report_office_open_failure(app: &AppHandle, error: &str) {
    eprintln!("Unable to open VisualTeX offline Office Session: {error}");
    let detail = error.chars().take(800).collect::<String>();
    let message = format!(
        "VisualTeX could not open this Office formula. The Office document was not modified.\n\nDetails: {detail}"
    );
    let shown = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        app.dialog()
            .message(message)
            .title("VisualTeX Office formula")
            .kind(MessageDialogKind::Error)
            .show(|_| {});
    }));
    if shown.is_err() {
        eprintln!("Unable to show the recoverable VisualTeX Office error dialog");
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
/// macOS press-and-hold: holding a letter asks for the accent menu instead of
/// repeating the key. The formula editor shows no such menu, so a held letter
/// was typed once. Register the key as off for VisualTeX only, as code editors
/// do: the registration domain is not saved, and a value the user set wins.
#[cfg(target_os = "macos")]
fn prefer_key_repeat_over_accent_menu() {
    use objc2::runtime::AnyObject;
    use objc2::{class, msg_send};
    use objc2_foundation::NSString;
    let key = NSString::from_str("ApplePressAndHoldEnabled");
    unsafe {
        let off: *mut AnyObject = msg_send![class!(NSNumber), numberWithBool: false];
        let defaults: *mut AnyObject = msg_send![class!(NSUserDefaults), standardUserDefaults];
        if off.is_null() || defaults.is_null() {
            return;
        }
        let registration: *mut AnyObject =
            msg_send![class!(NSDictionary), dictionaryWithObject: off, forKey: &*key];
        if registration.is_null() {
            return;
        }
        let _: () = msg_send![defaults, registerDefaults: registration];
    }
}

pub fn run() {
    #[cfg(target_os = "macos")]
    prefer_key_repeat_over_accent_menu();
    if let Some(status) = office::omml_batch::run_cli_if_requested() {
        std::process::exit(status);
    }
    // an AI agent's client: relay its stdio to the running companion's /mcp
    if let Some(status) = office::automation::stdio::run_if_requested() {
        std::process::exit(status);
    }
    if let Some(status) = office::macos_offline::run_image_ink_center_cli_if_requested() {
        std::process::exit(status);
    }
    if let Some(status) = office::macos_offline_installer::run_maintenance_cli_if_requested() {
        std::process::exit(status);
    }
    let background_mode = office::background::is_background_mode();
    let application_started_at = std::time::Instant::now();
    let maintenance_install = std::env::args_os().any(|argument| {
        argument
            == std::ffi::OsStr::new(office::macos_offline_installer::MAINTENANCE_INSTALL_ARGUMENT)
    });
    let initial_office_url =
        std::env::args().find(|argument| argument.starts_with("visualtex://office/open?session="));
    let ocr_state = ocr::OcrState::default();
    let office_ocr_state = ocr_state.clone();
    let quick_ocr_state = quick_ocr::QuickOcrState::default();
    let app = tauri::Builder::default()
        .plugin(tauri_plugin_single_instance::init(
            |app, arguments, _cwd| {
                if arguments.iter().any(|argument| {
                    argument == office::background::BACKGROUND_ARGUMENT
                        || argument == office::macos_offline_installer::MAINTENANCE_INSTALL_ARGUMENT
                }) {
                    return;
                }
                #[cfg(target_os = "macos")]
                if let Some(url) = arguments
                    .iter()
                    .find(|argument| argument.starts_with("visualtex://office/open?session="))
                {
                    if let Err(error) = office::macos_offline::handle_open_url_safely(app, url) {
                        report_office_open_failure(app, &error);
                    }
                    return;
                }
                #[cfg(target_os = "macos")]
                match office::macos_offline::consume_fast_open_request_off_main(
                    app,
                    FAST_OPEN_MAIN_THREAD_WAIT,
                ) {
                    Ok(Some(true)) => return,
                    // A pending read means macOS is holding the Office sandbox
                    // behind its privacy prompt; show the workspace meanwhile.
                    Ok(Some(false)) | Ok(None) => {}
                    Err(error) => {
                        eprintln!("Unable to consume VisualTeX Office fast-open request: {error}");
                        return;
                    }
                }
                let _ = office::background::reveal_main_window(app);
            },
        ))
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(ocr_state)
        .manage(quick_ocr_state)
        .setup(move |app| {
            let setup_result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(
                || -> Result<(), Box<dyn std::error::Error>> {
                    if maintenance_install {
                        match office::macos_offline_installer::install(app.handle()) {
                            Ok(status) => {
                                println!(
                                    "{}",
                                    serde_json::to_string(&status)
                                        .unwrap_or_else(|_| "{\"ok\":true}".to_string())
                                );
                                std::process::exit(0);
                            }
                            Err(error) => {
                                eprintln!("VisualTeX native Office installation failed: {error}");
                                std::process::exit(1);
                            }
                        }
                    }

                    #[cfg(not(debug_assertions))]
                    {
                        #[cfg(target_os = "macos")]
                        spawn_background_task("visualtex-url-handler-repair", || {
                            // LaunchServices repair is maintenance. Let a cold Office
                            // URL hydrate its resident editor before lsregister scans
                            // the application bundle.
                            std::thread::sleep(std::time::Duration::from_secs(2));
                            if let Err(error) = claim_production_visualtex_url_handler() {
                                eprintln!("Unable to claim the production visualtex:// handler: {error}");
                            }
                        });
                    }

                    // Set the application icon while setup is still on AppKit's main
                    // thread. Background-agent launches may stay accessory-only for a
                    // long time, but the first later transition to a regular app must
                    // already have the real VisualTeX Dock icon installed.
                    office::background::install_application_icon(app.handle())
                        .map_err(std::io::Error::other)?;
                    restore_main_window_size(app.handle()).map_err(std::io::Error::other)?;
                    if let Err(error) = quick_ocr::initialize(app.handle()) {
                        eprintln!("VisualTeX quick OCR initialization warning: {error}");
                    }
                    if let Err(error) = office::hotkeys::initialize(app.handle()) {
                        eprintln!("VisualTeX Office hotkey initialization warning: {error}");
                    }
                    let office_state = office::initialize(app.handle(), office_ocr_state.clone())
                        .map_err(std::io::Error::other)?;
                    if let Err(error) = office::powerpoint_native::start_double_click_monitor(
                        app.handle().clone(),
                        office_state.powerpoint_interactions.clone(),
                    ) {
                        eprintln!("Unable to start PowerPoint double-click monitor: {error}");
                    }
                    app.manage(office_state.clone());
                    // AI agents' operations run the add-ins' macros from this app's main thread
                    office::automation::hosts::set_app(app.handle().clone());
                    if initial_office_url.is_some() || background_mode {
                        // Hide the desktop shell before WebKit prewarming. Otherwise a
                        // cold Office URL can flash the main editor while the resident
                        // formula WebViews initialize.
                        office::background::hide_main_window(app.handle())
                            .map_err(std::io::Error::other)?;
                    }
                    #[cfg(target_os = "macos")]
                    office::macos_offline::start_fast_open_inbox_watcher(app.handle().clone());
                    #[cfg(target_os = "macos")]
                    {
                        office::wps::bridge::start(
                            app.handle().clone(),
                            office_state.install_token.as_str().to_string(),
                            office::wps::installer::addin_log_path(&office_state),
                        );
                        office::wps::installer::auto_install_in_background(app.handle().clone(), office_state.clone());
                    }
                    if let Err(error) =
                        office::macos_offline::prewarm_office_editor_windows(app.handle())
                    {
                        // Prewarming is an optimization. handle_open_url still creates
                        // the fixed host window lazily if WebKit was unavailable here.
                        eprintln!("Unable to prewarm VisualTeX Office editors: {error}");
                    }
                    if let Err(error) = anywhere::initialize(app.handle()) {
                        eprintln!("VisualTeX Anywhere initialization warning: {error}");
                    }
                    #[cfg(not(target_os = "macos"))]
                    office::lifecycle::start(office_state);
                    // The Mac add-ins need no companion server; AI agents reach /mcp on it.
                    #[cfg(target_os = "macos")]
                    if office::automation::settings::settings(&office_state.paths.root).enabled {
                        office::lifecycle::start(office_state.clone());
                    }
                    if let Some(url) = initial_office_url.as_deref() {
                        if let Err(error) =
                            office::macos_offline::handle_open_url_safely(app.handle(), url)
                        {
                            // A malformed, stale, or otherwise unusable Office Session is
                            // an interaction failure, not an application startup failure.
                            // Keep the resident process alive so correcting the document or
                            // retrying another formula never enters a relaunch crash loop.
                            report_office_open_failure(app.handle(), &error);
                        }
                    } else if !background_mode {
                        office::background::reveal_main_window(app.handle())
                            .map_err(std::io::Error::other)?;
                    }
                    Ok(())
                },
            ));
            match setup_result {
                Ok(result) => result,
                Err(payload) => Err(std::io::Error::other(format!(
                    "VisualTeX recovered from an internal startup error: {}",
                    panic_payload_detail(payload.as_ref())
                ))
                .into()),
            }
        })
        .invoke_handler(tauri::generate_handler![
            write_export_file,
            read_image_file,
            document_export::read_import_file,
            document_export::fetch_image_url,
            document_export::pdf_print_open,
            document_export::pdf_print_save,
            document_export::pdf_print_close,
            copy_png_to_clipboard,
            document_library::document_library_list,
            document_library::document_library_read,
            document_library::document_library_create,
            document_library::document_library_save,
            document_library::document_library_rename,
            document_library::document_library_duplicate,
            document_library::document_library_move,
            document_library::document_library_trash,
            document_library::document_library_restore,
            document_library::document_library_delete,
            document_library::document_library_empty_trash,
            document_library::document_library_create_folder,
            document_library::document_library_rename_folder,
            document_library::document_library_delete_folder,
            document_library::document_library_search,
            document_library::document_library_reveal,
            quick_ocr::capture_quick_ocr_screenshot,
            quick_ocr::wait_for_quick_ocr_system_screenshot,
            quick_ocr::configure_silent_ocr,
            office::hotkeys::configure_office_hotkeys,
            anywhere::anywhere_ready,
            anywhere::anywhere_layout,
            anywhere::anywhere_commit,
            anywhere::anywhere_commit_image,
            anywhere::anywhere_configure_hotkey,
            anywhere::anywhere_capture_screenshot,
            anywhere::anywhere_read_clipboard,
            anywhere::anywhere_dismiss,
            anywhere::anywhere_marked,
            anywhere::anywhere_status,
            anywhere::anywhere_log,
            set_app_theme,
            get_app_window_configuration,
            apply_app_window_configuration,
            switch_main_window_mode,
            set_main_window_keypad_mode,
            minimize_visualtex_main_window,
            system_math_glyphs::probe_macos_math_fonts,
            system_math_glyphs::list_system_font_families,
            system_math_glyphs::extract_macos_math_glyph,
            ocr::get_ocr_provider_configuration,
            ocr::save_ocr_provider_configuration,
            ocr::get_ocr_runtime_status,
            ocr::install_ocr_runtime,
            ocr::recognize_formula_image,
            ocr::prewarm_ocr_model,
            ocr::cancel_ocr_recognition,
            ocr::restart_ocr_worker,
            ocr::reset_ocr_runtime,
            ocr::install_optional_ocr_model,
            ocr::remove_optional_ocr_model,
            office::lifecycle::get_office_companion_status,
            office::automation::get_mcp_status,
            office::automation::set_mcp_settings,
            office::automation::reset_mcp_access,
            office::automation::get_mcp_audit,
            office::automation::render::mcp_render_result,
            office::lifecycle::start_office_companion,
            office::lifecycle::stop_office_companion,
            office::lifecycle::set_office_background_start,
            office::lifecycle::open_word,
            office::lifecycle::open_powerpoint,
            office::macos_offline::get_macos_offline_document_import_request,
            office::macos_offline::report_macos_offline_latex_redraw_stage,
            office::macos_offline::resolve_macos_offline_latex_redraw_font_sizes,
            office::macos_offline::focus_macos_offline_document_import_target,
            office::macos_offline::restore_macos_offline_document_import_window,
            office::macos_offline::get_macos_offline_document_import_progress,
            office::macos_offline::commit_macos_offline_document_import,
            office::macos_offline::cancel_macos_offline_document_import,
            office::macos_offline::note_macos_agent_session,
            office::macos_offline::get_macos_offline_office_session,
            office::macos_offline::report_macos_offline_office_editor_prewarm_diagnostic,
            office::macos_offline::report_macos_offline_office_editor_prewarmed,
            office::macos_offline::update_macos_offline_office_session,
            office::macos_offline::delete_macos_offline_office_session,
            office::macos_offline::commit_macos_offline_office_session,
            office::macos_offline::cancel_macos_offline_office_session,
            office::macos_offline::get_macos_offline_office_editor_activation,
            office::macos_offline::report_macos_offline_office_editor_ready,
            office::macos_offline::present_macos_offline_office_editor_window,
            office::macos_offline::close_macos_offline_office_editor_window,
            office::macos_offline::get_macos_offline_plugin_health,
            office::macos_offline_installer::get_macos_offline_office_install_status,
            office::macos_offline_installer::install_macos_offline_office_addins,
            office::macos_offline_installer::repair_macos_offline_office_addins,
            office::macos_offline_installer::uninstall_macos_offline_office_addins,
            office::wps::installer::get_wps_addin_status,
            office::wps::installer::install_wps_addin,
            office::wps::installer::uninstall_wps_addin,
            office::wps::installer::open_wps_office,
            office::macos_offline_installer::get_macos_office_hosts_running,
            office::macos_offline_installer::register_macos_powerpoint_addin,
            office::macos_offline_installer::request_quit_macos_office_hosts_for_addin_update,
            office::macos_offline_installer::open_macos_office_addin_maintenance_log,
            office::macos_offline_installer::reveal_macos_powerpoint_addin,
            office::macos_offline_installer::open_macos_powerpoint_addin_tutorial,
            office::macos_offline_installer::get_macos_major_version,
            office::macos_offline_installer::open_macos_full_disk_access_settings
        ])
        .build(tauri::generate_context!());
    let app = match app {
        Ok(app) => app,
        Err(error) => {
            // A Tauri/AppKit setup failure must not unwind through Objective-C's
            // didFinishLaunching callback, which macOS reports as SIGABRT.
            eprintln!("Unable to build VisualTeX: {error}");
            return;
        }
    };

    app.run(move |app, event| match event {
        #[cfg(target_os = "macos")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::Resized(size),
            ..
        } => {
            if label == "main" {
                schedule_persist_main_window_size(app, size.width, size.height);
            } else {
                office::macos_offline::schedule_persist_office_editor_window_size(
                    app,
                    &label,
                    size.width,
                    size.height,
                );
            }
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::CloseRequested { api, .. },
            ..
        } if label == "main" => {
            api.prevent_close();
            let _ = office::background::hide_main_window(app);
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::Opened { urls } => {
            for url in urls {
                if url.scheme() == "visualtex" {
                    if let Err(error) =
                        office::macos_offline::handle_open_url_safely(app, url.as_str())
                    {
                        report_office_open_failure(app, &error);
                    }
                }
            }
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::Reopen { .. } => {
            // Office fast-open writes a validated request into the host sandbox,
            // then asks LaunchServices to deliver this Reopen event to the resident
            // app. Consume that inbox before any startup/Dock reveal logic so the
            // editor can open without a Terminal prompt, custom URL, or /tmp socket.
            // The inbox read runs off the AppKit thread: macOS may hold it until
            // the user answers the "access data from other apps" prompt.
            match office::macos_offline::consume_fast_open_request_off_main(
                app,
                FAST_OPEN_MAIN_THREAD_WAIT,
            ) {
                Ok(Some(true)) => return,
                Ok(Some(false)) => {}
                // Still pending: the worker opens the editor once macOS allows
                // the read, so do not reveal the workspace over it.
                Ok(None) => return,
                Err(error) => {
                    eprintln!("Unable to consume VisualTeX Office fast-open request on Reopen: {error}");
                    report_office_open_failure(app, &error);
                    return;
                }
            }
            // LaunchServices can emit Reopen for the initial background launch.
            // Never reveal the desktop workspace during that startup window.
            if background_mode
                && application_started_at.elapsed() < std::time::Duration::from_secs(2)
            {
                return;
            }
            // A native Office launch starts a short-lived second VisualTeX
            // process. macOS can deliver Reopen before the single-instance URL
            // arguments establish the active formula Session. Defer the Dock
            // behavior briefly so that launch cannot reveal the desktop main
            // window over a hydrating Office editor.
            let app = app.clone();
            spawn_background_task("visualtex-office-reopen", move || {
                std::thread::sleep(std::time::Duration::from_millis(150));
                if office::macos_offline::focus_open_office_editor(&app) {
                    return;
                }
                // Word and PowerPoint write request.json before asking
                // LaunchServices to open VisualTeX. On a slower machine the
                // Reopen event can arrive before handle_open_url has registered
                // the active editor Session. Do not mistake that Office launch
                // for an explicit Dock click and reveal the desktop workspace.
                if office::macos_offline::has_recent_office_editor_request(
                    std::time::Duration::from_secs(3),
                ) {
                    return;
                }
                let _ = office::background::reveal_main_window(&app);
            });
        }
        tauri::RunEvent::ExitRequested { .. } => {
            #[cfg(target_os = "macos")]
            if let Err(error) = office::background::pause_launch_agent_for_quit() {
                eprintln!("Unable to pause VisualTeX Office background service: {error}");
            }
        }
        _ => {}
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn export_file_write_is_atomic_and_overwrites_existing_output() {
        let directory = tempfile::tempdir().expect("temporary export directory");
        let target = directory.path().join("formula.svg");
        let first = "<svg xmlns=\"http://www.w3.org/2000/svg\"></svg>";
        let second = "<svg xmlns=\"http://www.w3.org/2000/svg\"><path/></svg>";

        write_export_file(
            target.to_string_lossy().into_owned(),
            BASE64_STANDARD.encode(first.as_bytes()),
        )
        .expect("first export should succeed");
        assert_eq!(fs::read_to_string(&target).unwrap(), first);

        write_export_file(
            target.to_string_lossy().into_owned(),
            BASE64_STANDARD.encode(second.as_bytes()),
        )
        .expect("second export should replace the existing file");
        assert_eq!(fs::read_to_string(&target).unwrap(), second);
        assert!(fs::read_dir(directory.path()).unwrap().all(|entry| {
            !entry
                .unwrap()
                .file_name()
                .to_string_lossy()
                .contains("visualtex-")
        }));

        let unsupported = directory.path().join("formula.exe");
        assert!(write_export_file(
            unsupported.to_string_lossy().into_owned(),
            BASE64_STANDARD.encode(b"not allowed"),
        )
        .is_err());
    }
}
