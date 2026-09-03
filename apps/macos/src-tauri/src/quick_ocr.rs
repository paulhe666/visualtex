use crate::{OcrImageRequest, OcrState, ALLOWED_MODELS, MAX_IMAGE_BYTES};
use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::Serialize;
use serde_json::Value;
use std::collections::HashMap;
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tauri::utils::config::BackgroundThrottlingPolicy;
use tauri::{AppHandle, Emitter, Manager, WebviewUrl, WebviewWindowBuilder};

const QUICK_OCR_HUD_LABEL: &str = "quick-ocr-hud";
const QUICK_OCR_HUD_EVENT: &str = "quick-ocr-status";
const DEFAULT_SILENT_OCR_MODEL: &str = "PP-FormulaNet_plus-M";
const SYSTEM_SCREENSHOT_WAIT_TIMEOUT: Duration = Duration::from_secs(60);
const DEFAULT_SILENT_OCR_COPY_FORMAT: &str = "display-dollar";
const ALLOWED_SILENT_OCR_COPY_FORMATS: &[&str] = &[
    "raw",
    "inline-dollar",
    "inline-text-double-dollar",
    "inline-paren",
    "display-dollar",
    "display-bracket",
    "equation",
    "equation-star",
    "align",
    "align-star",
    "aligned",
    "gather",
    "gather-star",
    "multline",
    "multline-star",
    "equation-split",
    "equation-star-split",
];

#[derive(Clone)]
pub(crate) struct QuickOcrState {
    silent_enabled: Arc<AtomicBool>,
    silent_busy: Arc<AtomicBool>,
    model: Arc<Mutex<String>>,
    copy_format: Arc<Mutex<String>>,
}

impl Default for QuickOcrState {
    fn default() -> Self {
        Self {
            silent_enabled: Arc::new(AtomicBool::new(false)),
            silent_busy: Arc::new(AtomicBool::new(false)),
            model: Arc::new(Mutex::new(DEFAULT_SILENT_OCR_MODEL.to_string())),
            copy_format: Arc::new(Mutex::new(DEFAULT_SILENT_OCR_COPY_FORMAT.to_string())),
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct QuickOcrCapture {
    data_base64: String,
    extension: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct QuickOcrHudPayload {
    status: &'static str,
    message: String,
    progress: u8,
}

fn capture_path() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    std::env::temp_dir().join(format!(
        "visualtex-quick-ocr-{}-{nonce}.png",
        std::process::id()
    ))
}

fn capture_selection_png() -> Result<Option<Vec<u8>>, String> {
    #[cfg(target_os = "macos")]
    {
        let path = capture_path();
        let _ = fs::remove_file(&path);
        let output = Command::new("/usr/sbin/screencapture")
            .arg("-i")
            .arg("-x")
            .arg("-t")
            .arg("png")
            .arg(&path)
            .output()
            .map_err(|error| format!("Unable to start macOS screenshot selection: {error}"))?;

        if !output.status.success() {
            let _ = fs::remove_file(&path);
            let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
            return if stderr.is_empty() {
                Ok(None)
            } else {
                Err(format!("macOS screenshot capture failed: {stderr}"))
            };
        }
        if !path.is_file() {
            return Ok(None);
        }
        let bytes = fs::read(&path)
            .map_err(|error| format!("Unable to read the captured screenshot: {error}"));
        let _ = fs::remove_file(&path);
        let bytes = bytes?;
        if bytes.is_empty() {
            return Ok(None);
        }
        if bytes.len() > MAX_IMAGE_BYTES {
            return Err("The captured screenshot is too large for OCR".to_string());
        }
        Ok(Some(bytes))
    }

    #[cfg(not(target_os = "macos"))]
    {
        Err("Quick OCR screenshot capture is currently supported on macOS only".to_string())
    }
}

#[cfg(target_os = "macos")]
fn expand_home_path(value: &str) -> PathBuf {
    let trimmed = value.trim();
    if trimmed == "~" {
        return std::env::var_os("HOME")
            .map(PathBuf::from)
            .unwrap_or_else(|| PathBuf::from("/tmp"));
    }
    if let Some(relative) = trimmed.strip_prefix("~/") {
        if let Some(home) = std::env::var_os("HOME") {
            return PathBuf::from(home).join(relative);
        }
    }
    PathBuf::from(trimmed)
}

#[cfg(target_os = "macos")]
fn screenshot_directory() -> PathBuf {
    let default = std::env::var_os("HOME")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("/tmp"))
        .join("Desktop");
    let output = Command::new("/usr/bin/defaults")
        .args(["read", "com.apple.screencapture", "location"])
        .output();
    let Ok(output) = output else {
        return default;
    };
    if !output.status.success() {
        return default;
    }
    let value = String::from_utf8_lossy(&output.stdout).trim().to_string();
    if value.is_empty() {
        return default;
    }
    let candidate = expand_home_path(&value);
    if candidate.is_dir() {
        candidate
    } else {
        default
    }
}

#[cfg(target_os = "macos")]
fn supported_system_screenshot_extension(path: &Path) -> Option<String> {
    let extension = path.extension()?.to_str()?.to_ascii_lowercase();
    matches!(
        extension.as_str(),
        "png" | "jpg" | "jpeg" | "webp" | "bmp" | "tif" | "tiff"
    )
    .then_some(extension)
}

#[cfg(target_os = "macos")]
fn screenshot_directory_snapshot(directory: &Path) -> HashMap<PathBuf, (SystemTime, u64)> {
    let mut snapshot = HashMap::new();
    let Ok(entries) = fs::read_dir(directory) else {
        return snapshot;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if supported_system_screenshot_extension(&path).is_none() {
            continue;
        }
        let Ok(metadata) = entry.metadata() else {
            continue;
        };
        if !metadata.is_file() {
            continue;
        }
        snapshot.insert(
            path,
            (metadata.modified().unwrap_or(UNIX_EPOCH), metadata.len()),
        );
    }
    snapshot
}

#[cfg(target_os = "macos")]
fn read_stable_screenshot_file(path: &Path) -> Result<Option<(Vec<u8>, String)>, String> {
    let Some(extension) = supported_system_screenshot_extension(path) else {
        return Ok(None);
    };
    let first = fs::metadata(path)
        .map_err(|error| format!("Unable to inspect the system screenshot: {error}"))?;
    if first.len() == 0 {
        return Ok(None);
    }
    if first.len() > MAX_IMAGE_BYTES as u64 {
        return Err("The system screenshot is too large for OCR".to_string());
    }
    std::thread::sleep(Duration::from_millis(120));
    let second = match fs::metadata(path) {
        Ok(metadata) => metadata,
        Err(_) => return Ok(None),
    };
    if second.len() != first.len() || second.len() == 0 {
        return Ok(None);
    }
    let bytes = fs::read(path)
        .map_err(|error| format!("Unable to read the system screenshot: {error}"))?;
    if bytes.is_empty() {
        return Ok(None);
    }
    Ok(Some((bytes, extension)))
}

#[cfg(target_os = "macos")]
fn copy_nsdata_bytes(data: &objc2_foundation::NSData) -> Option<Vec<u8>> {
    use std::ffi::c_void;
    use std::ptr::NonNull;

    let length = data.length() as usize;
    if length == 0 || length > MAX_IMAGE_BYTES {
        return None;
    }
    let mut bytes = vec![0u8; length];
    let pointer = NonNull::new(bytes.as_mut_ptr().cast::<c_void>())?;
    unsafe {
        data.getBytes_length(pointer, length);
    }
    Some(bytes)
}

#[cfg(target_os = "macos")]
fn pasteboard_image_if_changed(initial_change_count: isize) -> Option<(Vec<u8>, String)> {
    use objc2_app_kit::{NSPasteboard, NSPasteboardTypePNG, NSPasteboardTypeTIFF};

    let pasteboard = NSPasteboard::generalPasteboard();
    if pasteboard.changeCount() == initial_change_count {
        return None;
    }
    unsafe {
        if let Some(data) = pasteboard.dataForType(NSPasteboardTypePNG) {
            if let Some(bytes) = copy_nsdata_bytes(&data) {
                return Some((bytes, "png".to_string()));
            }
        }
        if let Some(data) = pasteboard.dataForType(NSPasteboardTypeTIFF) {
            if let Some(bytes) = copy_nsdata_bytes(&data) {
                return Some((bytes, "tiff".to_string()));
            }
        }
    }
    None
}

#[cfg(target_os = "macos")]
struct SystemScreenshotBaseline {
    directory: PathBuf,
    files: HashMap<PathBuf, (SystemTime, u64)>,
    pasteboard_change_count: isize,
}

#[cfg(target_os = "macos")]
fn create_system_screenshot_baseline() -> SystemScreenshotBaseline {
    use objc2_app_kit::NSPasteboard;

    let directory = screenshot_directory();
    let files = screenshot_directory_snapshot(&directory);
    let pasteboard_change_count = NSPasteboard::generalPasteboard().changeCount();
    SystemScreenshotBaseline {
        directory,
        files,
        pasteboard_change_count,
    }
}

#[cfg(target_os = "macos")]
fn wait_for_next_system_screenshot(
    baseline: SystemScreenshotBaseline,
) -> Result<Option<(Vec<u8>, String)>, String> {
    let deadline = Instant::now() + SYSTEM_SCREENSHOT_WAIT_TIMEOUT;

    while Instant::now() < deadline {
        if let Some(capture) = pasteboard_image_if_changed(baseline.pasteboard_change_count) {
            return Ok(Some(capture));
        }

        let mut newest_candidate: Option<(PathBuf, SystemTime)> = None;
        if let Ok(entries) = fs::read_dir(&baseline.directory) {
            for entry in entries.flatten() {
                let path = entry.path();
                if supported_system_screenshot_extension(&path).is_none() {
                    continue;
                }
                let Ok(metadata) = entry.metadata() else {
                    continue;
                };
                if !metadata.is_file() || metadata.len() == 0 {
                    continue;
                }
                let modified = metadata.modified().unwrap_or(UNIX_EPOCH);
                let changed = baseline
                    .files
                    .get(&path)
                    .map(|(old_modified, old_len)| {
                        modified > *old_modified || metadata.len() != *old_len
                    })
                    .unwrap_or(true);
                if !changed {
                    continue;
                }
                if newest_candidate
                    .as_ref()
                    .is_none_or(|(_, current_modified)| modified > *current_modified)
                {
                    newest_candidate = Some((path, modified));
                }
            }
        }
        if let Some((path, _)) = newest_candidate {
            if let Some(capture) = read_stable_screenshot_file(&path)? {
                return Ok(Some(capture));
            }
        }
        std::thread::sleep(Duration::from_millis(120));
    }

    Ok(None)
}

#[cfg(not(target_os = "macos"))]
struct SystemScreenshotBaseline;

#[cfg(not(target_os = "macos"))]
fn create_system_screenshot_baseline() -> SystemScreenshotBaseline {
    SystemScreenshotBaseline
}

#[cfg(not(target_os = "macos"))]
fn wait_for_next_system_screenshot(
    _baseline: SystemScreenshotBaseline,
) -> Result<Option<(Vec<u8>, String)>, String> {
    Err("Waiting for the next system screenshot is currently supported on macOS only".to_string())
}

fn write_text_clipboard(text: &str) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        let mut child = Command::new("/usr/bin/pbcopy")
            .stdin(Stdio::piped())
            .spawn()
            .map_err(|error| format!("Unable to open the macOS clipboard: {error}"))?;
        child
            .stdin
            .as_mut()
            .ok_or_else(|| "Unable to access the macOS clipboard input".to_string())?
            .write_all(text.as_bytes())
            .map_err(|error| format!("Unable to write LaTeX to the clipboard: {error}"))?;
        let status = child
            .wait()
            .map_err(|error| format!("Unable to finish the clipboard write: {error}"))?;
        if !status.success() {
            return Err("macOS rejected the LaTeX clipboard write".to_string());
        }
        Ok(())
    }

    #[cfg(not(target_os = "macos"))]
    {
        Err("Silent OCR clipboard output is currently supported on macOS only".to_string())
    }
}

fn ensure_hud_window(app: &AppHandle) -> Result<(), String> {
    if app.get_webview_window(QUICK_OCR_HUD_LABEL).is_some() {
        return Ok(());
    }
    WebviewWindowBuilder::new(
        app,
        QUICK_OCR_HUD_LABEL,
        WebviewUrl::App("index.html?view=quick-ocr-hud".into()),
    )
    .title("VisualTeX OCR")
    .inner_size(360.0, 96.0)
    .center()
    .resizable(false)
    .decorations(false)
    .always_on_top(true)
    .focused(false)
    .focusable(false)
    .skip_taskbar(true)
    .visible(false)
    .background_throttling(BackgroundThrottlingPolicy::Disabled)
    .build()
    .map_err(|error| format!("Unable to initialize the quick OCR status window: {error}"))?;
    Ok(())
}

fn emit_hud(app: &AppHandle, status: &'static str, message: impl Into<String>, progress: u8) {
    if ensure_hud_window(app).is_err() {
        return;
    }
    let Some(window) = app.get_webview_window(QUICK_OCR_HUD_LABEL) else {
        return;
    };
    let payload = QuickOcrHudPayload {
        status,
        message: message.into(),
        progress: progress.min(100),
    };
    let _ = window.emit(QUICK_OCR_HUD_EVENT, payload);
    let _ = window.show();
}

fn schedule_hud_hide(app: AppHandle, delay: Duration) {
    tauri::async_runtime::spawn(async move {
        tokio::time::sleep(delay).await;
        if let Some(window) = app.get_webview_window(QUICK_OCR_HUD_LABEL) {
            let _ = window.hide();
        }
    });
}

pub(crate) fn handle_ocr_progress(app: &AppHandle, response: &Value) {
    let Some(state) = app.try_state::<QuickOcrState>() else {
        return;
    };
    if !state.silent_busy.load(Ordering::SeqCst) {
        return;
    }
    let stage = response
        .get("stage")
        .and_then(Value::as_str)
        .unwrap_or_default();
    let message = response
        .get("message")
        .and_then(Value::as_str)
        .unwrap_or("正在识别公式…");
    let progress = match stage {
        "preprocess" => 38,
        "model" => 58,
        "inference" => 76,
        _ => 48,
    };
    emit_hud(app, "running", message, progress);
}

fn current_model(state: &QuickOcrState) -> String {
    state
        .model
        .lock()
        .map(|model| model.clone())
        .unwrap_or_else(|_| DEFAULT_SILENT_OCR_MODEL.to_string())
}

fn current_copy_format(state: &QuickOcrState) -> String {
    state
        .copy_format
        .lock()
        .map(|format| format.clone())
        .unwrap_or_else(|_| DEFAULT_SILENT_OCR_COPY_FORMAT.to_string())
}

fn wrap_environment(name: &str, body: &str) -> String {
    format!("\\begin{{{name}}}\n{body}\n\\end{{{name}}}")
}

fn starts_with_ascii_at(source: &[u8], index: usize, value: &[u8]) -> bool {
    source
        .get(index..index.saturating_add(value.len()))
        .is_some_and(|slice| slice == value)
}

fn escaped_at(source: &[u8], index: usize) -> bool {
    let mut cursor = index;
    let mut count = 0usize;
    while cursor > 0 && source[cursor - 1] == b'\\' {
        cursor -= 1;
        count += 1;
    }
    count % 2 == 1
}

fn read_environment_token(source: &[u8], index: usize) -> Option<(bool, usize)> {
    let (is_begin, prefix) = if starts_with_ascii_at(source, index, b"\\begin{") {
        (true, b"\\begin{".as_slice())
    } else if starts_with_ascii_at(source, index, b"\\end{") {
        (false, b"\\end{".as_slice())
    } else {
        return None;
    };
    let mut cursor = index + prefix.len();
    while cursor < source.len() && source[cursor] != b'}' {
        let byte = source[cursor];
        if !byte.is_ascii_alphabetic() && byte != b'*' {
            return None;
        }
        cursor += 1;
    }
    (cursor < source.len()).then_some((is_begin, cursor + 1))
}

fn top_level_relation_index(latex: &str) -> Option<usize> {
    const RELATION_COMMANDS: &[&[u8]] = &[
        b"\\Longleftrightarrow",
        b"\\Longrightarrow",
        b"\\Leftrightarrow",
        b"\\Rightarrow",
        b"\\leftrightarrow",
        b"\\rightarrow",
        b"\\leftarrow",
        b"\\subseteq",
        b"\\supseteq",
        b"\\notin",
        b"\\approx",
        b"\\equiv",
        b"\\simeq",
        b"\\propto",
        b"\\mapsto",
        b"\\subset",
        b"\\supset",
        b"\\cong",
        b"\\neq",
        b"\\leq",
        b"\\geq",
        b"\\sim",
        b"\\to",
        b"\\ne",
        b"\\le",
        b"\\ge",
        b"\\in",
    ];
    let source = latex.as_bytes();
    let mut brace_depth = 0usize;
    let mut environment_depth = 0usize;
    let mut index = 0usize;
    while index < source.len() {
        if source[index] == b'\\' {
            if let Some((is_begin, end)) = read_environment_token(source, index) {
                if is_begin {
                    environment_depth += 1;
                } else {
                    environment_depth = environment_depth.saturating_sub(1);
                }
                index = end;
                continue;
            }
        }
        let byte = source[index];
        if byte == b'{' && !escaped_at(source, index) {
            brace_depth += 1;
            index += 1;
            continue;
        }
        if byte == b'}' && !escaped_at(source, index) {
            brace_depth = brace_depth.saturating_sub(1);
            index += 1;
            continue;
        }
        if brace_depth == 0 && environment_depth == 0 {
            if byte == b'&' && !escaped_at(source, index) {
                return None;
            }
            if matches!(byte, b'=' | b'<' | b'>') {
                return Some(index);
            }
            if byte == b'\\' {
                for command in RELATION_COMMANDS {
                    if !starts_with_ascii_at(source, index, command) {
                        continue;
                    }
                    let next = source.get(index + command.len()).copied();
                    if next.is_some_and(|value| value.is_ascii_alphabetic()) {
                        continue;
                    }
                    return Some(index);
                }
            }
        }
        index += 1;
    }
    None
}

fn add_alignment_marker(latex: &str) -> String {
    let Some(index) = top_level_relation_index(latex) else {
        return latex.to_string();
    };
    format!("{}&{}", &latex[..index], &latex[index..])
}

fn format_rows(lines: &[String], align_relations: bool) -> String {
    lines
        .iter()
        .enumerate()
        .map(|(index, line)| {
            let content = if align_relations {
                add_alignment_marker(line)
            } else {
                line.clone()
            };
            if index + 1 < lines.len() {
                format!("{content} \\\\")
            } else {
                content
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn balanced_brace_end(source: &[u8], opening_index: usize) -> Option<usize> {
    if source.get(opening_index) != Some(&b'{') {
        return None;
    }
    let mut depth = 0usize;
    let mut index = opening_index;
    while index < source.len() {
        if source[index] == b'{' && !escaped_at(source, index) {
            depth += 1;
        } else if source[index] == b'}' && !escaped_at(source, index) {
            depth = depth.saturating_sub(1);
            if depth == 0 {
                return Some(index + 1);
            }
        }
        index += 1;
    }
    None
}

fn format_inline_text_double_dollar(latex: &str) -> String {
    let source = latex.as_bytes();
    let mut result = String::new();
    let mut math_start = 0usize;
    let mut brace_depth = 0usize;
    let mut environment_depth = 0usize;
    let mut index = 0usize;

    while index < source.len() {
        if source[index] == b'\\' {
            if let Some((is_begin, end)) = read_environment_token(source, index) {
                if is_begin {
                    environment_depth += 1;
                } else {
                    environment_depth = environment_depth.saturating_sub(1);
                }
                index = end;
                continue;
            }
        }
        if brace_depth == 0
            && environment_depth == 0
            && starts_with_ascii_at(source, index, b"\\text{")
        {
            let opening = index + b"\\text".len();
            if let Some(end) = balanced_brace_end(source, opening) {
                let math = latex[math_start..index].trim();
                if !math.is_empty() {
                    result.push_str("$$");
                    result.push_str(math);
                    result.push_str("$$");
                }
                result.push_str(&latex[opening + 1..end - 1]);
                math_start = end;
                index = end;
                continue;
            }
        }
        if source[index] == b'{' && !escaped_at(source, index) {
            brace_depth += 1;
        } else if source[index] == b'}' && !escaped_at(source, index) {
            brace_depth = brace_depth.saturating_sub(1);
        }
        index += 1;
    }

    let math = latex[math_start..].trim();
    if !math.is_empty() {
        result.push_str("$$");
        result.push_str(math);
        result.push_str("$$");
    }
    result
}

fn format_silent_ocr_latex(lines: &[String], format: &str) -> String {
    let lines = lines
        .iter()
        .map(|line| line.trim())
        .filter(|line| !line.is_empty())
        .map(ToOwned::to_owned)
        .collect::<Vec<_>>();
    if lines.is_empty() {
        return String::new();
    }
    match format {
        "raw" => lines.join("\n"),
        "inline-dollar" => lines
            .iter()
            .map(|line| format!("${line}$"))
            .collect::<Vec<_>>()
            .join("\n"),
        "inline-text-double-dollar" => lines
            .iter()
            .map(|line| format_inline_text_double_dollar(line))
            .collect::<Vec<_>>()
            .join("\n"),
        "inline-paren" => lines
            .iter()
            .map(|line| format!("\\({line}\\)"))
            .collect::<Vec<_>>()
            .join("\n"),
        "display-dollar" => lines
            .iter()
            .map(|line| format!("$$\n{line}\n$$"))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "display-bracket" => lines
            .iter()
            .map(|line| format!("\\[\n{line}\n\\]"))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "equation" => lines
            .iter()
            .map(|line| wrap_environment("equation", line))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "equation-star" => lines
            .iter()
            .map(|line| wrap_environment("equation*", line))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "align" => wrap_environment("align", &format_rows(&lines, true)),
        "align-star" => wrap_environment("align*", &format_rows(&lines, true)),
        "aligned" => format!(
            "\\[\n{}\n\\]",
            wrap_environment("aligned", &format_rows(&lines, true))
        ),
        "gather" => wrap_environment("gather", &format_rows(&lines, false)),
        "gather-star" => wrap_environment("gather*", &format_rows(&lines, false)),
        "multline" => wrap_environment("multline", &format_rows(&lines, false)),
        "multline-star" => wrap_environment("multline*", &format_rows(&lines, false)),
        "equation-split" => wrap_environment(
            "equation",
            &wrap_environment("split", &format_rows(&lines, true)),
        ),
        "equation-star-split" => wrap_environment(
            "equation*",
            &wrap_environment("split", &format_rows(&lines, true)),
        ),
        _ => lines
            .iter()
            .map(|line| format!("$$\n{line}\n$$"))
            .collect::<Vec<_>>()
            .join("\n\n"),
    }
}

async fn run_silent_ocr(app: AppHandle) {
    let Some(state_guard) = app.try_state::<QuickOcrState>() else {
        return;
    };
    let state = state_guard.inner().clone();
    if !state.silent_enabled.load(Ordering::SeqCst) {
        return;
    }
    if state.silent_busy.swap(true, Ordering::SeqCst) {
        emit_hud(&app, "running", "已有一项 OCR 正在进行", 45);
        return;
    }

    let result = async {
        let bytes = tauri::async_runtime::spawn_blocking(capture_selection_png)
            .await
            .map_err(|error| format!("Screenshot task failed: {error}"))??;
        let Some(bytes) = bytes else {
            return Ok::<Option<()>, String>(None);
        };

        emit_hud(&app, "running", "正在检查 OCR 提供器…", 22);
        let ocr = app
            .try_state::<OcrState>()
            .ok_or_else(|| "OCR runtime is unavailable".to_string())?
            .inner()
            .clone();
        let active_provider = ocr.active_provider(&app)?;
        let requested = current_model(&state);
        let model = if active_provider == crate::ocr_provider::LOCAL_PROVIDER {
            let runtime = ocr.runtime_status(app.clone(), false).await?;
            if !runtime.installed {
                return Err("请先在 VisualTeX 中安装 OCR 运行环境".to_string());
            }
            if runtime.installed_models.iter().any(|item| item == &requested) {
                requested
            } else if runtime
                .installed_models
                .iter()
                .any(|item| item == &runtime.default_model)
            {
                runtime.default_model.clone()
            } else {
                runtime
                    .installed_models
                    .first()
                    .cloned()
                    .ok_or_else(|| "没有可用的 OCR 模型".to_string())?
            }
        } else {
            requested
        };

        emit_hud(
            &app,
            "running",
            if active_provider == crate::ocr_provider::LOCAL_PROVIDER {
                "正在使用本地模型识别公式…"
            } else {
                "正在通过已配置的 OCR API 识别公式…"
            },
            36,
        );
        let recognition = ocr
            .recognize(
                app.clone(),
                OcrImageRequest {
                    bytes,
                    extension: "png".to_string(),
                    model,
                },
            )
            .await?;
        let latex_lines = recognition
            .formulas
            .iter()
            .map(|formula| formula.latex.trim().to_string())
            .filter(|latex| !latex.is_empty())
            .collect::<Vec<_>>();
        if latex_lines.is_empty() {
            return Err("OCR 没有返回可用公式".to_string());
        }
        let formatted_latex = format_silent_ocr_latex(&latex_lines, &current_copy_format(&state));
        if formatted_latex.is_empty() {
            return Err("OCR 源码格式化结果为空".to_string());
        }

        emit_hud(&app, "running", "正在按当前源码格式复制 LaTeX…", 92);
        write_text_clipboard(&formatted_latex)?;
        emit_hud(&app, "success", "识别完成，LaTeX 已复制到剪贴板", 100);
        Ok(Some(()))
    }
    .await;

    state.silent_busy.store(false, Ordering::SeqCst);
    match result {
        Ok(Some(())) => schedule_hud_hide(app, Duration::from_millis(1500)),
        Ok(None) => {
            if let Some(window) = app.get_webview_window(QUICK_OCR_HUD_LABEL) {
                let _ = window.hide();
            }
        }
        Err(error) => {
            emit_hud(&app, "error", error, 100);
            schedule_hud_hide(app, Duration::from_millis(3200));
        }
    }
}

#[cfg(target_os = "macos")]
mod mac_hotkey {
    use super::*;
    use std::ffi::c_void;
    use std::ptr;

    type OSStatus = i32;
    type EventTargetRef = *mut c_void;
    type EventHandlerCallRef = *mut c_void;
    type EventRef = *mut c_void;
    type EventHandlerRef = *mut c_void;
    type EventHotKeyRef = *mut c_void;

    #[repr(C)]
    struct EventTypeSpec {
        event_class: u32,
        event_kind: u32,
    }

    #[repr(C)]
    struct EventHotKeyId {
        signature: u32,
        id: u32,
    }

    type EventHandlerProc = unsafe extern "C" fn(
        EventHandlerCallRef,
        EventRef,
        *mut c_void,
    ) -> OSStatus;

    #[link(name = "Carbon", kind = "framework")]
    unsafe extern "C" {
        fn GetApplicationEventTarget() -> EventTargetRef;
        fn InstallEventHandler(
            target: EventTargetRef,
            handler: Option<EventHandlerProc>,
            num_types: u32,
            types: *const EventTypeSpec,
            user_data: *mut c_void,
            handler_ref: *mut EventHandlerRef,
        ) -> OSStatus;
        fn RegisterEventHotKey(
            key_code: u32,
            modifiers: u32,
            hot_key_id: EventHotKeyId,
            target: EventTargetRef,
            options: u32,
            hot_key_ref: *mut EventHotKeyRef,
        ) -> OSStatus;
        fn UnregisterEventHotKey(hot_key_ref: EventHotKeyRef) -> OSStatus;
    }

    const EVENT_CLASS_KEYBOARD: u32 = u32::from_be_bytes(*b"keyb");
    const EVENT_HOT_KEY_PRESSED: u32 = 5;
    const CMD_KEY: u32 = 1 << 8;
    const SHIFT_KEY: u32 = 1 << 9;
    const KEY_CODE_O: u32 = 0x1f;
    const HOTKEY_SIGNATURE: u32 = u32::from_be_bytes(*b"VTOC");

    static HOTKEY_APP: OnceLock<AppHandle> = OnceLock::new();
    static HOTKEY_REF: Mutex<usize> = Mutex::new(0);
    static HANDLER_INSTALLED: AtomicBool = AtomicBool::new(false);

    unsafe extern "C" fn hotkey_handler(
        _call: EventHandlerCallRef,
        _event: EventRef,
        _user_data: *mut c_void,
    ) -> OSStatus {
        if let Some(app) = HOTKEY_APP.get() {
            let app = app.clone();
            tauri::async_runtime::spawn(async move {
                run_silent_ocr(app).await;
            });
        }
        0
    }

    pub(super) fn initialize(app: &AppHandle) -> Result<(), String> {
        let _ = HOTKEY_APP.set(app.clone());
        if HANDLER_INSTALLED.swap(true, Ordering::SeqCst) {
            return Ok(());
        }
        let spec = EventTypeSpec {
            event_class: EVENT_CLASS_KEYBOARD,
            event_kind: EVENT_HOT_KEY_PRESSED,
        };
        let mut handler_ref: EventHandlerRef = ptr::null_mut();
        let status = unsafe {
            InstallEventHandler(
                GetApplicationEventTarget(),
                Some(hotkey_handler),
                1,
                &spec,
                ptr::null_mut(),
                &mut handler_ref,
            )
        };
        if status != 0 {
            HANDLER_INSTALLED.store(false, Ordering::SeqCst);
            return Err(format!("Unable to install the macOS silent OCR hotkey handler: {status}"));
        }
        Ok(())
    }

    pub(super) fn set_registered(enabled: bool) -> Result<(), String> {
        let mut hotkey = HOTKEY_REF
            .lock()
            .map_err(|_| "Silent OCR hotkey state is unavailable".to_string())?;
        if enabled {
            if *hotkey != 0 {
                return Ok(());
            }
            let mut reference: EventHotKeyRef = ptr::null_mut();
            let status = unsafe {
                RegisterEventHotKey(
                    KEY_CODE_O,
                    CMD_KEY | SHIFT_KEY,
                    EventHotKeyId {
                        signature: HOTKEY_SIGNATURE,
                        id: 1,
                    },
                    GetApplicationEventTarget(),
                    0,
                    &mut reference,
                )
            };
            if status != 0 || reference.is_null() {
                return Err(format!("Unable to register ⌘⇧O for silent OCR: {status}"));
            }
            *hotkey = reference as usize;
        } else if *hotkey != 0 {
            let reference = *hotkey as EventHotKeyRef;
            let status = unsafe { UnregisterEventHotKey(reference) };
            if status != 0 {
                return Err(format!("Unable to unregister the silent OCR hotkey: {status}"));
            }
            *hotkey = 0;
        }
        Ok(())
    }
}

pub(crate) fn initialize(app: &AppHandle) -> Result<(), String> {
    ensure_hud_window(app)?;
    #[cfg(target_os = "macos")]
    mac_hotkey::initialize(app)?;
    Ok(())
}

#[tauri::command]
pub(crate) async fn capture_quick_ocr_screenshot(
    app: AppHandle,
) -> Result<Option<QuickOcrCapture>, String> {
    if let Some(window) = app.get_webview_window("main") {
        window
            .minimize()
            .map_err(|error| format!("Unable to minimize VisualTeX for quick OCR: {error}"))?;
    }
    tokio::time::sleep(Duration::from_millis(220)).await;
    let capture = tauri::async_runtime::spawn_blocking(capture_selection_png)
        .await
        .map_err(|error| format!("Screenshot task failed: {error}"))?;
    let reveal_result = crate::office::background::reveal_main_window(&app);
    if let Err(error) = reveal_result {
        return Err(format!("Unable to restore VisualTeX after the screenshot: {error}"));
    }
    Ok(capture?.map(|bytes| QuickOcrCapture {
        data_base64: BASE64_STANDARD.encode(bytes),
        extension: "png".to_string(),
    }))
}

#[tauri::command]
pub(crate) async fn wait_for_quick_ocr_system_screenshot(
    app: AppHandle,
) -> Result<Option<QuickOcrCapture>, String> {
    let baseline = create_system_screenshot_baseline();
    if let Some(window) = app.get_webview_window("main") {
        window
            .minimize()
            .map_err(|error| format!("Unable to minimize VisualTeX while waiting for a system screenshot: {error}"))?;
    }
    tokio::time::sleep(Duration::from_millis(180)).await;
    let capture = tauri::async_runtime::spawn_blocking(move || wait_for_next_system_screenshot(baseline))
        .await
        .map_err(|error| format!("System screenshot watcher failed: {error}"))?;
    let reveal_result = crate::office::background::reveal_main_window(&app);
    if let Err(error) = reveal_result {
        return Err(format!("Unable to restore VisualTeX after waiting for the system screenshot: {error}"));
    }
    Ok(capture?.map(|(bytes, extension)| QuickOcrCapture {
        data_base64: BASE64_STANDARD.encode(bytes),
        extension,
    }))
}

#[cfg(target_os = "macos")]
async fn set_hotkey_registered_on_main(app: &AppHandle, enabled: bool) -> Result<(), String> {
    let (sender, receiver) = tokio::sync::oneshot::channel();
    let main_app = app.clone();
    app.run_on_main_thread(move || {
        let result = if enabled {
            ensure_hud_window(&main_app)
                .and_then(|_| mac_hotkey::initialize(&main_app))
                .and_then(|_| mac_hotkey::set_registered(true))
        } else {
            mac_hotkey::set_registered(false)
        };
        let _ = sender.send(result);
    })
    .map_err(|error| format!("Unable to schedule the silent OCR hotkey update: {error}"))?;
    receiver
        .await
        .map_err(|_| "Silent OCR hotkey update was interrupted".to_string())?
}

#[tauri::command]
pub(crate) async fn configure_silent_ocr(
    app: AppHandle,
    state: tauri::State<'_, QuickOcrState>,
    enabled: bool,
    model: String,
    copy_format: String,
) -> Result<(), String> {
    let normalized_model = model.trim();
    if !ALLOWED_MODELS.iter().any(|allowed| allowed == &normalized_model) {
        return Err("Unsupported silent OCR model".to_string());
    }
    *state
        .model
        .lock()
        .map_err(|_| "Silent OCR model state is unavailable".to_string())? = normalized_model.to_string();
    let normalized_copy_format = copy_format.trim();
    if !ALLOWED_SILENT_OCR_COPY_FORMATS
        .iter()
        .any(|allowed| allowed == &normalized_copy_format)
    {
        return Err("Unsupported silent OCR LaTeX copy format".to_string());
    }
    *state
        .copy_format
        .lock()
        .map_err(|_| "Silent OCR copy-format state is unavailable".to_string())? =
        normalized_copy_format.to_string();

    #[cfg(target_os = "macos")]
    set_hotkey_registered_on_main(&app, enabled).await?;
    #[cfg(not(target_os = "macos"))]
    if enabled {
        return Err("Silent OCR global capture is currently supported on macOS only".to_string());
    }

    state.silent_enabled.store(enabled, Ordering::SeqCst);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::format_silent_ocr_latex;

    #[test]
    fn silent_ocr_copy_format_matches_visualtex_source_wrappers() {
        let lines = vec!["x=y".to_string(), "a>b".to_string()];
        assert_eq!(
            format_silent_ocr_latex(&lines, "raw"),
            "x=y\na>b"
        );
        assert_eq!(
            format_silent_ocr_latex(&lines, "inline-dollar"),
            "$x=y$\n$a>b$"
        );
        assert_eq!(
            format_silent_ocr_latex(&lines, "display-dollar"),
            "$$\nx=y\n$$\n\n$$\na>b\n$$"
        );
        assert_eq!(
            format_silent_ocr_latex(&lines, "display-bracket"),
            "\\[\nx=y\n\\]\n\n\\[\na>b\n\\]"
        );
        assert_eq!(
            format_silent_ocr_latex(&lines, "equation-star"),
            "\\begin{equation*}\nx=y\n\\end{equation*}\n\n\\begin{equation*}\na>b\n\\end{equation*}"
        );
        assert_eq!(
            format_silent_ocr_latex(&lines, "align"),
            "\\begin{align}\nx&=y \\\\\na&>b\n\\end{align}"
        );
        assert_eq!(
            format_silent_ocr_latex(&["\\text{速度}v=t".to_string()], "inline-text-double-dollar"),
            "速度$$v=t$$"
        );
    }
}
