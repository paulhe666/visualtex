#![cfg(target_os = "windows")]

use crate::{OcrImageRequest, OcrState};
use serde::{Deserialize, Serialize};
use std::ffi::c_void;
use std::fs;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::mpsc::{self, Receiver, Sender};
use std::sync::{Mutex, OnceLock};
use std::thread;
use std::time::Duration;
use tauri::utils::config::BackgroundThrottlingPolicy;
use tauri::{AppHandle, Emitter, Manager, WebviewUrl, WebviewWindowBuilder};

const HOTKEY_ID: i32 = 0x5654; // "VT", within RegisterHotKey's application ID range
const MOD_ALT: u32 = 0x0001;
const MOD_CONTROL: u32 = 0x0002;
const MOD_SHIFT: u32 = 0x0004;
const MOD_NOREPEAT: u32 = 0x4000;
const VK_O: u32 = 0x4f;
const SILENT_OCR_HOTKEY_CANDIDATES: &[(u32, &str)] = &[
    (MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, "Ctrl+Alt+O"),
    (
        MOD_CONTROL | MOD_ALT | MOD_SHIFT | MOD_NOREPEAT,
        "Ctrl+Alt+Shift+O",
    ),
];
const WM_HOTKEY: u32 = 0x0312;
const PM_REMOVE: u32 = 0x0001;
const SILENT_OCR_CONFIG_FILE: &str = "silent-ocr.json";
const SILENT_OCR_HUD_LABEL: &str = "silent-ocr-hud";
const SILENT_OCR_HUD_EVENT: &str = "visualtex-silent-ocr-status";
const DEFAULT_SILENT_OCR_MODEL: &str = "PP-FormulaNet_plus-M";
const DEFAULT_SILENT_OCR_COPY_FORMAT: &str = "display-dollar";
const DEFAULT_SILENT_OCR_CAPTURE_MODE: &str = "windows";
const SILENT_OCR_CAPTURE_TIMEOUT_MS: u64 = 60_000;
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

#[repr(C)]
#[derive(Clone, Copy, Default)]
struct Point {
    x: i32,
    y: i32,
}

#[repr(C)]
#[derive(Clone, Copy)]
struct Msg {
    hwnd: *mut c_void,
    message: u32,
    w_param: usize,
    l_param: isize,
    time: u32,
    pt: Point,
    l_private: u32,
}

impl Default for Msg {
    fn default() -> Self {
        Self {
            hwnd: std::ptr::null_mut(),
            message: 0,
            w_param: 0,
            l_param: 0,
            time: 0,
            pt: Point::default(),
            l_private: 0,
        }
    }
}

#[link(name = "user32")]
unsafe extern "system" {
    fn RegisterHotKey(hwnd: *mut c_void, id: i32, modifiers: u32, virtual_key: u32) -> i32;
    fn UnregisterHotKey(hwnd: *mut c_void, id: i32) -> i32;
    fn PeekMessageW(
        message: *mut Msg,
        hwnd: *mut c_void,
        min_filter: u32,
        max_filter: u32,
        remove_message: u32,
    ) -> i32;
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SilentOcrConfiguration {
    enabled: bool,
    model: String,
    copy_format: String,
    #[serde(default = "default_silent_ocr_capture_mode")]
    capture_mode: String,
}

fn default_silent_ocr_capture_mode() -> String {
    DEFAULT_SILENT_OCR_CAPTURE_MODE.to_string()
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SilentOcrHudPayload {
    status: String,
    message: String,
    progress: u8,
}

impl Default for SilentOcrConfiguration {
    fn default() -> Self {
        Self {
            enabled: false,
            model: DEFAULT_SILENT_OCR_MODEL.to_string(),
            copy_format: DEFAULT_SILENT_OCR_COPY_FORMAT.to_string(),
            capture_mode: default_silent_ocr_capture_mode(),
        }
    }
}

enum HotkeyCommand {
    SetRegistered {
        enabled: bool,
        reply: Sender<Result<String, String>>,
    },
}

static COMMAND_SENDER: OnceLock<Sender<HotkeyCommand>> = OnceLock::new();
static CONFIGURATION: OnceLock<Mutex<SilentOcrConfiguration>> = OnceLock::new();
static HUD_STATUS: OnceLock<Mutex<SilentOcrHudPayload>> = OnceLock::new();
static SILENT_OCR_BUSY: AtomicBool = AtomicBool::new(false);
static HUD_GENERATION: AtomicU64 = AtomicU64::new(0);

fn configuration_store() -> &'static Mutex<SilentOcrConfiguration> {
    CONFIGURATION.get_or_init(|| Mutex::new(SilentOcrConfiguration::default()))
}

fn hud_status_store() -> &'static Mutex<SilentOcrHudPayload> {
    HUD_STATUS.get_or_init(|| {
        Mutex::new(SilentOcrHudPayload {
            status: "running".to_string(),
            message: "正在准备静默 OCR…".to_string(),
            progress: 8,
        })
    })
}

#[tauri::command]
pub(crate) fn get_silent_ocr_hud_status() -> SilentOcrHudPayload {
    hud_status_store()
        .lock()
        .map(|payload| payload.clone())
        .unwrap_or_else(|_| SilentOcrHudPayload {
            status: "running".to_string(),
            message: "正在准备静默 OCR…".to_string(),
            progress: 8,
        })
}

fn configuration_path(app: &AppHandle) -> Result<PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX app-data directory: {error}"))?;
    Ok(app_data.join("ocr").join(SILENT_OCR_CONFIG_FILE))
}

fn load_configuration(app: &AppHandle) -> SilentOcrConfiguration {
    let Ok(path) = configuration_path(app) else {
        return SilentOcrConfiguration::default();
    };
    let Ok(bytes) = fs::read(path) else {
        return SilentOcrConfiguration::default();
    };
    let Ok(mut configuration) = serde_json::from_slice::<SilentOcrConfiguration>(&bytes) else {
        return SilentOcrConfiguration::default();
    };
    if !crate::ALLOWED_MODELS.contains(&configuration.model.as_str()) {
        configuration.model = DEFAULT_SILENT_OCR_MODEL.to_string();
    }
    if !ALLOWED_SILENT_OCR_COPY_FORMATS.contains(&configuration.copy_format.as_str()) {
        configuration.copy_format = DEFAULT_SILENT_OCR_COPY_FORMAT.to_string();
    }
    configuration.capture_mode = crate::windows_quick_ocr::normalize_capture_mode(
        &configuration.capture_mode,
    )
    .unwrap_or(DEFAULT_SILENT_OCR_CAPTURE_MODE)
    .to_string();
    configuration
}

fn persist_configuration(app: &AppHandle, configuration: &SilentOcrConfiguration) {
    let Ok(path) = configuration_path(app) else {
        return;
    };
    let Some(parent) = path.parent() else {
        return;
    };
    if let Err(error) = fs::create_dir_all(parent) {
        eprintln!("Unable to create the silent OCR configuration directory: {error}");
        return;
    }
    match serde_json::to_vec_pretty(configuration) {
        Ok(bytes) => {
            if let Err(error) = fs::write(path, bytes) {
                eprintln!("Unable to persist the silent OCR configuration: {error}");
            }
        }
        Err(error) => eprintln!("Unable to serialize the silent OCR configuration: {error}"),
    }
}

fn current_configuration() -> SilentOcrConfiguration {
    configuration_store()
        .lock()
        .map(|configuration| configuration.clone())
        .unwrap_or_default()
}

fn ensure_hud_window(app: &AppHandle) -> Result<(), String> {
    if app.get_webview_window(SILENT_OCR_HUD_LABEL).is_some() {
        return Ok(());
    }
    let window = WebviewWindowBuilder::new(
        app,
        SILENT_OCR_HUD_LABEL,
        WebviewUrl::App("index.html?view=silent-ocr-hud".into()),
    )
    .title("VisualTeX OCR")
    .inner_size(380.0, 104.0)
    .center()
    .resizable(false)
    .decorations(false)
    .transparent(true)
    .shadow(false)
    .always_on_top(true)
    .focused(false)
    .focusable(false)
    .skip_taskbar(true)
    .visible(false)
    .background_throttling(BackgroundThrottlingPolicy::Disabled)
    .build()
    .map_err(|error| format!("Unable to initialize the silent OCR status window: {error}"))?;
    // WebView2 keeps its own opaque default background even when the native
    // window is transparent. Force both the window/webview background to full
    // alpha zero so no pale rectangle can remain behind the HUD card.
    window
        .set_background_color(Some(tauri::window::Color(0, 0, 0, 0)))
        .map_err(|error| format!("Unable to clear the silent OCR status background: {error}"))?;
    Ok(())
}

fn emit_hud(app: &AppHandle, status: &'static str, message: impl Into<String>, progress: u8) {
    let payload = SilentOcrHudPayload {
        status: status.to_string(),
        message: message.into(),
        progress: progress.min(100),
    };
    if let Ok(mut current) = hud_status_store().lock() {
        *current = payload.clone();
    }
    HUD_GENERATION.fetch_add(1, Ordering::SeqCst);
    if ensure_hud_window(app).is_err() {
        return;
    }
    let Some(window) = app.get_webview_window(SILENT_OCR_HUD_LABEL) else {
        return;
    };
    let _ = window.emit(SILENT_OCR_HUD_EVENT, payload);
    let _ = window.show();
}

fn silent_ocr_progress_percent(stage: &str) -> u8 {
    match stage {
        "api-submit" => 28,
        "api-queued" => 48,
        "api-inference" => 76,
        "api-result" => 92,
        "preprocess" => 38,
        "model" => 58,
        "inference" => 76,
        _ => 48,
    }
}

pub(crate) fn handle_ocr_progress(app: &AppHandle, response: &serde_json::Value) {
    if !SILENT_OCR_BUSY.load(Ordering::SeqCst) {
        return;
    }
    let stage = response
        .get("stage")
        .and_then(serde_json::Value::as_str)
        .unwrap_or_default();
    let message = response
        .get("message")
        .and_then(serde_json::Value::as_str)
        .unwrap_or("正在识别公式…");
    emit_hud(app, "running", message, silent_ocr_progress_percent(stage));
}

fn hide_hud(app: &AppHandle) {
    if let Some(window) = app.get_webview_window(SILENT_OCR_HUD_LABEL) {
        let _ = window.hide();
    }
}

fn schedule_hud_hide(app: AppHandle, delay: Duration) {
    let generation = HUD_GENERATION.load(Ordering::SeqCst);
    tauri::async_runtime::spawn(async move {
        tokio::time::sleep(delay).await;
        if HUD_GENERATION.load(Ordering::SeqCst) == generation {
            hide_hud(&app);
        }
    });
}

fn register_hotkey() -> Result<&'static str, String> {
    let mut failures = Vec::new();
    for (modifiers, shortcut) in SILENT_OCR_HOTKEY_CANDIDATES {
        let result = unsafe {
            RegisterHotKey(
                std::ptr::null_mut(),
                HOTKEY_ID,
                *modifiers,
                VK_O,
            )
        };
        if result != 0 {
            return Ok(*shortcut);
        }
        failures.push(format!("{shortcut}: {}", std::io::Error::last_os_error()));
    }
    Err(format!(
        "无法注册静默 OCR 全局快捷键；Ctrl+Alt+O 可能被其他软件占用，备用快捷键也不可用。{}",
        if failures.is_empty() {
            String::new()
        } else {
            format!(" ({})", failures.join("; "))
        }
    ))
}

fn unregister_hotkey() -> Result<(), String> {
    let result = unsafe { UnregisterHotKey(std::ptr::null_mut(), HOTKEY_ID) };
    if result != 0 {
        return Ok(());
    }
    let error = std::io::Error::last_os_error();
    if error.raw_os_error() == Some(1419) {
        return Ok(());
    }
    Err(format!("Unable to unregister the silent OCR hotkey: {error}"))
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

async fn run_silent_ocr(app: AppHandle, configuration: SilentOcrConfiguration) -> Result<(), String> {
    let capture_mode = configuration.capture_mode.clone();
    let capture = tauri::async_runtime::spawn_blocking(move || {
        crate::windows_quick_ocr::capture_windows_quick_ocr_bytes(
            &capture_mode,
            SILENT_OCR_CAPTURE_TIMEOUT_MS,
        )
    })
    .await
    .map_err(|error| format!("Silent OCR screenshot task failed: {error}"))??;
    let Some(bytes) = capture else {
        hide_hud(&app);
        return Ok(());
    };

    emit_hud(&app, "running", "正在检查 OCR 提供器…", 22);
    let ocr = app
        .try_state::<OcrState>()
        .ok_or_else(|| "OCR runtime is unavailable".to_string())?
        .inner()
        .clone();
    let active_provider = ocr.active_provider(&app)?;
    let requested = configuration.model;
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
    let formatted_latex = format_silent_ocr_latex(&latex_lines, &configuration.copy_format);
    if formatted_latex.is_empty() {
        return Err("OCR 源码格式化结果为空".to_string());
    }

    emit_hud(&app, "running", "正在按当前源码格式复制 LaTeX…", 92);
    tauri::async_runtime::spawn_blocking(move || {
        crate::windows_quick_ocr::write_clipboard_text(&formatted_latex)
    })
    .await
    .map_err(|error| format!("Silent OCR clipboard task failed: {error}"))??;
    emit_hud(&app, "success", "识别完成，LaTeX 已复制到剪贴板", 100);
    Ok(())
}

fn launch_silent_ocr(app: AppHandle) {
    let configuration = current_configuration();
    if !configuration.enabled {
        return;
    }
    if SILENT_OCR_BUSY.swap(true, Ordering::SeqCst) {
        emit_hud(&app, "running", "已有一项 OCR 正在进行", 45);
        schedule_hud_hide(app, Duration::from_millis(1400));
        return;
    }
    tauri::async_runtime::spawn(async move {
        let result = run_silent_ocr(app.clone(), configuration).await;
        SILENT_OCR_BUSY.store(false, Ordering::SeqCst);
        match result {
            Ok(()) => schedule_hud_hide(app, Duration::from_millis(2600)),
            Err(error) => {
                eprintln!("Windows silent OCR failed: {error}");
                emit_hud(&app, "error", error, 100);
                schedule_hud_hide(app, Duration::from_millis(3200));
            }
        }
    });
}

fn hotkey_thread(app: AppHandle, receiver: Receiver<HotkeyCommand>) {
    let mut registered_shortcut: Option<&'static str> = None;
    loop {
        match receiver.recv_timeout(Duration::from_millis(12)) {
            Ok(HotkeyCommand::SetRegistered { enabled, reply }) => {
                let result = if enabled == registered_shortcut.is_some() {
                    Ok(registered_shortcut.unwrap_or_default().to_string())
                } else if enabled {
                    register_hotkey().map(|shortcut| {
                        registered_shortcut = Some(shortcut);
                        shortcut.to_string()
                    })
                } else {
                    unregister_hotkey().map(|_| {
                        registered_shortcut = None;
                        String::new()
                    })
                };
                let _ = reply.send(result);
            }
            Err(mpsc::RecvTimeoutError::Timeout) => {}
            Err(mpsc::RecvTimeoutError::Disconnected) => {
                if registered_shortcut.is_some() {
                    let _ = unregister_hotkey();
                }
                return;
            }
        }

        let mut message = Msg::default();
        while unsafe {
            PeekMessageW(
                &mut message,
                std::ptr::null_mut(),
                0,
                0,
                PM_REMOVE,
            )
        } != 0
        {
            if message.message == WM_HOTKEY && message.w_param == HOTKEY_ID as usize {
                launch_silent_ocr(app.clone());
            }
        }
    }
}

pub(crate) fn initialize(app: &AppHandle) -> Result<(), String> {
    if COMMAND_SENDER.get().is_some() {
        return Ok(());
    }
    let loaded_configuration = load_configuration(app);
    {
        let mut configuration = configuration_store()
            .lock()
            .map_err(|_| "Silent OCR configuration state is unavailable".to_string())?;
        *configuration = loaded_configuration.clone();
    }

    let (sender, receiver) = mpsc::channel();
    COMMAND_SENDER
        .set(sender)
        .map_err(|_| "Silent OCR hotkey bridge is already initialized".to_string())?;
    let hotkey_app = app.clone();
    thread::Builder::new()
        .name("visualtex-silent-ocr-hotkey".to_string())
        .spawn(move || hotkey_thread(hotkey_app, receiver))
        .map_err(|error| format!("Unable to start the silent OCR hotkey bridge: {error}"))?;

    if loaded_configuration.enabled {
        if let Err(error) = set_registered(true) {
            eprintln!("Unable to restore the silent OCR hotkey at startup: {error}");
        }
    }
    Ok(())
}

fn set_registered(enabled: bool) -> Result<String, String> {
    let sender = COMMAND_SENDER
        .get()
        .ok_or_else(|| "Silent OCR hotkey bridge is unavailable".to_string())?;
    let (reply_sender, reply_receiver) = mpsc::channel();
    sender
        .send(HotkeyCommand::SetRegistered {
            enabled,
            reply: reply_sender,
        })
        .map_err(|_| "Silent OCR hotkey bridge has stopped".to_string())?;
    reply_receiver
        .recv_timeout(Duration::from_secs(2))
        .map_err(|_| "Silent OCR hotkey update timed out".to_string())?
}

pub(crate) fn configure(
    app: &AppHandle,
    enabled: bool,
    model: &str,
    copy_format: &str,
    capture_mode: &str,
) -> Result<String, String> {
    let normalized_model = model.trim();
    if !crate::ALLOWED_MODELS.contains(&normalized_model) {
        return Err("Unsupported silent OCR model".to_string());
    }
    let normalized_copy_format = copy_format.trim();
    if !ALLOWED_SILENT_OCR_COPY_FORMATS.contains(&normalized_copy_format) {
        return Err("Unsupported silent OCR LaTeX copy format".to_string());
    }
    let normalized_capture_mode =
        crate::windows_quick_ocr::normalize_capture_mode(capture_mode)?;

    let registered_shortcut = set_registered(enabled)?;
    let next = SilentOcrConfiguration {
        enabled,
        model: normalized_model.to_string(),
        copy_format: normalized_copy_format.to_string(),
        capture_mode: normalized_capture_mode.to_string(),
    };
    {
        let mut configuration = configuration_store()
            .lock()
            .map_err(|_| "Silent OCR configuration state is unavailable".to_string())?;
        *configuration = next.clone();
    }
    persist_configuration(app, &next);
    Ok(registered_shortcut)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn silent_ocr_copy_format_matches_visualtex_source_wrappers() {
        let lines = vec!["x=y".to_string(), "a>b".to_string()];
        assert_eq!(format_silent_ocr_latex(&lines, "raw"), "x=y\na>b");
        assert_eq!(
            format_silent_ocr_latex(&lines, "inline-dollar"),
            "$x=y$\n$a>b$"
        );
        assert_eq!(
            format_silent_ocr_latex(&lines, "display-bracket"),
            "\\[\nx=y\n\\]\n\n\\[\na>b\n\\]"
        );
        assert_eq!(
            format_silent_ocr_latex(&lines, "align"),
            "\\begin{align}\nx&=y \\\\\na&>b\n\\end{align}"
        );
    }

    #[test]
    fn inline_text_double_dollar_preserves_text_outside_math() {
        assert_eq!(
            format_silent_ocr_latex(
                &["\\text{速度}v=t".to_string()],
                "inline-text-double-dollar"
            ),
            "速度$$v=t$$"
        );
    }

    #[test]
    fn silent_ocr_hud_maps_local_and_remote_recognition_stages() {
        assert_eq!(silent_ocr_progress_percent("api-submit"), 28);
        assert_eq!(silent_ocr_progress_percent("api-queued"), 48);
        assert_eq!(silent_ocr_progress_percent("api-inference"), 76);
        assert_eq!(silent_ocr_progress_percent("api-result"), 92);
        assert_eq!(silent_ocr_progress_percent("preprocess"), 38);
        assert_eq!(silent_ocr_progress_percent("model"), 58);
        assert_eq!(silent_ocr_progress_percent("inference"), 76);
    }
}
