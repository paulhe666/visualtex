//! Settings > Keyboard shortcuts > Formula input method, "Switch to English input in formulas": a formula the
//! shortcut opens in VisualTeX's own document is typed in English. The window that has the keyboard (the page's, in
//! the WebView2's process) gets its input method's English mode, and the mode it had back when the formula ends.
//! (In another application VisualTeX's own input method takes the formula's keys: nothing is switched there.)

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Mutex, OnceLock};
use tauri::{AppHandle, Manager};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    GetForegroundWindow, GetGUIThreadInfo, GetWindowThreadProcessId, SendMessageTimeoutW, GUITHREADINFO, SMTO_ABORTIFHUNG,
};

/// Saved beside the shortcuts (hotkey.rs); on unless turned off.
const SETTING_FILE: &str = "anywhere-english-input.json";
/// The input method window's commands (IMM32; Chinese, Japanese and Korean input methods answer them, TSF ones too).
const WM_IME_CONTROL: u32 = 0x0283;
const IMC_GETCONVERSIONMODE: usize = 0x0001;
const IMC_SETCONVERSIONMODE: usize = 0x0002;
const IMC_GETOPENSTATUS: usize = 0x0005;
/// The conversion mode's own-language bit (Chinese, kana, Hangul): without it the input method types Latin letters.
const IME_CMODE_NATIVE: usize = 0x0001;
/// An input method that has not answered within this is left as it is.
const ANSWER_MS: u32 = 300;

#[link(name = "imm32")]
extern "system" {
    fn ImmGetDefaultIMEWnd(window: HWND) -> HWND;
}

static SETTING: OnceLock<AtomicBool> = OnceLock::new();

fn setting_file(app: &AppHandle) -> Option<std::path::PathBuf> {
    app.path().app_data_dir().ok().map(|folder| folder.join(SETTING_FILE))
}

fn setting(app: &AppHandle) -> &'static AtomicBool {
    SETTING.get_or_init(|| {
        let saved = setting_file(app).and_then(|path| std::fs::read_to_string(path).ok());
        AtomicBool::new(saved.map(|text| text.trim() != "false").unwrap_or(true))
    })
}

pub(super) fn enabled(app: &AppHandle) -> bool {
    setting(app).load(Ordering::SeqCst)
}

pub(super) fn set_enabled(app: &AppHandle, on: bool) -> Result<(), String> {
    setting(app).store(on, Ordering::SeqCst);
    let path = setting_file(app).ok_or_else(|| "No settings folder".to_string())?;
    if let Some(folder) = path.parent() {
        std::fs::create_dir_all(folder).map_err(|error| format!("The setting could not be saved: {error}"))?;
    }
    std::fs::write(&path, if on { "true" } else { "false" }).map_err(|error| format!("The setting could not be saved: {error}"))
}

/// The input method window switched to English, and the conversion mode it had.
struct Switched {
    ime_window: isize,
    mode: usize,
}

static SWITCHED: Mutex<Option<Switched>> = Mutex::new(None);

fn control(window: HWND, command: usize, value: usize) -> Option<usize> {
    let mut answer = 0usize;
    let sent = unsafe {
        SendMessageTimeoutW(window, WM_IME_CONTROL, WPARAM(command), LPARAM(value as isize), SMTO_ABORTIFHUNG, ANSWER_MS, Some(&mut answer))
    };
    (sent.0 != 0).then_some(answer)
}

/// English for the window that has the keyboard, when the setting is on and its input method types its own language
/// now: true when it was switched (or still is, for an earlier formula).
pub(super) fn select(app: &AppHandle) -> bool {
    if !enabled(app) {
        return false;
    }
    let Ok(mut switched) = SWITCHED.lock() else { return false };
    if switched.is_some() {
        return true;
    }
    let foreground = unsafe { GetForegroundWindow() };
    let thread = unsafe { GetWindowThreadProcessId(foreground, None) };
    let mut info = GUITHREADINFO { cbSize: std::mem::size_of::<GUITHREADINFO>() as u32, ..Default::default() };
    let focus = if thread != 0 && unsafe { GetGUIThreadInfo(thread, &mut info) }.is_ok() && !info.hwndFocus.is_invalid() {
        info.hwndFocus
    } else {
        foreground
    };
    let ime_window = unsafe { ImmGetDefaultIMEWnd(focus) };
    if ime_window.is_invalid() {
        super::log("english input: the window has no input method window");
        return false;
    }
    let (Some(open), Some(mode)) = (control(ime_window, IMC_GETOPENSTATUS, 0), control(ime_window, IMC_GETCONVERSIONMODE, 0)) else {
        super::log("english input: the input method did not answer");
        return false;
    };
    if open == 0 || mode & IME_CMODE_NATIVE == 0 {
        super::log(&format!("english input: already Latin letters (open={open} mode={mode:#x})"));
        return false;
    }
    // Latin letters, half-width, the Latin punctuation (the mode's other bits are its own language's)
    if control(ime_window, IMC_SETCONVERSIONMODE, 0) != Some(0) {
        super::log(&format!("english input: the input method refused (mode={mode:#x})"));
        return false;
    }
    let now = control(ime_window, IMC_GETCONVERSIONMODE, 0);
    super::log(&format!("english input: mode {mode:#x} -> {now:?}"));
    *switched = Some(Switched { ime_window: ime_window.0 as isize, mode });
    true
}

/// The input method window switched by `select` gets its mode back.
pub(super) fn restore() {
    let Some(switched) = SWITCHED.lock().ok().and_then(|mut switched| switched.take()) else { return };
    let given = control(HWND(switched.ime_window as *mut _), IMC_SETCONVERSIONMODE, switched.mode);
    super::log(&format!("english input: mode {:#x} back ({})", switched.mode, if given == Some(0) { "ok" } else { "refused" }));
}
