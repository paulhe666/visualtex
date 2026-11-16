//! VisualTeX Anywhere on Windows: a global shortcut turns the field that has the keyboard, in
//! any application, into a formula input, through VisualTeX's own input method
//! (apps/windows/input-method), as on the Mac (apps/macos/src-tauri/src/anywhere).
//!
//! VisualTeX's input method is none of the user's input methods (it is registered but disabled:
//! no input switcher offers it). The shortcut asks it, in the thread of the field that has the
//! keyboard, for a formula (ime::open_in_field): there it writes down the input method the field
//! has, takes its place for that thread alone and reports the caret; VisualTeX then opens the
//! formula: the field holds the formula's space (its composition), the TIP forwards every key to
//! this page, the formula is drawn in the field's line by a transparent window that never takes
//! the keyboard, and Enter replaces the space with the formula's text (or pastes its image). When
//! the formula ends, however it ends, the TIP gives the field the input method it had back.

mod deliver;
mod hotkey;
mod ime;
mod panel;
mod screen_colors;

use serde::{Deserialize, Serialize};
use std::sync::Mutex;
use std::time::Duration;
use tauri::{AppHandle, Emitter, Manager};

pub(crate) const WINDOW_LABEL: &str = "anywhere";
pub(crate) const PANEL_TITLE: &str = "VisualTeX Anywhere";
const ACTIVATE_EVENT: &str = "visualtex-anywhere-activate";
const SUSPEND_EVENT: &str = "visualtex-anywhere-suspend";
/// A key the input method took from the field, for the editor.
const KEY_EVENT: &str = "visualtex-anywhere-key";
/// The field's text size or colour, learnt after the formula opened.
const STYLE_EVENT: &str = "visualtex-anywhere-style";
/// A shortcut of the open formula, pressed (registered with Windows while a formula is open).
const SHORTCUT_EVENT: &str = "visualtex-anywhere-shortcut";
/// Where the field put the formula's pieces (it holds them in its text and wraps them as its own).
const PIECES_EVENT: &str = "visualtex-anywhere-pieces";
/// The most pieces a formula is cut into for the field.
const MAX_PIECES: usize = 512;
/// The shortcut asked the field's thread for a formula; a field that has not reported its caret
/// by then does not take text.
const INPUT_METHOD_WAIT: Duration = Duration::from_millis(1200);
/// The input method disables its profile when a formula ends; VisualTeX makes sure of it this long
/// after (the field's application may have gone away with the formula open).
const PROFILE_CHECK_DELAY: Duration = Duration::from_secs(3);
const MAX_PENDING_KEYS: usize = 256;
/// A page that has not answered an activation after this long is reloaded.
const PAGE_ANSWER_WAIT: Duration = Duration::from_millis(1500);
/// What a line of the field's text is to its font size (Chrome reports the line box).
const LINE_TO_FONT: f64 = 1.19;
/// What an application's own caret (the one it gave IMM32, "source": "imm", or the one it draws,
/// "pixels") is to its font size: the font's ascent and descent, 1.32 em in Windows' UI fonts
/// (Segoe UI, Microsoft YaHei).
const IMM_CARET_TO_FONT: f64 = 1.32;
/// Applications that report a made-up caret: one line of the default font under their real one
/// (WeChat, typing through IMM32, gives Windows nothing truer). Their real caret is read off the
/// screen whenever it shows and remembered by window (it is hidden while a formula is open).
const MADE_UP_CARET_CLIENTS: &[&str] = &["Weixin"];
/// The shortest believable line of a field's text, CSS pixels (a 9 px font's): a shorter rectangle
/// is not the text's (a hidden caret, a clipped or made-up rectangle) and does not size the formula.
const MIN_TEXT_HEIGHT: f64 = 11.0;
/// The tallest, CSS pixels.
const MAX_TEXT_HEIGHT: f64 = 96.0;
/// How far the anchor's height may be from the field's first report and still size the formula
/// (an empty line first reports its whole line box, taller than its text). Outside this the
/// anchor's rectangle is one the field made up (applications typing through IMM32, such as Qt's,
/// leave it to Windows), and only its place is used.
const ANCHOR_RATIO: std::ops::RangeInclusive<f64> = 0.6..=1.6;

/// A rectangle in physical screen pixels (top-left origin).
#[derive(Clone, Copy, Debug, Default, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ScreenRect {
    pub(crate) x: f64,
    pub(crate) y: f64,
    pub(crate) width: f64,
    pub(crate) height: f64,
}

/// The strip's size and the point of it that sits on the formula's start and baseline, in CSS
/// pixels from its top left.
#[derive(Clone, Copy, Debug, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PanelLayout {
    pub(crate) width: f64,
    pub(crate) height: f64,
    pub(crate) anchor_x: f64,
    pub(crate) anchor_y: f64,
}

impl PanelLayout {
    const EMPTY: Self = Self { width: 96.0, height: 44.0, anchor_x: 14.0, anchor_y: 22.0 };

    fn is_valid(&self) -> bool {
        let finite = [self.width, self.height, self.anchor_x, self.anchor_y].iter().all(|value| value.is_finite());
        finite
            && (16.0..=4096.0).contains(&self.width)
            && (16.0..=4096.0).contains(&self.height)
            && (0.0..=self.width).contains(&self.anchor_x)
            && (0.0..=self.height).contains(&self.anchor_y)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
enum AnchorSource {
    Caret,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct Activation {
    generation: u64,
    /// The caret's height in CSS pixels.
    caret_height: f64,
    anchor_source: AnchorSource,
    app_name: String,
    accessibility_trusted: bool,
    resume: bool,
    input_method: bool,
    /// The display shortcut opened the formula: its row is a display row (else a new formula
    /// starts inline).
    display: bool,
    font_size: Option<f64>,
    text_color: Option<String>,
    x_height: Option<f64>,
    line_height: Option<f64>,
    line_descent: Option<f64>,
    background_color: Option<String>,
}

/// How the field's text looks where the formula goes (CSS pixels).
#[derive(Clone, Debug, Default, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
struct FieldText {
    font_size: Option<f64>,
    text_color: Option<String>,
    x_height: Option<f64>,
    line_height: Option<f64>,
    line_descent: Option<f64>,
    background_color: Option<String>,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct Suspension {
    generation: u64,
    keep_draft: bool,
}

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum CommitOutcome {
    Inserted,
    Copied,
    Stale,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnywhereStatus {
    accessibility_trusted: bool,
    shortcut: String,
    input_method: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Phase {
    Hidden,
    Activating,
    Presented,
}

/// An application's real caret (seen on the screen) relative to the made-up one it reports, in
/// physical pixels, by its window.
#[derive(Clone, Copy, Debug, PartialEq)]
struct CaretCorrection {
    window: isize,
    /// The real caret's top minus the reported rectangle's top.
    offset: f64,
    height: f64,
    /// Where the reported rectangle was when the caret was seen.
    at: (f64, f64),
}

/// How the formula reaches the field: through VisualTeX's input method, or (a field without it)
/// typed in VisualTeX's own strip and pasted, as on macOS.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Delivery {
    InputMethod,
    Paste,
}

struct Session {
    generation: u64,
    app_name: String,
    /// The caret (physical pixels) and the formula's place in the line.
    caret: Option<ScreenRect>,
    inline_anchor: Option<ScreenRect>,
    /// The scale of the monitor the formula is on.
    scale: f64,
    layout: PanelLayout,
    ready: bool,
    phase: Phase,
    draft: bool,
    reloaded_for: u64,
    /// The display shortcut opened this session.
    display: bool,
    awaiting_input_method: u64,
    pending_keys: Vec<serde_json::Value>,
    field_text: FieldText,
    /// The height of the field's text (CSS pixels) that sized this formula first: the anchor's
    /// height is checked against it. None: the field told nothing believable (the system's text
    /// size is used until it does).
    reference_height: Option<f64>,
    color_read: u64,
    /// The field's window when the shortcut was pressed (the screenshot selection takes the
    /// foreground for a moment; the pasting strip gives it back).
    field_window: isize,
    delivery: Delivery,
    /// The screenshot selection has the keyboard for a moment: the pasting strip stays open.
    holding: bool,
    /// The real carets of applications that make theirs up, last seen (a few windows).
    corrections: Vec<CaretCorrection>,
}

static SESSION: Mutex<Session> = Mutex::new(Session {
    generation: 0,
    app_name: String::new(),
    caret: None,
    inline_anchor: None,
    scale: 1.0,
    layout: PanelLayout::EMPTY,
    ready: false,
    phase: Phase::Hidden,
    draft: false,
    reloaded_for: 0,
    display: false,
    awaiting_input_method: 0,
    pending_keys: Vec::new(),
    field_text: FieldText {
        font_size: None,
        text_color: None,
        x_height: None,
        line_height: None,
        line_descent: None,
        background_color: None,
    },
    reference_height: None,
    color_read: 0,
    field_window: 0,
    delivery: Delivery::InputMethod,
    holding: false,
    corrections: Vec::new(),
});

static LOG_PATH: std::sync::OnceLock<std::path::PathBuf> = std::sync::OnceLock::new();
const MAX_LOG_BYTES: u64 = 512 * 1024;

/// One line of the diagnostics log (app data, anywhere.log). Formulas are not logged.
pub(crate) fn log(line: &str) {
    use std::io::Write;
    let Some(path) = LOG_PATH.get() else { return };
    if std::fs::metadata(path).map(|metadata| metadata.len() > MAX_LOG_BYTES).unwrap_or(false) {
        let _ = std::fs::remove_file(path);
    }
    let millis = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|value| value.as_millis())
        .unwrap_or_default();
    if let Ok(mut file) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(file, "epochMs={millis} {line}");
    }
}

fn session() -> Result<std::sync::MutexGuard<'static, Session>, String> {
    SESSION.lock().map_err(|_| "VisualTeX Anywhere state is unavailable".to_string())
}

/// At startup: the shortcut, the input method's link, and the formula's window (created a moment
/// later, when the input method is installed).
pub(crate) fn initialize(app: &AppHandle) -> Result<(), String> {
    if let Ok(directory) = app.path().app_data_dir() {
        let _ = std::fs::create_dir_all(&directory);
        let _ = LOG_PATH.set(directory.join("anywhere.log"));
    }
    log("started");
    ime::start_link(app)?;
    if let Err(error) = hotkey::register_saved(app) {
        log(&format!("shortcut not registered: {error}"));
    }
    let app = app.clone();
    let _ = std::thread::Builder::new().name("visualtex-anywhere-start".to_string()).spawn(move || {
        if !ime::is_installed() {
            log("input method not installed");
            return;
        }
        // (a formula's field that went away with VisualTeX's last run left the profile enabled)
        ime::disable_profile();
        std::thread::sleep(Duration::from_secs(3));
        let prewarm = app.clone();
        let _ = app.run_on_main_thread(move || {
            if let Err(error) = panel::ensure(&prewarm) {
                log(&format!("formula window unavailable: {error}"));
            }
        });
    });
    Ok(())
}

fn with_panel(app: &AppHandle, work: impl FnOnce(&tauri::WebviewWindow) + Send + 'static) {
    let target = app.clone();
    let _ = app.run_on_main_thread(move || match panel::ensure(&target) {
        Ok(window) => work(&window),
        Err(error) => log(&format!("formula window unavailable: {error}")),
    });
}

/// A shortcut: open a formula in the field that has the keyboard (`display`: its row is a display
/// row), or close the open one (keeping it for the next press).
pub(super) fn toggle(app: &AppHandle, display: bool) {
    if session().map(|session| session.phase != Phase::Hidden).unwrap_or(false) {
        hide(app, true, "shortcut");
        return;
    }
    log(&format!("shortcut display={display}"));
    let app = app.clone();
    let _ = std::thread::Builder::new()
        .name("visualtex-anywhere-activate".to_string())
        .spawn(move || begin_input_method(&app, display));
}

fn begin_input_method(app: &AppHandle, display: bool) {
    let started = std::time::Instant::now();
    let field_window = unsafe { windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow() }.0 as isize;
    if !ime::is_installed() {
        log("shortcut: the input method is not installed: pasting strip");
        begin_paste(app, display, field_window);
        return;
    }
    let generation = {
        let Ok(mut session) = session() else { return };
        session.generation += 1;
        session.phase = Phase::Activating;
        session.caret = None;
        session.inline_anchor = None;
        session.display = display;
        session.awaiting_input_method = session.generation;
        session.pending_keys.clear();
        session.field_text = FieldText::default();
        session.reference_height = None;
        session.field_window = field_window;
        session.delivery = Delivery::InputMethod;
        session.generation
    };
    // a new formula: the field that has the keyboard answers (no earlier formula's field)
    ime::pin(None);
    with_panel(app, |_| {});
    formula_keys(true);
    // The input method, in the field's own thread, takes the field for this formula and reports
    // its caret ("activate"): nothing is switched from here.
    let asked = ime::open_in_field(field_window, generation);
    log(&format!(
        "input method asked gen={generation} window={field_window:x} {} ms={}",
        asked.as_ref().map(|_| "ok".to_string()).unwrap_or_else(|error| error.clone()),
        started.elapsed().as_millis()
    ));
    if asked.is_err() {
        give_up(app, generation);
        return;
    }
    let app = app.clone();
    let _ = std::thread::Builder::new().name("visualtex-ime-wait".to_string()).spawn(move || {
        std::thread::sleep(INPUT_METHOD_WAIT);
        let waiting = session()
            .map(|session| session.awaiting_input_method == generation && session.generation == generation)
            .unwrap_or(false);
        if waiting {
            log(&format!("input method did not take the field gen={generation}"));
            give_up(&app, generation);
        }
    });
}

/// The field did not take the input method (the input method gives up by itself and leaves the
/// field its own): the formula opens in the pasting strip instead, while that field still has the
/// keyboard.
fn give_up(app: &AppHandle, generation: u64) {
    let (display, field_window) = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation || session.awaiting_input_method != generation {
            return;
        }
        session.awaiting_input_method = 0;
        session.phase = Phase::Hidden;
        (session.display, session.field_window)
    };
    formula_keys(false);
    let foreground = unsafe { windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow() }.0 as isize;
    if field_window != 0 && foreground == field_window {
        begin_paste(app, display, field_window);
    }
}

/// A formula closed: a moment later (no other one open) the input method's profile is made sure to
/// be disabled.
fn check_profile_later(generation: u64) {
    let _ = std::thread::Builder::new().name("visualtex-ime-profile".to_string()).spawn(move || {
        std::thread::sleep(PROFILE_CHECK_DELAY);
        let idle = session().map(|session| session.generation == generation && session.phase == Phase::Hidden).unwrap_or(false);
        if idle {
            ime::disable_profile();
        }
    });
}

/// The caret of the field's window (physical pixels): its application's system caret, else the
/// focused text's caret through UI Automation (Chrome, Edge and Electron applications show no
/// system caret), else the mouse pointer's place, a line of the system's text tall. A caret that
/// is no line of text (a hidden one-pixel caret some applications keep for accessibility tools)
/// is passed over.
fn field_caret(field_window: isize) -> ScreenRect {
    let believable = |source: &str, caret: Option<ScreenRect>| -> Option<ScreenRect> {
        let caret = caret?;
        let (scale, _) = panel::monitor_at(caret.x, caret.y + caret.height / 2.0);
        let height = caret.height / scale;
        if (MIN_TEXT_HEIGHT..=MAX_TEXT_HEIGHT).contains(&height) {
            return Some(caret);
        }
        log(&format!("pasting strip: the {source} caret is {height:.1} CSS px tall, not a line of text: not used"));
        None
    };
    if let Some(caret) = believable("system", system_caret(field_window)).or_else(|| believable("UI Automation", ui_automation_caret())) {
        return caret;
    }
    use windows::Win32::Foundation::POINT;
    use windows::Win32::UI::WindowsAndMessaging::GetCursorPos;
    log("pasting strip: no caret found, placed at the pointer");
    let mut pointer = POINT::default();
    let _ = unsafe { GetCursorPos(&mut pointer) };
    let (scale, _) = panel::monitor_at(pointer.x as f64, pointer.y as f64);
    let height = system_text_size() * LINE_TO_FONT * scale;
    ScreenRect { x: pointer.x as f64, y: pointer.y as f64 - height / 2.0, width: 1.0, height }
}

/// The system's text size (its message font, as dialogs and menus use it) in CSS pixels: what the
/// formula is sized from when the field tells nothing believable.
fn system_text_size() -> f64 {
    use windows::Win32::UI::HiDpi::SystemParametersInfoForDpi;
    use windows::Win32::UI::WindowsAndMessaging::{NONCLIENTMETRICSW, SPI_GETNONCLIENTMETRICS};
    let mut metrics = NONCLIENTMETRICSW { cbSize: std::mem::size_of::<NONCLIENTMETRICSW>() as u32, ..Default::default() };
    let read = unsafe {
        SystemParametersInfoForDpi(SPI_GETNONCLIENTMETRICS.0, metrics.cbSize, Some(&mut metrics as *mut NONCLIENTMETRICSW as *mut _), 0, 96)
    };
    // at 96 DPI a CSS pixel is a pixel; a negative height is the font's size, a positive one its cell
    let height = metrics.lfMessageFont.lfHeight.unsigned_abs() as f64;
    if read.is_ok() && (8.0..=48.0).contains(&height) {
        height
    } else {
        12.0
    }
}

fn system_caret(field_window: isize) -> Option<ScreenRect> {
    use windows::Win32::Foundation::{HWND, POINT};
    use windows::Win32::Graphics::Gdi::ClientToScreen;
    use windows::Win32::UI::HiDpi::LogicalToPhysicalPointForPerMonitorDPI;
    use windows::Win32::UI::WindowsAndMessaging::{GetCursorPos, GetGUIThreadInfo, GetWindowThreadProcessId, GUITHREADINFO};
    let thread = unsafe { GetWindowThreadProcessId(HWND(field_window as *mut _), None) };
    let mut info = GUITHREADINFO { cbSize: std::mem::size_of::<GUITHREADINFO>() as u32, ..Default::default() };
    if thread != 0 && unsafe { GetGUIThreadInfo(thread, &mut info) }.is_ok() && !info.hwndCaret.is_invalid() {
        let caret = info.rcCaret;
        let mut top = POINT { x: caret.left, y: caret.top };
        let mut bottom = POINT { x: caret.left, y: caret.bottom };
        unsafe {
            let _ = ClientToScreen(info.hwndCaret, &mut top);
            let _ = ClientToScreen(info.hwndCaret, &mut bottom);
            let _ = LogicalToPhysicalPointForPerMonitorDPI(Some(info.hwndCaret), &mut top);
            let _ = LogicalToPhysicalPointForPerMonitorDPI(Some(info.hwndCaret), &mut bottom);
        }
        let height = (bottom.y - top.y) as f64;
        if height > 0.0 && height < 400.0 {
            return Some(ScreenRect { x: top.x as f64, y: top.y as f64, width: 1.0, height });
        }
    }
    None
}

/// The focused element's text caret through UI Automation (physical pixels, this process being
/// per-monitor DPI aware): the caret range's rectangle, else the character after it (its left
/// edge), else the one before it (its right edge, at the end of the text).
fn ui_automation_caret() -> Option<ScreenRect> {
    use windows::core::BOOL;
    use windows::Win32::System::Com::{CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED};
    use windows::Win32::UI::Accessibility::{
        CUIAutomation, IUIAutomation, IUIAutomationTextPattern, IUIAutomationTextPattern2, IUIAutomationTextRange,
        TextPatternRangeEndpoint_Start, TextUnit_Character, UIA_TextPattern2Id, UIA_TextPatternId,
    };
    unsafe {
        let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
        let automation: IUIAutomation = CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;
        let focused = automation.GetFocusedElement().ok()?;
        let caret: IUIAutomationTextRange = match focused.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id) {
            Ok(pattern) => {
                let mut active = BOOL(0);
                pattern.GetCaretRange(&mut active).ok()?
            }
            Err(_) => {
                let pattern = focused.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId).ok()?;
                pattern.GetSelection().ok()?.GetElement(0).ok()?
            }
        };
        let rect = |range: &IUIAutomationTextRange| -> Option<[f64; 4]> {
            let values = safe_array_doubles(range.GetBoundingRectangles().ok()?);
            (values.len() >= 4 && values[3] > 0.0).then(|| [values[0], values[1], values[2], values[3]])
        };
        let found = if let Some([left, top, _, height]) = rect(&caret) {
            Some((left, top, height))
        } else {
            let next = caret.Clone().ok()?;
            let _ = next.ExpandToEnclosingUnit(TextUnit_Character);
            if let Some([left, top, _, height]) = rect(&next) {
                Some((left, top, height))
            } else {
                let previous = caret.Clone().ok()?;
                let _ = previous.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, -1);
                rect(&previous).map(|[left, top, width, height]| (left + width, top, height))
            }
        };
        let (x, y, height) = found?;
        (height > 0.0 && height < 400.0).then(|| ScreenRect { x, y, width: 1.0, height })
    }
}

/// The doubles of a SAFEARRAY that UI Automation returned (and the array freed).
unsafe fn safe_array_doubles(array: *mut windows::Win32::System::Com::SAFEARRAY) -> Vec<f64> {
    use windows::Win32::System::Ole::{SafeArrayAccessData, SafeArrayDestroy, SafeArrayGetLBound, SafeArrayGetUBound, SafeArrayUnaccessData};
    if array.is_null() {
        return Vec::new();
    }
    let mut values = Vec::new();
    if let (Ok(lower), Ok(upper)) = (SafeArrayGetLBound(array, 1), SafeArrayGetUBound(array, 1)) {
        let mut data: *mut std::ffi::c_void = std::ptr::null_mut();
        if upper >= lower && SafeArrayAccessData(array, &mut data).is_ok() {
            let count = (upper - lower + 1) as usize;
            values.extend_from_slice(std::slice::from_raw_parts(data as *const f64, count));
            let _ = SafeArrayUnaccessData(array);
        }
    }
    let _ = SafeArrayDestroy(array);
    values
}

/// A field without the input method: the formula opens in VisualTeX's own strip on the caret's
/// line, which takes the keyboard; Enter pastes the formula into the field (as on macOS).
fn begin_paste(app: &AppHandle, display: bool, field_window: isize) {
    let caret = field_caret(field_window);
    let (scale, _) = panel::monitor_at(caret.x, caret.y + caret.height / 2.0);
    let (activation, ready) = {
        let Ok(mut session) = session() else { return };
        session.generation += 1;
        session.phase = Phase::Activating;
        session.delivery = Delivery::Paste;
        session.caret = Some(caret);
        session.inline_anchor = None;
        session.scale = scale;
        session.display = display;
        session.awaiting_input_method = 0;
        session.pending_keys.clear();
        session.field_text = FieldText::default();
        session.reference_height = None;
        session.field_window = field_window;
        session.app_name = String::new();
        (pending_activation(&session), session.ready)
    };
    log(&format!("pasting strip gen={} caret={:.0},{:.0},{:.0}", activation.as_ref().map(|a| a.generation).unwrap_or(0), caret.x, caret.y, caret.height));
    formula_keys(true);
    let target = app.clone();
    with_panel(app, move |_| {
        if let (Some(activation), true) = (activation, ready) {
            send_activation(&target, activation);
        }
    });
}

/// The pasting strip lost the keyboard (a click elsewhere): it closes, keeping the formula for the
/// next shortcut.
pub(super) fn handle_focus_lost(app: &AppHandle) {
    let open = session()
        .map(|session| session.phase == Phase::Presented && session.delivery == Delivery::Paste && !session.holding)
        .unwrap_or(false);
    if open {
        hide_with(app, true, "strip lost the keyboard", false);
    }
}

/// The input method's report of the field: where the formula goes and how its text looks there.
struct FieldGeometry {
    caret: ScreenRect,
    baseline: f64,
    scale: f64,
    /// How the field's text looks, when the report is believable (else all None).
    text: FieldText,
    /// The height of the field's text in CSS pixels, when believable: not clipped, and a line of
    /// text tall (MIN_TEXT_HEIGHT, MAX_TEXT_HEIGHT).
    measured: Option<f64>,
    /// The report, for anywhere.log.
    report: String,
}

fn input_method_geometry(message: &serde_json::Value) -> Option<FieldGeometry> {
    let values: Vec<f64> = message.get("caret")?.as_array()?.iter().filter_map(|value| value.as_f64()).collect();
    let &[x, y, width, height] = values.as_slice() else { return None };
    if !(x.is_finite() && y.is_finite() && height.is_finite()) || height <= 0.0 || height > 600.0 {
        return None;
    }
    if x.abs() > 40000.0 || y.abs() > 40000.0 {
        return None;
    }
    let (scale, _) = panel::monitor_at(x, y + height / 2.0);
    // An empty line reports its whole line box (Chrome, Electron): the height of a line of this
    // field's text, learnt by the input method where the line had text, is the text's there (as
    // on macOS); the text sits in the middle of the taller box.
    let text_before = message.get("textBefore").and_then(|value| value.as_bool()).unwrap_or(true);
    let learnt = message.get("textHeight").and_then(|value| value.as_f64());
    let text_height = match learnt {
        Some(learnt) if !text_before && learnt < height && height <= learnt * 1.32 => learnt,
        _ => height,
    };
    let clipped = message.get("clipped").and_then(|value| value.as_bool()).unwrap_or(false);
    let source = message.get("source").and_then(|value| value.as_str()).filter(|source| matches!(*source, "imm" | "pixels"));
    let imm = source.is_some();
    let css_height = text_height / scale;
    // A clipped rectangle, or one shorter or taller than a line of text, places the formula
    // without sizing it.
    let measured = (!clipped && (MIN_TEXT_HEIGHT..=MAX_TEXT_HEIGHT).contains(&css_height)).then_some(css_height);
    let descent = text_height * 0.22;
    let below_text = (height - text_height) / 2.0;
    let line_to_font = if imm { IMM_CARET_TO_FONT } else { LINE_TO_FONT };
    Some(FieldGeometry {
        caret: ScreenRect { x, y, width, height },
        baseline: y + height - below_text - descent,
        scale,
        text: measured.map(|line| text_of_line(line, line_to_font)).unwrap_or_default(),
        measured,
        report: format!(
            "caret={x:.0},{y:.0},{height:.0} scale={scale:.2} clipped={clipped} textBefore={text_before} learnt={} text={css_height:.1}px{}{}",
            learnt.map(|value| format!("{value:.0}")).unwrap_or_else(|| "-".to_string()),
            source.map(|source| format!(" source={source}")).unwrap_or_default(),
            message.get("madeUp").and_then(|value| value.as_str()).map(|made_up| format!(" ({made_up})")).unwrap_or_default(),
        ),
    })
}

/// How text whose line (its caret) is `line` CSS pixels tall looks, the line being `line_to_font`
/// times its font size: its font size, the line and how far the line reaches below the baseline.
fn text_of_line(line: f64, line_to_font: f64) -> FieldText {
    FieldText {
        font_size: Some((line / line_to_font).clamp(9.0, 64.0)),
        line_height: Some(line),
        line_descent: Some(line * 0.22),
        ..FieldText::default()
    }
}

/// Whether the anchor's height (CSS pixels, None: not believable) sizes the formula, given the
/// height that sized it first (None: the system's text size is in use).
fn anchor_sizes(measured: Option<f64>, reference: Option<f64>) -> bool {
    match (measured, reference) {
        (Some(height), Some(reference)) => ANCHOR_RATIO.contains(&(height / reference)),
        (Some(_), None) => true,
        (None, _) => false,
    }
}

fn describe_text(text: &FieldText) -> String {
    format!(
        "font={} line={}",
        text.font_size.map(|value| format!("{value:.1}")).unwrap_or_else(|| "-".to_string()),
        text.line_height.map(|value| format!("{value:.1}")).unwrap_or_else(|| "-".to_string())
    )
}

fn inline_anchor(geometry: &FieldGeometry) -> ScreenRect {
    ScreenRect { x: geometry.caret.x, y: geometry.baseline, width: geometry.caret.width, height: 0.0 }
}

fn pending_activation(session: &Session) -> Option<Activation> {
    let caret = session.caret?;
    Some(Activation {
        generation: session.generation,
        caret_height: caret.height / session.scale,
        anchor_source: AnchorSource::Caret,
        app_name: session.app_name.clone(),
        accessibility_trusted: true,
        resume: session.draft,
        input_method: session.delivery == Delivery::InputMethod,
        display: session.display,
        font_size: session.field_text.font_size,
        text_color: session.field_text.text_color.clone(),
        x_height: session.field_text.x_height,
        line_height: session.field_text.line_height,
        line_descent: session.field_text.line_descent,
        background_color: session.field_text.background_color.clone(),
    })
}

/// A message from the input method (its link threads).
pub(super) fn input_method_message(app: &AppHandle, message: &serde_json::Value) {
    let kind = message.get("t").and_then(|value| value.as_str()).unwrap_or("");
    // (a message about an earlier formula, from a field still closing it)
    if let Some(about) = message.get("gen").and_then(|value| value.as_u64()).filter(|about| *about != 0) {
        if session().map(|session| session.generation != about).unwrap_or(false) && kind != "hello" {
            return;
        }
    }
    match kind {
        "hello" => log(&format!(
            "input method {}",
            message.get("version").and_then(|value| value.as_str()).unwrap_or("?")
        )),
        "activate" | "geometry" => caret_reported(app, message),
        "key" => input_method_key(app, message),
        "pieces" => input_method_pieces(app, message),
        "deactivate" | "disconnected" => {
            let open = session()
                .map(|session| {
                    session.phase != Phase::Hidden && session.awaiting_input_method == 0 && session.delivery == Delivery::InputMethod
                })
                .unwrap_or(false);
            if open {
                hide_with(app, true, if kind == "deactivate" { "input method left the field" } else { "input method gone" }, false);
            }
        }
        "forced" => {
            if session().map(|session| session.phase != Phase::Hidden && session.delivery == Delivery::InputMethod).unwrap_or(false) {
                hide_with(app, false, "field kept the text", false);
            }
        }
        _ => {}
    }
}

/// Where the field put the formula's pieces (physical pixels from the first piece's start, y
/// down; with where the held space starts and ends): for the page, in CSS pixels.
fn input_method_pieces(app: &AppHandle, message: &serde_json::Value) {
    let Some((generation, scale)) = session()
        .ok()
        .filter(|session| session.phase != Phase::Hidden && session.delivery == Delivery::InputMethod)
        .map(|session| (session.generation, session.scale))
    else {
        return;
    };
    let scale = if scale.is_finite() && scale > 0.0 { scale } else { 1.0 };
    let pair = |value: &serde_json::Value| -> Option<[f64; 2]> {
        let pair = value.as_array()?;
        let (x, y) = (pair.first()?.as_f64()?, pair.get(1)?.as_f64()?);
        (x.is_finite() && y.is_finite() && x.abs() < 40000.0 && y.abs() < 40000.0).then_some([x / scale, y / scale])
    };
    let places: Vec<[f64; 2]> = message
        .get("places")
        .and_then(|value| value.as_array())
        .map(|places| places.iter().take(MAX_PIECES).map_while(pair).collect())
        .unwrap_or_default();
    let start = message.get("start").and_then(pair);
    let end = message.get("end").and_then(pair);
    // an end the field did not tell (the TIP's estimate)
    let estimated = message.get("end").and_then(|value| value.as_array()).map(|end| end.len() > 2).unwrap_or(false);
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        let _ = window.emit(
            PIECES_EVENT,
            serde_json::json!({ "generation": generation, "places": places, "start": start, "end": end, "endEstimated": estimated }),
        );
    }
}

/// While a formula is open its shortcuts are VisualTeX's: the silent OCR shortcut steps aside (it
/// may be the same chord) and the formula's shortcuts are registered with Windows; when it closes,
/// the other way round.
fn formula_keys(on: bool) {
    if on {
        crate::windows_silent_ocr_hotkey::suspend_for_formula(true);
        hotkey::formula_shortcuts(true);
    } else {
        hotkey::formula_shortcuts(false);
        crate::windows_silent_ocr_hotkey::suspend_for_formula(false);
    }
}

/// A shortcut of the open formula was pressed (through Windows): the page runs it.
pub(super) fn formula_shortcut(app: &AppHandle, action: &str) {
    let generation = session().ok().filter(|session| session.phase != Phase::Hidden).map(|session| session.generation);
    let Some(generation) = generation else { return };
    log(&format!("formula shortcut {action} gen={generation}"));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        let _ = window.emit(SHORTCUT_EVENT, serde_json::json!({ "generation": generation, "action": action }));
    }
}

/// A report of an application that makes its caret up (MADE_UP_CARET_CLIENTS): the same report with
/// its real caret, "source": "pixels" ("madeUp" says what it replaced, for anywhere.log). The caret
/// is looked for on the screen where it shows: a moment after the field got the keyboard (its
/// caret is drawn then), and when a formula opens (before its anchor hides the caret); it is
/// remembered by window for the reports where it cannot show (the formula's anchor). None: another
/// application, no formula opening, or its caret not seen yet.
fn with_real_caret(message: &serde_json::Value) -> Option<serde_json::Value> {
    let client = message.get("client").and_then(|value| value.as_str())?;
    if !MADE_UP_CARET_CLIENTS.iter().any(|name| name.eq_ignore_ascii_case(client)) {
        return None;
    }
    let values: Vec<f64> = message.get("caret")?.as_array()?.iter().filter_map(|value| value.as_f64()).collect();
    let &[x, y, width, height] = values.as_slice() else { return None };
    if !(x.is_finite() && y.is_finite() && height.is_finite()) {
        return None;
    }
    let rect = ScreenRect { x, y, width, height };
    let window = unsafe { windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow() }.0 as isize;
    let (phase, opening) = session()
        .map(|session| (session.phase, session.phase == Phase::Activating && session.awaiting_input_method == session.generation))
        .ok()?;
    if phase == Phase::Hidden {
        // learn only: the field has just been given the keyboard; its caret is drawn a frame later
        // and blinks (fading, it looks shorter): the tallest over a blink, a few looks apart
        let client = client.to_string();
        let _ = std::thread::Builder::new().name("visualtex-anywhere-caret".to_string()).spawn(move || {
            let mut tallest: Option<screen_colors::CaretBar> = None;
            for _ in 0..CARET_LOOKS {
                std::thread::sleep(Duration::from_millis(CARET_LOOK_MS));
                if let Some(bar) = screen_colors::caret_bar_above(rect) {
                    if tallest.map(|tallest| bar.height > tallest.height).unwrap_or(true) {
                        tallest = Some(bar);
                    }
                }
            }
            if let Some(bar) = tallest {
                remember_caret(&client, CaretCorrection { window, offset: bar.top - y, height: bar.height, at: (x, y) });
            }
        });
        return None;
    }
    let known = session()
        .ok()?
        .corrections
        .iter()
        .rev()
        .find(|known| known.window == window && (known.at.0 - x).abs() < 800.0 && (known.at.1 - y).abs() < 400.0)
        .copied();
    let seen = if opening {
        screen_colors::caret_bar_above(rect).map(|bar| CaretCorrection { window, offset: bar.top - y, height: bar.height, at: (x, y) })
    } else {
        None
    };
    let (correction, fresh) = opening_caret(seen, known)?;
    if fresh {
        remember_caret(client, correction);
    }
    let caret = serde_json::json!([x, y + correction.offset, width.max(1.0), correction.height]);
    let mut corrected = message.clone();
    corrected["caret"] = caret.clone();
    corrected["line"] = caret;
    corrected["clipped"] = serde_json::json!(false);
    corrected["source"] = serde_json::json!("pixels");
    corrected["madeUp"] = serde_json::json!(format!(
        "made up {x:.0},{y:.0},{height:.0}; caret {}",
        if fresh { "seen" } else { "remembered" }
    ));
    Some(corrected)
}

/// Looks at a field's caret after it got the keyboard: how many, how far apart (more than a blink).
const CARET_LOOKS: usize = 6;
const CARET_LOOK_MS: u64 = 130;

/// The caret to use as a formula opens, from the one `seen` now and the one `known` in that window
/// (offsets from the same reported rectangle); true: the one seen (to remember). A caret seen
/// inside the known one and shorter was caught fading (its blink): the known one stays.
fn opening_caret(seen: Option<CaretCorrection>, known: Option<CaretCorrection>) -> Option<(CaretCorrection, bool)> {
    match (seen, known) {
        (Some(seen), Some(known)) => {
            let inside = seen.offset >= known.offset - 2.0 && seen.offset + seen.height <= known.offset + known.height + 2.0;
            if inside && seen.height < known.height - 2.0 {
                Some((known, false))
            } else {
                Some((seen, true))
            }
        }
        (Some(seen), None) => Some((seen, true)),
        (None, Some(known)) => Some((known, false)),
        (None, None) => None,
    }
}

/// Keeps the real caret seen in a window (the last few windows).
fn remember_caret(client: &str, correction: CaretCorrection) {
    let Ok(mut session) = session() else { return };
    let previous = session.corrections.iter().find(|known| known.window == correction.window).copied();
    session.corrections.retain(|known| known.window != correction.window);
    session.corrections.push(correction);
    if session.corrections.len() > 8 {
        session.corrections.remove(0);
    }
    drop(session);
    let changed = previous
        .map(|previous| (previous.offset - correction.offset).abs() > 1.0 || (previous.height - correction.height).abs() > 1.0)
        .unwrap_or(true);
    if changed {
        log(&format!(
            "caret of {client} seen: {:.0} px tall, {:.0} px from the reported rectangle's top",
            correction.height, correction.offset
        ));
    }
}

/// The field took the input method (or answered with its geometry): open the formula the
/// shortcut asked for, or follow the formula's place.
fn caret_reported(app: &AppHandle, message: &serde_json::Value) {
    let corrected = with_real_caret(message);
    let message = corrected.as_ref().unwrap_or(message);
    let geometry = input_method_geometry(message);
    let client = message.get("client").and_then(|value| value.as_str()).unwrap_or("").to_string();
    let kind = message.get("t").and_then(|value| value.as_str()).unwrap_or("");
    enum Next {
        Open(Option<Activation>, bool),
        Follow(ScreenRect, PanelLayout, Option<(u64, FieldText)>),
        Nothing,
    }
    // what the report did to the formula's size, for anywhere.log (written after the lock)
    let mut sizing: Option<String> = None;
    let next = {
        let Ok(mut session) = session() else { return };
        if session.awaiting_input_method == session.generation && session.phase == Phase::Activating {
            match geometry.as_ref() {
                Some(geometry) => {
                    session.awaiting_input_method = 0;
                    session.app_name = client.clone();
                    session.caret = Some(geometry.caret);
                    session.scale = geometry.scale;
                    // The field's text size when it tells a believable one, else the system's.
                    session.reference_height = geometry.measured;
                    session.field_text = match geometry.measured {
                        Some(_) => geometry.text.clone(),
                        None => text_of_line(system_text_size() * LINE_TO_FONT, LINE_TO_FONT),
                    };
                    // The field's colours now, before anything of VisualTeX's is on the screen
                    // (read after the formula shows, its own strip could be taken for the field).
                    if let Some(colors) = screen_colors::field_colors(geometry.caret) {
                        log(&format!(
                            "field colour gen={} text={} background={} (before the formula)",
                            session.generation, colors.text, colors.background
                        ));
                        session.field_text.text_color = Some(colors.text);
                        session.field_text.background_color = Some(colors.background);
                        session.color_read = session.generation;
                    }
                    sizing = Some(format!(
                        "size gen={} {kind} {} -> {} ({})",
                        session.generation,
                        geometry.report,
                        describe_text(&session.field_text),
                        if geometry.measured.is_some() { "field" } else { "not believable: system text size" }
                    ));
                    // In the line from the start, at the caret (where the composition begins),
                    // as on macOS; the composition's own place follows when the field reports it.
                    session.inline_anchor = Some(inline_anchor(geometry));
                    Next::Open(pending_activation(&session), session.ready)
                }
                // Activated without a caret yet: ask again (the TIP reports it with the anchor).
                None => Next::Nothing,
            }
        } else if let (Phase::Presented | Phase::Activating, Some(geometry), Delivery::InputMethod) =
            (session.phase, geometry.as_ref(), session.delivery)
        {
            if kind_is_anchor(message) && !ime::from_other_field(message) {
                session.caret = Some(geometry.caret);
                session.scale = geometry.scale;
                let anchor = inline_anchor(geometry);
                session.inline_anchor = Some(anchor);
                // The anchor's height sizes the formula only when it is believable and near the
                // height that sized it first; otherwise the anchor only places it.
                let sizes = anchor_sizes(geometry.measured, session.reference_height);
                let mut text = session.field_text.clone();
                if sizes {
                    session.reference_height = session.reference_height.or(geometry.measured);
                    text.font_size = geometry.text.font_size;
                    text.line_height = geometry.text.line_height;
                    text.line_descent = geometry.text.line_descent;
                }
                let restyle = (text != session.field_text).then(|| (session.generation, text.clone()));
                sizing = Some(format!(
                    "size gen={} anchor {} reference={} -> {}{}",
                    session.generation,
                    geometry.report,
                    session.reference_height.map(|value| format!("{value:.1}px")).unwrap_or_else(|| "-".to_string()),
                    if sizes { "sizes" } else { "places only" },
                    restyle
                        .as_ref()
                        .map(|(_, text)| format!(", {} -> {}", describe_text(&session.field_text), describe_text(text)))
                        .unwrap_or_default()
                ));
                session.field_text = text;
                let read_color = (session.color_read != session.generation && session.phase == Phase::Presented).then(|| {
                    session.color_read = session.generation;
                    (session.generation, geometry.caret)
                });
                if let Some((generation, caret)) = read_color {
                    read_field_color(app, generation, caret);
                }
                if session.phase == Phase::Presented {
                    Next::Follow(anchor, session.layout, restyle)
                } else {
                    Next::Nothing
                }
            } else {
                Next::Nothing
            }
        } else {
            Next::Nothing
        }
    };
    if let Some(line) = sizing {
        log(&line);
    }
    match next {
        Next::Open(activation, ready) => {
            // the formula's messages go to this field's TIP until it closes
            ime::pin(Some(message));
            ime::send(serde_json::json!({ "t": "session", "on": true }));
            log(&format!(
                "session open client={client:?} caret={}",
                geometry
                    .as_ref()
                    .map(|geometry| format!("{:.0},{:.0},{:.0} scale={:.2}", geometry.caret.x, geometry.caret.y, geometry.caret.height, geometry.scale))
                    .unwrap_or_else(|| "-".to_string()),
            ));
            if let (Some(activation), true) = (activation, ready) {
                send_activation(app, activation);
            }
        }
        Next::Follow(anchor, layout, restyle) => {
            if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
                panel::reframe(&window, anchor, layout, panel::Placement::InLine);
                if let Some((generation, text)) = restyle {
                    emit_style(&window, generation, &text);
                }
            }
        }
        Next::Nothing => {}
    }
}

/// The anchor's place (the TIP's "geometry" while a formula is open), not a mere focus report.
fn kind_is_anchor(message: &serde_json::Value) -> bool {
    message.get("t").and_then(|value| value.as_str()) == Some("geometry")
}

fn emit_style(window: &tauri::WebviewWindow, generation: u64, text: &FieldText) {
    let mut payload = serde_json::to_value(text).unwrap_or_default();
    if let Some(object) = payload.as_object_mut() {
        object.insert("generation".to_string(), serde_json::json!(generation));
    }
    let _ = window.emit(STYLE_EVENT, payload);
}

/// Reads the formula's colour off the screen (from the field's background) and gives it to the
/// page.
fn read_field_color(app: &AppHandle, generation: u64, caret: ScreenRect) {
    let app = app.clone();
    let _ = std::thread::Builder::new().name("visualtex-anywhere-color".to_string()).spawn(move || {
        let window = app.get_webview_window(WINDOW_LABEL);
        if let Some(window) = window.as_ref() {
            panel::exclude_from_capture(window, true);
            // the next composed frame leaves the formula out
            std::thread::sleep(Duration::from_millis(50));
        }
        let colors = screen_colors::field_colors(caret);
        if let Some(window) = window.as_ref() {
            panel::exclude_from_capture(window, false);
        }
        log(&format!(
            "field colour gen={generation} text={:?} background={:?}",
            colors.as_ref().map(|colors| colors.text.as_str()),
            colors.as_ref().map(|colors| colors.background.as_str()),
        ));
        let Some(colors) = colors else { return };
        let text = {
            let Ok(mut session) = session() else { return };
            if session.generation != generation {
                return;
            }
            session.field_text.text_color = Some(colors.text);
            session.field_text.background_color = Some(colors.background);
            session.field_text.clone()
        };
        if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
            emit_style(&window, generation, &text);
        }
    });
}

/// A key the input method took: to the editor once the formula is up (until then it waits, so
/// fast typing after the shortcut is not lost).
fn input_method_key(app: &AppHandle, message: &serde_json::Value) {
    let mut key = message.clone();
    let deliver_now = {
        let Ok(mut session) = session() else { return };
        if session.phase == Phase::Hidden {
            return;
        }
        if let Some(object) = key.as_object_mut() {
            object.insert("generation".to_string(), serde_json::json!(session.generation));
        }
        if session.phase == Phase::Presented && session.pending_keys.is_empty() {
            true
        } else {
            if session.pending_keys.len() < MAX_PENDING_KEYS {
                session.pending_keys.push(key.clone());
            }
            false
        }
    };
    if deliver_now {
        if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
            let _ = window.emit(KEY_EVENT, key);
        }
    }
}

fn flush_pending_keys(app: &AppHandle, generation: u64) {
    let keys = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation {
            return;
        }
        std::mem::take(&mut session.pending_keys)
    };
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        for key in keys {
            let _ = window.emit(KEY_EVENT, key);
        }
    }
}

/// Hands the session to the page and watches for its answer (the strip's layout report).
fn send_activation(app: &AppHandle, activation: Activation) {
    let generation = activation.generation;
    let Some(window) = app.get_webview_window(WINDOW_LABEL) else { return };
    if let Err(error) = window.emit(ACTIVATE_EVENT, activation) {
        log(&format!("activation gen={generation} not sent: {error}"));
        return;
    }
    let app = app.clone();
    let _ = std::thread::Builder::new().name("visualtex-anywhere-watch".to_string()).spawn(move || {
        std::thread::sleep(PAGE_ANSWER_WAIT);
        recover_unanswered(&app, generation);
    });
}

/// The page did not answer an activation: reload it once, or give the session up.
fn recover_unanswered(app: &AppHandle, generation: u64) {
    let reload = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation || session.phase != Phase::Activating {
            return;
        }
        if session.reloaded_for == generation {
            session.phase = Phase::Hidden;
            false
        } else {
            session.reloaded_for = generation;
            session.ready = false;
            true
        }
    };
    let Some(window) = app.get_webview_window(WINDOW_LABEL) else { return };
    if reload {
        log(&format!("page did not answer gen={generation}: reloading it"));
        panel::park(&window);
        let _ = window.reload();
    } else {
        log(&format!("page did not answer gen={generation} after a reload: session dropped"));
        panel::park(&window);
        ime::send(serde_json::json!({ "t": "cancel" }));
        ime::pin(None);
        formula_keys(false);
    }
}

/// Shows the formula at the caret, once per session.
fn present(app: &AppHandle, generation: u64) {
    let (caret, layout, inline, read_color, delivery) = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation || session.phase != Phase::Activating || !session.ready {
            return;
        }
        let Some(caret) = session.caret else { return };
        session.phase = Phase::Presented;
        let read_color = session.color_read != generation && session.delivery == Delivery::InputMethod;
        if read_color {
            session.color_read = generation;
        }
        (caret, session.layout, session.inline_anchor, read_color, session.delivery)
    };
    log(&format!(
        "present gen={generation} layout={:.0}x{:.0} inline={} delivery={delivery:?}",
        layout.width,
        layout.height,
        inline.is_some()
    ));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        match (delivery, inline) {
            (Delivery::Paste, _) => panel::present_focused(&window, caret, layout),
            (_, Some(anchor)) => panel::present(&window, anchor, layout, panel::Placement::InLine),
            (_, None) => panel::present(&window, caret, layout, panel::Placement::BelowLine),
        }
    }
    flush_pending_keys(app, generation);
    if read_color {
        read_field_color(app, generation, caret);
    }
}

fn hide(app: &AppHandle, keep_draft: bool, reason: &str) {
    hide_with(app, keep_draft, reason, true);
}

/// Hides the formula. `tell_input_method` drops the field's anchor and hands the field back to
/// the input method it came from (not when the input method itself ended the session); for the
/// pasting strip it gives the field its keyboard back (not when the strip lost it to a click).
fn hide_with(app: &AppHandle, keep_draft: bool, reason: &str, tell_input_method: bool) {
    let (generation, delivery, field_window) = {
        let Ok(mut session) = session() else { return };
        if session.phase == Phase::Hidden {
            return;
        }
        session.phase = Phase::Hidden;
        session.draft = keep_draft;
        session.awaiting_input_method = 0;
        session.pending_keys.clear();
        (session.generation, session.delivery, session.field_window)
    };
    if delivery == Delivery::Paste {
        formula_keys(false);
        log(&format!("hide gen={generation} keep_draft={keep_draft} ({reason}, pasting strip)"));
        if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
            panel::park(&window);
            let _ = window.emit(SUSPEND_EVENT, Suspension { generation, keep_draft });
        }
        if tell_input_method {
            let _ = std::thread::Builder::new()
                .name("visualtex-anywhere-refocus".to_string())
                .spawn(move || deliver::restore_foreground(field_window));
        }
        return;
    }
    if tell_input_method {
        ime::send(serde_json::json!({ "t": "cancel" }));
    }
    ime::pin(None);
    formula_keys(false);
    check_profile_later(generation);
    log(&format!("hide gen={generation} keep_draft={keep_draft} ({reason})"));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        panel::park(&window);
        let _ = window.emit(SUSPEND_EVENT, Suspension { generation, keep_draft });
    }
}

/// The page mounted its editor: from now on it takes activations directly.
#[tauri::command]
pub(crate) fn anywhere_ready(app: AppHandle) -> Result<(), String> {
    let pending = {
        let mut session = session()?;
        session.ready = true;
        match session.phase {
            Phase::Hidden => None,
            Phase::Activating | Phase::Presented => {
                session.draft = false;
                if session.phase == Phase::Presented {
                    session.phase = Phase::Activating;
                }
                pending_activation(&session)
            }
        }
    };
    log("page ready");
    match pending {
        Some(activation) => send_activation(&app, activation),
        None => {
            if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
                panel::park(&window);
            }
        }
    }
    Ok(())
}

/// The strip measured itself: show the formula there, or follow its new size.
#[tauri::command]
pub(crate) fn anywhere_layout(app: AppHandle, generation: u64, layout: PanelLayout) -> Result<(), String> {
    if !layout.is_valid() {
        return Err("The VisualTeX Anywhere layout is invalid".to_string());
    }
    let (phase, anchor, placement) = {
        let mut session = session()?;
        if session.generation != generation {
            return Ok(());
        }
        session.layout = layout;
        match (session.delivery, session.inline_anchor) {
            (Delivery::Paste, _) => (session.phase, session.caret, panel::Placement::OnLine),
            (_, Some(inline)) => (session.phase, Some(inline), panel::Placement::InLine),
            (_, None) => (session.phase, session.caret, panel::Placement::BelowLine),
        }
    };
    match (phase, anchor, app.get_webview_window(WINDOW_LABEL)) {
        (Phase::Activating, _, _) => present(&app, generation),
        (Phase::Presented, Some(anchor), Some(window)) => panel::reframe(&window, anchor, layout, placement),
        _ => {}
    }
    Ok(())
}

/// The Image output mode: the anchor goes, the field goes back to its input method, and the PNG
/// is pasted where the anchor was.
#[tauri::command]
pub(crate) async fn anywhere_commit_image(app: AppHandle, generation: u64, data_base64: String) -> Result<CommitOutcome, String> {
    use base64::Engine as _;
    let png = base64::engine::general_purpose::STANDARD
        .decode(data_base64.trim())
        .map_err(|error| format!("The formula image could not be read: {error}"))?;
    if png.len() > 50 * 1024 * 1024 || !png.starts_with(b"\x89PNG\r\n\x1a\n") {
        return Err("The formula image is not a PNG".to_string());
    }
    let (delivery, field_window) = {
        let mut session = session()?;
        if session.generation != generation || session.phase != Phase::Presented {
            return Ok(CommitOutcome::Stale);
        }
        session.phase = Phase::Hidden;
        session.draft = false;
        session.pending_keys.clear();
        (session.delivery, session.field_window)
    };
    if delivery == Delivery::InputMethod {
        ime::send(serde_json::json!({ "t": "cancel" }));
        check_profile_later(generation);
    }
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        panel::park(&window);
    }
    formula_keys(false);
    let bytes = png.len();
    let pasted = tauri::async_runtime::spawn_blocking(move || {
        if delivery == Delivery::Paste {
            // The strip had the keyboard: the field gets it back first.
            deliver::restore_foreground(field_window);
        }
        // The field drops the anchor first.
        std::thread::sleep(Duration::from_millis(120));
        deliver::paste_png(png)
    })
    .await
    .map_err(|error| format!("VisualTeX Anywhere insertion was interrupted: {error}"))?;
    log(&format!("commit image gen={generation} bytes={bytes} paste={pasted:?}"));
    pasted?;
    Ok(CommitOutcome::Inserted)
}

/// Enter: the field's anchor becomes the formula's text; the field goes back to its input method.
#[tauri::command]
pub(crate) async fn anywhere_commit(app: AppHandle, generation: u64, text: String) -> Result<CommitOutcome, String> {
    if text.len() > 1024 * 1024 {
        return Err("The formula is too long to insert".to_string());
    }
    let (delivery, field_window) = {
        let mut session = session()?;
        if session.generation != generation || session.phase != Phase::Presented {
            return Ok(CommitOutcome::Stale);
        }
        session.phase = Phase::Hidden;
        session.draft = false;
        session.pending_keys.clear();
        (session.delivery, session.field_window)
    };
    if delivery == Delivery::Paste {
        if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
            panel::park(&window);
        }
        formula_keys(false);
        let chars = text.chars().count();
        let pasted = tauri::async_runtime::spawn_blocking(move || {
            deliver::restore_foreground(field_window);
            std::thread::sleep(Duration::from_millis(80));
            deliver::paste_text(&text)
        })
        .await
        .map_err(|error| format!("VisualTeX Anywhere insertion was interrupted: {error}"))?;
        log(&format!("commit gen={generation} chars={chars} paste={pasted:?}"));
        pasted?;
        return Ok(CommitOutcome::Inserted);
    }
    let sent = ime::send(serde_json::json!({ "t": "commit", "text": text }));
    ime::pin(None);
    check_profile_later(generation);
    log(&format!("commit gen={generation} chars={} input_method={sent}", text.chars().count()));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        panel::park(&window);
    }
    formula_keys(false);
    if sent {
        Ok(CommitOutcome::Inserted)
    } else {
        // The field is gone: the formula is on the clipboard instead.
        let copied = tauri::async_runtime::spawn_blocking(move || crate::windows_quick_ocr::write_clipboard_text(&text))
            .await
            .map_err(|error| format!("Clipboard task failed: {error}"))?;
        copied?;
        Ok(CommitOutcome::Copied)
    }
}

/// Escape (keep_draft false) or the strip's own close.
#[tauri::command]
pub(crate) fn anywhere_dismiss(app: AppHandle, generation: u64, keep_draft: bool) -> Result<(), String> {
    if session()?.generation == generation {
        hide(&app, keep_draft, "page");
    }
    Ok(())
}

/// The formula's text for the field (if the field ends the composition itself), its caret (CSS
/// pixels from the formula's start, x and down to the caret's line: the input method places pinyin
/// candidates and the 中/英 indicator there) and the widths of the pieces TeX may break it into
/// (CSS pixels: the field holds each in its text and wraps them as its own text).
#[tauri::command]
pub(crate) fn anywhere_marked(
    generation: u64,
    text: String,
    width: f64,
    caret_x: Option<f64>,
    caret_y: Option<f64>,
    pieces: Option<Vec<f64>>,
) -> Result<(), String> {
    if text.len() > 256 * 1024 {
        return Err("The formula is too long to show".to_string());
    }
    let scale = {
        let session = session()?;
        if session.generation != generation || session.phase == Phase::Hidden {
            return Ok(());
        }
        session.scale
    };
    if width.is_finite() {
        let pieces: Vec<f64> = pieces
            .unwrap_or_default()
            .into_iter()
            .take(MAX_PIECES)
            .map(|piece| if piece.is_finite() { (piece * scale).clamp(0.0, 16384.0) } else { 0.0 })
            .collect();
        let _ = (caret_x, caret_y);
        ime::send(serde_json::json!({
            "t": "inline",
            "width": (width * scale).clamp(0.0, 16384.0),
            "text": text,
            "pieces": pieces,
        }));
    }
    Ok(())
}

/// The formula's caret entered or left text (\text{}). (The input method types formulas only: it
/// has nothing to do there; the page still tells.)
#[tauri::command]
pub(crate) fn anywhere_mode(generation: u64, text: bool) -> Result<(), String> {
    let _ = (generation, text);
    Ok(())
}

/// The page's own diagnostics (commit outcomes, errors), into anywhere.log.
#[tauri::command]
pub(crate) fn anywhere_log(line: String) {
    let line: String = line.chars().take(1600).map(|c| if c.is_control() { ' ' } else { c }).collect();
    log(&format!("page {line}"));
}

/// Screenshot OCR while a formula is open: the user selects an area (Windows' screen snip). The
/// selection takes the foreground for a moment; the field keeps its formula meanwhile and gets
/// the keyboard back after. Returns the PNG (base64), None when cancelled.
#[tauri::command]
pub(crate) async fn anywhere_capture_screenshot(app: AppHandle) -> Result<Option<String>, String> {
    use base64::Engine as _;
    let (field_window, delivery) = {
        let mut session = session()?;
        session.holding = true;
        (session.field_window, session.delivery)
    };
    // The keyboard goes back where it was: the field (input method), or the strip.
    let back_to = match delivery {
        Delivery::InputMethod => field_window,
        Delivery::Paste => app
            .get_webview_window(WINDOW_LABEL)
            .and_then(|window| window.hwnd().ok())
            .map(|hwnd| hwnd.0 as isize)
            .unwrap_or(0),
    };
    ime::send(serde_json::json!({ "t": "hold", "on": true }));
    let capture = tauri::async_runtime::spawn_blocking(move || {
        let mode = crate::windows_silent_ocr_hotkey::configured_capture_mode();
        let capture = crate::windows_quick_ocr::capture_windows_quick_ocr_bytes(&mode, 60_000);
        deliver::restore_foreground(back_to);
        capture
    })
    .await
    .map_err(|error| format!("Screenshot task failed: {error}"));
    if let Ok(mut session) = session() {
        session.holding = false;
    }
    ime::send(serde_json::json!({ "t": "hold", "on": false }));
    let capture = capture??;
    Ok(capture.map(|bytes| base64::engine::general_purpose::STANDARD.encode(bytes)))
}

/// Ctrl+V in an open formula: the clipboard's image (base64, for OCR) or text.
#[derive(Serialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub(crate) enum AnywhereClipboard {
    Image {
        #[serde(rename = "dataBase64")]
        data_base64: String,
        extension: String,
    },
    Text {
        text: String,
    },
    Nothing,
}

#[tauri::command]
pub(crate) async fn anywhere_read_clipboard() -> Result<AnywhereClipboard, String> {
    use base64::Engine as _;
    let content = tauri::async_runtime::spawn_blocking(deliver::read_clipboard)
        .await
        .map_err(|error| format!("Clipboard task failed: {error}"))??;
    Ok(match content {
        deliver::ClipboardContent::Image { bytes, extension } => {
            if bytes.len() > 20 * 1024 * 1024 {
                return Err("The copied image is larger than 20 MB".to_string());
            }
            AnywhereClipboard::Image { data_base64: base64::engine::general_purpose::STANDARD.encode(bytes), extension }
        }
        deliver::ClipboardContent::Text(text) => AnywhereClipboard::Text { text },
        deliver::ClipboardContent::Nothing => AnywhereClipboard::Nothing,
    })
}

/// Settings > Shortcuts: a global shortcut that opens and closes a formula (`action`:
/// "anywhere-toggle", the default, or "anywhere-open-display"; a null chord removes the display
/// one).
#[tauri::command]
pub(crate) async fn anywhere_configure_hotkey(
    app: AppHandle,
    chord: serde_json::Value,
    action: Option<String>,
) -> Result<String, String> {
    let slot = match action.as_deref() {
        None => hotkey::Slot::Toggle,
        Some(action) => hotkey::Slot::from_action(action).ok_or_else(|| format!("Unknown shortcut {action}"))?,
    };
    let chord: Option<hotkey::Chord> =
        serde_json::from_value(chord).map_err(|error| format!("Unreadable shortcut: {error}"))?;
    tauri::async_runtime::spawn_blocking(move || hotkey::configure(&app, slot, chord))
        .await
        .map_err(|error| format!("Shortcut task failed: {error}"))?
}

/// The page's shortcuts that act while a formula is open (Settings > Keyboard shortcuts): registered
/// with Windows while a formula is open, so that they work whatever the application keeps.
#[tauri::command]
pub(crate) fn anywhere_formula_shortcuts(shortcuts: Vec<hotkey::FormulaShortcut>) {
    hotkey::set_formula_shortcuts(shortcuts);
}

#[tauri::command]
pub(crate) fn anywhere_status() -> AnywhereStatus {
    AnywhereStatus { accessibility_trusted: true, shortcut: hotkey::shortcut_label(), input_method: ime::is_installed() }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn layouts_must_fit_their_anchor_and_stay_finite() {
        assert!(PanelLayout::EMPTY.is_valid());
        let mut layout = PanelLayout::EMPTY;
        layout.anchor_x = layout.width + 1.0;
        assert!(!layout.is_valid());
        layout = PanelLayout::EMPTY;
        layout.height = f64::NAN;
        assert!(!layout.is_valid());
    }

    #[test]
    fn geometry_reads_the_anchor_in_physical_pixels() {
        let geometry = input_method_geometry(&serde_json::json!({ "caret": [100, 200, 0, 24] })).unwrap();
        assert_eq!(geometry.caret.height, 24.0);
        assert!(geometry.baseline < 224.0 && geometry.baseline > 200.0);
        assert!(input_method_geometry(&serde_json::json!({ "caret": [0, 0, 0, 0] })).is_none());
    }

    #[test]
    fn rectangles_that_are_no_line_of_text_place_but_do_not_size() {
        // a line of text at any scale up to 2: sizes, the caret its line
        let line = input_method_geometry(&serde_json::json!({ "caret": [100, 200, 0, 30] })).unwrap();
        assert!(line.measured.is_some());
        assert_eq!(line.text.line_height, line.measured);
        assert!(line.text.font_size.unwrap() >= 9.0);
        // a hidden or made-up rectangle a few pixels tall
        let tiny = input_method_geometry(&serde_json::json!({ "caret": [100, 200, 0, 4] })).unwrap();
        assert!(tiny.measured.is_none() && tiny.text.font_size.is_none() && tiny.text.line_height.is_none());
        assert_eq!(tiny.caret.height, 4.0);
        // cut to the part the field shows
        let clipped = input_method_geometry(&serde_json::json!({ "caret": [100, 200, 0, 30], "clipped": true })).unwrap();
        assert!(clipped.measured.is_none() && clipped.text.font_size.is_none());
    }

    #[test]
    fn the_anchor_sizes_only_near_the_first_report() {
        assert!(anchor_sizes(Some(18.0), Some(20.0)));
        // an empty line's line box first, the text's height on the anchor
        assert!(anchor_sizes(Some(18.4), Some(26.7)));
        // a made-up rectangle much shorter or taller than the field's text
        assert!(!anchor_sizes(Some(11.0), Some(20.0)));
        assert!(!anchor_sizes(Some(40.0), Some(20.0)));
        assert!(!anchor_sizes(None, Some(20.0)));
        // the first believable report sizes a formula opened at the system's size
        assert!(anchor_sizes(Some(14.0), None));
    }

    #[test]
    fn the_system_text_size_is_a_text_size() {
        let size = system_text_size();
        assert!((8.0..=48.0).contains(&size));
        let text = text_of_line(size * LINE_TO_FONT, LINE_TO_FONT);
        assert!(text.font_size.unwrap() >= 9.0 && text.line_descent.unwrap() < text.line_height.unwrap());
    }

    #[test]
    fn a_fading_caret_does_not_shrink_the_known_one() {
        let caret = |offset: f64, height: f64| CaretCorrection { window: 1, offset, height, at: (0.0, 0.0) };
        // WeChat: 27 px learnt; at the opening a fading frame shows 20 px of it
        assert_eq!(opening_caret(Some(caret(-21.0, 20.0)), Some(caret(-26.0, 27.0))), Some((caret(-26.0, 27.0), false)));
        // the whole caret again, or a taller one (a larger font): the new one
        assert_eq!(opening_caret(Some(caret(-27.0, 27.0)), Some(caret(-26.0, 27.0))), Some((caret(-27.0, 27.0), true)));
        assert_eq!(opening_caret(Some(caret(-33.0, 34.0)), Some(caret(-26.0, 27.0))), Some((caret(-33.0, 34.0), true)));
        // only one of them
        assert_eq!(opening_caret(None, Some(caret(-26.0, 27.0))), Some((caret(-26.0, 27.0), false)));
        assert_eq!(opening_caret(Some(caret(-26.0, 27.0)), None), Some((caret(-26.0, 27.0), true)));
        assert_eq!(opening_caret(None, None), None);
    }

    #[test]
    fn an_imm32_caret_is_its_fonts_height() {
        // Qt's text cursor in WeChat: 28 physical pixels at 150 %, a 14 px font
        let caret = serde_json::json!({ "caret": [100, 200, 2, 42], "source": "imm" });
        let line = serde_json::json!({ "caret": [100, 200, 2, 42] });
        let imm = input_method_geometry(&caret).unwrap().text.font_size.unwrap();
        let text = input_method_geometry(&line).unwrap().text.font_size.unwrap();
        assert!((imm / text - LINE_TO_FONT / IMM_CARET_TO_FONT).abs() < 1e-9);
        assert!(input_method_geometry(&caret).unwrap().report.contains("source=imm"));
    }
}
