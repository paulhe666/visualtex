//! VisualTeX Anywhere on Windows: a global shortcut turns the field that has the keyboard, in
//! any application, into a formula input, through VisualTeX's own input method
//! (apps/windows/input-method), as on the Mac (apps/macos/src-tauri/src/anywhere).
//!
//! The shortcut switches the user's input method to VisualTeX's (unless it is in use already).
//! Its TIP, in the field's application, reports the caret; VisualTeX then opens the formula
//! there: the field holds an anchor (its composition), the TIP forwards every key to this page,
//! the formula is drawn in the field's line by a transparent window that never takes the
//! keyboard, and Enter replaces the anchor with the formula's text (or pastes its image) and
//! hands the field back to the input method it had.

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
/// Pinyin spelled in the formula's text: the editor's composition.
const COMPOSE_EVENT: &str = "visualtex-anywhere-compose";
/// The shortcut switched the field to the input method; a field that has not activated it by
/// then does not take text (or the switch failed).
const INPUT_METHOD_WAIT: Duration = Duration::from_millis(1200);
/// A TIP reconnects within a second of VisualTeX starting.
const IME_RELINK_WAIT: Duration = Duration::from_millis(1500);
const MAX_PENDING_KEYS: usize = 256;
/// A page that has not answered an activation after this long is reloaded.
const PAGE_ANSWER_WAIT: Duration = Duration::from_millis(1500);
/// What a line of the field's text is to its font size (Chrome reports the line box).
const LINE_TO_FONT: f64 = 1.19;

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
    /// The input method the field goes back to after the formula.
    return_to: Option<serde_json::Value>,
    awaiting_input_method: u64,
    pending_keys: Vec<serde_json::Value>,
    field_text: FieldText,
    color_read: u64,
    /// The field's window when the shortcut was pressed (the screenshot selection takes the
    /// foreground for a moment).
    field_window: isize,
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
    return_to: None,
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
    color_read: 0,
    field_window: 0,
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

/// At startup: the shortcut, the input method's link, VisualTeX Input, and the formula's window
/// (created a moment later, when the input method is installed).
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
        if !ime::is_enabled() {
            log("input method not installed");
            return;
        }
        ime::start_input_server();
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

/// The shortcut: open a formula in the field that has the keyboard, or close the open one
/// (keeping it for the next press).
pub(super) fn toggle(app: &AppHandle) {
    if session().map(|session| session.phase != Phase::Hidden).unwrap_or(false) {
        hide(app, true, "shortcut");
        return;
    }
    let app = app.clone();
    let _ = std::thread::Builder::new()
        .name("visualtex-anywhere-activate".to_string())
        .spawn(move || begin_input_method(&app));
}

fn begin_input_method(app: &AppHandle) {
    let started = std::time::Instant::now();
    if !ime::is_enabled() {
        log("shortcut: the input method is not installed");
        return;
    }
    let previous = ime::current_profile();
    let already = previous.as_ref().map(ime::is_visualtex).unwrap_or(false);
    let field_window = unsafe { windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow() }.0 as isize;
    let generation = {
        let Ok(mut session) = session() else { return };
        session.generation += 1;
        session.phase = Phase::Activating;
        session.caret = None;
        session.inline_anchor = None;
        session.return_to = previous.clone().filter(|_| !already);
        session.awaiting_input_method = session.generation;
        session.pending_keys.clear();
        session.field_text = FieldText::default();
        session.field_window = field_window;
        session.generation
    };
    with_panel(app, |_| {});
    crate::windows_silent_ocr_hotkey::suspend_for_formula(true);
    let switched = if already {
        // Typing with VisualTeX already: its field answers with the caret. A TIP relinks
        // within a second of VisualTeX starting.
        let deadline = std::time::Instant::now() + IME_RELINK_WAIT;
        while !ime::is_linked() && std::time::Instant::now() < deadline {
            std::thread::sleep(Duration::from_millis(40));
        }
        ime::send(serde_json::json!({ "t": "caret" }))
    } else {
        match ime::switch_in() {
            Ok(()) => true,
            Err(error) => {
                log(&format!("input method switch failed gen={generation}: {error}"));
                false
            }
        }
    };
    log(&format!(
        "input method requested gen={generation} from={} switched={switched} linked={} ms={}",
        previous.as_ref().map(|profile| profile.to_string()).unwrap_or_else(|| "?".to_string()),
        ime::is_linked(),
        started.elapsed().as_millis()
    ));
    if !switched {
        give_up(app, generation, false);
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
            give_up(&app, generation, !already);
        }
    });
}

/// The field did not take the input method: back to the input method it had.
fn give_up(app: &AppHandle, generation: u64, switch_back: bool) {
    let return_to = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation {
            return;
        }
        session.awaiting_input_method = 0;
        session.phase = Phase::Hidden;
        session.return_to.take()
    };
    if let (true, Some(profile)) = (switch_back, return_to) {
        ime::switch_to(&profile);
    }
    crate::windows_silent_ocr_hotkey::suspend_for_formula(false);
    let _ = app;
}

/// The input method's report of the field: where the formula goes and how its text looks there.
struct FieldGeometry {
    caret: ScreenRect,
    baseline: f64,
    scale: f64,
    text: FieldText,
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
    let css_height = height / scale;
    let font_size = (css_height / LINE_TO_FONT).clamp(6.0, 64.0);
    let descent = height * 0.22;
    Some(FieldGeometry {
        caret: ScreenRect { x, y, width, height },
        baseline: y + height - descent,
        scale,
        text: FieldText {
            font_size: Some(font_size),
            text_color: None,
            x_height: None,
            line_height: Some(css_height.clamp(10.0, 64.0)),
            line_descent: Some(descent / scale),
            background_color: None,
        },
    })
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
        input_method: true,
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
    match kind {
        "hello" => log(&format!(
            "input method {}",
            message.get("version").and_then(|value| value.as_str()).unwrap_or("?")
        )),
        "activate" | "geometry" => caret_reported(app, message),
        "key" => input_method_key(app, message),
        "compose" => {
            let generation = session().ok().filter(|session| session.phase == Phase::Presented).map(|session| session.generation);
            if let (Some(generation), Some(window)) = (generation, app.get_webview_window(WINDOW_LABEL)) {
                let mut payload = message.clone();
                if let Some(object) = payload.as_object_mut() {
                    object.insert("generation".to_string(), serde_json::json!(generation));
                }
                let _ = window.emit(COMPOSE_EVENT, payload);
            }
        }
        "deactivate" | "disconnected" => {
            let open = session().map(|session| session.phase != Phase::Hidden && session.awaiting_input_method == 0).unwrap_or(false);
            if open {
                hide_with(app, true, if kind == "deactivate" { "input method left the field" } else { "input method gone" }, false);
            }
        }
        "forced" => {
            if session().map(|session| session.phase != Phase::Hidden).unwrap_or(false) {
                hide_with(app, false, "field kept the text", false);
            }
        }
        _ => {}
    }
}

/// The field took the input method (or answered with its geometry): open the formula the
/// shortcut asked for, or follow the formula's place.
fn caret_reported(app: &AppHandle, message: &serde_json::Value) {
    let geometry = input_method_geometry(message);
    let client = message.get("client").and_then(|value| value.as_str()).unwrap_or("").to_string();
    enum Next {
        Open(Option<Activation>, bool),
        Follow(ScreenRect, PanelLayout, Option<(u64, FieldText)>),
        Nothing,
    }
    let next = {
        let Ok(mut session) = session() else { return };
        if session.awaiting_input_method == session.generation && session.phase == Phase::Activating {
            match geometry.as_ref() {
                Some(geometry) => {
                    session.awaiting_input_method = 0;
                    session.app_name = client.clone();
                    session.caret = Some(geometry.caret);
                    session.scale = geometry.scale;
                    session.field_text = geometry.text.clone();
                    session.inline_anchor = None;
                    Next::Open(pending_activation(&session), session.ready)
                }
                // Activated without a caret yet: ask again (the TIP reports it with the anchor).
                None => Next::Nothing,
            }
        } else if let (Phase::Presented | Phase::Activating, Some(geometry)) = (session.phase, geometry.as_ref()) {
            if kind_is_anchor(message) {
                session.caret = Some(geometry.caret);
                session.scale = geometry.scale;
                let anchor = inline_anchor(geometry);
                session.inline_anchor = Some(anchor);
                let mut text = session.field_text.clone();
                text.font_size = geometry.text.font_size.or(text.font_size);
                text.line_height = geometry.text.line_height.or(text.line_height);
                text.line_descent = geometry.text.line_descent.or(text.line_descent);
                let restyle = (text != session.field_text).then(|| (session.generation, text.clone()));
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
    match next {
        Next::Open(activation, ready) => {
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
        let colors = screen_colors::field_colors(caret);
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
        let return_to = session().ok().and_then(|mut session| session.return_to.take());
        ime::send(serde_json::json!({ "t": "cancel", "returnTo": return_to }));
        crate::windows_silent_ocr_hotkey::suspend_for_formula(false);
    }
}

/// Shows the formula at the caret, once per session.
fn present(app: &AppHandle, generation: u64) {
    let (caret, layout, inline, read_color) = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation || session.phase != Phase::Activating || !session.ready {
            return;
        }
        let Some(caret) = session.caret else { return };
        session.phase = Phase::Presented;
        let read_color = session.color_read != generation;
        if read_color {
            session.color_read = generation;
        }
        (caret, session.layout, session.inline_anchor, read_color)
    };
    log(&format!("present gen={generation} layout={:.0}x{:.0} inline={}", layout.width, layout.height, inline.is_some()));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        match inline {
            Some(anchor) => panel::present(&window, anchor, layout, panel::Placement::InLine),
            None => panel::present(&window, caret, layout, panel::Placement::BelowLine),
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
/// the input method it came from (not when the input method itself ended the session).
fn hide_with(app: &AppHandle, keep_draft: bool, reason: &str, tell_input_method: bool) {
    let (generation, return_to) = {
        let Ok(mut session) = session() else { return };
        if session.phase == Phase::Hidden {
            return;
        }
        session.phase = Phase::Hidden;
        session.draft = keep_draft;
        session.awaiting_input_method = 0;
        session.pending_keys.clear();
        (session.generation, session.return_to.take())
    };
    if tell_input_method {
        ime::send(serde_json::json!({ "t": "cancel", "returnTo": return_to }));
    } else if let Some(profile) = return_to {
        // The field is gone: the user's input method comes back for the next one.
        ime::switch_to(&profile);
    }
    crate::windows_silent_ocr_hotkey::suspend_for_formula(false);
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
        match session.inline_anchor {
            Some(inline) => (session.phase, Some(inline), panel::Placement::InLine),
            None => (session.phase, session.caret, panel::Placement::BelowLine),
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
    let return_to = {
        let mut session = session()?;
        if session.generation != generation || session.phase != Phase::Presented {
            return Ok(CommitOutcome::Stale);
        }
        session.phase = Phase::Hidden;
        session.draft = false;
        session.pending_keys.clear();
        session.return_to.take()
    };
    ime::send(serde_json::json!({ "t": "cancel", "returnTo": return_to }));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        panel::park(&window);
    }
    crate::windows_silent_ocr_hotkey::suspend_for_formula(false);
    let bytes = png.len();
    let pasted = tauri::async_runtime::spawn_blocking(move || {
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
    let return_to = {
        let mut session = session()?;
        if session.generation != generation || session.phase != Phase::Presented {
            return Ok(CommitOutcome::Stale);
        }
        session.phase = Phase::Hidden;
        session.draft = false;
        session.pending_keys.clear();
        session.return_to.take()
    };
    let sent = ime::send(serde_json::json!({ "t": "commit", "text": text, "returnTo": return_to }));
    log(&format!("commit gen={generation} chars={} input_method={sent}", text.chars().count()));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        panel::park(&window);
    }
    crate::windows_silent_ocr_hotkey::suspend_for_formula(false);
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

/// The formula's text for the field (if the field ends the composition itself) and its caret
/// (CSS pixels from the formula's start: the input method places pinyin candidates there).
#[tauri::command]
pub(crate) fn anywhere_marked(generation: u64, text: String, width: f64, caret_x: Option<f64>) -> Result<(), String> {
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
        ime::send(serde_json::json!({
            "t": "inline",
            "width": (width * scale).clamp(0.0, 16384.0),
            "text": text,
            "caretX": caret_x.filter(|value| value.is_finite()).unwrap_or(0.0) * scale,
        }));
    }
    Ok(())
}

/// The formula's caret entered or left text (\text{}).
#[tauri::command]
pub(crate) fn anywhere_mode(generation: u64, text: bool) -> Result<(), String> {
    let current = {
        let session = session()?;
        session.generation == generation && session.phase != Phase::Hidden
    };
    if current {
        ime::send(serde_json::json!({ "t": "mode", "text": text }));
    }
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
pub(crate) async fn anywhere_capture_screenshot() -> Result<Option<String>, String> {
    use base64::Engine as _;
    let field_window = session()?.field_window;
    ime::send(serde_json::json!({ "t": "hold", "on": true }));
    let capture = tauri::async_runtime::spawn_blocking(move || {
        let mode = crate::windows_silent_ocr_hotkey::configured_capture_mode();
        let capture = crate::windows_quick_ocr::capture_windows_quick_ocr_bytes(&mode, 60_000);
        deliver::restore_foreground(field_window);
        capture
    })
    .await
    .map_err(|error| format!("Screenshot task failed: {error}"));
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

/// Settings > Shortcuts: the global shortcut that opens and closes a formula.
#[tauri::command]
pub(crate) async fn anywhere_configure_hotkey(app: AppHandle, chord: serde_json::Value) -> Result<String, String> {
    let chord: hotkey::Chord = serde_json::from_value(chord).map_err(|error| format!("Unreadable shortcut: {error}"))?;
    tauri::async_runtime::spawn_blocking(move || hotkey::configure(&app, chord))
        .await
        .map_err(|error| format!("Shortcut task failed: {error}"))?
}

#[tauri::command]
pub(crate) fn anywhere_status() -> AnywhereStatus {
    AnywhereStatus { accessibility_trusted: true, shortcut: hotkey::shortcut_label(), input_method: ime::is_enabled() }
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
}
