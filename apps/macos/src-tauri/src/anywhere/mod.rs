//! VisualTeX Anywhere: a global shortcut turns the text caret of the frontmost
//! application into a formula input, the way an input method takes over
//! typing. A borderless composition strip appears on the caret line, the
//! formula is written in it with the full VisualTeX editor, and Enter types
//! the finished formula into the application.
//!
//! The strip lives in a resident, non-activating panel (the Office formula
//! editor's window class, see vendor/tao-office-panel): the application keeps
//! the foreground and its text field keeps its caret and selection, so
//! delivering the formula is an ordinary paste.

#[cfg(target_os = "macos")]
mod ax;
#[cfg(target_os = "macos")]
mod deliver;
#[cfg(target_os = "macos")]
mod hotkey;
#[cfg(target_os = "macos")]
mod ime;
#[cfg(target_os = "macos")]
mod panel;
#[cfg(target_os = "macos")]
mod screen_colors;

use serde::{Deserialize, Serialize};
use tauri::AppHandle;
#[cfg(target_os = "macos")]
use std::sync::Mutex;
#[cfg(target_os = "macos")]
use std::time::Duration;
#[cfg(target_os = "macos")]
use tauri::{Emitter, Manager};

pub(crate) const WINDOW_LABEL: &str = "anywhere";
/// tao builds windows with this title as non-activating panels.
pub(crate) const PANEL_TITLE: &str = "VisualTeX Anywhere";
#[cfg(target_os = "macos")]
const ACTIVATE_EVENT: &str = "visualtex-anywhere-activate";
#[cfg(target_os = "macos")]
const SUSPEND_EVENT: &str = "visualtex-anywhere-suspend";
/// A key the input method took from the field, for the editor.
#[cfg(target_os = "macos")]
const KEY_EVENT: &str = "visualtex-anywhere-key";
/// The field's text size or colour, learnt after the formula opened.
#[cfg(target_os = "macos")]
const STYLE_EVENT: &str = "visualtex-anywhere-style";
/// Pinyin spelled in the formula's text: the editor's composition.
#[cfg(target_os = "macos")]
const COMPOSE_EVENT: &str = "visualtex-anywhere-compose";
/// The shortcut switched the field to the input method; a field that has
/// not activated it by then does not take text (or the switch failed).
#[cfg(target_os = "macos")]
const INPUT_METHOD_WAIT: Duration = Duration::from_millis(900);
/// The input method reconnects to VisualTeX within a second of being started
/// or of VisualTeX starting (input-method/Sources/AppLink.swift).
#[cfg(target_os = "macos")]
const IME_RELINK_WAIT: Duration = Duration::from_millis(1500);
/// Keys typed before the card is up wait for it, up to this many.
#[cfg(target_os = "macos")]
const MAX_PENDING_KEYS: usize = 256;
/// A session shows the panel when the strip reports its layout, which the
/// page does once it has rendered the activation. A page that has not
/// answered after this long is broken (a render error that unmounted it, a
/// lost WebContent process) and is reloaded; an invisible panel that takes
/// the keyboard is never shown.
#[cfg(target_os = "macos")]
const PAGE_ANSWER_WAIT: Duration = Duration::from_millis(600);
/// What a line of the field's text is to its font size (Chrome reports the line box).
#[cfg(target_os = "macos")]
const LINE_TO_FONT: f64 = 1.19;
/// The shortest believable line of a field's text, points (a 9 pt font's): a
/// shorter rectangle is not the text's and does not size the formula.
#[cfg(target_os = "macos")]
pub(super) const MIN_TEXT_HEIGHT: f64 = 11.0;
/// The tallest, points.
#[cfg(target_os = "macos")]
const MAX_TEXT_HEIGHT: f64 = 96.0;
/// How far the anchor's height may be from the field's first report and
/// still size the formula (an empty line first reports its whole line box,
/// taller than its text); outside this only its place is used.
#[cfg(target_os = "macos")]
const ANCHOR_RATIO: std::ops::RangeInclusive<f64> = 0.6..=1.6;
/// The system's text size (NSFont.systemFontSize), points: what the formula
/// is sized from when the field tells nothing believable.
#[cfg(target_os = "macos")]
const SYSTEM_TEXT_SIZE: f64 = 13.0;

/// A rectangle in global screen coordinates as Accessibility and Quartz
/// report them: points, origin at the top left of the main display, y down.
#[derive(Clone, Copy, Debug, Default, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ScreenRect {
    pub(crate) x: f64,
    pub(crate) y: f64,
    pub(crate) width: f64,
    pub(crate) height: f64,
}

/// What the strip stands on: the text caret, the focused field when its
/// application does not report the caret, or the mouse pointer.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum AnchorSource {
    Caret,
    Field,
    Pointer,
}

/// The application that owned the caret when the shortcut was pressed.
#[derive(Clone, Debug)]
pub(crate) struct Target {
    pub(crate) app_name: String,
    pub(crate) anchor: ScreenRect,
    pub(crate) source: AnchorSource,
}

/// The strip's size and the point of it that sits on the caret line (where
/// the formula starts, at its vertical centre), in points from its top left.
#[derive(Clone, Copy, Debug, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PanelLayout {
    pub(crate) width: f64,
    pub(crate) height: f64,
    pub(crate) anchor_x: f64,
    pub(crate) anchor_y: f64,
}

impl PanelLayout {
    /// An empty strip, until the page measures its own.
    const EMPTY: Self = Self {
        width: 96.0,
        height: 44.0,
        anchor_x: 14.0,
        anchor_y: 22.0,
    };

    fn is_valid(&self) -> bool {
        let finite = [self.width, self.height, self.anchor_x, self.anchor_y]
            .iter()
            .all(|value| value.is_finite());
        finite
            && (16.0..=4096.0).contains(&self.width)
            && (16.0..=4096.0).contains(&self.height)
            && (0.0..=self.width).contains(&self.anchor_x)
            && (0.0..=self.height).contains(&self.anchor_y)
    }
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct Activation {
    generation: u64,
    /// The caret's height in points: the host's line height, which sets the
    /// formula's size.
    caret_height: f64,
    anchor_source: AnchorSource,
    app_name: String,
    accessibility_trusted: bool,
    /// The previous session hid without a commit or cancel: keep its formula.
    resume: bool,
    /// VisualTeX's input method types the formula: the editor takes its keys
    /// from the input method, and the formula is drawn in the field's line.
    input_method: bool,
    /// The display shortcut opened the formula: its row is a display row
    /// (else a new formula starts inline).
    display: bool,
    /// The field's font size in points, when the field tells it.
    font_size: Option<f64>,
    /// The field's text colour (#rrggbbaa), when the field tells it.
    text_color: Option<String>,
    /// The field font's x-height over its size.
    x_height: Option<f64>,
    /// The field's line (its caret) in points: height, and how far it
    /// reaches below the baseline.
    line_height: Option<f64>,
    line_descent: Option<f64>,
    /// The field's background (#rrggbbaa), as read off the screen.
    background_color: Option<String>,
}

/// How the field's text looks where the formula goes (input method).
#[cfg(target_os = "macos")]
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
    /// Pasted into the application that owned the caret.
    Inserted,
    /// Without the Accessibility permission VisualTeX cannot paste for the
    /// user: the formula is on the clipboard and the strip says so.
    Copied,
    /// The session ended (another shortcut press, a newer session) first.
    Stale,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnywhereStatus {
    accessibility_trusted: bool,
    shortcut: String,
}

#[cfg(target_os = "macos")]
struct Session {
    generation: u64,
    target: Option<Target>,
    layout: PanelLayout,
    /// The page has mounted its editor and listens for activations.
    ready: bool,
    phase: Phase,
    /// The strip still holds the formula of a session that hid without a
    /// commit or cancel.
    draft: bool,
    /// The generation whose unanswered activation reloaded the page: one
    /// reload per session.
    reloaded_for: u64,
    delivery: Delivery,
    /// The input source the field goes back to after the formula (the
    /// shortcut switched it to the input method).
    return_to: Option<String>,
    /// The display shortcut opened this session.
    display: bool,
    /// The generation whose shortcut switched the field to the input method
    /// and waits for its "activate".
    awaiting_input_method: u64,
    /// Keys the input method took before the card was up.
    pending_keys: Vec<serde_json::Value>,
    /// The field's text style where the formula goes.
    field_text: FieldText,
    /// The height of the field's text (points) that sized this formula first:
    /// the anchor's height is checked against it. None: the field told
    /// nothing believable (the system's text size is used until it does).
    reference_height: Option<f64>,
    /// Where the formula is drawn in the field's line: x at the formula's
    /// start (the input method's marked text), y on the line's baseline.
    inline_anchor: Option<ScreenRect>,
    /// The generation whose field colour was read off the screen (once).
    color_read: u64,
    /// The generation whose formula colour comes from the screen: it stays,
    /// whatever the field reports later.
    color_from_screen: u64,
}

/// How the formula reaches the application. Paste: the strip sits on the
/// caret line and Enter pastes (Accessibility). InputMethod: VisualTeX's
/// input method has the field, which shows the formula's text as marked
/// text; the card sits under the line and Enter inserts through the input
/// method.
#[cfg(target_os = "macos")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Delivery {
    Paste,
    InputMethod,
}

/// Hidden: nothing to show. Activating: the shortcut was pressed and the
/// strip is laying out; its layout report shows the panel. Presented: the
/// strip is up. Only an activation leads to Presented, so a late layout
/// report never brings back a strip that was closed.
#[cfg(target_os = "macos")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Phase {
    Hidden,
    Activating,
    Presented,
}

#[cfg(target_os = "macos")]
static SESSION: Mutex<Session> = Mutex::new(Session {
    generation: 0,
    target: None,
    layout: PanelLayout::EMPTY,
    ready: false,
    phase: Phase::Hidden,
    draft: false,
    reloaded_for: 0,
    delivery: Delivery::Paste,
    return_to: None,
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
    inline_anchor: None,
    color_read: 0,
    color_from_screen: 0,
});

#[cfg(target_os = "macos")]
static LOG_PATH: std::sync::OnceLock<std::path::PathBuf> = std::sync::OnceLock::new();
#[cfg(target_os = "macos")]
const MAX_LOG_BYTES: u64 = 512 * 1024;

/// One line of the diagnostics log (app data, anywhere.log): timings, where
/// the strip stood and what each commit did. Formulas are not logged.
#[cfg(target_os = "macos")]
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

#[cfg(target_os = "macos")]
fn session() -> Result<std::sync::MutexGuard<'static, Session>, String> {
    SESSION
        .lock()
        .map_err(|_| "VisualTeX Anywhere state is unavailable".to_string())
}

/// At startup (main thread): register the shortcut and prewarm the panel.
pub(crate) fn initialize(app: &AppHandle) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        if let Ok(directory) = app.path().app_data_dir() {
            let _ = std::fs::create_dir_all(&directory);
            let _ = LOG_PATH.set(directory.join("anywhere.log"));
        }
        hotkey::register(app)?;
        panel::ensure(app)?;
        let installer = app.clone();
        let _ = std::thread::Builder::new()
            .name("visualtex-ime-install".to_string())
            .spawn(move || {
                if let Err(error) = ime::start_link(&installer) {
                    log(&format!("input method link unavailable: {error}"));
                }
                if let Err(error) = ime::install(&installer) {
                    log(&format!("input method not installed: {error}"));
                }
                // The formula takes the field's colour from the screen when
                // the field does not tell it (screen_colors): with the input
                // method in use, macOS is asked once for Screen Recording.
                if ime::is_enabled(&installer) && !screen_colors::can_read_screen() {
                    let _ = installer.run_on_main_thread(|| {
                        let asked = screen_colors::request_access();
                        log(&format!("screen recording requested granted={asked}"));
                    });
                }
            });
        log("started");
        Ok(())
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = app;
        Ok(())
    }
}

/// A shortcut: show the strip at the caret (`display`: the formula's row is
/// a display row), or hide it (keeping its formula) when it is up, as
/// switching an input method off keeps nothing half-typed in the document.
#[cfg(target_os = "macos")]
fn toggle(app: &AppHandle, display: bool) {
    if session().map(|session| session.phase != Phase::Hidden).unwrap_or(false) {
        hide(app, true, "shortcut");
        return;
    }
    log(&format!("shortcut display={display}"));
    let app = app.clone();
    let spawned = std::thread::Builder::new()
        .name("visualtex-anywhere-activate".to_string())
        .spawn(move || {
            if ime::is_enabled(&app) {
                begin_input_method(&app, display);
            } else {
                activate(&app, display);
            }
        });
    if let Err(error) = spawned {
        eprintln!("Unable to start VisualTeX Anywhere: {error}");
    }
}

#[cfg(target_os = "macos")]
fn activate(app: &AppHandle, display: bool) {
    let started = std::time::Instant::now();
    let trusted = ax::is_trusted();
    if !trusted {
        ax::request_trust();
    }
    let target = ax::capture_target(trusted);
    log(&format!(
        "activate trusted={trusted} source={:?} anchor={:.0},{:.0},{:.0} app={:?} capture_ms={}",
        target.source,
        target.anchor.x,
        target.anchor.y,
        target.anchor.height,
        target.app_name,
        started.elapsed().as_millis()
    ));
    let window = match panel::ensure(app) {
        Ok(window) => window,
        Err(error) => {
            eprintln!("VisualTeX Anywhere panel is unavailable: {error}");
            return;
        }
    };
    let (activation, ready) = {
        let Ok(mut session) = session() else { return };
        session.generation += 1;
        session.phase = Phase::Activating;
        session.target = Some(target.clone());
        session.delivery = Delivery::Paste;
        session.return_to = None;
        session.display = display;
        session.awaiting_input_method = 0;
        session.pending_keys.clear();
        session.field_text = FieldText::default();
        session.reference_height = None;
        session.inline_anchor = None;
        (pending_activation(&session, trusted), session.ready)
    };
    // A cold (or reloading) panel takes the activation when its page reports
    // ready.
    if let (Some(activation), true) = (activation, ready) {
        send_activation(app, &window, activation);
    }
}

/// The activation of the current session, for its page.
#[cfg(target_os = "macos")]
fn pending_activation(session: &Session, trusted: bool) -> Option<Activation> {
    let target = session.target.as_ref()?;
    Some(Activation {
        generation: session.generation,
        caret_height: target.anchor.height,
        anchor_source: target.source,
        app_name: target.app_name.clone(),
        accessibility_trusted: trusted,
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

/// The shortcut with the input method enabled: switch the field to it; its
/// "activate" (input_method_message) opens the formula there.
#[cfg(target_os = "macos")]
fn begin_input_method(app: &AppHandle, display: bool) {
    let started = std::time::Instant::now();
    let previous = ime::current_source(app);
    let already = previous.as_deref() == Some(ime::SOURCE_ID);
    let generation = {
        let Ok(mut session) = session() else { return };
        session.generation += 1;
        session.phase = Phase::Activating;
        session.target = None;
        session.delivery = Delivery::InputMethod;
        session.return_to = previous.clone().filter(|_| !already);
        session.display = display;
        session.awaiting_input_method = session.generation;
        session.pending_keys.clear();
        session.field_text = FieldText::default();
        session.reference_height = None;
        session.inline_anchor = None;
        session.generation
    };
    if let Err(error) = panel::ensure(app) {
        log(&format!("input method session gen={generation}: no panel: {error}"));
        return;
    }
    let switched = if already {
        // Typing with VisualTeX already (chosen in the input menu): its
        // field answers with the caret. The input method may be relinking
        // (VisualTeX restarted, or macOS has just started it for the
        // shortcut's own modifier keys): it reconnects within a second.
        let deadline = std::time::Instant::now() + IME_RELINK_WAIT;
        while !ime::is_linked() && std::time::Instant::now() < deadline {
            std::thread::sleep(Duration::from_millis(40));
        }
        ime::send(serde_json::json!({ "t": "caret" }))
    } else {
        match ime::switch_in(app, previous.as_deref().unwrap_or("com.apple.keylayout.ABC")) {
            Ok(()) => true,
            Err(error) => {
                log(&format!("input method switch failed gen={generation}: {error}"));
                false
            }
        }
    };
    log(&format!(
        "input method requested gen={generation} from={} switched={switched} linked={} ms={}",
        previous.as_deref().unwrap_or("?"),
        ime::is_linked(),
        started.elapsed().as_millis()
    ));
    if !switched {
        fall_back_to_paste(app, generation, false, display);
        return;
    }
    let app = app.clone();
    let _ = std::thread::Builder::new()
        .name("visualtex-ime-wait".to_string())
        .spawn(move || {
            std::thread::sleep(INPUT_METHOD_WAIT);
            let waiting = session()
                .map(|session| session.awaiting_input_method == generation && session.generation == generation)
                .unwrap_or(false);
            if waiting {
                log(&format!("input method did not take the field gen={generation}"));
                fall_back_to_paste(&app, generation, !already, display);
            }
        });
}

/// The field did not take the input method: back to the input source it
/// had, and the pasting strip instead.
#[cfg(target_os = "macos")]
fn fall_back_to_paste(app: &AppHandle, generation: u64, switch_back: bool, display: bool) {
    {
        let Ok(mut session) = session() else { return };
        if session.generation != generation {
            return;
        }
        session.awaiting_input_method = 0;
        session.phase = Phase::Hidden;
        session.delivery = Delivery::Paste;
        session.return_to = None;
    }
    if switch_back {
        ime::switch_back();
    }
    activate(app, display);
}

/// The input method's report of the field: where the formula goes (the
/// marked text, or the caret's line before there is any) and how its text
/// looks there.
#[cfg(target_os = "macos")]
struct FieldGeometry {
    /// Quartz coordinates (y down).
    caret: ScreenRect,
    /// The line's baseline, Quartz y.
    baseline: f64,
    /// How the field's text looks (its size fields None when the report is
    /// not believable).
    text: FieldText,
    /// The height of the field's text in points, when believable (a line of
    /// text tall: MIN_TEXT_HEIGHT, MAX_TEXT_HEIGHT).
    measured: Option<f64>,
    /// The report, for anywhere.log.
    report: String,
}

/// A rectangle from the input method (AppKit screen coordinates, y up) in
/// the Quartz coordinates the panel code uses, with the baseline from the
/// field's font descent.
#[cfg(target_os = "macos")]
fn input_method_geometry(message: &serde_json::Value) -> Option<FieldGeometry> {
    let values: Vec<f64> = message.get("caret")?.as_array()?.iter().filter_map(|value| value.as_f64()).collect();
    let &[x, y, width, height] = values.as_slice() else { return None };
    if !(x.is_finite() && y.is_finite() && height.is_finite()) || (x == 0.0 && y == 0.0 && height == 0.0) || height <= 0.0 {
        return None;
    }
    // A rectangle far off any screen is a stale answer (the input method
    // drops those; this keeps an old one from moving the formula away).
    if x.abs() > 20000.0 || y.abs() > 20000.0 || height > 400.0 {
        return None;
    }
    let number = |key: &str| message.get(key).and_then(|value| value.as_f64()).filter(|value| value.is_finite() && *value > 0.0);
    let reported_size = number("fontSize").filter(|size| (4.0..=200.0).contains(size));
    // Chrome and Electron report a default 12 pt font, not the page's: a
    // line much taller than that font's own height says so. Their font is
    // then taken from the height of a line of their text, which an empty
    // line does not give (it reports its whole line box): the height the
    // input method learnt in this application where the line had text.
    let reported_box = match (number("ascent"), number("descent"), reported_size) {
        (Some(ascent), Some(descent), _) => ascent + descent,
        (_, _, Some(size)) => size * 1.15,
        _ => 0.0,
    };
    let field_tells_font = reported_box > 0.0 && height < reported_box * 1.25;
    let text_height = if field_tells_font {
        height
    } else {
        let text_before = message.get("textBefore").and_then(|value| value.as_bool()).unwrap_or(true);
        match number("textHeight") {
            Some(learnt) if !text_before && learnt < height && height <= learnt * 1.32 => learnt,
            _ => height,
        }
    };
    // A rectangle shorter or taller than a line of text places the formula
    // without sizing it.
    let measured = (MIN_TEXT_HEIGHT..=MAX_TEXT_HEIGHT).contains(&text_height).then_some(text_height);
    let font_size = measured.and_then(|_| {
        if field_tells_font {
            reported_size
        } else {
            Some((text_height / LINE_TO_FONT).clamp(9.0, 64.0))
        }
    });
    let scale = match (font_size, reported_size) {
        (Some(size), Some(reported)) if reported > 0.0 => size / reported,
        _ => 1.0,
    };
    let descent = number("descent")
        .map(|descent| descent * scale)
        .filter(|descent| *descent < text_height)
        .unwrap_or(text_height * 0.22);
    let main_height = panel::main_display_height();
    let top = main_height - (y + height);
    // Text sits in the middle of a taller line box.
    let below_text = (height - text_height) / 2.0;
    Some(FieldGeometry {
        caret: ScreenRect {
            x,
            y: top,
            width,
            height: height.clamp(10.0, 64.0),
        },
        baseline: top + height - below_text - descent,
        text: FieldText {
            font_size,
            text_color: message
                .get("color")
                .and_then(|value| value.as_str())
                .filter(|color| color.len() == 9 && color.starts_with('#'))
                .map(str::to_string),
            x_height: number("xHeight").filter(|ratio| (0.3..=0.8).contains(ratio)),
            line_height: measured,
            line_descent: measured.map(|_| descent),
            background_color: None,
        },
        measured,
        report: format!(
            "caret={x:.0},{y:.0},{height:.0} font={} fieldFont={field_tells_font} textBefore={} learnt={} text={text_height:.1}pt",
            reported_size.map(|value| format!("{value:.1}")).unwrap_or_else(|| "-".to_string()),
            message.get("textBefore").and_then(|value| value.as_bool()).unwrap_or(true),
            number("textHeight").map(|value| format!("{value:.0}")).unwrap_or_else(|| "-".to_string())
        ),
    })
}

/// How text whose line (its caret) is `line` points tall looks: its font
/// size, the line and how far it reaches below the baseline.
#[cfg(target_os = "macos")]
fn text_of_line(line: f64) -> FieldText {
    FieldText {
        font_size: Some((line / LINE_TO_FONT).clamp(9.0, 64.0)),
        line_height: Some(line),
        line_descent: Some(line * 0.22),
        ..FieldText::default()
    }
}

/// The field's text where the formula opens: as the field tells it when the
/// report is believable, else the system's text size in the field's colour.
#[cfg(target_os = "macos")]
fn opening_text(geometry: &FieldGeometry) -> FieldText {
    if geometry.measured.is_some() {
        return geometry.text.clone();
    }
    FieldText {
        text_color: geometry.text.text_color.clone(),
        x_height: geometry.text.x_height,
        ..text_of_line(SYSTEM_TEXT_SIZE * LINE_TO_FONT)
    }
}

/// Whether the anchor's height (points, None: not believable) sizes the
/// formula, given the height that sized it first (None: the system's).
#[cfg(target_os = "macos")]
fn anchor_sizes(measured: Option<f64>, reference: Option<f64>) -> bool {
    match (measured, reference) {
        (Some(height), Some(reference)) => ANCHOR_RATIO.contains(&(height / reference)),
        (Some(_), None) => true,
        (None, _) => false,
    }
}

#[cfg(target_os = "macos")]
fn describe_text(text: &FieldText) -> String {
    format!(
        "font={} line={}",
        text.font_size.map(|value| format!("{value:.1}")).unwrap_or_else(|| "-".to_string()),
        text.line_height.map(|value| format!("{value:.1}")).unwrap_or_else(|| "-".to_string())
    )
}

#[cfg(target_os = "macos")]
fn inline_anchor(geometry: &FieldGeometry) -> ScreenRect {
    ScreenRect {
        x: geometry.caret.x,
        y: geometry.baseline,
        width: geometry.caret.width,
        height: 0.0,
    }
}

/// A message from the input method (its link thread).
#[cfg(target_os = "macos")]
pub(crate) fn input_method_message(app: &AppHandle, message: &serde_json::Value) {
    let kind = message.get("t").and_then(|value| value.as_str()).unwrap_or("");
    match kind {
        "hello" => log(&format!(
            "input method {}",
            message.get("version").and_then(|value| value.as_str()).unwrap_or("?")
        )),
        "activate" | "geometry" => input_method_caret_reported(app, message),
        "start" => input_method_started(app, message),
        "key" => input_method_key(app, message),
        "compose" => {
            let generation = session()
                .ok()
                .filter(|session| session.phase == Phase::Presented && session.delivery == Delivery::InputMethod)
                .map(|session| session.generation);
            if let (Some(generation), Some(window)) = (generation, app.get_webview_window(WINDOW_LABEL)) {
                let mut payload = message.clone();
                if let Some(object) = payload.as_object_mut() {
                    object.insert("generation".to_string(), serde_json::json!(generation));
                }
                let _ = window.emit(COMPOSE_EVENT, payload);
            }
        }
        "deactivate" | "disconnected" => {
            let open = session()
                .map(|session| session.phase != Phase::Hidden && session.delivery == Delivery::InputMethod)
                .unwrap_or(false);
            if open {
                hide_with(app, true, if kind == "deactivate" { "input method left the field" } else { "input method gone" }, false);
            }
        }
        "forced" => {
            // The field ended the composition and kept its text.
            let open = session()
                .map(|session| session.phase != Phase::Hidden && session.delivery == Delivery::InputMethod)
                .unwrap_or(false);
            if open {
                hide_with(app, false, "field kept the text", false);
            }
        }
        _ => {}
    }
}

/// The field took the input method (or answered with its geometry): open
/// the formula the shortcut asked for, or follow the formula's place.
#[cfg(target_os = "macos")]
fn input_method_caret_reported(app: &AppHandle, message: &serde_json::Value) {
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
        if session.delivery != Delivery::InputMethod {
            Next::Nothing
        } else if session.awaiting_input_method == session.generation && session.phase == Phase::Activating {
            session.awaiting_input_method = 0;
            let anchor = geometry
                .as_ref()
                .map(|geometry| geometry.caret)
                .unwrap_or_else(|| ax::capture_target(ax::is_trusted()).anchor);
            session.field_text = geometry.as_ref().map(opening_text).unwrap_or_default();
            session.reference_height = geometry.as_ref().and_then(|geometry| geometry.measured);
            if let Some(geometry) = geometry.as_ref() {
                sizing = Some(format!(
                    "size gen={} {kind} {} -> {} ({})",
                    session.generation,
                    geometry.report,
                    describe_text(&session.field_text),
                    if geometry.measured.is_some() { "field" } else { "not believable: system text size" }
                ));
            }
            session.inline_anchor = geometry.as_ref().map(inline_anchor);
            session.target = Some(Target {
                app_name: client.clone(),
                anchor,
                source: if geometry.is_some() { AnchorSource::Caret } else { AnchorSource::Field },
            });
            Next::Open(pending_activation(&session, true), session.ready)
        } else if let (Phase::Presented | Phase::Activating, Some(geometry)) = (session.phase, geometry.as_ref()) {
            if let Some(target) = session.target.as_mut() {
                target.anchor = geometry.caret;
            }
            let anchor = inline_anchor(geometry);
            session.inline_anchor = Some(anchor);
            // A style the field tells only now (its font with the marked text).
            // The anchor's height sizes the formula only when it is believable
            // and near the height that sized it first.
            let sizes = anchor_sizes(geometry.measured, session.reference_height);
            let mut text = session.field_text.clone();
            let reported = &geometry.text;
            if sizes {
                session.reference_height = session.reference_height.or(geometry.measured);
                text.font_size = reported.font_size.or(text.font_size);
                text.x_height = reported.x_height.or(text.x_height);
                text.line_height = reported.line_height.or(text.line_height);
                text.line_descent = reported.line_descent.or(text.line_descent);
            }
            if session.color_from_screen != session.generation {
                text.text_color = reported.text_color.clone().or(text.text_color);
            }
            text.background_color = reported.background_color.clone().or(text.background_color);
            let restyle = (text != session.field_text).then(|| (session.generation, text.clone()));
            sizing = Some(format!(
                "size gen={} anchor {} reference={} -> {}{}",
                session.generation,
                geometry.report,
                session.reference_height.map(|value| format!("{value:.1}pt")).unwrap_or_else(|| "-".to_string()),
                if sizes { "sizes" } else { "places only" },
                restyle
                    .as_ref()
                    .filter(|(_, new)| new.font_size != session.field_text.font_size || new.line_height != session.field_text.line_height)
                    .map(|(_, new)| format!(", {} -> {}", describe_text(&session.field_text), describe_text(new)))
                    .unwrap_or_default()
            ));
            // The formula's colour from the field's background, once the
            // formula's place is known (if it was not read when it showed).
            let read_color = (session.color_read != session.generation && session.phase == Phase::Presented)
                .then(|| {
                    session.color_read = session.generation;
                    (session.generation, geometry.caret)
                });
            session.field_text = text;
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
    };
    if let Some(line) = sizing {
        log(&line);
    }
    match next {
        Next::Open(activation, ready) => {
            ime::send(serde_json::json!({ "t": "session", "on": true }));
            log(&format!(
                "input method session open client={client:?} caret={} font={:?} color={:?}",
                geometry
                    .as_ref()
                    .map(|geometry| format!("{:.0},{:.0},{:.0} baseline={:.1}", geometry.caret.x, geometry.caret.y, geometry.caret.height, geometry.baseline))
                    .unwrap_or_else(|| "-".to_string()),
                geometry.as_ref().and_then(|geometry| geometry.text.font_size),
                geometry.as_ref().and_then(|geometry| geometry.text.text_color.clone()),
            ));
            if let (Some(activation), true, Some(window)) = (activation, ready, app.get_webview_window(WINDOW_LABEL)) {
                send_activation(app, &window, activation);
            }
        }
        Next::Follow(anchor, layout, restyle) => {
            if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
                panel::reframe(&window, anchor, layout, panel::Placement::InLine);
                if let Some((generation, text)) = restyle {
                    let mut payload = serde_json::to_value(&text).unwrap_or_default();
                    if let Some(object) = payload.as_object_mut() {
                        object.insert("generation".to_string(), serde_json::json!(generation));
                    }
                    let _ = window.emit(STYLE_EVENT, payload);
                }
            }
        }
        Next::Nothing => {}
    }
}

/// Reads the formula's colour off the screen (screen_colors: from the
/// field's background, so it does not depend on the text before the caret)
/// and gives it to the page; without the permission the field's own colour
/// or the system appearance's stays.
#[cfg(target_os = "macos")]
fn read_field_color(app: &AppHandle, generation: u64, caret: ScreenRect) {
    let app = app.clone();
    let _ = std::thread::Builder::new()
        .name("visualtex-anywhere-color".to_string())
        .spawn(move || {
            let Some(window) = app.get_webview_window(WINDOW_LABEL) else { return };
            let Some(number) = panel::window_number(&window) else { return };
            let colors = screen_colors::field_colors(caret, number);
            log(&format!(
                "field colour gen={generation} screen={} text={:?} background={:?}",
                screen_colors::can_read_screen(),
                colors.as_ref().map(|colors| colors.text.as_str()),
                colors.as_ref().map(|colors| colors.background.as_str()),
            ));
            let Some(colors) = colors else { return };
            let text = {
                let Ok(mut session) = session() else { return };
                if session.generation != generation {
                    return;
                }
                session.color_from_screen = generation;
                session.field_text.text_color = Some(colors.text);
                session.field_text.background_color = Some(colors.background);
                session.field_text.clone()
            };
            let mut payload = serde_json::to_value(&text).unwrap_or_default();
            if let Some(object) = payload.as_object_mut() {
                object.insert("generation".to_string(), serde_json::json!(generation));
            }
            let _ = window.emit(STYLE_EVENT, payload);
        });
}

/// VisualTeX was chosen in the input menu and a character key started a
/// formula: a session without an input source to go back to.
#[cfg(target_os = "macos")]
fn input_method_started(app: &AppHandle, message: &serde_json::Value) {
    let geometry = input_method_geometry(message);
    let client = message.get("client").and_then(|value| value.as_str()).unwrap_or("").to_string();
    let (activation, ready) = {
        let Ok(mut session) = session() else { return };
        if session.phase != Phase::Hidden {
            return;
        }
        session.generation += 1;
        session.phase = Phase::Activating;
        session.delivery = Delivery::InputMethod;
        session.return_to = None;
        session.display = false;
        session.awaiting_input_method = 0;
        session.pending_keys.clear();
        session.field_text = geometry.as_ref().map(opening_text).unwrap_or_default();
        session.reference_height = geometry.as_ref().and_then(|geometry| geometry.measured);
        session.inline_anchor = geometry.as_ref().map(inline_anchor);
        session.target = Some(Target {
            app_name: client,
            anchor: geometry
                .as_ref()
                .map(|geometry| geometry.caret)
                .unwrap_or_else(|| ax::capture_target(ax::is_trusted()).anchor),
            source: AnchorSource::Caret,
        });
        (pending_activation(&session, true), session.ready)
    };
    if let Err(error) = panel::ensure(app) {
        log(&format!("input method start: no panel: {error}"));
        return;
    }
    if let (Some(activation), true, Some(window)) = (activation, ready, app.get_webview_window(WINDOW_LABEL)) {
        send_activation(app, &window, activation);
    }
}

/// A key the input method took: to the editor once the card is up (until
/// then it waits, so fast typing after the shortcut is not lost).
#[cfg(target_os = "macos")]
fn input_method_key(app: &AppHandle, message: &serde_json::Value) {
    let mut key = message.clone();
    let deliver_now = {
        let Ok(mut session) = session() else { return };
        if session.delivery != Delivery::InputMethod || session.phase == Phase::Hidden {
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

/// The card is up: the keys typed before it, in order.
#[cfg(target_os = "macos")]
fn flush_pending_keys(app: &AppHandle, generation: u64) {
    let keys = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation {
            return;
        }
        std::mem::take(&mut session.pending_keys)
    };
    if keys.is_empty() {
        return;
    }
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        for key in keys {
            let _ = window.emit(KEY_EVENT, key);
        }
    }
}

/// Hands the session to the page and watches for its answer (the strip's
/// layout report, which shows the panel).
#[cfg(target_os = "macos")]
fn send_activation(app: &AppHandle, window: &tauri::WebviewWindow, activation: Activation) {
    let generation = activation.generation;
    if let Err(error) = window.emit(ACTIVATE_EVENT, activation) {
        log(&format!("activation gen={generation} not sent: {error}"));
        return;
    }
    let app = app.clone();
    let watched = std::thread::Builder::new()
        .name("visualtex-anywhere-watch".to_string())
        .spawn(move || {
            std::thread::sleep(PAGE_ANSWER_WAIT);
            recover_unanswered(&app, generation);
        });
    if let Err(error) = watched {
        eprintln!("Unable to watch VisualTeX Anywhere: {error}");
    }
}

/// The page did not answer an activation: reload it once (the activation is
/// sent again when it reports ready), or give the session up.
#[cfg(target_os = "macos")]
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
        panel::park_for_reload(&window);
        if let Err(error) = window.reload() {
            log(&format!("page reload failed: {error}"));
        }
    } else {
        log(&format!("page did not answer gen={generation} after a reload: session dropped"));
        panel::park(&window);
    }
}

/// Orders the panel in at the caret, once per session.
#[cfg(target_os = "macos")]
fn present(app: &AppHandle, generation: u64) {
    let (target, layout, delivery, inline, read_color) = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation || session.phase != Phase::Activating || !session.ready {
            return;
        }
        let Some(target) = session.target.clone() else { return };
        session.phase = Phase::Presented;
        // The formula's colour from the field's background (screen_colors).
        let read_color = session.delivery == Delivery::InputMethod && session.color_read != generation;
        if read_color {
            session.color_read = generation;
        }
        (target, session.layout, session.delivery, session.inline_anchor, read_color)
    };
    log(&format!(
        "present gen={generation} {delivery:?} layout={:.0}x{:.0} anchor_in_strip={:.0},{:.0}",
        layout.width, layout.height, layout.anchor_x, layout.anchor_y
    ));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        match delivery {
            Delivery::Paste => panel::present(&window, target.anchor, layout),
            Delivery::InputMethod => match inline {
                Some(anchor) => panel::present_passive(&window, anchor, layout, panel::Placement::InLine),
                None => panel::present_passive(&window, target.anchor, layout, panel::Placement::BelowLine),
            },
        }
    }
    if delivery == Delivery::InputMethod {
        flush_pending_keys(app, generation);
    }
    if read_color {
        read_field_color(app, generation, target.anchor);
    }
}

#[cfg(target_os = "macos")]
fn hide(app: &AppHandle, keep_draft: bool, reason: &str) {
    hide_with(app, keep_draft, reason, true);
}

/// Hides the strip. With the input method, `tell_input_method` clears the
/// field's marked text and hands it back to the input source it came from
/// (not when the input method itself ended the session).
#[cfg(target_os = "macos")]
fn hide_with(app: &AppHandle, keep_draft: bool, reason: &str, tell_input_method: bool) {
    let (generation, delivery, return_to) = {
        let Ok(mut session) = session() else { return };
        if session.phase == Phase::Hidden {
            return;
        }
        session.phase = Phase::Hidden;
        session.draft = keep_draft;
        session.awaiting_input_method = 0;
        session.pending_keys.clear();
        (session.generation, session.delivery, session.return_to.take())
    };
    if delivery == Delivery::InputMethod && tell_input_method {
        ime::send(serde_json::json!({ "t": "cancel", "returnTo": return_to }));
    }
    log(&format!("hide gen={generation} keep_draft={keep_draft} ({reason})"));
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        panel::park(&window);
        let _ = window.emit(SUSPEND_EVENT, Suspension { generation, keep_draft });
    }
}

/// The panel lost the keyboard (the user clicked elsewhere): hide it and keep
/// the formula for the next shortcut press.
#[cfg(target_os = "macos")]
pub(crate) fn handle_focus_lost(app: &AppHandle) {
    hide(app, true, "focus lost");
}

/// The page mounted its editor: from now on it takes activations directly.
#[tauri::command]
pub(crate) fn anywhere_ready(app: AppHandle) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        let (pending, phase) = {
            let mut session = session()?;
            session.ready = true;
            let pending = match session.phase {
                Phase::Hidden => None,
                // The shortcut was pressed while the page was loading or
                // reloading, or the page remounted after an error while the
                // strip was up: it starts the session over.
                Phase::Activating | Phase::Presented => {
                    session.draft = false;
                    pending_activation(&session, ax::is_trusted())
                }
            };
            (pending, session.phase)
        };
        log(&format!("page ready phase={phase:?}"));
        let window = app
            .get_webview_window(WINDOW_LABEL)
            .ok_or_else(|| "VisualTeX Anywhere panel is unavailable".to_string())?;
        match pending {
            Some(activation) => send_activation(&app, &window, activation),
            None => panel::park(&window),
        }
        Ok(())
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = app;
        Ok(())
    }
}

/// The strip measured itself: show the panel there, or follow its new size.
#[tauri::command]
pub(crate) fn anywhere_layout(app: AppHandle, generation: u64, layout: PanelLayout) -> Result<(), String> {
    if !layout.is_valid() {
        return Err("The VisualTeX Anywhere layout is invalid".to_string());
    }
    #[cfg(target_os = "macos")]
    {
        let (phase, anchor, placement) = {
            let mut session = session()?;
            if session.generation != generation {
                return Ok(());
            }
            session.layout = layout;
            let caret = session.target.as_ref().map(|target| target.anchor);
            match (session.delivery, session.inline_anchor) {
                (Delivery::Paste, _) => (session.phase, caret, panel::Placement::OnLine),
                (Delivery::InputMethod, Some(inline)) => (session.phase, Some(inline), panel::Placement::InLine),
                (Delivery::InputMethod, None) => (session.phase, caret, panel::Placement::BelowLine),
            }
        };
        match (phase, anchor, app.get_webview_window(WINDOW_LABEL)) {
            (Phase::Activating, _, _) => present(&app, generation),
            (Phase::Presented, Some(anchor), Some(window)) => panel::reframe(&window, anchor, layout, placement),
            _ => {}
        }
        Ok(())
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (app, generation);
        Ok(())
    }
}

/// Enter in the strip: type the formula into the application that owned the
/// caret. One call is one delivery attempt.
/// The Image output mode: the formula as a PNG, pasted into the
/// application (with the input method, its anchor is dropped and the field
/// goes back to its input source first).
#[tauri::command]
pub(crate) async fn anywhere_commit_image(app: AppHandle, generation: u64, data_base64: String) -> Result<CommitOutcome, String> {
    use base64::Engine as _;
    let png = base64::engine::general_purpose::STANDARD
        .decode(data_base64.trim())
        .map_err(|error| format!("The formula image could not be read: {error}"))?;
    if png.len() > 50 * 1024 * 1024 || !png.starts_with(b"\x89PNG\r\n\x1a\n") {
        return Err("The formula image is not a PNG".to_string());
    }
    #[cfg(target_os = "macos")]
    {
        let input_method = {
            let mut session = session()?;
            if session.generation != generation || session.phase != Phase::Presented {
                log(&format!("commit image gen={generation} stale (current {} {:?})", session.generation, session.phase));
                return Ok(CommitOutcome::Stale);
            }
            session.phase = Phase::Hidden;
            session.draft = false;
            session.pending_keys.clear();
            (session.delivery == Delivery::InputMethod).then(|| session.return_to.take())
        };
        let window = app
            .get_webview_window(WINDOW_LABEL)
            .ok_or_else(|| "VisualTeX Anywhere panel is unavailable".to_string())?;
        if let Some(return_to) = input_method {
            // The field drops the formula's anchor and goes back to its
            // input source; the image is pasted where the anchor was.
            ime::send(serde_json::json!({ "t": "cancel", "returnTo": return_to }));
            panel::park(&window);
        } else {
            if !ax::is_trusted() {
                ax::request_trust();
                return Err("VisualTeX needs the Accessibility permission to paste the image".to_string());
            }
            panel::park_and_wait(&window)?;
        }
        let worker = app.clone();
        let bytes = png.len();
        let pasted = tauri::async_runtime::spawn_blocking(move || deliver::paste_png(&worker, png))
            .await
            .map_err(|error| format!("VisualTeX Anywhere insertion was interrupted: {error}"))?;
        log(&format!("commit image gen={generation} bytes={bytes} paste={pasted:?}"));
        pasted?;
        Ok(CommitOutcome::Inserted)
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (app, generation, png);
        Err("VisualTeX Anywhere is available on macOS only".to_string())
    }
}

#[tauri::command]
pub(crate) async fn anywhere_commit(app: AppHandle, generation: u64, text: String) -> Result<CommitOutcome, String> {
    if text.len() > 1024 * 1024 {
        return Err("The formula is too long to insert".to_string());
    }
    #[cfg(target_os = "macos")]
    {
        let input_method = {
            let mut session = session()?;
            if session.generation != generation || session.phase != Phase::Presented {
                log(&format!("commit gen={generation} stale (current {} {:?})", session.generation, session.phase));
                return Ok(CommitOutcome::Stale);
            }
            if session.delivery == Delivery::InputMethod {
                session.phase = Phase::Hidden;
                session.draft = false;
                session.pending_keys.clear();
                Some(session.return_to.take())
            } else {
                None
            }
        };
        if let Some(return_to) = input_method {
            // The field's marked text becomes the formula's text; the field
            // goes back to the input source it came from.
            let sent = ime::send(serde_json::json!({ "t": "commit", "text": text, "returnTo": return_to }));
            log(&format!("commit gen={generation} chars={} input_method={sent}", text.chars().count()));
            if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
                panel::park(&window);
            }
            return if sent {
                Ok(CommitOutcome::Inserted)
            } else {
                Err("VisualTeX's input method is not connected".to_string())
            };
        }
        if !ax::is_trusted() {
            ax::request_trust();
            let copied = deliver::copy(&app, text);
            log(&format!("commit gen={generation} untrusted copy={copied:?}"));
            copied?;
            return Ok(CommitOutcome::Copied);
        }
        {
            let mut session = session()?;
            if session.generation != generation || session.phase != Phase::Presented {
                return Ok(CommitOutcome::Stale);
            }
            session.phase = Phase::Hidden;
            session.draft = false;
        }
        let window = app
            .get_webview_window(WINDOW_LABEL)
            .ok_or_else(|| "VisualTeX Anywhere panel is unavailable".to_string())?;
        // The application's window must have the keyboard back before ⌘V.
        panel::park_and_wait(&window)?;
        let worker = app.clone();
        let length = text.chars().count();
        let pasted = tauri::async_runtime::spawn_blocking(move || deliver::paste(&worker, text))
            .await
            .map_err(|error| format!("VisualTeX Anywhere insertion was interrupted: {error}"))?;
        log(&format!("commit gen={generation} chars={length} paste={pasted:?}"));
        pasted?;
        Ok(CommitOutcome::Inserted)
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (app, generation);
        Err("VisualTeX Anywhere is available on macOS only".to_string())
    }
}

/// Escape (keep_draft false) or the strip's own close.
#[tauri::command]
pub(crate) fn anywhere_dismiss(app: AppHandle, generation: u64, keep_draft: bool) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        let current = session()?.generation == generation;
        if current {
            hide(&app, keep_draft, "page");
        }
        Ok(())
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (app, generation, keep_draft);
        Ok(())
    }
}

/// The formula's text for the field (the input method shows it as marked
/// text while the formula is open).
#[tauri::command]
pub(crate) fn anywhere_marked(generation: u64, text: String, width: f64, caret_x: Option<f64>) -> Result<(), String> {
    if text.len() > 256 * 1024 {
        return Err("The formula is too long to show".to_string());
    }
    #[cfg(target_os = "macos")]
    {
        let current = {
            let session = session()?;
            session.generation == generation && session.phase != Phase::Hidden && session.delivery == Delivery::InputMethod
        };
        if current && width.is_finite() {
            // The field holds the formula's width; `text` is what it gets if
            // it ends the composition itself.
            ime::send(serde_json::json!({
                "t": "inline",
                "width": width.clamp(0.0, 4096.0),
                "text": text,
                "caretX": caret_x.filter(|value| value.is_finite()).unwrap_or(0.0),
            }));
        }
        Ok(())
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (generation, text, width, caret_x);
        Ok(())
    }
}

/// The formula's caret entered or left text (\text{}): the input method
/// spells pinyin there.
#[tauri::command]
pub(crate) fn anywhere_mode(generation: u64, text: bool) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        let current = {
            let session = session()?;
            session.generation == generation && session.phase != Phase::Hidden && session.delivery == Delivery::InputMethod
        };
        if current {
            ime::send(serde_json::json!({ "t": "mode", "text": text }));
        }
        Ok(())
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (generation, text);
        Ok(())
    }
}

/// The page's own diagnostics (commit outcomes, errors), into anywhere.log.
#[tauri::command]
pub(crate) fn anywhere_log(line: String) {
    #[cfg(target_os = "macos")]
    {
        let line: String = line.chars().take(1600).map(|c| if c.is_control() { ' ' } else { c }).collect();
        log(&format!("page {line}"));
    }
    #[cfg(not(target_os = "macos"))]
    let _ = line;
}

/// Screenshot OCR while a formula is open: the user selects an area (macOS's
/// screenshot selection); no window moves, so the field keeps the keyboard
/// and the formula stays open. Returns the PNG (base64), None when cancelled.
#[tauri::command]
pub(crate) async fn anywhere_capture_screenshot() -> Result<Option<String>, String> {
    use base64::Engine as _;
    let capture = tauri::async_runtime::spawn_blocking(crate::quick_ocr::capture_selection_png)
        .await
        .map_err(|error| format!("Screenshot task failed: {error}"))??;
    Ok(capture.map(|bytes| base64::engine::general_purpose::STANDARD.encode(bytes)))
}

/// ⌘V in an open formula: the clipboard's image (base64, for OCR) or text.
#[derive(Serialize)]
#[serde(rename_all = "camelCase", tag = "kind")]
pub(crate) enum AnywhereClipboard {
    Image {
        #[serde(rename = "dataBase64")]
        data_base64: String,
        extension: String,
    },
    Text { text: String },
    Nothing,
}

#[tauri::command]
pub(crate) async fn anywhere_read_clipboard(app: AppHandle) -> Result<AnywhereClipboard, String> {
    #[cfg(target_os = "macos")]
    {
        use base64::Engine as _;
        let content = tauri::async_runtime::spawn_blocking(move || deliver::read_clipboard(&app))
            .await
            .map_err(|error| format!("Clipboard task failed: {error}"))??;
        Ok(match content {
            deliver::ClipboardContent::Image { bytes, extension } => {
                if bytes.len() > 20 * 1024 * 1024 {
                    return Err("The copied image is larger than 20 MB".to_string());
                }
                AnywhereClipboard::Image {
                    data_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
                    extension,
                }
            }
            deliver::ClipboardContent::Text(text) => AnywhereClipboard::Text { text },
            deliver::ClipboardContent::Nothing => AnywhereClipboard::Nothing,
        })
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = app;
        Ok(AnywhereClipboard::Nothing)
    }
}

/// Settings > Shortcuts: a global shortcut that opens and closes a formula
/// (`action`: "anywhere-toggle", the default, or "anywhere-open-display";
/// a null chord removes the display one).
#[tauri::command]
pub(crate) async fn anywhere_configure_hotkey(
    app: AppHandle,
    chord: serde_json::Value,
    action: Option<String>,
) -> Result<String, String> {
    #[cfg(target_os = "macos")]
    {
        let slot = match action.as_deref() {
            None => hotkey::Slot::Toggle,
            Some(action) => hotkey::Slot::from_action(action).ok_or_else(|| format!("Unknown shortcut {action}"))?,
        };
        let chord: Option<hotkey::AnywhereChord> =
            serde_json::from_value(chord).map_err(|error| format!("Unreadable shortcut: {error}"))?;
        hotkey::configure(&app, slot, chord)
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (app, chord, action);
        Err("VisualTeX Anywhere is available on macOS only".to_string())
    }
}

#[tauri::command]
pub(crate) fn anywhere_status() -> AnywhereStatus {
    #[cfg(target_os = "macos")]
    {
        AnywhereStatus {
            accessibility_trusted: ax::is_trusted(),
            shortcut: hotkey::shortcut_label(),
        }
    }
    #[cfg(not(target_os = "macos"))]
    {
        AnywhereStatus {
            accessibility_trusted: false,
            shortcut: String::new(),
        }
    }
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
        layout = PanelLayout::EMPTY;
        layout.width = 8.0;
        assert!(!layout.is_valid());
    }

    #[test]
    #[cfg(target_os = "macos")]
    fn the_anchor_sizes_only_near_the_first_report() {
        assert!(anchor_sizes(Some(18.0), Some(19.0)));
        // an empty line's line box first, the text's height on the anchor
        assert!(anchor_sizes(Some(16.0), Some(24.0)));
        // a made-up rectangle much shorter or taller than the field's text
        assert!(!anchor_sizes(Some(11.0), Some(20.0)));
        assert!(!anchor_sizes(Some(40.0), Some(20.0)));
        assert!(!anchor_sizes(None, Some(20.0)));
        // the first believable report sizes a formula opened at the system's size
        assert!(anchor_sizes(Some(14.0), None));
    }
}
