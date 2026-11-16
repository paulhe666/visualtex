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
mod screen_caret;
#[cfg(target_os = "macos")]
mod screen_colors;

use serde::{Deserialize, Serialize};
use tauri::AppHandle;
#[cfg(target_os = "macos")]
use std::sync::atomic::{AtomicU64, Ordering};
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
/// Where the field put the formula's pieces (it wraps them as its text).
#[cfg(target_os = "macos")]
const PIECES_EVENT: &str = "visualtex-anywhere-pieces";
/// Where the panel stands on the screen, each time its frame is set (panel.rs): the page cannot read it from its window.
#[cfg(target_os = "macos")]
const FRAME_EVENT: &str = "visualtex-anywhere-frame";
/// The most pieces a formula is cut into for the field (a longer formula's
/// last piece holds the rest).
const MAX_PIECES: usize = 512;
/// The shortcut asked the input method where the caret is; a field that takes text answers within
/// this (else what has the keyboard is no text field of the input method's: the pasting strip).
#[cfg(target_os = "macos")]
const INPUT_METHOD_WAIT: Duration = Duration::from_millis(500);
/// The input method reconnects to VisualTeX within a second of being started
/// or of VisualTeX starting (input-method/Sources/AppLink.swift).
#[cfg(target_os = "macos")]
const IME_RELINK_WAIT: Duration = Duration::from_millis(1500);
/// Keys typed before the card is up wait for it, up to this many.
#[cfg(target_os = "macos")]
const MAX_PENDING_KEYS: usize = 256;
/// How long a paste waits for the input method to say it gave the keys back.
#[cfg(target_os = "macos")]
const KEYS_RELEASE_WAIT: Duration = Duration::from_millis(300);
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
    /// The focused field's frame, as Accessibility reports it.
    pub(crate) field: Option<ScreenRect>,
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
    /// The field's application has been running since before the input method was one of the
    /// input sources (macOS gives an application its input methods when it starts): the strip
    /// names it, to be reopened; then its formulas are typed in its line.
    reopen_app: Option<String>,
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
    /// The display shortcut opened this session.
    display: bool,
    /// The generation whose shortcut asked the input method for the field's caret
    /// and waits for its answer.
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
    /// The app's input source before the strip took an English layout (given
    /// back when it hides).
    strip_source: Option<String>,
    /// The frontmost application (bundle identifier) when the shortcut was
    /// pressed: the field the input method must answer from.
    field_app: Option<String>,
    /// The strip's activation: the field's application to reopen (see
    /// Activation::reopen_app).
    reopen_app: Option<String>,
    /// The generation whose pasting strip has its keys taken by the input
    /// method, as the in-line formula has (before the strip is up they wait
    /// for it): typed so, the formula needs no English input source and the
    /// user's input source is never switched.
    grab_keys: u64,
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
    strip_source: None,
    field_app: None,
    reopen_app: None,
    grab_keys: 0,
});

/// The application whose field has VisualTeX's input method now (its bundle
/// identifier, the "client" of the input method's activate; None after its
/// deactivate).
#[cfg(target_os = "macos")]
static IME_FIELD: Mutex<Option<String>> = Mutex::new(None);

#[cfg(target_os = "macos")]
fn note_ime_field(kind: &str, client: Option<&str>) {
    let Ok(mut field) = IME_FIELD.lock() else { return };
    match kind {
        "activate" => *field = client.map(str::to_string),
        "deactivate" if field.as_deref() == client => *field = None,
        "disconnected" => *field = None,
        _ => {}
    }
}

#[cfg(target_os = "macos")]
fn ime_field() -> Option<String> {
    IME_FIELD.lock().ok().and_then(|field| field.clone())
}

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
                if ime::installed() && !screen_colors::can_read_screen() {
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
            if ime::installed() {
                begin_input_method(&app, display);
            } else {
                activate(&app, display, None);
            }
        });
    if let Err(error) = spawned {
        eprintln!("Unable to start VisualTeX Anywhere: {error}");
    }
}

/// The pasting strip: what has the keyboard is no text field the input method reaches
/// (`reopen_app`: because its application must be opened again first, see Activation::reopen_app).
/// The strip stands on the caret Accessibility reports; where it reports none (WeChat draws its own
/// text, QQ's Chromium reports the box but no place in it), on the caret found blinking on the
/// screen (screen_caret), looked for alongside; else on the field, or where that application's
/// caret was found last, or low in its window. Until the strip is up, the input method takes the
/// keys typed (they wait for the strip, nothing reaches the field).
#[cfg(target_os = "macos")]
fn activate(app: &AppHandle, display: bool, reopen_app: Option<String>) {
    let started = std::time::Instant::now();
    let trusted = ax::is_trusted();
    if !trusted {
        ax::request_trust();
    }
    let generation = {
        let Ok(mut session) = session() else { return };
        session.generation += 1;
        session.phase = Phase::Activating;
        session.target = None;
        session.delivery = Delivery::Paste;
        session.reopen_app = reopen_app;
        session.display = display;
        session.awaiting_input_method = 0;
        session.pending_keys.clear();
        session.field_text = FieldText::default();
        session.reference_height = None;
        session.inline_anchor = None;
        session.grab_keys = 0;
        session.generation
    };
    // (VisualTeX's own shortcuts excepted)
    let grabbed = ime::is_linked()
        && ime::send(serde_json::json!({ "t": "grab", "on": true, "skip": hotkey::registered_chords() }))
        && session()
            .map(|mut session| {
                let current = session.generation == generation;
                if current {
                    session.grab_keys = generation;
                }
                current
            })
            .unwrap_or(false);
    if grabbed {
        keep_keys(generation);
    }
    let search = screen_caret::Search::start(ax::frontmost_pid());
    let mut target = ax::capture_target(trusted);
    let mut screen = String::new();
    if target.source == AnchorSource::Caret {
        search.cancel();
    } else {
        let window = search.window();
        let field = if target.source == AnchorSource::Field { target.field } else { None };
        match search.finish(field) {
            Ok(found) => {
                target.anchor = found.caret;
                target.source = AnchorSource::Caret;
                if let Some(window) = window {
                    remember_caret(&target.app_name, window.frame(), found.caret);
                }
                screen = format!(" screen caret {}", found.report);
            }
            Err(why) => {
                screen = format!(" screen caret: {why}");
                if target.source == AnchorSource::Pointer {
                    if let Some(window) = window {
                        target.anchor = recalled_caret(&target.app_name, window.frame()).unwrap_or_else(|| low_in(window.frame()));
                        target.source = AnchorSource::Field;
                    }
                }
            }
        }
    }
    log(&format!(
        "activate trusted={trusted} source={:?} anchor={:.0},{:.0},{:.0} app={:?} capture_ms={}{screen}",
        target.source,
        target.anchor.x,
        target.anchor.y,
        target.anchor.height,
        target.app_name,
        started.elapsed().as_millis()
    ));
    let current = |session: &Session| session.generation == generation && session.phase == Phase::Activating;
    // (the shortcut pressed again while the caret was looked for: the strip is not wanted)
    if !session().map(|session| current(&*session)).unwrap_or(false) {
        return;
    }
    let window = match panel::ensure(app) {
        Ok(window) => window,
        Err(error) => {
            eprintln!("VisualTeX Anywhere panel is unavailable: {error}");
            hide_with(app, false, "no panel", false);
            return;
        }
    };
    // The card takes the field's light or dark look (QQ's dark theme on a light system: a dark card).
    let colors = panel::window_number(&window).and_then(|number| screen_colors::field_colors(target.anchor, number));
    if let Some(colors) = colors.as_ref() {
        log(&format!("card colour gen={generation} background={} (screen)", colors.background));
    }
    // Without the input method to take its keys, the strip's formula is typed in English: the
    // input source an ASCII-capable layout while it is up (the one it had, back when it ends).
    let strip_source = if grabbed { None } else { ime::select_ascii(app) };
    let (activation, ready) = {
        let Ok(mut session) = session() else { return };
        if !current(&*session) {
            drop(session);
            if let Some(source) = strip_source {
                ime::restore(app, &source);
            }
            return;
        }
        session.target = Some(target);
        if strip_source.is_some() {
            session.strip_source = strip_source;
        }
        if let Some(colors) = colors {
            session.field_text.text_color = Some(colors.text);
            session.field_text.background_color = Some(colors.background);
        }
        (pending_activation(&session, trusted), session.ready)
    };
    // A cold (or reloading) panel takes the activation when its page reports
    // ready.
    if let (Some(activation), true) = (activation, ready) {
        send_activation(app, &window, activation);
    }
}

/// Where each application's caret was found on the screen last, within its
/// window: (window size, offset from the window's top left, caret height).
#[cfg(target_os = "macos")]
static FOUND_CARETS: Mutex<Vec<(String, (f64, f64), ScreenRect)>> = Mutex::new(Vec::new());

#[cfg(target_os = "macos")]
fn remember_caret(app_name: &str, window: ScreenRect, caret: ScreenRect) {
    let Ok(mut carets) = FOUND_CARETS.lock() else { return };
    carets.retain(|(name, _, _)| name != app_name);
    carets.push((
        app_name.to_string(),
        (window.width, window.height),
        ScreenRect { x: caret.x - window.x, y: caret.y - window.y, width: 0.0, height: caret.height },
    ));
}

/// The application's caret where it was found last, if its window has kept its size.
#[cfg(target_os = "macos")]
fn recalled_caret(app_name: &str, window: ScreenRect) -> Option<ScreenRect> {
    let carets = FOUND_CARETS.lock().ok()?;
    let (_, size, offset) = carets.iter().find(|(name, _, _)| name == app_name)?;
    ((size.0 - window.width).abs() < 2.0 && (size.1 - window.height).abs() < 2.0).then(|| ScreenRect {
        x: window.x + offset.x,
        y: window.y + offset.y,
        width: 0.0,
        height: offset.height,
    })
}

/// Low in the window, where a chat's message box is.
#[cfg(target_os = "macos")]
fn low_in(window: ScreenRect) -> ScreenRect {
    ScreenRect {
        x: window.x + window.width * 0.5,
        y: window.y + (window.height - 110.0).max(window.height * 0.6),
        width: 0.0,
        height: 18.0,
    }
}

/// The input method's "grab" requests that want an answer, and the last one answered.
#[cfg(target_os = "macos")]
static GRAB_REQUESTS: AtomicU64 = AtomicU64::new(0);
#[cfg(target_os = "macos")]
static GRAB_ANSWERED: AtomicU64 = AtomicU64::new(0);

/// The input method gives the strip's keys back. `wait`: until it says it has (at most
/// KEYS_RELEASE_WAIT), for a paste that follows: the ⌘V it holds would take the paste's ⌘V.
#[cfg(target_os = "macos")]
fn release_keys(wait: bool) {
    let id = GRAB_REQUESTS.fetch_add(1, Ordering::SeqCst) + 1;
    if !ime::send(serde_json::json!({ "t": "grab", "on": false, "id": id })) || !wait {
        return;
    }
    let deadline = std::time::Instant::now() + KEYS_RELEASE_WAIT;
    while GRAB_ANSWERED.load(Ordering::SeqCst) < id && std::time::Instant::now() < deadline {
        std::thread::sleep(Duration::from_millis(4));
    }
    if GRAB_ANSWERED.load(Ordering::SeqCst) < id {
        log("input method did not confirm it gave the keys back");
    }
}

/// While a pasting strip has its keys taken, the input method is told so every second: it gives
/// them back by itself three seconds after the last word (a VisualTeX that hangs keeps no keys).
/// A strip that never comes up keeps them at most a few seconds.
#[cfg(target_os = "macos")]
fn keep_keys(generation: u64) {
    let spawned = std::thread::Builder::new()
        .name("visualtex-keys-kept".to_string())
        .spawn(move || {
            let started = std::time::Instant::now();
            loop {
                std::thread::sleep(Duration::from_secs(1));
                let holding = session()
                    .map(|session| {
                        session.generation == generation
                            && session.grab_keys == generation
                            && (session.phase == Phase::Presented
                                || (session.phase == Phase::Activating && started.elapsed() < Duration::from_secs(4)))
                    })
                    .unwrap_or(false);
                if !holding {
                    break;
                }
                ime::send(serde_json::json!({ "t": "grab", "on": true, "keep": true, "skip": hotkey::registered_chords() }));
            }
        });
    if let Err(error) = spawned {
        log(&format!("keys kept: no thread: {error}"));
    }
}

/// A pasting strip's commit: its keys are given back (and confirmed) before the paste, and an input
/// source it switched to (without the input method) is given back while it still has the keyboard.
#[cfg(target_os = "macos")]
async fn end_pasting_strip(app: &AppHandle, grabbed: bool, strip_source: Option<String>) {
    if grabbed {
        let _ = tauri::async_runtime::spawn_blocking(|| release_keys(true)).await;
    }
    if let Some(source) = strip_source {
        let worker = app.clone();
        let _ = tauri::async_runtime::spawn_blocking(move || ime::restore(&worker, &source)).await;
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
        reopen_app: session.reopen_app.clone().filter(|_| session.delivery == Delivery::Paste),
    })
}

/// The shortcut with the input method selected: the input method has the field that has the
/// keyboard already (it is selected beside the user's input source); it is asked where the caret
/// is, and its answer (input_method_message) opens the formula there. Nothing is switched.
#[cfg(target_os = "macos")]
fn begin_input_method(app: &AppHandle, display: bool) {
    let started = std::time::Instant::now();
    let field_app = ax::frontmost_bundle();
    let known_field = ime_field();
    let generation = {
        let Ok(mut session) = session() else { return };
        session.generation += 1;
        session.phase = Phase::Activating;
        session.target = None;
        session.delivery = Delivery::InputMethod;
        session.field_app = field_app.clone();
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
    // The input method's process may be gone (it is started with VisualTeX; a crash, an update):
    // started again, it links within a second.
    if !ime::is_linked() {
        ime::launch();
        let deadline = std::time::Instant::now() + IME_RELINK_WAIT;
        while !ime::is_linked() && std::time::Instant::now() < deadline {
            std::thread::sleep(Duration::from_millis(40));
        }
    }
    // (linked a moment ago: it says whether it is a selected input source right after)
    let deadline = std::time::Instant::now() + Duration::from_millis(400);
    while ime::is_linked() && !ime::is_selected() && std::time::Instant::now() < deadline {
        std::thread::sleep(Duration::from_millis(20));
    }
    let asked = ime::is_selected() && ime::send(serde_json::json!({ "t": "caret", "app": field_app }));
    log(&format!(
        "input method asked gen={generation} field={} ime_field={} linked={} selected={} asked={asked} ms={}",
        field_app.as_deref().unwrap_or("?"),
        known_field.as_deref().unwrap_or("-"),
        ime::is_linked(),
        ime::is_selected(),
        started.elapsed().as_millis()
    ));
    if !asked {
        fall_back_to_paste(app, generation, display);
        return;
    }
    let app = app.clone();
    let _ = std::thread::Builder::new()
        .name("visualtex-ime-wait".to_string())
        .spawn(move || {
            std::thread::sleep(INPUT_METHOD_WAIT);
            let awaiting = session()
                .map(|session| session.awaiting_input_method == generation && session.generation == generation)
                .unwrap_or(false);
            if awaiting {
                log(&format!("input method did not answer gen={generation}"));
                fall_back_to_paste(&app, generation, display);
            }
        });
}

/// No text field of the input method's has the keyboard: the pasting strip instead.
#[cfg(target_os = "macos")]
fn fall_back_to_paste(app: &AppHandle, generation: u64, display: bool) {
    {
        let Ok(mut session) = session() else { return };
        if session.generation != generation || session.awaiting_input_method != generation {
            return;
        }
        session.awaiting_input_method = 0;
        session.phase = Phase::Hidden;
        session.delivery = Delivery::Paste;
    }
    // An application running since before the input method was an input source has not got it
    // (a sandboxed one never gets it: reopening it changes nothing).
    let reopen = match (ax::frontmost_started(), ime::selected_since(app)) {
        (Some((name, started)), Some(since)) if started < since && !ax::frontmost_sandboxed() => Some(name),
        _ => None,
    };
    if let Some(name) = reopen.as_ref() {
        log(&format!("gen={generation}: {name} has been running since before the input method: to be reopened"));
    }
    activate(app, display, reopen);
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
    note_ime_field(kind, message.get("client").and_then(|value| value.as_str()));
    match kind {
        "hello" | "selected" => {
            let selected = message.get("selected").and_then(|value| value.as_bool()).unwrap_or(false);
            ime::note_selected(app, selected);
            log(&format!(
                "input method {} selected={selected}",
                message.get("version").and_then(|value| value.as_str()).unwrap_or("")
            ));
        }
        "geometry" => input_method_caret_reported(app, message),
        "nofield" => {
            // the shortcut's question: no field of the input method's has the keyboard
            let (generation, display) = session().map(|session| (session.generation, session.display)).unwrap_or((0, false));
            log(&format!("input method has no field gen={generation}"));
            fall_back_to_paste(app, generation, display);
        }
        "key" => input_method_key(app, message),
        "grab" => {
            // its answer to release_keys
            if let Some(id) = message.get("id").and_then(|value| value.as_u64()) {
                GRAB_ANSWERED.fetch_max(id, Ordering::SeqCst);
            }
        }
        "pieces" => input_method_pieces(app, message),
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
                hide_with(app, false, "field kept the text", true);
            }
        }
        _ => {}
    }
}

/// The input method's answer with the field's geometry: open the formula the
/// shortcut asked for, or follow the formula's place.
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
        // VisualTeX's own windows (its strip) are not the field the shortcut was
        // pressed in.
        let own_window = client == ime::APP_IDENTIFIER && session.field_app.as_deref() != Some(ime::APP_IDENTIFIER);
        if session.delivery != Delivery::InputMethod || own_window {
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
                field: None,
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
        Next::Open(mut activation, ready) => {
            // The field's colours now, before the field shows anything for the formula: an editor
            // that draws what is being composed on a background of its own (VS Code) gave the cover
            // that colour when the screen was read afterwards, and the cover showed as a band.
            let colors = geometry.as_ref().and_then(|geometry| {
                let number = app.get_webview_window(WINDOW_LABEL).and_then(|window| panel::window_number(&window))?;
                screen_colors::field_colors(geometry.caret, number)
            });
            if let Some(colors) = colors {
                if let Ok(mut session) = session() {
                    if Some(session.generation) == activation.as_ref().map(|activation| activation.generation) {
                        session.color_read = session.generation;
                        session.color_from_screen = session.generation;
                        session.field_text.text_color = Some(colors.text.clone());
                        session.field_text.background_color = Some(colors.background.clone());
                    }
                }
                log(&format!("field colour text={} background={} (before the formula)", colors.text, colors.background));
                if let Some(activation) = activation.as_mut() {
                    activation.text_color = Some(colors.text);
                    activation.background_color = Some(colors.background);
                }
            }
            // (the formula's keys are the input method's to take, but VisualTeX's own shortcuts)
            ime::send(serde_json::json!({ "t": "session", "on": true, "skip": hotkey::registered_chords() }));
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

/// Where the field put the formula's pieces (points from the first piece's
/// start, y down): for the page, which draws each piece there.
#[cfg(target_os = "macos")]
fn input_method_pieces(app: &AppHandle, message: &serde_json::Value) {
    let Some(generation) = session()
        .ok()
        .filter(|session| session.phase != Phase::Hidden && session.delivery == Delivery::InputMethod)
        .map(|session| session.generation)
    else {
        return;
    };
    let places: Vec<[f64; 2]> = message
        .get("places")
        .and_then(|value| value.as_array())
        .map(|places| {
            places
                .iter()
                .take(MAX_PIECES)
                .map_while(|place| {
                    let pair = place.as_array()?;
                    let x = pair.first()?.as_f64()?;
                    let y = pair.get(1)?.as_f64()?;
                    (x.is_finite() && y.is_finite() && x.abs() < 20000.0 && y.abs() < 20000.0).then_some([x, y])
                })
                .collect()
        })
        .unwrap_or_default();
    // where the held space starts (its anchor) and ends (the formula's last line), when the field tells
    let point = |key: &str| {
        message
            .get(key)
            .and_then(|value| value.as_array())
            .and_then(|pair| Some([pair.first()?.as_f64()?, pair.get(1)?.as_f64()?]))
            .filter(|[x, y]| x.is_finite() && y.is_finite() && x.abs() < 20000.0 && y.abs() < 20000.0)
    };
    let (start, end) = (point("start"), point("end"));
    // the field's own caret after the space: how far its line's bottom is below the last line's, and its height
    let caret_line = point("caretLine");
    // an end the field did not tell (the input method's estimate)
    let estimated = message
        .get("end")
        .and_then(|value| value.as_array())
        .map(|end| end.len() > 2)
        .unwrap_or(false);
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        let _ = window.emit(
            PIECES_EVENT,
            serde_json::json!({ "generation": generation, "places": places, "start": start, "end": end, "endEstimated": estimated, "caretLine": caret_line }),
        );
    }
}

/// A key the input method took: to the editor once the card is up (until
/// then it waits, so fast typing after the shortcut is not lost).
#[cfg(target_os = "macos")]
fn input_method_key(app: &AppHandle, message: &serde_json::Value) {
    let mut key = message.clone();
    let deliver_now = {
        let Ok(mut session) = session() else { return };
        let grabbed = session.delivery == Delivery::Paste && session.grab_keys == session.generation;
        if (session.delivery != Delivery::InputMethod && !grabbed) || session.phase == Phase::Hidden {
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
    let (reload, dropped) = {
        let Ok(mut session) = session() else { return };
        if session.generation != generation || session.phase != Phase::Activating {
            return;
        }
        if session.reloaded_for == generation {
            session.phase = Phase::Hidden;
            let grabbed = session.grab_keys == generation;
            session.grab_keys = 0;
            (false, Some((grabbed, session.strip_source.take())))
        } else {
            session.reloaded_for = generation;
            session.ready = false;
            (true, None)
        }
    };
    if let Some((grabbed, strip_source)) = dropped {
        if grabbed {
            release_keys(false);
        }
        if let Some(source) = strip_source {
            ime::restore(app, &source);
        }
    }
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
    if delivery == Delivery::Paste {
        // the keys typed before the strip was up (the input method keeps taking them while it is)
        let grabbed = session()
            .map(|session| session.generation == generation && session.grab_keys == generation)
            .unwrap_or(false);
        if grabbed {
            flush_pending_keys(app, generation);
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
/// field's marked text (not when the input method itself ended the session).
#[cfg(target_os = "macos")]
fn hide_with(app: &AppHandle, keep_draft: bool, reason: &str, tell_input_method: bool) {
    let (generation, delivery, strip_source, grabbed) = {
        let Ok(mut session) = session() else { return };
        if session.phase == Phase::Hidden {
            return;
        }
        session.phase = Phase::Hidden;
        session.draft = keep_draft;
        session.awaiting_input_method = 0;
        session.pending_keys.clear();
        let grabbed = session.grab_keys == session.generation;
        session.grab_keys = 0;
        (session.generation, session.delivery, session.strip_source.take(), grabbed)
    };
    if grabbed {
        release_keys(false);
    }
    if let Some(source) = strip_source {
        ime::restore(app, &source);
    }
    if delivery == Delivery::InputMethod && tell_input_method {
        ime::send(serde_json::json!({ "t": "cancel" }));
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
/// application (with the input method, the formula's space is dropped first).
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
        let (input_method, grabbed, strip_source) = {
            let mut session = session()?;
            if session.generation != generation || session.phase != Phase::Presented {
                log(&format!("commit image gen={generation} stale (current {} {:?})", session.generation, session.phase));
                return Ok(CommitOutcome::Stale);
            }
            session.phase = Phase::Hidden;
            session.draft = false;
            session.pending_keys.clear();
            let grabbed = session.grab_keys == session.generation;
            session.grab_keys = 0;
            (session.delivery == Delivery::InputMethod, grabbed, session.strip_source.take())
        };
        end_pasting_strip(&app, grabbed, strip_source).await;
        let window = app
            .get_webview_window(WINDOW_LABEL)
            .ok_or_else(|| "VisualTeX Anywhere panel is unavailable".to_string())?;
        if input_method {
            // The field drops the formula's space; the image is pasted where it was.
            ime::send(serde_json::json!({ "t": "cancel" }));
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
                true
            } else {
                false
            }
        };
        if input_method {
            // The field's marked text becomes the formula's text.
            let sent = ime::send(serde_json::json!({ "t": "commit", "text": text }));
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
        let (grabbed, strip_source) = {
            let mut session = session()?;
            if session.generation != generation || session.phase != Phase::Presented {
                return Ok(CommitOutcome::Stale);
            }
            session.phase = Phase::Hidden;
            session.draft = false;
            session.pending_keys.clear();
            let grabbed = session.grab_keys == session.generation;
            session.grab_keys = 0;
            (grabbed, session.strip_source.take())
        };
        end_pasting_strip(&app, grabbed, strip_source).await;
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
    #[cfg(target_os = "macos")]
    {
        let current = {
            let session = session()?;
            session.generation == generation && session.phase != Phase::Hidden && session.delivery == Delivery::InputMethod
        };
        let _ = (caret_x, caret_y);
        if current && width.is_finite() {
            // The field holds the formula's width; `text` is what it gets if
            // it ends the composition itself.
            // The pieces TeX may break the formula into (points wide): the
            // field holds each and wraps them as its own text.
            let pieces: Vec<f64> = pieces
                .unwrap_or_default()
                .into_iter()
                .take(MAX_PIECES)
                .map(|piece| if piece.is_finite() { piece.clamp(0.0, 8192.0) } else { 0.0 })
                .collect();
            ime::send(serde_json::json!({
                "t": "inline",
                "width": width.clamp(0.0, 4096.0),
                "text": text,
                "pieces": pieces,
            }));
        }
        Ok(())
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (generation, text, width, caret_x, caret_y, pieces);
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
