//! One engine per input context (a field of an application, as ibus-daemon creates them). VisualTeX
//! is a whole input method: Chinese through pinyin (RIME) and English (Shift switches), with formulas
//! as one more mode. While a formula is open every key goes to VisualTeX, which edits the formula and
//! draws it in the line itself, from where the field's caret was: the field holds the formula's space
//! in its preedit text (space.rs), so it wraps the formula as it wraps its own text; where the field
//! put each piece is read back by moving the preedit's cursor to it (IBus reports nothing but the
//! cursor's place). Return inserts the formula's text and, when the formula opened from another input
//! method, hands the field back to it. (The macOS input method, InputController.swift, does the same
//! through IMK.)

use crate::{ibus, keys, link::Link, rime, space};
use serde_json::{json, Value as Json};
use std::collections::HashMap;
use std::sync::{Arc, Mutex, Weak};
use std::time::{Duration, Instant};
use tokio::sync::Notify;
use zbus::object_server::SignalEmitter;
use zbus::zvariant::{OwnedObjectPath, OwnedValue, Value};
use zbus::Connection;

/// IBus capabilities (ibustypes.h).
const CAP_PREEDIT_TEXT: u32 = 1 << 0;

/// The formula's shortcut inside the input method (Ctrl+Alt+M; with Shift a display formula), so a
/// formula opens on Wayland too, where VisualTeX cannot take a global shortcut.
const FORMULA_KEY: u32 = 0x6d;

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Rect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

impl Rect {
    fn payload(&self) -> Json {
        json!([self.x, self.y, self.width, self.height])
    }
    fn bottom(&self) -> f64 {
        self.y + self.height
    }
    /// A place on some screen: positive size, not the origin a field answers before it knows.
    fn is_plausible(&self) -> bool {
        self.height > 0.0 && self.height < 400.0 && !(self.x == 0.0 && self.y == 0.0) && self.x.abs() < 30000.0 && self.y.abs() < 30000.0
    }
}

/// What one context shows and is doing.
struct Context {
    path: OwnedObjectPath,
    client: String,
    caps: u32,
    rime_session: rime::SessionId,
    /// The preedit text the field shows (pinyin's spelling, or the formula's space).
    marked: String,
    /// The cursor's place as the field last reported it (screen, y down).
    cursor: Option<Rect>,
    cursor_reports: u64,
    /// The text before the cursor on its line (surrounding text), when the field tells it.
    text_before: Option<bool>,
    /// A formula is open: its keys belong to VisualTeX.
    composing: bool,
    /// What the formula inserts, for a field that ends the composition itself.
    formula_text: String,
    formula_text_mode: bool,
    formula_pinyin: bool,
    ascii_before_formula: Option<bool>,
    /// The formula's pieces (pixels wide, from VisualTeX) and the preedit text holding them.
    pieces: Vec<f64>,
    space: space::Layout,
    filler: f64,
    joint: f64,
    calibrations: u32,
    /// The widths were measured in this field (else estimated: the space is held a margin short).
    widths_known: bool,
    /// The piece the field began the last line of the space with, as it was last read (space.rs).
    last_line_start: usize,
    /// The last full reading: the space's text then, each piece's place by its start, and when (places
    /// of pieces whose text before them has not changed since are not read again).
    probed: Option<(Vec<u16>, std::collections::HashMap<usize, Rect>, Instant)>,
    hold_generation: u64,
    last_hold: Instant,
    /// Where the formula starts (the first piece's place), once read.
    held_rect: Option<Rect>,
    /// The line where the formula opened, to tell a real place from a stale one.
    formula_start: Option<Rect>,
    last_places: Vec<[f64; 2]>,
    last_end: Option<Vec<f64>>,
    /// The formula's caret from its start (pixels): the candidate list goes there.
    caret_offset: (f64, f64),
    /// RIME's input mode last shown in the panel.
    shown_ascii: Option<bool>,
}

/// What the input method knows across its contexts.
pub struct Shared {
    pub conn: Connection,
    pub link: Link,
    contexts: Mutex<HashMap<String, Arc<Mutex<Context>>>>,
    notifiers: Mutex<HashMap<String, Arc<Notify>>>,
    /// The context whose field has the keyboard: VisualTeX's messages go to it.
    active: Mutex<Option<String>>,
    /// A filler's and a joint's width by application, learnt from where the field put the pieces.
    widths: Mutex<HashMap<String, (f64, f64)>>,
    /// The height of a line of text by application, measured where text stands before the caret.
    heights: Mutex<HashMap<String, f64>>,
    next_engine: Mutex<u64>,
    /// The field with the keyboard reports its cursor only at key events (Chromium, Electron).
    needs_key: std::sync::atomic::AtomicBool,
}

impl Shared {
    pub fn new(conn: Connection, link: Link) -> Arc<Shared> {
        Arc::new(Shared {
            conn,
            link,
            contexts: Mutex::new(HashMap::new()),
            notifiers: Mutex::new(HashMap::new()),
            active: Mutex::new(None),
            widths: Mutex::new(HashMap::new()),
            heights: Mutex::new(HashMap::new()),
            next_engine: Mutex::new(0),
            needs_key: std::sync::atomic::AtomicBool::new(false),
        })
    }

    fn context(&self, path: &str) -> Option<Arc<Mutex<Context>>> {
        self.contexts.lock().ok()?.get(path).cloned()
    }

    fn notifier(&self, path: &str) -> Arc<Notify> {
        let mut notifiers = self.notifiers.lock().unwrap();
        notifiers.entry(path.to_string()).or_insert_with(|| Arc::new(Notify::new())).clone()
    }

    fn active_context(&self) -> Option<Arc<Mutex<Context>>> {
        let path = self.active.lock().ok()?.clone()?;
        self.context(&path)
    }

    /// A new engine for ibus-daemon's CreateEngine: its object path.
    pub fn create(&self) -> OwnedObjectPath {
        let number = {
            let mut next = self.next_engine.lock().unwrap();
            *next += 1;
            *next
        };
        let path = OwnedObjectPath::try_from(format!("/org/freedesktop/IBus/Engine/VisualTeX/{number}")).unwrap();
        let context = Context {
            path: path.clone(),
            client: String::new(),
            caps: CAP_PREEDIT_TEXT,
            rime_session: 0,
            marked: String::new(),
            cursor: None,
            cursor_reports: 0,
            text_before: None,
            composing: false,
            formula_text: String::new(),
            formula_text_mode: false,
            formula_pinyin: false,
            ascii_before_formula: None,
            pieces: Vec::new(),
            space: space::Layout::default(),
            filler: 0.0,
            joint: 0.0,
            calibrations: 0,
            widths_known: false,
            last_line_start: 0,
            probed: None,
            hold_generation: 0,
            last_hold: Instant::now(),
            held_rect: None,
            formula_start: None,
            last_places: Vec::new(),
            last_end: None,
            caret_offset: (0.0, 0.0),
            shown_ascii: None,
        };
        self.contexts.lock().unwrap().insert(path.to_string(), Arc::new(Mutex::new(context)));
        path
    }

    pub fn destroy(&self, path: &str) {
        if let Some(context) = self.contexts.lock().unwrap().remove(path) {
            if let (Ok(context), Some(engine)) = (context.lock(), rime::engine()) {
                engine.destroy_session(context.rime_session);
            }
        }
        self.notifiers.lock().unwrap().remove(path);
        let mut active = self.active.lock().unwrap();
        if active.as_deref() == Some(path) {
            *active = None;
        }
    }
}

/// A signal to the field (through ibus-daemon).
enum Out {
    Commit(String),
    /// the preedit text, its cursor (characters) and the part underlined
    Preedit { text: String, cursor: u32, underline: Option<(u32, u32)> },
    /// the formula's space while its places are read: the anchor on a background or not (a change a
    /// field must draw, so it reports the cursor: see probe_once)
    Probe { text: String, cursor: u32, marked: bool },
    HidePreedit,
    Table { candidates: Vec<String>, labels: Vec<String>, cursor: u32 },
    HideTable,
    Mode(bool),
}

async fn emit(conn: &Connection, path: &OwnedObjectPath, outs: Vec<Out>) {
    if outs.is_empty() {
        return;
    }
    let Ok(emitter) = SignalEmitter::new(conn, path.clone()) else { return };
    for out in outs {
        let result = match out {
            Out::Commit(text) => Engine::commit_text(&emitter, ibus::text(&text, None)).await,
            Out::Preedit { text, cursor, underline } => {
                let length = text.chars().count() as u32;
                let value = ibus::text(&text, underline.map(|(start, end)| (start.min(length), end.min(length))));
                // mode 0: IBUS_ENGINE_PREEDIT_CLEAR (a field losing the focus drops it)
                Engine::update_preedit_text(&emitter, value, cursor.min(length), length > 0, 0).await
            }
            Out::Probe { text, cursor, marked } => {
                let length = text.chars().count() as u32;
                let value = ibus::text_with_background(&text, None, marked.then_some((0, 1, 0x00ff_ffff)));
                Engine::update_preedit_text(&emitter, value, cursor.min(length), length > 0, 0).await
            }
            Out::HidePreedit => {
                let _ = Engine::update_preedit_text(&emitter, ibus::text("", None), 0, false, 0).await;
                Engine::hide_preedit_text(&emitter).await
            }
            Out::Table { candidates, labels, cursor } => {
                Engine::update_lookup_table(&emitter, ibus::lookup_table(&candidates, &labels, cursor), true).await
            }
            Out::HideTable => Engine::hide_lookup_table(&emitter).await,
            Out::Mode(ascii) => Engine::update_property(&emitter, ibus::mode_property(ascii)).await,
        };
        if let Err(error) = result {
            crate::log::write(&format!("signal failed: {error}"));
        }
    }
}

// MARK: Pinyin and English (RIME)

fn ensure_rime(context: &mut Context) -> rime::SessionId {
    if context.rime_session == 0 {
        if let Some(engine) = rime::engine() {
            context.rime_session = engine.create_session();
            // a formula opened before RIME was ready starts in English as every formula does
            if context.rime_session != 0 && context.composing {
                set_formula_input_mode(context, true);
            }
        }
    }
    context.rime_session
}

fn set_formula_input_mode(context: &mut Context, on: bool) {
    let session = ensure_rime(context);
    let Some(engine) = rime::engine().filter(|_| session != 0) else { return };
    if on {
        if context.ascii_before_formula.is_none() {
            context.ascii_before_formula = Some(engine.is_ascii(session));
        }
        engine.set_ascii(session, true);
    } else if let Some(ascii) = context.ascii_before_formula.take() {
        engine.set_ascii(session, ascii);
    }
}

/// RIME's commit into the field, its spelling as preedit text, its candidates in the list.
fn show_pinyin(context: &mut Context, outs: &mut Vec<Out>) {
    let Some(engine) = rime::engine() else { return };
    let session = context.rime_session;
    if let Some(commit) = engine.take_commit(session).filter(|text| !text.is_empty()) {
        outs.push(Out::Commit(commit));
        context.marked.clear();
    }
    let state = engine.state(session);
    if state.is_composing() {
        context.marked = state.preedit.clone();
        let length = state.preedit.chars().count() as u32;
        outs.push(Out::Preedit { text: state.preedit.clone(), cursor: state.cursor as u32, underline: Some((0, length)) });
        if state.candidates.is_empty() {
            outs.push(Out::HideTable);
        } else {
            outs.push(Out::Table { candidates: state.candidates, labels: state.labels, cursor: state.highlighted as u32 });
        }
    } else {
        if !context.marked.is_empty() {
            context.marked.clear();
            outs.push(Out::HidePreedit);
        }
        outs.push(Out::HideTable);
    }
    show_mode(context, outs);
}

/// Pinyin in the formula's text: RIME's spelling is the editor's composition (shown in the
/// formula), its commit ends the composition with the chosen characters.
fn show_formula_pinyin(context: &mut Context, link: &Link, outs: &mut Vec<Out>) {
    let Some(engine) = rime::engine() else { return };
    let session = context.rime_session;
    if let Some(commit) = engine.take_commit(session).filter(|text| !text.is_empty()) {
        link.send(json!({ "t": "compose", "phase": "end", "text": commit }));
        context.formula_pinyin = false;
    }
    let state = engine.state(session);
    if state.is_composing() {
        link.send(json!({ "t": "compose", "phase": "update", "text": state.preedit }));
        context.formula_pinyin = true;
        outs.push(Out::Table { candidates: state.candidates, labels: state.labels, cursor: state.highlighted as u32 });
    } else {
        if context.formula_pinyin {
            link.send(json!({ "t": "compose", "phase": "end", "text": "" }));
        }
        context.formula_pinyin = false;
        outs.push(Out::HideTable);
    }
    show_mode(context, outs);
}

fn show_mode(context: &mut Context, outs: &mut Vec<Out>) {
    let Some(engine) = rime::engine() else { return };
    if context.rime_session == 0 {
        return;
    }
    let ascii = engine.is_ascii(context.rime_session);
    if context.shown_ascii != Some(ascii) {
        context.shown_ascii = Some(ascii);
        outs.push(Out::Mode(ascii));
    }
}

/// Drops pinyin being spelled in the formula (the formula closed).
fn drop_formula_pinyin(context: &mut Context, link: &Link, outs: &mut Vec<Out>) {
    let Some(engine) = rime::engine() else { return };
    if !(context.formula_pinyin || engine.state(context.rime_session).is_composing()) {
        return;
    }
    engine.clear(context.rime_session);
    if context.formula_pinyin {
        link.send(json!({ "t": "compose", "phase": "end", "text": "" }));
    }
    context.formula_pinyin = false;
    outs.push(Out::HideTable);
}

/// What is being spelled goes into the field as it is (focus moved, a formula opens).
fn finish_pinyin(context: &mut Context, outs: &mut Vec<Out>) {
    let Some(engine) = rime::engine() else { return };
    if context.rime_session == 0 || !engine.state(context.rime_session).is_composing() {
        return;
    }
    match engine.commit_composition(context.rime_session).filter(|text| !text.is_empty()) {
        Some(text) => outs.push(Out::Commit(text)),
        None => outs.push(Out::HidePreedit),
    }
    context.marked.clear();
    outs.push(Out::HideTable);
}

// MARK: The field's geometry

/// Where the formula goes and how the field's text looks there, for VisualTeX: the cursor's place
/// (screen, y down), the text height learnt in this application, whether text stands before the
/// cursor.
fn geometry(context: &Context, shared: &Shared) -> Json {
    let mut message = json!({ "client": context.client, "platform": "linux" });
    let rect = context.held_rect.or(context.cursor);
    if let Some(rect) = rect {
        message["caret"] = rect.payload();
        message["line"] = rect.payload();
    }
    if let Some(before) = context.text_before {
        message["textBefore"] = json!(before);
    }
    if let Some(height) = shared.heights.lock().ok().and_then(|heights| heights.get(&context.client).copied()) {
        message["textHeight"] = json!(height);
    }
    message["preedit"] = json!(context.caps & CAP_PREEDIT_TEXT != 0);
    message
}

fn note_text_height(context: &Context, shared: &Shared, rect: Rect) {
    if rect.height >= 6.0 && rect.height <= 120.0 && context.text_before == Some(true) {
        if let Ok(mut heights) = shared.heights.lock() {
            heights.insert(context.client.clone(), rect.height);
        }
    }
}

/// A filler's and a joint's width before the field has shown one: from its line's height (a line
/// is about 1.2 font sizes tall; a braille blank is about 0.6 em wide in the fonts fields fall back
/// to, a word space about 0.3 em).
fn estimated_widths(context: &Context, shared: &Shared) -> (f64, f64) {
    if let Some(widths) = shared.widths.lock().ok().and_then(|widths| widths.get(&context.client).copied()) {
        return widths;
    }
    let height = shared
        .heights
        .lock()
        .ok()
        .and_then(|heights| heights.get(&context.client).copied())
        .or(context.cursor.map(|rect| rect.height))
        .filter(|height| (6.0..=120.0).contains(height))
        .unwrap_or(17.0);
    let size = height / 1.19;
    (size * 0.6, size * 0.32)
}

// MARK: The engine object

pub struct Engine {
    pub shared: Arc<Shared>,
    pub path: OwnedObjectPath,
}

impl Engine {
    fn context(&self) -> Option<Arc<Mutex<Context>>> {
        self.shared.context(self.path.as_str())
    }

    async fn send_outs(&self, outs: Vec<Out>) {
        emit(&self.shared.conn, &self.path, outs).await;
    }

    fn activate_now(&self) {
        if let Ok(mut active) = self.shared.active.lock() {
            *active = Some(self.path.to_string());
        }
    }
}

#[zbus::interface(name = "org.freedesktop.IBus.Engine")]
impl Engine {
    async fn process_key_event(&self, keyval: u32, keycode: u32, state: u32) -> bool {
        // the key pressed to have Chromium report its cursor (see probe_once)
        if crate::x11::is_probe_key(keycode) {
            return true;
        }
        let Some(context) = self.context() else { return false };
        self.activate_now();
        let link = self.shared.link.clone();
        let mut outs = Vec::new();
        let handled = {
            let mut context = context.lock().unwrap();
            let release = keys::is_release(state);
            let shift_key = keyval == keys::SHIFT_L || keyval == keys::SHIFT_R;
            if context.composing {
                if shift_key {
                    // Shift switches Chinese and English in a formula too
                    let session = ensure_rime(&mut context);
                    let handled = rime::engine().map(|engine| engine.process(session, keyval, state & keys::RIME_MASK)).unwrap_or(false);
                    show_formula_pinyin(&mut context, &link, &mut outs);
                    handled
                } else if release {
                    true
                } else if state & (keys::SUPER | keys::MOD4 | keys::META | keys::HYPER) != 0 {
                    // the desktop's own shortcuts
                    false
                } else {
                    let session = ensure_rime(&mut context);
                    // in Chinese the keys spell pinyin and the chosen characters become the
                    // formula's text; in English (RIME takes nothing then) and for what RIME does not
                    // take they are the formula's. RIME sees them all, or a Shift around a "+" typed
                    // into the formula would switch it to Chinese (it saw Shift pressed alone).
                    let pinyin = !keys::has_command(state)
                        && rime::engine().map(|engine| engine.process(session, keyval, state & keys::RIME_MASK)).unwrap_or(false);
                    if pinyin {
                        show_formula_pinyin(&mut context, &link, &mut outs);
                    } else if !keys::is_modifier(keyval) {
                        link.send(keys::message(keyval, keycode, state));
                    }
                    true
                }
            } else if !release
                && keyval == FORMULA_KEY
                && state & keys::CONTROL != 0
                && state & keys::MOD1 != 0
                && link.is_connected()
            {
                // the formula's shortcut: VisualTeX opens a formula in this field
                finish_pinyin(&mut context, &mut outs);
                let mut message = geometry(&context, &self.shared);
                message["t"] = json!("start");
                message["display"] = json!(state & keys::SHIFT != 0);
                link.send(message);
                true
            } else {
                let session = ensure_rime(&mut context);
                if session == 0 {
                    false
                } else {
                    let handled = rime::engine().map(|engine| engine.process(session, keyval, state & keys::RIME_MASK)).unwrap_or(false);
                    show_pinyin(&mut context, &mut outs);
                    handled
                }
            }
        };
        self.send_outs(outs).await;
        handled
    }

    async fn set_cursor_location(&self, x: i32, y: i32, w: i32, h: i32) {
        let Some(context) = self.context() else { return };
        {
            let mut context = context.lock().unwrap();
            context.cursor = Some(Rect { x: x as f64, y: y as f64, width: w as f64, height: h as f64 });
            context.cursor_reports += 1;
            if std::env::var_os("VISUALTEX_IBUS_DEBUG").is_some() {
                crate::log::write(&format!("cursor {x},{y},{w},{h}"));
            }
        }
        self.shared.notifier(self.path.as_str()).notify_waiters();
    }

    async fn set_cursor_location_relative(&self, x: i32, y: i32, w: i32, h: i32) {
        if std::env::var_os("VISUALTEX_IBUS_DEBUG").is_some() {
            crate::log::write(&format!("cursor relative {x},{y},{w},{h}"));
        }
    }

    async fn process_hand_writing_event(&self, _coordinates: Vec<f64>) {}

    async fn cancel_hand_writing(&self, _n_strokes: u32) {}

    async fn set_capabilities(&self, caps: u32) {
        if let Some(context) = self.context() {
            context.lock().unwrap().caps = caps;
        }
    }

    async fn property_activate(&self, name: String, _state: u32) {
        if name != ibus::MODE_PROPERTY {
            return;
        }
        let Some(context) = self.context() else { return };
        let mut outs = Vec::new();
        {
            let mut context = context.lock().unwrap();
            let session = ensure_rime(&mut context);
            if let Some(engine) = rime::engine().filter(|_| session != 0) {
                let ascii = engine.is_ascii(session);
                engine.set_ascii(session, !ascii);
                show_mode(&mut context, &mut outs);
            }
        }
        self.send_outs(outs).await;
    }

    async fn property_show(&self, _name: String) {}

    async fn property_hide(&self, _name: String) {}

    async fn candidate_clicked(&self, index: u32, _button: u32, _state: u32) {
        let Some(context) = self.context() else { return };
        let link = self.shared.link.clone();
        let mut outs = Vec::new();
        {
            let mut context = context.lock().unwrap();
            let session = context.rime_session;
            if rime::engine().map(|engine| engine.select_on_page(session, index as usize)).unwrap_or(false) {
                if context.composing {
                    show_formula_pinyin(&mut context, &link, &mut outs);
                } else {
                    show_pinyin(&mut context, &mut outs);
                }
            }
        }
        self.send_outs(outs).await;
    }

    async fn focus_in(&self) {
        self.focused(String::new()).await;
    }

    async fn focus_in_id(&self, _object_path: String, client: String) {
        self.focused(client).await;
    }

    async fn focus_out(&self) {
        self.unfocused().await;
    }

    async fn focus_out_id(&self, _object_path: String) {
        self.unfocused().await;
    }

    async fn reset(&self) {
        // The field ended the composition itself (a click elsewhere in it): it gets the formula's
        // text, not its space; what was being spelled is dropped, as fields expect after a reset.
        let Some(context) = self.context() else { return };
        let link = self.shared.link.clone();
        let mut outs = Vec::new();
        {
            let mut context = context.lock().unwrap();
            if context.composing {
                crate::log::write(&format!("reset while composing text={}", context.formula_text.chars().count()));
                if !context.formula_text.is_empty() {
                    outs.push(Out::Commit(context.formula_text.clone()));
                } else {
                    outs.push(Out::HidePreedit);
                }
                context.marked.clear();
                context.formula_text.clear();
                drop_formula_pinyin(&mut context, &link, &mut outs);
                set_formula_input_mode(&mut context, false);
                context.composing = false;
                context.hold_generation += 1;
                link.send(json!({ "t": "forced" }));
            } else if let Some(engine) = rime::engine() {
                engine.clear(context.rime_session);
                context.marked.clear();
                outs.push(Out::HideTable);
            }
        }
        self.send_outs(outs).await;
    }

    async fn enable(&self) {
        let Some(context) = self.context() else { return };
        let mut outs = Vec::new();
        {
            let mut context = context.lock().unwrap();
            ensure_rime(&mut context);
            context.shown_ascii = None;
            show_mode(&mut context, &mut outs);
        }
        if let Ok(emitter) = SignalEmitter::new(&self.shared.conn, self.path.clone()) {
            let ascii = matches!(outs.first(), Some(Out::Mode(true)));
            let _ = Engine::register_properties(&emitter, ibus::property_list(vec![ibus::mode_property(ascii)])).await;
        }
        self.send_outs(outs).await;
    }

    async fn disable(&self) {
        self.unfocused().await;
    }

    async fn page_up(&self) {
        self.key_to_rime(0xff55).await;
    }

    async fn page_down(&self) {
        self.key_to_rime(0xff56).await;
    }

    async fn cursor_up(&self) {
        self.key_to_rime(0xff52).await;
    }

    async fn cursor_down(&self) {
        self.key_to_rime(0xff54).await;
    }

    async fn set_surrounding_text(&self, text: OwnedValue, cursor_pos: u32, _anchor_pos: u32) {
        let Some(context) = self.context() else { return };
        // an IBusText: ("IBusText", attachments, text, attributes)
        let content = match &*text {
            Value::Structure(structure) => structure.fields().get(2).and_then(|field| match field {
                Value::Str(text) => Some(text.as_str().to_string()),
                _ => None,
            }),
            _ => None,
        };
        let Some(content) = content else { return };
        let before = content.chars().take(cursor_pos as usize).last();
        context.lock().unwrap().text_before = Some(matches!(before, Some(c) if c != '\n'));
    }

    async fn panel_extension_received(&self, _event: OwnedValue) {}

    async fn panel_extension_register_keys(&self, _data: OwnedValue) {}

    #[zbus(property)]
    fn content_type(&self) -> (u32, u32) {
        (0, 0)
    }

    #[zbus(property)]
    fn set_content_type(&mut self, _value: (u32, u32)) {}

    /// FocusInId and FocusOutId are understood (they name the client).
    #[zbus(property)]
    fn focus_id(&self) -> (bool,) {
        (true,)
    }

    #[zbus(property)]
    fn active_surrounding_text(&self) -> (bool,) {
        (true,)
    }

    #[zbus(signal)]
    async fn commit_text(emitter: &SignalEmitter<'_>, text: Value<'_>) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn update_preedit_text(emitter: &SignalEmitter<'_>, text: Value<'_>, cursor_pos: u32, visible: bool, mode: u32) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn show_preedit_text(emitter: &SignalEmitter<'_>) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn hide_preedit_text(emitter: &SignalEmitter<'_>) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn update_auxiliary_text(emitter: &SignalEmitter<'_>, text: Value<'_>, visible: bool) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn update_lookup_table(emitter: &SignalEmitter<'_>, table: Value<'_>, visible: bool) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn show_lookup_table(emitter: &SignalEmitter<'_>) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn hide_lookup_table(emitter: &SignalEmitter<'_>) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn register_properties(emitter: &SignalEmitter<'_>, props: Value<'_>) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn update_property(emitter: &SignalEmitter<'_>, prop: Value<'_>) -> zbus::Result<()>;

    #[zbus(signal)]
    async fn forward_key_event(emitter: &SignalEmitter<'_>, keyval: u32, keycode: u32, state: u32) -> zbus::Result<()>;
}

impl Engine {
    async fn key_to_rime(&self, keyval: u32) {
        let Some(context) = self.context() else { return };
        let link = self.shared.link.clone();
        let mut outs = Vec::new();
        {
            let mut context = context.lock().unwrap();
            let session = context.rime_session;
            if rime::engine().map(|engine| engine.process(session, keyval, 0)).unwrap_or(false) {
                if context.composing {
                    show_formula_pinyin(&mut context, &link, &mut outs);
                } else {
                    show_pinyin(&mut context, &mut outs);
                }
            }
        }
        self.send_outs(outs).await;
    }

    async fn focused(&self, client: String) {
        let Some(context) = self.context() else { return };
        self.activate_now();
        // another field: whether it needs the probe key is learnt again
        self.shared.needs_key.store(false, std::sync::atomic::Ordering::SeqCst);
        let reports = {
            let mut context = context.lock().unwrap();
            if !client.is_empty() {
                context.client = client;
            }
            context.cursor_reports
        };
        // the field tells its cursor's place right after: VisualTeX hears of the field with it
        let notifier = self.shared.notifier(self.path.as_str());
        let deadline = Instant::now() + Duration::from_millis(150);
        loop {
            let notified = notifier.notified();
            if context.lock().unwrap().cursor_reports != reports {
                break;
            }
            let left = deadline.saturating_duration_since(Instant::now());
            if left.is_zero() || tokio::time::timeout(left, notified).await.is_err() {
                break;
            }
        }
        let message = {
            let context = context.lock().unwrap();
            let mut message = geometry(&context, &self.shared);
            message["t"] = json!("activate");
            crate::log::write(&format!(
                "activate client={} caret={:?} linked={}",
                context.client,
                context.cursor.map(|rect| (rect.x, rect.y, rect.height)),
                self.shared.link.is_connected()
            ));
            message
        };
        self.shared.link.send(message);
    }

    async fn unfocused(&self) {
        let Some(context) = self.context() else { return };
        let link = self.shared.link.clone();
        let mut outs = Vec::new();
        let client = {
            let mut context = context.lock().unwrap();
            crate::log::write(&format!("deactivate client={} composing={}", context.client, context.composing));
            if context.composing {
                if !context.marked.is_empty() {
                    outs.push(Out::HidePreedit);
                }
                context.marked.clear();
                drop_formula_pinyin(&mut context, &link, &mut outs);
                set_formula_input_mode(&mut context, false);
                context.composing = false;
                context.hold_generation += 1;
            } else {
                finish_pinyin(&mut context, &mut outs);
            }
            context.formula_text.clear();
            outs.push(Out::HideTable);
            context.client.clone()
        };
        if let Ok(mut active) = self.shared.active.lock() {
            if active.as_deref() == Some(self.path.as_str()) {
                *active = None;
            }
        }
        link.send(json!({ "t": "deactivate", "client": client }));
        self.send_outs(outs).await;
    }
}

// MARK: VisualTeX's messages

/// A message from VisualTeX, for the context whose field has the keyboard.
pub async fn receive(shared: Arc<Shared>, message: Json) {
    let kind = message.get("t").and_then(Json::as_str).unwrap_or("").to_string();
    let Some(context_arc) = shared.active_context() else {
        if kind != "disconnected" {
            crate::log::write(&format!("message {kind} without an active field"));
        }
        return;
    };
    let link = shared.link.clone();
    let mut outs = Vec::new();
    let mut hold_now = false;
    let mut return_to: Option<String> = None;
    let path = {
        let mut context = context_arc.lock().unwrap();
        match kind.as_str() {
            "disconnected" => {
                if context.composing {
                    outs.push(Out::HidePreedit);
                    context.marked.clear();
                    set_formula_input_mode(&mut context, false);
                    context.composing = false;
                    context.hold_generation += 1;
                }
            }
            "session" => {
                // VisualTeX opened (its shortcut) or closed a formula for this field; what was
                // being spelled is typed first
                let on = message.get("on").and_then(Json::as_bool).unwrap_or(false);
                if on {
                    finish_pinyin(&mut context, &mut outs);
                } else {
                    drop_formula_pinyin(&mut context, &link, &mut outs);
                }
                // a formula starts in English (its letters are variables); Shift switches to
                // Chinese for text; the input mode comes back after
                set_formula_input_mode(&mut context, on);
                context.composing = on;
                context.formula_text_mode = false;
                context.formula_start = if on { context.cursor } else { None };
                context.held_rect = None;
                context.hold_generation += 1;
                context.pieces.clear();
                context.last_places.clear();
                context.last_end = None;
                context.calibrations = 0;
                context.last_line_start = 0;
                context.probed = None;
                context.caret_offset = (0.0, 0.0);
                context.space = space::Layout::default();
                if on {
                    let (filler, joint) = estimated_widths(&context, &shared);
                    context.filler = filler;
                    context.joint = joint;
                    context.widths_known = shared.widths.lock().ok().map(|widths| widths.contains_key(&context.client)).unwrap_or(false);
                    hold_now = true;
                } else if !context.marked.is_empty() {
                    context.marked.clear();
                    outs.push(Out::HidePreedit);
                }
                crate::log::write(&format!("session {}", if on { "on" } else { "off" }));
            }
            "mode" => {
                context.formula_text_mode = message.get("text").and_then(Json::as_bool).unwrap_or(false);
            }
            "inline" => {
                context.formula_text = message.get("text").and_then(Json::as_str).unwrap_or("").to_string();
                context.caret_offset = (
                    message.get("caretX").and_then(Json::as_f64).unwrap_or(0.0),
                    message.get("caretY").and_then(Json::as_f64).unwrap_or(0.0),
                );
                if let (Some(pieces), true) = (message.get("pieces").and_then(Json::as_array), context.composing) {
                    let widths: Vec<f64> = pieces.iter().filter_map(Json::as_f64).map(|width| width.clamp(0.0, 20000.0)).collect();
                    if widths != context.pieces {
                        context.pieces = widths;
                        hold_now = true;
                    }
                }
            }
            "commit" => {
                let text = message.get("text").and_then(Json::as_str).unwrap_or("").to_string();
                drop_formula_pinyin(&mut context, &link, &mut outs);
                set_formula_input_mode(&mut context, false);
                context.formula_text_mode = false;
                outs.push(Out::HidePreedit);
                outs.push(Out::Commit(text.clone()));
                crate::log::write(&format!("commit length={} return={:?}", text.chars().count(), message.get("returnTo")));
                context.marked.clear();
                context.formula_text.clear();
                context.composing = false;
                context.hold_generation += 1;
                context.space = space::Layout::default();
                return_to = message.get("returnTo").and_then(Json::as_str).map(str::to_string);
            }
            "cancel" => {
                drop_formula_pinyin(&mut context, &link, &mut outs);
                set_formula_input_mode(&mut context, false);
                context.formula_text_mode = false;
                if !context.marked.is_empty() {
                    outs.push(Out::HidePreedit);
                }
                context.marked.clear();
                context.formula_text.clear();
                context.composing = false;
                context.hold_generation += 1;
                context.space = space::Layout::default();
                crate::log::write(&format!("cancel return={:?}", message.get("returnTo")));
                return_to = message.get("returnTo").and_then(Json::as_str).map(str::to_string);
            }
            "caret" => {
                let mut reply = geometry(&context, &shared);
                reply["t"] = json!("geometry");
                link.send(reply);
            }
            _ => {}
        }
        context.path.clone()
    };
    emit(&shared.conn, &path, outs).await;
    if hold_now {
        hold(shared.clone(), context_arc.clone()).await;
    }
    if let Some(engine) = return_to.filter(|engine| !engine.is_empty() && engine != "visualtex") {
        // the field goes back to the input method the formula's shortcut took it from
        let result = shared
            .conn
            .call_method(Some("org.freedesktop.IBus"), "/org/freedesktop/IBus", Some("org.freedesktop.IBus"), "SetGlobalEngine", &(engine.as_str(),))
            .await;
        if let Err(error) = result {
            crate::log::write(&format!("could not hand the field back to {engine}: {error}"));
        }
    }
}

// MARK: Where the formula goes in the line

/// The field holds the formula's pieces (space.rs); where each went is read back (probe). (Boxed:
/// a calibration in the probe holds again.)
fn hold(shared: Arc<Shared>, context_arc: Arc<Mutex<Context>>) -> std::pin::Pin<Box<dyn std::future::Future<Output = ()> + Send>> {
    Box::pin(async move {
        let (path, text, generation) = {
            let mut context = context_arc.lock().unwrap();
            if !context.composing {
                return;
            }
            // held short of the formula by as much as estimated widths may be off (5% a filler, a
            // quarter of a joint)
            let unit = context.filler.max(0.5);
            let fillers = context.pieces.iter().sum::<f64>() / unit;
            let joints = context.pieces.len().saturating_sub(1) as f64;
            let margin = if context.widths_known { 0.15 } else { 0.15 + fillers * 0.05 + joints * 0.25 * context.joint / unit };
            let line_start = if context.last_line_start < context.pieces.len() { context.last_line_start } else { 0 };
            let layout = space::layout(&context.pieces, context.filler, context.joint, margin, line_start);
            let text = String::from_utf16_lossy(&layout.text);
            if layout == context.space && context.marked == text {
                return;
            }
            crate::log::write(&format!(
                "hold pieces={} fillers={} filler={:.2} joint={:.2}",
                layout.starts.len(),
                layout.counts.iter().sum::<usize>(),
                context.filler,
                context.joint
            ));
            context.space = layout;
            context.marked = text.clone();
            context.last_hold = Instant::now();
            context.hold_generation += 1;
            (context.path.clone(), text, context.hold_generation)
        };
        emit(&shared.conn, &path, vec![Out::Preedit { text, cursor: 0, underline: None }]).await;
        let weak = Arc::downgrade(&context_arc);
        tokio::spawn(probe(shared, weak, generation));
    })
}

/// Waits until the field reports its cursor again (or `wait` passes); its place then.
async fn next_cursor(shared: &Shared, context: &Arc<Mutex<Context>>, path: &str, reports: u64, wait: Duration) -> Option<Rect> {
    let notifier = shared.notifier(path);
    let deadline = Instant::now() + wait;
    loop {
        let notified = notifier.notified();
        {
            let context = context.lock().unwrap();
            if context.cursor_reports != reports {
                return context.cursor;
            }
        }
        let left = deadline.saturating_duration_since(Instant::now());
        if left.is_zero() || tokio::time::timeout(left, notified).await.is_err() {
            return context.lock().unwrap().cursor;
        }
    }
}

/// What a reading of the pieces' places found: the formula changed since (stop), the places are not
/// settled yet (read again soon), or they are (follow them, slowly).
enum Probe {
    Stale,
    Settle,
    Done,
}

/// Follows where the field puts the formula's pieces (probe_once) while this space is held.
async fn probe(shared: Arc<Shared>, weak: Weak<Mutex<Context>>, generation: u64) {
    // the field lays the new text out first (and typing holds a new one every few milliseconds)
    let mut delay = Duration::from_millis(40);
    let mut settling = 0;
    loop {
        tokio::time::sleep(delay).await;
        let Some(context_arc) = weak.upgrade() else { return };
        match probe_once(&shared, &context_arc, generation).await {
            Probe::Stale => return,
            Probe::Settle if settling < 12 => {
                settling += 1;
                delay = Duration::from_millis(130);
            }
            // the field may still move the formula (scrolling, a late layout): followed slowly
            _ => delay = Duration::from_secs(1),
        }
    }
}

/// Reads where the field put each piece: the preedit's cursor at each piece's start and at the
/// end, the field reporting the cursor's place each time; then back before the formula. The
/// places go to VisualTeX (pieces), the widths are calibrated from them.
async fn probe_once(shared: &Arc<Shared>, context_arc: &Arc<Mutex<Context>>, generation: u64) -> Probe {
    let (path, text, starts, counts, length) = {
        let context = context_arc.lock().unwrap();
        if context.hold_generation != generation || !context.composing {
            return Probe::Stale;
        }
        let length = context.space.text.len();
        (context.path.clone(), context.marked.clone(), context.space.starts.clone(), context.space.counts.clone(), length)
    };
    let started = Instant::now();
    let mut variant = false;
    let mut indices = starts.clone();
    indices.push(length);
    // places of pieces whose text before them is as at the last full reading a moment ago (typing at
    // the formula's end changes its last piece only): kept, not read again (each reading moves the
    // field's cursor, in Chromium with a key: a formula of many pieces was read for seconds a key)
    let reuse: Vec<Option<Rect>> = {
        let context = context_arc.lock().unwrap();
        match &context.probed {
            Some((text, at, when)) if when.elapsed() < Duration::from_millis(1500) => {
                let common = text.iter().zip(context.space.text.iter()).take_while(|(a, b)| a == b).count();
                starts.iter().map(|start| if *start <= common { at.get(start).copied() } else { None }).collect()
            }
            _ => vec![None; starts.len()],
        }
    };
    let mut places: Vec<Option<Rect>> = Vec::new();
    for (position, &index) in indices.iter().enumerate() {
        if let Some(rect) = reuse.get(position).copied().flatten() {
            places.push(Some(rect));
            continue;
        }
        let reports = {
            let context = context_arc.lock().unwrap();
            if context.hold_generation != generation || !context.composing {
                return Probe::Stale;
            }
            context.cursor_reports
        };
        // the anchor goes on a background and off (under the formula's cover): Chromium reports
        // the cursor with a new frame, and a cursor moved in a composition that looks the same draws
        // none (it underlines a composition itself when no underline is given, so not an underline)
        variant = !variant;
        emit(&shared.conn, &path, vec![Out::Probe { text: text.clone(), cursor: index as u32, marked: variant }]).await;
        // most fields report the cursor's new place at once; Chromium and Electron only at a key
        // event: then the probe key is pressed (and swallowed when it comes back)
        let needs_key = shared.needs_key.load(std::sync::atomic::Ordering::SeqCst);
        let first_wait = if needs_key { 0 } else { 60 };
        let mut rect = None;
        let mut reported = false;
        if first_wait > 0 {
            rect = next_cursor(&shared, &context_arc, path.as_str(), reports, Duration::from_millis(first_wait)).await;
            reported = context_arc.lock().unwrap().cursor_reports != reports;
        }
        if !reported && crate::x11::press_probe_key() {
            // Chromium tells the place its renderer last sent, which is behind a moment after the
            // preedit changed: the key is pressed again until the place moves from the last piece's
            let previous = places.last().copied().flatten();
            let deadline = Instant::now() + Duration::from_millis(probe_wait().max(400));
            loop {
                let before = context_arc.lock().unwrap().cursor_reports;
                let answer = next_cursor(&shared, &context_arc, path.as_str(), before, Duration::from_millis(80)).await;
                if context_arc.lock().unwrap().cursor_reports != before {
                    reported = true;
                    rect = answer;
                    if answer != previous || previous.is_none() {
                        break;
                    }
                }
                if Instant::now() >= deadline || !crate::x11::press_probe_key() {
                    break;
                }
            }
            if reported && !needs_key {
                shared.needs_key.store(true, std::sync::atomic::Ordering::SeqCst);
                crate::log::write("the field reports its cursor at key events: probe key");
            }
        } else if !reported {
            rect = next_cursor(&shared, &context_arc, path.as_str(), reports, Duration::from_millis(probe_wait())).await;
        }
        places.push(rect.filter(Rect::is_plausible));
    }
    // back before the formula, where the field's own caret is covered
    let reports = context_arc.lock().unwrap().cursor_reports;
    emit(&shared.conn, &path, vec![Out::Preedit { text: text.clone(), cursor: 0, underline: None }]).await;
    let anchor = next_cursor(&shared, &context_arc, path.as_str(), reports, Duration::from_millis(150)).await;
    let elapsed = started.elapsed().as_millis();
    let mut context = context_arc.lock().unwrap();
    if context.hold_generation != generation || !context.composing {
        return Probe::Stale;
    }
    let end = places.pop().flatten();
    // each piece after the one before on its line, or on a lower line: a field that answers one
    // place for every part of the preedit places nothing
    let mut rects: Vec<Rect> = Vec::new();
    let mut stop: Option<String> = None;
    for (index, place) in places.iter().enumerate() {
        let Some(rect) = *place else {
            stop = Some(format!("piece {index} has no place"));
            break;
        };
        if let Some(before) = rects.last() {
            let same_line = (rect.y - before.y).abs() < before.height * 0.5;
            if if same_line { rect.x <= before.x } else { rect.y < before.y } {
                stop = Some(format!("piece {index} at {:.0},{:.0} does not follow {:.0},{:.0}", rect.x, rect.y, before.x, before.y));
                break;
            }
        }
        rects.push(rect);
    }
    if let Some(stop) = &stop {
        crate::log::write(&format!("probe: {stop} ms={elapsed}"));
    }
    let Some(first) = rects.first().copied() else { return Probe::Settle };
    if rects.len() == starts.len() {
        // where the last line begins (the first of the pieces on the last piece's line), and the
        // reading, for the next one
        let last = rects[rects.len() - 1];
        let mut line_start = rects.len() - 1;
        while line_start > 0 && (rects[line_start - 1].y - last.y).abs() < last.height * 0.5 {
            line_start -= 1;
        }
        context.last_line_start = line_start;
        let at: std::collections::HashMap<usize, Rect> = starts.iter().copied().zip(rects.iter().copied()).collect();
        context.probed = Some((context.space.text.clone(), at, Instant::now()));
    }
    if let Some(start) = context.formula_start {
        if (first.y - start.y).abs() > (first.height * 4.0).max(80.0) {
            crate::log::write("probe: the first piece is far from where the formula opened");
            return Probe::Settle;
        }
    }
    note_text_height(&context, &shared, first);
    if context.held_rect != Some(first) {
        context.held_rect = Some(first);
        let mut message = geometry(&context, &shared);
        message["t"] = json!("geometry");
        message["caret"] = json!([first.x, first.y, 0.0, first.height]);
        shared.link.send(message);
        crate::log::write(&format!("anchor {:.0},{:.0},{:.0} ms={elapsed}", first.x, first.y, first.height));
    }
    // the places from the first piece's start on its line (pixels, y down, by the lines' bottoms)
    let round = |value: f64| (value * 100.0).round() / 100.0;
    let places: Vec<[f64; 2]> = rects.iter().map(|rect| [round(rect.x - first.x), round(rect.bottom() - first.bottom())]).collect();
    let all_placed = rects.len() == starts.len();
    let end_place: Option<Vec<f64>> = match (end, all_placed) {
        (Some(end), true) => Some(vec![round(end.x - first.x), round(end.bottom() - first.bottom())]),
        (None, true) => {
            let last = rects.last().copied().unwrap_or(first);
            let x = last.x + *counts.last().unwrap_or(&1) as f64 * context.filler;
            Some(vec![round(x - first.x), round(last.bottom() - first.bottom()), 1.0])
        }
        _ => None,
    };
    if places != context.last_places || end_place != context.last_end {
        context.last_places = places.clone();
        context.last_end = end_place.clone();
        let mut message = json!({ "t": "pieces", "places": places, "count": starts.len() });
        if let Some(end) = &end_place {
            message["end"] = json!(end);
        }
        if let Some(anchor) = anchor.filter(|anchor| (anchor.y - first.y).abs() < first.height * 0.5) {
            message["start"] = json!([round(anchor.x - first.x), 0.0]);
        }
        shared.link.send(message);
    }
    // the filler's and the joint's widths as the field draws them, from the pieces' places on their
    // lines: start(next) − start = count × filler + joint, end − start(last) = count(last) × filler
    if !all_placed || context.last_hold.elapsed() < Duration::from_millis(120) {
        return Probe::Settle;
    }
    let mut equations: Vec<(f64, f64, f64)> = Vec::new();
    for index in 0..rects.len().saturating_sub(1) {
        let (here, next) = (rects[index], rects[index + 1]);
        if (next.y - here.y).abs() < here.height * 0.5 && next.x > here.x {
            equations.push((counts[index] as f64, 1.0, next.x - here.x));
        }
    }
    if let (Some(end), Some(last)) = (end, rects.last()) {
        if (end.y - last.y).abs() < last.height * 0.5 && end.x > last.x {
            equations.push((*counts.last().unwrap_or(&1) as f64, 0.0, end.x - last.x));
        }
    }
    let Some((filler, joint)) = solve(&equations) else { return Probe::Done };
    if !(0.5..40.0).contains(&filler) || !(0.0..40.0).contains(&joint) {
        return Probe::Done;
    }
    let fillers: f64 = counts.iter().sum::<usize>() as f64;
    context.widths_known = true;
    let filler_changed = (filler - context.filler).abs() * fillers.max(1.0) > 1.5;
    let joint_changed = (joint - context.joint).abs() > 1.5;
    let client = context.client.clone();
    if !filler_changed && !joint_changed {
        if let Ok(mut widths) = shared.widths.lock() {
            widths.insert(client, (filler, joint));
        }
        return Probe::Done;
    }
    // past three, the widths stay those the space is held with
    if context.calibrations >= 3 {
        return Probe::Done;
    }
    context.calibrations += 1;
    crate::log::write(&format!(
        "widths filler {:.2} -> {:.2}, joint {:.2} -> {:.2} from {} readings",
        context.filler,
        filler,
        context.joint,
        joint,
        equations.len()
    ));
    context.filler = filler;
    context.joint = joint;
    if let Ok(mut widths) = shared.widths.lock() {
        widths.insert(client, (filler, joint));
    }
    drop(context);
    tokio::spawn(hold(shared.clone(), context_arc.clone()));
    Probe::Stale
}

/// Least squares for a × filler + b × joint = d over the readings (the joint alone when only one
/// count was read is not told apart from the filler: then the joint keeps its estimate).
fn solve(equations: &[(f64, f64, f64)]) -> Option<(f64, f64)> {
    let (mut aa, mut ab, mut bb, mut ad, mut bd) = (0.0, 0.0, 0.0, 0.0, 0.0);
    for &(a, b, d) in equations {
        aa += a * a;
        ab += a * b;
        bb += b * b;
        ad += a * d;
        bd += b * d;
    }
    let determinant = aa * bb - ab * ab;
    if determinant.abs() > 1e-6 {
        let filler = (ad * bb - bd * ab) / determinant;
        let joint = (aa * bd - ab * ad) / determinant;
        return Some((filler, joint.max(0.0)));
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn widths_from_pieces_on_one_line() {
        // three pieces of 4, 2 and 5 fillers 7 wide, joints 3 wide
        let equations = vec![(4.0, 1.0, 31.0), (2.0, 1.0, 17.0), (5.0, 0.0, 35.0)];
        let (filler, joint) = solve(&equations).unwrap();
        assert!((filler - 7.0).abs() < 1e-9 && (joint - 3.0).abs() < 1e-9);
    }

    #[test]
    fn one_reading_tells_nothing_apart() {
        assert!(solve(&[(4.0, 1.0, 31.0)]).is_none());
    }
}

/// How long a reading waits for the field to report the cursor (milliseconds).
fn probe_wait() -> u64 {
    std::env::var("VISUALTEX_PROBE_WAIT").ok().and_then(|value| value.parse().ok()).unwrap_or(150)
}
