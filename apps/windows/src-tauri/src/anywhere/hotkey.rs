//! The system-wide shortcuts that open and close a formula (Settings > Keyboard shortcuts >
//! Formula input method): Ctrl+Alt+M starts it as an inline row, Ctrl+Alt+Shift+M as a display
//! row (by default). Windows hotkeys on a thread of their own, kept in app data
//! (anywhere-hotkey.json, anywhere-display-hotkey.json) so they are registered before the page
//! loads.

use serde::{Deserialize, Serialize};
use std::sync::mpsc::{self, Sender};
use std::sync::{Mutex, OnceLock};
use std::time::Duration;
use tauri::{AppHandle, Manager};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    RegisterHotKey, UnregisterHotKey, HOT_KEY_MODIFIERS, MOD_ALT, MOD_CONTROL, MOD_NOREPEAT, MOD_SHIFT, MOD_WIN,
};
use windows::Win32::UI::WindowsAndMessaging::{GetMessageW, PostThreadMessageW, MSG, WM_APP, WM_HOTKEY};

const WM_RECONFIGURE: u32 = WM_APP + 0x41;

/// The system-wide shortcuts: each opens and closes a formula, starting it as an inline or as a
/// display row.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Slot {
    Toggle,
    Display,
}

impl Slot {
    const ALL: [Slot; 2] = [Slot::Toggle, Slot::Display];

    /// Settings' action id (src/shortcuts/anywhereHotkeys.ts).
    pub(crate) fn from_action(action: &str) -> Option<Self> {
        match action {
            "anywhere-toggle" => Some(Slot::Toggle),
            "anywhere-open-display" => Some(Slot::Display),
            _ => None,
        }
    }

    fn index(self) -> usize {
        self as usize
    }

    /// "VA" and "VB", in RegisterHotKey's application range.
    fn hotkey_id(self) -> i32 {
        0x5641 + self.index() as i32
    }

    fn config_file(self) -> &'static str {
        match self {
            Slot::Toggle => "anywhere-hotkey.json",
            Slot::Display => "anywhere-display-hotkey.json",
        }
    }

    /// Ctrl+Alt+M; with Shift it opens a display formula ($ and $$).
    fn default_chord(self) -> Chord {
        Chord { code: "KeyM".to_string(), ctrl_key: true, alt_key: true, shift_key: self == Slot::Display, meta_key: false }
    }

    fn display(self) -> bool {
        self == Slot::Display
    }
}

/// A chord as the page describes it (DOM `code` and modifiers).
#[derive(Clone, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Chord {
    pub(crate) code: String,
    #[serde(default)]
    pub(crate) ctrl_key: bool,
    #[serde(default)]
    pub(crate) alt_key: bool,
    #[serde(default)]
    pub(crate) shift_key: bool,
    #[serde(default)]
    pub(crate) meta_key: bool,
}

impl Chord {
    fn modifiers(&self) -> HOT_KEY_MODIFIERS {
        let mut modifiers = MOD_NOREPEAT;
        if self.ctrl_key {
            modifiers |= MOD_CONTROL;
        }
        if self.alt_key {
            modifiers |= MOD_ALT;
        }
        if self.shift_key {
            modifiers |= MOD_SHIFT;
        }
        if self.meta_key {
            modifiers |= MOD_WIN;
        }
        modifiers
    }

    pub(crate) fn label(&self) -> String {
        let mut parts = Vec::new();
        if self.ctrl_key {
            parts.push("Ctrl".to_string());
        }
        if self.alt_key {
            parts.push("Alt".to_string());
        }
        if self.shift_key {
            parts.push("Shift".to_string());
        }
        if self.meta_key {
            parts.push("Win".to_string());
        }
        let key = self
            .code
            .strip_prefix("Key")
            .or_else(|| self.code.strip_prefix("Digit"))
            .unwrap_or(&self.code)
            .to_string();
        parts.push(key);
        parts.join("+")
    }
}

/// The Windows virtual key of a DOM `code`.
fn virtual_key(code: &str) -> Option<u32> {
    if let Some(letter) = code.strip_prefix("Key") {
        let ch = letter.chars().next()?;
        return (letter.len() == 1 && ch.is_ascii_uppercase()).then_some(ch as u32);
    }
    if let Some(digit) = code.strip_prefix("Digit") {
        let ch = digit.chars().next()?;
        return (digit.len() == 1 && ch.is_ascii_digit()).then_some(ch as u32);
    }
    if let Some(digit) = code.strip_prefix("Numpad") {
        if let Ok(value) = digit.parse::<u32>() {
            return (value <= 9).then_some(0x60 + value);
        }
    }
    if let Some(number) = code.strip_prefix('F') {
        if let Ok(value) = number.parse::<u32>() {
            return (1..=24).contains(&value).then_some(0x6f + value);
        }
    }
    Some(match code {
        "Space" => 0x20,
        "Enter" => 0x0d,
        "Tab" => 0x09,
        "PageUp" => 0x21,
        "PageDown" => 0x22,
        "End" => 0x23,
        "Home" => 0x24,
        "ArrowLeft" => 0x25,
        "ArrowUp" => 0x26,
        "ArrowRight" => 0x27,
        "ArrowDown" => 0x28,
        "Insert" => 0x2d,
        "Delete" => 0x2e,
        "Semicolon" => 0xba,
        "Equal" => 0xbb,
        "Comma" => 0xbc,
        "Minus" => 0xbd,
        "Period" => 0xbe,
        "Slash" => 0xbf,
        "Backquote" => 0xc0,
        "BracketLeft" => 0xdb,
        "Backslash" => 0xdc,
        "BracketRight" => 0xdd,
        "Quote" => 0xde,
        "NumpadMultiply" => 0x6a,
        "NumpadAdd" => 0x6b,
        "NumpadSubtract" => 0x6d,
        "NumpadDecimal" => 0x6e,
        "NumpadDivide" => 0x6f,
        _ => return None,
    })
}

type Request = (Slot, Option<Chord>, Sender<Result<String, String>>);

struct Hotkeys {
    thread: u32,
    /// Each slot's registered chord.
    current: [Option<Chord>; 2],
    request: Option<Request>,
}

static STATE: OnceLock<Mutex<Hotkeys>> = OnceLock::new();

fn state() -> &'static Mutex<Hotkeys> {
    STATE.get_or_init(|| Mutex::new(Hotkeys { thread: 0, current: [None, None], request: None }))
}

fn config_path(app: &AppHandle, slot: Slot) -> Option<std::path::PathBuf> {
    app.path().app_data_dir().ok().map(|directory| directory.join(slot.config_file()))
}

/// The slot's saved chord: the default when none was saved (or it is unreadable), none when the
/// user cleared it ("null"; not the inline one).
fn saved_chord(app: &AppHandle, slot: Slot) -> Option<Chord> {
    let Some(bytes) = config_path(app, slot).and_then(|path| std::fs::read(path).ok()) else {
        return Some(slot.default_chord());
    };
    match serde_json::from_slice::<Option<Chord>>(&bytes) {
        Ok(None) if slot != Slot::Toggle => None,
        Ok(Some(chord)) if virtual_key(&chord.code).is_some() => Some(chord),
        _ => Some(slot.default_chord()),
    }
}

fn register(slot: Slot, chord: &Chord) -> Result<(), String> {
    let key = virtual_key(&chord.code).ok_or_else(|| format!("{} cannot be a system-wide shortcut", chord.code))?;
    unsafe { RegisterHotKey(None, slot.hotkey_id(), chord.modifiers(), key) }.map_err(|error| {
        format!("Windows refused the shortcut {} (another program may use it): {error}", chord.label())
    })
}

/// Unregisters the slot's chord and registers `chord` (None: none); the old one comes back on
/// failure.
fn set_chord(slot: Slot, chord: Option<Chord>) -> Result<String, String> {
    let previous = state().lock().ok().and_then(|state| state.current[slot.index()].clone());
    unsafe {
        let _ = UnregisterHotKey(None, slot.hotkey_id());
    }
    let Some(chord) = chord else {
        if let Ok(mut state) = state().lock() {
            state.current[slot.index()] = None;
        }
        return Ok(String::new());
    };
    match register(slot, &chord) {
        Ok(()) => {
            let label = chord.label();
            if let Ok(mut state) = state().lock() {
                state.current[slot.index()] = Some(chord);
            }
            Ok(label)
        }
        Err(error) => {
            if let Some(previous) = previous {
                let _ = register(slot, &previous);
            }
            Err(error)
        }
    }
}

/// Registers the saved shortcuts on their own thread, which then waits for them.
pub(super) fn register_saved(app: &AppHandle) -> Result<(), String> {
    let chords = Slot::ALL.map(|slot| saved_chord(app, slot));
    let app = app.clone();
    let (ready_sender, ready) = mpsc::channel();
    std::thread::Builder::new()
        .name("visualtex-anywhere-hotkey".to_string())
        .spawn(move || {
            if let Ok(mut state) = state().lock() {
                state.thread = unsafe { GetCurrentThreadId() };
            }
            let mut toggle = Ok(String::new());
            for (slot, chord) in Slot::ALL.into_iter().zip(chords) {
                let Some(chord) = chord else { continue };
                let fallback = slot.default_chord();
                let mut result = set_chord(slot, Some(chord.clone()));
                if let (Err(error), true) = (&result, chord != fallback) {
                    // A saved chord Windows no longer accepts: the default instead.
                    super::log(&format!("{error}; using {}", fallback.label()));
                    result = set_chord(slot, Some(fallback));
                }
                match (slot, result) {
                    (Slot::Toggle, result) => toggle = result,
                    (_, Ok(label)) => super::log(&format!("display shortcut {label}")),
                    (_, Err(error)) => super::log(&error),
                }
            }
            let _ = ready_sender.send(toggle);
            let mut message = MSG::default();
            while unsafe { GetMessageW(&mut message, None, 0, 0) }.as_bool() {
                match message.message {
                    WM_HOTKEY => {
                        if let Some(slot) = Slot::ALL.into_iter().find(|slot| message.wParam.0 == slot.hotkey_id() as usize) {
                            super::toggle(&app, slot.display());
                        }
                    }
                    WM_RECONFIGURE => {
                        let request = state().lock().ok().and_then(|mut state| state.request.take());
                        if let Some((slot, chord, reply)) = request {
                            let _ = reply.send(set_chord(slot, chord));
                        }
                    }
                    _ => {}
                }
            }
        })
        .map_err(|error| format!("Unable to start the formula shortcut: {error}"))?;
    match ready.recv_timeout(Duration::from_secs(2)) {
        Ok(Ok(label)) => {
            super::log(&format!("shortcut {label}"));
            Ok(())
        }
        Ok(Err(error)) => Err(error),
        Err(_) => Err("The formula shortcut did not start".to_string()),
    }
}

/// Settings > Shortcuts: registers the slot's `chord` (now, on the hotkey thread), or removes it
/// (None, not the inline one), and keeps it.
pub(super) fn configure(app: &AppHandle, slot: Slot, chord: Option<Chord>) -> Result<String, String> {
    match &chord {
        None if slot == Slot::Toggle => return Err("The shortcut that opens formulas cannot be removed".to_string()),
        None => {}
        Some(chord) => {
            if virtual_key(&chord.code).is_none() {
                return Err(format!("{} cannot be a system-wide shortcut", chord.code));
            }
            if !(chord.ctrl_key || chord.alt_key || chord.meta_key) {
                return Err("Add Ctrl or Alt to the shortcut".to_string());
            }
        }
    }
    let (reply, answer) = mpsc::channel();
    let thread = {
        let mut state = state().lock().map_err(|_| "The formula shortcut is unavailable".to_string())?;
        if state.current[slot.index()] == chord {
            return Ok(chord.as_ref().map(Chord::label).unwrap_or_default());
        }
        state.request = Some((slot, chord.clone(), reply));
        state.thread
    };
    if thread == 0 {
        return Err("The formula shortcut is unavailable".to_string());
    }
    unsafe { PostThreadMessageW(thread, WM_RECONFIGURE, Default::default(), Default::default()) }
        .map_err(|error| format!("The formula shortcut is unavailable: {error}"))?;
    let label = answer
        .recv_timeout(Duration::from_secs(2))
        .map_err(|_| "The formula shortcut did not answer".to_string())??;
    if let Some(path) = config_path(app, slot) {
        if let Ok(bytes) = serde_json::to_vec(&chord) {
            let _ = std::fs::write(path, bytes);
        }
    }
    super::log(&format!("shortcut {slot:?} set to {}", if label.is_empty() { "none" } else { &label }));
    Ok(label)
}

/// The label of the chord that opens an inline formula (Ctrl+Alt+M).
pub(super) fn shortcut_label() -> String {
    state()
        .lock()
        .ok()
        .and_then(|state| state.current[Slot::Toggle.index()].as_ref().map(Chord::label))
        .unwrap_or_default()
}
