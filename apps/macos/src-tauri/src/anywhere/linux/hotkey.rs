//! The formula's shortcuts on Linux. On X11 VisualTeX takes them globally (a grab, as macOS's Carbon
//! hot keys): pressed in any field, they switch it to VisualTeX's input method and open a formula
//! there. Under Wayland an application cannot take a global shortcut, but VisualTeX's input method
//! sees every key of the field it types into: it takes the shortcuts there (Settings' chords, sent to
//! it whenever it links) and asks VisualTeX for a formula.

use serde::{Deserialize, Serialize};
use std::sync::Mutex;
use tauri::{AppHandle, Manager};

/// The shortcuts: each opens and closes a formula, starting it as an inline or as a display row.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Slot {
    Toggle,
    Display,
}

impl Slot {
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

    fn chord_file(self) -> &'static str {
        match self {
            Slot::Toggle => "anywhere-hotkey.json",
            Slot::Display => "anywhere-display-hotkey.json",
        }
    }

    /// Ctrl+Alt+M ("math"), with Shift a display formula, as on macOS (⌃⌥M).
    fn default_chord(self) -> AnywhereChord {
        AnywhereChord {
            code: "KeyM".to_string(),
            ctrl_key: true,
            alt_key: true,
            shift_key: self == Slot::Display,
            meta_key: false,
        }
    }
}

/// A key combination as Settings records it (KeyboardEvent.code and its modifiers).
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnywhereChord {
    code: String,
    #[serde(default)]
    ctrl_key: bool,
    #[serde(default)]
    alt_key: bool,
    #[serde(default)]
    shift_key: bool,
    #[serde(default)]
    meta_key: bool,
}

impl AnywhereChord {
    /// Ctrl+Alt+Shift+Super and the key, as Linux menus write it.
    pub(crate) fn label(&self) -> String {
        let mut parts: Vec<String> = Vec::new();
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
            parts.push("Super".to_string());
        }
        let key = self
            .code
            .strip_prefix("Key")
            .or_else(|| self.code.strip_prefix("Digit"))
            .unwrap_or(&self.code);
        parts.push(key.to_string());
        parts.join("+")
    }
}

static CHORDS: Mutex<[Option<AnywhereChord>; 2]> = Mutex::new([None, None]);

/// The X11 grabs (none under Wayland): the manager and the hot key of each slot.
struct Grabs {
    manager: global_hotkey::GlobalHotKeyManager,
    keys: [Option<global_hotkey::hotkey::HotKey>; 2],
}

// the manager stays on the main thread that made it; only its handle is kept here
struct SendGrabs(Grabs);
unsafe impl Send for SendGrabs {}

static GRABS: Mutex<Option<SendGrabs>> = Mutex::new(None);

fn wayland() -> bool {
    std::env::var_os("WAYLAND_DISPLAY").is_some()
}

fn hot_key(chord: &AnywhereChord) -> Option<global_hotkey::hotkey::HotKey> {
    use global_hotkey::hotkey::{Code, HotKey, Modifiers};
    use std::str::FromStr;
    let code = Code::from_str(&chord.code).ok()?;
    let mut modifiers = Modifiers::empty();
    if chord.ctrl_key {
        modifiers |= Modifiers::CONTROL;
    }
    if chord.alt_key {
        modifiers |= Modifiers::ALT;
    }
    if chord.shift_key {
        modifiers |= Modifiers::SHIFT;
    }
    if chord.meta_key {
        modifiers |= Modifiers::SUPER;
    }
    Some(HotKey::new(Some(modifiers), code))
}

/// The slots' chords as X11 grabs (main thread); the presses open and close formulas.
fn grab_all(app: &AppHandle) {
    if wayland() {
        return;
    }
    let chords = CHORDS.lock().map(|chords| chords.clone()).unwrap_or_default();
    let Ok(mut grabs) = GRABS.lock() else { return };
    if grabs.is_none() {
        match global_hotkey::GlobalHotKeyManager::new() {
            Ok(manager) => *grabs = Some(SendGrabs(Grabs { manager, keys: [None, None] })),
            Err(error) => {
                super::log(&format!("global shortcuts unavailable: {error}"));
                return;
            }
        }
        let app = app.clone();
        let _ = std::thread::Builder::new().name("visualtex-anywhere-hotkey".to_string()).spawn(move || {
            let receiver = global_hotkey::GlobalHotKeyEvent::receiver();
            while let Ok(event) = receiver.recv() {
                if event.state != global_hotkey::HotKeyState::Pressed {
                    continue;
                }
                let display = GRABS
                    .lock()
                    .ok()
                    .and_then(|grabs| grabs.as_ref().map(|grabs| grabs.0.keys[Slot::Display.index()].map(|key| key.id()) == Some(event.id)))
                    .unwrap_or(false);
                super::toggle(&app, display);
            }
        });
    }
    let Some(SendGrabs(grabs)) = grabs.as_mut() else { return };
    for slot in [Slot::Toggle, Slot::Display] {
        if let Some(old) = grabs.keys[slot.index()].take() {
            let _ = grabs.manager.unregister(old);
        }
        if let Some(key) = chords[slot.index()].as_ref().and_then(hot_key) {
            match grabs.manager.register(key) {
                Ok(()) => grabs.keys[slot.index()] = Some(key),
                Err(error) => super::log(&format!("shortcut {slot:?} not taken: {error}")),
            }
        }
    }
}

fn grab_on_main(app: &AppHandle) {
    let worker = app.clone();
    let _ = app.run_on_main_thread(move || grab_all(&worker));
}

fn chord_file(app: &AppHandle, slot: Slot) -> Option<std::path::PathBuf> {
    let directory = app.path().app_data_dir().ok()?;
    std::fs::create_dir_all(&directory).ok()?;
    Some(directory.join(slot.chord_file()))
}

/// At startup: the chords Settings kept (or the defaults).
pub(super) fn register(app: &AppHandle) -> Result<(), String> {
    let mut chords = CHORDS.lock().map_err(|_| "shortcut state unavailable".to_string())?;
    for slot in [Slot::Toggle, Slot::Display] {
        let kept = chord_file(app, slot)
            .and_then(|path| std::fs::read_to_string(path).ok())
            .and_then(|text| serde_json::from_str::<Option<AnywhereChord>>(&text).ok());
        chords[slot.index()] = match kept {
            Some(chord) => chord,
            None => Some(slot.default_chord()),
        };
    }
    drop(chords);
    grab_on_main(app);
    Ok(())
}

/// The chords, for the input method (on every link and change).
pub(super) fn tell_input_method() {
    let Ok(chords) = CHORDS.lock() else { return };
    super::ime::send(serde_json::json!({
        "t": "shortcuts",
        "toggle": chords[Slot::Toggle.index()],
        "display": chords[Slot::Display.index()],
    }));
}

pub(super) fn shortcut_label() -> String {
    CHORDS
        .lock()
        .ok()
        .and_then(|chords| chords[Slot::Toggle.index()].as_ref().map(AnywhereChord::label))
        .unwrap_or_else(|| Slot::Toggle.default_chord().label())
}

/// Settings > Shortcuts: keeps the slot's chord (None removes it, not for the inline one) and tells
/// the input method; returns its label.
pub(crate) fn configure(app: &AppHandle, slot: Slot, chord: Option<AnywhereChord>) -> Result<String, String> {
    if chord.is_none() && slot == Slot::Toggle {
        return Err("The shortcut that opens formulas cannot be removed".to_string());
    }
    if let Ok(mut chords) = CHORDS.lock() {
        chords[slot.index()] = chord.clone();
    }
    if let Some(path) = chord_file(app, slot) {
        if let Ok(text) = serde_json::to_string(&chord) {
            let _ = std::fs::write(path, text);
        }
    }
    tell_input_method();
    grab_on_main(app);
    let label = chord.as_ref().map(AnywhereChord::label).unwrap_or_default();
    super::log(&format!("shortcut {slot:?} set to {}", if label.is_empty() { "none" } else { &label }));
    Ok(label)
}
