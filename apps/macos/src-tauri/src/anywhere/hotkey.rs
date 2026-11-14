//! The Anywhere shortcuts, registered with Carbon like the Office hotkeys
//! (office/hotkeys.rs) under their own signature. A registered hotkey needs no
//! permission and takes the key combination before the frontmost application.
//! One opens and closes a formula (an inline row), the other does the same
//! with a display row.
//!
//! The user sets them in Settings > Shortcuts (default Control-Option-M and
//! Control-Option-Shift-M); the main window sends them here
//! (configure_anywhere_hotkey) and they are kept in the application data
//! folder so the next start registers them before any page has loaded.

use serde::{Deserialize, Serialize};
use std::ffi::c_void;
use std::path::PathBuf;
use std::ptr;
use std::sync::{Mutex, OnceLock};
use tauri::{AppHandle, Manager};

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
#[derive(Clone, Copy)]
struct EventHotKeyId {
    signature: u32,
    id: u32,
}

type EventHandlerProc = unsafe extern "C" fn(EventHandlerCallRef, EventRef, *mut c_void) -> OSStatus;

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
    fn GetEventParameter(
        event: EventRef,
        name: u32,
        desired_type: u32,
        actual_type: *mut u32,
        buffer_size: usize,
        actual_size: *mut usize,
        data: *mut c_void,
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
const EVENT_PARAM_DIRECT_OBJECT: u32 = u32::from_be_bytes(*b"----");
const TYPE_EVENT_HOT_KEY_ID: u32 = u32::from_be_bytes(*b"hkid");
const EVENT_NOT_HANDLED: OSStatus = -9874;
const HOTKEY_SIGNATURE: u32 = u32::from_be_bytes(*b"VTAW");
const CMD_KEY: u32 = 1 << 8;
const SHIFT_KEY: u32 = 1 << 9;
const OPTION_KEY: u32 = 1 << 11;
const CONTROL_KEY: u32 = 1 << 12;

/// The system-wide shortcuts: each opens and closes a formula, starting it
/// as an inline or as a display row.
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

    fn hotkey_id(self) -> u32 {
        self.index() as u32 + 1
    }

    fn chord_file(self) -> &'static str {
        match self {
            Slot::Toggle => "anywhere-hotkey.json",
            Slot::Display => "anywhere-display-hotkey.json",
        }
    }

    /// Control-Option-M ("math"): macOS and common applications leave it
    /// free, unlike Control-Space (input sources) or Option-Space (launchers).
    /// Shift with it opens a display formula ($ and $$).
    fn default_chord(self) -> AnywhereChord {
        AnywhereChord {
            code: "KeyM".to_string(),
            ctrl_key: true,
            alt_key: true,
            shift_key: self == Slot::Display,
            meta_key: false,
        }
    }

    fn display(self) -> bool {
        self == Slot::Display
    }
}

/// A key combination as Settings records it (KeyboardEvent.code and its
/// modifiers).
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
    fn modifiers(&self) -> u32 {
        (if self.meta_key { CMD_KEY } else { 0 })
            | (if self.shift_key { SHIFT_KEY } else { 0 })
            | (if self.alt_key { OPTION_KEY } else { 0 })
            | (if self.ctrl_key { CONTROL_KEY } else { 0 })
    }

    /// ⌃⌥⇧⌘ and the key, as macOS menus write it.
    pub(crate) fn label(&self) -> String {
        let mut label = String::new();
        if self.ctrl_key {
            label.push('⌃');
        }
        if self.alt_key {
            label.push('⌥');
        }
        if self.shift_key {
            label.push('⇧');
        }
        if self.meta_key {
            label.push('⌘');
        }
        let key = self
            .code
            .strip_prefix("Key")
            .or_else(|| self.code.strip_prefix("Digit"))
            .unwrap_or(&self.code);
        label.push_str(key);
        label
    }
}

static APP: OnceLock<AppHandle> = OnceLock::new();
/// Each slot's registered combination and its Carbon reference (as an address).
static REGISTERED: Mutex<[Option<(AnywhereChord, usize)>; 2]> = Mutex::new([None, None]);

unsafe extern "C" fn hotkey_handler(_call: EventHandlerCallRef, event: EventRef, _user_data: *mut c_void) -> OSStatus {
    let mut hotkey_id = EventHotKeyId { signature: 0, id: 0 };
    let status = unsafe {
        GetEventParameter(
            event,
            EVENT_PARAM_DIRECT_OBJECT,
            TYPE_EVENT_HOT_KEY_ID,
            ptr::null_mut(),
            std::mem::size_of::<EventHotKeyId>(),
            ptr::null_mut(),
            (&mut hotkey_id as *mut EventHotKeyId).cast(),
        )
    };
    if status != 0 || hotkey_id.signature != HOTKEY_SIGNATURE {
        return EVENT_NOT_HANDLED;
    }
    let Some(slot) = Slot::ALL.into_iter().find(|slot| slot.hotkey_id() == hotkey_id.id) else {
        return EVENT_NOT_HANDLED;
    };
    if let Some(app) = APP.get() {
        super::toggle(app, slot.display());
    }
    0
}

fn chord_file(app: &AppHandle, slot: Slot) -> Option<PathBuf> {
    app.path().app_data_dir().ok().map(|directory| directory.join(slot.chord_file()))
}

/// The slot's saved combination: the default when none was saved (or it is
/// unreadable), none when the user cleared it ("null").
fn saved_chord(app: &AppHandle, slot: Slot) -> Option<AnywhereChord> {
    let Some(text) = chord_file(app, slot).and_then(|path| std::fs::read_to_string(path).ok()) else {
        return Some(slot.default_chord());
    };
    match serde_json::from_str::<Option<AnywhereChord>>(&text) {
        Ok(None) if slot != Slot::Toggle => None,
        Ok(Some(chord)) if crate::office::hotkeys::carbon_key_code(&chord.code).is_some() => Some(chord),
        _ => Some(slot.default_chord()),
    }
}

/// Main thread, once at startup: the handler, and the saved combinations.
pub(super) fn register(app: &AppHandle) -> Result<(), String> {
    if APP.set(app.clone()).is_err() {
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
        return Err(format!("Unable to install the VisualTeX Anywhere shortcut handler: {status}"));
    }
    let mut result = Ok(());
    for slot in Slot::ALL {
        let Some(chord) = saved_chord(app, slot) else { continue };
        let registered = set_chord(slot, Some(&chord)).or_else(|error| {
            // A saved combination macOS no longer accepts: the default instead.
            let fallback = slot.default_chord();
            if chord != fallback {
                super::log(&format!("{error}; using {}", fallback.label()));
                set_chord(slot, Some(&fallback))
            } else {
                Err(error)
            }
        });
        if let Err(error) = registered {
            if slot == Slot::Toggle {
                result = Err(error);
            } else {
                super::log(&error);
            }
        }
    }
    result
}

fn register_chord(slot: Slot, chord: &AnywhereChord) -> Result<usize, String> {
    let key_code = crate::office::hotkeys::carbon_key_code(&chord.code)
        .ok_or_else(|| format!("The key {} cannot be a global shortcut", chord.code))?;
    if chord.modifiers() & (CONTROL_KEY | OPTION_KEY | CMD_KEY) == 0 {
        return Err("A global shortcut needs Control, Option or Command".to_string());
    }
    let mut reference: EventHotKeyRef = ptr::null_mut();
    let status = unsafe {
        RegisterEventHotKey(
            key_code,
            chord.modifiers(),
            EventHotKeyId {
                signature: HOTKEY_SIGNATURE,
                id: slot.hotkey_id(),
            },
            GetApplicationEventTarget(),
            0,
            &mut reference,
        )
    };
    if status != 0 || reference.is_null() {
        return Err(format!(
            "macOS did not accept {} as the VisualTeX Anywhere shortcut (status {status}); another application may use it",
            chord.label()
        ));
    }
    Ok(reference as usize)
}

/// Main thread: replaces the slot's registered combination, or removes it
/// (None); the previous one stays when macOS refuses the new one.
fn set_chord(slot: Slot, chord: Option<&AnywhereChord>) -> Result<(), String> {
    let mut registered = REGISTERED.lock().map_err(|_| "The shortcut state is unavailable".to_string())?;
    let registered = &mut registered[slot.index()];
    if registered.as_ref().map(|(current, _)| current) == chord {
        return Ok(());
    }
    let previous = registered.take();
    if let Some((_, reference)) = &previous {
        unsafe { UnregisterEventHotKey(*reference as EventHotKeyRef) };
    }
    let Some(chord) = chord else {
        return Ok(());
    };
    match register_chord(slot, chord) {
        Ok(reference) => {
            *registered = Some((chord.clone(), reference));
            Ok(())
        }
        Err(error) => {
            if let Some((previous_chord, _)) = previous {
                if let Ok(reference) = register_chord(slot, &previous_chord) {
                    *registered = Some((previous_chord, reference));
                }
            }
            Err(error)
        }
    }
}

/// The label of the combination that opens an inline formula (⌃⌥M).
pub(super) fn shortcut_label() -> String {
    REGISTERED
        .lock()
        .ok()
        .and_then(|registered| registered[Slot::Toggle.index()].as_ref().map(|(chord, _)| chord.label()))
        .unwrap_or_else(|| Slot::Toggle.default_chord().label())
}

/// Settings > Shortcuts: registers the slot's `chord` (main thread), or
/// removes it (None, not for the inline one), and keeps it for the next
/// start; returns its label.
pub(crate) fn configure(app: &AppHandle, slot: Slot, chord: Option<AnywhereChord>) -> Result<String, String> {
    if chord.is_none() && slot == Slot::Toggle {
        return Err("The shortcut that opens formulas cannot be removed".to_string());
    }
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    let wanted = chord.clone();
    app.run_on_main_thread(move || {
        let _ = sender.send(set_chord(slot, wanted.as_ref()));
    })
    .map_err(|error| format!("Unable to reach the main thread: {error}"))?;
    receiver
        .recv_timeout(std::time::Duration::from_secs(3))
        .map_err(|_| "macOS did not answer the shortcut registration".to_string())??;
    if let Some(path) = chord_file(app, slot) {
        if let Ok(text) = serde_json::to_string(&chord) {
            let _ = std::fs::write(path, text);
        }
    }
    let label = chord.as_ref().map(AnywhereChord::label).unwrap_or_default();
    super::log(&format!("shortcut {slot:?} set to {}", if label.is_empty() { "none" } else { &label }));
    Ok(label)
}
