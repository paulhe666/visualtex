use serde::Deserialize;
use std::collections::HashSet;
use tauri::AppHandle;

#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, PartialEq)]
#[serde(rename_all = "kebab-case")]
enum OfficeHotkeyAction {
    WordImageInline,
    WordImageDisplay,
    WordOmmlInline,
    WordOmmlDisplay,
    PowerpointSvgNew,
    PowerpointOmmlInline,
    PowerpointOmmlDisplay,
}

impl OfficeHotkeyAction {
    fn numeric_id(self) -> u32 {
        match self {
            Self::WordImageInline => 1,
            Self::WordImageDisplay => 2,
            Self::WordOmmlInline => 3,
            Self::WordOmmlDisplay => 4,
            Self::PowerpointSvgNew => 5,
            Self::PowerpointOmmlInline => 6,
            Self::PowerpointOmmlDisplay => 7,
        }
    }

    fn from_numeric_id(value: u32) -> Option<Self> {
        Some(match value {
            1 => Self::WordImageInline,
            2 => Self::WordImageDisplay,
            3 => Self::WordOmmlInline,
            4 => Self::WordOmmlDisplay,
            5 => Self::PowerpointSvgNew,
            6 => Self::PowerpointOmmlInline,
            7 => Self::PowerpointOmmlDisplay,
            _ => return None,
        })
    }

    fn bundle_id(self) -> &'static str {
        match self {
            Self::WordImageInline
            | Self::WordImageDisplay
            | Self::WordOmmlInline
            | Self::WordOmmlDisplay => "com.microsoft.Word",
            Self::PowerpointSvgNew | Self::PowerpointOmmlInline | Self::PowerpointOmmlDisplay => {
                "com.microsoft.Powerpoint"
            }
        }
    }

    fn host(self) -> super::sessions::OfficeHost {
        match self {
            Self::WordImageInline
            | Self::WordImageDisplay
            | Self::WordOmmlInline
            | Self::WordOmmlDisplay => super::sessions::OfficeHost::Word,
            Self::PowerpointSvgNew | Self::PowerpointOmmlInline | Self::PowerpointOmmlDisplay => {
                super::sessions::OfficeHost::Powerpoint
            }
        }
    }

    fn macro_name(self) -> &'static str {
        match self {
            Self::WordImageInline => "VisualTeX_CreateInline",
            Self::WordImageDisplay => "VisualTeX_CreateDisplay",
            Self::WordOmmlInline => "VisualTeX_CreateNativeInline",
            Self::WordOmmlDisplay => "VisualTeX_CreateNativeDisplay",
            Self::PowerpointSvgNew => "VisualTeX_NewFormula",
            Self::PowerpointOmmlInline => "VisualTeX_NewInlineNativeEquation",
            Self::PowerpointOmmlDisplay => "VisualTeX_NewDisplayNativeEquation",
        }
    }
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OfficeHotkeyChord {
    code: String,
    ctrl_key: bool,
    alt_key: bool,
    shift_key: bool,
    meta_key: bool,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OfficeHotkeyBinding {
    action_id: OfficeHotkeyAction,
    chord: OfficeHotkeyChord,
}

#[cfg(target_os = "macos")]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct ValidatedBinding {
    action: OfficeHotkeyAction,
    key_code: u32,
    modifiers: u32,
}

fn protected_code(chord: &OfficeHotkeyChord) -> bool {
    if !(chord.meta_key || chord.ctrl_key) || chord.alt_key {
        return false;
    }
    match chord.code.as_str() {
        "KeyC" | "KeyX" | "KeyV" | "KeyS" => !chord.shift_key,
        "KeyG" => chord.meta_key && !chord.ctrl_key && !chord.shift_key,
        "KeyK" => chord.shift_key,
        "KeyZ" | "KeyY" | "KeyN" | "KeyO" | "Comma" | "Digit0" | "Equal" | "Minus" => true,
        _ => false,
    }
}

#[cfg(target_os = "macos")]
fn validate_bindings(bindings: Vec<OfficeHotkeyBinding>) -> Result<Vec<ValidatedBinding>, String> {
    if bindings.len() > 7 {
        return Err("Office hotkeys contain too many bindings".to_string());
    }
    let mut actions = HashSet::new();
    let mut chords = HashSet::new();
    let mut validated = Vec::with_capacity(bindings.len());
    for binding in bindings {
        if !binding.chord.ctrl_key && !binding.chord.alt_key && !binding.chord.meta_key {
            return Err("An Office hotkey must include Ctrl, Option, or Command".to_string());
        }
        if protected_code(&binding.chord) {
            return Err("This shortcut is reserved by VisualTeX or macOS".to_string());
        }
        let key_code = mac_hotkey::key_code(&binding.chord.code).ok_or_else(|| {
            format!(
                "Unsupported macOS global-hotkey key: {}",
                binding.chord.code
            )
        })?;
        let modifiers = mac_hotkey::modifiers(&binding.chord);
        if !actions.insert(binding.action_id) {
            return Err("An Office action has more than one hotkey".to_string());
        }
        if !chords.insert((key_code, modifiers)) {
            return Err("The same shortcut is assigned to more than one Office action".to_string());
        }
        validated.push(ValidatedBinding {
            action: binding.action_id,
            key_code,
            modifiers,
        });
    }
    Ok(validated)
}

#[cfg(target_os = "macos")]
fn frontmost_bundle_id() -> Option<String> {
    use objc2_app_kit::NSWorkspace;
    NSWorkspace::sharedWorkspace()
        .frontmostApplication()
        .and_then(|application| application.bundleIdentifier())
        .map(|identifier| identifier.to_string())
}

#[cfg(target_os = "macos")]
fn execute(action: OfficeHotkeyAction) {
    if frontmost_bundle_id().as_deref() != Some(action.bundle_id()) {
        return;
    }
    if let Err(error) =
        super::macos_offline::run_office_hotkey_macro(action.host(), action.macro_name())
    {
        eprintln!("VisualTeX Office hotkey failed: {error}");
    }
}

#[cfg(target_os = "macos")]
mod mac_hotkey {
    use super::*;
    use std::ffi::c_void;
    use std::ptr;
    use std::sync::{
        atomic::{AtomicBool, Ordering},
        Mutex,
    };

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

    type EventHandlerProc =
        unsafe extern "C" fn(EventHandlerCallRef, EventRef, *mut c_void) -> OSStatus;

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
    const HOTKEY_SIGNATURE: u32 = u32::from_be_bytes(*b"VTOF");
    const CMD_KEY: u32 = 1 << 8;
    const SHIFT_KEY: u32 = 1 << 9;
    const OPTION_KEY: u32 = 1 << 11;
    const CONTROL_KEY: u32 = 1 << 12;

    #[derive(Clone, Copy)]
    struct Registration {
        binding: ValidatedBinding,
        reference: usize,
    }

    static REGISTRATIONS: Mutex<Vec<Registration>> = Mutex::new(Vec::new());
    static HANDLER_INSTALLED: AtomicBool = AtomicBool::new(false);

    unsafe extern "C" fn hotkey_handler(
        _call: EventHandlerCallRef,
        event: EventRef,
        _user_data: *mut c_void,
    ) -> OSStatus {
        let mut hotkey_id = EventHotKeyId {
            signature: 0,
            id: 0,
        };
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
        let Some(action) = OfficeHotkeyAction::from_numeric_id(hotkey_id.id) else {
            return EVENT_NOT_HANDLED;
        };
        tauri::async_runtime::spawn_blocking(move || execute(action));
        0
    }

    pub(super) fn initialize() -> Result<(), String> {
        if HANDLER_INSTALLED.swap(true, Ordering::SeqCst) {
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
            HANDLER_INSTALLED.store(false, Ordering::SeqCst);
            return Err(format!(
                "Unable to install the Office hotkey handler: {status}"
            ));
        }
        Ok(())
    }

    fn unregister_all(registrations: &mut Vec<Registration>) -> Result<(), String> {
        let mut first_error = None;
        for registration in registrations.drain(..) {
            let status = unsafe { UnregisterEventHotKey(registration.reference as EventHotKeyRef) };
            if status != 0 && first_error.is_none() {
                first_error = Some(format!("Unable to unregister an Office hotkey: {status}"));
            }
        }
        first_error.map_or(Ok(()), Err)
    }

    fn register_all(bindings: &[ValidatedBinding]) -> Result<Vec<Registration>, String> {
        let mut registrations = Vec::with_capacity(bindings.len());
        for binding in bindings {
            let mut reference: EventHotKeyRef = ptr::null_mut();
            let status = unsafe {
                RegisterEventHotKey(
                    binding.key_code,
                    binding.modifiers,
                    EventHotKeyId {
                        signature: HOTKEY_SIGNATURE,
                        id: binding.action.numeric_id(),
                    },
                    GetApplicationEventTarget(),
                    0,
                    &mut reference,
                )
            };
            if status != 0 || reference.is_null() {
                let _ = unregister_all(&mut registrations);
                return Err(format!(
                    "Unable to register an Office global hotkey (macOS status {status})"
                ));
            }
            registrations.push(Registration {
                binding: *binding,
                reference: reference as usize,
            });
        }
        Ok(registrations)
    }

    pub(super) fn set_bindings(bindings: Vec<ValidatedBinding>) -> Result<(), String> {
        initialize()?;
        let mut current = REGISTRATIONS
            .lock()
            .map_err(|_| "Office hotkey state is unavailable".to_string())?;
        let previous = current.iter().map(|item| item.binding).collect::<Vec<_>>();
        unregister_all(&mut current)?;
        match register_all(&bindings) {
            Ok(next) => {
                *current = next;
                Ok(())
            }
            Err(error) => {
                match register_all(&previous) {
                    Ok(restored) => *current = restored,
                    Err(restore_error) => {
                        return Err(format!(
                            "{error}; previous bindings could not be restored: {restore_error}"
                        ));
                    }
                }
                Err(error)
            }
        }
    }

    pub(super) fn modifiers(chord: &OfficeHotkeyChord) -> u32 {
        (if chord.meta_key { CMD_KEY } else { 0 })
            | (if chord.shift_key { SHIFT_KEY } else { 0 })
            | (if chord.alt_key { OPTION_KEY } else { 0 })
            | (if chord.ctrl_key { CONTROL_KEY } else { 0 })
    }

    pub(super) fn key_code(code: &str) -> Option<u32> {
        Some(match code {
            "KeyA" => 0x00,
            "KeyS" => 0x01,
            "KeyD" => 0x02,
            "KeyF" => 0x03,
            "KeyH" => 0x04,
            "KeyG" => 0x05,
            "KeyZ" => 0x06,
            "KeyX" => 0x07,
            "KeyC" => 0x08,
            "KeyV" => 0x09,
            "KeyB" => 0x0b,
            "KeyQ" => 0x0c,
            "KeyW" => 0x0d,
            "KeyE" => 0x0e,
            "KeyR" => 0x0f,
            "KeyY" => 0x10,
            "KeyT" => 0x11,
            "Digit1" => 0x12,
            "Digit2" => 0x13,
            "Digit3" => 0x14,
            "Digit4" => 0x15,
            "Digit6" => 0x16,
            "Digit5" => 0x17,
            "Equal" => 0x18,
            "Digit9" => 0x19,
            "Digit7" => 0x1a,
            "Minus" => 0x1b,
            "Digit8" => 0x1c,
            "Digit0" => 0x1d,
            "BracketRight" => 0x1e,
            "KeyO" => 0x1f,
            "KeyU" => 0x20,
            "BracketLeft" => 0x21,
            "KeyI" => 0x22,
            "KeyP" => 0x23,
            "Enter" => 0x24,
            "KeyL" => 0x25,
            "KeyJ" => 0x26,
            "Quote" => 0x27,
            "KeyK" => 0x28,
            "Semicolon" => 0x29,
            "Backslash" => 0x2a,
            "Comma" => 0x2b,
            "Slash" => 0x2c,
            "KeyN" => 0x2d,
            "KeyM" => 0x2e,
            "Period" => 0x2f,
            "Tab" => 0x30,
            "Space" => 0x31,
            "Backquote" => 0x32,
            "Backspace" => 0x33,
            "Escape" => 0x35,
            "F17" => 0x40,
            "F18" => 0x4f,
            "F19" => 0x50,
            "F20" => 0x5a,
            "F5" => 0x60,
            "F6" => 0x61,
            "F7" => 0x62,
            "F3" => 0x63,
            "F8" => 0x64,
            "F9" => 0x65,
            "F11" => 0x67,
            "F13" => 0x69,
            "F16" => 0x6a,
            "F14" => 0x6b,
            "F10" => 0x6d,
            "F12" => 0x6f,
            "F15" => 0x71,
            "Home" => 0x73,
            "PageUp" => 0x74,
            "Delete" => 0x75,
            "F4" => 0x76,
            "End" => 0x77,
            "F2" => 0x78,
            "PageDown" => 0x79,
            "F1" => 0x7a,
            "ArrowLeft" => 0x7b,
            "ArrowRight" => 0x7c,
            "ArrowDown" => 0x7d,
            "ArrowUp" => 0x7e,
            _ => return None,
        })
    }
}

pub(crate) fn initialize(_app: &AppHandle) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    mac_hotkey::initialize()?;
    Ok(())
}

#[tauri::command]
pub(crate) async fn configure_office_hotkeys(
    app: AppHandle,
    bindings: Vec<OfficeHotkeyBinding>,
) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        let validated = validate_bindings(bindings)?;
        let (sender, receiver) = tokio::sync::oneshot::channel();
        app.run_on_main_thread(move || {
            let _ = sender.send(mac_hotkey::set_bindings(validated));
        })
        .map_err(|error| format!("Unable to schedule the Office hotkey update: {error}"))?;
        return receiver
            .await
            .map_err(|_| "Office hotkey update was interrupted".to_string())?;
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = app;
        if bindings.is_empty() {
            Ok(())
        } else {
            Err("Office global hotkeys are currently supported on macOS only".to_string())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn office_actions_map_to_the_public_macros_used_by_the_ribbon() {
        let expected = [
            (
                OfficeHotkeyAction::WordImageInline,
                "VisualTeX_CreateInline",
            ),
            (
                OfficeHotkeyAction::WordImageDisplay,
                "VisualTeX_CreateDisplay",
            ),
            (
                OfficeHotkeyAction::WordOmmlInline,
                "VisualTeX_CreateNativeInline",
            ),
            (
                OfficeHotkeyAction::WordOmmlDisplay,
                "VisualTeX_CreateNativeDisplay",
            ),
            (OfficeHotkeyAction::PowerpointSvgNew, "VisualTeX_NewFormula"),
            (
                OfficeHotkeyAction::PowerpointOmmlInline,
                "VisualTeX_NewInlineNativeEquation",
            ),
            (
                OfficeHotkeyAction::PowerpointOmmlDisplay,
                "VisualTeX_NewDisplayNativeEquation",
            ),
        ];
        for (action, macro_name) in expected {
            assert_eq!(action.macro_name(), macro_name);
        }
    }
}
