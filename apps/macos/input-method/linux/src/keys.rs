//! IBus key events (X keysyms, evdev key codes, X modifier masks) as VisualTeX's editor page takes
//! them: the DOM KeyboardEvent fields `key` and `code`, the modifiers and the text a character key
//! types (the macOS input method's KeyMapping.swift).

use serde_json::{json, Value};

pub const SHIFT: u32 = 1 << 0;
pub const LOCK: u32 = 1 << 1;
pub const CONTROL: u32 = 1 << 2;
pub const MOD1: u32 = 1 << 3;
pub const MOD4: u32 = 1 << 6;
pub const SUPER: u32 = 1 << 26;
pub const HYPER: u32 = 1 << 27;
pub const META: u32 = 1 << 28;
pub const RELEASE: u32 = 1 << 30;

/// The modifiers RIME is told about (ibus-rime passes the same).
pub const RIME_MASK: u32 = SHIFT | LOCK | CONTROL | MOD1 | SUPER | RELEASE;

pub const SHIFT_L: u32 = 0xffe1;
pub const SHIFT_R: u32 = 0xffe2;
#[cfg(test)]
pub const RETURN: u32 = 0xff0d;

pub fn is_release(state: u32) -> bool {
    state & RELEASE != 0
}

pub fn has_command(state: u32) -> bool {
    state & (CONTROL | MOD1 | MOD4 | SUPER | HYPER | META) != 0
}

pub fn is_modifier(keyval: u32) -> bool {
    (0xffe1..=0xffee).contains(&keyval) || keyval == 0xfe03 || keyval == 0xff7e
}

/// evdev key codes (IBus's keycode) to DOM `code` values.
fn code(keycode: u32) -> &'static str {
    match keycode {
        1 => "Escape",
        2 => "Digit1",
        3 => "Digit2",
        4 => "Digit3",
        5 => "Digit4",
        6 => "Digit5",
        7 => "Digit6",
        8 => "Digit7",
        9 => "Digit8",
        10 => "Digit9",
        11 => "Digit0",
        12 => "Minus",
        13 => "Equal",
        14 => "Backspace",
        15 => "Tab",
        16 => "KeyQ",
        17 => "KeyW",
        18 => "KeyE",
        19 => "KeyR",
        20 => "KeyT",
        21 => "KeyY",
        22 => "KeyU",
        23 => "KeyI",
        24 => "KeyO",
        25 => "KeyP",
        26 => "BracketLeft",
        27 => "BracketRight",
        28 => "Enter",
        29 => "ControlLeft",
        30 => "KeyA",
        31 => "KeyS",
        32 => "KeyD",
        33 => "KeyF",
        34 => "KeyG",
        35 => "KeyH",
        36 => "KeyJ",
        37 => "KeyK",
        38 => "KeyL",
        39 => "Semicolon",
        40 => "Quote",
        41 => "Backquote",
        42 => "ShiftLeft",
        43 => "Backslash",
        44 => "KeyZ",
        45 => "KeyX",
        46 => "KeyC",
        47 => "KeyV",
        48 => "KeyB",
        49 => "KeyN",
        50 => "KeyM",
        51 => "Comma",
        52 => "Period",
        53 => "Slash",
        54 => "ShiftRight",
        55 => "NumpadMultiply",
        56 => "AltLeft",
        57 => "Space",
        58 => "CapsLock",
        59 => "F1",
        60 => "F2",
        61 => "F3",
        62 => "F4",
        63 => "F5",
        64 => "F6",
        65 => "F7",
        66 => "F8",
        67 => "F9",
        68 => "F10",
        69 => "NumLock",
        71 => "Numpad7",
        72 => "Numpad8",
        73 => "Numpad9",
        74 => "NumpadSubtract",
        75 => "Numpad4",
        76 => "Numpad5",
        77 => "Numpad6",
        78 => "NumpadAdd",
        79 => "Numpad1",
        80 => "Numpad2",
        81 => "Numpad3",
        82 => "Numpad0",
        83 => "NumpadDecimal",
        86 => "IntlBackslash",
        87 => "F11",
        88 => "F12",
        96 => "NumpadEnter",
        97 => "ControlRight",
        98 => "NumpadDivide",
        100 => "AltRight",
        102 => "Home",
        103 => "ArrowUp",
        104 => "PageUp",
        105 => "ArrowLeft",
        106 => "ArrowRight",
        107 => "End",
        108 => "ArrowDown",
        109 => "PageDown",
        110 => "Insert",
        111 => "Delete",
        117 => "NumpadEqual",
        125 => "MetaLeft",
        126 => "MetaRight",
        _ => "",
    }
}

/// Keys whose DOM `key` is a name, not the character they type.
fn named(keyval: u32) -> Option<&'static str> {
    Some(match keyval {
        0xff0d | 0xff8d => "Enter",
        0xff09 | 0xfe20 => "Tab",
        0xff08 => "Backspace",
        0xff1b => "Escape",
        0xffff | 0xff9f => "Delete",
        0xff50 | 0xff95 => "Home",
        0xff57 | 0xff9c => "End",
        0xff55 | 0xff9a => "PageUp",
        0xff56 | 0xff9b => "PageDown",
        0xff51 | 0xff96 => "ArrowLeft",
        0xff52 | 0xff97 => "ArrowUp",
        0xff53 | 0xff98 => "ArrowRight",
        0xff54 | 0xff99 => "ArrowDown",
        0xff63 | 0xff9e => "Insert",
        0xff0b => "Clear",
        0xffbe => "F1",
        0xffbf => "F2",
        0xffc0 => "F3",
        0xffc1 => "F4",
        0xffc2 => "F5",
        0xffc3 => "F6",
        0xffc4 => "F7",
        0xffc5 => "F8",
        0xffc6 => "F9",
        0xffc7 => "F10",
        0xffc8 => "F11",
        0xffc9 => "F12",
        0xffe1 | 0xffe2 => "Shift",
        0xffe3 | 0xffe4 => "Control",
        0xffe9 | 0xffea => "Alt",
        0xffeb | 0xffec | 0xffe7 | 0xffe8 => "Meta",
        0xffe5 => "CapsLock",
        _ => return None,
    })
}

/// The character a keysym stands for (Latin-1, Unicode keysyms and the keypad's).
pub fn character(keyval: u32) -> Option<char> {
    match keyval {
        0x20..=0x7e | 0xa0..=0xff => char::from_u32(keyval),
        0x0100_0100..=0x0110_ffff => char::from_u32(keyval - 0x0100_0000),
        0xffb0..=0xffb9 => char::from_u32(u32::from(b'0') + keyval - 0xffb0),
        0xffaa => Some('*'),
        0xffab => Some('+'),
        0xffad => Some('-'),
        0xffae => Some('.'),
        0xffaf => Some('/'),
        0xffbd => Some('='),
        0xff80 => Some(' '),
        _ => None,
    }
}

/// The message for the editor page (as the macOS input method sends it).
pub fn message(keyval: u32, keycode: u32, state: u32) -> Value {
    let control = state & CONTROL != 0;
    let meta = state & (SUPER | MOD4 | META | HYPER) != 0;
    let text = if named(keyval).is_none() && !control && !meta {
        character(keyval).filter(|c| !c.is_control()).map(String::from)
    } else {
        None
    };
    let key = named(keyval)
        .map(str::to_string)
        .or_else(|| text.clone())
        .or_else(|| character(keyval).map(String::from))
        .unwrap_or_else(|| "Unidentified".to_string());
    let mut message = json!({
        "t": "key",
        "key": key,
        "code": code(keycode),
        "shift": state & SHIFT != 0,
        "alt": state & MOD1 != 0,
        "ctrl": control,
        "meta": meta,
        "repeat": false,
    });
    if let Some(text) = text {
        message["text"] = Value::String(text);
    }
    message
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_letter_types_itself_and_names_its_code() {
        let message = message(u32::from(b'x'), 45, 0);
        assert_eq!(message["key"], "x");
        assert_eq!(message["code"], "KeyX");
        assert_eq!(message["text"], "x");
    }

    #[test]
    fn control_keys_type_nothing() {
        let message = message(u32::from(b'z'), 44, CONTROL);
        assert_eq!(message["key"], "z");
        assert!(message.get("text").is_none());
        assert_eq!(message["ctrl"], true);
    }

    #[test]
    fn return_is_named() {
        let message = message(RETURN, 28, 0);
        assert_eq!(message["key"], "Enter");
        assert!(message.get("text").is_none());
    }

    #[test]
    fn unicode_keysyms_type_their_character() {
        assert_eq!(character(0x0100_03b1), Some('α'));
        assert_eq!(character(0xffb7), Some('7'));
    }
}
