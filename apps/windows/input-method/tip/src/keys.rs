//! Windows key events as VisualTeX's editor page receives them: the DOM
//! KeyboardEvent fields (key, code, modifiers) and, for a character key, the
//! text it types (as apps/macos/input-method/Sources/KeyMapping.swift).

use serde_json::{json, Value};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetKeyState, GetKeyboardLayout, GetKeyboardState, MapVirtualKeyExW, ToUnicodeEx, MAPVK_VK_TO_VSC_EX,
    VK_CONTROL, VK_LWIN, VK_MENU, VK_RWIN, VK_SHIFT,
};

/// DOM `code` by virtual key (US positions; letters and digits by name).
fn code(vk: u32) -> Option<String> {
    let name = match vk {
        0x41..=0x5a => return Some(format!("Key{}", char::from_u32(vk)?)),
        0x30..=0x39 => return Some(format!("Digit{}", vk - 0x30)),
        0x60..=0x69 => return Some(format!("Numpad{}", vk - 0x60)),
        0x70..=0x7b => return Some(format!("F{}", vk - 0x6f)),
        0x08 => "Backspace",
        0x09 => "Tab",
        0x0d => "Enter",
        0x1b => "Escape",
        0x20 => "Space",
        0x21 => "PageUp",
        0x22 => "PageDown",
        0x23 => "End",
        0x24 => "Home",
        0x25 => "ArrowLeft",
        0x26 => "ArrowUp",
        0x27 => "ArrowRight",
        0x28 => "ArrowDown",
        0x2d => "Insert",
        0x2e => "Delete",
        0x6a => "NumpadMultiply",
        0x6b => "NumpadAdd",
        0x6d => "NumpadSubtract",
        0x6e => "NumpadDecimal",
        0x6f => "NumpadDivide",
        0xba => "Semicolon",
        0xbb => "Equal",
        0xbc => "Comma",
        0xbd => "Minus",
        0xbe => "Period",
        0xbf => "Slash",
        0xc0 => "Backquote",
        0xdb => "BracketLeft",
        0xdc => "Backslash",
        0xdd => "BracketRight",
        0xde => "Quote",
        0xe2 => "IntlBackslash",
        _ => return None,
    };
    Some(name.to_string())
}

/// DOM `key` for keys that are named rather than typed.
fn named(vk: u32) -> Option<&'static str> {
    Some(match vk {
        0x08 => "Backspace",
        0x09 => "Tab",
        0x0d => "Enter",
        0x1b => "Escape",
        0x21 => "PageUp",
        0x22 => "PageDown",
        0x23 => "End",
        0x24 => "Home",
        0x25 => "ArrowLeft",
        0x26 => "ArrowUp",
        0x27 => "ArrowRight",
        0x28 => "ArrowDown",
        0x2d => "Insert",
        0x2e => "Delete",
        0x70..=0x7b => return None,
        _ => return None,
    })
}

pub(crate) struct Modifiers {
    pub shift: bool,
    pub ctrl: bool,
    pub alt: bool,
    pub win: bool,
}

pub(crate) fn modifiers() -> Modifiers {
    let down = |vk: u16| unsafe { GetKeyState(vk as i32) } < 0;
    Modifiers {
        shift: down(VK_SHIFT.0),
        ctrl: down(VK_CONTROL.0),
        alt: down(VK_MENU.0),
        win: down(VK_LWIN.0) || down(VK_RWIN.0),
    }
}

/// The text a key types with the current layout and modifiers (nil for a
/// named key or a Control chord). Asks without changing dead-key state.
fn text(vk: u32, scan: u32, mods: &Modifiers) -> Option<String> {
    if named(vk).is_some() || (mods.ctrl && !mods.alt) {
        return None;
    }
    let mut state = [0u8; 256];
    unsafe { GetKeyboardState(&mut state).ok()? };
    let mut buffer = [0u16; 8];
    let layout = unsafe { GetKeyboardLayout(0) };
    // 0x4: do not change the keyboard state (Windows 10 1607 and later)
    let count = unsafe { ToUnicodeEx(vk, scan, &state, &mut buffer, 0x4, Some(layout)) };
    if count <= 0 {
        return None;
    }
    let text = String::from_utf16_lossy(&buffer[..count as usize]);
    text.chars().all(|ch| !ch.is_control()).then_some(text)
}

/// The key message for VisualTeX's page (the macOS input method's fields).
pub(crate) fn message(vk: u32, lparam: isize, repeat: bool) -> Value {
    let mods = modifiers();
    let layout = unsafe { GetKeyboardLayout(0) };
    let scan = ((lparam >> 16) & 0xff) as u32 | if (lparam >> 24) & 1 == 1 { 0xe000 } else { 0 };
    let scan = if scan == 0 { unsafe { MapVirtualKeyExW(vk, MAPVK_VK_TO_VSC_EX, Some(layout)) } } else { scan };
    let text = text(vk, scan, &mods);
    let key = named(vk)
        .map(str::to_string)
        .or_else(|| text.clone())
        .or_else(|| match vk {
            0x41..=0x5a => char::from_u32(vk + if mods.shift { 0 } else { 32 }).map(|ch| ch.to_string()),
            0x30..=0x39 => char::from_u32(vk).map(|ch| ch.to_string()),
            0x20 => Some(" ".to_string()),
            _ => None,
        })
        .unwrap_or_else(|| "Unidentified".to_string());
    let mut message = json!({
        "t": "key",
        "key": key,
        "code": code(vk).unwrap_or_default(),
        "shift": mods.shift,
        "alt": mods.alt,
        "ctrl": mods.ctrl,
        "meta": false,
        "repeat": repeat,
    });
    if let Some(text) = text {
        message["text"] = Value::String(text);
    }
    message
}

/// Keys VisualTeX never takes from the application: the Windows key, and
/// Alt chords Windows itself uses (Alt+Tab, Alt+F4, Alt+Space menus).
pub(crate) fn belongs_to_system(vk: u32) -> bool {
    let mods = modifiers();
    if mods.win {
        return true;
    }
    mods.alt && !mods.ctrl && matches!(vk, 0x09 | 0x73 | 0x20 | 0x1b)
}

/// Control chords that act on the open formula rather than the application:
/// Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z (undo, redo), Ctrl+V (paste: an image is
/// recognized), Ctrl+Enter chords (rows), and every Ctrl+Alt chord (the
/// formula's shortcuts, Settings > Keyboard shortcuts).
pub(crate) fn belongs_to_formula(vk: u32) -> bool {
    let mods = modifiers();
    if !mods.ctrl {
        return true;
    }
    mods.alt || matches!(vk, 0x5a | 0x59 | 0x56 | 0x0d)
}
