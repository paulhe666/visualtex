//! Windows key events as librime's keys: X11 keysyms and modifier masks (as
//! input-method/Sources/RimeKeys.swift does on the Mac).

use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetKeyState, GetKeyboardLayout, GetKeyboardState, ToUnicodeEx, VK_CAPITAL, VK_CONTROL, VK_LCONTROL, VK_LMENU,
    VK_LSHIFT, VK_MENU, VK_RCONTROL, VK_RMENU, VK_RSHIFT, VK_SHIFT,
};

pub(crate) const SHIFT_MASK: i32 = 1 << 0;
pub(crate) const LOCK_MASK: i32 = 1 << 1;
pub(crate) const CONTROL_MASK: i32 = 1 << 2;
pub(crate) const ALT_MASK: i32 = 1 << 3;
pub(crate) const RELEASE_MASK: i32 = 1 << 30;

pub(crate) const SHIFT_L: i32 = 0xffe1;
pub(crate) const SHIFT_R: i32 = 0xffe2;

fn extended(lparam: isize) -> bool {
    (lparam >> 24) & 1 == 1
}

fn scan(lparam: isize) -> u32 {
    ((lparam >> 16) & 0xff) as u32
}

pub(crate) fn is_shift(vk: u32) -> bool {
    vk == VK_SHIFT.0 as u32 || vk == VK_LSHIFT.0 as u32 || vk == VK_RSHIFT.0 as u32
}

pub(crate) fn is_control(vk: u32) -> bool {
    vk == VK_CONTROL.0 as u32 || vk == VK_LCONTROL.0 as u32 || vk == VK_RCONTROL.0 as u32
}

pub(crate) fn is_modifier(vk: u32) -> bool {
    is_shift(vk)
        || is_control(vk)
        || [VK_MENU.0, VK_LMENU.0, VK_RMENU.0, VK_CAPITAL.0, 0x5b, 0x5c].contains(&(vk as u16))
}

/// The modifier mask of the moment (with `release` for a key-up).
pub(crate) fn mask(release: bool) -> i32 {
    let down = |vk: u16| unsafe { GetKeyState(vk as i32) } < 0;
    let mut mask = 0;
    if down(VK_SHIFT.0) {
        mask |= SHIFT_MASK;
    }
    if unsafe { GetKeyState(VK_CAPITAL.0 as i32) } & 1 == 1 {
        mask |= LOCK_MASK;
    }
    if down(VK_CONTROL.0) {
        mask |= CONTROL_MASK;
    }
    if down(VK_MENU.0) {
        mask |= ALT_MASK;
    }
    if release {
        mask |= RELEASE_MASK;
    }
    mask
}

/// The character a key types with Shift and Caps Lock but without Control or Alt (librime takes
/// the modifiers from the mask).
fn character(vk: u32, lparam: isize) -> Option<char> {
    let mut state = [0u8; 256];
    unsafe { GetKeyboardState(&mut state).ok()? };
    for vk in [VK_CONTROL, VK_LCONTROL, VK_RCONTROL, VK_MENU, VK_LMENU, VK_RMENU] {
        state[vk.0 as usize] = 0;
    }
    let mut buffer = [0u16; 8];
    let layout = unsafe { GetKeyboardLayout(0) };
    // 0x4: leaves the keyboard's dead-key state alone (Windows 10 1607 and later).
    let count = unsafe { ToUnicodeEx(vk, scan(lparam), &state, &mut buffer, 0x4, Some(layout)) };
    if count != 1 {
        return None;
    }
    char::from_u32(buffer[0] as u32)
}

/// The keysym of a key, None for keys librime does not know.
pub(crate) fn keysym(vk: u32, lparam: isize) -> Option<i32> {
    let named = match vk {
        0x08 => 0xff08, // BackSpace
        0x09 => 0xff09, // Tab
        0x0d => {
            if extended(lparam) {
                0xff8d // KP_Enter
            } else {
                0xff0d // Return
            }
        }
        0x1b => 0xff1b, // Escape
        0x20 => 0x0020, // space
        0x21 => 0xff55, // Page_Up
        0x22 => 0xff56, // Page_Down
        0x23 => 0xff57, // End
        0x24 => 0xff50, // Home
        0x25 => 0xff51, // Left
        0x26 => 0xff52, // Up
        0x27 => 0xff53, // Right
        0x28 => 0xff54, // Down
        0x2d => 0xff63, // Insert
        0x2e => 0xffff, // Delete
        0x60..=0x69 => 0xffb0 + (vk as i32 - 0x60), // KP_0..KP_9
        0x6a => 0xffaa,
        0x6b => 0xffab,
        0x6d => 0xffad,
        0x6e => 0xffae,
        0x6f => 0xffaf,
        0x70..=0x7b => 0xffbe + (vk as i32 - 0x70), // F1..F12
        0x10 | 0xa0 | 0xa1 => {
            // Shift: left or right by scan code (0x36 is the right one).
            if vk == 0xa1 || (vk == 0x10 && scan(lparam) == 0x36) {
                SHIFT_R
            } else {
                SHIFT_L
            }
        }
        0x11 | 0xa2 | 0xa3 => {
            if vk == 0xa3 || (vk == 0x11 && extended(lparam)) {
                0xffe4
            } else {
                0xffe3
            }
        }
        0x12 | 0xa4 | 0xa5 => {
            if vk == 0xa5 || (vk == 0x12 && extended(lparam)) {
                0xffea
            } else {
                0xffe9
            }
        }
        0x14 => 0xffe5, // Caps_Lock
        _ => 0,
    };
    if named != 0 {
        return Some(named);
    }
    let character = character(vk, lparam)?;
    let value = character as u32;
    if (0x20..0x7f).contains(&value) {
        return Some(value as i32);
    }
    // Other characters: X11 Unicode keysyms.
    (value >= 0xa0).then_some(0x0100_0000 + value as i32)
}
