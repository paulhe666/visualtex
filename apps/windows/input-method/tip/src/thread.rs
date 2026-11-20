//! What the input method keeps for a field's thread, where every build of it finds it.
//!
//! An application that runs across an update of VisualTeX has two builds of this DLL: the one that
//! was in its threads before, whose text service Windows keeps for a thread's life, and the new one
//! the app puts there. What one of them is asked (the formula to open) and what it wrote down (the
//! input method the field had) must reach the other, so none of it lives in a build's own memory:
//! it is kept as properties of a message-only window of a system class ("Static", one per thread,
//! found by its title), which any build reads and writes.

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{HANDLE, HWND};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, FindWindowExW, GetPropW, GetWindowThreadProcessId, RemovePropW, SetPropW, HWND_MESSAGE,
    WINDOW_EX_STYLE, WINDOW_STYLE,
};

const TITLE: PCWSTR = w!("VisualTeXFormulaThread");

/// This thread's window (made when `create`).
fn window(create: bool) -> Option<HWND> {
    let thread = unsafe { GetCurrentThreadId() };
    let mut after: Option<HWND> = None;
    loop {
        let Ok(found) = (unsafe { FindWindowExW(Some(HWND_MESSAGE), after, w!("Static"), TITLE) }) else { break };
        if found.is_invalid() {
            break;
        }
        if unsafe { GetWindowThreadProcessId(found, None) } == thread {
            return Some(found);
        }
        after = Some(found);
    }
    if !create {
        return None;
    }
    unsafe {
        CreateWindowExW(WINDOW_EX_STYLE(0), w!("Static"), TITLE, WINDOW_STYLE(0), 0, 0, 0, 0, Some(HWND_MESSAGE), None, None, None)
    }
    .ok()
}

/// A value kept for this thread (0: none).
pub(crate) fn get(name: PCWSTR) -> isize {
    window(false).map(|window| unsafe { GetPropW(window, name) }.0 as isize).unwrap_or(0)
}

pub(crate) fn set(name: PCWSTR, value: isize) {
    let Some(window) = window(value != 0) else { return };
    unsafe {
        if value == 0 {
            let _ = RemovePropW(window, name);
        } else {
            let _ = SetPropW(window, name, Some(HANDLE(value as *mut _)));
        }
    }
}

/// The value, taken away.
pub(crate) fn take(name: PCWSTR) -> isize {
    let value = get(name);
    if value != 0 {
        set(name, 0);
    }
    value
}
