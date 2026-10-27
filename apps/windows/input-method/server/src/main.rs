//! VisualTeX Input: the part of VisualTeX's Windows input method that is one process per user
//! session (as Weasel's server is for RIME): librime with VisualTeX's pinyin schema, shared by
//! every application that types with VisualTeX (its user dictionary is a single database), and
//! the candidate window. The TIP (../tip) in each application talks to it over a pipe and starts
//! it when it is not running.

#![windows_subsystem = "windows"]

mod candidates;
mod log;
mod pipe;
mod rime;
mod security;

use windows::core::HSTRING;
use windows::Win32::Foundation::{GetLastError, ERROR_ALREADY_EXISTS};
use windows::Win32::System::Threading::CreateMutexW;
use windows::Win32::UI::HiDpi::{SetProcessDpiAwarenessContext, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2};
use windows::Win32::UI::WindowsAndMessaging::{DispatchMessageW, GetMessageW, TranslateMessage, MSG};

fn main() {
    // One per user session.
    let user = std::env::var("USERNAME").unwrap_or_default();
    let _instance = unsafe { CreateMutexW(None, true, &HSTRING::from(format!("Local\\VisualTeXInput-{user}"))) };
    if unsafe { GetLastError() } == ERROR_ALREADY_EXISTS {
        return;
    }
    unsafe {
        let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    }
    log::write(&format!("started {}", env!("CARGO_PKG_VERSION")));
    // RIME starts now (its maintenance runs in the background).
    if rime::engine().is_none() {
        log::write("RIME unavailable: only formulas can be typed");
    }
    candidates::create();
    let _ = std::thread::Builder::new().name("visualtex-input-pipe".to_string()).spawn(pipe::serve);
    let mut message = MSG::default();
    unsafe {
        while GetMessageW(&mut message, None, 0, 0).as_bool() {
            let _ = TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}
