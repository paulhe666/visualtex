//! The link to the VisualTeX app: a named pipe the app serves
//! (\\.\pipe\VisualTeXAnywhere-<user>), newline-delimited JSON both ways.
//! A reader thread hands each message to the TIP's window on the
//! application's UI thread, where TSF may be used.

use serde_json::Value;
use std::sync::Arc;
use std::thread;
use windows::core::HSTRING;
use windows::Win32::Foundation::{CloseHandle, GENERIC_READ, GENERIC_WRITE, HANDLE, HWND, LPARAM, WPARAM};
use windows::Win32::Storage::FileSystem::{
    CreateFileW, ReadFile, WriteFile, FILE_FLAG_OVERLAPPED, FILE_SHARE_NONE, OPEN_EXISTING,
};
use windows::Win32::System::Threading::{CreateEventW, INFINITE};
use windows::Win32::System::IO::{GetOverlappedResult, OVERLAPPED};
use windows::Win32::UI::WindowsAndMessaging::{PostMessageW, WM_APP};

/// A message from VisualTeX (lparam: Box<String>).
pub(crate) const WM_LINK_MESSAGE: u32 = WM_APP + 0x51;
/// The link closed (VisualTeX quit or restarted).
pub(crate) const WM_LINK_CLOSED: u32 = WM_APP + 0x52;

pub(crate) fn pipe_name() -> String {
    let user = std::env::var("USERNAME").unwrap_or_default();
    let safe: String = user.chars().map(|ch| if ch.is_alphanumeric() { ch } else { '_' }).collect();
    format!("\\\\.\\pipe\\VisualTeXAnywhere-{safe}")
}

struct Pipe(HANDLE);
unsafe impl Send for Pipe {}
unsafe impl Sync for Pipe {}

impl Drop for Pipe {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseHandle(self.0);
        }
    }
}

pub(crate) struct Link {
    pipe: Arc<Pipe>,
}

/// Overlapped I/O on the pipe: one read and one write may be in flight
/// together (a synchronous handle would serialize them).
fn transfer(pipe: HANDLE, write: Option<&[u8]>, read: Option<&mut [u8]>) -> Option<usize> {
    let event = unsafe { CreateEventW(None, true, false, None) }.ok()?;
    let mut overlapped = OVERLAPPED { hEvent: event, ..Default::default() };
    let mut count = 0u32;
    let started = unsafe {
        match (write, read) {
            (Some(bytes), _) => WriteFile(pipe, Some(bytes), None, Some(&mut overlapped)),
            (None, Some(buffer)) => ReadFile(pipe, Some(buffer), None, Some(&mut overlapped)),
            _ => return None,
        }
    };
    let pending = started.is_err()
        && started.as_ref().err().map(|error| error.code().0 as u32 == 0x800703E5).unwrap_or(false); // ERROR_IO_PENDING
    let result = if started.is_ok() || pending {
        unsafe { GetOverlappedResult(pipe, &overlapped, &mut count, true) }.ok().map(|_| count as usize)
    } else {
        None
    };
    unsafe {
        let _ = CloseHandle(event);
    }
    let _ = INFINITE;
    result
}

#[derive(Clone, Copy)]
struct Window(isize);
unsafe impl Send for Window {}

impl Link {
    /// Connects to VisualTeX (None when it is not running) and starts
    /// reading for `window`.
    pub(crate) fn connect(window: HWND) -> Option<Link> {
        let handle = unsafe {
            CreateFileW(
                &HSTRING::from(pipe_name()),
                (GENERIC_READ | GENERIC_WRITE).0,
                FILE_SHARE_NONE,
                None,
                OPEN_EXISTING,
                FILE_FLAG_OVERLAPPED,
                None,
            )
        }
        .ok()?;
        let pipe = Arc::new(Pipe(handle));
        let reader = pipe.clone();
        let target = Window(window.0 as isize);
        let spawned = thread::Builder::new().name("visualtex-tip-link".to_string()).spawn(move || {
            let mut pending: Vec<u8> = Vec::new();
            let mut buffer = vec![0u8; 64 * 1024];
            loop {
                let Some(count) = transfer(reader.0, None, Some(&mut buffer)) else { break };
                if count == 0 {
                    continue;
                }
                pending.extend_from_slice(&buffer[..count]);
                while let Some(end) = pending.iter().position(|byte| *byte == b'\n') {
                    let line: Vec<u8> = pending.drain(..=end).collect();
                    let text = String::from_utf8_lossy(&line[..line.len() - 1]).into_owned();
                    let message = Box::into_raw(Box::new(text));
                    let posted = unsafe {
                        PostMessageW(Some(HWND(target.0 as *mut _)), WM_LINK_MESSAGE, WPARAM(0), LPARAM(message as isize))
                    };
                    if posted.is_err() {
                        drop(unsafe { Box::from_raw(message) });
                        return;
                    }
                }
            }
            unsafe {
                let _ = PostMessageW(Some(HWND(target.0 as *mut _)), WM_LINK_CLOSED, WPARAM(0), LPARAM(0));
            }
        });
        if spawned.is_err() {
            return None;
        }
        Some(Link { pipe })
    }

    /// Sends a message; false when the link is gone.
    pub(crate) fn send(&self, message: &Value) -> bool {
        let Ok(mut line) = serde_json::to_vec(message) else { return false };
        line.push(b'\n');
        transfer(self.pipe.0, Some(&line), None) == Some(line.len())
    }
}

impl Drop for Link {
    /// Stops the reader (its read in flight is cancelled; it then lets go of
    /// the pipe, which closes).
    fn drop(&mut self) {
        unsafe {
            let _ = windows::Win32::System::IO::CancelIoEx(self.pipe.0, None);
        }
    }
}
