//! The link to VisualTeX Input (VisualTeXInput.exe beside this DLL: RIME and the candidate
//! window, one process per user session; see ../server). One request, one reply, on the
//! application's UI thread, so every call is short and bounded: a server that does not answer in
//! time is dropped and tried again a little later, and the key goes to the application.

use serde_json::{json, Value};
use std::cell::{Cell, RefCell};
use std::path::PathBuf;
use std::time::{Duration, Instant};
use windows::core::{HSTRING, PWSTR};
use windows::Win32::Foundation::{CloseHandle, GENERIC_READ, GENERIC_WRITE, HANDLE, MAX_PATH, WAIT_OBJECT_0};
use windows::Win32::Storage::FileSystem::{CreateFileW, FILE_FLAG_OVERLAPPED, FILE_SHARE_NONE, OPEN_EXISTING};
use windows::Win32::System::LibraryLoader::GetModuleFileNameW;
use windows::Win32::System::Pipes::{SetNamedPipeHandleState, TransactNamedPipe, WaitNamedPipeW, PIPE_READMODE_MESSAGE};
use windows::Win32::System::Threading::{
    CreateEventW, CreateProcessW, WaitForSingleObject, CREATE_BREAKAWAY_FROM_JOB, DETACHED_PROCESS, PROCESS_INFORMATION,
    STARTUPINFOW,
};
use windows::Win32::System::IO::{CancelIoEx, GetOverlappedResult, OVERLAPPED};

/// How long a reply may take (the application's UI thread waits for it).
const REPLY_WAIT_MS: u32 = 400;
/// After a failure, the next attempt waits this long.
const RETRY_AFTER: Duration = Duration::from_secs(2);

pub(crate) fn pipe_name() -> String {
    let user = std::env::var("USERNAME").unwrap_or_default();
    let safe: String = user.chars().map(|ch| if ch.is_ascii_alphanumeric() { ch } else { '_' }).collect();
    format!("\\\\.\\pipe\\VisualTeXInput-{safe}")
}

/// What RIME did with a key.
#[derive(Clone, Debug, Default)]
pub(crate) struct Reply {
    pub handled: bool,
    pub commit: Option<String>,
    pub preedit: String,
    /// The caret in `preedit`, UTF-16 units.
    pub cursor: usize,
    pub ascii: bool,
}

impl Reply {
    fn from(value: &Value) -> Reply {
        Reply {
            handled: value.get("handled").and_then(Value::as_bool).unwrap_or(false),
            commit: value.get("commit").and_then(Value::as_str).filter(|text| !text.is_empty()).map(str::to_string),
            preedit: value.get("preedit").and_then(Value::as_str).unwrap_or("").to_string(),
            cursor: value.get("cursor").and_then(Value::as_u64).unwrap_or(0) as usize,
            ascii: value.get("ascii").and_then(Value::as_bool).unwrap_or(false),
        }
    }
}

#[derive(Default)]
pub(crate) struct Engine {
    pipe: Cell<isize>,
    session: Cell<u64>,
    retry_at: Cell<Option<Instant>>,
    launched_at: RefCell<Option<Instant>>,
}

fn server_path() -> Option<PathBuf> {
    let mut buffer = vec![0u16; MAX_PATH as usize * 4];
    let length = unsafe { GetModuleFileNameW(Some(crate::dll_instance().into()), &mut buffer) } as usize;
    let module = PathBuf::from(String::from_utf16_lossy(&buffer[..length]));
    Some(module.parent()?.join("VisualTeXInput.exe"))
}

/// Starts VisualTeX Input (outside the application's job, when it may leave it: the server
/// outlives the application that started it).
fn launch_server() -> bool {
    let Some(path) = server_path() else { return false };
    if !path.exists() {
        crate::log::write("VisualTeXInput.exe is missing");
        return false;
    }
    let mut command: Vec<u16> = format!("\"{}\"", path.display()).encode_utf16().chain([0]).collect();
    let startup = STARTUPINFOW { cb: std::mem::size_of::<STARTUPINFOW>() as u32, ..Default::default() };
    for flags in [DETACHED_PROCESS | CREATE_BREAKAWAY_FROM_JOB, DETACHED_PROCESS] {
        let mut process = PROCESS_INFORMATION::default();
        let started = unsafe {
            CreateProcessW(
                &HSTRING::from(path.as_os_str()),
                Some(PWSTR(command.as_mut_ptr())),
                None,
                None,
                false,
                flags,
                None,
                None,
                &startup,
                &mut process,
            )
        };
        if started.is_ok() {
            unsafe {
                let _ = CloseHandle(process.hProcess);
                let _ = CloseHandle(process.hThread);
            }
            crate::log::write("started VisualTeX Input");
            return true;
        }
    }
    crate::log::write("VisualTeX Input could not be started");
    false
}

fn open_pipe() -> Option<HANDLE> {
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
    let mode = PIPE_READMODE_MESSAGE;
    if unsafe { SetNamedPipeHandleState(handle, Some(&mode), None, None) }.is_err() {
        unsafe {
            let _ = CloseHandle(handle);
        }
        return None;
    }
    Some(handle)
}

/// One request and its reply, or None (timed out or broken).
fn transact(pipe: HANDLE, request: &[u8]) -> Option<Vec<u8>> {
    let event = unsafe { CreateEventW(None, true, false, None) }.ok()?;
    let mut overlapped = OVERLAPPED { hEvent: event, ..Default::default() };
    let mut buffer = vec![0u8; 64 * 1024];
    let mut read = 0u32;
    let started = unsafe {
        TransactNamedPipe(
            pipe,
            Some(request.as_ptr().cast()),
            request.len() as u32,
            Some(buffer.as_mut_ptr().cast()),
            buffer.len() as u32,
            &mut read,
            Some(&mut overlapped),
        )
    };
    let pending = started.as_ref().err().map(|error| error.code().0 as u32 == 0x800703E5).unwrap_or(false); // ERROR_IO_PENDING
    let result = if started.is_ok() {
        Some(read as usize)
    } else if pending {
        if unsafe { WaitForSingleObject(event, REPLY_WAIT_MS) } == WAIT_OBJECT_0 {
            unsafe { GetOverlappedResult(pipe, &overlapped, &mut read, false) }.ok().map(|_| read as usize)
        } else {
            unsafe {
                let _ = CancelIoEx(pipe, Some(&overlapped));
                let _ = GetOverlappedResult(pipe, &overlapped, &mut read, true);
            }
            None
        }
    } else {
        None
    };
    unsafe {
        let _ = CloseHandle(event);
    }
    result.map(|length| {
        buffer.truncate(length);
        buffer
    })
}

impl Engine {
    fn handle(&self) -> Option<HANDLE> {
        let raw = self.pipe.get();
        (raw != 0).then_some(HANDLE(raw as *mut _))
    }

    fn drop_pipe(&self) {
        if let Some(pipe) = self.handle() {
            unsafe {
                let _ = CloseHandle(pipe);
            }
        }
        self.pipe.set(0);
        self.session.set(0);
    }

    fn connect(&self) -> Option<HANDLE> {
        if let Some(pipe) = self.handle() {
            return Some(pipe);
        }
        if self.retry_at.get().map(|at| Instant::now() < at).unwrap_or(false) {
            return None;
        }
        let pipe = open_pipe().or_else(|| {
            // Not running: start it (at most every 10 seconds) and wait a moment for its pipe.
            let recently = self.launched_at.borrow().map(|at| at.elapsed() < Duration::from_secs(10)).unwrap_or(false);
            if recently {
                return None;
            }
            *self.launched_at.borrow_mut() = Some(Instant::now());
            if !launch_server() {
                return None;
            }
            let name = HSTRING::from(pipe_name());
            let deadline = Instant::now() + Duration::from_millis(1200);
            while Instant::now() < deadline {
                if unsafe { WaitNamedPipeW(&name, 100) }.as_bool() {
                    if let Some(pipe) = open_pipe() {
                        return Some(pipe);
                    }
                }
                std::thread::sleep(Duration::from_millis(50));
            }
            None
        });
        match pipe {
            Some(pipe) => {
                self.pipe.set(pipe.0 as isize);
                self.retry_at.set(None);
                Some(pipe)
            }
            None => {
                self.retry_at.set(Some(Instant::now() + RETRY_AFTER));
                None
            }
        }
    }

    fn raw_call(&self, request: &Value) -> Option<Value> {
        let pipe = self.connect()?;
        let bytes = serde_json::to_vec(request).ok()?;
        match transact(pipe, &bytes) {
            Some(reply) => serde_json::from_slice(&reply).ok(),
            None => {
                crate::log::write("VisualTeX Input did not answer");
                self.drop_pipe();
                self.retry_at.set(Some(Instant::now() + RETRY_AFTER));
                None
            }
        }
    }

    fn ensure_session(&self) -> Option<u64> {
        if self.session.get() != 0 {
            return Some(self.session.get());
        }
        let reply = self.raw_call(&json!({ "op": "create" }))?;
        let session = reply.get("session").and_then(Value::as_u64).filter(|session| *session != 0)?;
        self.session.set(session);
        Some(session)
    }

    /// A request about this thread's RIME session (created as needed, again if the server
    /// restarted).
    pub(crate) fn call(&self, mut request: Value) -> Option<Value> {
        for _ in 0..2 {
            let session = self.ensure_session()?;
            request["session"] = json!(session);
            let reply = self.raw_call(&request)?;
            if reply.get("error").and_then(Value::as_str) == Some("session") {
                self.session.set(0);
                continue;
            }
            return Some(reply);
        }
        None
    }

    /// A request that needs no session.
    pub(crate) fn notify(&self, request: Value) {
        let _ = self.raw_call(&request);
    }

    /// A request about this thread's session only when there is one (hiding the candidates,
    /// dropping a spelling): no connection or session is made for it.
    pub(crate) fn call_existing(&self, mut request: Value) {
        let session = self.session.get();
        if session == 0 || self.handle().is_none() {
            return;
        }
        request["session"] = json!(session);
        let _ = self.raw_call(&request);
    }

    pub(crate) fn key(&self, keysym: i32, mask: i32) -> Option<Reply> {
        self.call(json!({ "op": "key", "keysym": keysym, "mask": mask })).map(|reply| Reply::from(&reply))
    }

    pub(crate) fn close(&self) {
        if self.session.get() != 0 && self.handle().is_some() {
            let _ = self.raw_call(&json!({ "op": "destroy", "session": self.session.get() }));
        }
        self.drop_pipe();
    }
}
