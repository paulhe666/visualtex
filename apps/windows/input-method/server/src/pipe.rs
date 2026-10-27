//! The pipe the TIP talks to (\\.\pipe\VisualTeXInput-<user>): one message-mode connection per
//! application thread that types with VisualTeX, one JSON request and one JSON reply at a time.
//!
//! Requests (`op`), each answered with a JSON object:
//! - hello: {ready, version}
//! - create: a RIME session for the thread {session} (0 until RIME is ready)
//! - key {session, keysym, mask}: the key to RIME → the session's state (see `state`)
//! - commit {session}: what is being spelled, committed as spelled → {commit}
//! - clear {session}: drops what is being spelled
//! - ascii {session, on?}: English (true) or Chinese, set when `on` is given → {ascii}
//! - show {session, caret}: the session's candidates under `caret` ([x, y, w, h], physical pixels)
//! - hide {session}
//! - indicator {text, caret}: "中" / "英" for a moment
//! - destroy {session}
//! - quit: the server exits (uninstall, update)

use crate::candidates::{self, Indicator, Rect, Show};
use crate::rime::{self, SessionId};
use serde_json::{json, Value};
use std::sync::Mutex;
use windows::core::HSTRING;
use windows::Win32::Foundation::{CloseHandle, ERROR_MORE_DATA, ERROR_PIPE_CONNECTED, HANDLE};
use windows::Win32::Storage::FileSystem::{ReadFile, WriteFile, PIPE_ACCESS_DUPLEX};
use windows::Win32::System::Pipes::{
    ConnectNamedPipe, CreateNamedPipeW, DisconnectNamedPipe, PIPE_READMODE_MESSAGE, PIPE_TYPE_MESSAGE,
    PIPE_UNLIMITED_INSTANCES, PIPE_WAIT,
};

/// librime calls one at a time.
static RIME: Mutex<()> = Mutex::new(());

pub fn pipe_name() -> String {
    let user = std::env::var("USERNAME").unwrap_or_default();
    let safe: String = user.chars().map(|ch| if ch.is_ascii_alphanumeric() { ch } else { '_' }).collect();
    format!("\\\\.\\pipe\\VisualTeXInput-{safe}")
}

struct Pipe(HANDLE);
unsafe impl Send for Pipe {}

/// Accepts connections forever (the calling thread).
pub fn serve() {
    let security = match crate::security::pipe_security() {
        Ok(security) => security,
        Err(error) => {
            crate::log::write(&format!("pipe security unavailable: {error}"));
            return;
        }
    };
    let name = HSTRING::from(pipe_name());
    loop {
        let pipe = unsafe {
            CreateNamedPipeW(
                &name,
                PIPE_ACCESS_DUPLEX,
                PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
                PIPE_UNLIMITED_INSTANCES,
                64 * 1024,
                64 * 1024,
                0,
                Some(&security.attributes),
            )
        };
        if pipe.is_invalid() {
            crate::log::write("pipe not created");
            std::thread::sleep(std::time::Duration::from_secs(1));
            continue;
        }
        let connected = unsafe { ConnectNamedPipe(pipe, None) };
        let connected = connected.is_ok()
            || connected.as_ref().err().map(|error| error.code() == ERROR_PIPE_CONNECTED.to_hresult()).unwrap_or(false);
        if !connected {
            unsafe {
                let _ = CloseHandle(pipe);
            }
            continue;
        }
        let client = Pipe(pipe);
        let _ = std::thread::Builder::new().name("visualtex-input-client".to_string()).spawn(move || {
            let client = client;
            serve_client(client.0);
        });
    }
}

fn read_message(pipe: HANDLE) -> Option<Vec<u8>> {
    let mut message = Vec::new();
    let mut buffer = vec![0u8; 16 * 1024];
    loop {
        let mut read = 0u32;
        let result = unsafe { ReadFile(pipe, Some(&mut buffer), Some(&mut read), None) };
        message.extend_from_slice(&buffer[..read as usize]);
        match result {
            Ok(()) => return Some(message),
            Err(error) if error.code() == ERROR_MORE_DATA.to_hresult() => continue,
            Err(_) => return None,
        }
    }
}

fn serve_client(pipe: HANDLE) {
    let mut sessions: Vec<SessionId> = Vec::new();
    while let Some(request) = read_message(pipe) {
        let reply = match serde_json::from_slice::<Value>(&request) {
            Ok(request) => handle(&request, &mut sessions),
            Err(_) => json!({ "error": "unreadable" }),
        };
        let Ok(bytes) = serde_json::to_vec(&reply) else { break };
        let mut written = 0u32;
        if unsafe { WriteFile(pipe, Some(&bytes), Some(&mut written), None) }.is_err() {
            break;
        }
    }
    // The application (or its thread) went away: its sessions go too.
    if let Some(engine) = rime::engine() {
        let _guard = RIME.lock();
        for session in &sessions {
            candidates::hide(*session);
            engine.destroy_session(*session);
        }
    }
    unsafe {
        let _ = DisconnectNamedPipe(pipe);
        let _ = CloseHandle(pipe);
    }
}

fn session_of(request: &Value) -> SessionId {
    request.get("session").and_then(Value::as_u64).unwrap_or(0) as SessionId
}

/// The session's state after a key: whether RIME took it, what it committed, what is being
/// spelled (and its caret, UTF-16), English or Chinese, and how many candidates.
fn state(engine: &rime::Engine, session: SessionId, handled: bool) -> Value {
    let commit = engine.take_commit(session);
    let state = engine.state(session);
    let mut reply = json!({
        "handled": handled,
        "preedit": state.preedit,
        "cursor": state.cursor,
        "ascii": engine.is_ascii(session),
        "candidates": state.candidates.len(),
    });
    if let Some(commit) = commit.filter(|commit| !commit.is_empty()) {
        reply["commit"] = Value::String(commit);
    }
    reply
}

fn handle(request: &Value, sessions: &mut Vec<SessionId>) -> Value {
    let op = request.get("op").and_then(Value::as_str).unwrap_or("");
    if op == "quit" {
        crate::log::write("quit requested");
        std::process::exit(0);
    }
    if op == "indicator" {
        if let Some(caret) = Rect::from_json(request.get("caret")) {
            let text = request.get("text").and_then(Value::as_str).unwrap_or("").to_string();
            candidates::indicate(Indicator { text, caret });
        }
        return json!({});
    }
    let Some(engine) = rime::engine() else {
        return json!({ "error": "rime unavailable", "ready": false });
    };
    let _guard = RIME.lock();
    let session = session_of(request);
    let known = sessions.contains(&session) && engine.exists(session);
    match op {
        "hello" => json!({ "ready": engine.is_ready(), "version": env!("CARGO_PKG_VERSION") }),
        "create" => {
            let session = engine.create_session();
            if session != 0 {
                sessions.push(session);
            }
            json!({ "session": session })
        }
        _ if !known => json!({ "error": "session" }),
        "key" => {
            let keysym = request.get("keysym").and_then(Value::as_i64).unwrap_or(0) as i32;
            let mask = request.get("mask").and_then(Value::as_i64).unwrap_or(0) as i32;
            let handled = engine.process(session, keysym, mask);
            state(engine, session, handled)
        }
        "commit" => {
            let commit = engine.commit_composition(session).unwrap_or_default();
            candidates::hide(session);
            json!({ "commit": commit })
        }
        "clear" => {
            engine.clear(session);
            candidates::hide(session);
            json!({})
        }
        "ascii" => {
            if let Some(on) = request.get("on").and_then(Value::as_bool) {
                engine.set_ascii(session, on);
            }
            json!({ "ascii": engine.is_ascii(session) })
        }
        "show" => {
            let state = engine.state(session);
            match Rect::from_json(request.get("caret")) {
                Some(caret) if !state.candidates.is_empty() => candidates::show(Show {
                    owner: session,
                    candidates: state.candidates,
                    labels: state.labels,
                    highlighted: state.highlighted,
                    caret,
                }),
                _ => candidates::hide(session),
            }
            json!({})
        }
        "hide" => {
            candidates::hide(session);
            json!({})
        }
        "destroy" => {
            candidates::hide(session);
            engine.destroy_session(session);
            sessions.retain(|known| *known != session);
            json!({})
        }
        _ => json!({ "error": "unknown op" }),
    }
}
