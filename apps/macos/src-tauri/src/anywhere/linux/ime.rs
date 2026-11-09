//! VisualTeX's input method on Linux: an IBus engine (input-method/linux, visualtex-ibus), installed
//! with the package (/usr/share/ibus/component/visualtex.xml). This side switches the field to it and
//! back (IBus's global engine, through the `ibus` command) and links with it over a Unix socket in the
//! app data folder (newline-delimited JSON, as on macOS); the engine's end must run as this user and
//! be VisualTeX's engine program.

use std::io::{BufRead, BufReader, Write};
use std::os::unix::fs::PermissionsExt;
use std::os::unix::net::{UnixListener, UnixStream};
use std::path::PathBuf;
use std::process::Command;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;
use std::thread;
use std::time::{Duration, Instant};
use tauri::{AppHandle, Manager};

/// The engine's name in IBus.
pub(crate) const SOURCE_ID: &str = "visualtex";
/// VisualTeX's own windows as the engine names their client.
pub(super) const APP_IDENTIFIER: &str = "visualtex";
/// Where a field goes back to when nothing else is known.
pub(super) const DEFAULT_SOURCE: &str = "xkb:us::eng";
const SOCKET_NAME: &str = "anywhere-ime.sock";
const ENGINE_PROGRAM: &str = "visualtex-ibus";

/// The last engine other than VisualTeX's seen, and the one the shortcut switched from.
static LAST_OTHER: Mutex<Option<String>> = Mutex::new(None);
static SWITCHED_FROM: Mutex<Option<String>> = Mutex::new(None);

fn component_dirs() -> Vec<PathBuf> {
    let mut dirs: Vec<PathBuf> = std::env::var("IBUS_COMPONENT_PATH")
        .map(|value| value.split(':').filter(|dir| !dir.is_empty()).map(PathBuf::from).collect())
        .unwrap_or_default();
    dirs.push(PathBuf::from("/usr/share/ibus/component"));
    dirs.push(PathBuf::from("/usr/local/share/ibus/component"));
    if let Some(data) = std::env::var_os("XDG_DATA_HOME")
        .map(PathBuf::from)
        .or_else(|| std::env::var_os("HOME").map(|home| PathBuf::from(home).join(".local/share")))
    {
        dirs.push(data.join("ibus/component"));
    }
    dirs
}

/// The engine is installed for IBus.
pub(super) fn installed() -> bool {
    component_dirs().iter().any(|dir| dir.join("visualtex.xml").is_file())
}

/// The shortcut can type through the input method: it is installed and IBus runs.
pub(super) fn is_enabled(_app: &AppHandle) -> bool {
    installed() && current_engine().is_some()
}

fn ibus(arguments: &[&str]) -> Option<String> {
    let output = Command::new("ibus").args(arguments).output().ok()?;
    output.status.success().then(|| String::from_utf8_lossy(&output.stdout).trim().to_string())
}

fn current_engine() -> Option<String> {
    let engine = ibus(&["engine"])?;
    if engine.is_empty() {
        return None;
    }
    if engine != SOURCE_ID {
        if let Ok(mut last) = LAST_OTHER.lock() {
            *last = Some(engine.clone());
        }
    }
    Some(engine)
}

pub(super) fn current_source(_app: &AppHandle) -> Option<String> {
    current_engine()
}

pub(super) fn recent_other_source(_app: &AppHandle) -> Option<String> {
    LAST_OTHER.lock().ok().and_then(|last| last.clone())
}

/// The pasting strip types its formula itself: nothing to switch.
pub(super) fn select_ascii(_app: &AppHandle) -> Option<String> {
    None
}

pub(super) fn restore(_app: &AppHandle, _id: &str) {}

/// Switches the field to VisualTeX's engine (from `previous`), waiting until the field takes it
/// (`taken`) or a moment passes.
pub(super) fn switch_in(_app: &AppHandle, previous: &str, taken: &dyn Fn() -> bool) -> Result<(), String> {
    if let Ok(mut from) = SWITCHED_FROM.lock() {
        *from = Some(previous.to_string()).filter(|engine| engine != SOURCE_ID);
    }
    // `ibus engine` fails when it cannot set the keyboard layout after switching (setxkbmap):
    // whether the engine is VisualTeX's tells
    Command::new("ibus")
        .args(["engine", SOURCE_ID])
        .output()
        .map_err(|error| format!("ibus: {error}"))?;
    if current_engine().as_deref() != Some(SOURCE_ID) {
        return Err(format!("IBus did not switch to {SOURCE_ID}"));
    }
    let deadline = Instant::now() + Duration::from_millis(800);
    while !taken() && Instant::now() < deadline {
        thread::sleep(Duration::from_millis(20));
    }
    Ok(())
}

/// The field back on the engine the shortcut took it from.
pub(super) fn switch_back() {
    let previous = SWITCHED_FROM.lock().ok().and_then(|mut from| from.take());
    if let Some(engine) = previous {
        let _ = Command::new("ibus").args(["engine", &engine]).output();
    }
}

/// The package installs the engine for every user; nothing to do here.
pub(super) fn install(_app: &AppHandle) -> Result<(), String> {
    if installed() {
        Ok(())
    } else {
        Err("VisualTeX's IBus engine is not installed (visualtex.xml is in no IBus component folder)".to_string())
    }
}

// MARK: The link

static LINK: Mutex<Option<(u64, UnixStream)>> = Mutex::new(None);
static NEXT_CONNECTION: AtomicU64 = AtomicU64::new(1);

fn socket_path(app: &AppHandle) -> Result<PathBuf, String> {
    let directory = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("no application data folder: {error}"))?;
    std::fs::create_dir_all(&directory).map_err(|error| format!("{}: {error}", directory.display()))?;
    Ok(directory.join(SOCKET_NAME))
}

/// Listens for the engine (one connection at a time: the newest).
pub(super) fn start_link(app: &AppHandle) -> Result<(), String> {
    let path = socket_path(app)?;
    let _ = std::fs::remove_file(&path);
    let listener = UnixListener::bind(&path).map_err(|error| format!("{}: {error}", path.display()))?;
    let _ = std::fs::set_permissions(&path, std::fs::Permissions::from_mode(0o600));
    let app = app.clone();
    thread::Builder::new()
        .name("visualtex-ime-link".to_string())
        .spawn(move || {
            for stream in listener.incoming() {
                let Ok(stream) = stream else { continue };
                if !peer_is_input_method(&stream) {
                    super::log("input method link refused: the peer is not VisualTeX's input method");
                    continue;
                }
                let Ok(writer) = stream.try_clone() else { continue };
                let id = NEXT_CONNECTION.fetch_add(1, Ordering::Relaxed);
                if let Ok(mut link) = LINK.lock() {
                    *link = Some((id, writer));
                }
                super::log(&format!("input method linked ({id})"));
                super::hotkey::tell_input_method();
                let app = app.clone();
                let _ = thread::Builder::new()
                    .name("visualtex-ime-read".to_string())
                    .spawn(move || read_messages(&app, id, stream));
            }
        })
        .map_err(|error| format!("Unable to start the input method link: {error}"))?;
    Ok(())
}

fn read_messages(app: &AppHandle, id: u64, stream: UnixStream) {
    for line in BufReader::new(stream).lines() {
        let Ok(line) = line else { break };
        if let Ok(message) = serde_json::from_str::<serde_json::Value>(&line) {
            super::input_method_message(app, &message);
        }
    }
    let current = LINK
        .lock()
        .map(|mut link| {
            let current = link.as_ref().map(|(link_id, _)| *link_id == id).unwrap_or(false);
            if current {
                *link = None;
            }
            current
        })
        .unwrap_or(false);
    if current {
        super::log(&format!("input method link closed ({id})"));
        super::input_method_message(app, &serde_json::json!({ "t": "disconnected" }));
    }
}

pub(super) fn is_linked() -> bool {
    LINK.lock().map(|link| link.is_some()).unwrap_or(false)
}

/// Sends a message to the engine; false when it is not linked.
pub(super) fn send(message: serde_json::Value) -> bool {
    let Ok(mut link) = LINK.lock() else { return false };
    let Some((_, stream)) = link.as_mut() else { return false };
    let Ok(mut line) = serde_json::to_vec(&message) else { return false };
    line.push(b'\n');
    if stream.write_all(&line).is_err() {
        *link = None;
        return false;
    }
    true
}

/// The socket's peer runs as this user and is VisualTeX's engine program.
fn peer_is_input_method(stream: &UnixStream) -> bool {
    use std::os::unix::io::AsRawFd;
    let mut credentials = libc::ucred { pid: 0, uid: 0, gid: 0 };
    let mut length = std::mem::size_of::<libc::ucred>() as libc::socklen_t;
    let status = unsafe {
        libc::getsockopt(
            stream.as_raw_fd(),
            libc::SOL_SOCKET,
            libc::SO_PEERCRED,
            (&mut credentials as *mut libc::ucred).cast(),
            &mut length,
        )
    };
    if status != 0 || credentials.uid != unsafe { libc::getuid() } {
        return false;
    }
    let program = std::fs::read_link(format!("/proc/{}/exe", credentials.pid)).unwrap_or_default();
    program
        .file_name()
        .map(|name| name.to_string_lossy().starts_with(ENGINE_PROGRAM))
        .unwrap_or(false)
}
