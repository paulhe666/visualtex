//! VisualTeX's macOS input method (input-method/, installed into
//! ~/Library/Input Methods): installing it, the switch into it that the
//! shortcut makes, and the socket it talks to VisualTeX over.
//!
//! With the input method the formula is typed the way any input method
//! types: the field keeps the keyboard and shows the formula's text as
//! marked text, VisualTeX's card under the line shows the formula itself,
//! and Return inserts the text through the field's own text input system
//! (no pasteboard, no synthesized ⌘V).
//!
//! macOS gives an input source selected from another process to the
//! frontmost application only when that application is activated again.
//! The application's own text input system switches at once, so the
//! shortcut makes VisualTeX the most recently used other input source and
//! then sends the system's "Select the previous input source" shortcut,
//! which the application handles itself. The input method switches back
//! the same way, through its client (selectMode).

use std::ffi::{c_void, CStr, CString};
use std::io::{BufRead, BufReader, Write};
use std::os::unix::fs::PermissionsExt;
use std::os::unix::io::AsRawFd;
use std::os::unix::net::{UnixListener, UnixStream};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::ptr;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{mpsc, Mutex};
use std::thread;
use std::time::Duration;
use tauri::{AppHandle, Manager};

/// The input source (and bundle) identifier, input-method/Info.plist.
pub(crate) const SOURCE_ID: &str = "com.visualtex.inputmethod.VisualTeX";
const BUNDLE_NAME: &str = "VisualTeX Input.app";
const EXECUTABLE: &str = "VisualTeXInput";
/// input-method/Sources/AppLink.swift
const SOCKET_NAME: &str = "anywhere-ime.sock";
const APP_IDENTIFIER: &str = "com.visualtex.studio";
/// "Select the previous input source" in com.apple.symbolichotkeys.
const SELECT_PREVIOUS_HOTKEY: &str = "60";
/// Its default binding: Control-Space.
const DEFAULT_SHORTCUT: Shortcut = Shortcut { key: 0x31, flags: 0x0004_0000 };

type CFTypeRef = *const c_void;
type CFIndex = isize;

const CF_STRING_ENCODING_UTF8: u32 = 0x0800_0100;
const CF_NUMBER_SINT64: CFIndex = 4;

#[repr(C)]
struct CFCallBacks {
    _opaque: [u8; 0],
}

#[link(name = "CoreFoundation", kind = "framework")]
unsafe extern "C" {
    fn CFRelease(value: CFTypeRef);
    fn CFRetain(value: CFTypeRef) -> CFTypeRef;
    fn CFGetTypeID(value: CFTypeRef) -> usize;
    fn CFStringGetTypeID() -> usize;
    fn CFNumberGetTypeID() -> usize;
    fn CFBooleanGetTypeID() -> usize;
    fn CFArrayGetTypeID() -> usize;
    fn CFDictionaryGetTypeID() -> usize;
    fn CFStringCreateWithBytes(allocator: CFTypeRef, bytes: *const u8, length: CFIndex, encoding: u32, external: u8) -> CFTypeRef;
    fn CFStringGetCString(string: CFTypeRef, buffer: *mut i8, size: CFIndex, encoding: u32) -> u8;
    fn CFNumberGetValue(number: CFTypeRef, kind: CFIndex, value: *mut c_void) -> u8;
    fn CFBooleanGetValue(boolean: CFTypeRef) -> u8;
    fn CFArrayGetCount(array: CFTypeRef) -> CFIndex;
    fn CFArrayGetValueAtIndex(array: CFTypeRef, index: CFIndex) -> CFTypeRef;
    fn CFDictionaryGetValue(dictionary: CFTypeRef, key: CFTypeRef) -> CFTypeRef;
    fn CFDictionaryCreate(
        allocator: CFTypeRef,
        keys: *const CFTypeRef,
        values: *const CFTypeRef,
        count: CFIndex,
        key_callbacks: *const CFCallBacks,
        value_callbacks: *const CFCallBacks,
    ) -> CFTypeRef;
    fn CFDataCreate(allocator: CFTypeRef, bytes: *const u8, length: CFIndex) -> CFTypeRef;
    fn CFURLCreateFromFileSystemRepresentation(allocator: CFTypeRef, path: *const u8, length: CFIndex, directory: u8) -> CFTypeRef;
    fn CFPreferencesCopyAppValue(key: CFTypeRef, application: CFTypeRef) -> CFTypeRef;
    fn CFPreferencesAppSynchronize(application: CFTypeRef) -> u8;
    static kCFTypeDictionaryKeyCallBacks: CFCallBacks;
    static kCFTypeDictionaryValueCallBacks: CFCallBacks;
}

#[link(name = "Carbon", kind = "framework")]
unsafe extern "C" {
    static kTISPropertyInputSourceID: CFTypeRef;
    static kTISPropertyInputSourceIsEnabled: CFTypeRef;
    fn TISCreateInputSourceList(properties: CFTypeRef, include_all_installed: u8) -> CFTypeRef;
    fn TISGetInputSourceProperty(source: CFTypeRef, key: CFTypeRef) -> CFTypeRef;
    fn TISSelectInputSource(source: CFTypeRef) -> i32;
    fn TISCopyCurrentKeyboardInputSource() -> CFTypeRef;
    fn TISRegisterInputSource(location: CFTypeRef) -> i32;
}

#[link(name = "Security", kind = "framework")]
unsafe extern "C" {
    static kSecGuestAttributeAudit: CFTypeRef;
    fn SecCodeCopySelf(flags: u32, code: *mut CFTypeRef) -> i32;
    fn SecCodeCopyStaticCode(code: CFTypeRef, flags: u32, static_code: *mut CFTypeRef) -> i32;
    fn SecCodeCopyDesignatedRequirement(code: CFTypeRef, flags: u32, requirement: *mut CFTypeRef) -> i32;
    fn SecRequirementCopyString(requirement: CFTypeRef, flags: u32, text: *mut CFTypeRef) -> i32;
    fn SecRequirementCreateWithString(text: CFTypeRef, flags: u32, requirement: *mut CFTypeRef) -> i32;
    fn SecCodeCopyGuestWithAttributes(host: CFTypeRef, attributes: CFTypeRef, flags: u32, guest: *mut CFTypeRef) -> i32;
    fn SecCodeCheckValidity(code: CFTypeRef, flags: u32, requirement: CFTypeRef) -> i32;
}

/// A Core Foundation object owned here (released on drop).
struct Owned(CFTypeRef);

impl Owned {
    fn new(value: CFTypeRef) -> Option<Self> {
        (!value.is_null()).then_some(Self(value))
    }
}

impl Drop for Owned {
    fn drop(&mut self) {
        unsafe { CFRelease(self.0) }
    }
}

fn cf_string(text: &str) -> Option<Owned> {
    Owned::new(unsafe {
        CFStringCreateWithBytes(ptr::null(), text.as_ptr(), text.len() as CFIndex, CF_STRING_ENCODING_UTF8, 0)
    })
}

fn is_type(value: CFTypeRef, type_id: usize) -> bool {
    !value.is_null() && unsafe { CFGetTypeID(value) } == type_id
}

fn string_value(value: CFTypeRef) -> Option<String> {
    if !is_type(value, unsafe { CFStringGetTypeID() }) {
        return None;
    }
    let mut buffer = vec![0i8; 1024];
    let ok = unsafe { CFStringGetCString(value, buffer.as_mut_ptr(), buffer.len() as CFIndex, CF_STRING_ENCODING_UTF8) };
    (ok != 0).then(|| unsafe { CStr::from_ptr(buffer.as_ptr()) }.to_string_lossy().into_owned())
}

fn number_value(value: CFTypeRef) -> Option<i64> {
    if is_type(value, unsafe { CFBooleanGetTypeID() }) {
        return Some(i64::from(unsafe { CFBooleanGetValue(value) }));
    }
    if !is_type(value, unsafe { CFNumberGetTypeID() }) {
        return None;
    }
    let mut number = 0i64;
    let ok = unsafe { CFNumberGetValue(value, CF_NUMBER_SINT64, (&mut number as *mut i64).cast()) };
    (ok != 0).then_some(number)
}

/// `dictionary[key]`, borrowed from the dictionary.
fn dictionary_value(dictionary: CFTypeRef, key: &str) -> CFTypeRef {
    if !is_type(dictionary, unsafe { CFDictionaryGetTypeID() }) {
        return ptr::null();
    }
    let Some(key) = cf_string(key) else { return ptr::null() };
    unsafe { CFDictionaryGetValue(dictionary, key.0) }
}

fn array_values(array: CFTypeRef) -> Vec<CFTypeRef> {
    if !is_type(array, unsafe { CFArrayGetTypeID() }) {
        return Vec::new();
    }
    let count = unsafe { CFArrayGetCount(array) };
    (0..count).map(|index| unsafe { CFArrayGetValueAtIndex(array, index) }).collect()
}

/// A preference of another application's domain, read fresh.
fn preference(domain: &str, key: &str) -> Option<Owned> {
    let domain = cf_string(domain)?;
    let key = cf_string(key)?;
    unsafe {
        CFPreferencesAppSynchronize(domain.0);
        Owned::new(CFPreferencesCopyAppValue(key.0, domain.0))
    }
}

// MARK: Text Input Sources

/// Runs Text Input Sources work on the main thread: on macOS 27 HIToolbox
/// asserts its queue there and stops the app when called from another.
fn on_main<T: Send + 'static>(app: &AppHandle, work: impl FnOnce() -> T + Send + 'static) -> Option<T> {
    if objc2_foundation::MainThreadMarker::new().is_some() {
        return Some(work());
    }
    let (sender, receiver) = mpsc::sync_channel(1);
    app.run_on_main_thread(move || {
        let _ = sender.send(work());
    })
    .ok()?;
    receiver.recv_timeout(Duration::from_secs(2)).ok()
}

fn input_source(id: &str, include_all_installed: bool) -> Option<Owned> {
    let key = unsafe { kTISPropertyInputSourceID };
    let value = cf_string(id)?;
    let filter = Owned::new(unsafe {
        CFDictionaryCreate(
            ptr::null(),
            &key,
            &value.0,
            1,
            &kCFTypeDictionaryKeyCallBacks,
            &kCFTypeDictionaryValueCallBacks,
        )
    })?;
    let list = Owned::new(unsafe { TISCreateInputSourceList(filter.0, u8::from(include_all_installed)) })?;
    let first = *array_values(list.0).first()?;
    Owned::new(unsafe { CFRetain(first) })
}

/// The user added VisualTeX in System Settings > Keyboard > Input Sources
/// (an application cannot: on macOS 27 TISEnableInputSource opens that page).
pub(super) fn is_enabled(app: &AppHandle) -> bool {
    on_main(app, || {
        input_source(SOURCE_ID, true)
            .map(|source| number_value(unsafe { TISGetInputSourceProperty(source.0, kTISPropertyInputSourceIsEnabled) }) == Some(1))
            .unwrap_or(false)
    })
    .unwrap_or(false)
}

/// The selected keyboard input source.
pub(super) fn current_source(app: &AppHandle) -> Option<String> {
    on_main(app, || {
        let current = Owned::new(unsafe { TISCopyCurrentKeyboardInputSource() })?;
        string_value(unsafe { TISGetInputSourceProperty(current.0, kTISPropertyInputSourceID) })
    })
    .flatten()
}

fn select(app: &AppHandle, id: &str) -> bool {
    let id = id.to_string();
    on_main(app, move || {
        input_source(&id, false).map(|source| unsafe { TISSelectInputSource(source.0) } == 0).unwrap_or(false)
    })
    .unwrap_or(false)
}

/// The input source used before the current one is VisualTeX
/// (com.apple.HIToolbox AppleInputSourceHistory, most recent first).
fn visualtex_is_previous() -> bool {
    let Some(history) = preference("com.apple.HIToolbox", "AppleInputSourceHistory") else { return false };
    array_values(history.0)
        .get(1)
        .and_then(|entry| string_value(dictionary_value(*entry, "Bundle ID")))
        .as_deref()
        == Some(SOURCE_ID)
}

#[derive(Clone, Copy, Debug)]
pub(super) struct Shortcut {
    key: u16,
    flags: u64,
}

/// The user's "Select the previous input source" shortcut; None when it is
/// turned off.
fn select_previous_shortcut() -> Option<Shortcut> {
    let Some(hotkeys) = preference("com.apple.symbolichotkeys", "AppleSymbolicHotKeys") else {
        return Some(DEFAULT_SHORTCUT);
    };
    let entry = dictionary_value(hotkeys.0, SELECT_PREVIOUS_HOTKEY);
    if entry.is_null() {
        return Some(DEFAULT_SHORTCUT);
    }
    if number_value(dictionary_value(entry, "enabled")) == Some(0) {
        return None;
    }
    let parameters = array_values(dictionary_value(dictionary_value(entry, "value"), "parameters"));
    match (
        parameters.get(1).and_then(|value| number_value(*value)),
        parameters.get(2).and_then(|value| number_value(*value)),
    ) {
        (Some(key), Some(flags)) if (0..=0x7f).contains(&key) => Some(Shortcut {
            key: key as u16,
            flags: flags as u64,
        }),
        _ => Some(DEFAULT_SHORTCUT),
    }
}

/// Switches the frontmost application to VisualTeX from `previous` (see the
/// module comment). The input method's "activate" message confirms it.
pub(super) fn switch_in(app: &AppHandle, previous: &str) -> Result<(), String> {
    let shortcut = select_previous_shortcut()
        .ok_or_else(|| "the \"Select the previous input source\" shortcut is turned off".to_string())?;
    if !visualtex_is_previous() {
        // Both selections are made from here: the application keeps its
        // input source and only the history changes.
        if !select(app, SOURCE_ID) || !select(app, previous) {
            return Err("macOS refused to select the VisualTeX input source".to_string());
        }
        thread::sleep(Duration::from_millis(30));
    }
    // The shortcut that opened the formula is still held: Control-Option
    // with the system shortcut's Control would be another shortcut.
    super::deliver::wait_for_modifiers_released(Duration::from_millis(600));
    super::deliver::post_shortcut(shortcut.key, shortcut.flags)
}

/// Switches back when the input method could not take the field (no text
/// field had the keyboard): the same shortcut, once more.
pub(super) fn switch_back() {
    if let Some(shortcut) = select_previous_shortcut() {
        let _ = super::deliver::post_shortcut(shortcut.key, shortcut.flags);
    }
}

// MARK: Installation

/// Copies the input method bundled with the app into ~/Library/Input Methods
/// when it is missing or another build, and registers it with Text Input
/// Sources. Enabling it is the user's (System Settings).
pub(super) fn install(app: &AppHandle) -> Result<(), String> {
    let bundled = app
        .path()
        .resource_dir()
        .map_err(|error| format!("no resource folder: {error}"))?
        .join("input-method")
        .join(BUNDLE_NAME);
    if !bundled.is_dir() {
        return Err(format!("the input method is not bundled ({})", bundled.display()));
    }
    let home = std::env::var_os("HOME").map(PathBuf::from).ok_or("HOME is not set")?;
    let folder = home.join("Library").join("Input Methods");
    let installed = folder.join(BUNDLE_NAME);
    let changed = build_id(&bundled) != build_id(&installed);
    if changed {
        std::fs::create_dir_all(&folder).map_err(|error| format!("Input Methods folder: {error}"))?;
        if installed.exists() {
            std::fs::remove_dir_all(&installed).map_err(|error| format!("old input method: {error}"))?;
        }
        let copied = Command::new("/usr/bin/ditto")
            .arg(&bundled)
            .arg(&installed)
            .status()
            .map_err(|error| format!("ditto: {error}"))?;
        if !copied.success() {
            return Err(format!("ditto failed: {copied}"));
        }
        // A running copy keeps the old code; macOS starts the new one when
        // it is next selected.
        let _ = Command::new("/usr/bin/pkill")
            .arg("-f")
            .arg(format!("Input Methods/{BUNDLE_NAME}/Contents/MacOS/{EXECUTABLE}"))
            .status();
    }
    let registered = on_main(app, move || register(&installed)).unwrap_or(false);
    super::log(&format!(
        "input method installed changed={changed} registered={registered} enabled={}",
        is_enabled(app)
    ));
    Ok(())
}

/// input-method build output: Contents/Resources/build-id.
fn build_id(bundle: &Path) -> Option<String> {
    std::fs::read_to_string(bundle.join("Contents").join("Resources").join("build-id"))
        .ok()
        .map(|text| text.trim().to_string())
}

fn register(bundle: &Path) -> bool {
    let Ok(path) = CString::new(bundle.as_os_str().as_encoded_bytes()) else { return false };
    let bytes = path.as_bytes();
    let Some(url) = Owned::new(unsafe {
        CFURLCreateFromFileSystemRepresentation(ptr::null(), bytes.as_ptr(), bytes.len() as CFIndex, 1)
    }) else {
        return false;
    };
    unsafe { TISRegisterInputSource(url.0) == 0 }
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

/// Listens for the input method (one connection at a time: the newest).
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

/// Sends a message to the input method; false when it is not linked.
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

/// The socket's peer is VisualTeX's input method: signed with its identifier
/// and, when this app carries a certificate, by the same certificate (this
/// app's designated requirement with the input method's identifier).
fn peer_is_input_method(stream: &UnixStream) -> bool {
    const SOL_LOCAL: i32 = 0;
    const LOCAL_PEERTOKEN: i32 = 0x006;
    let mut token = [0u8; 32];
    let mut length = token.len() as libc::socklen_t;
    let status = unsafe {
        libc::getsockopt(
            stream.as_raw_fd(),
            SOL_LOCAL,
            LOCAL_PEERTOKEN,
            token.as_mut_ptr().cast(),
            &mut length,
        )
    };
    if status != 0 {
        return false;
    }
    let Some(token_data) = Owned::new(unsafe { CFDataCreate(ptr::null(), token.as_ptr(), length as CFIndex) }) else {
        return false;
    };
    let key = unsafe { kSecGuestAttributeAudit };
    let Some(attributes) = Owned::new(unsafe {
        CFDictionaryCreate(
            ptr::null(),
            &key,
            &token_data.0,
            1,
            &kCFTypeDictionaryKeyCallBacks,
            &kCFTypeDictionaryValueCallBacks,
        )
    }) else {
        return false;
    };
    let mut guest: CFTypeRef = ptr::null();
    if unsafe { SecCodeCopyGuestWithAttributes(ptr::null(), attributes.0, 0, &mut guest) } != 0 {
        return false;
    }
    let Some(guest) = Owned::new(guest) else { return false };
    let Some(requirement_text) = cf_string(&input_method_requirement()) else { return false };
    let mut requirement: CFTypeRef = ptr::null();
    if unsafe { SecRequirementCreateWithString(requirement_text.0, 0, &mut requirement) } != 0 {
        return false;
    }
    let Some(requirement) = Owned::new(requirement) else { return false };
    unsafe { SecCodeCheckValidity(guest.0, 0, requirement.0) == 0 }
}

/// This app's designated requirement for the input method's identifier, or
/// the identifier alone for an ad-hoc signed (development) build.
fn input_method_requirement() -> String {
    let identifier_only = format!("identifier \"{SOURCE_ID}\"");
    let mut me: CFTypeRef = ptr::null();
    if unsafe { SecCodeCopySelf(0, &mut me) } != 0 {
        return identifier_only;
    }
    let Some(me) = Owned::new(me) else { return identifier_only };
    let mut static_code: CFTypeRef = ptr::null();
    if unsafe { SecCodeCopyStaticCode(me.0, 0, &mut static_code) } != 0 {
        return identifier_only;
    }
    let Some(static_code) = Owned::new(static_code) else { return identifier_only };
    let mut requirement: CFTypeRef = ptr::null();
    if unsafe { SecCodeCopyDesignatedRequirement(static_code.0, 0, &mut requirement) } != 0 {
        return identifier_only;
    }
    let Some(requirement) = Owned::new(requirement) else { return identifier_only };
    let mut text: CFTypeRef = ptr::null();
    if unsafe { SecRequirementCopyString(requirement.0, 0, &mut text) } != 0 {
        return identifier_only;
    }
    let Some(text) = Owned::new(text).and_then(|text| string_value(text.0)) else { return identifier_only };
    let own_identifier = format!("identifier \"{APP_IDENTIFIER}\"");
    if text.contains(&own_identifier) && text.contains("certificate") {
        text.replacen(&own_identifier, &identifier_only, 1)
    } else {
        identifier_only
    }
}
