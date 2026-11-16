//! VisualTeX's macOS input method (input-method/, installed into ~/Library/Input Methods) from
//! the app's side: installing it, making it one of the selected input sources, and the socket it
//! talks to VisualTeX over.
//!
//! It is a palette input method: selected beside the user's keyboard input source (as the system's
//! press-and-hold accents are), never in its place, and hidden in the system's menus. So the
//! shortcut switches nothing and nothing is switched back: the input method already has the field
//! that has the keyboard, VisualTeX asks it where the caret is and opens the formula there. The
//! field keeps the keyboard and holds the formula's space as marked text, VisualTeX draws the
//! formula in the line, and Return inserts the text through the field's own text input system (no
//! pasteboard, no synthesized ⌘V).
//!
//! macOS 27 lets no application enable an input source (TISEnableInputSource opens System
//! Settings). A palette that no menu shows cannot be added there by hand either, so VisualTeX writes
//! it into the list of enabled third-party input sources itself (com.apple.inputsources), as System
//! Settings records one, and selects it.

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
pub(crate) const SOURCE_ID: &str = "com.visualtex.inputmethod.Formula";
const BUNDLE_NAME: &str = "VisualTeX Formula.app";
const EXECUTABLE: &str = "VisualTeXFormula";
/// The earlier input method (a keyboard input source with pinyin of its own, which took part in
/// the user's switching between input sources): taken out when found.
const OLD_SOURCE_ID: &str = "com.visualtex.inputmethod.VisualTeX";
const OLD_BUNDLE_NAME: &str = "VisualTeX Input.app";
const OLD_EXECUTABLE: &str = "VisualTeXInput";
/// Where macOS records the third-party input sources the user has enabled.
const ENABLED_DOMAIN: &str = "com.apple.inputsources";
const ENABLED_KEY: &str = "AppleEnabledThirdPartyInputSources";
/// input-method/Sources/AppLink.swift
const SOCKET_NAME: &str = "anywhere-ime.sock";
pub(super) const APP_IDENTIFIER: &str = "com.visualtex.studio";
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
    fn CFPreferencesSetAppValue(key: CFTypeRef, value: CFTypeRef, application: CFTypeRef);
    fn CFPreferencesAppSynchronize(application: CFTypeRef) -> u8;
    fn CFArrayCreateMutable(allocator: CFTypeRef, capacity: CFIndex, callbacks: *const CFCallBacks) -> CFTypeRef;
    fn CFArrayAppendValue(array: CFTypeRef, value: CFTypeRef);
    static kCFTypeDictionaryKeyCallBacks: CFCallBacks;
    static kCFTypeDictionaryValueCallBacks: CFCallBacks;
    static kCFTypeArrayCallBacks: CFCallBacks;
}

#[link(name = "Carbon", kind = "framework")]
unsafe extern "C" {
    static kTISPropertyInputSourceID: CFTypeRef;
    static kTISPropertyInputSourceIsEnabled: CFTypeRef;
    static kTISPropertyInputSourceIsSelected: CFTypeRef;
    fn TISCreateInputSourceList(properties: CFTypeRef, include_all_installed: u8) -> CFTypeRef;
    fn TISGetInputSourceProperty(source: CFTypeRef, key: CFTypeRef) -> CFTypeRef;
    fn TISSelectInputSource(source: CFTypeRef) -> i32;
    fn TISDeselectInputSource(source: CFTypeRef) -> i32;
    fn TISCopyCurrentKeyboardInputSource() -> CFTypeRef;
    fn TISCopyCurrentASCIICapableKeyboardLayoutInputSource() -> CFTypeRef;
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

fn source_flag(source: &Owned, key: CFTypeRef) -> bool {
    number_value(unsafe { TISGetInputSourceProperty(source.0, key) }) == Some(1)
}

/// The input method is one of the selected input sources: the field that has the keyboard is its
/// field too.
pub(super) fn is_enabled(app: &AppHandle) -> bool {
    on_main(app, || {
        input_source(SOURCE_ID, true)
            .map(|source| unsafe { source_flag(&source, kTISPropertyInputSourceIsEnabled) && source_flag(&source, kTISPropertyInputSourceIsSelected) })
            .unwrap_or(false)
    })
    .unwrap_or(false)
}

/// The third-party input sources macOS has as enabled, with `add` among them (as System Settings
/// records a palette input method) and without `remove`; true when the list changed.
fn write_enabled_sources(add: Option<&str>, remove: Option<&str>) -> bool {
    let (Some(domain), Some(key)) = (cf_string(ENABLED_DOMAIN), cf_string(ENABLED_KEY)) else { return false };
    let existing = preference(ENABLED_DOMAIN, ENABLED_KEY);
    let entries = existing.as_ref().map(|list| array_values(list.0)).unwrap_or_default();
    let bundle_of = |entry: CFTypeRef| string_value(dictionary_value(entry, "Bundle ID")).unwrap_or_default();
    let has_added = add.map(|id| entries.iter().any(|entry| bundle_of(*entry) == id)).unwrap_or(true);
    let has_removed = remove.map(|id| entries.iter().any(|entry| bundle_of(*entry) == id)).unwrap_or(false);
    if has_added && !has_removed {
        return false;
    }
    let Some(list) = Owned::new(unsafe { CFArrayCreateMutable(ptr::null(), 0, &kCFTypeArrayCallBacks) }) else { return false };
    for entry in &entries {
        if remove.map(|id| bundle_of(*entry) == id).unwrap_or(false) {
            continue;
        }
        unsafe { CFArrayAppendValue(list.0, *entry) };
    }
    if let (false, Some(id)) = (has_added, add) {
        let (Some(bundle_key), Some(kind_key), Some(bundle), Some(kind)) =
            (cf_string("Bundle ID"), cf_string("InputSourceKind"), cf_string(id), cf_string("Non Keyboard Input Method"))
        else {
            return false;
        };
        let keys = [bundle_key.0, kind_key.0];
        let values = [bundle.0, kind.0];
        let Some(entry) = Owned::new(unsafe {
            CFDictionaryCreate(ptr::null(), keys.as_ptr(), values.as_ptr(), 2, &kCFTypeDictionaryKeyCallBacks, &kCFTypeDictionaryValueCallBacks)
        }) else {
            return false;
        };
        unsafe { CFArrayAppendValue(list.0, entry.0) };
    }
    unsafe {
        CFPreferencesSetAppValue(key.0, list.0, domain.0);
        CFPreferencesAppSynchronize(domain.0) != 0
    }
}

/// Makes the input method one of the selected input sources (enabled first, see the module
/// comment); selecting a palette leaves the keyboard input source as it is. (enabled, selected)
fn enable_and_select(app: &AppHandle) -> (bool, bool) {
    let state = |app: &AppHandle| {
        on_main(app, || {
            input_source(SOURCE_ID, true)
                .map(|source| unsafe { (source_flag(&source, kTISPropertyInputSourceIsEnabled), source_flag(&source, kTISPropertyInputSourceIsSelected)) })
        })
        .flatten()
        .unwrap_or((false, false))
    };
    let (mut enabled, mut selected) = state(app);
    if !enabled {
        let written = write_enabled_sources(Some(SOURCE_ID), None);
        // (Text Input Sources reads the list again a moment later)
        for _ in 0..20 {
            (enabled, selected) = state(app);
            if enabled {
                break;
            }
            thread::sleep(Duration::from_millis(100));
        }
        super::log(&format!("input method enabled: written={written} enabled={enabled}"));
    }
    if enabled && !selected {
        let status = on_main(app, || input_source(SOURCE_ID, true).map(|source| unsafe { TISSelectInputSource(source.0) })).flatten();
        (enabled, selected) = state(app);
        super::log(&format!("input method selected: status={status:?} selected={selected}"));
    }
    (enabled, selected)
}

/// The earlier input method goes: off the selected and the enabled input sources, its process and
/// its bundle. (It was a keyboard input source: one the user's input source switching went through.)
fn retire_old(app: &AppHandle) {
    let Some(home) = std::env::var_os("HOME").map(PathBuf::from) else { return };
    let bundle = home.join("Library").join("Input Methods").join(OLD_BUNDLE_NAME);
    let listed = write_enabled_sources(None, Some(OLD_SOURCE_ID));
    if !listed && !bundle.exists() {
        return;
    }
    // A field still typing with it goes to the keyboard layout the user types Latin letters with.
    let was_current = on_main(app, || {
        let current = Owned::new(unsafe { TISCopyCurrentKeyboardInputSource() })?;
        let id = string_value(unsafe { TISGetInputSourceProperty(current.0, kTISPropertyInputSourceID) })?;
        if id != OLD_SOURCE_ID {
            return Some(false);
        }
        let ascii = Owned::new(unsafe { TISCopyCurrentASCIICapableKeyboardLayoutInputSource() })?;
        unsafe { TISSelectInputSource(ascii.0) };
        Some(true)
    })
    .flatten()
    .unwrap_or(false);
    let _ = on_main(app, || input_source(OLD_SOURCE_ID, true).map(|source| unsafe { TISDeselectInputSource(source.0) }));
    let _ = Command::new("/usr/bin/pkill")
        .arg("-f")
        .arg(format!("Input Methods/{OLD_BUNDLE_NAME}/Contents/MacOS/{OLD_EXECUTABLE}"))
        .status();
    let removed = !bundle.exists() || std::fs::remove_dir_all(&bundle).is_ok();
    super::log(&format!("earlier input method taken out: listed={listed} current={was_current} bundle_removed={removed}"));
}

/// Starts the input method's process (macOS starts a palette input method for nobody); it links to
/// VisualTeX within a second.
pub(super) fn launch() {
    let Some(home) = std::env::var_os("HOME").map(PathBuf::from) else { return };
    let bundle = home.join("Library").join("Input Methods").join(BUNDLE_NAME);
    if !bundle.is_dir() {
        return;
    }
    let started = Command::new("/usr/bin/open").arg("-gj").arg(&bundle).status();
    super::log(&format!("input method started: {}", started.map(|status| status.success()).unwrap_or(false)));
}

/// The input method is in ~/Library/Input Methods (the app put it there).
pub(super) fn installed() -> bool {
    std::env::var_os("HOME")
        .map(|home| Path::new(&home).join("Library").join("Input Methods").join(BUNDLE_NAME).exists())
        .unwrap_or(false)
}

/// The app's own input source becomes an ASCII-capable keyboard layout (the
/// strip's formula is typed in English): the one it had, to give back, when
/// it was not one.
pub(super) fn select_ascii(app: &AppHandle) -> Option<String> {
    on_main(app, || {
        let current = Owned::new(unsafe { TISCopyCurrentKeyboardInputSource() })?;
        let current_id = string_value(unsafe { TISGetInputSourceProperty(current.0, kTISPropertyInputSourceID) })?;
        let ascii = Owned::new(unsafe { TISCopyCurrentASCIICapableKeyboardLayoutInputSource() })?;
        let ascii_id = string_value(unsafe { TISGetInputSourceProperty(ascii.0, kTISPropertyInputSourceID) })?;
        if ascii_id == current_id {
            return None;
        }
        (unsafe { TISSelectInputSource(ascii.0) } == 0).then_some(current_id)
    })
    .flatten()
}

/// The app's own input source back to `id` (after select_ascii).
pub(super) fn restore(app: &AppHandle, id: &str) {
    let _ = select(app, id);
}

fn select(app: &AppHandle, id: &str) -> bool {
    let id = id.to_string();
    on_main(app, move || {
        input_source(&id, false).map(|source| unsafe { TISSelectInputSource(source.0) } == 0).unwrap_or(false)
    })
    .unwrap_or(false)
}

// MARK: Installation

/// Copies the input method bundled with the app into ~/Library/Input Methods when it is missing or
/// another build, registers it with Text Input Sources, makes it one of the selected input sources,
/// starts it, and takes the earlier input method out.
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
        // The new build is copied aside and swapped in at once: never is the input method's bundle missing or half
        // there (an application that reaches for an input method while it cannot be started writes it off for as
        // long as it runs).
        let staging = folder.join(".VisualTeX Formula.staging");
        if staging.exists() {
            std::fs::remove_dir_all(&staging).map_err(|error| format!("{}: {error}", staging.display()))?;
        }
        let copied = Command::new("/usr/bin/ditto")
            .arg(&bundled)
            .arg(&staging)
            .status()
            .map_err(|error| format!("ditto: {error}"))?;
        if !copied.success() {
            return Err(format!("ditto failed: {copied}"));
        }
        swap_into_place(&staging, &installed)?;
        // (the old build, now under the staging name; a running copy keeps the files it has open)
        let _ = std::fs::remove_dir_all(&staging);
        // A running copy keeps the old code: it goes, and the new one is started below.
        let _ = Command::new("/usr/bin/pkill")
            .arg("-f")
            .arg(format!("Input Methods/{BUNDLE_NAME}/Contents/MacOS/{EXECUTABLE}"))
            .status();
    }
    let registered = on_main(app, move || register(&installed)).unwrap_or(false);
    retire_old(app);
    let (enabled, selected) = enable_and_select(app);
    super::log(&format!("input method installed changed={changed} registered={registered} enabled={enabled} selected={selected}"));
    if changed {
        thread::sleep(Duration::from_millis(300));
    }
    if !is_linked() {
        launch();
    }
    Ok(())
}

/// Puts the bundle at `fresh` where `installed` is in one step (the two exchanged when `installed` exists).
fn swap_into_place(fresh: &Path, installed: &Path) -> Result<(), String> {
    if !installed.exists() {
        return std::fs::rename(fresh, installed).map_err(|error| format!("{}: {error}", installed.display()));
    }
    let from = CString::new(fresh.as_os_str().as_encoded_bytes()).map_err(|error| error.to_string())?;
    let to = CString::new(installed.as_os_str().as_encoded_bytes()).map_err(|error| error.to_string())?;
    if unsafe { libc::renamex_np(from.as_ptr(), to.as_ptr(), libc::RENAME_SWAP) } == 0 {
        return Ok(());
    }
    // A volume that cannot exchange: two renames, the bundle away for an instant only.
    let error = std::io::Error::last_os_error();
    let old = fresh.with_extension("old");
    let _ = std::fs::remove_dir_all(&old);
    std::fs::rename(installed, &old).map_err(|rename| format!("swap: {error}; {}: {rename}", installed.display()))?;
    if let Err(rename) = std::fs::rename(fresh, installed) {
        let _ = std::fs::rename(&old, installed);
        return Err(format!("{}: {rename}", installed.display()));
    }
    std::fs::rename(&old, fresh).map_err(|rename| format!("{}: {rename}", fresh.display()))
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
