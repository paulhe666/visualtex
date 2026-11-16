//! VisualTeX's Windows input method from the app's side (apps/windows/input-method):
//! - opening a formula in the field that has the keyboard: the input method's DLL is put into
//!   that field's thread (a hook on that one thread, taken off again at once), where it makes a
//!   window; the app posts the request to that window, and everything else happens there, in the
//!   field's own thread (its input method written down, VisualTeX's activated for that thread
//!   alone, and handed back when the formula ends). The app switches no input method itself;
//! - the link: a pipe the input method connects to when a formula opens
//!   (\\.\pipe\VisualTeXAnywhere-<user>, newline-delimited JSON both ways). Messages go to the
//!   input method of the open formula's field (pinned);
//! - its registration: where it is installed, and that its profile is disabled while no formula
//!   is open (the input switcher lists enabled profiles; the input method disables it itself when
//!   a formula ends, this is for a field's application that went away with a formula open).

use serde_json::{json, Value};
use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use tauri::AppHandle;
use windows::core::{s, GUID, HSTRING, PCWSTR, PWSTR};
use windows::Win32::Foundation::{
    CloseHandle, ERROR_IO_PENDING, ERROR_PIPE_CONNECTED, HANDLE, HLOCAL, HMODULE, HWND, LPARAM, LRESULT, LocalFree, WAIT_OBJECT_0, WPARAM,
};
use windows::Win32::Security::Authorization::{
    ConvertSidToStringSidW, ConvertStringSecurityDescriptorToSecurityDescriptorW, SDDL_REVISION_1,
};
use windows::Win32::Security::{GetTokenInformation, TokenUser, PSECURITY_DESCRIPTOR, SECURITY_ATTRIBUTES, TOKEN_QUERY, TOKEN_USER};
use windows::Win32::Storage::FileSystem::{ReadFile, WriteFile, FILE_FLAG_OVERLAPPED, PIPE_ACCESS_DUPLEX};
use windows::Win32::System::Com::{CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED};
use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};
use windows::Win32::System::Pipes::{ConnectNamedPipe, CreateNamedPipeW, DisconnectNamedPipe, PIPE_TYPE_BYTE, PIPE_UNLIMITED_INSTANCES, PIPE_WAIT};
use windows::Win32::System::Registry::{RegGetValueW, HKEY_LOCAL_MACHINE, RRF_RT_REG_SZ};
use windows::Win32::System::Threading::{CreateEventW, GetCurrentProcess, OpenProcessToken, WaitForSingleObject, INFINITE};
use windows::Win32::System::IO::{GetOverlappedResult, OVERLAPPED};
use windows::Win32::UI::TextServices::{CLSID_TF_InputProcessorProfiles, ITfInputProcessorProfiles};
use windows::Win32::UI::WindowsAndMessaging::{
    FindWindowExW, GetGUIThreadInfo, GetWindowThreadProcessId, PostMessageW, PostThreadMessageW, SetWindowsHookExW,
    UnhookWindowsHookEx, GUITHREADINFO, HWND_MESSAGE, WH_GETMESSAGE, WM_APP, WM_NULL,
};

/// The TIP's CLSID and profile (input-method/tip/src/lib.rs).
pub(super) const CLSID_TEXT_SERVICE: GUID = GUID::from_u128(0xb6e0a6c4_6f52_4c7e_9b7a_2f1e5a3d9c41);
pub(super) const GUID_PROFILE: GUID = GUID::from_u128(0xd3a1f7b2_8c4e_4e2a_a5d6_6b9e0c1f2a73);
const LANGID_ZH_CN: u16 = 0x0804;
/// The request to open a formula, for the window the input method makes in a field's thread
/// (input-method/tip/src/host.rs).
const WM_OPEN: u32 = WM_APP + 0x60;
/// A thread that pumps messages makes the window within this.
const HOST_WINDOW_WAIT: std::time::Duration = std::time::Duration::from_millis(400);

// MARK: The link

struct Client {
    pipe: HANDLE,
}

unsafe impl Send for Client {}
unsafe impl Sync for Client {}

impl Drop for Client {
    fn drop(&mut self) {
        unsafe {
            let _ = DisconnectNamedPipe(self.pipe);
            let _ = CloseHandle(self.pipe);
        }
    }
}

#[derive(Default)]
struct Clients {
    all: HashMap<u64, Arc<Client>>,
    /// The client that reported last.
    active: Option<u64>,
    /// The client of the open formula's field.
    pinned: Option<u64>,
}

static CLIENTS: OnceLock<Mutex<Clients>> = OnceLock::new();
static NEXT_CLIENT: AtomicU64 = AtomicU64::new(1);

fn clients() -> &'static Mutex<Clients> {
    CLIENTS.get_or_init(|| Mutex::new(Clients::default()))
}

pub(super) fn pipe_name() -> String {
    let user = std::env::var("USERNAME").unwrap_or_default();
    let safe: String = user.chars().map(|ch| if ch.is_alphanumeric() { ch } else { '_' }).collect();
    format!("\\\\.\\pipe\\VisualTeXAnywhere-{safe}")
}

fn user_sid() -> Option<String> {
    unsafe {
        let mut token = HANDLE::default();
        OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &mut token).ok()?;
        let mut length = 0u32;
        let _ = GetTokenInformation(token, TokenUser, None, 0, &mut length);
        let mut buffer = vec![0u8; length as usize];
        let read = GetTokenInformation(token, TokenUser, Some(buffer.as_mut_ptr().cast()), length, &mut length);
        let _ = CloseHandle(token);
        read.ok()?;
        let user = &*(buffer.as_ptr() as *const TOKEN_USER);
        let mut text = PWSTR::null();
        ConvertSidToStringSidW(user.User.Sid, &mut text).ok()?;
        let sid = text.to_string().ok();
        let _ = LocalFree(Some(HLOCAL(text.0.cast())));
        sid
    }
}

/// Only this user's processes (with their sandboxed ones) may connect: what VisualTeX sends is
/// typed into the user's fields.
fn pipe_attributes() -> Option<(PSECURITY_DESCRIPTOR, SECURITY_ATTRIBUTES)> {
    let user = user_sid()?;
    let sddl = format!("D:P(A;;GA;;;SY)(A;;GA;;;{user})(A;;GRGW;;;AC)(A;;GRGW;;;S-1-15-2-2)S:(ML;;NW;;;LW)");
    let mut descriptor = PSECURITY_DESCRIPTOR::default();
    unsafe {
        ConvertStringSecurityDescriptorToSecurityDescriptorW(&HSTRING::from(sddl), SDDL_REVISION_1, &mut descriptor, None).ok()?;
    }
    Some((
        descriptor,
        SECURITY_ATTRIBUTES {
            nLength: std::mem::size_of::<SECURITY_ATTRIBUTES>() as u32,
            lpSecurityDescriptor: descriptor.0,
            bInheritHandle: false.into(),
        },
    ))
}

/// Overlapped I/O on a client's pipe: its read and a write may be in flight together.
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
    let pending = started.as_ref().err().map(|error| error.code() == ERROR_IO_PENDING.to_hresult()).unwrap_or(false);
    let result = if started.is_ok() || pending {
        unsafe { GetOverlappedResult(pipe, &overlapped, &mut count, true) }.ok().map(|_| count as usize)
    } else {
        None
    };
    unsafe {
        let _ = CloseHandle(event);
    }
    result
}

/// Serves the pipe (its own thread): each TIP that connects gets a reader thread.
pub(super) fn start_link(app: &AppHandle) -> Result<(), String> {
    let app = app.clone();
    std::thread::Builder::new()
        .name("visualtex-anywhere-pipe".to_string())
        .spawn(move || {
            // The descriptor lives as long as the pipe server (this thread).
            let Some((_descriptor, attributes)) = pipe_attributes() else {
                super::log("input method link: no pipe security");
                return;
            };
            let name = HSTRING::from(pipe_name());
            loop {
                let pipe = unsafe {
                    CreateNamedPipeW(
                        &name,
                        PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED,
                        PIPE_TYPE_BYTE | PIPE_WAIT,
                        PIPE_UNLIMITED_INSTANCES,
                        64 * 1024,
                        64 * 1024,
                        0,
                        Some(&attributes),
                    )
                };
                if pipe.is_invalid() {
                    super::log("input method link: pipe not created");
                    std::thread::sleep(std::time::Duration::from_secs(2));
                    continue;
                }
                if !wait_for_client(pipe) {
                    unsafe {
                        let _ = CloseHandle(pipe);
                    }
                    continue;
                }
                let id = NEXT_CLIENT.fetch_add(1, Ordering::SeqCst);
                let client = Arc::new(Client { pipe });
                if let Ok(mut clients) = clients().lock() {
                    clients.all.insert(id, client.clone());
                }
                let reader_app = app.clone();
                let _ = std::thread::Builder::new()
                    .name("visualtex-anywhere-client".to_string())
                    .spawn(move || read_client(reader_app, id, client));
            }
        })
        .map(|_| ())
        .map_err(|error| format!("Unable to start the input method link: {error}"))
}

fn wait_for_client(pipe: HANDLE) -> bool {
    let Ok(event) = (unsafe { CreateEventW(None, true, false, None) }) else { return false };
    let mut overlapped = OVERLAPPED { hEvent: event, ..Default::default() };
    let connected = unsafe { ConnectNamedPipe(pipe, Some(&mut overlapped)) };
    let result = match connected {
        Ok(()) => true,
        Err(error) if error.code() == ERROR_PIPE_CONNECTED.to_hresult() => true,
        Err(error) if error.code() == ERROR_IO_PENDING.to_hresult() => {
            let mut count = 0u32;
            unsafe { WaitForSingleObject(event, INFINITE) == WAIT_OBJECT_0 && GetOverlappedResult(pipe, &overlapped, &mut count, false).is_ok() }
        }
        Err(_) => false,
    };
    unsafe {
        let _ = CloseHandle(event);
    }
    result
}

fn read_client(app: AppHandle, id: u64, client: Arc<Client>) {
    let mut pending: Vec<u8> = Vec::new();
    let mut buffer = vec![0u8; 64 * 1024];
    loop {
        let Some(count) = transfer(client.pipe, None, Some(&mut buffer)) else { break };
        if count == 0 {
            continue;
        }
        pending.extend_from_slice(&buffer[..count]);
        while let Some(end) = pending.iter().position(|byte| *byte == b'\n') {
            let line: Vec<u8> = pending.drain(..=end).collect();
            let Ok(mut message) = serde_json::from_slice::<Value>(&line[..line.len() - 1]) else { continue };
            let kind = message.get("t").and_then(Value::as_str).unwrap_or("");
            let forward = {
                let Ok(mut clients) = clients().lock() else { return };
                match kind {
                    // The field this client serves has the keyboard.
                    "activate" | "geometry" | "key" => {
                        clients.active = Some(id);
                        true
                    }
                    "deactivate" => {
                        let was_active = clients.active == Some(id);
                        if was_active {
                            clients.active = None;
                        }
                        was_active || clients.pinned == Some(id)
                    }
                    "hello" => true,
                    _ => clients.active == Some(id) || clients.pinned == Some(id),
                }
            };
            if forward {
                // which TIP said it (the formula pins its field's)
                if let Some(object) = message.as_object_mut() {
                    object.insert("link".to_string(), json!(id));
                }
                super::input_method_message(&app, &message);
            }
        }
    }
    let was_active = clients()
        .lock()
        .map(|mut clients| {
            clients.all.remove(&id);
            let active = clients.active == Some(id) || clients.pinned == Some(id);
            if clients.active == Some(id) {
                clients.active = None;
            }
            if clients.pinned == Some(id) {
                clients.pinned = None;
            }
            active
        })
        .unwrap_or(false);
    if was_active {
        super::input_method_message(&app, &json!({ "t": "disconnected" }));
    }
}

/// Sends a message to the TIP of the open formula's field, else to the one whose field has the
/// keyboard; false when none is linked.
pub(super) fn send(message: Value) -> bool {
    let client = {
        let Ok(clients) = clients().lock() else { return false };
        clients
            .pinned
            .and_then(|id| clients.all.get(&id).cloned())
            .or_else(|| clients.active.and_then(|id| clients.all.get(&id).cloned()))
    };
    let Some(client) = client else { return false };
    let Ok(mut line) = serde_json::to_vec(&message) else { return false };
    line.push(b'\n');
    transfer(client.pipe, Some(&line), None) == Some(line.len())
}

/// The open formula's field is the one whose TIP reported `message` (its "link"); None: no formula.
pub(super) fn pin(message: Option<&Value>) {
    let link = message.and_then(|message| message.get("link")).and_then(Value::as_u64);
    if let Ok(mut clients) = clients().lock() {
        clients.pinned = link.filter(|id| clients.all.contains_key(id));
    }
}

/// `message` came from another TIP than the open formula's field's.
pub(super) fn from_other_field(message: &Value) -> bool {
    let link = message.get("link").and_then(Value::as_u64);
    clients().lock().map(|clients| matches!((clients.pinned, link), (Some(pinned), Some(link)) if pinned != link)).unwrap_or(false)
}

// MARK: Opening a formula in a field

type HookProcedure = unsafe extern "system" fn(i32, WPARAM, LPARAM) -> LRESULT;

struct Tip {
    module: HMODULE,
    hook: HookProcedure,
    /// The class of the window this build of the input method makes in a field's thread
    /// ("VisualTeXFormulaHost." and the build: the name of the folder it is installed in).
    window_class: Vec<u16>,
}

unsafe impl Send for Tip {}
unsafe impl Sync for Tip {}

/// The registered input method's DLL, loaded here once: its hook procedure is what Windows runs in
/// the field's thread.
fn tip() -> Option<&'static Tip> {
    static TIP: OnceLock<Option<Tip>> = OnceLock::new();
    TIP.get_or_init(|| {
        let path = tip_path()?;
        let module = unsafe { LoadLibraryW(&HSTRING::from(path.as_os_str())) }.ok()?;
        let hook = unsafe { GetProcAddress(module, s!("VisualTeXHook")) }?;
        // (an earlier build's DLL, a whole input method with pinyin, has no such procedure)
        let build = path.parent().and_then(|folder| folder.file_name()).map(|name| name.to_string_lossy().to_lowercase()).unwrap_or_default();
        let mut window_class: Vec<u16> = format!("VisualTeXFormulaHost.{build}").encode_utf16().collect();
        window_class.push(0);
        Some(Tip { module, hook: unsafe { std::mem::transmute::<unsafe extern "system" fn() -> isize, HookProcedure>(hook) }, window_class })
    })
    .as_ref()
}

/// The input method is installed: registered, and this build's kind (one that opens formulas in a
/// field's thread).
pub(super) fn is_installed() -> bool {
    tip().is_some()
}

/// The thread of the control that has the keyboard in `field_window`'s application.
fn focus_thread(field_window: isize) -> u32 {
    let thread = unsafe { GetWindowThreadProcessId(HWND(field_window as *mut _), None) };
    let mut info = GUITHREADINFO { cbSize: std::mem::size_of::<GUITHREADINFO>() as u32, ..Default::default() };
    if thread != 0 && unsafe { GetGUIThreadInfo(thread, &mut info) }.is_ok() && !info.hwndFocus.is_invalid() {
        let focus = unsafe { GetWindowThreadProcessId(info.hwndFocus, None) };
        if focus != 0 {
            return focus;
        }
    }
    thread
}

/// The window of this build of the input method in `thread` (None: it has none yet).
fn host_window(tip: &Tip, thread: u32) -> Option<HWND> {
    let mut after: Option<HWND> = None;
    loop {
        let window = unsafe { FindWindowExW(Some(HWND_MESSAGE), after, PCWSTR(tip.window_class.as_ptr()), PCWSTR::null()) }.ok()?;
        if window.is_invalid() {
            return None;
        }
        if unsafe { GetWindowThreadProcessId(window, None) } == thread {
            return Some(window);
        }
        after = Some(window);
    }
}

/// Opens formula `generation` in the field that has the keyboard in `field_window`: asks the input
/// method's window in that field's thread, putting the input method there first when the thread has
/// none. Its "activate" message (the caret) follows over the link. Err: the thread does not take it
/// (a 32-bit or elevated application, one that pumps no messages).
pub(super) fn open_in_field(field_window: isize, generation: u64) -> Result<(), String> {
    let tip = tip().ok_or("the input method is not installed")?;
    let thread = focus_thread(field_window);
    if thread == 0 {
        return Err("no window has the keyboard".to_string());
    }
    let window = match host_window(tip, thread) {
        Some(window) => window,
        None => {
            let hook = unsafe { SetWindowsHookExW(WH_GETMESSAGE, Some(tip.hook), Some(tip.module.into()), thread) }
                .map_err(|error| format!("the field's thread does not take the input method: {error}"))?;
            // the thread runs the hook when it next takes a message: one is posted for that
            let _ = unsafe { PostThreadMessageW(thread, WM_NULL, WPARAM(0), LPARAM(0)) };
            let deadline = std::time::Instant::now() + HOST_WINDOW_WAIT;
            let mut window = None;
            while window.is_none() && std::time::Instant::now() < deadline {
                std::thread::sleep(std::time::Duration::from_millis(4));
                window = host_window(tip, thread);
            }
            let _ = unsafe { UnhookWindowsHookEx(hook) };
            window.ok_or("the field's thread did not start the input method")?
        }
    };
    unsafe { PostMessageW(Some(window), WM_OPEN, WPARAM(generation as usize), LPARAM(0)) }
        .map_err(|error| format!("the input method's window did not take the request: {error}"))
}

/// VisualTeX's profile is disabled (no input switcher offers it). The input method does this itself
/// when a formula ends; here for one whose application went away meanwhile.
pub(super) fn disable_profile() {
    let _ = std::thread::spawn(|| unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        let profiles: windows::core::Result<ITfInputProcessorProfiles> =
            CoCreateInstance(&CLSID_TF_InputProcessorProfiles, None, CLSCTX_INPROC_SERVER);
        if let Ok(profiles) = profiles {
            let enabled = profiles.IsEnabledLanguageProfile(&CLSID_TEXT_SERVICE, LANGID_ZH_CN, &GUID_PROFILE).map(|enabled| enabled.as_bool()).unwrap_or(false);
            if enabled {
                let result = profiles.EnableLanguageProfile(&CLSID_TEXT_SERVICE, LANGID_ZH_CN, &GUID_PROFILE, false);
                super::log(&format!("input method profile was enabled: disabled {:?}", result.map_err(|error| error.code())));
            }
        }
    })
    .join();
}

// MARK: Registration

fn guid_text(guid: &GUID) -> String {
    format!("{{{guid:?}}}")
}

/// The input method's DLL, from its COM registration.
fn tip_path() -> Option<PathBuf> {
    let key = HSTRING::from(format!("SOFTWARE\\Classes\\CLSID\\{}\\InprocServer32", guid_text(&CLSID_TEXT_SERVICE)));
    let mut buffer = vec![0u16; 1024];
    let mut size = (buffer.len() * 2) as u32;
    let status = unsafe {
        RegGetValueW(HKEY_LOCAL_MACHINE, &key, None, RRF_RT_REG_SZ, None, Some(buffer.as_mut_ptr().cast()), Some(&mut size))
    };
    if status.is_err() {
        return None;
    }
    let length = buffer.iter().position(|unit| *unit == 0).unwrap_or(buffer.len());
    let path = PathBuf::from(String::from_utf16_lossy(&buffer[..length]));
    path.is_file().then_some(path)
}
