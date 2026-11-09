//! VisualTeX's Windows input method from the app's side (apps/windows/input-method):
//! - the link: a pipe the TIP in every application connects to (\\.\pipe\VisualTeXAnywhere-<user>,
//!   newline-delimited JSON both ways). Messages go to the TIP whose field has the keyboard (the
//!   last one to report it), or while a formula is open to the TIP of its field (pinned: another
//!   application's TIP that reports itself meanwhile, reconnecting or getting a focus of its own,
//!   must not take the formula's messages, which left the formula's field without its anchor and
//!   the card without its keys);
//! - its TSF profile: whether it is installed, which input method the field has, switching the
//!   field to VisualTeX (the shortcut) and back;
//! - VisualTeX Input (the per-user RIME server beside the TIP), started with VisualTeX.

use serde_json::{json, Value};
use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use tauri::AppHandle;
use windows::core::{GUID, HSTRING, PWSTR};
use windows::Win32::Foundation::{CloseHandle, ERROR_IO_PENDING, ERROR_PIPE_CONNECTED, HANDLE, HLOCAL, LocalFree, WAIT_OBJECT_0};
use windows::Win32::Security::Authorization::{
    ConvertSidToStringSidW, ConvertStringSecurityDescriptorToSecurityDescriptorW, SDDL_REVISION_1,
};
use windows::Win32::Security::{GetTokenInformation, TokenUser, PSECURITY_DESCRIPTOR, SECURITY_ATTRIBUTES, TOKEN_QUERY, TOKEN_USER};
use windows::Win32::Storage::FileSystem::{ReadFile, WriteFile, FILE_FLAG_OVERLAPPED, PIPE_ACCESS_DUPLEX};
use windows::Win32::System::Com::{CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED};
use windows::Win32::System::Pipes::{ConnectNamedPipe, CreateNamedPipeW, DisconnectNamedPipe, PIPE_TYPE_BYTE, PIPE_UNLIMITED_INSTANCES, PIPE_WAIT};
use windows::Win32::System::Registry::{RegGetValueW, HKEY_LOCAL_MACHINE, RRF_RT_REG_SZ};
use windows::Win32::System::Threading::{CreateEventW, GetCurrentProcess, OpenProcessToken, WaitForSingleObject, INFINITE};
use windows::Win32::System::IO::{GetOverlappedResult, OVERLAPPED};
use windows::Win32::UI::Input::KeyboardAndMouse::HKL;
use windows::Win32::UI::TextServices::{
    CLSID_TF_InputProcessorProfiles, ITfInputProcessorProfileMgr, GUID_TFCAT_TIP_KEYBOARD, TF_INPUTPROCESSORPROFILE,
    TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE, TF_IPPMF_FORSESSION, TF_IPP_FLAG_ACTIVE, TF_IPP_FLAG_ENABLED,
    TF_PROFILETYPE_INPUTPROCESSOR, TF_PROFILETYPE_KEYBOARDLAYOUT,
};

/// The TIP's CLSID and profile (input-method/tip/src/lib.rs).
pub(super) const CLSID_TEXT_SERVICE: GUID = GUID::from_u128(0xb6e0a6c4_6f52_4c7e_9b7a_2f1e5a3d9c41);
pub(super) const GUID_PROFILE: GUID = GUID::from_u128(0xd3a1f7b2_8c4e_4e2a_a5d6_6b9e0c1f2a73);
const LANGID_ZH_CN: u16 = 0x0804;

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
    /// The client whose field has the keyboard.
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
                    "activate" | "geometry" | "key" | "compose" => {
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

pub(super) fn is_linked() -> bool {
    clients().lock().map(|clients| clients.active.is_some()).unwrap_or(false)
}

// MARK: Profiles (a COM apartment of its own)

fn with_profiles<T>(work: impl FnOnce(&ITfInputProcessorProfileMgr) -> T + Send + 'static) -> Option<T>
where
    T: Send + 'static,
{
    std::thread::spawn(move || {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }
        let manager: windows::core::Result<ITfInputProcessorProfileMgr> =
            unsafe { CoCreateInstance(&CLSID_TF_InputProcessorProfiles, None, CLSCTX_INPROC_SERVER) };
        manager.ok().map(|manager| work(&manager))
    })
    .join()
    .ok()
    .flatten()
}

fn guid_text(guid: &GUID) -> String {
    format!("{{{guid:?}}}")
}

/// VisualTeX's input method is installed (registered) and enabled for the user.
pub(super) fn is_enabled() -> bool {
    with_profiles(|manager| {
        let mut profile = TF_INPUTPROCESSORPROFILE::default();
        let found = unsafe {
            manager.GetProfile(
                TF_PROFILETYPE_INPUTPROCESSOR,
                LANGID_ZH_CN,
                &CLSID_TEXT_SERVICE,
                &GUID_PROFILE,
                HKL::default(),
                &mut profile,
            )
        };
        found.is_ok() && profile.dwFlags & TF_IPP_FLAG_ENABLED != 0
    })
    .unwrap_or(false)
}

/// The keyboard input method the user types with now, as the TIP can activate it again
/// ({type, langid, clsid, profile, hkl}); None when it cannot be read.
pub(super) fn current_profile() -> Option<Value> {
    with_profiles(|manager| {
        let mut profile = TF_INPUTPROCESSORPROFILE::default();
        unsafe { manager.GetActiveProfile(&GUID_TFCAT_TIP_KEYBOARD, &mut profile) }.ok()?;
        Some(json!({
            "type": profile.dwProfileType,
            "langid": profile.langid,
            "clsid": guid_text(&profile.clsid),
            "profile": guid_text(&profile.guidProfile),
            "hkl": profile.hkl.0 as i64,
            "active": profile.dwFlags & TF_IPP_FLAG_ACTIVE != 0,
        }))
    })
    .flatten()
}

pub(super) fn is_visualtex(profile: &Value) -> bool {
    profile.get("type").and_then(Value::as_u64) == Some(TF_PROFILETYPE_INPUTPROCESSOR as u64)
        && profile.get("clsid").and_then(Value::as_str).map(|clsid| clsid.eq_ignore_ascii_case(&guid_text(&CLSID_TEXT_SERVICE))).unwrap_or(false)
}

/// Switches the user's input method to VisualTeX: the foreground field's thread activates the
/// TIP, which reports "activate" with the caret.
pub(super) fn switch_in() -> Result<(), String> {
    with_profiles(|manager| unsafe {
        manager
            .ActivateProfile(
                TF_PROFILETYPE_INPUTPROCESSOR,
                LANGID_ZH_CN,
                &CLSID_TEXT_SERVICE,
                &GUID_PROFILE,
                HKL::default(),
                TF_IPPMF_FORSESSION | TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE,
            )
            .map_err(|error| error.to_string())
    })
    .unwrap_or_else(|| Err("The Windows input method settings are unavailable".to_string()))
}

/// Switches back to `profile` (the field did not take VisualTeX).
pub(super) fn switch_to(profile: &Value) {
    let profile = profile.clone();
    let _ = with_profiles(move |manager| {
        let kind = profile.get("type").and_then(Value::as_u64).unwrap_or(0) as u32;
        let langid = profile.get("langid").and_then(Value::as_u64).unwrap_or(0) as u16;
        let parse = |key: &str| {
            profile
                .get(key)
                .and_then(Value::as_str)
                .map(|text| text.trim_matches(|ch| ch == '{' || ch == '}').replace('-', ""))
                .and_then(|hex| u128::from_str_radix(&hex, 16).ok())
                .map(GUID::from_u128)
                .unwrap_or_default()
        };
        let hkl = HKL(profile.get("hkl").and_then(Value::as_i64).unwrap_or(0) as isize as *mut _);
        let flags = TF_IPPMF_FORSESSION | TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE;
        unsafe {
            if kind == TF_PROFILETYPE_KEYBOARDLAYOUT {
                let _ = manager.ActivateProfile(TF_PROFILETYPE_KEYBOARDLAYOUT, langid, &GUID::zeroed(), &GUID::zeroed(), hkl, flags);
            } else {
                let _ = manager.ActivateProfile(TF_PROFILETYPE_INPUTPROCESSOR, langid, &parse("clsid"), &parse("profile"), HKL::default(), flags);
            }
        }
    });
}

// MARK: VisualTeX Input

/// Where the input method is installed: the TIP's folder, from its COM registration.
pub(super) fn install_directory() -> Option<PathBuf> {
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
    PathBuf::from(String::from_utf16_lossy(&buffer[..length])).parent().map(PathBuf::from)
}

/// Starts VisualTeX Input when it is installed and not running (it keeps itself to one per
/// session), so pinyin is ready before the first application asks for it.
pub(super) fn start_input_server() {
    let Some(directory) = install_directory() else { return };
    let program = directory.join("VisualTeXInput.exe");
    if !program.exists() {
        return;
    }
    use std::os::windows::process::CommandExt;
    const DETACHED_PROCESS: u32 = 0x0000_0008;
    match std::process::Command::new(&program).creation_flags(DETACHED_PROCESS).spawn() {
        Ok(_) => super::log("VisualTeX Input started"),
        Err(error) => super::log(&format!("VisualTeX Input not started: {error}")),
    }
}
