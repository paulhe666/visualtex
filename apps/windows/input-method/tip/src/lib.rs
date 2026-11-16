//! VisualTeX's Windows input method: a TSF text input processor (TIP) that types formulas, and
//! nothing else. It is not one of the user's input methods: its profile is registered but disabled
//! and hidden, so no input switcher offers it and the user's own input methods and the switching
//! between them are as they were without it.
//!
//! VisualTeX's shortcut is the only way in. The app puts this DLL into the thread of the field
//! that has the keyboard (VisualTeXHook, on that one thread) and asks its window there for a
//! formula (host.rs): the field's input method is written down, this TIP is activated for that
//! thread alone, the field holds the formula's space as its composition, the TIP forwards every key
//! to VisualTeX over a named pipe, VisualTeX draws the formula over the field's line, and on Return
//! the TIP replaces the composition with the formula's text (service.rs). Then, as after every
//! other end of a formula, the thread gets the input method it had back and the profile is disabled
//! again. Messages are newline-delimited JSON (see apps/windows/src-tauri/src/anywhere).

mod display;
mod edit;
mod factory;
mod field;
mod host;
mod keys;
mod link;
mod log;
mod register;
mod service;
mod space;

use std::ffi::c_void;
use std::sync::atomic::{AtomicIsize, Ordering};
use windows::core::{BOOL, GUID, HRESULT};
use windows::Win32::Foundation::{CLASS_E_CLASSNOTAVAILABLE, E_POINTER, HINSTANCE, LPARAM, LRESULT, S_FALSE, S_OK, TRUE, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::CallNextHookEx;
use windows::Win32::System::SystemServices::DLL_PROCESS_ATTACH;

/// The text service's CLSID.
pub(crate) const CLSID_TEXT_SERVICE: GUID = GUID::from_u128(0xb6e0a6c4_6f52_4c7e_9b7a_2f1e5a3d9c41);
/// Its language profile (registered under Simplified Chinese; activated whatever the field's language).
pub(crate) const GUID_PROFILE: GUID = GUID::from_u128(0xd3a1f7b2_8c4e_4e2a_a5d6_6b9e0c1f2a73);
pub(crate) const LANGID_ZH_CN: u16 = 0x0804;
pub(crate) const DESCRIPTION: &str = "VisualTeX";

static INSTANCE: AtomicIsize = AtomicIsize::new(0);

pub(crate) fn dll_instance() -> HINSTANCE {
    HINSTANCE(INSTANCE.load(Ordering::Relaxed) as *mut c_void)
}

#[no_mangle]
extern "system" fn DllMain(instance: HINSTANCE, reason: u32, _reserved: *mut c_void) -> BOOL {
    if reason == DLL_PROCESS_ATTACH {
        INSTANCE.store(instance.0 as isize, Ordering::Relaxed);
    }
    TRUE
}

#[no_mangle]
extern "system" fn DllGetClassObject(class: *const GUID, interface: *const GUID, object: *mut *mut c_void) -> HRESULT {
    if class.is_null() || interface.is_null() || object.is_null() {
        return E_POINTER;
    }
    unsafe { *object = std::ptr::null_mut() };
    if unsafe { *class } != CLSID_TEXT_SERVICE {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
    factory::class_object(interface, object)
}

/// The hook the app sets on the thread of the field that has the keyboard (WH_GETMESSAGE), for as
/// long as it takes this thread to run it once: it makes the thread's window (host.rs), which the
/// app then talks to.
#[no_mangle]
extern "system" fn VisualTeXHook(code: i32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    if code >= 0 {
        host::ensure();
    }
    unsafe { CallNextHookEx(None, code, wparam, lparam) }
}

/// The DLL stays loaded: TSF keeps its objects for the thread's life and an
/// unload under a live pipe thread would be fatal to the host.
#[no_mangle]
extern "system" fn DllCanUnloadNow() -> HRESULT {
    S_FALSE
}

#[no_mangle]
extern "system" fn DllRegisterServer() -> HRESULT {
    match register::register() {
        Ok(()) => S_OK,
        Err(error) => error.code(),
    }
}

#[no_mangle]
extern "system" fn DllUnregisterServer() -> HRESULT {
    match register::unregister() {
        Ok(()) => S_OK,
        Err(error) => error.code(),
    }
}
