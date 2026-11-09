//! VisualTeX's Windows input method: a TSF text input processor (TIP), as the macOS input method
//! (apps/macos/input-method) is there: pinyin and English through RIME (in VisualTeX Input, the
//! server beside this DLL), and formulas typed into any application with the VisualTeX app.
//!
//! The TIP lives in the host application's process. While a formula is open (VisualTeX's
//! shortcut switched the field to this TIP and opened it), the field holds one selected hair
//! space as its composition, the TIP forwards every key to VisualTeX over a named pipe, VisualTeX
//! draws the formula over the field's line, and on Return the TIP replaces the composition with
//! the formula's text. Messages are the macOS input method's (newline-delimited JSON; see
//! apps/windows/src-tauri/src/anywhere).

mod display;
mod edit;
mod engine;
mod factory;
mod keys;
mod link;
mod log;
mod profile;
mod register;
mod rimekeys;
mod service;
mod space;

use std::ffi::c_void;
use std::sync::atomic::{AtomicIsize, Ordering};
use windows::core::{BOOL, GUID, HRESULT};
use windows::Win32::Foundation::{CLASS_E_CLASSNOTAVAILABLE, E_POINTER, HINSTANCE, S_FALSE, S_OK, TRUE};
use windows::Win32::System::SystemServices::DLL_PROCESS_ATTACH;

/// The text service's CLSID.
pub(crate) const CLSID_TEXT_SERVICE: GUID = GUID::from_u128(0xb6e0a6c4_6f52_4c7e_9b7a_2f1e5a3d9c41);
/// Its language profile (Simplified Chinese: it types pinyin too).
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
