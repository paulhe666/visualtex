//! The part of the input method that lives in a field's thread before (and after) the text service
//! is active there: one message-only window.
//!
//! VisualTeX's profile is never among the user's input methods and stays disabled, so nothing but
//! VisualTeX's shortcut reaches it (the input switcher lists enabled profiles only, and Windows
//! activates a disabled profile only from the thread it is activated in). At the shortcut the app
//! puts this DLL into the thread of the field that has the keyboard (a hook on that one thread, taken
//! off again at once), which makes the window; the app then posts WM_OPEN to it. Here, in the field's
//! own thread, the input method the field has is written down, VisualTeX's profile is enabled and
//! activated for this thread alone, and when the formula ends, however it ends, the field gets the
//! input method it had back and the profile is disabled again.
//!
//! The window also receives the text service's messages (the app's link, its timers).

use crate::{log, service, CLSID_TEXT_SERVICE, GUID_PROFILE, LANGID_ZH_CN};
use std::cell::{Cell, RefCell};
use windows::core::{w, GUID};
use windows::Win32::Foundation::{HMODULE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::System::LibraryLoader::{
    GetModuleHandleExW, GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_PIN,
};
use windows::Win32::UI::Input::KeyboardAndMouse::HKL;
use windows::Win32::UI::TextServices::{
    CLSID_TF_InputProcessorProfiles, ITfInputProcessorProfileMgr, GUID_TFCAT_TIP_KEYBOARD, TF_INPUTPROCESSORPROFILE,
    TF_IPPMF_DISABLEPROFILE, TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE, TF_IPPMF_ENABLEPROFILE,
    TF_PROFILETYPE_INPUTPROCESSOR, TF_PROFILETYPE_KEYBOARDLAYOUT,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, KillTimer, PostMessageW, RegisterClassW, SetTimer, HWND_MESSAGE, WINDOW_EX_STYLE,
    WINDOW_STYLE, WM_APP, WM_TIMER, WNDCLASSW,
};

/// The app opens a formula in this thread's field (wparam: the formula's generation).
pub(crate) const WM_OPEN: u32 = WM_APP + 0x60;
/// The formula ended: the field gets its input method back.
const WM_HAND_BACK: u32 = WM_APP + 0x61;

/// The text service did not come up after the profile was activated (a thread without TSF).
const ACTIVATE_TIMER: usize = 4;
const ACTIVATE_WAIT_MS: u32 = 1500;

/// The input method a field had: what ITfInputProcessorProfileMgr::ActivateProfile takes.
#[derive(Clone, Debug)]
struct Profile {
    kind: u32,
    langid: u16,
    clsid: GUID,
    profile: GUID,
    hkl: isize,
}

impl Profile {
    fn is_visualtex(&self) -> bool {
        self.kind == TF_PROFILETYPE_INPUTPROCESSOR && self.clsid == CLSID_TEXT_SERVICE
    }
}

thread_local! {
    static WINDOW: Cell<isize> = const { Cell::new(0) };
    /// The formula the app asked for, until the text service takes it.
    static PENDING: Cell<u64> = const { Cell::new(0) };
    /// The input method the field had before the formula.
    static PREVIOUS: RefCell<Option<Profile>> = const { RefCell::new(None) };
    /// VisualTeX's profile is activated in this thread (by WM_OPEN) and not yet handed back.
    static HELD: Cell<bool> = const { Cell::new(false) };
}

/// This thread's window (0: none yet).
pub(crate) fn window() -> HWND {
    HWND(WINDOW.with(Cell::get) as *mut _)
}

/// The formula the app asked for in this thread, taken by the text service when it comes up.
pub(crate) fn take_pending() -> u64 {
    PENDING.with(|pending| pending.replace(0))
}

fn profiles() -> Option<ITfInputProcessorProfileMgr> {
    unsafe { CoCreateInstance(&CLSID_TF_InputProcessorProfiles, None, CLSCTX_INPROC_SERVER) }.ok()
}

/// The keyboard input method this thread types with now.
fn active_profile(manager: &ITfInputProcessorProfileMgr) -> Option<Profile> {
    let mut profile = TF_INPUTPROCESSORPROFILE::default();
    unsafe { manager.GetActiveProfile(&GUID_TFCAT_TIP_KEYBOARD, &mut profile) }.ok()?;
    Some(Profile {
        kind: profile.dwProfileType,
        langid: profile.langid,
        clsid: profile.clsid,
        profile: profile.guidProfile,
        hkl: profile.hkl.0 as isize,
    })
}

/// Activates `profile` for this thread only.
fn activate(manager: &ITfInputProcessorProfileMgr, profile: &Profile) -> windows::core::Result<()> {
    unsafe {
        if profile.kind == TF_PROFILETYPE_KEYBOARDLAYOUT {
            manager.ActivateProfile(
                TF_PROFILETYPE_KEYBOARDLAYOUT,
                profile.langid,
                &GUID::zeroed(),
                &GUID::zeroed(),
                HKL(profile.hkl as *mut _),
                TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE,
            )
        } else {
            manager.ActivateProfile(
                TF_PROFILETYPE_INPUTPROCESSOR,
                profile.langid,
                &profile.clsid,
                &profile.profile,
                HKL::default(),
                TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE,
            )
        }
    }
}

/// VisualTeX's profile is disabled again (off every input switcher).
fn disable(manager: &ITfInputProcessorProfileMgr) {
    let result = unsafe {
        manager.DeactivateProfile(
            TF_PROFILETYPE_INPUTPROCESSOR,
            LANGID_ZH_CN,
            &CLSID_TEXT_SERVICE,
            &GUID_PROFILE,
            HKL::default(),
            TF_IPPMF_DISABLEPROFILE,
        )
    };
    if let Err(error) = result {
        log::write(&format!("profile not disabled: {error}"));
    }
}

/// The app's WM_OPEN: a formula opens in this thread's field.
fn open(generation: u64) {
    if generation == 0 {
        return;
    }
    // The text service is up already (a formula just closed, its hand-back still on its way; or the
    // shortcut again): it takes the new formula as it is.
    if service::begin_in_this_thread(generation) {
        log::write(&format!("open gen={generation}: the text service is active"));
        return;
    }
    let Some(manager) = profiles() else {
        log::write(&format!("open gen={generation}: no TSF profiles"));
        return;
    };
    let previous = active_profile(&manager).filter(|profile| !profile.is_visualtex());
    PREVIOUS.with(|slot| {
        // (a formula opened over one whose hand-back has not run keeps the first one's input method)
        if previous.is_some() || !HELD.with(Cell::get) {
            *slot.borrow_mut() = previous.clone();
        }
    });
    PENDING.with(|pending| pending.set(generation));
    HELD.with(|held| held.set(true));
    let result = unsafe {
        manager.ActivateProfile(
            TF_PROFILETYPE_INPUTPROCESSOR,
            LANGID_ZH_CN,
            &CLSID_TEXT_SERVICE,
            &GUID_PROFILE,
            HKL::default(),
            TF_IPPMF_ENABLEPROFILE | TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE,
        )
    };
    log::write(&format!(
        "open gen={generation}: activated {:?} from {}",
        result.as_ref().map(|_| "ok").map_err(|error| error.code()),
        previous.as_ref().map(|profile| format!("{profile:?}")).unwrap_or_else(|| "-".to_string())
    ));
    if result.is_err() {
        hand_back_now();
        return;
    }
    // a thread without TSF never starts the text service: the profile does not stay enabled
    unsafe {
        SetTimer(Some(window()), ACTIVATE_TIMER, ACTIVATE_WAIT_MS, None);
    }
}

/// The formula ended (or never began): the field's input method comes back, after the edit that
/// ended it (posted: the text service may be inside one of TSF's calls).
pub(crate) fn hand_back() {
    let window = window();
    if window.0.is_null() {
        hand_back_now();
        return;
    }
    unsafe {
        let _ = PostMessageW(Some(window), WM_HAND_BACK, WPARAM(0), LPARAM(0));
    }
}

fn hand_back_now() {
    PENDING.with(|pending| pending.set(0));
    unsafe {
        let _ = KillTimer(Some(window()), ACTIVATE_TIMER);
    }
    let held = HELD.with(|held| held.replace(false));
    let Some(manager) = profiles() else { return };
    let previous = PREVIOUS.with(|slot| slot.borrow_mut().take());
    // Only a thread still typing with VisualTeX is given its input method back: the user may have
    // chosen another one meanwhile, which stays.
    let still_ours = active_profile(&manager).map(|profile| profile.is_visualtex()).unwrap_or(false);
    // (a thread that got VisualTeX without the shortcut, from the switcher while a formula was open
    // elsewhere, has no input method written down: disabling the profile takes it off)
    if !held && !still_ours {
        return;
    }
    if let (true, Some(profile)) = (still_ours, previous.as_ref()) {
        let result = activate(&manager, profile);
        log::write(&format!("handed back to {profile:?}: {:?}", result.map_err(|error| error.code())));
    } else {
        log::write(&format!("hand back: still ours={still_ours} previous={}", previous.is_some()));
    }
    // (disabling the profile also takes it off a thread that still has it: Windows then gives the
    // thread the language's default input method)
    disable(&manager);
}

extern "system" fn window_proc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match message {
        WM_OPEN => {
            open(wparam.0 as u64);
            LRESULT(0)
        }
        WM_HAND_BACK => {
            hand_back_now();
            LRESULT(0)
        }
        WM_TIMER if wparam.0 == ACTIVATE_TIMER => {
            unsafe {
                let _ = KillTimer(Some(window), ACTIVATE_TIMER);
            }
            if PENDING.with(Cell::get) != 0 {
                log::write("the text service did not start in this thread");
                hand_back_now();
            }
            LRESULT(0)
        }
        _ => {
            if service::window_message(message, wparam, lparam) {
                return LRESULT(0);
            }
            unsafe { DefWindowProcW(window, message, wparam, lparam) }
        }
    }
}

/// The class of the window, with the protocol's number: the app looks for it by name in the field's
/// thread, and one made by another build's DLL (still loaded in a long-running application) is not
/// this one.
pub(crate) const WINDOW_CLASS: windows::core::PCWSTR = w!("VisualTeXFormulaHost.2");

/// Makes this thread's window (once). The DLL is pinned: the window's procedure lives in it, and the
/// hook that brought it here is taken off again.
pub(crate) fn ensure() {
    if WINDOW.with(Cell::get) != 0 {
        return;
    }
    use std::sync::Once;
    static REGISTER: Once = Once::new();
    REGISTER.call_once(|| {
        let mut module = HMODULE::default();
        unsafe {
            let _ = GetModuleHandleExW(
                GET_MODULE_HANDLE_EX_FLAG_PIN | GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
                windows::core::PCWSTR(window_proc as *const u16),
                &mut module,
            );
        }
        let class = WNDCLASSW {
            lpfnWndProc: Some(window_proc),
            hInstance: crate::dll_instance(),
            lpszClassName: WINDOW_CLASS,
            ..Default::default()
        };
        unsafe { RegisterClassW(&class) };
    });
    let created = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE(0),
            WINDOW_CLASS,
            w!(""),
            WINDOW_STYLE(0),
            0,
            0,
            0,
            0,
            Some(HWND_MESSAGE),
            None,
            Some(crate::dll_instance()),
            None,
        )
    };
    if let Ok(created) = created {
        WINDOW.with(|window| window.set(created.0 as isize));
        log::write("host window made");
    }
}
