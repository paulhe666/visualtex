//! Handing the field back to the input method it had before VisualTeX's shortcut switched it to
//! VisualTeX (the app sends that profile as `returnTo`). The switch is made here, on the field's
//! own thread, as the macOS input method has the field's client select it.

use serde_json::Value;
use windows::core::GUID;
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::UI::Input::KeyboardAndMouse::HKL;
use windows::Win32::UI::TextServices::{
    CLSID_TF_InputProcessorProfiles, ITfInputProcessorProfileMgr, TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE,
    TF_PROFILETYPE_INPUTPROCESSOR, TF_PROFILETYPE_KEYBOARDLAYOUT,
};

fn guid(value: Option<&Value>) -> GUID {
    value
        .and_then(Value::as_str)
        .map(|text| text.trim_matches(|ch| ch == '{' || ch == '}'))
        .and_then(|text| u128::from_str_radix(&text.replace('-', ""), 16).ok())
        .map(GUID::from_u128)
        .unwrap_or_default()
}

/// Activates the profile described by `profile` ({type, langid, clsid, profile, hkl}).
pub(crate) fn activate(profile: &Value) {
    let kind = profile.get("type").and_then(Value::as_u64).unwrap_or(0) as u32;
    let langid = profile.get("langid").and_then(Value::as_u64).unwrap_or(0) as u16;
    let hkl = profile.get("hkl").and_then(Value::as_i64).unwrap_or(0) as isize;
    if langid == 0 {
        return;
    }
    let manager: windows::core::Result<ITfInputProcessorProfileMgr> =
        unsafe { CoCreateInstance(&CLSID_TF_InputProcessorProfiles, None, CLSCTX_INPROC_SERVER) };
    let Ok(manager) = manager else { return };
    let result = unsafe {
        if kind == TF_PROFILETYPE_KEYBOARDLAYOUT {
            manager.ActivateProfile(
                TF_PROFILETYPE_KEYBOARDLAYOUT,
                langid,
                &GUID::zeroed(),
                &GUID::zeroed(),
                HKL(hkl as *mut _),
                TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE,
            )
        } else {
            manager.ActivateProfile(
                TF_PROFILETYPE_INPUTPROCESSOR,
                langid,
                &guid(profile.get("clsid")),
                &guid(profile.get("profile")),
                HKL::default(),
                TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE,
            )
        }
    };
    crate::log::write(&format!("hand back to {profile} -> {result:?}"));
}
