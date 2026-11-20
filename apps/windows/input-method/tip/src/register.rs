//! Registration (regsvr32, elevated: the COM class and the TSF profile live in HKLM): the
//! in-process COM server, the keyboard profile (not enabled, and hidden in the settings: it is no
//! input method of the user's, host.rs enables it for the time of a formula) and the categories
//! that let TSF load it in every kind of application.

use crate::{dll_instance, CLSID_TEXT_SERVICE, DESCRIPTION, GUID_PROFILE, LANGID_ZH_CN};
use windows::core::{Result, GUID, HSTRING, PCWSTR};
use windows::Win32::Foundation::{ERROR_SUCCESS, MAX_PATH};
use windows::Win32::UI::Input::KeyboardAndMouse::HKL;
use windows::Win32::System::Com::{CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED};
use windows::Win32::System::LibraryLoader::GetModuleFileNameW;
use windows::Win32::System::Registry::{
    RegCloseKey, RegCreateKeyExW, RegDeleteTreeW, RegSetValueExW, HKEY, HKEY_LOCAL_MACHINE, KEY_WRITE,
    REG_OPTION_NON_VOLATILE, REG_SZ,
};
use windows::Win32::UI::TextServices::{
    CLSID_TF_CategoryMgr, CLSID_TF_InputProcessorProfiles, ITfCategoryMgr, ITfInputProcessorProfileMgr,
    GUID_TFCAT_TIPCAP_IMMERSIVESUPPORT, GUID_TFCAT_TIPCAP_SYSTRAYSUPPORT, GUID_TFCAT_TIPCAP_UIELEMENTENABLED,
    GUID_TFCAT_DISPLAYATTRIBUTEPROVIDER, GUID_TFCAT_TIP_KEYBOARD, TF_RP_HIDDENINSETTINGUI,
};

const CATEGORIES: [GUID; 5] = [
    GUID_TFCAT_TIP_KEYBOARD,
    GUID_TFCAT_DISPLAYATTRIBUTEPROVIDER,
    GUID_TFCAT_TIPCAP_IMMERSIVESUPPORT,
    GUID_TFCAT_TIPCAP_SYSTRAYSUPPORT,
    GUID_TFCAT_TIPCAP_UIELEMENTENABLED,
];

fn clsid_key() -> String {
    format!("SOFTWARE\\Classes\\CLSID\\{{{:?}}}", CLSID_TEXT_SERVICE)
}

fn module_path() -> Vec<u16> {
    let mut buffer = vec![0u16; MAX_PATH as usize * 4];
    let length = unsafe { GetModuleFileNameW(Some(dll_instance().into()), &mut buffer) } as usize;
    buffer.truncate(length);
    buffer
}

fn set_string(key: HKEY, name: Option<&str>, value: &[u16]) -> Result<()> {
    let mut data: Vec<u16> = value.to_vec();
    data.push(0);
    let bytes = unsafe { std::slice::from_raw_parts(data.as_ptr().cast::<u8>(), data.len() * 2) };
    let name = name.map(HSTRING::from);
    let status = unsafe {
        RegSetValueExW(
            key,
            name.as_ref().map(|name| PCWSTR(name.as_ptr())).unwrap_or(PCWSTR::null()),
            None,
            REG_SZ,
            Some(bytes),
        )
    };
    if status != ERROR_SUCCESS {
        return Err(status.to_hresult().into());
    }
    Ok(())
}

fn create_key(path: &str) -> Result<HKEY> {
    let mut key = HKEY::default();
    let status = unsafe {
        RegCreateKeyExW(
            HKEY_LOCAL_MACHINE,
            &HSTRING::from(path),
            None,
            None,
            REG_OPTION_NON_VOLATILE,
            KEY_WRITE,
            None,
            &mut key,
            None,
        )
    };
    if status != ERROR_SUCCESS {
        return Err(status.to_hresult().into());
    }
    Ok(key)
}

fn register_com_server() -> Result<()> {
    let class = create_key(&clsid_key())?;
    let description: Vec<u16> = DESCRIPTION.encode_utf16().collect();
    let result = set_string(class, None, &description);
    unsafe {
        let _ = RegCloseKey(class);
    }
    result?;
    let server = create_key(&format!("{}\\InprocServer32", clsid_key()))?;
    let result = set_string(server, None, &module_path()).and_then(|()| {
        let model: Vec<u16> = "Apartment".encode_utf16().collect();
        set_string(server, Some("ThreadingModel"), &model)
    });
    unsafe {
        let _ = RegCloseKey(server);
    }
    result
}

pub(crate) fn register() -> Result<()> {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
    register_com_server()?;
    let profiles: ITfInputProcessorProfileMgr =
        unsafe { CoCreateInstance(&CLSID_TF_InputProcessorProfiles, None, CLSCTX_INPROC_SERVER)? };
    let description: Vec<u16> = DESCRIPTION.encode_utf16().collect();
    let icon = module_path();
    unsafe {
        profiles.RegisterProfile(
            &CLSID_TEXT_SERVICE,
            LANGID_ZH_CN,
            &GUID_PROFILE,
            &description,
            &icon,
            0,
            HKL::default(),
            0,
            false,
            TF_RP_HIDDENINSETTINGUI,
        )?;
    }
    let categories: ITfCategoryMgr = unsafe { CoCreateInstance(&CLSID_TF_CategoryMgr, None, CLSCTX_INPROC_SERVER)? };
    for category in CATEGORIES {
        unsafe { categories.RegisterCategory(&CLSID_TEXT_SERVICE, &category, &CLSID_TEXT_SERVICE)? };
    }
    crate::log::write("registered");
    Ok(())
}

pub(crate) fn unregister() -> Result<()> {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
    if let Ok(categories) =
        unsafe { CoCreateInstance::<_, ITfCategoryMgr>(&CLSID_TF_CategoryMgr, None, CLSCTX_INPROC_SERVER) }
    {
        for category in CATEGORIES {
            unsafe {
                let _ = categories.UnregisterCategory(&CLSID_TEXT_SERVICE, &category, &CLSID_TEXT_SERVICE);
            }
        }
    }
    if let Ok(profiles) = unsafe {
        CoCreateInstance::<_, ITfInputProcessorProfileMgr>(&CLSID_TF_InputProcessorProfiles, None, CLSCTX_INPROC_SERVER)
    } {
        unsafe {
            let _ = profiles.UnregisterProfile(&CLSID_TEXT_SERVICE, LANGID_ZH_CN, &GUID_PROFILE, 0);
        }
    }
    unsafe {
        let _ = RegDeleteTreeW(HKEY_LOCAL_MACHINE, &HSTRING::from(clsid_key()));
    }
    Ok(())
}
