//! Who may open the input method's pipe: this user, and the sandboxed applications this user
//! runs (AppContainers such as the Start menu's search, Store apps and browser sandboxes, which
//! have the TIP loaded too), at low integrity as well.

use windows::core::{Result, HSTRING, PWSTR};
use windows::Win32::Foundation::{CloseHandle, HANDLE, HLOCAL, LocalFree};
use windows::Win32::Security::Authorization::{
    ConvertSidToStringSidW, ConvertStringSecurityDescriptorToSecurityDescriptorW, SDDL_REVISION_1,
};
use windows::Win32::Security::{GetTokenInformation, TokenUser, PSECURITY_DESCRIPTOR, SECURITY_ATTRIBUTES, TOKEN_QUERY, TOKEN_USER};
use windows::Win32::System::Threading::{GetCurrentProcess, OpenProcessToken};

/// This process's user as an SDDL SID string (S-1-5-21-...).
pub fn user_sid() -> Option<String> {
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

/// Security attributes for the pipe (kept alive by the returned owner).
pub struct PipeSecurity {
    descriptor: PSECURITY_DESCRIPTOR,
    pub attributes: SECURITY_ATTRIBUTES,
}

unsafe impl Send for PipeSecurity {}

impl Drop for PipeSecurity {
    fn drop(&mut self) {
        unsafe {
            let _ = LocalFree(Some(HLOCAL(self.descriptor.0)));
        }
    }
}

pub fn pipe_security() -> Result<PipeSecurity> {
    let user = user_sid().unwrap_or_else(|| "OW".to_string());
    // Full access for SYSTEM and this user; read/write for all (and all restricted) application
    // packages; the low-integrity label lets sandboxed processes write.
    let sddl = format!("D:P(A;;GA;;;SY)(A;;GA;;;{user})(A;;GRGW;;;AC)(A;;GRGW;;;S-1-15-2-2)S:(ML;;NW;;;LW)");
    let mut descriptor = PSECURITY_DESCRIPTOR::default();
    unsafe {
        ConvertStringSecurityDescriptorToSecurityDescriptorW(&HSTRING::from(sddl), SDDL_REVISION_1, &mut descriptor, None)?;
    }
    Ok(PipeSecurity {
        descriptor,
        attributes: SECURITY_ATTRIBUTES {
            nLength: std::mem::size_of::<SECURITY_ATTRIBUTES>() as u32,
            lpSecurityDescriptor: descriptor.0,
            bInheritHandle: false.into(),
        },
    })
}
