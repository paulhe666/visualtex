#![cfg(target_os = "windows")]

//! MathType's WPS startup template.
//!
//! Each start of MathType in the user's session (its editor, or an OLE edit of a MathType
//! equation from Word, PowerPoint or WPS) copies "MathType Commands 2016.dotm" from MathType's
//! "Office Support\32" folder into WPS's startup folder when no file of that name is there. In
//! 64-bit WPS that template cannot run: its 32-bit Declare statements do not compile (PtrSafe),
//! its VBA Extensibility reference may point to a removed WPS folder, and MathType's 64-bit
//! template fails there too (run-time error 75). So every WPS start showed VBA error boxes,
//! which also block COM calls until they are dismissed.
//!
//! VisualTeX keeps a macro-free placeholder under that name instead: WPS loads it silently and
//! MathType leaves an existing file alone. MathType's copy is kept in VisualTeX's office folder.
//! MathType equations in WPS keep working through VisualTeX's add-in. Deleting the placeholder
//! lets MathType install its template again (and VisualTeX replaces it again).

use std::fs;
use std::path::{Path, PathBuf};
use std::time::Duration;
use windows::core::HSTRING;
use windows::Win32::System::Registry::{RegGetValueW, HKEY, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, RRF_RT_REG_SZ};

const TEMPLATE: &str = "MathType Commands 2016.dotm";
const PLACEHOLDER: &[u8] = include_bytes!("wps_mathtype_placeholder.dotm");
/// MathType can put its template back whenever it starts; a placeholder left in place keeps it
/// out, so a slow check is enough.
const CHECK_INTERVAL: Duration = Duration::from_secs(10 * 60);
/// WPS keeps its startup templates open while it runs.
const BUSY_RETRY: Duration = Duration::from_secs(60);
const ERROR_SHARING_VIOLATION: i32 = 32;
const ERROR_LOCK_VIOLATION: i32 = 33;

#[derive(Debug, PartialEq, Eq)]
enum Outcome {
    /// No 64-bit WPS, or no MathType.
    NotNeeded,
    /// The placeholder is in place.
    InPlace,
    /// The placeholder was written (MathType's template, if any, moved aside).
    Written { moved: bool },
    /// WPS has the template open: try again later.
    Busy,
    Failed(String),
}

/// Checks now and then, from a background thread, for the rest of the session.
pub fn start() {
    let _ = std::thread::Builder::new()
        .name("visualtex-wps-mathtype-template".to_string())
        .spawn(|| {
            let mut last = None;
            loop {
                let outcome = check();
                let wait = if outcome == Outcome::Busy { BUSY_RETRY } else { CHECK_INTERVAL };
                let report = match &outcome {
                    Outcome::Written { moved } => Some(if *moved {
                        "WPS startup: MathType's VBA template (cannot run in 64-bit WPS) moved aside; placeholder written".to_string()
                    } else {
                        "WPS startup: placeholder for MathType's VBA template written".to_string()
                    }),
                    Outcome::Busy if last != Some(Outcome::Busy) => {
                        Some("WPS startup: MathType's VBA template is open in WPS; replacing it once WPS closes".to_string())
                    }
                    Outcome::Failed(error) => Some(format!("WPS startup: MathType template check failed: {error}")),
                    _ => None,
                };
                if let Some(message) = report {
                    crate::app_lifecycle::append_lifecycle_log(message);
                }
                last = Some(outcome);
                std::thread::sleep(wait);
            }
        });
}

fn check() -> Outcome {
    if !wps_is_64_bit() || !mathtype_installed() {
        return Outcome::NotNeeded;
    }
    let Some(startup) = std::env::var_os("APPDATA").map(|root| PathBuf::from(root).join("kingsoft").join("wps").join("startup")) else {
        return Outcome::NotNeeded;
    };
    place(&startup, &backup_directory())
}

/// Puts the placeholder at `startup\TEMPLATE`, moving another file of that name into `backup`.
fn place(startup: &Path, backup: &Path) -> Outcome {
    let target = startup.join(TEMPLATE);
    let mut moved = false;
    match fs::read(&target) {
        Ok(existing) if existing == PLACEHOLDER => return Outcome::InPlace,
        Ok(_) => {
            if let Err(error) = move_aside(&target, backup) {
                return if is_busy(&error) { Outcome::Busy } else { Outcome::Failed(format!("moving {}: {error}", target.display())) };
            }
            moved = true;
        }
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
        Err(error) if is_busy(&error) => return Outcome::Busy,
        Err(error) => return Outcome::Failed(format!("reading {}: {error}", target.display())),
    }
    if let Err(error) = fs::create_dir_all(startup).and_then(|_| fs::write(&target, PLACEHOLDER)) {
        return Outcome::Failed(format!("writing {}: {error}", target.display()));
    }
    Outcome::Written { moved }
}

fn move_aside(target: &Path, backup: &Path) -> std::io::Result<()> {
    fs::create_dir_all(backup)?;
    let kept = backup.join(TEMPLATE);
    if kept.exists() {
        clear_read_only(&kept);
        fs::remove_file(&kept)?;
    }
    // Renaming does not need write access to the file (MathType's copy is read-only), and WPS
    // holding it open makes it fail with a sharing violation.
    match fs::rename(target, &kept) {
        Ok(()) => Ok(()),
        Err(error) if is_busy(&error) => Err(error),
        Err(_) => {
            fs::copy(target, &kept)?;
            clear_read_only(target);
            fs::remove_file(target)
        }
    }
}

fn clear_read_only(path: &Path) {
    if let Ok(metadata) = fs::metadata(path) {
        let mut permissions = metadata.permissions();
        if permissions.readonly() {
            #[allow(clippy::permissions_set_readonly_false)]
            permissions.set_readonly(false);
            let _ = fs::set_permissions(path, permissions);
        }
    }
}

fn is_busy(error: &std::io::Error) -> bool {
    matches!(error.raw_os_error(), Some(ERROR_SHARING_VIOLATION | ERROR_LOCK_VIOLATION))
}

fn backup_directory() -> PathBuf {
    std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("office")
        .join("mathtype-wps")
}

fn mathtype_installed() -> bool {
    std::env::var_os("ProgramFiles(x86)")
        .map(|root| PathBuf::from(root).join("MathType").join("MathType.exe").is_file())
        .unwrap_or(false)
}

/// WPS Office's wps.exe (per-machine or per-user installation) is a 64-bit program.
fn wps_is_64_bit() -> bool {
    for root in [HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER] {
        if let Some(install) = registry_string(root, "SOFTWARE\\Kingsoft\\Office\\6.0\\Common", "InstallRoot") {
            if let Some(machine) = pe_machine(&Path::new(&install).join("office6").join("wps.exe")) {
                return machine == 0x8664;
            }
        }
    }
    false
}

fn registry_string(root: HKEY, key: &str, value: &str) -> Option<String> {
    let key = HSTRING::from(key);
    let value = HSTRING::from(value);
    let mut buffer = vec![0u16; 1024];
    let mut size = (buffer.len() * 2) as u32;
    let status = unsafe {
        RegGetValueW(root, &key, &value, RRF_RT_REG_SZ, None, Some(buffer.as_mut_ptr().cast()), Some(&mut size))
    };
    if status.is_err() {
        return None;
    }
    let length = buffer.iter().position(|unit| *unit == 0).unwrap_or(buffer.len());
    Some(String::from_utf16_lossy(&buffer[..length]))
}

/// The machine type in a PE file's header.
pub(crate) fn pe_machine(path: &Path) -> Option<u16> {
    use std::io::{Read, Seek, SeekFrom};
    let mut file = fs::File::open(path).ok()?;
    let mut dos = [0u8; 64];
    file.read_exact(&mut dos).ok()?;
    if &dos[..2] != b"MZ" {
        return None;
    }
    let offset = u32::from_le_bytes([dos[60], dos[61], dos[62], dos[63]]) as u64;
    file.seek(SeekFrom::Start(offset)).ok()?;
    let mut header = [0u8; 6];
    file.read_exact(&mut header).ok()?;
    if &header[..4] != b"PE\0\0" {
        return None;
    }
    Some(u16::from_le_bytes([header[4], header[5]]))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn scratch(name: &str) -> PathBuf {
        let directory = std::env::temp_dir().join(format!("visualtex-wps-mathtype-{name}-{}", std::process::id()));
        let _ = fs::remove_dir_all(&directory);
        directory
    }

    #[test]
    fn placeholder_is_a_macro_free_template() {
        let text = String::from_utf8_lossy(PLACEHOLDER);
        assert!(PLACEHOLDER.starts_with(b"PK"));
        assert!(text.contains("[Content_Types].xml"));
        assert!(!text.contains("vbaProject"));
    }

    #[test]
    fn writes_the_placeholder_and_keeps_mathtypes_template() {
        let root = scratch("place");
        let startup = root.join("startup");
        let backup = root.join("backup");
        fs::create_dir_all(&startup).unwrap();
        fs::write(startup.join(TEMPLATE), b"mathtype template").unwrap();
        let mut permissions = fs::metadata(startup.join(TEMPLATE)).unwrap().permissions();
        permissions.set_readonly(true);
        fs::set_permissions(startup.join(TEMPLATE), permissions).unwrap();

        assert_eq!(place(&startup, &backup), Outcome::Written { moved: true });
        assert_eq!(fs::read(startup.join(TEMPLATE)).unwrap(), PLACEHOLDER);
        assert_eq!(fs::read(backup.join(TEMPLATE)).unwrap(), b"mathtype template");
        assert_eq!(place(&startup, &backup), Outcome::InPlace);

        // A newer MathType copy replaces the kept one.
        fs::write(startup.join(TEMPLATE), b"second copy").unwrap();
        assert_eq!(place(&startup, &backup), Outcome::Written { moved: true });
        assert_eq!(fs::read(backup.join(TEMPLATE)).unwrap(), b"second copy");
        clear_read_only(&backup.join(TEMPLATE));
        let _ = fs::remove_dir_all(&root);
    }

    #[test]
    fn writes_the_placeholder_before_mathtype_does() {
        let root = scratch("missing");
        let startup = root.join("startup");
        assert_eq!(place(&startup, &root.join("backup")), Outcome::Written { moved: false });
        assert_eq!(fs::read(startup.join(TEMPLATE)).unwrap(), PLACEHOLDER);
        let _ = fs::remove_dir_all(&root);
    }

    #[test]
    fn a_template_open_in_wps_is_busy() {
        use std::os::windows::fs::OpenOptionsExt;
        let root = scratch("busy");
        let startup = root.join("startup");
        fs::create_dir_all(&startup).unwrap();
        fs::write(startup.join(TEMPLATE), b"mathtype template").unwrap();
        // WPS opens its startup templates without sharing.
        let held = fs::OpenOptions::new().read(true).share_mode(0).open(startup.join(TEMPLATE)).unwrap();
        assert_eq!(place(&startup, &root.join("backup")), Outcome::Busy);
        drop(held);
        assert_eq!(place(&startup, &root.join("backup")), Outcome::Written { moved: true });
        let _ = fs::remove_dir_all(&root);
    }
}
