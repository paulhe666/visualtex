use crate::office::background::OfficeBackgroundStatus;
use crate::office::manifest::{
    manifest_version, render_manifest, ManifestHost, LEGACY_POWERPOINT_MANIFEST_FILE,
    LEGACY_WORD_MANIFEST_FILE,
};
use crate::office::state::{OfficeCompanionStatus, OfficePaths, OFFICE_UI_VERSION};
use serde::Serialize;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Output};
use uuid::Uuid;

const CERTIFICATE_COMMON_NAME: &str = "VisualTeX Local Office Companion";

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeHostInstallStatus {
    pub application_installed: bool,
    pub manifest_installed: bool,
    pub manifest_version: Option<String>,
    pub manifest_path: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CertificateInstallStatus {
    pub certificate_exists: bool,
    pub private_key_exists: bool,
    pub trusted: bool,
    pub keychain_path: String,
    pub last_error: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeIntegrationStatus {
    pub word: OfficeHostInstallStatus,
    pub powerpoint: OfficeHostInstallStatus,
    pub expected_manifest_version: String,
    pub certificate: CertificateInstallStatus,
    pub background: OfficeBackgroundStatus,
    pub companion: OfficeCompanionStatus,
    pub office_ui_version: String,
}

fn user_home() -> Result<PathBuf, String> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .filter(|path| path.is_absolute())
        .ok_or_else(|| "Unable to resolve the current user's home directory".to_string())
}

fn application_paths(home: &Path, host: ManifestHost) -> [PathBuf; 2] {
    let name = match host {
        ManifestHost::Word => "Microsoft Word.app",
        ManifestHost::PowerPoint => "Microsoft PowerPoint.app",
    };
    [
        PathBuf::from("/Applications").join(name),
        home.join("Applications").join(name),
    ]
}

fn application_installed(home: &Path, host: ManifestHost) -> bool {
    application_paths(home, host)
        .iter()
        .any(|path| path.is_dir())
}

pub(crate) fn manifest_directory(home: &Path, host: ManifestHost) -> PathBuf {
    match host {
        ManifestHost::Word => home.join("Library/Containers/com.microsoft.Word/Data/Documents/wef"),
        ManifestHost::PowerPoint => {
            home.join("Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef")
        }
    }
}

pub(crate) fn manifest_path(home: &Path, host: ManifestHost) -> PathBuf {
    manifest_directory(home, host).join(host.file_name())
}

fn legacy_manifest_path(home: &Path, host: ManifestHost) -> PathBuf {
    let file_name = match host {
        ManifestHost::Word => LEGACY_WORD_MANIFEST_FILE,
        ManifestHost::PowerPoint => LEGACY_POWERPOINT_MANIFEST_FILE,
    };
    manifest_directory(home, host).join(file_name)
}

fn login_keychain(home: &Path) -> PathBuf {
    let database = home.join("Library/Keychains/login.keychain-db");
    if database.exists() {
        database
    } else {
        home.join("Library/Keychains/login.keychain")
    }
}

fn command_error(program: &str, output: &Output) -> String {
    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    format!(
        "{program} exited with {}{}{}",
        output.status,
        if stderr.is_empty() { "" } else { ": " },
        if stderr.is_empty() { &stdout } else { &stderr }
    )
}

fn run_security(arguments: &[&str]) -> Result<Output, String> {
    Command::new("/usr/bin/security")
        .args(arguments)
        .output()
        .map_err(|error| format!("Unable to run macOS security tool: {error}"))
}

fn read_manifest_version(path: &Path) -> Option<String> {
    let xml = fs::read_to_string(path).ok()?;
    let start = xml.find("<Version>")? + "<Version>".len();
    let end = xml[start..].find("</Version>")? + start;
    let value = xml[start..end].trim();
    (!value.is_empty()).then(|| value.to_string())
}

#[cfg(unix)]
fn sync_directory(path: &Path) -> Result<(), String> {
    let directory = fs::File::open(path)
        .map_err(|error| format!("Unable to open {} for sync: {error}", path.display()))?;
    directory
        .sync_all()
        .map_err(|error| format!("Unable to sync {}: {error}", path.display()))
}

#[cfg(not(unix))]
fn sync_directory(_path: &Path) -> Result<(), String> {
    Ok(())
}

#[cfg(unix)]
fn set_mode(path: &Path, mode: u32) -> Result<(), String> {
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(path, fs::Permissions::from_mode(mode))
        .map_err(|error| format!("Unable to set permissions on {}: {error}", path.display()))
}

#[cfg(not(unix))]
fn set_mode(_path: &Path, _mode: u32) -> Result<(), String> {
    Ok(())
}

fn atomic_write(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let directory = path
        .parent()
        .ok_or_else(|| format!("Manifest path has no parent: {}", path.display()))?;
    fs::create_dir_all(directory).map_err(|error| {
        format!(
            "Unable to create Office manifest directory {}: {error}",
            directory.display()
        )
    })?;
    set_mode(directory, 0o700)?;
    let temporary = directory.join(format!(
        ".{}.{}.tmp",
        path.file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("VisualTeX.xml"),
        Uuid::new_v4()
    ));
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&temporary)
        .map_err(|error| format!("Unable to create {}: {error}", temporary.display()))?;
    file.write_all(bytes)
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    file.sync_all()
        .map_err(|error| format!("Unable to sync {}: {error}", temporary.display()))?;
    set_mode(&temporary, 0o644)?;
    fs::rename(&temporary, path).map_err(|error| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to install manifest {}: {error}", path.display())
    })?;
    set_mode(path, 0o644)?;
    sync_directory(directory)
}

pub(crate) fn install_manifest_at(home: &Path, host: ManifestHost) -> Result<PathBuf, String> {
    let manifest = render_manifest(host)?;
    let path = manifest_path(home, host);
    // Office's own macOS registrar prefixes sideloaded manifests with their
    // add-in GUID. Preserve the registered file when its bytes are unchanged:
    // replacing the inode on every repair needlessly invalidates Office's Wef
    // registration and command cache.
    let current = fs::read(&path).ok();
    if current.as_deref() != Some(manifest.as_bytes()) {
        atomic_write(&path, manifest.as_bytes())?;
    }

    // Migrate installations created before GUID-prefixed registration. Keep a
    // single VisualTeX manifest for each permanent <Id> in the host catalog.
    let legacy = legacy_manifest_path(home, host);
    if legacy != path && legacy.exists() {
        fs::remove_file(&legacy)
            .map_err(|error| format!("Unable to remove {}: {error}", legacy.display()))?;
        if let Some(directory) = legacy.parent() {
            sync_directory(directory)?;
        }
    }
    Ok(path)
}

pub(crate) fn uninstall_manifest_at(home: &Path, host: ManifestHost) -> Result<(), String> {
    for path in [manifest_path(home, host), legacy_manifest_path(home, host)] {
        if path.exists() {
            fs::remove_file(&path)
                .map_err(|error| format!("Unable to remove {}: {error}", path.display()))?;
            if let Some(directory) = path.parent() {
                sync_directory(directory)?;
            }
        }
    }
    Ok(())
}

pub fn install_available_manifests() -> Result<(), String> {
    let home = user_home()?;
    for host in [ManifestHost::Word, ManifestHost::PowerPoint] {
        if application_installed(&home, host) {
            install_manifest_at(&home, host)?;
        }
    }
    Ok(())
}

pub fn uninstall_manifests() -> Result<(), String> {
    let home = user_home()?;
    uninstall_manifest_at(&home, ManifestHost::Word)?;
    uninstall_manifest_at(&home, ManifestHost::PowerPoint)
}

pub fn remove_trusted_certificate(paths: &OfficePaths) -> Result<(), String> {
    let home = user_home()?;
    let keychain = login_keychain(&home);
    if !keychain.exists() {
        return Ok(());
    }
    let output = run_security(&[
        "delete-certificate",
        "-c",
        CERTIFICATE_COMMON_NAME,
        keychain
            .to_str()
            .ok_or_else(|| "Login Keychain path is not valid UTF-8".to_string())?,
    ])?;
    if output.status.success() {
        return Ok(());
    }
    let message = command_error("security delete-certificate", &output);
    if message.contains("could not be found")
        || message.contains("The specified item could not be found")
    {
        Ok(())
    } else if paths.certificate.exists() {
        Err(message)
    } else {
        Ok(())
    }
}

pub fn trust_certificate(paths: &OfficePaths) -> Result<(), String> {
    let home = user_home()?;
    let keychain = login_keychain(&home);
    if !keychain.exists() {
        return Err(format!(
            "The user login Keychain does not exist: {}",
            keychain.display()
        ));
    }
    let _ = remove_trusted_certificate(paths);
    let output = run_security(&[
        "add-trusted-cert",
        "-d",
        "-r",
        "trustRoot",
        "-p",
        "ssl",
        "-k",
        keychain
            .to_str()
            .ok_or_else(|| "Login Keychain path is not valid UTF-8".to_string())?,
        paths
            .certificate
            .to_str()
            .ok_or_else(|| "Certificate path is not valid UTF-8".to_string())?,
    ])?;
    if output.status.success() {
        Ok(())
    } else {
        Err(command_error("security add-trusted-cert", &output))
    }
}

pub fn certificate_trust_status(paths: &OfficePaths) -> Result<bool, String> {
    if !paths.certificate.is_file() {
        return Ok(false);
    }
    let output = run_security(&[
        "verify-cert",
        "-c",
        paths
            .certificate
            .to_str()
            .ok_or_else(|| "Certificate path is not valid UTF-8".to_string())?,
        "-p",
        "ssl",
        "-n",
        "127.0.0.1",
        "-L",
        "-q",
    ])?;
    Ok(output.status.success())
}

pub fn verify_companion_health() -> Result<(), String> {
    let output = Command::new("/usr/bin/curl")
        .args([
            "--fail",
            "--silent",
            "--show-error",
            "--max-time",
            "5",
            "https://127.0.0.1:43127/health",
        ])
        .output()
        .map_err(|error| format!("Unable to run system curl: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        Err(command_error("curl --fail", &output))
    }
}

fn host_status(home: &Path, host: ManifestHost) -> OfficeHostInstallStatus {
    let path = manifest_path(home, host);
    OfficeHostInstallStatus {
        application_installed: application_installed(home, host),
        manifest_installed: path.is_file(),
        manifest_version: read_manifest_version(&path),
        manifest_path: path.display().to_string(),
    }
}

pub fn integration_status(
    paths: &OfficePaths,
    background: OfficeBackgroundStatus,
    companion: OfficeCompanionStatus,
) -> Result<OfficeIntegrationStatus, String> {
    let home = user_home()?;
    let keychain = login_keychain(&home);
    let trust = certificate_trust_status(paths);
    Ok(OfficeIntegrationStatus {
        word: host_status(&home, ManifestHost::Word),
        powerpoint: host_status(&home, ManifestHost::PowerPoint),
        expected_manifest_version: manifest_version(),
        certificate: CertificateInstallStatus {
            certificate_exists: paths.certificate.is_file(),
            private_key_exists: paths.private_key.is_file(),
            trusted: trust.as_ref().copied().unwrap_or(false),
            keychain_path: keychain.display().to_string(),
            last_error: trust.err(),
        },
        background,
        companion,
        office_ui_version: OFFICE_UI_VERSION.to_string(),
    })
}

pub fn open_office_application(host: ManifestHost) -> Result<(), String> {
    let application = match host {
        ManifestHost::Word => "Microsoft Word",
        ManifestHost::PowerPoint => "Microsoft PowerPoint",
    };
    let status = Command::new("/usr/bin/open")
        .args(["-a", application])
        .status()
        .map_err(|error| format!("Unable to launch {application}: {error}"))?;
    if status.success() {
        Ok(())
    } else {
        Err(format!(
            "macOS open failed to launch {application}: {status}"
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::office::manifest::{
        manifest_version, POWERPOINT_MANIFEST_FILE, WORD_MANIFEST_FILE,
    };
    use tempfile::TempDir;

    #[test]
    fn install_and_uninstall_touch_only_visualtex_manifest() {
        let temp = TempDir::new().unwrap();
        let home = temp.path();
        let word_directory = manifest_directory(home, ManifestHost::Word);
        fs::create_dir_all(&word_directory).unwrap();
        let other = word_directory.join("Other.Plugin.xml");
        fs::write(&other, "<OfficeApp/>").unwrap();

        let installed = install_manifest_at(home, ManifestHost::Word).unwrap();
        assert_eq!(installed.file_name().unwrap(), WORD_MANIFEST_FILE);
        assert!(installed.is_file());
        let expected_version = manifest_version();
        assert_eq!(
            read_manifest_version(&installed).as_deref(),
            Some(expected_version.as_str())
        );

        uninstall_manifest_at(home, ManifestHost::Word).unwrap();
        assert!(!installed.exists());
        assert!(other.is_file());
        assert!(word_directory.is_dir());
    }

    #[test]
    fn word_and_powerpoint_manifests_use_separate_directories_and_names() {
        let temp = TempDir::new().unwrap();
        let home = temp.path();
        let word = install_manifest_at(home, ManifestHost::Word).unwrap();
        let powerpoint = install_manifest_at(home, ManifestHost::PowerPoint).unwrap();
        assert_ne!(word, powerpoint);
        assert_eq!(word.file_name().unwrap(), WORD_MANIFEST_FILE);
        assert_eq!(powerpoint.file_name().unwrap(), POWERPOINT_MANIFEST_FILE);
        assert!(word
            .to_string_lossy()
            .contains("com.microsoft.Word/Data/Documents/wef"));
        assert!(powerpoint
            .to_string_lossy()
            .contains("com.microsoft.Powerpoint/Data/Documents/wef"));
    }

    #[test]
    fn install_migrates_legacy_unprefixed_manifest_and_preserves_registration_inode() {
        let temp = TempDir::new().unwrap();
        let home = temp.path();
        let legacy = legacy_manifest_path(home, ManifestHost::Word);
        fs::create_dir_all(legacy.parent().unwrap()).unwrap();
        fs::write(&legacy, "<OfficeApp>legacy</OfficeApp>").unwrap();

        let installed = install_manifest_at(home, ManifestHost::Word).unwrap();
        assert!(!legacy.exists());
        let first = fs::metadata(&installed).unwrap();
        install_manifest_at(home, ManifestHost::Word).unwrap();
        let second = fs::metadata(&installed).unwrap();

        #[cfg(unix)]
        {
            use std::os::unix::fs::MetadataExt;
            assert_eq!(first.ino(), second.ino());
        }
    }
}
