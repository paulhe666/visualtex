use crate::office::background::OfficeBackgroundStatus;
use crate::office::manifest::ManifestHost;
use crate::office::state::{OfficeCompanionStatus, OfficePaths, OFFICE_UI_VERSION};
use serde::Serialize;

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
    pub certificate: CertificateInstallStatus,
    pub background: OfficeBackgroundStatus,
    pub companion: OfficeCompanionStatus,
    pub office_ui_version: String,
}

pub fn install_available_manifests() -> Result<(), String> {
    Err("Use the Windows OLE or VSTO installer on Windows".to_string())
}

pub fn uninstall_manifests() -> Result<(), String> {
    Ok(())
}

pub fn trust_certificate(_paths: &OfficePaths) -> Result<(), String> {
    Ok(())
}

pub fn remove_trusted_certificate(_paths: &OfficePaths) -> Result<(), String> {
    Ok(())
}

pub fn verify_companion_health() -> Result<(), String> {
    Ok(())
}

pub fn integration_status(
    paths: &OfficePaths,
    background: OfficeBackgroundStatus,
    companion: OfficeCompanionStatus,
) -> Result<OfficeIntegrationStatus, String> {
    let platform = if cfg!(target_os = "windows") {
        "Windows"
    } else {
        std::env::consts::OS
    };
    let empty_host = |name: &str| OfficeHostInstallStatus {
        application_installed: false,
        manifest_installed: false,
        manifest_version: None,
        manifest_path: format!("{platform} {name} integration is managed separately"),
    };
    Ok(OfficeIntegrationStatus {
        word: empty_host("Word"),
        powerpoint: empty_host("PowerPoint"),
        certificate: CertificateInstallStatus {
            certificate_exists: paths.certificate.is_file(),
            private_key_exists: paths.private_key.is_file(),
            trusted: paths.certificate.is_file(),
            keychain_path: String::new(),
            last_error: None,
        },
        background,
        companion,
        office_ui_version: OFFICE_UI_VERSION.to_string(),
    })
}

pub fn open_office_application(host: ManifestHost) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        let executable = match host {
            ManifestHost::Word => "winword.exe",
            ManifestHost::PowerPoint => "powerpnt.exe",
        };
        return std::process::Command::new("cmd.exe")
            .args(["/C", "start", "", executable])
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("Unable to launch {executable}: {error}"));
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = host;
        Err("Office application launching is unsupported on this platform".to_string())
    }
}
