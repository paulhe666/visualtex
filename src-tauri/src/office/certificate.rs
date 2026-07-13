use crate::office::state::{OfficePaths, OFFICE_PORT, OFFICE_PROTOCOL_VERSION};
use getrandom::fill as random_fill;
use rcgen::{CertificateParams, DistinguishedName, DnType, KeyPair, SanType};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::net::{IpAddr, Ipv4Addr};
use std::path::Path;
use time::{Duration, OffsetDateTime};
use uuid::Uuid;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct InstallConfig {
    install_token: String,
    protocol_version: u32,
    port: u16,
    created_at: i64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CertificateMetadata {
    pub common_name: String,
    pub dns_names: Vec<String>,
    pub ip_addresses: Vec<String>,
    pub not_before: i64,
    pub not_after: i64,
    pub sha256_fingerprint: String,
    #[serde(default)]
    pub key_algorithm: String,
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

#[cfg(unix)]
fn sync_directory(path: &Path) -> Result<(), String> {
    let directory = fs::File::open(path).map_err(|error| {
        format!(
            "Unable to open directory {} for sync: {error}",
            path.display()
        )
    })?;
    directory
        .sync_all()
        .map_err(|error| format!("Unable to sync directory {}: {error}", path.display()))
}

#[cfg(not(unix))]
fn sync_directory(_path: &Path) -> Result<(), String> {
    // Windows does not allow opening a directory with std::fs::File solely
    // to call sync_all(). The temporary file itself has already been flushed
    // before the atomic rename, so no extra directory sync is available here.
    Ok(())
}

fn atomic_write(path: &Path, bytes: &[u8], mode: u32) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("Path has no parent: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let temporary = parent.join(format!(
        ".{}.{}.tmp",
        path.file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("visualtex"),
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
    set_mode(&temporary, mode)?;
    fs::rename(&temporary, path).map_err(|error| {
        let _ = fs::remove_file(&temporary);
        format!(
            "Unable to replace {} with {}: {error}",
            path.display(),
            temporary.display()
        )
    })?;
    set_mode(path, mode)?;
    sync_directory(parent)
}

fn new_install_token() -> Result<String, String> {
    let mut bytes = [0_u8; 32];
    random_fill(&mut bytes)
        .map_err(|error| format!("Unable to generate install token: {error}"))?;
    Ok(hex::encode(bytes))
}

fn read_install_config(path: &Path) -> Option<InstallConfig> {
    let bytes = fs::read(path).ok()?;
    let config = serde_json::from_slice::<InstallConfig>(&bytes).ok()?;
    (config.install_token.len() == 64
        && config.protocol_version == OFFICE_PROTOCOL_VERSION
        && config.port == OFFICE_PORT)
        .then_some(config)
}

fn ensure_install_config(paths: &OfficePaths) -> Result<String, String> {
    if let Some(config) = read_install_config(&paths.install) {
        set_mode(&paths.install, 0o600)?;
        return Ok(config.install_token);
    }

    let config = InstallConfig {
        install_token: new_install_token()?,
        protocol_version: OFFICE_PROTOCOL_VERSION,
        port: OFFICE_PORT,
        created_at: OffsetDateTime::now_utc().unix_timestamp(),
    };
    let json = serde_json::to_vec_pretty(&config)
        .map_err(|error| format!("Unable to encode install config: {error}"))?;
    atomic_write(&paths.install, &json, 0o600)?;
    Ok(config.install_token)
}

fn expected_key_algorithm() -> &'static str {
    #[cfg(target_os = "windows")]
    {
        "rsa-2048-sha256"
    }
    #[cfg(not(target_os = "windows"))]
    {
        "ecdsa-p256-sha256"
    }
}

fn certificate_metadata_matches_platform(paths: &OfficePaths) -> bool {
    fs::read(&paths.certificate_metadata)
        .ok()
        .and_then(|bytes| serde_json::from_slice::<CertificateMetadata>(&bytes).ok())
        .is_some_and(|metadata| metadata.key_algorithm == expected_key_algorithm())
}

fn ensure_certificate(paths: &OfficePaths) -> Result<(), String> {
    if paths.certificate.is_file()
        && paths.private_key.is_file()
        && paths.certificate_metadata.is_file()
        && certificate_metadata_matches_platform(paths)
    {
        set_mode(&paths.certificate, 0o644)?;
        set_mode(&paths.private_key, 0o600)?;
        set_mode(&paths.certificate_metadata, 0o644)?;
        return Ok(());
    }

    for path in [
        &paths.certificate,
        &paths.private_key,
        &paths.certificate_metadata,
    ] {
        if path.exists() {
            fs::remove_file(path)
                .map_err(|error| format!("Unable to remove stale {}: {error}", path.display()))?;
        }
    }

    let now = OffsetDateTime::now_utc();
    let not_before = now - Duration::days(1);
    let not_after = now + Duration::days(3650);
    let mut params = CertificateParams::default();
    params.not_before = not_before;
    params.not_after = not_after;
    params.distinguished_name = DistinguishedName::new();
    params
        .distinguished_name
        .push(DnType::OrganizationName, "VisualTeX");
    params
        .distinguished_name
        .push(DnType::CommonName, "VisualTeX Local Office Companion");
    params.subject_alt_names = vec![
        SanType::DnsName(
            "localhost"
                .try_into()
                .map_err(|error| format!("Invalid localhost SAN: {error}"))?,
        ),
        SanType::IpAddress(IpAddr::V4(Ipv4Addr::LOCALHOST)),
    ];

    #[cfg(target_os = "windows")]
    let key_pair = KeyPair::generate_for(&rcgen::PKCS_RSA_SHA256)
        .map_err(|error| format!("Unable to generate Office TLS RSA private key: {error}"))?;
    #[cfg(not(target_os = "windows"))]
    let key_pair = KeyPair::generate()
        .map_err(|error| format!("Unable to generate Office TLS private key: {error}"))?;
    let certificate = params
        .self_signed(&key_pair)
        .map_err(|error| format!("Unable to generate Office TLS certificate: {error}"))?;
    let certificate_pem = certificate.pem();
    let private_key_pem = key_pair.serialize_pem();
    let fingerprint = Sha256::digest(certificate.der().as_ref());
    let metadata = CertificateMetadata {
        common_name: "VisualTeX Local Office Companion".to_string(),
        dns_names: vec!["localhost".to_string()],
        ip_addresses: vec!["127.0.0.1".to_string()],
        not_before: not_before.unix_timestamp(),
        not_after: not_after.unix_timestamp(),
        sha256_fingerprint: hex::encode(fingerprint),
        key_algorithm: expected_key_algorithm().to_string(),
    };
    let metadata_json = serde_json::to_vec_pretty(&metadata)
        .map_err(|error| format!("Unable to encode certificate metadata: {error}"))?;

    atomic_write(&paths.private_key, private_key_pem.as_bytes(), 0o600)?;
    atomic_write(&paths.certificate, certificate_pem.as_bytes(), 0o644)?;
    atomic_write(&paths.certificate_metadata, &metadata_json, 0o644)
}

pub fn ensure_office_install(paths: &OfficePaths) -> Result<String, String> {
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create Office data directory: {error}"))?;
    set_mode(&paths.root, 0o700)?;
    fs::create_dir_all(&paths.sessions)
        .map_err(|error| format!("Unable to create Office session directory: {error}"))?;
    fs::create_dir_all(&paths.recovery)
        .map_err(|error| format!("Unable to create Office recovery directory: {error}"))?;
    set_mode(&paths.sessions, 0o700)?;
    set_mode(&paths.recovery, 0o700)?;

    ensure_certificate(paths)?;
    ensure_install_config(paths)
}

pub fn regenerate_certificate(paths: &OfficePaths) -> Result<(), String> {
    for path in [
        &paths.certificate,
        &paths.private_key,
        &paths.certificate_metadata,
    ] {
        if path.exists() {
            fs::remove_file(path)
                .map_err(|error| format!("Unable to remove {}: {error}", path.display()))?;
        }
    }
    ensure_certificate(paths)
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;
    use x509_parser::extensions::GeneralName;
    use x509_parser::pem::parse_x509_pem;
    use x509_parser::prelude::FromDer;
    use x509_parser::prelude::X509Certificate;

    fn paths(temp: &TempDir) -> OfficePaths {
        let root = temp.path().join("office");
        OfficePaths {
            certificate: root.join("localhost-cert.pem"),
            private_key: root.join("localhost-key.pem"),
            certificate_metadata: root.join("certificate.json"),
            install: root.join("install.json"),
            sessions: root.join("sessions"),
            recovery: root.join("recovery"),
            formula_cache: root.join("formulas"),
            ui_root: temp.path().join("ui"),
            root,
        }
    }

    #[test]
    fn generated_certificate_contains_required_sans_and_validity() {
        let temp = TempDir::new().expect("temp dir");
        let paths = paths(&temp);
        ensure_office_install(&paths).expect("office install");
        let bytes = fs::read(&paths.certificate).expect("certificate pem");
        let (_, pem) = parse_x509_pem(&bytes).expect("parse pem");
        let (_, certificate) = X509Certificate::from_der(&pem.contents).expect("parse x509");
        let san = certificate
            .subject_alternative_name()
            .expect("SAN extension")
            .expect("SAN must exist");
        let has_localhost = san
            .value
            .general_names
            .iter()
            .any(|name| matches!(name, GeneralName::DNSName(value) if *value == "localhost"));
        let has_loopback =
            san.value.general_names.iter().any(
                |name| matches!(name, GeneralName::IPAddress(value) if *value == [127, 0, 0, 1]),
            );
        assert!(has_localhost);
        assert!(has_loopback);
        let validity = certificate.validity();
        assert!(validity.not_before.timestamp() <= OffsetDateTime::now_utc().unix_timestamp());
        assert!(
            validity.not_after.timestamp()
                > (OffsetDateTime::now_utc() + Duration::days(365)).unix_timestamp()
        );
    }

    #[cfg(unix)]
    #[test]
    fn office_private_files_have_restrictive_permissions() {
        use std::os::unix::fs::PermissionsExt;
        let temp = TempDir::new().expect("temp dir");
        let paths = paths(&temp);
        ensure_office_install(&paths).expect("office install");
        assert_eq!(
            fs::metadata(&paths.root).unwrap().permissions().mode() & 0o777,
            0o700
        );
        assert_eq!(
            fs::metadata(&paths.private_key)
                .unwrap()
                .permissions()
                .mode()
                & 0o777,
            0o600
        );
        assert_eq!(
            fs::metadata(&paths.install).unwrap().permissions().mode() & 0o777,
            0o600
        );
        assert_eq!(
            fs::metadata(&paths.certificate)
                .unwrap()
                .permissions()
                .mode()
                & 0o777,
            0o644
        );
    }

    #[test]
    fn install_tokens_are_random_and_persisted() {
        let first_temp = TempDir::new().expect("first temp dir");
        let second_temp = TempDir::new().expect("second temp dir");
        let first_paths = paths(&first_temp);
        let second_paths = paths(&second_temp);
        let first = ensure_office_install(&first_paths).expect("first install");
        let repeated = ensure_office_install(&first_paths).expect("repeat install");
        let second = ensure_office_install(&second_paths).expect("second install");
        assert_eq!(first.len(), 64);
        assert_eq!(first, repeated);
        assert_ne!(first, second);
    }
}
