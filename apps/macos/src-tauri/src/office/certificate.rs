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
        .map_err(|error| format!("Unable to generate companion token: {error}"))?;
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
        .map_err(|error| format!("Unable to encode companion config: {error}"))?;
    atomic_write(&paths.install, &json, 0o600)?;
    Ok(config.install_token)
}

fn certificate_metadata_matches_platform(paths: &OfficePaths) -> bool {
    fs::read(&paths.certificate_metadata)
        .ok()
        .and_then(|bytes| serde_json::from_slice::<CertificateMetadata>(&bytes).ok())
        .is_some_and(|metadata| metadata.key_algorithm == "ecdsa-p256-sha256")
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
        .push(DnType::CommonName, "VisualTeX Local Companion");
    params.subject_alt_names = vec![
        SanType::DnsName(
            "localhost"
                .try_into()
                .map_err(|error| format!("Invalid localhost SAN: {error}"))?,
        ),
        SanType::IpAddress(IpAddr::V4(Ipv4Addr::LOCALHOST)),
    ];

    let key_pair = KeyPair::generate()
        .map_err(|error| format!("Unable to generate companion TLS private key: {error}"))?;
    let certificate = params
        .self_signed(&key_pair)
        .map_err(|error| format!("Unable to generate companion TLS certificate: {error}"))?;
    let certificate_pem = certificate.pem();
    let private_key_pem = key_pair.serialize_pem();
    let fingerprint = Sha256::digest(certificate.der().as_ref());
    let metadata = CertificateMetadata {
        common_name: "VisualTeX Local Companion".to_string(),
        dns_names: vec!["localhost".to_string()],
        ip_addresses: vec!["127.0.0.1".to_string()],
        not_before: not_before.unix_timestamp(),
        not_after: not_after.unix_timestamp(),
        sha256_fingerprint: hex::encode(fingerprint),
        key_algorithm: "ecdsa-p256-sha256".to_string(),
    };
    let metadata_json = serde_json::to_vec_pretty(&metadata)
        .map_err(|error| format!("Unable to encode certificate metadata: {error}"))?;

    atomic_write(&paths.private_key, private_key_pem.as_bytes(), 0o600)?;
    atomic_write(&paths.certificate, certificate_pem.as_bytes(), 0o644)?;
    atomic_write(&paths.certificate_metadata, &metadata_json, 0o644)
}

/// Prepare the private local companion runtime. This creates a loopback TLS
/// certificate and an API token, but does not trust the certificate in the
/// system keychain and does not install any Office.js manifest.
pub fn ensure_companion_runtime(paths: &OfficePaths) -> Result<String, String> {
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create companion data directory: {error}"))?;
    set_mode(&paths.root, 0o700)?;
    for directory in [&paths.sessions, &paths.recovery, &paths.formula_cache] {
        fs::create_dir_all(directory)
            .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
        set_mode(directory, 0o700)?;
    }
    ensure_certificate(paths)?;
    ensure_install_config(paths)
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

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
            root,
        }
    }

    #[test]
    fn companion_runtime_generates_persistent_private_credentials() {
        let temp = TempDir::new().expect("temp dir");
        let paths = paths(&temp);
        let first = ensure_companion_runtime(&paths).expect("runtime");
        let repeated = ensure_companion_runtime(&paths).expect("repeat runtime");
        assert_eq!(first.len(), 64);
        assert_eq!(first, repeated);
        assert!(paths.certificate.is_file());
        assert!(paths.private_key.is_file());
        assert!(paths.certificate_metadata.is_file());
    }
}
