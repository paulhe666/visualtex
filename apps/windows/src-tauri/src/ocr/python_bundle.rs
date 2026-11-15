use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::fs::{self, File};
use std::io::{self, Read};
use std::path::{Path, PathBuf};
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Manager};
use uuid::Uuid;
use zip::ZipArchive;

const WINDOWS_PYTHON_RESOURCE: &str = "ocr-python/windows-x64";
const EXPECTED_PYTHON_VERSION: &str = "3.12.10";
const MAX_ARCHIVE_ENTRIES: usize = 20_000;
const MAX_UNPACKED_BYTES: u64 = 256 * 1024 * 1024;

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BundleFileRecord {
    pub name: String,
    pub size: u64,
    pub sha256: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WindowsWheelhouseManifest {
    pub lock: BundleFileRecord,
    pub files: Vec<BundleFileRecord>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WindowsOpenMpRuntimeManifest {
    pub file: BundleFileRecord,
    pub version: String,
    pub source: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WindowsAppLocalRuntimeManifest {
    pub open_mp: WindowsOpenMpRuntimeManifest,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WindowsPythonBundleManifest {
    pub schema_version: u32,
    pub platform: String,
    pub architecture: String,
    pub python_version: String,
    pub pip_version: String,
    pub archive: BundleFileRecord,
    pub app_local_runtime: WindowsAppLocalRuntimeManifest,
    pub wheelhouse: WindowsWheelhouseManifest,
}

#[derive(Debug, Clone)]
pub struct WindowsOfflineInstallAssets {
    pub wheelhouse: PathBuf,
    pub lockfile: PathBuf,
    pub manifest: WindowsPythonBundleManifest,
}

#[derive(Debug, Clone)]
pub struct WindowsPythonBundle {
    root: PathBuf,
    pub manifest: WindowsPythonBundleManifest,
}

#[cfg(debug_assertions)]
fn development_bundle_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("resources")
        .join("ocr-python")
        .join("windows-x64")
}

fn parse_manifest(path: &Path) -> Result<WindowsPythonBundleManifest, String> {
    let content = fs::read(path).map_err(|error| {
        format!(
            "Unable to read bundled Python manifest {}: {error}",
            path.display()
        )
    })?;
    let manifest: WindowsPythonBundleManifest = serde_json::from_slice(&content)
        .map_err(|error| format!("Bundled Python manifest is invalid: {error}"))?;
    if manifest.schema_version != 2 {
        return Err(format!(
            "Unsupported bundled Python manifest schema: {}",
            manifest.schema_version
        ));
    }
    if manifest.platform != "windows" || manifest.architecture != "x64" {
        return Err(format!(
            "Bundled Python target mismatch: {}/{}",
            manifest.platform, manifest.architecture
        ));
    }
    if manifest.python_version != EXPECTED_PYTHON_VERSION {
        return Err(format!(
            "Bundled Python version mismatch: {} (expected {EXPECTED_PYTHON_VERSION})",
            manifest.python_version
        ));
    }
    validate_resource_name(&manifest.archive.name, "bundled Python archive")?;
    validate_resource_name(
        &manifest.app_local_runtime.open_mp.file.name,
        "app-local Microsoft OpenMP runtime",
    )?;
    if !manifest
        .app_local_runtime
        .open_mp
        .file
        .name
        .eq_ignore_ascii_case("vcomp140.dll")
        || manifest.app_local_runtime.open_mp.file.size == 0
    {
        return Err("Bundled Python manifest has an invalid app-local vcomp140.dll record".to_string());
    }
    validate_resource_name(&manifest.wheelhouse.lock.name, "OCR dependency lock")?;
    if manifest.wheelhouse.files.is_empty() {
        return Err("Bundled OCR wheelhouse manifest contains no wheels".to_string());
    }
    let mut names = std::collections::BTreeSet::new();
    for record in &manifest.wheelhouse.files {
        validate_resource_name(&record.name, "OCR dependency wheel")?;
        if !record.name.to_ascii_lowercase().ends_with(".whl") {
            return Err(format!("Bundled OCR dependency is not a wheel: {}", record.name));
        }
        if !names.insert(record.name.to_ascii_lowercase()) {
            return Err(format!("Bundled OCR wheelhouse contains a duplicate: {}", record.name));
        }
    }
    Ok(manifest)
}

fn validate_resource_name(name: &str, label: &str) -> Result<(), String> {
    if name.is_empty() || name.contains('/') || name.contains('\\') || name == "." || name == ".." {
        return Err(format!("Unsafe {label} name: {name}"));
    }
    Ok(())
}

pub fn locate_bundle(app: &AppHandle) -> Result<WindowsPythonBundle, String> {
    if let Ok(root) = app
        .path()
        .resolve(WINDOWS_PYTHON_RESOURCE, BaseDirectory::Resource)
    {
        let manifest_path = root.join("manifest.json");
        if manifest_path.is_file() {
            return Ok(WindowsPythonBundle {
                manifest: parse_manifest(&manifest_path)?,
                root,
            });
        }
    }

    #[cfg(debug_assertions)]
    {
        let root = development_bundle_root();
        let manifest_path = root.join("manifest.json");
        if manifest_path.is_file() {
            return Ok(WindowsPythonBundle {
                manifest: parse_manifest(&manifest_path)?,
                root,
            });
        }
    }

    Err("The bundled Windows x64 Python 3.12 runtime is missing. Reinstall VisualTeX using the complete installer.".to_string())
}

fn sha256_file(path: &Path) -> Result<String, String> {
    let mut file = File::open(path)
        .map_err(|error| format!("Unable to open {} for verification: {error}", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = vec![0_u8; 1024 * 1024];
    loop {
        let count = file
            .read(&mut buffer)
            .map_err(|error| format!("Unable to verify {}: {error}", path.display()))?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(hex::encode_upper(digest.finalize()))
}

fn verify_file_record(root: &Path, record: &BundleFileRecord, label: &str) -> Result<PathBuf, String> {
    let path = root.join(&record.name);
    let metadata = fs::metadata(&path)
        .map_err(|error| format!("Bundled {label} is missing {}: {error}", path.display()))?;
    if !metadata.is_file() || metadata.len() != record.size {
        return Err(format!("Bundled {label} size mismatch: {}", path.display()));
    }
    let actual = sha256_file(&path)?;
    if !actual.eq_ignore_ascii_case(&record.sha256) {
        return Err(format!(
            "Bundled {label} checksum mismatch for {}. Expected {}, actual {}",
            path.display(), record.sha256, actual
        ));
    }
    Ok(path)
}

fn verify_archive(bundle: &WindowsPythonBundle) -> Result<PathBuf, String> {
    verify_file_record(&bundle.root, &bundle.manifest.archive, "Python archive")
}

fn verify_wheelhouse(bundle: &WindowsPythonBundle) -> Result<WindowsOfflineInstallAssets, String> {
    let wheelhouse = bundle.root.join("wheelhouse");
    if !wheelhouse.is_dir() {
        return Err(format!(
            "Bundled OCR wheelhouse directory is missing: {}",
            wheelhouse.display()
        ));
    }
    let lockfile = verify_file_record(
        &bundle.root,
        &bundle.manifest.wheelhouse.lock,
        "OCR dependency lock",
    )?;
    let mut expected = std::collections::BTreeSet::new();
    for record in &bundle.manifest.wheelhouse.files {
        verify_file_record(&wheelhouse, record, "OCR dependency wheel")?;
        expected.insert(record.name.to_ascii_lowercase());
    }
    let actual = fs::read_dir(&wheelhouse)
        .map_err(|error| format!("Unable to enumerate bundled OCR wheelhouse: {error}"))?
        .filter_map(Result::ok)
        .filter(|entry| entry.file_type().is_ok_and(|kind| kind.is_file()))
        .map(|entry| entry.file_name().to_string_lossy().to_ascii_lowercase())
        .collect::<std::collections::BTreeSet<_>>();
    if actual != expected {
        return Err("Bundled OCR wheelhouse contains missing or unexpected files".to_string());
    }
    Ok(WindowsOfflineInstallAssets {
        wheelhouse,
        lockfile,
        manifest: bundle.manifest.clone(),
    })
}

fn extract_archive(archive_path: &Path, destination: &Path) -> Result<(), String> {
    fs::create_dir_all(destination).map_err(|error| {
        format!(
            "Unable to create bundled Python staging directory {}: {error}",
            destination.display()
        )
    })?;
    let file = File::open(archive_path)
        .map_err(|error| format!("Unable to open {}: {error}", archive_path.display()))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|error| format!("Bundled Python ZIP is invalid: {error}"))?;
    if archive.len() > MAX_ARCHIVE_ENTRIES {
        return Err(format!(
            "Bundled Python ZIP contains too many entries: {}",
            archive.len()
        ));
    }

    let mut unpacked_bytes = 0_u64;
    for index in 0..archive.len() {
        let mut entry = archive
            .by_index(index)
            .map_err(|error| format!("Unable to read bundled Python ZIP entry: {error}"))?;
        let relative = entry
            .enclosed_name()
            .ok_or_else(|| format!("Unsafe bundled Python ZIP path: {}", entry.name()))?
            .to_path_buf();
        unpacked_bytes = unpacked_bytes
            .checked_add(entry.size())
            .ok_or_else(|| "Bundled Python unpacked size overflow".to_string())?;
        if unpacked_bytes > MAX_UNPACKED_BYTES {
            return Err(format!(
                "Bundled Python expands beyond the {} MiB safety limit",
                MAX_UNPACKED_BYTES / 1024 / 1024
            ));
        }
        if entry.unix_mode().is_some_and(|mode| mode & 0o170000 == 0o120000) {
            return Err(format!(
                "Symbolic links are not allowed in bundled Python: {}",
                entry.name()
            ));
        }
        let output = destination.join(relative);
        if entry.is_dir() {
            fs::create_dir_all(&output)
                .map_err(|error| format!("Unable to create {}: {error}", output.display()))?;
            continue;
        }
        if let Some(parent) = output.parent() {
            fs::create_dir_all(parent)
                .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
        }
        let mut target = File::create(&output)
            .map_err(|error| format!("Unable to create {}: {error}", output.display()))?;
        io::copy(&mut entry, &mut target)
            .map_err(|error| format!("Unable to extract {}: {error}", output.display()))?;
    }
    Ok(())
}

fn verify_extracted_tree(root: &Path, manifest: &WindowsPythonBundleManifest) -> Result<(), String> {
    for required in [
        "python.exe",
        "python312.dll",
        "python312.zip",
        "python312._pth",
        "vcomp140.dll",
        "visualtex-python.json",
        "Lib/site-packages/pip/__init__.py",
        "Lib/site-packages/sitecustomize.py",
    ] {
        let path = root.join(required);
        if !path.is_file() {
            return Err(format!(
                "Bundled Python extraction is incomplete; missing {}",
                path.display()
            ));
        }
    }
    let openmp = root.join(&manifest.app_local_runtime.open_mp.file.name);
    let openmp_metadata = fs::metadata(&openmp)
        .map_err(|error| format!("Unable to inspect bundled app-local OpenMP runtime: {error}"))?;
    if openmp_metadata.len() != manifest.app_local_runtime.open_mp.file.size {
        return Err(format!(
            "Bundled app-local OpenMP runtime size mismatch: {}",
            openmp.display()
        ));
    }
    let openmp_sha256 = sha256_file(&openmp)?;
    if !openmp_sha256.eq_ignore_ascii_case(&manifest.app_local_runtime.open_mp.file.sha256) {
        return Err(format!(
            "Bundled app-local OpenMP runtime checksum mismatch: {}",
            openmp.display()
        ));
    }

    let metadata: serde_json::Value = serde_json::from_slice(
        &fs::read(root.join("visualtex-python.json"))
            .map_err(|error| format!("Unable to read bundled Python metadata: {error}"))?,
    )
    .map_err(|error| format!("Bundled Python metadata is invalid: {error}"))?;
    if metadata
        .get("pythonVersion")
        .and_then(serde_json::Value::as_str)
        != Some(manifest.python_version.as_str())
        || metadata
            .get("architecture")
            .and_then(serde_json::Value::as_str)
            != Some("x64")
    {
        return Err("Bundled Python metadata does not match the manifest".to_string());
    }
    Ok(())
}

fn install_resolved_bundle(
    bundle: WindowsPythonBundle,
    destination: &Path,
) -> Result<WindowsPythonBundleManifest, String> {
    let archive = verify_archive(&bundle)?;
    let parent = destination
        .parent()
        .ok_or_else(|| "Bundled Python destination has no parent directory".to_string())?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;

    let suffix = Uuid::new_v4();
    let staging = parent.join(format!(".python-installing-{suffix}"));
    let backup = parent.join(format!(".python-backup-{suffix}"));
    fs::remove_dir_all(&staging).ok();
    fs::remove_dir_all(&backup).ok();

    let result = (|| {
        extract_archive(&archive, &staging)?;
        verify_extracted_tree(&staging, &bundle.manifest)?;
        if destination.exists() {
            fs::rename(destination, &backup).map_err(|error| {
                format!("Unable to back up the existing private Python runtime: {error}")
            })?;
        }
        if let Err(error) = fs::rename(&staging, destination) {
            if backup.exists() {
                let _ = fs::rename(&backup, destination);
            }
            return Err(format!(
                "Unable to activate the bundled private Python runtime: {error}"
            ));
        }
        fs::remove_dir_all(&backup).ok();
        Ok(())
    })();

    if result.is_err() {
        fs::remove_dir_all(&staging).ok();
        if backup.exists() && !destination.exists() {
            let _ = fs::rename(&backup, destination);
        }
    }
    result?;
    Ok(bundle.manifest)
}

pub fn bundle_available(app: &AppHandle) -> bool {
    locate_bundle(app).is_ok()
}

pub fn offline_install_assets_from_root(
    root: &Path,
) -> Result<WindowsOfflineInstallAssets, String> {
    let bundle = WindowsPythonBundle {
        manifest: parse_manifest(&root.join("manifest.json"))?,
        root: root.to_path_buf(),
    };
    verify_archive(&bundle)?;
    verify_wheelhouse(&bundle)
}

pub fn locate_offline_install_assets(
    app: &AppHandle,
) -> Result<WindowsOfflineInstallAssets, String> {
    let bundle = locate_bundle(app)?;
    verify_archive(&bundle)?;
    verify_wheelhouse(&bundle)
}

pub fn install_bundle(
    app: &AppHandle,
    destination: &Path,
) -> Result<WindowsPythonBundleManifest, String> {
    install_resolved_bundle(locate_bundle(app)?, destination)
}

pub fn install_bundle_from_root(
    bundle_root: &Path,
    destination: &Path,
) -> Result<WindowsPythonBundleManifest, String> {
    let manifest = parse_manifest(&bundle_root.join("manifest.json"))?;
    install_resolved_bundle(
        WindowsPythonBundle {
            root: bundle_root.to_path_buf(),
            manifest,
        },
        destination,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::write::SimpleFileOptions;

    #[test]
    fn rejects_zip_path_traversal() {
        let root = tempfile::tempdir().unwrap();
        let archive_path = root.path().join("unsafe.zip");
        let file = File::create(&archive_path).unwrap();
        let mut writer = zip::ZipWriter::new(file);
        writer
            .start_file("../escape.txt", SimpleFileOptions::default())
            .unwrap();
        writer.write_all(b"unsafe").unwrap();
        writer.finish().unwrap();
        let destination = root.path().join("out");
        let error = extract_archive(&archive_path, &destination).unwrap_err();
        assert!(error.contains("Unsafe bundled Python ZIP path"));
        assert!(!root.path().join("escape.txt").exists());
    }
}
