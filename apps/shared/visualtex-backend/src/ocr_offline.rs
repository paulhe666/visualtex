use flate2::read::GzDecoder;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::{Component, Path, PathBuf};
use tar::{Archive, EntryType};
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Manager};
use uuid::Uuid;

pub const OFFLINE_DEFAULT_MODEL: &str = "PP-FormulaNet_plus-M";
const OFFLINE_RUNTIME_LAYOUT_VERSION: u32 = 3;
const OFFLINE_BUNDLE_RESOURCE: &str = "ocr/offline/macos-arm64";
const INSTALLED_MANIFEST_FILE: &str = "offline-manifest.json";
const KNOWN_MODELS: &[&str] = &[
    "PP-FormulaNet_plus-S",
    "PP-FormulaNet_plus-M",
    "PP-FormulaNet_plus-L",
];
const OPTIONAL_MODELS: &[&str] = &["PP-FormulaNet_plus-S", "PP-FormulaNet_plus-L"];
const MAX_OPTIONAL_MODEL_PACK_BYTES: u64 = 1024 * 1024 * 1024;
const MAX_ARCHIVE_ENTRIES: usize = 200_000;
const MAX_ARCHIVE_UNPACKED_BYTES: u64 = 3 * 1024 * 1024 * 1024;

fn expected_model_hashes(model: &str) -> Option<BTreeMap<&'static str, &'static str>> {
    let values = match model {
        "PP-FormulaNet_plus-S" => [
            (
                "inference.json",
                "01238434e33df83588e2627f350559b576e34551d2b2ffea148345032de56c00",
            ),
            (
                "inference.pdiparams",
                "e464f94412feaa98f8791eacc84684f887b3569e30e80c52b8112e9cf7d4069b",
            ),
            (
                "inference.yml",
                "96062655d94c21d39274328dbc82c1a487e66addb8425f5a7fd5b7dfb2421ec3",
            ),
        ],
        "PP-FormulaNet_plus-M" => [
            (
                "inference.json",
                "8333a7f650766a748e273c550d278601dd19dfeee1c4b01038ff632f134d9884",
            ),
            (
                "inference.pdiparams",
                "f16ef9b5c8227da70d3ec969a5195f4d62c1154427b883f4d6cff07633654041",
            ),
            (
                "inference.yml",
                "87b5f3d7f2b2fe553627d77b37f496608ca150ebd0ef62d362591edca47b5538",
            ),
        ],
        "PP-FormulaNet_plus-L" => [
            (
                "inference.json",
                "ad259c4b896d99aa3479336b9121112fb40ff1ababfbf8765a3428a3b86df582",
            ),
            (
                "inference.pdiparams",
                "4245c39c181d1d21e472bc85c7434df9b23f177be46552c0542bf153addbc355",
            ),
            (
                "inference.yml",
                "afc92a2737268da0499c37b0b6741da268c369fd7424667fcfeb8fa6c7b22d30",
            ),
        ],
        _ => return None,
    };
    Some(values.into_iter().collect())
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfflineFileRecord {
    pub name: String,
    pub size: u64,
    pub sha256: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfflineArchiveSet {
    pub runtime: OfflineFileRecord,
    pub default_model: OfflineFileRecord,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfflineModelManifest {
    pub name: String,
    pub files: BTreeMap<String, OfflineFileRecord>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfflineBundleManifest {
    pub schema_version: u32,
    pub platform: String,
    pub architecture: String,
    pub build_fingerprint: String,
    pub runtime_layout_version: u32,
    pub archives: OfflineArchiveSet,
    pub default_model: OfflineModelManifest,
}

#[derive(Debug, Clone)]
pub struct OfflineBundle {
    pub root: PathBuf,
    pub manifest: OfflineBundleManifest,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
struct OptionalModelPackManifest {
    schema_version: u32,
    platform: String,
    architecture: String,
    model: String,
    files: BTreeMap<String, OfflineFileRecord>,
}

#[cfg(debug_assertions)]
fn development_bundle_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("dist-ocr")
        .join("macos-arm64")
}

fn read_manifest(root: &Path) -> Result<OfflineBundleManifest, String> {
    let path = root.join("manifest.json");
    let content = fs::read_to_string(&path).map_err(|error| {
        format!(
            "Unable to read offline OCR manifest {}: {error}",
            path.display()
        )
    })?;
    let manifest: OfflineBundleManifest = serde_json::from_str(&content)
        .map_err(|error| format!("Offline OCR manifest is invalid: {error}"))?;
    if manifest.schema_version != 1 {
        return Err(format!(
            "Unsupported offline OCR manifest schema: {}",
            manifest.schema_version
        ));
    }
    if manifest.platform != "macos" || manifest.architecture != "arm64" {
        return Err(format!(
            "Offline OCR bundle target mismatch: {}/{}",
            manifest.platform, manifest.architecture
        ));
    }
    if manifest.runtime_layout_version != OFFLINE_RUNTIME_LAYOUT_VERSION {
        return Err(format!(
            "Offline OCR runtime layout mismatch: {} (expected {})",
            manifest.runtime_layout_version, OFFLINE_RUNTIME_LAYOUT_VERSION
        ));
    }
    if manifest.default_model.name != OFFLINE_DEFAULT_MODEL {
        return Err(format!(
            "Offline OCR default model mismatch: {}",
            manifest.default_model.name
        ));
    }
    Ok(manifest)
}

pub fn locate_bundle(app: &AppHandle) -> Result<OfflineBundle, String> {
    if let Ok(root) = app
        .path()
        .resolve(OFFLINE_BUNDLE_RESOURCE, BaseDirectory::Resource)
    {
        if root.join("manifest.json").is_file() {
            return Ok(OfflineBundle {
                manifest: read_manifest(&root)?,
                root,
            });
        }
    }
    #[cfg(debug_assertions)]
    {
        let root = development_bundle_root();
        if root.join("manifest.json").is_file() {
            return Ok(OfflineBundle {
                manifest: read_manifest(&root)?,
                root,
            });
        }
    }
    Err(
        "The bundled offline OCR runtime is missing. Reinstall VisualTeX using the complete macOS package."
            .to_string(),
    )
}

pub fn bundle_available(app: &AppHandle) -> bool {
    locate_bundle(app).is_ok()
}

fn write_synced_file(path: &Path, data: &[u8]) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("File has no parent directory: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let mut file = File::create(path)
        .map_err(|error| format!("Unable to create {}: {error}", path.display()))?;
    if let Err(error) = file.write_all(data).and_then(|_| file.sync_all()) {
        drop(file);
        fs::remove_file(path).ok();
        return Err(format!("Unable to write {}: {error}", path.display()));
    }
    Ok(())
}

#[cfg(unix)]
fn sync_directory(path: &Path) -> Result<(), String> {
    File::open(path)
        .and_then(|directory| directory.sync_all())
        .map_err(|error| format!("Unable to sync {}: {error}", path.display()))
}

#[cfg(not(unix))]
fn sync_directory(_path: &Path) -> Result<(), String> {
    Ok(())
}

fn sha256_file(path: &Path) -> Result<String, String> {
    let mut file = File::open(path).map_err(|error| {
        format!(
            "Unable to open {} for verification: {error}",
            path.display()
        )
    })?;
    let mut digest = Sha256::new();
    let mut buffer = [0_u8; 1024 * 1024];
    loop {
        let count = file
            .read(&mut buffer)
            .map_err(|error| format!("Unable to verify {}: {error}", path.display()))?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(hex::encode(digest.finalize()))
}

fn verify_record(root: &Path, record: &OfflineFileRecord) -> Result<PathBuf, String> {
    if record.name.contains('/') || record.name.contains('\\') || record.name.is_empty() {
        return Err(format!("Unsafe offline OCR archive name: {}", record.name));
    }
    let path = root.join(&record.name);
    let metadata = fs::metadata(&path)
        .map_err(|error| format!("Missing offline OCR archive {}: {error}", path.display()))?;
    if !metadata.is_file() || metadata.len() != record.size {
        return Err(format!(
            "Offline OCR archive size mismatch for {}",
            path.display()
        ));
    }
    let actual = sha256_file(&path)?;
    if actual != record.sha256 {
        return Err(format!(
            "Offline OCR archive checksum mismatch for {}",
            path.display()
        ));
    }
    Ok(path)
}

fn validate_archive_path(path: &Path) -> Result<(), String> {
    if path.as_os_str().is_empty() || path.is_absolute() {
        return Err(format!("Unsafe archive path: {}", path.display()));
    }
    for component in path.components() {
        match component {
            Component::Normal(_) | Component::CurDir => {}
            Component::ParentDir | Component::RootDir | Component::Prefix(_) => {
                return Err(format!("Unsafe archive path: {}", path.display()));
            }
        }
    }
    Ok(())
}

fn validate_archive_link(entry_path: &Path, link: &Path) -> Result<(), String> {
    if link.is_absolute() {
        return Err(format!(
            "Unsafe archive link {} -> {}",
            entry_path.display(),
            link.display()
        ));
    }
    let mut depth = entry_path
        .parent()
        .map(|path| path.components().count())
        .unwrap_or(0);
    for component in link.components() {
        match component {
            Component::Normal(_) => depth += 1,
            Component::CurDir => {}
            Component::ParentDir => {
                if depth == 0 {
                    return Err(format!(
                        "Archive link escapes destination: {} -> {}",
                        entry_path.display(),
                        link.display()
                    ));
                }
                depth -= 1;
            }
            Component::RootDir | Component::Prefix(_) => {
                return Err(format!(
                    "Unsafe archive link {} -> {}",
                    entry_path.display(),
                    link.display()
                ));
            }
        }
    }
    Ok(())
}

fn extract_archive(archive_path: &Path, destination: &Path) -> Result<(), String> {
    fs::create_dir_all(destination).map_err(|error| {
        format!(
            "Unable to create OCR extraction directory {}: {error}",
            destination.display()
        )
    })?;
    let file = File::open(archive_path)
        .map_err(|error| format!("Unable to open {}: {error}", archive_path.display()))?;
    let decoder = GzDecoder::new(file);
    let mut archive = Archive::new(decoder);
    let entries = archive
        .entries()
        .map_err(|error| format!("Unable to read {}: {error}", archive_path.display()))?;
    let mut entry_count = 0_usize;
    let mut unpacked_bytes = 0_u64;
    for entry in entries {
        entry_count += 1;
        if entry_count > MAX_ARCHIVE_ENTRIES {
            return Err(format!(
                "OCR archive has more than {MAX_ARCHIVE_ENTRIES} entries"
            ));
        }
        let mut entry =
            entry.map_err(|error| format!("Unable to read OCR archive entry: {error}"))?;
        let path = entry
            .path()
            .map_err(|error| format!("Invalid OCR archive path: {error}"))?
            .into_owned();
        validate_archive_path(&path)?;
        let entry_size = entry
            .header()
            .size()
            .map_err(|error| format!("Invalid OCR archive entry size: {error}"))?;
        unpacked_bytes = unpacked_bytes
            .checked_add(entry_size)
            .ok_or_else(|| "OCR archive unpacked size overflow".to_string())?;
        if unpacked_bytes > MAX_ARCHIVE_UNPACKED_BYTES {
            return Err(format!(
                "OCR archive expands beyond the {} GiB safety limit",
                MAX_ARCHIVE_UNPACKED_BYTES / 1024 / 1024 / 1024
            ));
        }
        let entry_type = entry.header().entry_type();
        if matches!(
            entry_type,
            EntryType::Block | EntryType::Char | EntryType::Fifo
        ) {
            return Err(format!(
                "Unsupported special file in OCR archive: {}",
                path.display()
            ));
        }
        if entry_type.is_symlink() || entry_type.is_hard_link() {
            let link = entry
                .link_name()
                .map_err(|error| format!("Invalid OCR archive link: {error}"))?
                .ok_or_else(|| format!("OCR archive link has no target: {}", path.display()))?;
            validate_archive_link(&path, &link)?;
        }
        let unpacked = entry
            .unpack_in(destination)
            .map_err(|error| format!("Unable to extract {}: {error}", path.display()))?;
        if !unpacked {
            return Err(format!(
                "OCR archive entry escaped its destination: {}",
                path.display()
            ));
        }
    }
    Ok(())
}

fn verify_installed_tree(root: &Path, manifest: &OfflineBundleManifest) -> Result<(), String> {
    let python = root.join("python/bin/python3");
    if !python.is_file() {
        return Err("Offline OCR archive did not install python/bin/python3".to_string());
    }
    let model_root = root
        .join("cache/paddlex/official_models")
        .join(&manifest.default_model.name);
    for (name, expected) in &manifest.default_model.files {
        let file = model_root.join(name);
        let metadata = fs::metadata(&file).map_err(|error| {
            format!(
                "Offline OCR model file is missing {}: {error}",
                file.display()
            )
        })?;
        if metadata.len() != expected.size || sha256_file(&file)? != expected.sha256 {
            return Err(format!(
                "Offline OCR model verification failed: {}",
                file.display()
            ));
        }
    }
    Ok(())
}

#[cfg(unix)]
fn ensure_python_executable(root: &Path) -> Result<(), String> {
    use std::os::unix::fs::PermissionsExt;
    let python = root.join("python/bin/python3");
    let mut permissions = fs::metadata(&python)
        .map_err(|error| format!("Unable to inspect bundled Python: {error}"))?
        .permissions();
    permissions.set_mode(0o755);
    fs::set_permissions(&python, permissions)
        .map_err(|error| format!("Unable to mark bundled Python executable: {error}"))
}

#[cfg(not(unix))]
fn ensure_python_executable(_root: &Path) -> Result<(), String> {
    Ok(())
}

fn preserve_optional_models(old_root: &Path, new_root: &Path) -> Result<(), String> {
    for model in KNOWN_MODELS {
        if *model == OFFLINE_DEFAULT_MODEL {
            continue;
        }
        let source = old_root.join("cache/paddlex/official_models").join(model);
        if !source.is_dir() {
            continue;
        }
        let target = new_root.join("cache/paddlex/official_models").join(model);
        if target.exists() {
            continue;
        }
        copy_directory_without_links(&source, &target)?;
    }
    Ok(())
}

fn copy_directory_without_links(source: &Path, target: &Path) -> Result<(), String> {
    fs::create_dir_all(target)
        .map_err(|error| format!("Unable to create {}: {error}", target.display()))?;
    for item in fs::read_dir(source)
        .map_err(|error| format!("Unable to read {}: {error}", source.display()))?
    {
        let item = item.map_err(|error| format!("Unable to read model directory: {error}"))?;
        let file_type = item
            .file_type()
            .map_err(|error| format!("Unable to inspect model file: {error}"))?;
        let destination = target.join(item.file_name());
        if file_type.is_symlink() {
            return Err(format!(
                "Refusing to preserve a symbolic link from the OCR model cache: {}",
                item.path().display()
            ));
        }
        if file_type.is_dir() {
            copy_directory_without_links(&item.path(), &destination)?;
        } else if file_type.is_file() {
            fs::copy(item.path(), &destination).map_err(|error| {
                format!(
                    "Unable to preserve optional OCR model {}: {error}",
                    item.path().display()
                )
            })?;
        }
    }
    Ok(())
}

pub fn install_bundle(
    app: &AppHandle,
    destination_root: &Path,
    mut progress: impl FnMut(&str, u8, &str, Option<String>),
) -> Result<OfflineBundleManifest, String> {
    let bundle = locate_bundle(app)?;
    progress("offline-verify", 8, "正在校验离线 OCR 安装包", None);
    let runtime_archive = verify_record(&bundle.root, &bundle.manifest.archives.runtime)?;
    let model_archive = verify_record(&bundle.root, &bundle.manifest.archives.default_model)?;

    let parent = destination_root
        .parent()
        .ok_or_else(|| "OCR runtime destination has no parent directory".to_string())?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create OCR data directory: {error}"))?;
    let suffix = Uuid::new_v4();
    let staging = parent.join(format!(".ocr-runtime-installing-{suffix}"));
    let backup = parent.join(format!(".ocr-runtime-backup-{suffix}"));
    fs::remove_dir_all(&staging).ok();
    fs::remove_dir_all(&backup).ok();

    let result = (|| {
        progress(
            "offline-runtime",
            20,
            "正在解压内置 Python 与 OCR 依赖",
            None,
        );
        extract_archive(&runtime_archive, &staging)?;
        progress(
            "offline-model",
            65,
            "正在安装内置 PP-FormulaNet M 模型",
            None,
        );
        extract_archive(&model_archive, &staging.join("cache"))?;
        if destination_root.is_dir() {
            preserve_optional_models(destination_root, &staging)?;
        }
        for directory in ["input", "processed", "logs", "tmp"] {
            fs::create_dir_all(staging.join(directory)).map_err(|error| {
                format!("Unable to create offline OCR {directory} directory: {error}")
            })?;
        }
        ensure_python_executable(&staging)?;
        verify_installed_tree(&staging, &bundle.manifest)?;
        fs::write(
            staging.join(INSTALLED_MANIFEST_FILE),
            serde_json::to_vec_pretty(&bundle.manifest)
                .map_err(|error| format!("Unable to serialize OCR manifest: {error}"))?,
        )
        .map_err(|error| format!("Unable to write installed OCR manifest: {error}"))?;

        progress("offline-activate", 88, "正在启用离线 OCR 运行环境", None);
        if destination_root.exists() {
            fs::rename(destination_root, &backup)
                .map_err(|error| format!("Unable to back up the existing OCR runtime: {error}"))?;
        }
        if let Err(error) = fs::rename(&staging, destination_root) {
            if backup.exists() {
                let _ = fs::rename(&backup, destination_root);
            }
            return Err(format!(
                "Unable to activate the offline OCR runtime: {error}"
            ));
        }
        fs::remove_dir_all(&backup).ok();
        Ok(())
    })();

    if result.is_err() {
        fs::remove_dir_all(&staging).ok();
        if backup.exists() && !destination_root.exists() {
            let _ = fs::rename(&backup, destination_root);
        }
    }
    result?;
    progress(
        "offline-complete",
        96,
        "离线 OCR 文件安装完成，正在验证",
        None,
    );
    Ok(bundle.manifest)
}

fn collect_regular_files(root: &Path) -> Result<Vec<PathBuf>, String> {
    let mut files = Vec::new();
    let mut pending = vec![root.to_path_buf()];
    while let Some(directory) = pending.pop() {
        for entry in fs::read_dir(&directory)
            .map_err(|error| format!("Unable to read model pack directory: {error}"))?
        {
            let entry =
                entry.map_err(|error| format!("Unable to read model pack entry: {error}"))?;
            let metadata = fs::symlink_metadata(entry.path())
                .map_err(|error| format!("Unable to inspect model pack entry: {error}"))?;
            if metadata.file_type().is_symlink() {
                return Err(format!(
                    "Symbolic links are not allowed in OCR model packs: {}",
                    entry.path().display()
                ));
            }
            if metadata.is_dir() {
                pending.push(entry.path());
            } else if metadata.is_file() {
                files.push(entry.path());
            } else {
                return Err(format!(
                    "Unsupported file type in OCR model pack: {}",
                    entry.path().display()
                ));
            }
        }
    }
    files.sort();
    Ok(files)
}

fn verify_optional_pack(pack_root: &Path) -> Result<OptionalModelPackManifest, String> {
    let manifest_path = pack_root.join("pack-manifest.json");
    let manifest_content = fs::read_to_string(&manifest_path).map_err(|error| {
        format!(
            "Unable to read optional OCR model manifest {}: {error}",
            manifest_path.display()
        )
    })?;
    let manifest: OptionalModelPackManifest = serde_json::from_str(&manifest_content)
        .map_err(|error| format!("Optional OCR model manifest is invalid: {error}"))?;
    if manifest.schema_version != 1
        || manifest.platform != "macos"
        || manifest.architecture != "arm64"
    {
        return Err("Optional OCR model pack target or schema is invalid".to_string());
    }
    if !OPTIONAL_MODELS.contains(&manifest.model.as_str()) {
        return Err(format!(
            "Only the optional S and L model packs can be imported; found {}",
            manifest.model
        ));
    }
    let expected = expected_model_hashes(&manifest.model)
        .ok_or_else(|| format!("Unknown OCR model: {}", manifest.model))?;
    if manifest.files.len() != expected.len() {
        return Err("Optional OCR model manifest has an unexpected file set".to_string());
    }

    let model_root = pack_root
        .join("paddlex/official_models")
        .join(&manifest.model);
    let mut allowed = vec![manifest_path.clone()];
    for (name, expected_hash) in expected {
        let record = manifest
            .files
            .get(name)
            .ok_or_else(|| format!("Optional OCR model manifest is missing {name}"))?;
        if record.name != name || record.sha256 != expected_hash {
            return Err(format!(
                "Optional OCR model manifest checksum is invalid for {name}"
            ));
        }
        let file = model_root.join(name);
        let metadata = fs::metadata(&file).map_err(|error| {
            format!(
                "Optional OCR model file is missing {}: {error}",
                file.display()
            )
        })?;
        if !metadata.is_file()
            || metadata.len() != record.size
            || sha256_file(&file)? != expected_hash
        {
            return Err(format!(
                "Optional OCR model verification failed: {}",
                file.display()
            ));
        }
        allowed.push(file);
    }
    allowed.sort();
    let actual = collect_regular_files(pack_root)?;
    if actual != allowed {
        return Err("Optional OCR model pack contains unexpected files".to_string());
    }
    Ok(manifest)
}

fn remove_path(path: &Path) -> Result<(), String> {
    if !path.exists() {
        return Ok(());
    }
    let metadata = fs::symlink_metadata(path).map_err(|error| {
        format!(
            "Unable to inspect {} during rollback: {error}",
            path.display()
        )
    })?;
    if metadata.is_dir() && !metadata.file_type().is_symlink() {
        fs::remove_dir_all(path).map_err(|error| {
            format!(
                "Unable to remove {} during rollback: {error}",
                path.display()
            )
        })
    } else {
        fs::remove_file(path).map_err(|error| {
            format!(
                "Unable to remove {} during rollback: {error}",
                path.display()
            )
        })
    }
}

struct OptionalModelActivation<'a> {
    incoming: &'a Path,
    target: &'a Path,
    backup: &'a Path,
    metadata_incoming: &'a Path,
    metadata_target: &'a Path,
    metadata_backup: &'a Path,
    models_root: &'a Path,
    manifest_directory: &'a Path,
}

impl OptionalModelActivation<'_> {
    fn activate(&self, mut sync: impl FnMut(&Path) -> Result<(), String>) -> Result<(), String> {
        let had_target = self.target.exists();
        let had_metadata = self.metadata_target.exists();
        let mut target_activated = false;
        let mut metadata_activated = false;

        let activation = (|| {
            if had_target {
                fs::rename(self.target, self.backup).map_err(|error| {
                    format!("Unable to back up the existing OCR model: {error}")
                })?;
            }
            if had_metadata {
                fs::rename(self.metadata_target, self.metadata_backup).map_err(|error| {
                    format!("Unable to back up the existing OCR model metadata: {error}")
                })?;
            }
            fs::rename(self.incoming, self.target)
                .map_err(|error| format!("Unable to activate the OCR model pack: {error}"))?;
            target_activated = true;
            fs::rename(self.metadata_incoming, self.metadata_target)
                .map_err(|error| format!("Unable to activate OCR model metadata: {error}"))?;
            metadata_activated = true;
            sync(self.models_root)?;
            sync(self.manifest_directory)?;
            Ok(())
        })();

        if let Err(error) = activation {
            let mut rollback_errors = Vec::new();
            if metadata_activated {
                if let Err(rollback_error) = remove_path(self.metadata_target) {
                    rollback_errors.push(rollback_error);
                }
            }
            if had_metadata && self.metadata_backup.exists() {
                if let Err(rollback_error) = fs::rename(self.metadata_backup, self.metadata_target)
                {
                    rollback_errors.push(format!(
                        "Unable to restore OCR model metadata {}: {rollback_error}",
                        self.metadata_target.display()
                    ));
                }
            }
            if target_activated {
                if let Err(rollback_error) = remove_path(self.target) {
                    rollback_errors.push(rollback_error);
                }
            }
            if had_target && self.backup.exists() {
                if let Err(rollback_error) = fs::rename(self.backup, self.target) {
                    rollback_errors.push(format!(
                        "Unable to restore OCR model {}: {rollback_error}",
                        self.target.display()
                    ));
                }
            }
            if let Err(rollback_error) = remove_path(self.incoming) {
                rollback_errors.push(rollback_error);
            }
            if let Err(rollback_error) = remove_path(self.metadata_incoming) {
                rollback_errors.push(rollback_error);
            }
            let _ = sync(self.models_root);
            let _ = sync(self.manifest_directory);
            if rollback_errors.is_empty() {
                return Err(error);
            }
            return Err(format!(
                "{error}; transaction rollback also failed: {}",
                rollback_errors.join("; ")
            ));
        }

        remove_path(self.backup)?;
        remove_path(self.metadata_backup)?;
        Ok(())
    }
}

pub fn install_optional_model_pack(
    package_path: &Path,
    runtime_root: &Path,
) -> Result<String, String> {
    if package_path.extension().and_then(|value| value.to_str()) != Some("vtxocrmodel") {
        return Err("Select a VisualTeX .vtxocrmodel package".to_string());
    }
    let package_metadata = fs::symlink_metadata(package_path).map_err(|error| {
        format!(
            "Unable to inspect OCR model package {}: {error}",
            package_path.display()
        )
    })?;
    if !package_metadata.is_file() || package_metadata.file_type().is_symlink() {
        return Err("The OCR model package must be a regular file".to_string());
    }
    if package_metadata.len() == 0 || package_metadata.len() > MAX_OPTIONAL_MODEL_PACK_BYTES {
        return Err("The OCR model package size is invalid".to_string());
    }
    let runtime_python = runtime_root.join("python/bin/python3");
    let legacy_python = runtime_root.join("venv/bin/python");
    if !runtime_python.is_file() && !legacy_python.is_file() {
        return Err(
            "Install the VisualTeX offline OCR runtime before importing a model pack".to_string(),
        );
    }

    let parent = runtime_root
        .parent()
        .ok_or_else(|| "OCR runtime has no parent directory".to_string())?;
    let suffix = Uuid::new_v4();
    let staging = parent.join(format!(".ocr-model-pack-{suffix}"));
    fs::remove_dir_all(&staging).ok();
    extract_archive(package_path, &staging)?;
    let pack_root = staging.join("visualtex-model-pack");
    let manifest = match verify_optional_pack(&pack_root) {
        Ok(manifest) => manifest,
        Err(error) => {
            fs::remove_dir_all(&staging).ok();
            return Err(error);
        }
    };

    let models_root = runtime_root.join("cache/paddlex/official_models");
    fs::create_dir_all(&models_root)
        .map_err(|error| format!("Unable to create OCR model directory: {error}"))?;
    let manifest_directory = runtime_root.join("model-packs");
    fs::create_dir_all(&manifest_directory)
        .map_err(|error| format!("Unable to create model pack metadata directory: {error}"))?;

    let incoming = models_root.join(format!(".{}-installing-{suffix}", manifest.model));
    let target = models_root.join(&manifest.model);
    let backup = models_root.join(format!(".{}-backup-{suffix}", manifest.model));
    let metadata_target = manifest_directory.join(format!("{}.json", manifest.model));
    let metadata_incoming =
        manifest_directory.join(format!(".{}-installing-{suffix}.json", manifest.model));
    let metadata_backup =
        manifest_directory.join(format!(".{}-backup-{suffix}.json", manifest.model));
    let metadata_bytes = serde_json::to_vec_pretty(&manifest)
        .map_err(|error| format!("Unable to serialize model pack metadata: {error}"))?;
    write_synced_file(&metadata_incoming, &metadata_bytes)?;

    let source = pack_root
        .join("paddlex/official_models")
        .join(&manifest.model);
    if let Err(error) = fs::rename(&source, &incoming) {
        fs::remove_file(&metadata_incoming).ok();
        fs::remove_dir_all(&staging).ok();
        return Err(format!("Unable to stage OCR model pack: {error}"));
    }
    let activation = OptionalModelActivation {
        incoming: &incoming,
        target: &target,
        backup: &backup,
        metadata_incoming: &metadata_incoming,
        metadata_target: &metadata_target,
        metadata_backup: &metadata_backup,
        models_root: &models_root,
        manifest_directory: &manifest_directory,
    }
    .activate(sync_directory);
    fs::remove_dir_all(&staging).ok();
    activation?;
    Ok(manifest.model)
}

pub fn remove_optional_model(runtime_root: &Path, model: &str) -> Result<(), String> {
    if !OPTIONAL_MODELS.contains(&model) {
        return Err("Only optional S or L models can be removed".to_string());
    }
    let target = runtime_root
        .join("cache/paddlex/official_models")
        .join(model);
    if target.exists() {
        fs::remove_dir_all(&target)
            .map_err(|error| format!("Unable to remove optional OCR model {model}: {error}"))?;
    }
    let metadata = runtime_root
        .join("model-packs")
        .join(format!("{model}.json"));
    if metadata.exists() {
        fs::remove_file(metadata)
            .map_err(|error| format!("Unable to remove model pack metadata: {error}"))?;
    }
    Ok(())
}

pub fn installed_models(root: &Path) -> Vec<String> {
    let model_root = root.join("cache/paddlex/official_models");
    KNOWN_MODELS
        .iter()
        .filter(|model| {
            let directory = model_root.join(model);
            directory.join("inference.json").is_file()
                && directory.join("inference.pdiparams").is_file()
                && directory.join("inference.yml").is_file()
        })
        .map(|model| (*model).to_string())
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    fn write_archive(path: &Path, entries: &[(&str, &[u8])]) {
        let owned = entries
            .iter()
            .map(|(name, data)| ((*name).to_string(), (*data).to_vec()))
            .collect::<Vec<_>>();
        write_owned_archive(path, &owned);
    }

    fn write_owned_archive(path: &Path, entries: &[(String, Vec<u8>)]) {
        let file = File::create(path).unwrap();
        let encoder = flate2::write::GzEncoder::new(file, flate2::Compression::default());
        let mut builder = tar::Builder::new(encoder);
        for (name, data) in entries {
            let mut header = tar::Header::new_gnu();
            header.set_size(data.len() as u64);
            header.set_mode(0o644);
            header.set_cksum();
            builder
                .append_data(&mut header, name, data.as_slice())
                .unwrap();
        }
        let encoder = builder.into_inner().unwrap();
        encoder.finish().unwrap();
    }

    #[test]
    fn archive_extraction_rejects_path_traversal() {
        let temp = TempDir::new().unwrap();
        let archive = temp.path().join("unsafe.tar.gz");
        let file = File::create(&archive).unwrap();
        let encoder = flate2::write::GzEncoder::new(file, flate2::Compression::default());
        let mut builder = tar::Builder::new(encoder);
        let mut header = tar::Header::new_gnu();
        header.set_size(3);
        header.set_mode(0o644);
        let malicious_path = b"../escape";
        header.as_mut_bytes()[..malicious_path.len()].copy_from_slice(malicious_path);
        header.set_cksum();
        builder.append(&header, &b"bad"[..]).unwrap();
        let encoder = builder.into_inner().unwrap();
        encoder.finish().unwrap();
        let result = extract_archive(&archive, &temp.path().join("output"));
        assert!(result.is_err());
        assert!(!temp.path().join("escape").exists());
    }

    #[test]
    fn archive_extraction_accepts_normal_files() {
        let temp = TempDir::new().unwrap();
        let archive = temp.path().join("safe.tar.gz");
        write_archive(&archive, &[("python/bin/python3", b"python")]);
        let output = temp.path().join("output");
        extract_archive(&archive, &output).unwrap();
        assert_eq!(
            fs::read(output.join("python/bin/python3")).unwrap(),
            b"python"
        );
    }

    #[test]
    fn installed_model_detection_requires_all_files() {
        let temp = TempDir::new().unwrap();
        let model = temp
            .path()
            .join("cache/paddlex/official_models/PP-FormulaNet_plus-M");
        fs::create_dir_all(&model).unwrap();
        fs::write(model.join("inference.json"), b"json").unwrap();
        fs::write(model.join("inference.pdiparams"), b"params").unwrap();
        assert!(installed_models(temp.path()).is_empty());
        fs::write(model.join("inference.yml"), b"yml").unwrap();
        assert_eq!(installed_models(temp.path()), vec![OFFLINE_DEFAULT_MODEL]);
    }

    #[test]
    fn optional_model_pack_rejects_wrong_extension_and_forged_files() {
        let temp = TempDir::new().unwrap();
        let runtime = temp.path().join("ocr-runtime");
        let python = runtime.join("python/bin/python3");
        fs::create_dir_all(python.parent().unwrap()).unwrap();
        fs::write(&python, b"python").unwrap();

        let wrong_extension = temp.path().join("model.tar.gz");
        fs::write(&wrong_extension, b"not a model pack").unwrap();
        assert!(install_optional_model_pack(&wrong_extension, &runtime)
            .unwrap_err()
            .contains(".vtxocrmodel"));

        let model = "PP-FormulaNet_plus-S";
        let records = expected_model_hashes(model)
            .unwrap()
            .into_iter()
            .map(|(name, sha256)| {
                (
                    name.to_string(),
                    OfflineFileRecord {
                        name: name.to_string(),
                        size: 3,
                        sha256: sha256.to_string(),
                    },
                )
            })
            .collect();
        let manifest = OptionalModelPackManifest {
            schema_version: 1,
            platform: "macos".to_string(),
            architecture: "arm64".to_string(),
            model: model.to_string(),
            files: records,
        };
        let root = format!("visualtex-model-pack/paddlex/official_models/{model}");
        let entries = vec![
            (
                "visualtex-model-pack/pack-manifest.json".to_string(),
                serde_json::to_vec_pretty(&manifest).unwrap(),
            ),
            (format!("{root}/inference.json"), b"bad".to_vec()),
            (format!("{root}/inference.pdiparams"), b"bad".to_vec()),
            (format!("{root}/inference.yml"), b"bad".to_vec()),
        ];
        let forged = temp.path().join("forged.vtxocrmodel");
        write_owned_archive(&forged, &entries);
        let error = install_optional_model_pack(&forged, &runtime).unwrap_err();
        assert!(error.contains("verification failed"));
        assert!(!runtime
            .join("cache/paddlex/official_models/PP-FormulaNet_plus-S")
            .exists());
    }

    #[test]
    fn optional_model_activation_rolls_back_model_and_metadata_on_sync_failure() {
        let temp = TempDir::new().unwrap();
        let models_root = temp.path().join("models");
        let manifest_root = temp.path().join("metadata");
        fs::create_dir_all(&models_root).unwrap();
        fs::create_dir_all(&manifest_root).unwrap();

        let target = models_root.join("PP-FormulaNet_plus-S");
        let incoming = models_root.join(".PP-FormulaNet_plus-S-installing");
        let backup = models_root.join(".PP-FormulaNet_plus-S-backup");
        fs::create_dir_all(&target).unwrap();
        fs::write(target.join("marker"), b"old-model").unwrap();
        fs::create_dir_all(&incoming).unwrap();
        fs::write(incoming.join("marker"), b"new-model").unwrap();

        let metadata_target = manifest_root.join("PP-FormulaNet_plus-S.json");
        let metadata_incoming = manifest_root.join(".PP-FormulaNet_plus-S-installing.json");
        let metadata_backup = manifest_root.join(".PP-FormulaNet_plus-S-backup.json");
        fs::write(&metadata_target, b"old-metadata").unwrap();
        fs::write(&metadata_incoming, b"new-metadata").unwrap();

        let mut sync_calls = 0;
        let result = OptionalModelActivation {
            incoming: &incoming,
            target: &target,
            backup: &backup,
            metadata_incoming: &metadata_incoming,
            metadata_target: &metadata_target,
            metadata_backup: &metadata_backup,
            models_root: &models_root,
            manifest_directory: &manifest_root,
        }
        .activate(|_| {
            sync_calls += 1;
            if sync_calls == 1 {
                Err("injected directory sync failure".to_string())
            } else {
                Ok(())
            }
        });

        assert!(result
            .unwrap_err()
            .contains("injected directory sync failure"));
        assert_eq!(fs::read(target.join("marker")).unwrap(), b"old-model");
        assert_eq!(fs::read(&metadata_target).unwrap(), b"old-metadata");
        assert!(!incoming.exists());
        assert!(!backup.exists());
        assert!(!metadata_incoming.exists());
        assert!(!metadata_backup.exists());
    }

    #[test]
    fn optional_model_activation_removes_new_files_when_no_previous_install_exists() {
        let temp = TempDir::new().unwrap();
        let models_root = temp.path().join("models");
        let manifest_root = temp.path().join("metadata");
        fs::create_dir_all(&models_root).unwrap();
        fs::create_dir_all(&manifest_root).unwrap();

        let target = models_root.join("PP-FormulaNet_plus-L");
        let incoming = models_root.join(".PP-FormulaNet_plus-L-installing");
        let backup = models_root.join(".PP-FormulaNet_plus-L-backup");
        fs::create_dir_all(&incoming).unwrap();
        fs::write(incoming.join("marker"), b"new-model").unwrap();

        let metadata_target = manifest_root.join("PP-FormulaNet_plus-L.json");
        let metadata_incoming = manifest_root.join(".PP-FormulaNet_plus-L-installing.json");
        let metadata_backup = manifest_root.join(".PP-FormulaNet_plus-L-backup.json");
        fs::write(&metadata_incoming, b"new-metadata").unwrap();

        let result = OptionalModelActivation {
            incoming: &incoming,
            target: &target,
            backup: &backup,
            metadata_incoming: &metadata_incoming,
            metadata_target: &metadata_target,
            metadata_backup: &metadata_backup,
            models_root: &models_root,
            manifest_directory: &manifest_root,
        }
        .activate(|_| Err("injected directory sync failure".to_string()));

        assert!(result.is_err());
        assert!(!target.exists());
        assert!(!metadata_target.exists());
        assert!(!incoming.exists());
        assert!(!metadata_incoming.exists());
    }

    #[test]
    fn optional_model_removal_never_removes_default_model() {
        let temp = TempDir::new().unwrap();
        let model_root = temp.path().join("cache/paddlex/official_models");
        for model in [OFFLINE_DEFAULT_MODEL, "PP-FormulaNet_plus-S"] {
            let directory = model_root.join(model);
            fs::create_dir_all(&directory).unwrap();
            for file in ["inference.json", "inference.pdiparams", "inference.yml"] {
                fs::write(directory.join(file), b"model").unwrap();
            }
        }
        let metadata = temp.path().join("model-packs/PP-FormulaNet_plus-S.json");
        fs::create_dir_all(metadata.parent().unwrap()).unwrap();
        fs::write(&metadata, b"{}").unwrap();

        assert!(remove_optional_model(temp.path(), OFFLINE_DEFAULT_MODEL).is_err());
        assert!(model_root.join(OFFLINE_DEFAULT_MODEL).is_dir());

        remove_optional_model(temp.path(), "PP-FormulaNet_plus-S").unwrap();
        assert!(!model_root.join("PP-FormulaNet_plus-S").exists());
        assert!(!metadata.exists());
        assert!(model_root.join(OFFLINE_DEFAULT_MODEL).is_dir());
    }
}
