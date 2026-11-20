use serde::{Deserialize, Serialize};
use std::fs::{self, File, OpenOptions};
use std::io::{Read, Write};
use std::path::{Component, Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use uuid::Uuid;

pub const STORAGE_DIRECTORY_NAME: &str = "VisualTeX-OCR";
pub const STORAGE_CONFIG_FILE: &str = "ocr-storage.json";
pub const STORAGE_MARKER_FILE: &str = ".visualtex-ocr-root.json";
pub const RUNTIME_INSTALL_MIN_FREE_BYTES: u64 = 2 * 1024 * 1024 * 1024;

const STORAGE_CONFIG_SCHEMA: u32 = 1;
const STORAGE_MARKER_SCHEMA: u32 = 1;
const STORAGE_PRODUCT: &str = "VisualTeX";
const STORAGE_KIND: &str = "ocr-runtime";

#[derive(Debug, Clone)]
pub struct StorageResolution {
    pub root: PathBuf,
    pub config_path: PathBuf,
    pub source: String,
    pub managed: bool,
}

#[derive(Debug, Clone)]
pub struct StorageChange {
    pub source_root: PathBuf,
    pub target_root: PathBuf,
    pub reset_source: bool,
    pub adopted_existing: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StorageConfig {
    schema_version: u32,
    runtime_root: String,
    updated_at_ms: u128,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StorageMarker {
    schema_version: u32,
    product: String,
    kind: String,
    runtime_id: String,
    created_at_ms: u128,
}

fn now_ms() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default()
}

fn normalized_absolute(path: &Path) -> Result<PathBuf, String> {
    if !path.is_absolute() {
        return Err(format!(
            "OCR storage path must be absolute: {}",
            path.display()
        ));
    }
    let mut normalized = PathBuf::new();
    for component in path.components() {
        match component {
            Component::Prefix(prefix) => normalized.push(prefix.as_os_str()),
            Component::RootDir => normalized.push(component.as_os_str()),
            Component::CurDir => {}
            Component::ParentDir => {
                if !normalized.pop() {
                    return Err(format!(
                        "OCR storage path escapes its filesystem root: {}",
                        path.display()
                    ));
                }
            }
            Component::Normal(value) => normalized.push(value),
        }
    }
    if normalized.parent().is_none() {
        return Err(
            "A drive or filesystem root cannot be used as the OCR storage directory".to_string(),
        );
    }
    Ok(normalized)
}

fn comparison_key(path: &Path) -> String {
    let value = path.to_string_lossy().replace('/', "\\");
    if cfg!(windows) {
        value.trim_end_matches('\\').to_ascii_lowercase()
    } else {
        value.trim_end_matches('\\').to_string()
    }
}

fn same_path(left: &Path, right: &Path) -> bool {
    comparison_key(left) == comparison_key(right)
}

fn is_within(path: &Path, parent: &Path) -> bool {
    let path_key = comparison_key(path);
    let parent_key = comparison_key(parent);
    path_key == parent_key
        || path_key
            .strip_prefix(&parent_key)
            .is_some_and(|suffix| suffix.starts_with('\\'))
}

fn validate_not_protected(path: &Path, protected_roots: &[PathBuf]) -> Result<(), String> {
    for protected in protected_roots {
        let protected = normalized_absolute(protected)?;
        if is_within(path, &protected) || is_within(&protected, path) {
            return Err(format!(
                "OCR storage cannot overlap the VisualTeX program or resource directory: {}",
                protected.display()
            ));
        }
    }
    Ok(())
}

#[cfg(windows)]
fn metadata_is_reparse_point(metadata: &fs::Metadata) -> bool {
    use std::os::windows::fs::MetadataExt;
    const FILE_ATTRIBUTE_REPARSE_POINT: u32 = 0x0400;
    metadata.file_attributes() & FILE_ATTRIBUTE_REPARSE_POINT != 0
}

#[cfg(not(windows))]
fn metadata_is_reparse_point(metadata: &fs::Metadata) -> bool {
    metadata.file_type().is_symlink()
}

fn reject_reparse_root(path: &Path) -> Result<(), String> {
    if let Ok(metadata) = fs::symlink_metadata(path) {
        if metadata.file_type().is_symlink() || metadata_is_reparse_point(&metadata) {
            return Err(format!(
                "OCR storage root cannot be a symbolic link or Windows reparse point: {}",
                path.display()
            ));
        }
        if !metadata.is_dir() {
            return Err(format!(
                "OCR storage path is not a directory: {}",
                path.display()
            ));
        }
    }
    Ok(())
}

fn read_json<T: for<'de> Deserialize<'de>>(path: &Path) -> Result<T, String> {
    let mut file =
        File::open(path).map_err(|error| format!("Unable to open {}: {error}", path.display()))?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)
        .map_err(|error| format!("Unable to read {}: {error}", path.display()))?;
    serde_json::from_slice(&bytes)
        .map_err(|error| format!("Unable to parse {}: {error}", path.display()))
}

fn write_atomic(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("Path has no parent directory: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let temporary = parent.join(format!(
        ".{}-{}.tmp",
        path.file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("visualtex-ocr"),
        Uuid::new_v4()
    ));
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&temporary)
        .map_err(|error| format!("Unable to create {}: {error}", temporary.display()))?;
    file.write_all(bytes)
        .and_then(|_| file.sync_all())
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    drop(file);

    let backup = parent.join(format!(
        ".{}-{}.backup",
        path.file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("visualtex-ocr"),
        Uuid::new_v4()
    ));
    let had_existing = path.exists();
    if had_existing {
        if let Err(error) = fs::rename(path, &backup) {
            let _ = fs::remove_file(&temporary);
            return Err(format!(
                "Unable to preserve the previous {} before replacement: {error}",
                path.display()
            ));
        }
    }

    if let Err(error) = fs::rename(&temporary, path) {
        let restore_error = if had_existing {
            fs::rename(&backup, path).err()
        } else {
            None
        };
        let _ = fs::remove_file(&temporary);
        return Err(match restore_error {
            Some(restore_error) => format!(
                "Unable to activate {}: {error}. Restoring the previous file also failed: {restore_error}",
                path.display()
            ),
            None => format!("Unable to activate {}: {error}", path.display()),
        });
    }
    if had_existing {
        let _ = fs::remove_file(&backup);
    }
    Ok(())
}

fn read_config(config_path: &Path) -> Result<Option<PathBuf>, String> {
    if !config_path.is_file() {
        return Ok(None);
    }
    let config: StorageConfig = read_json(config_path)?;
    if config.schema_version != STORAGE_CONFIG_SCHEMA {
        return Err(format!(
            "Unsupported OCR storage configuration schema {} in {}",
            config.schema_version,
            config_path.display()
        ));
    }
    Ok(Some(normalized_absolute(Path::new(&config.runtime_root))?))
}

fn quarantine_invalid_config(config_path: &Path, error: &str) -> Result<(), String> {
    if !config_path.exists() {
        return Ok(());
    }
    let parent = config_path.parent().ok_or_else(|| {
        format!(
            "OCR storage configuration has no parent: {}",
            config_path.display()
        )
    })?;
    let quarantine = parent.join(format!(
        "ocr-storage.invalid-{}-{}.json",
        now_ms(),
        Uuid::new_v4()
    ));
    fs::rename(config_path, &quarantine).map_err(|rename_error| {
        format!(
            "The OCR storage configuration is invalid ({error}) and could not be quarantined from {} to {}: {rename_error}",
            config_path.display(),
            quarantine.display()
        )
    })
}

pub fn write_config(config_path: &Path, runtime_root: &Path) -> Result<(), String> {
    let runtime_root = normalized_absolute(runtime_root)?;
    let config = StorageConfig {
        schema_version: STORAGE_CONFIG_SCHEMA,
        runtime_root: runtime_root.display().to_string(),
        updated_at_ms: now_ms(),
    };
    let bytes = serde_json::to_vec_pretty(&config)
        .map_err(|error| format!("Unable to serialize OCR storage configuration: {error}"))?;
    write_atomic(config_path, &bytes)
}

fn marker_path(root: &Path) -> PathBuf {
    root.join(STORAGE_MARKER_FILE)
}

pub fn marker_is_valid(root: &Path) -> bool {
    let Ok(marker) = read_json::<StorageMarker>(&marker_path(root)) else {
        return false;
    };
    marker.schema_version == STORAGE_MARKER_SCHEMA
        && marker.product == STORAGE_PRODUCT
        && marker.kind == STORAGE_KIND
        && !marker.runtime_id.trim().is_empty()
}

fn directory_contains_payload(path: &Path) -> bool {
    let Ok(entries) = fs::read_dir(path) else {
        return false;
    };
    entries.filter_map(Result::ok).any(|entry| {
        let Ok(metadata) = fs::symlink_metadata(entry.path()) else {
            return false;
        };
        if metadata.file_type().is_symlink() {
            return false;
        }
        metadata.is_file() || (metadata.is_dir() && directory_contains_payload(&entry.path()))
    })
}

pub fn looks_like_runtime_root(root: &Path) -> bool {
    if marker_is_valid(root)
        || root.join("python").join("python.exe").is_file()
        || root.join("venv").join("Scripts").join("python.exe").is_file()
        || root.join("venv").join("bin").join("python").is_file()
        || root.join("runtime-status.json").is_file()
        || root.join("install-status.json").is_file()
    {
        return true;
    }

    [
        "PP-FormulaNet_plus-S",
        "PP-FormulaNet_plus-M",
        "PP-FormulaNet_plus-L",
    ]
    .iter()
    .any(|model| {
        directory_contains_payload(
            &root
                .join("cache")
                .join("paddlex")
                .join("official_models")
                .join(model),
        )
    })
}

pub fn has_payload(root: &Path) -> bool {
    let Ok(entries) = fs::read_dir(root) else {
        return false;
    };
    entries.filter_map(Result::ok).any(|entry| {
        let name = entry.file_name().to_string_lossy().to_string();
        if name.eq_ignore_ascii_case(STORAGE_MARKER_FILE)
            || name.eq_ignore_ascii_case("logs")
        {
            return false;
        }
        let Ok(metadata) = fs::symlink_metadata(entry.path()) else {
            return false;
        };
        if metadata.file_type().is_symlink() {
            return false;
        }
        metadata.is_file() || (metadata.is_dir() && directory_contains_payload(&entry.path()))
    })
}

pub fn ensure_marker(root: &Path) -> Result<(), String> {
    let root = normalized_absolute(root)?;
    reject_reparse_root(&root)?;
    fs::create_dir_all(&root).map_err(|error| {
        format!(
            "Unable to create OCR storage root {}: {error}",
            root.display()
        )
    })?;
    if marker_path(&root).exists() {
        if marker_is_valid(&root) {
            return Ok(());
        }
        return Err(format!(
            "OCR storage marker is invalid or belongs to another product: {}",
            marker_path(&root).display()
        ));
    }
    if has_payload(&root) && !looks_like_runtime_root(&root) {
        return Err(format!(
            "The selected OCR storage folder contains unrelated files and has no VisualTeX marker: {}",
            root.display()
        ));
    }
    let marker = StorageMarker {
        schema_version: STORAGE_MARKER_SCHEMA,
        product: STORAGE_PRODUCT.to_string(),
        kind: STORAGE_KIND.to_string(),
        runtime_id: Uuid::new_v4().to_string(),
        created_at_ms: now_ms(),
    };
    let bytes = serde_json::to_vec_pretty(&marker)
        .map_err(|error| format!("Unable to serialize OCR storage marker: {error}"))?;
    write_atomic(&marker_path(&root), &bytes)
}

pub fn resolve(
    config_path: &Path,
    legacy_root: &Path,
    default_root: &Path,
    protected_roots: &[PathBuf],
) -> Result<StorageResolution, String> {
    let config_path = normalized_absolute(config_path)?;
    let legacy_root = normalized_absolute(legacy_root)?;
    let default_root = normalized_absolute(default_root)?;

    let configured_root = match read_config(&config_path) {
        Ok(root) => root,
        Err(error) => {
            quarantine_invalid_config(&config_path, &error)?;
            None
        }
    };
    if let Some(root) = configured_root {
        validate_not_protected(&root, protected_roots)?;
        reject_reparse_root(&root)?;
        return Ok(StorageResolution {
            managed: marker_is_valid(&root),
            root,
            config_path,
            source: "configured".to_string(),
        });
    }

    let (root, source) = if legacy_root.exists()
        && (has_payload(&legacy_root) || looks_like_runtime_root(&legacy_root))
    {
        (legacy_root, "legacy")
    } else {
        (default_root, "default")
    };
    validate_not_protected(&root, protected_roots)?;
    reject_reparse_root(&root)?;
    if root.exists() || source == "legacy" {
        ensure_marker(&root)?;
    }
    write_config(&config_path, &root)?;
    Ok(StorageResolution {
        managed: marker_is_valid(&root),
        root,
        config_path,
        source: source.to_string(),
    })
}

fn selected_target(selected_directory: &Path) -> PathBuf {
    let already_root = selected_directory
        .file_name()
        .and_then(|value| value.to_str())
        .is_some_and(|name| {
            name.eq_ignore_ascii_case(STORAGE_DIRECTORY_NAME)
                || name.eq_ignore_ascii_case("ocr-runtime")
        })
        || marker_is_valid(selected_directory)
        || looks_like_runtime_root(selected_directory);
    if already_root {
        selected_directory.to_path_buf()
    } else {
        selected_directory.join(STORAGE_DIRECTORY_NAME)
    }
}

fn nearest_existing_ancestor(path: &Path) -> Option<PathBuf> {
    let mut current = path.to_path_buf();
    loop {
        if current.exists() {
            return Some(current);
        }
        if !current.pop() {
            return None;
        }
    }
}

#[cfg(windows)]
pub fn available_space_bytes(path: &Path) -> Option<u64> {
    use std::os::windows::ffi::OsStrExt;
    type Bool = i32;
    #[link(name = "kernel32")]
    extern "system" {
        fn GetDiskFreeSpaceExW(
            directory_name: *const u16,
            free_bytes_available: *mut u64,
            total_number_of_bytes: *mut u64,
            total_number_of_free_bytes: *mut u64,
        ) -> Bool;
    }

    let existing = nearest_existing_ancestor(path)?;
    let wide = existing
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let mut available = 0_u64;
    let mut total = 0_u64;
    let mut free = 0_u64;
    let success =
        unsafe { GetDiskFreeSpaceExW(wide.as_ptr(), &mut available, &mut total, &mut free) };
    (success != 0).then_some(available)
}

#[cfg(not(windows))]
pub fn available_space_bytes(_path: &Path) -> Option<u64> {
    None
}

pub fn ensure_available_space(
    path: &Path,
    required_bytes: u64,
    purpose: &str,
) -> Result<(), String> {
    let Some(available) = available_space_bytes(path) else {
        return Ok(());
    };
    if available < required_bytes {
        return Err(format!(
            "OCR_STORAGE_NO_SPACE: {purpose} requires at least {:.2} GB free at {}, but only {:.2} GB is available. Choose another OCR storage location or free disk space.",
            required_bytes as f64 / 1024.0 / 1024.0 / 1024.0,
            path.display(),
            available as f64 / 1024.0 / 1024.0 / 1024.0,
        ));
    }
    Ok(())
}

pub fn configure(
    config_path: &Path,
    current_root: &Path,
    selected_directory: &Path,
    protected_roots: &[PathBuf],
) -> Result<StorageChange, String> {
    let _config_path = normalized_absolute(config_path)?;
    let current_root = normalized_absolute(current_root)?;
    let selected_directory = normalized_absolute(selected_directory)?;
    reject_reparse_root(&selected_directory)?;

    let selected_is_explicit_root = selected_directory
        .file_name()
        .and_then(|value| value.to_str())
        .is_some_and(|name| {
            name.eq_ignore_ascii_case(STORAGE_DIRECTORY_NAME)
                || name.eq_ignore_ascii_case("ocr-runtime")
        });
    if !selected_directory.is_dir()
        && !(selected_is_explicit_root
            && selected_directory
                .parent()
                .is_some_and(|parent| parent.is_dir()))
    {
        return Err(format!(
            "Select an existing parent directory, or an OCR root whose parent already exists: {}",
            selected_directory.display()
        ));
    }

    let target_root = normalized_absolute(&selected_target(&selected_directory))?;
    validate_not_protected(&target_root, protected_roots)?;
    reject_reparse_root(&target_root)?;

    let source_has_payload = has_payload(&current_root);
    if same_path(&current_root, &target_root) {
        ensure_marker(&target_root)?;
        return Ok(StorageChange {
            source_root: current_root,
            target_root,
            reset_source: false,
            adopted_existing: source_has_payload,
        });
    }
    if is_within(&target_root, &current_root) || is_within(&current_root, &target_root) {
        return Err(
            "The new OCR storage path cannot be inside the current OCR storage path, or contain it"
                .to_string(),
        );
    }

    let target_has_payload = has_payload(&target_root);
    if target_has_payload {
        if !marker_is_valid(&target_root) && !looks_like_runtime_root(&target_root) {
            return Err(format!(
                "The selected destination contains unrelated files and is not a VisualTeX OCR environment: {}",
                target_root.display()
            ));
        }
        ensure_marker(&target_root)?;
        return Ok(StorageChange {
            source_root: current_root,
            target_root,
            reset_source: source_has_payload,
            adopted_existing: true,
        });
    }

    fs::create_dir_all(&target_root)
        .map_err(|error| format!("Unable to create {}: {error}", target_root.display()))?;
    ensure_marker(&target_root)?;
    Ok(StorageChange {
        source_root: current_root,
        target_root,
        reset_source: source_has_payload,
        adopted_existing: false,
    })
}

pub fn friendly_storage_error(error: String, root: &Path) -> String {
    let lowered = error.to_ascii_lowercase();
    if lowered.contains("no space left on device")
        || lowered.contains("errno 28")
        || lowered.contains("not enough space")
        || lowered.contains("os error 112")
        || lowered.contains("ocr_storage_no_space")
    {
        let available = available_space_bytes(root)
            .map(|bytes| format!("{:.2} GB", bytes as f64 / 1024.0 / 1024.0 / 1024.0))
            .unwrap_or_else(|| "unknown".to_string());
        return format!(
            "OCR 存储位置空间不足。当前路径：{}；可用空间：{}。请在 OCR 设置中更换到空间充足的磁盘，或清理该磁盘后重试。\n\n原始错误：{}",
            root.display(),
            available,
            error
        );
    }
    error
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn legacy_runtime_is_adopted_and_persisted() {
        let temp = TempDir::new().unwrap();
        let config = temp.path().join("config/ocr-storage.json");
        let legacy = temp.path().join("legacy/ocr-runtime");
        let default = temp.path().join("default/ocr-runtime");
        fs::create_dir_all(legacy.join("python")).unwrap();
        fs::write(legacy.join("python/python.exe"), b"python").unwrap();

        let resolved = resolve(&config, &legacy, &default, &[]).unwrap();
        assert_eq!(resolved.root, legacy);
        assert_eq!(resolved.source, "legacy");
        assert!(marker_is_valid(&resolved.root));

        let second = resolve(&config, &legacy, &default, &[]).unwrap();
        assert_eq!(second.root, resolved.root);
        assert_eq!(second.source, "configured");
    }

    #[test]
    fn changing_storage_requires_reset_and_never_copies_the_source() {
        let temp = TempDir::new().unwrap();
        let config = temp.path().join("config/ocr-storage.json");
        let source = temp.path().join("legacy/ocr-runtime");
        let selected = temp.path().join("external");
        fs::create_dir_all(source.join("python/Lib/site-packages")).unwrap();
        fs::create_dir_all(&selected).unwrap();
        fs::write(source.join("python/python.exe"), b"python").unwrap();
        fs::write(
            source.join("python/Lib/site-packages/paddle.pyd"),
            b"paddle",
        )
        .unwrap();
        ensure_marker(&source).unwrap();

        let change = configure(&config, &source, &selected, &[]).unwrap();
        assert_eq!(change.target_root, selected.join(STORAGE_DIRECTORY_NAME));
        assert!(change.reset_source);
        assert!(!change.adopted_existing);
        assert!(!change.target_root.join("python/python.exe").exists());
        assert!(source.join("python/python.exe").is_file());
        assert!(marker_is_valid(&change.target_root));
    }

    #[test]
    fn unrelated_non_empty_destination_is_rejected() {
        let temp = TempDir::new().unwrap();
        let config = temp.path().join("config/ocr-storage.json");
        let source = temp.path().join("source");
        let selected = temp.path().join("selected");
        let target = selected.join(STORAGE_DIRECTORY_NAME);
        fs::create_dir_all(&source).unwrap();
        fs::create_dir_all(&target).unwrap();
        fs::write(target.join("personal.txt"), b"do not touch").unwrap();

        let error = configure(&config, &source, &selected, &[]).unwrap_err();
        assert!(error.contains("unrelated files"));
        assert!(target.join("personal.txt").is_file());
    }

    #[test]
    fn ordinary_parent_with_python_folder_is_not_an_ocr_root() {
        let temp = TempDir::new().unwrap();
        let config = temp.path().join("config/ocr-storage.json");
        let source = temp.path().join("source");
        let selected = temp.path().join("Roaming");
        fs::create_dir_all(selected.join("Python")).unwrap();
        fs::create_dir_all(&source).unwrap();

        assert!(!looks_like_runtime_root(&selected));
        let change = configure(&config, &source, &selected, &[]).unwrap();
        assert_eq!(change.target_root, selected.join(STORAGE_DIRECTORY_NAME));
    }

    #[test]
    fn empty_managed_cache_does_not_count_as_ocr_data() {
        let temp = TempDir::new().unwrap();
        let config = temp.path().join("config/ocr-storage.json");
        let source = temp.path().join("source");
        let selected = temp.path().join("selected");
        fs::create_dir_all(source.join("cache/paddlex/official_models")).unwrap();
        fs::create_dir_all(&selected).unwrap();
        ensure_marker(&source).unwrap();

        assert!(!has_payload(&source));
        let change = configure(&config, &source, &selected, &[]).unwrap();
        assert!(!change.reset_source);
    }

    #[test]
    fn managed_existing_destination_can_replace_current_storage() {
        let temp = TempDir::new().unwrap();
        let config = temp.path().join("config/ocr-storage.json");
        let source = temp.path().join("source");
        let selected = temp.path().join("selected");
        let target = selected.join(STORAGE_DIRECTORY_NAME);
        fs::create_dir_all(source.join("python")).unwrap();
        fs::write(source.join("python/python.exe"), b"source-python").unwrap();
        fs::create_dir_all(&target).unwrap();
        ensure_marker(&target).unwrap();
        fs::write(target.join("runtime-status.json"), b"{}").unwrap();

        let change = configure(&config, &source, &selected, &[]).unwrap();
        assert!(change.reset_source);
        assert!(change.adopted_existing);
        assert_eq!(change.target_root, target);
    }

    #[test]
    fn protected_program_directory_is_rejected() {
        let temp = TempDir::new().unwrap();
        let config = temp.path().join("config/ocr-storage.json");
        let source = temp.path().join("source");
        let program = temp.path().join("VisualTeX");
        fs::create_dir_all(&source).unwrap();
        fs::create_dir_all(&program).unwrap();

        let error =
            configure(&config, &source, &program, std::slice::from_ref(&program)).unwrap_err();
        assert!(error.contains("cannot overlap"));
    }
}
