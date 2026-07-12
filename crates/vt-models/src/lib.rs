use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::io::{Read, Write};
use std::path::{Component, Path, PathBuf};

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use tempfile::TempDir;
use walkdir::WalkDir;

const MANIFEST_FILE: &str = "visualtex-model.json";
const ACTIVE_MODELS_FILE: &str = "active-models.json";
const MAX_FILES: usize = 50_000;
const MAX_PACKAGE_BYTES: u64 = 24 * 1024 * 1024 * 1024;
const COPY_BUFFER_BYTES: usize = 1024 * 1024;

#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ModelKind {
    FormulaOcr,
    LayoutOcr,
    TextOcr,
    TableOcr,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ModelFileDigest {
    pub path: PathBuf,
    pub sha256: String,
    pub byte_len: u64,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ModelPackageManifest {
    pub schema_version: u32,
    pub id: String,
    pub kind: ModelKind,
    pub version: String,
    pub backend: String,
    pub entrypoint: PathBuf,
    #[serde(default)]
    pub files: Vec<ModelFileDigest>,
    #[serde(default)]
    pub metadata: BTreeMap<String, String>,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct InstalledModelPackage {
    pub manifest: ModelPackageManifest,
    pub install_path: PathBuf,
    pub installed_sha256: String,
    pub total_bytes: u64,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ModelPackageInspection {
    pub manifest: ModelPackageManifest,
    pub source_path: PathBuf,
    pub computed_files: Vec<ModelFileDigest>,
    pub package_sha256: String,
    pub total_bytes: u64,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ModelSelection {
    pub id: String,
    pub version: String,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ActiveModelRegistry {
    schema_version: u32,
    active: BTreeMap<ModelKind, ModelSelection>,
}

impl Default for ActiveModelRegistry {
    fn default() -> Self {
        Self {
            schema_version: 1,
            active: BTreeMap::new(),
        }
    }
}

#[derive(Debug, thiserror::Error)]
pub enum ModelPackageError {
    #[error("model package directory does not exist: {0}")]
    MissingSource(PathBuf),
    #[error("model package manifest is missing: {0}")]
    MissingManifest(PathBuf),
    #[error("model package manifest schema {0} is unsupported")]
    UnsupportedSchema(u32),
    #[error("invalid model package identifier: {0:?}")]
    InvalidId(String),
    #[error("invalid model package version: {0:?}")]
    InvalidVersion(String),
    #[error("invalid model package path: {0}")]
    InvalidPath(PathBuf),
    #[error("invalid entrypoint for {kind:?}: {path}")]
    InvalidEntrypoint { kind: ModelKind, path: PathBuf },
    #[error("symbolic links are not allowed in model packages: {0}")]
    Symlink(PathBuf),
    #[error("model package contains too many files")]
    TooManyFiles,
    #[error("model package exceeds the size limit")]
    TooLarge,
    #[error("manifest references a missing file: {0}")]
    MissingFile(PathBuf),
    #[error("manifest contains duplicate file path: {0}")]
    DuplicateFile(PathBuf),
    #[error("checksum mismatch for {path}: expected {expected}, got {actual}")]
    ChecksumMismatch {
        path: PathBuf,
        expected: String,
        actual: String,
    },
    #[error("file length mismatch for {path}: expected {expected}, got {actual}")]
    LengthMismatch {
        path: PathBuf,
        expected: u64,
        actual: u64,
    },
    #[error("model package is already installed: {0}")]
    AlreadyInstalled(PathBuf),
    #[error("installed package is missing its manifest: {0}")]
    BrokenInstall(PathBuf),
    #[error("model is not installed: {kind:?} {id}@{version}")]
    ModelNotInstalled {
        kind: ModelKind,
        id: String,
        version: String,
    },
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Json(#[from] serde_json::Error),
    #[error(transparent)]
    WalkDir(#[from] walkdir::Error),
}

pub fn inspect_package(
    source: impl AsRef<Path>,
) -> Result<ModelPackageInspection, ModelPackageError> {
    let source = source.as_ref();
    if !source.is_dir() {
        return Err(ModelPackageError::MissingSource(source.to_path_buf()));
    }
    let source = source.canonicalize()?;
    let manifest_path = source.join(MANIFEST_FILE);
    if !manifest_path.is_file() {
        return Err(ModelPackageError::MissingManifest(manifest_path));
    }
    let manifest: ModelPackageManifest = serde_json::from_slice(&fs::read(&manifest_path)?)?;
    validate_manifest(&manifest)?;

    let computed_files = compute_package_files(&source)?;
    let computed_by_path = computed_files
        .iter()
        .map(|file| (file.path.clone(), file))
        .collect::<BTreeMap<_, _>>();

    if !manifest.files.is_empty() {
        let mut declared = BTreeSet::new();
        for expected in &manifest.files {
            validate_relative_path(&expected.path)?;
            if !declared.insert(expected.path.clone()) {
                return Err(ModelPackageError::DuplicateFile(expected.path.clone()));
            }
            let actual = computed_by_path
                .get(&expected.path)
                .ok_or_else(|| ModelPackageError::MissingFile(expected.path.clone()))?;
            if !expected.sha256.eq_ignore_ascii_case(&actual.sha256) {
                return Err(ModelPackageError::ChecksumMismatch {
                    path: expected.path.clone(),
                    expected: expected.sha256.clone(),
                    actual: actual.sha256.clone(),
                });
            }
            if expected.byte_len != actual.byte_len {
                return Err(ModelPackageError::LengthMismatch {
                    path: expected.path.clone(),
                    expected: expected.byte_len,
                    actual: actual.byte_len,
                });
            }
        }
    }

    let entrypoint = source.join(&manifest.entrypoint);
    if !entrypoint.is_file() && !entrypoint.is_dir() {
        return Err(ModelPackageError::MissingFile(manifest.entrypoint.clone()));
    }
    let total_bytes = computed_files.iter().map(|file| file.byte_len).sum();
    let package_sha256 = package_digest(&computed_files);
    Ok(ModelPackageInspection {
        manifest,
        source_path: source,
        computed_files,
        package_sha256,
        total_bytes,
    })
}

pub fn install_package(
    source: impl AsRef<Path>,
    install_root: impl AsRef<Path>,
) -> Result<InstalledModelPackage, ModelPackageError> {
    let inspection = inspect_package(source)?;
    let install_root = install_root.as_ref();
    fs::create_dir_all(install_root)?;
    let install_root = install_root.canonicalize()?;
    let destination = install_root
        .join(&inspection.manifest.id)
        .join(&inspection.manifest.version);
    if destination.exists() {
        return Err(ModelPackageError::AlreadyInstalled(destination));
    }

    let staging_parent = install_root.join(".staging");
    fs::create_dir_all(&staging_parent)?;
    let staging = TempDir::new_in(&staging_parent)?;
    let staged_package = staging.path().join("package");
    fs::create_dir(&staged_package)?;
    copy_verified_package(&inspection, &staged_package)?;

    let staged_inspection = inspect_package(&staged_package)?;
    if staged_inspection.package_sha256 != inspection.package_sha256 {
        return Err(ModelPackageError::ChecksumMismatch {
            path: PathBuf::from("<package>"),
            expected: inspection.package_sha256,
            actual: staged_inspection.package_sha256,
        });
    }
    if let Some(parent) = destination.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::rename(&staged_package, &destination)?;
    sync_directory(destination.parent().unwrap_or(&install_root))?;

    let installed = InstalledModelPackage {
        manifest: staged_inspection.manifest,
        install_path: destination,
        installed_sha256: staged_inspection.package_sha256,
        total_bytes: staged_inspection.total_bytes,
    };
    if active_model(&install_root, installed.manifest.kind)?.is_none() {
        set_active_model(
            &install_root,
            installed.manifest.kind,
            &installed.manifest.id,
            &installed.manifest.version,
        )?;
    }
    Ok(installed)
}

pub fn list_installed(
    install_root: impl AsRef<Path>,
) -> Result<Vec<InstalledModelPackage>, ModelPackageError> {
    let root = install_root.as_ref();
    if !root.exists() {
        return Ok(Vec::new());
    }
    if !root.is_dir() {
        return Err(ModelPackageError::InvalidPath(root.to_path_buf()));
    }
    let root = root.canonicalize()?;
    let mut packages = Vec::new();
    for id_entry in fs::read_dir(root)? {
        let id_entry = id_entry?;
        if !id_entry.file_type()?.is_dir() || id_entry.file_name() == ".staging" {
            continue;
        }
        for version_entry in fs::read_dir(id_entry.path())? {
            let version_entry = version_entry?;
            if !version_entry.file_type()?.is_dir() {
                continue;
            }
            let path = version_entry.path();
            if !path.join(MANIFEST_FILE).is_file() {
                return Err(ModelPackageError::BrokenInstall(path));
            }
            let inspection = inspect_package(&path)?;
            packages.push(InstalledModelPackage {
                manifest: inspection.manifest,
                install_path: path,
                installed_sha256: inspection.package_sha256,
                total_bytes: inspection.total_bytes,
            });
        }
    }
    packages.sort_by(|left, right| {
        left.manifest
            .id
            .cmp(&right.manifest.id)
            .then_with(|| left.manifest.version.cmp(&right.manifest.version))
    });
    Ok(packages)
}

pub fn active_model(
    install_root: impl AsRef<Path>,
    kind: ModelKind,
) -> Result<Option<InstalledModelPackage>, ModelPackageError> {
    let root = install_root.as_ref();
    if !root.exists() {
        return Ok(None);
    }
    let registry = load_active_registry(root)?;
    let Some(selection) = registry.active.get(&kind) else {
        return Ok(None);
    };
    Ok(list_installed(root)?.into_iter().find(|package| {
        package.manifest.kind == kind
            && package.manifest.id == selection.id
            && package.manifest.version == selection.version
    }))
}

pub fn set_active_model(
    install_root: impl AsRef<Path>,
    kind: ModelKind,
    id: &str,
    version: &str,
) -> Result<InstalledModelPackage, ModelPackageError> {
    let root = install_root.as_ref();
    let installed = list_installed(root)?
        .into_iter()
        .find(|package| {
            package.manifest.kind == kind
                && package.manifest.id == id
                && package.manifest.version == version
        })
        .ok_or_else(|| ModelPackageError::ModelNotInstalled {
            kind,
            id: id.to_owned(),
            version: version.to_owned(),
        })?;
    let mut registry = load_active_registry(root)?;
    registry.active.insert(
        kind,
        ModelSelection {
            id: id.to_owned(),
            version: version.to_owned(),
        },
    );
    save_active_registry(root, &registry)?;
    Ok(installed)
}

pub fn clear_active_model(
    install_root: impl AsRef<Path>,
    kind: ModelKind,
) -> Result<bool, ModelPackageError> {
    let root = install_root.as_ref();
    if !root.exists() {
        return Ok(false);
    }
    let mut registry = load_active_registry(root)?;
    let removed = registry.active.remove(&kind).is_some();
    if removed {
        save_active_registry(root, &registry)?;
    }
    Ok(removed)
}

pub fn remove_installed(
    install_root: impl AsRef<Path>,
    id: &str,
    version: &str,
) -> Result<bool, ModelPackageError> {
    validate_component(id).map_err(|_| ModelPackageError::InvalidId(id.to_owned()))?;
    validate_component(version)
        .map_err(|_| ModelPackageError::InvalidVersion(version.to_owned()))?;
    let root = install_root.as_ref();
    if !root.exists() {
        return Ok(false);
    }
    let root = root.canonicalize()?;
    let target = root.join(id).join(version);
    if !target.exists() {
        return Ok(false);
    }
    let canonical = target.canonicalize()?;
    if !canonical.starts_with(&root) {
        return Err(ModelPackageError::InvalidPath(target));
    }
    let active_kinds = list_installed(&root)?
        .into_iter()
        .filter(|package| package.manifest.id == id && package.manifest.version == version)
        .map(|package| package.manifest.kind)
        .collect::<Vec<_>>();
    fs::remove_dir_all(&canonical)?;
    if let Some(parent) = canonical.parent()
        && parent.read_dir()?.next().is_none()
    {
        fs::remove_dir(parent)?;
    }
    for kind in active_kinds {
        if active_model(&root, kind)?.is_none() {
            clear_active_model(&root, kind)?;
        }
    }
    sync_directory(&root)?;
    Ok(true)
}

fn load_active_registry(root: &Path) -> Result<ActiveModelRegistry, ModelPackageError> {
    let path = root.join(ACTIVE_MODELS_FILE);
    if !path.exists() {
        return Ok(ActiveModelRegistry::default());
    }
    let registry: ActiveModelRegistry = serde_json::from_slice(&fs::read(path)?)?;
    if registry.schema_version != 1 {
        return Err(ModelPackageError::UnsupportedSchema(
            registry.schema_version,
        ));
    }
    Ok(registry)
}

fn save_active_registry(
    root: &Path,
    registry: &ActiveModelRegistry,
) -> Result<(), ModelPackageError> {
    fs::create_dir_all(root)?;
    let root = root.canonicalize()?;
    let staging_parent = root.join(".staging");
    fs::create_dir_all(&staging_parent)?;
    let staging = TempDir::new_in(&staging_parent)?;
    let staged_file = staging.path().join(ACTIVE_MODELS_FILE);
    let bytes = serde_json::to_vec_pretty(registry)?;
    let mut file = fs::OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&staged_file)?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.sync_all()?;

    let destination = root.join(ACTIVE_MODELS_FILE);
    let backup = root.join(format!(
        ".{ACTIVE_MODELS_FILE}.{}.backup",
        std::process::id()
    ));
    if backup.exists() {
        fs::remove_file(&backup)?;
    }
    if destination.exists() {
        fs::rename(&destination, &backup)?;
        if let Err(error) = fs::rename(&staged_file, &destination) {
            let _ = fs::rename(&backup, &destination);
            return Err(error.into());
        }
        fs::remove_file(&backup)?;
    } else {
        fs::rename(&staged_file, &destination)?;
    }
    sync_directory(&root)?;
    Ok(())
}

fn validate_manifest(manifest: &ModelPackageManifest) -> Result<(), ModelPackageError> {
    if manifest.schema_version != 1 {
        return Err(ModelPackageError::UnsupportedSchema(
            manifest.schema_version,
        ));
    }
    validate_component(&manifest.id)
        .map_err(|_| ModelPackageError::InvalidId(manifest.id.clone()))?;
    validate_component(&manifest.version)
        .map_err(|_| ModelPackageError::InvalidVersion(manifest.version.clone()))?;
    if manifest.backend.trim().is_empty() {
        return Err(ModelPackageError::InvalidId(manifest.backend.clone()));
    }
    validate_relative_path(&manifest.entrypoint)?;
    if manifest.kind == ModelKind::LayoutOcr
        && !matches!(
            manifest
                .entrypoint
                .extension()
                .and_then(|value| value.to_str()),
            Some("yaml" | "yml")
        )
    {
        return Err(ModelPackageError::InvalidEntrypoint {
            kind: manifest.kind,
            path: manifest.entrypoint.clone(),
        });
    }
    Ok(())
}

fn validate_component(value: &str) -> Result<(), ()> {
    if value.is_empty()
        || value.len() > 128
        || value.starts_with('.')
        || value.chars().any(|character| {
            !(character.is_ascii_alphanumeric() || matches!(character, '-' | '_' | '.'))
        })
    {
        return Err(());
    }
    Ok(())
}

fn validate_relative_path(path: &Path) -> Result<(), ModelPackageError> {
    if path.as_os_str().is_empty()
        || path.is_absolute()
        || path.components().any(|component| {
            matches!(
                component,
                Component::ParentDir | Component::RootDir | Component::Prefix(_)
            )
        })
    {
        return Err(ModelPackageError::InvalidPath(path.to_path_buf()));
    }
    Ok(())
}

fn compute_package_files(root: &Path) -> Result<Vec<ModelFileDigest>, ModelPackageError> {
    let mut files = Vec::new();
    let mut total_bytes = 0_u64;
    for entry in WalkDir::new(root).follow_links(false) {
        let entry = entry?;
        if entry.path() == root {
            continue;
        }
        let relative = entry
            .path()
            .strip_prefix(root)
            .map_err(|_| ModelPackageError::InvalidPath(entry.path().to_path_buf()))?
            .to_path_buf();
        validate_relative_path(&relative)?;
        if entry.file_type().is_symlink() {
            return Err(ModelPackageError::Symlink(relative));
        }
        if !entry.file_type().is_file() {
            continue;
        }
        if files.len() >= MAX_FILES {
            return Err(ModelPackageError::TooManyFiles);
        }
        let metadata = fs::metadata(entry.path())?;
        total_bytes = total_bytes.saturating_add(metadata.len());
        if total_bytes > MAX_PACKAGE_BYTES {
            return Err(ModelPackageError::TooLarge);
        }
        files.push(ModelFileDigest {
            path: relative,
            sha256: sha256_file(entry.path())?,
            byte_len: metadata.len(),
        });
    }
    files.sort_by(|left, right| left.path.cmp(&right.path));
    Ok(files)
}

fn copy_verified_package(
    inspection: &ModelPackageInspection,
    destination: &Path,
) -> Result<(), ModelPackageError> {
    for expected in &inspection.computed_files {
        let source = inspection.source_path.join(&expected.path);
        let target = destination.join(&expected.path);
        if let Some(parent) = target.parent() {
            fs::create_dir_all(parent)?;
        }
        let mut input = fs::File::open(&source)?;
        let mut output = fs::OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&target)?;
        let mut hasher = Sha256::new();
        let mut copied = 0_u64;
        let mut buffer = vec![0_u8; COPY_BUFFER_BYTES];
        loop {
            let read = input.read(&mut buffer)?;
            if read == 0 {
                break;
            }
            output.write_all(&buffer[..read])?;
            hasher.update(&buffer[..read]);
            copied += read as u64;
        }
        output.sync_all()?;
        let actual = format!("{:x}", hasher.finalize());
        if actual != expected.sha256 || copied != expected.byte_len {
            return Err(ModelPackageError::ChecksumMismatch {
                path: expected.path.clone(),
                expected: expected.sha256.clone(),
                actual,
            });
        }
    }
    sync_directory(destination)?;
    Ok(())
}

fn sha256_file(path: &Path) -> Result<String, ModelPackageError> {
    let mut file = fs::File::open(path)?;
    let mut hasher = Sha256::new();
    let mut buffer = vec![0_u8; COPY_BUFFER_BYTES];
    loop {
        let read = file.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        hasher.update(&buffer[..read]);
    }
    Ok(format!("{:x}", hasher.finalize()))
}

fn package_digest(files: &[ModelFileDigest]) -> String {
    let mut hasher = Sha256::new();
    for file in files {
        hasher.update(file.path.to_string_lossy().as_bytes());
        hasher.update([0]);
        hasher.update(file.byte_len.to_le_bytes());
        hasher.update([0]);
        hasher.update(file.sha256.as_bytes());
        hasher.update([b'\n']);
    }
    format!("{:x}", hasher.finalize())
}

fn sync_directory(_path: &Path) -> Result<(), std::io::Error> {
    #[cfg(unix)]
    {
        fs::File::open(_path)?.sync_all()?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    fn create_package(root: &Path, id: &str, version: &str) {
        fs::create_dir_all(root.join("inference")).unwrap();
        fs::write(root.join("inference/model.pdmodel"), b"model").unwrap();
        fs::write(root.join("inference/model.pdiparams"), b"parameters").unwrap();
        let manifest = ModelPackageManifest {
            schema_version: 1,
            id: id.to_owned(),
            kind: ModelKind::FormulaOcr,
            version: version.to_owned(),
            backend: "paddleocr-formula".to_owned(),
            entrypoint: PathBuf::from("inference"),
            files: Vec::new(),
            metadata: BTreeMap::from([("architecture".into(), "PP-FormulaNet".into())]),
        };
        fs::write(
            root.join(MANIFEST_FILE),
            serde_json::to_vec_pretty(&manifest).unwrap(),
        )
        .unwrap();
    }

    #[test]
    fn installs_lists_and_removes_verified_directory_package() {
        let temp = tempdir().unwrap();
        let source = temp.path().join("source");
        let installed = temp.path().join("installed");
        create_package(&source, "formula-local", "1.0.0");
        let inspection = inspect_package(&source).unwrap();
        assert_eq!(inspection.manifest.kind, ModelKind::FormulaOcr);
        assert_eq!(inspection.computed_files.len(), 3);
        assert!(inspection.total_bytes > 0);

        let package = install_package(&source, &installed).unwrap();
        assert_eq!(
            package.install_path,
            installed
                .join("formula-local/1.0.0")
                .canonicalize()
                .unwrap()
        );
        assert_eq!(list_installed(&installed).unwrap(), vec![package.clone()]);
        assert_eq!(
            active_model(&installed, ModelKind::FormulaOcr).unwrap(),
            Some(package.clone())
        );
        assert!(matches!(
            install_package(&source, &installed),
            Err(ModelPackageError::AlreadyInstalled(_))
        ));

        let source_v2 = temp.path().join("source-v2");
        create_package(&source_v2, "formula-local", "2.0.0");
        let package_v2 = install_package(&source_v2, &installed).unwrap();
        assert_eq!(
            active_model(&installed, ModelKind::FormulaOcr).unwrap(),
            Some(package.clone()),
            "installing another version must not silently switch the active model"
        );
        assert_eq!(
            set_active_model(&installed, ModelKind::FormulaOcr, "formula-local", "2.0.0").unwrap(),
            package_v2
        );
        assert_eq!(
            active_model(&installed, ModelKind::FormulaOcr)
                .unwrap()
                .unwrap()
                .manifest
                .version,
            "2.0.0"
        );

        assert!(remove_installed(&installed, "formula-local", "2.0.0").unwrap());
        assert_eq!(
            active_model(&installed, ModelKind::FormulaOcr).unwrap(),
            None
        );
        assert!(remove_installed(&installed, "formula-local", "1.0.0").unwrap());
        assert!(list_installed(&installed).unwrap().is_empty());
    }

    #[test]
    fn rejects_checksum_mismatch_and_path_escape() {
        let temp = tempdir().unwrap();
        let source = temp.path().join("source");
        create_package(&source, "formula-local", "1.0.0");
        let mut manifest: ModelPackageManifest =
            serde_json::from_slice(&fs::read(source.join(MANIFEST_FILE)).unwrap()).unwrap();
        manifest.files.push(ModelFileDigest {
            path: PathBuf::from("inference/model.pdmodel"),
            sha256: "00".repeat(32),
            byte_len: 5,
        });
        fs::write(
            source.join(MANIFEST_FILE),
            serde_json::to_vec_pretty(&manifest).unwrap(),
        )
        .unwrap();
        assert!(matches!(
            inspect_package(&source),
            Err(ModelPackageError::ChecksumMismatch { .. })
        ));

        manifest.files.clear();
        manifest.entrypoint = PathBuf::from("../outside");
        fs::write(
            source.join(MANIFEST_FILE),
            serde_json::to_vec_pretty(&manifest).unwrap(),
        )
        .unwrap();
        assert!(matches!(
            inspect_package(&source),
            Err(ModelPackageError::InvalidPath(_))
        ));
    }

    #[cfg(unix)]
    #[test]
    fn rejects_symlinks() {
        use std::os::unix::fs::symlink;

        let temp = tempdir().unwrap();
        let source = temp.path().join("source");
        create_package(&source, "formula-local", "1.0.0");
        symlink("model.pdmodel", source.join("inference/link")).unwrap();
        assert!(matches!(
            inspect_package(&source),
            Err(ModelPackageError::Symlink(_))
        ));
    }
}
