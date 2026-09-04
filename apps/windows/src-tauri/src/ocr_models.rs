use flate2::read::GzDecoder;
use reqwest::{Client, Response};
use reqwest::header::{CONTENT_LENGTH, CONTENT_RANGE, RANGE};
use reqwest::StatusCode;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, BTreeSet};
use std::fs::{self, File, OpenOptions};
use std::io::{Read, Write};
use std::path::{Component, Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use tokio::sync::Notify;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tar::{Archive, EntryType};
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Manager};
use uuid::Uuid;

pub const KNOWN_MODELS: &[&str] = &[
    "PP-FormulaNet_plus-S",
    "PP-FormulaNet_plus-M",
    "PP-FormulaNet_plus-L",
];

const MODEL_CATALOG_RESOURCE: &str = "ocr-models/windows-x64/catalog.json";
const MAX_MODEL_PACK_BYTES: u64 = 1024 * 1024 * 1024;
const MAX_ARCHIVE_ENTRIES: usize = 64;
const MAX_ARCHIVE_UNPACKED_BYTES: u64 = 1024 * 1024 * 1024;
const DOWNLOAD_PROGRESS_INTERVAL: Duration = Duration::from_millis(500);
const DOWNLOAD_BUFFER_BYTES: usize = 1024 * 1024;
const DOWNLOAD_RETRIES: usize = 3;

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
pub struct ModelFileRecord {
    pub name: String,
    pub size: u64,
    pub sha256: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ModelPackManifest {
    pub schema_version: u32,
    pub platform: String,
    pub architecture: String,
    pub model: String,
    pub files: BTreeMap<String, ModelFileRecord>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ModelCatalogEntry {
    pub model: String,
    pub url: String,
    pub size: u64,
    pub sha256: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ModelCatalog {
    pub schema_version: u32,
    pub platform: String,
    pub architecture: String,
    pub entries: Vec<ModelCatalogEntry>,
}

#[derive(Debug, Clone, Default)]
pub struct ModelInventory {
    pub installed: Vec<String>,
    pub damaged: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum ModelDownloadState {
    Idle,
    Downloading,
    Verifying,
    Installing,
    Complete,
    Cancelled,
    Failed,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ModelDownloadSnapshot {
    pub model: String,
    pub state: ModelDownloadState,
    pub downloaded_bytes: u64,
    pub total_bytes: u64,
    pub percent: u8,
    pub speed_bytes_per_second: u64,
    pub eta_seconds: Option<u64>,
    pub message: String,
    pub error: Option<String>,
}

impl ModelDownloadSnapshot {
    pub(crate) fn new(model: &str, total_bytes: u64) -> Self {
        Self {
            model: model.to_string(),
            state: ModelDownloadState::Idle,
            downloaded_bytes: 0,
            total_bytes,
            percent: 0,
            speed_bytes_per_second: 0,
            eta_seconds: None,
            message: "OCR model download has not started".to_string(),
            error: None,
        }
    }
}

#[derive(Default)]
pub struct ModelDownloadControl {
    running: AtomicBool,
    cancel_generation: AtomicU64,
    cancel_notify: Notify,
    snapshot: Mutex<Option<ModelDownloadSnapshot>>,
}

impl ModelDownloadControl {
    pub fn begin(self: &Arc<Self>) -> Result<ModelDownloadLease, String> {
        self.running
            .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
            .map_err(|_| "An OCR model download is already running".to_string())?;
        let generation = self.cancel_generation.load(Ordering::SeqCst);
        Ok(ModelDownloadLease {
            control: self.clone(),
            generation,
        })
    }

    pub fn is_running(&self) -> bool {
        self.running.load(Ordering::SeqCst)
    }

    pub fn cancel(&self) -> bool {
        let running = self.running.load(Ordering::SeqCst);
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        self.cancel_notify.notify_waiters();
        self.cancel_notify.notify_one();
        if running {
            if let Ok(mut guard) = self.snapshot.lock() {
                if let Some(snapshot) = guard.as_mut() {
                    snapshot.state = ModelDownloadState::Cancelled;
                    snapshot.speed_bytes_per_second = 0;
                    snapshot.eta_seconds = None;
                    snapshot.message =
                        "OCR 模型下载已立即取消，.part 文件已保留以便续传".to_string();
                    snapshot.error = None;
                }
            }
        }
        running
    }

    pub fn is_cancelled(&self, generation: u64) -> bool {
        self.cancel_generation.load(Ordering::SeqCst) != generation
    }

    pub async fn cancelled(&self, generation: u64) {
        loop {
            if self.is_cancelled(generation) {
                return;
            }
            self.cancel_notify.notified().await;
        }
    }

    pub fn set_snapshot(&self, snapshot: ModelDownloadSnapshot) -> Result<(), String> {
        let mut guard = self
            .snapshot
            .lock()
            .map_err(|_| "OCR model download status lock is unavailable".to_string())?;
        *guard = Some(snapshot);
        Ok(())
    }

    pub fn snapshot(&self) -> Result<Option<ModelDownloadSnapshot>, String> {
        self.snapshot
            .lock()
            .map(|snapshot| snapshot.clone())
            .map_err(|_| "OCR model download status lock is unavailable".to_string())
    }
}

pub struct ModelDownloadLease {
    control: Arc<ModelDownloadControl>,
    generation: u64,
}

impl ModelDownloadLease {
    pub fn generation(&self) -> u64 {
        self.generation
    }

    pub fn control(&self) -> &Arc<ModelDownloadControl> {
        &self.control
    }
}

impl Drop for ModelDownloadLease {
    fn drop(&mut self) {
        self.control.running.store(false, Ordering::SeqCst);
    }
}

fn now_ms() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default()
}

fn cancellation_error() -> String {
    "OCR model download was cancelled".to_string()
}

fn ensure_not_cancelled(
    cancellation: Option<(&ModelDownloadControl, u64)>,
) -> Result<(), String> {
    if cancellation.is_some_and(|(control, generation)| control.is_cancelled(generation)) {
        return Err(cancellation_error());
    }
    Ok(())
}

fn sha256_file_with_cancel(
    path: &Path,
    cancellation: Option<(&ModelDownloadControl, u64)>,
) -> Result<String, String> {
    let mut file = File::open(path)
        .map_err(|error| format!("Unable to open {} for SHA-256 verification: {error}", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = vec![0_u8; DOWNLOAD_BUFFER_BYTES];
    loop {
        ensure_not_cancelled(cancellation)?;
        let count = file
            .read(&mut buffer)
            .map_err(|error| format!("Unable to read {} for SHA-256 verification: {error}", path.display()))?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    ensure_not_cancelled(cancellation)?;
    Ok(hex::encode(digest.finalize()))
}

fn sha256_file(path: &Path) -> Result<String, String> {
    sha256_file_with_cancel(path, None)
}

fn validate_archive_path(path: &Path) -> Result<(), String> {
    if path.as_os_str().is_empty() || path.is_absolute() {
        return Err(format!("Unsafe OCR model archive path: {}", path.display()));
    }
    for component in path.components() {
        match component {
            Component::Normal(_) | Component::CurDir => {}
            Component::ParentDir | Component::RootDir | Component::Prefix(_) => {
                return Err(format!("Unsafe OCR model archive path: {}", path.display()));
            }
        }
    }
    Ok(())
}

fn extract_archive_with_cancel(
    archive_path: &Path,
    destination: &Path,
    cancellation: Option<(&ModelDownloadControl, u64)>,
) -> Result<(), String> {
    fs::create_dir_all(destination)
        .map_err(|error| format!("Unable to create OCR model staging directory: {error}"))?;
    let file = File::open(archive_path)
        .map_err(|error| format!("Unable to open OCR model package {}: {error}", archive_path.display()))?;
    let decoder = GzDecoder::new(file);
    let mut archive = Archive::new(decoder);
    let entries = archive
        .entries()
        .map_err(|error| format!("Unable to read OCR model package: {error}"))?;
    let mut entry_count = 0_usize;
    let mut unpacked_bytes = 0_u64;
    let mut buffer = vec![0_u8; DOWNLOAD_BUFFER_BYTES];
    for entry in entries {
        ensure_not_cancelled(cancellation)?;
        entry_count += 1;
        if entry_count > MAX_ARCHIVE_ENTRIES {
            return Err("OCR model package contains too many archive entries".to_string());
        }
        let mut entry = entry.map_err(|error| format!("Unable to read OCR model package entry: {error}"))?;
        let path = entry
            .path()
            .map_err(|error| format!("OCR model package contains an invalid path: {error}"))?
            .into_owned();
        validate_archive_path(&path)?;
        let entry_type = entry.header().entry_type();
        if entry_type.is_symlink()
            || entry_type.is_hard_link()
            || matches!(entry_type, EntryType::Block | EntryType::Char | EntryType::Fifo)
        {
            return Err(format!(
                "OCR model package contains an unsupported file type: {}",
                path.display()
            ));
        }
        let entry_size = entry
            .header()
            .size()
            .map_err(|error| format!("OCR model package entry size is invalid: {error}"))?;
        unpacked_bytes = unpacked_bytes
            .checked_add(entry_size)
            .ok_or_else(|| "OCR model package unpacked size overflow".to_string())?;
        if unpacked_bytes > MAX_ARCHIVE_UNPACKED_BYTES {
            return Err("OCR model package expands beyond the 1 GiB safety limit".to_string());
        }

        let target = destination.join(&path);
        if entry_type.is_dir() {
            fs::create_dir_all(&target).map_err(|error| {
                format!("Unable to create OCR model package directory {}: {error}", target.display())
            })?;
            continue;
        }
        if !entry_type.is_file() {
            return Err(format!(
                "OCR model package contains an unsupported file type: {}",
                path.display()
            ));
        }
        if let Some(parent) = target.parent() {
            fs::create_dir_all(parent).map_err(|error| {
                format!("Unable to create OCR model package directory {}: {error}", parent.display())
            })?;
        }
        let mut output = File::create(&target).map_err(|error| {
            format!("Unable to create OCR model package file {}: {error}", target.display())
        })?;
        loop {
            ensure_not_cancelled(cancellation)?;
            let count = entry.read(&mut buffer).map_err(|error| {
                format!("Unable to extract OCR model package file {}: {error}", target.display())
            })?;
            if count == 0 {
                break;
            }
            output.write_all(&buffer[..count]).map_err(|error| {
                format!("Unable to write OCR model package file {}: {error}", target.display())
            })?;
        }
        output.sync_all().map_err(|error| {
            format!("Unable to flush OCR model package file {}: {error}", target.display())
        })?;
    }
    ensure_not_cancelled(cancellation)?;
    Ok(())
}

fn extract_archive(archive_path: &Path, destination: &Path) -> Result<(), String> {
    extract_archive_with_cancel(archive_path, destination, None)
}

fn collect_regular_files(root: &Path) -> Result<Vec<PathBuf>, String> {
    let mut files = Vec::new();
    let mut pending = vec![root.to_path_buf()];
    while let Some(directory) = pending.pop() {
        for entry in fs::read_dir(&directory)
            .map_err(|error| format!("Unable to read OCR model package directory: {error}"))?
        {
            let entry = entry.map_err(|error| format!("Unable to read OCR model package entry: {error}"))?;
            let metadata = fs::symlink_metadata(entry.path())
                .map_err(|error| format!("Unable to inspect OCR model package entry: {error}"))?;
            if metadata.file_type().is_symlink() {
                return Err(format!(
                    "Symbolic links are not allowed in OCR model packages: {}",
                    entry.path().display()
                ));
            }
            if metadata.is_dir() {
                pending.push(entry.path());
            } else if metadata.is_file() {
                files.push(entry.path());
            } else {
                return Err(format!(
                    "Unsupported file type in OCR model package: {}",
                    entry.path().display()
                ));
            }
        }
    }
    files.sort();
    Ok(files)
}

fn verify_pack_with_cancel(
    pack_root: &Path,
    cancellation: Option<(&ModelDownloadControl, u64)>,
) -> Result<ModelPackManifest, String> {
    ensure_not_cancelled(cancellation)?;
    let manifest_path = pack_root.join("pack-manifest.json");
    let content = fs::read(&manifest_path).map_err(|error| {
        format!(
            "Unable to read OCR model package manifest {}: {error}",
            manifest_path.display()
        )
    })?;
    let manifest: ModelPackManifest = serde_json::from_slice(&content)
        .map_err(|error| format!("OCR model package manifest is invalid: {error}"))?;
    if manifest.schema_version != 1
        || manifest.platform != "windows"
        || manifest.architecture != "x64"
    {
        return Err("OCR model package target or schema is invalid".to_string());
    }
    if !KNOWN_MODELS.contains(&manifest.model.as_str()) {
        return Err(format!("Unsupported OCR model package: {}", manifest.model));
    }
    let expected = expected_model_hashes(&manifest.model)
        .ok_or_else(|| format!("Unknown OCR model: {}", manifest.model))?;
    if manifest.files.len() != expected.len() {
        return Err("OCR model package manifest has an unexpected file set".to_string());
    }

    let model_root = pack_root
        .join("paddlex")
        .join("official_models")
        .join(&manifest.model);
    let mut allowed = vec![manifest_path.clone()];
    for (name, expected_hash) in expected {
        let record = manifest
            .files
            .get(name)
            .ok_or_else(|| format!("OCR model package manifest is missing {name}"))?;
        if record.name != name || !record.sha256.eq_ignore_ascii_case(expected_hash) {
            return Err(format!("OCR model package manifest checksum is invalid for {name}"));
        }
        let file = model_root.join(name);
        let metadata = fs::metadata(&file)
            .map_err(|error| format!("OCR model package file is missing {}: {error}", file.display()))?;
        if !metadata.is_file() || metadata.len() != record.size {
            return Err(format!("OCR model package size verification failed: {}", file.display()));
        }
        let actual = sha256_file_with_cancel(&file, cancellation)?;
        if !actual.eq_ignore_ascii_case(expected_hash) {
            return Err(format!("OCR model package SHA-256 verification failed: {}", file.display()));
        }
        allowed.push(file);
    }
    allowed.sort();
    let actual = collect_regular_files(pack_root)?;
    if actual != allowed {
        return Err("OCR model package contains unexpected files".to_string());
    }
    ensure_not_cancelled(cancellation)?;
    Ok(manifest)
}

fn verify_pack(pack_root: &Path) -> Result<ModelPackManifest, String> {
    verify_pack_with_cancel(pack_root, None)
}

fn model_root(runtime_root: &Path) -> PathBuf {
    runtime_root.join("cache").join("paddlex").join("official_models")
}

fn quarantine_root(runtime_root: &Path) -> PathBuf {
    runtime_root.join("quarantine").join("models")
}

fn quarantine_path(runtime_root: &Path, model: &str, source: &Path) -> Result<PathBuf, String> {
    let root = quarantine_root(runtime_root);
    fs::create_dir_all(&root)
        .map_err(|error| format!("Unable to create OCR model quarantine directory: {error}"))?;
    let target = root.join(format!("{model}-{}-{}", now_ms(), Uuid::new_v4()));
    fs::rename(source, &target).map_err(|error| {
        format!(
            "Unable to isolate damaged OCR model {} into {}: {error}",
            source.display(),
            target.display()
        )
    })?;
    Ok(target)
}

fn verify_installed_model_with_cancel(
    directory: &Path,
    model: &str,
    cancellation: Option<(&ModelDownloadControl, u64)>,
) -> Result<(), String> {
    ensure_not_cancelled(cancellation)?;
    let expected = expected_model_hashes(model)
        .ok_or_else(|| format!("Unsupported OCR model directory: {model}"))?;
    let expected_names = expected.keys().copied().collect::<BTreeSet<_>>();
    let actual_names = fs::read_dir(directory)
        .map_err(|error| format!("Unable to read OCR model directory {}: {error}", directory.display()))?
        .filter_map(Result::ok)
        .filter(|entry| entry.file_type().is_ok_and(|kind| kind.is_file()))
        .map(|entry| entry.file_name().to_string_lossy().to_string())
        .collect::<BTreeSet<_>>();
    if actual_names != expected_names.iter().map(|name| (*name).to_string()).collect() {
        return Err("OCR model directory contains missing or unexpected files".to_string());
    }
    for (name, expected_hash) in expected {
        let path = directory.join(name);
        let actual = sha256_file_with_cancel(&path, cancellation)?;
        if !actual.eq_ignore_ascii_case(expected_hash) {
            return Err(format!("OCR model SHA-256 mismatch: {}", path.display()));
        }
    }
    ensure_not_cancelled(cancellation)?;
    Ok(())
}

fn verify_installed_model(directory: &Path, model: &str) -> Result<(), String> {
    verify_installed_model_with_cancel(directory, model, None)
}

pub fn inspect_models(runtime_root: &Path) -> Result<ModelInventory, String> {
    let root = model_root(runtime_root);
    fs::create_dir_all(&root)
        .map_err(|error| format!("Unable to create OCR model directory: {error}"))?;
    let mut inventory = ModelInventory::default();
    for model in KNOWN_MODELS {
        let directory = root.join(model);
        if !directory.exists() {
            continue;
        }
        if !directory.is_dir() {
            let _ = quarantine_path(runtime_root, model, &directory)?;
            inventory.damaged.push((*model).to_string());
            continue;
        }
        match verify_installed_model(&directory, model) {
            Ok(()) => inventory.installed.push((*model).to_string()),
            Err(_) => {
                let _ = quarantine_path(runtime_root, model, &directory)?;
                inventory.damaged.push((*model).to_string());
            }
        }
    }
    Ok(inventory)
}

fn write_atomic(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("Path has no parent directory: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let temporary = path.with_extension(format!("tmp-{}", Uuid::new_v4()));
    let mut file = File::create(&temporary)
        .map_err(|error| format!("Unable to create {}: {error}", temporary.display()))?;
    file.write_all(bytes)
        .and_then(|_| file.sync_all())
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    if path.exists() {
        fs::remove_file(path)
            .map_err(|error| format!("Unable to replace {}: {error}", path.display()))?;
    }
    fs::rename(&temporary, path)
        .map_err(|error| format!("Unable to activate {}: {error}", path.display()))
}

fn install_model_pack_with_cancel(
    package_path: &Path,
    runtime_root: &Path,
    cancellation: Option<(&ModelDownloadControl, u64)>,
) -> Result<String, String> {
    ensure_not_cancelled(cancellation)?;
    if package_path.extension().and_then(|value| value.to_str()) != Some("vtxocrmodel") {
        return Err("Select a VisualTeX .vtxocrmodel package".to_string());
    }
    let metadata = fs::symlink_metadata(package_path).map_err(|error| {
        format!("Unable to inspect OCR model package {}: {error}", package_path.display())
    })?;
    if !metadata.is_file() || metadata.file_type().is_symlink() {
        return Err("The OCR model package must be a regular file".to_string());
    }
    if metadata.len() == 0 || metadata.len() > MAX_MODEL_PACK_BYTES {
        return Err("The OCR model package size is invalid".to_string());
    }
    let required_space = metadata
        .len()
        .saturating_mul(2)
        .saturating_add(512 * 1024 * 1024);
    crate::ocr_storage::ensure_available_space(
        runtime_root,
        required_space,
        "Staging, verifying, and installing the selected OCR model package",
    )?;

    let staging_root = runtime_root
        .join("staging")
        .join(format!("model-{}", Uuid::new_v4()));
    fs::remove_dir_all(&staging_root).ok();
    if let Err(error) = extract_archive_with_cancel(package_path, &staging_root, cancellation) {
        fs::remove_dir_all(&staging_root).ok();
        return Err(error);
    }
    let pack_root = staging_root.join("visualtex-model-pack");
    let manifest = match verify_pack_with_cancel(&pack_root, cancellation) {
        Ok(manifest) => manifest,
        Err(error) => {
            if cancellation.is_some_and(|(control, generation)| control.is_cancelled(generation)) {
                fs::remove_dir_all(&staging_root).ok();
            } else {
                let _ = quarantine_path(runtime_root, "invalid-package", &staging_root);
            }
            return Err(error);
        }
    };

    let models_root = model_root(runtime_root);
    fs::create_dir_all(&models_root)
        .map_err(|error| format!("Unable to create OCR model directory: {error}"))?;
    let suffix = Uuid::new_v4();
    let source = pack_root
        .join("paddlex")
        .join("official_models")
        .join(&manifest.model);
    let incoming = models_root.join(format!(".{}-installing-{suffix}", manifest.model));
    let target = models_root.join(&manifest.model);
    let backup = models_root.join(format!(".{}-backup-{suffix}", manifest.model));
    ensure_not_cancelled(cancellation)?;
    fs::rename(&source, &incoming)
        .map_err(|error| format!("Unable to stage OCR model files: {error}"))?;
    if let Err(error) = verify_installed_model_with_cancel(&incoming, &manifest.model, cancellation) {
        fs::remove_dir_all(&incoming).ok();
        fs::remove_dir_all(&staging_root).ok();
        return Err(error);
    }

    let metadata_root = runtime_root.join("model-packs");
    fs::create_dir_all(&metadata_root)
        .map_err(|error| format!("Unable to create OCR model metadata directory: {error}"))?;
    let metadata_target = metadata_root.join(format!("{}.json", manifest.model));
    let metadata_bytes = serde_json::to_vec_pretty(&manifest)
        .map_err(|error| format!("Unable to serialize OCR model metadata: {error}"))?;

    ensure_not_cancelled(cancellation)?;
    let had_target = target.exists();
    if had_target {
        fs::rename(&target, &backup)
            .map_err(|error| format!("Unable to back up the existing OCR model: {error}"))?;
    }
    if let Err(error) = fs::rename(&incoming, &target) {
        if had_target && backup.exists() {
            let _ = fs::rename(&backup, &target);
        }
        return Err(format!("Unable to atomically activate the OCR model: {error}"));
    }
    if let Err(error) = write_atomic(&metadata_target, &metadata_bytes) {
        let _ = fs::remove_dir_all(&target);
        if had_target && backup.exists() {
            let _ = fs::rename(&backup, &target);
        }
        return Err(error);
    }
    fs::remove_dir_all(&backup).ok();
    fs::remove_dir_all(&staging_root).ok();
    Ok(manifest.model)
}

pub fn install_model_pack(package_path: &Path, runtime_root: &Path) -> Result<String, String> {
    install_model_pack_with_cancel(package_path, runtime_root, None)
}

pub fn remove_model(runtime_root: &Path, model: &str) -> Result<(), String> {
    if !KNOWN_MODELS.contains(&model) {
        return Err(format!("Unsupported OCR model: {model}"));
    }
    let target = model_root(runtime_root).join(model);
    if target.exists() {
        fs::remove_dir_all(&target)
            .map_err(|error| format!("Unable to remove OCR model {model}: {error}"))?;
    }
    let metadata = runtime_root.join("model-packs").join(format!("{model}.json"));
    if metadata.exists() {
        fs::remove_file(metadata)
            .map_err(|error| format!("Unable to remove OCR model metadata: {error}"))?;
    }
    Ok(())
}

#[cfg(debug_assertions)]
fn development_catalog_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("resources")
        .join("ocr-models")
        .join("windows-x64")
        .join("catalog.json")
}

fn parse_catalog(path: &Path) -> Result<ModelCatalog, String> {
    let content = fs::read(path)
        .map_err(|error| format!("Unable to read OCR model catalog {}: {error}", path.display()))?;
    let catalog: ModelCatalog = serde_json::from_slice(&content)
        .map_err(|error| format!("OCR model catalog is invalid: {error}"))?;
    if catalog.schema_version != 1
        || catalog.platform != "windows"
        || catalog.architecture != "x64"
    {
        return Err("OCR model catalog target or schema is invalid".to_string());
    }
    let mut seen = BTreeSet::new();
    for entry in &catalog.entries {
        if !KNOWN_MODELS.contains(&entry.model.as_str()) || !seen.insert(entry.model.clone()) {
            return Err(format!("OCR model catalog contains an invalid model: {}", entry.model));
        }
        if !entry.url.starts_with("https://")
            || entry.size == 0
            || entry.size > MAX_MODEL_PACK_BYTES
            || entry.sha256.len() != 64
            || !entry.sha256.chars().all(|value| value.is_ascii_hexdigit())
        {
            return Err(format!("OCR model catalog entry is invalid: {}", entry.model));
        }
    }
    Ok(catalog)
}

pub fn load_catalog(app: &AppHandle) -> Result<ModelCatalog, String> {
    if let Ok(path) = app
        .path()
        .resolve(MODEL_CATALOG_RESOURCE, BaseDirectory::Resource)
    {
        if path.is_file() {
            return parse_catalog(&path);
        }
    }
    #[cfg(debug_assertions)]
    {
        let path = development_catalog_path();
        if path.is_file() {
            return parse_catalog(&path);
        }
    }
    Err("The OCR model download catalog is not included in this VisualTeX build. Import a verified .vtxocrmodel package instead.".to_string())
}

pub fn catalog_entry(app: &AppHandle, model: &str) -> Result<ModelCatalogEntry, String> {
    if !KNOWN_MODELS.contains(&model) {
        return Err(format!("Unsupported OCR model: {model}"));
    }
    load_catalog(app)?
        .entries
        .into_iter()
        .find(|entry| entry.model == model)
        .ok_or_else(|| format!("The OCR model catalog has no download for {model}"))
}

fn parse_content_range_total(value: &str) -> Option<u64> {
    value.rsplit('/').next()?.parse().ok()
}

fn response_total(response: &Response, existing: u64, expected: u64) -> Result<u64, String> {
    if response.status() == StatusCode::PARTIAL_CONTENT {
        let total = response
            .headers()
            .get(CONTENT_RANGE)
            .and_then(|value| value.to_str().ok())
            .and_then(parse_content_range_total)
            .unwrap_or(expected);
        if total != expected {
            return Err(format!(
                "OCR model server reported total size {total}, expected {expected}"
            ));
        }
        return Ok(total);
    }
    if response.status() == StatusCode::OK && existing == 0 {
        let total = response
            .headers()
            .get(CONTENT_LENGTH)
            .and_then(|value| value.to_str().ok())
            .and_then(|value| value.parse().ok())
            .unwrap_or(expected);
        if total != expected {
            return Err(format!(
                "OCR model server reported total size {total}, expected {expected}"
            ));
        }
        return Ok(total);
    }
    Err(format!(
        "OCR model server did not honor the requested byte range; HTTP {}",
        response.status()
    ))
}

fn publish_download_snapshot(
    control: &ModelDownloadControl,
    snapshot: &mut ModelDownloadSnapshot,
    callback: &mut impl FnMut(&ModelDownloadSnapshot),
) -> Result<(), String> {
    control.set_snapshot(snapshot.clone())?;
    callback(snapshot);
    Ok(())
}

fn publish_download_cancelled<T>(
    control: &ModelDownloadControl,
    snapshot: &mut ModelDownloadSnapshot,
    callback: &mut impl FnMut(&ModelDownloadSnapshot),
) -> Result<T, String> {
    snapshot.state = ModelDownloadState::Cancelled;
    snapshot.speed_bytes_per_second = 0;
    snapshot.eta_seconds = None;
    snapshot.message = "OCR 模型下载已立即取消，.part 文件已保留以便续传".to_string();
    snapshot.error = None;
    publish_download_snapshot(control, snapshot, callback)?;
    Err(cancellation_error())
}

pub(crate) async fn download_once(
    client: &Client,
    entry: &ModelCatalogEntry,
    part_path: &Path,
    control: &ModelDownloadControl,
    generation: u64,
    snapshot: &mut ModelDownloadSnapshot,
    callback: &mut impl FnMut(&ModelDownloadSnapshot),
) -> Result<(), String> {
    let mut existing = fs::metadata(part_path).map(|metadata| metadata.len()).unwrap_or(0);
    if existing > entry.size {
        fs::remove_file(part_path)
            .map_err(|error| format!("Unable to reset oversized OCR model partial file: {error}"))?;
        existing = 0;
    }
    if existing == entry.size {
        snapshot.state = ModelDownloadState::Downloading;
        snapshot.downloaded_bytes = existing;
        snapshot.total_bytes = entry.size;
        snapshot.percent = 100;
        snapshot.speed_bytes_per_second = 0;
        snapshot.eta_seconds = Some(0);
        snapshot.message = format!("{} 的 .part 文件已完整，正在继续校验", entry.model);
        snapshot.error = None;
        publish_download_snapshot(control, snapshot, callback)?;
        return Ok(());
    }

    let mut request = client.get(&entry.url);
    if existing > 0 {
        request = request.header(RANGE, format!("bytes={existing}-"));
    }
    let mut response = tokio::select! {
        _ = control.cancelled(generation) => {
            return publish_download_cancelled(control, snapshot, callback);
        }
        result = request.send() => {
            result.map_err(|error| format!("Unable to connect to the OCR model server: {error}"))?
        }
    };
    if existing > 0 && response.status() == StatusCode::OK {
        fs::remove_file(part_path)
            .map_err(|error| format!("Unable to restart OCR model download: {error}"))?;
        existing = 0;
        response = tokio::select! {
            _ = control.cancelled(generation) => {
                return publish_download_cancelled(control, snapshot, callback);
            }
            result = client.get(&entry.url).send() => {
                result.map_err(|error| format!("Unable to restart OCR model download: {error}"))?
            }
        };
    }
    response
        .error_for_status_ref()
        .map_err(|error| format!("OCR model download failed: {error}"))?;
    let total = response_total(&response, existing, entry.size)?;
    let mut output = OpenOptions::new()
        .create(true)
        .append(existing > 0)
        .truncate(existing == 0)
        .write(true)
        .open(part_path)
        .map_err(|error| format!("Unable to open OCR model partial file: {error}"))?;

    snapshot.state = ModelDownloadState::Downloading;
    snapshot.downloaded_bytes = existing;
    snapshot.total_bytes = total;
    snapshot.message = if existing > 0 {
        format!("正在从 {} 字节处继续下载 {}", existing, entry.model)
    } else {
        format!("正在下载 {}", entry.model)
    };
    snapshot.error = None;
    publish_download_snapshot(control, snapshot, callback)?;

    let started = Instant::now();
    let mut last_update = Instant::now() - DOWNLOAD_PROGRESS_INTERVAL;
    let mut downloaded_this_attempt = 0_u64;
    loop {
        let chunk = tokio::select! {
            _ = control.cancelled(generation) => {
                return publish_download_cancelled(control, snapshot, callback);
            }
            result = response.chunk() => {
                result.map_err(|error| format!("OCR model download stream failed: {error}"))?
            }
        };
        let Some(chunk) = chunk else {
            break;
        };
        if control.is_cancelled(generation) {
            return publish_download_cancelled(control, snapshot, callback);
        }
        output
            .write_all(&chunk)
            .map_err(|error| format!("Unable to write OCR model partial file: {error}"))?;
        downloaded_this_attempt = downloaded_this_attempt.saturating_add(chunk.len() as u64);
        snapshot.downloaded_bytes = existing.saturating_add(downloaded_this_attempt);
        if last_update.elapsed() >= DOWNLOAD_PROGRESS_INTERVAL {
            let elapsed = started.elapsed().as_secs_f64().max(0.001);
            let speed = (downloaded_this_attempt as f64 / elapsed) as u64;
            let remaining = total.saturating_sub(snapshot.downloaded_bytes);
            snapshot.speed_bytes_per_second = speed;
            snapshot.eta_seconds = (speed > 0).then_some(remaining / speed);
            snapshot.percent = ((snapshot.downloaded_bytes.saturating_mul(100) / total).min(100)) as u8;
            snapshot.message = format!("正在下载 {}", entry.model);
            publish_download_snapshot(control, snapshot, callback)?;
            last_update = Instant::now();
        }
    }
    if control.is_cancelled(generation) {
        return publish_download_cancelled(control, snapshot, callback);
    }
    output
        .sync_all()
        .map_err(|error| format!("Unable to flush OCR model partial file: {error}"))?;
    let final_size = fs::metadata(part_path)
        .map_err(|error| format!("Unable to inspect OCR model partial file: {error}"))?
        .len();
    if final_size != entry.size {
        return Err(format!(
            "OCR model download is incomplete: {final_size} of {} bytes",
            entry.size
        ));
    }
    snapshot.downloaded_bytes = final_size;
    snapshot.percent = 100;
    snapshot.speed_bytes_per_second = 0;
    snapshot.eta_seconds = Some(0);
    Ok(())
}

pub async fn download_and_install_model(
    app: &AppHandle,
    runtime_root: &Path,
    model: &str,
    control: &ModelDownloadControl,
    generation: u64,
    mut callback: impl FnMut(&ModelDownloadSnapshot),
) -> Result<String, String> {
    let entry = catalog_entry(app, model)?;
    let required_space = entry
        .size
        .saturating_mul(2)
        .saturating_add(512 * 1024 * 1024);
    crate::ocr_storage::ensure_available_space(
        runtime_root,
        required_space,
        &format!("Downloading, staging, and installing the {} OCR model", entry.model),
    )?;
    let downloads_root = runtime_root.join("downloads");
    fs::create_dir_all(&downloads_root)
        .map_err(|error| format!("Unable to create OCR model download directory: {error}"))?;
    let part_path = downloads_root.join(format!("{}.vtxocrmodel.part", entry.model));
    let package_path = downloads_root.join(format!("{}.vtxocrmodel", entry.model));
    let mut snapshot = ModelDownloadSnapshot::new(model, entry.size);
    control.set_snapshot(snapshot.clone())?;
    if control.is_cancelled(generation) {
        return publish_download_cancelled(control, &mut snapshot, &mut callback);
    }

    let client = Client::builder()
        .user_agent("VisualTeX/1.2.6 OCR model manager")
        .connect_timeout(Duration::from_secs(30))
        .timeout(Duration::from_secs(3 * 60 * 60))
        .build()
        .map_err(|error| format!("Unable to initialize OCR model downloader: {error}"))?;

    let mut last_error = None;
    for attempt in 1..=DOWNLOAD_RETRIES {
        match download_once(
            &client,
            &entry,
            &part_path,
            control,
            generation,
            &mut snapshot,
            &mut callback,
        )
        .await
        {
            Ok(()) => {
                last_error = None;
                break;
            }
            Err(error) if control.is_cancelled(generation) => return Err(error),
            Err(error) => {
                last_error = Some(error.clone());
                snapshot.state = ModelDownloadState::Downloading;
                snapshot.message = format!(
                    "{} 下载中断，正在进行第 {}/{} 次重试",
                    entry.model,
                    attempt.min(DOWNLOAD_RETRIES),
                    DOWNLOAD_RETRIES
                );
                snapshot.error = Some(error);
                publish_download_snapshot(control, &mut snapshot, &mut callback)?;
                if attempt < DOWNLOAD_RETRIES {
                    tokio::select! {
                        _ = control.cancelled(generation) => {
                            return publish_download_cancelled(control, &mut snapshot, &mut callback);
                        }
                        _ = tokio::time::sleep(Duration::from_secs(attempt as u64)) => {}
                    }
                }
            }
        }
    }
    if let Some(error) = last_error {
        snapshot.state = ModelDownloadState::Failed;
        snapshot.message = "OCR 模型下载失败，.part 文件已保留，可重试续传".to_string();
        snapshot.error = Some(error.clone());
        publish_download_snapshot(control, &mut snapshot, &mut callback)?;
        return Err(error);
    }

    snapshot.state = ModelDownloadState::Verifying;
    snapshot.message = format!("正在校验 {} 的文件大小与 SHA-256", entry.model);
    snapshot.error = None;
    publish_download_snapshot(control, &mut snapshot, &mut callback)?;
    let actual = match sha256_file_with_cancel(&part_path, Some((control, generation))) {
        Ok(actual) => actual,
        Err(_) if control.is_cancelled(generation) => {
            return publish_download_cancelled(control, &mut snapshot, &mut callback);
        }
        Err(error) => return Err(error),
    };
    if !actual.eq_ignore_ascii_case(&entry.sha256) {
        let quarantine = downloads_root
            .join("quarantine")
            .join(format!("{}-{}-{}.bad", entry.model, now_ms(), Uuid::new_v4()));
        if let Some(parent) = quarantine.parent() {
            fs::create_dir_all(parent).ok();
        }
        let _ = fs::rename(&part_path, &quarantine);
        let error = format!(
            "OCR model package SHA-256 mismatch for {}. Expected {}, actual {}. The damaged file was isolated.",
            entry.model, entry.sha256, actual
        );
        snapshot.state = ModelDownloadState::Failed;
        snapshot.message = "OCR 模型包校验失败，损坏文件已隔离".to_string();
        snapshot.error = Some(error.clone());
        publish_download_snapshot(control, &mut snapshot, &mut callback)?;
        return Err(error);
    }
    if control.is_cancelled(generation) {
        return publish_download_cancelled(control, &mut snapshot, &mut callback);
    }
    if package_path.exists() {
        fs::remove_file(&package_path)
            .map_err(|error| format!("Unable to replace downloaded OCR model package: {error}"))?;
    }
    fs::rename(&part_path, &package_path)
        .map_err(|error| format!("Unable to finalize OCR model package: {error}"))?;

    snapshot.state = ModelDownloadState::Installing;
    snapshot.message = format!("正在暂存并原子激活 {}", entry.model);
    publish_download_snapshot(control, &mut snapshot, &mut callback)?;
    let installed = match install_model_pack_with_cancel(
        &package_path,
        runtime_root,
        Some((control, generation)),
    ) {
        Ok(installed) => installed,
        Err(_) if control.is_cancelled(generation) => {
            let _ = fs::rename(&package_path, &part_path);
            return publish_download_cancelled(control, &mut snapshot, &mut callback);
        }
        Err(error) => return Err(error),
    };
    let _ = fs::remove_file(&package_path);

    snapshot.state = ModelDownloadState::Complete;
    snapshot.message = format!("{} 已下载、校验并安装完成", installed);
    snapshot.error = None;
    publish_download_snapshot(control, &mut snapshot, &mut callback)?;
    Ok(installed)
}

#[cfg(test)]
mod tests {
    use super::*;
    use flate2::write::GzEncoder;
    use std::net::TcpListener;
    use std::thread;
    use tempfile::TempDir;

    fn write_archive(path: &Path, entries: &[(String, Vec<u8>)]) {
        let file = File::create(path).unwrap();
        let encoder = GzEncoder::new(file, flate2::Compression::default());
        let mut builder = tar::Builder::new(encoder);
        for (name, bytes) in entries {
            let mut header = tar::Header::new_gnu();
            header.set_size(bytes.len() as u64);
            header.set_mode(0o644);
            header.set_cksum();
            builder
                .append_data(&mut header, name, bytes.as_slice())
                .unwrap();
        }
        builder.into_inner().unwrap().finish().unwrap();
    }

    #[test]
    fn package_rejects_path_traversal() {
        let temp = TempDir::new().unwrap();
        let archive = temp.path().join("unsafe.vtxocrmodel");
        let file = File::create(&archive).unwrap();
        let encoder = GzEncoder::new(file, flate2::Compression::default());
        let mut builder = tar::Builder::new(encoder);
        let mut header = tar::Header::new_gnu();
        header.set_size(3);
        header.set_mode(0o644);
        let malicious = b"../escape";
        header.as_mut_bytes()[..malicious.len()].copy_from_slice(malicious);
        header.set_cksum();
        builder.append(&header, &b"bad"[..]).unwrap();
        builder.into_inner().unwrap().finish().unwrap();
        let result = extract_archive(&archive, &temp.path().join("out"));
        assert!(result.is_err());
        assert!(!temp.path().join("escape").exists());
    }

    #[test]
    fn damaged_model_is_quarantined_and_not_reported_as_installed() {
        let temp = TempDir::new().unwrap();
        let directory = model_root(temp.path()).join("PP-FormulaNet_plus-M");
        fs::create_dir_all(&directory).unwrap();
        fs::write(directory.join("inference.json"), b"bad").unwrap();
        fs::write(directory.join("inference.pdiparams"), b"bad").unwrap();
        fs::write(directory.join("inference.yml"), b"bad").unwrap();
        let inventory = inspect_models(temp.path()).unwrap();
        assert!(inventory.installed.is_empty());
        assert_eq!(inventory.damaged, vec!["PP-FormulaNet_plus-M"]);
        assert!(!directory.exists());
        assert!(quarantine_root(temp.path()).is_dir());
    }

    #[test]
    fn forged_model_pack_is_rejected() {
        let temp = TempDir::new().unwrap();
        let model = "PP-FormulaNet_plus-S";
        let files = expected_model_hashes(model)
            .unwrap()
            .into_iter()
            .map(|(name, hash)| {
                (
                    name.to_string(),
                    ModelFileRecord {
                        name: name.to_string(),
                        size: 3,
                        sha256: hash.to_string(),
                    },
                )
            })
            .collect();
        let manifest = ModelPackManifest {
            schema_version: 1,
            platform: "windows".to_string(),
            architecture: "x64".to_string(),
            model: model.to_string(),
            files,
        };
        let root = format!("visualtex-model-pack/paddlex/official_models/{model}");
        let archive = temp.path().join("forged.vtxocrmodel");
        write_archive(
            &archive,
            &[
                (
                    "visualtex-model-pack/pack-manifest.json".to_string(),
                    serde_json::to_vec_pretty(&manifest).unwrap(),
                ),
                (format!("{root}/inference.json"), b"bad".to_vec()),
                (format!("{root}/inference.pdiparams"), b"bad".to_vec()),
                (format!("{root}/inference.yml"), b"bad".to_vec()),
            ],
        );
        let error = install_model_pack(&archive, temp.path()).unwrap_err();
        assert!(error.contains("SHA-256"));
        assert!(inspect_models(temp.path()).unwrap().installed.is_empty());
    }

    #[tokio::test]
    async fn interrupted_download_resumes_with_http_range() {
        let payload = b"0123456789abcdef".to_vec();
        let listener = TcpListener::bind("127.0.0.1:0").unwrap();
        let address = listener.local_addr().unwrap();
        let expected_payload = payload.clone();
        let server = thread::spawn(move || {
            let (mut stream, _) = listener.accept().unwrap();
            let mut request = [0_u8; 4096];
            let count = stream.read(&mut request).unwrap();
            let request = String::from_utf8_lossy(&request[..count]);
            assert!(request.contains("Range: bytes=5-") || request.contains("range: bytes=5-"));
            let remainder = &expected_payload[5..];
            write!(
                stream,
                "HTTP/1.1 206 Partial Content\r\nContent-Length: {}\r\nContent-Range: bytes 5-15/16\r\nConnection: close\r\n\r\n",
                remainder.len()
            )
            .unwrap();
            stream.write_all(remainder).unwrap();
        });

        let temp = TempDir::new().unwrap();
        let part = temp.path().join("model.vtxocrmodel.part");
        fs::write(&part, &payload[..5]).unwrap();
        let entry = ModelCatalogEntry {
            model: "PP-FormulaNet_plus-S".to_string(),
            url: format!("http://{address}/model.vtxocrmodel"),
            size: payload.len() as u64,
            sha256: "0".repeat(64),
        };
        let client = Client::builder().build().unwrap();
        let control = ModelDownloadControl::default();
        let mut snapshot = ModelDownloadSnapshot::new(&entry.model, entry.size);
        download_once(
            &client,
            &entry,
            &part,
            &control,
            0,
            &mut snapshot,
            &mut |_| {},
        )
        .await
        .unwrap();
        server.join().unwrap();
        assert_eq!(fs::read(part).unwrap(), payload);
        assert_eq!(snapshot.downloaded_bytes, 16);
        assert_eq!(snapshot.percent, 100);
    }

    #[tokio::test]
    async fn stalled_network_read_is_cancelled_immediately() {
        let listener = TcpListener::bind("127.0.0.1:0").unwrap();
        let address = listener.local_addr().unwrap();
        let (stop_tx, stop_rx) = std::sync::mpsc::channel();
        let server = thread::spawn(move || {
            let (mut stream, _) = listener.accept().unwrap();
            let mut request = [0_u8; 4096];
            let _ = stream.read(&mut request).unwrap();
            write!(
                stream,
                "HTTP/1.1 200 OK\r\nContent-Length: 1048576\r\nConnection: close\r\n\r\n"
            )
            .unwrap();
            stream.flush().unwrap();
            let _ = stop_rx.recv_timeout(Duration::from_secs(5));
        });

        let temp = TempDir::new().unwrap();
        let part = temp.path().join("stalled.vtxocrmodel.part");
        let entry = ModelCatalogEntry {
            model: "PP-FormulaNet_plus-S".to_string(),
            url: format!("http://{address}/stalled.vtxocrmodel"),
            size: 1024 * 1024,
            sha256: "0".repeat(64),
        };
        let client = Client::builder().build().unwrap();
        let control = Arc::new(ModelDownloadControl::default());
        let lease = control.begin().unwrap();
        let generation = lease.generation();
        let cancel_control = control.clone();
        tokio::spawn(async move {
            tokio::time::sleep(Duration::from_millis(100)).await;
            assert!(cancel_control.cancel());
        });

        let mut snapshot = ModelDownloadSnapshot::new(&entry.model, entry.size);
        let started = Instant::now();
        let error = download_once(
            &client,
            &entry,
            &part,
            control.as_ref(),
            generation,
            &mut snapshot,
            &mut |_| {},
        )
        .await
        .unwrap_err();
        let elapsed = started.elapsed();
        stop_tx.send(()).unwrap();
        server.join().unwrap();

        assert!(error.to_ascii_lowercase().contains("cancel"));
        assert_eq!(snapshot.state, ModelDownloadState::Cancelled);
        assert!(
            elapsed < Duration::from_secs(1),
            "cancellation took too long: {elapsed:?}"
        );
        drop(lease);
    }

    #[test]
    fn content_range_total_is_parsed() {
        assert_eq!(parse_content_range_total("bytes 100-199/1000"), Some(1000));
        assert_eq!(parse_content_range_total("bytes */1000"), Some(1000));
        assert_eq!(parse_content_range_total("invalid"), None);
    }
}
