//! In-app updates (Windows).
//!
//! VisualTeX's download server keeps a signed manifest (MANIFEST_URL, beside the release files): the newest
//! version, its release notes and, for each installer flavour (full and no-OCR), the installer's address,
//! size, SHA-256 and minisign signature. The manifest's signature and the installer's are checked against the public key
//! compiled into VisualTeX (update-public-key.txt, as `tauri signer generate` writes it): nothing unsigned
//! runs, and a manifest only offers a version newer than the running one.
//!
//! The user starts each step: the download (into the install folder's updates\, resumable), then the
//! installation. Before the installer runs it is checked again while held open against writes;
//! updates\pending.json notes the update; the installer starts with /UPDATE /P /R (no questions, the parts
//! installed last time, VisualTeX started again at the end) and VisualTeX quits. The next start deletes
//! updates\ when the new version runs, or reports that the update did not finish and keeps the checked
//! installer for another try.

use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};
use minisign_verify::{PublicKey, Signature};
use reqwest::header::{CACHE_CONTROL, CONTENT_LENGTH, CONTENT_RANGE, RANGE};
use reqwest::{Client, StatusCode, Url};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, HashSet};
use std::fs::{self, File, OpenOptions};
use std::future::Future;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Mutex, MutexGuard};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tauri::{AppHandle, Emitter, Manager};
use tokio::sync::Notify;

use crate::app_lifecycle::append_lifecycle_log;

const MANIFEST_URL: &str =
    "https://download.visualtex.pauljianliao.com/visualtex-downloads/updates/windows/latest.json";
const MANIFEST_FORMAT: &str = "visualtex-update-1";
const PRODUCT: &str = "VisualTeX";
/// The release key's public half, as `tauri signer generate` writes it (<key>.pub). A test build names
/// another key at compile time (VISUALTEX_UPDATE_PUBLIC_KEY). No key: no in-app updates.
const RELEASE_PUBLIC_KEY: &str = include_str!("../update-public-key.txt");
/// the check's error when this build has no key (the page then asks GitHub, as before)
pub const KEY_MISSING: &str = "UPDATE_KEY_MISSING";
const CANCELLED: &str = "UPDATE_CANCELLED";
const UPDATES_FOLDER: &str = "updates";
const PENDING_FILE: &str = "pending.json";
/// the installer writes it before the files and renames it after them (hooks.nsh VisualTeXOpenFileList)
const UNFINISHED_LIST: &str = "installed-files.new.txt";
const STATUS_EVENT: &str = "visualtex-update-status";
const MAX_MANIFEST_BYTES: usize = 512 * 1024;
const MAX_INSTALLER_BYTES: u64 = 4 << 30;
/// free space the installation needs besides the download
const INSTALL_HEADROOM_BYTES: u64 = 512 << 20;
const DOWNLOAD_ATTEMPTS: u64 = 3;
const STATUS_INTERVAL: Duration = Duration::from_millis(250);

/// Microsoft Office programs the installer asks to close before it replaces the add-ins
/// (scripts/office_process_guard.ps1), with the names shown to the user
const MICROSOFT_OFFICE_PROCESSES: [(&str, &str); 9] = [
    ("winword.exe", "Word"),
    ("powerpnt.exe", "PowerPoint"),
    ("excel.exe", "Excel"),
    ("outlook.exe", "Outlook"),
    ("onenote.exe", "OneNote"),
    ("msaccess.exe", "Access"),
    ("mspub.exe", "Publisher"),
    ("visio.exe", "Visio"),
    ("winproj.exe", "Project"),
];
/// WPS Office programs that load the add-ins
const WPS_PROCESSES: [&str; 3] = ["wps.exe", "wpp.exe", "et.exe"];

/// The file on the server: the manifest's bytes (base64) and their signature (Tauri signer format), in one
/// file so that a reader never gets a manifest with another manifest's signature.
#[derive(Deserialize)]
struct Envelope {
    format: String,
    manifest: String,
    signature: String,
}

#[derive(Deserialize, Clone, Debug)]
struct Manifest {
    product: String,
    version: String,
    #[serde(default)]
    name: String,
    #[serde(default)]
    notes: String,
    #[serde(default)]
    pub_date: String,
    #[serde(default)]
    release_page: String,
    #[serde(default)]
    platforms: BTreeMap<String, Asset>,
}

#[derive(Deserialize, Serialize, Clone, Debug, PartialEq, Eq)]
struct Asset {
    url: String,
    size: u64,
    sha256: String,
    signature: String,
}

#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Pending {
    from: String,
    to: String,
    installer: String,
    started_at_ms: u64,
}

#[derive(Serialize, Clone, Copy, Debug, Default, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum Phase {
    #[default]
    Idle,
    Downloading,
    Verifying,
    Ready,
    Installing,
    Failed,
}

#[derive(Serialize, Clone, Debug, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct UpdateStatus {
    phase: Phase,
    /// the version the download is for
    version: String,
    downloaded: u64,
    total: u64,
    error: String,
}

#[derive(Serialize, Clone, Debug, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct UpdateOutcome {
    /// "updated" or "unfinished"
    kind: &'static str,
    version: String,
}

#[derive(Clone, Debug)]
struct Offer {
    version: String,
    asset: Asset,
    file_name: String,
}

#[derive(Default)]
struct Inner {
    status: UpdateStatus,
    offer: Option<Offer>,
    /// what the last update did, told to the page once
    outcome: Option<UpdateOutcome>,
    /// an update to this version wrote the program but did not finish: offered again although not newer
    repair_version: Option<String>,
}

#[derive(Default)]
pub struct UpdaterState {
    inner: Mutex<Inner>,
    /// a download or a cancel moves it on; a download whose number is no longer current stops
    generation: AtomicU64,
    cancelled: Notify,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct UpdateCheck {
    current_version: String,
    latest_version: String,
    release_url: String,
    release_name: String,
    release_notes: String,
    published_at: String,
    update_available: bool,
    /// an installer for this build's flavour is offered (otherwise only the release page)
    can_install: bool,
    size: u64,
    status: UpdateStatus,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct InstallStart {
    started: bool,
    /// Office programs to close first (none when the installer started)
    running_apps: Vec<String>,
}

fn lock(state: &UpdaterState) -> MutexGuard<'_, Inner> {
    state.inner.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

/// Changes the status (only for the download GENERATION when one is given) and tells the page.
fn update_status(app: &AppHandle, generation: Option<u64>, change: impl FnOnce(&mut UpdateStatus)) {
    let state = app.state::<UpdaterState>();
    let status = {
        let mut inner = lock(&state);
        if generation.is_some_and(|generation| state.generation.load(Ordering::SeqCst) != generation) {
            return;
        }
        change(&mut inner.status);
        inner.status.clone()
    };
    let _ = app.emit(STATUS_EVENT, status);
}

fn now_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|elapsed| elapsed.as_millis() as u64)
        .unwrap_or(0)
}

/// The installer flavour this build belongs to: the no-OCR build is compiled with VISUALTEX_NO_OCR_BUNDLE=1
/// (scripts/tauri_build.mjs).
fn target() -> &'static str {
    match option_env!("VISUALTEX_NO_OCR_BUNDLE") {
        Some("1") => "windows-x86_64-noocr",
        _ => "windows-x86_64",
    }
}

fn install_dir() -> Result<PathBuf, String> {
    std::env::current_exe()
        .ok()
        .and_then(|executable| executable.parent().map(Path::to_path_buf))
        .ok_or_else(|| "VisualTeX could not find its install folder.".to_string())
}

fn updates_dir() -> Result<PathBuf, String> {
    Ok(install_dir()?.join(UPDATES_FOLDER))
}

/// The manifest's address; VISUALTEX_UPDATE_MANIFEST_URL names another one (tests). Whatever it names, the
/// manifest needs a valid signature from the compiled-in key.
fn manifest_url() -> String {
    std::env::var("VISUALTEX_UPDATE_MANIFEST_URL")
        .ok()
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| MANIFEST_URL.to_string())
}

/// https, or plain http from this computer (tests)
fn secure_address(url: &Url) -> bool {
    url.scheme() == "https"
        || (url.scheme() == "http" && matches!(url.host_str(), Some("127.0.0.1" | "localhost")))
}

/// Installers come from the manifest's own server.
fn same_origin(manifest: &str, installer: &str) -> bool {
    let (Ok(manifest), Ok(installer)) = (Url::parse(manifest), Url::parse(installer)) else {
        return false;
    };
    secure_address(&installer)
        && installer.username().is_empty()
        && installer.password().is_none()
        && manifest.scheme() == installer.scheme()
        && manifest.host_str() == installer.host_str()
        && manifest.port_or_known_default() == installer.port_or_known_default()
}

fn public_key() -> Result<PublicKey, String> {
    let configured = option_env!("VISUALTEX_UPDATE_PUBLIC_KEY")
        .unwrap_or(RELEASE_PUBLIC_KEY)
        .trim();
    if configured.is_empty() {
        return Err(KEY_MISSING.to_string());
    }
    minisign_text(configured)
        .and_then(|text| PublicKey::decode(&text).map_err(|error| error.to_string()))
        .map_err(|error| format!("VisualTeX's update key is unreadable: {error}"))
}

/// Tauri's signer writes keys and signatures as base64 of the minisign text; minisign text is taken as it is.
fn minisign_text(value: &str) -> Result<String, String> {
    let value = value.trim();
    if value.starts_with("untrusted comment:") {
        return Ok(value.to_string());
    }
    let bytes = BASE64.decode(value).map_err(|error| error.to_string())?;
    String::from_utf8(bytes).map_err(|error| error.to_string())
}

fn signature_of(value: &str) -> Result<Signature, String> {
    minisign_text(value)
        .and_then(|text| Signature::decode(&text).map_err(|error| error.to_string()))
        .map_err(|error| format!("A signature in the update manifest is unreadable: {error}"))
}

fn version_number(part: &str) -> Option<u64> {
    if part.is_empty() || !part.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    part.parse().ok()
}

/// major.minor.patch, digits only
fn parse_version(value: &str) -> Option<(u64, u64, u64)> {
    let mut parts = value.trim().split('.');
    let version = (
        version_number(parts.next()?)?,
        version_number(parts.next()?)?,
        version_number(parts.next()?)?,
    );
    if parts.next().is_some() {
        return None;
    }
    Some(version)
}

fn is_newer(candidate: &str, current: &str) -> bool {
    matches!(
        (parse_version(candidate), parse_version(current)),
        (Some(candidate), Some(current)) if candidate > current
    )
}

/// "VisualTeX_1.3.2_x64-setup.exe" -> "1.3.2"
fn version_in_name(name: &str) -> Option<String> {
    name.split('_')
        .find(|part| parse_version(part).is_some())
        .map(str::to_string)
}

/// The manifest, once its signature checked out.
fn open_manifest(body: &[u8], key: &PublicKey) -> Result<Manifest, String> {
    let envelope: Envelope = serde_json::from_slice(body)
        .map_err(|error| format!("The update manifest is unreadable: {error}"))?;
    if envelope.format != MANIFEST_FORMAT {
        return Err(format!(
            "The update manifest has an unknown format: {}",
            envelope.format
        ));
    }
    let bytes = BASE64
        .decode(envelope.manifest.trim())
        .map_err(|error| format!("The update manifest is unreadable: {error}"))?;
    let signature = signature_of(&envelope.signature)?;
    key.verify(&bytes, &signature, false)
        .map_err(|error| format!("The update manifest's signature is not valid ({error})."))?;
    let manifest: Manifest = serde_json::from_slice(&bytes)
        .map_err(|error| format!("The update manifest is unreadable: {error}"))?;
    if manifest.product != PRODUCT || parse_version(&manifest.version).is_none() {
        return Err("The update manifest is not for this program.".to_string());
    }
    Ok(manifest)
}

/// The installer's file name in updates\, once the manifest's entry for it checked out.
fn checked_asset(asset: &Asset, manifest_url: &str) -> Result<String, String> {
    if asset.size == 0 || asset.size > MAX_INSTALLER_BYTES {
        return Err("The update manifest gives an impossible installer size.".to_string());
    }
    if asset.sha256.len() != 64 || !asset.sha256.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err("The update manifest gives no valid SHA-256 for the installer.".to_string());
    }
    signature_of(&asset.signature)?;
    if !same_origin(manifest_url, &asset.url) {
        return Err("The update manifest names an installer on another server.".to_string());
    }
    let name = Url::parse(&asset.url)
        .ok()
        .and_then(|url| url.path_segments().and_then(|segments| segments.last().map(str::to_string)))
        .unwrap_or_default();
    let valid = name.len() <= 120
        && name.to_ascii_lowercase().ends_with(".exe")
        && !name.starts_with('.')
        && name
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'));
    if !valid {
        return Err("The update manifest names an installer file VisualTeX does not accept.".to_string());
    }
    Ok(name)
}

/// Size, SHA-256 and signature of an installer, read once from start to end.
fn verify_installer(mut reader: impl Read, asset: &Asset, key: &PublicKey) -> Result<(), String> {
    let signature = signature_of(&asset.signature)?;
    let mut verifier = key.verify_stream(&signature).map_err(|error| {
        format!("The installer's signature does not fit VisualTeX's update key ({error}).")
    })?;
    let mut digest = Sha256::new();
    let mut length = 0_u64;
    let mut buffer = vec![0_u8; 1 << 20];
    loop {
        let read = reader
            .read(&mut buffer)
            .map_err(|error| format!("Unable to read the downloaded installer: {error}"))?;
        if read == 0 {
            break;
        }
        length += read as u64;
        if length > asset.size {
            return Err("The downloaded installer is larger than the manifest says.".to_string());
        }
        digest.update(&buffer[..read]);
        verifier.update(&buffer[..read]);
    }
    if length != asset.size {
        return Err(format!(
            "The downloaded installer has {length} of {} bytes.",
            asset.size
        ));
    }
    if !hex::encode(digest.finalize()).eq_ignore_ascii_case(&asset.sha256) {
        return Err("The downloaded installer's SHA-256 does not match the manifest.".to_string());
    }
    verifier
        .finalize()
        .map_err(|error| format!("The downloaded installer's signature is not valid ({error})."))
}

/// Opened for reading while no other program can change or delete it.
fn open_held(path: &Path) -> Result<File, String> {
    use std::os::windows::fs::OpenOptionsExt;
    const FILE_SHARE_READ: u32 = 0x1;
    OpenOptions::new()
        .read(true)
        .share_mode(FILE_SHARE_READ)
        .open(path)
        .map_err(|error| format!("Unable to open the downloaded installer: {error}"))
}

fn client() -> Result<Client, String> {
    Client::builder()
        .user_agent(concat!("VisualTeX/", env!("CARGO_PKG_VERSION"), " updater"))
        .connect_timeout(Duration::from_secs(20))
        .read_timeout(Duration::from_secs(60))
        .build()
        .map_err(|error| format!("Unable to start the update connection: {error}"))
}

async fn fetch_manifest(client: &Client, address: &str) -> Result<Vec<u8>, String> {
    let url = Url::parse(address).map_err(|_| "The update address is not valid.".to_string())?;
    if !secure_address(&url) {
        return Err("The update address is not secure.".to_string());
    }
    let mut response = client
        .get(url)
        .header(CACHE_CONTROL, "no-cache")
        .send()
        .await
        .map_err(|error| format!("Unable to reach the update server: {error}"))?;
    if !response.status().is_success() {
        return Err(format!(
            "The update server answered HTTP {}.",
            response.status().as_u16()
        ));
    }
    let mut body = Vec::new();
    while let Some(chunk) = response
        .chunk()
        .await
        .map_err(|error| format!("The update manifest did not arrive: {error}"))?
    {
        body.extend_from_slice(&chunk);
        if body.len() > MAX_MANIFEST_BYTES {
            return Err("The update manifest is too large.".to_string());
        }
    }
    Ok(body)
}

#[tauri::command]
pub async fn check_for_update(app: AppHandle) -> Result<UpdateCheck, String> {
    let key = public_key()?;
    let address = manifest_url();
    let manifest = open_manifest(&fetch_manifest(&client()?, &address).await?, &key)?;
    let current = app.package_info().version.to_string();
    let state = app.state::<UpdaterState>();
    let repair = lock(&state).repair_version.as_deref() == Some(manifest.version.as_str());
    let update_available = is_newer(&manifest.version, &current) || repair;
    let offer = match manifest.platforms.get(target()) {
        Some(asset) if update_available => Some(Offer {
            version: manifest.version.clone(),
            asset: asset.clone(),
            file_name: checked_asset(asset, &address)?,
        }),
        _ => None,
    };
    let size = offer.as_ref().map_or(0, |offer| offer.asset.size);
    // a download that finished earlier (checked again before it runs)
    let downloaded = match (&offer, updates_dir()) {
        (Some(offer), Ok(dir)) => fs::metadata(dir.join(&offer.file_name))
            .is_ok_and(|metadata| metadata.is_file() && metadata.len() == offer.asset.size),
        _ => false,
    };
    let status = {
        let mut inner = lock(&state);
        let busy = matches!(
            inner.status.phase,
            Phase::Downloading | Phase::Verifying | Phase::Installing
        );
        if !busy {
            inner.status = match &offer {
                Some(offer) => UpdateStatus {
                    phase: if downloaded { Phase::Ready } else { Phase::Idle },
                    version: offer.version.clone(),
                    downloaded: if downloaded { offer.asset.size } else { 0 },
                    total: offer.asset.size,
                    error: String::new(),
                },
                None => UpdateStatus::default(),
            };
            inner.offer = offer;
        }
        inner.status.clone()
    };
    let _ = app.emit(STATUS_EVENT, status.clone());
    Ok(UpdateCheck {
        current_version: current,
        latest_version: manifest.version.clone(),
        release_url: manifest.release_page.clone(),
        release_name: if manifest.name.trim().is_empty() {
            format!("VisualTeX v{}", manifest.version)
        } else {
            manifest.name.clone()
        },
        release_notes: manifest.notes,
        published_at: manifest.pub_date,
        update_available,
        can_install: size > 0,
        size,
        status,
    })
}

#[tauri::command]
pub fn download_update(app: AppHandle) -> Result<UpdateStatus, String> {
    let state = app.state::<UpdaterState>();
    let (offer, generation, status) = {
        let mut inner = lock(&state);
        if !matches!(inner.status.phase, Phase::Idle | Phase::Failed) {
            return Ok(inner.status.clone());
        }
        let offer = inner
            .offer
            .clone()
            .ok_or_else(|| "There is no update to download.".to_string())?;
        let generation = state.generation.fetch_add(1, Ordering::SeqCst) + 1;
        inner.status = UpdateStatus {
            phase: Phase::Downloading,
            version: offer.version.clone(),
            downloaded: 0,
            total: offer.asset.size,
            error: String::new(),
        };
        (offer, generation, inner.status.clone())
    };
    let _ = app.emit(STATUS_EVENT, status.clone());
    let task_app = app.clone();
    tauri::async_runtime::spawn(async move {
        let result = download(&task_app, &offer, generation).await;
        match result {
            Ok(()) => {
                append_lifecycle_log(format!("update {} downloaded and checked", offer.version));
                update_status(&task_app, Some(generation), |status| {
                    status.phase = Phase::Ready;
                    status.downloaded = offer.asset.size;
                    status.error.clear();
                });
            }
            Err(error) if error == CANCELLED => {}
            Err(error) => {
                append_lifecycle_log(format!("update {} download failed: {error}", offer.version));
                update_status(&task_app, Some(generation), |status| {
                    status.phase = Phase::Failed;
                    status.error = error;
                });
            }
        }
    });
    Ok(status)
}

/// Runs FUTURE until it ends or the download it belongs to is cancelled.
async fn until_cancelled<T>(
    app: &AppHandle,
    generation: u64,
    future: impl Future<Output = T>,
) -> Result<T, String> {
    let state = app.state::<UpdaterState>();
    // created before the check: a cancel between the two still wakes it
    let cancelled = state.cancelled.notified();
    if state.generation.load(Ordering::SeqCst) != generation {
        return Err(CANCELLED.to_string());
    }
    tokio::select! {
        _ = cancelled => Err(CANCELLED.to_string()),
        value = future => Ok(value),
    }
}

/// Earlier downloads (other versions, other names) go when a new one starts.
fn remove_other_downloads(dir: &Path, keep: &str) {
    let part = format!("{keep}.part");
    let Ok(entries) = fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let name = entry.file_name().to_string_lossy().into_owned();
        if name == keep || name == part || name == PENDING_FILE {
            continue;
        }
        let path = entry.path();
        let _ = if path.is_dir() {
            fs::remove_dir_all(&path)
        } else {
            fs::remove_file(&path)
        };
    }
}

async fn download(app: &AppHandle, offer: &Offer, generation: u64) -> Result<(), String> {
    let dir = updates_dir()?;
    fs::create_dir_all(&dir)
        .map_err(|error| format!("Unable to create the updates folder: {error}"))?;
    remove_other_downloads(&dir, &offer.file_name);
    let installer = dir.join(&offer.file_name);
    let part = dir.join(format!("{}.part", offer.file_name));
    if installer.is_file() {
        let key = public_key()?;
        let checked = open_held(&installer)
            .and_then(|held| verify_installer(&held, &offer.asset, &key));
        if checked.is_ok() {
            return Ok(());
        }
        let _ = fs::remove_file(&installer);
    }
    let existing = fs::metadata(&part)
        .map(|metadata| metadata.len())
        .unwrap_or(0)
        .min(offer.asset.size);
    let needed = offer.asset.size - existing + INSTALL_HEADROOM_BYTES;
    if let Some(available) = crate::ocr::storage::available_space_bytes(&dir) {
        if available < needed {
            return Err(format!(
                "Not enough free disk space: the update needs {} MB, {} MB are free.",
                needed >> 20,
                available >> 20
            ));
        }
    }
    let client = client()?;
    let mut failure = String::new();
    for attempt in 1..=DOWNLOAD_ATTEMPTS {
        match download_once(app, &client, offer, &part, generation).await {
            Ok(()) => {
                failure.clear();
                break;
            }
            Err(error) if error == CANCELLED => return Err(error),
            Err(error) => {
                failure = error;
                if attempt < DOWNLOAD_ATTEMPTS {
                    until_cancelled(app, generation, tokio::time::sleep(Duration::from_secs(2 * attempt)))
                        .await?;
                }
            }
        }
    }
    if !failure.is_empty() {
        return Err(failure);
    }
    update_status(app, Some(generation), |status| status.phase = Phase::Verifying);
    let asset = offer.asset.clone();
    let checked_part = part.clone();
    let checked = tauri::async_runtime::spawn_blocking(move || {
        let key = public_key()?;
        let held = open_held(&checked_part)?;
        verify_installer(&held, &asset, &key)
    })
    .await
    .map_err(|error| format!("The installer check stopped: {error}"))?;
    if let Err(error) = checked {
        let _ = fs::remove_file(&part);
        return Err(error);
    }
    if state_generation(app) != generation {
        return Err(CANCELLED.to_string());
    }
    fs::rename(&part, &installer)
        .map_err(|error| format!("Unable to keep the downloaded installer: {error}"))
}

fn state_generation(app: &AppHandle) -> u64 {
    app.state::<UpdaterState>().generation.load(Ordering::SeqCst)
}

/// One request: the rest of PART (Range) or the whole file.
async fn download_once(
    app: &AppHandle,
    client: &Client,
    offer: &Offer,
    part: &Path,
    generation: u64,
) -> Result<(), String> {
    let size = offer.asset.size;
    let mut existing = fs::metadata(part).map(|metadata| metadata.len()).unwrap_or(0);
    if existing > size {
        fs::remove_file(part).map_err(|error| format!("Unable to restart the download: {error}"))?;
        existing = 0;
    }
    if existing == size {
        return Ok(());
    }
    let mut request = client
        .get(&offer.asset.url)
        .header(CACHE_CONTROL, "no-cache");
    if existing > 0 {
        request = request.header(RANGE, format!("bytes={existing}-"));
    }
    let mut response = until_cancelled(app, generation, request.send())
        .await?
        .map_err(|error| format!("Unable to reach the update server: {error}"))?;
    let status = response.status();
    if existing > 0 && status == StatusCode::OK {
        // the server sent the whole file: written from the start
        existing = 0;
    } else if existing > 0 && status != StatusCode::PARTIAL_CONTENT {
        let _ = fs::remove_file(part);
        return Err(format!(
            "The update server answered HTTP {} to a resumed download.",
            status.as_u16()
        ));
    }
    if !status.is_success() {
        return Err(format!("The update server answered HTTP {}.", status.as_u16()));
    }
    let header = |name| {
        response
            .headers()
            .get(name)
            .and_then(|value| value.to_str().ok())
            .map(str::to_string)
    };
    if status == StatusCode::PARTIAL_CONTENT {
        // bytes <first>-<last>/<total>
        let range = header(CONTENT_RANGE).unwrap_or_default();
        let fits = range
            .strip_prefix("bytes ")
            .and_then(|rest| rest.split_once('/'))
            .and_then(|(span, total)| {
                let first = span.split('-').next()?.parse::<u64>().ok()?;
                Some(first == existing && total.parse::<u64>().ok()? == size)
            })
            .unwrap_or(false);
        if !fits {
            let _ = fs::remove_file(part);
            return Err("The update server resumed the download at the wrong place.".to_string());
        }
    } else if header(CONTENT_LENGTH)
        .and_then(|value| value.parse::<u64>().ok())
        .is_some_and(|length| length != size)
    {
        return Err("The update server offers another file than the manifest names.".to_string());
    }
    let mut output = OpenOptions::new()
        .create(true)
        .write(true)
        .append(existing > 0)
        .truncate(existing == 0)
        .open(part)
        .map_err(|error| format!("Unable to write the download: {error}"))?;
    let mut downloaded = existing;
    update_status(app, Some(generation), |status| {
        status.phase = Phase::Downloading;
        status.downloaded = downloaded;
        status.total = size;
        status.error.clear();
    });
    let mut reported = Instant::now();
    loop {
        let chunk = until_cancelled(app, generation, response.chunk())
            .await?
            .map_err(|error| format!("The download was interrupted: {error}"))?;
        let Some(chunk) = chunk else {
            break;
        };
        downloaded += chunk.len() as u64;
        if downloaded > size {
            drop(output);
            let _ = fs::remove_file(part);
            return Err("The update server sent more than the manifest names.".to_string());
        }
        output
            .write_all(&chunk)
            .map_err(|error| format!("Unable to write the download: {error}"))?;
        if reported.elapsed() >= STATUS_INTERVAL {
            reported = Instant::now();
            update_status(app, Some(generation), |status| status.downloaded = downloaded);
        }
    }
    output
        .sync_all()
        .map_err(|error| format!("Unable to write the download: {error}"))?;
    update_status(app, Some(generation), |status| status.downloaded = downloaded);
    if downloaded != size {
        return Err(format!(
            "The download stopped at {} of {} MB.",
            downloaded >> 20,
            size >> 20
        ));
    }
    Ok(())
}

#[tauri::command]
pub fn cancel_update_download(app: AppHandle) -> UpdateStatus {
    let state = app.state::<UpdaterState>();
    let (status, part) = {
        let mut inner = lock(&state);
        if !matches!(inner.status.phase, Phase::Downloading | Phase::Verifying) {
            return inner.status.clone();
        }
        state.generation.fetch_add(1, Ordering::SeqCst);
        state.cancelled.notify_waiters();
        inner.status.phase = Phase::Idle;
        inner.status.downloaded = 0;
        inner.status.error.clear();
        let part = inner.offer.as_ref().and_then(|offer| {
            updates_dir()
                .ok()
                .map(|dir| dir.join(format!("{}.part", offer.file_name)))
        });
        (inner.status.clone(), part)
    };
    // what came so far goes; the download may hold the file a moment longer
    if let Some(part) = part {
        std::thread::spawn(move || {
            for _ in 0..40 {
                if fs::remove_file(&part).is_ok() || !part.exists() {
                    break;
                }
                std::thread::sleep(Duration::from_millis(250));
            }
            if let Some(dir) = part.parent() {
                let _ = fs::remove_dir(dir);
            }
        });
    }
    let _ = app.emit(STATUS_EVENT, status.clone());
    status
}

#[tauri::command]
pub fn install_update(app: AppHandle, force: bool) -> Result<InstallStart, String> {
    let state = app.state::<UpdaterState>();
    let offer = {
        let inner = lock(&state);
        match (&inner.offer, inner.status.phase) {
            (Some(offer), Phase::Ready) => offer.clone(),
            _ => return Err("The update is not downloaded yet.".to_string()),
        }
    };
    if !force {
        let running_apps = running_office_apps();
        if !running_apps.is_empty() {
            return Ok(InstallStart {
                started: false,
                running_apps,
            });
        }
    }
    let dir = updates_dir()?;
    let installer = dir.join(&offer.file_name);
    let key = public_key()?;
    // checked again, and started while still held open against writes
    let held = open_held(&installer)?;
    if let Err(error) = verify_installer(&held, &offer.asset, &key) {
        drop(held);
        let _ = fs::remove_file(&installer);
        update_status(&app, None, |status| {
            status.phase = Phase::Failed;
            status.downloaded = 0;
            status.error = error.clone();
        });
        return Err(error);
    }
    let pending = Pending {
        from: app.package_info().version.to_string(),
        to: offer.version.clone(),
        installer: offer.file_name.clone(),
        started_at_ms: now_ms(),
    };
    let pending_path = dir.join(PENDING_FILE);
    let text = serde_json::to_vec_pretty(&pending).map_err(|error| error.to_string())?;
    fs::write(&pending_path, text).map_err(|error| format!("Unable to note the update: {error}"))?;
    if let Err(error) = start_installer(&installer) {
        let _ = fs::remove_file(&pending_path);
        return Err(error);
    }
    drop(held);
    append_lifecycle_log(format!(
        "update {} -> {}: installer started, VisualTeX quits",
        pending.from, pending.to
    ));
    update_status(&app, None, |status| status.phase = Phase::Installing);
    let quitting = app.clone();
    std::thread::spawn(move || {
        // the page shows that the installation started before the window goes
        std::thread::sleep(Duration::from_millis(300));
        quitting.exit(0);
    });
    Ok(InstallStart {
        started: true,
        running_apps: Vec::new(),
    })
}

/// /UPDATE: overwrite in place, no uninstall first; /P: progress only, no questions; /R: start VisualTeX at
/// the end. Its working folder is not updates\, which the next start deletes.
fn start_installer(installer: &Path) -> Result<(), String> {
    use std::os::windows::process::CommandExt;
    const CREATE_BREAKAWAY_FROM_JOB: u32 = 0x0100_0000;
    let run = |flags: u32| {
        std::process::Command::new(installer)
            .args(["/UPDATE", "/P", "/R"])
            .current_dir(std::env::temp_dir())
            .creation_flags(flags)
            .spawn()
    };
    // a job VisualTeX may run in must not end the installer when VisualTeX quits
    run(CREATE_BREAKAWAY_FROM_JOB)
        .or_else(|_| run(0))
        .map(|_| ())
        .map_err(|error| format!("Unable to start the installer: {error}"))
}

/// The Office programs the installer would ask to close, for the hosts the add-ins are installed for.
fn running_office_apps() -> Vec<String> {
    let (microsoft_office, wps) = crate::office::windows_backend::office_hosts_chosen();
    let running = running_process_names();
    let mut apps = Vec::new();
    if microsoft_office {
        for (process, name) in MICROSOFT_OFFICE_PROCESSES {
            if running.contains(process) {
                apps.push(name.to_string());
            }
        }
    }
    if wps && WPS_PROCESSES.iter().any(|process| running.contains(*process)) {
        apps.push("WPS Office".to_string());
    }
    apps
}

fn running_process_names() -> HashSet<String> {
    use windows::Win32::Foundation::CloseHandle;
    use windows::Win32::System::Diagnostics::ToolHelp::{
        CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W, TH32CS_SNAPPROCESS,
    };
    let mut names = HashSet::new();
    unsafe {
        let Ok(snapshot) = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) else {
            return names;
        };
        let mut entry = PROCESSENTRY32W {
            dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
            ..Default::default()
        };
        let mut more = Process32FirstW(snapshot, &mut entry).is_ok();
        while more {
            let length = entry
                .szExeFile
                .iter()
                .position(|&unit| unit == 0)
                .unwrap_or(entry.szExeFile.len());
            names.insert(String::from_utf16_lossy(&entry.szExeFile[..length]).to_ascii_lowercase());
            more = Process32NextW(snapshot, &mut entry).is_ok();
        }
        let _ = CloseHandle(snapshot);
    }
    names
}

#[tauri::command]
pub fn get_update_status(app: AppHandle) -> UpdateStatus {
    lock(&app.state::<UpdaterState>()).status.clone()
}

/// What the last update did, once.
#[tauri::command]
pub fn take_update_outcome(app: AppHandle) -> Option<UpdateOutcome> {
    lock(&app.state::<UpdaterState>()).outcome.take()
}

/// What the last update did: pending.json names it. It finished when this is its version and the
/// installer renamed its file list (which it does after the last file). Returns the outcome and the
/// installer to keep for another try.
fn settle(install_dir: &Path, current: &str) -> (Option<UpdateOutcome>, Option<String>) {
    let pending_path = install_dir.join(UPDATES_FOLDER).join(PENDING_FILE);
    let Ok(bytes) = fs::read(&pending_path) else {
        return (None, None);
    };
    let _ = fs::remove_file(&pending_path);
    let Ok(pending) = serde_json::from_slice::<Pending>(&bytes) else {
        return (None, None);
    };
    if pending.to == current && !install_dir.join(UNFINISHED_LIST).exists() {
        (
            Some(UpdateOutcome {
                kind: "updated",
                version: pending.to,
            }),
            None,
        )
    } else {
        (
            Some(UpdateOutcome {
                kind: "unfinished",
                version: pending.to,
            }),
            Some(pending.installer),
        )
    }
}

/// Deletes what updates\ holds except a download for a newer version (or KEEP), trying again while a file
/// is in use (the installer that started this version may still be closing); the folder goes when empty.
fn tidy_updates(dir: &Path, current: &str, keep: Option<&str>) {
    for _ in 0..60 {
        let Ok(entries) = fs::read_dir(dir) else {
            return;
        };
        let mut busy = false;
        for entry in entries.flatten() {
            let name = entry.file_name().to_string_lossy().into_owned();
            let installer = name.strip_suffix(".part").unwrap_or(&name);
            let wanted = keep == Some(installer)
                || version_in_name(installer).is_some_and(|version| is_newer(&version, current));
            if wanted {
                continue;
            }
            let path = entry.path();
            let removed = if path.is_dir() {
                fs::remove_dir_all(&path)
            } else {
                fs::remove_file(&path)
            };
            if removed.is_err() && path.exists() {
                busy = true;
            }
        }
        let _ = fs::remove_dir(dir);
        if !busy {
            return;
        }
        std::thread::sleep(Duration::from_secs(1));
    }
}

/// At start: what the last update did, for the page to tell; updates\ keeps only a download that can still
/// be used.
pub fn start(app: &AppHandle) {
    let Ok(install_dir) = install_dir() else {
        return;
    };
    let current = app.package_info().version.to_string();
    let (outcome, keep) = settle(&install_dir, &current);
    if let Some(outcome) = &outcome {
        append_lifecycle_log(format!(
            "the update to {} {}",
            outcome.version,
            if outcome.kind == "updated" { "finished" } else { "did not finish" }
        ));
    }
    {
        let state = app.state::<UpdaterState>();
        let mut inner = lock(&state);
        if let Some(outcome) = &outcome {
            if outcome.kind == "unfinished" && outcome.version == current {
                inner.repair_version = Some(current.clone());
            }
        }
        inner.outcome = outcome;
    }
    let dir = install_dir.join(UPDATES_FOLDER);
    if dir.is_dir() {
        std::thread::spawn(move || tidy_updates(&dir, &current, keep.as_deref()));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    // a key and signatures made with `tauri signer` (test key, password visualtex-dev-test)
    const TEST_PUBLIC_KEY: &str = "dW50cnVzdGVkIGNvbW1lbnQ6IG1pbmlzaWduIHB1YmxpYyBrZXk6IDlBQ0ZBMUJEMjNCMkY5MjAKUldRZytiSWp2YUhQbWxSNmkwa2pjTFRzdWtuMm5tVkY0RVBpTTRoNG9tN2Nxd1JFbTNhYUhZWmoK";
    // signature of the 12 bytes "hello update"
    const SAMPLE_SIGNATURE: &str = "dW50cnVzdGVkIGNvbW1lbnQ6IHNpZ25hdHVyZSBmcm9tIHRhdXJpIHNlY3JldCBrZXkKUlVRZytiSWp2YUhQbXVxOHRIcWY0V0ppRFNaOW9nVHVrc0hMSWVHaEY3Ly9zb1pqMzdZdDBScjZqZjV2Q1Q2cmQwK0Y3Vlc0TkRXOVVOTS9NYVhnWE5MSkE5dDVwMmhCelFJPQp0cnVzdGVkIGNvbW1lbnQ6IHRpbWVzdGFtcDoxNzk1NTAzNzg0CWZpbGU6c2FtcGxlLmJpbgo3WHBma0JxTEFYMm9ZU1ZRU0h6U21sUkN6aFVnNTl3SHNtWm5GUGlSMjg0OHRBZ0g2MzIzenhpdkFoS2JLZVZ1aEJmS3ZFalpDWml2VEcxRG8zQXNDUT09Cg==";
    const SAMPLE: &[u8] = b"hello update";

    fn test_key() -> PublicKey {
        PublicKey::decode(&minisign_text(TEST_PUBLIC_KEY).unwrap()).unwrap()
    }

    fn sample_asset() -> Asset {
        Asset {
            url: "https://download.visualtex.pauljianliao.com/visualtex-downloads/releases/v9.0.0/VisualTeX_9.0.0_x64-setup.exe".into(),
            size: SAMPLE.len() as u64,
            sha256: hex::encode(Sha256::digest(SAMPLE)),
            signature: SAMPLE_SIGNATURE.into(),
        }
    }

    #[test]
    fn versions_compare_by_number() {
        assert!(is_newer("1.3.10", "1.3.9"));
        assert!(is_newer("2.0.0", "1.99.99"));
        assert!(!is_newer("1.3.1", "1.3.1"));
        assert!(!is_newer("1.3.0", "1.3.1"));
        assert!(!is_newer("1.4", "1.3.1"));
        assert!(!is_newer("1.4.0-beta", "1.3.1"));
        assert!(!is_newer("+2.0.0", "1.3.1"));
        assert_eq!(version_in_name("VisualTeX_1.3.2_x64-setup.exe").as_deref(), Some("1.3.2"));
        assert_eq!(version_in_name("VisualTeX_1.3.2_x64-no-ocr-setup.exe.part").as_deref(), Some("1.3.2"));
        assert_eq!(version_in_name("pending.json"), None);
    }

    #[test]
    fn a_signed_installer_checks_out_and_any_change_does_not() {
        let key = test_key();
        let asset = sample_asset();
        assert_eq!(verify_installer(SAMPLE, &asset, &key), Ok(()));
        assert!(verify_installer(&b"hello updatf"[..], &asset, &key).is_err());
        assert!(verify_installer(&b"hello update!"[..], &asset, &key).is_err());
        // a file with the right digest in the manifest but no matching signature
        let other = b"hello other!";
        let forged = Asset { sha256: hex::encode(Sha256::digest(other)), ..asset.clone() };
        assert!(verify_installer(&other[..], &forged, &key)
            .unwrap_err()
            .contains("signature"));
    }

    #[test]
    fn installers_come_only_from_the_manifest_server() {
        let asset = sample_asset();
        assert_eq!(
            checked_asset(&asset, MANIFEST_URL).as_deref(),
            Ok("VisualTeX_9.0.0_x64-setup.exe")
        );
        let elsewhere = Asset { url: "https://example.com/VisualTeX_9.0.0_x64-setup.exe".into(), ..asset.clone() };
        assert!(checked_asset(&elsewhere, MANIFEST_URL).is_err());
        let plain = Asset { url: asset.url.replace("https://", "http://"), ..asset.clone() };
        assert!(checked_asset(&plain, MANIFEST_URL).is_err());
        let script = Asset { url: asset.url.replace("setup.exe", "setup.cmd"), ..asset.clone() };
        assert!(checked_asset(&script, MANIFEST_URL).is_err());
        let local = Asset { url: "http://127.0.0.1:8870/VisualTeX_9.0.0_x64-setup.exe".into(), ..asset };
        assert!(checked_asset(&local, "http://127.0.0.1:8870/latest.json").is_ok());
        assert!(checked_asset(&local, "http://127.0.0.1:8871/latest.json").is_err());
    }

    #[test]
    fn an_unsigned_or_changed_manifest_is_refused() {
        let key = test_key();
        let manifest = br#"{"product":"VisualTeX","version":"9.0.0","platforms":{}}"#;
        let envelope = |bytes: &[u8]| {
            serde_json::to_vec(&serde_json::json!({
                "format": MANIFEST_FORMAT,
                "manifest": BASE64.encode(bytes),
                "signature": SAMPLE_SIGNATURE,
            }))
            .unwrap()
        };
        // the sample signature is for other bytes
        assert!(open_manifest(&envelope(manifest), &key)
            .unwrap_err()
            .contains("signature"));
        // the bytes it is for are not a manifest
        assert!(open_manifest(&envelope(SAMPLE), &key).is_err());
    }

    #[test]
    fn the_next_start_tells_what_the_update_did() {
        let temp = TempDir::new().unwrap();
        let root = temp.path().join("VisualTeX");
        let updates = root.join(UPDATES_FOLDER);
        fs::create_dir_all(&updates).unwrap();
        let pending = |to: &str| {
            let text = serde_json::to_vec(&Pending {
                from: "1.3.1".into(),
                to: to.into(),
                installer: format!("VisualTeX_{to}_x64-setup.exe"),
                started_at_ms: 1,
            })
            .unwrap();
            fs::write(updates.join(PENDING_FILE), text).unwrap();
        };

        pending("1.3.2");
        let (outcome, keep) = settle(&root, "1.3.2");
        assert_eq!(outcome.unwrap().kind, "updated");
        assert_eq!(keep, None);
        assert!(!updates.join(PENDING_FILE).exists());

        // the installer stopped before its last file
        pending("1.3.2");
        fs::write(root.join(UNFINISHED_LIST), b"").unwrap();
        let (outcome, keep) = settle(&root, "1.3.2");
        assert_eq!(outcome.unwrap().kind, "unfinished");
        assert_eq!(keep.as_deref(), Some("VisualTeX_1.3.2_x64-setup.exe"));
        fs::remove_file(root.join(UNFINISHED_LIST)).unwrap();

        // the version did not change (installer cancelled or failed)
        pending("1.3.2");
        assert_eq!(settle(&root, "1.3.1").0.unwrap().kind, "unfinished");
        assert_eq!(settle(&root, "1.3.1"), (None, None));
    }

    /// The release script's last check (scripts/prepare_windows_update.mjs): the folder VISUALTEX_RELEASE_DIR holds
    /// latest.json and the installers it names, taken as VisualTeX takes them from the server: the manifest signed
    /// with this build's key (VISUALTEX_UPDATE_PUBLIC_KEY or update-public-key.txt), both flavours there under their
    /// own names, each installer's size, SHA-256 and signature right. VISUALTEX_RELEASE_VERSION: the version it must be.
    /// Run: cargo test --lib the_release_folder_checks_out -- --ignored
    #[test]
    #[ignore]
    fn the_release_folder_checks_out() {
        let dir = PathBuf::from(std::env::var("VISUALTEX_RELEASE_DIR").expect("VISUALTEX_RELEASE_DIR names the release folder"));
        let key = public_key().expect("this build has an update key");
        let body = fs::read(dir.join("latest.json")).expect("latest.json in the release folder");
        let manifest = open_manifest(&body, &key).expect("the manifest");
        let expected = std::env::var("VISUALTEX_RELEASE_VERSION").unwrap_or_default();
        if !expected.is_empty() {
            assert_eq!(manifest.version, expected, "the manifest's version");
        }
        for (target, no_ocr) in [("windows-x86_64", false), ("windows-x86_64-noocr", true)] {
            let asset = manifest.platforms.get(target).unwrap_or_else(|| panic!("the manifest has no {target} installer"));
            let name = checked_asset(asset, MANIFEST_URL).unwrap_or_else(|error| panic!("{target}: {error}"));
            assert_eq!(version_in_name(&name).as_deref(), Some(manifest.version.as_str()), "{target}: the version in {name}");
            assert_eq!(name.contains("no-ocr"), no_ocr, "{target}: {name} is the other flavour");
            let file = File::open(dir.join(&name)).unwrap_or_else(|error| panic!("{name}: {error}"));
            verify_installer(file, asset, &key).unwrap_or_else(|error| panic!("{name}: {error}"));
            println!("{target}: {name}, {} bytes, checked", asset.size);
        }
    }

    #[test]
    fn only_a_download_that_can_still_be_used_stays() {
        let temp = TempDir::new().unwrap();
        let updates = temp.path().join(UPDATES_FOLDER);
        fs::create_dir_all(&updates).unwrap();
        for name in [
            "VisualTeX_1.3.2_x64-setup.exe",
            "VisualTeX_1.3.3_x64-setup.exe.part",
            "VisualTeX_1.3.4_x64-setup.exe",
            "stray.tmp",
        ] {
            fs::write(updates.join(name), b"x").unwrap();
        }
        tidy_updates(&updates, "1.3.3", None);
        let mut left: Vec<String> = fs::read_dir(&updates)
            .unwrap()
            .flatten()
            .map(|entry| entry.file_name().to_string_lossy().into_owned())
            .collect();
        left.sort();
        assert_eq!(left, vec!["VisualTeX_1.3.4_x64-setup.exe".to_string()]);

        tidy_updates(&updates, "1.3.4", Some("VisualTeX_1.3.4_x64-setup.exe"));
        assert!(updates.join("VisualTeX_1.3.4_x64-setup.exe").exists());
        tidy_updates(&updates, "1.3.4", None);
        assert!(!updates.exists());
    }
}
