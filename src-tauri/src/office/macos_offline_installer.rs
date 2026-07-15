use crate::office::macos_offline::offline_root;
use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::HashMap;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use tauri::{AppHandle, Manager};
use uuid::Uuid;

const WORD_APP_NAME: &str = "Microsoft Word.app";
const POWERPOINT_APP_NAME: &str = "Microsoft PowerPoint.app";
const OFFICE_GROUP_CONTAINER: &str = "Library/Group Containers/UBF8T346G9.Office";
const WORD_ADDIN_NAME: &str = "VisualTeX.dotm";
const POWERPOINT_ADDIN_NAME: &str = "VisualTeX.ppam";
const WORD_SCRIPT_NAME: &str = "VisualTeXWord.scpt";
const POWERPOINT_SCRIPT_NAME: &str = "VisualTeXPowerPoint.scpt";
const ADDIN_MANIFEST_NAME: &str = "addins.json";
const WORD_VBA_ENTRY: &str = "word/vbaProject.bin";
const POWERPOINT_VBA_ENTRY: &str = "ppt/vbaProject.bin";
const CUSTOM_UI_ENTRY: &str = "customUI/customUI14.xml";
const PLACEHOLDER_PNG_BASE64: &str =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineHostInstallStatus {
    application_installed: bool,
    files_installed: bool,
    loaded: bool,
    plugin_version: Option<String>,
    install_paths: Vec<String>,
    health_path: String,
    last_error: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineOfficeInstallStatus {
    word: MacOfflineHostInstallStatus,
    powerpoint: MacOfflineHostInstallStatus,
    compiled_artifacts_available: bool,
    resource_root: String,
    powerpoint_addin_path: String,
    word_script_path: String,
    powerpoint_script_path: String,
    tutorial_path: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct PluginHealthFile {
    loaded: bool,
    plugin_version: Option<String>,
    host: Option<String>,
    timestamp: Option<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct AddinManifest {
    schema_version: u32,
    plugin_version: String,
    files: HashMap<String, AddinManifestFile>,
}

#[derive(Debug, Deserialize)]
#[serde(deny_unknown_fields)]
struct AddinManifestFile {
    sha256: String,
}

#[derive(Debug)]
struct FileBackup {
    path: PathBuf,
    bytes: Option<Vec<u8>>,
}

fn user_home() -> Result<PathBuf, String> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .filter(|value| value.is_absolute())
        .ok_or_else(|| "Unable to resolve the current user's home directory".to_string())
}

fn application_paths(home: &Path, app_name: &str) -> [PathBuf; 2] {
    [
        PathBuf::from("/Applications").join(app_name),
        home.join("Applications").join(app_name),
    ]
}

fn application_installed(app_name: &str) -> bool {
    user_home()
        .map(|home| {
            application_paths(&home, app_name)
                .into_iter()
                .any(|path| path.is_dir())
        })
        .unwrap_or(false)
}

fn resource_root(app: &AppHandle) -> Result<PathBuf, String> {
    let bundled = app
        .path()
        .resource_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX resources: {error}"))?
        .join("office/macos-offline");
    if bundled.join("PROTOCOL.md").is_file() {
        return Ok(bundled);
    }
    let development = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .ok_or_else(|| "Unable to resolve the VisualTeX source root".to_string())?
        .join("office/macos-offline");
    if development.join("PROTOCOL.md").is_file() {
        return Ok(development);
    }
    Err("VisualTeX macOS offline add-in resources are missing".to_string())
}

fn powerpoint_addin_path() -> Result<PathBuf, String> {
    Ok(offline_root()?.join("OfficeAddins/VisualTeX.ppam"))
}

fn placeholder_path() -> Result<PathBuf, String> {
    Ok(offline_root()?.join("OfficeAddins/resources/placeholder.png"))
}

fn word_script_path() -> Result<PathBuf, String> {
    Ok(user_home()?.join(format!(
        "Library/Application Scripts/com.microsoft.Word/{WORD_SCRIPT_NAME}"
    )))
}

fn powerpoint_script_path() -> Result<PathBuf, String> {
    Ok(user_home()?.join(format!(
        "Library/Application Scripts/com.microsoft.Powerpoint/{POWERPOINT_SCRIPT_NAME}"
    )))
}

fn health_path(host: &str) -> Result<PathBuf, String> {
    Ok(offline_root()?
        .join("OfficePluginStatus")
        .join(format!("{host}.json")))
}

fn directory_name_matches(value: &Path, expected: &str) -> bool {
    value
        .file_name()
        .and_then(|name| name.to_str())
        .map(|name| {
            name.trim_end_matches(".localized")
                .eq_ignore_ascii_case(expected)
        })
        .unwrap_or(false)
}

fn collect_word_startup_paths(
    directory: &Path,
    depth: usize,
    output: &mut Vec<PathBuf>,
) -> Result<(), String> {
    if depth > 8 || !directory.is_dir() {
        return Ok(());
    }
    let entries = fs::read_dir(directory)
        .map_err(|error| format!("Unable to inspect {}: {error}", directory.display()))?;
    for entry in entries {
        let entry = entry.map_err(|error| {
            format!("Unable to inspect an Office startup directory entry: {error}")
        })?;
        let path = entry.path();
        if !path.is_dir() {
            continue;
        }
        let parent_is_startup = path
            .parent()
            .map(|parent| directory_name_matches(parent, "Startup"))
            .unwrap_or(false);
        if parent_is_startup && directory_name_matches(&path, "Word") {
            output.push(path);
            continue;
        }
        collect_word_startup_paths(&path, depth + 1, output)?;
    }
    Ok(())
}

pub(crate) fn discover_word_startup_paths() -> Result<Vec<PathBuf>, String> {
    let root = user_home()?.join(OFFICE_GROUP_CONTAINER);
    if !root.is_dir() {
        return Ok(Vec::new());
    }
    let mut paths = Vec::new();
    collect_word_startup_paths(&root, 0, &mut paths)?;
    paths.sort();
    paths.dedup();
    Ok(paths)
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

fn atomic_write(path: &Path, bytes: &[u8], mode: u32) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("Path has no parent: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let temporary = parent.join(format!(
        ".{}.{}.tmp",
        path.file_name().and_then(|value| value.to_str()).unwrap_or("visualtex"),
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
        format!("Unable to replace {}: {error}", path.display())
    })?;
    set_mode(path, mode)
}

fn copy_atomic(source: &Path, destination: &Path, mode: u32) -> Result<(), String> {
    let bytes = fs::read(source)
        .map_err(|error| format!("Unable to read {}: {error}", source.display()))?;
    atomic_write(destination, &bytes, mode)
}

fn verify_copy(source: &Path, destination: &Path) -> Result<(), String> {
    let source_bytes = fs::read(source)
        .map_err(|error| format!("Unable to verify {}: {error}", source.display()))?;
    let destination_bytes = fs::read(destination)
        .map_err(|error| format!("Unable to verify {}: {error}", destination.display()))?;
    if source_bytes == destination_bytes {
        Ok(())
    } else {
        Err(format!(
            "Installed file {} does not match its VisualTeX source",
            destination.display()
        ))
    }
}

fn capture_backup(path: &Path) -> Result<FileBackup, String> {
    let bytes = match fs::read(path) {
        Ok(bytes) => Some(bytes),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => None,
        Err(error) => return Err(format!("Unable to back up {}: {error}", path.display())),
    };
    Ok(FileBackup {
        path: path.to_path_buf(),
        bytes,
    })
}

fn restore_backups(backups: &[FileBackup]) -> Result<(), String> {
    let mut failures = Vec::new();
    for backup in backups.iter().rev() {
        let result = match backup.bytes.as_deref() {
            Some(bytes) => atomic_write(&backup.path, bytes, 0o600),
            None => remove_if_exists(&backup.path),
        };
        if let Err(error) = result {
            failures.push(error);
        }
    }
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("; "))
    }
}

fn bytes_contain(haystack: &[u8], needle: &[u8]) -> bool {
    !needle.is_empty() && haystack.windows(needle.len()).any(|window| window == needle)
}

fn sha256_hex(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}

#[cfg(target_os = "macos")]
fn validate_zip_entries(path: &Path, required_entries: &[&str]) -> Result<(), String> {
    let integrity = Command::new("/usr/bin/unzip")
        .arg("-tqq")
        .arg(path)
        .output()
        .map_err(|error| format!("Unable to validate {} with unzip: {error}", path.display()))?;
    if !integrity.status.success() {
        let detail = String::from_utf8_lossy(&integrity.stderr).trim().to_string();
        return Err(if detail.is_empty() {
            format!("Compiled add-in {} is a corrupt OOXML package", path.display())
        } else {
            format!("Compiled add-in {} is corrupt: {detail}", path.display())
        });
    }
    let listing = Command::new("/usr/bin/unzip")
        .arg("-Z1")
        .arg(path)
        .output()
        .map_err(|error| format!("Unable to inspect {} with unzip: {error}", path.display()))?;
    if !listing.status.success() {
        return Err(format!("Unable to list OOXML entries in {}", path.display()));
    }
    let listing_text = String::from_utf8_lossy(&listing.stdout);
    let entries = listing_text
        .lines()
        .map(str::trim)
        .collect::<std::collections::HashSet<_>>();
    for entry in required_entries {
        if !entries.contains(entry) {
            return Err(format!(
                "Compiled add-in {} does not contain required OOXML entry {entry}",
                path.display()
            ));
        }
    }
    Ok(())
}

#[cfg(not(target_os = "macos"))]
fn validate_zip_entries(_path: &Path, _required_entries: &[&str]) -> Result<(), String> {
    Ok(())
}

fn validate_compiled_addin(
    path: &Path,
    expected_name: &str,
    expected_vba_entry: &str,
) -> Result<Vec<u8>, String> {
    if path.file_name().and_then(|value| value.to_str()) != Some(expected_name) {
        return Err(format!("Compiled add-in must be named {expected_name}"));
    }
    let bytes = fs::read(path)
        .map_err(|error| format!("Compiled add-in {} is missing: {error}", path.display()))?;
    if bytes.len() < 1024 {
        return Err(format!(
            "Compiled add-in {} is too small to contain a VBA project",
            path.display()
        ));
    }
    if !bytes.starts_with(b"PK\x03\x04") {
        return Err(format!(
            "Compiled add-in {} is not a valid OOXML package",
            path.display()
        ));
    }
    let required_entries = [expected_vba_entry, CUSTOM_UI_ENTRY];
    for entry in required_entries {
        if !bytes_contain(&bytes, entry.as_bytes()) {
            return Err(format!(
                "Compiled add-in {} does not contain required OOXML entry {entry}",
                path.display()
            ));
        }
    }
    validate_zip_entries(path, &required_entries)?;
    Ok(bytes)
}

fn validate_compiled_artifacts(root: &Path) -> Result<(PathBuf, PathBuf), String> {
    let manifest_path = root.join("resources").join(ADDIN_MANIFEST_NAME);
    let manifest_bytes = fs::read(&manifest_path)
        .map_err(|error| format!("Compiled add-in manifest is missing: {error}"))?;
    if manifest_bytes.is_empty() || manifest_bytes.len() > 64 * 1024 {
        return Err("Compiled add-in manifest has an invalid size".to_string());
    }
    let manifest: AddinManifest = serde_json::from_slice(&manifest_bytes)
        .map_err(|error| format!("Compiled add-in manifest is invalid: {error}"))?;
    if manifest.schema_version != 1 || manifest.plugin_version != env!("CARGO_PKG_VERSION") {
        return Err("Compiled add-in manifest version does not match VisualTeX".to_string());
    }

    let (word, powerpoint) = source_artifacts(root);
    for (path, name, vba_entry) in [
        (&word, WORD_ADDIN_NAME, WORD_VBA_ENTRY),
        (&powerpoint, POWERPOINT_ADDIN_NAME, POWERPOINT_VBA_ENTRY),
    ] {
        let bytes = validate_compiled_addin(path, name, vba_entry)?;
        let expected = manifest
            .files
            .get(name)
            .ok_or_else(|| format!("Compiled add-in manifest is missing {name}"))?;
        if expected.sha256.len() != 64
            || !expected.sha256.bytes().all(|value| value.is_ascii_hexdigit())
            || !expected.sha256.eq_ignore_ascii_case(&sha256_hex(&bytes))
        {
            return Err(format!("Compiled add-in checksum failed for {name}"));
        }
    }
    Ok((word, powerpoint))
}

fn compile_applescript(source: &Path, destination: &Path) -> Result<(), String> {
    let parent = destination
        .parent()
        .ok_or_else(|| "AppleScript destination has no parent".to_string())?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let temporary = parent.join(format!(
        ".{}.{}.tmp.scpt",
        destination
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap_or("VisualTeX"),
        Uuid::new_v4()
    ));
    let output = Command::new("/usr/bin/osacompile")
        .arg("-o")
        .arg(&temporary)
        .arg(source)
        .output()
        .map_err(|error| format!("Unable to launch osacompile: {error}"))?;
    if !output.status.success() {
        let _ = fs::remove_file(&temporary);
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        return Err(if detail.is_empty() {
            format!("Unable to compile {}", source.display())
        } else {
            format!("Unable to compile {}: {detail}", source.display())
        });
    }
    fs::rename(&temporary, destination).map_err(|error| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to install {}: {error}", destination.display())
    })?;
    set_mode(destination, 0o600)
}

fn health_is_current(host: &str, health: &PluginHealthFile) -> bool {
    let host_matches = health.host.as_deref() == Some(host);
    let version_matches = health.plugin_version.as_deref() == Some(env!("CARGO_PKG_VERSION"));
    let has_timestamp = health.timestamp.as_deref().is_some_and(|value| {
        !value.is_empty() && value.len() <= 64 && !value.chars().any(char::is_control)
    });
    health.loaded && host_matches && version_matches && has_timestamp
}

fn read_health(host: &str) -> Result<(bool, Option<String>), String> {
    let path = health_path(host)?;
    let bytes = match fs::read(&path) {
        Ok(value) => value,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok((false, None)),
        Err(error) => return Err(format!("Unable to read {}: {error}", path.display())),
    };
    let health: PluginHealthFile = serde_json::from_slice(&bytes)
        .map_err(|error| format!("{} contains invalid JSON: {error}", path.display()))?;
    let loaded = health_is_current(host, &health);
    Ok((loaded, health.plugin_version))
}

fn source_artifacts(root: &Path) -> (PathBuf, PathBuf) {
    (
        root.join("resources/VisualTeX.dotm"),
        root.join("resources/VisualTeX.ppam"),
    )
}

fn compiled_artifacts_available(root: &Path) -> bool {
    validate_compiled_artifacts(root).is_ok()
}

fn host_status(
    host: &str,
    app_name: &str,
    install_paths: Vec<PathBuf>,
) -> Result<MacOfflineHostInstallStatus, String> {
    let (loaded, plugin_version, last_error) = match read_health(host) {
        Ok((loaded, plugin_version)) => (loaded, plugin_version, None),
        Err(error) => (false, None, Some(error)),
    };
    let files_installed = !install_paths.is_empty() && install_paths.iter().all(|path| path.is_file());
    Ok(MacOfflineHostInstallStatus {
        application_installed: application_installed(app_name),
        files_installed,
        loaded,
        plugin_version,
        install_paths: install_paths
            .iter()
            .map(|path| path.display().to_string())
            .collect(),
        health_path: health_path(host)?.display().to_string(),
        last_error,
    })
}

pub fn status(app: &AppHandle) -> Result<MacOfflineOfficeInstallStatus, String> {
    let root = resource_root(app)?;
    let mut word_paths = discover_word_startup_paths()?
        .into_iter()
        .map(|path| path.join(WORD_ADDIN_NAME))
        .collect::<Vec<_>>();
    let powerpoint_path = powerpoint_addin_path()?;
    let word_script = word_script_path()?;
    let powerpoint_script = powerpoint_script_path()?;
    let placeholder = placeholder_path()?;
    word_paths.extend([word_script.clone(), placeholder.clone()]);
    Ok(MacOfflineOfficeInstallStatus {
        word: host_status("word", WORD_APP_NAME, word_paths)?,
        powerpoint: host_status(
            "powerpoint",
            POWERPOINT_APP_NAME,
            vec![
                powerpoint_path.clone(),
                powerpoint_script.clone(),
                placeholder.clone(),
            ],
        )?,
        compiled_artifacts_available: compiled_artifacts_available(&root),
        resource_root: root.display().to_string(),
        powerpoint_addin_path: powerpoint_path.display().to_string(),
        word_script_path: word_script.display().to_string(),
        powerpoint_script_path: powerpoint_script.display().to_string(),
        tutorial_path: root
            .join("POWERPOINT_INSTALL.md")
            .display()
            .to_string(),
    })
}

pub fn install(app: &AppHandle) -> Result<MacOfflineOfficeInstallStatus, String> {
    if !cfg!(target_os = "macos") {
        return Err("The native offline Office add-ins are available only on macOS".to_string());
    }
    let root = resource_root(app)?;
    let (word_source, powerpoint_source) = validate_compiled_artifacts(&root)?;

    let word_startup_paths = discover_word_startup_paths()?;
    if word_startup_paths.is_empty() {
        return Err(
            "Microsoft Word Startup directory was not found. Start Word once, quit it, and run Repair."
                .to_string(),
        );
    }
    let word_destinations = word_startup_paths
        .iter()
        .map(|startup| startup.join(WORD_ADDIN_NAME))
        .collect::<Vec<_>>();
    let powerpoint_destination = powerpoint_addin_path()?;
    let placeholder_destination = placeholder_path()?;
    let word_script_destination = word_script_path()?;
    let powerpoint_script_destination = powerpoint_script_path()?;
    let word_health = health_path("word")?;
    let powerpoint_health = health_path("powerpoint")?;

    let mut mutation_paths = word_destinations.clone();
    mutation_paths.extend([
        powerpoint_destination.clone(),
        placeholder_destination.clone(),
        word_script_destination.clone(),
        powerpoint_script_destination.clone(),
        word_health.clone(),
        powerpoint_health.clone(),
    ]);
    mutation_paths.sort();
    mutation_paths.dedup();
    let backups = mutation_paths
        .iter()
        .map(|path| capture_backup(path))
        .collect::<Result<Vec<_>, _>>()?;

    let install_result = (|| -> Result<(), String> {
        for (startup, destination) in word_startup_paths.iter().zip(&word_destinations) {
            fs::create_dir_all(startup)
                .map_err(|error| format!("Unable to create {}: {error}", startup.display()))?;
            copy_atomic(&word_source, destination, 0o600)?;
            verify_copy(&word_source, destination)?;
        }
        copy_atomic(&powerpoint_source, &powerpoint_destination, 0o600)?;
        verify_copy(&powerpoint_source, &powerpoint_destination)?;

        let placeholder = BASE64_STANDARD
            .decode(PLACEHOLDER_PNG_BASE64)
            .map_err(|error| format!("Unable to decode the VisualTeX placeholder: {error}"))?;
        atomic_write(&placeholder_destination, &placeholder, 0o600)?;
        if fs::read(&placeholder_destination)
            .map_err(|error| format!("Unable to verify the VisualTeX placeholder: {error}"))?
            != placeholder
        {
            return Err("Installed VisualTeX placeholder does not match its source".to_string());
        }

        compile_applescript(
            &root.join("word/VisualTeXWord.scpt"),
            &word_script_destination,
        )?;
        compile_applescript(
            &root.join("powerpoint/VisualTeXPowerPoint.scpt"),
            &powerpoint_script_destination,
        )?;
        if fs::metadata(&word_script_destination)
            .map(|value| value.len() == 0)
            .unwrap_or(true)
            || fs::metadata(&powerpoint_script_destination)
                .map(|value| value.len() == 0)
                .unwrap_or(true)
        {
            return Err("Compiled VisualTeX AppleScriptTask file is empty".to_string());
        }

        // Old health records must never make a freshly updated file look loaded.
        // Word/PowerPoint recreate these only after the new VBA project executes.
        remove_if_exists(&word_health)?;
        remove_if_exists(&powerpoint_health)?;
        Ok(())
    })();

    if let Err(error) = install_result {
        return match restore_backups(&backups) {
            Ok(()) => Err(format!("{error}. The previous VisualTeX add-ins were restored.")),
            Err(rollback_error) => Err(format!(
                "{error}. VisualTeX could not fully restore the previous files: {rollback_error}"
            )),
        };
    }
    status(app)
}

pub fn uninstall(app: &AppHandle) -> Result<MacOfflineOfficeInstallStatus, String> {
    if !cfg!(target_os = "macos") {
        return Err("The native offline Office add-ins are available only on macOS".to_string());
    }
    for startup in discover_word_startup_paths()? {
        remove_if_exists(&startup.join(WORD_ADDIN_NAME))?;
    }
    remove_if_exists(&powerpoint_addin_path()?)?;
    remove_if_exists(&word_script_path()?)?;
    remove_if_exists(&powerpoint_script_path()?)?;
    remove_if_exists(&placeholder_path()?)?;
    remove_if_exists(&health_path("word")?)?;
    remove_if_exists(&health_path("powerpoint")?)?;
    status(app)
}

fn remove_if_exists(path: &Path) -> Result<(), String> {
    match fs::remove_file(path) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(format!("Unable to remove {}: {error}", path.display())),
    }
}

pub fn reveal_powerpoint_addin() -> Result<(), String> {
    let path = powerpoint_addin_path()?;
    if !path.is_file() {
        return Err("VisualTeX.ppam is not installed. Install or repair the offline add-in first."
            .to_string());
    }
    let status = Command::new("/usr/bin/open")
        .arg("-R")
        .arg(&path)
        .status()
        .map_err(|error| format!("Unable to reveal VisualTeX.ppam in Finder: {error}"))?;
    if status.success() {
        Ok(())
    } else {
        Err("Finder could not reveal VisualTeX.ppam".to_string())
    }
}

pub fn open_powerpoint_tutorial(app: &AppHandle) -> Result<(), String> {
    let tutorial = resource_root(app)?.join("POWERPOINT_INSTALL.md");
    if !tutorial.is_file() {
        return Err("The VisualTeX PowerPoint installation tutorial is missing".to_string());
    }
    let status = Command::new("/usr/bin/open")
        .arg(&tutorial)
        .status()
        .map_err(|error| format!("Unable to open the PowerPoint tutorial: {error}"))?;
    if status.success() {
        Ok(())
    } else {
        Err("macOS could not open the PowerPoint installation tutorial".to_string())
    }
}

#[tauri::command]
pub fn get_macos_offline_office_install_status(
    app: AppHandle,
) -> Result<MacOfflineOfficeInstallStatus, String> {
    status(&app)
}

#[tauri::command]
pub fn install_macos_offline_office_addins(
    app: AppHandle,
) -> Result<MacOfflineOfficeInstallStatus, String> {
    install(&app)
}

#[tauri::command]
pub fn repair_macos_offline_office_addins(
    app: AppHandle,
) -> Result<MacOfflineOfficeInstallStatus, String> {
    install(&app)
}

#[tauri::command]
pub fn uninstall_macos_offline_office_addins(
    app: AppHandle,
) -> Result<MacOfflineOfficeInstallStatus, String> {
    uninstall(&app)
}

#[tauri::command]
pub fn reveal_macos_powerpoint_addin() -> Result<(), String> {
    reveal_powerpoint_addin()
}

#[tauri::command]
pub fn open_macos_powerpoint_addin_tutorial(app: AppHandle) -> Result<(), String> {
    open_powerpoint_tutorial(&app)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn health_requires_exact_host_current_version_and_timestamp() {
        let current = PluginHealthFile {
            loaded: true,
            plugin_version: Some(env!("CARGO_PKG_VERSION").to_string()),
            host: Some("word".to_string()),
            timestamp: Some("2026-07-15T12:34:56".to_string()),
        };
        assert!(health_is_current("word", &current));

        let mut stale = current.clone();
        stale.plugin_version = Some("0.0.0".to_string());
        assert!(!health_is_current("word", &stale));

        let mut wrong_host = current.clone();
        wrong_host.host = Some("powerpoint".to_string());
        assert!(!health_is_current("word", &wrong_host));

        let mut missing_timestamp = current.clone();
        missing_timestamp.timestamp = Some(String::new());
        assert!(!health_is_current("word", &missing_timestamp));

        let mut malformed_timestamp = current;
        malformed_timestamp.timestamp = Some("bad\ntimestamp".to_string());
        assert!(!health_is_current("word", &malformed_timestamp));
    }

    #[test]
    fn application_detection_checks_system_and_user_application_directories() {
        let home = Path::new("/Users/tester");
        let paths = application_paths(home, WORD_APP_NAME);
        assert_eq!(paths[0], PathBuf::from("/Applications/Microsoft Word.app"));
        assert_eq!(
            paths[1],
            PathBuf::from("/Users/tester/Applications/Microsoft Word.app")
        );
    }

    #[test]
    fn file_backups_restore_previous_content_and_remove_new_files() {
        let root = std::env::temp_dir().join(format!("visualtex-rollback-test-{}", Uuid::new_v4()));
        fs::create_dir_all(&root).expect("temp directory should be created");
        let existing = root.join("existing.bin");
        let newly_created = root.join("new.bin");
        fs::write(&existing, b"old").expect("old file should be written");
        let backups = vec![
            capture_backup(&existing).expect("existing backup"),
            capture_backup(&newly_created).expect("missing backup"),
        ];
        atomic_write(&existing, b"new", 0o600).expect("existing file should change");
        atomic_write(&newly_created, b"created", 0o600).expect("new file should be created");
        restore_backups(&backups).expect("rollback should succeed");
        assert_eq!(fs::read(&existing).unwrap(), b"old");
        assert!(!newly_created.exists());
        let _ = fs::remove_dir_all(root);
    }

    #[test]
    fn localized_startup_names_are_matched_without_locale_assumptions() {
        assert!(directory_name_matches(Path::new("Startup.localized"), "Startup"));
        assert!(directory_name_matches(Path::new("Word"), "Word"));
        assert!(!directory_name_matches(Path::new("PowerPoint"), "Word"));
    }

    #[test]
    fn compiled_addin_validation_rejects_missing_or_fake_files() {
        let root = std::env::temp_dir().join(format!("visualtex-addin-test-{}", Uuid::new_v4()));
        fs::create_dir_all(&root).expect("temp directory should be created");
        let fake = root.join(WORD_ADDIN_NAME);
        fs::write(&fake, b"PK\x03\x04fake").expect("fake file should be written");
        assert!(validate_compiled_addin(&fake, WORD_ADDIN_NAME, WORD_VBA_ENTRY).is_err());
        let _ = fs::remove_dir_all(root);
    }
}
