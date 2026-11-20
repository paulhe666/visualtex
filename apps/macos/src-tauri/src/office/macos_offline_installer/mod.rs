//! Installing the Word and PowerPoint add-ins on macOS: the compiled add-ins are checked, copied into the Office
//! folders with backups, and their health read back for the settings. mod.rs keeps the names and types, the paths,
//! the status, the install and uninstall steps and the Tauri commands; each area is a child module.

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::HashMap;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tauri::{AppHandle, Manager};
use uuid::Uuid;

mod files;
mod health;
mod locations;
mod maintenance;
mod math_fonts;
mod office_apps;
mod powerpoint;
mod validation;

use files::*;
use health::*;
pub(crate) use locations::*;
pub(crate) use maintenance::*;
use math_fonts::*;
pub(crate) use math_fonts::install_math_fonts_with_wps_addin;
pub(crate) use office_apps::*;
pub(crate) use powerpoint::*;
use validation::*;

pub const MAINTENANCE_INSTALL_ARGUMENT: &str = "--install-macos-office-addins";

const WORD_APP_NAME: &str = "Microsoft Word.app";
const POWERPOINT_APP_NAME: &str = "Microsoft PowerPoint.app";
const WORD_PROCESS_NAME: &str = "Microsoft Word";
const POWERPOINT_PROCESS_NAME: &str = "Microsoft PowerPoint";
const OFFICE_GROUP_CONTAINER: &str = "Library/Group Containers/UBF8T346G9.Office";
const WORD_ADDIN_NAME: &str = "VisualTeX.dotm";
const POWERPOINT_ADDIN_NAME: &str = "VisualTeX.ppam";
const WORD_SCRIPT_NAME: &str = "VisualTeXWord.scpt";
const POWERPOINT_SCRIPT_NAME: &str = "VisualTeXPowerPoint.scpt";
const OFFICE_MAINTENANCE_LOG_NAME: &str = "office-addin-maintenance.log";
const ADDIN_MANIFEST_NAME: &str = "addins.json";
const LEGACY_WORD_MANIFEST_ID: &str = "d6fcb260-4c37-4f73-a173-cf24674f81f2";
const LEGACY_POWERPOINT_MANIFEST_ID: &str = "a6d13cf2-54e8-4dfa-a20c-15de864ab3c5";
const WORD_VBA_ENTRY: &str = "word/vbaProject.bin";
const POWERPOINT_VBA_ENTRY: &str = "ppt/vbaProject.bin";
const WORD_VBA_SOURCE_REVISION: &str = "word-office-performance-20260801-r130";
const POWERPOINT_VBA_SOURCE_REVISION: &str = "powerpoint-native-omml-edit-20260926-r1";
const CUSTOM_UI_ENTRY: &str = "customUI/customUI14.xml";
const CONTENT_TYPES_ENTRY: &str = "[Content_Types].xml";
const CONTENT_TYPES_ZIP_PATTERN: &str = "\\[Content_Types\\].xml";
const WORD_MAIN_PART: &str = "/word/document.xml";
const WORD_MAIN_CONTENT_TYPE: &str =
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml";
const POWERPOINT_MAIN_PART: &str = "/ppt/presentation.xml";
const POWERPOINT_MAIN_CONTENT_TYPE: &str =
    "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml";
// A pending Word formula must not appear as a black square while the native
// editor is open. Keep a one-pixel transparent InlineShape as the durable
// transaction target; the VBA adapter moves the caret after it immediately.
const PLACEHOLDER_PNG_BASE64: &str =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGBgAAAABQABpfZFQAAAAABJRU5ErkJggg==";
// The math fonts of Word's OMML font menu (office/macos-offline/fonts, pinned
// in addins.json), installed into the user's font folder with the add-ins.
// TrueType builds: Word's PDF export embeds TrueType math fonts and rasterizes
// CFF ones. The Windows add-in uses the same families, so a document shows the
// same fonts on both platforms.
const BUNDLED_MATH_FONTS: [&str; 3] = [
    "VisualTeXLatinModernMath.ttf",
    "VisualTeXTeXGyreTermesMath.ttf",
    "STIXTwoMath-Regular.ttf",
];
const BUNDLED_MATH_FONT_DIRECTORY: &str = "fonts";
const USER_FONT_DIRECTORY: &str = "Library/Fonts";

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineHostInstallStatus {
    application_installed: bool,
    application_running: bool,
    files_present: bool,
    files_installed: bool,
    health_reported: bool,
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

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase", default)]
struct PluginHealthFile {
    loaded: Option<bool>,
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

#[derive(Debug)]
struct PluginHealthStatus {
    reported: bool,
    loaded: bool,
    plugin_version: Option<String>,
}

fn user_home() -> Result<PathBuf, String> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .filter(|value| value.is_absolute())
        .ok_or_else(|| "Unable to resolve the current user's home directory".to_string())
}

fn offline_root() -> Result<PathBuf, String> {
    Ok(user_home()?.join(OFFICE_GROUP_CONTAINER).join("VisualTeX"))
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
    #[cfg(debug_assertions)]
    {
        let development = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .ok_or_else(|| "Unable to resolve the VisualTeX source root".to_string())?
            .join("office/macos-offline");
        if development.join("PROTOCOL.md").is_file() {
            return Ok(development);
        }
    }
    Err("VisualTeX macOS offline add-in resources are missing".to_string())
}

fn powerpoint_addin_path() -> Result<PathBuf, String> {
    Ok(offline_root()?.join("OfficeAddins/VisualTeX.ppam"))
}

fn legacy_group_placeholder_path() -> Result<PathBuf, String> {
    Ok(offline_root()?.join("OfficeAddins/resources/placeholder.png"))
}

fn word_runtime_root() -> Result<PathBuf, String> {
    crate::office::macos_offline::runtime_root(crate::office::sessions::OfficeHost::Word)
}

fn powerpoint_runtime_root() -> Result<PathBuf, String> {
    crate::office::macos_offline::runtime_root(crate::office::sessions::OfficeHost::Powerpoint)
}

fn compatibility_placeholder_path() -> Result<PathBuf, String> {
    // The currently reviewed DOTM combines the Application Scripts runtime
    // root with the older OfficeAddins/resources suffix. Keep this exact path
    // installed until the next manually compiled VTProtocol is promoted.
    Ok(word_runtime_root()?.join("OfficeAddins/resources/placeholder.png"))
}

fn canonical_placeholder_path() -> Result<PathBuf, String> {
    Ok(
        user_home()?
            .join("Library/Application Scripts/com.microsoft.Word/VisualTeXPlaceholder.png"),
    )
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
    crate::office::macos_offline::health_path(host)
}

/// Returned by `status` when an Office path probe timed out. A timed-out probe
/// says nothing about whether the add-in is installed: macOS may be holding the
/// read behind its "access data from other apps" prompt, or Office may still be
/// coordinating its containers. Reporting "not installed" then would open a
/// false repair prompt.
pub const OFFICE_STATUS_UNAVAILABLE_ERROR: &str =
    "Office add-in status is temporarily unavailable: macOS did not let VisualTeX read the Office folders in time.";

static OFFICE_PATH_PROBE_TIMEOUTS: AtomicU64 = AtomicU64::new(0);

fn clean_visualtex_owned_office_state() -> Result<(), String> {
    crate::office::background::uninstall_launch_agent()?;

    let home = user_home()?;
    let office_group = home.join(OFFICE_GROUP_CONTAINER);
    let mut word_artifacts = discover_word_startup_artifacts()?;
    for user_content in ["User Content.localized", "User Content"] {
        for startup in ["Startup.localized", "Startup"] {
            word_artifacts.push(
                office_group
                    .join(user_content)
                    .join(startup)
                    .join("Word")
                    .join(WORD_ADDIN_NAME),
            );
        }
    }
    word_artifacts.extend([
        home.join("Library/Application Support/Microsoft/Office/Startup/Word")
            .join(WORD_ADDIN_NAME),
        home.join(
            "Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/Startup/Word",
        )
        .join(WORD_ADDIN_NAME),
    ]);
    word_artifacts.sort();
    word_artifacts.dedup();
    for artifact in word_artifacts {
        remove_if_exists(&artifact)?;
    }

    let mut powerpoint_artifacts = discover_powerpoint_addin_artifacts()?;
    powerpoint_artifacts.push(powerpoint_addin_path()?);
    for user_content in ["User Content.localized", "User Content"] {
        for addins in ["Add-Ins.localized", "Add-Ins", "Addins"] {
            powerpoint_artifacts.push(
                office_group
                    .join(user_content)
                    .join(addins)
                    .join(POWERPOINT_ADDIN_NAME),
            );
        }
    }
    powerpoint_artifacts.extend([
        home.join("Library/Application Support/Microsoft/Office/PowerPoint Add-Ins")
            .join(POWERPOINT_ADDIN_NAME),
        home.join(
            "Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/Microsoft/Office/PowerPoint Add-Ins",
        )
        .join(POWERPOINT_ADDIN_NAME),
    ]);
    powerpoint_artifacts.sort();
    powerpoint_artifacts.dedup();
    for artifact in powerpoint_artifacts {
        remove_if_exists(&artifact)?;
    }

    for file in [
        word_script_path()?,
        powerpoint_script_path()?,
        canonical_placeholder_path()?,
        compatibility_placeholder_path()?,
        legacy_group_placeholder_path()?,
        health_path("word")?,
        health_path("powerpoint")?,
    ] {
        remove_if_exists(&file)?;
    }
    for legacy_file in discover_legacy_visualtex_wef_files()? {
        remove_if_exists(&legacy_file)?;
    }

    remove_dir_if_exists(&word_runtime_root()?)?;
    remove_dir_if_exists(&powerpoint_runtime_root()?)?;

    let root = offline_root()?;
    for directory in ["OfficeSessions", "OfficePluginStatus", "NativeDocuments"] {
        remove_dir_if_exists(&root.join(directory))?;
    }
    remove_dir_if_exists(&home.join("Library/Logs/VisualTeX"))?;
    Ok(())
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

fn host_status(
    host: &str,
    app_name: &str,
    process_name: &str,
    install_paths: Vec<PathBuf>,
    files_present: bool,
    files_installed: bool,
) -> Result<MacOfflineHostInstallStatus, String> {
    let application_running = application_running(process_name);
    let macro_responsive = files_installed
        && application_running
        && crate::office::macos_offline::refresh_health_signal(host);
    let (mut health_reported, mut loaded, mut plugin_version, last_error) = match read_health(host)
    {
        Ok(health) => (health.reported, health.loaded, health.plugin_version, None),
        Err(error) => (false, false, None, Some(error)),
    };
    if macro_responsive {
        // A successful fixed macro invocation proves that the installed DOTM or
        // PPAM is loaded. Do not display a false warning merely because Office
        // delayed or sandboxed the optional JSON health write.
        health_reported = true;
        loaded = true;
        if plugin_version.is_none() {
            plugin_version = Some(env!("CARGO_PKG_VERSION").to_string());
        }
    }
    // A stale health record is only historical evidence that an add-in once
    // loaded. It must never make a missing or byte-mismatched add-in look
    // currently installed after an app/plugin update.
    if !files_installed {
        loaded = false;
    }
    Ok(MacOfflineHostInstallStatus {
        application_installed: application_installed(app_name),
        application_running,
        files_present,
        files_installed,
        health_reported,
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
    status_at(&resource_root(app)?)
}

fn status_at(root: &Path) -> Result<MacOfflineOfficeInstallStatus, String> {
    let probe_timeouts_before = OFFICE_PATH_PROBE_TIMEOUTS.load(Ordering::Relaxed);
    let status = status_snapshot(root)?;
    if OFFICE_PATH_PROBE_TIMEOUTS.load(Ordering::Relaxed) != probe_timeouts_before {
        return Err(OFFICE_STATUS_UNAVAILABLE_ERROR.to_string());
    }
    Ok(status)
}

fn status_snapshot(root: &Path) -> Result<MacOfflineOfficeInstallStatus, String> {
    let (word_source, powerpoint_source) = source_artifacts(root);
    let word_destinations = discover_word_startup_paths()?
        .into_iter()
        .map(|path| path.join(WORD_ADDIN_NAME))
        .collect::<Vec<_>>();
    let discovered_word_artifacts = discover_word_startup_artifacts()?;
    let powerpoint_path = powerpoint_addin_path()?;
    let powerpoint_destinations = vec![powerpoint_path.clone()];
    let discovered_powerpoint_artifacts = discover_powerpoint_addin_artifacts()?;
    let word_script = word_script_path()?;
    let powerpoint_script = powerpoint_script_path()?;
    let compatibility_placeholder = compatibility_placeholder_path()?;
    let canonical_placeholder = canonical_placeholder_path()?;
    let word_support_paths = vec![
        word_script.clone(),
        compatibility_placeholder.clone(),
        canonical_placeholder.clone(),
    ];
    let powerpoint_support_paths = vec![powerpoint_script.clone()];
    let word_files_present = !discovered_word_artifacts.is_empty();
    let powerpoint_files_present = !discovered_powerpoint_artifacts.is_empty();
    let word_files_installed = word_support_paths
        .iter()
        .all(|path| bounded_path_test(path, "-f"))
        && addin_installation_matches(&word_source, &word_destinations, &discovered_word_artifacts);
    let powerpoint_files_installed = powerpoint_support_paths
        .iter()
        .all(|path| bounded_path_test(path, "-f"))
        && addin_installation_matches(
            &powerpoint_source,
            &powerpoint_destinations,
            &discovered_powerpoint_artifacts,
        );
    let mut word_paths = word_destinations.clone();
    word_paths.extend(discovered_word_artifacts.iter().cloned());
    word_paths.extend(word_support_paths);
    word_paths.sort();
    word_paths.dedup();
    let mut powerpoint_paths = powerpoint_destinations.clone();
    powerpoint_paths.extend(discovered_powerpoint_artifacts.iter().cloned());
    powerpoint_paths.extend(powerpoint_support_paths);
    powerpoint_paths.sort();
    powerpoint_paths.dedup();
    Ok(MacOfflineOfficeInstallStatus {
        word: host_status(
            "word",
            WORD_APP_NAME,
            WORD_PROCESS_NAME,
            word_paths,
            word_files_present,
            word_files_installed,
        )?,
        powerpoint: host_status(
            "powerpoint",
            POWERPOINT_APP_NAME,
            POWERPOINT_PROCESS_NAME,
            powerpoint_paths,
            powerpoint_files_present,
            powerpoint_files_installed,
        )?,
        compiled_artifacts_available: compiled_artifacts_available(root),
        resource_root: root.display().to_string(),
        powerpoint_addin_path: powerpoint_path.display().to_string(),
        word_script_path: word_script.display().to_string(),
        powerpoint_script_path: powerpoint_script.display().to_string(),
        tutorial_path: root.join("POWERPOINT_INSTALL.md").display().to_string(),
    })
}

fn install_impl(root: &Path, operation: &str) -> Result<MacOfflineOfficeInstallStatus, String> {
    let result = install_addins_impl(root, operation);
    // The fonts of Word's OMML font menu, whatever happened to the add-ins (for
    // example Word still open): a font problem is logged, never reported as an
    // add-in failure.
    if cfg!(target_os = "macos") {
        install_bundled_math_fonts(root, operation);
    }
    result
}

fn install_addins_impl(
    root: &Path,
    operation: &str,
) -> Result<MacOfflineOfficeInstallStatus, String> {
    if !cfg!(target_os = "macos") {
        return Err("The native offline Office add-ins are available only on macOS".to_string());
    }
    let (word_source, powerpoint_source) = maintenance_step(
        operation,
        "validate-packaged-addins",
        || validate_compiled_artifacts(root),
    )?;
    maintenance_step(operation, "office-process-check", ensure_office_hosts_stopped)?;
    maintenance_step(operation, "office-folder-access", probe_office_folder_access)?;
    maintenance_step(operation, "backup-packaged-addins", || {
        backup_compiled_artifacts_to_scratch(&word_source, &powerpoint_source)
    })?;

    let word_startup_paths = word_startup_install_paths()?;
    let word_destinations = word_startup_paths
        .iter()
        .map(|startup| startup.join(WORD_ADDIN_NAME))
        .collect::<Vec<_>>();
    let powerpoint_destination = powerpoint_addin_path()?;
    let compatibility_placeholder_destination = compatibility_placeholder_path()?;
    let canonical_placeholder_destination = canonical_placeholder_path()?;
    let legacy_group_placeholder = legacy_group_placeholder_path()?;
    let word_script_destination = word_script_path()?;
    let powerpoint_script_destination = powerpoint_script_path()?;
    let word_health = health_path("word")?;
    let powerpoint_health = health_path("powerpoint")?;
    let legacy_wef_files = discover_legacy_visualtex_wef_files()?;
    let existing_word_artifacts = discover_word_startup_artifacts()?;
    let existing_powerpoint_artifacts = discover_powerpoint_addin_artifacts()?;

    let mut mutation_paths = word_destinations.clone();
    mutation_paths.extend([
        powerpoint_destination.clone(),
        compatibility_placeholder_destination.clone(),
        canonical_placeholder_destination.clone(),
        legacy_group_placeholder.clone(),
        word_script_destination.clone(),
        powerpoint_script_destination.clone(),
        word_health.clone(),
        powerpoint_health.clone(),
    ]);
    mutation_paths.extend(legacy_wef_files.iter().cloned());
    mutation_paths.extend(existing_word_artifacts.iter().cloned());
    mutation_paths.extend(existing_powerpoint_artifacts.iter().cloned());
    mutation_paths.sort();
    mutation_paths.dedup();
    let backups = mutation_paths
        .iter()
        .map(|path| capture_backup(path))
        .collect::<Result<Vec<_>, _>>()?;

    let install_result = (|| -> Result<(), String> {
        maintenance_step(operation, "cleanup-existing-addins", || {
            clean_visualtex_owned_office_state()
        })?;

        append_office_maintenance_log(operation, "copy-word-addin", "begin", "");
        for (startup, destination) in word_startup_paths.iter().zip(&word_destinations) {
            fs::create_dir_all(startup)
                .map_err(|error| format!("Unable to create {}: {error}", startup.display()))?;
            copy_atomic(&word_source, destination, 0o600)?;
            verify_copy(&word_source, destination)?;
        }
        append_office_maintenance_log(operation, "copy-word-addin", "ok", "");

        append_office_maintenance_log(operation, "copy-powerpoint-addin", "begin", "");
        copy_atomic(&powerpoint_source, &powerpoint_destination, 0o600)?;
        verify_copy(&powerpoint_source, &powerpoint_destination)?;
        verify_addin_installation(
            &word_source,
            &word_destinations,
            &discover_word_startup_artifacts()?,
            "Word",
        )?;
        verify_addin_installation(
            &powerpoint_source,
            &[powerpoint_destination.clone()],
            &discover_powerpoint_addin_artifacts()?,
            "PowerPoint",
        )?;
        append_office_maintenance_log(operation, "copy-powerpoint-addin", "ok", "");

        let placeholder = BASE64_STANDARD
            .decode(PLACEHOLDER_PNG_BASE64)
            .map_err(|error| format!("Unable to decode the VisualTeX placeholder: {error}"))?;
        for destination in [
            &compatibility_placeholder_destination,
            &canonical_placeholder_destination,
        ] {
            atomic_write(destination, &placeholder, 0o600)?;
            if fs::read(destination)
                .map_err(|error| format!("Unable to verify the VisualTeX placeholder: {error}"))?
                != placeholder
            {
                return Err(format!(
                    "Installed VisualTeX placeholder does not match its source: {}",
                    destination.display()
                ));
            }
        }

        append_office_maintenance_log(operation, "compile-applescript", "begin", "");
        compile_applescript(
            &root.join("word/VisualTeXWord.scpt"),
            &word_script_destination,
        )?;
        compile_applescript(
            &root.join("powerpoint/VisualTeXPowerPoint.scpt"),
            &powerpoint_script_destination,
        )?;
        append_office_maintenance_log(operation, "compile-applescript", "ok", "");
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

        // The retired Office.js manifests survive removal of the desktop app in
        // each Office sandbox. Remove only VisualTeX's two fixed legacy IDs so
        // the native DOTM/PPAM Ribbon cannot be confused with the old web add-in.
        for legacy_file in &legacy_wef_files {
            remove_if_exists(legacy_file)?;
        }
        Ok(())
    })();

    if let Err(error) = install_result {
        return match restore_backups(&backups) {
            Ok(()) => Err(format!(
                "{error}. The previous VisualTeX add-ins were restored."
            )),
            Err(rollback_error) => Err(format!(
                "{error}. VisualTeX could not fully restore the previous files: {rollback_error}"
            )),
        };
    }
    maintenance_step(operation, "post-install-status", || status_at(root))
}

pub fn install(app: &AppHandle) -> Result<MacOfflineOfficeInstallStatus, String> {
    run_maintenance_operation("install", || {
        let root = maintenance_step("install", "resolve-resources", || resource_root(app))?;
        install_impl(&root, "install")
    })
}

fn uninstall_impl() -> Result<(), String> {
    if !cfg!(target_os = "macos") {
        return Err("The native offline Office add-ins are available only on macOS".to_string());
    }
    maintenance_step("uninstall", "office-process-check", ensure_office_hosts_stopped)?;
    maintenance_step("uninstall", "office-folder-access", probe_office_folder_access)?;
    // Uninstall success must not depend on a second status probe. On machines
    // with a wedged Office container, cleanup can succeed even though probing
    // that container is temporarily unavailable.
    maintenance_step("uninstall", "cleanup-installed-addins", || {
        clean_visualtex_owned_office_state()
    })
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfficeHostsRunning {
    word: bool,
    powerpoint: bool,
}

#[tauri::command]
pub async fn get_macos_office_hosts_running() -> Result<MacOfficeHostsRunning, String> {
    tauri::async_runtime::spawn_blocking(|| MacOfficeHostsRunning {
        word: application_running(WORD_PROCESS_NAME),
        powerpoint: application_running(POWERPOINT_PROCESS_NAME),
    })
    .await
    .map_err(|error| format!("Office process check failed: {error}"))
}

#[tauri::command]
pub async fn register_macos_powerpoint_addin(
    retry: Option<bool>,
) -> Result<PowerPointRegistrationResult, String> {
    let retry = retry.unwrap_or(false);
    tauri::async_runtime::spawn_blocking(move || register_powerpoint_addin(retry))
        .await
        .map_err(|error| format!("PowerPoint registration worker failed: {error}"))?
}

#[tauri::command]
pub async fn get_macos_offline_office_install_status(
    app: AppHandle,
) -> Result<MacOfflineOfficeInstallStatus, String> {
    let worker = tauri::async_runtime::spawn_blocking(move || status(&app));
    match tokio::time::timeout(Duration::from_secs(5), worker).await {
        Ok(result) => result
            .map_err(|error| format!("Office installation status worker failed: {error}"))?,
        Err(_) => Err(
            "Timed out reading native Office add-in status. Quit Word and PowerPoint, then retry."
                .to_string(),
        ),
    }
}

// Install, repair and uninstall run in a child process with a time limit, so
// a wait inside macOS never freezes the app (Tauri runs synchronous commands
// on the main thread).
#[tauri::command]
pub async fn install_macos_offline_office_addins(
    app: AppHandle,
) -> Result<serde_json::Value, String> {
    run_maintenance_command(app, "install").await
}

#[tauri::command]
pub async fn repair_macos_offline_office_addins(
    app: AppHandle,
) -> Result<serde_json::Value, String> {
    run_maintenance_command(app, "repair").await
}

#[tauri::command]
pub async fn uninstall_macos_offline_office_addins(
    app: AppHandle,
) -> Result<serde_json::Value, String> {
    run_maintenance_command(app, "uninstall").await
}

/// The macOS major version (27 for macOS 27), from sw_vers.
fn macos_major_version() -> Option<u32> {
    let output = Command::new("/usr/bin/sw_vers")
        .arg("-productVersion")
        .stdin(Stdio::null())
        .output()
        .ok()?;
    String::from_utf8_lossy(&output.stdout)
        .trim()
        .split('.')
        .next()?
        .parse()
        .ok()
}

/// macOS 27 lists Full Disk Access per app under Files & Folders; earlier
/// releases have a Full Disk Access list of its own.
fn full_disk_access_settings_url(macos_major: Option<u32>) -> &'static str {
    if macos_major.is_some_and(|major| major >= 27) {
        "x-apple.systempreferences:com.apple.preference.security?Privacy_FilesAndFolders"
    } else {
        "x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles"
    }
}

#[tauri::command]
pub fn get_macos_major_version() -> Option<u32> {
    macos_major_version()
}

#[tauri::command]
pub fn open_macos_full_disk_access_settings() -> Result<(), String> {
    let url = full_disk_access_settings_url(macos_major_version());
    let status = Command::new("/usr/bin/open")
        .arg(url)
        .status()
        .map_err(|error| format!("Unable to open System Settings: {error}"))?;
    if status.success() {
        Ok(())
    } else {
        Err(format!("Unable to open System Settings ({status})"))
    }
}

#[tauri::command]
pub fn request_quit_macos_office_hosts_for_addin_update() -> Result<(), String> {
    request_office_hosts_quit_for_update()
}

#[tauri::command]
pub fn open_macos_office_addin_maintenance_log() -> Result<(), String> {
    open_office_maintenance_log()
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
mod tests;
