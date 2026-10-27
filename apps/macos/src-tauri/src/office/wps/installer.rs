//! Installs the VisualTeX add-in into WPS Office for Mac.
//!
//! WPS loads JS add-ins listed in `publish.xml` inside its sandbox container
//! (`.kingsoft/wps/jsaddons`). The add-in installed there is a small bootstrap
//! that loads its logic from the running VisualTeX (bridge.rs), so it rarely
//! changes; WPS asks the user to trust an add-in again whenever its files do.

use super::bridge::{self, WPS_BRIDGE_PORT};
use crate::office::state::OfficeCompanionState;
use serde::Serialize;
use serde_json::Value;
use std::fs;
use std::io::ErrorKind;
use std::path::{Path, PathBuf};
use std::sync::mpsc;
use std::time::Duration;

/// macOS refused the WPS container (App Data protection); the Settings page
/// recognizes this text and shows the Full Disk Access steps.
pub(crate) const WPS_FOLDER_ACCESS_DENIED: &str = "macOS denied VisualTeX access to the WPS add-in folder";
const ADDIN_NAME: &str = "VisualTeX";
/// The bootstrap's version; it names the add-in folder (WPS requires
/// `name_version`). Bump it only when the installed files change.
const ADDIN_VERSION: &str = "1";
const WPS_BUNDLE_IDS: [&str; 2] = ["com.kingsoft.wpsoffice.mac", "com.kingsoft.wpsoffice.mac.global"];
const WPS_APPLICATIONS: [&str; 2] = ["/Applications/WPS Office.app", "/Applications/wpsoffice.app"];
const AUTO_INSTALL_DISABLED_FILE: &str = "auto-install-disabled";
const FILESYSTEM_TIMEOUT: Duration = Duration::from_secs(15);

const INDEX_HTML: &str = include_str!("../../../../office/wps/addin/index.html");
const MAIN_JS: &str = include_str!("../../../../office/wps/addin/main.js");
const RIBBON_XML: &str = include_str!("../../../../office/wps/addin/ribbon.xml");
const IMAGES: [(&str, &[u8]); 17] = [
    ("vtxImageInline.svg", include_bytes!("../../../../office/wps/addin/images/vtxImageInline.svg")),
    ("vtxImageDisplay.svg", include_bytes!("../../../../office/wps/addin/images/vtxImageDisplay.svg")),
    ("vtxOmmlInline.svg", include_bytes!("../../../../office/wps/addin/images/vtxOmmlInline.svg")),
    ("vtxOmmlDisplay.svg", include_bytes!("../../../../office/wps/addin/images/vtxOmmlDisplay.svg")),
    ("vtxEditSelected.svg", include_bytes!("../../../../office/wps/addin/images/vtxEditSelected.svg")),
    // The other buttons share the Word add-in's ribbon icons.
    ("vtxConvertNative.svg", WORD_ICON_CONVERT_OMML),
    ("vtxConvertImage.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/convert-image.svg")),
    ("vtxConvertDocument.svg", WORD_ICON_CONVERT_OMML),
    ("vtxUpdateNumbers.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/update-numbers.svg")),
    ("vtxNumberingFormat.svg", WORD_ICON_NUMBERING_FORMAT),
    ("vtxCrossReference.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/cross-reference.svg")),
    ("vtxBatchNumbering.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/batch-numbering.svg")),
    ("vtxUpdateEquationFormat.svg", WORD_ICON_NUMBERING_FORMAT),
    ("vtxDocumentImport.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/document-import.svg")),
    ("vtxOpen.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/open-visualtex.svg")),
    ("vtxRedrawSelection.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/redraw-selection.svg")),
    ("vtxRedrawDocument.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/redraw-document.svg")),
];
const WORD_ICON_CONVERT_OMML: &[u8] = include_bytes!("../../../../office/macos-offline/word/ribbon-icons/convert-omml.svg");
const WORD_ICON_NUMBERING_FORMAT: &[u8] =
    include_bytes!("../../../../office/macos-offline/word/ribbon-icons/numbering-format.svg");

/// Brings the running WPS forward, as the Word add-in's document operations
/// bring Word forward. Only a WPS that runs is activated: `open -b` would
/// launch one that does not.
pub(crate) fn activate_running_wps() {
    for bundle_id in WPS_BUNDLE_IDS {
        let running = std::process::Command::new("/usr/bin/lsappinfo")
            .args(["find", &format!("bundleid={bundle_id}")])
            .output()
            .map(|output| output.status.success() && !String::from_utf8_lossy(&output.stdout).trim().is_empty())
            .unwrap_or(false);
        if running {
            let _ = std::process::Command::new("/usr/bin/open").args(["-b", bundle_id]).output();
            return;
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WpsContainerStatus {
    pub bundle_id: String,
    pub addin_path: String,
    pub installed: bool,
    pub current: bool,
    pub error: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WpsAddinStatus {
    pub wps_installed: bool,
    pub containers: Vec<WpsContainerStatus>,
    pub bridge_running: bool,
    pub bridge_error: Option<String>,
    pub port: u16,
    pub auto_install_disabled: bool,
}

fn home() -> Result<PathBuf, String> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .filter(|path| path.is_absolute())
        .ok_or_else(|| "Unable to resolve the current user's home directory".to_string())
}

/// The jsaddons folders of the WPS builds that have run at least once.
fn jsaddons_roots(home: &Path) -> Vec<(String, PathBuf)> {
    WPS_BUNDLE_IDS
        .iter()
        .map(|bundle_id| {
            let data = home.join("Library/Containers").join(bundle_id).join("Data");
            (bundle_id.to_string(), data)
        })
        .filter(|(_, data)| data.is_dir())
        .map(|(bundle_id, data)| (bundle_id, data.join(".kingsoft/wps/jsaddons")))
        .collect()
}

fn addin_folder_name() -> String {
    format!("{ADDIN_NAME}_{ADDIN_VERSION}")
}

fn publish_entry() -> String {
    format!(r#"<jsplugin name="{ADDIN_NAME}" enable="enable" url="file://" type="wps" version="{ADDIN_VERSION}"/>"#)
}

fn config_js(token: &str) -> String {
    let config = serde_json::json!({ "port": WPS_BRIDGE_PORT, "token": token });
    format!("window.VISUALTEX_WPS_CONFIG = {config};\n")
}

fn addin_files(token: &str) -> Vec<(String, Vec<u8>)> {
    let mut files = vec![
        ("index.html".to_string(), INDEX_HTML.as_bytes().to_vec()),
        ("main.js".to_string(), MAIN_JS.as_bytes().to_vec()),
        ("ribbon.xml".to_string(), RIBBON_XML.as_bytes().to_vec()),
        ("config.js".to_string(), config_js(token).into_bytes()),
    ];
    for (name, bytes) in IMAGES {
        files.push((format!("images/{name}"), bytes.to_vec()));
    }
    files
}

fn is_our_publish_line(line: &str) -> bool {
    line.contains(&format!(r#"name="{ADDIN_NAME}""#))
}

/// Our entry as WPS keeps it. Once the user allows the add-in, WPS rewrites
/// publish.xml (attribute order, indentation) and turns enable="enable" into
/// "enable_dev", so the entry is matched by its attributes, not its text.
fn our_publish_line(text: &str) -> Option<&str> {
    text.lines().find(|line| is_our_publish_line(line))
}

fn publish_line_is_current(line: &str) -> bool {
    let version = format!(r#"version="{ADDIN_VERSION}""#);
    [version.as_str(), r#"url="file://""#, r#"type="wps""#]
        .iter()
        .all(|attribute| line.contains(attribute))
}

fn publish_line_is_enabled(line: &str) -> bool {
    line.contains(r#"enable="enable""#) || line.contains(r#"enable="enable_dev""#)
}

/// `publish.xml` with our entry added or replaced and other add-ins' entries
/// kept, or None when it already lists this add-in version. Rewriting a
/// listed entry would make WPS ask the user to allow the add-in again, so an
/// entry turned off in WPS is only replaced when `reenable` (Install button).
fn updated_publish_xml(existing: Option<&str>, reenable: bool) -> Option<String> {
    let entry = publish_entry();
    match existing.filter(|text| text.contains("</jsplugins>")) {
        Some(text) => {
            if our_publish_line(text)
                .is_some_and(|line| publish_line_is_current(line) && (!reenable || publish_line_is_enabled(line)))
            {
                return None;
            }
            let kept: Vec<&str> = text.lines().filter(|line| !is_our_publish_line(line)).collect();
            Some(kept.join("\n").replacen("</jsplugins>", &format!("  {entry}\n</jsplugins>"), 1))
        }
        None => Some(format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<jsplugins>\n  {entry}\n</jsplugins>\n")),
    }
}

fn publish_xml_without_us(existing: &str) -> String {
    existing
        .lines()
        .filter(|line| !is_our_publish_line(line))
        .collect::<Vec<_>>()
        .join("\n")
}

/// WPS records each trusted add-in in authaddin.json with a signature of its
/// files; a stale record makes WPS skip the add-in silently. Returns the file
/// without VisualTeX's records, or None when there is nothing to remove.
fn authaddin_without_us(text: &str) -> Option<String> {
    let mut json: Value = serde_json::from_str(text).ok()?;
    let mut changed = false;
    for (_, group) in json.as_object_mut()?.iter_mut() {
        let Some(group) = group.as_object_mut() else { continue };
        let ours: Vec<String> = group
            .iter()
            .filter(|(_, entry)| entry.get("name").and_then(Value::as_str) == Some(ADDIN_NAME))
            .map(|(key, _)| key.clone())
            .collect();
        for key in &ours {
            group.remove(key);
            changed = true;
        }
        if !ours.is_empty() {
            if let Some(Value::String(names)) = group.get_mut("namelist") {
                *names = names
                    .split(';')
                    .filter(|name| !name.is_empty() && !ours.iter().any(|key| key == name))
                    .collect::<Vec<_>>()
                    .join(";");
            }
        }
    }
    changed.then(|| serde_json::to_string_pretty(&json).unwrap_or_else(|_| text.to_string()))
}

fn access_error(error: std::io::Error, path: &Path) -> String {
    if error.kind() == ErrorKind::PermissionDenied || error.raw_os_error() == Some(1) {
        format!("{WPS_FOLDER_ACCESS_DENIED} ({}): {error}", path.display())
    } else {
        format!("Unable to change {}: {error}", path.display())
    }
}

fn addin_is_current(jsaddons: &Path, token: &str) -> bool {
    let folder = jsaddons.join(addin_folder_name());
    let files_match = addin_files(token)
        .iter()
        .all(|(name, bytes)| fs::read(folder.join(name)).map(|actual| &actual == bytes).unwrap_or(false));
    let listed = fs::read_to_string(jsaddons.join("publish.xml"))
        .map(|text| our_publish_line(&text).is_some_and(publish_line_is_current))
        .unwrap_or(false);
    files_match && listed
}

/// Writes the add-in and lists it in publish.xml. WPS keeps a trust record
/// with a signature of the add-in's files; it is dropped when the files
/// change (a stale record makes WPS skip the add-in silently) or when the
/// user re-enables the add-in, so WPS asks again instead.
fn install_into(jsaddons: &Path, token: &str, reenable: bool) -> Result<(), String> {
    fs::create_dir_all(jsaddons).map_err(|error| access_error(error, jsaddons))?;
    let folder = jsaddons.join(addin_folder_name());
    let mut files_changed = false;
    for (name, bytes) in addin_files(token) {
        let path = folder.join(&name);
        if fs::read(&path).map(|actual| actual == bytes).unwrap_or(false) {
            continue;
        }
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent).map_err(|error| access_error(error, parent))?;
        }
        fs::write(&path, &bytes).map_err(|error| access_error(error, &path))?;
        files_changed = true;
    }
    // Older bootstrap versions live in their own name_version folders.
    if let Ok(entries) = fs::read_dir(jsaddons) {
        for entry in entries.flatten() {
            let name = entry.file_name().to_string_lossy().to_string();
            if name.starts_with(&format!("{ADDIN_NAME}_")) && name != addin_folder_name() {
                let _ = fs::remove_dir_all(entry.path());
            }
        }
    }
    let publish = jsaddons.join("publish.xml");
    let existing = fs::read_to_string(&publish).ok();
    let mut relisted = false;
    if let Some(updated) = updated_publish_xml(existing.as_deref(), reenable) {
        fs::write(&publish, updated).map_err(|error| access_error(error, &publish))?;
        relisted = true;
    }
    if files_changed || (reenable && relisted) {
        let auth = jsaddons.join("authaddin.json");
        if let Some(updated) = fs::read_to_string(&auth).ok().as_deref().and_then(authaddin_without_us) {
            fs::write(&auth, updated).map_err(|error| access_error(error, &auth))?;
        }
    }
    Ok(())
}

fn uninstall_from(jsaddons: &Path) -> Result<(), String> {
    if let Ok(entries) = fs::read_dir(jsaddons) {
        for entry in entries.flatten() {
            if entry.file_name().to_string_lossy().starts_with(&format!("{ADDIN_NAME}_")) {
                fs::remove_dir_all(entry.path()).map_err(|error| access_error(error, &entry.path()))?;
            }
        }
    }
    let publish = jsaddons.join("publish.xml");
    if let Ok(text) = fs::read_to_string(&publish) {
        if text.lines().any(is_our_publish_line) {
            fs::write(&publish, publish_xml_without_us(&text)).map_err(|error| access_error(error, &publish))?;
        }
    }
    let auth = jsaddons.join("authaddin.json");
    if let Some(updated) = fs::read_to_string(&auth).ok().as_deref().and_then(authaddin_without_us) {
        fs::write(&auth, updated).map_err(|error| access_error(error, &auth))?;
    }
    Ok(())
}

/// Runs container I/O on its own thread: on some macOS versions the first
/// access to another app's container waits for the user to answer a prompt.
fn with_timeout<T: Send + 'static>(work: impl FnOnce() -> T + Send + 'static) -> Result<T, String> {
    let (sender, receiver) = mpsc::sync_channel(1);
    std::thread::Builder::new()
        .name("visualtex-wps-addin".to_string())
        .spawn(move || {
            let _ = sender.send(work());
        })
        .map_err(|error| format!("Unable to start the WPS add-in installer: {error}"))?;
    receiver.recv_timeout(FILESYSTEM_TIMEOUT).map_err(|_| {
        format!("{WPS_FOLDER_ACCESS_DENIED}: macOS did not answer in time. If it asks whether VisualTeX may access data from other apps, allow it and try again.")
    })
}

fn settings_dir(state: &OfficeCompanionState) -> PathBuf {
    state.paths.root.join("wps")
}

fn auto_install_disabled(state: &OfficeCompanionState) -> bool {
    settings_dir(state).join(AUTO_INSTALL_DISABLED_FILE).is_file()
}

fn set_auto_install_disabled(state: &OfficeCompanionState, disabled: bool) {
    let path = settings_dir(state).join(AUTO_INSTALL_DISABLED_FILE);
    if disabled {
        let _ = fs::create_dir_all(settings_dir(state));
        let _ = fs::write(path, b"1");
    } else {
        let _ = fs::remove_file(path);
    }
}

pub(crate) fn addin_log_path(state: &OfficeCompanionState) -> PathBuf {
    settings_dir(state).join("addin.log")
}

fn status_of(state: &OfficeCompanionState, errors: &[(String, String)]) -> Result<WpsAddinStatus, String> {
    let token = state.install_token.as_str().to_string();
    let roots = jsaddons_roots(&home()?);
    let containers = with_timeout(move || {
        roots
            .into_iter()
            .map(|(bundle_id, jsaddons)| WpsContainerStatus {
                installed: jsaddons.join(addin_folder_name()).join("main.js").is_file(),
                current: addin_is_current(&jsaddons, &token),
                addin_path: jsaddons.join(addin_folder_name()).display().to_string(),
                bundle_id,
                error: None,
            })
            .collect::<Vec<_>>()
    })?
    .into_iter()
    .map(|mut container| {
        container.error = errors
            .iter()
            .find(|(bundle_id, _)| *bundle_id == container.bundle_id)
            .map(|(_, error)| error.clone());
        container
    })
    .collect::<Vec<_>>();
    let bridge = bridge::status();
    Ok(WpsAddinStatus {
        wps_installed: WPS_APPLICATIONS.iter().any(|path| Path::new(path).is_dir()) || !containers.is_empty(),
        containers,
        bridge_running: bridge.running,
        bridge_error: bridge.error,
        port: WPS_BRIDGE_PORT,
        auto_install_disabled: auto_install_disabled(state),
    })
}

fn install_all(state: &OfficeCompanionState, reenable: bool) -> Vec<(String, String)> {
    let token = state.install_token.as_str().to_string();
    let roots = match home() {
        Ok(home) => jsaddons_roots(&home),
        Err(error) => return vec![(String::new(), error)],
    };
    roots
        .into_iter()
        .filter_map(|(bundle_id, jsaddons)| {
            let token = token.clone();
            let result = with_timeout(move || install_into(&jsaddons, &token, reenable)).and_then(|result| result);
            result.err().map(|error| (bundle_id, error))
        })
        .collect()
}

/// At startup: install or update the add-in in every WPS container, unless
/// the user uninstalled it. Failures only show up in Settings.
pub(crate) fn auto_install_in_background(state: OfficeCompanionState) {
    crate::spawn_background_task("visualtex-wps-addin-install", move || {
        std::thread::sleep(Duration::from_secs(4));
        if auto_install_disabled(&state) {
            return;
        }
        for (bundle_id, error) in install_all(&state, false) {
            eprintln!("VisualTeX could not install the WPS add-in for {bundle_id}: {error}");
        }
    });
}

#[tauri::command]
pub async fn get_wps_addin_status(state: tauri::State<'_, OfficeCompanionState>) -> Result<WpsAddinStatus, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || status_of(&state, &[]))
        .await
        .map_err(|error| format!("WPS add-in status failed: {error}"))?
}

#[tauri::command]
pub async fn install_wps_addin(state: tauri::State<'_, OfficeCompanionState>) -> Result<WpsAddinStatus, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || {
        set_auto_install_disabled(&state, false);
        let errors = install_all(&state, true);
        if let Some((_, error)) = errors.first() {
            return Err(error.clone());
        }
        status_of(&state, &errors)
    })
    .await
    .map_err(|error| format!("WPS add-in installation failed: {error}"))?
}

#[tauri::command]
pub async fn uninstall_wps_addin(state: tauri::State<'_, OfficeCompanionState>) -> Result<WpsAddinStatus, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || {
        set_auto_install_disabled(&state, true);
        for (_, jsaddons) in jsaddons_roots(&home()?) {
            with_timeout(move || uninstall_from(&jsaddons))??;
        }
        status_of(&state, &[])
    })
    .await
    .map_err(|error| format!("WPS add-in removal failed: {error}"))?
}

#[tauri::command]
pub fn open_wps_office() -> Result<(), String> {
    for bundle_id in WPS_BUNDLE_IDS {
        let status = std::process::Command::new("/usr/bin/open")
            .args(["-b", bundle_id])
            .status()
            .map_err(|error| format!("Unable to open WPS Office: {error}"))?;
        if status.success() {
            return Ok(());
        }
    }
    Err("WPS Office is not installed".to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn publish_xml_keeps_other_add_ins_and_replaces_ours() {
        let existing = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<jsplugins>\n  <jsplugin name=\"Zotero\" enable=\"enable_dev\" url=\"file://\" type=\"wps\" version=\"2\"/>\n  <jsplugin name=\"VisualTeX\" enable=\"enable\" url=\"file://\" type=\"wps\" version=\"0\"/>\n</jsplugins>";
        let merged = updated_publish_xml(Some(existing), false).unwrap();
        assert!(merged.contains("name=\"Zotero\""));
        assert_eq!(merged.matches("name=\"VisualTeX\"").count(), 1);
        assert!(merged.contains(&publish_entry()));
        assert!(merged.trim_end().ends_with("</jsplugins>"));
        assert!(updated_publish_xml(Some(&merged), false).is_none());
        let fresh = updated_publish_xml(None, false).unwrap();
        assert!(fresh.starts_with("<?xml"));
        assert!(fresh.contains(&publish_entry()));
        let removed = publish_xml_without_us(&merged);
        assert!(!removed.contains("VisualTeX"));
        assert!(removed.contains("Zotero"));
    }

    #[test]
    fn the_entry_wps_rewrote_after_the_user_allowed_it_is_kept() {
        let allowed = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<jsplugins>\n    <jsplugin url=\"file://\" name=\"VisualTeX\" version=\"1\" type=\"wps\" enable=\"enable_dev\"/>\n</jsplugins>\n";
        assert!(updated_publish_xml(Some(allowed), false).is_none());
        assert!(updated_publish_xml(Some(allowed), true).is_none());
        let turned_off = allowed.replace("enable_dev", "disable");
        assert!(updated_publish_xml(Some(&turned_off), false).is_none());
        let relisted = updated_publish_xml(Some(&turned_off), true).unwrap();
        assert!(relisted.contains(&publish_entry()) && !relisted.contains("disable"));
    }

    #[test]
    fn only_our_trust_records_are_removed() {
        let auth = r#"{"wps":{"aa":{"name":"VisualTeX","md5":"x"},"bb":{"name":"Zotero","md5":"y"},"namelist":"aa;bb"},"wpp":{"cc":{"name":"Other"},"namelist":"cc"}}"#;
        let updated: Value = serde_json::from_str(&authaddin_without_us(auth).unwrap()).unwrap();
        assert!(updated["wps"].get("aa").is_none());
        assert_eq!(updated["wps"]["bb"]["name"], "Zotero");
        assert_eq!(updated["wps"]["namelist"], "bb");
        assert_eq!(updated["wpp"]["namelist"], "cc");
        assert!(authaddin_without_us(r#"{"wps":{"bb":{"name":"Zotero"},"namelist":"bb"}}"#).is_none());
        assert!(authaddin_without_us("not json").is_none());
    }

    #[test]
    fn install_is_idempotent_and_uninstall_leaves_other_add_ins() {
        let temp = TempDir::new().unwrap();
        let jsaddons = temp.path().join("jsaddons");
        fs::create_dir_all(jsaddons.join("Zotero_2")).unwrap();
        fs::create_dir_all(jsaddons.join("VisualTeX_0")).unwrap();
        fs::write(jsaddons.join("publish.xml"), "<jsplugins>\n  <jsplugin name=\"Zotero\" url=\"file://\" type=\"wps\" version=\"2\"/>\n</jsplugins>").unwrap();
        fs::write(jsaddons.join("authaddin.json"), r#"{"wps":{"aa":{"name":"VisualTeX"},"bb":{"name":"Zotero"},"namelist":"aa;bb"}}"#).unwrap();
        install_into(&jsaddons, "token-1", false).unwrap();
        assert!(addin_is_current(&jsaddons, "token-1"));
        assert!(!addin_is_current(&jsaddons, "token-2"));
        assert!(!jsaddons.join("VisualTeX_0").exists());
        assert!(jsaddons.join("Zotero_2").exists());
        let config = fs::read_to_string(jsaddons.join("VisualTeX_1/config.js")).unwrap();
        assert!(config.contains("\"token\":\"token-1\"") && config.contains("43129"));
        let auth = fs::read_to_string(jsaddons.join("authaddin.json")).unwrap();
        assert!(!auth.contains("VisualTeX") && auth.contains("Zotero"));

        // The user allowed the add-in: WPS rewrote our entry and trusts the
        // files. Neither a restart nor the Install button may undo that.
        let publish = fs::read_to_string(jsaddons.join("publish.xml")).unwrap().replace(
            &publish_entry(),
            r#"<jsplugin url="file://" name="VisualTeX" version="1" type="wps" enable="enable_dev"/>"#,
        );
        fs::write(jsaddons.join("publish.xml"), &publish).unwrap();
        let trusted = r#"{"wps":{"cc":{"name":"VisualTeX","md5":"x"},"bb":{"name":"Zotero"},"namelist":"cc;bb"}}"#;
        fs::write(jsaddons.join("authaddin.json"), trusted).unwrap();
        for reenable in [false, true] {
            install_into(&jsaddons, "token-1", reenable).unwrap();
            assert_eq!(fs::read_to_string(jsaddons.join("publish.xml")).unwrap(), publish);
            assert_eq!(fs::read_to_string(jsaddons.join("authaddin.json")).unwrap(), trusted);
        }
        assert!(addin_is_current(&jsaddons, "token-1"));

        // New files (another token) invalidate WPS's trust record.
        install_into(&jsaddons, "token-2", false).unwrap();
        assert!(!fs::read_to_string(jsaddons.join("authaddin.json")).unwrap().contains("VisualTeX"));
        assert_eq!(fs::read_to_string(jsaddons.join("publish.xml")).unwrap(), publish);

        uninstall_from(&jsaddons).unwrap();
        assert!(!jsaddons.join("VisualTeX_1").exists());
        let publish = fs::read_to_string(jsaddons.join("publish.xml")).unwrap();
        assert!(!publish.contains("VisualTeX") && publish.contains("Zotero"));
    }
}
