//! Installs the VisualTeX add-ins into WPS Office for Mac: one for WPS Writer
//! and one for WPS Presentation.
//!
//! WPS loads JS add-ins listed in `publish.xml` inside its sandbox container
//! (`.kingsoft/wps/jsaddons`); each entry names the kind of document it is
//! for. Each add-in installed there is a small bootstrap that loads its logic
//! from the running VisualTeX (bridge.rs), so it rarely changes; WPS asks the
//! user to trust an add-in again whenever its files do.

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

/// An add-in VisualTeX installs into WPS.
struct WpsAddin {
    /// The name in publish.xml and authaddin.json.
    name: &'static str,
    /// The bootstrap's version; with the name it names the add-in folder (WPS
    /// requires `name_version`). Bump it only when the installed files change.
    version: &'static str,
    /// The kind of document WPS loads the add-in for: "wps" (Writer) or "wpp"
    /// (Presentation).
    kind: &'static str,
}

const WRITER: WpsAddin = WpsAddin { name: "VisualTeX", version: "1", kind: "wps" };
const PRESENTATION: WpsAddin = WpsAddin { name: "VisualTeXSlides", version: "1", kind: "wpp" };
const ADDINS: [&WpsAddin; 2] = [&WRITER, &PRESENTATION];
const WPS_BUNDLE_IDS: [&str; 2] = ["com.kingsoft.wpsoffice.mac", "com.kingsoft.wpsoffice.mac.global"];
const WPS_APPLICATIONS: [&str; 2] = ["/Applications/WPS Office.app", "/Applications/wpsoffice.app"];
const AUTO_INSTALL_DISABLED_FILE: &str = "auto-install-disabled";
const FILESYSTEM_TIMEOUT: Duration = Duration::from_secs(15);

const INDEX_HTML: &str = include_str!("../../../../office/wps/addin/index.html");
const MAIN_JS: &str = include_str!("../../../../office/wps/addin/main.js");
const RIBBON_XML: &str = include_str!("../../../../office/wps/addin/ribbon.xml");
const SLIDES_MAIN_JS: &str = include_str!("../../../../office/wps/addin-wpp/main.js");
const SLIDES_RIBBON_XML: &str = include_str!("../../../../office/wps/addin-wpp/ribbon.xml");
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
/// The Presentation add-in's buttons, drawn like the Writer add-in's.
const SLIDES_IMAGES: [(&str, &[u8]); 7] = [
    ("vtxToLatex.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/document-import.svg")),
    ("vtxOmmlInline.svg", include_bytes!("../../../../office/wps/addin/images/vtxOmmlInline.svg")),
    ("vtxOmmlDisplay.svg", include_bytes!("../../../../office/wps/addin/images/vtxOmmlDisplay.svg")),
    ("vtxNewFormula.svg", include_bytes!("../../../../office/wps/addin/images/vtxImageDisplay.svg")),
    ("vtxEditSelected.svg", include_bytes!("../../../../office/wps/addin/images/vtxEditSelected.svg")),
    ("vtxDeleteSelected.svg", include_bytes!("../../../../office/wps/addin-wpp/images/vtxDeleteSelected.svg")),
    ("vtxOpen.svg", include_bytes!("../../../../office/macos-offline/word/ribbon-icons/open-visualtex.svg")),
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

impl WpsAddin {
    fn folder_name(&self) -> String {
        format!("{}_{}", self.name, self.version)
    }

    /// Folders of this add-in's other versions start with this.
    fn folder_prefix(&self) -> String {
        format!("{}_", self.name)
    }

    fn publish_entry(&self) -> String {
        format!(
            r#"<jsplugin name="{}" enable="enable" url="file://" type="{}" version="{}"/>"#,
            self.name, self.kind, self.version
        )
    }

    fn files(&self, token: &str) -> Vec<(String, Vec<u8>)> {
        let (main_js, ribbon_xml, images): (&str, &str, &[(&str, &[u8])]) = if self.kind == "wpp" {
            (SLIDES_MAIN_JS, SLIDES_RIBBON_XML, &SLIDES_IMAGES)
        } else {
            (MAIN_JS, RIBBON_XML, &IMAGES)
        };
        let mut files = vec![
            ("index.html".to_string(), INDEX_HTML.as_bytes().to_vec()),
            ("main.js".to_string(), main_js.as_bytes().to_vec()),
            ("ribbon.xml".to_string(), ribbon_xml.as_bytes().to_vec()),
            ("config.js".to_string(), config_js(token).into_bytes()),
        ];
        for (name, bytes) in images {
            files.push((format!("images/{name}"), bytes.to_vec()));
        }
        files
    }

    /// `name="VisualTeX"` does not match `name="VisualTeXSlides"`.
    fn is_publish_line(&self, line: &str) -> bool {
        line.contains(&format!(r#"name="{}""#, self.name))
    }

    /// The entry as WPS keeps it. Once the user allows the add-in, WPS
    /// rewrites publish.xml (attribute order, indentation) and turns
    /// enable="enable" into "enable_dev", so the entry is matched by its
    /// attributes, not its text.
    fn publish_line<'a>(&self, text: &'a str) -> Option<&'a str> {
        text.lines().find(|line| self.is_publish_line(line))
    }

    fn publish_line_is_current(&self, line: &str) -> bool {
        let version = format!(r#"version="{}""#, self.version);
        let kind = format!(r#"type="{}""#, self.kind);
        [version.as_str(), r#"url="file://""#, kind.as_str()]
            .iter()
            .all(|attribute| line.contains(attribute))
    }
}

fn config_js(token: &str) -> String {
    let config = serde_json::json!({ "port": WPS_BRIDGE_PORT, "token": token });
    format!("window.VISUALTEX_WPS_CONFIG = {config};\n")
}

fn publish_line_is_enabled(line: &str) -> bool {
    line.contains(r#"enable="enable""#) || line.contains(r#"enable="enable_dev""#)
}

/// `publish.xml` with the add-in's entry added or replaced and other
/// add-ins' entries kept, or None when it already lists this add-in version.
/// Rewriting a listed entry would make WPS ask the user to allow the add-in
/// again, so an entry turned off in WPS is only replaced when `reenable`
/// (Install button).
fn updated_publish_xml(addin: &WpsAddin, existing: Option<&str>, reenable: bool) -> Option<String> {
    let entry = addin.publish_entry();
    match existing.filter(|text| text.contains("</jsplugins>")) {
        Some(text) => {
            if addin.publish_line(text).is_some_and(|line| {
                addin.publish_line_is_current(line) && (!reenable || publish_line_is_enabled(line))
            }) {
                return None;
            }
            let kept: Vec<&str> = text.lines().filter(|line| !addin.is_publish_line(line)).collect();
            Some(kept.join("\n").replacen("</jsplugins>", &format!("  {entry}\n</jsplugins>"), 1))
        }
        None => Some(format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<jsplugins>\n  {entry}\n</jsplugins>\n")),
    }
}

fn publish_xml_without(addin: &WpsAddin, existing: &str) -> String {
    existing
        .lines()
        .filter(|line| !addin.is_publish_line(line))
        .collect::<Vec<_>>()
        .join("\n")
}

/// WPS records each trusted add-in in authaddin.json with a signature of its
/// files; a stale record makes WPS skip the add-in silently. Returns the file
/// without the add-in's records, or None when there is nothing to remove.
fn authaddin_without(addin: &WpsAddin, text: &str) -> Option<String> {
    let mut json: Value = serde_json::from_str(text).ok()?;
    let mut changed = false;
    for (_, group) in json.as_object_mut()?.iter_mut() {
        let Some(group) = group.as_object_mut() else { continue };
        let ours: Vec<String> = group
            .iter()
            .filter(|(_, entry)| entry.get("name").and_then(Value::as_str) == Some(addin.name))
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

fn addin_is_current(addin: &WpsAddin, jsaddons: &Path, token: &str) -> bool {
    let folder = jsaddons.join(addin.folder_name());
    let files_match = addin
        .files(token)
        .iter()
        .all(|(name, bytes)| fs::read(folder.join(name)).map(|actual| &actual == bytes).unwrap_or(false));
    let listed = fs::read_to_string(jsaddons.join("publish.xml"))
        .map(|text| addin.publish_line(&text).is_some_and(|line| addin.publish_line_is_current(line)))
        .unwrap_or(false);
    files_match && listed
}

fn addins_are_current(jsaddons: &Path, token: &str) -> bool {
    ADDINS.iter().all(|addin| addin_is_current(addin, jsaddons, token))
}

/// Writes the add-ins and lists them in publish.xml. WPS keeps a trust record
/// with a signature of each add-in's files; it is dropped when the files
/// change (a stale record makes WPS skip the add-in silently) or when the
/// user re-enables the add-in, so WPS asks again instead.
fn install_into(jsaddons: &Path, token: &str, reenable: bool) -> Result<(), String> {
    fs::create_dir_all(jsaddons).map_err(|error| access_error(error, jsaddons))?;
    for addin in ADDINS {
        install_addin(addin, jsaddons, token, reenable)?;
    }
    Ok(())
}

fn install_addin(addin: &WpsAddin, jsaddons: &Path, token: &str, reenable: bool) -> Result<(), String> {
    let folder = jsaddons.join(addin.folder_name());
    let mut files_changed = false;
    for (name, bytes) in addin.files(token) {
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
            if name.starts_with(&addin.folder_prefix()) && name != addin.folder_name() {
                let _ = fs::remove_dir_all(entry.path());
            }
        }
    }
    let publish = jsaddons.join("publish.xml");
    let existing = fs::read_to_string(&publish).ok();
    let mut relisted = false;
    if let Some(updated) = updated_publish_xml(addin, existing.as_deref(), reenable) {
        fs::write(&publish, updated).map_err(|error| access_error(error, &publish))?;
        relisted = true;
    }
    if files_changed || (reenable && relisted) {
        let auth = jsaddons.join("authaddin.json");
        if let Some(updated) = fs::read_to_string(&auth).ok().as_deref().and_then(|text| authaddin_without(addin, text)) {
            fs::write(&auth, updated).map_err(|error| access_error(error, &auth))?;
        }
    }
    Ok(())
}

fn uninstall_from(jsaddons: &Path) -> Result<(), String> {
    for addin in ADDINS {
        if let Ok(entries) = fs::read_dir(jsaddons) {
            for entry in entries.flatten() {
                if entry.file_name().to_string_lossy().starts_with(&addin.folder_prefix()) {
                    fs::remove_dir_all(entry.path()).map_err(|error| access_error(error, &entry.path()))?;
                }
            }
        }
        let publish = jsaddons.join("publish.xml");
        if let Ok(text) = fs::read_to_string(&publish) {
            if text.lines().any(|line| addin.is_publish_line(line)) {
                fs::write(&publish, publish_xml_without(addin, &text))
                    .map_err(|error| access_error(error, &publish))?;
            }
        }
        let auth = jsaddons.join("authaddin.json");
        if let Some(updated) = fs::read_to_string(&auth).ok().as_deref().and_then(|text| authaddin_without(addin, text)) {
            fs::write(&auth, updated).map_err(|error| access_error(error, &auth))?;
        }
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
                installed: ADDINS
                    .iter()
                    .all(|addin| jsaddons.join(addin.folder_name()).join("main.js").is_file()),
                current: addins_are_current(&jsaddons, &token),
                addin_path: jsaddons.join(WRITER.folder_name()).display().to_string(),
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
/// the user uninstalled it, with the math fonts of its equation font menu.
/// Failures only show up in Settings.
pub(crate) fn auto_install_in_background(app: tauri::AppHandle, state: OfficeCompanionState) {
    crate::spawn_background_task("visualtex-wps-addin-install", move || {
        std::thread::sleep(Duration::from_secs(4));
        if auto_install_disabled(&state) {
            return;
        }
        for (bundle_id, error) in install_all(&state, false) {
            eprintln!("VisualTeX could not install the WPS add-in for {bundle_id}: {error}");
        }
        if wps_present() {
            crate::office::macos_offline_installer::install_math_fonts_with_wps_addin(&app);
        }
    });
}

/// WPS is installed (an app bundle or a container of it).
fn wps_present() -> bool {
    WPS_APPLICATIONS.iter().any(|path| Path::new(path).is_dir())
        || home().map(|home| !jsaddons_roots(&home).is_empty()).unwrap_or(false)
}

#[tauri::command]
pub async fn get_wps_addin_status(state: tauri::State<'_, OfficeCompanionState>) -> Result<WpsAddinStatus, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || status_of(&state, &[]))
        .await
        .map_err(|error| format!("WPS add-in status failed: {error}"))?
}

#[tauri::command]
pub async fn install_wps_addin(
    app: tauri::AppHandle,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<WpsAddinStatus, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || {
        set_auto_install_disabled(&state, false);
        let errors = install_all(&state, true);
        crate::office::macos_offline_installer::install_math_fonts_with_wps_addin(&app);
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
        let merged = updated_publish_xml(&WRITER, Some(existing), false).unwrap();
        assert!(merged.contains("name=\"Zotero\""));
        assert_eq!(merged.matches("name=\"VisualTeX\"").count(), 1);
        assert!(merged.contains(&WRITER.publish_entry()));
        assert!(merged.trim_end().ends_with("</jsplugins>"));
        assert!(updated_publish_xml(&WRITER, Some(&merged), false).is_none());
        let fresh = updated_publish_xml(&WRITER, None, false).unwrap();
        assert!(fresh.starts_with("<?xml"));
        assert!(fresh.contains(&WRITER.publish_entry()));
        let removed = publish_xml_without(&WRITER, &merged);
        assert!(!removed.contains("VisualTeX"));
        assert!(removed.contains("Zotero"));
    }

    #[test]
    fn the_presentation_add_in_has_its_own_entry() {
        assert_eq!(
            PRESENTATION.publish_entry(),
            r#"<jsplugin name="VisualTeXSlides" enable="enable" url="file://" type="wpp" version="1"/>"#
        );
        let writer_only = updated_publish_xml(&WRITER, None, false).unwrap();
        // The Writer entry is not the Presentation add-in's, and the other way round.
        assert!(PRESENTATION.publish_line(&writer_only).is_none());
        let both = updated_publish_xml(&PRESENTATION, Some(&writer_only), false).unwrap();
        assert!(WRITER.publish_line(&both).is_some_and(|line| WRITER.publish_line_is_current(line)));
        assert!(PRESENTATION.publish_line(&both).is_some_and(|line| PRESENTATION.publish_line_is_current(line)));
        assert!(updated_publish_xml(&WRITER, Some(&both), false).is_none());
        // A Writer-typed entry under the Presentation name is not current.
        let wrong_kind = both.replace(r#"type="wpp""#, r#"type="wps""#);
        assert!(updated_publish_xml(&PRESENTATION, Some(&wrong_kind), false).is_some());
        let without_slides = publish_xml_without(&PRESENTATION, &both);
        assert!(without_slides.contains(r#"name="VisualTeX""#) && !without_slides.contains("VisualTeXSlides"));
        assert!(!PRESENTATION.folder_name().starts_with(&WRITER.folder_prefix()));
        let names: Vec<String> = PRESENTATION.files("t").into_iter().map(|(name, _)| name).collect();
        for expected in ["index.html", "main.js", "ribbon.xml", "config.js", "images/vtxNewFormula.svg", "images/vtxDeleteSelected.svg"] {
            assert!(names.iter().any(|name| name == expected), "{expected}");
        }
        let ribbon = String::from_utf8(PRESENTATION.files("t").into_iter().find(|(name, _)| name == "ribbon.xml").unwrap().1).unwrap();
        assert!(ribbon.contains("vtxNewFormula") && ribbon.contains("vtxOmmlDisplay") && !ribbon.contains("vtxConvertNative"));
        for (name, _) in SLIDES_IMAGES {
            let id = name.trim_end_matches(".svg");
            assert!(ribbon.contains(&format!("id=\"{id}\"")), "{id} has a button");
        }
    }

    #[test]
    fn the_entry_wps_rewrote_after_the_user_allowed_it_is_kept() {
        let allowed = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<jsplugins>\n    <jsplugin url=\"file://\" name=\"VisualTeX\" version=\"1\" type=\"wps\" enable=\"enable_dev\"/>\n</jsplugins>\n";
        assert!(updated_publish_xml(&WRITER, Some(allowed), false).is_none());
        assert!(updated_publish_xml(&WRITER, Some(allowed), true).is_none());
        let turned_off = allowed.replace("enable_dev", "disable");
        assert!(updated_publish_xml(&WRITER, Some(&turned_off), false).is_none());
        let relisted = updated_publish_xml(&WRITER, Some(&turned_off), true).unwrap();
        assert!(relisted.contains(&WRITER.publish_entry()) && !relisted.contains("disable"));
    }

    #[test]
    fn only_our_trust_records_are_removed() {
        let auth = r#"{"wps":{"aa":{"name":"VisualTeX","md5":"x"},"bb":{"name":"Zotero","md5":"y"},"namelist":"aa;bb"},"wpp":{"cc":{"name":"Other"},"dd":{"name":"VisualTeXSlides"},"namelist":"cc;dd"}}"#;
        let updated: Value = serde_json::from_str(&authaddin_without(&WRITER, auth).unwrap()).unwrap();
        assert!(updated["wps"].get("aa").is_none());
        assert_eq!(updated["wps"]["bb"]["name"], "Zotero");
        assert_eq!(updated["wps"]["namelist"], "bb");
        assert_eq!(updated["wpp"]["namelist"], "cc;dd");
        let updated: Value = serde_json::from_str(&authaddin_without(&PRESENTATION, auth).unwrap()).unwrap();
        assert!(updated["wpp"].get("dd").is_none());
        assert_eq!(updated["wpp"]["namelist"], "cc");
        assert_eq!(updated["wps"]["namelist"], "aa;bb");
        assert!(authaddin_without(&WRITER, r#"{"wps":{"bb":{"name":"Zotero"},"namelist":"bb"}}"#).is_none());
        assert!(authaddin_without(&WRITER, "not json").is_none());
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
        assert!(addins_are_current(&jsaddons, "token-1"));
        assert!(!addins_are_current(&jsaddons, "token-2"));
        assert!(!jsaddons.join("VisualTeX_0").exists());
        assert!(jsaddons.join("Zotero_2").exists());
        for folder in ["VisualTeX_1", "VisualTeXSlides_1"] {
            let config = fs::read_to_string(jsaddons.join(folder).join("config.js")).unwrap();
            assert!(config.contains("\"token\":\"token-1\"") && config.contains("43129"), "{folder}");
        }
        let slides_main = fs::read_to_string(jsaddons.join("VisualTeXSlides_1/main.js")).unwrap();
        assert!(slides_main.contains("host=wpp"));
        let auth = fs::read_to_string(jsaddons.join("authaddin.json")).unwrap();
        assert!(!auth.contains("VisualTeX") && auth.contains("Zotero"));

        // The user allowed both add-ins: WPS rewrote the entries and trusts
        // the files. Neither a restart nor the Install button may undo that.
        let publish = fs::read_to_string(jsaddons.join("publish.xml"))
            .unwrap()
            .replace(
                &WRITER.publish_entry(),
                r#"<jsplugin url="file://" name="VisualTeX" version="1" type="wps" enable="enable_dev"/>"#,
            )
            .replace(
                &PRESENTATION.publish_entry(),
                r#"<jsplugin url="file://" name="VisualTeXSlides" version="1" type="wpp" enable="enable_dev"/>"#,
            );
        fs::write(jsaddons.join("publish.xml"), &publish).unwrap();
        let trusted = r#"{"wps":{"cc":{"name":"VisualTeX","md5":"x"},"bb":{"name":"Zotero"},"namelist":"cc;bb"},"wpp":{"dd":{"name":"VisualTeXSlides","md5":"y"},"namelist":"dd"}}"#;
        fs::write(jsaddons.join("authaddin.json"), trusted).unwrap();
        for reenable in [false, true] {
            install_into(&jsaddons, "token-1", reenable).unwrap();
            assert_eq!(fs::read_to_string(jsaddons.join("publish.xml")).unwrap(), publish);
            assert_eq!(fs::read_to_string(jsaddons.join("authaddin.json")).unwrap(), trusted);
        }
        assert!(addins_are_current(&jsaddons, "token-1"));

        // New files (another token) invalidate WPS's trust records.
        install_into(&jsaddons, "token-2", false).unwrap();
        assert!(!fs::read_to_string(jsaddons.join("authaddin.json")).unwrap().contains("VisualTeX"));
        assert_eq!(fs::read_to_string(jsaddons.join("publish.xml")).unwrap(), publish);

        uninstall_from(&jsaddons).unwrap();
        assert!(!jsaddons.join("VisualTeX_1").exists());
        assert!(!jsaddons.join("VisualTeXSlides_1").exists());
        let publish = fs::read_to_string(jsaddons.join("publish.xml")).unwrap();
        assert!(!publish.contains("VisualTeX") && publish.contains("Zotero"));
    }

    #[test]
    fn an_existing_writer_install_gains_the_presentation_add_in_and_keeps_its_trust() {
        let temp = TempDir::new().unwrap();
        let jsaddons = temp.path().join("jsaddons");
        fs::create_dir_all(&jsaddons).unwrap();
        install_addin(&WRITER, &jsaddons, "token-1", false).unwrap();
        let publish = fs::read_to_string(jsaddons.join("publish.xml")).unwrap().replace(
            &WRITER.publish_entry(),
            r#"<jsplugin url="file://" name="VisualTeX" version="1" type="wps" enable="enable_dev"/>"#,
        );
        fs::write(jsaddons.join("publish.xml"), &publish).unwrap();
        let trusted = r#"{"wps":{"cc":{"name":"VisualTeX","md5":"x"},"namelist":"cc"}}"#;
        fs::write(jsaddons.join("authaddin.json"), trusted).unwrap();
        install_into(&jsaddons, "token-1", false).unwrap();
        let updated = fs::read_to_string(jsaddons.join("publish.xml")).unwrap();
        assert!(updated.contains(r#"<jsplugin url="file://" name="VisualTeX" version="1" type="wps" enable="enable_dev"/>"#));
        assert!(updated.contains(&PRESENTATION.publish_entry()));
        assert_eq!(fs::read_to_string(jsaddons.join("authaddin.json")).unwrap(), trusted);
        assert!(addins_are_current(&jsaddons, "token-1"));
    }

    /// One ribbon element: its name, boxStyle, size and child elements.
    struct RibbonNode {
        name: String,
        style: String,
        large: bool,
        children: Vec<RibbonNode>,
    }

    fn ribbon_tree(xml: &str) -> RibbonNode {
        use quick_xml::{events::Event, Reader};
        fn node(start: &quick_xml::events::BytesStart) -> RibbonNode {
            let attribute = |key: &str| {
                start
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.as_ref() == key.as_bytes())
                    .map(|attribute| String::from_utf8_lossy(&attribute.value).into_owned())
                    .unwrap_or_default()
            };
            RibbonNode {
                name: String::from_utf8_lossy(start.local_name().as_ref()).into_owned(),
                style: attribute("boxStyle"),
                large: attribute("size") == "large",
                children: Vec::new(),
            }
        }
        let mut reader = Reader::from_str(xml);
        let mut stack = vec![RibbonNode { name: String::new(), style: String::new(), large: false, children: Vec::new() }];
        loop {
            match reader.read_event().unwrap() {
                Event::Start(start) => stack.push(node(&start)),
                Event::Empty(start) => stack.last_mut().unwrap().children.push(node(&start)),
                Event::End(_) => {
                    let done = stack.pop().unwrap();
                    stack.last_mut().unwrap().children.push(done);
                }
                Event::Eof => break,
                _ => {}
            }
        }
        stack.pop().unwrap()
    }

    /// How WPS lays the columns out: it packs consecutive vertical boxes (and small controls beside them) into one
    /// column, and it centres a column's controls between the rows; a box of only size boxes gets tighter rows than
    /// buttons. So in a group, small controls stand in columns of two or three rows, side-by-side columns stand in a
    /// horizontal box with the same number of rows each, and every box has a button or menu row.
    #[test]
    fn the_tabs_keep_wps_columns() {
        fn groups<'a>(node: &'a RibbonNode, out: &mut Vec<&'a RibbonNode>) {
            if node.name == "group" {
                out.push(node);
            }
            for child in &node.children {
                groups(child, out);
            }
        }
        let column_rows = |node: &RibbonNode| {
            let rows = node.children.len();
            (node.name == "box"
                && node.style == "vertical"
                && (2..=3).contains(&rows)
                && node.children.iter().all(|child| child.name != "box"))
            .then_some(rows)
        };
        let buttons = |columns: &[&RibbonNode]| {
            columns.iter().flat_map(|column| column.children.iter()).any(|control| {
                !matches!(control.name.as_str(), "comboBox" | "dropDown" | "editBox" | "labelControl")
            })
        };
        for ribbon in [RIBBON_XML, SLIDES_RIBBON_XML] {
            let tree = ribbon_tree(ribbon);
            let mut found = Vec::new();
            groups(&tree, &mut found);
            assert!(found.len() >= 3);
            for group in found {
                let mut previous_column = false;
                for child in &group.children {
                    let is_column = child.name == "box" && child.style == "vertical";
                    if child.name == "box" && child.style == "horizontal" {
                        let columns: Vec<&RibbonNode> = child.children.iter().collect();
                        let rows: Vec<Option<usize>> = columns.iter().map(|column| column_rows(column)).collect();
                        assert!(columns.len() >= 2 && rows.iter().all(Option::is_some), "columns of two or three rows");
                        assert!(rows.windows(2).all(|pair| pair[0] == pair[1]), "columns side by side have the same rows");
                        assert!(buttons(&columns), "a box of only size boxes gets tighter rows");
                    } else if is_column {
                        assert!(!previous_column, "two columns side by side outside a horizontal box");
                        assert!(column_rows(child).is_some() && buttons(&[child]), "a column of two or three with a button");
                    } else {
                        assert!(child.large, "a small control outside a column");
                    }
                    previous_column = is_column;
                }
            }
        }
    }

    /// A tab's layout, group by group: "L" for a large control, "C<n>" for a column of n small controls. Office for
    /// Mac stacks consecutive small controls three to a column by itself; WPS needs a vertical box for each column.
    fn layout(xml: &str) -> Vec<Vec<String>> {
        fn flush(small: &mut usize, out: &mut Vec<String>) {
            while *small > 0 {
                let rows = (*small).min(3);
                out.push(format!("C{rows}"));
                *small -= rows;
            }
        }
        fn group_layout(group: &RibbonNode, out: &mut Vec<String>) {
            let mut small = 0;
            for child in &group.children {
                match (child.name.as_str(), child.style.as_str()) {
                    ("box", "horizontal") => {
                        flush(&mut small, out);
                        for column in &child.children {
                            let rows = column.children.iter().filter(|row| row.name != "labelControl").count();
                            out.push(format!("C{rows}"));
                        }
                    }
                    ("box", "vertical") => {
                        flush(&mut small, out);
                        let rows = child.children.iter().filter(|row| row.name != "labelControl").count();
                        out.push(format!("C{rows}"));
                    }
                    _ if child.large => {
                        flush(&mut small, out);
                        out.push("L".to_string());
                    }
                    _ => small += 1,
                }
            }
            flush(&mut small, out);
        }
        fn walk(node: &RibbonNode, out: &mut Vec<Vec<String>>) {
            if node.name == "group" {
                let mut group = Vec::new();
                group_layout(node, &mut group);
                out.push(group);
                return;
            }
            for child in &node.children {
                walk(child, out);
            }
        }
        let mut out = Vec::new();
        walk(&ribbon_tree(xml), &mut out);
        out
    }

    /// WPS's tabs are the Office for Mac tabs: the same groups with the same large buttons and columns.
    #[test]
    fn the_tabs_have_the_office_tabs_layout() {
        let word = include_str!("../../../../office/macos-offline/word/customUI14.xml");
        let powerpoint = include_str!("../../../../office/macos-offline/powerpoint/customUI14.xml");
        assert_eq!(layout(RIBBON_XML), layout(word));
        assert_eq!(layout(SLIDES_RIBBON_XML), layout(powerpoint));
        assert_eq!(
            layout(RIBBON_XML),
            vec![
                vec!["L"; 5],
                vec!["C3", "C3", "C3"],
                vec!["L", "L"],
                vec!["C3", "L"],
                vec!["L"],
            ]
        );
    }
}
