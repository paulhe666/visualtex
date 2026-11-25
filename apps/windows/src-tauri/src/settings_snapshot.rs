//! A copy of the app's settings in the settings folder (settings-snapshot.json, next to VisualTeX's other
//! settings in %APPDATA%\com.visualtex.studio). Most settings live in the main window's localStorage, inside
//! the WebView folder (%LOCALAPPDATA%\com.visualtex.studio\EBWebView), which is mostly cache and which an
//! uninstall deletes. The main window saves its localStorage here whenever it changes
//! (src/desktop/settingsSnapshot.ts); when VisualTeX starts and a saved key is missing from localStorage,
//! the main window puts it back before the page's own scripts run. Keys the page has are never overwritten.

use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::{AppHandle, Manager};

pub(crate) const FILE_NAME: &str = "settings-snapshot.json";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const SCHEMA_VERSION: u32 = 1;

#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SnapshotFile {
    schema_version: u32,
    saved_at_ms: u128,
    entries: BTreeMap<String, String>,
}

fn snapshot_path(app: &AppHandle) -> Option<PathBuf> {
    app.path().app_data_dir().ok().map(|folder| folder.join(FILE_NAME))
}

fn now_ms() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|elapsed| elapsed.as_millis())
        .unwrap_or_default()
}

/// Writes the page's localStorage (a JSON object of strings) as the snapshot, replacing the old one in
/// one step so a crash cannot leave half a file.
pub(crate) fn write_snapshot(path: &Path, snapshot: &str) -> Result<(), String> {
    if snapshot.len() > MAX_BYTES {
        return Err(format!("The settings are larger than {} MB", MAX_BYTES / 1024 / 1024));
    }
    let entries: BTreeMap<String, String> = serde_json::from_str(snapshot)
        .map_err(|error| format!("The settings are not a list of saved values: {error}"))?;
    let file = SnapshotFile {
        schema_version: SCHEMA_VERSION,
        saved_at_ms: now_ms(),
        entries,
    };
    let bytes = serde_json::to_vec_pretty(&file).map_err(|error| error.to_string())?;
    if let Some(folder) = path.parent() {
        fs::create_dir_all(folder)
            .map_err(|error| format!("Unable to create {}: {error}", folder.display()))?;
    }
    let temporary = path.with_extension("json.tmp");
    fs::write(&temporary, &bytes)
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    fs::rename(&temporary, path).map_err(|error| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to replace {}: {error}", path.display())
    })
}

/// The script that puts saved settings back into an empty or partial localStorage, or None without a
/// usable snapshot. It runs only on the app's own pages (not on an external page the window might show).
pub(crate) fn restore_script_from(path: &Path) -> Option<String> {
    let text = fs::read_to_string(path).ok()?;
    let file: SnapshotFile = serde_json::from_str(&text).ok()?;
    if file.schema_version != SCHEMA_VERSION || file.entries.is_empty() {
        return None;
    }
    let saved = serde_json::to_string(&file.entries).ok()?;
    Some(format!(
        "(function(){{try{{var place=window.location;\
if(!(place.protocol==='tauri:'||place.hostname==='tauri.localhost'))return;\
var saved={saved};var storage=window.localStorage;\
for(var key in saved){{if(Object.prototype.hasOwnProperty.call(saved,key)&&storage.getItem(key)===null){{storage.setItem(key,saved[key]);}}}}\
}}catch(error){{}}}})();"
    ))
}

pub(crate) fn restore_script(app: &AppHandle) -> Option<String> {
    restore_script_from(&snapshot_path(app)?)
}

#[tauri::command]
pub(crate) fn save_settings_snapshot(app: AppHandle, snapshot: String) -> Result<(), String> {
    let path = snapshot_path(&app).ok_or("VisualTeX's settings folder is unavailable")?;
    write_snapshot(&path, &snapshot)
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn a_saved_snapshot_becomes_a_script_that_fills_only_missing_keys() {
        let temp = TempDir::new().unwrap();
        let path = temp.path().join("settings").join(FILE_NAME);
        write_snapshot(&path, r#"{"visualtex-editor":"{\"theme\":\"dark\"}","custom tiles":"[\"a\\u2028b\"]"}"#)
            .unwrap();
        let script = restore_script_from(&path).unwrap();
        assert!(script.contains(r#""visualtex-editor":"{\"theme\":\"dark\"}""#));
        assert!(script.contains("storage.getItem(key)===null"));
        assert!(script.contains("place.hostname==='tauri.localhost'"));
        assert!(!path.with_extension("json.tmp").exists());
    }

    #[test]
    fn bad_or_oversized_settings_are_refused_and_keep_the_old_snapshot() {
        let temp = TempDir::new().unwrap();
        let path = temp.path().join(FILE_NAME);
        write_snapshot(&path, r#"{"a":"1"}"#).unwrap();
        assert!(write_snapshot(&path, r#"{"a":1}"#).is_err());
        assert!(write_snapshot(&path, "not json").is_err());
        let huge = format!("{{\"a\":\"{}\"}}", "x".repeat(MAX_BYTES));
        assert!(write_snapshot(&path, &huge).is_err());
        assert!(restore_script_from(&path).unwrap().contains(r#"{"a":"1"}"#));
    }

    #[test]
    fn no_script_without_a_snapshot_or_with_an_empty_one() {
        let temp = TempDir::new().unwrap();
        let path = temp.path().join(FILE_NAME);
        assert!(restore_script_from(&path).is_none());
        write_snapshot(&path, "{}").unwrap();
        assert!(restore_script_from(&path).is_none());
        fs::write(&path, "{broken").unwrap();
        assert!(restore_script_from(&path).is_none());
    }
}
