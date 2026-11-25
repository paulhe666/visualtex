//! Settings > Keyboard shortcuts > Formula input method, "Switch to English input in formulas" (as on Windows): a
//! formula the shortcut opens is typed in English, the input source an ASCII-capable keyboard layout while it is open
//! and the one it had back when it ends. In VisualTeX's own document, and in the pasting strip without the input
//! method (with it, the input method takes the formula's keys: nothing is switched).

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Mutex, OnceLock};
use tauri::{AppHandle, Manager};

/// Saved beside the shortcuts (hotkey.rs); on unless turned off.
const SETTING_FILE: &str = "anywhere-english-input.json";

static SETTING: OnceLock<AtomicBool> = OnceLock::new();

fn setting_file(app: &AppHandle) -> Option<std::path::PathBuf> {
    app.path().app_data_dir().ok().map(|folder| folder.join(SETTING_FILE))
}

fn setting(app: &AppHandle) -> &'static AtomicBool {
    SETTING.get_or_init(|| {
        let saved = setting_file(app).and_then(|path| std::fs::read_to_string(path).ok());
        AtomicBool::new(saved.map(|text| text.trim() != "false").unwrap_or(true))
    })
}

pub(super) fn enabled(app: &AppHandle) -> bool {
    setting(app).load(Ordering::SeqCst)
}

pub(super) fn set_enabled(app: &AppHandle, on: bool) -> Result<(), String> {
    setting(app).store(on, Ordering::SeqCst);
    let path = setting_file(app).ok_or_else(|| "No settings folder".to_string())?;
    if let Some(folder) = path.parent() {
        std::fs::create_dir_all(folder).map_err(|error| format!("The setting could not be saved: {error}"))?;
    }
    std::fs::write(&path, if on { "true" } else { "false" }).map_err(|error| format!("The setting could not be saved: {error}"))
}

/// The input source the document's formula switched from, to give back.
static DOCUMENT_SOURCE: Mutex<Option<String>> = Mutex::new(None);

/// English for a formula the shortcut opened in the main window's document, when the setting is on and the input
/// source is not an ASCII-capable layout already: true when it was switched (or still is, for an earlier formula).
pub(super) fn select_for_document(app: &AppHandle) -> bool {
    if !enabled(app) {
        return false;
    }
    let Ok(mut source) = DOCUMENT_SOURCE.lock() else { return false };
    if source.is_some() {
        return true;
    }
    *source = super::ime::select_ascii(app);
    super::log(&format!("english input (document): {}", source.as_deref().map(|id| format!("from {id}")).unwrap_or_else(|| "not switched".to_string())));
    source.is_some()
}

/// The input source the document's formula switched from, back.
pub(super) fn restore_for_document(app: &AppHandle) {
    let Some(id) = DOCUMENT_SOURCE.lock().ok().and_then(|mut source| source.take()) else { return };
    super::ime::restore(app, &id);
    super::log(&format!("english input (document): {id} back"));
}
