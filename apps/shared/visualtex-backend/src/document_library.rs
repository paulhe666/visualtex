//! The document library's commands (document mode of the main window). The documents are files (document_files.rs):
//! Markdown files in the documents folder — `Documents/VisualTeX` unless the user chose another one, which is
//! remembered in the app's data folder — and files elsewhere, edited where they are. The app's data folder
//! (`documents/`) also holds the records of files elsewhere and the library of earlier versions, whose documents are
//! converted once (the page does that: it reads and writes documents) and stay there.
use crate::document_files::{self as files, DocumentContent, DocumentEntry, FileState, LegacyDocument, Library, OpenedDocument, Places, SaveOutcome};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use tauri::{AppHandle, Emitter, Manager};

const CONFIG_FILE: &str = "library.json";
const MIGRATION_FILE: &str = "migration.json";
const DEFAULT_FOLDER: &str = "VisualTeX";
/// Tells the page that files are waiting to be opened (`document_pending_files`).
pub const OPEN_FILES_EVENT: &str = "document-open-files";

#[derive(Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Config {
    /// the documents folder the user chose (none: the default)
    #[serde(default)]
    root: Option<String>,
}

fn data_dir(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app.path().app_data_dir().map_err(|e| e.to_string())?.join("documents"))
}

fn default_root(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app.path().document_dir().map_err(|e| e.to_string())?.join(DEFAULT_FOLDER))
}

fn places(app: &AppHandle) -> Result<Places, String> {
    let data = data_dir(app)?;
    let config: Config = fs::read_to_string(data.join(CONFIG_FILE)).ok().and_then(|text| serde_json::from_str(&text).ok()).unwrap_or_default();
    let root = match config.root.map(PathBuf::from).filter(|root| root.is_absolute()) {
        Some(root) => root,
        None => default_root(app)?,
    };
    Ok(Places { root, data })
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Moved {
    library: Library,
    id: String,
}

#[tauri::command]
pub fn document_library_list(app: AppHandle) -> Result<Library, String> {
    files::list(&places(&app)?)
}

#[tauri::command]
pub fn document_library_read(app: AppHandle, id: String) -> Result<OpenedDocument, String> {
    files::read(&places(&app)?, &id)
}

#[tauri::command]
pub fn document_library_check(app: AppHandle, id: String) -> Result<FileState, String> {
    files::check(&places(&app)?, &id)
}

#[tauri::command]
pub fn document_library_create(app: AppHandle, folder: Option<String>, name: String, extension: String, content: DocumentContent) -> Result<DocumentEntry, String> {
    files::create(&places(&app)?, folder.as_deref(), &name, &extension, &content)
}

#[tauri::command]
pub fn document_library_save(app: AppHandle, id: String, content: DocumentContent, base: Option<String>, force: bool) -> Result<SaveOutcome, String> {
    files::save(&places(&app)?, &id, &content, base.as_deref(), force)
}

#[tauri::command]
pub fn document_library_save_as(app: AppHandle, path: String, content: DocumentContent) -> Result<DocumentEntry, String> {
    files::save_as(&places(&app)?, Path::new(&path), &content)
}

/// A file anywhere on the disk as a document: of the documents folder when it is in it, else a recent file.
#[tauri::command]
pub fn document_library_open_file(app: AppHandle, path: String) -> Result<DocumentEntry, String> {
    files::open_outside(&places(&app)?, Path::new(&path))
}

#[tauri::command]
pub fn document_library_forget(app: AppHandle, id: String) -> Result<Library, String> {
    files::forget(&places(&app)?, &id)
}

#[tauri::command]
pub fn document_library_rename(app: AppHandle, id: String, title: String) -> Result<Moved, String> {
    let (library, id) = files::rename(&places(&app)?, &id, &title)?;
    Ok(Moved { library, id })
}

#[tauri::command]
pub fn document_library_duplicate(app: AppHandle, id: String, title: String) -> Result<DocumentEntry, String> {
    files::duplicate(&places(&app)?, &id, &title)
}

#[tauri::command]
pub fn document_library_move(app: AppHandle, id: String, folder: Option<String>) -> Result<Moved, String> {
    let (library, id) = files::move_to(&places(&app)?, &id, folder.as_deref())?;
    Ok(Moved { library, id })
}

#[tauri::command]
pub fn document_library_trash(app: AppHandle, id: String) -> Result<Library, String> {
    files::trash(&places(&app)?, &id)
}

#[tauri::command]
pub fn document_library_restore(app: AppHandle, id: String) -> Result<Library, String> {
    files::restore(&places(&app)?, &id)
}

#[tauri::command]
pub fn document_library_delete(app: AppHandle, id: String) -> Result<Library, String> {
    files::delete(&places(&app)?, &id)
}

#[tauri::command]
pub fn document_library_empty_trash(app: AppHandle) -> Result<Library, String> {
    files::empty_trash(&places(&app)?)
}

#[tauri::command]
pub fn document_library_create_folder(app: AppHandle, name: String) -> Result<Moved, String> {
    let (library, id) = files::create_folder(&places(&app)?, &name)?;
    Ok(Moved { library, id })
}

#[tauri::command]
pub fn document_library_rename_folder(app: AppHandle, id: String, name: String) -> Result<Library, String> {
    files::rename_folder(&places(&app)?, &id, &name)
}

#[tauri::command]
pub fn document_library_delete_folder(app: AppHandle, id: String) -> Result<Library, String> {
    files::delete_folder(&places(&app)?, &id)
}

#[tauri::command]
pub fn document_library_search(app: AppHandle, query: String) -> Result<Vec<String>, String> {
    files::search(&places(&app)?, &query)
}

/// Another documents folder (none: the default one again). The documents stay where they are: the folder chosen is
/// read as it is.
#[tauri::command]
pub fn document_library_set_root(app: AppHandle, path: Option<String>) -> Result<Library, String> {
    let data = data_dir(&app)?;
    let root = match path.map(PathBuf::from) {
        Some(root) if root.is_absolute() => {
            fs::create_dir_all(&root).map_err(|e| format!("Unable to use the folder: {e}"))?;
            Some(root.to_string_lossy().into_owned())
        }
        Some(_) => return Err("Invalid folder".to_string()),
        None => None,
    };
    let bytes = serde_json::to_vec_pretty(&Config { root }).map_err(|e| e.to_string())?;
    files::write_atomically(&data.join(CONFIG_FILE), &bytes)?;
    files::list(&places(&app)?)
}

/// Shows the document's file in Finder / Explorer.
#[tauri::command]
pub fn document_library_reveal(app: AppHandle, id: String) -> Result<(), String> {
    let path = files::document_path(&places(&app)?, &id)?;
    if !path.is_file() {
        return Err("The document's file is gone".to_string());
    }
    // Explorer reads /select,"path" itself: the argument goes as written (a user name may hold spaces)
    #[cfg(target_os = "windows")]
    let spawned = {
        use std::os::windows::process::CommandExt;
        std::process::Command::new("explorer.exe").raw_arg(format!("/select,\"{}\"", path.display())).spawn()
    };
    #[cfg(not(target_os = "windows"))]
    let spawned = std::process::Command::new("/usr/bin/open").arg("-R").arg(&path).spawn();
    spawned.map(|_| ()).map_err(|e| format!("Unable to show the file: {e}"))
}

// ------------------------------------------------------------------------- the library of earlier versions

/// The documents earlier versions kept in the app's data folder, for their conversion to Markdown files.
#[tauri::command]
pub fn document_library_legacy_list(app: AppHandle) -> Result<Vec<LegacyDocument>, String> {
    Ok(files::legacy_list(&data_dir(&app)?))
}

#[tauri::command]
pub fn document_library_legacy_read(app: AppHandle, id: String) -> Result<String, String> {
    files::legacy_read(&data_dir(&app)?, &id)
}

/// What the conversion of the earlier library did (none: it has not run).
#[tauri::command]
pub fn document_library_migration(app: AppHandle) -> Result<Option<Value>, String> {
    Ok(fs::read_to_string(data_dir(&app)?.join(MIGRATION_FILE)).ok().and_then(|text| serde_json::from_str(&text).ok()))
}

#[tauri::command]
pub fn document_library_set_migration(app: AppHandle, state: Value) -> Result<(), String> {
    let bytes = serde_json::to_vec_pretty(&state).map_err(|e| e.to_string())?;
    files::write_atomically(&data_dir(&app)?.join(MIGRATION_FILE), &bytes)
}

// ------------------------------------------------------------------------------- files the system hands over

static PENDING_FILES: Mutex<Vec<String>> = Mutex::new(Vec::new());

/// Documents the system asked the app to open (a double click, "Open with", a second start with a path): kept until
/// the page takes them, and the page is told. With no window yet the page finds them when it starts.
pub fn queue_files(app: &AppHandle, paths: Vec<PathBuf>) {
    if paths.is_empty() {
        return;
    }
    {
        let mut pending = PENDING_FILES.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
        for path in paths {
            let text = path.to_string_lossy().into_owned();
            if !pending.contains(&text) {
                pending.push(text);
            }
        }
    }
    let _ = app.emit(OPEN_FILES_EVENT, ());
}

/// The files waiting to be opened, taken by the page (each is handed over once).
#[tauri::command]
pub fn document_pending_files() -> Vec<String> {
    std::mem::take(&mut *PENDING_FILES.lock().unwrap_or_else(|poisoned| poisoned.into_inner()))
}
