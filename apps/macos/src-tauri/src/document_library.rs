//! The document library of the main window's document mode: every document is a `.vtxdoc` file in the app's data
//! folder (`documents/`), and `index.json` beside them keeps titles, folders, times and the recently deleted ones.
//! Files that appear without an index entry (copied in by hand) are taken in; entries whose file is gone are dropped.
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::{AppHandle, Manager};

const INDEX_FILE: &str = "index.json";
const EXTENSION: &str = "vtxdoc";
/// Recently deleted documents are kept this long.
const TRASH_KEEP_MS: u64 = 30 * 24 * 60 * 60 * 1000;
const MAX_DOCUMENT_BYTES: usize = 200 * 1024 * 1024;

static LOCK: Mutex<()> = Mutex::new(());

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentEntry {
    pub id: String,
    /// The shown name: the user's, or (until renamed) the document's first heading or line.
    pub title: String,
    /// Renamed by the user: the title no longer follows the text.
    #[serde(default)]
    pub title_set: bool,
    #[serde(default)]
    pub folder: Option<String>,
    pub created: u64,
    pub modified: u64,
    #[serde(default)]
    pub trashed: Option<u64>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentFolder {
    pub id: String,
    pub name: String,
    pub created: u64,
}

#[derive(Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Index {
    #[serde(default)]
    folders: Vec<DocumentFolder>,
    #[serde(default)]
    documents: Vec<DocumentEntry>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Library {
    pub root: String,
    pub folders: Vec<DocumentFolder>,
    pub documents: Vec<DocumentEntry>,
}

fn now() -> u64 {
    SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_millis() as u64).unwrap_or(0)
}

fn new_id() -> String {
    uuid::Uuid::new_v4().simple().to_string()
}

fn valid_id(id: &str) -> Result<&str, String> {
    if !id.is_empty() && id.len() <= 64 && id.bytes().all(|b| b.is_ascii_alphanumeric() || b == b'-' || b == b'_') {
        Ok(id)
    } else {
        Err("Invalid document id".to_string())
    }
}

fn clean_name(name: &str) -> String {
    name.chars().filter(|c| !c.is_control()).collect::<String>().trim().chars().take(200).collect()
}

fn write_atomically(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let parent = path.parent().ok_or("No parent directory")?;
    fs::create_dir_all(parent).map_err(|e| format!("Unable to create the documents folder: {e}"))?;
    let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("document");
    let temporary = parent.join(format!(".{name}.{}.tmp", std::process::id()));
    fs::write(&temporary, bytes).map_err(|e| format!("Unable to write the document: {e}"))?;
    fs::rename(&temporary, path).map_err(|e| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to save the document: {e}")
    })
}

fn document_path(root: &Path, id: &str) -> PathBuf {
    root.join(format!("{id}.{EXTENSION}"))
}

fn file_time(path: &Path) -> u64 {
    fs::metadata(path)
        .and_then(|m| m.modified())
        .ok()
        .and_then(|t| t.duration_since(UNIX_EPOCH).ok())
        .map(|d| d.as_millis() as u64)
        .unwrap_or_else(now)
}

/// The index, reconciled with the files on disk; the long-deleted ones are removed for good.
fn load(root: &Path) -> Result<Index, String> {
    fs::create_dir_all(root).map_err(|e| format!("Unable to create the documents folder: {e}"))?;
    let mut index: Index = fs::read_to_string(root.join(INDEX_FILE))
        .ok()
        .and_then(|text| serde_json::from_str(&text).ok())
        .unwrap_or_default();
    let mut changed = false;
    let before = index.documents.len();
    index.documents.retain(|doc| valid_id(&doc.id).is_ok() && document_path(root, &doc.id).is_file());
    changed |= index.documents.len() != before;
    let cutoff = now().saturating_sub(TRASH_KEEP_MS);
    let expired: Vec<String> = index
        .documents
        .iter()
        .filter(|doc| doc.trashed.is_some_and(|time| time < cutoff))
        .map(|doc| doc.id.clone())
        .collect();
    for id in &expired {
        let _ = fs::remove_file(document_path(root, id));
    }
    if !expired.is_empty() {
        index.documents.retain(|doc| !expired.contains(&doc.id));
        changed = true;
    }
    let folders: Vec<String> = index.folders.iter().map(|f| f.id.clone()).collect();
    for doc in index.documents.iter_mut() {
        if doc.folder.as_ref().is_some_and(|folder| !folders.contains(folder)) {
            doc.folder = None;
            changed = true;
        }
    }
    if let Ok(entries) = fs::read_dir(root) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().and_then(|e| e.to_str()) != Some(EXTENSION) {
                continue;
            }
            let Some(stem) = path.file_stem().and_then(|s| s.to_str()) else { continue };
            if stem.starts_with('.') || index.documents.iter().any(|doc| doc.id == stem) {
                continue;
            }
            // A file put here by hand: an id-like name is kept, any other is renamed to a new id.
            let (id, title) = if valid_id(stem).is_ok() && stem.len() >= 16 {
                (stem.to_string(), String::new())
            } else {
                let id = new_id();
                if fs::rename(&path, document_path(root, &id)).is_err() {
                    continue;
                }
                (id, clean_name(stem))
            };
            let time = file_time(&document_path(root, &id));
            index.documents.push(DocumentEntry {
                title_set: !title.is_empty(),
                title,
                id,
                folder: None,
                created: time,
                modified: time,
                trashed: None,
            });
            changed = true;
        }
    }
    if changed {
        store(root, &index)?;
    }
    Ok(index)
}

fn store(root: &Path, index: &Index) -> Result<(), String> {
    let bytes = serde_json::to_vec_pretty(index).map_err(|e| e.to_string())?;
    write_atomically(&root.join(INDEX_FILE), &bytes)
}

fn library(root: &Path, index: Index) -> Library {
    let mut documents = index.documents;
    documents.sort_by(|a, b| b.modified.cmp(&a.modified));
    Library { root: root.to_string_lossy().into_owned(), folders: index.folders, documents }
}

fn entry_mut<'a>(index: &'a mut Index, id: &str) -> Result<&'a mut DocumentEntry, String> {
    index.documents.iter_mut().find(|doc| doc.id == id).ok_or_else(|| "The document is not in the library".to_string())
}

fn check_size(text: &str) -> Result<(), String> {
    if text.len() > MAX_DOCUMENT_BYTES {
        Err("The document is too large".to_string())
    } else {
        Ok(())
    }
}

pub fn list(root: &Path) -> Result<Library, String> {
    let _guard = LOCK.lock().unwrap_or_else(|p| p.into_inner());
    Ok(library(root, load(root)?))
}

pub fn read(root: &Path, id: &str) -> Result<String, String> {
    let id = valid_id(id)?;
    fs::read_to_string(document_path(root, id)).map_err(|e| format!("Unable to read the document: {e}"))
}

pub fn create(root: &Path, folder: Option<String>, text: &str, title: Option<String>) -> Result<DocumentEntry, String> {
    check_size(text)?;
    let _guard = LOCK.lock().unwrap_or_else(|p| p.into_inner());
    let mut index = load(root)?;
    let folder = folder.filter(|id| index.folders.iter().any(|f| &f.id == id));
    let id = new_id();
    write_atomically(&document_path(root, &id), text.as_bytes())?;
    let time = now();
    let title = title.map(|t| clean_name(&t)).unwrap_or_default();
    let entry = DocumentEntry { id, title_set: !title.is_empty(), title, folder, created: time, modified: time, trashed: None };
    index.documents.push(entry.clone());
    store(root, &index)?;
    Ok(entry)
}

/// Saves the text; `title` (the document's first heading or line) names it while the user has not.
pub fn save(root: &Path, id: &str, text: &str, title: &str) -> Result<DocumentEntry, String> {
    let id = valid_id(id)?;
    check_size(text)?;
    let _guard = LOCK.lock().unwrap_or_else(|p| p.into_inner());
    let mut index = load(root)?;
    let entry = entry_mut(&mut index, id)?;
    write_atomically(&document_path(root, id), text.as_bytes())?;
    entry.modified = now();
    if !entry.title_set {
        entry.title = clean_name(title);
    }
    let entry = entry.clone();
    store(root, &index)?;
    Ok(entry)
}

fn update(root: &Path, id: &str, change: impl FnOnce(&mut Index, &str) -> Result<(), String>) -> Result<Library, String> {
    let id = valid_id(id)?;
    let _guard = LOCK.lock().unwrap_or_else(|p| p.into_inner());
    let mut index = load(root)?;
    change(&mut index, id)?;
    store(root, &index)?;
    Ok(library(root, index))
}

/// An empty name gives the title back to the document's text.
pub fn rename(root: &Path, id: &str, title: &str) -> Result<Library, String> {
    update(root, id, |index, id| {
        let entry = entry_mut(index, id)?;
        entry.title = clean_name(title);
        entry.title_set = !entry.title.is_empty();
        Ok(())
    })
}

pub fn duplicate(root: &Path, id: &str, title: &str) -> Result<DocumentEntry, String> {
    let id = valid_id(id)?;
    let text = read(root, id)?;
    let folder = {
        let _guard = LOCK.lock().unwrap_or_else(|p| p.into_inner());
        load(root)?.documents.iter().find(|doc| doc.id == id).and_then(|doc| doc.folder.clone())
    };
    create(root, folder, &text, Some(title.to_string()))
}

pub fn move_to(root: &Path, id: &str, folder: Option<String>) -> Result<Library, String> {
    update(root, id, |index, id| {
        let folder = folder.filter(|folder| index.folders.iter().any(|f| &f.id == folder));
        entry_mut(index, id)?.folder = folder;
        Ok(())
    })
}

pub fn trash(root: &Path, id: &str) -> Result<Library, String> {
    update(root, id, |index, id| {
        entry_mut(index, id)?.trashed = Some(now());
        Ok(())
    })
}

pub fn restore(root: &Path, id: &str) -> Result<Library, String> {
    update(root, id, |index, id| {
        entry_mut(index, id)?.trashed = None;
        Ok(())
    })
}

pub fn delete(root: &Path, id: &str) -> Result<Library, String> {
    update(root, id, |index, id| {
        index.documents.retain(|doc| doc.id != id);
        let _ = fs::remove_file(document_path(root, id));
        Ok(())
    })
}

pub fn empty_trash(root: &Path) -> Result<Library, String> {
    let _guard = LOCK.lock().unwrap_or_else(|p| p.into_inner());
    let mut index = load(root)?;
    for doc in index.documents.iter().filter(|doc| doc.trashed.is_some()) {
        let _ = fs::remove_file(document_path(root, &doc.id));
    }
    index.documents.retain(|doc| doc.trashed.is_none());
    store(root, &index)?;
    Ok(library(root, index))
}

pub fn create_folder(root: &Path, name: &str) -> Result<(Library, String), String> {
    let name = clean_name(name);
    if name.is_empty() {
        return Err("The folder needs a name".to_string());
    }
    let _guard = LOCK.lock().unwrap_or_else(|p| p.into_inner());
    let mut index = load(root)?;
    let id = new_id();
    index.folders.push(DocumentFolder { id: id.clone(), name, created: now() });
    store(root, &index)?;
    Ok((library(root, index), id))
}

pub fn rename_folder(root: &Path, id: &str, name: &str) -> Result<Library, String> {
    let name = clean_name(name);
    if name.is_empty() {
        return Err("The folder needs a name".to_string());
    }
    update(root, id, |index, id| {
        let folder = index.folders.iter_mut().find(|f| f.id == id).ok_or("The folder is not in the library")?;
        folder.name = name;
        Ok(())
    })
}

/// The folder goes; its documents stay, outside any folder.
pub fn delete_folder(root: &Path, id: &str) -> Result<Library, String> {
    update(root, id, |index, id| {
        index.folders.retain(|f| f.id != id);
        for doc in index.documents.iter_mut().filter(|doc| doc.folder.as_deref() == Some(id)) {
            doc.folder = None;
        }
        Ok(())
    })
}

fn collect_text(value: &Value, out: &mut String) {
    match value {
        Value::Object(map) => {
            for (key, item) in map {
                match (key.as_str(), item) {
                    ("text" | "latex" | "alt" | "title" | "href", Value::String(text)) => {
                        out.push_str(text);
                        out.push('\n');
                    }
                    _ => collect_text(item, out),
                }
            }
        }
        Value::Array(items) => items.iter().for_each(|item| collect_text(item, out)),
        _ => {}
    }
}

/// The ids of the documents (not deleted) whose title or text contains every word of the query.
pub fn search(root: &Path, query: &str) -> Result<Vec<String>, String> {
    let words: Vec<String> = query.split_whitespace().map(|w| w.to_lowercase()).collect();
    if words.is_empty() {
        return Ok(Vec::new());
    }
    let index = {
        let _guard = LOCK.lock().unwrap_or_else(|p| p.into_inner());
        load(root)?
    };
    let mut found = Vec::new();
    for doc in index.documents.iter().filter(|doc| doc.trashed.is_none()) {
        let mut text = doc.title.clone();
        text.push('\n');
        if let Ok(source) = fs::read_to_string(document_path(root, &doc.id)) {
            match serde_json::from_str::<Value>(&source) {
                Ok(value) => collect_text(&value, &mut text),
                Err(_) => text.push_str(&source),
            }
        }
        let text = text.to_lowercase();
        if words.iter().all(|word| text.contains(word.as_str())) {
            found.push(doc.id.clone());
        }
    }
    Ok(found)
}

fn root(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app.path().app_data_dir().map_err(|e| e.to_string())?.join("documents"))
}

#[tauri::command]
pub fn document_library_list(app: AppHandle) -> Result<Library, String> {
    list(&root(&app)?)
}

#[tauri::command]
pub fn document_library_read(app: AppHandle, id: String) -> Result<String, String> {
    read(&root(&app)?, &id)
}

#[tauri::command]
pub fn document_library_create(
    app: AppHandle,
    folder: Option<String>,
    text: String,
    title: Option<String>,
) -> Result<DocumentEntry, String> {
    create(&root(&app)?, folder, &text, title)
}

#[tauri::command]
pub fn document_library_save(app: AppHandle, id: String, text: String, title: String) -> Result<DocumentEntry, String> {
    save(&root(&app)?, &id, &text, &title)
}

#[tauri::command]
pub fn document_library_rename(app: AppHandle, id: String, title: String) -> Result<Library, String> {
    rename(&root(&app)?, &id, &title)
}

#[tauri::command]
pub fn document_library_duplicate(app: AppHandle, id: String, title: String) -> Result<DocumentEntry, String> {
    duplicate(&root(&app)?, &id, &title)
}

#[tauri::command]
pub fn document_library_move(app: AppHandle, id: String, folder: Option<String>) -> Result<Library, String> {
    move_to(&root(&app)?, &id, folder)
}

#[tauri::command]
pub fn document_library_trash(app: AppHandle, id: String) -> Result<Library, String> {
    trash(&root(&app)?, &id)
}

#[tauri::command]
pub fn document_library_restore(app: AppHandle, id: String) -> Result<Library, String> {
    restore(&root(&app)?, &id)
}

#[tauri::command]
pub fn document_library_delete(app: AppHandle, id: String) -> Result<Library, String> {
    delete(&root(&app)?, &id)
}

#[tauri::command]
pub fn document_library_empty_trash(app: AppHandle) -> Result<Library, String> {
    empty_trash(&root(&app)?)
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FolderCreated {
    library: Library,
    id: String,
}

#[tauri::command]
pub fn document_library_create_folder(app: AppHandle, name: String) -> Result<FolderCreated, String> {
    let (library, id) = create_folder(&root(&app)?, &name)?;
    Ok(FolderCreated { library, id })
}

#[tauri::command]
pub fn document_library_rename_folder(app: AppHandle, id: String, name: String) -> Result<Library, String> {
    rename_folder(&root(&app)?, &id, &name)
}

#[tauri::command]
pub fn document_library_delete_folder(app: AppHandle, id: String) -> Result<Library, String> {
    delete_folder(&root(&app)?, &id)
}

#[tauri::command]
pub fn document_library_search(app: AppHandle, query: String) -> Result<Vec<String>, String> {
    search(&root(&app)?, &query)
}

/// Shows the document's file in Finder.
#[tauri::command]
pub fn document_library_reveal(app: AppHandle, id: String) -> Result<(), String> {
    let path = document_path(&root(&app)?, valid_id(&id)?);
    if !path.is_file() {
        return Err("The document is not in the library".to_string());
    }
    std::process::Command::new("/usr/bin/open")
        .arg("-R")
        .arg(&path)
        .spawn()
        .map(|_| ())
        .map_err(|e| format!("Unable to show the file: {e}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const DOC: &str = r#"{"version":1,"doc":{"type":"doc","content":[{"type":"heading","content":[{"type":"text","text":"Fourier 变换"}]},{"type":"math_inline","attrs":{"latex":"\\hat f(\\xi)"}}]}}"#;

    #[test]
    fn create_save_rename_and_search() {
        let dir = tempfile::TempDir::new().unwrap();
        let root = dir.path();
        assert!(list(root).unwrap().documents.is_empty());
        let a = create(root, None, DOC, None).unwrap();
        assert_eq!(a.title, "");
        let saved = save(root, &a.id, DOC, "Fourier 变换").unwrap();
        assert_eq!(saved.title, "Fourier 变换");
        assert_eq!(read(root, &a.id).unwrap(), DOC);
        rename(root, &a.id, "笔记").unwrap();
        assert_eq!(save(root, &a.id, DOC, "Other").unwrap().title, "笔记");
        rename(root, &a.id, "  ").unwrap();
        assert_eq!(save(root, &a.id, DOC, "Other").unwrap().title, "Other");
        assert_eq!(search(root, "变换").unwrap(), vec![a.id.clone()]);
        assert_eq!(search(root, "hat F").unwrap(), vec![a.id.clone()]);
        assert!(search(root, "missing").unwrap().is_empty());
        assert!(read(root, "../index").is_err());
    }

    #[test]
    fn folders_trash_and_files_on_disk() {
        let dir = tempfile::TempDir::new().unwrap();
        let root = dir.path();
        let (_, folder) = create_folder(root, "课程").unwrap();
        let a = create(root, Some(folder.clone()), DOC, Some("A".into())).unwrap();
        assert_eq!(a.folder.as_deref(), Some(folder.as_str()));
        let b = duplicate(root, &a.id, "A 副本").unwrap();
        assert_eq!(b.folder.as_deref(), Some(folder.as_str()));
        let lib = delete_folder(root, &folder).unwrap();
        assert!(lib.folders.is_empty() && lib.documents.iter().all(|d| d.folder.is_none()));
        trash(root, &a.id).unwrap();
        assert!(search(root, "变换").unwrap() == vec![b.id.clone()]);
        restore(root, &a.id).unwrap();
        trash(root, &b.id).unwrap();
        let lib = empty_trash(root).unwrap();
        assert_eq!(lib.documents.len(), 1);
        assert!(!document_path(root, &b.id).exists());
        // a file copied in by hand is taken in, named after its file
        fs::write(root.join("讲义.vtxdoc"), DOC).unwrap();
        let lib = list(root).unwrap();
        assert_eq!(lib.documents.len(), 2);
        assert!(lib.documents.iter().any(|d| d.title == "讲义" && d.title_set));
        // its file gone: the entry goes
        fs::remove_file(document_path(root, &a.id)).unwrap();
        assert_eq!(list(root).unwrap().documents.len(), 1);
    }
}
