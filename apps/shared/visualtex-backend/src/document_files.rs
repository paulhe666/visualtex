//! Documents as files: the documents folder is a visible folder of Markdown files (`.md`; a `.vtxdoc` of an earlier
//! version is listed too), its folders are its subfolders. What VisualTeX remembers about a document besides its
//! text (the document's memory, a JSON record made by the editor) is kept outside the file: for a document of the
//! folder in the hidden `.visualtex` subfolder under the document's relative path, for a file opened from elsewhere
//! in the app's data folder, found again by the file's path or, after the file was moved, by its content.
//! Deleted documents go to the hidden `.trash` subfolder with their record and pictures, and for good after 30 days.
//!
//! A document's id is its path: relative to the documents folder with "/" (`Notes/Proof.md`), or `file:` and the
//! absolute path of a file elsewhere.
use base64::Engine;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::HashSet;
use std::fs;
use std::path::{Component, Path, PathBuf};
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};

const MEMORY_DIR: &str = ".visualtex";
const TRASH_DIR: &str = ".trash";
const ASSETS_DIR: &str = "assets";
const OUTSIDE_INDEX: &str = "outside.json";
const OUTSIDE_MEMORY_DIR: &str = "memory";
/// Recently deleted documents are kept this long.
const TRASH_KEEP_MS: u64 = 30 * 24 * 60 * 60 * 1000;
const MAX_DOCUMENT_BYTES: usize = 200 * 1024 * 1024;
/// A documents folder is read this deep and no further than this many documents (a folder chosen by mistake).
const MAX_DEPTH: usize = 8;
const MAX_DOCUMENTS: usize = 5000;
const MAX_RECENT: usize = 30;
const OUTSIDE_PREFIX: &str = "file:";

static LOCK: Mutex<()> = Mutex::new(());

fn lock() -> std::sync::MutexGuard<'static, ()> {
    LOCK.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

/// Where documents and what is remembered about them live.
#[derive(Debug, Clone)]
pub struct Places {
    /// the documents folder
    pub root: PathBuf,
    /// the app's data folder for documents: the records of files outside the documents folder
    pub data: PathBuf,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentEntry {
    pub id: String,
    /// The file's name without its extension.
    pub title: String,
    /// The subfolder of the documents folder it is in (relative, "/" separated).
    pub folder: Option<String>,
    pub created: u64,
    pub modified: u64,
    pub trashed: Option<u64>,
    /// The file.
    pub path: String,
    /// "markdown", or "vtxdoc" for a document of an earlier version.
    pub kind: String,
    /// A file outside the documents folder, edited where it is.
    pub outside: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentFolder {
    /// Its path relative to the documents folder.
    pub id: String,
    pub name: String,
    pub created: u64,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Library {
    pub root: String,
    pub folders: Vec<DocumentFolder>,
    /// The documents folder's documents and the recently deleted ones, newest first.
    pub documents: Vec<DocumentEntry>,
    /// Files opened from elsewhere, the last opened first.
    pub recent: Vec<DocumentEntry>,
}

/// A document read: its bytes, its memory record and what its next save is checked against.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OpenedDocument {
    pub entry: DocumentEntry,
    pub base64: String,
    pub memory: Option<String>,
    /// SHA-256 of the file's bytes.
    pub fingerprint: String,
    /// The folder the file is in (its pictures' paths start there).
    pub dir: String,
    /// The files in `assets` beside it.
    pub assets: Vec<String>,
}

/// A picture file to write beside a document: its path as the document's text names it.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AssetFile {
    pub path: String,
    pub base64: String,
}

/// What a save writes: the file's bytes, the memory record (none: the document needs none) and new pictures.
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentContent {
    pub base64: String,
    #[serde(default)]
    pub memory: Option<String>,
    #[serde(default)]
    pub assets: Vec<AssetFile>,
    /// The modification time the file gets (a converted document keeps the time of the one it came from).
    #[serde(default)]
    pub modified: Option<u64>,
}

/// A save's outcome: written, or left alone because the file is no longer what the editor read.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase", tag = "status")]
pub enum SaveOutcome {
    Saved { entry: DocumentEntry, fingerprint: String },
    Changed { fingerprint: String },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FileState {
    pub exists: bool,
    pub fingerprint: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct TrashInfo {
    /// where it was, relative to the documents folder
    path: String,
    trashed: u64,
}

/// What is remembered about a file outside the documents folder.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OutsideFile {
    path: String,
    /// SHA-256 of its bytes when it was last read or saved here: finds the file again after it was moved or renamed
    fingerprint: String,
    /// its memory record's file in the data folder
    #[serde(default)]
    record: Option<String>,
    opened: u64,
    /// taken out of the recent files (its record stays: opened again, it is listed again)
    #[serde(default)]
    unlisted: bool,
}

#[derive(Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OutsideIndex {
    #[serde(default)]
    files: Vec<OutsideFile>,
}

fn now() -> u64 {
    SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_millis() as u64).unwrap_or(0)
}

fn millis(time: std::io::Result<SystemTime>) -> Option<u64> {
    time.ok().and_then(|t| t.duration_since(UNIX_EPOCH).ok()).map(|d| d.as_millis() as u64)
}

pub fn fingerprint(bytes: &[u8]) -> String {
    Sha256::digest(bytes).iter().map(|byte| format!("{byte:02x}")).collect()
}

fn decode(base64: &str) -> Result<Vec<u8>, String> {
    let bytes = base64::engine::general_purpose::STANDARD.decode(base64).map_err(|_| "Invalid document data".to_string())?;
    if bytes.len() > MAX_DOCUMENT_BYTES {
        return Err("The document is too large".to_string());
    }
    Ok(bytes)
}

fn encode(bytes: &[u8]) -> String {
    base64::engine::general_purpose::STANDARD.encode(bytes)
}

/// Writes through a temporary file beside the target, so a reader never sees half a file. A link is written through to
/// the file it points at (it stays a link), a file keeps its permissions, and a read-only file is not written.
pub fn write_atomically(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let link = fs::symlink_metadata(path).map(|metadata| metadata.file_type().is_symlink()).unwrap_or(false);
    let resolved = if link { fs::canonicalize(path).map_err(|e| format!("Unable to save the document: {e}"))? } else { path.to_path_buf() };
    let path = resolved.as_path();
    let permissions = fs::metadata(path).ok().map(|metadata| metadata.permissions());
    if permissions.as_ref().is_some_and(|permissions| permissions.readonly()) {
        return Err("The file is read-only".to_string());
    }
    let parent = path.parent().ok_or("No parent directory")?;
    fs::create_dir_all(parent).map_err(|e| format!("Unable to create the folder: {e}"))?;
    let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("document");
    let temporary = parent.join(format!(".{name}.{}.tmp", std::process::id()));
    fs::write(&temporary, bytes).map_err(|e| format!("Unable to write the document: {e}"))?;
    if let Some(permissions) = permissions {
        let _ = fs::set_permissions(&temporary, permissions);
    }
    fs::rename(&temporary, path).map_err(|e| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to save the document: {e}")
    })
}

/// A folder the file managers do not show: a leading dot is enough elsewhere, Windows wants the attribute.
fn hidden_dir(path: &Path) -> Result<(), String> {
    let existed = path.is_dir();
    fs::create_dir_all(path).map_err(|e| format!("Unable to create the folder: {e}"))?;
    #[cfg(not(target_os = "windows"))]
    let _ = existed;
    #[cfg(target_os = "windows")]
    if !existed {
        use std::os::windows::ffi::OsStrExt;
        use windows::core::PCWSTR;
        use windows::Win32::Storage::FileSystem::{SetFileAttributesW, FILE_ATTRIBUTE_HIDDEN};
        let wide: Vec<u16> = path.as_os_str().encode_wide().chain(std::iter::once(0)).collect();
        // SAFETY: `wide` is a NUL-terminated path that lives through the call.
        let _ = unsafe { SetFileAttributesW(PCWSTR(wide.as_ptr()), FILE_ATTRIBUTE_HIDDEN) };
    }
    Ok(())
}

/// A name as a file name: nothing a file system refuses, not a device name, not ending in a dot or space.
pub fn file_name(name: &str) -> String {
    let cleaned: String = name
        .chars()
        .map(|c| if c.is_control() || matches!(c, '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|') { '-' } else { c })
        .collect();
    let trimmed: String = cleaned.trim().trim_end_matches(['.', ' ']).trim_start_matches('.').chars().take(120).collect();
    let trimmed = trimmed.trim().to_string();
    let stem = trimmed.split('.').next().unwrap_or("").to_ascii_uppercase();
    let device = matches!(stem.as_str(), "CON" | "PRN" | "AUX" | "NUL")
        || (stem.len() == 4 && (stem.starts_with("COM") || stem.starts_with("LPT")) && stem.as_bytes()[3].is_ascii_digit());
    if device {
        format!("_{trimmed}")
    } else {
        trimmed
    }
}

fn extension_of(path: &Path) -> Option<&'static str> {
    match path.extension().and_then(|e| e.to_str()).map(|e| e.to_ascii_lowercase()).as_deref() {
        Some("md") => Some("md"),
        Some("markdown") => Some("markdown"),
        Some("vtxdoc") => Some("vtxdoc"),
        _ => None,
    }
}

fn kind_of(path: &Path) -> &'static str {
    if extension_of(path) == Some("vtxdoc") {
        "vtxdoc"
    } else {
        "markdown"
    }
}

/// `<stem>.<extension>` in `dir`, or the first of `<stem> 2`, `<stem> 3`… that is free.
fn free_path(dir: &Path, stem: &str, extension: &str) -> PathBuf {
    let mut path = dir.join(format!("{stem}.{extension}"));
    let mut n = 2;
    while path.exists() {
        path = dir.join(format!("{stem} {n}.{extension}"));
        n += 1;
    }
    path
}

fn relative_id(root: &Path, path: &Path) -> Option<String> {
    let relative = path.strip_prefix(root).ok()?;
    let parts: Vec<String> = relative.components().map(|c| c.as_os_str().to_string_lossy().into_owned()).collect();
    if parts.is_empty() {
        None
    } else {
        Some(parts.join("/"))
    }
}

/// A relative path of the documents folder as its parts: no "..", nothing absolute, nothing hidden but the trash.
fn relative_parts(id: &str) -> Result<Vec<&str>, String> {
    let parts: Vec<&str> = id.split('/').collect();
    let invalid = || "Invalid document path".to_string();
    if id.is_empty() || id.len() > 1024 || id.contains('\\') || id.contains('\0') {
        return Err(invalid());
    }
    for (index, part) in parts.iter().enumerate() {
        let hidden = part.starts_with('.') && !(index == 0 && *part == TRASH_DIR);
        if part.is_empty() || *part == "." || *part == ".." || hidden || part.contains(':') {
            return Err(invalid());
        }
    }
    Ok(parts)
}

#[derive(Debug, Clone, PartialEq)]
enum Place {
    /// a document of the documents folder
    Library,
    /// a recently deleted one (`.trash/<item>/<file>`)
    Trash,
    /// a file elsewhere
    Outside,
}

struct Located {
    path: PathBuf,
    place: Place,
}

fn locate(places: &Places, id: &str) -> Result<Located, String> {
    if let Some(outside) = id.strip_prefix(OUTSIDE_PREFIX) {
        let path = PathBuf::from(outside);
        if !path.is_absolute() || extension_of(&path).is_none() {
            return Err("Invalid document path".to_string());
        }
        return Ok(Located { path, place: Place::Outside });
    }
    let parts = relative_parts(id)?;
    let mut path = places.root.clone();
    for part in &parts {
        path.push(part);
    }
    if extension_of(&path).is_none() {
        return Err("Invalid document path".to_string());
    }
    let place = if parts[0] == TRASH_DIR {
        if parts.len() != 3 {
            return Err("Invalid document path".to_string());
        }
        Place::Trash
    } else {
        Place::Library
    };
    Ok(Located { path, place })
}

/// The file of the document `id`.
pub fn document_path(places: &Places, id: &str) -> Result<PathBuf, String> {
    Ok(locate(places, id)?.path)
}

/// Whether a path names a document by its extension (.md, .markdown, .vtxdoc).
pub fn is_document_path(path: &Path) -> bool {
    extension_of(path).is_some()
}

/// The documents among a start's arguments (a double click and "Open with" pass the file's path): existing files
/// with a document's extension, a relative path taken from `cwd`. The program's own path and the flags are not.
pub fn document_arguments<I, S>(args: I, cwd: Option<&Path>) -> Vec<PathBuf>
where
    I: IntoIterator<Item = S>,
    S: AsRef<str>,
{
    let mut paths: Vec<PathBuf> = Vec::new();
    for arg in args.into_iter().skip(1) {
        let arg = arg.as_ref();
        if arg.is_empty() || arg.starts_with('-') {
            continue;
        }
        let given = PathBuf::from(arg);
        let path = match (given.is_absolute(), cwd) {
            (true, _) => given,
            (false, Some(cwd)) => cwd.join(given),
            (false, None) => continue,
        };
        if is_document_path(&path) && path.is_file() && !paths.contains(&path) {
            paths.push(path);
        }
    }
    paths
}

fn entry_for(places: &Places, path: &Path, trashed: Option<u64>) -> Option<DocumentEntry> {
    let metadata = fs::metadata(path).ok()?;
    let modified = millis(metadata.modified()).unwrap_or_else(now);
    let created = millis(metadata.created()).unwrap_or(modified).min(modified);
    let title = path.file_stem()?.to_string_lossy().into_owned();
    let (id, folder, outside) = match relative_id(&places.root, path) {
        Some(id) => {
            let folder = if trashed.is_some() { None } else { id.rfind('/').map(|at| id[..at].to_string()) };
            (id, folder, false)
        }
        None => (format!("{OUTSIDE_PREFIX}{}", path.to_string_lossy()), None, true),
    };
    Some(DocumentEntry { id, title, folder, created, modified, trashed, path: path.to_string_lossy().into_owned(), kind: kind_of(path).to_string(), outside })
}

// ----------------------------------------------------------------------------------------------- memory records

/// Where a document of the documents folder keeps its memory: `.visualtex/<relative path>.json`.
fn library_memory_path(places: &Places, path: &Path) -> Option<PathBuf> {
    let relative = path.strip_prefix(&places.root).ok()?;
    let mut record = places.root.join(MEMORY_DIR).join(relative).into_os_string();
    record.push(".json");
    Some(PathBuf::from(record))
}

fn trash_memory_path(path: &Path) -> PathBuf {
    path.with_file_name("memory.json")
}

fn load_outside(places: &Places) -> OutsideIndex {
    fs::read_to_string(places.data.join(OUTSIDE_INDEX)).ok().and_then(|text| serde_json::from_str(&text).ok()).unwrap_or_default()
}

fn store_outside(places: &Places, index: &OutsideIndex) -> Result<(), String> {
    let bytes = serde_json::to_vec_pretty(index).map_err(|e| e.to_string())?;
    write_atomically(&places.data.join(OUTSIDE_INDEX), &bytes)
}

fn outside_record_path(places: &Places, record: &str) -> PathBuf {
    places.data.join(OUTSIDE_MEMORY_DIR).join(record)
}

/// The index entry of a file elsewhere: by its path; or, for a file not known under its path, the entry whose file
/// had these bytes (the file was moved or renamed: the entry follows it; copied: the copy gets the record too).
fn outside_entry<'a>(places: &Places, index: &'a mut OutsideIndex, path: &Path, print: &str) -> &'a mut OutsideFile {
    let text = path.to_string_lossy().into_owned();
    if let Some(at) = index.files.iter().position(|file| file.path == text) {
        return &mut index.files[at];
    }
    if let Some(at) = index.files.iter().position(|file| file.fingerprint == print) {
        if !Path::new(&index.files[at].path).exists() {
            index.files[at].path = text;
            return &mut index.files[at];
        }
        let mut copy = index.files[at].clone();
        copy.path = text;
        if let Some(record) = &copy.record {
            let name = format!("{}.json", uuid::Uuid::new_v4().simple());
            copy.record = fs::copy(outside_record_path(places, record), outside_record_path(places, &name)).ok().map(|_| name);
        }
        index.files.push(copy);
        return index.files.last_mut().unwrap();
    }
    index.files.push(OutsideFile { path: text, fingerprint: print.to_string(), record: None, opened: now(), unlisted: false });
    index.files.last_mut().unwrap()
}

fn read_memory(places: &Places, located: &Located, print: &str) -> Option<String> {
    match located.place {
        Place::Library => fs::read_to_string(library_memory_path(places, &located.path)?).ok(),
        Place::Trash => fs::read_to_string(trash_memory_path(&located.path)).ok(),
        Place::Outside => {
            let mut index = load_outside(places);
            let entry = outside_entry(places, &mut index, &located.path, print);
            entry.opened = now();
            entry.fingerprint = print.to_string();
            entry.unlisted = false;
            let memory = entry.record.as_ref().and_then(|record| fs::read_to_string(outside_record_path(places, record)).ok());
            let _ = store_outside(places, &index);
            memory
        }
    }
}

fn write_memory(places: &Places, located: &Located, memory: Option<&str>, print: &str) -> Result<(), String> {
    match located.place {
        Place::Library | Place::Trash => {
            let record = if located.place == Place::Library {
                library_memory_path(places, &located.path).ok_or("Invalid document path")?
            } else {
                trash_memory_path(&located.path)
            };
            match memory {
                Some(text) => {
                    hidden_dir(&places.root.join(MEMORY_DIR))?;
                    write_atomically(&record, text.as_bytes())
                }
                None => {
                    let _ = fs::remove_file(&record);
                    Ok(())
                }
            }
        }
        Place::Outside => {
            let mut index = load_outside(places);
            let entry = outside_entry(places, &mut index, &located.path, print);
            entry.fingerprint = print.to_string();
            match memory {
                Some(text) => {
                    let name = entry.record.clone().unwrap_or_else(|| format!("{}.json", uuid::Uuid::new_v4().simple()));
                    write_atomically(&outside_record_path(places, &name), text.as_bytes())?;
                    entry.record = Some(name);
                }
                None => {
                    if let Some(record) = entry.record.take() {
                        let _ = fs::remove_file(outside_record_path(places, &record));
                    }
                }
            }
            store_outside(places, &index)
        }
    }
}

/// A document's memory record goes with it to `to` (both in the documents folder).
fn move_library_memory(places: &Places, from: &Path, to: &Path) {
    let (Some(old), Some(new)) = (library_memory_path(places, from), library_memory_path(places, to)) else { return };
    if !old.is_file() {
        return;
    }
    if let Some(parent) = new.parent() {
        let _ = fs::create_dir_all(parent);
    }
    let _ = fs::rename(&old, &new);
}

// ------------------------------------------------------------------------------------------------------ pictures

/// A picture path that is the document's own: relative, inside the document's folder.
fn own_asset(path: &str) -> Option<PathBuf> {
    let normal = path.replace('\\', "/");
    if normal.is_empty() || normal.starts_with('/') || normal.contains(':') || normal.contains('\0') {
        return None;
    }
    let relative = Path::new(&normal);
    let mut depth = 0;
    for part in relative.components() {
        match part {
            Component::Normal(_) => depth += 1,
            Component::CurDir => {}
            _ => return None,
        }
    }
    if depth == 0 {
        None
    } else {
        Some(relative.components().filter(|part| matches!(part, Component::Normal(_))).collect())
    }
}

fn percent_decoded(text: &str) -> String {
    let bytes = text.as_bytes();
    let hex = |byte: u8| (byte as char).to_digit(16).map(|digit| digit as u8);
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let (Some(high), Some(low)) = (hex(bytes[i + 1]), hex(bytes[i + 2])) {
                out.push(high * 16 + low);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8(out).unwrap_or_else(|_| text.to_string())
}

/// The paths of the pictures a Markdown text shows (`![…](path)`, `![…](<path>)`), as written.
pub fn picture_paths(text: &str) -> Vec<String> {
    let mut paths = Vec::new();
    let bytes = text.as_bytes();
    let mut i = 0;
    while i + 1 < bytes.len() {
        if bytes[i] != b'!' || bytes[i + 1] != b'[' {
            i += 1;
            continue;
        }
        // the description, to its "](" on the same line
        let mut j = i + 2;
        while j < bytes.len() && bytes[j] != b'\n' && !(bytes[j] == b']' && bytes[j - 1] != b'\\') {
            j += 1;
        }
        if j + 1 >= bytes.len() || bytes[j] != b']' || bytes[j + 1] != b'(' {
            i += 2;
            continue;
        }
        let mut start = j + 2;
        while start < bytes.len() && bytes[start] == b' ' {
            start += 1;
        }
        let (from, to) = if start < bytes.len() && bytes[start] == b'<' {
            let mut end = start + 1;
            while end < bytes.len() && bytes[end] != b'>' && bytes[end] != b'\n' {
                end += 1;
            }
            (start + 1, end)
        } else {
            let mut end = start;
            while end < bytes.len() && !matches!(bytes[end], b')' | b' ' | b'\n' | b'\t') {
                end += 1;
            }
            (start, end)
        };
        if to > from {
            let path = text[from..to].to_string();
            if !paths.contains(&path) {
                paths.push(path);
            }
        }
        i = to.max(i + 2);
    }
    paths
}

/// The document's own picture files: the paths its text and its memory record name, as files of its folder.
fn own_assets(dir: &Path, text: &str, memory: Option<&str>) -> Vec<PathBuf> {
    let mut named = picture_paths(text);
    let record: Option<serde_json::Value> = memory.and_then(|text| serde_json::from_str(text).ok());
    if let Some(images) = record.as_ref().and_then(|record| record.get("images")).and_then(|images| images.as_object()) {
        named.extend(images.values().filter_map(|path| path.as_str()).map(str::to_string));
    }
    let mut found: Vec<PathBuf> = Vec::new();
    for path in named {
        for candidate in [path.clone(), percent_decoded(&path)] {
            if let Some(relative) = own_asset(&candidate) {
                if dir.join(&relative).is_file() && !found.contains(&relative) {
                    found.push(relative);
                }
            }
        }
    }
    found
}

/// Whether another document of `dir` names the picture `relative`.
fn used_by_another(dir: &Path, relative: &Path, except: &Path) -> bool {
    let name = relative.to_string_lossy().replace('\\', "/");
    let Ok(entries) = fs::read_dir(dir) else { return false };
    entries.flatten().any(|entry| {
        let path = entry.path();
        path != except && extension_of(&path).is_some() && fs::read(&path).map(|bytes| String::from_utf8_lossy(&bytes).contains(&name)).unwrap_or(false)
    })
}

/// The document's pictures copied to `to` (the same relative paths). A file already there with the same bytes is the
/// picture; a file of that name with other bytes stops the copy before anything is copied (the document would show it
/// instead of its own, and its own would then be removed), as does a picture that cannot be copied (the copies made
/// are removed again).
fn copy_assets(from: &Path, to: &Path, assets: &[PathBuf]) -> Result<(), String> {
    if from == to {
        return Ok(());
    }
    let mut wanted = Vec::new();
    for relative in assets {
        let source = from.join(relative);
        let target = to.join(relative);
        if target.exists() {
            let same = matches!((fs::read(&source), fs::read(&target)), (Ok(mine), Ok(there)) if mine == there);
            if !same {
                return Err(format!("Another picture is named {} there", relative.to_string_lossy().replace('\\', "/")));
            }
            continue;
        }
        wanted.push((source, target));
    }
    let mut copied: Vec<PathBuf> = Vec::new();
    for (source, target) in wanted {
        let done = match target.parent() {
            Some(parent) => fs::create_dir_all(parent),
            None => Ok(()),
        }
        .and_then(|_| fs::copy(&source, &target).map(|_| ()));
        if let Err(error) = done {
            for path in &copied {
                let _ = fs::remove_file(path);
            }
            return Err(format!("Unable to copy the picture {}: {error}", source.display()));
        }
        copied.push(target);
    }
    Ok(())
}

/// The pictures of a document that left `dir` are removed there, unless another document of `dir` shows them.
fn drop_assets(dir: &Path, assets: &[PathBuf], left: &Path) {
    for relative in assets {
        if !used_by_another(dir, relative, left) {
            let _ = fs::remove_file(dir.join(relative));
            // (an emptied assets folder goes too; remove_dir refuses one that still has files)
            let mut parent = relative.parent();
            while let Some(folder) = parent.filter(|p| !p.as_os_str().is_empty()) {
                if fs::remove_dir(dir.join(folder)).is_err() {
                    break;
                }
                parent = folder.parent();
            }
        }
    }
}

fn write_assets(dir: &Path, assets: &[AssetFile]) -> Result<(), String> {
    // a file with a picture's name is never replaced: it is the picture already, or another picture's name was given
    // (then nothing is written: every name is checked first)
    let mut wanted = Vec::new();
    for asset in assets {
        let relative = own_asset(&asset.path).ok_or("Invalid picture path")?;
        let target = dir.join(relative);
        let bytes = decode(&asset.base64)?;
        match fs::read(&target) {
            Ok(there) if there == bytes => continue,
            Ok(_) => return Err(format!("Another picture is named {}", asset.path)),
            Err(_) => wanted.push((target, bytes)),
        }
    }
    for (target, bytes) in wanted {
        write_atomically(&target, &bytes)?;
    }
    Ok(())
}

fn asset_names(dir: &Path) -> Vec<String> {
    let mut names: Vec<String> = fs::read_dir(dir.join(ASSETS_DIR))
        .map(|entries| entries.flatten().filter(|e| e.path().is_file()).map(|e| e.file_name().to_string_lossy().into_owned()).collect())
        .unwrap_or_default();
    names.sort();
    names
}

// ------------------------------------------------------------------------------------------------------- listing

fn trash_items(places: &Places) -> Vec<(PathBuf, TrashInfo)> {
    let mut items = Vec::new();
    let Ok(entries) = fs::read_dir(places.root.join(TRASH_DIR)) else { return items };
    let cutoff = now().saturating_sub(TRASH_KEEP_MS);
    for entry in entries.flatten() {
        let dir = entry.path();
        if !dir.is_dir() {
            continue;
        }
        let info: Option<TrashInfo> = fs::read_to_string(dir.join("info.json")).ok().and_then(|text| serde_json::from_str(&text).ok());
        let Some(info) = info else { continue };
        if info.trashed < cutoff {
            let _ = fs::remove_dir_all(&dir);
            continue;
        }
        let file = fs::read_dir(&dir).ok().and_then(|files| files.flatten().map(|f| f.path()).find(|p| p.is_file() && extension_of(p).is_some()));
        if let Some(file) = file {
            items.push((file, info));
        }
    }
    items
}

fn walk(places: &Places, dir: &Path, depth: usize, folders: &mut Vec<DocumentFolder>, documents: &mut Vec<DocumentEntry>) {
    let Ok(entries) = fs::read_dir(dir) else { return };
    for entry in entries.flatten() {
        let path = entry.path();
        let name = entry.file_name().to_string_lossy().into_owned();
        if name.starts_with('.') {
            continue;
        }
        if path.is_dir() {
            // (a folder of pictures is not a folder of documents)
            if name == ASSETS_DIR || depth >= MAX_DEPTH {
                continue;
            }
            if let Some(id) = relative_id(&places.root, &path) {
                let created = fs::metadata(&path).ok().and_then(|m| millis(m.created()).or_else(|| millis(m.modified()))).unwrap_or_else(now);
                folders.push(DocumentFolder { name: id.clone(), id, created });
            }
            walk(places, &path, depth + 1, folders, documents);
        } else if extension_of(&path).is_some() && documents.len() < MAX_DOCUMENTS {
            if let Some(entry) = entry_for(places, &path, None) {
                documents.push(entry);
            }
        }
    }
}

fn library(places: &Places) -> Result<Library, String> {
    fs::create_dir_all(&places.root).map_err(|e| format!("Unable to create the documents folder: {e}"))?;
    let mut folders = Vec::new();
    let mut documents = Vec::new();
    walk(places, &places.root, 0, &mut folders, &mut documents);
    for (file, info) in trash_items(places) {
        if let Some(entry) = entry_for(places, &file, Some(info.trashed)) {
            documents.push(entry);
        }
    }
    documents.sort_by(|a, b| b.modified.cmp(&a.modified).then_with(|| a.id.cmp(&b.id)));
    folders.sort_by(|a, b| a.id.to_lowercase().cmp(&b.id.to_lowercase()));
    let mut outside = load_outside(places).files;
    outside.sort_by(|a, b| b.opened.cmp(&a.opened));
    let recent = outside
        .iter()
        .filter(|file| !file.unlisted && relative_id(&places.root, Path::new(&file.path)).is_none())
        .filter_map(|file| entry_for(places, Path::new(&file.path), None))
        .take(MAX_RECENT)
        .collect();
    Ok(Library { root: places.root.to_string_lossy().into_owned(), folders, documents, recent })
}

pub fn list(places: &Places) -> Result<Library, String> {
    let _guard = lock();
    library(places)
}

// ----------------------------------------------------------------------------------------------- reading, saving

pub fn read(places: &Places, id: &str) -> Result<OpenedDocument, String> {
    let _guard = lock();
    let located = locate(places, id)?;
    let bytes = fs::read(&located.path).map_err(|e| format!("Unable to read the document: {e}"))?;
    if bytes.len() > MAX_DOCUMENT_BYTES {
        return Err("The document is too large".to_string());
    }
    let print = fingerprint(&bytes);
    let memory = read_memory(places, &located, &print);
    let trashed = if located.place == Place::Trash { trash_items(places).into_iter().find(|(file, _)| file == &located.path).map(|(_, info)| info.trashed) } else { None };
    let entry = entry_for(places, &located.path, trashed).ok_or("Unable to read the document")?;
    let dir = located.path.parent().ok_or("No parent directory")?;
    Ok(OpenedDocument { entry, base64: encode(&bytes), memory, fingerprint: print, dir: dir.to_string_lossy().into_owned(), assets: asset_names(dir) })
}

/// The file's state now: what a save or a reload is decided by.
pub fn check(places: &Places, id: &str) -> Result<FileState, String> {
    let located = locate(places, id)?;
    Ok(match fs::read(&located.path) {
        Ok(bytes) => FileState { exists: true, fingerprint: fingerprint(&bytes) },
        Err(_) => FileState { exists: false, fingerprint: String::new() },
    })
}

fn write_document(places: &Places, located: &Located, content: &DocumentContent) -> Result<String, String> {
    let bytes = decode(&content.base64)?;
    let print = fingerprint(&bytes);
    let dir = located.path.parent().ok_or("No parent directory")?;
    write_assets(dir, &content.assets)?;
    // the record first: a file without its record's latest state reads as text alone where the two differ
    write_memory(places, located, content.memory.as_deref(), &print)?;
    write_atomically(&located.path, &bytes)?;
    if let Some(time) = content.modified {
        if let Ok(file) = fs::OpenOptions::new().write(true).open(&located.path) {
            let _ = file.set_modified(UNIX_EPOCH + std::time::Duration::from_millis(time));
        }
    }
    Ok(print)
}

/// A new document in `folder` of the documents folder, named `name` (or the first free "name 2", "name 3"…).
pub fn create(places: &Places, folder: Option<&str>, name: &str, extension: &str, content: &DocumentContent) -> Result<DocumentEntry, String> {
    let _guard = lock();
    let extension = if extension == "vtxdoc" { "vtxdoc" } else { "md" };
    let mut dir = places.root.clone();
    if let Some(folder) = folder.filter(|f| !f.is_empty()) {
        for part in relative_parts(folder)? {
            dir.push(part);
        }
    }
    fs::create_dir_all(&dir).map_err(|e| format!("Unable to create the documents folder: {e}"))?;
    let stem = file_name(name);
    if stem.is_empty() {
        return Err("The document needs a name".to_string());
    }
    let path = free_path(&dir, &stem, extension);
    write_document(places, &Located { path: path.clone(), place: Place::Library }, content)?;
    entry_for(places, &path, None).ok_or_else(|| "Unable to create the document".to_string())
}

/// Saves the document if its file is still what the editor read (`base`, the fingerprint it got), or anyway when
/// `force`. A file that is gone is written again.
pub fn save(places: &Places, id: &str, content: &DocumentContent, base: Option<&str>, force: bool) -> Result<SaveOutcome, String> {
    let _guard = lock();
    let located = locate(places, id)?;
    if !force {
        if let (Some(base), Ok(current)) = (base, fs::read(&located.path)) {
            let print = fingerprint(&current);
            if print != base {
                return Ok(SaveOutcome::Changed { fingerprint: print });
            }
        }
    }
    let print = write_document(places, &located, content)?;
    let entry = entry_for(places, &located.path, None).ok_or("Unable to save the document")?;
    Ok(SaveOutcome::Saved { entry, fingerprint: print })
}

/// A document written to a file the user chose (Save as): the text, its record and its pictures beside it.
pub fn save_as(places: &Places, path: &Path, content: &DocumentContent) -> Result<DocumentEntry, String> {
    let _guard = lock();
    if !path.is_absolute() || extension_of(path).is_none() {
        return Err("Invalid document path".to_string());
    }
    let place = if relative_id(&places.root, path).is_some() { Place::Library } else { Place::Outside };
    write_document(places, &Located { path: path.to_path_buf(), place }, content)?;
    entry_for(places, path, None).ok_or_else(|| "Unable to save the document".to_string())
}

/// A file elsewhere taken into the recent files (a file of the documents folder is just itself).
pub fn open_outside(places: &Places, path: &Path) -> Result<DocumentEntry, String> {
    let _guard = lock();
    if !path.is_absolute() || extension_of(path).is_none() {
        return Err("Not a Markdown or VisualTeX document".to_string());
    }
    let bytes = fs::read(path).map_err(|e| format!("Unable to read the document: {e}"))?;
    if relative_id(&places.root, path).is_none() {
        let mut index = load_outside(places);
        let print = fingerprint(&bytes);
        let entry = outside_entry(places, &mut index, path, &print);
        entry.opened = now();
        entry.unlisted = false;
        store_outside(places, &index)?;
    }
    entry_for(places, path, None).ok_or_else(|| "Unable to read the document".to_string())
}

/// A file elsewhere leaves the recent files. The file stays, and so does its record: opened again, the document has
/// what VisualTeX remembers of it.
pub fn forget(places: &Places, id: &str) -> Result<Library, String> {
    let _guard = lock();
    let located = locate(places, id)?;
    if located.place == Place::Outside {
        let mut index = load_outside(places);
        let text = located.path.to_string_lossy().into_owned();
        for file in index.files.iter_mut().filter(|file| file.path == text) {
            file.unlisted = true;
        }
        store_outside(places, &index)?;
    }
    library(places)
}

// ------------------------------------------------------------------------------- renaming, moving, deleting

fn text_of(path: &Path) -> String {
    fs::read(path).map(|bytes| String::from_utf8_lossy(&bytes).into_owned()).unwrap_or_default()
}

/// The document's file, record and pictures go to `target` (a free path in the documents folder).
fn relocate(places: &Places, from: &Path, target: &Path) -> Result<(), String> {
    let old_dir = from.parent().ok_or("No parent directory")?.to_path_buf();
    let new_dir = target.parent().ok_or("No parent directory")?.to_path_buf();
    let memory = library_memory_path(places, from).and_then(|record| fs::read_to_string(record).ok());
    let assets = own_assets(&old_dir, &text_of(from), memory.as_deref());
    fs::create_dir_all(&new_dir).map_err(|e| format!("Unable to create the folder: {e}"))?;
    copy_assets(&old_dir, &new_dir, &assets)?;
    fs::rename(from, target).map_err(|e| format!("Unable to move the document: {e}"))?;
    move_library_memory(places, from, target);
    if old_dir != new_dir {
        drop_assets(&old_dir, &assets, from);
    }
    Ok(())
}

fn library_document(places: &Places, id: &str) -> Result<PathBuf, String> {
    let located = locate(places, id)?;
    if located.place != Place::Library || !located.path.is_file() {
        return Err("The document is not in the documents folder".to_string());
    }
    Ok(located.path)
}

/// Renames the document's file (its record follows); the new id comes back with the library.
pub fn rename(places: &Places, id: &str, title: &str) -> Result<(Library, String), String> {
    let _guard = lock();
    let located = locate(places, id)?;
    let stem = file_name(title);
    if stem.is_empty() {
        return Err("The document needs a name".to_string());
    }
    let extension = extension_of(&located.path).unwrap_or("md");
    let dir = located.path.parent().ok_or("No parent directory")?;
    let same = located.path.file_stem().map(|s| s.to_string_lossy() == stem.as_str()).unwrap_or(false);
    if same {
        return Ok((library(places)?, id.to_string()));
    }
    // (a change of letter case only is the same file on Windows and macOS: no "name 2")
    let only_case = located.path.file_stem().map(|s| s.to_string_lossy().to_lowercase() == stem.to_lowercase()).unwrap_or(false);
    let target = if only_case { dir.join(format!("{stem}.{extension}")) } else { free_path(dir, &stem, extension) };
    match located.place {
        Place::Library => {
            fs::rename(&located.path, &target).map_err(|e| format!("Unable to rename the document: {e}"))?;
            move_library_memory(places, &located.path, &target);
        }
        Place::Outside => {
            fs::rename(&located.path, &target).map_err(|e| format!("Unable to rename the document: {e}"))?;
            let mut index = load_outside(places);
            let old = located.path.to_string_lossy().into_owned();
            for file in index.files.iter_mut().filter(|file| file.path == old) {
                file.path = target.to_string_lossy().into_owned();
            }
            store_outside(places, &index)?;
        }
        Place::Trash => return Err("The document is deleted".to_string()),
    }
    let new_id = entry_for(places, &target, None).map(|entry| entry.id).ok_or("Unable to rename the document")?;
    Ok((library(places)?, new_id))
}

/// A copy beside the document, with its record (its pictures are the same files).
pub fn duplicate(places: &Places, id: &str, title: &str) -> Result<DocumentEntry, String> {
    let _guard = lock();
    let located = locate(places, id)?;
    if located.place == Place::Trash {
        return Err("The document is deleted".to_string());
    }
    let bytes = fs::read(&located.path).map_err(|e| format!("Unable to read the document: {e}"))?;
    let print = fingerprint(&bytes);
    let memory = read_memory(places, &located, &print);
    let stem = file_name(title);
    if stem.is_empty() {
        return Err("The document needs a name".to_string());
    }
    let dir = located.path.parent().ok_or("No parent directory")?;
    let target = free_path(dir, &stem, extension_of(&located.path).unwrap_or("md"));
    let copy = Located { path: target.clone(), place: located.place.clone() };
    write_memory(places, &copy, memory.as_deref(), &print)?;
    write_atomically(&target, &bytes)?;
    entry_for(places, &target, None).ok_or_else(|| "Unable to copy the document".to_string())
}

/// Moves the document to another folder of the documents folder (none: the folder itself), with its record and its
/// pictures.
pub fn move_to(places: &Places, id: &str, folder: Option<&str>) -> Result<(Library, String), String> {
    let _guard = lock();
    let from = library_document(places, id)?;
    let mut dir = places.root.clone();
    if let Some(folder) = folder.filter(|f| !f.is_empty()) {
        for part in relative_parts(folder)? {
            dir.push(part);
        }
        if !dir.is_dir() {
            return Err("The folder is not in the documents folder".to_string());
        }
    }
    if from.parent() == Some(dir.as_path()) {
        return Ok((library(places)?, id.to_string()));
    }
    let stem = from.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
    let target = free_path(&dir, &stem, extension_of(&from).unwrap_or("md"));
    relocate(places, &from, &target)?;
    let new_id = relative_id(&places.root, &target).ok_or("Unable to move the document")?;
    Ok((library(places)?, new_id))
}

/// The document goes to the recently deleted ones, with its record and its pictures.
pub fn trash(places: &Places, id: &str) -> Result<Library, String> {
    let _guard = lock();
    let from = library_document(places, id)?;
    let old_dir = from.parent().ok_or("No parent directory")?.to_path_buf();
    hidden_dir(&places.root.join(TRASH_DIR))?;
    let item = places.root.join(TRASH_DIR).join(uuid::Uuid::new_v4().simple().to_string());
    fs::create_dir_all(&item).map_err(|e| format!("Unable to delete the document: {e}"))?;
    let record = library_memory_path(places, &from);
    let memory = record.as_ref().and_then(|record| fs::read_to_string(record).ok());
    let assets = own_assets(&old_dir, &text_of(&from), memory.as_deref());
    if let Err(error) = copy_assets(&old_dir, &item, &assets) {
        let _ = fs::remove_dir_all(&item);
        return Err(error);
    }
    let info = TrashInfo { path: id.to_string(), trashed: now() };
    write_atomically(&item.join("info.json"), &serde_json::to_vec_pretty(&info).map_err(|e| e.to_string())?)?;
    let target = item.join(from.file_name().ok_or("Invalid document path")?);
    fs::rename(&from, &target).map_err(|e| format!("Unable to delete the document: {e}"))?;
    if let (Some(record), Some(text)) = (record, memory) {
        let _ = fs::write(trash_memory_path(&target), text);
        let _ = fs::remove_file(record);
    }
    drop_assets(&old_dir, &assets, &from);
    library(places)
}

/// A recently deleted document goes back where it was (the documents folder itself when its folder is gone).
pub fn restore(places: &Places, id: &str) -> Result<Library, String> {
    let _guard = lock();
    let located = locate(places, id)?;
    if located.place != Place::Trash || !located.path.is_file() {
        return Err("The document is not deleted".to_string());
    }
    let item = located.path.parent().ok_or("No parent directory")?.to_path_buf();
    let info: Option<TrashInfo> = fs::read_to_string(item.join("info.json")).ok().and_then(|text| serde_json::from_str(&text).ok());
    let mut dir = places.root.clone();
    if let Some(info) = &info {
        if let Ok(parts) = relative_parts(&info.path) {
            let mut was = places.root.clone();
            for part in &parts[..parts.len() - 1] {
                was.push(part);
            }
            if was.is_dir() {
                dir = was;
            }
        }
    }
    let stem = located.path.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
    let target = free_path(&dir, &stem, extension_of(&located.path).unwrap_or("md"));
    let memory = fs::read_to_string(trash_memory_path(&located.path)).ok();
    let assets = own_assets(&item, &text_of(&located.path), memory.as_deref());
    copy_assets(&item, &dir, &assets)?;
    fs::rename(&located.path, &target).map_err(|e| format!("Unable to restore the document: {e}"))?;
    if let Some(text) = memory {
        write_memory(places, &Located { path: target, place: Place::Library }, Some(&text), "")?;
    }
    let _ = fs::remove_dir_all(&item);
    library(places)
}

/// A recently deleted document is removed for good.
pub fn delete(places: &Places, id: &str) -> Result<Library, String> {
    let _guard = lock();
    let located = locate(places, id)?;
    if located.place != Place::Trash {
        return Err("The document is not deleted".to_string());
    }
    if let Some(item) = located.path.parent() {
        let _ = fs::remove_dir_all(item);
    }
    library(places)
}

pub fn empty_trash(places: &Places) -> Result<Library, String> {
    let _guard = lock();
    if let Ok(entries) = fs::read_dir(places.root.join(TRASH_DIR)) {
        for entry in entries.flatten() {
            let _ = fs::remove_dir_all(entry.path());
        }
    }
    library(places)
}

// ------------------------------------------------------------------------------------------------------- folders

fn folder_path(places: &Places, id: &str) -> Result<PathBuf, String> {
    let mut path = places.root.clone();
    for part in relative_parts(id)? {
        path.push(part);
    }
    if !path.is_dir() || id.starts_with(TRASH_DIR) {
        return Err("The folder is not in the documents folder".to_string());
    }
    Ok(path)
}

pub fn create_folder(places: &Places, name: &str) -> Result<(Library, String), String> {
    let _guard = lock();
    let name = file_name(name);
    if name.is_empty() {
        return Err("The folder needs a name".to_string());
    }
    let mut path = places.root.join(&name);
    let mut n = 2;
    while path.exists() {
        path = places.root.join(format!("{name} {n}"));
        n += 1;
    }
    fs::create_dir_all(&path).map_err(|e| format!("Unable to create the folder: {e}"))?;
    let id = relative_id(&places.root, &path).ok_or("Unable to create the folder")?;
    Ok((library(places)?, id))
}

/// Renames a folder; its documents' records follow.
pub fn rename_folder(places: &Places, id: &str, name: &str) -> Result<Library, String> {
    let _guard = lock();
    let path = folder_path(places, id)?;
    let name = file_name(name);
    if name.is_empty() {
        return Err("The folder needs a name".to_string());
    }
    let target = path.with_file_name(&name);
    if target != path {
        if target.exists() && target.to_string_lossy().to_lowercase() != path.to_string_lossy().to_lowercase() {
            return Err("A folder with this name exists".to_string());
        }
        fs::rename(&path, &target).map_err(|e| format!("Unable to rename the folder: {e}"))?;
        let records = places.root.join(MEMORY_DIR);
        if let (Ok(old), Ok(new)) = (path.strip_prefix(&places.root), target.strip_prefix(&places.root)) {
            if records.join(old).is_dir() {
                let _ = fs::rename(records.join(old), records.join(new));
            }
        }
    }
    library(places)
}

fn documents_under(dir: &Path, depth: usize, out: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(dir) else { return };
    for entry in entries.flatten() {
        let path = entry.path();
        let name = entry.file_name().to_string_lossy().into_owned();
        if name.starts_with('.') {
            continue;
        }
        if path.is_dir() {
            if name != ASSETS_DIR && depth < MAX_DEPTH {
                documents_under(&path, depth + 1, out);
            }
        } else if extension_of(&path).is_some() {
            out.push(path);
        }
    }
}

fn remove_empty_dirs(dir: &Path) -> bool {
    let Ok(entries) = fs::read_dir(dir) else { return false };
    let mut empty = true;
    for entry in entries.flatten() {
        let path = entry.path();
        if !(path.is_dir() && remove_empty_dirs(&path)) {
            empty = false;
        }
    }
    empty && fs::remove_dir(dir).is_ok()
}

/// The folder goes; its documents stay, in the documents folder itself (with their records and pictures). A folder
/// that holds other files stays with them.
pub fn delete_folder(places: &Places, id: &str) -> Result<Library, String> {
    let _guard = lock();
    let path = folder_path(places, id)?;
    let mut documents = Vec::new();
    documents_under(&path, 0, &mut documents);
    for from in documents {
        let stem = from.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
        let target = free_path(&places.root, &stem, extension_of(&from).unwrap_or("md"));
        relocate(places, &from, &target)?;
    }
    remove_empty_dirs(&path);
    if let Ok(relative) = path.strip_prefix(&places.root) {
        remove_empty_dirs(&places.root.join(MEMORY_DIR).join(relative));
    }
    library(places)
}

// -------------------------------------------------------------------------------------------------------- search

/// The ids of the documents (not deleted) whose name or text contains every word of the query.
pub fn search(places: &Places, query: &str) -> Result<Vec<String>, String> {
    let words: Vec<String> = query.split_whitespace().map(|w| w.to_lowercase()).collect();
    if words.is_empty() {
        return Ok(Vec::new());
    }
    let documents = {
        let _guard = lock();
        library(places)?.documents
    };
    let mut found = Vec::new();
    for doc in documents.iter().filter(|doc| doc.trashed.is_none()) {
        let mut text = doc.title.clone();
        text.push('\n');
        text.push_str(&text_of(Path::new(&doc.path)));
        let text = text.to_lowercase();
        if words.iter().all(|word| text.contains(word.as_str())) {
            found.push(doc.id.clone());
        }
    }
    Ok(found)
}

// ------------------------------------------------------------------------- the library of earlier versions

/// A document of the library earlier versions kept in the app's data folder (`documents/<id>.vtxdoc`, `index.json`).
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct LegacyDocument {
    pub id: String,
    pub title: String,
    /// its folder's name there
    pub folder: Option<String>,
    pub created: u64,
    pub modified: u64,
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
struct LegacyEntry {
    id: String,
    #[serde(default)]
    title: String,
    #[serde(default)]
    folder: Option<String>,
    #[serde(default)]
    created: u64,
    #[serde(default)]
    modified: u64,
    #[serde(default)]
    trashed: Option<u64>,
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
struct LegacyFolder {
    id: String,
    #[serde(default)]
    name: String,
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
struct LegacyIndex {
    #[serde(default)]
    folders: Vec<LegacyFolder>,
    #[serde(default)]
    documents: Vec<LegacyEntry>,
}

fn legacy_id(id: &str) -> Result<&str, String> {
    if !id.is_empty() && id.len() <= 64 && id.bytes().all(|b| b.is_ascii_alphanumeric() || b == b'-' || b == b'_') {
        Ok(id)
    } else {
        Err("Invalid document id".to_string())
    }
}

/// The documents of the earlier library that were not deleted there: its index's, and files without an entry.
pub fn legacy_list(old: &Path) -> Vec<LegacyDocument> {
    let index: LegacyIndex = fs::read_to_string(old.join("index.json")).ok().and_then(|text| serde_json::from_str(&text).ok()).unwrap_or_default();
    let mut seen: HashSet<String> = HashSet::new();
    let mut documents = Vec::new();
    for entry in &index.documents {
        seen.insert(entry.id.clone());
        if entry.trashed.is_some() || legacy_id(&entry.id).is_err() || !old.join(format!("{}.vtxdoc", entry.id)).is_file() {
            continue;
        }
        let folder = entry.folder.as_ref().and_then(|id| index.folders.iter().find(|f| &f.id == id)).map(|f| f.name.clone()).filter(|name| !name.trim().is_empty());
        documents.push(LegacyDocument { id: entry.id.clone(), title: entry.title.clone(), folder, created: entry.created, modified: entry.modified });
    }
    if let Ok(entries) = fs::read_dir(old) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().and_then(|e| e.to_str()) != Some("vtxdoc") {
                continue;
            }
            let Some(stem) = path.file_stem().and_then(|s| s.to_str()) else { continue };
            if stem.starts_with('.') || seen.contains(stem) || legacy_id(stem).is_err() {
                continue;
            }
            let time = fs::metadata(&path).ok().and_then(|m| millis(m.modified())).unwrap_or_else(now);
            documents.push(LegacyDocument { id: stem.to_string(), title: String::new(), folder: None, created: time, modified: time });
        }
    }
    documents.sort_by(|a, b| a.modified.cmp(&b.modified));
    documents
}

pub fn legacy_read(old: &Path, id: &str) -> Result<String, String> {
    fs::read_to_string(old.join(format!("{}.vtxdoc", legacy_id(id)?))).map_err(|e| format!("Unable to read the document: {e}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn places(dir: &tempfile::TempDir) -> Places {
        Places { root: dir.path().join("Documents").join("VisualTeX"), data: dir.path().join("data") }
    }

    fn content(text: &str, memory: Option<&str>) -> DocumentContent {
        DocumentContent { base64: encode(text.as_bytes()), memory: memory.map(str::to_string), ..Default::default() }
    }

    fn with_assets(text: &str, memory: Option<&str>, assets: &[(&str, &[u8])]) -> DocumentContent {
        DocumentContent {
            base64: encode(text.as_bytes()),
            memory: memory.map(str::to_string),
            assets: assets.iter().map(|(path, bytes)| AssetFile { path: path.to_string(), base64: encode(bytes) }).collect(),
            modified: None,
        }
    }

    fn text(places: &Places, id: &str) -> String {
        String::from_utf8(decode(&read(places, id).unwrap().base64).unwrap()).unwrap()
    }

    const MEMORY: &str = r#"{"format":"visualtex-document-memory","version":1,"settings":{"paper":"a5"},"blocks":[],"images":{"vtx-image:1.png":"assets/in memory.png"}}"#;

    #[test]
    fn names_and_collisions() {
        assert_eq!(file_name("  a/b:c*?.md.  "), "a-b-c--.md");
        assert_eq!(file_name("证明：第一章"), "证明：第一章");
        assert_eq!(file_name("CON"), "_CON");
        assert_eq!(file_name("com1.txt"), "_com1.txt");
        assert_eq!(file_name("..hidden"), "hidden");
        assert_eq!(file_name("\u{1}\n"), "--");
        assert_eq!(file_name(" . "), "");
        let dir = tempfile::TempDir::new().unwrap();
        let places = places(&dir);
        let a = create(&places, None, "Notes", "md", &content("one\n", None)).unwrap();
        let b = create(&places, None, "Notes", "md", &content("two\n", None)).unwrap();
        let c = create(&places, None, "Notes", "md", &content("three\n", None)).unwrap();
        assert_eq!((a.id.as_str(), b.id.as_str(), c.id.as_str()), ("Notes.md", "Notes 2.md", "Notes 3.md"));
        assert_eq!((a.title.as_str(), a.kind.as_str(), a.outside, a.folder.clone()), ("Notes", "markdown", false, None));
        assert!(create(&places, None, " . ", "md", &content("", None)).is_err());
        // a rename to a name in use takes the next free one; to its own name nothing; to another letter case that name
        let (_, renamed) = rename(&places, "Notes 3.md", "Notes").unwrap();
        assert_eq!(renamed, "Notes 4.md");
        assert_eq!(rename(&places, "Notes.md", "Notes").unwrap().1, "Notes.md");
        assert_eq!(rename(&places, "Notes.md", "NOTES").unwrap().1, "NOTES.md");
        assert_eq!(text(&places, "NOTES.md"), "one\n");
        // ids cannot leave the documents folder or name hidden files
        for id in ["../x.md", "/etc/x.md", "a/../x.md", ".visualtex/x.md", "a\\b.md", "x.txt", "", ".trash/x.md"] {
            assert!(read(&places, id).is_err(), "{id}");
        }
        assert!(create(&places, Some("../up"), "x", "md", &content("", None)).is_err());
    }

    #[test]
    fn save_checks_the_file_on_disk_and_writes_atomically() {
        let dir = tempfile::TempDir::new().unwrap();
        let places = places(&dir);
        let entry = create(&places, None, "Doc", "md", &content("first\n", Some(MEMORY))).unwrap();
        let opened = read(&places, &entry.id).unwrap();
        assert_eq!(opened.memory.as_deref(), Some(MEMORY));
        assert_eq!(opened.fingerprint, fingerprint(b"first\n"));
        assert!(places.root.join(".visualtex").join("Doc.md.json").is_file());
        // unchanged on disk: written
        let saved = save(&places, &entry.id, &content("second\n", None), Some(&opened.fingerprint), false).unwrap();
        let SaveOutcome::Saved { fingerprint: print, .. } = saved else { panic!("not saved") };
        assert_eq!(print, fingerprint(b"second\n"));
        assert!(!places.root.join(".visualtex").join("Doc.md.json").exists(), "a document without memory has no record");
        // changed on disk by another program: left alone, the new state reported
        fs::write(places.root.join("Doc.md"), "elsewhere\n").unwrap();
        let outcome = save(&places, &entry.id, &content("third\n", Some(MEMORY)), Some(&print), false).unwrap();
        let SaveOutcome::Changed { fingerprint: theirs } = outcome else { panic!("not refused") };
        assert_eq!(theirs, fingerprint(b"elsewhere\n"));
        assert_eq!(text(&places, &entry.id), "elsewhere\n");
        assert_eq!(check(&places, &entry.id).unwrap().fingerprint, theirs);
        // overwrite
        assert!(matches!(save(&places, &entry.id, &content("third\n", Some(MEMORY)), Some(&print), true).unwrap(), SaveOutcome::Saved { .. }));
        assert_eq!(text(&places, &entry.id), "third\n");
        // gone from the disk: written again
        fs::remove_file(places.root.join("Doc.md")).unwrap();
        assert!(!check(&places, &entry.id).unwrap().exists);
        assert!(matches!(save(&places, &entry.id, &content("fourth\n", None), Some(&print), false).unwrap(), SaveOutcome::Saved { .. }));
        // no temporary file is left, bytes are kept as given (a byte order mark, CRLF)
        let bytes = [0xef, 0xbb, 0xbf, b'a', b'\r', b'\n'];
        save(&places, &entry.id, &DocumentContent { base64: encode(&bytes), ..Default::default() }, None, false).unwrap();
        assert_eq!(fs::read(places.root.join("Doc.md")).unwrap(), bytes);
        let names: Vec<String> = fs::read_dir(&places.root).unwrap().flatten().map(|e| e.file_name().to_string_lossy().into_owned()).collect();
        assert!(names.iter().all(|name| !name.ends_with(".tmp")), "{names:?}");
        // a failed write leaves the file as it was
        assert!(write_atomically(&places.root.join("Doc.md").join("below-a-file"), b"x").is_err());
        assert_eq!(fs::read(places.root.join("Doc.md")).unwrap(), bytes);
    }

    #[cfg(unix)]
    #[test]
    fn a_link_stays_a_link_and_a_file_keeps_its_permissions() {
        use std::os::unix::fs::{symlink, PermissionsExt};
        let dir = tempfile::TempDir::new().unwrap();
        let real = dir.path().join("notes").join("real.md");
        fs::create_dir_all(real.parent().unwrap()).unwrap();
        fs::write(&real, "one\n").unwrap();
        fs::set_permissions(&real, fs::Permissions::from_mode(0o640)).unwrap();
        let link = dir.path().join("link.md");
        symlink(&real, &link).unwrap();
        // saved through the link: the file it points at is written, the link stays
        write_atomically(&link, b"two\n").unwrap();
        assert!(fs::symlink_metadata(&link).unwrap().file_type().is_symlink());
        assert_eq!(fs::read(&real).unwrap(), b"two\n");
        assert_eq!(fs::metadata(&real).unwrap().permissions().mode() & 0o777, 0o640);
        // a file made read-only is left alone
        fs::set_permissions(&real, fs::Permissions::from_mode(0o444)).unwrap();
        assert_eq!(write_atomically(&link, b"three\n").unwrap_err(), "The file is read-only");
        assert_eq!(fs::read(&real).unwrap(), b"two\n");
        let left: Vec<String> = fs::read_dir(real.parent().unwrap()).unwrap().flatten().map(|e| e.file_name().to_string_lossy().into_owned()).collect();
        assert_eq!(left, vec!["real.md"]);
    }

    #[test]
    fn folders_moves_and_the_memory_and_pictures_that_follow() {
        let dir = tempfile::TempDir::new().unwrap();
        let places = places(&dir);
        let (_, folder) = create_folder(&places, "课程").unwrap();
        assert_eq!(folder, "课程");
        assert_eq!(create_folder(&places, "课程").unwrap().1, "课程 2");
        let body = "![a](assets/a.png) and ![b](<assets/b c.png>) and ![far](../far.png) and ![web](https://example.com/x.png)\n";
        let doc = create(&places, None, "Paper", "md", &with_assets(body, Some(MEMORY), &[("assets/a.png", b"A"), ("assets/b c.png", b"B"), ("assets/in memory.png", b"M")])).unwrap();
        // another document of the same folder shows a.png too
        create(&places, None, "Other", "md", &content("![a](assets/a.png)\n", None)).unwrap();
        assert_eq!(read(&places, &doc.id).unwrap().assets, vec!["a.png", "b c.png", "in memory.png"]);
        // a picture file is never replaced by a save: the same bytes are there already, other bytes are refused
        save(&places, &doc.id, &with_assets(body, Some(MEMORY), &[("assets/a.png", b"A")]), None, false).unwrap();
        assert!(save(&places, &doc.id, &with_assets(body, Some(MEMORY), &[("assets/a.png", b"other bytes")]), None, false).is_err());
        assert_eq!(fs::read(places.root.join("assets").join("a.png")).unwrap(), b"A");
        // (and none of the others is written then, nor the document)
        let refused = create(&places, None, "Refused", "md", &with_assets("x\n", None, &[("assets/new.png", b"N"), ("assets/a.png", b"other bytes")]));
        assert_eq!(refused.unwrap_err(), "Another picture is named assets/a.png");
        assert!(!places.root.join("assets").join("new.png").exists() && !places.root.join("Refused.md").exists());
        // a converted document keeps the time of the one it came from
        let dated = create(&places, None, "Dated", "md", &DocumentContent { base64: encode(b"x\n"), modified: Some(1_600_000_000_000), ..Default::default() }).unwrap();
        assert_eq!(dated.modified, 1_600_000_000_000);
        trash(&places, &dated.id).unwrap();
        empty_trash(&places).unwrap();
        assert!(save(&places, &doc.id, &with_assets(body, None, &[("../escape.png", b"x")]), None, false).is_err());
        // moved into the folder: text, record and pictures there; a.png stays for the other document
        let (lib, moved) = move_to(&places, &doc.id, Some(&folder)).unwrap();
        assert_eq!(moved, "课程/Paper.md");
        assert_eq!(lib.documents.iter().find(|d| d.id == moved).unwrap().folder.as_deref(), Some("课程"));
        assert_eq!(read(&places, &moved).unwrap().memory.as_deref(), Some(MEMORY));
        assert!(places.root.join(".visualtex").join("课程").join("Paper.md.json").is_file());
        assert!(!places.root.join(".visualtex").join("Paper.md.json").exists());
        for name in ["a.png", "b c.png", "in memory.png"] {
            assert!(places.root.join("课程").join("assets").join(name).is_file(), "{name}");
        }
        assert!(places.root.join("assets").join("a.png").is_file());
        assert!(!places.root.join("assets").join("b c.png").exists());
        // renamed: the record follows
        let (_, renamed) = rename(&places, &moved, "Thesis").unwrap();
        assert_eq!(renamed, "课程/Thesis.md");
        assert_eq!(read(&places, &renamed).unwrap().memory.as_deref(), Some(MEMORY));
        // duplicated: the copy has the record
        let copy = duplicate(&places, &renamed, "Thesis copy").unwrap();
        assert_eq!(copy.id, "课程/Thesis copy.md");
        assert_eq!(read(&places, &copy.id).unwrap().memory.as_deref(), Some(MEMORY));
        // the folder renamed: records follow
        rename_folder(&places, &folder, "Courses").unwrap();
        assert_eq!(read(&places, "Courses/Thesis.md").unwrap().memory.as_deref(), Some(MEMORY));
        assert!(rename_folder(&places, "Courses", "课程 2").is_err());
        // the folder deleted: its documents stay, in the documents folder, with records and pictures
        let lib = delete_folder(&places, "Courses").unwrap();
        assert!(lib.folders.iter().all(|f| f.id != "Courses"));
        assert!(!places.root.join("Courses").exists());
        assert_eq!(read(&places, "Thesis.md").unwrap().memory.as_deref(), Some(MEMORY));
        assert_eq!(read(&places, "Thesis copy.md").unwrap().memory.as_deref(), Some(MEMORY));
        assert!(places.root.join("assets").join("b c.png").is_file());
        assert_eq!(search(&places, "far thesis").unwrap().len(), 2);
        assert!(search(&places, "missing").unwrap().is_empty());
    }

    #[test]
    fn trash_restore_and_delete() {
        let dir = tempfile::TempDir::new().unwrap();
        let places = places(&dir);
        let (_, folder) = create_folder(&places, "Notes").unwrap();
        let doc = create(&places, Some(&folder), "Proof", "md", &with_assets("![p](assets/p.png)\n", Some(MEMORY), &[("assets/p.png", b"P")])).unwrap();
        let lib = trash(&places, &doc.id).unwrap();
        let trashed = lib.documents.iter().find(|d| d.trashed.is_some()).unwrap().clone();
        assert!(trashed.id.starts_with(".trash/") && trashed.id.ends_with("/Proof.md"));
        assert!(!places.root.join("Notes").join("Proof.md").exists());
        assert!(!places.root.join("Notes").join("assets").exists(), "its pictures went with it");
        assert!(!places.root.join(".visualtex").join("Notes").join("Proof.md.json").exists());
        assert!(search(&places, "p").unwrap().is_empty());
        assert_eq!(read(&places, &trashed.id).unwrap().memory.as_deref(), Some(MEMORY));
        // a new document took its name meanwhile: restored beside it
        create(&places, Some(&folder), "Proof", "md", &content("new\n", None)).unwrap();
        let lib = restore(&places, &trashed.id).unwrap();
        assert!(lib.documents.iter().all(|d| d.trashed.is_none()));
        assert_eq!(text(&places, "Notes/Proof 2.md"), "![p](assets/p.png)\n");
        assert_eq!(read(&places, "Notes/Proof 2.md").unwrap().memory.as_deref(), Some(MEMORY));
        assert_eq!(fs::read(places.root.join("Notes").join("assets").join("p.png")).unwrap(), b"P");
        // deleted for good; the trash emptied; long-deleted ones go by themselves
        trash(&places, "Notes/Proof 2.md").unwrap();
        trash(&places, "Notes/Proof.md").unwrap();
        let lib = list(&places).unwrap();
        let ids: Vec<String> = lib.documents.iter().filter(|d| d.trashed.is_some()).map(|d| d.id.clone()).collect();
        assert_eq!(ids.len(), 2);
        assert!(delete(&places, "Notes/x.md").is_err());
        let lib = delete(&places, &ids[0]).unwrap();
        assert_eq!(lib.documents.len(), 1);
        let item = places.root.join(".trash").join(ids[1].split('/').nth(1).unwrap());
        fs::write(item.join("info.json"), r#"{"path":"Notes/Proof.md","trashed":1}"#).unwrap();
        assert!(list(&places).unwrap().documents.is_empty());
        assert!(!item.exists());
        let doc = create(&places, None, "Last", "md", &content("x\n", None)).unwrap();
        trash(&places, &doc.id).unwrap();
        assert!(empty_trash(&places).unwrap().documents.is_empty());
    }

    #[test]
    fn a_picture_of_the_same_name_is_never_taken_for_the_documents_own() {
        let dir = tempfile::TempDir::new().unwrap();
        let places = places(&dir);
        let (_, folder) = create_folder(&places, "Notes").unwrap();
        let figure = |at: &Path| fs::read(at.join("assets").join("figure.png")).unwrap();
        let a = create(&places, None, "A", "md", &with_assets("![f](assets/figure.png)\n", Some(MEMORY), &[("assets/figure.png", b"mine")])).unwrap();
        create(&places, Some(&folder), "B", "md", &with_assets("![f](assets/figure.png)\n", None, &[("assets/figure.png", b"theirs")])).unwrap();
        // moved into Notes, A would show B's picture and its own would be removed: refused, nothing moved
        assert!(move_to(&places, &a.id, Some(&folder)).unwrap_err().contains("assets/figure.png"));
        assert!(places.root.join("A.md").is_file());
        assert_eq!((figure(&places.root), figure(&places.root.join("Notes"))), (b"mine".to_vec(), b"theirs".to_vec()));
        assert_eq!(read(&places, "A.md").unwrap().memory.as_deref(), Some(MEMORY));
        // Notes deleted: B cannot come into the documents folder, so it stays where it is with its picture
        assert!(delete_folder(&places, &folder).is_err());
        assert!(places.root.join("Notes").join("B.md").is_file());
        assert_eq!(figure(&places.root.join("Notes")), b"theirs");
        // A deleted, another picture of that name put there meanwhile: A is not restored over it and keeps its own
        let trashed = trash(&places, &a.id).unwrap().documents.into_iter().find(|d| d.trashed.is_some()).unwrap();
        create(&places, None, "C", "md", &with_assets("![f](assets/figure.png)\n", None, &[("assets/figure.png", b"other")])).unwrap();
        assert!(restore(&places, &trashed.id).is_err());
        assert_eq!(read(&places, &trashed.id).unwrap().memory.as_deref(), Some(MEMORY));
        assert_eq!(figure(Path::new(&trashed.path).parent().unwrap()), b"mine");
        assert_eq!(figure(&places.root), b"other");
        // the same bytes under the same name: the picture is that file, and the move goes through
        let d = create(&places, None, "D", "md", &with_assets("![g](assets/same.png)\n", None, &[("assets/same.png", b"same")])).unwrap();
        create(&places, Some(&folder), "E", "md", &with_assets("![g](assets/same.png)\n", None, &[("assets/same.png", b"same")])).unwrap();
        assert_eq!(move_to(&places, &d.id, Some(&folder)).unwrap().1, "Notes/D.md");
        assert!(!places.root.join("assets").join("same.png").exists());
        assert_eq!(fs::read(places.root.join("Notes").join("assets").join("same.png")).unwrap(), b"same");
        // nothing half done is left in the trash
        let items: Vec<PathBuf> = fs::read_dir(places.root.join(".trash")).unwrap().flatten().map(|e| e.path()).collect();
        assert_eq!(items.len(), 1);
    }

    #[test]
    fn files_elsewhere_keep_their_memory_in_the_data_folder() {
        let dir = tempfile::TempDir::new().unwrap();
        let places = places(&dir);
        let elsewhere = dir.path().join("Desktop");
        fs::create_dir_all(&elsewhere).unwrap();
        let file = elsewhere.join("paper.md");
        fs::write(&file, "# Paper\n").unwrap();
        let entry = open_outside(&places, &file).unwrap();
        assert!(entry.outside && entry.id == format!("file:{}", file.display()));
        assert!(open_outside(&places, &elsewhere.join("notes.txt")).is_err());
        let opened = read(&places, &entry.id).unwrap();
        assert_eq!(opened.memory, None);
        // saved in place, the record in the data folder: nothing appears beside the file
        let saved = save(&places, &entry.id, &with_assets("# Paper\n\n![x](assets/x.png)\n", Some(MEMORY), &[("assets/x.png", b"X")]), Some(&opened.fingerprint), false).unwrap();
        assert!(matches!(saved, SaveOutcome::Saved { .. }));
        let beside: Vec<String> = fs::read_dir(&elsewhere).unwrap().flatten().map(|e| e.file_name().to_string_lossy().into_owned()).collect();
        assert_eq!(beside.len(), 2, "{beside:?}");
        assert!(elsewhere.join("assets").join("x.png").is_file());
        assert_eq!(read(&places, &entry.id).unwrap().memory.as_deref(), Some(MEMORY));
        assert_eq!(list(&places).unwrap().recent.len(), 1);
        assert!(list(&places).unwrap().documents.is_empty());
        // renamed and moved by another program: found again by its content
        let moved = dir.path().join("Archive").join("renamed.md");
        fs::create_dir_all(moved.parent().unwrap()).unwrap();
        fs::rename(&file, &moved).unwrap();
        let again = open_outside(&places, &moved).unwrap();
        assert_eq!(read(&places, &again.id).unwrap().memory.as_deref(), Some(MEMORY));
        let recent = list(&places).unwrap().recent;
        assert_eq!(recent.len(), 1);
        assert_eq!(recent[0].id, again.id);
        // copied: the copy has the record too, the original keeps its own
        let copy = dir.path().join("Archive").join("copy.md");
        fs::copy(&moved, &copy).unwrap();
        assert_eq!(read(&places, &format!("file:{}", copy.display())).unwrap().memory.as_deref(), Some(MEMORY));
        save(&places, &format!("file:{}", copy.display()), &content("changed\n", None), None, false).unwrap();
        assert_eq!(read(&places, &again.id).unwrap().memory.as_deref(), Some(MEMORY));
        // its text changed elsewhere and it was moved: nothing to find it by, the text alone
        fs::write(&moved, "rewritten\n").unwrap();
        let lost = dir.path().join("Archive").join("lost.md");
        fs::rename(&moved, &lost).unwrap();
        assert_eq!(read(&places, &format!("file:{}", lost.display())).unwrap().memory, None);
        // forgotten: out of the recent files, the file stays
        let lib = forget(&places, &format!("file:{}", copy.display())).unwrap();
        assert!(lib.recent.iter().all(|d| !d.path.ends_with("copy.md")));
        assert!(copy.is_file());
        // a file of the documents folder opened by its path is the library's document
        let inside = create(&places, None, "Inside", "md", &content("x\n", None)).unwrap();
        assert_eq!(open_outside(&places, Path::new(&inside.path)).unwrap().id, "Inside.md");
        // saved as a file elsewhere, and as one of the documents folder
        let target = elsewhere.join("saved as.md");
        let saved = save_as(&places, &target, &with_assets("![x](assets/y.png)\n", Some(MEMORY), &[("assets/y.png", b"Y")])).unwrap();
        assert!(saved.outside);
        assert_eq!(read(&places, &saved.id).unwrap().memory.as_deref(), Some(MEMORY));
        assert!(elsewhere.join("assets").join("y.png").is_file());
        assert_eq!(save_as(&places, &places.root.join("In library.md"), &content("x\n", Some(MEMORY))).unwrap().id, "In library.md");
        assert!(places.root.join(".visualtex").join("In library.md.json").is_file());
        // taken out of the recent files: its record stays, and opened again it is listed again with it
        assert!(forget(&places, &saved.id).unwrap().recent.iter().all(|d| d.id != saved.id));
        assert_eq!(read(&places, &saved.id).unwrap().memory.as_deref(), Some(MEMORY));
        assert!(list(&places).unwrap().recent.iter().any(|d| d.id == saved.id));
        forget(&places, &saved.id).unwrap();
        assert_eq!(open_outside(&places, &target).unwrap().id, saved.id);
        assert!(list(&places).unwrap().recent.iter().any(|d| d.id == saved.id));
    }

    #[test]
    fn listing_skips_hidden_folders_and_pictures() {
        let dir = tempfile::TempDir::new().unwrap();
        let places = places(&dir);
        create(&places, None, "A", "md", &with_assets("![x](assets/x.png)\n", Some(MEMORY), &[("assets/x.png", b"X")])).unwrap();
        fs::create_dir_all(places.root.join("Sub").join("Deeper")).unwrap();
        fs::write(places.root.join("Sub").join("Deeper").join("B.markdown"), "b\n").unwrap();
        fs::write(places.root.join("Sub").join("old.vtxdoc"), "{}").unwrap();
        fs::write(places.root.join("Sub").join("readme.txt"), "t\n").unwrap();
        fs::create_dir_all(places.root.join(".obsidian")).unwrap();
        fs::write(places.root.join(".obsidian").join("hidden.md"), "h\n").unwrap();
        let lib = list(&places).unwrap();
        let mut ids: Vec<&str> = lib.documents.iter().map(|d| d.id.as_str()).collect();
        ids.sort();
        assert_eq!(ids, vec!["A.md", "Sub/Deeper/B.markdown", "Sub/old.vtxdoc"]);
        assert_eq!(lib.folders.iter().map(|f| f.id.as_str()).collect::<Vec<_>>(), vec!["Sub", "Sub/Deeper"]);
        assert_eq!(lib.documents.iter().find(|d| d.id == "Sub/old.vtxdoc").unwrap().kind, "vtxdoc");
        assert_eq!(picture_paths("![a](x.png) ![b\\]c](<y z.png> \"t\") `![no]` ![d](https://e/f.png)\n![broken](\n"), vec!["x.png", "y z.png", "https://e/f.png"]);
    }

    #[test]
    fn documents_among_start_arguments() {
        let dir = tempfile::TempDir::new().unwrap();
        let note = dir.path().join("my note.md");
        let other = dir.path().join("Other.MARKDOWN");
        let old = dir.path().join("old.vtxdoc");
        let text = dir.path().join("readme.txt");
        for file in [&note, &other, &old, &text] {
            fs::write(file, "x").unwrap();
        }
        let s = |path: &Path| path.to_string_lossy().into_owned();
        let exe = "C:\\Program Files\\VisualTeX\\VisualTeX.exe".to_string();
        // the program itself, flags, other files, missing files and repeats are not documents to open
        let args = vec![exe.clone(), "--office-autostart".into(), s(&note), s(&text), s(&dir.path().join("missing.md")), s(&other), s(&note), s(&old)];
        assert_eq!(document_arguments(&args, None), vec![note.clone(), other.clone(), old.clone()]);
        assert!(document_arguments(vec![exe.clone()], None).is_empty());
        assert!(document_arguments(Vec::<String>::new(), None).is_empty());
        // the first argument is the program even when it is a document's path
        assert!(document_arguments(vec![s(&note)], None).is_empty());
        // a relative path is taken from the folder the start was made in
        assert_eq!(document_arguments(vec![exe.clone(), "my note.md".into()], Some(dir.path())), vec![note.clone()]);
        assert!(document_arguments(vec![exe, "my note.md".into()], None).is_empty());
        assert!(is_document_path(Path::new("a.Md")) && !is_document_path(Path::new("a.mdx")) && !is_document_path(Path::new("md")));
    }

    #[test]
    fn the_earlier_library_is_listed_for_conversion() {
        let dir = tempfile::TempDir::new().unwrap();
        let old = dir.path().join("documents");
        fs::create_dir_all(&old).unwrap();
        fs::write(old.join("index.json"), r#"{"folders":[{"id":"f1","name":"课程","created":1}],"documents":[
            {"id":"aaaaaaaaaaaaaaaa","title":"First","titleSet":true,"folder":"f1","created":1,"modified":5,"trashed":null},
            {"id":"bbbbbbbbbbbbbbbb","title":"Deleted","created":1,"modified":6,"trashed":3},
            {"id":"cccccccccccccccc","title":"No file","created":1,"modified":7}]}"#).unwrap();
        fs::write(old.join("aaaaaaaaaaaaaaaa.vtxdoc"), "A").unwrap();
        fs::write(old.join("bbbbbbbbbbbbbbbb.vtxdoc"), "B").unwrap();
        fs::write(old.join("dddddddddddddddd.vtxdoc"), "D").unwrap();
        let documents = legacy_list(&old);
        assert_eq!(documents.iter().map(|d| d.id.as_str()).collect::<Vec<_>>().len(), 2);
        let first = documents.iter().find(|d| d.id == "aaaaaaaaaaaaaaaa").unwrap();
        assert_eq!((first.title.as_str(), first.folder.as_deref()), ("First", Some("课程")));
        assert!(documents.iter().any(|d| d.id == "dddddddddddddddd"));
        assert_eq!(legacy_read(&old, "aaaaaaaaaaaaaaaa").unwrap(), "A");
        assert!(legacy_read(&old, "../index").is_err());
        assert!(legacy_list(&dir.path().join("missing")).is_empty());
    }
}
