//! The files an update leaves behind. The installer lists every file it writes into the install folder
//! (installed-files.txt, hooks.nsh VisualTeXRecordInstalledFile) and keeps the list of the version it
//! replaced (installed-files.previous.txt). An update only overwrites files, and web asset names change
//! with every build, so at the next start the files only the previous version had are deleted here,
//! then its list. Nothing outside the install folder is touched, and nothing the installer did not list.

use std::collections::HashSet;
use std::fs;
use std::io::ErrorKind;
use std::path::{Component, Path, PathBuf};

pub(crate) const CURRENT_LIST: &str = "installed-files.txt";
pub(crate) const PREVIOUS_LIST: &str = "installed-files.previous.txt";
/// The installer lists the uninstaller last: a list that names it is complete.
const LAST_ENTRY: &str = "uninstall.exe";

/// The paths a list names, as the installer writes them (UTF-16 with a byte-order mark) or as a retry
/// list is written back (UTF-8).
fn read_list(path: &Path) -> std::io::Result<Vec<String>> {
    let bytes = fs::read(path)?;
    let text = if let Some(rest) = bytes.strip_prefix(&[0xFF, 0xFE]) {
        let units: Vec<u16> = rest
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        String::from_utf16_lossy(&units)
    } else {
        String::from_utf8_lossy(bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(&bytes)).into_owned()
    };
    Ok(text
        .lines()
        .map(|line| line.trim().replace('/', "\\"))
        .filter(|line| !line.is_empty())
        .collect())
}

/// A listed path that stays inside the install folder: relative, without a drive or "..".
fn safe_relative(entry: &str) -> Option<PathBuf> {
    if entry.contains(':') {
        return None;
    }
    let mut clean = PathBuf::new();
    for component in Path::new(entry).components() {
        match component {
            Component::Normal(part) => clean.push(part),
            Component::CurDir => {}
            _ => return None,
        }
    }
    (!clean.as_os_str().is_empty()).then_some(clean)
}

#[derive(Debug, Default, PartialEq)]
pub(crate) struct Cleanup {
    pub removed: Vec<String>,
    /// still there (in use): kept in the previous list and tried again at the next start
    pub failed: Vec<String>,
}

/// Deletes the files the previous version's list names and the current list does not, then the folders
/// that leaves empty. The previous list goes when everything was deleted; otherwise it keeps only what is
/// left. None when there is nothing to compare (no update since the last cleanup, or a development build).
pub(crate) fn remove_obsolete(install_dir: &Path) -> std::io::Result<Option<Cleanup>> {
    let previous_path = install_dir.join(PREVIOUS_LIST);
    let current_path = install_dir.join(CURRENT_LIST);
    if !previous_path.is_file() || !current_path.is_file() {
        return Ok(None);
    }
    let current: HashSet<String> = read_list(&current_path)?
        .into_iter()
        .map(|entry| entry.to_lowercase())
        .collect();
    if !current.contains(LAST_ENTRY) {
        // not a complete list: deleting against it could remove what this version needs
        return Ok(None);
    }
    let mut cleanup = Cleanup::default();
    let mut folders = HashSet::new();
    for entry in read_list(&previous_path)? {
        if current.contains(&entry.to_lowercase()) {
            continue;
        }
        let Some(relative) = safe_relative(&entry) else {
            continue;
        };
        let target = install_dir.join(&relative);
        match fs::symlink_metadata(&target) {
            Ok(metadata) if !metadata.is_dir() => match fs::remove_file(&target) {
                Ok(()) => cleanup.removed.push(entry),
                Err(error) if error.kind() == ErrorKind::NotFound => {}
                Err(_) => cleanup.failed.push(entry),
            },
            // already gone, or a folder: only listed files are removed
            _ => {}
        }
        let mut parent = relative.parent();
        while let Some(folder) = parent.filter(|folder| !folder.as_os_str().is_empty()) {
            folders.insert(folder.to_path_buf());
            parent = folder.parent();
        }
    }
    let mut folders: Vec<PathBuf> = folders.into_iter().collect();
    folders.sort_by_key(|folder| std::cmp::Reverse(folder.components().count()));
    for folder in folders {
        // removes only empty folders: anything else the folder still holds keeps it
        let _ = fs::remove_dir(install_dir.join(folder));
    }
    if cleanup.failed.is_empty() {
        fs::remove_file(&previous_path)?;
    } else {
        fs::write(&previous_path, cleanup.failed.join("\r\n"))?;
    }
    Ok(Some(cleanup))
}

/// Once per start, from the installed program's folder, off the startup path.
pub(crate) fn start() {
    let Some(install_dir) = std::env::current_exe()
        .ok()
        .and_then(|executable| executable.parent().map(Path::to_path_buf))
    else {
        return;
    };
    std::thread::spawn(move || match remove_obsolete(&install_dir) {
        Ok(Some(cleanup)) => crate::app_lifecycle::append_lifecycle_log(format!(
            "removed {} file(s) only the previous version had; {} still in use",
            cleanup.removed.len(),
            cleanup.failed.len()
        )),
        Ok(None) => {}
        Err(error) => crate::app_lifecycle::append_lifecycle_log(format!(
            "could not remove the previous version's files: {error}"
        )),
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    fn utf16_list(entries: &[&str]) -> Vec<u8> {
        let mut bytes = vec![0xFF, 0xFE];
        for entry in entries {
            for unit in format!("{entry}\r\n").encode_utf16() {
                bytes.extend_from_slice(&unit.to_le_bytes());
            }
        }
        bytes
    }

    fn touch(root: &Path, relative: &str) {
        let path = root.join(relative);
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        fs::write(path, b"x").unwrap();
    }

    #[test]
    fn files_only_the_previous_version_had_are_removed() {
        let temp = TempDir::new().unwrap();
        let root = temp.path().join("VisualTeX");
        for file in [
            "visualtex.exe",
            "uninstall.exe",
            "office\\assets\\app-OLD.js",
            "office\\assets\\app-NEW.js",
            "windows-office\\VisualTeX-NativeRender-x64.msi",
            "logs\\visualtex.log",
        ] {
            touch(&root, file);
        }
        fs::write(
            root.join(PREVIOUS_LIST),
            utf16_list(&[
                "visualtex.exe",
                "office\\assets\\app-OLD.js",
                "windows-office\\VisualTeX-NativeRender-x64.msi",
                "uninstall.exe",
            ]),
        )
        .unwrap();
        fs::write(
            root.join(CURRENT_LIST),
            utf16_list(&["visualtex.exe", "office\\assets\\app-NEW.js", "uninstall.exe"]),
        )
        .unwrap();

        let cleanup = remove_obsolete(&root).unwrap().unwrap();
        assert_eq!(cleanup.failed, Vec::<String>::new());
        assert_eq!(cleanup.removed.len(), 2);
        assert!(!root.join("office\\assets\\app-OLD.js").exists());
        assert!(root.join("office\\assets\\app-NEW.js").is_file());
        // the folder only the previous version used goes once it is empty
        assert!(!root.join("windows-office").exists());
        // files the installer never listed (logs) stay
        assert!(root.join("logs\\visualtex.log").is_file());
        assert!(!root.join(PREVIOUS_LIST).exists());
        assert!(remove_obsolete(&root).unwrap().is_none());
    }

    #[test]
    fn nothing_outside_the_install_folder_and_nothing_against_an_incomplete_list() {
        let temp = TempDir::new().unwrap();
        let root = temp.path().join("VisualTeX");
        touch(&root, "visualtex.exe");
        touch(temp.path(), "outside.txt");
        touch(&root, "old.dll");
        fs::write(
            root.join(PREVIOUS_LIST),
            utf16_list(&["..\\outside.txt", "C:\\Windows\\win.ini", "old.dll", "visualtex.exe"]),
        )
        .unwrap();
        // an installation that stopped halfway: no uninstaller in its list
        fs::write(root.join(CURRENT_LIST), utf16_list(&["visualtex.exe"])).unwrap();
        assert!(remove_obsolete(&root).unwrap().is_none());
        assert!(root.join("old.dll").is_file());

        fs::write(root.join(CURRENT_LIST), utf16_list(&["visualtex.exe", "uninstall.exe"])).unwrap();
        let cleanup = remove_obsolete(&root).unwrap().unwrap();
        assert_eq!(cleanup.removed, vec!["old.dll".to_string()]);
        assert!(temp.path().join("outside.txt").is_file());
        assert!(root.join("visualtex.exe").is_file());
    }

    #[test]
    fn names_compare_without_case_and_with_either_slash() {
        let temp = TempDir::new().unwrap();
        let root = temp.path().join("VisualTeX");
        touch(&root, "Office\\Dialog\\index.html");
        fs::write(root.join(PREVIOUS_LIST), "office/dialog/INDEX.html\r\n").unwrap();
        fs::write(root.join(CURRENT_LIST), utf16_list(&["OFFICE\\dialog\\index.html", "uninstall.exe"])).unwrap();
        let cleanup = remove_obsolete(&root).unwrap().unwrap();
        assert!(cleanup.removed.is_empty());
        assert!(root.join("Office\\Dialog\\index.html").is_file());
    }
}
