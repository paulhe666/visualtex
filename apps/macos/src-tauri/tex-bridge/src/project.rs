//! A LaTeX project's files (the folder is the truth): listed, read and written for the editor, and copied into
//! the document's build folder (the user's folder gets no .aux/.log/.synctex).

use std::collections::HashMap;
use std::path::{Component, Path, PathBuf};
use std::time::UNIX_EPOCH;

use serde::Serialize;

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectFile {
    /// "/"-separated, relative to the project folder
    pub path: String,
    pub size: u64,
    /// "tex" (sources the editor opens), "bib", "image", "other"
    pub kind: String,
}

const SKIP_DIRS: &[&str] = &[".git", ".svn", ".hg", "node_modules", "__MACOSX", ".vscode", ".idea", "_minted", ".texpadtmp", "auto"];
/// TeX's own output: never copied (a stale .aux from the user's own runs would confuse the first run here).
const OUTPUT_EXT: &[&str] = &[
    "aux", "log", "out", "toc", "lof", "lot", "fls", "fdb_latexmk", "xdv", "dvi", "blg", "bcf", "nav", "snm", "vrb", "synctex", "idx", "ilg", "ind", "glo", "gls", "glg", "ist", "loa", "thm", "brf", "run.xml",
];
const MAX_FILES: usize = 4000;
const MAX_DEPTH: usize = 8;
const MAX_FILE_SIZE: u64 = 200 * 1024 * 1024;

pub fn kind_of(path: &str) -> &'static str {
    let lower = path.to_ascii_lowercase();
    let ext = lower.rsplit('.').next().unwrap_or("");
    match ext {
        "tex" | "ltx" | "sty" | "cls" | "def" | "cfg" | "clo" | "fd" | "dtx" | "ins" | "bbx" | "cbx" | "lbx" | "dbx" | "bst" | "tikz" | "pgf" | "lua" | "txt" | "md" | "csv" | "dat" | "latexmkrc" => "tex",
        "bib" => "bib",
        "png" | "jpg" | "jpeg" | "pdf" | "eps" | "ps" | "svg" | "gif" | "bmp" | "webp" => "image",
        _ => "other",
    }
}

fn is_output(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    if lower.ends_with(".synctex.gz") || lower.ends_with(".synctex(busy)") || lower.ends_with(".run.xml") {
        return true;
    }
    let ext = lower.rsplit('.').next().unwrap_or("");
    OUTPUT_EXT.contains(&ext)
}

/// The project's files (outputs and tool folders left out), sorted.
pub fn list_files(root: &Path) -> Vec<ProjectFile> {
    let mut out = vec![];
    walk(root, root, 0, &mut out);
    out.sort_by(|a, b| a.path.cmp(&b.path));
    out
}

fn walk(root: &Path, dir: &Path, depth: usize, out: &mut Vec<ProjectFile>) {
    if depth > MAX_DEPTH || out.len() >= MAX_FILES {
        return;
    }
    let Ok(entries) = std::fs::read_dir(dir) else { return };
    let mut entries: Vec<_> = entries.flatten().collect();
    entries.sort_by_key(|e| e.file_name());
    for e in entries {
        let name = e.file_name().to_string_lossy().to_string();
        if name.starts_with('.') || name.starts_with("_vtx_") {
            continue;
        }
        let Ok(meta) = e.metadata() else { continue };
        let path = e.path();
        if meta.is_dir() {
            if SKIP_DIRS.contains(&name.as_str()) || name.starts_with("_minted") {
                continue;
            }
            walk(root, &path, depth + 1, out);
        } else if meta.is_file() && !is_output(&name) && meta.len() <= MAX_FILE_SIZE {
            if let Ok(rel) = path.strip_prefix(root) {
                let rel = rel.to_string_lossy().replace('\\', "/");
                out.push(ProjectFile { kind: kind_of(&rel).to_string(), path: rel, size: meta.len() });
                if out.len() >= MAX_FILES {
                    return;
                }
            }
        }
    }
}

/// A relative path that stays inside the project ("a/b.tex"); None for absolute paths or "..".
pub fn safe_rel(rel: &str) -> Option<PathBuf> {
    let p = Path::new(rel);
    let mut out = PathBuf::new();
    for c in p.components() {
        match c {
            Component::Normal(s) => out.push(s),
            Component::CurDir => {}
            _ => return None,
        }
    }
    if out.as_os_str().is_empty() {
        None
    } else {
        Some(out)
    }
}

/// The main file: the given one, else main.tex, else the .tex nearest the top with \documentclass and
/// \begin{document}.
pub fn find_main(root: &Path, files: &[ProjectFile]) -> Option<String> {
    let mut best: Option<(i64, String)> = None;
    for f in files {
        if !f.path.to_ascii_lowercase().ends_with(".tex") {
            continue;
        }
        let Ok(text) = std::fs::read(root.join(&f.path)) else { continue };
        let text = String::from_utf8_lossy(&text);
        let code = crate::engine::strip_comments(&text);
        let mut score: i64 = 0;
        if code.contains("\\documentclass") {
            score += 100;
        }
        if code.contains("\\begin{document}") {
            score += 50;
        }
        if code.contains("{subfiles}") || code.contains("{standalone}") {
            score -= 60;
        }
        let name = f.path.rsplit('/').next().unwrap_or("").to_ascii_lowercase();
        if ["main.tex", "thesis.tex", "paper.tex", "book.tex", "report.tex", "manuscript.tex"].contains(&name.as_str()) {
            score += 10;
        }
        score -= 5 * f.path.matches('/').count() as i64;
        if best.as_ref().map(|(s, _)| score > *s).unwrap_or(true) {
            best = Some((score, f.path.clone()));
        }
    }
    best.filter(|(s, _)| *s >= 100).map(|(_, p)| p)
}

/// What the build folder last got from the project: path → (size, modified).
pub type Manifest = HashMap<String, (u64, u128)>;

fn stamp(meta: &std::fs::Metadata) -> u128 {
    meta.modified().ok().and_then(|t| t.duration_since(UNIX_EPOCH).ok()).map(|d| d.as_nanos()).unwrap_or(0)
}

/// Copy the project into `work`: changed files only; `overlays` (unsaved editor text) win over the disk;
/// files gone from the project are removed. Returns the paths that changed.
pub fn sync(root: &Path, work: &Path, manifest: &mut Manifest, overlays: &HashMap<String, String>) -> std::io::Result<Vec<String>> {
    std::fs::create_dir_all(work)?;
    let files = list_files(root);
    let mut changed = vec![];
    let mut seen = std::collections::HashSet::new();
    for f in &files {
        let Some(rel) = safe_rel(&f.path) else { continue };
        seen.insert(f.path.clone());
        let dest = work.join(&rel);
        if let Some(text) = overlays.get(&f.path) {
            let same = std::fs::read(&dest).map(|b| b == text.as_bytes()).unwrap_or(false);
            if !same {
                if let Some(parent) = dest.parent() {
                    std::fs::create_dir_all(parent)?;
                }
                std::fs::write(&dest, text.as_bytes())?;
                changed.push(f.path.clone());
            }
            manifest.insert(f.path.clone(), (u64::MAX, 0));
            continue;
        }
        let src = root.join(&rel);
        let Ok(meta) = std::fs::metadata(&src) else { continue };
        let key = (meta.len(), stamp(&meta));
        if manifest.get(&f.path) == Some(&key) && dest.exists() {
            continue;
        }
        if let Some(parent) = dest.parent() {
            std::fs::create_dir_all(parent)?;
        }
        std::fs::copy(&src, &dest)?;
        manifest.insert(f.path.clone(), key);
        changed.push(f.path.clone());
    }
    // overlays for files not on disk yet (a new file not saved)
    for (path, text) in overlays {
        if seen.contains(path) {
            continue;
        }
        if let Some(rel) = safe_rel(path) {
            let dest = work.join(rel);
            if let Some(parent) = dest.parent() {
                std::fs::create_dir_all(parent)?;
            }
            if std::fs::read(&dest).map(|b| b != text.as_bytes()).unwrap_or(true) {
                std::fs::write(&dest, text.as_bytes())?;
                changed.push(path.clone());
            }
            manifest.insert(path.clone(), (u64::MAX, 0));
            seen.insert(path.clone());
        }
    }
    let gone: Vec<String> = manifest.keys().filter(|k| !seen.contains(*k)).cloned().collect();
    for g in gone {
        if let Some(rel) = safe_rel(&g) {
            let _ = std::fs::remove_file(work.join(rel));
        }
        manifest.remove(&g);
        changed.push(g);
    }
    Ok(changed)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rel_paths() {
        assert_eq!(safe_rel("a/b.tex"), Some(PathBuf::from("a").join("b.tex")));
        assert_eq!(safe_rel("./a.tex"), Some(PathBuf::from("a.tex")));
        assert!(safe_rel("../x").is_none());
        assert!(safe_rel("/etc/passwd").is_none());
        assert!(safe_rel("").is_none());
    }

    #[test]
    fn sync_copies_changes_only() {
        let base = std::env::temp_dir().join(format!("vtx-sync-{}", std::process::id()));
        let (root, work) = (base.join("root"), base.join("work"));
        let _ = std::fs::remove_dir_all(&base);
        std::fs::create_dir_all(root.join("chapters")).unwrap();
        std::fs::write(root.join("main.tex"), "\\documentclass{article}\\begin{document}\\input{chapters/a}\\end{document}").unwrap();
        std::fs::write(root.join("chapters/a.tex"), "A").unwrap();
        std::fs::write(root.join("main.aux"), "stale").unwrap();
        let mut m = Manifest::new();
        let first = sync(&root, &work, &mut m, &HashMap::new()).unwrap();
        assert_eq!(first.len(), 2);
        assert!(!work.join("main.aux").exists());
        assert!(sync(&root, &work, &mut m, &HashMap::new()).unwrap().is_empty());
        let mut over = HashMap::new();
        over.insert("chapters/a.tex".to_string(), "B".to_string());
        assert_eq!(sync(&root, &work, &mut m, &over).unwrap(), vec!["chapters/a.tex".to_string()]);
        assert_eq!(std::fs::read_to_string(work.join("chapters/a.tex")).unwrap(), "B");
        std::fs::remove_file(root.join("chapters/a.tex")).unwrap();
        let gone = sync(&root, &work, &mut m, &HashMap::new()).unwrap();
        assert_eq!(gone, vec!["chapters/a.tex".to_string()]);
        assert_eq!(find_main(&root, &list_files(&root)).as_deref(), Some("main.tex"));
        let _ = std::fs::remove_dir_all(&base);
    }
}
