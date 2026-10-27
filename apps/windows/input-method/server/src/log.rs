//! A small log for development: %LOCALAPPDATA%\VisualTeX\input-server.log, kept under 512 KB.

use std::io::Write;
use std::path::PathBuf;

fn path() -> Option<PathBuf> {
    let base = std::env::var_os("LOCALAPPDATA")?;
    let directory = PathBuf::from(base).join("VisualTeX");
    let _ = std::fs::create_dir_all(&directory);
    Some(directory.join("input-server.log"))
}

pub fn write(line: &str) {
    let Some(path) = path() else { return };
    if std::fs::metadata(&path).map(|meta| meta.len() > 512 * 1024).unwrap_or(false) {
        let _ = std::fs::remove_file(&path);
    }
    let Ok(mut file) = std::fs::OpenOptions::new().create(true).append(true).open(&path) else { return };
    let millis = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default();
    let _ = writeln!(file, "{millis} {line}");
}
