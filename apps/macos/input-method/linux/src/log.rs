//! The input method's diagnostics: $XDG_STATE_HOME/visualtex/ibus.log (kept under half a megabyte),
//! as VisualTeXInput.log on macOS. Formulas and typed text are not logged.

use std::io::Write;
use std::path::PathBuf;
use std::sync::OnceLock;

static PATH: OnceLock<Option<PathBuf>> = OnceLock::new();
const MAX_BYTES: u64 = 512 * 1024;

fn path() -> Option<&'static PathBuf> {
    PATH.get_or_init(|| {
        let base = std::env::var_os("XDG_STATE_HOME")
            .map(PathBuf::from)
            .or_else(|| std::env::var_os("HOME").map(|home| PathBuf::from(home).join(".local/state")))?;
        let directory = base.join("visualtex");
        std::fs::create_dir_all(&directory).ok()?;
        Some(directory.join("ibus.log"))
    })
    .as_ref()
}

pub fn write(line: &str) {
    let Some(path) = path() else { return };
    if std::fs::metadata(path).map(|metadata| metadata.len() > MAX_BYTES).unwrap_or(false) {
        let _ = std::fs::remove_file(path);
    }
    let millis = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|value| value.as_millis())
        .unwrap_or_default();
    if let Ok(mut file) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(file, "epochMs={millis} {line}");
    }
}
