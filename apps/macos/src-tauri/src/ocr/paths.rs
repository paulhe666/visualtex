//! Where the OCR runtime lives: the runtime folders and the worker script.

use super::*;

pub(super) fn runtime_paths(app: &AppHandle) -> Result<RuntimePaths, String> {
    let root = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve application data directory: {error}"))?
        .join("ocr-runtime");
    let offline_python = if cfg!(windows) {
        root.join("python").join("python.exe")
    } else {
        root.join("python").join("bin").join("python3")
    };
    let legacy_python = if cfg!(windows) {
        root.join("venv").join("Scripts").join("python.exe")
    } else {
        root.join("venv").join("bin").join("python")
    };
    let python = if offline_python.exists() || !legacy_python.exists() {
        offline_python
    } else {
        legacy_python
    };
    Ok(RuntimePaths {
        root: root.clone(),
        python,
        input: root.join("input"),
        processed: root.join("processed"),
        logs: root.join("logs"),
        cache: root.join("cache"),
        temp: root.join("tmp"),
    })
}
