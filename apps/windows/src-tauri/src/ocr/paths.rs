//! Where the OCR runtime lives: the storage configuration, the runtime folders, the preferred model and the worker script.

use super::*;

pub(super) fn ocr_storage_config_path(app: &AppHandle) -> Result<PathBuf, String> {
    #[cfg(windows)]
    if let Some(appdata) = env::var_os("APPDATA") {
        return Ok(PathBuf::from(appdata)
            .join("VisualTeX")
            .join(storage::STORAGE_CONFIG_FILE));
    }
    Ok(app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve application data directory: {error}"))?
        .join(storage::STORAGE_CONFIG_FILE))
}

pub(super) fn legacy_ocr_runtime_root(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve application data directory: {error}"))?
        .join("ocr-runtime"))
}

pub(super) fn default_ocr_runtime_root(app: &AppHandle) -> Result<PathBuf, String> {
    #[cfg(windows)]
    if let Some(local_appdata) = env::var_os("LOCALAPPDATA") {
        return Ok(PathBuf::from(local_appdata)
            .join("VisualTeXData")
            .join("ocr-runtime"));
    }
    legacy_ocr_runtime_root(app)
}

pub(super) fn protected_ocr_storage_roots(app: &AppHandle) -> Vec<PathBuf> {
    let mut roots = Vec::new();
    if let Ok(executable) = env::current_exe() {
        if let Some(parent) = executable.parent() {
            roots.push(parent.to_path_buf());
        }
    }
    if let Ok(resources) = app.path().resource_dir() {
        roots.push(resources);
    }
    roots
}

pub(super) fn runtime_paths_from_resolution(
    resolution: storage::StorageResolution,
) -> RuntimePaths {
    let root = resolution.root;
    let offline_python = if cfg!(windows) {
        root.join("python").join("python.exe")
    } else {
        root.join("python").join("bin").join("python3")
    };
    let venv = root.join("venv");
    let legacy_python = if cfg!(windows) {
        venv.join("Scripts").join("python.exe")
    } else {
        venv.join("bin").join("python")
    };
    let python = if offline_python.exists() || !legacy_python.exists() {
        offline_python
    } else {
        legacy_python.clone()
    };
    RuntimePaths {
        root: root.clone(),
        venv,
        python,
        input: root.join("input"),
        processed: root.join("processed"),
        logs: root.join("logs"),
        cache: root.join("cache"),
        temp: root.join("tmp"),
        storage_config_path: resolution.config_path,
        storage_source: resolution.source,
        storage_managed: resolution.managed,
    }
}

pub(super) fn runtime_paths(app: &AppHandle) -> Result<RuntimePaths, String> {
    let resolution = storage::resolve(
        &ocr_storage_config_path(app)?,
        &legacy_ocr_runtime_root(app)?,
        &default_ocr_runtime_root(app)?,
        &protected_ocr_storage_roots(app),
    )?;
    Ok(runtime_paths_from_resolution(resolution))
}

pub(super) fn read_preferred_ocr_model(paths: &RuntimePaths) -> Option<String> {
    let model = fs::read_to_string(paths.root.join(OCR_PREFERRED_MODEL_FILE))
        .ok()?
        .trim()
        .to_string();
    ALLOWED_MODELS
        .contains(&model.as_str())
        .then_some(model)
}

pub(super) fn write_preferred_ocr_model(paths: &RuntimePaths, model: &str) -> Result<(), String> {
    if !ALLOWED_MODELS.contains(&model) {
        return Err(format!("Unsupported PP-FormulaNet model: {model}"));
    }
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    fs::write(paths.root.join(OCR_PREFERRED_MODEL_FILE), model.as_bytes())
        .map_err(|error| format!("Unable to save the preferred OCR model: {error}"))
}

pub(super) fn worker_script_path(app: &AppHandle) -> Result<PathBuf, String> {
    if let Ok(path) = app.path().resolve("ocr/worker.py", BaseDirectory::Resource) {
        if path.exists() {
            return Ok(path);
        }
    }

    #[cfg(debug_assertions)]
    {
        let development_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("ocr")
            .join("worker.py");
        if development_path.exists() {
            return Ok(development_path);
        }
    }

    Err("Unable to locate bundled OCR worker.py".to_string())
}
