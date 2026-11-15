//! The runtime status: the persisted probe, the status the app shows, and the cached status.

use super::*;

pub(super) fn runtime_probe_cache_path(paths: &RuntimePaths) -> PathBuf {
    paths.root.join(OCR_RUNTIME_PROBE_CACHE_FILE)
}

pub(super) fn read_persisted_runtime_probe(paths: &RuntimePaths) -> Option<RuntimeProbe> {
    let content = fs::read(runtime_probe_cache_path(paths)).ok()?;
    let cache: RuntimeProbeCache = serde_json::from_slice(&content).ok()?;
    if cache.schema_version != OCR_RUNTIME_PROBE_CACHE_SCHEMA
        || cache.python_path != paths.python.display().to_string()
    {
        return None;
    }
    let (python_size, python_modified_ms) = python_runtime_signature(&paths.python).ok()?;
    if cache.python_size != python_size || cache.python_modified_ms != python_modified_ms {
        return None;
    }
    validate_runtime_probe(&cache.probe).ok()?;
    Some(cache.probe)
}

pub(super) fn write_persisted_runtime_probe(paths: &RuntimePaths, probe: &RuntimeProbe) -> Result<(), String> {
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    let (python_size, python_modified_ms) = python_runtime_signature(&paths.python)?;
    let cache = RuntimeProbeCache {
        schema_version: OCR_RUNTIME_PROBE_CACHE_SCHEMA,
        python_path: paths.python.display().to_string(),
        python_size,
        python_modified_ms,
        probe: probe.clone(),
    };
    let destination = runtime_probe_cache_path(paths);
    let temporary = paths.root.join(format!(
        ".{OCR_RUNTIME_PROBE_CACHE_FILE}.{}.tmp",
        std::process::id()
    ));
    fs::write(
        &temporary,
        serde_json::to_vec_pretty(&cache)
            .map_err(|error| format!("Unable to encode OCR runtime cache: {error}"))?,
    )
    .map_err(|error| format!("Unable to write OCR runtime cache: {error}"))?;
    if destination.exists() {
        fs::remove_file(&destination).ok();
    }
    fs::rename(&temporary, &destination)
        .map_err(|error| format!("Unable to activate OCR runtime cache: {error}"))
}

pub(super) fn get_runtime_status_inner(
    app: &AppHandle,
    force_refresh: bool,
) -> Result<OcrRuntimeStatus, String> {
    let paths = runtime_paths(app)?;
    let offline_bundle_available = crate::ocr_offline::bundle_available(app);
    let installed_models = crate::ocr_offline::installed_models(&paths.root);
    let default_model_installed = installed_models
        .iter()
        .any(|model| model == crate::ocr_offline::OFFLINE_DEFAULT_MODEL);
    let structural_error = if !paths.python.is_file() {
        Some("OCR offline runtime is not installed".to_string())
    } else if !default_model_installed {
        Some("The bundled PP-FormulaNet M model is not installed".to_string())
    } else {
        None
    };

    if structural_error.is_none() && !force_refresh {
        let cached_probe = read_persisted_runtime_probe(&paths);
        return Ok(OcrRuntimeStatus {
            installed: true,
            python_path: Some(paths.python.display().to_string()),
            python_version: cached_probe
                .as_ref()
                .map(|probe| probe.python_version.clone()),
            paddle_version: cached_probe
                .as_ref()
                .map(|probe| probe.paddle_version.clone()),
            paddleocr_version: cached_probe
                .as_ref()
                .map(|probe| probe.paddleocr_version.clone()),
            runtime_path: paths.root.display().to_string(),
            offline_bundle_available,
            installed_models,
            default_model: crate::ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: if cached_probe.is_some() {
                "PaddleOCR formula runtime is ready; the default M model can be reused in memory"
                    .to_string()
            } else {
                "OCR runtime files are ready; the default M model will be verified during background preloading"
                    .to_string()
            },
        });
    }

    let probe_result = structural_error.map_or_else(|| probe_runtime(&paths), Err);
    match probe_result {
        Ok(probe) => Ok(OcrRuntimeStatus {
            installed: true,
            python_path: Some(paths.python.display().to_string()),
            python_version: Some(probe.python_version),
            paddle_version: Some(probe.paddle_version),
            paddleocr_version: Some(probe.paddleocr_version),
            runtime_path: paths.root.display().to_string(),
            offline_bundle_available,
            installed_models,
            default_model: crate::ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: "PaddleOCR formula runtime is ready for offline recognition".to_string(),
        }),
        Err(error) => Ok(OcrRuntimeStatus {
            installed: false,
            python_path: paths
                .python
                .exists()
                .then(|| paths.python.display().to_string()),
            python_version: None,
            paddle_version: None,
            paddleocr_version: None,
            runtime_path: paths.root.display().to_string(),
            offline_bundle_available,
            installed_models,
            default_model: crate::ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: if offline_bundle_available {
                format!("Offline OCR package is ready to install. Current runtime: {error}")
            } else {
                error
            },
        }),
    }
}

pub(super) fn read_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
) -> Result<Option<OcrRuntimeStatus>, String> {
    cache
        .lock()
        .map(|status| status.clone())
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())
}

pub(super) fn write_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    status: Option<OcrRuntimeStatus>,
) -> Result<(), String> {
    let mut guard = cache
        .lock()
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())?;
    *guard = status;
    Ok(())
}
