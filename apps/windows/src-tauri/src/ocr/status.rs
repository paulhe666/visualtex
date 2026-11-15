//! The runtime status: the status manifest, the probe turned into a status, and the cached status.

use super::*;

pub(super) fn runtime_status_manifest_path(paths: &RuntimePaths) -> PathBuf {
    paths.root.join(OCR_RUNTIME_STATUS_FILE)
}

pub(super) fn write_runtime_status_manifest(paths: &RuntimePaths, probe: &RuntimeProbe) -> Result<(), String> {
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    let manifest = RuntimeStatusManifest {
        schema_version: OCR_RUNTIME_STATUS_SCHEMA,
        python_path: paths.python.display().to_string(),
        python_version: probe.python_version.clone(),
        paddle_version: probe.paddle_version.clone(),
        paddleocr_version: probe.paddleocr_version.clone(),
    };
    let content = serde_json::to_vec_pretty(&manifest)
        .map_err(|error| format!("Unable to serialize OCR runtime status: {error}"))?;
    let target = runtime_status_manifest_path(paths);
    let temporary = target.with_extension("json.tmp");
    fs::write(&temporary, content)
        .map_err(|error| format!("Unable to cache OCR runtime status: {error}"))?;
    if target.exists() {
        fs::remove_file(&target)
            .map_err(|error| format!("Unable to replace OCR runtime status: {error}"))?;
    }
    fs::rename(&temporary, &target)
        .map_err(|error| format!("Unable to publish OCR runtime status: {error}"))
}

pub(super) fn read_runtime_status_manifest(paths: &RuntimePaths) -> Option<RuntimeProbe> {
    let content = fs::read(runtime_status_manifest_path(paths)).ok()?;
    let manifest: RuntimeStatusManifest = serde_json::from_slice(&content).ok()?;
    if manifest.schema_version != OCR_RUNTIME_STATUS_SCHEMA
        || manifest.python_path != paths.python.display().to_string()
        || manifest.paddle_version != PADDLE_VERSION
        || manifest.paddleocr_version != PADDLEOCR_VERSION
    {
        return None;
    }
    Some(RuntimeProbe {
        python_version: manifest.python_version,
        paddle_version: manifest.paddle_version,
        paddleocr_version: manifest.paddleocr_version,
    })
}

pub(super) fn runtime_bundle_available(app: &AppHandle) -> bool {
    #[cfg(windows)]
    {
        python_bundle::bundle_available(app)
    }
    #[cfg(not(windows))]
    {
        crate::ocr_offline::bundle_available(app)
    }
}

pub(super) fn model_inventory(app: &AppHandle, paths: &RuntimePaths) -> (Vec<String>, Vec<String>, bool) {
    #[cfg(windows)]
    {
        let inventory = models::inspect_models(&paths.root).unwrap_or_default();
        let catalog_available = models::load_catalog(app).is_ok();
        (inventory.installed, inventory.damaged, catalog_available)
    }
    #[cfg(not(windows))]
    {
        (
            crate::ocr_offline::installed_models(&paths.root),
            Vec::new(),
            false,
        )
    }
}

pub(super) fn runtime_status_from_probe(
    app: &AppHandle,
    paths: &RuntimePaths,
    result: Result<RuntimeProbe, String>,
) -> OcrRuntimeStatus {
    let runtime_bundle_available = runtime_bundle_available(app);
    let offline_bundle_available = runtime_bundle_available;
    let (installed_models, damaged_models, model_catalog_available) = model_inventory(app, paths);
    match result {
        Ok(probe) => {
            let default_cached = installed_models
                .iter()
                .any(|model| model == crate::ocr_offline::OFFLINE_DEFAULT_MODEL);
            OcrRuntimeStatus {
                installed: true,
                python_path: Some(paths.python.display().to_string()),
                python_version: Some(probe.python_version),
                paddle_version: Some(probe.paddle_version),
                paddleocr_version: Some(probe.paddleocr_version),
                runtime_path: paths.root.display().to_string(),
                storage_config_path: paths.storage_config_path.display().to_string(),
                storage_source: paths.storage_source.clone(),
                storage_managed: paths.storage_managed || storage::marker_is_valid(&paths.root),
                storage_available_bytes: storage::available_space_bytes(&paths.root),
                storage_persistent_across_uninstall: true,
                runtime_bundle_available,
                offline_bundle_available,
                installed_models,
                damaged_models,
                model_catalog_available,
                default_model: DEFAULT_OCR_MODEL.to_string(),
                message: if default_cached {
                    "PaddleOCR formula runtime is ready; the M model is installed".to_string()
                } else {
                    "PaddleOCR formula runtime is ready; no default model is installed".to_string()
                },
            }
        }
        Err(error) => OcrRuntimeStatus {
            installed: false,
            python_path: paths
                .python
                .exists()
                .then(|| paths.python.display().to_string()),
            python_version: None,
            paddle_version: None,
            paddleocr_version: None,
            runtime_path: paths.root.display().to_string(),
            storage_config_path: paths.storage_config_path.display().to_string(),
            storage_source: paths.storage_source.clone(),
            storage_managed: paths.storage_managed || storage::marker_is_valid(&paths.root),
            storage_available_bytes: storage::available_space_bytes(&paths.root),
            storage_persistent_across_uninstall: true,
            runtime_bundle_available,
            offline_bundle_available,
            installed_models,
            damaged_models,
            model_catalog_available,
            default_model: DEFAULT_OCR_MODEL.to_string(),
            message: if runtime_bundle_available {
                format!("The bundled OCR runtime is ready to install. Current runtime: {error}")
            } else {
                error
            },
        },
    }
}

pub(super) fn get_runtime_status_fast(app: &AppHandle) -> Result<OcrRuntimeStatus, String> {
    let paths = runtime_paths(app)?;
    Ok(runtime_status_from_probe(
        app,
        &paths,
        probe_runtime_from_files(&paths),
    ))
}

pub(super) fn get_runtime_status_inner(app: &AppHandle) -> Result<OcrRuntimeStatus, String> {
    let paths = runtime_paths(app)?;
    Ok(runtime_status_from_probe(
        app,
        &paths,
        probe_runtime(&paths),
    ))
}

pub(super) fn should_run_full_runtime_probe(force_refresh: bool, installing: bool) -> bool {
    force_refresh && !installing
}

pub(super) fn reconcile_interrupted_install_snapshot(
    snapshot: &mut InstallSnapshot,
    installation_is_running: bool,
) -> bool {
    if installation_is_running
        || !matches!(
            snapshot.state,
            InstallState::Installing
                | InstallState::DependenciesInstalled
                | InstallState::Verifying
        )
    {
        return false;
    }

    let verification = matches!(snapshot.state, InstallState::Verifying)
        || snapshot.current_step.as_deref() == Some("verify");
    let step = snapshot
        .current_step
        .clone()
        .unwrap_or_else(|| "unknown".to_string());
    snapshot.state = if verification {
        InstallState::VerificationFailed
    } else {
        InstallState::InstallFailed
    };
    snapshot.message = "上一次 OCR 安装被中断".to_string();
    snapshot.detail = Some("已完成的依赖和步骤已保留，可从当前步骤继续安装".to_string());
    snapshot.error = Some(format!(
        "The previous VisualTeX process exited while OCR installation step '{step}' was active."
    ));
    snapshot.touch();
    true
}
