//! Installing the runtime from the offline bundle.

use super::*;

pub(super) fn install_runtime_inner(
    app: &AppHandle,
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
) -> Result<OcrRuntimeStatus, String> {
    stop_worker(worker, worker_pid)?;
    let paths = runtime_paths(app)?;
    crate::ocr_offline::install_bundle(app, &paths.root, |stage, percent, message, detail| {
        emit_progress(app, stage, percent, message, detail);
    })?;

    emit_progress(app, "verify", 97, "正在验证离线 PP-FormulaNet 接口", None);
    let status = get_runtime_status_inner(app, true)?;
    if !status.installed {
        return Err(status.message);
    }
    if !status
        .installed_models
        .iter()
        .any(|model| model == crate::ocr_offline::OFFLINE_DEFAULT_MODEL)
    {
        return Err("The bundled PP-FormulaNet M model was not installed".to_string());
    }
    emit_progress(
        app,
        "complete",
        100,
        "OCR 离线运行环境安装完成",
        Some("Python、PaddleOCR 与默认 M 模型均已内置，无需联网".to_string()),
    );
    Ok(status)
}
