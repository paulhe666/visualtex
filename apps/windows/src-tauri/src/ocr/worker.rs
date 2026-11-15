//! The PaddleOCR worker process: its JSON protocol, start, warm-up, recognition and stop.

use super::*;

impl Drop for OcrWorker {
    fn drop(&mut self) {
        let pid = self.child.id();
        let _ = self.child.kill();
        let _ = self.child.wait();
        let _ = self
            .pid_state
            .compare_exchange(pid, 0, Ordering::SeqCst, Ordering::SeqCst);
    }
}

pub(super) fn read_worker_json<R: BufRead>(
    reader: &mut R,
    closed_message: &str,
    response_name: &str,
) -> Result<Value, String> {
    loop {
        let mut bytes = Vec::new();
        let count = reader
            .read_until(b'\n', &mut bytes)
            .map_err(|error| format!("Unable to read {response_name}: {error}"))?;
        if count == 0 {
            return Err(closed_message.to_string());
        }

        while bytes
            .last()
            .is_some_and(|byte| matches!(*byte, b'\n' | b'\r'))
        {
            bytes.pop();
        }
        if bytes.iter().all(u8::is_ascii_whitespace) {
            continue;
        }

        let first_non_whitespace = bytes
            .iter()
            .copied()
            .find(|byte| !byte.is_ascii_whitespace());
        if first_non_whitespace != Some(b'{') {
            // Native dependencies occasionally write diagnostics to stdout.
            // Ignore those lines so they cannot corrupt the JSON protocol.
            continue;
        }

        match serde_json::from_slice(&bytes) {
            Ok(value) => return Ok(value),
            Err(error) => {
                // This fallback prevents a legacy Windows code page from
                // crashing the reader. New workers always emit ASCII-safe JSON.
                let output = String::from_utf8_lossy(&bytes);
                return serde_json::from_str(output.trim()).map_err(|lossy_error| {
                    format!(
                        "{response_name} returned invalid JSON: {error}; UTF-8-lossy parse: {lossy_error}; output={output:?}"
                    )
                });
            }
        }
    }
}

impl OcrWorker {
    pub(super) fn worker_failure(&mut self, message: impl AsRef<str>) -> String {
        let status = match self.child.try_wait() {
            Ok(Some(status)) => status.to_string(),
            Ok(None) => "still running".to_string(),
            Err(error) => format!("status unavailable: {error}"),
        };
        format!(
            "{}; worker_status={status}; log={}",
            message.as_ref(),
            self.log_path.display()
        )
    }

    pub(super) fn send(&mut self, app: &AppHandle, payload: &Value) -> Result<Value, String> {
        if let Some(status) = self.child.try_wait().map_err(|error| error.to_string())? {
            return Err(format!(
                "OCR worker exited unexpectedly: {status}; log={}",
                self.log_path.display()
            ));
        }

        serde_json::to_writer(&mut self.stdin, payload)
            .map_err(|error| format!("Unable to encode OCR request: {error}"))?;
        self.stdin
            .write_all(b"\n")
            .map_err(|error| self.worker_failure(format!("Unable to send OCR request: {error}")))?;
        self.stdin.flush().map_err(|error| {
            self.worker_failure(format!("Unable to flush OCR request: {error}"))
        })?;

        loop {
            let response = read_worker_json(
                &mut self.stdout,
                "OCR worker closed its output stream",
                "OCR response",
            )
            .map_err(|error| self.worker_failure(error))?;
            if response.get("event").and_then(Value::as_str) == Some("progress") {
                let _ = app.emit("ocr-recognition-progress", &response);
                if let Some(state) = app.try_state::<OcrState>() {
                    state.events.publish("ocr-recognition-progress", &response);
                }
                continue;
            }
            return Ok(response);
        }
    }
}

pub(super) fn cleanup_worker_temp(paths: &RuntimePaths) -> Result<(), String> {
    if paths.temp.exists() {
        fs::remove_dir_all(&paths.temp)
            .map_err(|error| format!("Unable to clean OCR temporary files: {error}"))?;
    }
    fs::create_dir_all(&paths.temp)
        .map_err(|error| format!("Unable to create OCR temporary directory: {error}"))
}

pub(super) fn terminate_worker_process(worker_pid: &AtomicU32) -> Result<bool, String> {
    let pid = worker_pid.swap(0, Ordering::SeqCst);
    if pid == 0 {
        return Ok(false);
    }

    #[cfg(unix)]
    {
        let result = unsafe { libc::kill(pid as i32, libc::SIGKILL) };
        if result == 0 {
            return Ok(true);
        }
        let error = std::io::Error::last_os_error();
        if error.raw_os_error() == Some(libc::ESRCH) {
            return Ok(false);
        }
        return Err(format!("Unable to terminate OCR worker {pid}: {error}"));
    }

    #[cfg(windows)]
    {
        let mut command = Command::new("taskkill");
        command
            .arg("/PID")
            .arg(pid.to_string())
            .arg("/T")
            .arg("/F");
        hide_windows_console(&mut command);
        let status = command
            .status()
            .map_err(|error| format!("Unable to terminate OCR worker {pid}: {error}"))?;
        return Ok(status.success());
    }

    #[allow(unreachable_code)]
    Ok(false)
}

pub(super) fn stop_worker(
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
) -> Result<(), String> {
    let terminate_result = terminate_worker_process(worker_pid);
    if let Ok(mut guard) = worker.lock() {
        guard.take();
    }
    terminate_result.map(|_| ())
}

pub(super) fn spawn_worker(
    app: &AppHandle,
    paths: &RuntimePaths,
    worker_pid: Arc<AtomicU32>,
) -> Result<OcrWorker, String> {
    ensure_private_python_isolation(paths)?;
    let script = worker_script_path(app)?;
    fs::create_dir_all(&paths.logs)
        .map_err(|error| format!("Unable to create OCR log directory: {error}"))?;
    fs::create_dir_all(&paths.cache)
        .map_err(|error| format!("Unable to create OCR cache directory: {error}"))?;
    fs::create_dir_all(&paths.temp)
        .map_err(|error| format!("Unable to create OCR temporary directory: {error}"))?;
    let log_path = paths.logs.join("worker.log");
    let mut log_file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log_path)
        .map_err(|error| format!("Unable to open OCR worker log: {error}"))?;
    writeln!(
        log_file,
        "\n===== VisualTeX OCR worker start: pid pending, unix_ms={} =====",
        SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|duration| duration.as_millis())
            .unwrap_or_default()
    )
    .map_err(|error| format!("Unable to initialize OCR worker log: {error}"))?;
    let log_file_error = log_file
        .try_clone()
        .map_err(|error| format!("Unable to clone OCR log handle: {error}"))?;

    let mut command = python_command(&paths.python);
    configure_ocr_runtime_cache_environment(&mut command, paths);
    command
        .arg(&script)
        .env("PYTHONUNBUFFERED", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .env("VISUALTEX_PARENT_PID", std::process::id().to_string())
        .env("VISUALTEX_OFFLINE_OCR", "1")
        .env("HF_HUB_OFFLINE", "1")
        .env("TRANSFORMERS_OFFLINE", "1")
        .env("MODELSCOPE_OFFLINE", "1")
        .env("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::from(log_file_error));
    hide_windows_console(&mut command);
    let mut child = command
        .spawn()
        .map_err(|error| format!("Unable to start OCR worker: {error}"))?;
    worker_pid.store(child.id(), Ordering::SeqCst);

    let stdin = child
        .stdin
        .take()
        .ok_or_else(|| "OCR worker stdin is unavailable".to_string())?;
    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "OCR worker stdout is unavailable".to_string())?;
    let mut worker = OcrWorker {
        child,
        stdin: BufWriter::new(stdin),
        stdout: BufReader::new(stdout),
        pid_state: worker_pid,
        log_path: log_path.clone(),
        loaded_model: None,
    };

    let ready = read_worker_json(
        &mut worker.stdout,
        "OCR worker closed before sending its ready signal",
        "OCR worker ready signal",
    )
    .map_err(|error| format!("{error}; log={}", log_path.display()))?;
    if ready.get("event").and_then(Value::as_str) != Some("ready") {
        return Err(format!("Unexpected OCR worker ready response: {ready}"));
    }
    Ok(worker)
}

pub(super) fn acquire_background_runtime_mutation(
    storage_change_running: &Arc<AtomicBool>,
    runtime_mutations: &Arc<AtomicU32>,
) -> Option<RuntimeMutationLease> {
    let deadline = Instant::now() + Duration::from_secs(20);
    loop {
        if storage_change_running.load(Ordering::SeqCst) {
            return None;
        }
        if runtime_mutations
            .compare_exchange(0, 1, Ordering::SeqCst, Ordering::SeqCst)
            .is_ok()
        {
            if storage_change_running.load(Ordering::SeqCst) {
                runtime_mutations.store(0, Ordering::SeqCst);
                return None;
            }
            return Some(RuntimeMutationLease {
                counter: runtime_mutations.clone(),
            });
        }
        if Instant::now() >= deadline {
            return None;
        }
        thread::sleep(Duration::from_millis(100));
    }
}

pub(super) fn schedule_worker_rewarm(
    app: AppHandle,
    worker_state: Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: Arc<AtomicU32>,
    runtime_status: Arc<Mutex<Option<OcrRuntimeStatus>>>,
    desired_warmup_model: Arc<Mutex<Option<String>>>,
    storage_change_running: Arc<AtomicBool>,
    runtime_mutations: Arc<AtomicU32>,
) {
    let _ = tauri::async_runtime::spawn(async move {
        let result = tauri::async_runtime::spawn_blocking(move || {
            let Some(_mutation) = acquire_background_runtime_mutation(
                &storage_change_running,
                &runtime_mutations,
            ) else {
                return Ok(());
            };
            let model = {
                let mut desired = desired_warmup_model
                    .lock()
                    .map_err(|_| "OCR warmup selection is unavailable".to_string())?;
                desired
                    .get_or_insert_with(|| DEFAULT_OCR_MODEL.to_string())
                    .clone()
            };

            // The recognition thread releases this lock as soon as the killed
            // worker pipe closes. A warmup task that was already running may
            // have recovered the worker itself, so preserve that live warmed
            // process instead of killing and loading it a second time.
            {
                let mut guard = worker_state
                    .lock()
                    .map_err(|_| "OCR worker lock is poisoned".to_string())?;
                let already_recovered = if let Some(worker) = guard.as_mut() {
                    worker.loaded_model.as_deref() == Some(model.as_str())
                        && worker
                            .child
                            .try_wait()
                            .map_err(|error| {
                                format!("Unable to inspect the recovered OCR worker: {error}")
                            })?
                            .is_none()
                } else {
                    false
                };
                if already_recovered {
                    return Ok(());
                }
                guard.take();
            }

            let paths = runtime_paths(&app)?;
            cleanup_worker_temp(&paths)?;
            warmup_worker(
                &app,
                &worker_state,
                &worker_pid,
                &runtime_status,
                &desired_warmup_model,
                &model,
            )
        })
        .await
        .map_err(|error| format!("OCR recovery task failed: {error}"))
        .and_then(|result| result);

        if let Err(error) = result {
            eprintln!("Unable to restore OCR prewarm after cancellation: {error}");
        }
    });
}

pub(super) fn warmup_model_is_current(
    desired_warmup_model: &Arc<Mutex<Option<String>>>,
    model: &str,
) -> Result<bool, String> {
    desired_warmup_model
        .lock()
        .map(|desired| desired.as_deref() == Some(model))
        .map_err(|_| "OCR warmup selection is unavailable".to_string())
}

pub(super) fn warmup_worker(
    app: &AppHandle,
    worker_state: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
    runtime_status: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    desired_warmup_model: &Arc<Mutex<Option<String>>>,
    model: &str,
) -> Result<(), String> {
    if !ALLOWED_MODELS.contains(&model) {
        return Err(format!("Unsupported PP-FormulaNet model: {model}"));
    }
    if !warmup_model_is_current(desired_warmup_model, model)? {
        return Ok(());
    }
    let paths = runtime_paths(app)?;
    let status = match read_cached_runtime_status(runtime_status)? {
        Some(status) => status,
        None => {
            let status = get_runtime_status_fast(app)?;
            write_cached_runtime_status(runtime_status, Some(status.clone()))?;
            status
        }
    };
    if !status.installed {
        return Err(format!("OCR runtime is not installed: {}", status.message));
    }
    if !status.installed_models.iter().any(|installed| installed == model) {
        return Err(format!(
            "OCR model {model} is not installed. Startup warmup will not download it automatically."
        ));
    }

    {
        let mut guard = worker_state
            .lock()
            .map_err(|_| "OCR worker lock is poisoned".to_string())?;
        if !warmup_model_is_current(desired_warmup_model, model)? {
            return Ok(());
        }
        if guard
            .as_ref()
            .and_then(|worker| worker.loaded_model.as_deref())
            == Some(model)
        {
            return Ok(());
        }
        if guard.is_none() {
            *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
        }
        let payload = json!({
            "id": format!("warmup-{}", std::process::id()),
            "action": "warmup",
            "model": model,
            "device": "cpu"
        });
        let first_result = guard
            .as_mut()
            .ok_or_else(|| "OCR worker failed to start".to_string())?
            .send(app, &payload);
        let response = match first_result {
            Ok(response) => response,
            Err(first_error) => {
                guard.take();
                *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
                guard
                    .as_mut()
                    .ok_or_else(|| "OCR worker failed to restart".to_string())?
                    .send(app, &payload)
                    .map_err(|second_error| {
                        format!(
                            "OCR warmup worker failed twice. First: {first_error}. Second: {second_error}"
                        )
                    })?
            }
        };
        if response.get("ok").and_then(Value::as_bool) != Some(true) {
            let error = response
                .get("error")
                .and_then(Value::as_str)
                .unwrap_or("OCR model warmup failed");
            return Err(error.to_string());
        }
        if let Some(worker) = guard.as_mut() {
            worker.loaded_model = Some(model.to_string());
        }
    }

    let refreshed = get_runtime_status_fast(app)?;
    write_cached_runtime_status(runtime_status, Some(refreshed))?;
    Ok(())
}

pub(super) fn run_recognition(
    app: &AppHandle,
    worker_state: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
    cancel_generation: &Arc<AtomicU64>,
    runtime_status: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    request: OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    let request_generation = cancel_generation.load(Ordering::SeqCst);

    if request.bytes.is_empty() {
        return Err("The selected image is empty".to_string());
    }
    if request.bytes.len() > MAX_IMAGE_BYTES {
        return Err("The image is larger than the 20 MB limit".to_string());
    }
    if !ALLOWED_MODELS.contains(&request.model.as_str()) {
        return Err(format!(
            "Unsupported PP-FormulaNet model: {}",
            request.model
        ));
    }

    let paths = runtime_paths(app)?;
    let status = match read_cached_runtime_status(runtime_status)? {
        Some(status) => status,
        None => {
            let status = get_runtime_status_fast(app)?;
            write_cached_runtime_status(runtime_status, Some(status.clone()))?;
            status
        }
    };
    if !status.installed {
        return Err(format!("OCR runtime is not installed: {}", status.message));
    }
    if !status
        .installed_models
        .iter()
        .any(|installed| installed == &request.model)
    {
        return Err(format!(
            "OCR model {} is not installed. VisualTeX will not download it during recognition.",
            request.model
        ));
    }
    if !paths.python.exists() {
        return Err("OCR runtime is not installed: Python executable is missing".to_string());
    }
    fs::create_dir_all(&paths.input)
        .map_err(|error| format!("Unable to create OCR input directory: {error}"))?;
    fs::create_dir_all(&paths.processed)
        .map_err(|error| format!("Unable to create OCR processed directory: {error}"))?;

    let extension = request
        .extension
        .trim_start_matches('.')
        .to_ascii_lowercase();
    let allowed_extensions = ["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"];
    if !allowed_extensions.contains(&extension.as_str()) {
        return Err(format!("Unsupported image type: .{extension}"));
    }

    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|error| error.to_string())?
        .as_nanos();
    let request_id = format!("{}-{nonce}", std::process::id());
    let input_path = paths.input.join(format!("{request_id}.{extension}"));
    let processed_path = paths.processed.join(format!("{request_id}.png"));
    fs::write(&input_path, &request.bytes)
        .map_err(|error| format!("Unable to save OCR input image: {error}"))?;

    let payload = json!({
        "id": request_id,
        "action": "recognize",
        "image_path": input_path,
        "processed_path": processed_path,
        "model": request.model,
        "device": "cpu"
    });

    if cancel_generation.load(Ordering::SeqCst) != request_generation {
        let _ = fs::remove_file(&input_path);
        return Err(OCR_CANCELLED.to_string());
    }

    let response_result = (|| -> Result<Value, String> {
        let mut next_wait_notice = Instant::now();
        let mut guard = loop {
            match worker_state.try_lock() {
                Ok(guard) => break guard,
                Err(TryLockError::WouldBlock) => {
                    if cancel_generation.load(Ordering::SeqCst) != request_generation {
                        return Err(OCR_CANCELLED.to_string());
                    }
                    let now = Instant::now();
                    if now >= next_wait_notice {
                        let message = if request.model == "PP-FormulaNet_plus-L" {
                            "正在等待高精度 L 模型后台准备；首次使用可能需要下载约 731.5 MB 并完成模型初始化，准备完成后会自动继续识别".to_string()
                        } else {
                            format!(
                                "正在等待 {} 模型后台准备，完成后会自动继续识别",
                                request.model
                            )
                        };
                        emit_recognition_progress(
                            app,
                            &request_id,
                            "model-wait",
                            message,
                            &request.model,
                        );
                        next_wait_notice = now + OCR_WORKER_WAIT_NOTICE_INTERVAL;
                    }
                    thread::sleep(OCR_WORKER_WAIT_POLL_INTERVAL);
                }
                Err(TryLockError::Poisoned(_)) => {
                    return Err("OCR worker lock is poisoned".to_string());
                }
            }
        };

        if guard.is_none() {
            *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
        }

        let first_result = guard
            .as_mut()
            .ok_or_else(|| "OCR worker failed to start".to_string())?
            .send(app, &payload);
        match first_result {
            Ok(response) => {
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }
                if response.get("ok").and_then(Value::as_bool) == Some(true) {
                    if let Some(worker) = guard.as_mut() {
                        worker.loaded_model = Some(request.model.clone());
                    }
                }
                Ok(response)
            }
            Err(first_error) => {
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }

                guard.take();
                *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
                let response = guard
                    .as_mut()
                    .ok_or_else(|| "OCR worker failed to restart".to_string())?
                    .send(app, &payload)
                    .map_err(|second_error| {
                        format!(
                            "OCR worker failed twice. First: {first_error}. Second: {second_error}"
                        )
                    })?;
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }
                if response.get("ok").and_then(Value::as_bool) == Some(true) {
                    if let Some(worker) = guard.as_mut() {
                        worker.loaded_model = Some(request.model.clone());
                    }
                }
                Ok(response)
            }
        }
    })();

    let _ = fs::remove_file(&input_path);
    let _ = fs::remove_file(&processed_path);
    let response_value = response_result?;

    let response: WorkerRecognitionResponse = serde_json::from_value(response_value)
        .map_err(|error| format!("Unable to decode OCR result: {error}"))?;
    if !response.ok {
        let mut message = response
            .error
            .unwrap_or_else(|| "PP-FormulaNet recognition failed".to_string());
        if let Some(details) = response.details {
            if !details.trim().is_empty() {
                let detail_tail = tail_text(&details, 3000);
                message.push('\n');
                message.push_str(detail_tail.trim());
            }
        }
        return Err(message);
    }

    let formulas = response.formulas.unwrap_or_default();
    if formulas.is_empty() {
        return Err("PP-FormulaNet returned no formulas".to_string());
    }

    Ok(OcrRecognitionResult {
        provider: provider::LOCAL_PROVIDER.to_string(),
        model: response.model.unwrap_or_else(|| request.model.clone()),
        elapsed_ms: response.elapsed_ms.unwrap_or_default(),
        processed_width: response.processed_width.unwrap_or_default(),
        processed_height: response.processed_height.unwrap_or_default(),
        background_inverted: response.background_inverted.unwrap_or(false),
        background_luminance: response.background_luminance.unwrap_or(255.0),
        formulas,
    })
}
