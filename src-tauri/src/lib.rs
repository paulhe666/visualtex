use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::env;
use std::fs::{self, OpenOptions};
use std::io::{BufRead, BufReader, BufWriter, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};
use std::sync::atomic::{AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Emitter, Manager, State};

mod office;

const PADDLE_VERSION: &str = "3.3.1";
const PADDLEOCR_VERSION: &str = "3.7.0";
const MAX_IMAGE_BYTES: usize = 20 * 1024 * 1024;
const OCR_CANCELLED: &str = "OCR_CANCELLED";
const ALLOWED_MODELS: &[&str] = &[
    "PP-FormulaNet_plus-S",
    "PP-FormulaNet_plus-M",
    "PP-FormulaNet_plus-L",
    "PP-FormulaNet-S",
    "PP-FormulaNet-L",
];

#[derive(Clone)]
struct RuntimePaths {
    root: PathBuf,
    venv: PathBuf,
    python: PathBuf,
    input: PathBuf,
    processed: PathBuf,
    logs: PathBuf,
    cache: PathBuf,
    temp: PathBuf,
}

struct OcrState {
    worker: Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: Arc<AtomicU32>,
    cancel_generation: Arc<AtomicU64>,
    runtime_status: Arc<Mutex<Option<OcrRuntimeStatus>>>,
}

impl Default for OcrState {
    fn default() -> Self {
        Self {
            worker: Arc::new(Mutex::new(None)),
            worker_pid: Arc::new(AtomicU32::new(0)),
            cancel_generation: Arc::new(AtomicU64::new(0)),
            runtime_status: Arc::new(Mutex::new(None)),
        }
    }
}

impl Drop for OcrState {
    fn drop(&mut self) {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let _ = terminate_worker_process(&self.worker_pid);
    }
}

struct OcrWorker {
    child: Child,
    stdin: BufWriter<ChildStdin>,
    stdout: BufReader<ChildStdout>,
    pid_state: Arc<AtomicU32>,
    log_path: PathBuf,
    loaded_model: Option<String>,
}

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

fn read_worker_json<R: BufRead>(
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
    fn worker_failure(&mut self, message: impl AsRef<str>) -> String {
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

    fn send(&mut self, app: &AppHandle, payload: &Value) -> Result<Value, String> {
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
                continue;
            }
            return Ok(response);
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct OcrInstallProgress {
    stage: String,
    percent: u8,
    message: String,
    detail: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct OcrRuntimeStatus {
    installed: bool,
    python_path: Option<String>,
    python_version: Option<String>,
    paddle_version: Option<String>,
    paddleocr_version: Option<String>,
    runtime_path: String,
    message: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OcrImageRequest {
    bytes: Vec<u8>,
    extension: String,
    model: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct OcrFormulaResult {
    latex: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct OcrRecognitionResult {
    model: String,
    elapsed_ms: u64,
    processed_width: u32,
    processed_height: u32,
    background_inverted: bool,
    background_luminance: f64,
    formulas: Vec<OcrFormulaResult>,
}

#[derive(Debug, Deserialize)]
struct WorkerRecognitionResponse {
    ok: bool,
    model: Option<String>,
    elapsed_ms: Option<u64>,
    processed_width: Option<u32>,
    processed_height: Option<u32>,
    background_inverted: Option<bool>,
    background_luminance: Option<f64>,
    formulas: Option<Vec<OcrFormulaResult>>,
    error: Option<String>,
    details: Option<String>,
}

#[derive(Debug, Deserialize)]
struct RuntimeProbe {
    python_version: String,
    paddle_version: String,
    paddleocr_version: String,
}

#[derive(Debug, Deserialize)]
struct PythonProbe {
    version: String,
    major: u8,
    minor: u8,
    machine: String,
    executable: String,
}

fn runtime_paths(app: &AppHandle) -> Result<RuntimePaths, String> {
    let root = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve application data directory: {error}"))?
        .join("ocr-runtime");
    let venv = root.join("venv");
    let python = if cfg!(windows) {
        venv.join("Scripts").join("python.exe")
    } else {
        venv.join("bin").join("python")
    };
    Ok(RuntimePaths {
        root: root.clone(),
        venv,
        python,
        input: root.join("input"),
        processed: root.join("processed"),
        logs: root.join("logs"),
        cache: root.join("cache"),
        temp: root.join("tmp"),
    })
}

fn worker_script_path(app: &AppHandle) -> Result<PathBuf, String> {
    if let Ok(path) = app.path().resolve("ocr/worker.py", BaseDirectory::Resource) {
        if path.exists() {
            return Ok(path);
        }
    }

    let development_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("ocr")
        .join("worker.py");
    if development_path.exists() {
        return Ok(development_path);
    }

    Err("Unable to locate bundled OCR worker.py".to_string())
}

fn emit_progress(
    app: &AppHandle,
    stage: &str,
    percent: u8,
    message: impl Into<String>,
    detail: Option<String>,
) {
    let _ = app.emit(
        "ocr-install-progress",
        OcrInstallProgress {
            stage: stage.to_string(),
            percent,
            message: message.into(),
            detail,
        },
    );
}

fn tail_text(value: &str, max_chars: usize) -> String {
    let total = value.chars().count();
    if total <= max_chars {
        return value.to_string();
    }
    value.chars().skip(total - max_chars).collect()
}

fn command_output(command: &mut Command, label: &str) -> Result<String, String> {
    let output = command
        .output()
        .map_err(|error| format!("Unable to start {label}: {error}"))?;
    if output.status.success() {
        return Ok(String::from_utf8_lossy(&output.stdout).trim().to_string());
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    let combined = format!("{stdout}\n{stderr}");
    let tail = tail_text(&combined, 6000);
    Err(format!(
        "{label} failed with {}:\n{}",
        output.status,
        tail.trim()
    ))
}

fn probe_python(
    program: &Path,
    prefix_args: &[String],
    label: &str,
) -> Result<PythonProbe, String> {
    let script = r#"import json, platform, sys; print(json.dumps({'version': platform.python_version(), 'major': sys.version_info.major, 'minor': sys.version_info.minor, 'machine': platform.machine(), 'executable': sys.executable}))"#;
    let mut command = Command::new(program);
    command
        .args(prefix_args)
        .arg("-c")
        .arg(script)
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8");
    let output = command_output(&mut command, &format!("Python version check ({label})"))?;
    serde_json::from_str(&output)
        .map_err(|error| format!("Python returned an invalid version response: {error}"))
}

fn find_system_python() -> Result<(PathBuf, PythonProbe), String> {
    let mut candidates: Vec<(String, PathBuf, Vec<String>)> = Vec::new();

    if let Ok(explicit) = env::var("VISUALTEX_PYTHON") {
        let explicit = explicit.trim().trim_matches('"');
        if !explicit.is_empty() {
            candidates.push((
                "VISUALTEX_PYTHON".to_string(),
                PathBuf::from(explicit),
                Vec::new(),
            ));
        }
    }

    if cfg!(windows) {
        // The Windows Python launcher usually exposes installed runtimes via
        // selectors rather than python3.13.exe-style command names. Probe each
        // supported version explicitly so a default Python 3.14 installation
        // cannot hide an installed 3.9–3.13 runtime.
        for minor in (9..=13).rev() {
            for launcher in ["py.exe", "py"] {
                for selector in [format!("-V:3.{minor}"), format!("-3.{minor}")] {
                    candidates.push((
                        format!("{launcher} {selector}"),
                        PathBuf::from(launcher),
                        vec![selector],
                    ));
                }
            }
        }
    }

    for name in [
        "python3.13",
        "python3.12",
        "python3.11",
        "python3.10",
        "python3.9",
        "python3",
    ] {
        candidates.push((name.to_string(), PathBuf::from(name), Vec::new()));
    }

    if cfg!(windows) {
        for name in ["python.exe", "python", "py.exe", "py"] {
            candidates.push((name.to_string(), PathBuf::from(name), Vec::new()));
        }
    }
    if cfg!(target_os = "macos") {
        for path in [
            "/opt/homebrew/bin/python3",
            "/usr/local/bin/python3",
            "/usr/bin/python3",
        ] {
            candidates.push((path.to_string(), PathBuf::from(path), Vec::new()));
        }
    }

    let mut failures = Vec::new();
    for (label, program, prefix_args) in candidates {
        match probe_python(&program, &prefix_args, &label) {
            Ok(probe) => {
                let version_ok = probe.major == 3 && (9..=13).contains(&probe.minor);
                let machine = probe.machine.to_ascii_lowercase();
                let architecture_ok = if cfg!(all(target_os = "macos", target_arch = "aarch64")) {
                    machine == "arm64" || machine == "aarch64"
                } else if cfg!(any(target_os = "windows", target_os = "linux")) {
                    matches!(machine.as_str(), "x86_64" | "amd64" | "x64")
                } else {
                    true
                };
                if version_ok && architecture_ok {
                    let executable = PathBuf::from(&probe.executable);
                    if executable.as_os_str().is_empty() {
                        failures.push(format!("{label}: Python did not report sys.executable"));
                        continue;
                    }
                    return Ok((executable, probe));
                }
                failures.push(format!(
                    "{label}: Python {}, architecture {}",
                    probe.version, probe.machine
                ));
            }
            Err(error) => failures.push(format!("{label}: {error}")),
        }
    }

    Err(format!(
        "未检测到可用于 OCR 的 64 位 Python 3.9–3.13。请安装 x64 Python 3.13，并启用 Python Launcher，然后完全重启 VisualTeX。Python 3.14 当前不兼容 PaddleOCR。\n\nNo compatible 64-bit Python 3.9–3.13 interpreter was found. Install x64 Python 3.13 with the Python Launcher enabled, then restart VisualTeX. Python 3.14 is not currently supported by PaddleOCR.\n\nChecked:\n{}",
        failures.join("\n")
    ))
}

fn probe_runtime(paths: &RuntimePaths) -> Result<RuntimeProbe, String> {
    if !paths.python.exists() {
        return Err("OCR virtual environment is not installed".to_string());
    }
    let script = r#"import json, platform; import paddle; import paddleocr; import tokenizers, imagesize, ftfy, wand; from importlib.metadata import version; from paddleocr import FormulaRecognition; print(json.dumps({'python_version': platform.python_version(), 'paddle_version': paddle.__version__, 'paddleocr_version': version('paddleocr')}))"#;
    let output = command_output(
        Command::new(&paths.python).arg("-c").arg(script),
        "OCR runtime verification",
    )?;
    serde_json::from_str(&output)
        .map_err(|error| format!("OCR runtime returned invalid version information: {error}"))
}

fn get_runtime_status_inner(app: &AppHandle) -> Result<OcrRuntimeStatus, String> {
    let paths = runtime_paths(app)?;
    match probe_runtime(&paths) {
        Ok(probe) => Ok(OcrRuntimeStatus {
            installed: true,
            python_path: Some(paths.python.display().to_string()),
            python_version: Some(probe.python_version),
            paddle_version: Some(probe.paddle_version),
            paddleocr_version: Some(probe.paddleocr_version),
            runtime_path: paths.root.display().to_string(),
            message: "PaddleOCR formula runtime is ready".to_string(),
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
            message: error,
        }),
    }
}

fn read_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
) -> Result<Option<OcrRuntimeStatus>, String> {
    cache
        .lock()
        .map(|status| status.clone())
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())
}

fn write_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    status: Option<OcrRuntimeStatus>,
) -> Result<(), String> {
    let mut guard = cache
        .lock()
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())?;
    *guard = status;
    Ok(())
}

fn cleanup_worker_temp(paths: &RuntimePaths) -> Result<(), String> {
    if paths.temp.exists() {
        fs::remove_dir_all(&paths.temp)
            .map_err(|error| format!("Unable to clean OCR temporary files: {error}"))?;
    }
    fs::create_dir_all(&paths.temp)
        .map_err(|error| format!("Unable to create OCR temporary directory: {error}"))
}

fn terminate_worker_process(worker_pid: &AtomicU32) -> Result<bool, String> {
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
        let status = Command::new("taskkill")
            .arg("/PID")
            .arg(pid.to_string())
            .arg("/T")
            .arg("/F")
            .status()
            .map_err(|error| format!("Unable to terminate OCR worker {pid}: {error}"))?;
        return Ok(status.success());
    }

    #[allow(unreachable_code)]
    Ok(false)
}

fn stop_worker(
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
) -> Result<(), String> {
    let terminate_result = terminate_worker_process(worker_pid);
    if let Ok(mut guard) = worker.lock() {
        guard.take();
    }
    terminate_result.map(|_| ())
}

fn install_runtime_inner(
    app: &AppHandle,
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
) -> Result<OcrRuntimeStatus, String> {
    stop_worker(worker, worker_pid)?;
    let paths = runtime_paths(app)?;
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    fs::create_dir_all(&paths.input)
        .map_err(|error| format!("Unable to create OCR input directory: {error}"))?;
    fs::create_dir_all(&paths.processed)
        .map_err(|error| format!("Unable to create OCR processed directory: {error}"))?;
    fs::create_dir_all(&paths.logs)
        .map_err(|error| format!("Unable to create OCR log directory: {error}"))?;
    fs::create_dir_all(&paths.cache)
        .map_err(|error| format!("Unable to create OCR cache directory: {error}"))?;
    cleanup_worker_temp(&paths)?;

    emit_progress(app, "python", 5, "正在查找兼容的 Python", None);
    let (system_python, probe) = find_system_python()?;
    emit_progress(
        app,
        "python",
        12,
        format!("使用 Python {}", probe.version),
        Some(system_python.display().to_string()),
    );

    emit_progress(app, "venv", 18, "正在创建独立 OCR 环境", None);
    command_output(
        Command::new(&system_python)
            .arg("-m")
            .arg("venv")
            .arg("--clear")
            .arg(&paths.venv),
        "Python virtual environment creation",
    )?;

    emit_progress(app, "pip", 28, "正在更新安装工具", None);
    command_output(
        Command::new(&paths.python)
            .arg("-m")
            .arg("pip")
            .arg("install")
            .arg("--disable-pip-version-check")
            .arg("--no-input")
            .arg("--upgrade")
            .arg("pip")
            .arg("setuptools")
            .arg("wheel"),
        "pip bootstrap",
    )?;

    emit_progress(
        app,
        "paddle",
        42,
        format!("正在安装 PaddlePaddle {PADDLE_VERSION}"),
        Some("首次安装需要下载约 100 MB 的框架文件".to_string()),
    );
    command_output(
        Command::new(&paths.python)
            .arg("-m")
            .arg("pip")
            .arg("install")
            .arg("--disable-pip-version-check")
            .arg("--no-input")
            .arg(format!("paddlepaddle=={PADDLE_VERSION}")),
        "PaddlePaddle installation",
    )?;

    emit_progress(
        app,
        "paddleocr",
        68,
        format!("正在安装 PaddleOCR {PADDLEOCR_VERSION}"),
        None,
    );
    command_output(
        Command::new(&paths.python)
            .arg("-m")
            .arg("pip")
            .arg("install")
            .arg("--disable-pip-version-check")
            .arg("--no-input")
            .arg(format!("paddleocr=={PADDLEOCR_VERSION}")),
        "PaddleOCR installation",
    )?;

    emit_progress(
        app,
        "formula-deps",
        84,
        "正在安装 PP-FormulaNet 公式解码依赖",
        Some("安装 tokenizers、ftfy、imagesize 和 Wand".to_string()),
    );
    command_output(
        Command::new(&paths.python)
            .arg("-m")
            .arg("pip")
            .arg("install")
            .arg("--disable-pip-version-check")
            .arg("--no-input")
            .arg("tokenizers==0.19.1")
            .arg("imagesize")
            .arg("ftfy")
            .arg("Wand"),
        "PP-FormulaNet dependency installation",
    )?;

    emit_progress(app, "verify", 92, "正在验证 PP-FormulaNet 接口", None);
    let status = get_runtime_status_inner(app)?;
    if !status.installed {
        return Err(status.message);
    }
    emit_progress(
        app,
        "complete",
        100,
        "OCR 运行环境安装完成",
        Some("模型权重会在第一次识别时自动下载".to_string()),
    );
    Ok(status)
}

fn spawn_worker(
    app: &AppHandle,
    paths: &RuntimePaths,
    worker_pid: Arc<AtomicU32>,
) -> Result<OcrWorker, String> {
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

    let mut child = Command::new(&paths.python)
        .arg(&script)
        .env("PYTHONUNBUFFERED", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .env("VISUALTEX_PARENT_PID", std::process::id().to_string())
        .env("PADDLE_PDX_MODEL_SOURCE", "BOS")
        .env("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
        .env("PADDLE_PDX_CACHE_HOME", paths.cache.join("paddlex"))
        .env("PADDLE_HOME", paths.cache.join("paddle"))
        .env("XDG_CACHE_HOME", &paths.cache)
        .env("TMPDIR", &paths.temp)
        .env("TMP", &paths.temp)
        .env("TEMP", &paths.temp)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::from(log_file_error))
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

fn run_recognition(
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
    if let Some(status) = read_cached_runtime_status(runtime_status)? {
        if !status.installed {
            return Err(format!("OCR runtime is not installed: {}", status.message));
        }
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
        let mut guard = worker_state
            .lock()
            .map_err(|_| "OCR worker lock is poisoned".to_string())?;

        let should_restart_for_model = guard
            .as_ref()
            .and_then(|worker| worker.loaded_model.as_deref())
            .is_some_and(|loaded_model| loaded_model != request.model);
        if should_restart_for_model {
            guard.take();
        }
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
        model: response.model.unwrap_or_else(|| request.model.clone()),
        elapsed_ms: response.elapsed_ms.unwrap_or_default(),
        processed_width: response.processed_width.unwrap_or_default(),
        processed_height: response.processed_height.unwrap_or_default(),
        background_inverted: response.background_inverted.unwrap_or(false),
        background_luminance: response.background_luminance.unwrap_or(255.0),
        formulas,
    })
}

#[tauri::command]
async fn get_ocr_runtime_status(
    app: AppHandle,
    state: State<'_, OcrState>,
    force_refresh: Option<bool>,
) -> Result<OcrRuntimeStatus, String> {
    let runtime_status = state.runtime_status.clone();
    if !force_refresh.unwrap_or(false) {
        if let Some(status) = read_cached_runtime_status(&runtime_status)? {
            return Ok(status);
        }
    }

    tauri::async_runtime::spawn_blocking(move || {
        let status = get_runtime_status_inner(&app)?;
        write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
        Ok(status)
    })
    .await
    .map_err(|error| format!("OCR runtime status task failed: {error}"))?
}

#[tauri::command]
async fn install_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    let worker = state.worker.clone();
    let worker_pid = state.worker_pid.clone();
    let runtime_status = state.runtime_status.clone();
    tauri::async_runtime::spawn_blocking(move || {
        let status = install_runtime_inner(&app, &worker, &worker_pid)?;
        write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
        Ok(status)
    })
    .await
    .map_err(|error| format!("OCR installer task failed: {error}"))?
}

#[tauri::command]
async fn recognize_formula_image(
    app: AppHandle,
    state: State<'_, OcrState>,
    request: OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    let worker = state.worker.clone();
    let worker_pid = state.worker_pid.clone();
    let cancel_generation = state.cancel_generation.clone();
    let runtime_status = state.runtime_status.clone();
    tauri::async_runtime::spawn_blocking(move || {
        run_recognition(
            &app,
            &worker,
            &worker_pid,
            &cancel_generation,
            &runtime_status,
            request,
        )
    })
    .await
    .map_err(|error| format!("OCR recognition task failed: {error}"))?
}

#[tauri::command]
fn cancel_ocr_recognition(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.cancel_generation.fetch_add(1, Ordering::SeqCst);
    terminate_worker_process(&state.worker_pid)?;
    cleanup_worker_temp(&runtime_paths(&app)?)
}

#[tauri::command]
async fn restart_ocr_worker(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.cancel_generation.fetch_add(1, Ordering::SeqCst);
    let worker = state.worker.clone();
    let worker_pid = state.worker_pid.clone();
    tauri::async_runtime::spawn_blocking(move || {
        stop_worker(&worker, &worker_pid)?;
        cleanup_worker_temp(&runtime_paths(&app)?)
    })
    .await
    .map_err(|error| format!("OCR restart task failed: {error}"))?
}

#[tauri::command]
async fn reset_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    state.cancel_generation.fetch_add(1, Ordering::SeqCst);
    let worker = state.worker.clone();
    let worker_pid = state.worker_pid.clone();
    let runtime_status = state.runtime_status.clone();
    tauri::async_runtime::spawn_blocking(move || {
        write_cached_runtime_status(&runtime_status, None)?;
        stop_worker(&worker, &worker_pid)?;
        let paths = runtime_paths(&app)?;
        if paths.root.exists() {
            fs::remove_dir_all(&paths.root)
                .map_err(|error| format!("Unable to remove OCR runtime: {error}"))?;
        }
        let status = get_runtime_status_inner(&app)?;
        write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
        Ok(status)
    })
    .await
    .map_err(|error| format!("OCR reset task failed: {error}"))?
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .manage(OcrState::default())
        .setup(|app| {
            let office_state = office::initialize(app.handle()).map_err(std::io::Error::other)?;
            app.manage(office_state.clone());
            office::start(office_state);
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            get_ocr_runtime_status,
            install_ocr_runtime,
            recognize_formula_image,
            cancel_ocr_recognition,
            restart_ocr_worker,
            reset_ocr_runtime,
            office::lifecycle::get_office_companion_status,
            office::lifecycle::start_office_companion,
            office::lifecycle::stop_office_companion,
            office::lifecycle::get_office_integration_status,
            office::lifecycle::install_office_integration,
            office::lifecycle::repair_office_integration,
            office::lifecycle::uninstall_office_integration,
            office::lifecycle::regenerate_office_certificate,
            office::lifecycle::open_word,
            office::lifecycle::open_powerpoint
        ])
        .run(tauri::generate_context!())
        .expect("error while running VisualTeX");
}

#[cfg(test)]
mod protocol_tests {
    use super::*;
    use std::io::{BufReader, Cursor};

    #[test]
    fn worker_protocol_skips_non_utf8_diagnostics_before_json() {
        let bytes = b"\xd5\xfd\xca\xbd\xc8\xd5\xd6\xbe\n{\"event\":\"ready\",\"ok\":true}\n";
        let mut reader = BufReader::new(Cursor::new(bytes));

        let value = read_worker_json(&mut reader, "closed", "test response")
            .expect("reader should skip non-protocol diagnostic bytes");

        assert_eq!(value.get("event").and_then(Value::as_str), Some("ready"));
    }

    #[test]
    fn worker_protocol_decodes_ascii_escaped_unicode() {
        let bytes = b"{\"message\":\"\\u6b63\\u5728\\u52a0\\u8f7d\"}\n";
        let mut reader = BufReader::new(Cursor::new(bytes));

        let value = read_worker_json(&mut reader, "closed", "test response")
            .expect("escaped Unicode JSON should parse");

        assert_eq!(
            value.get("message").and_then(Value::as_str),
            Some("正在加载")
        );
    }

    #[test]
    fn runtime_status_cache_round_trips_and_clears() {
        let cache = Arc::new(Mutex::new(None));
        let expected = OcrRuntimeStatus {
            installed: true,
            python_path: Some("/tmp/visualtex-python".to_string()),
            python_version: Some("3.13.0".to_string()),
            paddle_version: Some(PADDLE_VERSION.to_string()),
            paddleocr_version: Some(PADDLEOCR_VERSION.to_string()),
            runtime_path: "/tmp/visualtex-ocr".to_string(),
            message: "ready".to_string(),
        };

        write_cached_runtime_status(&cache, Some(expected.clone()))
            .expect("cache write should succeed");
        let cached = read_cached_runtime_status(&cache)
            .expect("cache read should succeed")
            .expect("cached status should exist");
        assert_eq!(cached.python_version, expected.python_version);
        assert!(cached.installed);

        write_cached_runtime_status(&cache, None).expect("cache clear should succeed");
        assert!(read_cached_runtime_status(&cache)
            .expect("cache read after clear should succeed")
            .is_none());
    }
}

#[cfg(all(test, unix))]
mod tests {
    use super::*;

    #[test]
    fn active_worker_can_be_terminated_without_taking_the_worker_lock() {
        let mut child = Command::new("sleep")
            .arg("30")
            .spawn()
            .expect("failed to start test process");
        let pid = child.id();
        let worker_pid = AtomicU32::new(pid);

        assert!(terminate_worker_process(&worker_pid).expect("termination failed"));
        let status = child.wait().expect("failed to wait for test process");

        assert!(!status.success());
        assert_eq!(worker_pid.load(Ordering::SeqCst), 0);
    }
}
