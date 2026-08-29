use serde::{Deserialize, Serialize};
use std::fs::{self, OpenOptions};
use std::io::{BufReader, Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, ExitStatus, Stdio};
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use std::sync::mpsc::{self, RecvTimeoutError};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

#[cfg(windows)]
use std::ffi::c_void;
#[cfg(windows)]
use std::os::windows::process::CommandExt;

#[cfg(windows)]
const CREATE_NO_WINDOW: u32 = 0x0800_0000;
#[cfg(windows)]
const MB_ERR_INVALID_CHARS: u32 = 0x0000_0008;
#[cfg(windows)]
const PROCESS_QUERY_LIMITED_INFORMATION: u32 = 0x0000_1000;

#[cfg(windows)]
pub fn windows_powershell_executable() -> Result<PathBuf, String> {
    let mut roots = Vec::new();
    for name in ["SystemRoot", "WINDIR"] {
        if let Some(value) = std::env::var_os(name) {
            let root = PathBuf::from(value);
            if !roots.iter().any(|existing| existing == &root) {
                roots.push(root);
            }
        }
    }
    let conventional = PathBuf::from(r"C:\Windows");
    if !roots.iter().any(|existing| existing == &conventional) {
        roots.push(conventional);
    }
    for root in roots {
        for system_directory in ["Sysnative", "System32"] {
            let candidate = root
                .join(system_directory)
                .join("WindowsPowerShell")
                .join("v1.0")
                .join("powershell.exe");
            if candidate.is_file() {
                return Ok(candidate);
            }
        }
    }
    Err("Windows PowerShell was not found under SystemRoot\\System32\\WindowsPowerShell\\v1.0; VisualTeX does not use PATH lookup for OCR maintenance".to_string())
}

#[cfg(windows)]
fn powershell_command() -> Result<Command, String> {
    Ok(Command::new(windows_powershell_executable()?))
}

#[cfg(windows)]
#[link(name = "kernel32")]
unsafe extern "system" {
    fn GetACP() -> u32;
    fn GetOEMCP() -> u32;
    fn MultiByteToWideChar(
        code_page: u32,
        flags: u32,
        multi_byte: *const i8,
        byte_count: i32,
        wide_char: *mut u16,
        wide_count: i32,
    ) -> i32;
    fn OpenProcess(
        desired_access: u32,
        inherit_handle: i32,
        process_id: u32,
    ) -> *mut c_void;
    fn GetModuleHandleW(module_name: *const u16) -> *mut c_void;
    fn GetProcAddress(module: *mut c_void, proc_name: *const i8) -> *mut c_void;
    fn CloseHandle(handle: *mut c_void) -> i32;
}

pub const INSTALL_STATUS_SCHEMA: u32 = 1;
pub const INSTALL_STATUS_FILE: &str = "install-status.json";
pub const INSTALL_PID_FILE: &str = "install-process.json";
pub const INSTALL_LOG_FILE: &str = "ocr-install.log";
pub const PREVIOUS_INSTALL_LOG_FILE: &str = "ocr-install.previous.log";
const INSTALL_ACTIVITY_POLL_INTERVAL: Duration = Duration::from_secs(2);
const INSTALL_STATUS_UPDATE_INTERVAL: Duration = Duration::from_secs(1);

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum InstallState {
    NotInstalled,
    Installing,
    InstallFailed,
    DependenciesInstalled,
    Verifying,
    VerificationFailed,
    Complete,
    Cancelled,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct InstallSnapshot {
    pub schema_version: u32,
    pub state: InstallState,
    pub current_step: Option<String>,
    pub completed_steps: Vec<String>,
    pub percent: u8,
    pub message: String,
    pub detail: Option<String>,
    pub error: Option<String>,
    pub log_path: String,
    pub updated_at_ms: u128,
}

impl InstallSnapshot {
    pub fn new(log_path: &Path) -> Self {
        Self {
            schema_version: INSTALL_STATUS_SCHEMA,
            state: InstallState::NotInstalled,
            current_step: None,
            completed_steps: Vec::new(),
            percent: 0,
            message: "OCR runtime is not installed".to_string(),
            detail: None,
            error: None,
            log_path: log_path.display().to_string(),
            updated_at_ms: now_ms(),
        }
    }

    pub fn touch(&mut self) {
        self.updated_at_ms = now_ms();
    }

    pub fn mark_step_complete(&mut self, step: &str) {
        if !self.completed_steps.iter().any(|item| item == step) {
            self.completed_steps.push(step.to_string());
        }
        self.touch();
    }

    pub fn step_complete(&self, step: &str) -> bool {
        self.completed_steps.iter().any(|item| item == step)
    }
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ActiveProcessRecord {
    pid: u32,
    runtime_root: String,
    started_at_ms: u128,
}

#[derive(Default)]
pub struct InstallControl {
    running: AtomicBool,
    cancel_generation: AtomicU64,
    active_pid: AtomicU32,
    snapshot: Mutex<Option<InstallSnapshot>>,
}

impl InstallControl {
    pub fn begin(self: &Arc<Self>) -> Result<InstallLease, String> {
        self.running
            .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
            .map_err(|_| "OCR installation is already running. Repeated install requests are blocked.".to_string())?;
        let generation = self.cancel_generation.load(Ordering::SeqCst);
        Ok(InstallLease {
            control: self.clone(),
            generation,
        })
    }

    pub fn is_running(&self) -> bool {
        self.running.load(Ordering::SeqCst)
    }

    pub fn cancel(&self) -> Result<bool, String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let pid = self.active_pid.load(Ordering::SeqCst);
        if pid == 0 {
            return Ok(false);
        }
        terminate_process_tree(pid)?;
        let _ = self
            .active_pid
            .compare_exchange(pid, 0, Ordering::SeqCst, Ordering::SeqCst);
        Ok(true)
    }

    pub fn cancellation_generation(&self) -> u64 {
        self.cancel_generation.load(Ordering::SeqCst)
    }

    pub fn active_pid(&self) -> u32 {
        self.active_pid.load(Ordering::SeqCst)
    }

    pub fn set_snapshot(&self, snapshot: InstallSnapshot) -> Result<(), String> {
        let mut guard = self
            .snapshot
            .lock()
            .map_err(|_| "OCR installation status lock is unavailable".to_string())?;
        *guard = Some(snapshot);
        Ok(())
    }

    pub fn snapshot(&self) -> Result<Option<InstallSnapshot>, String> {
        self.snapshot
            .lock()
            .map(|snapshot| snapshot.clone())
            .map_err(|_| "OCR installation status lock is unavailable".to_string())
    }

    pub fn update_activity(&self, runtime_root: &Path, snapshot: InstallSnapshot) -> Result<(), String> {
        self.set_snapshot(snapshot.clone())?;
        save_snapshot(runtime_root, &snapshot)
    }
}

pub struct InstallLease {
    control: Arc<InstallControl>,
    generation: u64,
}

impl InstallLease {
    pub fn generation(&self) -> u64 {
        self.generation
    }

    pub fn control(&self) -> &Arc<InstallControl> {
        &self.control
    }
}

impl Drop for InstallLease {
    fn drop(&mut self) {
        self.control.active_pid.store(0, Ordering::SeqCst);
        self.control.running.store(false, Ordering::SeqCst);
    }
}

#[derive(Debug, Clone, Copy)]
pub struct CommandLimits {
    pub step_timeout: Duration,
    pub idle_timeout: Duration,
}

#[derive(Debug)]
pub struct CommandCapture {
    pub status: ExitStatus,
    pub stdout: String,
    pub stderr: String,
    pub elapsed: Duration,
}

enum StreamMessage {
    Stdout(Vec<u8>),
    Stderr(Vec<u8>),
    Closed,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct FilesystemActivity {
    total_bytes: u64,
    newest_modified_ms: u128,
    file_count: u64,
}

#[derive(Debug, Default)]
struct PipProgressTracker {
    package: Option<String>,
    total_size: Option<String>,
    latest_progress: Option<String>,
    installing: bool,
    recent_text: String,
}

impl PipProgressTracker {
    fn observe(&mut self, text: &str) {
        self.recent_text.push_str(text);
        if self.recent_text.chars().count() > 32 * 1024 {
            self.recent_text = tail_chars(&self.recent_text, 16 * 1024);
        }
        let observed = self.recent_text.clone();
        for segment in observed.split(['\r', '\n']) {
            let line = segment.trim();
            if line.is_empty() {
                continue;
            }
            if let Some(download) = line.split_once("Downloading ").map(|(_, value)| value) {
                let package = download
                    .split(" (")
                    .next()
                    .unwrap_or(download)
                    .trim()
                    .trim_end_matches(':');
                if !package.is_empty() {
                    self.package = Some(package.to_string());
                }
                if let Some(start) = download.rfind(" (") {
                    let size_and_rest = &download[start + 2..];
                    if let Some(end) = size_and_rest.find(')') {
                        self.total_size = Some(size_and_rest[..end].to_string());
                    }
                }
                self.installing = false;
                self.latest_progress = Some(line.to_string());
            } else if line.contains("Installing collected packages")
                || line.starts_with("Installing ")
            {
                self.installing = true;
                self.latest_progress = Some(line.to_string());
            } else if line.contains('%')
                || (line.contains('/')
                    && [" KB", " MB", " GB", " kB"]
                        .iter()
                        .any(|unit| line.contains(unit)))
            {
                self.latest_progress = Some(line.to_string());
            }
        }
    }

    fn status(&self, elapsed: Duration) -> Option<(String, Option<String>)> {
        let waited = format_elapsed(elapsed);
        if self.installing {
            let package = self.package.as_deref().unwrap_or("已下载的 Python 依赖");
            return Some((
                format!("正在安装 {package}"),
                Some(format!("已等待 {waited} · 下载缓存和现有环境将被保留")),
            ));
        }
        let package = self.package.as_deref()?;
        let mut parts = Vec::new();
        if let Some(size) = &self.total_size {
            parts.push(format!("文件大小 {size}"));
        }
        parts.push(format!("已等待 {waited}"));
        if let Some(progress) = &self.latest_progress {
            if !progress.contains("Downloading ") {
                parts.push(progress.clone());
            }
        }
        Some((format!("正在下载 {package}"), Some(parts.join(" · "))))
    }
}

fn format_elapsed(elapsed: Duration) -> String {
    let total_seconds = elapsed.as_secs();
    let minutes = total_seconds / 60;
    let seconds = total_seconds % 60;
    if minutes == 0 {
        format!("{seconds} 秒")
    } else {
        format!("{minutes} 分 {seconds} 秒")
    }
}

fn filesystem_activity(runtime_root: &Path) -> FilesystemActivity {
    let mut activity = FilesystemActivity::default();
    let mut pending = vec![runtime_root.join("cache").join("pip"), runtime_root.join("tmp")];
    let mut visited = 0_usize;
    while let Some(path) = pending.pop() {
        if visited >= 50_000 {
            break;
        }
        let Ok(metadata) = fs::metadata(&path) else {
            continue;
        };
        if metadata.is_file() {
            visited += 1;
            activity.file_count = activity.file_count.saturating_add(1);
            activity.total_bytes = activity.total_bytes.saturating_add(metadata.len());
            if let Ok(modified) = metadata.modified() {
                if let Ok(duration) = modified.duration_since(UNIX_EPOCH) {
                    activity.newest_modified_ms =
                        activity.newest_modified_ms.max(duration.as_millis());
                }
            }
            continue;
        }
        let Ok(entries) = fs::read_dir(path) else {
            continue;
        };
        for entry in entries.flatten() {
            pending.push(entry.path());
        }
    }
    activity
}

fn poll_filesystem_activity(
    runtime_root: &Path,
    last_poll: &mut Instant,
    previous: &mut FilesystemActivity,
    last_activity: &mut Instant,
) -> bool {
    if last_poll.elapsed() < INSTALL_ACTIVITY_POLL_INTERVAL {
        return false;
    }
    *last_poll = Instant::now();
    let current = filesystem_activity(runtime_root);
    if current == *previous {
        return false;
    }
    *previous = current;
    *last_activity = Instant::now();
    true
}

fn update_live_install_status(
    control: &InstallControl,
    runtime_root: &Path,
    tracker: &PipProgressTracker,
    elapsed: Duration,
    filesystem_changed: bool,
) {
    let Ok(Some(mut snapshot)) = control.snapshot() else {
        return;
    };
    if let Some((message, detail)) = tracker.status(elapsed) {
        snapshot.message = message;
        snapshot.detail = detail;
    } else if filesystem_changed {
        snapshot.message = "正在下载或安装 Python 依赖".to_string();
        snapshot.detail = Some(format!(
            "下载缓存或临时文件仍在增长 · 已等待 {}",
            format_elapsed(elapsed)
        ));
    } else {
        return;
    }
    snapshot.error = None;
    snapshot.touch();
    let _ = control.update_activity(runtime_root, snapshot);
}

pub fn install_log_path(runtime_root: &Path) -> PathBuf {
    runtime_root.join("logs").join(INSTALL_LOG_FILE)
}

pub fn previous_install_log_path(runtime_root: &Path) -> PathBuf {
    runtime_root.join("logs").join(PREVIOUS_INSTALL_LOG_FILE)
}

pub fn begin_install_log_session(runtime_root: &Path) -> Result<(), String> {
    let path = install_log_path(runtime_root);
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)
            .map_err(|error| format!("Unable to create OCR installation log directory: {error}"))?;
    }
    let previous = previous_install_log_path(runtime_root);
    if path.is_file() {
        if previous.exists() {
            fs::remove_file(&previous).map_err(|error| {
                format!(
                    "Unable to replace the previous OCR installation log {}: {error}",
                    previous.display()
                )
            })?;
        }
        if let Err(rename_error) = fs::rename(&path, &previous) {
            fs::copy(&path, &previous).map_err(|copy_error| {
                format!(
                    "Unable to rotate OCR installation log {} to {}: rename={rename_error}; copy={copy_error}",
                    path.display(),
                    previous.display()
                )
            })?;
            fs::remove_file(&path).map_err(|error| {
                format!(
                    "Unable to clear the previous OCR installation log {} after rotation: {error}",
                    path.display()
                )
            })?;
        }
    }
    let mut file = OpenOptions::new()
        .create(true)
        .write(true)
        .truncate(true)
        .open(&path)
        .map_err(|error| format!("Unable to start a new OCR installation log: {error}"))?;
    writeln!(
        file,
        "===== VisualTeX OCR install session | unix_ms={} =====",
        now_ms()
    )
    .map_err(|error| format!("Unable to initialize OCR installation log: {error}"))
}

pub fn install_status_path(runtime_root: &Path) -> PathBuf {
    runtime_root.join(INSTALL_STATUS_FILE)
}

pub fn active_process_path(runtime_root: &Path) -> PathBuf {
    runtime_root.join(INSTALL_PID_FILE)
}

pub fn load_snapshot(runtime_root: &Path) -> Option<InstallSnapshot> {
    let content = fs::read(install_status_path(runtime_root)).ok()?;
    let snapshot: InstallSnapshot = serde_json::from_slice(&content).ok()?;
    (snapshot.schema_version == INSTALL_STATUS_SCHEMA).then_some(snapshot)
}

pub fn save_snapshot(runtime_root: &Path, snapshot: &InstallSnapshot) -> Result<(), String> {
    fs::create_dir_all(runtime_root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    let target = install_status_path(runtime_root);
    let temporary = target.with_extension("json.tmp");
    let content = serde_json::to_vec_pretty(snapshot)
        .map_err(|error| format!("Unable to serialize OCR installation status: {error}"))?;
    fs::write(&temporary, content)
        .map_err(|error| format!("Unable to write OCR installation status: {error}"))?;
    if target.exists() {
        fs::remove_file(&target)
            .map_err(|error| format!("Unable to replace OCR installation status: {error}"))?;
    }
    fs::rename(&temporary, &target)
        .map_err(|error| format!("Unable to publish OCR installation status: {error}"))
}

pub fn append_install_log(runtime_root: &Path, message: &str) -> Result<(), String> {
    let path = install_log_path(runtime_root);
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)
            .map_err(|error| format!("Unable to create OCR installation log directory: {error}"))?;
    }
    let mut file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .map_err(|error| format!("Unable to open OCR installation log: {error}"))?;
    writeln!(file, "{message}")
        .map_err(|error| format!("Unable to write OCR installation log: {error}"))
}

pub fn cleanup_runtime_processes(runtime_root: &Path) -> Result<Vec<u32>, String> {
    #[cfg(windows)]
    {
        let escaped_root = runtime_root.display().to_string().replace('\'', "''");
        let script = format!(
            "$root='{escaped_root}'.Replace('/',[char]92).ToLowerInvariant(); [Console]::OutputEncoding=[Text.UTF8Encoding]::new(); Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {{ $_.ProcessId -ne $PID -and $_.Name -match '^(python|pythonw|pip|pip3|cargo|rustc|maturin)(\\.exe)?$' -and ((($_.ExecutablePath + ' ' + $_.CommandLine).Replace('/',[char]92).ToLowerInvariant()).Contains($root)) }} | ForEach-Object {{ [Console]::WriteLine($_.ProcessId) }}"
        );
        let output = powershell_command()?
            .args(["-NoProfile", "-NonInteractive", "-Command", &script])
            .creation_flags(CREATE_NO_WINDOW)
            .env("PYTHONUTF8", "1")
            .env("PYTHONIOENCODING", "utf-8")
            .output()
            .map_err(|error| format!("Unable to scan residual OCR installation processes: {error}"))?;
        if !output.status.success() {
            return Err(format!(
                "Residual OCR process scan failed with {}. stdout:\n{}\nstderr:\n{}",
                output.status,
                decode_process_output(&output.stdout),
                decode_process_output(&output.stderr)
            ));
        }
        let mut terminated = Vec::new();
        for pid in decode_process_output(&output.stdout)
            .lines()
            .filter_map(|line| line.trim().parse::<u32>().ok())
        {
            if pid == 0 || pid == std::process::id() || terminated.contains(&pid) {
                continue;
            }
            match terminate_process_tree(pid) {
                Ok(()) => terminated.push(pid),
                Err(error) => {
                    if process_belongs_to_runtime(pid, runtime_root)? {
                        return Err(error);
                    }
                }
            }
        }
        Ok(terminated)
    }
    #[cfg(not(windows))]
    {
        let _ = runtime_root;
        Ok(Vec::new())
    }
}

pub fn cleanup_runtime_processes_for_install(runtime_root: &Path) -> Vec<u32> {
    best_effort_runtime_process_cleanup_for_install(
        runtime_root,
        cleanup_runtime_processes(runtime_root),
    )
}

pub(crate) fn best_effort_runtime_process_cleanup_for_install(
    runtime_root: &Path,
    result: Result<Vec<u32>, String>,
) -> Vec<u32> {
    match result {
        Ok(terminated) => terminated,
        Err(error) => {
            // Full-process inventory is only a defensive pre-install cleanup.
            // Enterprise AppLocker/WDAC policies can block VisualTeX from
            // launching PowerShell even though the application itself is allowed.
            // The explicitly recorded installer PID is validated separately using
            // native Win32 APIs, so this auxiliary inventory must not make OCR
            // installation require administrator elevation or policy changes.
            let warning = format!(
                "Residual OCR process inventory was skipped before installation: {error}"
            );
            eprintln!("{warning}");
            let _ = append_install_log(runtime_root, &warning);
            Vec::new()
        }
    }
}

pub fn cleanup_stale_process(runtime_root: &Path) -> Result<bool, String> {
    let path = active_process_path(runtime_root);
    let Ok(content) = fs::read(&path) else {
        return Ok(false);
    };
    let record: ActiveProcessRecord = match serde_json::from_slice(&content) {
        Ok(record) => record,
        Err(_) => {
            let _ = fs::remove_file(path);
            return Ok(false);
        }
    };
    let expected_root = runtime_root.display().to_string();
    if record.runtime_root != expected_root {
        return Err("Refusing to terminate an OCR process whose runtime root does not match this application environment".to_string());
    }
    if !process_belongs_to_runtime(record.pid, runtime_root)? {
        let _ = fs::remove_file(path);
        return Ok(false);
    }
    let terminated = terminate_process_tree(record.pid).is_ok();
    let _ = fs::remove_file(path);
    Ok(terminated)
}

fn write_active_process(runtime_root: &Path, pid: u32) -> Result<(), String> {
    let record = ActiveProcessRecord {
        pid,
        runtime_root: runtime_root.display().to_string(),
        started_at_ms: now_ms(),
    };
    let content = serde_json::to_vec_pretty(&record)
        .map_err(|error| format!("Unable to serialize OCR installation process record: {error}"))?;
    let target = active_process_path(runtime_root);
    let temporary = target.with_extension("json.tmp");
    fs::write(&temporary, content)
        .map_err(|error| format!("Unable to write OCR installation process record: {error}"))?;
    if target.exists() {
        fs::remove_file(&target)
            .map_err(|error| format!("Unable to replace OCR installation process record: {error}"))?;
    }
    fs::rename(&temporary, &target)
        .map_err(|error| format!("Unable to publish OCR installation process record: {error}"))
}

fn clear_active_process(runtime_root: &Path, control: &InstallControl, pid: u32) {
    let _ = control
        .active_pid
        .compare_exchange(pid, 0, Ordering::SeqCst, Ordering::SeqCst);
    let _ = fs::remove_file(active_process_path(runtime_root));
}

pub fn run_logged_command(
    command: &mut Command,
    label: &str,
    runtime_root: &Path,
    control: &InstallControl,
    generation: u64,
    limits: CommandLimits,
) -> Result<CommandCapture, String> {
    let log_path = install_log_path(runtime_root);
    if let Some(parent) = log_path.parent() {
        fs::create_dir_all(parent)
            .map_err(|error| format!("Unable to create OCR installation log directory: {error}"))?;
    }
    let mut log = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log_path)
        .map_err(|error| format!("Unable to open OCR installation log: {error}"))?;
    writeln!(
        log,
        "\n===== {label} | unix_ms={} =====",
        now_ms()
    )
    .map_err(|error| format!("Unable to initialize OCR installation log: {error}"))?;
    writeln!(log, "command={command:?}")
        .map_err(|error| format!("Unable to record OCR installation command: {error}"))?;

    command
        .env_remove("PYTHONPATH")
        .env_remove("PYTHONHOME")
        .env_remove("PYTHONUSERBASE")
        .env("PYTHONNOUSERSITE", "1")
        .env("PYTHONSAFEPATH", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .env("PYTHONUNBUFFERED", "1")
        .env("PIP_DISABLE_PIP_VERSION_CHECK", "1")
        .env("PIP_NO_INPUT", "1")
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    #[cfg(windows)]
    command.creation_flags(CREATE_NO_WINDOW);

    let started = Instant::now();
    let mut child = command
        .spawn()
        .map_err(|error| format!("Unable to start {label}: {error}; log={}", log_path.display()))?;
    let pid = child.id();
    control.active_pid.store(pid, Ordering::SeqCst);
    if let Err(error) = write_active_process(runtime_root, pid) {
        let _ = terminate_process_tree(pid);
        let _ = child.wait();
        clear_active_process(runtime_root, control, pid);
        return Err(error);
    }

    let stdout = match child.stdout.take() {
        Some(stdout) => stdout,
        None => {
            let _ = terminate_process_tree(pid);
            let _ = child.wait();
            clear_active_process(runtime_root, control, pid);
            return Err(format!("{label} stdout is unavailable"));
        }
    };
    let stderr = match child.stderr.take() {
        Some(stderr) => stderr,
        None => {
            let _ = terminate_process_tree(pid);
            let _ = child.wait();
            clear_active_process(runtime_root, control, pid);
            return Err(format!("{label} stderr is unavailable"));
        }
    };
    let (sender, receiver) = mpsc::channel();
    spawn_stream_reader(stdout, sender.clone(), true);
    spawn_stream_reader(stderr, sender.clone(), false);
    drop(sender);

    let mut stdout_bytes = Vec::new();
    let mut stderr_bytes = Vec::new();
    let mut last_activity = Instant::now();
    let mut last_filesystem_poll = Instant::now();
    let mut last_status_update = Instant::now() - INSTALL_STATUS_UPDATE_INTERVAL;
    let mut filesystem_state = filesystem_activity(runtime_root);
    let mut progress_tracker = PipProgressTracker::default();
    let mut closed_streams = 0usize;

    let result = (|| -> Result<CommandCapture, String> {
        loop {
            if control.cancellation_generation() != generation {
                let _ = terminate_process_tree(pid);
                let _ = child.wait();
                break Err(format!(
                    "{label} was cancelled. The entire installation process tree was terminated. log={}",
                    log_path.display()
                ));
            }
            if started.elapsed() > limits.step_timeout {
                let _ = terminate_process_tree(pid);
                let _ = child.wait();
                let download_detail = progress_tracker
                    .status(started.elapsed())
                    .map(|(message, detail)| {
                        format!(" {message}. {}", detail.unwrap_or_default())
                    })
                    .unwrap_or_default();
                break Err(format!(
                    "{label} timed out because the download or installation exceeded the installer limit of {} minutes.{download_detail} This is separate from pip's 30-second per-socket timeout. The pip cache and existing OCR environment were preserved; retry continues from the missing package. log={}",
                    limits.step_timeout.as_secs() / 60,
                    log_path.display()
                ));
            }
            if last_activity.elapsed() > limits.idle_timeout {
                let _ = terminate_process_tree(pid);
                let _ = child.wait();
                break Err(format!(
                    "{label} showed no stdout/stderr or pip cache/temp-file growth for {} seconds and was stopped as genuinely inactive. The pip cache and existing OCR environment were preserved. log={}",
                    limits.idle_timeout.as_secs(),
                    log_path.display()
                ));
            }

            match receiver.recv_timeout(Duration::from_millis(100)) {
                Ok(StreamMessage::Stdout(bytes)) => {
                    last_activity = Instant::now();
                    progress_tracker.observe(&decode_process_output(&bytes));
                    if last_status_update.elapsed() >= INSTALL_STATUS_UPDATE_INTERVAL {
                        update_live_install_status(
                            control,
                            runtime_root,
                            &progress_tracker,
                            started.elapsed(),
                            false,
                        );
                        last_status_update = Instant::now();
                    }
                    log_bytes(&mut log, "stdout", &bytes)?;
                    stdout_bytes.extend_from_slice(&bytes);
                }
                Ok(StreamMessage::Stderr(bytes)) => {
                    last_activity = Instant::now();
                    progress_tracker.observe(&decode_process_output(&bytes));
                    if last_status_update.elapsed() >= INSTALL_STATUS_UPDATE_INTERVAL {
                        update_live_install_status(
                            control,
                            runtime_root,
                            &progress_tracker,
                            started.elapsed(),
                            false,
                        );
                        last_status_update = Instant::now();
                    }
                    log_bytes(&mut log, "stderr", &bytes)?;
                    stderr_bytes.extend_from_slice(&bytes);
                }
                Ok(StreamMessage::Closed) => closed_streams += 1,
                Err(RecvTimeoutError::Disconnected) => closed_streams = 2,
                Err(RecvTimeoutError::Timeout) => {}
            }

            let filesystem_changed = poll_filesystem_activity(
                runtime_root,
                &mut last_filesystem_poll,
                &mut filesystem_state,
                &mut last_activity,
            );
            if filesystem_changed {
                update_live_install_status(
                    control,
                    runtime_root,
                    &progress_tracker,
                    started.elapsed(),
                    true,
                );
            }

            if control.cancellation_generation() != generation {
                let _ = terminate_process_tree(pid);
                let _ = child.wait();
                break Err(format!(
                    "{label} was cancelled. The entire installation process tree was terminated. log={}",
                    log_path.display()
                ));
            }

            if let Some(status) = child
                .try_wait()
                .map_err(|error| format!("Unable to query {label} process state: {error}"))?
            {
                let drain_deadline = Instant::now() + Duration::from_secs(2);
                while closed_streams < 2 && Instant::now() < drain_deadline {
                    match receiver.recv_timeout(Duration::from_millis(50)) {
                        Ok(StreamMessage::Stdout(bytes)) => {
                            log_bytes(&mut log, "stdout", &bytes)?;
                            stdout_bytes.extend_from_slice(&bytes);
                        }
                        Ok(StreamMessage::Stderr(bytes)) => {
                            log_bytes(&mut log, "stderr", &bytes)?;
                            stderr_bytes.extend_from_slice(&bytes);
                        }
                        Ok(StreamMessage::Closed) => closed_streams += 1,
                        Err(RecvTimeoutError::Disconnected) => break,
                        Err(RecvTimeoutError::Timeout) => {}
                    }
                }
                let capture = CommandCapture {
                    status,
                    stdout: decode_process_output(&stdout_bytes),
                    stderr: decode_process_output(&stderr_bytes),
                    elapsed: started.elapsed(),
                };
                writeln!(
                    log,
                    "exit_status={} elapsed_ms={} closed_streams={closed_streams}",
                    capture.status,
                    capture.elapsed.as_millis()
                )
                .map_err(|error| format!("Unable to finalize OCR installation log: {error}"))?;
                if capture.status.success() {
                    break Ok(capture);
                }
                break Err(format_command_failure(label, &capture, &log_path));
            }
        }
    })();

    if result.is_err() {
        if child.try_wait().ok().flatten().is_none() {
            let _ = terminate_process_tree(pid);
            let _ = child.wait();
        }
        let drain_deadline = Instant::now() + Duration::from_secs(2);
        while closed_streams < 2 && Instant::now() < drain_deadline {
            match receiver.recv_timeout(Duration::from_millis(50)) {
                Ok(StreamMessage::Stdout(bytes)) => {
                    let _ = log_bytes(&mut log, "stdout", &bytes);
                    stdout_bytes.extend_from_slice(&bytes);
                }
                Ok(StreamMessage::Stderr(bytes)) => {
                    let _ = log_bytes(&mut log, "stderr", &bytes);
                    stderr_bytes.extend_from_slice(&bytes);
                }
                Ok(StreamMessage::Closed) => closed_streams += 1,
                Err(RecvTimeoutError::Disconnected) => break,
                Err(RecvTimeoutError::Timeout) => {}
            }
        }
        let _ = writeln!(
            log,
            "aborted elapsed_ms={} captured_stdout_bytes={} captured_stderr_bytes={} closed_streams={closed_streams}",
            started.elapsed().as_millis(),
            stdout_bytes.len(),
            stderr_bytes.len()
        );
        let _ = log.flush();
    }

    clear_active_process(runtime_root, control, pid);
    result
}

fn spawn_stream_reader<R: Read + Send + 'static>(reader: R, sender: mpsc::Sender<StreamMessage>, stdout: bool) {
    thread::spawn(move || {
        let mut reader = BufReader::new(reader);
        let mut buffer = vec![0_u8; 8 * 1024];
        loop {
            match reader.read(&mut buffer) {
                Ok(0) => break,
                Ok(count) => {
                    let bytes = buffer[..count].to_vec();
                    let message = if stdout {
                        StreamMessage::Stdout(bytes)
                    } else {
                        StreamMessage::Stderr(bytes)
                    };
                    if sender.send(message).is_err() {
                        return;
                    }
                }
                Err(_) => break,
            }
        }
        let _ = sender.send(StreamMessage::Closed);
    });
}

fn log_bytes(log: &mut std::fs::File, stream: &str, bytes: &[u8]) -> Result<(), String> {
    let text = decode_process_output(bytes);
    write!(log, "[{stream}] {text}")
        .map_err(|error| format!("Unable to append OCR installation output: {error}"))?;
    log.flush()
        .map_err(|error| format!("Unable to flush OCR installation output: {error}"))
}

pub(crate) fn decode_process_output(bytes: &[u8]) -> String {
    if let Ok(text) = std::str::from_utf8(bytes) {
        return text.to_string();
    }

    #[cfg(windows)]
    {
        let system_code_pages = unsafe { [GetOEMCP(), GetACP(), 936] };
        let mut attempted = Vec::new();
        for code_page in system_code_pages {
            if code_page == 65001 || attempted.contains(&code_page) {
                continue;
            }
            attempted.push(code_page);
            if let Some(text) = decode_windows_code_page(bytes, code_page) {
                return text;
            }
        }
    }

    String::from_utf8_lossy(bytes).to_string()
}

#[cfg(windows)]
fn decode_windows_code_page(bytes: &[u8], code_page: u32) -> Option<String> {
    if bytes.is_empty() {
        return Some(String::new());
    }
    let byte_count = i32::try_from(bytes.len()).ok()?;
    let required = unsafe {
        MultiByteToWideChar(
            code_page,
            MB_ERR_INVALID_CHARS,
            bytes.as_ptr().cast(),
            byte_count,
            std::ptr::null_mut(),
            0,
        )
    };
    if required <= 0 {
        return None;
    }
    let mut wide = vec![0_u16; required as usize];
    let written = unsafe {
        MultiByteToWideChar(
            code_page,
            MB_ERR_INVALID_CHARS,
            bytes.as_ptr().cast(),
            byte_count,
            wide.as_mut_ptr(),
            required,
        )
    };
    if written <= 0 {
        return None;
    }
    Some(String::from_utf16_lossy(&wide[..written as usize]))
}

fn format_command_failure(label: &str, capture: &CommandCapture, log_path: &Path) -> String {
    let stdout = tail_chars(&capture.stdout, 6000);
    let stderr = tail_chars(&capture.stderr, 6000);
    format!(
        "{label} failed with {}.\nstdout:\n{}\nstderr:\n{}\nFull log: {}",
        capture.status,
        stdout.trim(),
        stderr.trim(),
        log_path.display()
    )
}

fn tail_chars(value: &str, max_chars: usize) -> String {
    let count = value.chars().count();
    if count <= max_chars {
        value.to_string()
    } else {
        value.chars().skip(count - max_chars).collect()
    }
}

#[cfg(windows)]
pub(crate) fn process_belongs_to_runtime(pid: u32, runtime_root: &Path) -> Result<bool, String> {
    // Do not use PowerShell/WMI here. This check is part of the stale-PID safety
    // boundary and must still work on enterprise machines where powershell.exe is
    // disabled by AppLocker/WDAC. Every command recorded by run_logged_command is
    // launched from VisualTeX's private OCR runtime, so its executable path is a
    // sufficient ownership check.
    let process = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, pid) };
    if process.is_null() {
        // The PID no longer exists, belongs to a protected process, or has already
        // been reused by something we cannot inspect. In all cases it is unsafe to
        // terminate it and safe to treat the VisualTeX PID record as stale.
        return Ok(false);
    }

    type QueryFullProcessImageNameWFn = unsafe extern "system" fn(
        *mut c_void,
        u32,
        *mut u16,
        *mut u32,
    ) -> i32;
    let module_name: Vec<u16> = "kernel32.dll\0".encode_utf16().collect();
    let module = unsafe { GetModuleHandleW(module_name.as_ptr()) };
    let proc = if module.is_null() {
        std::ptr::null_mut()
    } else {
        unsafe { GetProcAddress(module, b"QueryFullProcessImageNameW\0".as_ptr().cast()) }
    };
    if proc.is_null() {
        unsafe {
            CloseHandle(process);
        }
        return Ok(false);
    }
    let query_full_process_image_name: QueryFullProcessImageNameWFn =
        unsafe { std::mem::transmute(proc) };
    let mut executable_path = vec![0u16; 32_768];
    let mut path_size = executable_path.len() as u32;
    let query_result = unsafe {
        query_full_process_image_name(
            process,
            0,
            executable_path.as_mut_ptr(),
            &mut path_size,
        )
    };
    unsafe {
        CloseHandle(process);
    }
    if query_result == 0 || path_size == 0 {
        return Ok(false);
    }

    let process_text = String::from_utf16_lossy(&executable_path[..path_size as usize])
        .replace('/', "\\")
        .trim_end_matches('\\')
        .to_ascii_lowercase();
    let root_text = runtime_root
        .display()
        .to_string()
        .replace('/', "\\")
        .trim_end_matches('\\')
        .to_ascii_lowercase();
    let root_prefix = format!("{root_text}\\");
    Ok(process_text == root_text || process_text.starts_with(&root_prefix))
}

#[cfg(not(windows))]
fn process_belongs_to_runtime(_pid: u32, _runtime_root: &Path) -> Result<bool, String> {
    Ok(true)
}

pub fn terminate_process_tree(pid: u32) -> Result<(), String> {
    if pid == 0 {
        return Ok(());
    }
    #[cfg(windows)]
    {
        let status = Command::new("taskkill")
            .args(["/PID", &pid.to_string(), "/T", "/F"])
            .creation_flags(CREATE_NO_WINDOW)
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .status()
            .map_err(|error| format!("Unable to terminate OCR installation process tree {pid}: {error}"))?;
        if status.success() {
            return Ok(());
        }
        return Err(format!(
            "taskkill could not terminate OCR installation process tree {pid}: {status}"
        ));
    }
    #[cfg(unix)]
    {
        let result = unsafe { libc::kill(pid as i32, libc::SIGKILL) };
        if result == 0 {
            return Ok(());
        }
        let error = std::io::Error::last_os_error();
        if error.raw_os_error() == Some(libc::ESRCH) {
            return Ok(());
        }
        Err(format!("Unable to terminate OCR installation process {pid}: {error}"))
    }
}

fn now_ms() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    fn test_command(script: &str) -> Command {
        #[cfg(windows)]
        {
            let mut command = powershell_command().expect("absolute Windows PowerShell path");
            command.args(["-NoProfile", "-Command", script]);
            command
        }
        #[cfg(not(windows))]
        {
            let mut command = Command::new("sh");
            command.args(["-c", script]);
            command
        }
    }

    #[test]
    fn install_log_session_rotates_the_previous_attempt() {
        let root = tempdir().unwrap();
        append_install_log(root.path(), "older attempt").unwrap();
        begin_install_log_session(root.path()).unwrap();
        let current = fs::read_to_string(install_log_path(root.path())).unwrap();
        let previous = fs::read_to_string(previous_install_log_path(root.path())).unwrap();
        assert!(current.contains("VisualTeX OCR install session"));
        assert!(!current.contains("older attempt"));
        assert!(previous.contains("older attempt"));
    }

    #[test]
    fn policy_blocked_residual_scan_does_not_block_installation() {
        let root = tempdir().unwrap();
        let terminated = best_effort_runtime_process_cleanup_for_install(
            root.path(),
            Err("Unable to scan residual OCR installation processes: os error 786".to_string()),
        );
        assert!(terminated.is_empty());
        let log = fs::read_to_string(install_log_path(root.path())).unwrap();
        assert!(log.contains("inventory was skipped before installation"));
        assert!(log.contains("786"));
    }

    #[cfg(windows)]
    #[test]
    fn stale_pid_ownership_uses_native_process_image_path() {
        let executable = std::env::current_exe().unwrap();
        let executable_parent = executable.parent().unwrap();
        assert!(process_belongs_to_runtime(std::process::id(), executable_parent).unwrap());
        let unrelated = tempdir().unwrap();
        assert!(!process_belongs_to_runtime(std::process::id(), unrelated.path()).unwrap());
    }

    #[test]
    fn duplicate_installation_is_rejected() {
        let control = Arc::new(InstallControl::default());
        let first = control.begin().expect("first install should acquire the lock");
        let error = match control.begin() {
            Ok(_) => panic!("duplicate install should be rejected"),
            Err(error) => error,
        };
        assert!(error.contains("already running"));
        drop(first);
        assert!(control.begin().is_ok());
    }

    #[test]
    fn command_failure_includes_stdout_stderr_and_exit_code() {
        let root = tempdir().unwrap();
        let control = Arc::new(InstallControl::default());
        let lease = control.begin().unwrap();
        #[cfg(windows)]
        let script = "[Console]::OutputEncoding=[Text.UTF8Encoding]::new(); Write-Output '标准输出'; [Console]::Error.WriteLine('错误输出'); exit 7";
        #[cfg(not(windows))]
        let script = "printf 'standard output\\n'; printf 'error output\\n' >&2; exit 7";
        let error = run_logged_command(
            &mut test_command(script),
            "failing command",
            root.path(),
            lease.control(),
            lease.generation(),
            CommandLimits {
                step_timeout: Duration::from_secs(5),
                idle_timeout: Duration::from_secs(2),
            },
        )
        .expect_err("command should fail");
        assert!(error.contains("failed with"));
        assert!(error.contains("stdout"));
        assert!(error.contains("stderr"));
        let log = fs::read_to_string(install_log_path(root.path())).unwrap();
        assert!(log.contains("exit_status"));
    }

    #[test]
    fn command_timeout_terminates_instead_of_hanging() {
        let root = tempdir().unwrap();
        let control = Arc::new(InstallControl::default());
        let lease = control.begin().unwrap();
        #[cfg(windows)]
        let script = "Start-Sleep -Seconds 5";
        #[cfg(not(windows))]
        let script = "sleep 5";
        let error = run_logged_command(
            &mut test_command(script),
            "slow command",
            root.path(),
            lease.control(),
            lease.generation(),
            CommandLimits {
                step_timeout: Duration::from_millis(350),
                idle_timeout: Duration::from_secs(2),
            },
        )
        .expect_err("command should time out");
        assert!(error.contains("timed out"));
        assert_eq!(lease.control().active_pid(), 0);
    }

    #[test]
    fn utf8_output_is_preserved_in_install_log() {
        let root = tempdir().unwrap();
        let control = Arc::new(InstallControl::default());
        let lease = control.begin().unwrap();
        #[cfg(windows)]
        let script = "[Console]::OutputEncoding=[Text.UTF8Encoding]::new(); Write-Output '正在安装依赖'";
        #[cfg(not(windows))]
        let script = "printf '正在安装依赖\\n'";
        run_logged_command(
            &mut test_command(script),
            "utf8 command",
            root.path(),
            lease.control(),
            lease.generation(),
            CommandLimits {
                step_timeout: Duration::from_secs(5),
                idle_timeout: Duration::from_secs(2),
            },
        )
        .expect("UTF-8 command should succeed");
        let log = fs::read_to_string(install_log_path(root.path())).unwrap();
        assert!(log.contains("正在安装依赖"));
    }

    #[cfg(windows)]
    #[test]
    fn windows_gbk_output_is_converted_to_utf8() {
        let gbk = b"\xd0\xc5\xcf\xa2: \xd3\xc3\xcc\xe1\xb9\xa9\xb5\xc4\xc4\xa3\xca\xbd\xce\xde\xb7\xa8\xd5\xd2\xb5\xbd\xce\xc4\xbc\xfe\xa1\xa3";
        assert_eq!(decode_process_output(gbk), "信息: 用提供的模式无法找到文件。");
    }

    #[test]
    fn snapshots_round_trip_and_keep_completed_steps() {
        let root = tempdir().unwrap();
        let mut snapshot = InstallSnapshot::new(&install_log_path(root.path()));
        snapshot.state = InstallState::InstallFailed;
        snapshot.current_step = Some("tokenizers".to_string());
        snapshot.mark_step_complete("paddle");
        save_snapshot(root.path(), &snapshot).unwrap();
        let loaded = load_snapshot(root.path()).expect("snapshot should load");
        assert_eq!(loaded.state, InstallState::InstallFailed);
        assert!(loaded.step_complete("paddle"));
        assert!(!loaded.step_complete("tokenizers"));
    }
}
