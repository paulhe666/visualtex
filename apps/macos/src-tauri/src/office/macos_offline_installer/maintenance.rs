//! The maintenance log, and install, repair and uninstall in a child process of the app: one at a time, with a
//! time limit, so a wait inside macOS never freezes the app.

use super::*;

pub(super) fn office_maintenance_log_path() -> Result<PathBuf, String> {
    Ok(user_home()?
        .join("Library/Logs/VisualTeX")
        .join(OFFICE_MAINTENANCE_LOG_NAME))
}

pub(super) fn append_office_maintenance_log(
    operation: &str,
    phase: &str,
    outcome: &str,
    detail: &str,
) {
    let Ok(path) = office_maintenance_log_path() else {
        return;
    };
    let Some(parent) = path.parent() else {
        return;
    };
    if fs::create_dir_all(parent).is_err() {
        return;
    }

    // Keep the diagnostic file useful without allowing it to grow forever.
    if fs::metadata(&path)
        .map(|metadata| metadata.len() > 2 * 1024 * 1024)
        .unwrap_or(false)
    {
        let rotated = path.with_extension("log.1");
        let _ = fs::remove_file(&rotated);
        let _ = fs::rename(&path, rotated);
    }

    let Ok(mut file) = OpenOptions::new().create(true).append(true).open(&path) else {
        return;
    };
    let unix_ms = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default();
    let process_snapshot = phase == "operation" || outcome == "error";
    let word_running = process_snapshot
        .then(|| application_running(WORD_PROCESS_NAME))
        .map(|value| value.to_string())
        .unwrap_or_else(|| "not-probed".to_string());
    let powerpoint_running = process_snapshot
        .then(|| application_running(POWERPOINT_PROCESS_NAME))
        .map(|value| value.to_string())
        .unwrap_or_else(|| "not-probed".to_string());
    let _ = writeln!(
        file,
        "\n===== VisualTeX Office add-in maintenance =====\nunix_ms={unix_ms}\nversion={}\noperation={operation}\nphase={phase}\noutcome={outcome}\nword_running={word_running}\npowerpoint_running={powerpoint_running}\ndetail:\n{detail}\n",
        env!("CARGO_PKG_VERSION"),
    );
}

pub(super) fn maintenance_step<T>(
    operation: &str,
    phase: &str,
    action: impl FnOnce() -> Result<T, String>,
) -> Result<T, String> {
    append_office_maintenance_log(operation, phase, "begin", "");
    match action() {
        Ok(value) => {
            append_office_maintenance_log(operation, phase, "ok", "");
            Ok(value)
        }
        Err(error) => {
            append_office_maintenance_log(operation, phase, "error", &error);
            Err(error)
        }
    }
}

pub(super) fn run_maintenance_operation<T>(
    operation: &str,
    action: impl FnOnce() -> Result<T, String>,
) -> Result<T, String> {
    let started = Instant::now();
    append_office_maintenance_log(operation, "operation", "begin", "");
    match action() {
        Ok(value) => {
            append_office_maintenance_log(
                operation,
                "operation",
                "ok",
                &format!("elapsed_ms={}", started.elapsed().as_millis()),
            );
            Ok(value)
        }
        Err(error) => {
            append_office_maintenance_log(
                operation,
                "operation",
                "error",
                &format!(
                    "elapsed_ms={}\nerror={error}",
                    started.elapsed().as_millis()
                ),
            );
            Err(error)
        }
    }
}

pub fn open_office_maintenance_log() -> Result<(), String> {
    let path = office_maintenance_log_path()?;
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)
            .map_err(|error| format!("Unable to create Office log directory: {error}"))?;
    }
    if !path.exists() {
        let mut file = OpenOptions::new()
            .create(true)
            .append(true)
            .open(&path)
            .map_err(|error| format!("Unable to create Office maintenance log: {error}"))?;
        writeln!(
            file,
            "VisualTeX Office add-in maintenance log\nversion={}\n",
            env!("CARGO_PKG_VERSION")
        )
        .map_err(|error| format!("Unable to initialize Office maintenance log: {error}"))?;
    }

    let mut child = Command::new("/usr/bin/open")
        .arg(&path)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .map_err(|error| format!("Unable to open Office maintenance log: {error}"))?;
    let deadline = Instant::now() + Duration::from_secs(2);
    loop {
        match child.try_wait() {
            Ok(Some(status)) if status.success() => return Ok(()),
            Ok(Some(status)) => {
                return Err(format!(
                    "macOS could not open the Office maintenance log (status {status})"
                ))
            }
            Ok(None) if Instant::now() < deadline => {
                thread::sleep(Duration::from_millis(20));
            }
            Ok(None) | Err(_) => {
                let _ = child.kill();
                let _ = child.wait();
                return Err("Timed out opening the Office maintenance log".to_string());
            }
        }
    }
}

/// The add-in maintenance child: `visualtex --office-addin-maintenance
/// <install|repair|uninstall> <resource root>`. It runs before the Tauri
/// builder, so the single-instance plugin never hands it to the running app.
pub const MAINTENANCE_CHILD_ARGUMENT: &str = "--office-addin-maintenance";

/// How long the app waits for one install, repair or uninstall. A normal run
/// takes up to about half a minute (24 s measured); a longer one is waiting on
/// macOS.
pub(super) const MAINTENANCE_TIMEOUT: Duration = Duration::from_secs(90);

pub const MAINTENANCE_TIMEOUT_ERROR: &str =
    "The Office add-in operation timed out: macOS did not let VisualTeX change the Office folders in time. If macOS asks whether VisualTeX may access data from other apps, click Allow, then try again.";

pub(super) static MAINTENANCE_IN_FLIGHT: AtomicBool = AtomicBool::new(false);

pub(super) struct MaintenanceInFlight;

impl MaintenanceInFlight {
    pub(super) fn acquire() -> Result<Self, String> {
        if MAINTENANCE_IN_FLIGHT.swap(true, Ordering::AcqRel) {
            return Err("Another VisualTeX Office add-in operation is still running.".to_string());
        }
        Ok(Self)
    }
}

impl Drop for MaintenanceInFlight {
    fn drop(&mut self) {
        MAINTENANCE_IN_FLIGHT.store(false, Ordering::Release);
    }
}

pub fn run_maintenance_cli_if_requested() -> Option<i32> {
    let arguments = std::env::args_os().collect::<Vec<_>>();
    if arguments.get(1).map(|value| value.as_os_str())
        != Some(std::ffi::OsStr::new(MAINTENANCE_CHILD_ARGUMENT))
    {
        return None;
    }
    let operation = arguments
        .get(2)
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_string();
    let Some(root) = arguments.get(3).map(PathBuf::from) else {
        eprintln!("{MAINTENANCE_CHILD_ARGUMENT} requires an operation and a resource folder");
        return Some(2);
    };
    let result = match operation.as_str() {
        "install" | "repair" => {
            run_maintenance_operation(&operation, || install_impl(&root, &operation)).and_then(
                |status| {
                    serde_json::to_string(&status).map_err(|error| {
                        format!("Unable to report the Office add-in status: {error}")
                    })
                },
            )
        }
        "uninstall" => run_maintenance_operation("uninstall", uninstall_impl)
            .map(|()| "null".to_string()),
        _ => Err(format!(
            "Unknown Office add-in maintenance operation: {operation}"
        )),
    };
    match result {
        Ok(output) => {
            println!("{output}");
            Some(0)
        }
        Err(error) => {
            eprintln!("{error}");
            Some(1)
        }
    }
}

pub(super) fn spawn_pipe_reader<R: std::io::Read + Send + 'static>(
    pipe: Option<R>,
) -> thread::JoinHandle<String> {
    thread::spawn(move || {
        let mut text = String::new();
        if let Some(mut pipe) = pipe {
            let _ = pipe.read_to_string(&mut text);
        }
        text
    })
}

pub(super) struct ChildOutput {
    pub(super) status: std::process::ExitStatus,
    pub(super) stdout: String,
    pub(super) stderr: String,
}

/// Waits for a child whose stdout and stderr are piped, and kills it once
/// `timeout` has passed (None). The pipes are drained while it runs.
pub(super) fn wait_for_child_output(
    mut child: std::process::Child,
    timeout: Duration,
) -> Result<Option<ChildOutput>, String> {
    let stdout = spawn_pipe_reader(child.stdout.take());
    let stderr = spawn_pipe_reader(child.stderr.take());
    let deadline = Instant::now() + timeout;
    let status = loop {
        match child.try_wait() {
            Ok(Some(status)) => break status,
            Ok(None) if Instant::now() < deadline => thread::sleep(Duration::from_millis(50)),
            Ok(None) => {
                let _ = child.kill();
                let _ = child.wait();
                return Ok(None);
            }
            Err(error) => {
                let _ = child.kill();
                let _ = child.wait();
                return Err(error.to_string());
            }
        }
    };
    Ok(Some(ChildOutput {
        status,
        stdout: stdout.join().unwrap_or_default(),
        stderr: stderr.join().unwrap_or_default(),
    }))
}

/// Runs one maintenance operation in a child process and stops it after
/// `timeout`. The first write into Microsoft's group container can wait inside
/// open() until the user answers macOS's "access data from other apps"
/// prompt, and a thread cannot be taken out of that wait; a process can.
pub(super) fn run_maintenance_child(
    executable: &Path,
    operation: &str,
    root: &Path,
    timeout: Duration,
) -> Result<String, String> {
    let child = Command::new(executable)
        .arg(MAINTENANCE_CHILD_ARGUMENT)
        .arg(operation)
        .arg(root)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|error| format!("Unable to start the Office add-in {operation}: {error}"))?;
    let Some(output) = wait_for_child_output(child, timeout)
        .map_err(|error| format!("Unable to wait for the Office add-in {operation}: {error}"))?
    else {
        return Err(MAINTENANCE_TIMEOUT_ERROR.to_string());
    };
    if output.status.success() {
        Ok(output.stdout.trim().to_string())
    } else if output.stderr.trim().is_empty() {
        Err(format!("The Office add-in {operation} failed ({})", output.status))
    } else {
        Err(output.stderr.trim().to_string())
    }
}

pub(super) async fn run_maintenance_command(
    app: AppHandle,
    operation: &'static str,
) -> Result<serde_json::Value, String> {
    if !cfg!(target_os = "macos") {
        return Err("The native offline Office add-ins are available only on macOS".to_string());
    }
    let root = resource_root(&app)?;
    let executable = std::env::current_exe()
        .map_err(|error| format!("Unable to locate the VisualTeX executable: {error}"))?;
    let worker = tauri::async_runtime::spawn_blocking(move || {
        let _in_flight = MaintenanceInFlight::acquire()?;
        let result = run_maintenance_child(&executable, operation, &root, MAINTENANCE_TIMEOUT);
        if matches!(&result, Err(error) if error == MAINTENANCE_TIMEOUT_ERROR) {
            append_office_maintenance_log(
                operation,
                "operation",
                "timeout",
                &format!("stopped after {} s", MAINTENANCE_TIMEOUT.as_secs()),
            );
        }
        result
    });
    let output = worker
        .await
        .map_err(|error| format!("Office add-in {operation} worker failed: {error}"))??;
    serde_json::from_str(&output)
        .map_err(|error| format!("The Office add-in {operation} returned an unreadable status: {error}"))
}
