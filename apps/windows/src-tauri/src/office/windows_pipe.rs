#![cfg(target_os = "windows")]

use getrandom::fill as fill_random;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::fs::{File, OpenOptions};
use std::io::{BufRead, BufReader, Write};
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::{mpsc, Mutex};
use std::thread;
use std::time::{Duration, Instant};
use uuid::Uuid;

const CREATE_NO_WINDOW: u32 = 0x0800_0000;
pub const WINDOWS_OFFICE_PROTOCOL_VERSION: u32 = 1;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PipeRequest {
    pub protocol_version: u32,
    pub id: String,
    pub method: String,
    #[serde(default)]
    pub params: Value,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PipeError {
    pub code: String,
    pub message: String,
    #[serde(default)]
    pub retryable: bool,
    #[serde(default)]
    pub details: Option<Value>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PipeResponse {
    pub protocol_version: u32,
    pub id: String,
    pub ok: bool,
    #[serde(default)]
    pub result: Option<Value>,
    #[serde(default)]
    pub error: Option<PipeError>,
}

enum PipeProcess {
    NotStarted,
    Running(Child),
    Stopped,
}

pub struct WindowsPipeClient {
    pipe_name: String,
    token: String,
    executable: PathBuf,
    temp_root: PathBuf,
    log_root: PathBuf,
    child: Mutex<PipeProcess>,
}

impl WindowsPipeClient {
    pub fn new(executable: PathBuf, temp_root: PathBuf, log_root: PathBuf) -> Result<Self, String> {
        let sid = current_user_sid()?;
        let mut token = [0_u8; 32];
        fill_random(&mut token)
            .map_err(|error| format!("Unable to create Office pipe token: {error}"))?;
        Ok(Self {
            pipe_name: format!(r"\\.\pipe\VisualTeX.OfficeBridge.{sid}"),
            token: hex::encode(token),
            executable,
            temp_root,
            log_root,
            child: Mutex::new(PipeProcess::NotStarted),
        })
    }

    pub fn pipe_name(&self) -> &str {
        &self.pipe_name
    }

    pub fn is_healthy(&self) -> bool {
        self.request_method("health", json!({})).is_ok()
    }

    pub fn request_value(&self, request: Value) -> Result<Value, String> {
        let request: PipeRequest = serde_json::from_value(request)
            .map_err(|error| format!("Invalid Windows Office request: {error}"))?;
        self.request(request)
    }

    pub fn request_method(&self, method: &str, params: Value) -> Result<Value, String> {
        self.request(PipeRequest {
            protocol_version: WINDOWS_OFFICE_PROTOCOL_VERSION,
            id: Uuid::new_v4().to_string(),
            method: method.to_string(),
            params,
        })
    }

    pub fn events_after(&self, cursor: u64) -> Vec<Value> {
        self.request_method("events.after", json!({ "cursor": cursor }))
            .ok()
            .and_then(|value| value.get("events").cloned())
            .and_then(|value| value.as_array().cloned())
            .unwrap_or_default()
    }

    pub fn shutdown(&self) -> Result<(), String> {
        let mut guard = self
            .child
            .lock()
            .map_err(|_| "Windows Office bridge process lock is poisoned".to_string())?;
        let PipeProcess::Running(mut child) = std::mem::replace(&mut *guard, PipeProcess::Stopped)
        else {
            return Ok(());
        };
        if child
            .try_wait()
            .map_err(|error| error.to_string())?
            .is_some()
        {
            return Ok(());
        }
        let pipe_name = self.pipe_name.clone();
        let token = self.token.clone();
        let (sender, receiver) = mpsc::channel();
        // Only the already-owned child can receive shutdown. A stuck pipe must
        // not prevent its owner from releasing the process during application exit.
        let request = PipeRequest {
            protocol_version: WINDOWS_OFFICE_PROTOCOL_VERSION,
            id: Uuid::new_v4().to_string(),
            method: "shutdown".to_string(),
            params: json!({}),
        };
        thread::spawn(move || {
            let result = exchange(&pipe_name, &token, request, Duration::from_secs(2));
            let _ = sender.send(result);
        });
        let deadline = Instant::now() + Duration::from_secs(2);
        let _ = receiver.recv_timeout(Duration::from_secs(2));
        while Instant::now() < deadline {
            if child
                .try_wait()
                .map_err(|error| error.to_string())?
                .is_some()
            {
                return Ok(());
            }
            thread::sleep(Duration::from_millis(25));
        }
        if child
            .try_wait()
            .map_err(|error| error.to_string())?
            .is_none()
        {
            child
                .kill()
                .map_err(|error| format!("Unable to stop owned Office sidecar: {error}"))?;
        }
        child.wait().map_err(|error| error.to_string())?;
        Ok(())
    }

    fn request(&self, request: PipeRequest) -> Result<Value, String> {
        if request.protocol_version != WINDOWS_OFFICE_PROTOCOL_VERSION {
            return Err("Unsupported Windows Office bridge protocol version".to_string());
        }
        self.ensure_started()?;
        exchange(
            &self.pipe_name,
            &self.token,
            request,
            Duration::from_secs(5),
        )
    }

    fn ensure_started(&self) -> Result<(), String> {
        let mut guard = self
            .child
            .lock()
            .map_err(|_| "Windows Office bridge process lock is poisoned".to_string())?;
        if matches!(*guard, PipeProcess::Stopped) {
            return Err("Windows Office bridge has shut down".to_string());
        }
        if let PipeProcess::Running(child) = &mut *guard {
            match child.try_wait() {
                Ok(None) => return Ok(()),
                Ok(Some(_)) => {
                    *guard = PipeProcess::NotStarted;
                }
                Err(error) => {
                    return Err(format!("Unable to inspect owned Office sidecar: {error}"))
                }
            }
        }

        if !self.executable.is_file() {
            return Err(format!(
                "Windows Office bridge executable is missing: {}",
                self.executable.display()
            ));
        }
        std::fs::create_dir_all(&self.temp_root)
            .map_err(|error| format!("Unable to create Office temp directory: {error}"))?;
        std::fs::create_dir_all(&self.log_root)
            .map_err(|error| format!("Unable to create Office log directory: {error}"))?;

        let child = Command::new(&self.executable)
            .arg("--parent-pid")
            .arg(std::process::id().to_string())
            .arg("--pipe-name")
            .arg(&self.pipe_name)
            .arg("--token")
            .arg(&self.token)
            .arg("--temp-root")
            .arg(&self.temp_root)
            .arg("--log-root")
            .arg(&self.log_root)
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .creation_flags(CREATE_NO_WINDOW)
            .spawn()
            .map_err(|error| format!("Unable to start Windows Office bridge: {error}"))?;
        *guard = PipeProcess::Running(child);
        Ok(())
    }
}

fn exchange(
    pipe_name: &str,
    token: &str,
    request: PipeRequest,
    timeout: Duration,
) -> Result<Value, String> {
    let mut file = connect_with_retry(pipe_name, timeout)?;
    let reader_file = file
        .try_clone()
        .map_err(|error| format!("Unable to clone Office pipe handle: {error}"))?;
    let mut reader = BufReader::new(reader_file);

    let handshake = PipeRequest {
        protocol_version: WINDOWS_OFFICE_PROTOCOL_VERSION,
        id: Uuid::new_v4().to_string(),
        method: "handshake".to_string(),
        params: json!({ "token": token }),
    };
    write_json_line(&mut file, &handshake)?;
    let handshake_response = read_response(&mut reader)?;
    if !handshake_response.ok {
        return Err(response_error(
            handshake_response,
            "Office bridge authentication failed",
        ));
    }

    write_json_line(&mut file, &request)?;
    let response = read_response(&mut reader)?;
    if !response.ok {
        return Err(response_error(
            response,
            "Windows Office bridge request failed",
        ));
    }
    Ok(response.result.unwrap_or(Value::Null))
}

fn connect_with_retry(pipe_name: &str, timeout: Duration) -> Result<File, String> {
    let deadline = Instant::now() + timeout;
    loop {
        match OpenOptions::new().read(true).write(true).open(pipe_name) {
            Ok(file) => return Ok(file),
            Err(error) if Instant::now() < deadline => {
                let _ = error;
                thread::sleep(Duration::from_millis(75));
            }
            Err(error) => {
                return Err(format!(
                    "Unable to connect to Windows Office pipe {}: {error}",
                    pipe_name
                ))
            }
        }
    }
}
impl Drop for WindowsPipeClient {
    fn drop(&mut self) {
        let _ = self.shutdown();
    }
}

fn write_json_line(file: &mut File, value: &impl Serialize) -> Result<(), String> {
    serde_json::to_writer(&mut *file, value)
        .map_err(|error| format!("Unable to serialize Office pipe request: {error}"))?;
    file.write_all(b"\n")
        .map_err(|error| format!("Unable to write Office pipe request: {error}"))?;
    file.flush()
        .map_err(|error| format!("Unable to flush Office pipe request: {error}"))
}

fn read_response(reader: &mut BufReader<File>) -> Result<PipeResponse, String> {
    let mut line = String::new();
    reader
        .read_line(&mut line)
        .map_err(|error| format!("Unable to read Office pipe response: {error}"))?;
    if line.trim().is_empty() {
        return Err("Windows Office bridge closed the pipe without a response".to_string());
    }
    serde_json::from_str(&line)
        .map_err(|error| format!("Invalid Windows Office bridge response: {error}"))
}

fn response_error(response: PipeResponse, fallback: &str) -> String {
    response
        .error
        .map(|error| format!("{}: {}", error.code, error.message))
        .unwrap_or_else(|| fallback.to_string())
}

fn current_user_sid() -> Result<String, String> {
    let output = Command::new("whoami.exe")
        .args(["/user", "/fo", "csv", "/nh"])
        .creation_flags(CREATE_NO_WINDOW)
        .output()
        .map_err(|error| format!("Unable to query current Windows user SID: {error}"))?;
    if !output.status.success() {
        return Err("whoami.exe failed while querying the current user SID".to_string());
    }
    let text = String::from_utf8_lossy(&output.stdout);
    let fields = text
        .trim()
        .split(',')
        .map(|value| value.trim().trim_matches('"'))
        .collect::<Vec<_>>();
    let sid = fields
        .last()
        .filter(|value| value.starts_with("S-1-"))
        .ok_or_else(|| format!("Unable to parse current Windows user SID from: {text}"))?;
    Ok((*sid).to_string())
}

pub fn locate_sidecar(app: Option<&tauri::AppHandle>) -> PathBuf {
    let bundled_filename = "visualtex-windows-office-bridge.exe";
    let development_filename = "visualtex-windows-office-bridge-x86_64-pc-windows-msvc.exe";
    let mut candidates = Vec::new();
    if let Ok(current) = std::env::current_exe() {
        if let Some(parent) = current.parent() {
            candidates.push(parent.join(bundled_filename));
            candidates.push(parent.join(development_filename));
            candidates.push(parent.join("binaries").join(development_filename));
        }
    }
    if let Some(app) = app {
        use tauri::Manager;
        if let Ok(resource) = app.path().resource_dir() {
            candidates.push(resource.join(bundled_filename));
            candidates.push(resource.join(development_filename));
            candidates.push(resource.join("binaries").join(development_filename));
        }
    }
    candidates
        .into_iter()
        .find(|path| path.is_file())
        .unwrap_or_else(|| Path::new("src-tauri/binaries").join(development_filename))
}
