//! The add-ins' health signal: written by the app, read back for the settings.

use super::*;

#[cfg(target_os = "macos")]
pub(crate) fn refresh_health_signal(host: &str) -> bool {
    let (process_name, script) = match host {
        "word" => (
            "Microsoft Word",
            r#"tell application "Microsoft Word" to run VB macro macro name "VisualTeX_InitializeWordHost""#,
        ),
        "powerpoint" => (
            "Microsoft PowerPoint",
            r#"tell application "Microsoft PowerPoint" to run VB macro macro name "Auto_Open" list of parameters {}"#,
        ),
        _ => return false,
    };
    let running = Command::new("/usr/bin/pgrep")
        .args(["-x", process_name])
        .output()
        .is_ok_and(|output| output.status.success());
    if !running {
        return false;
    }
    let mut child = match Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(script)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
    {
        Ok(child) => child,
        Err(_) => return false,
    };
    let deadline = Instant::now() + Duration::from_millis(1000);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => return status.success(),
            Ok(None) if Instant::now() < deadline => {
                std::thread::sleep(Duration::from_millis(20));
            }
            Ok(None) | Err(_) => {
                let _ = child.kill();
                let _ = child.wait();
                return false;
            }
        }
    }
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn refresh_health_signal(_host: &str) -> bool {
    false
}

pub(crate) fn health_path(host: &str) -> Result<PathBuf, String> {
    Ok(runtime_root(host_from_request_name(host)?)?
        .join("OfficePluginStatus")
        .join(format!("{host}.json")))
}

pub(super) fn read_health(host: &str) -> Result<MacOfflinePluginHealth, String> {
    let path = health_path(host)?;
    let fallback = || MacOfflinePluginHealth {
        loaded: false,
        plugin_version: None,
        source_revision: None,
        host: host.to_string(),
        timestamp: None,
        status_path: path.display().to_string(),
    };
    let bytes = match fs::read(&path) {
        Ok(bytes) => bytes,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(fallback()),
        Err(error) => return Err(format!("Unable to read {} health: {error}", host)),
    };
    let value: Value = serde_json::from_slice(&bytes)
        .map_err(|error| format!("{host} health file contains invalid JSON: {error}"))?;
    let plugin_version = value
        .get("pluginVersion")
        .and_then(Value::as_str)
        .map(str::to_string);
    let source_revision = value
        .get("sourceRevision")
        .and_then(Value::as_str)
        .map(str::to_string);
    let reported_host = value
        .get("host")
        .and_then(Value::as_str)
        .unwrap_or(host)
        .to_string();
    let timestamp = value
        .get("timestamp")
        .and_then(Value::as_str)
        .filter(|value| {
            !value.is_empty() && value.len() <= 64 && !value.chars().any(char::is_control)
        })
        .map(str::to_string);
    let loaded = value
        .get("loaded")
        .and_then(Value::as_bool)
        .unwrap_or(false)
        && plugin_version.as_deref() == Some(env!("CARGO_PKG_VERSION"))
        && reported_host == host
        && timestamp.is_some();
    Ok(MacOfflinePluginHealth {
        loaded,
        plugin_version,
        source_revision,
        host: reported_host,
        timestamp,
        status_path: path.display().to_string(),
    })
}
