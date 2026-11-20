//! Word and PowerPoint as applications: whether they are installed and running, and asking them to quit.

use super::*;

pub(super) fn application_paths(home: &Path, app_name: &str) -> [PathBuf; 2] {
    [
        PathBuf::from("/Applications").join(app_name),
        home.join("Applications").join(app_name),
    ]
}

pub(super) fn application_bundle_identifier(app_name: &str) -> Option<&'static str> {
    match app_name {
        WORD_APP_NAME => Some("com.microsoft.Word"),
        POWERPOINT_APP_NAME => Some("com.microsoft.Powerpoint"),
        _ => None,
    }
}

#[cfg(target_os = "macos")]
pub(super) fn application_installed(app_name: &str) -> bool {
    let standard_location_found = user_home()
        .map(|home| {
            application_paths(&home, app_name)
                .into_iter()
                .any(|path| bounded_path_test(&path, "-d"))
        })
        .unwrap_or(false);
    if standard_location_found {
        return true;
    }

    let Some(bundle_identifier) = application_bundle_identifier(app_name) else {
        return false;
    };
    let query = format!("kMDItemCFBundleIdentifier == '{bundle_identifier}'");
    let mut child = match Command::new("/usr/bin/mdfind")
        .arg(query)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
    {
        Ok(child) => child,
        Err(_) => return false,
    };
    let deadline = Instant::now() + Duration::from_millis(1500);
    loop {
        match child.try_wait() {
            Ok(Some(_)) => {
                let Ok(output) = child.wait_with_output() else {
                    return false;
                };
                if !output.status.success() {
                    return false;
                }
                return String::from_utf8_lossy(&output.stdout)
                    .lines()
                    .map(str::trim)
                    .filter(|value| !value.is_empty())
                    .map(Path::new)
                    .any(|path| bounded_path_test(path, "-d"));
            }
            Ok(None) if Instant::now() < deadline => {
                thread::sleep(Duration::from_millis(20));
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
pub(super) fn application_installed(app_name: &str) -> bool {
    user_home()
        .map(|home| {
            application_paths(&home, app_name)
                .into_iter()
                .any(|path| path.is_dir())
        })
        .unwrap_or(false)
}

#[cfg(target_os = "macos")]
pub(super) fn application_running(process_name: &str) -> bool {
    let mut child = match Command::new("/usr/bin/pgrep")
        .args(["-x", process_name])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
    {
        Ok(child) => child,
        Err(_) => return false,
    };
    let deadline = Instant::now() + Duration::from_millis(500);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => return status.success(),
            Ok(None) if Instant::now() < deadline => {
                thread::sleep(Duration::from_millis(20));
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
pub(super) fn application_running(_process_name: &str) -> bool {
    false
}

pub(super) fn ensure_office_hosts_stopped() -> Result<(), String> {
    let running = [
        (WORD_PROCESS_NAME, application_running(WORD_PROCESS_NAME)),
        (
            POWERPOINT_PROCESS_NAME,
            application_running(POWERPOINT_PROCESS_NAME),
        ),
    ]
    .into_iter()
    .filter_map(|(name, is_running)| is_running.then_some(name))
    .collect::<Vec<_>>();
    if running.is_empty() {
        Ok(())
    } else {
        Err(format!(
            "Fully quit {} with Command-Q before installing, repairing, or uninstalling VisualTeX add-ins. Office keeps already-loaded VBA code in memory, so changing files while it is running cannot safely validate the add-in state.",
            running.join(" and ")
        ))
    }
}

#[cfg(target_os = "macos")]
pub(super) fn request_office_application_quit(process_name: &str, bundle_identifier: &str) -> Result<(), String> {
    if !application_running(process_name) {
        return Ok(());
    }
    let script = format!("tell application id \"{bundle_identifier}\" to quit");
    let mut child = Command::new("/usr/bin/osascript")
        .args(["-e", &script])
        .spawn()
        .map_err(|error| format!("Unable to request {process_name} to quit: {error}"))?;
    crate::spawn_background_task("visualtex-office-quit-wait", move || {
        let _ = child.wait();
    });
    Ok(())
}

#[cfg(not(target_os = "macos"))]
pub(super) fn request_office_application_quit(_process_name: &str, _bundle_identifier: &str) -> Result<(), String> {
    Ok(())
}

pub fn request_office_hosts_quit_for_update() -> Result<(), String> {
    if !cfg!(target_os = "macos") {
        return Err("Office add-in update assistance is available only on macOS".to_string());
    }
    request_office_application_quit(WORD_PROCESS_NAME, "com.microsoft.Word")?;
    request_office_application_quit(POWERPOINT_PROCESS_NAME, "com.microsoft.Powerpoint")?;
    Ok(())
}
