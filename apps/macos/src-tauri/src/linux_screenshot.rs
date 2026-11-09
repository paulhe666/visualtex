//! A screenshot of an area the user selects, on Linux: the desktop's screenshot portal
//! (xdg-desktop-portal, GNOME and KDE on X11 and Wayland), else a screenshot program found on the
//! system (gnome-screenshot, spectacle, grim with slurp, maim, scrot, ImageMagick's import).

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::Duration;
use zbus::zvariant::{OwnedObjectPath, OwnedValue, Value};

/// The PNG of the selected area; None when the user cancelled.
pub(crate) fn capture_selection_png(max_bytes: usize) -> Result<Option<Vec<u8>>, String> {
    let file = match portal_screenshot() {
        Ok(file) => file,
        Err(portal_error) => match program_screenshot()? {
            Some(file) => Some(file),
            None if portal_error.is_empty() => None,
            None => return Err(portal_error),
        },
    };
    let Some(file) = file else { return Ok(None) };
    let bytes = std::fs::read(&file).map_err(|error| format!("Unable to read the screenshot: {error}"));
    let _ = std::fs::remove_file(&file);
    let bytes = bytes?;
    if bytes.is_empty() {
        return Ok(None);
    }
    if bytes.len() > max_bytes {
        return Err("The captured screenshot is too large for OCR".to_string());
    }
    Ok(Some(bytes))
}

/// The portal's interactive screenshot: Ok(None) when cancelled, Err when there is no portal.
fn portal_screenshot() -> Result<Option<PathBuf>, String> {
    let connection = zbus::blocking::Connection::session().map_err(|error| format!("no session bus: {error}"))?;
    let token = format!("visualtex{}", std::process::id());
    let sender = connection
        .unique_name()
        .map(|name| name.trim_start_matches(':').replace('.', "_"))
        .unwrap_or_default();
    let request_path = format!("/org/freedesktop/portal/desktop/request/{sender}/{token}");
    let request = zbus::blocking::Proxy::new(
        &connection,
        "org.freedesktop.portal.Desktop",
        request_path.as_str(),
        "org.freedesktop.portal.Request",
    )
    .map_err(|error| error.to_string())?;
    let mut responses = request.receive_signal("Response").map_err(|error| error.to_string())?;
    let mut options: HashMap<&str, Value> = HashMap::new();
    options.insert("interactive", Value::from(true));
    options.insert("handle_token", Value::from(token.as_str()));
    let reply: OwnedObjectPath = connection
        .call_method(
            Some("org.freedesktop.portal.Desktop"),
            "/org/freedesktop/portal/desktop",
            Some("org.freedesktop.portal.Screenshot"),
            "Screenshot",
            &("", options),
        )
        .map_err(|error| format!("no screenshot portal: {error}"))?
        .body()
        .deserialize()
        .map_err(|error| error.to_string())?;
    if reply.as_str() != request_path {
        // an older portal answers on a path of its own
        let other = zbus::blocking::Proxy::new(&connection, "org.freedesktop.portal.Desktop", reply.as_str(), "org.freedesktop.portal.Request")
            .map_err(|error| error.to_string())?;
        responses = other.receive_signal("Response").map_err(|error| error.to_string())?;
    }
    let (sender_done, receiver) = std::sync::mpsc::sync_channel(1);
    std::thread::spawn(move || {
        if let Some(message) = responses.next() {
            let _ = sender_done.send(message.body().deserialize::<(u32, HashMap<String, OwnedValue>)>().ok());
        }
    });
    let (response, results) = receiver
        .recv_timeout(Duration::from_secs(300))
        .map_err(|_| String::new())?
        .ok_or_else(|| "The screenshot portal answered nothing readable".to_string())?;
    if response != 0 {
        return Ok(None);
    }
    let uri: String = results
        .get("uri")
        .and_then(|value| String::try_from(value.clone()).ok())
        .ok_or_else(|| "The screenshot portal gave no file".to_string())?;
    let path = uri.strip_prefix("file://").map(percent_decode).ok_or_else(|| format!("Unreadable screenshot address {uri}"))?;
    Ok(Some(PathBuf::from(path)))
}

fn percent_decode(text: &str) -> String {
    let bytes = text.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' && index + 2 < bytes.len() {
            if let Ok(byte) = u8::from_str_radix(&text[index + 1..index + 3], 16) {
                out.push(byte);
                index += 3;
                continue;
            }
        }
        out.push(bytes[index]);
        index += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

fn has(program: &str) -> bool {
    std::env::var_os("PATH")
        .map(|paths| std::env::split_paths(&paths).any(|directory| directory.join(program).is_file()))
        .unwrap_or(false)
}

/// A screenshot program's area selection into a file; None when none ran or the user cancelled.
fn program_screenshot() -> Result<Option<PathBuf>, String> {
    let file = std::env::temp_dir().join(format!("visualtex-ocr-{}.png", std::process::id()));
    let _ = std::fs::remove_file(&file);
    let path = file.to_string_lossy().to_string();
    let wayland = std::env::var_os("WAYLAND_DISPLAY").is_some();
    let attempts: Vec<(&str, Vec<String>)> = vec![
        ("gnome-screenshot", vec!["-a".into(), "-f".into(), path.clone()]),
        ("spectacle", vec!["-r".into(), "-b".into(), "-n".into(), "-o".into(), path.clone()]),
        ("maim", vec!["-s".into(), path.clone()]),
        ("scrot", vec!["-s".into(), "-o".into(), path.clone()]),
        ("import", vec![path.clone()]),
    ];
    if wayland && has("grim") && has("slurp") {
        let area = Command::new("slurp").output().map_err(|error| error.to_string())?;
        if !area.status.success() {
            return Ok(None);
        }
        let geometry = String::from_utf8_lossy(&area.stdout).trim().to_string();
        let status = Command::new("grim").args(["-g", &geometry, &path]).status().map_err(|error| error.to_string())?;
        return Ok(status.success().then_some(file).filter(|file| Path::new(file).is_file()));
    }
    for (program, arguments) in attempts {
        if !has(program) {
            continue;
        }
        let status = Command::new(program).args(&arguments).status().map_err(|error| format!("{program}: {error}"))?;
        return Ok((status.success() && file.is_file()).then_some(file));
    }
    Err("No screenshot tool: install gnome-screenshot, spectacle or maim (or a desktop with xdg-desktop-portal)".to_string())
}
