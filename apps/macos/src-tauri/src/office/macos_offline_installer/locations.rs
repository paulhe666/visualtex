//! Where the add-ins go: the Word startup folders and the PowerPoint add-in folders, looked at with bounded
//! file-system probes, the files earlier versions left there, and whether macOS lets VisualTeX in at all.

use super::*;

pub(super) fn is_legacy_visualtex_wef_file(path: &Path) -> bool {
    let Some(name) = path.file_name().and_then(|value| value.to_str()) else {
        return false;
    };
    let name = name.to_ascii_lowercase();
    name.contains(LEGACY_WORD_MANIFEST_ID)
        || name.contains(LEGACY_POWERPOINT_MANIFEST_ID)
        || name.starts_with("visualtex.word.xml")
        || name.starts_with("visualtex.powerpoint.xml")
}

pub(super) fn collect_legacy_visualtex_wef_files(
    directory: &Path,
    depth: usize,
    output: &mut Vec<PathBuf>,
) -> Result<(), String> {
    if depth > 4 {
        return Ok(());
    }
    // Never call fs::read_dir directly inside an Office sandbox. A few macOS
    // builds can leave these container directories coordinated indefinitely.
    // bounded_directory_entries delegates enumeration to a killable child and
    // treats an inaccessible sandbox as "not discoverable" rather than hanging.
    for path in bounded_directory_entries(directory, "")? {
        if bounded_path_test(&path, "-d") {
            collect_legacy_visualtex_wef_files(&path, depth + 1, output)?;
        } else if bounded_path_test(&path, "-f") && is_legacy_visualtex_wef_file(&path) {
            output.push(path);
        }
    }
    Ok(())
}

pub(super) fn discover_legacy_visualtex_wef_files() -> Result<Vec<PathBuf>, String> {
    let home = user_home()?;
    let mut files = Vec::new();
    for host_container in ["com.microsoft.Word", "com.microsoft.Powerpoint"] {
        collect_legacy_visualtex_wef_files(
            &home
                .join("Library/Containers")
                .join(host_container)
                .join("Data/Documents/wef"),
            0,
            &mut files,
        )?;
    }
    files.sort();
    files.dedup();
    Ok(files)
}

pub(super) fn note_office_path_probe_timeout() {
    OFFICE_PATH_PROBE_TIMEOUTS.fetch_add(1, Ordering::Relaxed);
}

#[cfg(target_os = "macos")]
pub(super) fn bounded_path_test(path: &Path, flag: &str) -> bool {
    let mut child = match Command::new("/bin/test")
        .arg(flag)
        .arg(path)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
    {
        Ok(child) => child,
        Err(_) => return false,
    };
    let deadline = Instant::now() + Duration::from_millis(300);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => return status.success(),
            Ok(None) if Instant::now() < deadline => {
                thread::sleep(Duration::from_millis(20));
            }
            Ok(None) | Err(_) => {
                let _ = child.kill();
                let _ = child.wait();
                note_office_path_probe_timeout();
                return false;
            }
        }
    }
}

#[cfg(not(target_os = "macos"))]
pub(super) fn bounded_path_test(path: &Path, flag: &str) -> bool {
    match flag {
        "-d" => path.is_dir(),
        "-f" => path.is_file(),
        _ => false,
    }
}

pub(super) fn push_existing_directory(paths: &mut Vec<PathBuf>, path: PathBuf) {
    if bounded_path_test(&path, "-d") {
        paths.push(path);
    }
}

pub(super) fn canonical_word_startup_path(home: &Path) -> PathBuf {
    home.join(OFFICE_GROUP_CONTAINER)
        .join("User Content.localized/Startup.localized/Word")
}

pub(super) fn known_word_startup_paths(home: &Path) -> Vec<PathBuf> {
    let office_group = home.join(OFFICE_GROUP_CONTAINER);
    let mut paths = Vec::new();
    for user_content in ["User Content.localized", "User Content"] {
        for startup in ["Startup.localized", "Startup"] {
            push_existing_directory(
                &mut paths,
                office_group.join(user_content).join(startup).join("Word"),
            );
        }
    }
    for path in [
        home.join("Library/Application Support/Microsoft/Office/Startup/Word"),
        home.join(
            "Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/Startup/Word",
        ),
    ] {
        push_existing_directory(&mut paths, path);
    }
    paths
}

pub(crate) fn discover_word_startup_paths() -> Result<Vec<PathBuf>, String> {
    let home = user_home()?;
    let mut paths = known_word_startup_paths(&home);
    // Never recursively walk the Office group container from the settings/status
    // path. On some Macs FileProvider/TCC coordination can block a synchronous
    // read_dir indefinitely. The stable Office startup locations above cover the
    // supported installations; word_startup_install_paths supplies the canonical
    // location when Office has not created it yet.
    paths.sort();
    paths.dedup();
    Ok(paths)
}

pub(super) fn word_startup_install_paths() -> Result<Vec<PathBuf>, String> {
    let mut paths = discover_word_startup_paths()?;
    if paths.is_empty() {
        // Some Office releases do not create the Startup directory until a
        // template has been installed. Use Microsoft's stable group-container
        // location as a safe fallback instead of requiring Word to be opened
        // once before installation.
        paths.push(canonical_word_startup_path(&user_home()?));
    }
    Ok(paths)
}

pub(super) fn is_visualtex_addin_artifact(path: &Path, extension: &str) -> bool {
    path.file_name()
        .and_then(|value| value.to_str())
        .map(|value| {
            let value = value.to_ascii_lowercase();
            let value = value.trim_start_matches('.');
            let value = value.strip_prefix("~$").unwrap_or(value);
            let Some(suffix) = value.strip_prefix("visualtex") else {
                return false;
            };
            let separator_is_owned = suffix
                .chars()
                .next()
                .is_some_and(|character| matches!(character, '.' | ' ' | '(' | '-' | '_'));
            separator_is_owned && suffix.contains(extension)
        })
        .unwrap_or(false)
}

pub(super) fn is_visualtex_word_startup_artifact(path: &Path) -> bool {
    is_visualtex_addin_artifact(path, ".dotm")
}

#[cfg(target_os = "macos")]
pub(super) fn bounded_directory_entries(
    directory: &Path,
    _canonical_name: &str,
) -> Result<Vec<PathBuf>, String> {
    // Office can leave its group-container directories coordinated for tens of
    // seconds after a document or host closes. Enumerate in a killable child.
    // Permission-denied, missing, or timed-out containers are a soft "unknown"
    // status; status/repair UI must remain responsive.
    let mut child = Command::new("/bin/ls")
        .args(["-1A"])
        .arg(directory)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .map_err(|error| format!("Unable to inspect {}: {error}", directory.display()))?;
    let deadline = Instant::now() + Duration::from_millis(900);
    loop {
        match child.try_wait() {
            Ok(Some(_)) => {
                let output = child.wait_with_output().map_err(|error| {
                    format!("Unable to inspect {}: {error}", directory.display())
                })?;
                if !output.status.success() {
                    return Ok(Vec::new());
                }
                return Ok(String::from_utf8_lossy(&output.stdout)
                    .lines()
                    .map(|name| directory.join(name))
                    .collect());
            }
            Ok(None) if Instant::now() < deadline => {
                thread::sleep(Duration::from_millis(20));
            }
            Ok(None) | Err(_) => {
                let _ = child.kill();
                let _ = child.wait();
                note_office_path_probe_timeout();
                return Ok(Vec::new());
            }
        }
    }
}

#[cfg(not(target_os = "macos"))]
pub(super) fn bounded_directory_entries(
    directory: &Path,
    _canonical_name: &str,
) -> Result<Vec<PathBuf>, String> {
    fs::read_dir(directory)
        .map_err(|error| format!("Unable to inspect {}: {error}", directory.display()))?
        .map(|entry| {
            entry
                .map(|entry| entry.path())
                .map_err(|error| format!("Unable to inspect {}: {error}", directory.display()))
        })
        .collect()
}

pub(super) fn discover_word_startup_artifacts() -> Result<Vec<PathBuf>, String> {
    let mut artifacts = Vec::new();
    for startup in discover_word_startup_paths()? {
        for path in bounded_directory_entries(&startup, WORD_ADDIN_NAME)? {
            if bounded_path_test(&path, "-f") && is_visualtex_word_startup_artifact(&path) {
                artifacts.push(path);
            }
        }
    }
    artifacts.sort();
    artifacts.dedup();
    Ok(artifacts)
}

pub(super) fn known_powerpoint_addin_directories() -> Result<Vec<PathBuf>, String> {
    let home = user_home()?;
    let office_group = home.join(OFFICE_GROUP_CONTAINER);
    let mut directories = Vec::new();
    push_existing_directory(&mut directories, offline_root()?.join("OfficeAddins"));
    for user_content in ["User Content.localized", "User Content"] {
        for addins in ["Add-Ins.localized", "Add-Ins", "Addins"] {
            push_existing_directory(
                &mut directories,
                office_group.join(user_content).join(addins),
            );
        }
    }
    for path in [
        home.join("Library/Application Support/Microsoft/Office/PowerPoint Add-Ins"),
        home.join(
            "Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/Microsoft/Office/PowerPoint Add-Ins",
        ),
    ] {
        push_existing_directory(&mut directories, path);
    }
    directories.sort();
    directories.dedup();
    Ok(directories)
}

pub(super) fn discover_powerpoint_addin_artifacts() -> Result<Vec<PathBuf>, String> {
    let mut artifacts = Vec::new();
    for directory in known_powerpoint_addin_directories()? {
        for path in bounded_directory_entries(&directory, POWERPOINT_ADDIN_NAME)? {
            if bounded_path_test(&path, "-f") && is_visualtex_addin_artifact(&path, ".ppam") {
                artifacts.push(path);
            }
        }
    }
    artifacts.sort();
    artifacts.dedup();
    Ok(artifacts)
}

/// Start of the error when macOS refuses VisualTeX Microsoft's group
/// container. macOS 27 never asks for another developer's container (it shows
/// "Data Access Blocked"); earlier releases ask once whether VisualTeX may
/// access data from other apps. Either way Full Disk Access lets it in.
pub const OFFICE_FOLDER_ACCESS_DENIED: &str = "macOS denied VisualTeX access to the Office folders";

/// Writes and removes one file in VisualTeX's folder in Microsoft's group
/// container before anything else there, so that a refusal is reported as
/// such instead of as whichever file the install touched first.
pub(super) fn probe_office_folder_access() -> Result<(), String> {
    let folder = offline_root()?;
    let probe = folder.join(".visualtex-access-probe");
    let result = fs::create_dir_all(&folder)
        .and_then(|()| fs::write(&probe, b""))
        .and_then(|()| fs::remove_file(&probe));
    match result {
        Ok(()) => Ok(()),
        Err(error)
            if error.raw_os_error() == Some(1)
                || error.kind() == std::io::ErrorKind::PermissionDenied =>
        {
            Err(format!("{OFFICE_FOLDER_ACCESS_DENIED}: {error}"))
        }
        Err(error) => Err(format!("Unable to write in {}: {error}", folder.display())),
    }
}
