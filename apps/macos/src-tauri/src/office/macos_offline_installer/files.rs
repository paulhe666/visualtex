//! Writing the add-in files: atomic writes and copies that are read back, backups restored when an install fails,
//! and bounded removal and reads.

use super::*;

#[cfg(target_os = "macos")]
pub(super) fn bounded_remove(path: &Path, recursive: bool) -> Result<(), String> {
    let mut command = Command::new("/bin/rm");
    command.arg(if recursive { "-rf" } else { "-f" }).arg(path);
    let mut child = command
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .map_err(|error| format!("Unable to remove {}: {error}", path.display()))?;
    let deadline = Instant::now() + Duration::from_secs(2);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => {
                return if status.success() {
                    Ok(())
                } else {
                    Err(format!("Unable to remove {}", path.display()))
                };
            }
            Ok(None) if Instant::now() < deadline => {
                thread::sleep(Duration::from_millis(20));
            }
            Ok(None) | Err(_) => {
                let _ = child.kill();
                let _ = child.wait();
                return Err(format!(
                    "Timed out removing VisualTeX Office path {}",
                    path.display()
                ));
            }
        }
    }
}

pub(super) fn remove_dir_if_exists(path: &Path) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        return bounded_remove(path, true);
    }
    #[cfg(not(target_os = "macos"))]
    {
        let metadata = match fs::symlink_metadata(path) {
            Ok(metadata) => metadata,
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(()),
            Err(error) => return Err(format!("Unable to inspect {}: {error}", path.display())),
        };
        if metadata.file_type().is_symlink() || metadata.is_file() {
            return remove_if_exists(path);
        }
        fs::remove_dir_all(path)
            .map_err(|error| format!("Unable to remove {}: {error}", path.display()))
    }
}

pub(super) fn backup_compiled_artifacts_to_scratch(word: &Path, powerpoint: &Path) -> Result<(), String> {
    let backup_root = offline_root()?.join("Scratch/InstallBackup");
    fs::create_dir_all(&backup_root)
        .map_err(|error| format!("Unable to create {}: {error}", backup_root.display()))?;
    for (source, name) in [(word, WORD_ADDIN_NAME), (powerpoint, POWERPOINT_ADDIN_NAME)] {
        let destination = backup_root.join(name);
        copy_atomic(source, &destination, 0o600)?;
        verify_copy(source, &destination)?;
    }
    Ok(())
}

#[cfg(unix)]
pub(super) fn set_mode(path: &Path, mode: u32) -> Result<(), String> {
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(path, fs::Permissions::from_mode(mode))
        .map_err(|error| format!("Unable to set permissions on {}: {error}", path.display()))
}

#[cfg(not(unix))]
pub(super) fn set_mode(_path: &Path, _mode: u32) -> Result<(), String> {
    Ok(())
}

pub(super) fn atomic_write(path: &Path, bytes: &[u8], mode: u32) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("Path has no parent: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let temporary = parent.join(format!(
        ".{}.{}.tmp",
        path.file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("visualtex"),
        Uuid::new_v4()
    ));
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&temporary)
        .map_err(|error| format!("Unable to create {}: {error}", temporary.display()))?;
    file.write_all(bytes)
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    file.sync_all()
        .map_err(|error| format!("Unable to sync {}: {error}", temporary.display()))?;
    set_mode(&temporary, mode)?;
    fs::rename(&temporary, path).map_err(|error| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to replace {}: {error}", path.display())
    })?;
    set_mode(path, mode)
}

pub(super) fn copy_atomic(source: &Path, destination: &Path, mode: u32) -> Result<(), String> {
    let bytes = fs::read(source)
        .map_err(|error| format!("Unable to read {}: {error}", source.display()))?;
    atomic_write(destination, &bytes, mode)
}

#[cfg(target_os = "macos")]
pub(super) fn verify_copy(source: &Path, destination: &Path) -> Result<(), String> {
    let mut child = Command::new("/usr/bin/cmp")
        .arg("-s")
        .arg(source)
        .arg(destination)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .map_err(|error| format!("Unable to verify {}: {error}", destination.display()))?;
    let deadline = Instant::now() + Duration::from_millis(900);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => {
                return if status.success() {
                    Ok(())
                } else {
                    Err(format!(
                        "Installed file {} does not match its VisualTeX source",
                        destination.display()
                    ))
                };
            }
            Ok(None) if Instant::now() < deadline => {
                thread::sleep(Duration::from_millis(20));
            }
            Ok(None) => {
                let _ = child.kill();
                let _ = child.wait();
                // Never fall back to synchronous metadata on a coordinated Office
                // container: the metadata call itself can hang. A timed-out compare
                // is an unknown/mismatched state; the UI can retry without blocking.
                note_office_path_probe_timeout();
                return Err(format!(
                    "Timed out verifying installed VisualTeX file {}",
                    destination.display()
                ));
            }
            Err(error) => {
                let _ = child.kill();
                let _ = child.wait();
                return Err(format!(
                    "Unable to verify {}: {error}",
                    destination.display()
                ));
            }
        }
    }
}

#[cfg(not(target_os = "macos"))]
pub(super) fn verify_copy(source: &Path, destination: &Path) -> Result<(), String> {
    let source_bytes = fs::read(source)
        .map_err(|error| format!("Unable to verify {}: {error}", source.display()))?;
    let destination_bytes = fs::read(destination)
        .map_err(|error| format!("Unable to verify {}: {error}", destination.display()))?;
    if source_bytes == destination_bytes {
        Ok(())
    } else {
        Err(format!(
            "Installed file {} does not match its VisualTeX source",
            destination.display()
        ))
    }
}

pub(super) fn addin_installation_matches(
    source: &Path,
    expected_destinations: &[PathBuf],
    discovered_artifacts: &[PathBuf],
) -> bool {
    if expected_destinations.is_empty() {
        return false;
    }
    let mut expected = expected_destinations.to_vec();
    expected.sort();
    expected.dedup();
    let mut discovered = discovered_artifacts.to_vec();
    discovered.sort();
    discovered.dedup();
    if expected != discovered {
        return false;
    }
    expected
        .iter()
        .all(|destination| verify_copy(source, destination).is_ok())
}

pub(super) fn verify_addin_installation(
    source: &Path,
    expected_destinations: &[PathBuf],
    discovered_artifacts: &[PathBuf],
    host: &str,
) -> Result<(), String> {
    if addin_installation_matches(source, expected_destinations, discovered_artifacts) {
        Ok(())
    } else {
        Err(format!(
            "The installed VisualTeX {host} add-in does not exactly match the packaged version or a stale loadable copy remains"
        ))
    }
}

pub(super) fn capture_backup(path: &Path) -> Result<FileBackup, String> {
    let bytes = match fs::read(path) {
        Ok(bytes) => Some(bytes),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => None,
        Err(error) => return Err(format!("Unable to back up {}: {error}", path.display())),
    };
    Ok(FileBackup {
        path: path.to_path_buf(),
        bytes,
    })
}

pub(super) fn restore_backups(backups: &[FileBackup]) -> Result<(), String> {
    let mut failures = Vec::new();
    for backup in backups.iter().rev() {
        let result = match backup.bytes.as_deref() {
            Some(bytes) => atomic_write(&backup.path, bytes, 0o600),
            None => remove_if_exists(&backup.path),
        };
        if let Err(error) = result {
            failures.push(error);
        }
    }
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("; "))
    }
}

#[cfg(target_os = "macos")]
pub(super) fn bounded_read_small_file(path: &Path, maximum_bytes: usize) -> Option<Vec<u8>> {
    let mut child = Command::new("/bin/cat")
        .arg(path)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;
    let deadline = Instant::now() + Duration::from_millis(1000);
    loop {
        match child.try_wait() {
            Ok(Some(_)) => {
                let output = child.wait_with_output().ok()?;
                if !output.status.success() || output.stdout.len() > maximum_bytes {
                    return None;
                }
                return Some(output.stdout);
            }
            Ok(None) if Instant::now() < deadline => {
                thread::sleep(Duration::from_millis(20));
            }
            Ok(None) | Err(_) => {
                let _ = child.kill();
                let _ = child.wait();
                return None;
            }
        }
    }
}

#[cfg(not(target_os = "macos"))]
pub(super) fn bounded_read_small_file(path: &Path, maximum_bytes: usize) -> Option<Vec<u8>> {
    let bytes = fs::read(path).ok()?;
    (bytes.len() <= maximum_bytes).then_some(bytes)
}

pub(super) fn remove_if_exists(path: &Path) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        return bounded_remove(path, false);
    }
    #[cfg(not(target_os = "macos"))]
    {
        match fs::remove_file(path) {
            Ok(()) => Ok(()),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
            Err(error) => Err(format!("Unable to remove {}: {error}", path.display())),
        }
    }
}
