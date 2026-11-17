//! Fast open: the add-ins drop an open request into an inbox; the app claims it, answers readiness and watches the inbox.

use super::*;

pub(super) fn mark_fast_open_access_unavailable() {
    FAST_OPEN_ACCESS_BACKOFF_UNTIL_MS.store(
        epoch_ms().saturating_add(FAST_OPEN_ACCESS_BACKOFF_MS),
        Ordering::Relaxed,
    );
}

pub(super) fn fast_open_access_is_backed_off() -> bool {
    epoch_ms() < FAST_OPEN_ACCESS_BACKOFF_UNTIL_MS.load(Ordering::Relaxed)
}

pub(super) fn fast_open_inbox_roots_for_home(home: &Path) -> Vec<(OfficeHost, PathBuf)> {
    vec![
        (
            OfficeHost::Word,
            home.join(WORD_FAST_OPEN_INBOX_SUFFIX),
        ),
        (
            OfficeHost::Powerpoint,
            home.join(POWERPOINT_FAST_OPEN_INBOX_SUFFIX),
        ),
    ]
}

pub(super) fn fast_open_inbox_roots() -> Result<Vec<(OfficeHost, PathBuf)>, String> {
    Ok(fast_open_inbox_roots_for_home(&user_home()?))
}

/// The Word add-in's numbered-equation requests. The add-in writes the manifest into the fast-open inbox as
/// `<session>.number` (it cannot write into the Session folder, which is in Application Scripts); the answer is
/// `<session>.number-done` there, holding the result manifest the command line mode (`--office-omml-to-latex-batch`) writes,
/// or `error=<message>`. The helper script runs that command line in a process of its own and costs Word about half a second
/// for each call (AppleScriptTask); this answers in the resident within a poll interval.
const FAST_NUMBER_REQUEST_SUFFIX: &str = ".number";
const FAST_NUMBER_ANSWER_SUFFIX: &str = ".number-done";
const FAST_NUMBER_RESULT_FILE: &str = "omml-latex-result.txt";
const MAX_FAST_NUMBER_BYTES: u64 = 16 * 1024 * 1024;

pub(super) fn fast_number_session_id(path: &Path) -> Option<String> {
    let file_name = path.file_name()?.to_str()?;
    let session_id = file_name.strip_suffix(FAST_NUMBER_REQUEST_SUFFIX)?;
    if validate_uuid(session_id, "Fast-number Session id").is_ok() {
        Some(session_id.to_string())
    } else {
        None
    }
}

/// The result manifest of a request: the manifest goes to the Session folder as formula-restore-source.txt, where the command
/// line mode reads it.
pub(super) fn fast_number_answer(session_directory: &Path, manifest: &[u8]) -> String {
    let result = (|| -> Result<String, String> {
        fs::create_dir_all(session_directory)
            .map_err(|error| format!("Unable to create {}: {error}", session_directory.display()))?;
        let input = session_directory.join(FORMULA_RESTORE_SOURCE_FILE);
        atomic_write_runtime(&input, manifest, 0o600)?;
        let output = session_directory.join(FAST_NUMBER_RESULT_FILE);
        crate::office::omml_batch::convert_manifest(&input, &output)?;
        fs::read_to_string(&output)
            .map_err(|error| format!("Unable to read {}: {error}", output.display()))
    })();
    match result {
        Ok(text) => text,
        Err(error) => format!("error={}\n", error.replace(['\r', '\n'], " ")),
    }
}

/// Answers every number request waiting in `inbox_root` (`session_directory_of` gives a Session's folder). Returns how many.
pub(super) fn answer_fast_number_requests(
    inbox_root: &Path,
    session_directory_of: &dyn Fn(&str) -> Result<PathBuf, String>,
) -> Result<usize, String> {
    let entries = match fs::read_dir(inbox_root) {
        Ok(entries) => entries,
        Err(error) if error.kind() == std::io::ErrorKind::PermissionDenied => {
            mark_fast_open_access_unavailable();
            return Ok(0);
        }
        Err(_) => return Ok(0),
    };
    let now = SystemTime::now();
    let mut answered = 0;
    for entry in entries.flatten() {
        let path = entry.path();
        let Some(session_id) = fast_number_session_id(&path) else {
            continue;
        };
        let Ok(metadata) = fs::symlink_metadata(&path) else {
            continue;
        };
        if metadata.file_type().is_symlink() || !metadata.is_file() || metadata.len() == 0 || metadata.len() > MAX_FAST_NUMBER_BYTES {
            continue;
        }
        // The add-in publishes the request by an atomic rename, so it is complete when it appears. One left from long ago
        // (the add-in gave up on it) is not answered: the helper script's package would race it.
        let recent = match metadata.modified() {
            Ok(modified) => match now.duration_since(modified) {
                Ok(age) => age <= FAST_OPEN_MAX_AGE,
                Err(error) => error.duration() <= FAST_OPEN_FUTURE_TOLERANCE,
            },
            Err(_) => false,
        };
        let claim_path = inbox_root.join(format!(".{session_id}.{}.claim", Uuid::new_v4()));
        if fs::rename(&path, &claim_path).is_err() {
            continue;
        }
        let manifest = fs::read(&claim_path);
        let _ = fs::remove_file(&claim_path);
        if !recent {
            continue;
        }
        let answer = match (manifest, session_directory_of(&session_id)) {
            (Ok(manifest), Ok(directory)) => fast_number_answer(&directory, &manifest),
            (Err(error), _) => format!("error=Unable to read the request: {error}\n"),
            (_, Err(error)) => format!("error={}\n", error.replace(['\r', '\n'], " ")),
        };
        atomic_write_runtime(
            &inbox_root.join(format!("{session_id}{FAST_NUMBER_ANSWER_SUFFIX}")),
            answer.as_bytes(),
            0o600,
        )?;
        answered += 1;
    }
    Ok(answered)
}

pub(crate) fn consume_fast_number_requests() -> Result<usize, String> {
    if fast_open_access_is_backed_off() {
        return Ok(0);
    }
    let mut answered = 0;
    for (host, root) in fast_open_inbox_roots()? {
        if host != OfficeHost::Word {
            continue;
        }
        answered += answer_fast_number_requests(&root, &|session_id| {
            session_directory(host, session_id)
        })?;
    }
    Ok(answered)
}

pub(super) fn fast_open_session_id(path: &Path) -> Option<String> {
    let file_name = path.file_name()?.to_str()?;
    let session_id = file_name.strip_suffix(".json")?;
    if validate_uuid(session_id, "Fast-open Session id").is_ok() {
        Some(session_id.to_string())
    } else {
        None
    }
}

pub(super) fn fast_open_accept_path(inbox_root: &Path, session_id: &str) -> Result<PathBuf, String> {
    validate_uuid(session_id, "Fast-open Session id")?;
    Ok(inbox_root.join(format!("{session_id}.accepted")))
}

pub(super) fn fast_open_modified_is_recent(modified: SystemTime, now: SystemTime) -> bool {
    match now.duration_since(modified) {
        Ok(age) => (FAST_OPEN_MIN_STABLE_AGE..=FAST_OPEN_MAX_AGE).contains(&age),
        Err(error) => error.duration() <= FAST_OPEN_FUTURE_TOLERANCE,
    }
}

pub(super) fn fast_open_operation_supported(operation: &str) -> bool {
    matches!(
        operation,
        "formula"
            | "nativeToImage"
            | "imageToNative"
            | "documentImport"
            | "latexRedraw"
            | "formulaRestore"
    )
}

pub(super) fn persist_fast_open_claim(
    expected_host: OfficeHost,
    session_id: &str,
    claim_path: &Path,
) -> Result<(), String> {
    let metadata = fs::symlink_metadata(claim_path).map_err(|error| {
        format!(
            "Unable to inspect claimed Office fast-open request {}: {error}",
            claim_path.display()
        )
    })?;
    if metadata.file_type().is_symlink()
        || !metadata.is_file()
        || metadata.len() == 0
        || metadata.len() > MAX_REQUEST_BYTES
    {
        return Err("Office fast-open request has an invalid file type or size".to_string());
    }
    let bytes = fs::read(claim_path).map_err(|error| {
        format!(
            "Unable to read claimed Office fast-open request {}: {error}",
            claim_path.display()
        )
    })?;
    let request: MacOfflineSessionRequest = serde_json::from_slice(&bytes)
        .map_err(|error| format!("Office fast-open request contains invalid JSON: {error}"))?;
    validate_request(&request, session_id)?;
    if host_from_request_name(&request.host)? != expected_host {
        return Err("Office fast-open request host does not match its sandbox inbox".to_string());
    }
    let operation = request.operation.as_deref().unwrap_or("formula");
    if !fast_open_operation_supported(operation) {
        return Err("Office fast-open request uses an unsupported operation".to_string());
    }

    ensure_runtime_root(expected_host)?;
    atomic_write_runtime(&request_path(expected_host, session_id)?, &bytes, 0o600)
}

pub(crate) fn consume_fast_open_request(app: &AppHandle) -> Result<bool, String> {
    if fast_open_access_is_backed_off() {
        return Ok(false);
    }
    let now = SystemTime::now();
    let mut candidates = Vec::new();
    'host_roots: for (host, root) in fast_open_inbox_roots()? {
        let root_metadata = loop {
            match fs::symlink_metadata(&root) {
                Ok(metadata) => break metadata,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue 'host_roots,
                Err(error) if error.kind() == std::io::ErrorKind::PermissionDenied => {
                    mark_fast_open_access_unavailable();
                    continue 'host_roots;
                }
                Err(error) => {
                    return Err(format!(
                        "Unable to inspect Office fast-open inbox {}: {error}",
                        root.display()
                    ))
                }
            }
        };
        if root_metadata.file_type().is_symlink() || !root_metadata.is_dir() {
            return Err(format!(
                "Office fast-open inbox is not a real directory: {}",
                root.display()
            ));
        }
        let entries = loop {
            match fs::read_dir(&root) {
                Ok(entries) => break entries,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) if error.kind() == std::io::ErrorKind::PermissionDenied => {
                    mark_fast_open_access_unavailable();
                    continue 'host_roots;
                }
                Err(error) => {
                    return Err(format!(
                        "Unable to enumerate Office fast-open inbox {}: {error}",
                        root.display()
                    ))
                }
            }
        };
        'entries: for entry in entries {
            let entry = match entry {
                Ok(entry) => entry,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) if error.kind() == std::io::ErrorKind::PermissionDenied => {
                    mark_fast_open_access_unavailable();
                    continue 'host_roots;
                }
                Err(error) => {
                    return Err(format!(
                        "Unable to enumerate Office fast-open inbox {}: {error}",
                        root.display()
                    ))
                }
            };
            let path = entry.path();
            let Some(session_id) = fast_open_session_id(&path) else {
                continue;
            };
            let metadata = loop {
                match fs::symlink_metadata(&path) {
                    Ok(metadata) => break metadata,
                    Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                    Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue 'entries,
                    Err(error) if error.kind() == std::io::ErrorKind::PermissionDenied => {
                        mark_fast_open_access_unavailable();
                        continue 'host_roots;
                    }
                    Err(error) => {
                        return Err(format!(
                            "Unable to inspect Office fast-open request {}: {error}",
                            path.display()
                        ))
                    }
                }
            };
            if metadata.file_type().is_symlink()
                || !metadata.is_file()
                || metadata.len() == 0
                || metadata.len() > MAX_REQUEST_BYTES
            {
                continue;
            }
            let modified = match metadata.modified() {
                Ok(modified) => modified,
                Err(error) if error.kind() == std::io::ErrorKind::PermissionDenied => {
                    mark_fast_open_access_unavailable();
                    continue 'host_roots;
                }
                Err(error) => {
                    return Err(format!(
                        "Unable to read Office fast-open request timestamp {}: {error}",
                        path.display()
                    ))
                }
            };
            if !fast_open_modified_is_recent(modified, now) {
                continue;
            }
            candidates.push((modified, host, path, session_id));
        }
    }

    candidates.sort_by(|left, right| right.0.cmp(&left.0));
    for (_modified, host, path, session_id) in candidates {
        let parent = path
            .parent()
            .ok_or_else(|| "Office fast-open request has no inbox parent".to_string())?;
        let claim_path = parent.join(format!(
            ".{session_id}.{}.claim",
            Uuid::new_v4()
        ));
        match fs::rename(&path, &claim_path) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue,
            Err(error) if error.kind() == std::io::ErrorKind::PermissionDenied => {
                mark_fast_open_access_unavailable();
                continue;
            }
            Err(error) => {
                return Err(format!(
                    "Unable to claim Office fast-open request {}: {error}",
                    path.display()
                ))
            }
        }

        let persist_result = persist_fast_open_claim(host, &session_id, &claim_path);
        let _ = fs::remove_file(&claim_path);
        persist_result?;
        // The resident process owns ordinary formula activation for both Word
        // and PowerPoint. Once the sandbox request has been validated and
        // claimed there is no reason to ask LaunchServices to reopen VisualTeX;
        // doing so can raise the desktop workspace before the formula window.
        atomic_write_runtime(
            &fast_open_accept_path(parent, &session_id)?,
            b"accepted\n",
            0o600,
        )?;
        let url = format!("visualtex://office/open?session={session_id}");
        handle_open_url_safely(app, &url)?;
        return Ok(true);
    }
    Ok(false)
}

/// Runs `consume_fast_open_request` on a worker thread and waits at most
/// `wait` for it. The first read of an Office sandbox can block inside
/// `open()` until the user answers macOS's "access data from other apps"
/// prompt; the AppKit thread must never wait on that, or every later editor
/// request and the main window freeze with it. `Ok(None)` means the read is
/// still pending; the worker finishes the request on its own once macOS lets
/// it through.
pub(crate) fn consume_fast_open_request_off_main(
    app: &AppHandle,
    wait: Duration,
) -> Result<Option<bool>, String> {
    if FAST_OPEN_KICK_IN_FLIGHT.swap(true, Ordering::AcqRel) {
        return Ok(None);
    }
    let (sender, receiver) = mpsc::channel();
    let worker_app = app.clone();
    struct KickInFlight;
    impl Drop for KickInFlight {
        fn drop(&mut self) {
            FAST_OPEN_KICK_IN_FLIGHT.store(false, Ordering::Release);
        }
    }
    let spawned = crate::spawn_background_task("visualtex-office-fast-open-kick", move || {
        let in_flight = KickInFlight;
        let result = consume_fast_open_request(&worker_app);
        drop(in_flight);
        if let Err(mpsc::SendError(Err(error))) = sender.send(result) {
            eprintln!("Unable to consume VisualTeX Office fast-open request: {error}");
        }
    });
    if !spawned {
        FAST_OPEN_KICK_IN_FLIGHT.store(false, Ordering::Release);
        return Err("Unable to start the Office fast-open worker".to_string());
    }
    match receiver.recv_timeout(wait) {
        Ok(result) => result.map(Some),
        Err(_) => Ok(None),
    }
}

pub(super) fn fast_open_ready_heartbeat() -> Result<String, String> {
    let executable = std::env::current_exe()
        .map_err(|error| format!("Unable to resolve VisualTeX resident executable: {error}"))?;
    let executable = executable
        .to_str()
        .ok_or_else(|| "VisualTeX resident executable path is not valid UTF-8".to_string())?;
    if executable.contains(['\r', '\n']) {
        return Err("VisualTeX resident executable path contains a newline".to_string());
    }
    // numbers=1: this resident answers the Word add-in's number requests (<session>.number), which an older one does not:
    // the add-in asks only a resident that says so, instead of waiting for an answer that never comes.
    Ok(format!(
        "visualtex-fast-open-ready-v2\n{}\n{}\n{}\nnumbers=1\n",
        epoch_ms(),
        std::process::id(),
        executable
    ))
}

pub(super) fn refresh_fast_open_ready_markers() -> Result<(), String> {
    if fast_open_access_is_backed_off() {
        return Ok(());
    }
    let heartbeat = fast_open_ready_heartbeat()?;
    for (_host, root) in fast_open_inbox_roots()? {
        // FastOpen is an optimization only. If macOS denies direct access to an
        // Office sandbox, do not surface an error or keep hammering that path;
        // VBA will fall back to the AppleScriptTask/session URL transport.
        if fs::create_dir_all(&root).is_err() {
            mark_fast_open_access_unavailable();
            continue;
        }
        if set_mode(&root, 0o700).is_err() {
            mark_fast_open_access_unavailable();
            continue;
        }
        if atomic_write_runtime(
            &root.join(FAST_OPEN_READY_FILE),
            heartbeat.as_bytes(),
            0o600,
        )
        .is_err()
        {
            mark_fast_open_access_unavailable();
            continue;
        }
    }
    Ok(())
}

pub(crate) fn start_fast_open_inbox_watcher(app: AppHandle) {
    if FAST_OPEN_WATCHER_STARTED.set(()).is_err() {
        return;
    }
    crate::spawn_background_task("visualtex-office-fast-open", move || {
        let mut last_heartbeat = Instant::now()
            .checked_sub(FAST_OPEN_READY_HEARTBEAT_INTERVAL)
            .unwrap_or_else(Instant::now);
        loop {
            if last_heartbeat.elapsed() >= FAST_OPEN_READY_HEARTBEAT_INTERVAL {
                if let Err(error) = refresh_fast_open_ready_markers() {
                    eprintln!("Unable to refresh VisualTeX Office fast-open readiness: {error}");
                }
                last_heartbeat = Instant::now();
            }
            if let Err(error) = consume_fast_number_requests() {
                eprintln!("Unable to answer VisualTeX Word number request: {error}");
            }
            if let Err(error) = consume_fast_open_request(&app) {
                eprintln!("Unable to consume VisualTeX Office fast-open request: {error}");
                std::thread::sleep(Duration::from_millis(250));
                continue;
            }
            std::thread::sleep(FAST_OPEN_POLL_INTERVAL);
        }
    });
}
