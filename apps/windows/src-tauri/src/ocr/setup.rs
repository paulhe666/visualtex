//! Installing the runtime: the offline wheelhouse install step by step, package probes, the virtual environment
//! identity and resetting the runtime.

use super::*;

pub(super) fn remove_runtime_entry_once(path: &Path) -> std::io::Result<()> {
    let metadata = match fs::symlink_metadata(path) {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(()),
        Err(error) => return Err(error),
    };
    if metadata.is_dir() && !metadata.file_type().is_symlink() {
        fs::remove_dir_all(path)
    } else {
        fs::remove_file(path)
    }
}

pub(super) fn remove_runtime_entry_with_retry(path: &Path, runtime_root: &Path) -> Result<(), String> {
    let deadline = Instant::now() + Duration::from_secs(20);
    let mut attempt = 0_u32;
    let mut last_error = None;

    loop {
        attempt += 1;
        let _ = cleanup_stale_process(runtime_root);
        match remove_runtime_entry_once(path) {
            Ok(()) => return Ok(()),
            Err(error) => last_error = Some(error),
        }
        if Instant::now() >= deadline {
            let error = last_error
                .map(|error| error.to_string())
                .unwrap_or_else(|| "unknown Windows file access error".to_string());
            return Err(format!(
                "Unable to reset OCR runtime entry {} after {attempt} attempts. Another program may still be using this path. Close OCR-related Python processes or applications viewing files under the OCR runtime, then retry. Last error: {error}",
                path.display(),
            ));
        }
        let delay_ms = (attempt as u64 * 150).min(1_500);
        thread::sleep(Duration::from_millis(delay_ms));
    }
}

pub(super) fn reset_runtime_contents(runtime_root: &Path) -> Result<(), String> {
    if !runtime_root.exists() {
        return Ok(());
    }
    let entries = fs::read_dir(runtime_root)
        .map_err(|error| format!("Unable to inspect OCR runtime before reset: {error}"))?
        .collect::<Result<Vec<_>, _>>()
        .map_err(|error| format!("Unable to enumerate OCR runtime before reset: {error}"))?;

    for entry in entries {
        let name = entry.file_name().to_string_lossy().to_string();
        if name.eq_ignore_ascii_case("logs")
            || name.eq_ignore_ascii_case(storage::STORAGE_MARKER_FILE)
        {
            continue;
        }
        remove_runtime_entry_with_retry(&entry.path(), runtime_root)?;
    }
    fs::create_dir_all(runtime_root.join("logs"))
        .map_err(|error| format!("Unable to preserve OCR reset logs directory: {error}"))
}

pub(super) fn normalized_windows_path(value: impl AsRef<str>) -> String {
    value
        .as_ref()
        .replace('/', "\\")
        .trim_end_matches('\\')
        .to_ascii_lowercase()
}

pub(super) fn command_limits_for_step(step: &str) -> CommandLimits {
    match step {
        "paddle" | "paddleocr" => CommandLimits {
            step_timeout: Duration::from_secs(30 * 60),
            idle_timeout: Duration::from_secs(6 * 60),
        },
        "venv" | "pip-bootstrap" => CommandLimits {
            step_timeout: Duration::from_secs(10 * 60),
            idle_timeout: Duration::from_secs(2 * 60),
        },
        _ => CommandLimits {
            step_timeout: Duration::from_secs(15 * 60),
            idle_timeout: Duration::from_secs(4 * 60),
        },
    }
}

pub(super) fn install_command(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
    step: &str,
    command: &mut Command,
    label: &str,
) -> Result<CommandCapture, String> {
    configure_ocr_runtime_cache_environment(command, paths);
    command
        .env("PIP_DEFAULT_TIMEOUT", "30")
        .env("PIP_RETRIES", "2")
        .env("CARGO_TARGET_DIR", paths.temp.join("cargo-target"));
    run_logged_command(
        command,
        label,
        &paths.root,
        control,
        generation,
        command_limits_for_step(step),
    )
}

pub(super) fn probe_package_for_install(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
    distribution: &str,
    import_name: &str,
    expected_version: Option<&str>,
) -> Result<InstalledPackageProbe, String> {
    if import_name == "paddle" {
        ensure_paddle_cpu_compatibility()?;
        ensure_private_python_app_local_runtime(paths)?;
    }
    let script = r#"import importlib, importlib.metadata as metadata, json, platform, sys
name, import_name = sys.argv[1], sys.argv[2]
module = importlib.import_module(import_name)
print(json.dumps({
  'distribution': name,
  'version': metadata.version(name),
  'modulePath': getattr(module, '__file__', '') or '',
  'executable': sys.executable,
  'pythonVersion': platform.python_version(),
}))"#;
    let mut command = python_command(&paths.python);
    command
        .arg("-c")
        .arg(script)
        .arg(distribution)
        .arg(import_name);
    let capture = match install_command(
        paths,
        control,
        generation,
        "package-probe",
        &mut command,
        &format!("Verify Python package {distribution}"),
    ) {
        Ok(capture) => capture,
        Err(error) if import_name == "paddle" => {
            return Err(explain_paddle_import_failure(paths, &error));
        }
        Err(error) => return Err(error),
    };
    let json_line = capture
        .stdout
        .lines()
        .rev()
        .map(str::trim)
        .find(|line| line.starts_with('{'))
        .ok_or_else(|| {
            format!(
                "Package {distribution} verification returned no JSON. stdout:\n{}\nstderr:\n{}",
                capture.stdout, capture.stderr
            )
        })?;
    let probe: InstalledPackageProbe = serde_json::from_str(json_line).map_err(|error| {
        format!(
            "Package {distribution} verification returned invalid JSON: {error}\nstdout:\n{}\nstderr:\n{}",
            capture.stdout, capture.stderr
        )
    })?;
    validate_installed_package_probe(paths, distribution, expected_version, &probe)?;
    append_install_log(
        &paths.root,
        &format!(
            "verified package={distribution} version={} module={} executable={} python={}",
            probe.version, probe.module_path, probe.executable, probe.python_version
        ),
    )?;
    Ok(probe)
}

pub(super) fn validate_installed_package_probe(
    paths: &RuntimePaths,
    distribution: &str,
    expected_version: Option<&str>,
    probe: &InstalledPackageProbe,
) -> Result<(), String> {
    if probe.distribution.to_ascii_lowercase() != distribution.to_ascii_lowercase() {
        return Err(format!(
            "Package verification returned distribution '{}' while '{}' was requested",
            probe.distribution, distribution
        ));
    }
    if let Some(expected) = expected_version {
        if probe.version != expected {
            return Err(format!(
                "Package {distribution} version mismatch: installed {}, expected {expected}",
                probe.version
            ));
        }
    }
    let expected_python = normalized_windows_path(paths.python.display().to_string());
    let actual_python = normalized_windows_path(&probe.executable);
    if actual_python != expected_python {
        return Err(format!(
            "Package {distribution} was verified with the wrong interpreter. Expected {}, actual {}",
            paths.python.display(), probe.executable
        ));
    }
    let environment_root = active_python_environment_root(paths);
    let expected_root = normalized_windows_path(environment_root.display().to_string());
    let module_path = normalized_windows_path(&probe.module_path);
    if module_path.is_empty() || !module_path.starts_with(&expected_root) {
        return Err(format!(
            "Package {distribution} is not installed inside the active VisualTeX OCR Python environment {}. modulePath={}",
            environment_root.display(), probe.module_path
        ));
    }
    Ok(())
}

pub(super) fn offline_pip_install_arguments(
    assets: &python_bundle::WindowsOfflineInstallAssets,
) -> Vec<String> {
    vec![
        "-m".to_string(),
        "pip".to_string(),
        "install".to_string(),
        "--isolated".to_string(),
        "--no-index".to_string(),
        "--find-links".to_string(),
        assets.wheelhouse.display().to_string(),
        "--only-binary=:all:".to_string(),
        "--require-hashes".to_string(),
        "--no-deps".to_string(),
        "--no-compile".to_string(),
        "--no-warn-script-location".to_string(),
        "--disable-pip-version-check".to_string(),
        "--no-input".to_string(),
        "--progress-bar".to_string(),
        "raw".to_string(),
        "--requirement".to_string(),
        assets.lockfile.display().to_string(),
    ]
}

pub(super) fn install_offline_dependency_lock(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
    assets: &python_bundle::WindowsOfflineInstallAssets,
) -> Result<(), String> {
    let mut command = python_command(&paths.python);
    command
        .args(offline_pip_install_arguments(assets))
        .env("PIP_NO_INDEX", "1")
        .env("PIP_FIND_LINKS", &assets.wheelhouse)
        .env("PIP_REQUIRE_VIRTUALENV", "0")
        .env("PIP_DISABLE_PIP_VERSION_CHECK", "1");
    install_command(
        paths,
        control,
        generation,
        "offline-wheelhouse",
        &mut command,
        "Install the fixed VisualTeX OCR wheelhouse",
    )?;
    Ok(())
}

pub(super) fn pip_check_runtime(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<(), String> {
    let mut command = python_command(&paths.python);
    command
        .arg("-m")
        .arg("pip")
        .arg("check")
        .env("PIP_NO_INDEX", "1")
        .env("PIP_DISABLE_PIP_VERSION_CHECK", "1");
    install_command(
        paths,
        control,
        generation,
        "dependency-check",
        &mut command,
        "OCR private Python dependency check",
    )?;
    Ok(())
}

pub(super) fn ensure_private_dependency_closure(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<(), String> {
    pip_check_runtime(paths, control, generation).map_err(|error| {
        format!(
            "The fixed offline OCR wheelhouse did not produce a complete dependency closure. No index, system Python or user site-packages were used. {error}"
        )
    })
}

pub(super) fn verify_package_step(
    app: &AppHandle,
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
    snapshot: &mut InstallSnapshot,
    step: &str,
    percent: u8,
    message: &str,
    distribution: &str,
    import_name: &str,
    expected_version: Option<&str>,
) -> Result<(), String> {
    set_install_step(
        app,
        control,
        paths,
        snapshot,
        InstallState::Installing,
        Some(step),
        percent,
        message,
        Some(format!(
            "仅验证 {distribution} 来自 VisualTeX 私有 Python；此步骤不访问网络"
        )),
        None,
    )?;
    probe_package_for_install(
        paths,
        control,
        generation,
        distribution,
        import_name,
        expected_version,
    )?;
    snapshot.mark_step_complete(step);
    publish_install_snapshot(app, control, paths, snapshot)
}

pub(super) fn probe_runtime_for_install(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<RuntimeProbe, String> {
    ensure_paddle_cpu_compatibility()?;
    ensure_private_python_app_local_runtime(paths)?;
    let script = r#"import json, platform, sys
import paddle, paddleocr, tokenizers, imagesize, ftfy, wand
from importlib.metadata import version
from paddleocr import FormulaRecognition
print(json.dumps({
  'pythonVersion': platform.python_version(),
  'paddleVersion': paddle.__version__,
  'paddleocrVersion': version('paddleocr'),
  'executable': sys.executable,
}))"#;
    let mut command = python_command(&paths.python);
    command.arg("-c").arg(script);
    let capture = install_command(
        paths,
        control,
        generation,
        "verify",
        &mut command,
        "OCR runtime verification",
    )?;
    let probe = parse_runtime_probe_output(&capture.stdout, &capture.stderr)?;
    if probe.paddle_version != PADDLE_VERSION || probe.paddleocr_version != PADDLEOCR_VERSION {
        return Err(format!(
            "OCR runtime version mismatch: PaddlePaddle {}, PaddleOCR {}; expected {} and {}\nstdout:\n{}\nstderr:\n{}",
            probe.paddle_version,
            probe.paddleocr_version,
            PADDLE_VERSION,
            PADDLEOCR_VERSION,
            capture.stdout,
            capture.stderr
        ));
    }
    write_runtime_status_manifest(paths, &probe)?;
    Ok(probe)
}

pub(super) fn damaged_venv_python_error(error: &str) -> bool {
    let normalized = error.to_ascii_lowercase();
    !normalized.contains("cancelled")
        && !normalized.contains("timed out")
        && !normalized.contains("no stdout/stderr")
        && (normalized.contains("unable to start ocr virtual environment python check")
            || normalized.contains("ocr virtual environment python check failed with")
            || normalized.contains("returned invalid python information"))
}

pub(super) fn should_replace_with_bundled_python(error: &str) -> bool {
    let normalized = error.to_ascii_lowercase();
    damaged_venv_python_error(error)
        || normalized.contains("unsupported visualtex ocr python")
        || normalized.contains("32-bit interpreter")
        || normalized.contains("app-local microsoft openmp runtime")
}

pub(super) fn verify_venv_identity(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<PythonProbe, String> {
    ensure_private_python_isolation(paths)?;
    ensure_private_python_app_local_runtime(paths)?;
    let script = r#"import json, os, platform, site, struct, sys
user_site = site.getusersitepackages()
user_sites = [user_site] if isinstance(user_site, str) else list(user_site)
normalized = lambda value: os.path.normcase(os.path.abspath(value))
user_site_on_path = any(normalized(path) == normalized(candidate) for path in sys.path for candidate in user_sites if candidate)
if site.ENABLE_USER_SITE or user_site_on_path:
    raise RuntimeError(f'VisualTeX OCR Python is not isolated from user site-packages: enabled={site.ENABLE_USER_SITE}, onPath={user_site_on_path}, userSites={user_sites}, sys.path={sys.path}')
print(json.dumps({'version': platform.python_version(), 'major': sys.version_info.major, 'minor': sys.version_info.minor, 'bits': struct.calcsize('P') * 8, 'machine': platform.machine(), 'executable': sys.executable}))"#;
    let mut command = python_command(&paths.python);
    command.arg("-c").arg(script);
    let capture = install_command(
        paths,
        control,
        generation,
        "python",
        &mut command,
        "OCR virtual environment Python check",
    )?;
    let probe: PythonProbe = serde_json::from_str(capture.stdout.trim()).map_err(|error| {
        format!(
            "OCR virtual environment returned invalid Python information: {error}\nstdout:\n{}\nstderr:\n{}",
            capture.stdout, capture.stderr
        )
    })?;
    if probe.bits != 64 {
        return Err(format!(
            "The existing VisualTeX OCR environment uses a 32-bit interpreter: Python {} ({}-bit, platform {}). It must be replaced by the bundled x64 Python 3.12 runtime.",
            probe.version, probe.bits, probe.machine
        ));
    }
    if probe.major != 3 || !matches!(probe.minor, 9 | 10 | 11 | 12) {
        return Err(format!(
            "Unsupported VisualTeX OCR Python {} ({}-bit). It must be replaced by the bundled x64 Python 3.12 runtime.",
            probe.version, probe.bits
        ));
    }
    let expected = normalized_windows_path(paths.python.display().to_string());
    let actual = normalized_windows_path(&probe.executable);
    if expected != actual {
        return Err(format!(
            "OCR virtual environment resolved the wrong interpreter. Expected {}, actual {}",
            paths.python.display(), probe.executable
        ));
    }
    Ok(probe)
}

#[cfg(windows)]
pub(super) fn install_windows_runtime_inner(
    app: &AppHandle,
    initial_paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<OcrRuntimeStatus, String> {
    let mut effective_paths = initial_paths.clone();
    let paths = &mut effective_paths;
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    for directory in [
        &paths.input,
        &paths.processed,
        &paths.logs,
        &paths.cache,
        &paths.temp,
    ] {
        fs::create_dir_all(directory)
            .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    }
    if let Err(error) = begin_install_log_session(&paths.root) {
        eprintln!("Unable to rotate OCR installation log: {error}");
    }
    let stale_removed = cleanup_stale_process(&paths.root)?;
    cleanup_worker_temp(paths)?;
    if stale_removed {
        append_install_log(
            &paths.root,
            &format!(
                "terminated the recorded OCR installation process before starting a new task; recorded_process_removed={stale_removed}"
            ),
        )?;
    }

    let mut snapshot = load_snapshot(&paths.root)
        .unwrap_or_else(|| InstallSnapshot::new(&install_log_path(&paths.root)));
    snapshot.error = None;
    snapshot.detail = None;
    set_install_step(
        app,
        control,
        paths,
        &mut snapshot,
        InstallState::Installing,
        Some("start"),
        1,
        "正在检查可恢复的 OCR 安装环境",
        Some("不会清空已成功安装的 PaddlePaddle、PaddleOCR 或其他依赖".to_string()),
        None,
    )?;

    let installation_result = (|| -> Result<OcrRuntimeStatus, String> {
        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Installing,
            Some("cpu-check"),
            3,
            "正在检查处理器与 PaddlePaddle 的兼容性",
            Some("Windows x64 的 PaddlePaddle 3.3.1 需要 CPU 与操作系统共同提供 AVX".to_string()),
            None,
        )?;
        ensure_paddle_cpu_compatibility()?;
        append_install_log(&paths.root, "processor compatibility check passed: AVX is available")?;
        snapshot.mark_step_complete("cpu-check");

        if paths.python.is_file() {
            if let Ok(probe) = probe_runtime_for_install(paths, control, generation) {
                snapshot.completed_steps = vec![
                    "venv",
                    "pip-bootstrap",
                    "paddle",
                    "paddleocr",
                    "tokenizers",
                    "imagesize",
                    "ftfy",
                    "wand",
                    "verify",
                ]
                .into_iter()
                .map(str::to_string)
                .collect();
                set_install_step(
                    app,
                    control,
                    paths,
                    &mut snapshot,
                    InstallState::Complete,
                    Some("complete"),
                    100,
                    "已检测到完整 OCR 环境，无需重复安装",
                    Some(format!(
                        "Python {} · Paddle {} · PaddleOCR {}",
                        probe.python_version, probe.paddle_version, probe.paddleocr_version
                    )),
                    None,
                )?;
                return Ok(runtime_status_from_probe(app, paths, Ok(probe)));
            }
        }

        let existing_python_probe = if paths.python.is_file() {
            match verify_venv_identity(paths, control, generation) {
                Ok(probe) => Some(probe),
                Err(error) if should_replace_with_bundled_python(&error) => {
                    append_install_log(
                        &paths.root,
                        &format!(
                            "existing VisualTeX OCR Python is damaged and will be replaced by the bundled x64 Python 3.12 runtime: {error}"
                        ),
                    )?;
                    None
                }
                Err(error) => return Err(error),
            }
        } else {
            append_install_log(
                &paths.root,
                "no usable VisualTeX OCR Python was found; installing the bundled private x64 Python 3.12 runtime",
            )?;
            None
        };

        let python_probe = if let Some(probe) = existing_python_probe {
            probe
        } else {
            let damaged_root = active_python_environment_root(paths);
            if damaged_root.exists() {
                fs::remove_dir_all(&damaged_root).map_err(|error| {
                    format!(
                        "Unable to remove damaged OCR Python environment {}: {error}",
                        damaged_root.display()
                    )
                })?;
            }
            let legacy_python = paths.venv.join("Scripts").join("python.exe");
            if paths.venv.exists() && !legacy_python.is_file() {
                fs::remove_dir_all(&paths.venv).map_err(|error| {
                    format!("Unable to remove incomplete legacy OCR virtual environment: {error}")
                })?;
            }
            set_install_step(
                app,
                control,
                paths,
                &mut snapshot,
                InstallState::Installing,
                Some("python"),
                5,
                "正在校验内置的 64 位 Python 3.12",
                Some("该解释器独立于系统 Python，不写入 PATH，也不受系统 32 位 Python 影响".to_string()),
                None,
            )?;
            set_install_step(
                app,
                control,
                paths,
                &mut snapshot,
                InstallState::Installing,
                Some("python-runtime"),
                12,
                "正在安装 VisualTeX 私有 Python 运行时",
                Some(paths.root.join("python").display().to_string()),
                None,
            )?;
            let private_root = paths.root.join("python");
            let manifest = python_bundle::install_bundle(app, &private_root)?;
            paths.python = private_root.join("python.exe");
            append_install_log(
                &paths.root,
                &format!(
                    "activated bundled private Python {} ({}) at {}",
                    manifest.python_version,
                    manifest.architecture,
                    paths.python.display()
                ),
            )?;
            snapshot.mark_step_complete("python-runtime");
            verify_venv_identity(paths, control, generation)?
        };
        snapshot.mark_step_complete("venv");
        snapshot.mark_step_complete("python-runtime");
        publish_install_snapshot(app, control, paths, &mut snapshot)?;

        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Installing,
            Some("pip-bootstrap"),
            22,
            "正在确认 OCR 虚拟环境与 pip",
            Some(format!(
                "sys.executable={} · Python {}",
                python_probe.executable, python_probe.version
            )),
            None,
        )?;
        let mut pip_version_command = python_command(&paths.python);
        pip_version_command.arg("-m").arg("pip").arg("--version");
        let pip_capture = install_command(
            paths,
            control,
            generation,
            "pip-bootstrap",
            &mut pip_version_command,
            "OCR virtual environment pip check",
        )?;
        let environment_root = active_python_environment_root(paths);
        let expected_environment = normalized_windows_path(environment_root.display().to_string());
        if !normalized_windows_path(&pip_capture.stdout).contains(&expected_environment) {
            return Err(format!(
                "python -m pip resolved outside the active VisualTeX OCR Python environment. Expected path under {}, output: {}",
                environment_root.display(), pip_capture.stdout.trim()
            ));
        }
        append_install_log(
            &paths.root,
            &format!(
                "environment sys.executable={} python={} pip={}",
                python_probe.executable,
                python_probe.version,
                pip_capture.stdout.trim()
            ),
        )?;
        snapshot.mark_step_complete("pip-bootstrap");

        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Installing,
            Some("wheelhouse-verify"),
            26,
            "正在校验安装包内置的固定 OCR wheelhouse",
            Some("逐个校验 requirements lock、wheel 大小与 SHA-256；不会连接 PyPI".to_string()),
            None,
        )?;
        let offline_assets = python_bundle::locate_offline_install_assets(app)?;
        append_install_log(
            &paths.root,
            &format!(
                "verified bundled OCR wheelhouse files={} lock={} python={} architecture={}",
                offline_assets.manifest.wheelhouse.files.len(),
                offline_assets.lockfile.display(),
                offline_assets.manifest.python_version,
                offline_assets.manifest.architecture
            ),
        )?;
        snapshot.mark_step_complete("wheelhouse-verify");

        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Installing,
            Some("offline-wheelhouse"),
            32,
            "正在检查或安装本地 OCR 依赖闭包",
            Some("完整依赖已存在时会直接复用；否则从固定 wheelhouse 安装且不访问网络".to_string()),
            None,
        )?;
        let dependency_files_ready = probe_runtime_from_files(paths).is_ok();
        let dependency_closure_ready = if dependency_files_ready {
            match pip_check_runtime(paths, control, generation) {
                Ok(()) => true,
                Err(error) => {
                    append_install_log(
                        &paths.root,
                        &format!(
                            "existing private dependency closure cannot be reused and will be repaired from the fixed wheelhouse: {error}"
                        ),
                    )?;
                    false
                }
            }
        } else {
            false
        };
        if dependency_closure_ready {
            set_install_step(
                app,
                control,
                paths,
                &mut snapshot,
                InstallState::Installing,
                Some("offline-wheelhouse"),
                37,
                "已检测到完整的 OCR 依赖，跳过重复安装",
                Some("将直接重新验证 PaddlePaddle 原生 DLL；不会再次解压 68 个 wheel".to_string()),
                None,
            )?;
            append_install_log(
                &paths.root,
                "reused the existing fixed offline dependency closure; skipped pip reinstall",
            )?;
        } else {
            install_offline_dependency_lock(paths, control, generation, &offline_assets)?;
        }
        snapshot.mark_step_complete("offline-wheelhouse");
        publish_install_snapshot(app, control, paths, &mut snapshot)?;

        verify_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "paddle",
            38,
            &format!("正在检查或安装 PaddlePaddle {PADDLE_VERSION}"),
            "paddlepaddle",
            "paddle",
            Some(PADDLE_VERSION),
        )?;
        verify_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "paddleocr",
            66,
            &format!("正在检查或安装 PaddleOCR {PADDLEOCR_VERSION}"),
            "paddleocr",
            "paddleocr",
            Some(PADDLEOCR_VERSION),
        )?;
        verify_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "tokenizers",
            82,
            "正在安装并验证 tokenizers 预编译 wheel",
            "tokenizers",
            "tokenizers",
            Some("0.19.1"),
        )?;
        verify_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "imagesize",
            85,
            "正在安装并验证 imagesize",
            "imagesize",
            "imagesize",
            None,
        )?;
        verify_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "ftfy",
            88,
            "正在安装并验证 ftfy",
            "ftfy",
            "ftfy",
            None,
        )?;
        verify_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "wand",
            91,
            "正在安装并单独验证 Wand 导入",
            "Wand",
            "wand",
            None,
        )?;

        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Installing,
            Some("dependency-check"),
            92,
            "正在检查私有 Python 的完整依赖闭包",
            Some("pip check 仅检查刚才从本地固定 wheelhouse 安装的依赖；不会尝试联网修复".to_string()),
            None,
        )?;
        ensure_private_dependency_closure(paths, control, generation)?;
        snapshot.mark_step_complete("dependency-check");
        publish_install_snapshot(app, control, paths, &mut snapshot)?;

        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::DependenciesInstalled,
            Some("dependencies-installed"),
            93,
            "PP-FormulaNet 依赖已从固定离线 wheelhouse 安装并验证",
            Some("PaddlePaddle、PaddleOCR、tokenizers 及完整传递依赖均来自安装包本地资源".to_string()),
            None,
        )?;
        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Verifying,
            Some("verify"),
            94,
            "正在验证 PP-FormulaNet 接口",
            Some("No ccache found 等 Paddle 警告只写入日志，不作为失败原因".to_string()),
            None,
        )?;
        let probe = probe_runtime_for_install(paths, control, generation)?;
        snapshot.mark_step_complete("verify");
        let status = runtime_status_from_probe(app, paths, Ok(probe));
        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Complete,
            Some("complete"),
            100,
            "OCR 运行环境安装完成",
            Some("已写入 runtime-status.json；下次启动将直接恢复已安装状态".to_string()),
            None,
        )?;
        Ok(status)
    })();

    match installation_result {
        Ok(status) => Ok(status),
        Err(error) => {
            let error = storage::friendly_storage_error(error, &paths.root);
            let cancelled = error.to_ascii_lowercase().contains("cancelled");
            let verification = snapshot.current_step.as_deref() == Some("verify");
            let state = if cancelled {
                InstallState::Cancelled
            } else if verification {
                InstallState::VerificationFailed
            } else {
                InstallState::InstallFailed
            };
            let step = snapshot.current_step.clone();
            let failed_percent = snapshot.percent;
            set_install_step(
                app,
                control,
                paths,
                &mut snapshot,
                state,
                step.as_deref(),
                failed_percent,
                if cancelled {
                    "OCR 安装已取消"
                } else if verification {
                    "OCR 运行时验证失败"
                } else {
                    "OCR 依赖安装失败"
                },
                Some("已成功安装的环境和步骤会保留；可重试当前步骤或查看完整日志".to_string()),
                Some(error.clone()),
            )?;
            Err(error)
        }
    }
}

pub(super) fn install_runtime_inner(
    app: &AppHandle,
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
    install_control: &InstallControl,
    install_generation: u64,
) -> Result<OcrRuntimeStatus, String> {
    stop_worker(worker, worker_pid)?;
    let paths = runtime_paths(app)?;
    storage::ensure_marker(&paths.root)?;
    let required_space = if paths.python.exists() {
        512 * 1024 * 1024
    } else {
        storage::RUNTIME_INSTALL_MIN_FREE_BYTES
    };
    storage::ensure_available_space(
        &paths.root,
        required_space,
        "Installing the private Python 3.12 OCR runtime and fixed dependencies",
    )?;

    #[cfg(windows)]
    {
        return install_windows_runtime_inner(
            app,
            &paths,
            install_control,
            install_generation,
        )
        .map_err(|error| storage::friendly_storage_error(error, &paths.root));
    }

    #[cfg(not(windows))]
    {
        let _ = (install_control, install_generation);
        crate::ocr_offline::install_bundle(app, &paths.root, |stage, percent, message, detail| {
            emit_progress(app, stage, percent, message, detail);
        })?;

        emit_progress(app, "verify", 97, "正在验证离线 PP-FormulaNet 接口", None);
        let status = get_runtime_status_inner(app)?;
        if !status.installed {
            return Err(status.message);
        }
        if !status
            .installed_models
            .iter()
            .any(|model| model == crate::ocr_offline::OFFLINE_DEFAULT_MODEL)
        {
            return Err("The bundled PP-FormulaNet M model was not installed".to_string());
        }
        emit_progress(
            app,
            "complete",
            100,
            "OCR 离线运行环境安装完成",
            Some("Python、PaddleOCR 与默认 M 模型均已内置，无需联网".to_string()),
        );
        Ok(status)
    }
}
