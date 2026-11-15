//! The private Python environment: how its processes start, its isolation, and probing the installed packages.

use super::*;

pub(super) fn hide_windows_console(command: &mut Command) {
    #[cfg(windows)]
    {
        command.creation_flags(CREATE_NO_WINDOW);
    }
    #[cfg(not(windows))]
    {
        let _ = command;
    }
}

pub(super) fn configure_python_environment(command: &mut Command) {
    command
        .env_remove("PYTHONPATH")
        .env_remove("PYTHONHOME")
        .env_remove("PYTHONUSERBASE")
        .env("PYTHONNOUSERSITE", "1")
        .env("PYTHONSAFEPATH", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .env("PYTHONUNBUFFERED", "1");
}

pub(super) fn python_command(program: &Path) -> Command {
    let mut command = Command::new(program);
    command.args(["-I", "-X", "utf8"]);
    configure_python_environment(&mut command);
    command
}

pub(super) fn configure_ocr_runtime_cache_environment(command: &mut Command, paths: &RuntimePaths) {
    let huggingface = paths.cache.join("huggingface");
    command
        .env("PADDLE_PDX_CACHE_HOME", paths.cache.join("paddlex"))
        .env("PADDLE_HOME", paths.cache.join("paddle"))
        .env("MODELSCOPE_CACHE", paths.cache.join("modelscope"))
        .env("HF_HOME", &huggingface)
        .env("HF_HUB_CACHE", huggingface.join("hub"))
        .env("HUGGINGFACE_HUB_CACHE", huggingface.join("hub"))
        .env("XDG_CACHE_HOME", &paths.cache)
        .env("PIP_CACHE_DIR", paths.cache.join("pip"))
        .env("TMPDIR", &paths.temp)
        .env("TMP", &paths.temp)
        .env("TEMP", &paths.temp)
        .env("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True");
}

pub(super) fn command_output_detailed(command: &mut Command, label: &str) -> Result<(String, String), String> {
    configure_python_environment(command);
    hide_windows_console(command);
    let output = command
        .output()
        .map_err(|error| format!("Unable to start {label}: {error}"))?;
    let stdout = decode_process_output(&output.stdout);
    let stderr = decode_process_output(&output.stderr);
    if output.status.success() {
        return Ok((stdout, stderr));
    }

    let combined = format!("stdout:\n{stdout}\nstderr:\n{stderr}");
    let tail = tail_text(&combined, 6000);
    Err(format!(
        "{label} failed with {}:\n{}",
        output.status,
        tail.trim()
    ))
}

pub(super) fn read_pyvenv_python_version(paths: &RuntimePaths) -> Option<String> {
    let bundled_metadata = paths.root.join("python").join("visualtex-python.json");
    if let Ok(content) = fs::read(&bundled_metadata) {
        if let Ok(value) = serde_json::from_slice::<Value>(&content) {
            if let Some(version) = value.get("pythonVersion").and_then(Value::as_str) {
                if !version.trim().is_empty() {
                    return Some(version.trim().to_string());
                }
            }
        }
    }
    let candidates = [
        paths.venv.join("pyvenv.cfg"),
        paths.root.join("python").join("pyvenv.cfg"),
    ];
    for candidate in candidates {
        let Ok(content) = fs::read_to_string(candidate) else {
            continue;
        };
        for line in content.lines() {
            let Some((key, value)) = line.split_once('=') else {
                continue;
            };
            if matches!(key.trim(), "version" | "version_info") {
                let version = value
                    .trim()
                    .split('.')
                    .take(3)
                    .collect::<Vec<_>>()
                    .join(".");
                if !version.is_empty() {
                    return Some(version);
                }
            }
        }
    }
    None
}

pub(super) fn active_python_environment_root(paths: &RuntimePaths) -> PathBuf {
    let private_root = paths.root.join("python");
    let python = normalized_windows_path(paths.python.display().to_string());
    let private = normalized_windows_path(private_root.display().to_string());
    if python.starts_with(&private) {
        private_root
    } else {
        paths.venv.clone()
    }
}

pub(super) fn ensure_private_python_isolation(paths: &RuntimePaths) -> Result<(), String> {
    let private_root = paths.root.join("python");
    let active_root = active_python_environment_root(paths);
    if normalized_windows_path(active_root.display().to_string())
        != normalized_windows_path(private_root.display().to_string())
        || !paths.python.is_file()
    {
        return Ok(());
    }

    let site_packages = private_root.join("Lib").join("site-packages");
    fs::create_dir_all(&site_packages).map_err(|error| {
        format!(
            "Unable to create the private OCR Python site-packages directory: {error}"
        )
    })?;
    let target = site_packages.join("sitecustomize.py");
    if fs::read_to_string(&target).ok().as_deref() == Some(PRIVATE_PYTHON_SITE_CUSTOMIZE) {
        return Ok(());
    }
    let temporary = target.with_extension("py.tmp");
    fs::write(&temporary, PRIVATE_PYTHON_SITE_CUSTOMIZE.as_bytes()).map_err(|error| {
        format!("Unable to stage the private OCR Python isolation guard: {error}")
    })?;
    if target.exists() {
        fs::remove_file(&target).map_err(|error| {
            format!("Unable to replace the private OCR Python isolation guard: {error}")
        })?;
    }
    fs::rename(&temporary, &target).map_err(|error| {
        format!("Unable to activate the private OCR Python isolation guard: {error}")
    })
}

#[cfg(windows)]
pub(super) fn ensure_private_python_app_local_runtime(paths: &RuntimePaths) -> Result<(), String> {
    let private_root = paths.root.join("python");
    let active_root = active_python_environment_root(paths);
    if normalized_windows_path(active_root.display().to_string())
        != normalized_windows_path(private_root.display().to_string())
    {
        return Ok(());
    }
    let openmp = private_root.join("vcomp140.dll");
    if !openmp.is_file() {
        return Err(format!(
            "The existing VisualTeX OCR private Python is missing the app-local Microsoft OpenMP runtime {}. It must be replaced by the bundled self-contained x64 Python 3.12 runtime.",
            openmp.display()
        ));
    }
    Ok(())
}

#[cfg(not(windows))]
pub(super) fn ensure_private_python_app_local_runtime(_paths: &RuntimePaths) -> Result<(), String> {
    Ok(())
}

#[cfg(all(windows, any(target_arch = "x86", target_arch = "x86_64")))]
pub(super) fn paddle_cpu_supports_avx() -> bool {
    std::is_x86_feature_detected!("avx")
}

#[cfg(all(windows, not(any(target_arch = "x86", target_arch = "x86_64"))))]
pub(super) fn paddle_cpu_supports_avx() -> bool {
    false
}

#[cfg(not(windows))]
pub(super) fn paddle_cpu_supports_avx() -> bool {
    true
}

pub(super) fn ensure_paddle_cpu_compatibility() -> Result<(), String> {
    if paddle_cpu_supports_avx() {
        return Ok(());
    }
    Err(
        "This processor or virtual machine does not expose AVX instructions required by PaddlePaddle 3.3.1 for Windows x64. Reinstalling Python packages, copying DLL files or repairing the Visual C++ runtime cannot make this PaddlePaddle build run. Enable AVX passthrough in the virtual machine or use a computer with AVX support."
            .to_string(),
    )
}

pub(super) fn explain_paddle_import_failure(paths: &RuntimePaths, error: &str) -> String {
    if let Err(cpu_error) = ensure_paddle_cpu_compatibility() {
        return format!("{cpu_error}\nOriginal PaddlePaddle error:\n{error}");
    }
    let openmp = active_python_environment_root(paths).join("vcomp140.dll");
    if !openmp.is_file() {
        return format!(
            "PaddlePaddle could not load because the VisualTeX private Python runtime is missing its app-local Microsoft OpenMP dependency {}. Install a VisualTeX build that bundles vcomp140.dll or repair the private runtime.\nOriginal PaddlePaddle error:\n{error}",
            openmp.display()
        );
    }
    let normalized = error.to_ascii_lowercase();
    if normalized.contains("0xc0000005")
        || normalized.contains("initialization routine failed")
        || error.contains("初始化例程失败")
    {
        return format!(
            "PaddlePaddle's native Windows DLL initialization failed after VisualTeX verified AVX support and the app-local Microsoft OpenMP runtime {}. This is not a wheel download failure, and retrying will reuse the installed dependency closure instead of reinstalling all packages. Security software, virtualization settings or another injected native runtime may be blocking libpaddle.pyd.\nOriginal PaddlePaddle error:\n{error}",
            openmp.display()
        );
    }
    if normalized.contains("specified module could not be found")
        || error.contains("找不到指定的模块")
        || normalized.contains("dll load failed")
    {
        return format!(
            "PaddlePaddle's native dependency chain could not be loaded. VisualTeX found the app-local Microsoft OpenMP runtime at {}, so the remaining likely causes are a blocked/corrupted Paddle DLL or security software quarantine. The fixed installer will preserve the installed packages and report this native-load failure directly.\nOriginal PaddlePaddle error:\n{error}",
            openmp.display()
        );
    }
    error.to_string()
}

pub(super) fn site_packages_candidates(paths: &RuntimePaths) -> Vec<PathBuf> {
    let mut candidates = vec![
        paths.venv.join("Lib").join("site-packages"),
        paths.venv.join("lib").join("site-packages"),
        paths.root.join("python").join("Lib").join("site-packages"),
        paths.root.join("python").join("lib").join("site-packages"),
    ];
    let python_lib = paths.root.join("python").join("lib");
    if let Ok(entries) = fs::read_dir(python_lib) {
        for entry in entries.flatten() {
            if entry.file_type().is_ok_and(|kind| kind.is_dir()) {
                candidates.push(entry.path().join("site-packages"));
            }
        }
    }
    candidates
}

pub(super) fn dist_info_version(site_packages: &Path, distribution: &str) -> Option<String> {
    let prefix = format!("{}-", distribution.to_ascii_lowercase());
    let suffix = ".dist-info";
    for entry in fs::read_dir(site_packages).ok()?.flatten() {
        let name = entry.file_name().to_string_lossy().to_string();
        let lower = name.to_ascii_lowercase();
        if lower.starts_with(&prefix) && lower.ends_with(suffix) {
            let end = name.len().saturating_sub(suffix.len());
            return Some(name[prefix.len()..end].to_string());
        }
    }
    None
}

pub(super) fn python_import_file_exists(site_packages: &Path, import_name: &str) -> bool {
    site_packages.join(import_name).is_dir()
        || site_packages.join(format!("{import_name}.py")).is_file()
        || fs::read_dir(site_packages).is_ok_and(|entries| {
            entries.flatten().any(|entry| {
                let name = entry.file_name().to_string_lossy().to_ascii_lowercase();
                name.starts_with(&format!("{import_name}.")) && name.ends_with(".pyd")
            })
        })
}

pub(super) fn required_dependency_version(
    site_packages: &Path,
    distribution: &str,
    import_name: &str,
) -> Result<String, String> {
    if !python_import_file_exists(site_packages, import_name) {
        return Err(format!(
            "OCR dependency '{distribution}' is missing from the VisualTeX virtual environment. Resume installation from the {distribution} step."
        ));
    }
    dist_info_version(site_packages, distribution).ok_or_else(|| {
        format!(
            "OCR dependency '{distribution}' has no package metadata in the VisualTeX virtual environment. Resume installation from the {distribution} step."
        )
    })
}

pub(super) fn validate_formula_dependency_files(site_packages: &Path) -> Result<(), String> {
    let tokenizers = required_dependency_version(site_packages, "tokenizers", "tokenizers")?;
    if cfg!(windows) && tokenizers != "0.19.1" {
        return Err(format!(
            "OCR dependency tokenizers has version {tokenizers}; VisualTeX 1.2.7 requires the precompiled Windows wheel tokenizers 0.19.1."
        ));
    }
    required_dependency_version(site_packages, "imagesize", "imagesize")?;
    required_dependency_version(site_packages, "ftfy", "ftfy")?;
    required_dependency_version(site_packages, "Wand", "wand")?;
    Ok(())
}

pub(super) fn probe_runtime_from_files(paths: &RuntimePaths) -> Result<RuntimeProbe, String> {
    if !paths.python.is_file() {
        return Err("OCR Python executable is missing".to_string());
    }
    ensure_private_python_app_local_runtime(paths)?;

    let python_version = read_pyvenv_python_version(paths)
        .ok_or_else(|| "OCR Python version cache is missing".to_string())?;
    for site_packages in site_packages_candidates(paths) {
        if !site_packages.is_dir()
            || !python_import_file_exists(&site_packages, "paddle")
            || !python_import_file_exists(&site_packages, "paddleocr")
        {
            continue;
        }
        let Some(paddle_version) = dist_info_version(&site_packages, "paddlepaddle") else {
            continue;
        };
        let Some(paddleocr_version) = dist_info_version(&site_packages, "paddleocr") else {
            continue;
        };
        if paddle_version != PADDLE_VERSION || paddleocr_version != PADDLEOCR_VERSION {
            return Err(format!(
                "OCR runtime version mismatch: PaddlePaddle {paddle_version}, PaddleOCR {paddleocr_version}; expected {PADDLE_VERSION} and {PADDLEOCR_VERSION}"
            ));
        }
        validate_formula_dependency_files(&site_packages)?;
        if let Some(probe) = read_runtime_status_manifest(paths) {
            if probe.python_version == python_version {
                return Ok(probe);
            }
        }
        return Ok(RuntimeProbe {
            python_version,
            paddle_version,
            paddleocr_version,
        });
    }
    Err("OCR package metadata is missing; run a forced check or repair the runtime".to_string())
}

pub(super) fn parse_runtime_probe_output(stdout: &str, stderr: &str) -> Result<RuntimeProbe, String> {
    for line in stdout.lines().rev().map(str::trim).filter(|line| !line.is_empty()) {
        if let Ok(probe) = serde_json::from_str::<RuntimeProbe>(line) {
            return Ok(probe);
        }
    }
    Err(format!(
        "OCR runtime returned invalid version information. Expected pythonVersion/paddleVersion/paddleocrVersion (legacy snake_case is also accepted).\nRaw stdout:\n{}\nRaw stderr:\n{}",
        stdout.trim(),
        stderr.trim()
    ))
}

pub(super) fn probe_runtime(paths: &RuntimePaths) -> Result<RuntimeProbe, String> {
    if !paths.python.exists() {
        return Err("OCR runtime is not installed".to_string());
    }
    ensure_paddle_cpu_compatibility()?;
    ensure_private_python_isolation(paths)?;
    ensure_private_python_app_local_runtime(paths)?;
    let script = r#"import json, platform; import paddle; import paddleocr; import tokenizers, imagesize, ftfy, wand; from importlib.metadata import version; from paddleocr import FormulaRecognition; print(json.dumps({'pythonVersion': platform.python_version(), 'paddleVersion': paddle.__version__, 'paddleocrVersion': version('paddleocr')}))"#;
    let mut command = python_command(&paths.python);
    command.arg("-c").arg(script);
    configure_ocr_runtime_cache_environment(&mut command, paths);
    let (stdout, stderr) = command_output_detailed(
        &mut command,
        "OCR runtime verification",
    )?;
    let probe = parse_runtime_probe_output(&stdout, &stderr)?;
    if probe.paddle_version != PADDLE_VERSION || probe.paddleocr_version != PADDLEOCR_VERSION {
        return Err(format!(
            "OCR runtime version mismatch: PaddlePaddle {}, PaddleOCR {}; expected {} and {}\nRaw stdout:\n{}\nRaw stderr:\n{}",
            probe.paddle_version, probe.paddleocr_version, PADDLE_VERSION, PADDLEOCR_VERSION,
            stdout.trim(), stderr.trim()
        ));
    }
    write_runtime_status_manifest(paths, &probe)?;
    Ok(probe)
}
