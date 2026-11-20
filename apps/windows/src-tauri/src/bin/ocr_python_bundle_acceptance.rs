#[path = "../ocr_python_bundle.rs"]
mod ocr_python_bundle;

use serde_json::Value;
use std::path::PathBuf;
use std::process::Command;

fn main() {
    if let Err(error) = run() {
        eprintln!("VisualTeX bundled OCR Python acceptance failed: {error}");
        std::process::exit(1);
    }
}

fn run() -> Result<(), String> {
    let bundle_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("resources")
        .join("ocr-python")
        .join("windows-x64");
    let temporary = std::env::temp_dir()
        .join("VisualTeX 用户 空格路径")
        .join(format!(
            "ocr-python-acceptance-{}",
            uuid::Uuid::new_v4()
        ));
    std::fs::create_dir_all(&temporary)
        .map_err(|error| format!("Unable to create acceptance directory: {error}"))?;
    let runtime_root = temporary.join("ocr-runtime").join("python");
    let system_root = std::env::var_os("SystemRoot")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(r"C:\Windows"));
    let system32 = system_root.join("System32");
    let manifest =
        ocr_python_bundle::install_bundle_from_root(&bundle_root, &runtime_root)?;
    let python = runtime_root.join("python.exe");
    let app_local_openmp = runtime_root.join("vcomp140.dll");
    if !app_local_openmp.is_file() {
        return Err(format!(
            "Bundled private Python is missing the app-local Microsoft OpenMP runtime: {}",
            app_local_openmp.display()
        ));
    }

    let probe = Command::new(&python)
        .arg("-c")
        .arg(
            "import json, platform, struct, sys; print(json.dumps({'pythonVersion': platform.python_version(), 'bits': struct.calcsize('P') * 8, 'executable': sys.executable, 'prefix': sys.prefix}))",
        )
        .env_clear()
        .env("PATH", "")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .output()
        .map_err(|error| format!("Unable to start bundled private Python: {error}"))?;
    if !probe.status.success() {
        return Err(format!(
            "Bundled private Python probe failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            probe.status.code(),
            String::from_utf8_lossy(&probe.stdout),
            String::from_utf8_lossy(&probe.stderr)
        ));
    }
    let value: Value = serde_json::from_slice(&probe.stdout)
        .map_err(|error| format!("Bundled Python probe returned invalid JSON: {error}"))?;
    if value.get("pythonVersion").and_then(Value::as_str)
        != Some(manifest.python_version.as_str())
        || value.get("bits").and_then(Value::as_u64) != Some(64)
    {
        return Err(format!("Unexpected bundled Python probe: {value}"));
    }
    let executable = value
        .get("executable")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .replace('\\', "/")
        .to_ascii_lowercase();
    let expected = python
        .display()
        .to_string()
        .replace('\\', "/")
        .to_ascii_lowercase();
    if executable != expected {
        return Err(format!(
            "Bundled Python resolved the wrong executable. Expected {expected}, actual {executable}"
        ));
    }

    let offline_assets = ocr_python_bundle::offline_install_assets_from_root(&bundle_root)?;
    let offline_install = Command::new(&python)
        .args([
            "-I",
            "-X",
            "utf8",
            "-m",
            "pip",
            "install",
            "--isolated",
            "--no-index",
            "--find-links",
        ])
        .arg(&offline_assets.wheelhouse)
        .args([
            "--only-binary=:all:",
            "--require-hashes",
            "--disable-pip-version-check",
            "--no-input",
            "--requirement",
        ])
        .arg(&offline_assets.lockfile)
        .env_clear()
        .env("PATH", "")
        .env("PIP_NO_INDEX", "1")
        .env("PIP_FIND_LINKS", &offline_assets.wheelhouse)
        .env("PYTHONNOUSERSITE", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .output()
        .map_err(|error| format!("Unable to start fixed offline wheelhouse installation: {error}"))?;
    if !offline_install.status.success() {
        return Err(format!(
            "Fixed offline wheelhouse installation failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            offline_install.status.code(),
            String::from_utf8_lossy(&offline_install.stdout),
            String::from_utf8_lossy(&offline_install.stderr)
        ));
    }

    let dependency_check = Command::new(&python)
        .args(["-I", "-X", "utf8", "-m", "pip", "check"])
        .env_clear()
        .env("PATH", "")
        .env("PIP_NO_INDEX", "1")
        .env("PYTHONNOUSERSITE", "1")
        .output()
        .map_err(|error| format!("Unable to run offline dependency closure check: {error}"))?;
    if !dependency_check.status.success() {
        return Err(format!(
            "Offline dependency closure check failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            dependency_check.status.code(),
            String::from_utf8_lossy(&dependency_check.stdout),
            String::from_utf8_lossy(&dependency_check.stderr)
        ));
    }

    let isolated_profile = temporary.join("isolated-profile");
    let isolated_appdata = isolated_profile.join("AppData").join("Roaming");
    let isolated_local_appdata = isolated_profile.join("AppData").join("Local");
    std::fs::create_dir_all(&isolated_appdata)
        .and_then(|_| std::fs::create_dir_all(&isolated_local_appdata))
        .map_err(|error| format!("Unable to create isolated OCR profile: {error}"))?;

    let interface_probe = Command::new(&python)
        .args(["-I", "-X", "utf8", "-c"])
        .arg(
            r#"import ctypes, importlib.metadata as metadata, json, paddle, tokenizers
from paddleocr import FormulaRecognition
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
kernel32.GetModuleHandleW.argtypes = [ctypes.c_wchar_p]
kernel32.GetModuleHandleW.restype = ctypes.c_void_p
kernel32.GetModuleFileNameW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint32]
kernel32.GetModuleFileNameW.restype = ctypes.c_uint32
handle = kernel32.GetModuleHandleW('vcomp140.dll')
if not handle:
    raise RuntimeError('Paddle did not load vcomp140.dll')
buffer = ctypes.create_unicode_buffer(32768)
if not kernel32.GetModuleFileNameW(handle, buffer, len(buffer)):
    raise ctypes.WinError(ctypes.get_last_error())
print(json.dumps({'paddle': paddle.__version__, 'paddleocr': metadata.version('paddleocr'), 'tokenizers': metadata.version('tokenizers'), 'formulaRecognition': FormulaRecognition.__name__, 'vcompPath': buffer.value}))"#,
        )
        .env_clear()
        .env("SystemRoot", &system_root)
        .env("WINDIR", &system_root)
        .env("PATH", &system32)
        .env("USERPROFILE", &isolated_profile)
        .env("HOME", &isolated_profile)
        .env("APPDATA", &isolated_appdata)
        .env("LOCALAPPDATA", &isolated_local_appdata)
        .env("MODELSCOPE_CACHE", isolated_local_appdata.join("modelscope"))
        .env("PADDLE_PDX_CACHE_HOME", isolated_local_appdata.join("paddlex"))
        .env("VISUALTEX_OFFLINE_OCR", "1")
        .env("HF_HUB_OFFLINE", "1")
        .env("TRANSFORMERS_OFFLINE", "1")
        .env("MODELSCOPE_OFFLINE", "1")
        .env("PIP_NO_INDEX", "1")
        .env("PYTHONNOUSERSITE", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .output()
        .map_err(|error| format!("Unable to run FormulaRecognition interface probe: {error}"))?;
    if !interface_probe.status.success() {
        return Err(format!(
            "FormulaRecognition interface probe failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            interface_probe.status.code(),
            String::from_utf8_lossy(&interface_probe.stdout),
            String::from_utf8_lossy(&interface_probe.stderr)
        ));
    }
    let interface_value: Value = serde_json::from_slice(&interface_probe.stdout)
        .map_err(|error| format!("FormulaRecognition interface probe returned invalid JSON: {error}"))?;
    if interface_value.get("paddle").and_then(Value::as_str) != Some("3.3.1")
        || interface_value.get("paddleocr").and_then(Value::as_str) != Some("3.7.0")
        || interface_value.get("tokenizers").and_then(Value::as_str) != Some("0.19.1")
        || interface_value.get("formulaRecognition").and_then(Value::as_str)
            != Some("FormulaRecognition")
    {
        return Err(format!("Unexpected offline OCR interface probe: {interface_value}"));
    }
    let loaded_openmp = interface_value
        .get("vcompPath")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .replace('\\', "/")
        .to_ascii_lowercase();
    let expected_openmp = app_local_openmp
        .display()
        .to_string()
        .replace('\\', "/")
        .to_ascii_lowercase();
    if loaded_openmp != expected_openmp {
        return Err(format!(
            "Paddle loaded Microsoft OpenMP from outside the private runtime. Expected {expected_openmp}, actual {loaded_openmp}"
        ));
    }

    let fake_user_base = temporary.join("fake-user-base");
    let fake_python_path = temporary.join("hostile-python-path");
    std::fs::create_dir_all(&fake_python_path)
        .map_err(|error| format!("Unable to create hostile PYTHONPATH directory: {error}"))?;
    std::fs::write(
        fake_python_path.join("visualtex_pythonpath_pollution.py"),
        b"POLLUTED = True\n",
    )
    .map_err(|error| format!("Unable to create hostile PYTHONPATH package: {error}"))?;
    for relative in [
        PathBuf::from("Python312").join("site-packages"),
        PathBuf::from("Lib").join("site-packages"),
    ] {
        let site_packages = fake_user_base.join(relative);
        std::fs::create_dir_all(&site_packages)
            .map_err(|error| format!("Unable to create fake user site-packages: {error}"))?;
        std::fs::write(
            site_packages.join("visualtex_user_pollution.py"),
            b"POLLUTED = True\n",
        )
        .map_err(|error| format!("Unable to create fake user package: {error}"))?;
    }
    let isolation = Command::new(&python)
        .arg("-c")
        .arg(
            "import importlib.util, json, os, site, sys; user_site=site.getusersitepackages(); sites=[user_site] if isinstance(user_site,str) else list(user_site); norm=lambda value: os.path.normcase(os.path.abspath(value)); print(json.dumps({'userSiteEnabled': bool(site.ENABLE_USER_SITE), 'userSiteOnPath': any(norm(path)==norm(candidate) for path in sys.path for candidate in sites if candidate), 'pollutionVisible': importlib.util.find_spec('visualtex_user_pollution') is not None, 'sysPath': sys.path}))",
        )
        .env_clear()
        .env("PATH", "")
        .env("PYTHONUSERBASE", &fake_user_base)
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .output()
        .map_err(|error| format!("Unable to run bundled Python isolation probe: {error}"))?;
    if !isolation.status.success() {
        return Err(format!(
            "Bundled Python isolation probe failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            isolation.status.code(),
            String::from_utf8_lossy(&isolation.stdout),
            String::from_utf8_lossy(&isolation.stderr)
        ));
    }
    let isolation_value: Value = serde_json::from_slice(&isolation.stdout)
        .map_err(|error| format!("Bundled Python isolation probe returned invalid JSON: {error}"))?;
    if isolation_value
        .get("userSiteEnabled")
        .and_then(Value::as_bool)
        .unwrap_or(true)
        || isolation_value
            .get("userSiteOnPath")
            .and_then(Value::as_bool)
            .unwrap_or(true)
        || isolation_value
            .get("pollutionVisible")
            .and_then(Value::as_bool)
            .unwrap_or(true)
    {
        return Err(format!(
            "Bundled Python allowed user site-packages pollution: {isolation_value}"
        ));
    }

    let hostile_isolation = Command::new(&python)
        .args(["-I", "-X", "utf8", "-c"])
        .arg(
            "import importlib.util, json, os, sys; print(json.dumps({'userPollutionVisible': importlib.util.find_spec('visualtex_user_pollution') is not None, 'pythonPathPollutionVisible': importlib.util.find_spec('visualtex_pythonpath_pollution') is not None, 'sysPath': sys.path}))",
        )
        .env_clear()
        .env("PATH", &fake_python_path)
        .env("PYTHONPATH", &fake_python_path)
        .env("PYTHONHOME", temporary.join("invalid-python-home"))
        .env("PYTHONUSERBASE", &fake_user_base)
        .output()
        .map_err(|error| format!("Unable to run hostile-environment isolation probe: {error}"))?;
    if !hostile_isolation.status.success() {
        return Err(format!(
            "Hostile-environment isolation probe failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            hostile_isolation.status.code(),
            String::from_utf8_lossy(&hostile_isolation.stdout),
            String::from_utf8_lossy(&hostile_isolation.stderr)
        ));
    }
    let hostile_value: Value = serde_json::from_slice(&hostile_isolation.stdout)
        .map_err(|error| format!("Hostile-environment probe returned invalid JSON: {error}"))?;
    if hostile_value
        .get("userPollutionVisible")
        .and_then(Value::as_bool)
        .unwrap_or(true)
        || hostile_value
            .get("pythonPathPollutionVisible")
            .and_then(Value::as_bool)
            .unwrap_or(true)
    {
        return Err(format!(
            "Bundled Python accepted external PYTHONPATH/PYTHONHOME/PYTHONUSERBASE pollution: {hostile_value}"
        ));
    }

    let pip = Command::new(&python)
        .args(["-I", "-X", "utf8", "-m", "pip", "--version"])
        .env_clear()
        .env("PATH", &fake_python_path)
        .env("PYTHONPATH", &fake_python_path)
        .env("PYTHONHOME", temporary.join("invalid-python-home"))
        .env("PYTHONUSERBASE", &fake_user_base)
        .output()
        .map_err(|error| format!("Unable to run bundled pip: {error}"))?;
    if !pip.status.success() {
        return Err(format!(
            "Bundled pip failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            pip.status.code(),
            String::from_utf8_lossy(&pip.stdout),
            String::from_utf8_lossy(&pip.stderr)
        ));
    }
    let pip_output = String::from_utf8_lossy(&pip.stdout).replace('\\', "/");
    let expected_root = runtime_root.display().to_string().replace('\\', "/");
    if !pip_output
        .to_ascii_lowercase()
        .contains(&expected_root.to_ascii_lowercase())
    {
        return Err(format!(
            "Bundled pip resolved outside the private runtime: {pip_output}"
        ));
    }

    println!("VisualTeX bundled OCR Python acceptance passed");
    println!("python={}", manifest.python_version);
    println!("pip={}", manifest.pip_version);
    println!("runtime={}", runtime_root.display());
    println!("wheelhouse_files={}", offline_assets.manifest.wheelhouse.files.len());
    println!("offline_dependency_closure=true");
    println!("formula_recognition=true");
    println!("path_isolated=true");
    std::fs::remove_dir_all(&temporary).ok();
    Ok(())
}
