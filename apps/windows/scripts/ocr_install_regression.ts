import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import {
  beginOcrInstallGuard,
  endOcrInstallGuard,
  isOcrInstallActive,
  ocrInstallStatusToProgress,
  shouldDisplayRuntimeError,
} from "../src/ocr/ocrInstallState";
import type { OcrInstallStatus } from "../src/ocr/ocrService";

const root = path.resolve(import.meta.dirname, "..");
const rustSource = fs.readFileSync(
  path.join(root, "src-tauri", "src", "lib.rs"),
  "utf8",
);
const processSource = fs.readFileSync(
  path.join(root, "src-tauri", "src", "ocr_install.rs"),
  "utf8",
);
const pythonBundleSource = fs.readFileSync(
  path.join(root, "src-tauri", "src", "ocr_python_bundle.rs"),
  "utf8",
);
const storageSource = fs.readFileSync(
  path.join(root, "src-tauri", "src", "ocr_storage.rs"),
  "utf8",
);
const tauriConfigSource = fs.readFileSync(
  path.join(root, "src-tauri", "tauri.ocr-bundle.conf.json"),
  "utf8",
);
const privateSiteCustomizeSource = fs.readFileSync(
  path.join(root, "src-tauri", "ocr", "private_sitecustomize.py"),
  "utf8",
);
const dialogSource = fs.readFileSync(
  path.join(root, "src", "components", "OcrDialog.tsx"),
  "utf8",
);
const officeTransportSource = fs.readFileSync(
  path.join(root, "src", "office", "api", "ocrHttpTransport.ts"),
  "utf8",
);

const guard = { current: false };
assert.equal(beginOcrInstallGuard(guard), true);
assert.equal(beginOcrInstallGuard(guard), false, "rapid second click must be blocked");
endOcrInstallGuard(guard);
assert.equal(beginOcrInstallGuard(guard), true);

assert.equal(isOcrInstallActive("installing"), true);
assert.equal(isOcrInstallActive("dependenciesInstalled"), true);
assert.equal(isOcrInstallActive("verifying"), true);
assert.equal(isOcrInstallActive("installFailed"), false);
assert.equal(
  shouldDisplayRuntimeError("old missing tokenizers error", "installing"),
  false,
  "old runtime errors must not mix into an active install",
);
assert.equal(
  shouldDisplayRuntimeError("old missing tokenizers error", "installFailed"),
  false,
  "install failures are shown in the install panel, not as stale runtime errors",
);
assert.equal(shouldDisplayRuntimeError("recognition error", "complete"), true);

const failedStatus: OcrInstallStatus = {
  schemaVersion: 1,
  state: "installFailed",
  currentStep: "tokenizers",
  completedSteps: ["venv", "paddle", "paddleocr"],
  percent: 82,
  message: "tokenizers wheel installation failed",
  detail: "Retry the current step",
  error: "ModuleNotFoundError: No module named 'tokenizers'",
  logPath: "C:\\VisualTeX\\ocr-runtime\\logs\\ocr-install.log",
  updatedAtMs: Date.now(),
};
const progress = ocrInstallStatusToProgress(failedStatus);
assert.equal(progress.stage, "tokenizers");
assert.equal(progress.percent, 82);
assert.match(progress.error ?? "", /tokenizers/);
assert.deepEqual(failedStatus.completedSteps, ["venv", "paddle", "paddleocr"]);

assert.match(rustSource, /alias = "python_version"/);
assert.match(rustSource, /alias = "paddle_version"/);
assert.match(rustSource, /alias = "paddleocr_version"/);
assert.match(rustSource, /'pythonVersion': platform\.python_version\(\)/);
assert.match(rustSource, /struct\.calcsize\('P'\) \* 8/);
assert.match(rustSource, /probe\.bits != 64/);
assert.match(rustSource, /should_replace_with_bundled_python/);
assert.match(rustSource, /ocr_python_bundle::install_bundle/);
assert.match(rustSource, /private x64 Python 3\.12 runtime/);
assert.match(pythonBundleSource, /Bundled \{label\} checksum mismatch/);
assert.match(pythonBundleSource, /enclosed_name\(\)/);
assert.match(tauriConfigSource, /ocr-python\/windows-x64/);
assert.match(rustSource, /matches!\(probe\.minor, 9 \| 10 \| 11 \| 12\)/);
assert.match(rustSource, /Unsupported VisualTeX OCR Python/);
assert.doesNotMatch(
  rustSource,
  /\.arg\("--clear"\)/,
  "resume installation must not clear a valid virtual environment",
);
assert.match(rustSource, /"tokenizers"[\s\S]*?Some\("0\.19\.1"\)/);
assert.match(rustSource, /"imagesize"[\s\S]*?"imagesize"[\s\S]*?"imagesize"/);
assert.match(rustSource, /"ftfy"[\s\S]*?"ftfy"[\s\S]*?"ftfy"/);
assert.match(rustSource, /"Wand"[\s\S]*?"wand"[\s\S]*?"Wand"/);
assert.match(rustSource, /should_run_full_runtime_probe\(force_refresh, installing\)/);
assert.match(rustSource, /validate_formula_dependency_files/);
assert.match(rustSource, /requires the precompiled Windows wheel tokenizers 0\.19\.1/);
assert.match(rustSource, /reconcile_interrupted_install_snapshot/);
assert.match(rustSource, /上一次 OCR 安装被中断/);
assert.match(rustSource, /install_control\.set_snapshot\(InstallSnapshot::new/);
assert.match(rustSource, /runtime-status\.json/);
assert.match(rustSource, /No ccache found.*不作为失败原因/);
assert.match(rustSource, /python -m pip resolved outside/);
assert.match(rustSource, /modulePath/);
assert.match(rustSource, /fn python_command\(program: &Path\) -> Command/);
assert.match(rustSource, /\.args\(\["-I", "-X", "utf8"\]\)/);
assert.match(rustSource, /PYTHONNOUSERSITE/);
assert.match(rustSource, /PYTHONSAFEPATH/);
assert.match(rustSource, /env_remove\("PYTHONPATH"\)/);
assert.match(rustSource, /env_remove\("PYTHONHOME"\)/);
assert.match(rustSource, /env_remove\("PYTHONUSERBASE"\)/);
assert.match(rustSource, /ensure_private_python_isolation/);
assert.match(rustSource, /PRIVATE_PYTHON_SITE_CUSTOMIZE/);
assert.match(rustSource, /ensure_private_dependency_closure/);
assert.match(rustSource, /\.arg\("check"\)/);
assert.match(rustSource, /isolated from user site-packages/);
assert.doesNotMatch(
  rustSource,
  /Command::new\(&paths\.python\)/,
  "all active OCR Python entry points must use the centralized isolated python_command helper",
);
assert.match(processSource, /PYTHONNOUSERSITE/);
assert.match(processSource, /PYTHONSAFEPATH/);
assert.match(privateSiteCustomizeSource, /site\.ENABLE_USER_SITE = False/);
assert.match(privateSiteCustomizeSource, /sys\.path\[:\] =/);
assert.match(privateSiteCustomizeSource, /ctypes\.WinDLL\(_vcomp140_path\)/);
assert.match(pythonBundleSource, /Lib\/site-packages\/sitecustomize\.py/);
assert.match(pythonBundleSource, /vcomp140\.dll/);
assert.match(rustSource, /ensure_private_python_app_local_runtime/);
assert.match(rustSource, /app-local Microsoft OpenMP runtime/);
assert.match(rustSource, /is_x86_feature_detected!\("avx"\)/);
assert.match(rustSource, /skipped pip reinstall/);
assert.match(rustSource, /configure_ocr_storage_location/);
assert.match(rustSource, /open_ocr_storage_location/);
assert.match(rustSource, /\.join\("VisualTeXData"\)/);
assert.match(rustSource, /storage_persistent_across_uninstall: true/);
assert.match(rustSource, /begin_storage_change/);
assert.match(rustSource, /begin_runtime_mutation/);
assert.match(rustSource, /A cached runtime status is only safe while installation/);
assert.doesNotMatch(rustSource, /if !force_refresh \|\| installing/);
assert.match(dialogSource, /runtimeRequestGenerationRef/);
assert.match(dialogSource, /applyRuntimeStatus\(nextRuntime\)/);
assert.match(storageSource, /STORAGE_DIRECTORY_NAME: &str = "VisualTeX-OCR"/);
assert.match(storageSource, /STORAGE_CONFIG_FILE: &str = "ocr-storage\.json"/);
assert.match(storageSource, /STORAGE_MARKER_FILE: &str = "\.visualtex-ocr-root\.json"/);
assert.match(storageSource, /reset_source: source_has_payload/);
assert.match(storageSource, /directory_contains_payload/);
assert.match(storageSource, /root\.join\("python"\)\.join\("python\.exe"\)\.is_file\(\)/);
assert.doesNotMatch(storageSource, /current and selected locations both contain OCR data/);
assert.match(rustSource, /reset_runtime_contents\(&current\.root\)/);
assert.doesNotMatch(storageSource, /copy_tree|source_to_remove|visualtex-ocr-migrating/);
assert.match(storageSource, /GetDiskFreeSpaceExW/);
assert.match(storageSource, /ocr-storage\.invalid-/);

assert.match(rustSource, /--only-binary=:all:/);
assert.match(rustSource, /--no-deps/);
assert.match(rustSource, /--no-compile/);
assert.match(rustSource, /--no-warn-script-location/);
assert.match(processSource, /PYTHONUTF8/);
assert.match(processSource, /PYTHONIOENCODING/);
assert.match(processSource, /MultiByteToWideChar/);
assert.match(processSource, /GetOEMCP/);
assert.match(rustSource, /PIP_DEFAULT_TIMEOUT/);
assert.match(rustSource, /PIP_CACHE_DIR/);
assert.match(rustSource, /CARGO_TARGET_DIR/);
assert.match(processSource, /no stdout\/stderr or pip cache\/temp-file growth/);
assert.match(processSource, /exceeded the installer limit/);
assert.match(processSource, /Downloading /);
assert.match(processSource, /INSTALL_ACTIVITY_POLL_INTERVAL/);
assert.match(rustSource, /--progress-bar/);
assert.match(rustSource, /"raw"\.to_string\(\)/);
assert.match(processSource, /taskkill/);
assert.match(processSource, /\/T/);
assert.match(processSource, /\/F/);

assert.match(dialogSource, /Retry current step|重试当前步骤/);
assert.match(dialogSource, /View log|查看日志/);
assert.match(dialogSource, /Reset environment|重置环境/);
assert.match(dialogSource, /Cancel installation|取消安装/);
assert.match(dialogSource, /beginOcrInstallGuard/);
assert.match(dialogSource, /shouldDisplayRuntimeError/);
assert.match(dialogSource, /await refreshInstallStatus\(\)/);
assert.match(dialogSource, /Independent OCR storage|独立 OCR 存储位置/);
assert.match(dialogSource, /Change location|更改位置/);
assert.match(dialogSource, /Preserved after uninstall|卸载后保留并自动复用/);
assert.match(dialogSource, /configureOcrStorageLocation/);
assert.match(dialogSource, /storageLowForInitialInstall/);

for (const command of [
  "get_ocr_install_status",
  "cancel_ocr_install",
  "open_ocr_install_logs",
]) {
  assert.match(officeTransportSource, new RegExp(command));
}

console.log("VisualTeX OCR installer frontend and protocol regression passed");
