import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = (path) => readFile(path, "utf8");
const [
  tauriConfig,
  lifecycle,
  rustMain,
  officeLifecycle,
  officeServer,
  officeSessions,
  certificateScript,
  installerHooks,
  tauriBuild,
  platformBundle,
  frontendVerifier,
  embeddedVerifier,
] = await Promise.all([
  read("src-tauri/tauri.conf.json"),
  read("src-tauri/src/app_lifecycle.rs"),
  read("src-tauri/src/lib.rs"),
  read("src-tauri/src/office/lifecycle.rs"),
  read("src-tauri/src/office/server.rs"),
  read("src-tauri/src/office/sessions.rs"),
  read("scripts/ensure_windows_office_certificate.ps1"),
  read("src-tauri/windows/hooks.nsh"),
  read("scripts/tauri_build.mjs"),
  read("scripts/build_platform_bundle.mjs"),
  read("scripts/verify_frontend_dist.mjs"),
  read("scripts/verify_embedded_frontend_assets.mjs"),
]);

const config = JSON.parse(tauriConfig);
assert.deepEqual(config.app.windows, []);
assert.equal(
  config.build.beforeBuildCommand,
  "node scripts/verify_frontend_dist.mjs",
);

assert(lifecycle.includes("pub enum AppRunMode"));
assert(lifecycle.includes("Desktop"));
assert(lifecycle.includes("OfficeBackground"));
assert(lifecycle.includes("OfficeBootstrap"));
assert(lifecycle.includes("pub fn ensure_main_window"));
assert(lifecycle.includes('WebviewUrl::App("index.html".into())'));
assert(lifecycle.includes('app.get_webview_window("main")'));
assert(lifecycle.includes("window.show()"));
assert(lifecycle.includes("window.unminimize()"));
assert(lifecycle.includes("window.set_focus()"));
assert(lifecycle.includes("app.asset_resolver()"));
assert(lifecycle.includes("Office companion resources are a separate installed resource tree"));
assert(lifecycle.includes("main-asset diagnostic"));
assert(lifecycle.includes("app-lifecycle.log"));

assert(rustMain.includes("arguments_request_desktop"));
assert(rustMain.includes("single-instance desktop activation failed"));
assert(rustMain.includes("office-background mode started companion without requesting main index.html or creating any WebView"));
assert(rustMain.includes("office-bootstrap setup skipped all windows"));
assert(rustMain.includes("office::bootstrap_configuration"));
assert(rustMain.includes("OCR startup warmup scheduled"));
assert(rustMain.includes("schedule_startup_warmup"));
assert(rustMain.includes("destroy_main_window_for_background"));
assert(rustMain.includes("state.shutdown(app)"));
assert(rustMain.includes("office::server::stop"));
assert(rustMain.includes("app.exit(0)"));

assert(!officeLifecycle.includes("prewarm_desktop_session_window"));
assert(!officeServer.includes("prewarm_desktop_session_window"));
assert(officeServer.includes("open_desktop_session_window"));
assert(officeServer.includes("WebviewWindowBuilder::new"));
assert(officeServer.includes("creating hidden Office editor WebView"));
assert(officeServer.includes(".visible(false)"));
assert(officeServer.includes(".on_page_load"));
assert(officeServer.includes("navigating hidden reused Office editor WebView"));
assert(officeServer.includes("Office editor page-load reveal fallback"));
assert(officeLifecycle.includes("office-bootstrap completed without creating a WebView"));
assert(officeLifecycle.includes("VisualTeX Office Session Cleanup"));
assert(officeLifecycle.includes("background Session cleanup started"));
assert(officeLifecycle.includes("cleanup_expired_now"));
assert(!officeSessions.includes("store.cleanup_expired(now_ms())?;"));
assert(officeSessions.includes("session_file_is_definitely_fresh"));
assert(officeSessions.includes("SESSION_CLEANUP_INTERVAL_MS"));
assert(officeSessions.includes("Lock only the individual candidate"));

assert(certificateScript.includes('ArgumentList "--office-bootstrap"'));
assert(certificateScript.includes("WaitForExit(30000)"));
assert(certificateScript.includes("left a residual process"));
assert(!certificateScript.includes('ArgumentList "--office-background"'));
assert(!installerHooks.includes("-CompanionOnly"));
assert(installerHooks.includes("without leaving a resident VisualTeX process"));
assert(installerHooks.includes("RuntimeVerificationPending"));
assert(installerHooks.includes("Function VisualTeXStopPrivateMathTypeRuntime"));
assert(installerHooks.includes("Function un.VisualTeXStopPrivateMathTypeRuntime"));
assert(installerHooks.includes("Function VisualTeXPromptPrivateMathTypeRuntimeClosure"));
assert(installerHooks.includes("$INSTDIR\\mathtype-runtime"));
const runtimeGuard = await read("scripts/manage_private_mathtype_runtime.ps1");
assert(runtimeGuard.includes("Get-CimInstance -ClassName Win32_Process"));
assert(runtimeGuard.includes("$path.StartsWith($script:runtimePrefix"));
assert(runtimeGuard.includes("$current.CreationDate -ne $target.CreationDate"));
assert(runtimeGuard.includes("PRIVATE_RUNTIME_CHECK_FAILED"));
assert(installerHooks.includes('visualtex-runtime-maintenance.ps1'));
assert(installerHooks.includes('-InstallRoot "$INSTDIR" -Mode ${MODE}'));
assert(installerHooks.includes("Never terminate a user's separately installed"));
assert(installerHooks.includes("是否关闭 VisualTeX 私有 MathType Runtime 并继续安装"));
assert(installerHooks.includes("IDYES visualtex_close_private_mathtype_now IDNO visualtex_keep_private_mathtype"));
assert.equal(
  (installerHooks.match(/^\s*Call VisualTeXPromptPrivateMathTypeRuntimeClosure\s*$/gm) ?? []).length,
  2,
);
assert(installerHooks.includes('Call ${PREFIX}VisualTeXStopPrivateMathTypeRuntime'));
assert.equal(
  (installerHooks.match(/^\s*Call un\.VisualTeXPromptPrivateMathTypeRuntimeClosure\s*$/gm) ?? []).length,
  1,
);
assert(installerHooks.includes('${If} $0 != "10"'));
assert(!installerHooks.includes('RMDir /r "$INSTDIR\\mathtype-runtime"'));
assert(!installerHooks.includes('$$targets='));

assert(tauriBuild.includes("clean_windows_release_outputs.mjs"));
assert.equal((tauriBuild.match(/build:desktop/g) ?? []).length, 2); // Windows and non-Windows branches, one each.
assert(tauriBuild.includes('["build", "--no-bundle"'));
assert(tauriBuild.includes("verify_embedded_frontend_assets.mjs"));
assert(tauriBuild.includes("preparePatchedNsisTemplate"));
assert(tauriBuild.includes("visualtex-installer.nsi"));
assert(tauriBuild.includes('run(tauri, ["bundle", "--bundles", "nsis"'));
assert(!platformBundle.includes('run(npm, ["run", "build:desktop"])'));
assert(frontendVerifier.includes("dist/index.html references a missing asset"));
assert(embeddedVerifier.includes("current dist/index.html references"));
assert(tauriBuild.includes("Same version: always remove the installed payload before reinstalling."));
assert(tauriBuild.includes("Goto reinst_uninstall"));

console.log("Windows app lifecycle, bootstrap and deterministic release build smoke passed.");
