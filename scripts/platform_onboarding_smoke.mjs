import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  DEFAULT_ONBOARDING_STORAGE_KEY,
  WINDOWS_DESKTOP_ONBOARDING_STORAGE_KEY,
  detectDesktopPlatformFrom,
  onboardingStorageKey,
  shouldOpenOnboardingInitially,
  shouldShowMacOfficeFirstRun,
} from "../src/platform.ts";
import { tutorialSteps } from "../src/components/OnboardingTour.tsx";

assert.equal(detectDesktopPlatformFrom("MacIntel", "Mozilla/5.0 (Macintosh)"), "macos");
assert.equal(detectDesktopPlatformFrom("Win32", "Mozilla/5.0 (Windows NT 10.0)"), "windows");
assert.equal(detectDesktopPlatformFrom("Linux x86_64", "Mozilla/5.0 (X11; Linux)"), "other");

assert.equal(shouldShowMacOfficeFirstRun("macos", true, false), true);
assert.equal(shouldShowMacOfficeFirstRun("macos", true, true), false);
assert.equal(shouldShowMacOfficeFirstRun("windows", true, false), false);
assert.equal(shouldShowMacOfficeFirstRun("macos", false, false), false);
assert.equal(
  onboardingStorageKey("windows", true),
  WINDOWS_DESKTOP_ONBOARDING_STORAGE_KEY,
);
assert.equal(onboardingStorageKey("windows", false), DEFAULT_ONBOARDING_STORAGE_KEY);
assert.equal(onboardingStorageKey("macos", true), DEFAULT_ONBOARDING_STORAGE_KEY);
assert.equal(shouldOpenOnboardingInitially(false, true), false);
assert.equal(shouldOpenOnboardingInitially(false, false), true);
assert.equal(shouldOpenOnboardingInitially(true, false), false);

const macSteps = tutorialSteps("cn", "macos");
const windowsSteps = tutorialSteps("cn", "windows");
const otherSteps = tutorialSteps("cn", "other");
const macIds = macSteps.map((step) => step.id);
const windowsIds = windowsSteps.map((step) => step.id);
const otherIds = otherSteps.map((step) => step.id);

assert(macIds.includes("mac-office-enable"));
assert(macIds.includes("mac-office-manage"));
assert(!macIds.includes("windows-office-manage"));
assert(windowsIds.includes("windows-office-manage"));
assert(!windowsIds.includes("mac-office-enable"));
assert(!windowsIds.includes("mac-office-manage"));
assert(!otherIds.some((id) => id.includes("office")));
assert(macSteps.find((step) => step.id === "mac-office-enable")?.description.includes("加载项"));
assert(macSteps.find((step) => step.id === "mac-office-manage")?.description.includes("卸载"));
const windowsOfficeStep = windowsSteps.find((step) => step.id === "windows-office-manage");
assert(windowsOfficeStep?.title.includes("Word 和 PowerPoint"));
assert(windowsOfficeStep?.description.includes("OLE 或 OMML"));
assert(windowsOfficeStep?.description.includes("公式编号和插入引用"));
assert(windowsOfficeStep?.description.includes("转为原生 OLE"));
assert(!windowsOfficeStep?.description.includes("可信目录"));
assert(!windowsOfficeStep?.description.includes("伴侣服务"));
assert(!windowsOfficeStep?.description.includes("manifest"));

const appSource = await readFile("src/App.tsx", "utf8");
const firstRunSource = await readFile("src/components/MacOfficeFirstRunPrompt.tsx", "utf8");
const macSettingsSource = await readFile("src/components/MacOfficeIntegrationSettings.tsx", "utf8");
const windowsSettingsSource = await readFile("src/components/WindowsOfficeIntegrationSettings.tsx", "utf8");
const mainSource = await readFile("src-tauri/src/main.rs", "utf8");
const lifecycleSource = await readFile("src-tauri/src/office/lifecycle.rs", "utf8");
const windowsBackendSource = await readFile("src-tauri/src/office/windows_backend.rs", "utf8");
const hooksSource = await readFile("src-tauri/windows/hooks.nsh", "utf8");
const installOleSource = await readFile("scripts/install_windows_ole.ps1", "utf8");
const installVstoSource = await readFile("scripts/install_windows_vsto.ps1", "utf8");
const windowsBundleSource = await readFile("src-tauri/tauri.windows.conf.json", "utf8");
const certificateSource = await readFile("scripts/ensure_windows_office_certificate.ps1", "utf8");

assert(appSource.includes("<MacOfficeFirstRunPrompt"));
assert(appSource.includes("onboardingStorageKey("));
assert(appSource.indexOf("<MacOfficeFirstRunPrompt") < appSource.indexOf("<OnboardingTour"));
assert(firstRunSource.includes('invoke("install_office_integration")'));
assert(firstRunSource.includes("onComplete(false)"));
assert(macSettingsSource.includes('"set_office_background_start"'));
assert(windowsSettingsSource.includes('"set_office_background_start"'));
assert(mainSource.includes('#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]'));
assert(lifecycleSource.includes("pub fn set_office_background_start"));
assert(lifecycleSource.includes("background::install_launch_agent"));
assert(lifecycleSource.includes("set_background_start_enabled(enabled)"));
assert(lifecycleSource.includes("powershell_compatible_path"));
assert(lifecycleSource.includes('strip_prefix(r"\\\\?\\")'));
assert(lifecycleSource.includes("hidden_windows_command"));
assert(lifecycleSource.includes("CREATE_NO_WINDOW"));
assert(lifecycleSource.includes('"-WindowStyle",'));
assert(lifecycleSource.includes('"Hidden",'));
assert(lifecycleSource.includes('run_windows_script(&app, "install_windows_vsto.ps1", &[])'));
assert(!windowsBundleSource.includes('"../scripts/install_windows_ole.ps1"'));
assert(installVstoSource.includes("Assert-NoOfficeProcesses"));
assert(installVstoSource.includes("MSIRESTARTMANAGERCONTROL=Disable"));
assert(windowsBackendSource.includes("hidden_command"));
assert(windowsBackendSource.includes("CREATE_NO_WINDOW"));

assert(hooksSource.includes("${NSD_Check} $VisualTeXOfficeOleRadio"));
assert(hooksSource.includes('StrCpy $VisualTeXOfficeChoice "ole"'));
assert(hooksSource.includes('install_windows_ole.ps1'));
assert(hooksSource.includes("-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass"));
assert(installOleSource.includes("Ensure-VisualTeXCatalogShare"));
assert(installOleSource.includes("ConvertFrom-VisualTeXExtendedPath"));
assert(installOleSource.includes('$Path.StartsWith("\\\\?\\", [StringComparison]::OrdinalIgnoreCase)'));
assert(installOleSource.includes("$root = Split-Path -Parent $scriptRoot"));
assert(installOleSource.includes("[Text.UTF8Encoding]::new($false, $true)"));
assert(installOleSource.includes("[IO.File]::ReadAllText($Path, $utf8)"));
assert(installOleSource.includes("ensure_windows_office_certificate.ps1"));
assert(installOleSource.includes("Start-Process -FilePath $visualTeX -ArgumentList \"--office-background\""));
assert(installOleSource.includes("Install-TrustedCatalogAddinWithRetry Word"));
assert(installOleSource.includes("Install-TrustedCatalogAddinWithRetry PowerPoint"));
assert(installOleSource.includes("attempt $attempt of 2"));
assert(installOleSource.includes("Do not close the Office Add-ins window while setup is running"));
assert(installOleSource.includes("VisualTeX OLE Ribbon commands were not persisted"));
assert(installOleSource.includes("Close all $OfficeHost windows"));
assert(installOleSource.includes("$startedProcessId = $process.Id"));
assert(installOleSource.includes("Stop-Process -Force -ErrorAction SilentlyContinue"));
assert(certificateSource.includes("certutil.exe -user -f -addstore Root $certificatePath"));
assert(!certificateSource.includes("$rootStore.Add($certificate)"));

console.log("Platform onboarding, startup controls, and Windows OLE installer checks passed.");
