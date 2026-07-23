import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  DEFAULT_ONBOARDING_STORAGE_KEY,
  WINDOWS_DESKTOP_ONBOARDING_STORAGE_KEY,
  detectDesktopPlatformFrom,
  onboardingStorageKey,
  shouldOpenOnboardingInitially,
} from "../src/platform.ts";
import { tutorialSteps } from "../src/components/OnboardingTour.tsx";
import {
  VISUALTEX_QQ_GROUP_NUMBER,
  VISUALTEX_QQ_GROUP_QR_DATA_URL,
} from "../src/assets/visualtexQqGroup.ts";

assert.equal(detectDesktopPlatformFrom("Win32", "Mozilla/5.0 (Windows NT 10.0)"), "windows");
assert.equal(onboardingStorageKey("windows", true), WINDOWS_DESKTOP_ONBOARDING_STORAGE_KEY);
assert.equal(onboardingStorageKey("windows", false), DEFAULT_ONBOARDING_STORAGE_KEY);
assert.equal(shouldOpenOnboardingInitially(false, false), true);
assert.equal(shouldOpenOnboardingInitially(true, false), false);

const windowsSteps = tutorialSteps("cn", "windows");
const windowsIds = windowsSteps.map((step) => step.id);
assert(windowsIds.includes("windows-office-manage"));
assert(!windowsIds.some((id) => id.startsWith("mac-")));
const windowsOfficeStep = windowsSteps.find((step) => step.id === "windows-office-manage");
assert(windowsOfficeStep?.title.includes("Word 和 PowerPoint"));
assert(windowsOfficeStep?.description.includes("OLE 或 OMML"));
assert(windowsOfficeStep?.description.includes("公式编号和插入引用"));
assert(windowsOfficeStep?.description.includes("转为原生 OLE"));

const matrixFontsStep = windowsSteps.find((step) => step.id === "matrix-fonts");
const inputBehaviorStep = windowsSteps.find((step) => step.id === "input-behavior");
const exportStep = windowsSteps.find((step) => step.id === "export");
assert(matrixFontsStep?.description.includes("10×10"));
assert(matrixFontsStep?.description.includes("黑板粗体"));
assert(inputBehaviorStep?.description.includes("上标与下标"));
assert(inputBehaviorStep?.description.includes("按 Enter 结束"));
assert(inputBehaviorStep?.description.includes("微分 d"));
assert(exportStep?.description.includes("Markdown、SVG 或 PNG"));
assert(exportStep?.description.includes("自选路径"));

assert.equal(VISUALTEX_QQ_GROUP_NUMBER, "1045801770");
assert(VISUALTEX_QQ_GROUP_QR_DATA_URL.startsWith("data:image/png;base64,"));
assert(Buffer.from(VISUALTEX_QQ_GROUP_QR_DATA_URL.split(",")[1], "base64").length > 10_000);

const updateDialogSource = await readFile("src/components/UpdateDialog.tsx", "utf8");
const stylesSource = await readFile("src/styles.css", "utf8");
const windowsSettingsSource = await readFile("src/components/WindowsOfficeIntegrationSettings.tsx", "utf8");
const mainSource = await readFile("src-tauri/src/main.rs", "utf8");
const lifecycleSource = await readFile("src-tauri/src/office/lifecycle.rs", "utf8");
const windowsBackendSource = await readFile("src-tauri/src/office/windows_backend.rs", "utf8");
const hooksSource = await readFile("src-tauri/windows/hooks.nsh", "utf8");
const installOleSource = await readFile("scripts/install_windows_ole.ps1", "utf8");
const installVstoSource = await readFile("scripts/install_windows_vsto.ps1", "utf8");
const windowsBundleSource = await readFile("src-tauri/tauri.windows.conf.json", "utf8");
const certificateSource = await readFile("scripts/ensure_windows_office_certificate.ps1", "utf8");

assert(updateDialogSource.includes("update-community-card"));
assert(updateDialogSource.includes("VISUALTEX_QQ_GROUP_QR_DATA_URL"));
assert(updateDialogSource.includes("VISUALTEX_QQ_GROUP_NUMBER"));
assert(stylesSource.includes(".update-community-qr img"));
assert(stylesSource.includes(".onboarding-input-behavior-demo"));
assert(stylesSource.includes(".onboarding-export-demo"));
assert(windowsSettingsSource.includes('"set_office_background_start"'));
assert(mainSource.includes('#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]'));
assert(lifecycleSource.includes("pub fn set_office_background_start"));
assert(lifecycleSource.includes("set_background_start_enabled(enabled)"));
assert(lifecycleSource.includes("powershell_compatible_path"));
assert(lifecycleSource.includes("hidden_windows_command"));
assert(lifecycleSource.includes("CREATE_NO_WINDOW"));
assert(lifecycleSource.includes('run_windows_script(&app, "install_windows_vsto.ps1", &[])'));
assert(!windowsBundleSource.includes('"../scripts/install_windows_ole.ps1"'));
assert(installVstoSource.includes("Assert-NoOfficeProcesses"));
assert(installVstoSource.includes("MSIRESTARTMANAGERCONTROL=Disable"));
assert(windowsBackendSource.includes("hidden_command"));
assert(windowsBackendSource.includes("CREATE_NO_WINDOW"));
assert(hooksSource.includes("${NSD_Check} $VisualTeXOfficeOleRadio"));
assert(hooksSource.includes('StrCpy $VisualTeXOfficeChoice "ole"'));
assert(hooksSource.includes('install_windows_ole.ps1'));
assert(installOleSource.includes("Ensure-VisualTeXCatalogShare"));
assert(installOleSource.includes("ensure_windows_office_certificate.ps1"));
assert(certificateSource.includes("certutil.exe -user -f -addstore Root $certificatePath"));

console.log("Windows onboarding and Office integration controls passed.");
