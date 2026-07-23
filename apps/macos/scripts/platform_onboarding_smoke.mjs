import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  DEFAULT_ONBOARDING_STORAGE_KEY,
  MACOS_DESKTOP_ONBOARDING_STORAGE_KEY,
  detectDesktopPlatformFrom,
  onboardingStorageKey,
  shouldOpenOnboardingInitially,
  shouldShowMacOfficeFirstRun,
} from "../src/platform.ts";
import { tutorialSteps } from "../src/components/OnboardingTour.tsx";

assert.equal(detectDesktopPlatformFrom("MacIntel", "Mozilla/5.0 (Macintosh)"), "macos");
assert.equal(shouldShowMacOfficeFirstRun("macos", true, false), true);
assert.equal(shouldShowMacOfficeFirstRun("macos", true, true), false);
assert.equal(shouldShowMacOfficeFirstRun("macos", false, false), false);
assert.equal(onboardingStorageKey("macos", true), MACOS_DESKTOP_ONBOARDING_STORAGE_KEY);
assert.equal(onboardingStorageKey("macos", false), DEFAULT_ONBOARDING_STORAGE_KEY);
assert.equal(shouldOpenOnboardingInitially(false, true), false);
assert.equal(shouldOpenOnboardingInitially(false, false), true);
assert.equal(shouldOpenOnboardingInitially(true, false), false);

const macSteps = tutorialSteps("cn", "macos");
const macIds = macSteps.map((step) => step.id);
assert(macIds.includes("export"));
assert(macIds.includes("input-behavior"));
assert(macIds.includes("mac-word-plugin"));
assert(macIds.includes("mac-powerpoint-load"));
assert(macIds.includes("mac-powerpoint-use"));
assert(!macIds.includes("windows-office-manage"));
assert(macSteps.find((step) => step.id === "export")?.description.includes("Markdown"));
assert(macSteps.find((step) => step.id === "input-behavior")?.description.includes("Enter"));
assert(macSteps.find((step) => step.id === "mac-word-plugin")?.description.includes("OMML"));
assert(macSteps.find((step) => step.id === "mac-powerpoint-load")?.description.includes("PowerPoint 加载项"));
assert(macSteps.find((step) => step.id === "mac-powerpoint-use")?.description.includes("双击"));

const appSource = await readFile("src/App.tsx", "utf8");
const firstRunSource = await readFile("src/components/MacOfficeFirstRunPrompt.tsx", "utf8");
const macSettingsSource = await readFile("src/components/MacOfficeIntegrationSettings.tsx", "utf8");
const lifecycleSource = await readFile("src-tauri/src/office/lifecycle.rs", "utf8");

assert(appSource.includes("<MacOfficeFirstRunPrompt"));
assert(appSource.includes("onboardingStorageKey("));
assert(appSource.indexOf("<MacOfficeFirstRunPrompt") < appSource.indexOf("<OnboardingTour"));
assert(firstRunSource.includes('"install_macos_offline_office_addins"'));
assert(firstRunSource.includes("<PowerPointAddinGuide"));
assert(firstRunSource.includes("onComplete(false)"));
assert(macSettingsSource.includes('"install_macos_offline_office_addins"'));
assert(macSettingsSource.includes("<PowerPointAddinGuide"));
for (const obsolete of [
  "install_office_integration",
  "repair_office_integration",
  "uninstall_office_integration",
  "regenerate_office_certificate",
]) {
  assert(!firstRunSource.includes(obsolete));
  assert(!macSettingsSource.includes(obsolete));
  assert(!lifecycleSource.includes(obsolete));
}
assert(lifecycleSource.includes("pub fn set_office_background_start"));
assert(lifecycleSource.includes("background::install_launch_agent"));
assert(lifecycleSource.includes('open_office_application("Microsoft Word")'));
assert(lifecycleSource.includes('open_office_application("Microsoft PowerPoint")'));

console.log("macOS onboarding and native Office controls passed.");
