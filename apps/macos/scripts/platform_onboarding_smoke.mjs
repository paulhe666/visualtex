import assert from "node:assert/strict";
import {
  macOfficeAutoInstallErrorText,
  macOfficeFolderAccessDenied,
  macOfficeFullDiskAccessSteps,
  macOfficeTextIsFolderAccessDenied,
  macPowerPointRegistrationNeeded,
  resolveMacOfficeAutoInstall,
  resolveMacOfficeOnboarding,
} from "../src/components/macOfficeOnboarding.ts";

const absentHost = {
  applicationInstalled: false,
  applicationRunning: false,
  filesPresent: false,
  filesInstalled: false,
  healthReported: false,
  loaded: false,
  pluginVersion: null,
  installPaths: [],
  healthPath: "",
  lastError: null,
};
const missingAddin = { ...absentHost, applicationInstalled: true };
const staleAddin = { ...missingAddin, filesPresent: true };
const currentAddin = { ...staleAddin, filesInstalled: true };

function status(word = absentHost, powerpoint = absentHost) {
  return {
    word,
    powerpoint,
    compiledArtifactsAvailable: true,
    resourceRoot: "",
    powerpointAddinPath: "",
    wordScriptPath: "",
    powerpointScriptPath: "",
    tutorialPath: "",
  };
}

assert.equal(resolveMacOfficeOnboarding(status(), false), null, "no Office installation needs no prompt");
assert.equal(resolveMacOfficeOnboarding(status(currentAddin, currentAddin), false), null, "current add-ins need no prompt even on a fresh profile");
assert.equal(resolveMacOfficeOnboarding(status(currentAddin), true), null, "an absent PowerPoint must not require installation");
assert.equal(resolveMacOfficeOnboarding({ ...status(missingAddin), compiledArtifactsAvailable: false }, false), null, "do not offer installation without packaged add-ins");
assert.deepEqual(resolveMacOfficeOnboarding(status(missingAddin), false), {
  mode: "setup", powerpointRegistrationRequired: false,
});
assert.deepEqual(resolveMacOfficeOnboarding(status(absentHost, missingAddin), false), {
  mode: "setup", powerpointRegistrationRequired: true,
});
assert.deepEqual(resolveMacOfficeOnboarding(status(missingAddin), true), {
  mode: "repair", powerpointRegistrationRequired: false,
});
assert.deepEqual(resolveMacOfficeOnboarding(status(currentAddin, missingAddin), false), {
  mode: "repair", powerpointRegistrationRequired: true,
});
assert.deepEqual(resolveMacOfficeOnboarding(status(staleAddin), false), {
  mode: "update", powerpointRegistrationRequired: false,
});
assert.deepEqual(resolveMacOfficeOnboarding(status(missingAddin, staleAddin), true), {
  mode: "update", powerpointRegistrationRequired: false,
});
assert.equal(resolveMacOfficeOnboarding(status({ ...currentAddin, loaded: false, healthReported: false }), true), null, "a closed Office host does not make its current files stale");

const running = (host) => ({ ...host, applicationRunning: true });
assert.equal(resolveMacOfficeAutoInstall(status(currentAddin, currentAddin), false), null, "current add-ins are left alone");
assert.equal(resolveMacOfficeAutoInstall(status(missingAddin, missingAddin), false), "install", "a new user gets the add-ins without a prompt");
assert.equal(resolveMacOfficeAutoInstall(status(staleAddin, currentAddin), false), "install", "an older add-in is replaced without a prompt");
assert.equal(resolveMacOfficeAutoInstall(status(running(staleAddin), currentAddin), false), "wait-for-office", "loaded VBA is replaced only after Office quits");
assert.equal(resolveMacOfficeAutoInstall(status(currentAddin, running(staleAddin)), false), "wait-for-office", "PowerPoint running also defers the update");
assert.equal(resolveMacOfficeAutoInstall(status(missingAddin), true), null, "no automatic install after the user uninstalled");
assert.equal(resolveMacOfficeAutoInstall({ ...status(missingAddin), compiledArtifactsAvailable: false }, false), null, "nothing to install without packaged add-ins");
assert.equal(resolveMacOfficeAutoInstall(status(absentHost, absentHost), false), null, "no Office, nothing to install");
assert.equal(macPowerPointRegistrationNeeded(status(absentHost, running(currentAddin))), true, "open PowerPoint without VisualTeX loaded is registered");
assert.equal(macPowerPointRegistrationNeeded(status(absentHost, running({ ...currentAddin, loaded: true }))), false, "a loaded add-in needs nothing");
assert.equal(macPowerPointRegistrationNeeded(status(absentHost, currentAddin)), false, "PowerPoint is never launched to register");
assert.equal(macPowerPointRegistrationNeeded(status(absentHost, running(staleAddin))), false, "only the current PPAM is registered");

assert.match(macOfficeAutoInstallErrorText(new Error("Unable to create x: Operation not permitted (os error 1)"), true), /denied access/, "permission errors are explained");
assert.match(macOfficeAutoInstallErrorText("The Office add-in operation timed out: macOS ...", true), /click Allow/, "timeouts are explained");
assert.match(macOfficeAutoInstallErrorText(new Error("disk full"), true), /disk full/, "other errors are shown as they are");

// A refused Office folder (OFFICE_FOLDER_ACCESS_DENIED in the installer).
const denied = "macOS denied VisualTeX access to the Office folders: Operation not permitted (os error 1)";
assert.equal(macOfficeFolderAccessDenied(new Error(denied)), true, "a refused Office folder is recognised");
assert.equal(macOfficeFolderAccessDenied("Operation not permitted (os error 1)"), false, "other permission errors are not a refused folder");
for (const isEn of [true, false]) {
  const text = macOfficeAutoInstallErrorText(denied, isEn);
  assert.equal(macOfficeTextIsFolderAccessDenied(text), true, "a refused folder shows the Full Disk Access steps");
  assert.doesNotMatch(text, /click Allow|点「允许」/, "never asks the user to click a prompt macOS 27 does not show");
}
const steps27 = macOfficeFullDiskAccessSteps(true, 27, false).join(" ");
assert.match(steps27, /Files & Folders/, "macOS 27 keeps Full Disk Access under Files & Folders");
assert.doesNotMatch(steps27, /earlier VisualTeX/, "a new user has no old entry to remove");
assert.match(macOfficeFullDiskAccessSteps(true, 26, false).join(" "), /> Full Disk Access\./, "earlier macOS has its own Full Disk Access list");
assert.match(macOfficeFullDiskAccessSteps(true, 27, true).join(" "), /remove it first/, "an updating user removes the earlier entry first");
assert.match(macOfficeFullDiskAccessSteps(false, 27, true).join(" "), /先删除/, "the Chinese steps include the removal too");

console.log("macOS Office onboarding decisions passed.");
