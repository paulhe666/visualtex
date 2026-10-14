import assert from "node:assert/strict";
import { resolveMacOfficeOnboarding } from "../src/components/macOfficeOnboarding.ts";

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

console.log("macOS Office onboarding decisions passed.");
