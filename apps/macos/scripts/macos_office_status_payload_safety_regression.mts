import assert from "node:assert/strict";
import { decodeMacOfflineOfficeStatus } from "../src/components/macOfficeStatusValidation.ts";

const host = {
  applicationInstalled: true,
  applicationRunning: false,
  filesPresent: true,
  filesInstalled: true,
  healthReported: true,
  loaded: true,
  pluginVersion: "1.2.5",
  installPaths: ["/Users/test/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/VisualTeX.dotm"],
  healthPath: "/Users/test/Library/Application Support/VisualTeX/health.json",
  lastError: null,
};
const status = {
  word: host,
  powerpoint: { ...host, loaded: false },
  compiledArtifactsAvailable: true,
  resourceRoot: "/Applications/VisualTeX.app/Contents/Resources",
  powerpointAddinPath: "/tmp/VisualTeX.ppam",
  wordScriptPath: "/tmp/VisualTeXWord.scpt",
  powerpointScriptPath: "/tmp/VisualTeXPowerPoint.scpt",
  tutorialPath: "/tmp/tutorial.html",
};
assert.equal(decodeMacOfflineOfficeStatus(status), status);
assert.throws(
  () =>
    decodeMacOfflineOfficeStatus({
      ...status,
      word: { ...host, applicationRunning: "false" },
    }),
  /word\.applicationRunning/,
);
assert.throws(
  () =>
    decodeMacOfflineOfficeStatus({
      ...status,
      powerpoint: { ...host, installPaths: [null] },
    }),
  /powerpoint\.installPaths\[0\]/,
);
assert.throws(
  () => decodeMacOfflineOfficeStatus({ ...status, compiledArtifactsAvailable: 1 }),
  /compiledArtifactsAvailable/,
);

console.log("VisualTeX macOS Office status payload safety regression passed");
