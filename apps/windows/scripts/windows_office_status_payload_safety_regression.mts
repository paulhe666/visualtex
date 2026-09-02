import assert from "node:assert/strict";
import {
  decodeBooleanStatus,
  decodeOfficeCompanionStatus,
  decodeOfficePlatformStatus,
} from "../src/components/windowsOfficeStatusValidation.ts";

const platform = {
  platform: "windows",
  mode: "vsto",
  activeBackend: "vsto",
  oleBridgeHealthy: true,
  oleLocalServerHealthy: true,
  staticInstallVerified: true,
  wordFilesPresent: true,
  wordRegistryComplete: true,
  wordLoadEnabled: true,
  powerpointFilesPresent: true,
  powerpointRegistryComplete: true,
  powerpointLoadEnabled: true,
  vstoWordHealthy: true,
  vstoPowerpointHealthy: true,
  wordConnected: true,
  powerpointConnected: false,
  connectionVerificationAttempted: true,
  companionProcessRunning: true,
  companionPortListening: true,
  companionHttpsHealthy: true,
  companionCertificateMatches: true,
  companionProtocolMatches: true,
  officeRuntimeVerified: true,
  currentUserCertificateTrusted: true,
  backgroundStartEnabled: true,
  lastError: null,
};
assert.equal(decodeOfficePlatformStatus(platform), platform);
assert.throws(
  () => decodeOfficePlatformStatus({ ...platform, wordConnected: "yes" }),
  /wordConnected/,
);
assert.throws(
  () => decodeOfficePlatformStatus({ ...platform, mode: "legacy" }),
  /\.mode/,
);

const companion = {
  running: true,
  bindAddress: "127.0.0.1",
  port: 43127,
  certificatePath: "C:/VisualTeX/cert.pem",
  officeUiVersion: "1.2.5",
  protocolVersion: 1,
  lastError: null,
};
assert.equal(decodeOfficeCompanionStatus(companion), companion);
assert.throws(
  () => decodeOfficeCompanionStatus({ ...companion, port: 70000 }),
  /\.port/,
);
assert.throws(
  () => decodeOfficeCompanionStatus({ ...companion, protocolVersion: "1" }),
  /protocolVersion/,
);
assert.equal(decodeBooleanStatus(true, "flag"), true);
assert.throws(() => decodeBooleanStatus(1, "flag"), /flag/);

console.log("VisualTeX Windows Office status payload safety regression passed");
