export type WindowsOfficeMode = "auto" | "vsto";

export interface OfficePlatformStatus {
  platform: string;
  mode: WindowsOfficeMode;
  activeBackend: string;
  oleBridgeHealthy: boolean;
  oleLocalServerHealthy: boolean;
  staticInstallVerified: boolean;
  wordFilesPresent: boolean;
  wordRegistryComplete: boolean;
  wordLoadEnabled: boolean;
  powerpointFilesPresent: boolean;
  powerpointRegistryComplete: boolean;
  powerpointLoadEnabled: boolean;
  vstoWordHealthy: boolean;
  vstoPowerpointHealthy: boolean;
  wordConnected: boolean;
  powerpointConnected: boolean;
  connectionVerificationAttempted: boolean;
  companionProcessRunning: boolean;
  companionPortListening: boolean;
  companionHttpsHealthy: boolean;
  companionCertificateMatches: boolean;
  companionProtocolMatches: boolean;
  officeRuntimeVerified: boolean;
  currentUserCertificateTrusted: boolean;
  backgroundStartEnabled: boolean;
  lastError: string | null;
}

export interface OfficeCompanionStatus {
  running: boolean;
  bindAddress: string;
  port: number;
  certificatePath: string;
  officeUiVersion: string;
  protocolVersion: number;
  lastError: string | null;
}

type JsonRecord = Record<string, unknown>;

function invalid(path: string, expectation: string): never {
  throw new Error(
    `VisualTeX Windows Office status returned invalid data at ${path}; expected ${expectation}.`,
  );
}

function record(value: unknown, path: string): JsonRecord {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    invalid(path, "an object");
  }
  return value as JsonRecord;
}

function stringValue(value: unknown, path: string) {
  if (typeof value !== "string") invalid(path, "a string");
  return value;
}

function booleanValue(value: unknown, path: string) {
  if (typeof value !== "boolean") invalid(path, "a boolean");
  return value;
}

function nullableString(value: unknown, path: string) {
  if (value === null) return null;
  return stringValue(value, path);
}

function positiveInteger(value: unknown, path: string) {
  if (
    typeof value !== "number" ||
    !Number.isSafeInteger(value) ||
    value < 1
  ) {
    invalid(path, "a positive integer");
  }
  return value;
}

export function decodeOfficePlatformStatus(value: unknown): OfficePlatformStatus {
  const status = record(value, "officePlatformStatus");
  stringValue(status.platform, "officePlatformStatus.platform");
  if (status.mode !== "auto" && status.mode !== "vsto") {
    invalid("officePlatformStatus.mode", '"auto" or "vsto"');
  }
  stringValue(status.activeBackend, "officePlatformStatus.activeBackend");
  const booleanKeys = [
    "oleBridgeHealthy",
    "oleLocalServerHealthy",
    "staticInstallVerified",
    "wordFilesPresent",
    "wordRegistryComplete",
    "wordLoadEnabled",
    "powerpointFilesPresent",
    "powerpointRegistryComplete",
    "powerpointLoadEnabled",
    "vstoWordHealthy",
    "vstoPowerpointHealthy",
    "wordConnected",
    "powerpointConnected",
    "connectionVerificationAttempted",
    "companionProcessRunning",
    "companionPortListening",
    "companionHttpsHealthy",
    "companionCertificateMatches",
    "companionProtocolMatches",
    "officeRuntimeVerified",
    "currentUserCertificateTrusted",
    "backgroundStartEnabled",
  ] as const;
  for (const key of booleanKeys) {
    booleanValue(status[key], `officePlatformStatus.${key}`);
  }
  nullableString(status.lastError, "officePlatformStatus.lastError");
  return status as unknown as OfficePlatformStatus;
}

export function decodeOfficeCompanionStatus(value: unknown): OfficeCompanionStatus {
  const status = record(value, "officeCompanionStatus");
  booleanValue(status.running, "officeCompanionStatus.running");
  stringValue(status.bindAddress, "officeCompanionStatus.bindAddress");
  const port = positiveInteger(status.port, "officeCompanionStatus.port");
  if (port > 65_535) invalid("officeCompanionStatus.port", "a TCP port from 1 through 65535");
  stringValue(status.certificatePath, "officeCompanionStatus.certificatePath");
  stringValue(status.officeUiVersion, "officeCompanionStatus.officeUiVersion");
  positiveInteger(status.protocolVersion, "officeCompanionStatus.protocolVersion");
  nullableString(status.lastError, "officeCompanionStatus.lastError");
  return status as unknown as OfficeCompanionStatus;
}

export function decodeBooleanStatus(value: unknown, path: string): boolean {
  return booleanValue(value, path);
}
