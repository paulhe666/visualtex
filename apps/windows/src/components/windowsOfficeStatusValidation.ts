export type WindowsOfficeMode = "auto" | "vsto";

export interface OfficePlatformStatus {
  platform: string;
  mode: WindowsOfficeMode;
  activeBackend: string;
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
  // Word draws MathType's pictures with MathType's own fonts on this PC.
  mathtypeFontsHealthy: boolean;
  mathtypeFontsProblem: string | null;
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
    "mathtypeFontsHealthy",
  ] as const;
  for (const key of booleanKeys) {
    booleanValue(status[key], `officePlatformStatus.${key}`);
  }
  nullableString(status.mathtypeFontsProblem, "officePlatformStatus.mathtypeFontsProblem");
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

export interface WpsInstallation {
  path: string;
  version: string;
  architecture: "x86" | "x64";
  perUser: boolean;
  writerRegistered: boolean;
  presentationRegistered: boolean;
  writerWhitelisted: boolean;
  presentationWhitelisted: boolean;
  formulaOleRegistered: boolean;
}

/** A host that loaded the add-in (the add-in notes it at every start). */
export interface WpsHostLoad {
  at: string;
  version: string;
  architecture: string;
  enabled: boolean;
}

export interface WpsIntegrationStatus {
  installations: WpsInstallation[];
  hostEnabled: boolean;
  writerLoaded: WpsHostLoad | null;
  presentationLoaded: WpsHostLoad | null;
  companionError: string | null;
}

export interface NativeRenderStatus {
  installed: boolean;
  running: boolean;
}

function hostLoad(value: unknown, path: string): WpsHostLoad | null {
  if (value === null) return null;
  const load = record(value, path);
  stringValue(load.at, `${path}.at`);
  stringValue(load.version, `${path}.version`);
  stringValue(load.architecture, `${path}.architecture`);
  booleanValue(load.enabled, `${path}.enabled`);
  return load as unknown as WpsHostLoad;
}

export function decodeWpsIntegrationStatus(value: unknown): WpsIntegrationStatus {
  const status = record(value, "wpsIntegrationStatus");
  if (!Array.isArray(status.installations)) {
    invalid("wpsIntegrationStatus.installations", "an array");
  }
  status.installations.forEach((entry, index) => {
    const path = `wpsIntegrationStatus.installations[${index}]`;
    const installation = record(entry, path);
    stringValue(installation.path, `${path}.path`);
    stringValue(installation.version, `${path}.version`);
    if (installation.architecture !== "x86" && installation.architecture !== "x64") {
      invalid(`${path}.architecture`, '"x86" or "x64"');
    }
    for (const key of [
      "perUser",
      "writerRegistered",
      "presentationRegistered",
      "writerWhitelisted",
      "presentationWhitelisted",
      "formulaOleRegistered",
    ] as const) {
      booleanValue(installation[key], `${path}.${key}`);
    }
  });
  booleanValue(status.hostEnabled, "wpsIntegrationStatus.hostEnabled");
  hostLoad(status.writerLoaded, "wpsIntegrationStatus.writerLoaded");
  hostLoad(status.presentationLoaded, "wpsIntegrationStatus.presentationLoaded");
  nullableString(status.companionError, "wpsIntegrationStatus.companionError");
  return status as unknown as WpsIntegrationStatus;
}

export function decodeNativeRenderStatus(value: unknown): NativeRenderStatus {
  const status = record(value, "nativeRenderStatus");
  booleanValue(status.installed, "nativeRenderStatus.installed");
  booleanValue(status.running, "nativeRenderStatus.running");
  return status as unknown as NativeRenderStatus;
}

export function decodeBooleanStatus(value: unknown, path: string): boolean {
  return booleanValue(value, path);
}
