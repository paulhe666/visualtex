export interface MacOfflineHostStatus {
  applicationInstalled: boolean;
  applicationRunning: boolean;
  filesPresent: boolean;
  filesInstalled: boolean;
  healthReported: boolean;
  loaded: boolean;
  pluginVersion: string | null;
  installPaths: string[];
  healthPath: string;
  lastError: string | null;
}

export interface MacOfflineOfficeStatus {
  word: MacOfflineHostStatus;
  powerpoint: MacOfflineHostStatus;
  compiledArtifactsAvailable: boolean;
  resourceRoot: string;
  powerpointAddinPath: string;
  wordScriptPath: string;
  powerpointScriptPath: string;
  tutorialPath: string;
}

type JsonRecord = Record<string, unknown>;

function invalid(path: string, expectation: string): never {
  throw new Error(
    `VisualTeX macOS Office status returned invalid data at ${path}; expected ${expectation}.`,
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

function stringArray(value: unknown, path: string) {
  if (!Array.isArray(value)) invalid(path, "an array of strings");
  value.forEach((entry, index) => stringValue(entry, `${path}[${index}]`));
  return value as string[];
}

function decodeHostStatus(value: unknown, path: string): MacOfflineHostStatus {
  const status = record(value, path);
  for (const key of [
    "applicationInstalled",
    "applicationRunning",
    "filesPresent",
    "filesInstalled",
    "healthReported",
    "loaded",
  ] as const) {
    booleanValue(status[key], `${path}.${key}`);
  }
  nullableString(status.pluginVersion, `${path}.pluginVersion`);
  stringArray(status.installPaths, `${path}.installPaths`);
  stringValue(status.healthPath, `${path}.healthPath`);
  nullableString(status.lastError, `${path}.lastError`);
  return status as unknown as MacOfflineHostStatus;
}

export function decodeMacOfflineOfficeStatus(value: unknown): MacOfflineOfficeStatus {
  const status = record(value, "macOfflineOfficeStatus");
  decodeHostStatus(status.word, "macOfflineOfficeStatus.word");
  decodeHostStatus(status.powerpoint, "macOfflineOfficeStatus.powerpoint");
  booleanValue(
    status.compiledArtifactsAvailable,
    "macOfflineOfficeStatus.compiledArtifactsAvailable",
  );
  stringValue(status.resourceRoot, "macOfflineOfficeStatus.resourceRoot");
  stringValue(status.powerpointAddinPath, "macOfflineOfficeStatus.powerpointAddinPath");
  stringValue(status.wordScriptPath, "macOfflineOfficeStatus.wordScriptPath");
  stringValue(
    status.powerpointScriptPath,
    "macOfflineOfficeStatus.powerpointScriptPath",
  );
  stringValue(status.tutorialPath, "macOfflineOfficeStatus.tutorialPath");
  return status as unknown as MacOfflineOfficeStatus;
}

/**
 * True when the status check could not finish rather than finding a problem:
 * the frontend wait expired, or the backend could not read the Office folders
 * in time (macOS may be waiting on its "access data from other apps" prompt).
 * Such a result says nothing about the add-in files, so it must not open the
 * repair prompt.
 */
export function isMacOfficeStatusUnavailable(reason: unknown): boolean {
  const message =
    reason instanceof Error ? reason.message : typeof reason === "string" ? reason : "";
  return /timed out|temporarily unavailable/i.test(message);
}

export function macOfficeStatusUnavailableText(isEn: boolean): string {
  return isEn
    ? "VisualTeX could not read the Office folders in time. If macOS asks whether VisualTeX may access data from other apps, click Allow, then refresh."
    : "暂时读不到 Office 文件夹。如果 macOS 弹出「VisualTeX 想访问其他 App 的数据」，请点「允许」，然后刷新状态。";
}
