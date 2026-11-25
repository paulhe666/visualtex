/**
 * In-app updates where the platform has them (Windows: src-tauri/src/updater.rs). The platform module
 * "@platform/appUpdate" exports `nativeUpdater`, or null where an update goes through the release page.
 */
export type NativeUpdatePhase = "idle" | "downloading" | "verifying" | "ready" | "installing" | "failed";

export interface NativeUpdateStatus {
  phase: NativeUpdatePhase;
  /** the version the download is for */
  version: string;
  downloaded: number;
  total: number;
  error: string;
}

export interface NativeUpdateCheck {
  currentVersion: string;
  latestVersion: string;
  releaseUrl: string;
  releaseName: string;
  releaseNotes: string;
  publishedAt: string;
  updateAvailable: boolean;
  /** an installer for this build is offered (otherwise only the release page) */
  canInstall: boolean;
  size: number;
  status: NativeUpdateStatus;
}

export interface NativeUpdateOutcome {
  kind: "updated" | "unfinished";
  version: string;
}

export interface NativeInstallStart {
  started: boolean;
  /** Office programs to close first (none when the installer started) */
  runningApps: string[];
}

export interface NativeUpdater {
  check(): Promise<NativeUpdateCheck>;
  download(): Promise<NativeUpdateStatus>;
  cancel(): Promise<NativeUpdateStatus>;
  /** force: install although Office programs run (the installer then asks about them itself) */
  install(force: boolean): Promise<NativeInstallStart>;
  status(): Promise<NativeUpdateStatus>;
  /** what the last update did, once after it */
  takeOutcome(): Promise<NativeUpdateOutcome | null>;
  /** status changes; returns the function that stops listening */
  onStatus(listener: (status: NativeUpdateStatus) => void): () => void;
}

/** the check's answer when the build has no update key (updater.rs KEY_MISSING) */
export const NATIVE_UPDATE_KEY_MISSING = "UPDATE_KEY_MISSING";
