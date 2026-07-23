import { openUrl } from "@tauri-apps/plugin-opener";
import packageInfo from "../../package.json";

const LATEST_RELEASE_API =
  "https://api.github.com/repos/paulhe666/visualtex/releases/latest";

export const PROJECT_URL = "https://github.com/paulhe666/visualtex";
export const CURRENT_VERSION = packageInfo.version;

export interface UpdateCheckResult {
  currentVersion: string;
  latestVersion: string;
  releaseUrl: string;
  releaseName: string;
  releaseNotes: string;
  publishedAt: string;
  updateAvailable: boolean;
}

interface GitHubReleaseResponse {
  tag_name?: string;
  html_url?: string;
  name?: string;
  body?: string;
  published_at?: string;
  draft?: boolean;
  prerelease?: boolean;
}

const versionParts = (version: string) =>
  version
    .trim()
    .replace(/^v/i, "")
    .split("-")[0]
    .split(".")
    .map((part) => Number.parseInt(part, 10) || 0);

export function isNewerVersion(candidate: string, current: string): boolean {
  const next = versionParts(candidate);
  const installed = versionParts(current);
  const length = Math.max(next.length, installed.length);
  for (let index = 0; index < length; index += 1) {
    const nextPart = next[index] ?? 0;
    const installedPart = installed[index] ?? 0;
    if (nextPart !== installedPart) return nextPart > installedPart;
  }
  return false;
}

export async function checkForUpdates(): Promise<UpdateCheckResult> {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 8000);
  try {
    const response = await fetch(LATEST_RELEASE_API, {
      headers: {
        Accept: "application/vnd.github+json",
      },
      cache: "no-store",
      signal: controller.signal,
    });
    if (!response.ok) {
      throw new Error(`GitHub release request failed (${response.status})`);
    }

    const release = (await response.json()) as GitHubReleaseResponse;
    const latestVersion = release.tag_name?.replace(/^v/i, "");
    if (
      !latestVersion ||
      !release.html_url ||
      release.draft ||
      release.prerelease
    ) {
      throw new Error("No stable VisualTeX release was returned");
    }

    return {
      currentVersion: CURRENT_VERSION,
      latestVersion,
      releaseUrl: release.html_url,
      releaseName: release.name || `VisualTeX v${latestVersion}`,
      releaseNotes: release.body || "",
      publishedAt: release.published_at || "",
      updateAvailable: isNewerVersion(latestVersion, CURRENT_VERSION),
    };
  } finally {
    window.clearTimeout(timeout);
  }
}

export async function openReleasePage(url: string): Promise<void> {
  if ("__TAURI_INTERNALS__" in window) {
    await openUrl(url);
    return;
  }
  window.open(url, "_blank", "noopener,noreferrer");
}
