import { openUrl } from "@tauri-apps/plugin-opener";
import packageInfo from "../../package.json";

const LATEST_RELEASE_API =
  "https://api.github.com/repos/paulhe666/visualtex/releases/latest";
const VISUALTEX_GITHUB_HOST = "github.com";
const VISUALTEX_GITHUB_PATH = "/paulhe666/visualtex";

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
  tagName: string;
  htmlUrl: string;
  name: string;
  body: string;
  publishedAt: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function optionalString(
  source: Record<string, unknown>,
  key: string,
): string {
  const value = source[key];
  if (value === undefined || value === null) return "";
  if (typeof value !== "string") {
    throw new Error(`Invalid GitHub release field: ${key}`);
  }
  return value;
}

export function normalizeVisualTexGitHubUrl(value: string): string {
  let parsed: URL;
  try {
    parsed = new URL(value);
  } catch {
    throw new Error("Invalid VisualTeX GitHub URL");
  }
  const path = parsed.pathname.replace(/\/+$/, "").toLowerCase();
  const allowedPath =
    path === VISUALTEX_GITHUB_PATH ||
    path.startsWith(`${VISUALTEX_GITHUB_PATH}/`);
  if (
    parsed.protocol !== "https:" ||
    parsed.hostname.toLowerCase() !== VISUALTEX_GITHUB_HOST ||
    parsed.username ||
    parsed.password ||
    !allowedPath
  ) {
    throw new Error("Refused an untrusted VisualTeX release URL");
  }
  return parsed.toString();
}

export function parseStableGitHubRelease(
  value: unknown,
): GitHubReleaseResponse {
  if (!isRecord(value)) {
    throw new Error("Invalid GitHub release response");
  }
  if (
    (value.draft !== undefined && typeof value.draft !== "boolean") ||
    (value.prerelease !== undefined && typeof value.prerelease !== "boolean")
  ) {
    throw new Error("Invalid GitHub release stability fields");
  }
  if (value.draft === true || value.prerelease === true) {
    throw new Error("No stable VisualTeX release was returned");
  }

  const rawTag = optionalString(value, "tag_name").trim();
  const tagName = rawTag.replace(/^v/i, "");
  if (!/^\d+(?:\.\d+)+(?:-[0-9A-Za-z.-]+)?$/.test(tagName)) {
    throw new Error("Invalid VisualTeX release version");
  }
  const rawUrl = optionalString(value, "html_url").trim();
  if (!rawUrl) throw new Error("VisualTeX release URL is missing");

  return {
    tagName,
    htmlUrl: normalizeVisualTexGitHubUrl(rawUrl),
    name: optionalString(value, "name"),
    body: optionalString(value, "body"),
    publishedAt: optionalString(value, "published_at"),
  };
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

    const release = parseStableGitHubRelease(await response.json());
    return {
      currentVersion: CURRENT_VERSION,
      latestVersion: release.tagName,
      releaseUrl: release.htmlUrl,
      releaseName: release.name || `VisualTeX v${release.tagName}`,
      releaseNotes: release.body,
      publishedAt: release.publishedAt,
      updateAvailable: isNewerVersion(release.tagName, CURRENT_VERSION),
    };
  } finally {
    window.clearTimeout(timeout);
  }
}

export async function openReleasePage(url: string): Promise<void> {
  const trustedUrl = normalizeVisualTexGitHubUrl(url);
  if ("__TAURI_INTERNALS__" in window) {
    await openUrl(trustedUrl);
    return;
  }
  window.open(trustedUrl, "_blank", "noopener,noreferrer");
}
