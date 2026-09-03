import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";

export function createBrowserProfilePath(name) {
  return join(tmpdir(), `${name}-${process.pid}`);
}

export function resolveChromiumExecutable() {
  const configuredPath = process.env.CHROME_PATH?.trim();
  if (configuredPath) return configuredPath;

  const candidates =
    process.platform === "win32"
      ? [
          process.env.PROGRAMFILES &&
            join(
              process.env.PROGRAMFILES,
              "Google",
              "Chrome",
              "Application",
              "chrome.exe",
            ),
          process.env["PROGRAMFILES(X86)"] &&
            join(
              process.env["PROGRAMFILES(X86)"],
              "Google",
              "Chrome",
              "Application",
              "chrome.exe",
            ),
          process.env.LOCALAPPDATA &&
            join(
              process.env.LOCALAPPDATA,
              "Google",
              "Chrome",
              "Application",
              "chrome.exe",
            ),
          process.env.PROGRAMFILES &&
            join(
              process.env.PROGRAMFILES,
              "Microsoft",
              "Edge",
              "Application",
              "msedge.exe",
            ),
          process.env["PROGRAMFILES(X86)"] &&
            join(
              process.env["PROGRAMFILES(X86)"],
              "Microsoft",
              "Edge",
              "Application",
              "msedge.exe",
            ),
        ]
      : process.platform === "darwin"
        ? [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
          ]
        : [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/usr/bin/microsoft-edge",
          ];

  const installedPath = candidates.find(
    (candidate) => candidate && existsSync(candidate),
  );
  if (installedPath) return installedPath;

  throw new Error(
    "Chrome, Edge, or Chromium was not found. Set CHROME_PATH to the browser executable.",
  );
}
