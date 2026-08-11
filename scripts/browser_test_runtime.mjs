import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import process from "node:process";

export function resolveBrowserTestChromePath() {
  const candidates =
    process.platform === "win32"
      ? [
          "C:/Program Files/Google/Chrome/Application/chrome.exe",
          "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
        ]
      : process.platform === "darwin"
        ? ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]
        : [
            "/usr/bin/google-chrome",
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
          ];
  const chromePath =
    process.env.VISUALTEX_CHROME_PATH ??
    candidates.find((candidate) => existsSync(candidate));
  if (!chromePath) {
    throw new Error(
      "Google Chrome was not found. Set VISUALTEX_CHROME_PATH to run browser regressions.",
    );
  }
  return chromePath;
}

export function browserTestProfilePath(name, pid = process.pid) {
  return path.join(os.tmpdir(), `${name}-${pid}`);
}
