// What shared code (apps/shared/visualtex-app) needs to know about the app that builds it. Each app has its own
// src/platform/ with the same module names and exports; shared code imports them as "@platform/<module>".
import packageInfo from "../../package.json";
import type { AppFeatures } from "../../../shared/visualtex-app/src/runtime/appFeatures";

/** The web edition takes the macOS app's behaviour of the shared code (the paths that run in any browser engine). */
export const APP_PLATFORM: "windows" | "macos" = "macos";
/** Shortcuts follow the visitor's computer: ⌘ on a Mac, Ctrl elsewhere. */
export const KEYBOARD_PLATFORM: "windows" | "macos" =
  typeof navigator !== "undefined" && /Mac|iPhone|iPad|iPod/i.test(navigator.platform || navigator.userAgent)
    ? "macos"
    : "windows";
/** the app's name beside its version (the help manual) */
export const APP_LABEL = "Web";
export const APP_VERSION: string = packageInfo.version;
/** A browser has none of the desktop's system parts: no Office, no system-wide input, no screen capture, no local
 *  OCR models, no updates to check, no window to shrink. */
export const APP_FEATURES: AppFeatures = {
  office: false,
  anywhere: false,
  screenOcr: false,
  localOcr: false,
  updates: false,
  keypadWindow: false,
};
