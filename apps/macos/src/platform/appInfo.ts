// What shared code (apps/shared/visualtex-app) needs to know about the app that builds it. Each app has its own
// src/platform/ with the same module names and exports; shared code imports them as "@platform/<module>".
import packageInfo from "../../package.json";
import { DESKTOP_APP_FEATURES, type AppFeatures } from "../../../shared/visualtex-app/src/runtime/appFeatures";

export const APP_PLATFORM: "windows" | "macos" = "macos";
/** whose shortcut conventions apply (Ctrl or ⌘ as the primary modifier): this system's */
export const KEYBOARD_PLATFORM: "windows" | "macos" = "macos";
/** the app's name beside its version (the help manual) */
export const APP_LABEL = "macOS";
export const APP_VERSION: string = packageInfo.version;
export const APP_FEATURES: AppFeatures = DESKTOP_APP_FEATURES;
