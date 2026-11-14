// What shared code (apps/shared/visualtex-app) needs to know about the app that builds it. Each app has its own
// src/platform/ with the same module names and exports; shared code imports them as "@platform/<module>".
import packageInfo from "../../package.json";

export const APP_PLATFORM = "macos" as const;
export const APP_VERSION: string = packageInfo.version;
