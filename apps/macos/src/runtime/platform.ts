import { isTauriRuntimeAvailable } from "../office/shared/tauriTransport";

/** The desktop app on Linux: no Office (its settings and add-ins are macOS's and Windows's). */
export function isLinuxDesktop() {
  if (typeof navigator === "undefined" || !isTauriRuntimeAvailable()) return false;
  return /Linux/i.test(navigator.userAgent) && !/Android/i.test(navigator.userAgent);
}
