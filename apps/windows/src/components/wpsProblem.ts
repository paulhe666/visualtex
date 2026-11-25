import type { WpsIntegrationStatus } from "./windowsOfficeStatusValidation";

export interface WpsProblem {
  /** Repairing the Office integration (Settings > Office, or the prompt at startup) fixes it. */
  repair: boolean;
  text: string;
}

// What keeps VisualTeX out of WPS, most basic first; null when nothing does.
export function wpsProblem(wps: WpsIntegrationStatus, isEn: boolean): WpsProblem | null {
  if (wps.installations.length === 0) return null;
  if (!wps.hostEnabled) {
    return {
      repair: false,
      text: isEn
        ? "VisualTeX is turned off for WPS: WPS shows no VisualTeX tab. Turn it on, then reopen WPS."
        : "WPS 中的 VisualTeX 已关闭，WPS 不显示 VisualTeX 选项卡。开启后重新打开 WPS。",
    };
  }
  if (wps.companionError) {
    return {
      repair: true,
      text: (isEn ? "Registering the add-ins for WPS failed: " : "为 WPS 注册加载项失败：") +
        wps.companionError,
    };
  }
  const unregistered = wps.installations.find(
    (installation) =>
      !installation.writerRegistered || !installation.presentationRegistered,
  );
  if (unregistered) {
    const bits = unregistered.architecture === "x86" ? "32" : "64";
    return {
      repair: true,
      text: isEn
        ? `The VisualTeX add-ins are not registered for ${bits}-bit WPS ${unregistered.version}.`
        : `VisualTeX 加载项没有注册到 ${bits} 位 WPS ${unregistered.version}。`,
    };
  }
  if (
    wps.installations.some(
      (installation) =>
        !installation.writerWhitelisted || !installation.presentationWhitelisted,
    )
  ) {
    return {
      repair: true,
      text: isEn
        ? "VisualTeX is not in the WPS add-in whitelist."
        : "VisualTeX 不在 WPS 加载项白名单中。",
    };
  }
  if (!wps.writerLoaded && !wps.presentationLoaded) {
    return {
      repair: false,
      text: isEn
        ? "WPS has not loaded VisualTeX yet. Quit WPS completely and open it again; if there is still no VisualTeX tab, tick VisualTeX under COM Add-ins in WPS."
        : "WPS 尚未加载 VisualTeX。请完全退出 WPS 后重新打开；仍没有 VisualTeX 选项卡时，在 WPS 的“COM 加载项”中勾选 VisualTeX。",
    };
  }
  const lastLoad = [wps.writerLoaded, wps.presentationLoaded].find(Boolean);
  if (lastLoad && !lastLoad.enabled) {
    return {
      repair: false,
      text: isEn
        ? "Reopen WPS to show the VisualTeX tab."
        : "重新打开 WPS 后显示 VisualTeX 选项卡。",
    };
  }
  return null;
}

