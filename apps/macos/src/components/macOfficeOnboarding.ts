import type { MacOfflineOfficeStatus } from "./macOfficeStatusValidation";

export function resolveMacOfficeOnboarding(
  status: MacOfflineOfficeStatus,
  previouslyConfigured: boolean,
): { mode: "setup" | "update" | "repair"; powerpointRegistrationRequired: boolean } | null {
  if (!status.compiledArtifactsAvailable) return null;

  const hosts = [status.word, status.powerpoint];
  const outdatedHosts = hosts.filter(
    (host) => host.applicationInstalled && !host.filesInstalled,
  );
  if (outdatedHosts.length === 0) return null;

  const mode = outdatedHosts.some((host) => host.filesPresent)
    ? "update"
    : previouslyConfigured || hosts.some((host) => host.filesPresent)
      ? "repair"
      : "setup";

  return {
    mode,
    powerpointRegistrationRequired:
      status.powerpoint.applicationInstalled && !status.powerpoint.filesPresent,
  };
}

/** Set by Uninstall in Settings; Install or Repair clears it. */
export const MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY =
  "visualtex.office.macos.auto-install.disabled";

/**
 * What VisualTeX does on its own with the Office add-ins: install or update
 * them now, wait until Word and PowerPoint have quit (VBA already loaded in
 * Office cannot be replaced), or nothing.
 */
export function resolveMacOfficeAutoInstall(
  status: MacOfflineOfficeStatus,
  autoInstallDisabled: boolean,
): "install" | "wait-for-office" | null {
  if (autoInstallDisabled || !status.compiledArtifactsAvailable) return null;
  const needed = [status.word, status.powerpoint].some(
    (host) => host.applicationInstalled && !host.filesInstalled,
  );
  if (!needed) return null;
  return status.word.applicationRunning || status.powerpoint.applicationRunning
    ? "wait-for-office"
    : "install";
}

/** PowerPoint is open with the current PPAM installed but not loaded. */
export function macPowerPointRegistrationNeeded(status: MacOfflineOfficeStatus): boolean {
  return (
    status.powerpoint.applicationInstalled &&
    status.powerpoint.applicationRunning &&
    status.powerpoint.filesInstalled &&
    !status.powerpoint.loaded
  );
}

/**
 * "pending" once VisualTeX has registered the PPAM in PowerPoint, "confirmed"
 * once it has seen VisualTeX loaded there. While pending, a registered but
 * switched-off entry is VisualTeX's unfinished load (PowerPoint showed a
 * prompt or a dialog) and is loaded again; once confirmed, it was switched off
 * by the user and stays off.
 */
export const MAC_POWERPOINT_REGISTRATION_STORAGE_KEY =
  "visualtex.office.macos.powerpoint-registration";
export const MAC_POWERPOINT_REGISTRATION_RETRIES = 2;

/** Why VisualTeX's automatic install failed, in the words the prompt shows. */
export function macOfficeAutoInstallErrorText(reason: unknown, isEn: boolean): string {
  const raw =
    reason instanceof Error ? reason.message : typeof reason === "string" ? reason : "";
  if (/add-in operation timed out/i.test(raw)) {
    return isEn
      ? "VisualTeX could not install its Office add-ins automatically: macOS did not let it change the Office folders in time. If macOS asks whether VisualTeX may access data from other apps, click Allow, then repair."
      : "VisualTeX 未能自动安装 Office 插件：macOS 未及时允许修改 Office 文件夹。如果 macOS 询问「VisualTeX 想访问其他 App 的数据」，请点「允许」，然后修复。";
  }
  if (/Operation not permitted|Permission denied|EPERM/i.test(raw)) {
    return isEn
      ? "VisualTeX could not install its Office add-ins automatically: macOS denied access to the Office folders. Allow VisualTeX when macOS asks, then repair."
      : "VisualTeX 未能自动安装 Office 插件：macOS 拒绝访问 Office 文件夹。请在 macOS 询问时允许 VisualTeX，然后修复。";
  }
  return isEn
    ? `VisualTeX could not install its Office add-ins automatically: ${raw || "unknown error"}`
    : `VisualTeX 未能自动安装 Office 插件：${raw || "未知错误"}`;
}
