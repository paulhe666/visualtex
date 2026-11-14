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

/** Set once VisualTeX has installed its add-ins on this Mac. */
export const MAC_OFFICE_FIRST_RUN_STORAGE_KEY =
  "visualtex.office.macos.native-first-run.v1.2.0.completed";

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

/** Start of the installer's error when macOS refuses VisualTeX the Office folders. */
const MAC_OFFICE_FOLDER_DENIED_ERROR = "macOS denied VisualTeX access to the Office folders";

/** macOS refused VisualTeX the Office folders (OFFICE_FOLDER_ACCESS_DENIED in Rust). */
export function macOfficeFolderAccessDenied(reason: unknown): boolean {
  return errorMessage(reason).includes(MAC_OFFICE_FOLDER_DENIED_ERROR);
}

/**
 * What the prompt and Settings say when macOS refused the Office folders;
 * MacOfficeFullDiskAccessHelp shows the steps under it.
 */
export function macOfficeFolderAccessDeniedText(isEn: boolean): string {
  return isEn
    ? "macOS blocked access to the Office add-in folder. Turn on Full Disk Access for VisualTeX:"
    : "macOS 已阻止访问 Office 插件文件夹。请为 VisualTeX 开启「完全磁盘访问权限」：";
}

export function macOfficeTextIsFolderAccessDenied(text: string): boolean {
  return text === macOfficeFolderAccessDeniedText(true) || text === macOfficeFolderAccessDeniedText(false);
}

/**
 * The Full Disk Access steps. macOS 27 never asks for another developer's
 * folder and lists Full Disk Access per app under Files & Folders; earlier
 * releases have a Full Disk Access list. An entry left by an earlier VisualTeX
 * build belongs to that build's signature, so an updating user replaces it.
 */
export function macOfficeFullDiskAccessSteps(
  isEn: boolean,
  macosMajor: number | null,
  updating: boolean,
): string[] {
  const filesAndFolders = macosMajor === null || macosMajor >= 27;
  const steps = isEn
    ? [
        filesAndFolders
          ? "Open Settings: System Settings > Privacy & Security > Files & Folders."
          : "Open Settings: System Settings > Privacy & Security > Full Disk Access.",
        ...(updating
          ? ["Remove any earlier VisualTeX entry (select it, then click −)."]
          : []),
        filesAndFolders
          ? "Turn on Full Disk Access for VisualTeX. If it is not listed, click + and choose VisualTeX in Applications."
          : "Turn on VisualTeX. If it is not listed, click + and choose VisualTeX in Applications.",
        "Choose Quit & Reopen, or restart VisualTeX. The add-ins are then installed automatically.",
      ]
    : [
        filesAndFolders
          ? "打开设置：系统设置 > 隐私与安全性 > 文件与文件夹。"
          : "打开设置：系统设置 > 隐私与安全性 > 完全磁盘访问权限。",
        ...(updating
          ? ["移除列表中的旧版 VisualTeX（选中后点击 −）。"]
          : []),
        filesAndFolders
          ? "为 VisualTeX 开启「完全磁盘访问权限」。如列表中没有 VisualTeX，点击 + 并在“应用程序”中选择。"
          : "开启 VisualTeX。如列表中没有 VisualTeX，点击 + 并在“应用程序”中选择。",
        "选择「退出并重新打开」，或重新启动 VisualTeX。插件随后自动安装。",
      ];
  return steps;
}

function errorMessage(reason: unknown): string {
  return reason instanceof Error ? reason.message : typeof reason === "string" ? reason : "";
}

/** Why VisualTeX's automatic install failed, in the words the prompt shows. */
export function macOfficeAutoInstallErrorText(reason: unknown, isEn: boolean): string {
  const raw = errorMessage(reason);
  if (macOfficeFolderAccessDenied(raw)) return macOfficeFolderAccessDeniedText(isEn);
  if (/add-in operation timed out/i.test(raw)) {
    return isEn
      ? "Office add-in installation failed: macOS did not grant access to the Office folders in time. If macOS asks for access to data from other apps, click Allow, then repair."
      : "Office 插件安装失败：macOS 未及时授予 Office 文件夹访问权限。如 macOS 请求访问其他 App 的数据，点击「允许」后修复。";
  }
  if (/Operation not permitted|Permission denied|EPERM/i.test(raw)) {
    return isEn
      ? "Office add-in installation failed: macOS denied access to an Office folder. Quit Word and PowerPoint, then try again."
      : "Office 插件安装失败：macOS 拒绝访问 Office 文件夹。请退出 Word 和 PowerPoint 后重试。";
  }
  return isEn
    ? `Office add-in installation failed: ${raw || "unknown error"}`
    : `Office 插件安装失败：${raw || "未知错误"}`;
}
