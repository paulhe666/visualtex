import { useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  Download,
  ExternalLink,
  FileText,
  LoaderCircle,
  Presentation,
  RefreshCw,
  ShieldAlert,
} from "lucide-react";
import type { Language } from "../stores/editorStore";
import { writeLocalStorage } from "../runtime/safeStorage";
import {
  MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY,
  macOfficeFolderAccessDenied,
  macOfficeFolderAccessDeniedText,
  macOfficeTextIsFolderAccessDenied,
} from "./macOfficeOnboarding";
import { MacOfficeFullDiskAccessHelp } from "./MacOfficeFullDiskAccessHelp";
import { PowerPointAddinGuide } from "./PowerPointAddinGuide";
import {
  decodeMacOfflineOfficeStatus,
  isMacOfficeStatusUnavailable,
  macOfficeStatusUnavailableText,
  type MacOfflineOfficeStatus,
} from "./macOfficeStatusValidation";

interface Props {
  open: boolean;
  language: Language;
  mode?: "setup" | "update" | "repair";
  powerpointRegistrationRequired?: boolean;
  /** Why VisualTeX's own install failed, shown when the prompt opens. */
  initialError?: string;
  onComplete: (installed: boolean) => void;
}

function messageFrom(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

/** An install error in the prompt's words; a folder refusal gets its steps below. */
function installErrorText(error: unknown, isEn: boolean, fallback: string) {
  return macOfficeFolderAccessDenied(error)
    ? macOfficeFolderAccessDeniedText(isEn)
    : messageFrom(error, fallback);
}

function wait(milliseconds: number) {
  return new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds));
}

async function invokeWithTimeout<T>(command: string, timeoutMs = 6_000): Promise<T> {
  let timeoutId: number | undefined;
  try {
    return await Promise.race([
      invoke<T>(command),
      new Promise<T>((_, reject) => {
        timeoutId = window.setTimeout(() => {
          reject(new Error("Office status request timed out"));
        }, timeoutMs);
      }),
    ]);
  } finally {
    if (timeoutId !== undefined) window.clearTimeout(timeoutId);
  }
}

export function MacOfficeFirstRunPrompt({
  open,
  language,
  mode = "setup",
  powerpointRegistrationRequired = false,
  initialError = "",
  onComplete,
}: Props) {
  const [status, setStatus] = useState<MacOfflineOfficeStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState("");
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";

  const refresh = async () => {
    const next = decodeMacOfflineOfficeStatus(
      await invokeWithTimeout<unknown>("get_macos_offline_office_install_status"),
    );
    setStatus(next);
    return next;
  };

  useEffect(() => {
    if (!open) return;
    setError(initialError);
    setBusy("refresh");
    void refresh()
      .catch((reason) => {
        const raw = messageFrom(
          reason,
          isEn
            ? "Unable to inspect the native Office add-ins on this Mac."
            : "无法检测这台 Mac 上的原生 Office 加载项。",
        );
        setError(
          isMacOfficeStatusUnavailable(reason)
            ? macOfficeStatusUnavailableText(isEn)
            : raw,
        );
      })
      .finally(() => setBusy(null));
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [initialError, isEn, open]);

  if (!open) return null;

  const officeDetected = Boolean(
    status?.word.applicationInstalled || status?.powerpoint.applicationInstalled,
  );
  const nativeFilesReady = Boolean(
    status?.compiledArtifactsAvailable &&
      (!status.word.applicationInstalled || status.word.filesInstalled) &&
      (!status.powerpoint.applicationInstalled || status.powerpoint.filesInstalled),
  );
  const powerpointNeedsRegistration = Boolean(
    powerpointRegistrationRequired &&
      status?.powerpoint.applicationInstalled &&
      status.powerpoint.filesInstalled &&
      !status.powerpoint.loaded,
  );
  const officeHostsRunning = Boolean(
    status?.word.applicationRunning || status?.powerpoint.applicationRunning,
  );
  const updateRequired = Boolean(
    mode !== "setup" &&
      ((status?.word.applicationInstalled && !status.word.filesInstalled) ||
        (status?.powerpoint.applicationInstalled &&
          !status.powerpoint.filesInstalled)),
  );

  const install = async () => {
    setBusy("install");
    setError("");
    try {
      const next = decodeMacOfflineOfficeStatus(
        await invoke<unknown>("install_macos_offline_office_addins"),
      );
      writeLocalStorage(MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY, "false");
      setStatus(next);
      if (
        mode !== "setup" &&
        (!next.word.applicationInstalled || next.word.filesInstalled) &&
        (!next.powerpoint.applicationInstalled || next.powerpoint.filesInstalled)
      ) {
        onComplete(true);
      }
    } catch (reason) {
      setError(
        installErrorText(
          reason,
          isEn,
          isEn
            ? "VisualTeX could not install the native Word and PowerPoint add-ins."
            : "VisualTeX 无法安装 Word 和 PowerPoint 原生加载项。",
        ),
      );
    } finally {
      setBusy(null);
    }
  };

  const quitOfficeAndUpdate = async () => {
    setBusy("quit-update");
    setError("");
    try {
      await invoke("request_quit_macos_office_hosts_for_addin_update");
      const deadline = Date.now() + 120_000;
      while (Date.now() < deadline) {
        await wait(500);
        const current = await refresh();
        if (!current.word.applicationRunning && !current.powerpoint.applicationRunning) {
          const next = decodeMacOfflineOfficeStatus(
            await invoke<unknown>("install_macos_offline_office_addins"),
          );
          setStatus(next);
          if (
            (!next.word.applicationInstalled || next.word.filesInstalled) &&
            (!next.powerpoint.applicationInstalled || next.powerpoint.filesInstalled)
          ) {
            onComplete(true);
          }
          return;
        }
      }
      throw new Error(
        isEn
          ? "Timed out waiting for Word and PowerPoint to quit."
          : "等待 Word 和 PowerPoint 退出超时。",
      );
    } catch (reason) {
      setError(
        installErrorText(
          reason,
          isEn,
          isEn
            ? "VisualTeX could not finish the Office add-in update."
            : "VisualTeX 无法完成 Office 插件更新。",
        ),
      );
    } finally {
      setBusy(null);
    }
  };

  const runAction = async (name: string, command: string) => {
    setBusy(name);
    setError("");
    try {
      await invoke(command);
      if (command === "open_powerpoint") {
        window.setTimeout(() => void refresh().catch(() => undefined), 1200);
      }
    } catch (reason) {
      setError(
        messageFrom(
          reason,
          isEn ? "The requested Office action failed." : "Office 操作执行失败。",
        ),
      );
    } finally {
      setBusy(null);
    }
  };

  return (
    <div className="office-first-run-backdrop">
      <section
        ref={dialogRef}
        className="office-first-run-dialog is-native-office"
        role="dialog"
        aria-modal="true"
        aria-labelledby="office-first-run-title"
      >
        <header>
          <span><Download size={20} /></span>
          <div>
            <strong id="office-first-run-title">
              {mode === "update"
                ? isEn
                  ? "Update the VisualTeX Office add-ins"
                  : "更新 VisualTeX Office 插件"
                : mode === "repair"
                  ? isEn
                    ? "Repair the VisualTeX Office add-ins"
                    : "修复 VisualTeX Office 插件"
                  : isEn
                    ? "Set up VisualTeX for Word and PowerPoint"
                    : "配置 Word 与 PowerPoint 的 VisualTeX 插件"}
            </strong>
          </div>
        </header>

        <div className="office-first-run-hosts">
          <article className={status?.word.filesInstalled && status.word.loaded ? "is-loaded" : status?.word.filesInstalled ? "is-files-ready" : ""}>
            <FileText size={20} />
            <div>
              <strong>Microsoft Word · DOTM</strong>
              <span>
                {!status?.word.applicationInstalled
                  ? isEn ? "Word not detected" : "未检测到 Word"
                  : status.word.filesInstalled && status.word.loaded
                    ? isEn ? "Installed and loaded" : "已安装并加载"
                    : status.word.filesInstalled
                      ? isEn ? "Installed; restart Word" : "已安装，请重启 Word"
                      : status.word.filesPresent
                        ? isEn ? "Older add-in detected; update required" : "检测到旧插件，需要更新"
                        : isEn ? "Not installed" : "尚未安装"}
              </span>
            </div>
            {status?.word.filesInstalled && status.word.loaded ? <CheckCircle2 size={17} /> : status?.word.filesInstalled ? <ShieldAlert size={17} /> : null}
          </article>
          <article className={status?.powerpoint.filesInstalled && status.powerpoint.loaded ? "is-loaded" : status?.powerpoint.filesInstalled ? "is-files-ready" : ""}>
            <Presentation size={20} />
            <div>
              <strong>Microsoft PowerPoint · PPAM</strong>
              <span>
                {!status?.powerpoint.applicationInstalled
                  ? isEn ? "PowerPoint not detected" : "未检测到 PowerPoint"
                  : status.powerpoint.filesInstalled && status.powerpoint.loaded
                    ? isEn ? "Installed and loaded" : "已安装并加载"
                    : status.powerpoint.filesInstalled
                      ? powerpointRegistrationRequired
                        ? isEn ? "Installed; not registered" : "已安装，未登记"
                        : isEn ? "Installed" : "已安装"
                      : status.powerpoint.filesPresent
                        ? isEn ? "Older add-in detected; update required" : "检测到旧插件，需要更新"
                        : isEn ? "Not installed" : "尚未安装"}
              </span>
            </div>
            {status?.powerpoint.filesInstalled && status.powerpoint.loaded ? <CheckCircle2 size={17} /> : status?.powerpoint.filesInstalled ? <ShieldAlert size={17} /> : null}
          </article>
        </div>

        {powerpointNeedsRegistration ? (
          <div className="office-first-run-powerpoint-guide">
            <div className="office-first-run-note is-important">
              <p>
                {isEn
                  ? "PowerPoint has not registered the PPAM."
                  : "PowerPoint 尚未登记 PPAM。"}
              </p>
            </div>
            <PowerPointAddinGuide language={language} compact loaded={false} />
            <div className="office-first-run-guide-actions">
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => void runAction("reveal", "reveal_macos_powerpoint_addin")}
              >
                <ExternalLink size={15} />
                {isEn ? "Show PPAM in Finder" : "在 Finder 中显示 PPAM"}
              </button>
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => void runAction("powerpoint", "open_powerpoint")}
              >
                <Presentation size={15} />
                {isEn ? "Open PowerPoint" : "打开 PowerPoint"}
              </button>
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => {
                  setBusy("refresh");
                  void refresh()
                    .catch((reason) => {
                      const raw = messageFrom(
                        reason,
                        isEn
                          ? "Unable to refresh the native Office add-in status."
                          : "无法刷新原生 Office 加载项状态。",
                      );
                      setError(
                        isMacOfficeStatusUnavailable(reason)
                          ? macOfficeStatusUnavailableText(isEn)
                          : raw,
                      );
                    })
                    .finally(() => setBusy(null));
                }}
              >
                <RefreshCw size={15} className={busy === "refresh" ? "is-spinning" : ""} />
                {isEn ? "Refresh status" : "刷新状态"}
              </button>
            </div>
          </div>
        ) : (
          <div className="office-first-run-note">
            <p>
              {mode === "update"
                ? isEn ? "Updates the installed DOTM and PPAM." : "将更新已安装的 DOTM 和 PPAM。"
                : mode === "repair"
                  ? isEn ? "Restores missing DOTM and PPAM files." : "将恢复缺失的 DOTM 和 PPAM 文件。"
                  : powerpointRegistrationRequired
                    ? isEn
                      ? "PowerPoint requires one registration in Tools → PowerPoint Add-ins."
                      : "PowerPoint 需在“工具 → PowerPoint 加载项”中登记一次。"
                    : isEn ? "Installs the DOTM and PPAM." : "将安装 DOTM 和 PPAM。"}
            </p>
            {mode !== "setup" && updateRequired && officeHostsRunning && (
              <p className="is-warning">
                {isEn
                  ? "Save your Office documents first. Updating quits Word and PowerPoint."
                  : "请先保存 Office 文档。更新时将退出 Word 和 PowerPoint。"}
              </p>
            )}
            {!officeDetected && status && (
              <p className="is-warning">
                {isEn
                  ? "Word or PowerPoint not found."
                  : "未检测到 Word 或 PowerPoint。"}
              </p>
            )}
            {nativeFilesReady && (
              <p className="is-warning">
                {powerpointRegistrationRequired
                  ? isEn
                    ? "Add-in files installed. Restart Word and register the PowerPoint PPAM."
                    : "插件文件已安装。请重启 Word 并登记 PowerPoint PPAM。"
                  : isEn
                    ? "Add-in files ready. Reopen Word or PowerPoint."
                    : "插件文件已就绪。请重新打开 Word 或 PowerPoint。"}
              </p>
            )}
            {error && <p className="is-warning" role="alert">{error}</p>}
            {error && macOfficeTextIsFolderAccessDenied(error) && (
              <MacOfficeFullDiskAccessHelp isEn={isEn} updating={mode !== "setup"} />
            )}
          </div>
        )}

        {powerpointNeedsRegistration && error && (
          <div className="office-settings-warning" role="alert">
            <ShieldAlert size={15} />
            <span>{error}</span>
          </div>
        )}

        <footer>
          <button
            type="button"
            className="secondary-button"
            disabled={busy !== null}
            onClick={() => onComplete(false)}
          >
            {isEn ? "Later" : "稍后处理"}
          </button>
          <button
            type="button"
            className="primary-button"
            disabled={busy !== null || (!nativeFilesReady && !officeDetected)}
            onClick={() => {
              if (nativeFilesReady) {
                onComplete(true);
              } else if (mode !== "setup" && officeHostsRunning) {
                void quitOfficeAndUpdate();
              } else {
                void install();
              }
            }}
          >
            {busy === "install" || busy === "quit-update" ? (
              <LoaderCircle className="is-spinning" size={16} />
            ) : nativeFilesReady ? (
              <CheckCircle2 size={16} />
            ) : (
              <Download size={16} />
            )}
            {nativeFilesReady
              ? mode === "setup"
                ? isEn ? "Continue" : "继续"
                : isEn ? "Done" : "完成"
              : mode === "update"
                ? officeHostsRunning
                  ? isEn ? "Quit Office and update add-ins" : "退出 Office 并更新插件"
                  : isEn ? "Update DOTM and PPAM" : "更新 DOTM 和 PPAM"
                : mode === "repair"
                  ? officeHostsRunning
                    ? isEn ? "Quit Office and repair add-ins" : "退出 Office 并修复插件"
                    : isEn ? "Repair DOTM and PPAM" : "修复 DOTM 和 PPAM"
                  : isEn ? "Install DOTM and PPAM" : "安装 DOTM 和 PPAM"}
          </button>
        </footer>
      </section>
    </div>
  );
}
