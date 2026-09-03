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
import { PowerPointAddinGuide } from "./PowerPointAddinGuide";
import {
  decodeMacOfflineOfficeStatus,
  type MacOfflineOfficeStatus,
} from "./macOfficeStatusValidation";

interface Props {
  open: boolean;
  language: Language;
  mode?: "setup" | "update" | "repair";
  powerpointRegistrationRequired?: boolean;
  onComplete: (installed: boolean) => void;
}

function messageFrom(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

function wait(milliseconds: number) {
  return new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds));
}

export function MacOfficeFirstRunPrompt({
  open,
  language,
  mode = "setup",
  powerpointRegistrationRequired = false,
  onComplete,
}: Props) {
  const [status, setStatus] = useState<MacOfflineOfficeStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState("");
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";

  const refresh = async () => {
    const next = decodeMacOfflineOfficeStatus(
      await invoke<unknown>("get_macos_offline_office_install_status"),
    );
    setStatus(next);
    return next;
  };

  useEffect(() => {
    if (!open) return;
    setError("");
    setBusy("refresh");
    void refresh()
      .catch((reason) => {
        setError(
          messageFrom(
            reason,
            isEn
              ? "Unable to inspect the native Office add-ins on this Mac."
              : "无法检测这台 Mac 上的原生 Office 加载项。",
          ),
        );
      })
      .finally(() => setBusy(null));
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [isEn, open]);

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
        messageFrom(
          reason,
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
          ? "Timed out waiting for Word and PowerPoint to quit. Finish any Save prompts, then try again."
          : "等待 Word 和 PowerPoint 退出超时。请先处理 Office 的保存提示，然后重试。",
      );
    } catch (reason) {
      setError(
        messageFrom(
          reason,
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
            <p>
              {mode === "update"
                ? isEn
                  ? "The installed DOTM or PPAM belongs to an older VisualTeX build. You do not need to delete it manually; VisualTeX will replace the old add-ins with the versions bundled in this app."
                  : "检测到已安装的 DOTM 或 PPAM 属于旧版 VisualTeX。无需手动删除旧文件，VisualTeX 会直接用当前应用内置版本覆盖更新。"
                : mode === "repair"
                  ? isEn
                    ? "VisualTeX detected a missing or incomplete Office add-in installation. It will restore the required DOTM/PPAM files without asking you to re-register an already configured PowerPoint add-in."
                    : "检测到 Office 插件文件缺失或安装不完整。VisualTeX 会恢复所需 DOTM/PPAM；已经配置过的 PowerPoint 加载项不会因此要求重新登记。"
                  : isEn
                    ? "VisualTeX installs a Word DOTM template and a PowerPoint PPAM add-in. Both run locally and open the desktop formula editor when needed."
                    : "VisualTeX 会安装 Word DOTM 模板和 PowerPoint PPAM 加载项；两者都在本机运行，并在需要时打开桌面公式编辑器。"}
            </p>
          </div>
        </header>

        <div className="office-first-run-hosts">
          <article className={status?.word.filesInstalled && status.word.loaded ? "is-loaded" : status?.word.filesInstalled ? "is-files-ready" : ""}>
            <FileText size={20} />
            <div>
              <strong>Microsoft Word · DOTM</strong>
              <small>
                {!status?.word.applicationInstalled
                  ? isEn ? "Word not detected" : "未检测到 Word"
                  : status.word.filesInstalled && status.word.loaded
                    ? isEn ? "Installed and loaded" : "已安装并加载"
                    : status.word.filesInstalled
                      ? isEn ? "Installed; restart Word" : "已安装，请重启 Word"
                      : status.word.filesPresent
                        ? isEn ? "Older add-in detected; update required" : "检测到旧插件，需要更新"
                        : isEn ? "Not installed" : "尚未安装"}
              </small>
            </div>
            {status?.word.filesInstalled && status.word.loaded ? <CheckCircle2 size={17} /> : status?.word.filesInstalled ? <ShieldAlert size={17} /> : null}
          </article>
          <article className={status?.powerpoint.filesInstalled && status.powerpoint.loaded ? "is-loaded" : status?.powerpoint.filesInstalled ? "is-files-ready" : ""}>
            <Presentation size={20} />
            <div>
              <strong>Microsoft PowerPoint · PPAM</strong>
              <small>
                {!status?.powerpoint.applicationInstalled
                  ? isEn ? "PowerPoint not detected" : "未检测到 PowerPoint"
                  : status.powerpoint.filesInstalled && status.powerpoint.loaded
                    ? isEn ? "Installed and loaded" : "已安装并加载"
                    : status.powerpoint.filesInstalled
                      ? powerpointRegistrationRequired
                        ? isEn ? "Installed; register once" : "已安装，需要登记一次"
                        : isEn ? "Installed; registration is preserved" : "已安装，原有登记保持不变"
                      : status.powerpoint.filesPresent
                        ? isEn ? "Older add-in detected; update required" : "检测到旧插件，需要更新"
                        : isEn ? "Not installed" : "尚未安装"}
              </small>
            </div>
            {status?.powerpoint.filesInstalled && status.powerpoint.loaded ? <CheckCircle2 size={17} /> : status?.powerpoint.filesInstalled ? <ShieldAlert size={17} /> : null}
          </article>
        </div>

        {powerpointNeedsRegistration ? (
          <div className="office-first-run-powerpoint-guide">
            <div className="office-first-run-note is-important">
              <p>
                {isEn
                  ? "The PPAM file is ready, but PowerPoint has not registered it. Not seeing VisualTeX in the Add-ins list yet is expected. Click + first, then choose the PPAM file."
                  : "PPAM 文件已经准备好，但 PowerPoint 尚未登记它。此时在加载项列表里看不到 VisualTeX 是正常的；必须先点击左下角＋，再选择 PPAM 文件。"}
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
                      setError(
                        messageFrom(
                          reason,
                          isEn
                            ? "Unable to refresh the native Office add-in status."
                            : "无法刷新原生 Office 加载项状态。",
                        ),
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
                ? isEn
                  ? "VisualTeX updates the existing DOTM and PPAM in place. PowerPoint keeps the same registered PPAM path, so an update does not require registering the add-in again."
                  : "VisualTeX 会在原路径直接覆盖旧 DOTM 和 PPAM。PowerPoint 会继续使用原先登记的 PPAM 路径，版本更新不需要重新登记加载项。"
                : mode === "repair"
                  ? isEn
                    ? "VisualTeX will restore the missing Office files. Any PowerPoint PPAM that was already registered keeps the same path and does not need to be registered again."
                    : "VisualTeX 会恢复缺失的 Office 插件文件。已经登记过的 PowerPoint PPAM 路径保持不变，不需要重新登记。"
                  : powerpointRegistrationRequired
                    ? isEn
                      ? "Word loads VisualTeX automatically from its Startup folder after Word restarts. PowerPoint needs one manual registration through Tools → PowerPoint Add-ins; later updates keep the same PPAM path."
                      : "Word 重启后会从 Startup 目录自动加载 VisualTeX。PowerPoint 首次安装需要在“工具 → PowerPoint 加载项”中登记一次；后续更新不会要求重新登记。"
                    : isEn
                      ? "VisualTeX will install the required Office files. Existing PowerPoint registration is preserved."
                      : "VisualTeX 会安装所需 Office 插件文件；已有的 PowerPoint 登记状态会保留。"}
            </p>
            {mode !== "setup" && updateRequired && officeHostsRunning && (
              <p className="is-warning">
                {isEn
                  ? "Save your Office documents first. VisualTeX must let Word and PowerPoint quit normally before replacing loaded VBA add-ins. Click the update button below; any unsaved Office document will still receive its normal Save prompt."
                  : "请先保存 Office 文档。已加载的 VBA 插件必须在 Word 和 PowerPoint 正常退出后才能替换。点击下方更新按钮即可；如有未保存文档，Office 仍会正常弹出保存提示。"}
              </p>
            )}
            {!officeDetected && status && (
              <p className="is-warning">
                {isEn
                  ? "Word or PowerPoint was not found. Open the Office application once, then return to Settings to install the native add-ins."
                  : "未检测到 Word 或 PowerPoint。请先打开一次对应 Office 应用，再回到设置中安装原生加载项。"}
              </p>
            )}
            {nativeFilesReady && (
              <p className="is-warning">
                {powerpointRegistrationRequired
                  ? isEn
                    ? "The files are installed. Restart Word, then register the PowerPoint PPAM once to finish first-time setup."
                    : "插件文件已安装。请重新打开 Word，并完成一次 PowerPoint PPAM 登记以结束首次配置。"
                  : isEn
                    ? "The current add-in files are ready. Reopen Word or PowerPoint so Office loads the repaired/updated files."
                    : "当前插件文件已就绪。重新打开 Word 或 PowerPoint 后，Office 会加载修复/更新后的插件。"}
              </p>
            )}
            {error && <p className="is-warning" role="alert">{error}</p>}
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
