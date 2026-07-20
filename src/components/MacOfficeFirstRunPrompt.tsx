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

interface MacOfflineHostStatus {
  applicationInstalled: boolean;
  filesInstalled: boolean;
  loaded: boolean;
  pluginVersion: string | null;
  installPaths: string[];
  healthPath: string;
  lastError: string | null;
}

interface MacOfflineOfficeStatus {
  word: MacOfflineHostStatus;
  powerpoint: MacOfflineHostStatus;
  compiledArtifactsAvailable: boolean;
  resourceRoot: string;
  powerpointAddinPath: string;
  wordScriptPath: string;
  powerpointScriptPath: string;
  tutorialPath: string;
}

interface Props {
  open: boolean;
  language: Language;
  onComplete: (installed: boolean) => void;
}

function messageFrom(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

export function MacOfficeFirstRunPrompt({ open, language, onComplete }: Props) {
  const [status, setStatus] = useState<MacOfflineOfficeStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState("");
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";

  const refresh = async () => {
    const next = await invoke<MacOfflineOfficeStatus>(
      "get_macos_offline_office_install_status",
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
    status?.powerpoint.applicationInstalled &&
      status.powerpoint.filesInstalled &&
      !status.powerpoint.loaded,
  );

  const install = async () => {
    setBusy("install");
    setError("");
    try {
      const next = await invoke<MacOfflineOfficeStatus>(
        "install_macos_offline_office_addins",
      );
      setStatus(next);
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
              {isEn ? "Set up VisualTeX for Word and PowerPoint" : "配置 Word 与 PowerPoint 的 VisualTeX 插件"}
            </strong>
            <p>
              {isEn
                ? "VisualTeX installs a Word DOTM template and a PowerPoint PPAM add-in. Both run locally and open the desktop formula editor when needed."
                : "VisualTeX 会安装 Word DOTM 模板和 PowerPoint PPAM 加载项；两者都在本机运行，并在需要时打开桌面公式编辑器。"}
            </p>
          </div>
        </header>

        <div className="office-first-run-hosts">
          <article className={status?.word.loaded ? "is-loaded" : status?.word.filesInstalled ? "is-files-ready" : ""}>
            <FileText size={20} />
            <div>
              <strong>Microsoft Word · DOTM</strong>
              <small>
                {!status?.word.applicationInstalled
                  ? isEn ? "Word not detected" : "未检测到 Word"
                  : status.word.loaded
                    ? isEn ? "Installed and loaded" : "已安装并加载"
                    : status.word.filesInstalled
                      ? isEn ? "Installed; restart Word" : "已安装，请重启 Word"
                      : isEn ? "Not installed" : "尚未安装"}
              </small>
            </div>
            {status?.word.loaded ? <CheckCircle2 size={17} /> : status?.word.filesInstalled ? <ShieldAlert size={17} /> : null}
          </article>
          <article className={status?.powerpoint.loaded ? "is-loaded" : status?.powerpoint.filesInstalled ? "is-files-ready" : ""}>
            <Presentation size={20} />
            <div>
              <strong>Microsoft PowerPoint · PPAM</strong>
              <small>
                {!status?.powerpoint.applicationInstalled
                  ? isEn ? "PowerPoint not detected" : "未检测到 PowerPoint"
                  : status.powerpoint.loaded
                    ? isEn ? "Installed and loaded" : "已安装并加载"
                    : status.powerpoint.filesInstalled
                      ? isEn ? "Installed; register once" : "已安装，需要登记一次"
                      : isEn ? "Not installed" : "尚未安装"}
              </small>
            </div>
            {status?.powerpoint.loaded ? <CheckCircle2 size={17} /> : status?.powerpoint.filesInstalled ? <ShieldAlert size={17} /> : null}
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
                  void refresh().finally(() => setBusy(null));
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
              {isEn
                ? "Word loads VisualTeX automatically from its Startup folder after Word restarts. PowerPoint needs one manual registration through Tools → PowerPoint Add-ins; later updates keep the same PPAM path."
                : "Word 重启后会从 Startup 目录自动加载 VisualTeX。PowerPoint 需要在“工具 → PowerPoint 加载项”中手动登记一次；后续更新会继续覆盖同一个 PPAM 路径。"}
            </p>
            {!officeDetected && status && (
              <p className="is-warning">
                {isEn
                  ? "Word or PowerPoint was not found. Open the Office application once, then return to Settings to install the native add-ins."
                  : "未检测到 Word 或 PowerPoint。请先打开一次对应 Office 应用，再回到设置中安装原生加载项。"}
              </p>
            )}
            {nativeFilesReady && (
              <p className="is-warning">
                {isEn
                  ? "The files are installed, but this does not mean Word or PowerPoint has loaded them. Fully quit Word with Command-Q and reopen it; register the PPAM in PowerPoint once."
                  : "文件已安装不等于 Office 已加载。请用 ⌘Q 完全退出 Word 后重新打开；PowerPoint 还需要手动登记 PPAM 一次。"}
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
            onClick={() => nativeFilesReady ? onComplete(true) : void install()}
          >
            {busy === "install" ? <LoaderCircle className="is-spinning" size={16} /> : nativeFilesReady ? <CheckCircle2 size={16} /> : <Download size={16} />}
            {nativeFilesReady
              ? isEn ? "Continue" : "继续"
              : isEn ? "Install DOTM and PPAM" : "安装 DOTM 和 PPAM"}
          </button>
        </footer>
      </section>
    </div>
  );
}
