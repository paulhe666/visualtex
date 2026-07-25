import { useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  Circle,
  Download,
  ExternalLink,
  FileText,
  Presentation,
  RefreshCw,
  ShieldAlert,
  Trash2,
  Wrench,
} from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import { PowerPointAddinGuide } from "./PowerPointAddinGuide";

interface MacOfflineHostStatus {
  applicationInstalled: boolean;
  applicationRunning: boolean;
  filesInstalled: boolean;
  healthReported: boolean;
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

function errorMessage(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

function officeErrorSummary(raw: string, isEn: boolean) {
  if (/OfficePluginStatus|health file|invalid JSON|sourceRevision/i.test(raw)) {
    return isEn
      ? "An old add-in status record was ignored. Refresh the page; repair only if the Office buttons do not work."
      : "已忽略旧的插件状态记录。请先刷新；仅在 Office 按钮不可用时修复插件。";
  }
  if (/Fully quit|Command-Q/i.test(raw)) {
    return isEn
      ? "Quit Word and PowerPoint with Command-Q before installing or repairing."
      : "安装或修复前，请先使用 ⌘Q 完全退出 Word 和 PowerPoint。";
  }
  return isEn
    ? "The Office add-in operation failed. Refresh the status, then repair if needed."
    : "Office 插件操作失败。请先刷新状态，仍有问题时再点击修复。";
}

function StatusLine({
  ok,
  pending = false,
  children,
}: {
  ok: boolean;
  pending?: boolean;
  children: React.ReactNode;
}) {
  return (
    <div className="office-platform-status-line">
      {ok ? (
        <CheckCircle2 className="office-state-ok" size={15} />
      ) : pending ? (
        <Circle className="office-state-neutral" size={15} />
      ) : (
        <ShieldAlert className="office-state-warning" size={15} />
      )}
      <span>{children}</span>
    </div>
  );
}

export function MacOfficeIntegrationSettings() {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const [status, setStatus] = useState<MacOfflineOfficeStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [message, setMessage] = useState("");

  const refresh = useCallback(async () => {
    setBusy((value) => value ?? "refresh");
    try {
      const next = await invoke<MacOfflineOfficeStatus>(
        "get_macos_offline_office_install_status",
      );
      setStatus(next);
      setMessage("");
    } catch (error) {
      setMessage(
        errorMessage(
          error,
          isEn
            ? "Unable to read the native Office add-in status."
            : "无法读取原生 Office 加载项状态。",
        ),
      );
    } finally {
      setBusy((value) => (value === "refresh" ? null : value));
    }
  }, [isEn]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const run = useCallback(
    async (name: string, command: string) => {
      setBusy(name);
      setMessage("");
      try {
        await invoke(command);
        if (command !== "open_word" && command !== "open_powerpoint" && command !== "reveal_macos_powerpoint_addin") {
          await refresh();
        }
      } catch (error) {
        setMessage(
          errorMessage(
            error,
            isEn ? "The native Office operation failed." : "原生 Office 操作执行失败。",
          ),
        );
      } finally {
        setBusy(null);
      }
    },
    [isEn, refresh],
  );

  const powerpointNeedsVerification = Boolean(
    status?.powerpoint.applicationRunning &&
      status.powerpoint.filesInstalled &&
      !status.powerpoint.loaded,
  );
  const detailedError =
    message || status?.word.lastError || status?.powerpoint.lastError || "";

  return (
    <section className="settings-section office-integration-section">
      <div className="settings-section-heading office-settings-heading">
        <div>
          <strong>{isEn ? "Word and PowerPoint native add-ins" : "Word 与 PowerPoint 原生加载项"}</strong>
          <p>
            {isEn
              ? "Install, update and check the VisualTeX add-ins for Word and PowerPoint."
              : "安装、更新并检查 Word 和 PowerPoint 中的 VisualTeX 插件。"}
          </p>
        </div>
        <button
          type="button"
          className="icon-button compact"
          onClick={() => void refresh()}
          disabled={busy !== null}
          title={isEn ? "Refresh" : "刷新"}
        >
          <RefreshCw size={15} className={busy === "refresh" ? "is-spinning" : ""} />
        </button>
      </div>

      {!status ? (
        <div className="office-settings-loading">
          <RefreshCw size={16} className="is-spinning" />
          <span>{isEn ? "Reading native add-in status…" : "正在读取原生加载项状态…"}</span>
        </div>
      ) : (
        <div className="office-status-grid native-office-status-grid">
          <article className="office-status-card">
            <header>
              <strong><FileText size={16} /> Word</strong>
              {status.word.loaded && <CheckCircle2 className="office-state-ok" size={15} />}
            </header>
            <StatusLine ok={status.word.applicationInstalled}>
              {status.word.applicationInstalled
                ? isEn ? "Word is installed" : "Word 已安装"
                : isEn ? "Word was not found" : "未找到 Word"}
            </StatusLine>
            <StatusLine ok={status.word.filesInstalled}>
              {status.word.filesInstalled
                ? isEn ? "VisualTeX files match this app version" : "VisualTeX 插件文件为当前版本"
                : isEn ? "VisualTeX is not installed" : "VisualTeX 插件未安装"}
            </StatusLine>
            <StatusLine
              ok={status.word.loaded}
              pending={status.word.filesInstalled && !status.word.applicationRunning}
            >
              {status.word.loaded
                ? isEn ? "VisualTeX is loaded" : "VisualTeX 已加载"
                : !status.word.filesInstalled
                  ? isEn ? "Install the add-in first" : "请先安装插件"
                  : !status.word.applicationRunning
                    ? isEn ? "Open Word to check loading" : "打开 Word 后自动检查加载状态"
                    : isEn ? "Not confirmed; repair only if the buttons do not work" : "尚未确认；仅在按钮不可用时点击修复"}
            </StatusLine>
            <details className="office-install-paths">
              <summary>{isEn ? "Install location" : "安装位置"}</summary>
              <code>{status.word.installPaths[0] ?? "—"}</code>
            </details>
          </article>

          <article className="office-status-card">
            <header>
              <strong><Presentation size={16} /> PowerPoint</strong>
              {status.powerpoint.loaded && <CheckCircle2 className="office-state-ok" size={15} />}
            </header>
            <StatusLine ok={status.powerpoint.applicationInstalled}>
              {status.powerpoint.applicationInstalled
                ? isEn ? "PowerPoint is installed" : "PowerPoint 已安装"
                : isEn ? "PowerPoint was not found" : "未找到 PowerPoint"}
            </StatusLine>
            <StatusLine ok={status.powerpoint.filesInstalled}>
              {status.powerpoint.filesInstalled
                ? isEn ? "VisualTeX files match this app version" : "VisualTeX 插件文件为当前版本"
                : isEn ? "VisualTeX is not installed" : "VisualTeX 插件未安装"}
            </StatusLine>
            <StatusLine
              ok={status.powerpoint.loaded}
              pending={status.powerpoint.filesInstalled && !status.powerpoint.applicationRunning}
            >
              {status.powerpoint.loaded
                ? isEn ? "VisualTeX is loaded" : "VisualTeX 已加载"
                : !status.powerpoint.filesInstalled
                  ? isEn ? "Install the add-in first" : "请先安装插件"
                  : !status.powerpoint.applicationRunning
                    ? isEn ? "Open PowerPoint to check loading" : "打开 PowerPoint 后自动检查加载状态"
                    : isEn ? "Not confirmed; register the PPAM only if the ribbon is missing" : "尚未确认；仅在功能区缺失时重新登记 PPAM"}
            </StatusLine>
            <details className="office-install-paths">
              <summary>{isEn ? "Install location" : "安装位置"}</summary>
              <code>{status.powerpointAddinPath}</code>
            </details>
          </article>

          <div className="office-package-status">
            <StatusLine ok={status.compiledArtifactsAvailable}>
              {status.compiledArtifactsAvailable
                ? isEn ? "Installer resources are complete" : "安装包资源完整"
                : isEn ? "Installer resources are missing" : "安装包缺少插件资源"}
            </StatusLine>
          </div>
        </div>
      )}

      <div className="office-settings-actions">
        <button
          type="button"
          className="primary-button"
          disabled={busy !== null || !status?.compiledArtifactsAvailable}
          onClick={() => void run("install", "install_macos_offline_office_addins")}
        >
          <Download size={15} />
          {isEn ? "Install or update add-ins" : "安装或更新插件"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !status?.compiledArtifactsAvailable}
          onClick={() => void run("repair", "repair_macos_offline_office_addins")}
        >
          <Wrench size={15} />
          {isEn ? "Repair add-ins" : "修复插件"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !status?.word.applicationInstalled}
          onClick={() => void run("word", "open_word")}
        >
          <FileText size={15} />
          {isEn ? "Open Word" : "打开 Word"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !status?.powerpoint.applicationInstalled}
          onClick={() => void run("powerpoint", "open_powerpoint")}
        >
          <Presentation size={15} />
          {isEn ? "Open PowerPoint" : "打开 PowerPoint"}
        </button>
        <button
          type="button"
          className="secondary-button danger-subtle"
          disabled={busy !== null}
          onClick={() => void run("uninstall", "uninstall_macos_offline_office_addins")}
        >
          <Trash2 size={15} />
          {isEn ? "Uninstall add-ins" : "卸载插件"}
        </button>
      </div>

      {status?.powerpoint.applicationInstalled && !status.powerpoint.loaded && (
        <div className={`native-powerpoint-settings-guide${powerpointNeedsVerification ? " is-required" : ""}`}>
          <div className="settings-section-heading">
            <div>
              <strong>{isEn ? "Load VisualTeX in PowerPoint" : "在 PowerPoint 中加载 VisualTeX"}</strong>
              <p>
                {!status.powerpoint.filesInstalled
                  ? isEn
                    ? "Install the add-in first."
                    : "请先安装插件。"
                  : isEn
                    ? "Open PowerPoint and refresh. If the VisualTeX tab is still missing, add VisualTeX.ppam once."
                    : "先打开 PowerPoint 并刷新；如果仍没有 VisualTeX 选项卡，再手动添加一次 VisualTeX.ppam。"}
              </p>
            </div>
          </div>
          <PowerPointAddinGuide language={language} loaded={status.powerpoint.loaded} />
          <div className="office-settings-actions">
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || !status.powerpoint.filesInstalled}
              onClick={() => void run("reveal", "reveal_macos_powerpoint_addin")}
            >
              <ExternalLink size={15} />
              {isEn ? "Show VisualTeX.ppam in Finder" : "在 Finder 中显示 VisualTeX.ppam"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null}
              onClick={() => void refresh()}
            >
              <RefreshCw size={15} />
              {isEn ? "Check whether PowerPoint loaded it" : "检查 PowerPoint 是否已加载"}
            </button>
          </div>
        </div>
      )}

      {detailedError && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span>
            <strong>{officeErrorSummary(detailedError, isEn)}</strong>
            <details>
              <summary>{isEn ? "Technical details" : "技术详情"}</summary>
              <code>{detailedError}</code>
            </details>
          </span>
        </div>
      )}
    </section>
  );
}
