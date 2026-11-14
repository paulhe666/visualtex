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
import { readLocalStorage, writeLocalStorage } from "../runtime/safeStorage";
import {
  MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY,
  MAC_OFFICE_FIRST_RUN_STORAGE_KEY,
  macOfficeFolderAccessDenied,
  macOfficeFolderAccessDeniedText,
} from "./macOfficeOnboarding";
import { MacOfficeFullDiskAccessHelp } from "./MacOfficeFullDiskAccessHelp";
import { PowerPointAddinGuide } from "./PowerPointAddinGuide";
import {
  decodeMacOfflineOfficeStatus,
  type MacOfflineOfficeStatus,
} from "./macOfficeStatusValidation";

function errorMessage(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  if (error && typeof error === "object") {
    const candidate = error as Record<string, unknown>;
    for (const key of ["message", "error", "details"]) {
      const value = candidate[key];
      if (typeof value === "string" && value.trim()) return value;
    }
    try {
      const serialized = JSON.stringify(error);
      if (serialized && serialized !== "{}") return serialized;
    } catch {
      // Fall through to the localized fallback.
    }
  }
  return fallback;
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

function officeErrorSummary(raw: string, isEn: boolean) {
  if (macOfficeFolderAccessDenied(raw)) return macOfficeFolderAccessDeniedText(isEn);
  if (/add-in operation timed out/i.test(raw)) {
    return isEn
      ? "macOS did not allow VisualTeX to modify the Office folders in time."
      : "macOS 未及时允许 VisualTeX 修改 Office 文件夹。";
  }
  if (/timed out|timeout|temporarily unavailable/i.test(raw)) {
    return isEn
      ? "Office status check timed out."
      : "Office 状态读取超时。";
  }
  if (/OfficePluginStatus|health file|invalid JSON|sourceRevision/i.test(raw)) {
    return isEn
      ? "An outdated add-in status record was ignored."
      : "已忽略旧的插件状态记录。";
  }
  if (/Fully quit|Command-Q/i.test(raw)) {
    return isEn
      ? "Quit Word and PowerPoint with Command-Q first."
      : "请先使用 ⌘Q 退出 Word 和 PowerPoint。";
  }
  if (/Operation not permitted|Permission denied|EPERM|permission/i.test(raw)) {
    return isEn
      ? "macOS denied access to an Office add-in folder."
      : "macOS 拒绝访问 Office 插件目录。";
  }
  if (/still running|taking longer|operation is still/i.test(raw)) {
    return isEn
      ? "The Office add-in operation is taking unusually long."
      : "Office 插件操作耗时异常。";
  }
  return isEn
    ? "The Office add-in operation failed."
    : "Office 插件操作失败。";
}

export function StatusLine({
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
  const [openingLog, setOpeningLog] = useState(false);
  const [message, setMessage] = useState("");

  const refresh = useCallback(async () => {
    setBusy((value) => value ?? "refresh");
    try {
      const next = decodeMacOfflineOfficeStatus(
        await invokeWithTimeout<unknown>("get_macos_offline_office_install_status"),
      );
      setStatus(next);
      setMessage("");
    } catch (error) {
      const raw = errorMessage(
        error,
        isEn
          ? "Unable to read the native Office add-in status."
          : "无法读取原生 Office 加载项状态。",
      );
      setMessage(
        raw === "Office status request timed out"
          ? isEn
            ? "Office status request timed out."
            : "Office 状态读取超时。"
          : raw,
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
      const mutatesAddins =
        command === "install_macos_offline_office_addins" ||
        command === "repair_macos_offline_office_addins" ||
        command === "uninstall_macos_offline_office_addins";

      setBusy(name);
      setMessage("");
      let slowOperationTimer: number | undefined;
      if (mutatesAddins) {
        slowOperationTimer = window.setTimeout(() => {
          setMessage(
            isEn
              ? "Still working (timeout: 90 seconds)."
              : "仍在处理（超时：90 秒）。",
          );
        }, 5_000);
      }
      try {
        await invoke(command);
        if (mutatesAddins) {
          // Uninstall also stops VisualTeX from installing the add-ins again
          // on its own; installing or repairing turns that back on.
          writeLocalStorage(
            MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY,
            command === "uninstall_macos_offline_office_addins" ? "true" : "false",
          );
        }
        if (slowOperationTimer !== undefined) {
          window.clearTimeout(slowOperationTimer);
          slowOperationTimer = undefined;
        }
        if (command === "uninstall_macos_offline_office_addins") {
          // Cleanup success must not wait for a potentially unhealthy Office
          // container status probe. Refresh in the background instead.
          void refresh();
        } else if (
          command !== "open_word" &&
          command !== "open_powerpoint" &&
          command !== "reveal_macos_powerpoint_addin"
        ) {
          await refresh();
        }
      } catch (error) {
        if (slowOperationTimer !== undefined) {
          window.clearTimeout(slowOperationTimer);
        }
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

  const openMaintenanceLog = useCallback(async () => {
    setOpeningLog(true);
    try {
      await invoke("open_macos_office_addin_maintenance_log");
    } catch (error) {
      setMessage(
        errorMessage(
          error,
          isEn
            ? "Unable to open the Office add-in maintenance log."
            : "无法打开 Office 插件维护日志。",
        ),
      );
    } finally {
      setOpeningLog(false);
    }
  }, [isEn]);

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
          {busy === "refresh" ? (
            <RefreshCw size={16} className="is-spinning" />
          ) : (
            <ShieldAlert size={16} className="office-state-warning" />
          )}
          <span>
            {busy === "refresh"
              ? isEn
                ? "Reading native add-in status…"
                : "正在读取原生加载项状态…"
              : isEn
                ? "Office status check did not finish."
                : "Office 插件状态读取未完成。"}
          </span>
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
                    ? isEn ? "Word is not running" : "Word 未运行"
                    : isEn ? "Loading not confirmed" : "加载状态未确认"}
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
                    ? isEn ? "PowerPoint is not running" : "PowerPoint 未运行"
                    : isEn ? "Loading not confirmed" : "加载状态未确认"}
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

      {detailedError && (
        <div className="office-settings-warning" role="alert" aria-live="assertive">
          <ShieldAlert size={15} />
          <span>
            <strong>{officeErrorSummary(detailedError, isEn)}</strong>
            {macOfficeFolderAccessDenied(detailedError) && (
              <MacOfficeFullDiskAccessHelp
                isEn={isEn}
                updating={readLocalStorage(MAC_OFFICE_FIRST_RUN_STORAGE_KEY) === "true"}
              />
            )}
            <details open>
              <summary>{isEn ? "Error details" : "错误详情"}</summary>
              <code>{detailedError}</code>
            </details>
            <button
              type="button"
              className="secondary-button compact"
              disabled={openingLog}
              onClick={() => void openMaintenanceLog()}
            >
              <ExternalLink size={13} />
              {openingLog
                ? isEn
                  ? "Opening log…"
                  : "正在打开日志…"
                : isEn
                  ? "View diagnostic log"
                  : "查看诊断日志"}
            </button>
          </span>
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
              {isEn ? "Check load status" : "检查加载状态"}
            </button>
          </div>
        </div>
      )}

    </section>
  );
}
