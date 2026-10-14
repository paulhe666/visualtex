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
  if (/timed out|timeout/i.test(raw)) {
    return isEn
      ? "Office status checking timed out. You can still uninstall the add-ins, or quit Office and refresh."
      : "Office 状态读取超时。仍可直接卸载插件，或退出 Office 后重新刷新。";
  }
  if (/OfficePluginStatus|health file|invalid JSON|sourceRevision/i.test(raw)) {
    return isEn
      ? "An old add-in status record was ignored. Refresh the page; repair only if the Office buttons do not work."
      : "已忽略旧的插件状态记录。请先刷新；仅在 Office 按钮不可用时修复插件。";
  }
  if (/Fully quit|Command-Q/i.test(raw)) {
    return isEn
      ? "Quit Word and PowerPoint completely with Command-Q before installing, repairing, or uninstalling the add-ins."
      : "安装、修复或卸载插件前，请先使用 ⌘Q 完全退出 Word 和 PowerPoint。";
  }
  if (/Operation not permitted|Permission denied|EPERM|permission/i.test(raw)) {
    return isEn
      ? "macOS denied access to an Office add-in location. Check VisualTeX permissions, then retry."
      : "macOS 拒绝访问 Office 插件目录。请检查 VisualTeX 的系统权限后重试。";
  }
  if (/still running|taking longer|operation is still/i.test(raw)) {
    return isEn
      ? "The Office add-in operation is taking unusually long. It may be blocked by Office, macOS permissions, or a coordinated Office folder."
      : "Office 插件操作耗时异常，可能被 Office 进程、macOS 权限或 Office 目录协调状态阻塞。";
  }
  return isEn
    ? "The Office add-in operation failed. See the error details below and retry after resolving it."
    : "Office 插件操作失败。请查看下方错误详情，处理后再重试。";
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
              ? "The Office add-in operation is still running after 5 seconds. It may be blocked by Office, macOS permissions, or an Office folder."
              : "Office 插件操作超过 5 秒仍未完成，可能被 Office 进程、macOS 权限或 Office 目录阻塞。",
          );
        }, 5_000);
      }
      try {
        await invoke(command);
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
                ? "The Office status check did not finish. You can retry or uninstall the add-ins directly."
                : "Office 插件状态读取未完成。可以重试，或直接卸载插件。"}
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

      {detailedError && (
        <div className="office-settings-warning" role="alert" aria-live="assertive">
          <ShieldAlert size={15} />
          <span>
            <strong>{officeErrorSummary(detailedError, isEn)}</strong>
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
              {isEn ? "Check whether PowerPoint loaded it" : "检查 PowerPoint 是否已加载"}
            </button>
          </div>
        </div>
      )}

    </section>
  );
}
