import { useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  Download,
  ExternalLink,
  Play,
  RefreshCw,
  ShieldAlert,
  Square,
  Trash2,
  Wrench,
} from "lucide-react";
import { useEditorStore } from "../stores/editorStore";

export type WindowsOfficeMode = "auto" | "ole" | "vsto";

interface OfficePlatformStatus {
  platform: string;
  mode: WindowsOfficeMode;
  activeBackend: string;
  oleBridgeHealthy: boolean;
  vstoWordHealthy: boolean;
  vstoPowerpointHealthy: boolean;
  officeCatalogRegistered: boolean;
  currentUserCertificateTrusted: boolean;
  backgroundStartEnabled: boolean;
  lastError: string | null;
}

interface OfficeCompanionStatus {
  running: boolean;
  bindAddress: string;
  port: number;
  certificatePath: string;
  officeUiVersion: string;
  protocolVersion: number;
  lastError: string | null;
}

function errorMessage(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

function StatusLine({ ok, children }: { ok: boolean; children: React.ReactNode }) {
  return (
    <div className="office-platform-status-line">
      {ok ? (
        <CheckCircle2 className="office-state-ok" size={15} />
      ) : (
        <ShieldAlert className="office-state-warning" size={15} />
      )}
      <span>{children}</span>
    </div>
  );
}

export function WindowsOfficeIntegrationSettings() {
  const isEn = useEditorStore((state) => state.language) === "en";
  const [status, setStatus] = useState<OfficePlatformStatus | null>(null);
  const [companion, setCompanion] = useState<OfficeCompanionStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [message, setMessage] = useState("");

  const refresh = useCallback(async () => {
    setBusy((value) => value ?? "refresh");
    try {
      const [nextStatus, nextCompanion] = await Promise.all([
        invoke<OfficePlatformStatus>("get_office_platform_status"),
        invoke<OfficeCompanionStatus>("get_office_companion_status"),
      ]);
      setStatus(nextStatus);
      setCompanion(nextCompanion);
      setMessage("");
    } catch (error) {
      setMessage(
        errorMessage(
          error,
          isEn
            ? "Unable to read Windows Office integration status."
            : "无法读取 Windows Office 集成状态。",
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
    async (name: string, command: string, args?: Record<string, unknown>) => {
      setBusy(name);
      setMessage("");
      try {
        await invoke(command, args);
        if (command !== "open_word" && command !== "open_powerpoint") {
          await refresh();
        }
      } catch (error) {
        setMessage(
          errorMessage(
            error,
            isEn ? "Windows Office operation failed." : "Windows Office 操作失败。",
          ),
        );
      } finally {
        setBusy(null);
      }
    },
    [isEn, refresh],
  );

  return (
    <section className="settings-section office-integration-section">
      <div className="settings-section-heading office-settings-heading">
        <div>
          <strong>{isEn ? "Windows Office integration" : "Windows Office 集成"}</strong>
          <p>
            {isEn
              ? "The Windows release uses the OLE Bridge for Word and PowerPoint. Legacy VisualTeX VSTO add-ins are disabled to prevent duplicate ribbons."
              : "Windows 正式版使用 OLE Bridge 集成 Word 与 PowerPoint，并禁用旧版 VisualTeX VSTO 加载项以避免重复按钮。"}
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

      <div className="office-status-grid">
        <article className="office-status-card">
          <header>
            <strong>{isEn ? "Current backend" : "当前后端"}</strong>
            <StatusLine ok={Boolean(status && !status.lastError)}>
              {status?.activeBackend ?? "—"}
            </StatusLine>
          </header>
          <dl>
            <div><dt>{isEn ? "Selected mode" : "设置模式"}</dt><dd>OLE</dd></div>
            <div><dt>{isEn ? "Duplicate ribbon guard" : "重复按钮保护"}</dt><dd>{isEn ? "Legacy VSTO disabled" : "旧版 VSTO 已禁用"}</dd></div>
          </dl>
        </article>

        <article className="office-status-card">
          <header><strong>OLE Bridge</strong></header>
          <StatusLine ok={Boolean(status?.oleBridgeHealthy)}>
            {status?.oleBridgeHealthy
              ? isEn ? "Named pipe and STA backend healthy" : "命名管道与 STA 后端健康"
              : isEn ? "Unavailable or stopped" : "不可用或未启动"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.officeCatalogRegistered)}>
            {isEn ? "Office Catalog" : "Office Catalog"}
          </StatusLine>
        </article>

        <article className="office-status-card">
          <header><strong>{isEn ? "Local security" : "本地安全"}</strong></header>
          <StatusLine ok={Boolean(status?.currentUserCertificateTrusted)}>
            {isEn ? "Current-user HTTPS certificate" : "当前用户 HTTPS 证书"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.backgroundStartEnabled)}>
            {isEn ? "Background startup" : "后台启动"}
          </StatusLine>
        </article>

        <article className="office-status-card office-status-card-wide">
          <header><strong>{isEn ? "Session companion" : "Session 伴侣服务"}</strong></header>
          <StatusLine ok={Boolean(companion?.running)}>
            {companion?.running
              ? `${companion.bindAddress}:${companion.port}`
              : isEn ? "Stopped" : "已停止"}
          </StatusLine>
          <p>
            {isEn
              ? "OCR, formula rendering, cache and editor sessions remain in the shared HTTPS companion."
              : "OCR、公式渲染、缓存和编辑会话继续由共享 HTTPS 伴侣服务负责。"}
          </p>
        </article>
      </div>

      {(message || status?.lastError || companion?.lastError) && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span>{message || status?.lastError || companion?.lastError}</span>
        </div>
      )}

      <div className="office-settings-actions">
        <button
          type="button"
          className="primary-button"
          disabled={busy !== null}
          onClick={() => void run("install-ole", "install_windows_ole_integration")}
        >
          <Download size={15} />
          {isEn ? "Install / enable OLE" : "安装/启用 OLE"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("repair", "repair_windows_office_integration")}
        >
          <Wrench size={15} />
          {isEn ? "Repair OLE integration" : "修复 OLE 集成"}
        </button>
        <button
          type="button"
          className="secondary-button danger-subtle"
          disabled={busy !== null}
          onClick={() => void run("uninstall-ole", "uninstall_windows_ole_integration")}
        >
          <Trash2 size={15} />
          {isEn ? "Remove OLE manifest" : "移除 OLE manifest"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || Boolean(companion?.running)}
          onClick={() => void run("start", "start_office_companion")}
        >
          <Play size={15} />
          {isEn ? "Start companion" : "启动伴侣服务"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !companion?.running}
          onClick={() => void run("stop", "stop_office_companion")}
        >
          <Square size={14} />
          {isEn ? "Stop companion" : "停止伴侣服务"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null} onClick={() => void run("word", "open_word")}>
          <ExternalLink size={15} />{isEn ? "Open Word" : "打开 Word"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null} onClick={() => void run("powerpoint", "open_powerpoint")}>
          <ExternalLink size={15} />{isEn ? "Open PowerPoint" : "打开 PowerPoint"}
        </button>
      </div>
    </section>
  );
}
