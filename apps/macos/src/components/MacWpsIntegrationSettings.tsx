import { useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { CheckCircle2, Download, FileText, RefreshCw, ShieldAlert, Trash2 } from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import { MacOfficeFullDiskAccessHelp } from "./MacOfficeFullDiskAccessHelp";
import { StatusLine } from "./MacOfficeIntegrationSettings";

/** office::wps::installer::WpsAddinStatus */
interface WpsContainerStatus {
  bundleId: string;
  addinPath: string;
  installed: boolean;
  current: boolean;
  error: string | null;
}

interface WpsAddinStatus {
  wpsInstalled: boolean;
  containers: WpsContainerStatus[];
  bridgeRunning: boolean;
  bridgeError: string | null;
  port: number;
  autoInstallDisabled: boolean;
}

/** WPS_FOLDER_ACCESS_DENIED in office/wps/installer.rs. */
const WPS_FOLDER_DENIED_ERROR = "macOS denied VisualTeX access to the WPS add-in folder";

function decodeWpsStatus(value: unknown): WpsAddinStatus {
  const status = value as Partial<WpsAddinStatus> | null;
  if (
    !status ||
    typeof status.wpsInstalled !== "boolean" ||
    !Array.isArray(status.containers) ||
    typeof status.bridgeRunning !== "boolean" ||
    typeof status.port !== "number"
  ) {
    throw new Error("VisualTeX returned an invalid WPS add-in status");
  }
  return status as WpsAddinStatus;
}

function errorText(error: unknown): string {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string") return error;
  try {
    return JSON.stringify(error);
  } catch {
    return String(error);
  }
}

function wpsBuildName(bundleId: string, isEn: boolean) {
  if (bundleId.endsWith(".global")) return isEn ? "WPS Office (international)" : "WPS Office（国际版）";
  return "WPS Office";
}

export function MacWpsIntegrationSettings() {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const [status, setStatus] = useState<WpsAddinStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [message, setMessage] = useState("");

  const refresh = useCallback(async () => {
    setBusy((value) => value ?? "refresh");
    try {
      setStatus(decodeWpsStatus(await invoke<unknown>("get_wps_addin_status")));
      setMessage("");
    } catch (error) {
      setMessage(errorText(error));
    } finally {
      setBusy((value) => (value === "refresh" ? null : value));
    }
  }, []);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const run = async (name: string, command: string) => {
    setBusy(name);
    try {
      const result = await invoke<unknown>(command);
      if (result) setStatus(decodeWpsStatus(result));
      setMessage("");
    } catch (error) {
      setMessage(errorText(error));
    } finally {
      setBusy(null);
    }
  };

  const containerError = status?.containers.find((container) => container.error)?.error ?? "";
  const error = message || containerError || status?.bridgeError || "";
  const accessDenied = error.includes(WPS_FOLDER_DENIED_ERROR);

  return (
    <section className="settings-section office-integration-section">
      <div className="settings-section-heading office-settings-heading">
        <div>
          <strong>{isEn ? "WPS Office add-in" : "WPS Office 加载项"}</strong>
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

      {status && (
        <div className="office-status-grid native-office-status-grid">
          {(status.containers.length > 0
            ? status.containers
            : [{ bundleId: "com.kingsoft.wpsoffice.mac", addinPath: "", installed: false, current: false, error: null }]
          ).map((container) => (
            <article className="office-status-card" key={container.bundleId}>
              <header>
                <strong>
                  <FileText size={16} /> {wpsBuildName(container.bundleId, isEn)}
                </strong>
                {container.current && status.bridgeRunning && (
                  <CheckCircle2 className="office-state-ok" size={15} />
                )}
              </header>
              <StatusLine ok={status.wpsInstalled}>
                {status.wpsInstalled
                  ? isEn ? "WPS Office is installed" : "WPS Office 已安装"
                  : isEn ? "WPS Office was not found (open it once after installing)" : "未找到 WPS Office（安装后请先打开一次）"}
              </StatusLine>
              <StatusLine ok={container.current}>
                {container.current
                  ? isEn ? "The VisualTeX add-in is installed" : "VisualTeX 加载项已安装"
                  : container.installed
                    ? isEn ? "The add-in needs to be updated" : "加载项需要更新"
                    : isEn ? "The add-in is not installed" : "加载项未安装"}
              </StatusLine>
              <StatusLine ok={status.bridgeRunning}>
                {status.bridgeRunning
                  ? isEn
                    ? `Connected to WPS on port ${status.port}`
                    : `已准备好与 WPS 连接（端口 ${status.port}）`
                  : isEn ? "VisualTeX is not listening for WPS" : "VisualTeX 未能开启与 WPS 的连接"}
              </StatusLine>
              {container.addinPath && (
                <details className="office-install-paths">
                  <summary>{isEn ? "Install location" : "安装位置"}</summary>
                  <code>{container.addinPath}</code>
                </details>
              )}
            </article>
          ))}
        </div>
      )}

      <p className="wps-addin-note">
        {isEn
          ? "The VisualTeX tab appears in WPS Writer. The first time WPS loads the add-in it asks twice: choose Allow, then OK. Keep VisualTeX running while you use the tab."
          : "VisualTeX 选项卡会出现在 WPS 文字中。WPS 第一次加载这个加载项时会先后询问两次，请先选择「允许」，再选择「确定」。使用时请保持 VisualTeX 处于打开状态。"}
      </p>

      {error && (
        <div className="office-settings-warning" role="alert" aria-live="assertive">
          <ShieldAlert size={15} />
          <span>
            <strong>
              {accessDenied
                ? isEn
                  ? "macOS blocked VisualTeX from the folder where WPS keeps its add-ins. To let VisualTeX install the add-in, turn on Full Disk Access for VisualTeX:"
                  : "macOS 阻止了 VisualTeX 访问 WPS 存放加载项的文件夹。要让 VisualTeX 安装加载项，请为 VisualTeX 打开「完全磁盘访问权限」："
                : isEn
                  ? "The WPS add-in operation failed."
                  : "WPS 加载项操作失败。"}
            </strong>
            {accessDenied && <MacOfficeFullDiskAccessHelp isEn={isEn} updating={false} />}
            <details open={!accessDenied}>
              <summary>{isEn ? "Error details" : "错误详情"}</summary>
              <code>{error}</code>
            </details>
          </span>
        </div>
      )}

      <div className="office-settings-actions">
        <button
          type="button"
          className="primary-button"
          disabled={busy !== null || !status?.wpsInstalled}
          onClick={() => void run("install", "install_wps_addin")}
        >
          <Download size={15} />
          {isEn ? "Install or update the WPS add-in" : "安装或更新 WPS 加载项"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !status?.wpsInstalled}
          onClick={() => void run("open", "open_wps_office")}
        >
          <FileText size={15} />
          {isEn ? "Open WPS" : "打开 WPS"}
        </button>
        <button
          type="button"
          className="secondary-button danger-subtle"
          disabled={busy !== null || !status?.containers.some((container) => container.installed)}
          onClick={() => void run("uninstall", "uninstall_wps_addin")}
        >
          <Trash2 size={15} />
          {isEn ? "Uninstall the WPS add-in" : "卸载 WPS 加载项"}
        </button>
      </div>
    </section>
  );
}
