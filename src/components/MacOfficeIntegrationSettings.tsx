import { useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
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

  return (
    <section className="settings-section office-integration-section">
      <div className="settings-section-heading office-settings-heading">
        <div>
          <strong>{isEn ? "Word and PowerPoint native add-ins" : "Word 与 PowerPoint 原生加载项"}</strong>
          <p>
            {isEn
              ? "Word uses VisualTeX.dotm and PowerPoint uses VisualTeX.ppam. The add-ins communicate with the desktop app through local Office runtime files and the visualtex URL scheme."
              : "Word 使用 VisualTeX.dotm，PowerPoint 使用 VisualTeX.ppam。加载项通过 Office 本地运行文件和 visualtex URL Scheme 与桌面应用通信。"}
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
              <strong><FileText size={16} /> Word · VisualTeX.dotm</strong>
              {status.word.loaded && <CheckCircle2 className="office-state-ok" size={15} />}
            </header>
            <StatusLine ok={status.word.applicationInstalled}>
              {status.word.applicationInstalled
                ? isEn ? "Microsoft Word detected" : "已检测到 Microsoft Word"
                : isEn ? "Microsoft Word not detected" : "未检测到 Microsoft Word"}
            </StatusLine>
            <StatusLine ok={status.word.applicationRunning}>
              {status.word.applicationRunning
                ? isEn ? "Word is running" : "Word 正在运行"
                : isEn ? "Word is not running" : "Word 当前未运行"}
            </StatusLine>
            <StatusLine ok={status.word.filesInstalled}>
              {status.word.filesInstalled
                ? isEn ? "DOTM and AppleScriptTask installed" : "DOTM 与 AppleScriptTask 已安装"
                : isEn ? "Native Word files are missing" : "Word 原生文件尚未安装"}
            </StatusLine>
            <StatusLine ok={status.word.loaded}>
              {status.word.loaded
                ? isEn ? "Word confirmed that VisualTeX is loaded" : "Word 已确认加载 VisualTeX"
                : !status.word.filesInstalled
                  ? isEn ? "Waiting for installation" : "等待安装"
                  : !status.word.applicationRunning
                    ? isEn ? "Files are installed. Start Word to verify whether the DOTM loads" : "文件仅已安装；请启动 Word 后再验证 DOTM 是否加载"
                    : !status.word.healthReported
                      ? isEn ? "Word is running, but no VisualTeX load-confirmation file has been created. This does not mean the DOTM is missing; if formula buttons also fail, repair the native file bridge" : "Word 已启动，但尚未生成 VisualTeX 加载确认文件。这不代表 DOTM 未安装；如果公式按钮也报错，请修复原生文件桥，而不是反复重启 Word"
                      : isEn ? "Word is running, but the VisualTeX load report is stale or incompatible. Run Repair after quitting Word" : "Word 正在运行，但 VisualTeX 加载报告已过期或版本不匹配。请退出 Word 后执行修复"}
            </StatusLine>
            <p title={status.word.installPaths.join("\n")}>{status.word.installPaths[0] ?? "—"}</p>
          </article>

          <article className="office-status-card">
            <header>
              <strong><Presentation size={16} /> PowerPoint · VisualTeX.ppam</strong>
              {status.powerpoint.loaded && <CheckCircle2 className="office-state-ok" size={15} />}
            </header>
            <StatusLine ok={status.powerpoint.applicationInstalled}>
              {status.powerpoint.applicationInstalled
                ? isEn ? "Microsoft PowerPoint detected" : "已检测到 Microsoft PowerPoint"
                : isEn ? "Microsoft PowerPoint not detected" : "未检测到 Microsoft PowerPoint"}
            </StatusLine>
            <StatusLine ok={status.powerpoint.applicationRunning}>
              {status.powerpoint.applicationRunning
                ? isEn ? "PowerPoint is running" : "PowerPoint 正在运行"
                : isEn ? "PowerPoint is not running" : "PowerPoint 当前未运行"}
            </StatusLine>
            <StatusLine ok={status.powerpoint.filesInstalled}>
              {status.powerpoint.filesInstalled
                ? isEn ? "PPAM and AppleScriptTask installed" : "PPAM 与 AppleScriptTask 已安装"
                : isEn ? "Native PowerPoint files are missing" : "PowerPoint 原生文件尚未安装"}
            </StatusLine>
            <StatusLine ok={status.powerpoint.loaded}>
              {status.powerpoint.loaded
                ? isEn ? "PowerPoint confirmed that VisualTeX is loaded" : "PowerPoint 已确认加载 VisualTeX"
                : !status.powerpoint.filesInstalled
                  ? isEn ? "Waiting for installation" : "等待安装"
                  : !status.powerpoint.applicationRunning
                    ? isEn ? "Files are installed. Start PowerPoint to verify the fixed PPAM" : "文件仅已安装；请启动 PowerPoint 后验证固定路径中的 PPAM"
                    : !status.powerpoint.healthReported
                      ? isEn ? "PowerPoint is running, but no VisualTeX load-confirmation file has been created. This alone does not mean the PPAM is unregistered; if formula buttons also fail, repair the native file bridge" : "PowerPoint 已启动，但尚未生成 VisualTeX 加载确认文件。这不等于 PPAM 未登记；如果公式按钮也报错，请修复原生文件桥"
                      : isEn ? "PowerPoint is running, but the VisualTeX load report is stale or incompatible. Quit PowerPoint and run Repair" : "PowerPoint 正在运行，但 VisualTeX 加载报告已过期或版本不匹配。请退出 PowerPoint 后执行修复"}
            </StatusLine>
            <p title={status.powerpointAddinPath}>{status.powerpointAddinPath}</p>
          </article>

          <article className="office-status-card">
            <header><strong>{isEn ? "Packaged native resources" : "随应用打包的原生资源"}</strong></header>
            <StatusLine ok={status.compiledArtifactsAvailable}>
              {status.compiledArtifactsAvailable
                ? isEn ? "Reviewed DOTM and PPAM are available" : "已包含经过验收的 DOTM 与 PPAM"
                : isEn ? "Compiled Office add-ins are missing from this app" : "当前应用包中缺少已编译 Office 加载项"}
            </StatusLine>
            <p title={status.resourceRoot}>{status.resourceRoot}</p>
          </article>
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
          {isEn ? "Install DOTM and PPAM" : "安装 DOTM 和 PPAM"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !status?.compiledArtifactsAvailable}
          onClick={() => void run("repair", "repair_macos_offline_office_addins")}
        >
          <Wrench size={15} />
          {isEn ? "Repair native add-ins" : "修复原生加载项"}
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
          {isEn ? "Uninstall native add-ins" : "卸载原生加载项"}
        </button>
      </div>

      {status?.powerpoint.applicationInstalled && (
        <div className={`native-powerpoint-settings-guide${powerpointNeedsVerification ? " is-required" : ""}`}>
          <div className="settings-section-heading">
            <div>
              <strong>{isEn ? "Load VisualTeX in PowerPoint" : "在 PowerPoint 中加载 VisualTeX"}</strong>
              <p>
                {status.powerpoint.loaded
                  ? isEn
                    ? "PowerPoint has confirmed the fixed PPAM. Future updates keep the same path."
                    : "PowerPoint 已确认加载固定路径中的 PPAM，后续更新会继续使用同一路径。"
                  : !status.powerpoint.filesInstalled
                    ? isEn
                      ? "Install or repair the native add-ins before checking PowerPoint registration."
                      : "请先安装或修复原生加载项，再检查 PowerPoint 登记状态。"
                    : !status.powerpoint.applicationRunning
                      ? isEn
                        ? "The PPAM file is installed, but PowerPoint is not running. Open PowerPoint and refresh this page before deciding whether manual registration is needed."
                        : "PPAM 文件已经安装，但 PowerPoint 当前未运行。请先打开 PowerPoint 并刷新本页，再判断是否需要手动登记。"
                      : isEn
                        ? "PowerPoint is running, but VisualTeX has not been confirmed. Check the add-in list first; register the fixed VisualTeX.ppam only when it is actually absent."
                        : "PowerPoint 正在运行，但尚未确认 VisualTeX。请先检查加载项列表；只有列表中确实不存在时，才手动登记固定路径中的 VisualTeX.ppam。"}
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

      {(message || status?.word.lastError || status?.powerpoint.lastError) && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span>{message || status?.word.lastError || status?.powerpoint.lastError}</span>
        </div>
      )}
    </section>
  );
}
