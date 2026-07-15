import { useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  Download,
  ExternalLink,
  Play,
  RefreshCw,
  ShieldAlert,
  ShieldCheck,
  Square,
  ToggleLeft,
  ToggleRight,
  Trash2,
  Wrench,
} from "lucide-react";
import { useEditorStore } from "../stores/editorStore";

interface OfficeHostInstallStatus {
  applicationInstalled: boolean;
  manifestInstalled: boolean;
  manifestVersion: string | null;
  manifestPath: string;
}

interface CertificateInstallStatus {
  certificateExists: boolean;
  privateKeyExists: boolean;
  trusted: boolean;
  keychainPath: string;
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

interface OfficeBackgroundStatus {
  installed: boolean;
  loaded: boolean;
  runningInBackgroundMode: boolean;
  plistPath: string;
  executablePath: string;
  lastError: string | null;
}

interface OfficeIntegrationStatus {
  word: OfficeHostInstallStatus;
  powerpoint: OfficeHostInstallStatus;
  expectedManifestVersion: string;
  certificate: CertificateInstallStatus;
  background: OfficeBackgroundStatus;
  companion: OfficeCompanionStatus;
  officeUiVersion: string;
}

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

function HostCard({
  name,
  host,
  isEn,
  expectedVersion,
}: {
  name: string;
  host: OfficeHostInstallStatus;
  isEn: boolean;
  expectedVersion: string;
}) {
  const ready =
    host.applicationInstalled &&
    host.manifestInstalled &&
    host.manifestVersion === expectedVersion;
  return (
    <article className="office-status-card">
      <header>
        <strong>{name}</strong>
        {ready ? (
          <CheckCircle2 className="office-state-ok" size={15} />
        ) : (
          <ShieldAlert className="office-state-warning" size={15} />
        )}
      </header>
      <StatusLine ok={host.applicationInstalled}>
        {isEn ? "Application" : "应用"}: {host.applicationInstalled ? (isEn ? "installed" : "已安装") : (isEn ? "missing" : "未检测到")}
      </StatusLine>
      <StatusLine ok={host.manifestInstalled}>
        manifest: {host.manifestInstalled ? host.manifestVersion ?? "—" : isEn ? "missing" : "缺失"}
      </StatusLine>
      <p title={host.manifestPath}>{host.manifestPath}</p>
    </article>
  );
}

export function MacOfficeIntegrationSettings() {
  const isEn = useEditorStore((state) => state.language) === "en";
  const [status, setStatus] = useState<OfficeIntegrationStatus | null>(null);
  const [offlineStatus, setOfflineStatus] = useState<MacOfflineOfficeStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [message, setMessage] = useState("");

  const refresh = useCallback(async () => {
    setBusy((value) => value ?? "refresh");
    try {
      const [compatibility, offline] = await Promise.all([
        invoke<OfficeIntegrationStatus>("get_office_integration_status"),
        invoke<MacOfflineOfficeStatus>("get_macos_offline_office_install_status"),
      ]);
      setStatus(compatibility);
      setOfflineStatus(offline);
      setMessage("");
    } catch (error) {
      setMessage(
        errorMessage(
          error,
          isEn ? "Unable to read macOS Office status." : "无法读取 macOS Office 集成状态。",
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
            isEn ? "macOS Office operation failed." : "macOS Office 集成操作失败。",
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
          <strong>{isEn ? "macOS Office integration" : "macOS Office 集成"}</strong>
          <p>
            {isEn
              ? "Installs the local Word and PowerPoint add-in resources, HTTPS certificate, and background companion. Activate VisualTeX from Office Add-ins when needed."
              : "安装 Word 与 PowerPoint 本地加载项资源、HTTPS 证书和后台伴侣服务；需要时可从 Office 的“加载项”中启用 VisualTeX。"}
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
          <span>{isEn ? "Reading integration status…" : "正在读取集成状态…"}</span>
        </div>
      ) : (
        <div className="office-status-grid">
          <HostCard
            name="Microsoft Word"
            host={status.word}
            isEn={isEn}
            expectedVersion={status.expectedManifestVersion}
          />
          <HostCard
            name="Microsoft PowerPoint"
            host={status.powerpoint}
            isEn={isEn}
            expectedVersion={status.expectedManifestVersion}
          />

          <article className="office-status-card">
            <header>
              <strong>{isEn ? "Local companion" : "本地伴侣服务"}</strong>
            </header>
            <StatusLine ok={status.companion.running}>
              {status.companion.running
                ? `${status.companion.bindAddress}:${status.companion.port}`
                : isEn ? "Stopped" : "已停止"}
            </StatusLine>
            <dl>
              <div><dt>{isEn ? "Protocol" : "协议"}</dt><dd>{status.companion.protocolVersion}</dd></div>
              <div><dt>Office UI</dt><dd>{status.officeUiVersion}</dd></div>
            </dl>
          </article>

          <article className="office-status-card">
            <header>
              <strong>LaunchAgent</strong>
            </header>
            <StatusLine ok={status.background.installed}>
              {status.background.installed
                ? isEn ? "Installed" : "已安装"
                : isEn ? "Missing" : "缺失"}
            </StatusLine>
            <StatusLine ok={status.background.loaded}>
              {status.background.loaded
                ? isEn ? "Loaded for this login" : "当前登录会话已加载"
                : isEn ? "Not loaded" : "未加载"}
            </StatusLine>
            <p title={status.background.plistPath}>{status.background.plistPath || "—"}</p>
          </article>

          <article className="office-status-card">
            <header>
              <strong>{isEn ? "Login Keychain certificate" : "登录 Keychain 证书"}</strong>
              {status.certificate.trusted ? (
                <ShieldCheck className="office-state-ok" size={15} />
              ) : (
                <ShieldAlert className="office-state-warning" size={15} />
              )}
            </header>
            <StatusLine ok={status.certificate.certificateExists}>
              {isEn ? "Certificate file" : "证书文件"}
            </StatusLine>
            <StatusLine ok={status.certificate.privateKeyExists}>
              {isEn ? "Private key" : "私钥"}
            </StatusLine>
            <StatusLine ok={status.certificate.trusted}>
              {isEn ? "Trusted by login Keychain" : "已受登录 Keychain 信任"}
            </StatusLine>
            <p title={status.certificate.keychainPath}>{status.certificate.keychainPath}</p>
          </article>
        </div>
      )}

      <div className="settings-section-heading office-settings-heading">
        <div>
          <strong>{isEn ? "Native offline add-ins (staged)" : "原生离线加载项（分阶段启用）"}</strong>
          <p>
            {isEn
              ? "Word uses a global DOTM template and PowerPoint uses a fixed PPAM file. This route has no HTTPS, certificate, manifest, network, or Office.js dependency. The compatibility route above remains installed until native acceptance is complete."
              : "Word 使用全局 DOTM 模板，PowerPoint 使用固定路径 PPAM；该路线不依赖 HTTPS、证书、Manifest、网络或 Office.js。在原生验收全部完成前，上方兼容路线仍会保留。"}
          </p>
        </div>
      </div>

      {offlineStatus && (
        <div className="office-status-grid">
          <article className="office-status-card">
            <header><strong>Word · VisualTeX.dotm</strong></header>
            <StatusLine ok={offlineStatus.word.filesInstalled}>
              {offlineStatus.word.filesInstalled ? (isEn ? "Files installed" : "文件已安装") : (isEn ? "Files missing" : "文件缺失")}
            </StatusLine>
            <StatusLine ok={offlineStatus.word.loaded}>
              {offlineStatus.word.loaded ? (isEn ? "Loaded by Word" : "Word 已加载") : (isEn ? "Waiting for Word health signal" : "等待 Word 健康状态")}
            </StatusLine>
            <p title={offlineStatus.word.installPaths.join("\n")}>{offlineStatus.word.installPaths[0] ?? "—"}</p>
          </article>
          <article className="office-status-card">
            <header><strong>PowerPoint · VisualTeX.ppam</strong></header>
            <StatusLine ok={offlineStatus.powerpoint.filesInstalled}>
              {offlineStatus.powerpoint.filesInstalled ? (isEn ? "Files installed" : "文件已安装") : (isEn ? "Files missing" : "文件缺失")}
            </StatusLine>
            <StatusLine ok={offlineStatus.powerpoint.loaded}>
              {offlineStatus.powerpoint.loaded ? (isEn ? "Loaded by PowerPoint" : "PowerPoint 已加载") : (isEn ? "Manual registration or health signal required" : "需要手动登记或等待健康状态")}
            </StatusLine>
            <p title={offlineStatus.powerpointAddinPath}>{offlineStatus.powerpointAddinPath}</p>
          </article>
          <article className="office-status-card">
            <header><strong>{isEn ? "Compiled add-in package" : "已编译加载项包"}</strong></header>
            <StatusLine ok={offlineStatus.compiledArtifactsAvailable}>
              {offlineStatus.compiledArtifactsAvailable
                ? (isEn ? "DOTM and PPAM are available" : "DOTM 与 PPAM 已就绪")
                : (isEn ? "Reviewed VBA sources exist; compiled Office binaries are not packaged yet" : "VBA 源码已生成，但尚未打包真实 Office 编译产物")}
            </StatusLine>
            <p title={offlineStatus.resourceRoot}>{offlineStatus.resourceRoot}</p>
          </article>
        </div>
      )}

      <div className="office-settings-actions">
        <button type="button" className="secondary-button" disabled={busy !== null || !offlineStatus?.compiledArtifactsAvailable} onClick={() => void run("offline-install", "install_macos_offline_office_addins")}>
          <Download size={15} />{isEn ? "Install native offline add-ins" : "安装原生离线加载项"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null || !offlineStatus?.compiledArtifactsAvailable} onClick={() => void run("offline-repair", "repair_macos_offline_office_addins")}>
          <Wrench size={15} />{isEn ? "Repair native add-ins" : "修复原生加载项"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null || !offlineStatus?.powerpoint.filesInstalled} onClick={() => void run("offline-reveal", "reveal_macos_powerpoint_addin")}>
          <ExternalLink size={15} />{isEn ? "Show PPAM in Finder" : "在 Finder 中显示 PPAM"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null} onClick={() => void run("offline-tutorial", "open_macos_powerpoint_addin_tutorial")}>
          <ExternalLink size={15} />{isEn ? "Open PPAM tutorial" : "打开 PPAM 安装教程"}
        </button>
        <button type="button" className="secondary-button danger-subtle" disabled={busy !== null} onClick={() => void run("offline-uninstall", "uninstall_macos_offline_office_addins")}>
          <Trash2 size={15} />{isEn ? "Uninstall native add-ins" : "卸载原生加载项"}
        </button>
      </div>

      {(message || status?.certificate.lastError || status?.background.lastError || status?.companion.lastError || offlineStatus?.word.lastError || offlineStatus?.powerpoint.lastError) && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span>
            {message || status?.certificate.lastError || status?.background.lastError || status?.companion.lastError || offlineStatus?.word.lastError || offlineStatus?.powerpoint.lastError}
          </span>
        </div>
      )}

      <div className="office-settings-actions">
        <button type="button" className="primary-button" disabled={busy !== null} onClick={() => void run("install", "install_office_integration")}>
          <Download size={15} />{isEn ? "Install Office integration" : "安装 Office 集成"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null} onClick={() => void run("repair", "repair_office_integration")}>
          <Wrench size={15} />{isEn ? "Repair" : "修复 Office 集成"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run(
            "background-start",
            "set_office_background_start",
            { enabled: !status?.background.installed },
          )}
        >
          {status?.background.installed ? <ToggleRight size={15} /> : <ToggleLeft size={15} />}
          {status?.background.installed
            ? isEn ? "Disable startup" : "关闭开机启动"
            : isEn ? "Enable startup" : "启用开机启动"}
        </button>
        <button type="button" className="secondary-button danger-subtle" disabled={busy !== null} onClick={() => void run("uninstall", "uninstall_office_integration")}>
          <Trash2 size={15} />{isEn ? "Uninstall" : "卸载 Office 集成"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null} onClick={() => void run("certificate", "regenerate_office_certificate")}>
          <ShieldCheck size={15} />{isEn ? "Regenerate certificate" : "重新生成证书"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null || Boolean(status?.companion.running)} onClick={() => void run("start", "start_office_companion")}>
          <Play size={15} />{isEn ? "Start companion" : "启动伴侣服务"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null || !status?.companion.running} onClick={() => void run("stop", "stop_office_companion")}>
          <Square size={14} />{isEn ? "Stop companion" : "停止伴侣服务"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null || !status?.word.applicationInstalled} onClick={() => void run("word", "open_word")}>
          <ExternalLink size={15} />{isEn ? "Open Word" : "打开 Word"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null || !status?.powerpoint.applicationInstalled} onClick={() => void run("powerpoint", "open_powerpoint")}>
          <ExternalLink size={15} />{isEn ? "Open PowerPoint" : "打开 PowerPoint"}
        </button>
      </div>
    </section>
  );
}
