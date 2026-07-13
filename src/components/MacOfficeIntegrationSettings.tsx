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
  certificate: CertificateInstallStatus;
  background: OfficeBackgroundStatus;
  companion: OfficeCompanionStatus;
  officeUiVersion: string;
}

const EXPECTED_MANIFEST_VERSION = "1.0.18.0";

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
}: {
  name: string;
  host: OfficeHostInstallStatus;
  isEn: boolean;
}) {
  const ready =
    host.applicationInstalled &&
    host.manifestInstalled &&
    host.manifestVersion === EXPECTED_MANIFEST_VERSION;
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
  const [busy, setBusy] = useState<string | null>(null);
  const [message, setMessage] = useState("");

  const refresh = useCallback(async () => {
    setBusy((value) => value ?? "refresh");
    try {
      setStatus(await invoke<OfficeIntegrationStatus>("get_office_integration_status"));
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
    async (name: string, command: string) => {
      setBusy(name);
      setMessage("");
      try {
        await invoke(command);
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
              ? "The existing Office.js, AppleScript, login Keychain and LaunchAgent integration remains independent from Windows Office components."
              : "现有 Office.js、AppleScript、登录 Keychain 与 LaunchAgent 集成保持不变，并与 Windows Office 组件完全独立。"}
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
          <HostCard name="Microsoft Word" host={status.word} isEn={isEn} />
          <HostCard name="Microsoft PowerPoint" host={status.powerpoint} isEn={isEn} />

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

      {(message || status?.certificate.lastError || status?.background.lastError || status?.companion.lastError) && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span>
            {message || status?.certificate.lastError || status?.background.lastError || status?.companion.lastError}
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
