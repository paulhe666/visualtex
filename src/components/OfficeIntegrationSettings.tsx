import { useCallback, useEffect, useMemo, useState } from "react";
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

interface OfficeIntegrationStatus {
  word: OfficeHostInstallStatus;
  powerpoint: OfficeHostInstallStatus;
  certificate: CertificateInstallStatus;
  companion: OfficeCompanionStatus;
  officeUiVersion: string;
}

interface OcrRuntimeStatus {
  installed: boolean;
  pythonPath?: string | null;
  pythonVersion?: string | null;
  paddleVersion?: string | null;
  paddleocrVersion?: string | null;
  runtimePath?: string | null;
  message?: string | null;
}

type OfficeAction =
  | "install"
  | "repair"
  | "uninstall"
  | "certificate"
  | "start"
  | "stop"
  | "ocr"
  | "word"
  | "powerpoint";

const EXPECTED_MANIFEST_VERSION = "1.0.6.0";

function isDesktopRuntime() {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

function operationErrorMessage(error: unknown, fallback: string) {
  if (error instanceof Error) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

function StateMark({ ok }: { ok: boolean }) {
  return ok ? (
    <CheckCircle2 className="office-state-ok" size={15} />
  ) : (
    <ShieldAlert className="office-state-warning" size={15} />
  );
}

function hostSummary(
  host: OfficeHostInstallStatus,
  isEn: boolean,
) {
  if (!host.applicationInstalled) {
    return isEn ? "Application not found" : "未检测到应用";
  }
  if (!host.manifestInstalled) {
    return isEn ? "Add-in not installed" : "插件未安装";
  }
  if (host.manifestVersion !== EXPECTED_MANIFEST_VERSION) {
    return isEn ? "Manifest version mismatch" : "manifest 版本不匹配";
  }
  return isEn ? "Application and add-in ready" : "应用和插件均已就绪";
}

export function OfficeIntegrationSettings() {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const [status, setStatus] = useState<OfficeIntegrationStatus | null>(null);
  const [ocrStatus, setOcrStatus] = useState<OcrRuntimeStatus | null>(null);
  const [busy, setBusy] = useState<OfficeAction | "refresh" | null>(null);
  const [message, setMessage] = useState("");
  const desktopRuntime = isDesktopRuntime();

  const refresh = useCallback(async () => {
    if (!desktopRuntime) return;
    setBusy((value) => value ?? "refresh");
    try {
      const [office, ocr] = await Promise.all([
        invoke<OfficeIntegrationStatus>("get_office_integration_status"),
        invoke<OcrRuntimeStatus>("get_ocr_runtime_status").catch(() => null),
      ]);
      setStatus(office);
      setOcrStatus(ocr);
      setMessage("");
    } catch (error) {
      setMessage(
        operationErrorMessage(
          error,
          isEn
            ? "Unable to read Office integration status"
            : "无法读取 Office 集成状态",
        ),
      );
    } finally {
      setBusy((value) => (value === "refresh" ? null : value));
    }
  }, [desktopRuntime, isEn]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const run = useCallback(
    async (action: OfficeAction, command: string) => {
      setBusy(action);
      setMessage("");
      try {
        await invoke(command);
        if (action !== "word" && action !== "powerpoint") {
          await refresh();
        }
      } catch (error) {
        setMessage(
          operationErrorMessage(
            error,
            isEn
              ? "Office integration operation failed"
              : "Office 集成操作失败",
          ),
        );
      } finally {
        setBusy(null);
      }
    },
    [isEn, refresh],
  );

  const manifestMismatch = useMemo(
    () =>
      Boolean(
        status &&
          ((status.word.manifestInstalled &&
            status.word.manifestVersion !== EXPECTED_MANIFEST_VERSION) ||
            (status.powerpoint.manifestInstalled &&
              status.powerpoint.manifestVersion !== EXPECTED_MANIFEST_VERSION)),
      ),
    [status],
  );

  if (!desktopRuntime) return null;

  return (
    <section className="settings-section office-integration-section">
      <div className="settings-section-heading office-settings-heading">
        <div>
          <strong>{isEn ? "Office integration" : "Office 集成"}</strong>
          <p>
            {isEn
              ? "Install the local Word and PowerPoint bridge, certificate and companion service."
              : "安装本地 Word/PowerPoint 桥接插件、证书和伴侣服务。"}
          </p>
        </div>
        <button
          type="button"
          className="icon-button compact"
          onClick={() => void refresh()}
          disabled={busy !== null}
          aria-label={isEn ? "Refresh Office status" : "刷新 Office 状态"}
          title={isEn ? "Refresh Office status" : "刷新 Office 状态"}
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
          <article className="office-status-card">
            <header>
              <strong>Microsoft Word</strong>
              <StateMark
                ok={
                  status.word.applicationInstalled &&
                  status.word.manifestInstalled &&
                  status.word.manifestVersion === EXPECTED_MANIFEST_VERSION
                }
              />
            </header>
            <p>{hostSummary(status.word, isEn)}</p>
            <dl>
              <div>
                <dt>{isEn ? "Application" : "应用"}</dt>
                <dd>{status.word.applicationInstalled ? (isEn ? "Installed" : "已安装") : (isEn ? "Missing" : "未安装")}</dd>
              </div>
              <div>
                <dt>{isEn ? "Add-in" : "插件"}</dt>
                <dd>{status.word.manifestInstalled ? (isEn ? "Installed" : "已安装") : (isEn ? "Missing" : "未安装")}</dd>
              </div>
              <div>
                <dt>manifest</dt>
                <dd>{status.word.manifestVersion ?? "—"}</dd>
              </div>
            </dl>
          </article>

          <article className="office-status-card">
            <header>
              <strong>Microsoft PowerPoint</strong>
              <StateMark
                ok={
                  status.powerpoint.applicationInstalled &&
                  status.powerpoint.manifestInstalled &&
                  status.powerpoint.manifestVersion === EXPECTED_MANIFEST_VERSION
                }
              />
            </header>
            <p>{hostSummary(status.powerpoint, isEn)}</p>
            <dl>
              <div>
                <dt>{isEn ? "Application" : "应用"}</dt>
                <dd>{status.powerpoint.applicationInstalled ? (isEn ? "Installed" : "已安装") : (isEn ? "Missing" : "未安装")}</dd>
              </div>
              <div>
                <dt>{isEn ? "Add-in" : "插件"}</dt>
                <dd>{status.powerpoint.manifestInstalled ? (isEn ? "Installed" : "已安装") : (isEn ? "Missing" : "未安装")}</dd>
              </div>
              <div>
                <dt>manifest</dt>
                <dd>{status.powerpoint.manifestVersion ?? "—"}</dd>
              </div>
            </dl>
          </article>

          <article className="office-status-card">
            <header>
              <strong>{isEn ? "Local companion" : "本地伴侣服务"}</strong>
              <StateMark ok={status.companion.running} />
            </header>
            <p>
              {status.companion.running
                ? isEn
                  ? "Running on loopback only"
                  : "仅在本机回环地址运行"
                : isEn
                  ? "Stopped"
                  : "已停止"}
            </p>
            <dl>
              <div>
                <dt>{isEn ? "Address" : "地址"}</dt>
                <dd>{status.companion.bindAddress}:{status.companion.port}</dd>
              </div>
              <div>
                <dt>{isEn ? "Protocol" : "协议"}</dt>
                <dd>{status.companion.protocolVersion}</dd>
              </div>
              <div>
                <dt>{isEn ? "Office UI" : "Office UI"}</dt>
                <dd>{status.officeUiVersion}</dd>
              </div>
            </dl>
          </article>

          <article className="office-status-card">
            <header>
              <strong>{isEn ? "Certificate" : "本地 HTTPS 证书"}</strong>
              {status.certificate.trusted ? (
                <ShieldCheck className="office-state-ok" size={15} />
              ) : (
                <ShieldAlert className="office-state-warning" size={15} />
              )}
            </header>
            <p>
              {status.certificate.trusted
                ? isEn
                  ? "Trusted by the login Keychain"
                  : "已受登录 Keychain 信任"
                : isEn
                  ? "Not trusted"
                  : "尚未受信任"}
            </p>
            <dl>
              <div>
                <dt>{isEn ? "Certificate" : "证书"}</dt>
                <dd>{status.certificate.certificateExists ? (isEn ? "Present" : "存在") : (isEn ? "Missing" : "缺失")}</dd>
              </div>
              <div>
                <dt>{isEn ? "Private key" : "私钥"}</dt>
                <dd>{status.certificate.privateKeyExists ? (isEn ? "Present" : "存在") : (isEn ? "Missing" : "缺失")}</dd>
              </div>
            </dl>
          </article>

          <article className="office-status-card office-status-card-wide">
            <header>
              <strong>{isEn ? "Offline OCR" : "离线 OCR"}</strong>
              <StateMark ok={Boolean(ocrStatus?.installed)} />
            </header>
            <p>
              {ocrStatus?.installed
                ? isEn
                  ? "OCR runtime is installed. The default M model is checked by the offline package installer."
                  : "OCR runtime 已安装；默认 M 模型由离线资源安装器进一步校验。"
                : isEn
                  ? "OCR runtime is not installed"
                  : "OCR runtime 尚未安装"}
            </p>
            <dl>
              <div>
                <dt>runtime</dt>
                <dd>{ocrStatus?.installed ? (isEn ? "Available" : "可用") : (isEn ? "Missing" : "缺失")}</dd>
              </div>
              <div>
                <dt>PP-FormulaNet plus-M</dt>
                <dd>{isEn ? "Pending offline package check" : "等待离线资源包校验"}</dd>
              </div>
              <div>
                <dt>{isEn ? "Python" : "Python"}</dt>
                <dd>{ocrStatus?.pythonVersion ?? "—"}</dd>
              </div>
            </dl>
          </article>
        </div>
      )}

      {manifestMismatch && (
        <div className="office-settings-warning">
          <ShieldAlert size={15} />
          <span>
            {isEn
              ? "The installed manifest does not match VisualTeX 1.0.6. Use Repair Office Integration."
              : "已安装 manifest 与 VisualTeX 1.0.6 不匹配，请执行“修复 Office 集成”。"}
          </span>
        </div>
      )}

      {message && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span>{message}</span>
        </div>
      )}

      <div className="office-settings-actions">
        <button
          type="button"
          className="primary-button"
          disabled={busy !== null}
          onClick={() => void run("install", "install_office_integration")}
        >
          <Download size={15} />
          {isEn ? "Install Office integration" : "安装 Office 集成"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("repair", "repair_office_integration")}
        >
          <Wrench size={15} />
          {isEn ? "Repair" : "修复 Office 集成"}
        </button>
        <button
          type="button"
          className="secondary-button danger-subtle"
          disabled={busy !== null}
          onClick={() => void run("uninstall", "uninstall_office_integration")}
        >
          <Trash2 size={15} />
          {isEn ? "Uninstall" : "卸载 Office 集成"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("certificate", "regenerate_office_certificate")}
        >
          <ShieldCheck size={15} />
          {isEn ? "Regenerate certificate" : "重新生成证书"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || Boolean(status?.companion.running)}
          onClick={() => void run("start", "start_office_companion")}
        >
          <Play size={15} />
          {isEn ? "Start companion" : "启动伴侣服务"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !status?.companion.running}
          onClick={() => void run("stop", "stop_office_companion")}
        >
          <Square size={14} />
          {isEn ? "Stop companion" : "停止伴侣服务"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("ocr", "install_ocr_runtime")}
        >
          <Download size={15} />
          {isEn ? "Install offline OCR" : "安装离线 OCR"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !status?.word.applicationInstalled}
          onClick={() => void run("word", "open_word")}
        >
          <ExternalLink size={15} />
          {isEn ? "Open Word" : "打开 Word"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !status?.powerpoint.applicationInstalled}
          onClick={() => void run("powerpoint", "open_powerpoint")}
        >
          <ExternalLink size={15} />
          {isEn ? "Open PowerPoint" : "打开 PowerPoint"}
        </button>
      </div>
    </section>
  );
}
