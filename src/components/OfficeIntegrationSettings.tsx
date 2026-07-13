import { useCallback, useEffect, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open } from "@tauri-apps/plugin-dialog";
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

interface OcrRuntimeStatus {
  installed: boolean;
  pythonPath?: string | null;
  pythonVersion?: string | null;
  paddleVersion?: string | null;
  paddleocrVersion?: string | null;
  runtimePath?: string | null;
  offlineBundleAvailable?: boolean;
  installedModels?: string[];
  defaultModel?: string;
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
  | "ocr-model-install"
  | "ocr-model-remove"
  | "word"
  | "powerpoint";

const EXPECTED_MANIFEST_VERSION = "1.0.16.0";

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

  const installOptionalModel = useCallback(async () => {
    setBusy("ocr-model-install");
    setMessage("");
    try {
      const selected = await open({
        multiple: false,
        directory: false,
        filters: [
          {
            name: "VisualTeX OCR Model Pack",
            extensions: ["vtxocrmodel"],
          },
        ],
      });
      const packagePath =
        typeof selected === "string"
          ? selected
          : Array.isArray(selected)
            ? selected[0]
            : null;
      if (!packagePath) return;
      const next = await invoke<OcrRuntimeStatus>("install_optional_ocr_model", {
        packagePath,
      });
      setOcrStatus(next);
      setMessage(
        isEn
          ? "Optional OCR model installed and verified."
          : "可选 OCR 模型已完成校验并安装。",
      );
    } catch (error) {
      setMessage(
        operationErrorMessage(
          error,
          isEn ? "Unable to install OCR model pack" : "无法安装 OCR 模型包",
        ),
      );
    } finally {
      setBusy(null);
    }
  }, [isEn]);

  const removeOptionalModel = useCallback(
    async (model: string) => {
      const confirmed = window.confirm(
        isEn
          ? `Remove the optional offline model ${model}?`
          : `确定卸载可选离线模型 ${model} 吗？`,
      );
      if (!confirmed) return;
      setBusy("ocr-model-remove");
      setMessage("");
      try {
        const next = await invoke<OcrRuntimeStatus>("remove_optional_ocr_model", {
          model,
        });
        setOcrStatus(next);
        setMessage(
          isEn ? "Optional OCR model removed." : "可选 OCR 模型已卸载。",
        );
      } catch (error) {
        setMessage(
          operationErrorMessage(
            error,
            isEn ? "Unable to remove OCR model" : "无法卸载 OCR 模型",
          ),
        );
      } finally {
        setBusy(null);
      }
    },
    [isEn],
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
              ? "Install the local Word and PowerPoint bridge, certificate and companion service. In PowerPoint, double-click a VisualTeX formula to edit it."
              : "安装本地 Word/PowerPoint 桥接插件、证书和伴侣服务；在 PowerPoint 中双击 VisualTeX 公式即可重新编辑。"}
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
              <strong>{isEn ? "Login background service" : "登录后台服务"}</strong>
              <StateMark ok={status.background.installed} />
            </header>
            <p>
              {status.background.runningInBackgroundMode
                ? isEn
                  ? "This process was started by the Office background LaunchAgent."
                  : "当前进程由 Office 后台 LaunchAgent 启动。"
                : status.background.installed
                  ? isEn
                    ? "The current app is serving Office now; a hidden instance takes over after quit and starts at login."
                    : "当前应用正在提供 Office 服务；退出后由隐藏实例接管，并在登录时自动启动。"
                  : isEn
                    ? "Not configured"
                    : "尚未配置"}
            </p>
            <dl>
              <div>
                <dt>LaunchAgent</dt>
                <dd>
                  {status.background.installed
                    ? isEn
                      ? "Installed"
                      : "已安装"
                    : isEn
                      ? "Missing"
                      : "缺失"}
                </dd>
              </div>
              <div>
                <dt>{isEn ? "Current login session" : "当前登录会话"}</dt>
                <dd>
                  {status.background.loaded
                    ? isEn
                      ? "Loaded"
                      : "已加载"
                    : status.background.installed
                      ? isEn
                        ? "Handoff on quit"
                        : "退出后接管"
                      : isEn
                        ? "Not loaded"
                        : "未加载"}
                </dd>
              </div>
              <div>
                <dt>{isEn ? "Configuration" : "配置文件"}</dt>
                <dd title={status.background.plistPath}>
                  {status.background.plistPath || "—"}
                </dd>
              </div>
              {status.background.lastError && (
                <div>
                  <dt>{isEn ? "Error" : "错误"}</dt>
                  <dd title={status.background.lastError}>
                    {status.background.lastError}
                  </dd>
                </div>
              )}
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
                  ? "The self-contained Python, PaddleOCR and installed formula models are ready without a network connection."
                  : "自包含 Python、PaddleOCR 与已安装公式模型均已就绪，断网也可使用。"
                : ocrStatus?.offlineBundleAvailable
                  ? isEn
                    ? "The complete offline package is bundled with VisualTeX and is ready to install locally."
                    : "完整离线包已随 VisualTeX 内置，可直接在本机安装。"
                  : isEn
                    ? "The offline OCR package is missing from this application build."
                    : "当前应用包缺少离线 OCR 资源。"}
            </p>
            <dl>
              <div>
                <dt>{isEn ? "Offline package" : "离线资源包"}</dt>
                <dd>
                  {ocrStatus?.offlineBundleAvailable
                    ? isEn
                      ? "Bundled"
                      : "已内置"
                    : isEn
                      ? "Missing"
                      : "缺失"}
                </dd>
              </div>
              <div>
                <dt>{ocrStatus?.defaultModel ?? "PP-FormulaNet_plus-M"}</dt>
                <dd>
                  {ocrStatus?.installedModels?.includes(
                    ocrStatus?.defaultModel ?? "PP-FormulaNet_plus-M",
                  )
                    ? isEn
                      ? "Installed"
                      : "已安装"
                    : isEn
                      ? "Ready in package"
                      : "已包含在资源包中"}
                </dd>
              </div>
              <div>
                <dt>{isEn ? "Installed models" : "已安装模型"}</dt>
                <dd>
                  {ocrStatus?.installedModels?.length
                    ? ocrStatus.installedModels.join(", ")
                    : "—"}
                </dd>
              </div>
              <div>
                <dt>Python</dt>
                <dd>{ocrStatus?.pythonVersion ?? "—"}</dd>
              </div>
            </dl>
            <div className="office-model-pack-actions">
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => void installOptionalModel()}
              >
                <Download size={14} />
                {isEn ? "Import S/L model pack" : "导入 S/L 模型包"}
              </button>
              {ocrStatus?.installedModels
                ?.filter(
                  (model) =>
                    model !==
                    (ocrStatus.defaultModel ?? "PP-FormulaNet_plus-M"),
                )
                .map((model) => (
                  <button
                    type="button"
                    className="secondary-button danger-subtle"
                    disabled={busy !== null}
                    onClick={() => void removeOptionalModel(model)}
                    key={model}
                  >
                    <Trash2 size={14} />
                    {isEn ? `Remove ${model}` : `卸载 ${model}`}
                  </button>
                ))}
            </div>
          </article>
        </div>
      )}

      {manifestMismatch && (
        <div className="office-settings-warning">
          <ShieldAlert size={15} />
          <span>
            {isEn
              ? "The installed manifest does not match VisualTeX 1.0.16. Use Repair Office Integration."
              : "已安装 manifest 与 VisualTeX 1.0.16 不匹配，请执行“修复 Office 集成”。"}
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
          disabled={busy !== null || ocrStatus?.offlineBundleAvailable === false}
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
