import { useCallback, useEffect, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  ChevronDown,
  CircleAlert,
  CircleHelp,
  Download,
  ExternalLink,
  FolderOpen,
  Play,
  RefreshCw,
  Settings2,
  ShieldAlert,
  Square,
  ToggleLeft,
  ToggleRight,
  Trash2,
  Wrench,
  X,
} from "lucide-react";
import { useEditorStore } from "../stores/editorStore";

export type WindowsOfficeMode = "auto" | "vsto";

interface OfficePlatformStatus {
  platform: string;
  mode: WindowsOfficeMode;
  activeBackend: string;
  oleBridgeHealthy: boolean;
  oleLocalServerHealthy: boolean;
  staticInstallVerified: boolean;
  wordFilesPresent: boolean;
  wordRegistryComplete: boolean;
  wordLoadEnabled: boolean;
  powerpointFilesPresent: boolean;
  powerpointRegistryComplete: boolean;
  powerpointLoadEnabled: boolean;
  vstoWordHealthy: boolean;
  vstoPowerpointHealthy: boolean;
  wordConnected: boolean;
  powerpointConnected: boolean;
  connectionVerificationAttempted: boolean;
  companionProcessRunning: boolean;
  companionPortListening: boolean;
  companionHttpsHealthy: boolean;
  companionCertificateMatches: boolean;
  companionProtocolMatches: boolean;
  officeRuntimeVerified: boolean;
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
  const [confirmRuntimeTest, setConfirmRuntimeTest] = useState(false);
  const [forceCloseOffice, setForceCloseOffice] = useState(false);
  const [confirmUninstall, setConfirmUninstall] = useState(false);
  const [mathTypeDoubleClickEditEnabled, setMathTypeDoubleClickEditEnabled] =
    useState(true);

  const refresh = useCallback(async () => {
    setBusy((value) => value ?? "refresh");
    try {
      const [nextStatus, nextCompanion, nextMathTypeDoubleClickEditEnabled] =
        await Promise.all([
          invoke<OfficePlatformStatus>("get_office_platform_status"),
          invoke<OfficeCompanionStatus>("get_office_companion_status"),
          invoke<boolean>("get_mathtype_double_click_edit_enabled"),
        ]);
      setStatus(nextStatus);
      setCompanion(nextCompanion);
      setMathTypeDoubleClickEditEnabled(nextMathTypeDoubleClickEditEnabled);
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
        if (
          command !== "open_word" &&
          command !== "open_powerpoint" &&
          command !== "open_windows_office_logs"
        ) {
          await refresh();
        }
        return true;
      } catch (error) {
        setMessage(
          errorMessage(
            error,
            isEn ? "Windows Office operation failed." : "Windows Office 操作失败。",
          ),
        );
        return false;
      } finally {
        setBusy(null);
      }
    },
    [isEn, refresh],
  );

  const installHealthy = Boolean(
    status?.staticInstallVerified &&
      status.wordFilesPresent &&
      status.wordRegistryComplete &&
      status.wordLoadEnabled &&
      status.powerpointFilesPresent &&
      status.powerpointRegistryComplete &&
      status.powerpointLoadEnabled &&
      status.oleLocalServerHealthy,
  );
  const runtimeHealthy = Boolean(status?.officeRuntimeVerified);
  const officeConnectionsVerified = Boolean(
    status?.wordConnected && status?.powerpointConnected,
  );
  const integrationReady = Boolean(
    installHealthy &&
      runtimeHealthy &&
      officeConnectionsVerified &&
      !status?.lastError,
  );
  const hasInstalledComponents = Boolean(
    status?.wordFilesPresent ||
      status?.powerpointFilesPresent ||
      status?.wordRegistryComplete ||
      status?.powerpointRegistryComplete ||
      status?.oleLocalServerHealthy,
  );
  const verificationPending = Boolean(
    installHealthy &&
      runtimeHealthy &&
      !officeConnectionsVerified &&
      !status?.connectionVerificationAttempted &&
      !status?.lastError,
  );
  const installationNeedsRepair = hasInstalledComponents && !installHealthy;
  const statusCopy = useMemo(() => {
    if (!status) {
      return {
        title: isEn ? "Checking Office integration…" : "正在检查 Office 集成…",
        description: isEn
          ? "VisualTeX is reading the Word and PowerPoint add-in status."
          : "VisualTeX 正在读取 Word 和 PowerPoint 加载项状态。",
      };
    }
    if (integrationReady) {
      return {
        title: isEn ? "Office integration is ready" : "Office 集成可正常使用",
        description: isEn
          ? "Word and PowerPoint can create and edit VisualTeX formulas."
          : "Word 和 PowerPoint 已可创建、插入和编辑 VisualTeX 公式。",
      };
    }
    if (verificationPending) {
      return {
        title: isEn
          ? "Office integration is installed"
          : "Office 集成已安装，等待连接验证",
        description: isEn
          ? "Start Word and PowerPoint once to verify that both add-ins connect successfully."
          : "启动一次 Word 和 PowerPoint 验证加载项连接后，即可完成全部检查。",
      };
    }
    if (hasInstalledComponents) {
      return {
        title: installationNeedsRepair
          ? isEn
            ? "Office integration needs repair"
            : "Office 集成需要修复"
          : isEn
            ? "Office connection verification needs attention"
            : "Office 连接验证未通过",
        description: installationNeedsRepair
          ? isEn
            ? "Some installation components are incomplete. Repair the integration and try again."
            : "部分安装组件不完整，请修复 Office 集成后重试。"
          : isEn
            ? "Close Office and run the connection verification again."
            : "请关闭所有 Office 应用后重新验证连接，必要时可选择强制关闭。",
      };
    }
    return {
      title: isEn ? "Office integration is not installed" : "尚未安装 Office 集成",
      description: isEn
        ? "Install once to add VisualTeX tools to Word and PowerPoint."
        : "安装后即可在 Word 和 PowerPoint 中直接使用 VisualTeX。",
    };
  }, [
    hasInstalledComponents,
    installationNeedsRepair,
    integrationReady,
    isEn,
    status,
    verificationPending,
  ]);

  const diagnosticMessage =
    message ||
    (!verificationPending ? status?.lastError : null) ||
    companion?.lastError;

  const openRuntimeVerification = () => {
    setMessage("");
    setForceCloseOffice(false);
    setConfirmRuntimeTest(true);
  };

  const verifyRuntime = async () => {
    const succeeded = await run(
      "runtime-test",
      "test_windows_office_runtime",
      { forceCloseOffice },
    );
    if (succeeded) {
      setConfirmRuntimeTest(false);
      setForceCloseOffice(false);
    }
  };

  const uninstall = async () => {
    const succeeded = await run(
      "uninstall-ole",
      "uninstall_windows_ole_integration",
    );
    if (succeeded) setConfirmUninstall(false);
  };

  return (
    <section className="settings-section office-integration-section">
      <div className="office-settings-heading office-settings-heading-simple">
        <div>
          <h3>{isEn ? "Office integration" : "Office 集成"}</h3>
          <p>
            {isEn
              ? "Use VisualTeX directly in Microsoft Word and PowerPoint."
              : "在 Microsoft Word 和 PowerPoint 中直接插入和编辑公式。"}
          </p>
        </div>
        <button
          type="button"
          className="icon-button compact"
          onClick={() => void refresh()}
          disabled={busy !== null}
          aria-label={isEn ? "Refresh Office status" : "刷新 Office 状态"}
          title={isEn ? "Refresh status" : "刷新状态"}
        >
          <RefreshCw
            size={15}
            className={busy === "refresh" ? "is-spinning" : ""}
          />
        </button>
      </div>

      <div
        className={`office-summary-card ${
          integrationReady
            ? "is-ready"
            : verificationPending
              ? "is-pending"
              : hasInstalledComponents
                ? "needs-attention"
                : "is-not-installed"
        }`}
      >
        <span className="office-summary-icon" aria-hidden="true">
          {integrationReady ? (
            <CheckCircle2 size={22} />
          ) : verificationPending ? (
            <CircleHelp size={22} />
          ) : (
            <CircleAlert size={22} />
          )}
        </span>
        <div>
          <strong>{statusCopy.title}</strong>
          <p>{statusCopy.description}</p>
        </div>
      </div>

      <div className="office-mathtype-double-click-setting">
        <div>
          <strong>
            {isEn
              ? "Double-click MathType formulas with VisualTeX"
              : "双击 MathType 公式时使用 VisualTeX 编辑"}
          </strong>
          <p>
            {isEn
              ? "When enabled, double-clicking a MathType OLE formula opens VisualTeX instead of the native MathType editor. Disable it to compare or edit the same object directly in MathType."
              : "开启后，双击 MathType OLE 公式会优先进入 VisualTeX，而不是启动 MathType 原生编辑器；关闭后可直接用 MathType 验收或编辑同一个对象。"}
          </p>
        </div>
        <button
          type="button"
          className="secondary-button office-mathtype-double-click-toggle"
          disabled={busy !== null}
          aria-pressed={mathTypeDoubleClickEditEnabled}
          onClick={() =>
            void run(
              "mathtype-double-click",
              "set_mathtype_double_click_edit_enabled",
              { enabled: !mathTypeDoubleClickEditEnabled },
            )
          }
        >
          {mathTypeDoubleClickEditEnabled ? (
            <ToggleRight size={16} />
          ) : (
            <ToggleLeft size={16} />
          )}
          {mathTypeDoubleClickEditEnabled
            ? isEn
              ? "Enabled"
              : "已开启"
            : isEn
              ? "Disabled"
              : "已关闭"}
        </button>
      </div>

      {diagnosticMessage && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span className="office-settings-diagnostic">
            {diagnosticMessage}
          </span>
        </div>
      )}

      <div className="office-primary-actions">
        {!integrationReady && !hasInstalledComponents && (
          <button
            type="button"
            className="primary-button office-action-main"
            disabled={busy !== null}
            onClick={() =>
              void run("install-ole", "install_windows_ole_integration")
            }
          >
            <Download size={16} />
            {busy === "install-ole"
              ? isEn
                ? "Installing…"
                : "正在安装…"
              : isEn
                ? "Install Office integration"
                : "安装 Office 集成"}
          </button>
        )}
        {installationNeedsRepair && (
          <button
            type="button"
            className="primary-button office-action-main"
            disabled={busy !== null}
            onClick={() =>
              void run("repair", "repair_windows_office_integration")
            }
          >
            <Wrench size={16} />
            {busy === "repair"
              ? isEn
                ? "Repairing…"
                : "正在修复…"
              : isEn
                ? "Repair integration"
                : "修复 Office 集成"}
          </button>
        )}
        {hasInstalledComponents && installHealthy && !integrationReady && (
          <button
            type="button"
            className="primary-button office-action-main"
            disabled={busy !== null}
            onClick={openRuntimeVerification}
          >
            <CheckCircle2 size={16} />
            {busy === "runtime-test"
              ? isEn
                ? "Verifying…"
                : "正在验证…"
              : status?.connectionVerificationAttempted
                ? isEn
                  ? "Verify Office connection again"
                  : "重新验证 Office 连接"
                : isEn
                  ? "Verify Office connection"
                  : "验证 Office 连接"}
          </button>
        )}
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("word", "open_word")}
        >
          <ExternalLink size={15} />
          {isEn ? "Open Word" : "打开 Word"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("powerpoint", "open_powerpoint")}
        >
          <ExternalLink size={15} />
          {isEn ? "Open PowerPoint" : "打开 PowerPoint"}
        </button>
      </div>

      <details className="office-advanced-settings">
        <summary>
          <span>
            <Settings2 size={15} />
            {isEn ? "Advanced diagnostics" : "高级诊断与维护"}
          </span>
          <ChevronDown size={15} className="office-details-chevron" />
        </summary>

        <div className="office-advanced-content">
          <div className="office-status-grid office-status-grid-compact">
            <article className="office-status-card">
              <header>
                <strong>{isEn ? "Installation" : "安装状态"}</strong>
                <StatusLine ok={installHealthy}>
                  {installHealthy
                    ? isEn
                      ? "Complete"
                      : "完整"
                    : isEn
                      ? "Incomplete"
                      : "不完整"}
                </StatusLine>
              </header>
              <StatusLine ok={Boolean(status?.wordFilesPresent)}>
                {isEn ? "Word add-in files" : "Word 加载项文件"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordRegistryComplete)}>
                {isEn ? "Word registration" : "Word 注册信息"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordLoadEnabled)}>
                {isEn ? "Word add-in enabled" : "Word 加载项已启用"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointFilesPresent)}>
                {isEn ? "PowerPoint add-in files" : "PowerPoint 加载项文件"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointRegistryComplete)}>
                {isEn ? "PowerPoint registration" : "PowerPoint 注册信息"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointLoadEnabled)}>
                {isEn ? "PowerPoint add-in enabled" : "PowerPoint 加载项已启用"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.oleLocalServerHealthy)}>
                {isEn ? "Formula OLE service" : "公式 OLE 服务"}
              </StatusLine>
            </article>

            <article className="office-status-card">
              <header>
                <strong>{isEn ? "Runtime" : "运行状态"}</strong>
                <StatusLine ok={runtimeHealthy}>
                  {runtimeHealthy
                    ? isEn
                      ? "Available"
                      : "可用"
                    : isEn
                      ? "Unavailable"
                      : "不可用"}
                </StatusLine>
              </header>
              <StatusLine ok={Boolean(status?.companionProcessRunning)}>
                {isEn ? "Companion process" : "伴侣进程"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionPortListening)}>
                {isEn ? "Local port" : "本地端口"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionHttpsHealthy)}>
                {isEn ? "Local HTTPS connection" : "本地 HTTPS 连接"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionCertificateMatches)}>
                {isEn ? "Certificate" : "证书"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionProtocolMatches)}>
                {isEn ? "Protocol version" : "协议版本"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordConnected)}>
                {isEn ? "Word connection" : "Word 连接"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointConnected)}>
                {isEn ? "PowerPoint connection" : "PowerPoint 连接"}
              </StatusLine>
            </article>
          </div>

          <div className="office-diagnostic-meta">
            <span>
              {isEn ? "Backend" : "后端"}: {status?.activeBackend ?? "—"}
            </span>
            <span>
              {isEn ? "Companion" : "伴侣服务"}: {companion?.running
                ? `${companion.bindAddress}:${companion.port}`
                : isEn
                  ? "Stopped"
                  : "已停止"}
            </span>
          </div>

          <div className="office-secondary-actions">
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || !hasInstalledComponents}
              onClick={openRuntimeVerification}
            >
              <CheckCircle2 size={15} />
              {isEn ? "Verify Office connection" : "验证 Office 连接"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || !hasInstalledComponents}
              onClick={() =>
                void run("repair", "repair_windows_office_integration")
              }
            >
              <Wrench size={15} />
              {isEn ? "Repair integration" : "修复 Office 集成"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null}
              onClick={() => void run("open-logs", "open_windows_office_logs")}
            >
              <FolderOpen size={15} />
              {isEn ? "Open logs" : "打开诊断日志"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null}
              onClick={() =>
                void run("background-start", "set_office_background_start", {
                  enabled: !status?.backgroundStartEnabled,
                })
              }
            >
              {status?.backgroundStartEnabled ? (
                <ToggleRight size={15} />
              ) : (
                <ToggleLeft size={15} />
              )}
              {status?.backgroundStartEnabled
                ? isEn
                  ? "Disable startup"
                  : "关闭开机启动"
                : isEn
                  ? "Enable startup"
                  : "启用开机启动"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || Boolean(companion?.running)}
              onClick={() => void run("start", "start_office_companion")}
            >
              <Play size={15} />
              {isEn ? "Start service" : "启动服务"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || !companion?.running}
              onClick={() => void run("stop", "stop_office_companion")}
            >
              <Square size={14} />
              {isEn ? "Stop service" : "停止服务"}
            </button>
            <button
              type="button"
              className="secondary-button danger-subtle"
              disabled={busy !== null || !hasInstalledComponents}
              onClick={() => setConfirmUninstall(true)}
            >
              <Trash2 size={15} />
              {isEn ? "Uninstall integration" : "卸载 Office 集成"}
            </button>
          </div>
        </div>
      </details>

      {confirmRuntimeTest && (
        <div
          className="office-confirm-backdrop"
          role="presentation"
          onMouseDown={() => {
            if (busy !== null) return;
            setConfirmRuntimeTest(false);
            setForceCloseOffice(false);
          }}
        >
          <section
            className="office-confirm-dialog office-runtime-confirm-dialog"
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="office-runtime-test-title"
            aria-describedby="office-runtime-test-description"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <header>
              <span
                className="office-confirm-icon is-verification"
                aria-hidden="true"
              >
                <CheckCircle2 size={19} />
              </span>
              <div>
                <strong id="office-runtime-test-title">
                  {isEn ? "Verify Office connection" : "验证 Office 连接"}
                </strong>
                <p id="office-runtime-test-description">
                  {isEn
                    ? "VisualTeX will start Word and PowerPoint briefly to verify that both add-ins connect. Save your documents and close every Office application before continuing."
                    : "VisualTeX 将临时启动 Word 和 PowerPoint，检查两个加载项是否真正连接。继续前请保存文档，并关闭所有正在运行的 Office 应用。"}
                </p>
              </div>
              <button
                type="button"
                className="icon-button compact"
                disabled={busy !== null}
                onClick={() => {
                  setConfirmRuntimeTest(false);
                  setForceCloseOffice(false);
                }}
                aria-label={isEn ? "Cancel verification" : "取消验证"}
              >
                <X size={16} />
              </button>
            </header>

            <label className="office-force-close-option">
              <input
                type="checkbox"
                checked={forceCloseOffice}
                disabled={busy !== null}
                onChange={(event) => setForceCloseOffice(event.target.checked)}
              />
              <span>
                <strong>
                  {isEn
                    ? "Force-close running Word and PowerPoint"
                    : "强制关闭正在运行的 Word 和 PowerPoint"}
                </strong>
                <small>
                  {isEn
                    ? "Use only after saving. Unsaved Office changes may be lost."
                    : "仅在保存文档后使用；未保存的 Office 内容可能丢失。"}
                </small>
              </span>
            </label>

            {message && (
              <div className="office-confirm-inline-warning" role="alert">
                <ShieldAlert size={15} />
                <span>{message}</span>
              </div>
            )}

            <footer>
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => {
                  setConfirmRuntimeTest(false);
                  setForceCloseOffice(false);
                }}
              >
                {isEn ? "Cancel" : "取消"}
              </button>
              <button
                type="button"
                className="primary-button"
                disabled={busy !== null}
                onClick={() => void verifyRuntime()}
              >
                <CheckCircle2 size={15} />
                {busy === "runtime-test"
                  ? isEn
                    ? "Verifying…"
                    : "正在验证…"
                  : forceCloseOffice
                    ? isEn
                      ? "Force close and verify"
                      : "强制关闭并验证"
                    : isEn
                      ? "I have closed Office; verify"
                      : "我已关闭 Office，开始验证"}
              </button>
            </footer>
          </section>
        </div>
      )}

      {confirmUninstall && (
        <div
          className="office-confirm-backdrop"
          role="presentation"
          onMouseDown={() => busy === null && setConfirmUninstall(false)}
        >
          <section
            className="office-confirm-dialog"
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="office-uninstall-title"
            aria-describedby="office-uninstall-description"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <header>
              <span className="office-confirm-icon" aria-hidden="true">
                <Trash2 size={19} />
              </span>
              <div>
                <strong id="office-uninstall-title">
                  {isEn ? "Uninstall Office integration?" : "确定卸载 Office 集成？"}
                </strong>
                <p id="office-uninstall-description">
                  {isEn
                    ? "This removes the VisualTeX add-ins and OLE registration from Word and PowerPoint. Your formulas and VisualTeX documents will not be deleted."
                    : "这会从 Word 和 PowerPoint 中移除 VisualTeX 加载项及 OLE 注册，但不会删除已有公式或 VisualTeX 文档。"}
                </p>
              </div>
              <button
                type="button"
                className="icon-button compact"
                disabled={busy !== null}
                onClick={() => setConfirmUninstall(false)}
                aria-label={isEn ? "Cancel uninstall" : "取消卸载"}
              >
                <X size={16} />
              </button>
            </header>
            <footer>
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => setConfirmUninstall(false)}
              >
                {isEn ? "Cancel" : "取消"}
              </button>
              <button
                type="button"
                className="danger-button"
                disabled={busy !== null}
                onClick={() => void uninstall()}
              >
                <Trash2 size={15} />
                {busy === "uninstall-ole"
                  ? isEn
                    ? "Uninstalling…"
                    : "正在卸载…"
                  : isEn
                    ? "Uninstall"
                    : "确认卸载"}
              </button>
            </footer>
          </section>
        </div>
      )}
    </section>
  );
}
