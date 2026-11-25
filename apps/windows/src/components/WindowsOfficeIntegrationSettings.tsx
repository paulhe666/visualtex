import { useCallback, useEffect, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  ChevronDown,
  CircleAlert,
  CircleHelp,
  Copy,
  Download,
  ExternalLink,
  FolderOpen,
  Play,
  RefreshCw,
  Search,
  Settings2,
  ShieldAlert,
  Square,
  ToggleLeft,
  ToggleRight,
  Trash2,
  Wrench,
  X,
} from "lucide-react";
import { useEditorStore } from "../../../shared/visualtex-app/src/stores/editorStore";
import {
  decodeBooleanStatus,
  decodeNativeRenderStatus,
  decodeOfficeCompanionStatus,
  decodeOfficePlatformStatus,
  decodeWpsIntegrationStatus,
  type NativeRenderStatus,
  type OfficeCompanionStatus,
  type OfficePlatformStatus,
  type WindowsOfficeMode,
  type WpsIntegrationStatus,
} from "./windowsOfficeStatusValidation";
import { wpsProblem } from "./wpsProblem";
import {
  decodeWpsDiagnosis,
  wpsDiagnosisAllFine,
  wpsDiagnosisLines,
  wpsReportItems,
  type WpsDiagnosis,
} from "./wpsDiagnosis";
import { hostLoadText, officeDiagnosticReport } from "./officeDiagnosticReport";
import { officeScriptReason } from "./officeScriptReason";

export type { WindowsOfficeMode } from "./windowsOfficeStatusValidation";

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
  const [wps, setWps] = useState<WpsIntegrationStatus | null>(null);
  // why WPS does not show VisualTeX, looked up when the WPS card reports a problem
  const [wpsDiagnosis, setWpsDiagnosis] = useState<WpsDiagnosis | null>(null);
  // the full WPS diagnosis with WPS's own answer (Advanced diagnostics > Check WPS)
  const [wpsCheck, setWpsCheck] = useState<WpsDiagnosis | null>(null);
  const [reportCopied, setReportCopied] = useState(false);
  const [renderer, setRenderer] = useState<NativeRenderStatus | null>(null);
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
          invoke<unknown>("get_office_platform_status"),
          invoke<unknown>("get_office_companion_status"),
          invoke<unknown>("get_mathtype_double_click_edit_enabled"),
        ]);
      setStatus(decodeOfficePlatformStatus(nextStatus));
      setCompanion(decodeOfficeCompanionStatus(nextCompanion));
      setMathTypeDoubleClickEditEnabled(
        decodeBooleanStatus(
          nextMathTypeDoubleClickEditEnabled,
          "mathTypeDoubleClickEditEnabled",
        ),
      );
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
    }
    // WPS and the renderer service are read separately: a failure there must not hide
    // the Office status above.
    try {
      const [nextWps, nextRenderer] = await Promise.all([
        invoke<unknown>("get_wps_integration_status"),
        invoke<unknown>("get_native_render_status"),
      ]);
      setWps(decodeWpsIntegrationStatus(nextWps));
      setRenderer(decodeNativeRenderStatus(nextRenderer));
    } catch (error) {
      setWps(null);
      setRenderer(null);
      setMessage((value) =>
        value ||
        errorMessage(
          error,
          isEn ? "Unable to read the WPS Office status." : "无法读取 WPS Office 状态。",
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
          // a repair or another change makes the last WPS check out of date
          setWpsCheck(null);
          await refresh();
        }
        return true;
      } catch (error) {
        const details = errorMessage(
          error,
          isEn ? "Windows Office operation failed." : "Windows Office 操作失败。",
        );
        // the script's own cause first, then everything it printed
        const reason = officeScriptReason(error);
        setMessage(reason ? `${reason}\n\n${details}` : details);
        return false;
      } finally {
        setBusy(null);
      }
    },
    [isEn, refresh],
  );

  const componentsHealthy = Boolean(
    status?.staticInstallVerified &&
      status.wordFilesPresent &&
      status.wordRegistryComplete &&
      status.wordLoadEnabled &&
      status.powerpointFilesPresent &&
      status.powerpointRegistryComplete &&
      status.powerpointLoadEnabled &&
      status.oleLocalServerHealthy,
  );
  // MathType's pictures name MT Extra and the Euclid fonts; without them (or
  // with Office's older MT Extra) Word shows wrong symbols. Repair installs them.
  const mathTypeFontsHealthy = Boolean(status?.mathtypeFontsHealthy);
  const installHealthy = componentsHealthy && mathTypeFontsHealthy;
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
        description: "",
      };
    }
    if (integrationReady) {
      return {
        title: isEn ? "Office integration is ready" : "Office 集成正常",
        description: "",
      };
    }
    if (verificationPending) {
      return {
        title: isEn
          ? "Installed, connection not verified"
          : "已安装，连接未验证",
        description: "",
      };
    }
    if (installationNeedsRepair && componentsHealthy && !mathTypeFontsHealthy) {
      return {
        title: isEn ? "MathType fonts incomplete" : "MathType 字体不完整",
        description: "",
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
        description: "",
      };
    }
    return {
      title: isEn ? "Office integration is not installed" : "未安装 Office 集成",
      description: "",
    };
  }, [
    componentsHealthy,
    hasInstalledComponents,
    installationNeedsRepair,
    integrationReady,
    isEn,
    mathTypeFontsHealthy,
    status,
    verificationPending,
  ]);

  const wpsIssue = wps ? wpsProblem(wps, isEn) : null;
  const wpsNeedsDiagnosis = Boolean(wps && wps.hostEnabled && wpsProblem(wps, false));
  useEffect(() => {
    setWpsDiagnosis(null);
    if (!wpsNeedsDiagnosis) return;
    let cancelled = false;
    void invoke<unknown>("diagnose_wps_integration")
      .then((value) => {
        if (!cancelled) setWpsDiagnosis(decodeWpsDiagnosis(value));
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [wps, wpsNeedsDiagnosis]);
  // the precise causes replace the summary; with none found, the summary says what is wrong and that the
  // registration, the whitelist and the load test are fine
  // Check WPS loads the add-in to hear WPS's answer, which renews its load record: the card then shows what
  // the check found instead of "Ready"
  const wpsCheckCauses = wpsCheck ? wpsDiagnosisLines(wpsCheck, isEn) : [];
  const wpsLines = (() => {
    if (!wpsIssue) return wpsCheckCauses;
    const diagnosis = wpsCheck ?? wpsDiagnosis;
    if (!diagnosis) return [wpsIssue.text];
    const causes = wpsDiagnosisLines(diagnosis, isEn);
    return causes.length > 0 ? causes : [wpsIssue.text, wpsDiagnosisAllFine(isEn)];
  })();
  const wpsCheckItems = wpsCheck ? wpsReportItems(wpsCheck, isEn) : [];
  const wpsCheckOk = wpsCheckItems.every((item) => item.ok);

  const diagnosticMessage =
    message ||
    (!verificationPending ? status?.lastError : null) ||
    (status && !status.mathtypeFontsHealthy ? status.mathtypeFontsProblem : null) ||
    companion?.lastError;

  // asks a hidden WPS too (several seconds); the add-ins' load records it updates are read again afterwards
  const checkWps = async () => {
    setBusy("wps-check");
    setMessage("");
    try {
      setWpsCheck(decodeWpsDiagnosis(await invoke<unknown>("diagnose_wps_integration", { askWps: true })));
    } catch (error) {
      setMessage(errorMessage(error, isEn ? "Unable to check WPS." : "无法检查 WPS。"));
    } finally {
      setBusy(null);
    }
    void refresh();
  };

  const copyReport = async () => {
    const text = officeDiagnosticReport({
      isEn,
      title: statusCopy.title,
      installHealthy,
      runtimeHealthy,
      status,
      companion,
      renderer,
      wps,
      wpsLines,
      wpsCheck,
      message: diagnosticMessage ?? "",
    });
    try {
      // Windows' clipboard directly: the page's clipboard refuses while the window is not focused
      await invoke("write_windows_ocr_clipboard_text", { text }).catch(() =>
        navigator.clipboard.writeText(text),
      );
      setReportCopied(true);
      window.setTimeout(() => setReportCopied(false), 1600);
    } catch (error) {
      setMessage(errorMessage(error, isEn ? "Unable to copy the report." : "无法复制报告。"));
    }
  };

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
          {statusCopy.description && <p>{statusCopy.description}</p>}
        </div>
      </div>

      <div className="office-mathtype-double-click-setting">
        <div>
          <strong>
            {isEn
              ? "Double-click MathType formulas with VisualTeX"
              : "双击 MathType 公式时使用 VisualTeX 编辑"}
          </strong>
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

      {renderer?.installed && !renderer.running && (
        <div className="office-mathtype-double-click-setting">
          <div>
            <strong>
              {isEn
                ? "The MathType renderer service is not running: MathType equations cannot be inserted or converted."
                : "MathType 公式渲染服务未运行，无法插入或转换 MathType 公式。"}
            </strong>
          </div>
          <button
            type="button"
            className="secondary-button office-mathtype-double-click-toggle"
            disabled={busy !== null}
            onClick={() => void run("render-start", "start_native_render_service")}
          >
            <Play size={15} />
            {busy === "render-start"
              ? isEn
                ? "Starting…"
                : "正在启动…"
              : isEn
                ? "Start service"
                : "启动服务"}
          </button>
        </div>
      )}

      {wps && (
        <article className="office-status-card">
          <header>
            <strong>WPS Office</strong>
            {wps.installations.length > 0 && (
              <StatusLine ok={wpsLines.length === 0}>
                {wpsLines.length === 0
                  ? isEn
                    ? "Ready"
                    : "正常"
                  : isEn
                    ? "Needs attention"
                    : "需要处理"}
              </StatusLine>
            )}
          </header>
          {wps.installations.length === 0 && (
            <p>{isEn ? "WPS Office was not found." : "未检测到 WPS Office。"}</p>
          )}
          {wps.installations.map((installation) => (
            <div key={installation.path} className="office-wps-installation">
              <p title={installation.path}>
                {`WPS Office ${installation.version} · ${
                  installation.architecture === "x86" ? "32" : "64"
                }${isEn ? "-bit" : " 位"} · ${
                  installation.perUser
                    ? isEn
                      ? "current user"
                      : "当前用户"
                    : isEn
                      ? "all users"
                      : "所有用户"
                }`}
              </p>
              <StatusLine ok={installation.writerRegistered}>
                {installation.writerRegistered
                  ? isEn
                    ? "WPS Writer add-in registered"
                    : "WPS 文字加载项已注册"
                  : isEn
                    ? "WPS Writer add-in not registered"
                    : "WPS 文字加载项未注册"}
              </StatusLine>
              <StatusLine ok={installation.presentationRegistered}>
                {installation.presentationRegistered
                  ? isEn
                    ? "WPS Presentation add-in registered"
                    : "WPS 演示加载项已注册"
                  : isEn
                    ? "WPS Presentation add-in not registered"
                    : "WPS 演示加载项未注册"}
              </StatusLine>
              <StatusLine
                ok={installation.writerWhitelisted && installation.presentationWhitelisted}
              >
                {installation.writerWhitelisted && installation.presentationWhitelisted
                  ? isEn
                    ? "In the WPS add-in whitelist"
                    : "已加入 WPS 加载项白名单"
                  : isEn
                    ? "Not in the WPS add-in whitelist"
                    : "未加入 WPS 加载项白名单"}
              </StatusLine>
              <StatusLine ok={installation.formulaOleRegistered}>
                {isEn ? "Formula OLE service" : "公式 OLE 服务"}
              </StatusLine>
            </div>
          ))}
          {wps.installations.length > 0 && (
            <>
              <StatusLine ok={Boolean(wps.writerLoaded?.enabled)}>
                {hostLoadText(wps.writerLoaded, isEn ? "WPS Writer" : "WPS 文字", isEn)}
              </StatusLine>
              <StatusLine ok={Boolean(wps.presentationLoaded?.enabled)}>
                {hostLoadText(
                  wps.presentationLoaded,
                  isEn ? "WPS Presentation" : "WPS 演示",
                  isEn,
                )}
              </StatusLine>
              <div className="office-mathtype-double-click-setting">
                <div>
                  <strong>{isEn ? "Use VisualTeX in WPS" : "在 WPS 中使用 VisualTeX"}</strong>
                </div>
                <button
                  type="button"
                  className="secondary-button office-mathtype-double-click-toggle"
                  disabled={busy !== null}
                  aria-pressed={wps.hostEnabled}
                  onClick={() =>
                    void run("wps-host", "set_wps_host_enabled", {
                      enabled: !wps.hostEnabled,
                    })
                  }
                >
                  {wps.hostEnabled ? <ToggleRight size={16} /> : <ToggleLeft size={16} />}
                  {wps.hostEnabled
                    ? isEn
                      ? "Enabled"
                      : "已开启"
                    : isEn
                      ? "Disabled"
                      : "已关闭"}
                </button>
              </div>
              <details className="office-wps-help">
                <summary>
                  {isEn ? "VisualTeX tab not showing in WPS?" : "WPS 中看不到 VisualTeX 选项卡？"}
                </summary>
                <p>
                  {isEn
                    ? "WPS can turn off add-ins it thinks slow down its startup. Turn VisualTeX back on:"
                    : "WPS 会停用它认为拖慢启动速度的加载项。按以下步骤重新启用 VisualTeX："}
                </p>
                <ol>
                  <li>
                    {isEn
                      ? "In WPS Writer or WPS Presentation, open any document."
                      : "在 WPS 文字或 WPS 演示中打开任意文档。"}
                  </li>
                  <li>
                    {isEn ? "Choose Tools > COM Add-ins." : "点击“工具”选项卡中的“COM 加载项”。"}
                  </li>
                  <li>
                    {isEn
                      ? "Click the link in the notice about disabled third-party add-ins."
                      : "点击提示“此电脑已禁用可能拖慢启动速度的第三方加载项”中的链接。"}
                  </li>
                  <li>
                    {isEn ? "Turn VisualTeX on." : "找到 VisualTeX，将其启用。"}
                  </li>
                  <li>
                    {isEn
                      ? "Close WPS and open it again."
                      : "关闭 WPS 后重新打开。"}
                  </li>
                </ol>
              </details>
            </>
          )}
          {wpsLines.length > 0 && (
            <div className="office-settings-warning" role="alert">
              <ShieldAlert size={15} />
              <span className="office-settings-diagnostic">
                {wpsLines.join("\n")}
              </span>
            </div>
          )}
          {wpsIssue?.repair && (
            <div className="office-primary-actions">
              <button
                type="button"
                className="primary-button office-action-main"
                disabled={busy !== null}
                onClick={() => void run("repair", "repair_windows_office_integration")}
              >
                <Wrench size={16} />
                {busy === "repair"
                  ? isEn
                    ? "Repairing…"
                    : "正在修复…"
                  : isEn
                    ? "Repair WPS add-ins"
                    : "修复 WPS 加载项"}
              </button>
            </div>
          )}
        </article>
      )}

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
              <StatusLine ok={mathTypeFontsHealthy}>
                {isEn ? "MathType fonts" : "MathType 字体"}
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
              <StatusLine ok={Boolean(renderer?.running)}>
                {isEn ? "MathType renderer service" : "MathType 渲染服务"}
              </StatusLine>
            </article>

            {wps && wps.installations.length > 0 && (
              <article className="office-status-card office-status-card-wide">
                <header>
                  <strong>WPS</strong>
                  {wpsCheck ? (
                    <StatusLine ok={wpsCheckOk}>
                      {wpsCheckOk
                        ? isEn
                          ? "Ready"
                          : "正常"
                        : isEn
                          ? "Needs attention"
                          : "需要处理"}
                    </StatusLine>
                  ) : (
                    <span className="office-status-unchecked">
                      {isEn ? "Not checked" : "未检查"}
                    </span>
                  )}
                </header>
                {wpsCheckItems.map((item, index) => (
                  <StatusLine key={index} ok={item.ok}>
                    {item.text}
                  </StatusLine>
                ))}
              </article>
            )}
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
            {wps && wps.installations.length > 0 && (
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => void checkWps()}
              >
                <Search size={15} />
                {busy === "wps-check"
                  ? isEn
                    ? "Checking WPS…"
                    : "正在检查 WPS…"
                  : isEn
                    ? "Check WPS"
                    : "检查 WPS"}
              </button>
            )}
            <button
              type="button"
              className="secondary-button"
              disabled={busy === "wps-check"}
              onClick={() => void copyReport()}
            >
              <Copy size={15} />
              {reportCopied
                ? isEn
                  ? "Copied"
                  : "已复制"
                : isEn
                  ? "Copy report"
                  : "复制报告"}
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
                    ? "Word and PowerPoint will be started to verify the add-ins. Save and close all Office documents first."
                    : "将启动 Word 和 PowerPoint 验证加载项。请先保存并关闭所有 Office 文档。"}
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
                <span>
                  {isEn ? "Unsaved changes will be lost." : "未保存的内容将丢失。"}
                </span>
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
                      ? "Verify"
                      : "开始验证"}
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
                    ? "Removes the VisualTeX add-ins from Word and PowerPoint. Formulas in documents are kept."
                    : "将从 Word 和 PowerPoint 移除 VisualTeX 加载项。文档中的公式保留。"}
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
