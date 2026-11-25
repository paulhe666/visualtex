import { APP_VERSION } from "../platform/appInfo";
import type {
  NativeRenderStatus,
  OfficeCompanionStatus,
  OfficePlatformStatus,
  WpsHostLoad,
  WpsIntegrationStatus,
} from "./windowsOfficeStatusValidation";
import { wpsReportItems, type WpsDiagnosis } from "./wpsDiagnosis";

/** "WPS 文字已加载 VisualTeX · <time>", or that it has not (the WPS card and the report). */
export function hostLoadText(load: WpsHostLoad | null, host: string, isEn: boolean) {
  if (!load) return isEn ? `${host} has not loaded VisualTeX` : `${host}尚未加载 VisualTeX`;
  const off = load.enabled ? "" : isEn ? " (turned off)" : "（未启用）";
  return isEn
    ? `${host} loaded VisualTeX${off} · ${load.at}`
    : `${host}已加载 VisualTeX${off} · ${load.at}`;
}

export interface OfficeDiagnosticReportInput {
  isEn: boolean;
  /** the headline of Settings > Office */
  title: string;
  installHealthy: boolean;
  runtimeHealthy: boolean;
  status: OfficePlatformStatus | null;
  companion: OfficeCompanionStatus | null;
  renderer: NativeRenderStatus | null;
  wps: WpsIntegrationStatus | null;
  /** the WPS card's problem lines */
  wpsLines: string[];
  /** the last Check WPS */
  wpsCheck: WpsDiagnosis | null;
  /** the page's error message */
  message: string;
}

function timestamp(date: Date) {
  const pad = (value: number) => String(value).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

/** Settings > Office as plain text to paste into a message: what the page shows, the WPS check included. */
export function officeDiagnosticReport(input: OfficeDiagnosticReportInput, now = new Date()): string {
  const { isEn, status, companion, renderer, wps } = input;
  const colon = isEn ? ": " : "：";
  const mark = (ok: boolean) => (ok ? "[ok] " : "[!]  ");
  const label = (en: string, zh: string) => (isEn ? en : zh);
  const lines: string[] = [];
  lines.push(`VisualTeX ${APP_VERSION} Windows · ${label("Office diagnostic report", "Office 诊断报告")} · ${timestamp(now)}`);
  lines.push(`${label("Office integration", "Office 集成")}${colon}${input.title}`);

  lines.push(`${label("Installation", "安装状态")}${colon}${input.installHealthy ? label("Complete", "完整") : label("Incomplete", "不完整")}`);
  const install: [boolean | undefined, string, string][] = [
    [status?.wordFilesPresent, "Word add-in files", "Word 加载项文件"],
    [status?.wordRegistryComplete, "Word registration", "Word 注册信息"],
    [status?.wordLoadEnabled, "Word add-in enabled", "Word 加载项已启用"],
    [status?.powerpointFilesPresent, "PowerPoint add-in files", "PowerPoint 加载项文件"],
    [status?.powerpointRegistryComplete, "PowerPoint registration", "PowerPoint 注册信息"],
    [status?.powerpointLoadEnabled, "PowerPoint add-in enabled", "PowerPoint 加载项已启用"],
    [status?.oleLocalServerHealthy, "Formula OLE service", "公式 OLE 服务"],
    [status?.mathtypeFontsHealthy, "MathType fonts", "MathType 字体"],
  ];
  for (const [ok, en, zh] of install) lines.push(`  ${mark(Boolean(ok))}${label(en, zh)}`);

  lines.push(`${label("Runtime", "运行状态")}${colon}${input.runtimeHealthy ? label("Available", "可用") : label("Unavailable", "不可用")}`);
  const runtime: [boolean | undefined, string, string][] = [
    [status?.companionProcessRunning, "Companion process", "伴侣进程"],
    [status?.companionPortListening, "Local port", "本地端口"],
    [status?.companionHttpsHealthy, "Local HTTPS connection", "本地 HTTPS 连接"],
    [status?.companionCertificateMatches, "Certificate", "证书"],
    [status?.companionProtocolMatches, "Protocol version", "协议版本"],
    [status?.wordConnected, "Word connection", "Word 连接"],
    [status?.powerpointConnected, "PowerPoint connection", "PowerPoint 连接"],
    [renderer?.running, "MathType renderer service", "MathType 渲染服务"],
  ];
  for (const [ok, en, zh] of runtime) lines.push(`  ${mark(Boolean(ok))}${label(en, zh)}`);
  lines.push(
    `${label("Backend", "后端")}${colon}${status?.activeBackend ?? "—"} · ${label("Companion", "伴侣服务")}${colon}${
      companion?.running ? `${companion.bindAddress}:${companion.port}` : label("Stopped", "已停止")
    }`,
  );

  if (wps) {
    lines.push(`WPS Office${colon}${input.wpsLines.length === 0 ? label("Ready", "正常") : label("Needs attention", "需要处理")}`);
    if (wps.installations.length === 0) lines.push(`  ${label("WPS Office was not found.", "未检测到 WPS Office。")}`);
    for (const installation of wps.installations) {
      const bitness = installation.architecture === "x86" ? "32" : "64";
      lines.push(`  WPS Office ${installation.version} · ${bitness}${label("-bit", " 位")} · ${installation.perUser ? label("current user", "当前用户") : label("all users", "所有用户")} · ${installation.path}`);
      lines.push(`  ${mark(installation.writerRegistered)}${label("WPS Writer add-in registered", "WPS 文字加载项已注册")}`);
      lines.push(`  ${mark(installation.presentationRegistered)}${label("WPS Presentation add-in registered", "WPS 演示加载项已注册")}`);
      lines.push(`  ${mark(installation.writerWhitelisted && installation.presentationWhitelisted)}${label("In the WPS add-in whitelist", "已加入 WPS 加载项白名单")}`);
      lines.push(`  ${mark(installation.formulaOleRegistered)}${label("Formula OLE service", "公式 OLE 服务")}`);
    }
    if (wps.installations.length > 0) {
      lines.push(`  ${mark(Boolean(wps.writerLoaded?.enabled))}${hostLoadText(wps.writerLoaded, label("WPS Writer", "WPS 文字"), isEn)}`);
      lines.push(`  ${mark(Boolean(wps.presentationLoaded?.enabled))}${hostLoadText(wps.presentationLoaded, label("WPS Presentation", "WPS 演示"), isEn)}`);
      lines.push(`  ${label("Use VisualTeX in WPS", "在 WPS 中使用 VisualTeX")}${colon}${wps.hostEnabled ? label("Enabled", "已开启") : label("Disabled", "已关闭")}`);
    }
    for (const line of input.wpsLines) lines.push(`  ${mark(false)}${line}`);
  }

  if (input.wpsCheck) {
    const items = wpsReportItems(input.wpsCheck, isEn);
    const allOk = items.every((item) => item.ok);
    lines.push(`${label("WPS check", "WPS 检查")}${input.wpsCheck.checkedAt ? ` ${input.wpsCheck.checkedAt}` : ""}${colon}${allOk ? label("Ready", "正常") : label("Needs attention", "需要处理")}`);
    for (const item of items) lines.push(`  ${mark(item.ok)}${item.text}`);
  }

  if (input.message.trim()) lines.push(`${label("Message", "消息")}${colon}${input.message.trim()}`);
  return lines.join("\n");
}
