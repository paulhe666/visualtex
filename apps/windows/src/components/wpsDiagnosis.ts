/**
 * Why each WPS Office does or does not show VisualTeX, as windows_office_diagnosis.ps1 (-DiagnosisMode Wps)
 * finds it, in the words Settings > Office and the WPS repair prompt show under the problem.
 */
export interface WpsHostDiagnosis {
  host: "writer" | "presentation";
  /** the registration in WPS's own registry view */
  state: "ok" | "missing" | "disabled" | "server" | "fileMissing" | "wrongBitness";
  detail: {
    loadBehavior?: number;
    scope?: "user" | "machine";
    file?: string;
    quarantine?: string;
    fileArchitecture?: string;
  };
  whitelisted: boolean;
  /** "ok", "" (not tried: the registration is broken) or the error WPS meets creating the add-in */
  activation: string;
}

export interface WpsInstallationDiagnosis {
  path: string;
  architecture: string;
  version: string;
  registeredAt: string;
  hosts: WpsHostDiagnosis[];
  /** WPS processes of this installation started before the add-ins were registered for it */
  staleProcesses: { id: number; name: string; started: string }[];
}

/** What a WPS started for the question (hidden) says about one add-in: Settings > Office > Check WPS. */
export interface WpsAnswer {
  host: "writer" | "presentation";
  /** WPS's COM add-in list has the add-in */
  listed: boolean;
  /** WPS connected it by itself */
  connected: boolean;
  /** when it had not: whether it connected when asked; null when not asked */
  askedConnected: boolean | null;
  /** WPS's error when asked to connect it */
  answer: string;
  /** the question failed: WPS did not start or did not answer in time */
  error: string;
}

export interface WpsDiagnosis {
  installations: WpsInstallationDiagnosis[];
  /** WPS's own lists of add-ins it stopped loading that name VisualTeX ("WPS\\AddinsLF: {CLSID}") */
  wpsRecords: string[];
  /** only after Check WPS */
  wpsAnswers?: WpsAnswer[];
  /** other registry values and settings files of WPS that name the add-ins; only after Check WPS */
  wpsSettings?: string[];
  checkedAt?: string;
  error?: string;
}

function asArray(value: unknown): unknown[] {
  if (Array.isArray(value)) return value;
  return value === null || value === undefined ? [] : [value];
}

function text(value: unknown) {
  return typeof value === "string" ? value : "";
}

/** Lenient: a diagnosis explains, so a missing field only drops its line. */
export function decodeWpsDiagnosis(value: unknown): WpsDiagnosis {
  const record = (value ?? {}) as Record<string, unknown>;
  return {
    error: text(record.error) || text(record.wpsAnswersError) || undefined,
    wpsRecords: asArray(record.wpsRecords).map(text).filter(Boolean),
    wpsAnswers: "wpsAnswers" in record
      ? asArray(record.wpsAnswers).map((entry) => {
          const answer = (entry ?? {}) as Record<string, unknown>;
          return {
            host: answer.host === "presentation" ? "presentation" : "writer",
            listed: answer.listed === true,
            connected: answer.connected === true,
            askedConnected: typeof answer.askedConnected === "boolean" ? answer.askedConnected : null,
            answer: text(answer.answer),
            error: text(answer.error),
          } satisfies WpsAnswer;
        })
      : undefined,
    wpsSettings: "wpsSettings" in record ? asArray(record.wpsSettings).map(text).filter(Boolean) : undefined,
    checkedAt: text(record.checkedAt) || undefined,
    installations: asArray(record.installations).map((entry) => {
      const installation = (entry ?? {}) as Record<string, unknown>;
      return {
        path: text(installation.path),
        architecture: text(installation.architecture),
        version: text(installation.version),
        registeredAt: text(installation.registeredAt),
        hosts: asArray(installation.hosts).map((hostEntry) => {
          const host = (hostEntry ?? {}) as Record<string, unknown>;
          const detail = (host.detail ?? {}) as Record<string, unknown>;
          return {
            host: host.host === "presentation" ? "presentation" : "writer",
            state: (["ok", "missing", "disabled", "server", "fileMissing", "wrongBitness"].includes(text(host.state))
              ? text(host.state)
              : "ok") as WpsHostDiagnosis["state"],
            detail: {
              loadBehavior: typeof detail.loadBehavior === "number" ? detail.loadBehavior : undefined,
              scope: detail.scope === "user" ? "user" : detail.scope === "machine" ? "machine" : undefined,
              file: text(detail.file) || undefined,
              quarantine: text(detail.quarantine) || undefined,
              fileArchitecture: text(detail.fileArchitecture) || undefined,
            },
            whitelisted: host.whitelisted === true,
            activation: text(host.activation),
          };
        }),
        staleProcesses: asArray(installation.staleProcesses).map((processEntry) => {
          const process = (processEntry ?? {}) as Record<string, unknown>;
          return {
            id: typeof process.id === "number" ? process.id : 0,
            name: text(process.name),
            started: text(process.started),
          };
        }),
      };
    }),
  };
}

const bits = (architecture: string) => (architecture === "x86" ? "32" : "64");

/** One line per cause found, most basic first; empty when the diagnosis finds nothing wrong. */
export function wpsDiagnosisLines(diagnosis: WpsDiagnosis, isEn: boolean): string[] {
  const lines: string[] = [];
  if (diagnosis.wpsRecords.length > 0) {
    lines.push(isEn
      ? `WPS has VisualTeX on its own list of add-ins it stopped loading (${diagnosis.wpsRecords.join("; ")}). Repair the Office integration to clear it, then reopen WPS.`
      : `WPS 把 VisualTeX 记在它停止加载的加载项列表中（${diagnosis.wpsRecords.join("；")}）。修复 Office 集成会清除这条记录，然后重新打开 WPS。`);
  }
  const several = diagnosis.installations.length > 1;
  for (const installation of diagnosis.installations) {
    const wps = several ? ` ${bits(installation.architecture)}-bit ${installation.version}` : "";
    const wpsZh = several ? ` ${bits(installation.architecture)} 位 ${installation.version}` : "";
    for (const host of installation.hosts) {
      const name = isEn
        ? `${host.host === "writer" ? "WPS Writer" : "WPS Presentation"}${wps}`
        : `${host.host === "writer" ? "WPS 文字" : "WPS 演示"}${wpsZh}`;
      const { detail } = host;
      switch (host.state) {
        case "missing":
          lines.push(isEn
            ? `${name}: VisualTeX is not registered for ${bits(installation.architecture)}-bit WPS.`
            : `${name}：${bits(installation.architecture)} 位 WPS 中没有 VisualTeX 的注册。`);
          continue;
        case "disabled":
          lines.push(isEn
            ? `${name}: VisualTeX is turned off (LoadBehavior=${detail.loadBehavior ?? "?"}, ${detail.scope === "user" ? "for this user" : "for this computer"}). Tick VisualTeX under COM Add-ins in WPS, or repair the Office integration.`
            : `${name}：VisualTeX 已被停用（LoadBehavior=${detail.loadBehavior ?? "?"}，${detail.scope === "user" ? "当前用户" : "本机"}）。在 WPS 的“COM 加载项”中勾选 VisualTeX，或修复 Office 集成。`);
          continue;
        case "server":
          lines.push(isEn
            ? `${name}: the add-in's COM registration is incomplete.`
            : `${name}：VisualTeX 加载项的 COM 注册不完整。`);
          continue;
        case "fileMissing":
          lines.push(isEn
            ? `${name}: the add-in file is missing: ${detail.file ?? ""}${detail.quarantine ? ` (${detail.quarantine})` : ""}.`
            : `${name}：加载项文件不存在：${detail.file ?? ""}${detail.quarantine ? `（${detail.quarantine}）` : ""}。`);
          continue;
        case "wrongBitness":
          lines.push(isEn
            ? `${name}: the registered add-in is ${bits(detail.fileArchitecture ?? "")}-bit, WPS is ${bits(installation.architecture)}-bit.`
            : `${name}：注册的是 ${bits(detail.fileArchitecture ?? "")} 位加载项，WPS 是 ${bits(installation.architecture)} 位。`);
          continue;
        default:
          break;
      }
      if (!host.whitelisted) {
        lines.push(isEn
          ? `${name}: VisualTeX is not in WPS's add-in whitelist.`
          : `${name}：VisualTeX 不在 WPS 加载项白名单中。`);
      }
      if (host.activation && host.activation !== "ok") {
        lines.push(isEn
          ? `${name}: the add-in cannot start: ${host.activation}`
          : `${name}：加载项无法启动：${host.activation}`);
      }
    }
    if (installation.staleProcesses.length > 0) {
      const since = installation.staleProcesses.map((process) => process.started).sort()[0];
      lines.push(isEn
        ? `WPS${wps} has been running since ${since}, before VisualTeX registered its add-ins (${installation.registeredAt}). Quit WPS completely, including its icon in the notification area, and open it again.`
        : `WPS${wpsZh} 从 ${since} 起一直在运行，早于 VisualTeX 为它注册加载项（${installation.registeredAt}）。完全退出 WPS（包括通知区域中的 WPS）后重新打开。`);
    }
  }
  for (const answer of diagnosis.wpsAnswers ?? []) {
    if (answer.listed && answer.connected) continue;
    lines.push(`${wpsAnswerLabel(answer.host, isEn)}${isEn ? ": " : "："}${wpsAnswerText(answer, isEn)}`);
  }
  return lines;
}

function wpsAnswerLabel(host: WpsAnswer["host"], isEn: boolean) {
  return host === "writer"
    ? isEn ? "WPS Writer COM add-in" : "WPS 文字 COM 加载项"
    : isEn ? "WPS Presentation COM add-in" : "WPS 演示 COM 加载项";
}

function wpsAnswerText(answer: WpsAnswer, isEn: boolean) {
  if (answer.error) return isEn ? `WPS could not be asked: ${answer.error}` : `无法询问 WPS：${answer.error}`;
  if (!answer.listed) return isEn ? "not listed" : "未列出";
  if (answer.connected) return isEn ? "connected" : "已连接";
  if (answer.askedConnected) return isEn ? "not connected; connects when asked" : "未连接；请求时可以连接";
  return isEn
    ? `not connected; WPS answered: ${answer.answer || "no reason"}`
    : `未连接；WPS 回答：${answer.answer || "无说明"}`;
}

export interface WpsReportItem {
  ok: boolean;
  text: string;
}

/**
 * Every check of the WPS diagnosis, passed or not, for the WPS card under Settings > Office > Advanced
 * diagnostics and the copied report: registration, whitelist and load test per host, WPS processes older
 * than the registration, what WPS itself answers, its stop-loading lists and other entries naming VisualTeX.
 */
export function wpsReportItems(diagnosis: WpsDiagnosis, isEn: boolean): WpsReportItem[] {
  const items: WpsReportItem[] = [];
  const colon = isEn ? ": " : "：";
  if (diagnosis.error) items.push({ ok: false, text: (isEn ? "Diagnosis failed: " : "诊断失败：") + diagnosis.error });
  const several = diagnosis.installations.length > 1;
  for (const installation of diagnosis.installations) {
    const wps = several ? ` ${bits(installation.architecture)}-bit ${installation.version}` : "";
    const wpsZh = several ? ` ${bits(installation.architecture)} 位 ${installation.version}` : "";
    for (const host of installation.hosts) {
      const name = isEn
        ? `${host.host === "writer" ? "WPS Writer" : "WPS Presentation"}${wps}`
        : `${host.host === "writer" ? "WPS 文字" : "WPS 演示"}${wpsZh}`;
      if (host.state !== "ok") {
        const single: WpsDiagnosis = {
          installations: [{ ...installation, hosts: [host], staleProcesses: [] }],
          wpsRecords: [],
        };
        for (const line of wpsDiagnosisLines(single, isEn)) items.push({ ok: false, text: line });
        continue;
      }
      items.push({ ok: true, text: isEn ? `${name} registration` : `${name}注册信息` });
      items.push({ ok: host.whitelisted, text: isEn ? `${name} in the WPS add-in whitelist` : `${name}加载项白名单` });
      const loadTest = isEn ? `${name} load test` : `${name}加载测试`;
      items.push(host.activation === "ok"
        ? { ok: true, text: loadTest }
        : { ok: false, text: `${loadTest}${colon}${host.activation || (isEn ? "not run" : "未运行")}` });
    }
    if (installation.staleProcesses.length > 0) {
      const since = installation.staleProcesses.map((process) => process.started).sort()[0];
      items.push({
        ok: false,
        text: isEn
          ? `WPS${wps} running since ${since}, before the add-ins were registered (${installation.registeredAt})`
          : `WPS${wpsZh} 自 ${since} 起运行，早于加载项注册（${installation.registeredAt}）`,
      });
    }
  }
  for (const answer of diagnosis.wpsAnswers ?? []) {
    items.push({
      ok: answer.listed && answer.connected && !answer.error,
      text: `${wpsAnswerLabel(answer.host, isEn)}${colon}${wpsAnswerText(answer, isEn)}`,
    });
  }
  const records = isEn ? "WPS stop-loading lists" : "WPS 停止加载记录";
  items.push({
    ok: diagnosis.wpsRecords.length === 0,
    text: `${records}${colon}${diagnosis.wpsRecords.length === 0 ? (isEn ? "none" : "无") : diagnosis.wpsRecords.join(isEn ? "; " : "；")}`,
  });
  if (diagnosis.wpsSettings) {
    const other = isEn ? "Other WPS entries naming VisualTeX" : "WPS 其他记录";
    items.push({
      ok: diagnosis.wpsSettings.length === 0,
      text: `${other}${colon}${diagnosis.wpsSettings.length === 0 ? (isEn ? "none" : "无") : diagnosis.wpsSettings.join(isEn ? "; " : "；")}`,
    });
  }
  return items;
}

/**
 * The line for a WPS whose registration, whitelist and load test are all fine while WPS has not loaded
 * VisualTeX: the cause is on WPS's side.
 */
export function wpsDiagnosisAllFine(isEn: boolean) {
  return isEn
    ? "Registration, whitelist and load test are fine: WPS itself has not loaded VisualTeX."
    : "注册、白名单和加载测试都正常：是 WPS 自己没有加载 VisualTeX。";
}
