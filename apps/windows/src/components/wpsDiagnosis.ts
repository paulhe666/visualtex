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

export interface WpsDiagnosis {
  installations: WpsInstallationDiagnosis[];
  /** WPS's own lists of add-ins it stopped loading that name VisualTeX ("WPS\\AddinsLF: {CLSID}") */
  wpsRecords: string[];
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
    error: text(record.error) || undefined,
    wpsRecords: asArray(record.wpsRecords).map(text).filter(Boolean),
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
  return lines;
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
