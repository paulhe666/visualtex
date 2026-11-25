import assert from "node:assert/strict";
import {
  decodeWpsDiagnosis,
  wpsDiagnosisLines,
  wpsReportItems,
} from "../src/components/wpsDiagnosis.ts";
import { officeDiagnosticReport } from "../src/components/officeDiagnosticReport.ts";

// What windows_office_diagnosis.ps1 -DiagnosisMode Wps -AskWps prints on a PC where WPS loads both add-ins.
const healthy = {
  installations: [
    {
      path: "C:\\Users\\u\\AppData\\Local\\Kingsoft\\WPS Office\\12.1.0.29681\\office6\\wps.exe",
      architecture: "x64",
      version: "12.1.0.29681",
      registeredAt: "2026-11-23 20:16",
      hosts: [
        { host: "writer", state: "ok", detail: {}, whitelisted: true, activation: "ok" },
        { host: "presentation", state: "ok", detail: {}, whitelisted: true, activation: "ok" },
      ],
      staleProcesses: [],
    },
  ],
  wpsRecords: [],
  wpsAnswers: [
    { host: "writer", listed: true, connected: true, askedConnected: null, answer: "", error: "" },
    { host: "presentation", listed: true, connected: true, askedConnected: null, answer: "", error: "" },
  ],
  wpsSettings: [],
  checkedAt: "2026-11-23 20:45:39",
};

const good = decodeWpsDiagnosis(healthy);
assert.equal(good.wpsAnswers?.length, 2);
assert.deepEqual(wpsDiagnosisLines(good, false), [], "a healthy WPS has no causes");
const goodItems = wpsReportItems(good, false);
assert.ok(goodItems.every((item) => item.ok), JSON.stringify(goodItems));
assert.ok(goodItems.some((item) => item.text === "WPS 文字 COM 加载项：已连接"));
assert.ok(goodItems.some((item) => item.text === "WPS 停止加载记录：无"));
assert.ok(goodItems.some((item) => item.text === "WPS 其他记录：无"));

// Without Check WPS there is no answer and no settings line: the automatic check stays as it was.
const automatic = decodeWpsDiagnosis({ installations: healthy.installations, wpsRecords: [] });
assert.equal(automatic.wpsAnswers, undefined);
assert.equal(automatic.wpsSettings, undefined);
assert.ok(!wpsReportItems(automatic, false).some((item) => item.text.includes("COM 加载项")));

// PowerShell writes a one-element array as a bare object; WPS refusing, not listing and not answering.
const broken = decodeWpsDiagnosis({
  ...healthy,
  wpsRecords: ["WPS\\AddinsLF: {F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}"],
  wpsAnswers: { host: "writer", listed: true, connected: false, askedConnected: false, answer: "加载项加载失败", error: "" },
  wpsSettings: "%USERPROFILE%\\AppData\\Roaming\\kingsoft\\x\\mdbx.dat (2026-11-23 20:00): VisualTeX.WordVsto, user_remove",
});
assert.equal(broken.wpsAnswers?.length, 1);
const causes = wpsDiagnosisLines(broken, false);
assert.ok(causes.some((line) => line.includes("停止加载")), causes.join("\n"));
assert.ok(causes.includes("WPS 文字 COM 加载项：未连接；WPS 回答：加载项加载失败"), causes.join("\n"));
const brokenItems = wpsReportItems(broken, true);
assert.ok(brokenItems.some((item) => !item.ok && item.text === "WPS Writer COM add-in: not connected; WPS answered: 加载项加载失败"));
assert.ok(brokenItems.some((item) => !item.ok && item.text.startsWith("Other WPS entries naming VisualTeX: %USERPROFILE%")));

const notListed = decodeWpsDiagnosis({ ...healthy, wpsAnswers: [{ host: "presentation", listed: false, connected: false, askedConnected: null, answer: "", error: "" }] });
assert.ok(wpsDiagnosisLines(notListed, false).includes("WPS 演示 COM 加载项：未列出"));
const noAnswer = decodeWpsDiagnosis({ ...healthy, wpsAnswers: [{ host: "writer", listed: false, connected: false, askedConnected: null, answer: "", error: "WPS did not answer within 90 seconds" }] });
assert.ok(wpsDiagnosisLines(noAnswer, true).includes("WPS Writer COM add-in: WPS could not be asked: WPS did not answer within 90 seconds"));
const healedWhenAsked = decodeWpsDiagnosis({ ...healthy, wpsAnswers: [{ host: "writer", listed: true, connected: false, askedConnected: true, answer: "", error: "" }] });
assert.ok(wpsDiagnosisLines(healedWhenAsked, false).includes("WPS 文字 COM 加载项：未连接；请求时可以连接"));

// A broken registration lists its cause instead of the passed checks.
const missing = decodeWpsDiagnosis({
  ...healthy,
  installations: [{ ...healthy.installations[0], hosts: [{ host: "writer", state: "missing", detail: {}, whitelisted: false, activation: "" }] }],
});
const missingItems = wpsReportItems(missing, false);
assert.ok(missingItems.some((item) => !item.ok && item.text.includes("没有 VisualTeX 的注册")), JSON.stringify(missingItems));
assert.ok(!missingItems.some((item) => item.text.includes("加载测试")));

// The copied report: everything the page shows, the check included, with [ok] and [!] marks.
const report = officeDiagnosticReport(
  {
    isEn: false,
    title: "Office 集成正常",
    installHealthy: true,
    runtimeHealthy: false,
    status: null,
    companion: null,
    renderer: null,
    wps: {
      installations: [
        {
          path: healthy.installations[0].path,
          architecture: "x64",
          version: "12.1.0.29681",
          perUser: true,
          writerRegistered: true,
          presentationRegistered: true,
          writerWhitelisted: true,
          presentationWhitelisted: true,
          formulaOleRegistered: true,
        },
      ],
      writerLoaded: { at: "2026-11-23 20:42:22", enabled: true },
      presentationLoaded: null,
      hostEnabled: true,
      companionError: "",
    } as never,
    wpsLines: [],
    wpsCheck: broken,
    message: "",
  },
  new Date(2026, 10, 23, 21, 5, 9),
);
assert.match(report, /^VisualTeX \S+ Windows · Office 诊断报告 · 2026-11-23 21:05:09$/m);
assert.match(report, /^Office 集成：Office 集成正常$/m);
assert.match(report, /^ {2}\[ok\] WPS 文字已加载 VisualTeX · 2026-11-23 20:42:22$/m);
assert.match(report, /^ {2}\[!\] {2}WPS 演示尚未加载 VisualTeX$/m);
assert.match(report, /^WPS 检查 2026-11-23 20:45:39：需要处理$/m);
assert.match(report, /^ {2}\[!\] {2}WPS 文字 COM 加载项：未连接；WPS 回答：加载项加载失败$/m);
assert.match(report, /^ {2}\[!\] {2}WPS 停止加载记录：WPS\\AddinsLF: \{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1\}$/m);

console.log("WPS diagnosis report regression passed");
