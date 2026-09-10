import { installBrowserCompatibility } from "./runtime/browserCompatibility";
import { prepareEditorCrashSafeRestart } from "./runtime/editorCrashRecovery";
import { installFloatingLayerAutoAvoidance } from "./runtime/floatingLayerAutoAvoidance";

installBrowserCompatibility();
installFloatingLayerAutoAvoidance();

function describeBootError(value: unknown): string {
  if (value instanceof Error) return value.stack || value.message;
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function renderBootError(value: unknown) {
  const detail = describeBootError(value) || "Unknown startup error";
  console.error("VisualTeX frontend startup failed", value);
  const root = document.getElementById("root");
  if (!root) return;
  root.innerHTML = "";
  const panel = document.createElement("main");
  panel.setAttribute("role", "alert");
  panel.style.cssText =
    "min-height:100vh;display:grid;place-content:center;gap:12px;padding:32px;" +
    "background:#f5f6f8;color:#9f1239;font:14px/1.6 -apple-system,BlinkMacSystemFont,sans-serif;" +
    "white-space:pre-wrap;overflow:auto";
  const heading = document.createElement("strong");
  heading.textContent = "VisualTeX 前端启动失败";
  const message = document.createElement("code");
  message.textContent = detail;
  const actions = document.createElement("div");
  actions.style.cssText = "display:flex;gap:10px;flex-wrap:wrap";
  const reload = document.createElement("button");
  reload.type = "button";
  reload.textContent = "重新加载";
  reload.style.cssText =
    "padding:8px 14px;border:1px solid #be123c;border-radius:7px;background:#be123c;" +
    "color:white;font:600 13px -apple-system,BlinkMacSystemFont,sans-serif;cursor:pointer";
  reload.addEventListener("click", () => window.location.reload());
  const safeRestart = document.createElement("button");
  safeRestart.type = "button";
  safeRestart.textContent = "安全启动（保留设置）";
  safeRestart.style.cssText =
    "padding:8px 14px;border:1px solid #9f1239;border-radius:7px;background:white;" +
    "color:#9f1239;font:600 13px -apple-system,BlinkMacSystemFont,sans-serif;cursor:pointer";
  safeRestart.addEventListener("click", () => {
    try {
      prepareEditorCrashSafeRestart();
      window.location.reload();
    } catch (reason) {
      console.error("VisualTeX boot safe restart preparation failed", reason);
      window.alert("无法准备安全启动，原有数据未被清除。请把错误信息发送给开发者。");
    }
  });
  actions.append(reload, safeRestart);
  panel.append(heading, message, actions);
  root.append(panel);
}

void import("./desktop/main").catch(renderBootError);
