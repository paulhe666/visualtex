// Must run first: isolates the landing showcase's storage before editor modules load.
import { isLandingPreview } from "./runtime/landingPreview";
import { StrictMode, Suspense, lazy } from "react";
import { createRoot } from "react-dom/client";
import { installBrowserCompatibility } from "./runtime/browserCompatibility";
import { prepareEditorCrashSafeRestart } from "./runtime/editorCrashRecovery";
import { installFloatingLayerAutoAvoidance } from "./runtime/floatingLayerAutoAvoidance";
import { VisualTexErrorBoundary } from "./runtime/VisualTexErrorBoundary";
import { applyLandingDocumentMeta, detectLandingLang } from "./landing/i18n";
// Same stylesheets as apps/macos/src/desktop/main.tsx. The landing page also
// relies on their base rules (box sizing, body font, MathLive fonts).
import "mathlive/static.css";
import "./styles.css";
import "./styles-editor-parity.css";
import "./styles-windows-shared-latest.css";
import "./styles-macos-platform-overrides.css";
import "./landing/landing.css";
import "./tutorial/tutorial.css";

const EditorRoot = lazy(() => import("./web/EditorRoot"));
const LandingPage = lazy(() =>
  import("./landing/LandingPage").then((module) => ({
    default: module.LandingPage,
  })),
);
const TutorialPage = lazy(() =>
  import("./tutorial/TutorialPage").then((module) => ({
    default: module.TutorialPage,
  })),
);

installBrowserCompatibility();
installFloatingLayerAutoAvoidance();

const normalizedPath = window.location.pathname.replace(/\/+$/, "") || "/";
// Browser regressions synced from the macOS app open the editor at "/".
const editorAtRoot = import.meta.env.VITE_VISUALTEX_EDITOR_AT_ROOT === "1";
const showEditor =
  editorAtRoot ||
  isLandingPreview ||
  normalizedPath === "/editor" ||
  normalizedPath.startsWith("/editor/");

const showTutorial = !showEditor && normalizedPath === "/tutorial";

document.documentElement.dataset.page = showEditor ? "editor" : showTutorial ? "tutorial" : "landing";

if (showEditor) {
  document.documentElement.lang = "zh-CN";
  document.title = "VisualTeX 网页公式编辑器";
  const description = document.querySelector<HTMLMetaElement>('meta[name="description"]');
  if (description) {
    description.content = "免费使用 VisualTeX 网页公式编辑器，通过结构化输入创建、编辑和复制 LaTeX 数学公式。";
  }
} else if (!showTutorial) {
  applyLandingDocumentMeta(detectLandingLang());
}

const canonical = document.querySelector<HTMLLinkElement>('link[rel="canonical"]');
if (canonical) {
  canonical.href = showEditor
    ? "https://visualtex.pauljianliao.com/editor"
    : showTutorial
      ? "https://visualtex.pauljianliao.com/tutorial"
      : "https://visualtex.pauljianliao.com/";
}

function describeBootError(value: unknown): string {
  if (value instanceof Error) return value.stack || value.message;
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

// Same recovery panel as the macOS app when the editor bundle fails to load.
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

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX application root element.");

const render = () =>
  createRoot(root).render(
    <StrictMode>
      <VisualTexErrorBoundary>
        <Suspense fallback={<main className="route-loading" aria-label="Loading VisualTeX" />}>
          {showEditor ? <EditorRoot /> : showTutorial ? <TutorialPage /> : <LandingPage />}
        </Suspense>
      </VisualTexErrorBoundary>
    </StrictMode>,
  );

if (showEditor) {
  // Load the editor bundle up front so a failed chunk shows the recovery panel.
  import("./web/EditorRoot").then(render, renderBootError);
} else {
  render();
}
