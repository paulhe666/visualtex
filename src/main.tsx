import { StrictMode, Suspense, lazy } from "react";
import { installBrowserCompatibility } from "./runtime/browserCompatibility";
import { VisualTexErrorBoundary } from "./runtime/VisualTexErrorBoundary";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "./styles.css";
import "./styles-macos-main.css";
import "./styles-editor-parity.css";
import "./styles-latest-macos-ui.css";
import "./styles-windows-shared-latest.css";
import "./styles-custom-symbol-designer.css";
import "./landing/landing.css";
const App = lazy(() => import("./App"));
const LandingPage = lazy(() =>
  import("./landing/LandingPage").then((module) => ({
    default: module.LandingPage,
  })),
);

installBrowserCompatibility();

const normalizedPath = window.location.pathname.replace(/\/+$/, "") || "/";
const showEditor = normalizedPath === "/editor" || normalizedPath.startsWith("/editor/");

document.documentElement.dataset.page = showEditor ? "editor" : "landing";
document.documentElement.lang = "zh-CN";

document.title = showEditor
  ? "VisualTeX 网页公式编辑器"
  : "VisualTeX — 可视化 LaTeX 公式编辑器";

const description = document.querySelector<HTMLMetaElement>('meta[name="description"]');
if (description) {
  description.content = showEditor
    ? "免费使用 VisualTeX 网页公式编辑器，通过结构化输入创建、编辑和复制 LaTeX 数学公式。"
    : "VisualTeX 是面向数学、物理、工程、教学与科研写作的可视化 LaTeX 公式编辑器，提供网页端、桌面端、本地公式 OCR 与 Office 工作流。";
}

const canonical = document.querySelector<HTMLLinkElement>('link[rel="canonical"]');
if (canonical) {
  canonical.href = showEditor
    ? "https://visualtex.pauljianliao.com/editor"
    : "https://visualtex.pauljianliao.com/";
}

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <VisualTexErrorBoundary>
      <Suspense fallback={<main className="route-loading" aria-label="Loading VisualTeX" />}>
        {showEditor ? <App /> : <LandingPage />}
      </Suspense>
    </VisualTexErrorBoundary>
  </StrictMode>,
);
