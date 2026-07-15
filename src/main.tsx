import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "./styles.css";
import "./landing/landing.css";
import App from "./App";
import { LandingPage } from "./landing/LandingPage";

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
    {showEditor ? <App /> : <LandingPage />}
  </StrictMode>,
);
