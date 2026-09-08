import { useEffect, useRef } from "react";
import type { LucideIcon } from "lucide-react";
import {
  ArrowRight,
  ArrowUpRight,
  Download,
  Laptop,
  Monitor,
} from "lucide-react";
import { VisualTeXLogo } from "../components/VisualTeXLogo";
import { SupportCodes } from "./SupportCodes";

const VERSION = "1.2.4";
const DOWNLOAD_BASE = `https://download.visualtex.pauljianliao.com/visualtex-downloads/releases/v${VERSION}`;
const OCR_MODEL_BASE = "https://download.visualtex.pauljianliao.com/ppformula-model";
const RELEASES_URL = "https://github.com/paulhe666/visualtex/releases";

type PlatformId = "mac" | "windows";

type DownloadOption = {
  id: PlatformId;
  icon: LucideIcon;
  title: string;
  detail: string;
  href: string;
  action: string;
  secondaryHref?: string;
  secondaryAction?: string;
};

const downloads: readonly DownloadOption[] = [
  {
    id: "mac",
    icon: Laptop,
    title: "macOS",
    detail: "Apple Silicon · macOS 11+",
    href: `${DOWNLOAD_BASE}/VisualTeX_${VERSION}_aarch64.dmg`,
    action: "下载 DMG",
  },
  {
    id: "windows",
    icon: Monitor,
    title: "Windows",
    detail: "Windows 10/11 · x64",
    href: `${DOWNLOAD_BASE}/VisualTeX_${VERSION}_x64-setup.exe`,
    action: "下载安装程序",
  },
];

const ocrModels = [
  {
    id: "ocr-s",
    title: "OCR-S 模型",
    detail: "轻量版 · Windows x64 · 200.05 MB",
    href: `${OCR_MODEL_BASE}/VisualTeX_PP-FormulaNet_plus-S_windows-x64.vtxocrmodel`,
    action: "下载 OCR-S 模型",
    recommended: false,
  },
  {
    id: "ocr-m",
    title: "OCR-M 模型",
    detail: "均衡版 · Windows x64 · 425.83 MB",
    href: `${OCR_MODEL_BASE}/VisualTeX_PP-FormulaNet_plus-M_windows-x64.vtxocrmodel`,
    action: "下载 OCR-M 模型",
    recommended: true,
  },
  {
    id: "ocr-l",
    title: "OCR-L 模型",
    detail: "高精度版 · Windows x64 · 670.29 MB",
    href: `${OCR_MODEL_BASE}/VisualTeX_PP-FormulaNet_plus-L_windows-x64.vtxocrmodel`,
    action: "下载 OCR-L 模型",
    recommended: false,
  },
] as const;

const features = [
  { title: "可视化编辑", detail: "直接编辑分式、积分和矩阵，LaTeX 源码同步更新。", scope: "网页 / 桌面" },
  { title: "原生 MathType 公式", detail: "无需安装 MathType，即可插入和编辑原生公式。", scope: "Windows" },
  { title: "Word 与 PowerPoint", detail: "插入、修改公式，管理编号与交叉引用。", scope: "桌面" },
  { title: "图片转公式", detail: "粘贴图片识别；桌面端还支持离线 OCR。", scope: "网页 / 桌面" },
  { title: "LaTeX 源码", detail: "语法高亮、命令补全、多行编辑。", scope: "网页 / 桌面" },
  { title: "复制与导出", detail: "导出 LaTeX、SVG 和 PNG。", scope: "网页 / 桌面" },
] as const;

type PlatformDetection = {
  platform: PlatformId | "";
  isMobileDevice: boolean;
};

function detectPlatform(): PlatformDetection {
  const userAgent = navigator.userAgent.toLowerCase();
  const platform = navigator.platform.toLowerCase();
  const isIPadDesktopMode = platform.includes("mac") && navigator.maxTouchPoints > 1;
  const isMobileDevice = /android|iphone|ipad|ipod|mobile/.test(userAgent) || isIPadDesktopMode;

  if (isMobileDevice || userAgent.includes("cros")) {
    return { platform: "", isMobileDevice };
  }
  if (userAgent.includes("windows") || platform.startsWith("win")) {
    return { platform: "windows", isMobileDevice: false };
  }
  if (userAgent.includes("macintosh") || platform.startsWith("mac")) {
    return { platform: "mac", isMobileDevice: false };
  }
  return { platform: "", isMobileDevice: false };
}


function EditorPreview() {
  const viewportRef = useRef<HTMLDivElement>(null);
  const frameRef = useRef<HTMLIFrameElement>(null);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const resize = () => {
      if (frameRef.current) {
        frameRef.current.style.transform = `scale(${viewport.clientWidth / 1440})`;
      }
    };
    resize();
    const observer = new ResizeObserver(resize);
    observer.observe(viewport);
    return () => observer.disconnect();
  }, []);

  return (
    <figure className="landing-preview">
      <div className="landing-preview-label">
        <span>VisualTeX / 网页编辑器</span>
        <a href="/editor">打开编辑器 <ArrowUpRight size={16} aria-hidden="true" /></a>
      </div>
      <div className="landing-preview-viewport" ref={viewportRef}>
        <iframe
          ref={frameRef}
          className="landing-preview-frame"
          src="/editor?landing-preview=1"
          title="VisualTeX 网页公式编辑器预览"
          tabIndex={-1}
          aria-hidden="true"
        />
      </div>
    </figure>
  );
}

export function LandingPage() {
  const { platform: detectedPlatform, isMobileDevice } = detectPlatform();
  const orderedDownloads = [...downloads].sort(
    (left, right) => Number(right.id === detectedPlatform) - Number(left.id === detectedPlatform),
  );

  return (
    <div className="landing-page">
      <a className="landing-skip" href="#main">跳转到正文</a>
      <header className="landing-header">
        <div className="landing-container landing-header-inner">
          <a className="landing-brand" href="/" aria-label="VisualTeX 首页">
            <VisualTeXLogo />
            <span>VisualTeX</span>
          </a>
          <nav className="landing-nav" aria-label="主要导航">
            <a className="landing-nav-detail" href="#features">功能</a>
            <a href="#download">下载</a>
            <a className="landing-nav-detail" href={RELEASES_URL} target="_blank" rel="noreferrer">GitHub</a>
            <a className="landing-button landing-button-small" href="/editor">在线编辑 <ArrowUpRight size={16} aria-hidden="true" /></a>
          </nav>
        </div>
      </header>

      <main id="main">
        <section className="landing-hero landing-container" aria-labelledby="landing-title">
          <div className="landing-hero-copy">
            <div>
              <p className="landing-eyebrow">Visual LaTeX Editor</p>
              <h1 id="landing-title">可视化编辑，<br /><mark className="landing-mark landing-mark-lilac">LaTeX</mark> 同步。</h1>
            </div>
            <div className="landing-hero-intro">
              <p>在浏览器中编辑公式，<br />在 Word 与 PowerPoint 中继续使用。</p>
              <div className="landing-actions">
                <a className="landing-button" href="/editor">打开编辑器 <ArrowRight size={18} aria-hidden="true" /></a>
                <a className="landing-button landing-button-outline" href="#download">下载桌面端 <Download size={17} aria-hidden="true" /></a>
              </div>
              <p className="landing-platforms">Web · Windows · macOS</p>
            </div>
          </div>
          <EditorPreview />
        </section>

        <section className="landing-features" id="features" aria-labelledby="features-title">
          <div className="landing-container">
          <div className="landing-section-heading">
            <p className="landing-eyebrow">功能</p>
            <h2 id="features-title">专注<mark className="landing-mark landing-mark-coral">公式</mark>，连接<mark className="landing-mark landing-mark-lilac">文档</mark>。</h2>
          </div>
          <div className="landing-feature-grid">
            {features.map((feature, index) => (
              <article className="landing-feature" data-feature={index} key={feature.title}>
                <div className="landing-feature-meta">
                  <span>0{index + 1}</span><span>{feature.scope}</span>
                </div>
                <h3>{feature.title}</h3>
                <p>{feature.detail}</p>
              </article>
            ))}
          </div>
          </div>
        </section>

        <section className="landing-download" id="download" aria-labelledby="download-title">
          <div className="landing-container">
            <div className="landing-download-heading">
              <div>
                <p className="landing-eyebrow">Desktop</p>
                <h2 id="download-title">下载 VisualTeX</h2>
              </div>
              <a className="landing-text-link" href={RELEASES_URL} target="_blank" rel="noreferrer">全部版本与安装说明 <ArrowUpRight size={16} aria-hidden="true" /></a>
            </div>
            {isMobileDevice && <p className="landing-device-note">桌面安装包请在电脑上下载。</p>}
            <div className="landing-download-grid">
              {orderedDownloads.map((download) => {
                const Icon = download.icon;
                const recommended = download.id === detectedPlatform;
                return (
                  <article className="landing-download-item" key={download.id} aria-label={recommended ? `${download.title}，当前设备` : download.title}>
                    <div className="landing-download-title">
                      <Icon size={24} strokeWidth={1.5} aria-hidden="true" />
                      <h3>{download.title}</h3>
                      {recommended && <span className="landing-device-label">当前设备</span>}
                    </div>
                    <p>{download.detail}</p>
                    <div className="landing-download-bottom">
                      <a className="landing-button landing-download-action" href={download.href}><Download size={17} aria-hidden="true" />{download.action}</a>
                      <span className="landing-version">v{VERSION}</span>
                    </div>
                  </article>
                );
              })}
            </div>

            <section className="landing-models" aria-labelledby="models-title">
              <div className="landing-models-heading"><h3 id="models-title">Windows 离线 OCR 模型</h3><p>下载后在桌面端导入。</p></div>
              <div className="landing-ocr-model-grid">
                {ocrModels.map((model) => (
                  <article className="landing-model-row" key={model.id}>
                    <h3>{model.title}</h3>
                    <p>{model.detail}</p>
                    <a className="landing-button landing-model-download" href={model.href} aria-label={model.action}>下载 <Download size={16} aria-hidden="true" /></a>
                  </article>
                ))}
              </div>
            </section>
          </div>
        </section>
        <section className="landing-support landing-container" aria-labelledby="support-title">
          <div className="landing-support-heading">
            <h2 id="support-title">支持与交流</h2>
            <p>打赏自愿，不影响任何功能的使用。<br />QQ 交流群：1045801770</p>
          </div>
          <SupportCodes />
        </section>
      </main>

      <footer className="landing-footer landing-container">
        <a className="landing-brand" href="/"><VisualTeXLogo /><span>VisualTeX</span></a>
        <span>Visual LaTeX Editor</span>
        <a className="landing-text-link" href="https://github.com/paulhe666/visualtex" target="_blank" rel="noreferrer">GitHub <ArrowUpRight size={16} aria-hidden="true" /></a>
      </footer>
    </div>
  );
}

export default LandingPage;
