import type { LucideIcon } from "lucide-react";
import {
  ArrowRight,
  ArrowUpRight,
  Braces,
  Check,
  Code2,
  Download,
  FileImage,
  FileText,
  Github,
  Laptop,
  Monitor,
  Package,
  ScanLine,
  ShieldCheck,
} from "lucide-react";
import { VisualTeXLogo } from "../components/VisualTeXLogo";

const VERSION = "1.1.0";
const RELEASE_BASE = `https://github.com/paulhe666/visualtex/releases/download/v${VERSION}`;
const RELEASES_URL = "https://github.com/paulhe666/visualtex/releases";

type PlatformId = "mac" | "windows" | "linux";

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
    href: `${RELEASE_BASE}/VisualTeX_${VERSION}_aarch64.dmg`,
    action: "下载 DMG",
  },
  {
    id: "windows",
    icon: Monitor,
    title: "Windows",
    detail: "Windows 10/11 · x64",
    href: `${RELEASE_BASE}/VisualTeX_${VERSION}_x64-setup.exe`,
    action: "下载安装程序",
  },
  {
    id: "linux",
    icon: Package,
    title: "Linux",
    detail: "通用 AppImage · x64",
    href: `${RELEASE_BASE}/VisualTeX_${VERSION}_amd64.AppImage`,
    action: "下载 AppImage",
    secondaryHref: `${RELEASE_BASE}/VisualTeX_${VERSION}_amd64.deb`,
    secondaryAction: "下载 DEB",
  },
];

const features = [
  {
    icon: Braces,
    title: "结构化可视化编辑",
    description: "分式、根式、积分、求和、矩阵与上下标都能直接点击插入，并保持完整数学结构。",
  },
  {
    icon: Code2,
    title: "LaTeX 双向同步",
    description: "可视化公式与 LaTeX 源码同步编辑，支持多种常用环境与复制格式。",
  },
  {
    icon: ScanLine,
    title: "本地公式 OCR",
    description: "桌面端可把公式截图识别成可编辑 LaTeX，图片仅在本机处理。",
  },
  {
    icon: FileText,
    title: "Office 公式工作流",
    description: "在 Word 与 PowerPoint 中插入、更新并重新打开 VisualTeX 公式。",
  },
  {
    icon: ShieldCheck,
    title: "本地优先",
    description: "文档、历史记录与桌面 OCR 工作流优先保存在本机，减少不必要的数据上传。",
  },
  {
    icon: FileImage,
    title: "清晰公式导出",
    description: "适合课程作业、论文、讲义和演示文稿，兼顾公式质量与后续编辑。",
  },
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
  if (userAgent.includes("linux") || platform.includes("linux")) {
    return { platform: "linux", isMobileDevice: false };
  }
  return { platform: "", isMobileDevice: false };
}

function EditorPreview() {
  return (
    <div className="landing-preview-wrap">
      <div className="landing-orbit landing-orbit-one" />
      <div className="landing-orbit landing-orbit-two" />
      <figure className="landing-preview-window">
        <figcaption className="landing-preview-titlebar">
          <span className="landing-preview-brand"><VisualTeXLogo />VisualTeX</span>
          <span className="landing-preview-title">未命名公式 <Check size={11} /></span>
          <span className="landing-preview-actions"><i>＋</i><i>⌁</i><i>□</i><b>LaTeX 代码格式</b></span>
        </figcaption>
        <div className="landing-preview-body">
          <aside className="landing-preview-sidebar" aria-hidden="true">
            <strong>公式工具</strong>
            <div className="landing-preview-categories"><span className="is-active">☆ 常用</span><span>结构</span><span>微积分</span><span>矩阵</span></div>
            <div className="landing-preview-symbols">
              <span><b><i>1</i><i>x</i></b><small>分式</small></span>
              <span><b>√x</b><small>平方根</small></span>
              <span><b>x<sup>n</sup></b><small>上标</small></span>
              <span><b>x<sub>i</sub></b><small>下标</small></span>
              <span><b>∫f(x)</b><small>定积分</small></span>
              <span><b>Σa<sub>i</sub></b><small>求和</small></span>
            </div>
          </aside>
          <div className="landing-preview-workspace">
            <div className="landing-preview-workspace-head"><strong><Braces size={13} />可视化编辑</strong><span>−　100%　＋</span></div>
            <div className="landing-preview-canvas">
              <div className="landing-formula-card">
                <span className="landing-formula-index">01</span>
                <div className="landing-formula-expression">f(x) = <span className="landing-fraction"><span>1</span><span>σ√2π</span></span> e<sup>−(x−μ)²/2σ²</sup></div>
              </div>
              <button className="landing-add-row" aria-label="添加公式行">＋</button>
            </div>
            <div className="landing-source-panel"><Code2 size={12} /><span>展开 LaTeX 源码</span><i /></div>
          </div>
        </div>
      </figure>
    </div>
  );
}

export function LandingPage() {
  const { platform: detectedPlatform, isMobileDevice } = detectPlatform();
  const detectedPlatformName = downloads.find((download) => download.id === detectedPlatform)?.title;
  const orderedDownloads = [...downloads].sort(
    (left, right) => Number(right.id === detectedPlatform) - Number(left.id === detectedPlatform),
  );

  return (
    <div className="landing-page">
      <header className="landing-header">
        <div className="landing-header-inner">
          <a className="landing-brand" href="/" aria-label="VisualTeX 首页">
            <span className="landing-brand-mark"><VisualTeXLogo /></span>
            <span>VisualTeX</span>
          </a>

          <nav className="landing-nav" aria-label="主要导航">
            <a className="landing-nav-download" href="#download"><Download size={17} /><span>下载桌面端</span></a>
            <a className="landing-nav-editor" href="/editor"><span>打开网页端</span><ArrowRight size={17} /></a>
          </nav>
        </div>
      </header>

      <main>
        <section className="landing-hero">
          <div className="landing-hero-glow landing-hero-glow-one" />
          <div className="landing-hero-glow landing-hero-glow-two" />

          <div className="landing-hero-inner">
            <div className="landing-hero-copy">
              <div className="landing-eyebrow"><span className="landing-eyebrow-dot" />可视化 LaTeX 公式编辑器</div>
              <h1>让复杂公式编辑，<span>回归直觉。</span></h1>
              <p>VisualTeX 将结构化公式编辑、LaTeX 源码、本地公式 OCR 与 Office 工作流放在同一个界面中，让数学与科研写作更自然、更高效。</p>

              <div className="landing-hero-actions">
                <a className="landing-primary-action" href="/editor">立即使用网页端<ArrowRight size={19} /></a>
                <a className="landing-secondary-action" href="#download"><Download size={18} />下载桌面端</a>
              </div>

              <div className="landing-hero-meta">
                <span><Check size={15} />网页端无需安装</span>
                <span><Check size={15} />桌面端支持 Office 与本地 OCR</span>
              </div>
            </div>

            <EditorPreview />
          </div>
        </section>

        <section className="landing-section landing-features-section">
          <div className="landing-section-heading">
            <span>核心能力</span>
            <h2>一个更完整，也更克制的公式工作区</h2>
            <p>网页端负责快速编辑，桌面端进一步连接 Office、本地 OCR 与系统级工作流。</p>
          </div>

          <div className="landing-feature-grid">
            {features.map(({ icon: Icon, title, description }, index) => (
              <article className="landing-feature-card" key={title}>
                <div className="landing-feature-topline">
                  <span className="landing-feature-icon"><Icon size={21} /></span>
                  <span className="landing-feature-number">0{index + 1}</span>
                </div>
                <h3>{title}</h3>
                <p>{description}</p>
              </article>
            ))}
          </div>
        </section>

        <section className="landing-download-section" id="download">
          <div className="landing-section-heading landing-download-heading">
            <span>桌面应用</span>
            <h2>选择你的系统，直接开始使用</h2>
            <p>当前版本 v{VERSION}。安装包来自 VisualTeX 官方 GitHub Release。</p>
          </div>

          <p className="landing-device-note" role="status">
            {detectedPlatformName
              ? `已识别当前设备为 ${detectedPlatformName}，对应安装包已优先展示。`
              : isMobileDevice
                ? "移动设备可直接使用网页端；桌面安装包请在对应电脑上下载。"
                : "未能自动识别当前系统，请手动选择对应安装包。"}
          </p>

          <div className="landing-download-grid">
            {orderedDownloads.map((download) => {
              const Icon = download.icon;
              const recommended = download.id === detectedPlatform;
              return (
                <article
                  className={`landing-download-card${recommended ? " is-recommended" : ""}`}
                  key={download.id}
                  aria-label={recommended ? `${download.title}，当前设备推荐` : download.title}
                >
                  {recommended && <span className="landing-recommended-badge">为此设备推荐</span>}
                  <span className="landing-download-icon"><Icon size={25} /></span>
                  <h3>{download.title}</h3>
                  <p>{download.detail}</p>
                  <a className="landing-download-primary" href={download.href}><Download size={17} />{download.action}</a>
                  {download.secondaryHref && download.secondaryAction && (
                    <a className="landing-download-secondary" href={download.secondaryHref}>{download.secondaryAction}</a>
                  )}
                </article>
              );
            })}
          </div>

          <div className="landing-release-link">
            <Github size={18} />
            <span>需要旧版本、校验文件或安装说明？</span>
            <a href={RELEASES_URL} target="_blank" rel="noreferrer">查看 GitHub 全部发布<ArrowUpRight size={16} /></a>
          </div>
        </section>

      </main>

      <footer className="landing-footer">
        <div className="landing-footer-inner">
          <a className="landing-brand" href="/"><span className="landing-brand-mark"><VisualTeXLogo /></span><span>VisualTeX</span></a>
          <p>Visual LaTeX Formula Editor</p>
          <a href="https://github.com/paulhe666/visualtex" target="_blank" rel="noreferrer">GitHub<ArrowUpRight size={15} /></a>
        </div>
      </footer>
    </div>
  );
}

export default LandingPage;
