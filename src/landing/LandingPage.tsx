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
  Sparkles,
} from "lucide-react";
import { VisualTeXLogo } from "../components/VisualTeXLogo";

const VERSION = "1.1.0";
const RELEASE_BASE = `https://github.com/paulhe666/visualtex/releases/download/v${VERSION}`;
const RELEASES_URL = "https://github.com/paulhe666/visualtex/releases";

const downloads = [
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
] as const;

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

function detectPlatform() {
  const userAgent = navigator.userAgent.toLowerCase();
  if (userAgent.includes("mac")) return "mac";
  if (userAgent.includes("win")) return "windows";
  if (userAgent.includes("linux")) return "linux";
  return "";
}

function EditorPreview() {
  return (
    <div className="landing-preview-wrap" aria-label="VisualTeX 编辑器界面预览">
      <div className="landing-orbit landing-orbit-one" />
      <div className="landing-orbit landing-orbit-two" />

      <div className="landing-preview-window">
        <div className="landing-preview-titlebar">
          <div className="landing-window-dots" aria-hidden="true">
            <span />
            <span />
            <span />
          </div>
          <div className="landing-preview-title">未命名公式</div>
          <div className="landing-preview-status">
            <span />
            已保存
          </div>
        </div>

        <div className="landing-preview-toolbar">
          <button type="button">x²</button>
          <button type="button">√</button>
          <button type="button">∑</button>
          <button type="button">∫</button>
          <button type="button">α</button>
          <button type="button">矩阵</button>
        </div>

        <div className="landing-preview-body">
          <aside className="landing-preview-sidebar">
            <span className="is-active">常用</span>
            <span>结构</span>
            <span>符号</span>
            <span>矩阵</span>
          </aside>

          <div className="landing-preview-canvas">
            <div className="landing-formula-card">
              <span className="landing-formula-index">01</span>
              <div className="landing-formula-expression">
                x = <span className="landing-fraction"><span>−b ± √(b² − 4ac)</span><span>2a</span></span>
              </div>
            </div>
            <div className="landing-formula-card is-secondary">
              <span className="landing-formula-index">02</span>
              <div className="landing-formula-expression">
                ∫<span className="landing-integral-limits"><sup>∞</sup><sub>−∞</sub></span>
                e<sup>−x²</sup> dx = √π
              </div>
            </div>
            <div className="landing-source-panel">
              <div className="landing-source-heading">
                <span>LaTeX 源码</span>
                <span>实时同步</span>
              </div>
              <code>{String.raw`x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}`}</code>
            </div>
          </div>
        </div>
      </div>

      <div className="landing-float-card landing-float-card-top">
        <Sparkles size={17} />
        <div>
          <strong>直观编辑</strong>
          <span>点击即可插入结构</span>
        </div>
      </div>

      <div className="landing-float-card landing-float-card-bottom">
        <Code2 size={17} />
        <div>
          <strong>双向同步</strong>
          <span>公式与源码始终一致</span>
        </div>
      </div>
    </div>
  );
}

export function LandingPage() {
  const detectedPlatform = detectPlatform();

  return (
    <div className="landing-page">
      <header className="landing-header">
        <div className="landing-header-inner">
          <a className="landing-brand" href="/" aria-label="VisualTeX 首页">
            <span className="landing-brand-mark">
              <VisualTeXLogo />
            </span>
            <span>VisualTeX</span>
          </a>

          <nav className="landing-nav" aria-label="主要导航">
            <a className="landing-nav-download" href="#download">
              <Download size={17} />
              <span>下载应用端</span>
            </a>
            <a className="landing-nav-editor" href="/editor">
              <span>网页端使用</span>
              <ArrowRight size={17} />
            </a>
          </nav>
        </div>
      </header>

      <main>
        <section className="landing-hero">
          <div className="landing-hero-glow landing-hero-glow-one" />
          <div className="landing-hero-glow landing-hero-glow-two" />

          <div className="landing-hero-inner">
            <div className="landing-hero-copy">
              <div className="landing-eyebrow">
                <span className="landing-eyebrow-dot" />
                可视化 LaTeX 公式编辑器
              </div>

              <h1>
                让复杂公式编辑，
                <span>回归直觉。</span>
              </h1>

              <p>
                VisualTeX 将结构化公式编辑、LaTeX 源码、本地公式 OCR 与 Office 工作流放在同一个界面中，让数学与科研写作更自然、更高效。
              </p>

              <div className="landing-hero-actions">
                <a className="landing-primary-action" href="/editor">
                  立即使用网页端
                  <ArrowRight size={19} />
                </a>
                <a className="landing-secondary-action" href="#download">
                  <Download size={18} />
                  下载桌面端
                </a>
              </div>

              <div className="landing-hero-meta">
                <span><Check size={15} /> 网页端无需安装</span>
                <span><Check size={15} /> 桌面端支持 Office 与本地 OCR</span>
              </div>
            </div>

            <EditorPreview />
          </div>
        </section>

        <section className="landing-proof-strip" aria-label="VisualTeX 适用场景">
          <div className="landing-proof-inner">
            <span>适用于</span>
            <strong>数学</strong>
            <i />
            <strong>物理</strong>
            <i />
            <strong>工程</strong>
            <i />
            <strong>教学</strong>
            <i />
            <strong>科研写作</strong>
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

        <section className="landing-workflow-section">
          <div className="landing-workflow-inner">
            <div className="landing-workflow-copy">
              <span>从输入到使用</span>
              <h2>公式不只需要写出来，还应该顺畅地进入你的文档。</h2>
              <p>在网页中快速完成公式，在桌面端接入 Word 与 PowerPoint。VisualTeX 把编辑、导出与再次修改连接成连续工作流。</p>
              <a href="/editor">
                打开网页编辑器
                <ArrowRight size={18} />
              </a>
            </div>

            <div className="landing-workflow-steps">
              <div><span>01</span><strong>可视化输入</strong><p>直接选择结构与符号，不必从空白源码开始。</p></div>
              <div><span>02</span><strong>同步检查</strong><p>公式和 LaTeX 源码始终可见，修改结果立即反馈。</p></div>
              <div><span>03</span><strong>插入文档</strong><p>桌面端连接 Word 与 PowerPoint，并保留再次编辑能力。</p></div>
            </div>
          </div>
        </section>

        <section className="landing-download-section" id="download">
          <div className="landing-section-heading landing-download-heading">
            <span>桌面应用</span>
            <h2>选择你的系统，直接开始使用</h2>
            <p>当前版本 v{VERSION}。安装包来自 VisualTeX 官方 GitHub Release。</p>
          </div>

          <div className="landing-download-grid">
            {downloads.map(({ id, icon: Icon, title, detail, href, action, secondaryHref, secondaryAction }) => {
              const recommended = id === detectedPlatform;
              return (
                <article className={`landing-download-card${recommended ? " is-recommended" : ""}`} key={id}>
                  {recommended && <span className="landing-recommended-badge">为此设备推荐</span>}
                  <span className="landing-download-icon"><Icon size={25} /></span>
                  <h3>{title}</h3>
                  <p>{detail}</p>
                  <a className="landing-download-primary" href={href}>
                    <Download size={17} />
                    {action}
                  </a>
                  {secondaryHref && secondaryAction && (
                    <a className="landing-download-secondary" href={secondaryHref}>
                      {secondaryAction}
                    </a>
                  )}
                </article>
              );
            })}
          </div>

          <div className="landing-release-link">
            <Github size={18} />
            <span>需要旧版本、校验文件或安装说明？</span>
            <a href={RELEASES_URL} target="_blank" rel="noreferrer">
              查看 GitHub 全部发布
              <ArrowUpRight size={16} />
            </a>
          </div>
        </section>

        <section className="landing-final-cta">
          <div className="landing-final-cta-mark"><VisualTeXLogo /></div>
          <div>
            <span>无需安装</span>
            <h2>现在就在浏览器中创建你的第一个公式。</h2>
          </div>
          <a href="/editor">
            进入网页端
            <ArrowRight size={19} />
          </a>
        </section>
      </main>

      <footer className="landing-footer">
        <div className="landing-footer-inner">
          <a className="landing-brand" href="/">
            <span className="landing-brand-mark"><VisualTeXLogo /></span>
            <span>VisualTeX</span>
          </a>
          <p>Visual LaTeX Formula Editor</p>
          <a href="https://github.com/paulhe666/visualtex" target="_blank" rel="noreferrer">
            GitHub
            <ArrowUpRight size={15} />
          </a>
        </div>
      </footer>
    </div>
  );
}

export default LandingPage;
