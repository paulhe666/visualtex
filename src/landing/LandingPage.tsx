import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { ART_DEFS, ART_FIELD, ART_HERO, ART_MARKS, ART_RULE, ART_TEXT, type InkKey } from "./art.generated";
import { layoutFormulaField } from "./formulaField";
import { applyLandingDocumentMeta, detectLandingLang, saveLandingLang, type LandingLang } from "./i18n";
import { SupportCodes } from "./SupportCodes";

const VERSION = "1.2.7";
const DOWNLOAD_BASE = `https://download.visualtex.pauljianliao.com/visualtex-downloads/releases/v${VERSION}`;
const OCR_MODEL_BASE = "https://download.visualtex.pauljianliao.com/ppformula-model";
const RELEASES_URL = "https://github.com/paulhe666/visualtex/releases";
const REPO_URL = "https://github.com/paulhe666/visualtex";

type PlatformId = "mac" | "windows";

const downloads = [
  {
    id: "mac",
    name: "macName",
    meta: "macMeta",
    href: `${DOWNLOAD_BASE}/VisualTeX_${VERSION}_aarch64.dmg`,
    secondaryHref: `${DOWNLOAD_BASE}/VisualTeX_${VERSION}_aarch64-no-ocr.dmg`,
  },
  {
    id: "windows",
    name: "winName",
    meta: "winMeta",
    href: `${DOWNLOAD_BASE}/VisualTeX_${VERSION}_x64-setup.exe`,
    secondaryHref: `${DOWNLOAD_BASE}/VisualTeX_${VERSION}_x64-no-ocr-setup.exe`,
  },
] as const satisfies readonly { id: PlatformId; name: InkKey; meta: InkKey; href: string; secondaryHref: string }[];

const ocrModels = [
  { label: "ocrS", href: `${OCR_MODEL_BASE}/VisualTeX_PP-FormulaNet_plus-S_windows-x64.vtxocrmodel` },
  { label: "ocrM", href: `${OCR_MODEL_BASE}/VisualTeX_PP-FormulaNet_plus-M_windows-x64.vtxocrmodel` },
  { label: "ocrL", href: `${OCR_MODEL_BASE}/VisualTeX_PP-FormulaNet_plus-L_windows-x64.vtxocrmodel` },
] as const satisfies readonly { label: InkKey; href: string }[];

const features = ["feature1", "feature2", "feature3", "feature4", "feature5"] as const satisfies readonly InkKey[];

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

/** A string drawn as brush lettering (pre-rendered SVG with its own aria-label). */
function Ink({ k, lang }: { k: InkKey; lang: LandingLang }) {
  return <span className="landing-ink" dangerouslySetInnerHTML={{ __html: ART_TEXT[k][lang] }} />;
}

function Svg({ html, className }: { html: string; className?: string }) {
  return <span className={className} aria-hidden="true" dangerouslySetInnerHTML={{ __html: html }} />;
}

const arrow = (
  <svg className="landing-arrow" viewBox="0 0 24 24" aria-hidden="true">
    <path d="M4 12h15M13 6l6 6-6 6" />
  </svg>
);

/** A white sheet styled like a LaTeX theorem environment: label, equation number, QED box. */
function Card({ n, label, lang, children }: { n: number; label: InkKey; lang: LandingLang; children: ReactNode }) {
  return (
    <article className="landing-card" data-clear data-reveal>
      <header className="landing-card-head">
        <Ink k={label} lang={lang} />
        <Svg className="landing-eqno" html={ART_MARKS.eq[n - 1]} />
      </header>
      <Svg className="landing-rule" html={ART_RULE} />
      <div className="landing-card-body">{children}</div>
      <svg className="landing-qed" viewBox="0 0 16 16" aria-hidden="true"><rect x="1.5" y="1.5" width="13" height="13" /></svg>
    </article>
  );
}

function EditorPreview({ lang }: { lang: LandingLang }) {
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
    <>
      <div className="landing-frame" ref={viewportRef}>
        <iframe
          ref={frameRef}
          src="/editor?landing-preview=1"
          title={lang === "zh" ? "VisualTeX 网页编辑器预览" : "VisualTeX web editor preview"}
          loading="lazy"
          tabIndex={-1}
          aria-hidden="true"
          inert
        />
      </div>
      <div className="landing-actions">
        <a className="landing-btn landing-btn-solid" href="/editor"><Ink k="openEditor" lang={lang} />{arrow}</a>
      </div>
    </>
  );
}

export function LandingPage() {
  const [lang, setLang] = useState<LandingLang>(detectLandingLang);
  const pageRef = useRef<HTMLDivElement>(null);
  const fieldRef = useRef<HTMLDivElement>(null);
  const { platform: detectedPlatform, isMobileDevice } = detectPlatform();
  const orderedDownloads = [...downloads].sort(
    (left, right) => Number(right.id === detectedPlatform) - Number(left.id === detectedPlatform),
  );

  useEffect(() => applyLandingDocumentMeta(lang), [lang]);

  // The formula field is one continuous layer behind the whole page; re-pack it whenever the
  // page width or the language (and so the size of the foreground) changes.
  useLayoutEffect(() => {
    const page = pageRef.current, field = fieldRef.current;
    if (!page || !field) return;
    let width = -1, timer = 0;
    const build = () => { width = page.clientWidth; layoutFormulaField(page, field, ART_FIELD); };
    build();
    const observer = new ResizeObserver(() => {
      if (page.clientWidth === width) return;
      window.clearTimeout(timer);
      timer = window.setTimeout(build, 120);
    });
    observer.observe(page);
    return () => { observer.disconnect(); window.clearTimeout(timer); };
  }, [lang]);

  // Cards rise into place as they enter the viewport.
  useEffect(() => {
    const page = pageRef.current;
    if (!page) return;
    const observer = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        if (entry.isIntersecting) {
          entry.target.classList.add("is-in");
          observer.unobserve(entry.target);
        }
      }
    }, { rootMargin: "0px 0px -12% 0px" });
    page.querySelectorAll("[data-reveal]").forEach((el) => observer.observe(el));
    return () => observer.disconnect();
  }, []);

  const toggleLang = () => {
    const next = lang === "zh" ? "en" : "zh";
    saveLandingLang(next);
    setLang(next);
  };

  return (
    <div className="landing-page" ref={pageRef} lang={lang === "zh" ? "zh-CN" : "en"}>
      <svg className="landing-defs" aria-hidden="true" dangerouslySetInnerHTML={{ __html: ART_DEFS }} />
      <div className="landing-field" ref={fieldRef} aria-hidden="true" />
      <a className="landing-skip" href="#main">{lang === "zh" ? "跳转到正文" : "Skip to content"}</a>

      <nav className="landing-nav" aria-label={lang === "zh" ? "主要导航" : "Main navigation"}>
        <a href="/editor"><Ink k="navEditor" lang={lang} /></a>
        <a href="#download"><Ink k="navDownload" lang={lang} /></a>
        <button type="button" className="landing-lang" onClick={toggleLang}><Ink k="langSwitch" lang={lang} /></button>
      </nav>

      <div className="landing-fg">
        <header className="landing-hero">
          <h1 className="landing-word" dangerouslySetInnerHTML={{ __html: ART_HERO }} />
        </header>

        <main id="main">
          <section className="landing-section"><p className="landing-say" data-clear><Ink k="say1" lang={lang} /></p></section>

          <section className="landing-section">
            <Card n={1} label="figLabel" lang={lang}><EditorPreview lang={lang} /></Card>
          </section>

          <section className="landing-section"><p className="landing-say" data-clear><Ink k="say2" lang={lang} /></p></section>

          <section className="landing-section">
            <Card n={2} label="propLabel" lang={lang}>
              <ul className="landing-features">
                {features.map((key, index) => (
                  <li key={key}><Svg className="landing-num" html={ART_MARKS.roman[index]} /><Ink k={key} lang={lang} /></li>
                ))}
              </ul>
            </Card>
          </section>

          <section className="landing-section"><p className="landing-say" data-clear><Ink k="say3" lang={lang} /></p></section>

          <section className="landing-section" id="download">
            <Card n={3} label="thmLabel" lang={lang}>
              {isMobileDevice && <p className="landing-device-note"><Ink k="mobileNote" lang={lang} /></p>}
              <div className="landing-platforms">
                {orderedDownloads.map((download) => (
                  <div className="landing-platform" key={download.id}>
                    <div className="landing-platform-name">
                      <Ink k={download.name} lang={lang} />
                      {download.id === detectedPlatform && <span className="landing-device"><Ink k="thisDevice" lang={lang} /></span>}
                    </div>
                    <div className="landing-platform-meta"><Ink k={download.meta} lang={lang} /></div>
                    <div className="landing-platform-actions">
                      <a className="landing-btn landing-btn-solid" href={download.href}><Ink k="full" lang={lang} /></a>
                      <a className="landing-btn landing-btn-ghost" href={download.secondaryHref}><Ink k="lite" lang={lang} /></a>
                    </div>
                  </div>
                ))}
              </div>
              <div className="landing-models">
                <span className="landing-models-cap"><Ink k="ocrCaption" lang={lang} /></span>
                {ocrModels.map((model) => (
                  <a className="landing-chip" key={model.label} href={model.href}><Ink k={model.label} lang={lang} /></a>
                ))}
              </div>
              <a className="landing-releases" href={RELEASES_URL} target="_blank" rel="noreferrer"><Ink k="allReleases" lang={lang} />{arrow}</a>
            </Card>
          </section>

          <section className="landing-section">
            <Card n={4} label="remLabel" lang={lang}>
              <SupportCodes lang={lang} />
              <p className="landing-note"><Ink k="tips" lang={lang} /><Ink k="group" lang={lang} /></p>
            </Card>
          </section>
        </main>

        <footer className="landing-footer" data-clear>
          <Ink k="footerName" lang={lang} />
          <a href={REPO_URL} target="_blank" rel="noreferrer"><Ink k="footerGithub" lang={lang} /></a>
        </footer>
      </div>
    </div>
  );
}

export default LandingPage;
