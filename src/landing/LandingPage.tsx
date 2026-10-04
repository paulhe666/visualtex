import { useEffect, useLayoutEffect, useRef, useState, type MouseEvent, type ReactNode, type RefObject } from "react";
import { ART_DEFS, ART_FIELD, ART_HERO, ART_MARKS, ART_RULE, ART_SHAPES, ART_TEXT, type InkKey } from "./art.generated";
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

// Nav entries double as cross-references to the numbered sheets, like \ref in a paper.
const navRefs = [
  { href: "#features", label: "navFeatures", ref: "refProp" },
  { href: "#download", label: "navDownload", ref: "refThm" },
] as const satisfies readonly { href: string; label: InkKey; ref: InkKey }[];

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

/**
 * A link drawn as a brush stroke (solid) or a sketched outline (ghost). The stroke paints in
 * when its sheet appears and repaints on hover; a press leaves a blot (see useInkSplash).
 */
function BrushButton({ href, kind, seed, children, ...rest }: {
  href: string; kind: "solid" | "ghost"; seed: number; children: ReactNode; download?: boolean; target?: string; rel?: string;
}) {
  const shape = ART_SHAPES.swash[seed % ART_SHAPES.swash.length];
  return (
    <a className={`landing-btn landing-btn-${kind}`} href={href} {...rest}>
      <Svg className="landing-btn-swash" html={shape} />
      {kind === "solid" && <Svg className="landing-btn-swash landing-btn-swash-over" html={ART_SHAPES.swash[(seed + 1) % ART_SHAPES.swash.length]} />}
      {kind === "ghost" && <Svg className="landing-btn-box" html={ART_SHAPES.box[seed % ART_SHAPES.box.length]} />}
      <span className="landing-btn-label">{children}</span>
    </a>
  );
}

/** A white sheet styled like a LaTeX theorem environment: label, equation number, QED box. */
function Card({ n, label, lang, children }: { n: number; label: InkKey; lang: LandingLang; children: ReactNode }) {
  return (
    <article className="landing-card" data-clear data-reveal>
      <header className="landing-card-head">
        <Ink k={label} lang={lang} />
        <span className="landing-eqno">
          <Svg html={ART_MARKS.eq[n - 1]} />
          <Svg className="landing-eqno-loop" html={ART_SHAPES.loop[n % ART_SHAPES.loop.length]} />
        </span>
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
        <BrushButton href="/editor" kind="solid" seed={2}><Ink k="openEditor" lang={lang} />{arrow}</BrushButton>
      </div>
    </>
  );
}

/** Wipes the hero lettering in, letters first and then the long sweep. */
function useHeroWriting(hero: RefObject<HTMLElement | null>) {
  useLayoutEffect(() => {
    const svg = hero.current?.querySelector("svg");
    if (!svg || window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    const groups = [...svg.querySelectorAll<SVGGElement>("g[data-wipe]")].map((g) => {
      const [from, to] = g.dataset.wipe!.split(" ").map(Number);
      const id = g.getAttribute("mask")!.slice(5, -1);
      const mask = svg.querySelector(`#${id}`)!;
      return { from, to, body: mask.querySelector<SVGRectElement>(".wipe-body")!, edge: mask.querySelector<SVGRectElement>(".wipe-edge")! };
    });
    const timing = [{ start: 250, ms: 1500 }, { start: 1550, ms: 650 }];
    const ease = (t: number) => 1 - Math.pow(1 - t, 2.2);
    const set = (g: (typeof groups)[number], x: number) => {
      g.body.setAttribute("width", String(Math.max(0, x - Number(g.body.getAttribute("x")))));
      g.edge.setAttribute("x", String(x));
    };
    groups.forEach((g) => set(g, g.from));
    let frame = 0;
    const t0 = performance.now();
    const tick = (now: number) => {
      let done = true;
      groups.forEach((g, i) => {
        const t = Math.min(1, Math.max(0, (now - t0 - timing[i].start) / timing[i].ms));
        if (t < 1) done = false;
        set(g, g.from + (g.to + 200 - g.from) * ease(t));
      });
      if (!done) frame = requestAnimationFrame(tick);
    };
    frame = requestAnimationFrame(tick);
    return () => {
      cancelAnimationFrame(frame);
      groups.forEach((g) => set(g, g.to + 400));
    };
  }, [hero]);
}

/** Pressing a brush button leaves an ink blot that spreads and dries. */
function useInkSplash(page: RefObject<HTMLElement | null>) {
  useEffect(() => {
    const root = page.current;
    if (!root || window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    let count = 0;
    const onDown = (event: PointerEvent) => {
      const button = (event.target as Element).closest<HTMLElement>(".landing-btn");
      if (!button) return;
      const rect = button.getBoundingClientRect();
      const blot = document.createElement("span");
      blot.className = "landing-blot";
      blot.innerHTML = ART_SHAPES.blot[count++ % ART_SHAPES.blot.length];
      blot.style.left = `${event.clientX - rect.left}px`;
      blot.style.top = `${event.clientY - rect.top}px`;
      button.appendChild(blot);
      blot.addEventListener("animationend", () => blot.remove(), { once: true });
    };
    root.addEventListener("pointerdown", onDown);
    return () => root.removeEventListener("pointerdown", onDown);
  }, [page]);
}

export function LandingPage() {
  const [lang, setLang] = useState<LandingLang>(detectLandingLang);
  const [active, setActive] = useState("");
  const [floating, setFloating] = useState(false);
  const pageRef = useRef<HTMLDivElement>(null);
  const fieldRef = useRef<HTMLDivElement>(null);
  const heroRef = useRef<HTMLElement>(null);
  const { platform: detectedPlatform, isMobileDevice } = detectPlatform();
  const orderedDownloads = [...downloads].sort(
    (left, right) => Number(right.id === detectedPlatform) - Number(left.id === detectedPlatform),
  );

  useEffect(() => applyLandingDocumentMeta(lang), [lang]);
  useHeroWriting(heroRef);
  useInkSplash(pageRef);

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

  // Sheets and margin notes draw themselves in as they enter the viewport; a sheet's QED box
  // fills once it has been read to the end.
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
    const ends = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        if (entry.isIntersecting) {
          entry.target.closest(".landing-card")?.classList.add("is-done");
          ends.unobserve(entry.target);
        }
      }
    }, { rootMargin: "0px 0px -18% 0px" });
    page.querySelectorAll("[data-reveal]").forEach((el) => observer.observe(el));
    page.querySelectorAll(".landing-qed").forEach((el) => ends.observe(el));
    return () => { observer.disconnect(); ends.disconnect(); };
  }, [lang]);

  // Nav: becomes a paper slip once the hero is behind; the sheet in the middle of the screen
  // is underlined in it.
  useEffect(() => {
    const page = pageRef.current;
    if (!page) return;
    const onScroll = () => setFloating(window.scrollY > window.innerHeight * 0.6);
    onScroll();
    window.addEventListener("scroll", onScroll, { passive: true });
    const observer = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        if (entry.isIntersecting) setActive(`#${entry.target.id}`);
        else setActive((current) => (current === `#${entry.target.id}` ? "" : current));
      }
    }, { rootMargin: "-45% 0px -45% 0px" });
    navRefs.forEach(({ href }) => { const el = page.querySelector(href); if (el) observer.observe(el); });
    return () => { window.removeEventListener("scroll", onScroll); observer.disconnect(); };
  }, []);

  // Following a reference circles the equation number it points to, as a reader would.
  const followRef = (event: MouseEvent<HTMLAnchorElement>) => {
    const card = pageRef.current?.querySelector(`${event.currentTarget.getAttribute("href")} .landing-card`);
    if (!card) return;
    card.classList.remove("is-ref");
    void (card as HTMLElement).offsetWidth;
    card.classList.add("is-ref");
    window.setTimeout(() => card.classList.remove("is-ref"), 3200);
  };

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

      <nav className={`landing-nav${floating ? " is-floating" : ""}`} data-clear="fixed" aria-label={lang === "zh" ? "主要导航" : "Main navigation"}>
        <a href="/editor"><Ink k="navEditor" lang={lang} /></a>
        {navRefs.map(({ href, label, ref }) => (
          <a key={href} href={href} onClick={followRef} className={active === href ? "is-active" : undefined}>
            <Ink k={label} lang={lang} />
            <Svg className="landing-nav-ref" html={ART_TEXT[ref][lang]} />
            <Svg className="landing-nav-rule" html={ART_RULE} />
          </a>
        ))}
        <button type="button" className="landing-lang" onClick={toggleLang}><Ink k="langSwitch" lang={lang} /></button>
      </nav>

      <div className="landing-fg">
        <header className="landing-hero" ref={heroRef}>
          <h1 className="landing-word" dangerouslySetInnerHTML={{ __html: ART_HERO }} />
          <div className="landing-hero-cta" data-clear>
            <div className="landing-hero-note" aria-hidden="true">
              <Ink k="noteHero" lang={lang} />
              <Svg className="landing-anno-arrow" html={ART_SHAPES.arrows.hero.svg} />
            </div>
            <div className="landing-hero-buttons">
              <BrushButton href="/editor" kind="solid" seed={0}><Ink k="heroOpen" lang={lang} />{arrow}</BrushButton>
              <BrushButton href="#download" kind="ghost" seed={1}><Ink k="heroDownload" lang={lang} /></BrushButton>
            </div>
          </div>
        </header>

        <main id="main">
          <section className="landing-section"><p className="landing-say" data-clear><Ink k="say1" lang={lang} /></p></section>

          <section className="landing-section">
            <Card n={1} label="figLabel" lang={lang}>
              <EditorPreview lang={lang} />
            </Card>
          </section>

          <section className="landing-section"><p className="landing-say" data-clear><Ink k="say2" lang={lang} /></p></section>

          <section className="landing-section" id="features">
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
                {orderedDownloads.map((download, index) => (
                  <div className="landing-platform" key={download.id}>
                    <div className="landing-platform-name">
                      <Ink k={download.name} lang={lang} />
                      {download.id === detectedPlatform && <span className="landing-device"><Ink k="thisDevice" lang={lang} /></span>}
                    </div>
                    <div className="landing-platform-meta"><Ink k={download.meta} lang={lang} /></div>
                    <div className="landing-platform-actions">
                      <BrushButton href={download.href} kind="solid" seed={index + 3}><Ink k="full" lang={lang} /></BrushButton>
                      <BrushButton href={download.secondaryHref} kind="ghost" seed={index + 1}><Ink k="lite" lang={lang} /></BrushButton>
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
              <a className="landing-releases" href={RELEASES_URL} target="_blank" rel="noreferrer">
                <Ink k="allReleases" lang={lang} />{arrow}
                <Svg className="landing-nav-rule" html={ART_RULE} />
              </a>
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
          <a href={REPO_URL} target="_blank" rel="noreferrer"><Ink k="footerGithub" lang={lang} /><Svg className="landing-nav-rule" html={ART_RULE} /></a>
        </footer>
      </div>
    </div>
  );
}

export default LandingPage;
