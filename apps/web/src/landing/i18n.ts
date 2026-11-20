export type LandingLang = "zh" | "en";

const STORAGE_KEY = "visualtex.landing.lang";

/** Saved choice first, then the browser's preferred languages: any zh-* means Chinese. */
export function detectLandingLang(): LandingLang {
  try {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    if (saved === "zh" || saved === "en") return saved;
  } catch {
    // Storage can be unavailable (privacy mode); fall back to the browser language.
  }
  const preferred = navigator.languages?.length ? navigator.languages : [navigator.language];
  return preferred[0]?.toLowerCase().startsWith("zh") ? "zh" : "en";
}

export function saveLandingLang(lang: LandingLang) {
  try {
    window.localStorage.setItem(STORAGE_KEY, lang);
  } catch {
    // Ignore: the choice simply will not persist.
  }
}

const META: Record<LandingLang, { htmlLang: string; title: string; description: string }> = {
  zh: {
    htmlLang: "zh-CN",
    title: "VisualTeX — 可视化 LaTeX 公式编辑器",
    description:
      "VisualTeX 是面向数学、物理、工程、教学与科研写作的可视化 LaTeX 公式编辑器，提供网页端、桌面端、本地公式 OCR 与 Office 工作流。",
  },
  en: {
    htmlLang: "en",
    title: "VisualTeX — Visual LaTeX Formula Editor",
    description:
      "VisualTeX is a visual LaTeX formula editor for math, physics, engineering, teaching and research writing, with a web editor, desktop apps, local formula OCR and Office workflows.",
  },
};

export function applyLandingDocumentMeta(lang: LandingLang) {
  const meta = META[lang];
  document.documentElement.lang = meta.htmlLang;
  document.title = meta.title;
  const description = document.querySelector<HTMLMetaElement>('meta[name="description"]');
  if (description) description.content = meta.description;
}
