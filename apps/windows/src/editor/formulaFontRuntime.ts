import type { MathfieldElement } from "mathlive";
import {
  FORMULA_FONT_PREFERENCES_CHANGED_EVENT,
  formulaChineseFontFamily,
  formulaLetterFontFamilies,
  readPersistedFormulaFontPreferences,
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
} from "./formulaFontPreferences";

const globalStyleId = "visualtex-formula-font-runtime-style";
const shadowStyleId = "visualtex-formula-font-runtime-shadow-style";
let installed = false;

function currentFamilies() {
  const preferences = readPersistedFormulaFontPreferences();
  const letter = formulaLetterFontFamilies(
    preferences.formulaLetterFont ?? DEFAULT_FORMULA_LETTER_FONT,
  );
  const chinese = formulaChineseFontFamily(
    preferences.formulaChineseFont ?? DEFAULT_FORMULA_CHINESE_FONT,
  );
  return { letter, chinese };
}

export function buildFormulaPreviewFontCss(
  uprightFamily: string,
  italicFamily: string,
  chineseFamily: string,
) {
  return `
/* Font preferences may substitute a family, but they must never normalize
   MathLive's mathematical variant. Keep the same variant contract as the
   editable MathEditor: variables are italic, \\mathbf is bold upright and
   \\mathbfit/bold Greek are bold italic. Dedicated script/fraktur/blackboard/
   monospace/sans-serif classes are intentionally left to MathLive. */
.math-preview .ML__mathit,
.math-preview .ML__mathnormal,
.math-preview .ML__lcGreek,
.math-preview .ML__ucGreek,
.math-preview .ML__latin {
  font-family: var(--visualtex-formula-italic-font-family, ${italicFamily});
  font-style: italic;
}
.math-preview .ML__mathbf:not(.lcGreek) {
  font-family: var(--visualtex-formula-upright-font-family, ${uprightFamily});
  font-style: normal;
  font-weight: 700;
}
.math-preview .ML__mathbfit,
.math-preview .lcGreek.ML__mathbf {
  font-family: var(--visualtex-formula-italic-font-family, ${italicFamily});
  font-style: italic;
  font-weight: 700;
}
.math-preview .ML__mathrm,
.math-preview .ML__operator_name {
  font-family: var(--visualtex-formula-upright-font-family, ${uprightFamily});
  font-style: normal;
}
.math-preview .ML__text,
.math-preview .ML__text span {
  font-family: var(--visualtex-formula-chinese-font-family, ${chineseFamily}), var(--visualtex-formula-upright-font-family, ${uprightFamily});
}
`;
}

function runtimeCss() {
  const { letter, chinese } = currentFamilies();
  return buildFormulaPreviewFontCss(letter.upright, letter.italic, chinese);
}

function shadowCss() {
  const { letter, chinese } = currentFamilies();
  return `
.ML__mathit,
.ML__mathnormal,
.ML__lcGreek,
.ML__ucGreek,
.ML__latin {
  font-family: ${letter.italic} !important;
  font-style: italic !important;
}
.ML__mathbf:not(.lcGreek) {
  font-family: ${letter.upright} !important;
  font-style: normal !important;
  font-weight: 700 !important;
}
.ML__mathbfit,
.lcGreek.ML__mathbf {
  font-family: ${letter.italic} !important;
  font-style: italic !important;
  font-weight: 700 !important;
}
.ML__mathrm,
.ML__operator_name {
  font-family: ${letter.upright} !important;
  font-style: normal !important;
}
.ML__text,
.ML__text span,
.ML__textord {
  font-family: ${chinese}, ${letter.upright} !important;
}
`;
}

function installGlobalStyle() {
  if (typeof document === "undefined") return;
  let style = document.getElementById(globalStyleId) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement("style");
    style.id = globalStyleId;
    document.head.append(style);
  }
  style.textContent = runtimeCss();
}

function installShadowStyle(field: MathfieldElement) {
  const root = field.shadowRoot;
  if (!root) return;
  let style = root.getElementById(shadowStyleId) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement("style");
    style.id = shadowStyleId;
    root.append(style);
  }
  style.textContent = shadowCss();
}

function refreshAllMathfields() {
  installGlobalStyle();
  if (typeof document === "undefined") return;
  document
    .querySelectorAll<MathfieldElement>("math-field")
    .forEach(installShadowStyle);
}

export function installFormulaFontRuntime() {
  if (installed || typeof window === "undefined" || typeof document === "undefined") {
    return;
  }
  installed = true;
  const refresh = () => window.requestAnimationFrame(refreshAllMathfields);
  window.addEventListener(FORMULA_FONT_PREFERENCES_CHANGED_EVENT, refresh);
  const observer = new MutationObserver((records) => {
    for (const record of records) {
      for (const node of Array.from(record.addedNodes)) {
        if (!(node instanceof Element)) continue;
        if (node.matches("math-field")) {
          window.requestAnimationFrame(() => installShadowStyle(node as MathfieldElement));
        }
        node.querySelectorAll?.<MathfieldElement>("math-field").forEach((field) => {
          window.requestAnimationFrame(() => installShadowStyle(field));
        });
      }
    }
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });
  refreshAllMathfields();
}

installFormulaFontRuntime();
