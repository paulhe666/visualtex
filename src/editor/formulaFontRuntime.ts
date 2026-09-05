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

function runtimeCss() {
  const { letter, chinese } = currentFamilies();
  return `
.math-preview .ML__mathit,
.math-preview .ML__mathnormal,
.math-preview .ML__mathbf,
.math-preview .ML__mathbfit,
.math-preview .ML__mathrm,
.math-preview .ML__operator_name,
.math-preview .ML__lcGreek,
.math-preview .ML__ucGreek,
.math-preview .ML__latin {
  font-family: var(--visualtex-formula-italic-font-family, ${letter.italic});
}
.math-preview .ML__mathrm,
.math-preview .ML__operator_name {
  font-family: var(--visualtex-formula-upright-font-family, ${letter.upright});
}
.math-preview .ML__text,
.math-preview .ML__text span {
  font-family: var(--visualtex-formula-chinese-font-family, ${chinese}), var(--visualtex-formula-upright-font-family, ${letter.upright});
}
`;
}

function shadowCss() {
  const { letter, chinese } = currentFamilies();
  return `
.ML__mathit,
.ML__mathnormal,
.ML__mathbf,
.ML__mathbfit,
.ML__lcGreek,
.ML__ucGreek,
.ML__latin {
  font-family: ${letter.italic} !important;
}
.ML__mathrm,
.ML__operator_name {
  font-family: ${letter.upright} !important;
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
