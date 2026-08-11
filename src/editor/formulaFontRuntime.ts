import type { MathfieldElement } from "mathlive";
import { useEditorStore } from "../stores/editorStore";
import {
  formulaChineseFontFamily,
  formulaLetterFontFamilies,
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "./formulaFontPreferences";

const GLOBAL_STYLE_ID = "visualtex-formula-font-runtime-style";
const SHADOW_STYLE_ID = "visualtex-formula-font-runtime-shadow-style";
let installed = false;
let observer: MutationObserver | null = null;
let unsubscribeStore: (() => void) | null = null;

function cssText(letterFont: FormulaLetterFont, chineseFont: FormulaChineseFont) {
  const letter = formulaLetterFontFamilies(letterFont);
  const chinese = formulaChineseFontFamily(chineseFont);
  return `
:root {
  --visualtex-formula-upright-font-family: ${letter.upright};
  --visualtex-formula-italic-font-family: ${letter.italic};
  --visualtex-formula-chinese-font-family: ${chinese};
}
math-field {
  --_text-font-family: var(--visualtex-formula-chinese-font-family);
}
.ML__cmr, .ML__mathbf {
  font-family: var(--visualtex-formula-upright-font-family) !important;
}
.ML__mathit, .ML__mathbfit, .lcGreek.ML__mathbf {
  font-family: var(--visualtex-formula-italic-font-family) !important;
}
.ML__text {
  font-family: var(--visualtex-formula-chinese-font-family) !important;
}
`;
}

function currentFonts() {
  const state = useEditorStore.getState();
  return {
    letterFont: state.formulaLetterFont,
    chineseFont: state.formulaChineseFont,
  };
}

function applyGlobalStyle() {
  if (typeof document === "undefined") return;
  const { letterFont, chineseFont } = currentFonts();
  let style = document.getElementById(GLOBAL_STYLE_ID) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement("style");
    style.id = GLOBAL_STYLE_ID;
    document.head.append(style);
  }
  const css = cssText(letterFont, chineseFont);
  if (style.textContent !== css) style.textContent = css;
}

function applyShadowStyle(field: MathfieldElement) {
  const root = field.shadowRoot;
  if (!root) return;
  const { letterFont, chineseFont } = currentFonts();
  let style = root.getElementById(SHADOW_STYLE_ID) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement("style");
    style.id = SHADOW_STYLE_ID;
    root.append(style);
  }
  const css = cssText(letterFont, chineseFont);
  if (style.textContent !== css) style.textContent = css;
}

function refreshMathfields() {
  if (typeof document === "undefined") return;
  applyGlobalStyle();
  document.querySelectorAll<MathfieldElement>("math-field").forEach((field) => {
    applyShadowStyle(field);
  });
}

export function installFormulaFontRuntime() {
  if (installed || typeof document === "undefined") return;
  installed = true;
  refreshMathfields();
  unsubscribeStore = useEditorStore.subscribe((state, previous) => {
    if (
      state.formulaLetterFont !== previous.formulaLetterFont ||
      state.formulaChineseFont !== previous.formulaChineseFont
    ) {
      refreshMathfields();
    }
  });
  observer = new MutationObserver(() => refreshMathfields());
  observer.observe(document.documentElement, { childList: true, subtree: true });
}

export function disposeFormulaFontRuntime() {
  observer?.disconnect();
  observer = null;
  unsubscribeStore?.();
  unsubscribeStore = null;
  installed = false;
}

installFormulaFontRuntime();
