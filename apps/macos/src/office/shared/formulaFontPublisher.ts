import { invoke, isTauri } from "@tauri-apps/api/core";
import { FORMULA_FONT_PREFERENCES_CHANGED_EVENT } from "../../../../shared/visualtex-app/src/editor/formulaFontPreferences";
import { currentFormulaFontSettings } from "../../../../shared/visualtex-app/src/office/shared/formulaFontRefresh";

// The Office add-ins ask the app which VisualTeX formulas were drawn in other fonts than the ones it draws in now
// (Update Equation Format: the resident's <session>.fonts answer, the WPS bridge's /formulas/font-check). Those fonts
// are the Settings' in this app's storage, which only its windows can read: each window that may be the only one
// running (the main window, the resident Office editors) tells the backend at start and after every change, and
// the backend keeps them in office/formula-fonts.json for the next start (src-tauri office/formula_fonts.rs).

let lastPublished = "";

function publishCurrentFormulaFonts() {
  const fonts = currentFormulaFontSettings();
  const text = JSON.stringify(fonts);
  if (text === lastPublished) return;
  lastPublished = text;
  void invoke("macos_office_publish_formula_fonts", { fonts }).catch(() => {
    // published again at the next change or start
    lastPublished = "";
  });
}

/** Publishes the formula fonts now and after every change (in this window or another one). */
export function installFormulaFontPublisher(): () => void {
  if (typeof window === "undefined" || !isTauri()) return () => undefined;
  publishCurrentFormulaFonts();
  const onStorage = (event: StorageEvent) => {
    if (event.key === null || event.key.startsWith("visualtex.formula-")) publishCurrentFormulaFonts();
  };
  window.addEventListener(FORMULA_FONT_PREFERENCES_CHANGED_EVENT, publishCurrentFormulaFonts);
  window.addEventListener("storage", onStorage);
  return () => {
    window.removeEventListener(FORMULA_FONT_PREFERENCES_CHANGED_EVENT, publishCurrentFormulaFonts);
    window.removeEventListener("storage", onStorage);
  };
}
