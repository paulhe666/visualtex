/**
 * What an app that builds the shared code has besides the editor (its src/platform/appInfo.ts gives APP_FEATURES).
 * The desktop apps have all of it; the web edition (apps/web) runs in a browser and has none, so shared code leaves
 * out the controls, settings and shortcuts of the parts an app lacks.
 */
export interface AppFeatures {
  /** the Office and WPS add-ins and the AI agents' access to them (Settings → Office, AI agents; Office shortcuts) */
  office: boolean;
  /** VisualTeX Anywhere: formulas typed in other applications through system-wide shortcuts */
  anywhere: boolean;
  /** screenshot OCR from the screen: quick OCR and silent (background) OCR */
  screenOcr: boolean;
  /** OCR models that run on this computer (the recognizer menu lists them) */
  localOcr: boolean;
  /** update checks against the GitHub releases */
  updates: boolean;
  /** the keypad: the main window shrunk to the formula and its tools */
  keypadWindow: boolean;
}

export const DESKTOP_APP_FEATURES: AppFeatures = {
  office: true,
  anywhere: true,
  screenOcr: true,
  localOcr: true,
  updates: true,
  keypadWindow: true,
};
