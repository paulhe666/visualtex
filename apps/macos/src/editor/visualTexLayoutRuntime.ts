import { setVisualTexEditingCore, setVisualTexLayoutEngine } from "mathlive";
import { createVisualTexLayout, type LayoutData } from "../../../shared/tex-layout/src/index.ts";
import { createEditingCore } from "../../../shared/visualtex-core/src/index.ts";
import tfm from "../../../shared/tex-layout/data/tfm.json";
import meanings from "../../../shared/tex-layout/data/latex-meanings.json";
import templates from "../../../shared/tex-layout/data/latex-templates.json";
import onearg from "../../../shared/tex-layout/data/latex-onearg.json";
import textFonts from "../../../shared/tex-layout/data/text-fonts.json";
import fonts from "../../../shared/tex-layout/data/fonts.json";
import "../../../shared/tex-layout/data/fonts/fonts.css";
import profile from "../../../shared/tex-layout/data/profile.json";
import type NtxData from "./ntxProfileData";
import type { PathGlyphSet } from "../../../shared/tex-layout/src/path-font.ts";
import type { NativeMeasure } from "../../../shared/tex-layout/src/native-font.ts";
import { gunzipSync, strFromU8 } from "fflate";
import { RARE_INTEGRAL_GLYPHS_GZIP_BASE64 } from "../math/rareIntegralGlyphs.generatedData";
import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  type FormulaChineseFont,
  type FormulaLetterFont,
  FORMULA_FONT_PREFERENCES_CHANGED_EVENT,
  formulaChineseFontFamily,
  formulaLetterFontFamilies,
  readPersistedFormulaFontPreferences,
} from "./formulaFontPreferences";
import {
  FORMULA_SPACING_CHANGED_EVENT,
  activeFormulaSpacing,
  readPersistedFormulaSpacing,
  type FormulaSpacingPreferences,
} from "./formulaSpacingPreferences";

// Formulas are laid out and drawn exactly as LaTeX does (amsmath, 10pt), by
// the shared VisualTeX layout engine, in one of its font profiles: Computer
// Modern, or newtxtext + newtxmath for "Times New Roman" (every symbol,
// spacing and the caret as LaTeX with those packages has them). The editing
// core decides where the caret is shown and answers every geometry question
// from the engine's layout (docs/editing-core/DESIGN.md).
// Characters the TeX fonts lack (Chinese in \text{...}) are set in the formula
// Chinese font, with metrics measured from that font as XeTeX would. The
// other letter fonts set letters, digits and roman text the same way, over
// Computer Modern; symbols keep the TeX fonts.
let installed = false;
let chineseFamily: string | null = null;
let letterFamilies: { upright: string; italic: string } | null | undefined;
let profileChoice: "cm" | "ntx" | undefined;
let spacingPreferences: FormulaSpacingPreferences | null = null;
let fontRevision = 0;
const fontRevisionListeners = new Set<() => void>();

/** Changes whenever the formula fonts or spacing change (previews cache their markup by it). */
export function getFormulaFontRevision(): number {
  return fontRevision;
}

export function subscribeFormulaFontRevision(listener: () => void): () => void {
  fontRevisionListeners.add(listener);
  return () => fontRevisionListeners.delete(listener);
}

/** The editing core: the editor's caret, movement and geometry (docs/editing-core/DESIGN.md). */
export const visualTexEditingCore = createEditingCore();

type CoreMathfield = Parameters<typeof visualTexEditingCore.hitTest>[0];

/** The editing model of a field, for the core's decisions (null before the field mounts). */
export function editingModelOf(field: HTMLElement): CoreMathfield["model"] | null {
  return (field as unknown as { _mathfield?: CoreMathfield })._mathfield?.model ?? null;
}

/** Put a field's caret at the stop a page point means (clicks the app receives for a row). */
export function placeCaretAtPoint(field: HTMLElement, clientX: number, clientY: number): boolean {
  const mathfield = (field as unknown as { _mathfield?: CoreMathfield })._mathfield;
  if (!mathfield) return false;
  const stop = visualTexEditingCore.hitTest(mathfield, clientX, clientY);
  return stop !== null && visualTexEditingCore.goToStop(mathfield.model, stop);
}

type Profile = "cm" | "ntx";

// An image of a formula saved with its own fonts (Office conversion) is laid
// out in them: set only while that layout runs (withFormulaLayoutFonts).
let fontOverride: { letter: FormulaLetterFont; chinese: FormulaChineseFont } | null = null;

function currentLetterFont() {
  if (fontOverride) return fontOverride.letter;
  return readPersistedFormulaFontPreferences().formulaLetterFont ?? DEFAULT_FORMULA_LETTER_FONT;
}

// The newtx profile's data is loaded when first needed (it would slow the
// start for everyone else); until it is there, formulas use Computer Modern.
let ntxData: typeof NtxData | null = null;
let ntxLoading: Promise<void> | null = null;

function loadNtxProfile(): Promise<void> {
  ntxLoading ??= import("./ntxProfileData")
    .then((m) => {
      ntxData = m.default;
      profileChoice = undefined;
      relayoutFormulas();
    })
    .catch((error) => {
      ntxLoading = null;
      console.error("VisualTeX: Times New Roman (newtx) profile unavailable", error);
    });
  return ntxLoading;
}

/** Resolves once the chosen font's (or `letter`'s) profile is loaded (images are drawn only in the real one). */
export async function formulaLayoutReady(letter?: FormulaLetterFont): Promise<"cm" | "ntx"> {
  const font = letter ?? currentLetterFont();
  if (font === "times" && !ntxData) await loadNtxProfile();
  return font === "times" && ntxData ? "ntx" : letter ? "cm" : currentProfile();
}

/**
 * Run `layout` (synchronous) in the given fonts and that letter font's active
 * spacing configuration instead of the chosen ones; the profile must be
 * loaded (formulaLayoutReady). Open formulas are not laid out again.
 */
export function withFormulaLayoutFonts<T>(
  fonts: { letter: FormulaLetterFont; chinese: FormulaChineseFont },
  layout: () => T,
): T {
  const previous = fontOverride;
  fontOverride = fonts;
  try {
    return layout();
  } finally {
    fontOverride = previous;
  }
}

/** The engine profile of the chosen letter font: "Times New Roman" is LaTeX's newtxtext + newtxmath. */
function currentProfile(): Profile {
  if (fontOverride) return fontOverride.letter === "times" && ntxData ? "ntx" : "cm";
  if (profileChoice === undefined) {
    const wantsNtx = currentLetterFont() === "times";
    if (wantsNtx && !ntxData) void loadNtxProfile();
    profileChoice = wantsNtx && ntxData ? "ntx" : "cm";
  }
  return profileChoice;
}

function currentLetterFamilies(): { upright: string; italic: string } | null {
  if (fontOverride) return null; // both letter fonts are TeX profiles
  if (letterFamilies === undefined) {
    const font = currentLetterFont();
    // Times New Roman is the newtx profile once loaded; until then its letters over Computer Modern
    letterFamilies = font === "katex" || (font === "times" && ntxData) ? null : formulaLetterFontFamilies(font);
  }
  return letterFamilies;
}

/** The spacing adjustments of the chosen font's active configuration (Settings), null for LaTeX's. */
function currentSpacing() {
  if (spacingPreferences === null) spacingPreferences = readPersistedFormulaSpacing();
  return activeFormulaSpacing(spacingPreferences, currentLetterFont() === "times" ? "times" : "katex");
}

function currentChineseFamily(): string {
  if (fontOverride) return formulaChineseFontFamily(fontOverride.chinese);
  if (chineseFamily === null) {
    chineseFamily = formulaChineseFontFamily(
      readPersistedFormulaFontPreferences().formulaChineseFont ?? DEFAULT_FORMULA_CHINESE_FONT,
    );
  }
  return chineseFamily;
}

const MEASURE_PX = 200;
let measureContext: CanvasRenderingContext2D | null | undefined;

function measureNative(family: string, text: string, style = "") {
  if (measureContext === undefined) {
    measureContext = typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
  }
  if (!measureContext) return null;
  measureContext.font = `${style ? style + " " : ""}${MEASURE_PX}px ${family}`;
  const m = measureContext.measureText(text);
  return {
    width: m.width / MEASURE_PX,
    ascent: Math.max(0, m.actualBoundingBoxAscent) / MEASURE_PX,
    descent: Math.max(0, m.actualBoundingBoxDescent) / MEASURE_PX,
    italic: Math.max(0, m.actualBoundingBoxRight - m.width) / MEASURE_PX,
  };
}

// Rare integrals (\sumint, \intbar, ...) that no LaTeX package in the
// reference set defines: drawn from the STIX Two Math outlines the app ships.
function rareIntegralGlyphs(): PathGlyphSet | undefined {
  try {
    const bytes = Uint8Array.from(atob(RARE_INTEGRAL_GLYPHS_GZIP_BASE64), (c) => c.charCodeAt(0));
    const payload = JSON.parse(strFromU8(gunzipSync(bytes))) as Omit<PathGlyphSet, "name">;
    return { name: "stix-two-math-integrals", unitsPerEm: payload.unitsPerEm, glyphs: payload.glyphs };
  } catch (error) {
    console.error("VisualTeX: rare integral outlines unavailable", error);
    return undefined;
  }
}

function profileData(p: Profile): LayoutData {
  if (p === "cm") return { tfm, meanings, templates, onearg, textFonts, fonts, profile } as unknown as LayoutData;
  // the fonts newtx shares with Computer Modern (rsfs, ...) come from the main data
  const d = ntxData!;
  return {
    tfm: { ...tfm, ...d.tfm },
    meanings: d.meanings,
    templates: d.templates,
    onearg: d.onearg,
    textFonts: d.textFonts,
    fonts: [...fonts, ...d.fonts],
    profile: d.profile,
  } as unknown as LayoutData;
}

type Engine = ReturnType<typeof createVisualTexLayout>;

/**
 * The layout engine with the app's fonts; `measure` sizes the characters the
 * TeX fonts lack. It lays out in the profile of the chosen letter font: an
 * engine per profile, made when first needed, behind one object.
 */
export function createAppLayoutEngine(measure: NativeMeasure = measureNative): Engine {
  const engines = new Map<Profile, Engine>();
  let paths: PathGlyphSet | undefined | null = null;
  const engineFor = (p: Profile): Engine => {
    let e = engines.get(p);
    if (!e) {
      if (paths === null) paths = rareIntegralGlyphs();
      e = createVisualTexLayout(profileData(p), {
        measureNative: measure,
        fallbackFamily: currentChineseFamily,
        pathGlyphs: paths,
        caretState: visualTexEditingCore.caretState,
        letterFamilies: currentLetterFamilies,
        spacing: currentSpacing,
      });
      engines.set(p, e);
    }
    return e;
  };
  return new Proxy({} as Engine, {
    get(_target, key) {
      const e = engineFor(currentProfile()) as unknown as Record<PropertyKey, unknown>;
      const v = e[key];
      return typeof v === "function" ? (v as (...args: unknown[]) => unknown).bind(e) : v;
    },
  });
}

/** Lay out every open formula and preview again (the formula fonts changed). */
function relayoutFormulas(): void {
  chineseFamily = null;
  letterFamilies = undefined;
  profileChoice = undefined;
  spacingPreferences = null;
  fontRevision++;
  for (const listener of fontRevisionListeners) listener();
  if (typeof document === "undefined") return;
  document.querySelectorAll<HTMLElement>("math-field").forEach((field) => {
    (field as unknown as { visualTexRender?: () => void }).visualTexRender?.();
  });
}

export function installVisualTexLayout(): void {
  if (installed) return;
  installed = true;
  setVisualTexEditingCore(visualTexEditingCore);
  setVisualTexLayoutEngine(createAppLayoutEngine());
  if (typeof window !== "undefined") {
    // open formulas are laid out again in the new fonts
    window.addEventListener(FORMULA_FONT_PREFERENCES_CHANGED_EVENT, relayoutFormulas);
    window.addEventListener(FORMULA_SPACING_CHANGED_EVENT, relayoutFormulas);
  }
}

installVisualTexLayout();
