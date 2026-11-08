import { setVisualTexEditingCore, setVisualTexLayoutEngine } from "mathlive";
import { createVisualTexLayout, type LayoutData, type PiecePlacer } from "../../shared/tex-layout/src/index.ts";
import { createEditingCore } from "../../shared/visualtex-core/src/index.ts";
import tfm from "../../shared/tex-layout/data/tfm.json";
import meanings from "../../shared/tex-layout/data/latex-meanings.json";
import templates from "../../shared/tex-layout/data/latex-templates.json";
import onearg from "../../shared/tex-layout/data/latex-onearg.json";
import textFonts from "../../shared/tex-layout/data/text-fonts.json";
import fonts from "../../shared/tex-layout/data/fonts.json";
import "../../shared/tex-layout/data/fonts/fonts.css";
import profile from "../../shared/tex-layout/data/profile.json";
import type NtxData from "./ntxProfileData";
import type { PathGlyphSet } from "../../shared/tex-layout/src/path-font.ts";
import type { NativeMeasure } from "../../shared/tex-layout/src/native-font.ts";
import { gunzipSync, strFromU8 } from "fflate";
import { RARE_INTEGRAL_GLYPHS_GZIP_BASE64 } from "../math/rareIntegralGlyphs.generatedData";
import { customSymbolCssClass, customSymbolSvgMarkerClass, getActiveCustomSymbols } from "../math/customSymbolRegistry";

/** A custom operator's italic correction, by the class of its box (editor and export classes). */
function customSymbolOperatorItalic(className: string) {
  if (!className.startsWith("visualtex-custom-symbol-")) return null;
  const symbol = getActiveCustomSymbols().find(
    (item) => customSymbolCssClass(item) === className || customSymbolSvgMarkerClass(item) === className,
  );
  return symbol?.role === "operator" ? symbol.italicCorrectionEm ?? null : null;
}
import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  type FormulaChineseFont,
  type FormulaLetterFont,
  FORMULA_FONT_PREFERENCES_CHANGED_EVENT,
  FORMULA_FACES_STORAGE,
  formulaChineseFontFamily,
  formulaFaceFamilies,
  formulaLetterFontFamilies,
  readPersistedFormulaFaces,
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
// Modern, or newtxtext + newtxmath ("newtxmath", stored as "times"; every symbol,
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
type FaceFamilies = Record<"letter" | "greek" | "digit", { upright: string; italic: string } | null>;
let faceFamilies: FaceFamilies | undefined;
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

// VisualTeX Anywhere draws the caret as tall as the host field's line (setFormulaCaretBand);
// everywhere else formulas have the profile's caret.
let caretBand: { up: number; down: number } | null = null;

/**
 * The caret of this page's formulas: from the baseline `up` and `down`, in em of the formula's
 * font at text size (null: the profile's caret). Open formulas are drawn again.
 */
export function setFormulaCaretBand(band: { up: number; down: number } | null): void {
  const same =
    band === caretBand ||
    (band !== null && caretBand !== null && Math.abs(band.up - caretBand.up) < 1e-3 && Math.abs(band.down - caretBand.down) < 1e-3);
  if (same) return;
  caretBand = band;
  if (typeof document === "undefined") return;
  document.querySelectorAll<HTMLElement>("math-field").forEach((field) => {
    (field as unknown as { visualTexRender?: () => void }).visualTexRender?.();
  });
}

// VisualTeX Anywhere sets its formula on several lines, where the host field wraps its text
// (setFormulaLineLayout); everywhere else formulas stay on one line.
let lineLayout: ((mathfield: unknown) => PiecePlacer | null) | null = null;

/**
 * How this page's edited formulas go on several lines (null: one line); open formulas are drawn
 * again. The placer gets the widths of the pieces TeX may break the formula into (em) and gives
 * each piece's place (shared/tex-layout lines.ts).
 */
export function setFormulaLineLayout(layout: ((mathfield: unknown) => PiecePlacer | null) | null): void {
  lineLayout = layout;
  redrawOpenFormulas();
}

/** Draws every open formula again (its line layout changed). */
export function redrawOpenFormulas(): void {
  if (typeof document === "undefined") return;
  document.querySelectorAll<HTMLElement>("math-field").forEach((field) => {
    (field as unknown as { visualTexRender?: () => void }).visualTexRender?.();
  });
}

type CoreMathfield = Parameters<typeof visualTexEditingCore.hitTest>[0];

/** The editing model of a field, for the core's decisions (null before the field mounts). */
export function editingModelOf(field: HTMLElement): CoreMathfield["model"] | null {
  return (field as unknown as { _mathfield?: CoreMathfield })._mathfield?.model ?? null;
}

/**
 * Put a field's caret at the stop a page point means (clicks the app receives
 * for a row) and draw it there: a field that already has the focus is not
 * drawn again by focusing it.
 */
export function placeCaretAtPoint(field: HTMLElement, clientX: number, clientY: number): boolean {
  const mathfield = (field as unknown as { _mathfield?: CoreMathfield })._mathfield;
  if (!mathfield) return false;
  const stop = visualTexEditingCore.hitTest(mathfield, clientX, clientY);
  if (stop === null) return false;
  const drawn = field as unknown as { visualTexGoToStop?: (stop: number) => boolean };
  if (drawn.visualTexGoToStop) return drawn.visualTexGoToStop(stop);
  if (!visualTexEditingCore.goToStop(mathfield.model, stop)) return false;
  // goToStop moves the model only. A click beside the formula then left the
  // caret painted where it was (WebKit no longer refocuses, which used to
  // repaint): the element's position setter requests the render.
  (field as unknown as { position: number }).position = mathfield.model.position;
  return true;
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

/** Letters, Greek letters and digits in fonts of their own (Settings), over the math font. */
function currentFaceFamilies(): FaceFamilies {
  if (faceFamilies === undefined) {
    const faces = readPersistedFormulaFaces();
    faceFamilies = {
      letter: formulaFaceFamilies(faces.letter),
      greek: formulaFaceFamilies(faces.greek),
      digit: formulaFaceFamilies(faces.digit),
    };
  }
  return faceFamilies;
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
        faceFamilies: currentFaceFamilies,
        spacing: currentSpacing,
        classOperatorItalic: customSymbolOperatorItalic,
        caretBand: () => caretBand,
        lineLayout: (mathfield) => lineLayout?.(mathfield) ?? null,
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
  faceFamilies = undefined;
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
    // the faces chosen in another window (the main window's Settings, for an Office editor)
    window.addEventListener("storage", (event) => {
      if (event.key === FORMULA_FACES_STORAGE) relayoutFormulas();
    });
  }
}

installVisualTexLayout();
