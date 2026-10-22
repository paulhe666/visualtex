import { convertLatexToSvg } from "mathlive";
import type { GlyphTable, SvgFormula } from "../../../shared/tex-layout/src/index.ts";
import { expandVisualTexLatexForLayout } from "../editor/mathLiveIntegralCompatibility.ts";
import { formulaLayoutReady, withFormulaLayoutFonts } from "../editor/visualTexLayoutRuntime.ts";
import type { FormulaChineseFont, FormulaLetterFont } from "../editor/formulaFontPreferences.ts";
import { customSymbolArtworkForClass } from "../math/customSymbolRendering.ts";

// Exported, copied and Office images drawn by the editor's own layout engine:
// the formula is parsed and laid out exactly as the editor shows it (font
// profile, spacing configuration, custom symbols) and drawn with the outlines
// of the same fonts. The outlines of each font are loaded when an image first
// needs them.

type GlyphModule = () => Promise<unknown>;

const MAIN_GLYPHS = import.meta.glob("../../../shared/tex-layout/data/glyphs/*.json", { import: "default" }) as Record<string, GlyphModule>;
const NTX_GLYPHS = import.meta.glob("../../../shared/tex-layout/data/profiles/ntx/glyphs/*.json", { import: "default" }) as Record<string, GlyphModule>;

function byFont(modules: Record<string, GlyphModule>) {
  return new Map(Object.entries(modules).map(([path, load]) => [path.slice(path.lastIndexOf("/") + 1, -".json".length), load]));
}

const SOURCES = { cm: [byFont(MAIN_GLYPHS)], ntx: [byFont(NTX_GLYPHS), byFont(MAIN_GLYPHS)] };
const loaded = new Map<string, GlyphTable | null>();

function glyphKey(profile: "cm" | "ntx", tfm: string) {
  return `${profile}:${tfm}`;
}

async function loadGlyphs(profile: "cm" | "ntx", fonts: readonly string[]) {
  await Promise.all(
    fonts.map(async (tfm) => {
      const key = glyphKey(profile, tfm);
      if (loaded.has(key)) return;
      const load = SOURCES[profile].map((m) => m.get(tfm)).find(Boolean);
      loaded.set(key, load ? ((await load()) as GlyphTable) : null);
    }),
  );
}

export interface EngineSvgOptions {
  displayMode: boolean;
  /** size of the formula's text (LaTeX's 10pt is scaled to it) */
  fontSizePx: number;
  color: string;
  idPrefix?: string;
  /** the formula's own fonts (a saved formula's); absent: the chosen ones */
  fonts?: { letter: FormulaLetterFont; chinese: FormulaChineseFont };
}

/**
 * Lay out `source` (one formula, or rows already joined into one environment)
 * and draw it. Throws for commands the editor cannot typeset, as exports did
 * at the MathJax parser.
 */
export async function renderEngineSvg(source: string, options: EngineSvgOptions): Promise<SvgFormula> {
  const fonts = options.fonts;
  const profile = await formulaLayoutReady(fonts?.letter);
  const expanded = expandVisualTexLatexForLayout(source);
  const draw = () =>
    convertLatexToSvg(
      expanded,
      { defaultMode: options.displayMode ? "math" : "inline-math" },
      {
        pxPerEm: options.fontSizePx,
        color: options.color,
        idPrefix: options.idPrefix,
        glyphs: (tfm: string) => loaded.get(glyphKey(profile, tfm)) ?? undefined,
        classArtwork: customSymbolArtworkForClass,
      },
    ) as SvgFormula & { unknown: string[] };
  const render = () => (fonts ? withFormulaLayoutFonts(fonts, draw) : draw());
  let result = render();
  const unknown = [...new Set(result.unknown.filter(Boolean))];
  if (unknown.length) throw new Error(`Unsupported LaTeX: ${unknown.slice(0, 5).join(", ")}`);
  if (result.missing.length) {
    await loadGlyphs(profile, result.missing);
    result = render();
    const absent = result.missing.filter((tfm) => loaded.get(glyphKey(profile, tfm)) === null);
    if (absent.length) throw new Error(`VisualTeX has no outlines for the fonts ${absent.join(", ")}.`);
  }
  return result;
}
