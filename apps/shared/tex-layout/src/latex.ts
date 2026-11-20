// LaTeX's math setup: font families and parameter values for
// \documentclass{article} (10pt) with amsmath, amssymb and mathrsfs.
import { decimalToScaled } from "./arith.ts";
import { FIL, NORMAL, type GlueSpec } from "./nodes.ts";
import type { MathEnv } from "./mlist.ts";
import { TeXFont, type TfmData } from "./tfm.ts";

export type TfmLoader = (name: string) => TfmData;

const PT = 65536;

/** LaTeX defaults (latex.ltx / fontmath.ltx). */
export const LATEX_PARAMS = {
  nullDelimiterSpace: decimalToScaled("1.2"),
  scriptSpace: decimalToScaled("0.5"),
  delimiterFactor: 901,
  delimiterShortfall: 5 * PT,
  thinMuSkip: { width: 3 * PT, stretch: 0, stretchOrder: NORMAL, shrink: 0, shrinkOrder: NORMAL } as GlueSpec,
  medMuSkip: { width: 4 * PT, stretch: 2 * PT, stretchOrder: NORMAL, shrink: 4 * PT, shrinkOrder: NORMAL } as GlueSpec,
  thickMuSkip: { width: 5 * PT, stretch: 5 * PT, stretchOrder: NORMAL, shrink: 0, shrinkOrder: NORMAL } as GlueSpec,
  binOpPenalty: 700,
  relPenalty: 500,
};

/** \skewchar settings from LaTeX's font family declarations. */
export function skewCharFor(name: string): number {
  if (/^cmmib?\d+$/.test(name)) return 0o177;
  if (/^cmb?sy\d+$/.test(name)) return 0o60;
  return -1;
}

/** What a font profile's LaTeX setup changes in its fonts (tools/texprofile.py). */
export interface FontSetup {
  /** \skewchar of each font (by TFM name) */
  skewchar?: Record<string, number>;
  /** \fontdimen values a package assigned, by \fontname ("ntx-Regular-tlf-ot1 at 7.3pt"): number -> sp */
  params?: Record<string, Record<string, number>>;
}

export class FontCache {
  private readonly fonts = new Map<string, TeXFont>();
  private readonly load: TfmLoader;
  private readonly setup: FontSetup;
  /** setup.params by font name and size in sp */
  private readonly params = new Map<string, Record<string, number>>();

  constructor(load: TfmLoader, setup: FontSetup = {}) {
    this.load = load;
    this.setup = setup;
    for (const [spec, dims] of Object.entries(setup.params ?? {})) {
      const m = /^(\S+)(?: at ([0-9.]+)pt)?$/.exec(spec);
      if (m) this.params.set(m[1] + "@" + (m[2] ? decimalToScaled(m[2]) : new TeXFont(load(m[1])).size), dims);
    }
  }

  /** Font `name` at `size` sp (null = design size). Same name+size = same font. */
  get(name: string, size: number | null = null): TeXFont {
    const data = this.load(name);
    const probe = new TeXFont(data, size);
    const key = name + "@" + probe.size;
    let f = this.fonts.get(key);
    if (!f) {
      f = probe;
      f.skewChar = this.setup.skewchar?.[name] ?? skewCharFor(name);
      for (const [k, v] of Object.entries(this.params.get(key) ?? {})) f.params[Number(k)] = v;
      this.fonts.set(key, f);
    }
    return f;
  }

  /** Parse a \fontname result such as "cmex7 at 5.0pt". */
  byFontname(spec: string): TeXFont | null {
    if (spec === "nullfont") return null;
    const m = /^(\S+)(?: at ([0-9.]+)pt)?$/.exec(spec);
    if (!m) throw new Error("bad font name " + spec);
    return this.get(m[1], m[2] ? decimalToScaled(m[2]) : null);
  }
}

export type FamilyTable = Map<number, [TeXFont | null, TeXFont | null, TeXFont | null]>;

/** A family table from \fontname results: family -> [text, script, scriptscript]. */
export function familiesFrom(cache: FontCache, fams: Record<string, string[]>): FamilyTable {
  const t: FamilyTable = new Map();
  for (const [fam, [a, b, c]] of Object.entries(fams)) t.set(Number(fam), [cache.byFontname(a), cache.byFontname(b), cache.byFontname(c)]);
  return t;
}

/** The families LaTeX sets up at 10pt (as reported by \fontname). */
export function standardFamilies(cache: FontCache): FamilyTable {
  const t: FamilyTable = new Map();
  const set = (fam: number, a: string, b: string, c: string) =>
    t.set(fam, [cache.byFontname(a), cache.byFontname(b), cache.byFontname(c)]);
  set(0, "cmr10", "cmr7", "cmr5");
  set(1, "cmmi10", "cmmi7", "cmmi5");
  set(2, "cmsy10", "cmsy7", "cmsy5");
  set(3, "cmex10", "cmex7", "cmex7 at 5.0pt");
  set(4, "msam10", "msam7", "msam5");
  set(5, "msbm10", "msbm7", "msbm5");
  set(6, "rsfs10", "rsfs7", "rsfs5");
  return t;
}

export function makeEnv(fams: FamilyTable, params = LATEX_PARAMS): MathEnv {
  return {
    famFnt(fam: number, size: number) {
      const row = fams.get(fam);
      if (!row) return null;
      return row[size / 16] ?? null;
    },
    ...params,
  };
}

export { FIL };
