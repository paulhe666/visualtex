// Kernel atoms -> TeX noads with LaTeX semantics.
//
// Every construct is translated the way LaTeX (with amsmath, amssymb,
// mathtools, mathrsfs, bm, xcolor) would build it: math codes and
// \mathchardefs, TeX's group simplification (tex.web 1186), script attachment
// (tex.web 1176), and the macro definitions of \frac, \genfrac, \sqrt, \root,
// \text, \operatorname, \big, ... Constructs that LaTeX builds as boxes are
// built as boxes by running the TeX engine on the inner formulas.
//
// The atom objects themselves are used as node tags, so the editor can map
// every glyph and region back to the atom that produced it.
import { MAX_DIMEN, decimalToScaled, div, nxPlusY, remainder, scanDimen, xnOverD, xOverN } from "./arith.ts";
import { FontCache, LATEX_PARAMS, familiesFrom, makeEnv, type FamilyTable } from "./latex.ts";
import { MlistJson } from "./mlist-json.ts";
import {
  ACCENT_NOAD,
  BIN_NOAD,
  CHOICE_NODE,
  CLOSE_NOAD,
  DEFAULT_CODE,
  DISPLAY_STYLE,
  EMPTY,
  FRACTION_NOAD,
  INNER_NOAD,
  LEFT_NOAD,
  LIMITS,
  MATH_CHAR,
  MIDDLE_NOAD,
  MathTypesetter,
  NO_LIMITS,
  NULL_DELIMITER,
  OP_NOAD,
  OPEN_NOAD,
  ORD_NOAD,
  OVER_NOAD,
  PUNCT_NOAD,
  RADICAL_NOAD,
  REL_NOAD,
  RIGHT_NOAD,
  SCRIPT_SCRIPT_STYLE,
  SCRIPT_STYLE,
  STYLE_NODE,
  SUB_BOX,
  SUB_MLIST,
  TEXT_STYLE,
  UNDER_NOAD,
  VCENTER_NOAD,
  boxField,
  charField,
  emptyField,
  mlistField,
  newNoad,
  type ChoiceItem,
  type Delimiter,
  type Field,
  type MathEnv,
  type MItem,
  type Noad,
} from "./mlist.ts";
import {
  CHAR_NODE,
  GLUE_NODE,
  HLIST_NODE,
  KERN_NODE,
  MATH_NODE,
  MU_GLUE,
  NORMAL,
  RULE_NODE,
  VLIST_NODE,
  newCharacter,
  newKern,
  newNullBox,
  newRule,
  type BoxNode,
  type GlueSpec,
  type Node,
  type Tag,
} from "./nodes.ts";
import { EXACTLY, hpack, vpack, vpackage } from "./pack.ts";
import { VList, glue, halign, type AlignNoalign, type AlignRow, type VParams } from "./align.ts";
import { typesetText, type TextToken } from "./text.ts";
import type { TeXFont } from "./tfm.ts";
import { scalePermille, spacingKey, spacingParamScales, spacingPercent, type Spacing } from "./spacing.ts";

/** The parts of a kernel atom the adapter reads (duck-typed). */
export interface AtomLike {
  type: string;
  command?: string;
  value?: string;
  mode?: string;
  style?: Record<string, any>;
  branch(name: string): AtomLike[] | undefined;
  [key: string]: any;
}

export interface LatexData {
  meanings: {
    commands: Record<string, string>;
    robust: Record<string, string>;
    mathcode: Record<string, number>;
    delcode: Record<string, number>;
  };
  templates: Record<string, { mathlist: any[]; fontids: Record<string, string> }>;
  onearg: Record<string, { mathlist: any[]; fontids: Record<string, string> }>;
  textFonts: Record<string, string>;
  /** the font profile's LaTeX setup (tools/texprofile.py); Computer Modern when absent */
  profile?: ProfileData;
}

/** A font profile's math setup as LaTeX has it after its preamble (tools/texprofile.py). */
export interface ProfileData {
  /** family -> [text, script, scriptscript] font, as \fontname reports them */
  fams: Record<string, string[]>;
  /** the family of each math alphabet and symbol font (FamilyRoles) */
  roles: Record<string, number>;
  skewchar: Record<string, number>;
  /** \fontdimen values a package assigned (FontSetup.params) */
  params: Record<string, Record<string, number>>;
  /** NFSS defaults: encoding, \rmdefault, \sfdefault, \ttdefault, and the text, script, scriptscript sizes */
  text: { encoding: string; rm: string; sf: string; tt: string; sizes: string[] };
  /** the family \bm uses for each family that has a bold version */
  bold?: Record<string, number>;
  /** the families of \boldmath that differ: bm sets a character with no bold family in \boldmath, in a box */
  boldmath?: Record<string, string[]>;
}

/** What each math alphabet and symbol font sets \fam to. */
export interface FamilyRoles {
  rm: number;
  mi: number;
  sy: number;
  ex: number;
  bf: number;
  /** bm's bold math italic, symbols and (if the profile has one) large symbols */
  bm: number;
  bmsy: number;
  bmex?: number;
  sf: number;
  tt: number;
  it: number;
  cal: number;
  bb: number;
  scr: number;
  frak: number;
  esint?: number;
}

/** The closing delimiter of each opening one (the kernel's RIGHT_DELIM), and back. */
const RIGHT_FENCE: Record<string, string> = {
  "(": ")",
  "{": "}",
  "[": "]",
  "|": "|",
  "\\lbrace": "\\rbrace",
  "\\lparen": "\\rparen",
  "\\{": "\\}",
  "\\langle": "\\rangle",
  "\\lfloor": "\\rfloor",
  "\\lceil": "\\rceil",
  "\\vert": "\\vert",
  "\\lvert": "\\rvert",
  "\\Vert": "\\Vert",
  "\\lVert": "\\rVert",
  "\\lbrack": "\\rbrack",
  "\\ulcorner": "\\urcorner",
  "\\llcorner": "\\lrcorner",
  "\\lgroup": "\\rgroup",
  "\\lmoustache": "\\rmoustache",
};
const LEFT_FENCE: Record<string, string> = Object.fromEntries(Object.entries(RIGHT_FENCE).map(([l, r]) => [r, l]));

// Families as LaTeX allocates them with the Computer Modern preamble (see
// data/notes); a profile's own come from its data (ProfileData.roles).
export const FAM = {
  rm: 0,
  mi: 1,
  sy: 2,
  ex: 3,
  msa: 4,
  msb: 5,
  rsfs: 6,
  bf: 7,
  bm: 8,
  bmsy: 9,
  sf: 10,
  tt: 11,
  it: 12,
  frak: 13,
  /** esint.sty's largesymbolsA (only for the integrals amsmath does not define) */
  esint: 14,
  /** outline-drawn operators (rare integrals from STIX Two Math), see PathFont; not a TeX family */
  paths: 20,
  /**
   * The editor's letter font, when one other than Computer Modern is chosen
   * (a system font, measured): italic, upright, bold and bold italic letters
   * and digits. Not TeX families: no math code names them (see letterField).
   */
  letterIt: 21,
  letterUp: 22,
  letterBf: 23,
  letterBfIt: 24,
} as const;

const CM_ROLES: FamilyRoles = {
  rm: FAM.rm,
  mi: FAM.mi,
  sy: FAM.sy,
  ex: FAM.ex,
  bf: FAM.bf,
  bm: FAM.bm,
  bmsy: FAM.bmsy,
  sf: FAM.sf,
  tt: FAM.tt,
  it: FAM.it,
  cal: FAM.sy,
  bb: FAM.msb,
  scr: FAM.rsfs,
  frak: FAM.frak,
  esint: FAM.esint,
};
const CM_TEXT: ProfileData["text"] = { encoding: "OT1", rm: "cmr", sf: "cmss", tt: "cmtt", sizes: ["10", "7", "5"] };
/** Greek in the Computer Modern fonts: capitals at codes 0-10, lower case at 11-39 of the math italic fonts. */
const UPPER_GREEK = "ΓΔΘΛΞΠΣΥΦΨΩ";
const LOWER_GREEK = "αβγδϵζηθικλμνξπρστυϕχψωεϑϖϱςφ";

export function latexFamilies(cache: FontCache, profile?: ProfileData): FamilyTable {
  if (profile) return familiesFrom(cache, profile.fams);
  const t: FamilyTable = new Map();
  const set = (fam: number, a: string, b: string, c: string) =>
    t.set(fam, [cache.byFontname(a), cache.byFontname(b), cache.byFontname(c)]);
  set(FAM.rm, "cmr10", "cmr7", "cmr5");
  set(FAM.mi, "cmmi10", "cmmi7", "cmmi5");
  set(FAM.sy, "cmsy10", "cmsy7", "cmsy5");
  set(FAM.ex, "cmex10", "cmex7", "cmex7 at 5.0pt");
  set(FAM.msa, "msam10", "msam7", "msam5");
  set(FAM.msb, "msbm10", "msbm7", "msbm5");
  set(FAM.rsfs, "rsfs10", "rsfs7", "rsfs5");
  set(FAM.bf, "cmbx10", "cmbx7", "cmbx5");
  set(FAM.bm, "cmmib10", "cmmib7", "cmmib5");
  set(FAM.bmsy, "cmbsy10", "cmbsy7", "cmbsy5");
  set(FAM.sf, "cmss10", "cmss8 at 7.0pt", "cmss8 at 5.0pt");
  set(FAM.tt, "cmtt10", "cmtt8 at 7.0pt", "cmtt8 at 5.0pt");
  set(FAM.it, "cmti10", "cmti7", "cmti7 at 5.0pt");
  set(FAM.frak, "eufm10", "eufm7", "eufm5");
  set(FAM.esint, "esint10", "esint10 at 7.0pt", "esint10 at 5.0pt");
  return t;
}

/** Math alphabet command -> its role (FamilyRoles), the family set by \mathgroup. */
const ALPHABET_ROLE: Record<string, keyof FamilyRoles> = {
  "\\mathrm": "rm",
  "\\mathnormal": "mi",
  "\\mathcal": "cal",
  "\\mathbb": "bb",
  "\\mathscr": "scr",
  "\\mathbf": "bf",
  "\\mathsf": "sf",
  "\\mathtt": "tt",
  "\\mathit": "it",
  "\\mathfrak": "frak",
};

/** Context while converting a list. */
interface Ctx {
  /** \fam inside a math alphabet, -1 outside */
  fam: number;
  /** inside \bm{...}: map families to their bold versions */
  bold: boolean;
  /** inside \operatorname: \newmcodes@ math codes */
  operator: boolean;
  /** inside real display math (LaTeX's \if@display) */
  display: boolean;
}

const BASE_CTX: Ctx = { fam: -1, bold: false, operator: false, display: false };

/** amsmath arrow fills: [left piece, repeated piece, right piece] */
type ArrowFill = [string, string, string];
const FILL = {
  right: ["\\relbar", "\\relbar", "\\rightarrow"] as ArrowFill,
  left: ["\\leftarrow", "\\relbar", "\\relbar"] as ArrowFill,
  leftright: ["\\leftarrow", "\\relbar", "\\rightarrow"] as ArrowFill,
  Right: ["\\Relbar", "\\Relbar", "\\Rightarrow"] as ArrowFill,
  Left: ["\\Leftarrow", "\\Relbar", "\\Relbar"] as ArrowFill,
  LeftRight: ["\\Leftarrow", "\\Relbar", "\\Rightarrow"] as ArrowFill,
};
const OVER_ARROWS: Record<string, { fill: ArrowFill; under: boolean }> = {
  "\\overrightarrow": { fill: FILL.right, under: false },
  "\\overleftarrow": { fill: FILL.left, under: false },
  "\\overleftrightarrow": { fill: FILL.leftright, under: false },
  "\\underrightarrow": { fill: FILL.right, under: true },
  "\\underleftarrow": { fill: FILL.left, under: true },
  "\\underleftrightarrow": { fill: FILL.leftright, under: true },
};
const EXT_ARROWS: Record<string, { k: [number, number, number, number]; fill: ArrowFill }> = {
  "\\xrightarrow": { k: [0, 3, 5, 9], fill: FILL.right },
  "\\xleftarrow": { k: [3, 0, 9, 5], fill: FILL.left },
  // mathtools
  "\\xleftrightarrow": { k: [3, 3, 7, 7], fill: FILL.leftright },
  "\\xRightarrow": { k: [0, 3, 5, 9], fill: FILL.Right },
  "\\xLeftarrow": { k: [3, 0, 9, 5], fill: FILL.Left },
  "\\xLeftrightarrow": { k: [3, 3, 7, 7], fill: FILL.LeftRight },
};

const PT = 65536;

/**
 * The tag of a wrapper shared by a run of atoms (a \text run, a colour or font
 * group, a \colorbox): an object of its own, so the region of no single atom
 * spans the run (the editor puts the caret after an atom at its region's edge).
 */
export interface RunTag {
  run: AtomLike[];
}
const runTag = (run: AtomLike[]): RunTag => ({ run });

/** TeX's scanning of "<factor><dimen>" such as .6\dimen@ (tex.web 455). */
export function texScale(factor: string, v: number): number {
  const [ip, fp = ""] = factor.split(".");
  const intPart = ip === "" ? 0 : parseInt(ip, 10);
  const f = decimalToScaled("0." + fp);
  let negative = false;
  if (v < 0) {
    negative = true;
    v = -v;
  }
  const r = nxPlusY(intPart, v, xnOverD(v, f, 0o200000));
  return negative ? -r : r;
}

export interface NativeFonts {
  /** a font for characters missing from the TeX fonts (e.g. CJK), at a size in sp */
  fallback(size: number): TeXFont | null;
}

export class LatexAdapter {
  readonly data: LatexData;
  readonly cache: FontCache;
  readonly fams: FamilyTable;
  readonly env: MathEnv;
  /** the families of the math alphabets and symbol fonts */
  readonly role: FamilyRoles;
  /** families of the bold math version (ProfileData.boldmath) for bm's boxes */
  private readonly boldEnv: MathEnv | null;
  /** NFSS defaults of the profile (\text) */
  readonly text: ProfileData["text"] & { bold: string };
  native: NativeFonts | null = null;
  /** operators drawn from outlines: command -> character code in FAM.paths */
  pathOps = new Map<string, number>();
  /** atoms drawn with a contour oval over their integral signs (\oiiint), filled by each layout */
  readonly ovals = new Set<Tag>();
  /**
   * The formula is being edited (set by each layout): an empty slot gets an
   * empty box where its caret stands, and an empty last row (a line just
   * started with Shift+Enter) is kept, where LaTeX drops a trailing \\.
   */
  editor = false;
  /** the user's spacing adjustments (setSpacing), per mille: \arraycolsep, \arraystretch, \jot */
  private arrayColSep = 1000;
  private arrayRowSep = 1000;
  private alignRowSep = 1000;
  private spacingApplied = "";
  private readonly json: MlistJson;

  constructor(data: LatexData, cache: FontCache, fams?: FamilyTable) {
    this.data = data;
    this.cache = cache;
    this.fams = fams ?? latexFamilies(cache, data.profile);
    this.role = data.profile ? (data.profile.roles as unknown as FamilyRoles) : CM_ROLES;
    const text = data.profile?.text ?? CM_TEXT;
    // \bfdefault: the bold series the profile's text fonts have (Computer Modern's bx)
    this.text = { ...text, bold: data.textFonts[`${text.encoding}/${text.rm}/bx/n/${text.sizes[0]}`] ? "bx" : "b" };
    this.env = makeEnv(this.fams, LATEX_PARAMS);
    const boldmath = data.profile?.boldmath;
    this.boldEnv = boldmath ? makeEnv(new Map([...this.fams, ...familiesFrom(cache, boldmath)]), LATEX_PARAMS) : null;
    this.json = new MlistJson(cache);
    this.json.charField = (fam, code) => this.letterField(fam, code) ?? charField(fam, code);
  }

  /**
   * The user's spacing adjustments (spacing.ts), null for LaTeX's: they scale
   * the math font parameters and LaTeX's spacing values of every formula laid
   * out after this.
   */
  setSpacing(spacing: Spacing | null): void {
    const key = spacingKey(spacing);
    if (key === this.spacingApplied) return;
    this.spacingApplied = key;
    const { sy, ex } = spacingParamScales(spacing);
    const pm = (k: Parameters<typeof spacingPercent>[1]) => spacingPercent(spacing, k) * 10;
    const muskip = (g: typeof LATEX_PARAMS.thinMuSkip, m: number) => ({ ...g, width: scalePermille(g.width, m), stretch: scalePermille(g.stretch, m), shrink: scalePermille(g.shrink, m) });
    for (const env of [this.env, this.boldEnv]) {
      if (!env) continue;
      env.syScale = Object.keys(sy).length ? sy : undefined;
      env.exScale = Object.keys(ex).length ? ex : undefined;
      env.radicalGap = pm("radicalGap") === 1000 ? undefined : pm("radicalGap");
      env.scriptSpace = scalePermille(LATEX_PARAMS.scriptSpace, pm("scriptSpace"));
      env.delimiterFactor = Math.round((LATEX_PARAMS.delimiterFactor * pm("delimiterSize")) / 1000);
      env.thinMuSkip = muskip(LATEX_PARAMS.thinMuSkip, pm("thinSpace"));
      env.medMuSkip = muskip(LATEX_PARAMS.medMuSkip, pm("medSpace"));
      env.thickMuSkip = muskip(LATEX_PARAMS.thickMuSkip, pm("thickSpace"));
    }
    this.arrayColSep = pm("arrayColSep");
    this.arrayRowSep = pm("arrayRowSep");
    this.alignRowSep = pm("alignRowSep");
  }

  // ---------------------------------------------------------------- entry --
  /** A whole formula: `$\displaystyle ...$` or `$...$`. */
  formula(atoms: AtomLike[], display: boolean): MItem[] {
    this.ovals.clear();
    const out: MItem[] = [{ type: STYLE_NODE, style: display ? DISPLAY_STYLE : TEXT_STYLE }];
    out.push(...this.list(atoms, { ...BASE_CTX, display }));
    return out;
  }

  // ---------------------------------------------------------------- lists --
  /** A math list (the contents of a branch). */
  list(atoms: AtomLike[] | undefined, ctx: Ctx): MItem[] {
    const out: MItem[] = [];
    const visible = (atoms ?? []).filter((a) => a.type !== "first");
    // Background runs first (\colorbox), then style runs, as the kernel's
    // LaTeX serializer forms them.
    for (const run of this.styleRuns(visible, ctx, true)) {
      const bg = this.backgroundOf(run[0]);
      if (!bg) {
        for (const r of this.styleRuns(run, ctx, false)) this.styledRun(r, ctx, out);
        continue;
      }
      // xcolor \colorbox{c}{...}: its argument is set in text mode
      let content: BoxNode;
      if (run[0].mode === "text") content = this.textRunBox(run, 0);
      else {
        const items: MItem[] = [];
        for (const r of this.styleRuns(run, ctx, false)) this.styledRun(r, ctx, items);
        content = this.mathBox(items, TEXT_STYLE, runTag(run));
      }
      out.push(...this.colorBox(content, bg, runTag(run)));
    }
    const first = atoms?.[0];
    if (this.editor && out.length === 0 && first?.type === "first") {
      // an empty box, as TeX makes of {}, where the empty slot's caret stands
      const empty = newNullBox();
      empty.firstTag = first;
      out.push(this.boxNoad(empty, undefined));
    }
    return out;
  }

  /** A background colour the atom does not inherit from its parent (\colorbox). */
  private backgroundOf(a: AtomLike): string | null {
    const bg = a.style?.backgroundColor;
    if (!bg || bg === "none") return null;
    return a.parent?.style?.backgroundColor === bg ? null : bg;
  }

  /**
   * xcolor \color@b@x: \setbox0\hbox{\kern\fboxsep #1\kern\fboxsep}, its
   * height and depth grown by \fboxsep, then a rule of that size in the colour,
   * \hskip-\wd0, and the box.
   */
  private colorBox(content: BoxNode, color: string, tag: Tag): MItem[] {
    const fboxsep = 3 * PT;
    const box0 = hpack([newKern(fboxsep, tag), content, newKern(fboxsep, tag)], 0, undefined, tag);
    box0.height += fboxsep;
    box0.depth += fboxsep;
    const block = newRule(tag);
    block.width = box0.width;
    block.height = box0.height;
    block.depth = box0.depth;
    block.color = color;
    const back: Node = { type: GLUE_NODE, spec: { width: -box0.width, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag };
    return [block as unknown as MItem, back as unknown as MItem, this.boxNoad(box0, tag)];
  }

  /** Split atoms into runs that serialize to one style command (or one background). */
  private styleRuns(atoms: AtomLike[], ctx: Ctx | undefined, background: boolean): AtomLike[][] {
    const runs: AtomLike[][] = [];
    let cur: AtomLike[] = [];
    let key = "";
    for (const a of atoms) {
      const k = background ? (this.backgroundOf(a) ?? "") + "|" + (a.mode === "text" ? "text" : "math") : this.styleKey(a, ctx);
      if (cur.length > 0 && k !== key) {
        runs.push(cur);
        cur = [];
      }
      key = k;
      cur.push(a);
    }
    if (cur.length > 0) runs.push(cur);
    return runs;
  }

  private styleKey(a: AtomLike, ctx?: Ctx): string {
    if (a.type === "subsup") return "subsup"; // scripts attach outside style groups
    if (ctx?.operator) return "op";
    const s = a.style ?? {};
    if (a.mode === "text") return "text|" + (s.color ?? "") + "|" + (s.fontSeries ?? "") + "|" + (s.fontShape ?? "") + "|" + (s.fontFamily ?? "");
    if (a.type === "latex" || a.mode === "latex") return "latex";
    return [a.mode ?? "math", s.color ?? "", this.alphabetCommand(a) ?? "", this.boldCommand(a) ?? ""].join("|");
  }

  /** The math alphabet command the serializer emits for an atom's style. */
  private alphabetCommand(a: AtomLike): string | null {
    const s = a.style ?? {};
    let v: string | undefined = s.variant;
    if (v === undefined) return null;
    if (!["calligraphic", "fraktur", "double-struck", "script", "monospace", "sans-serif"].includes(v) && s.variantStyle)
      v += "-" + s.variantStyle;
    const map: Record<string, string> = {
      calligraphic: "\\mathcal",
      fraktur: "\\mathfrak",
      "double-struck": "\\mathbb",
      script: "\\mathscr",
      monospace: "\\mathtt",
      "sans-serif": "\\mathsf",
      "normal-up": "\\mathrm",
      "normal-italic": "\\mathnormal",
      "main-up": "\\mathrm",
      "main-italic": "\\mathit",
      "ams-up": "\\mathrm",
      "ams-italic": "\\mathit",
      "normal-bolditalic": "\\mathbfit",
      "main-bolditalic": "\\mathbfit",
      "ams-bolditalic": "\\mathbfit",
    };
    return map[v] ?? null;
  }

  private boldCommand(a: AtomLike): string | null {
    const s = a.style ?? {};
    if (a.mode !== "math" && a.mode !== undefined) return null;
    if (s.variantStyle !== "bold") return null;
    if (s.variant === "normal") return "\\mathbf";
    return "\\bm";
  }

  /** Convert a style run; wrap it in the group its style command makes. */
  private styledRun(run: AtomLike[], ctx: Ctx, out: MItem[]): void {
    const first = run[0];
    if (first.type === "subsup" || ctx.operator) {
      for (const a of run) this.atom(a, ctx, out);
      return;
    }
    if (first.mode === "text") {
      out.push(this.textRun(run));
      return;
    }
    if (first.type === "latex" || first.mode === "latex") {
      for (const a of run) out.push(this.latexInput(a));
      return;
    }
    const alphabet = this.alphabetCommand(first);
    const bold = this.boldCommand(first);
    const color = first.style?.color;
    if (!alphabet && !bold && !color) {
      for (const a of run) this.atom(a, ctx, out);
      return;
    }
    // The command wraps the run in braces (a math group); colour adds its
    // specials inside that group (they occupy no space).
    let inner: Ctx = ctx;
    if (alphabet === "\\mathbfit") inner = { ...inner, fam: this.role.bm };
    else if (alphabet) inner = { ...inner, fam: ALPHABET_ROLE[alphabet] ? this.role[ALPHABET_ROLE[alphabet]] ?? -1 : -1 };
    if (bold === "\\mathbf") inner = { ...inner, fam: this.role.bf };
    else if (bold === "\\bm") inner = { ...inner, bold: true };
    const items: MItem[] = [];
    for (const a of run) this.atom(a, inner, items);
    out.push(this.groupNoad(items, runTag(run)));
  }

  // ---------------------------------------------------------------- atoms --
  /** Convert one atom, appending to `out` (scripts attach TeX-style). */
  private atom(a: AtomLike, ctx: Ctx, out: MItem[]): void {
    if (this.commandAtom(a, ctx, out)) {
      if (a.branch("superscript") !== undefined || a.branch("subscript") !== undefined) this.attachScripts(out, a, ctx, false);
      return;
    }
    switch (a.type) {
      case "first":
        return;
      case "space":
        out.push(...this.controlSpace(a));
        break;
      case "macro":
        out.push(...this.macro(a, ctx));
        break;
      case "macro-argument":
        out.push(...this.list(a.branch("body"), ctx));
        break;
      case "subsup":
        this.attachScripts(out, a, ctx, true);
        return;
      case "group":
        out.push(this.groupNoad(this.list(a.branch("body"), ctx), a));
        break;
      case "genfrac":
        out.push(...this.genfrac(a, ctx));
        break;
      case "surd":
        out.push(...this.surd(a, ctx));
        break;
      case "accent":
        out.push(this.accent(a, ctx));
        break;
      case "leftright":
        out.push(this.leftRight(a, ctx));
        break;
      case "delim": // \middle inside \left...\right
        out.push(this.middle(a));
        break;
      case "sizeddelim":
        out.push(this.sizedDelim(a));
        break;
      case "extensible-symbol":
      case "operator":
        out.push(...this.operator(a, ctx));
        break;
      case "error":
        // a LaTeX command the kernel does not know (\idotsint ...): as LaTeX defines it
        if (a.command && this.knowsCommand(a.command)) out.push(...this.operator(a, ctx));
        else out.push(...this.fallback(a, ctx));
        break;
      case "spacing":
        out.push(...this.spacing(a, ctx));
        break;
      case "placeholder":
        out.push(this.placeholder(a, ctx));
        break;
      case "phantom":
        out.push(...this.phantom(a, ctx));
        break;
      case "box":
        out.push(...this.boxAtom(a, ctx));
        break;
      case "enclose":
        out.push(...this.enclose(a, ctx));
        break;
      case "overunder":
        out.push(...this.overUnder(a, ctx));
        break;
      case "array":
        out.push(...this.arrayEnv(a, ctx));
        break;
      case "mord":
      case "mbin":
      case "mrel":
      case "mopen":
      case "mclose":
      case "mpunct":
      case "minner":
      case "mop":
        out.push(...this.symbol(a, ctx));
        break;
      default:
        out.push(...this.fallback(a, ctx));
        break;
    }
    // Scripts carried by the atom itself (operators, groups, ...)
    if (a.branch("superscript") !== undefined || a.branch("subscript") !== undefined) this.attachScripts(out, a, ctx, false);
  }

  /** tex.web 1176: attach scripts to the last noad or a new empty one. */
  private attachScripts(out: MItem[], a: AtomLike, ctx: Ctx, fromSubsup: boolean): void {
    const sup = a.branch("superscript");
    const sub = a.branch("subscript");
    // TeX reads them in source order; the serializer emits _ then ^.
    for (const [which, atoms] of [["sub", sub], ["sup", sup]] as const) {
      if (atoms === undefined) continue;
      let last = out[out.length - 1] as Noad | undefined;
      const allowed = last && last.type >= ORD_NOAD && last.type < LEFT_NOAD;
      if (!allowed || (last as Noad)[which].mathType !== EMPTY) {
        last = newNoad(ORD_NOAD, a);
        out.push(last);
      }
      (last as Noad)[which] = this.field(atoms, ctx);
      if (fromSubsup) (last as Noad).scriptTag = a;
    }
    void fromSubsup;
  }

  /**
   * A braced field (tex.web 1186): a single plain ord noad is replaced by its
   * nucleus; otherwise the field is a sub-mlist (possibly empty).
   */
  field(atoms: AtomLike[] | undefined, ctx: Ctx): Field {
    const f = this.simplify(this.list(atoms, ctx));
    const first = atoms?.[0];
    if (first && first.type === "first") f.first = first;
    return f;
  }

  /** A branch as a sub-mlist field as it stands (a fraction's parts), knowing its branch start. */
  private listField(atoms: AtomLike[] | undefined, ctx: Ctx, list: MItem[] = this.list(atoms, ctx)): Field {
    const f = mlistField(list);
    const first = atoms?.[0];
    if (first && first.type === "first") f.first = first;
    return f;
  }

  private simplify(list: MItem[]): Field {
    if (list.length === 1) {
      const p = list[0] as Noad;
      if (p.type === ORD_NOAD && p.sub.mathType === EMPTY && p.sup.mathType === EMPTY)
        return p.nucleus.mathType === MATH_CHAR && p.nucleus.tag === undefined ? { ...p.nucleus, tag: p.tag } : { ...p.nucleus };
    }
    return mlistField(list);
  }

  /** A group `{...}` in a list: an ord noad (or the accent noad it contains). */
  private groupNoad(list: MItem[], tag: Tag): Noad {
    if (list.length === 1 && (list[0] as Noad).type === ACCENT_NOAD) return list[0] as Noad;
    const n = newNoad(ORD_NOAD, tag);
    n.nucleus = this.simplify(list);
    return n;
  }

  // ---------------------------------------------------- command dispatch --
  /**
   * Constructs identified by their command rather than their atom type
   * (the kernel leaves some of them untyped). Returns true if handled.
   */
  private commandAtom(a: AtomLike, ctx: Ctx, out: MItem[]): boolean {
    const cmd = a.command ?? "";
    const styles: Record<string, number> = {
      "\\displaystyle": DISPLAY_STYLE,
      "\\textstyle": TEXT_STYLE,
      "\\scriptstyle": SCRIPT_STYLE,
      "\\scriptscriptstyle": SCRIPT_SCRIPT_STYLE,
    };
    if (styles[cmd] !== undefined && a.branch("body") !== undefined) {
      // A style declaration: affects the rest of the current list (no group).
      out.push({ type: STYLE_NODE, style: styles[cmd], tag: a });
      out.push(...this.list(a.branch("body"), ctx));
      return true;
    }
    if ((cmd === "\\overline" || cmd === "\\underline") && a.branch("body") !== undefined) {
      const n = newNoad(cmd === "\\overline" ? OVER_NOAD : UNDER_NOAD, a);
      n.nucleus = this.field(a.branch("body"), ctx);
      out.push(n);
      return true;
    }
    const classes: Record<string, number> = {
      "\\mathord": ORD_NOAD,
      "\\mathop": OP_NOAD,
      "\\mathbin": BIN_NOAD,
      "\\mathrel": REL_NOAD,
      "\\mathopen": OPEN_NOAD,
      "\\mathclose": CLOSE_NOAD,
      "\\mathpunct": PUNCT_NOAD,
      "\\mathinner": INNER_NOAD,
    };
    if (classes[cmd] !== undefined && a.branch("body") !== undefined) {
      // \mathbin{...} etc.: a noad of that class with a braced nucleus
      const n = newNoad(classes[cmd], a);
      n.nucleus = this.field(a.branch("body"), ctx);
      if (cmd === "\\mathop" && a.explicitSubsupPlacement) {
        n.subtype = a.subsupPlacement === "over-under" ? LIMITS : a.subsupPlacement === "adjacent" ? NO_LIMITS : NORMAL;
      }
      out.push(n);
      return true;
    }
    if (cmd === "\\not" && a.branch("body") !== undefined && this.meaning("\\n@tch@r")) {
      // newtxmath: \not\foo is \notfoo or \nfoo when defined (\not= \ne, \not< \nless,
      // \not> \ngtr), else \n@tch@r\foo
      const body = (a.branch("body") ?? []).filter((x) => x.type !== "overlap" && x.type !== "first");
      const name = body.length === 1 ? body[0].command ?? body[0].value : undefined;
      const special: Record<string, string> = { "=": "\\ne", "<": "\\nless", ">": "\\ngtr" };
      const bare = name?.startsWith("\\") ? name.slice(1) : name;
      const target = name === undefined ? undefined : special[name] ?? ["\\not" + bare, "\\n" + bare].find((c) => this.meaning(c));
      if (target) {
        const mc = this.mathCodeOf(target, ctx);
        if (mc !== null) {
          out.push(this.mathChar(mc, ctx, a));
          return true;
        }
        const tpl = this.template(target, a);
        if (tpl) {
          out.push(...tpl);
          return true;
        }
      }
      out.push(this.mathChar(this.mathCodeOf("\\n@tch@r", ctx) ?? 0x3236, ctx, a));
      for (const b of body) this.atom(b, ctx, out);
      return true;
    }
    if (cmd === "\\not" && a.branch("body") !== undefined) {
      // \not = \mathchar"3236 followed by the negated symbol
      out.push(this.mathChar(this.mathCodeOf("\\not", ctx) ?? 0x3236, ctx, a));
      for (const b of (a.branch("body") ?? []).filter((x) => x.type !== "overlap")) this.atom(b, ctx, out);
      return true;
    }
    if (cmd === "\\operatorname" || cmd === "\\operatorname*") {
      out.push(this.operatorName(a, ctx, cmd === "\\operatorname*" || a.subsupPlacement === "over-under" && !!a.explicitSubsupPlacement));
      return true;
    }
    if (/^\\(?:class|htmlClass|cssId|htmlId|htmlData|style|htmlStyle|href)$/.test(cmd) && a.branch("body") !== undefined) {
      // HTML annotations (MathJax/MathLive): the content takes part in the
      // list as if it were not wrapped; the editor marks its box.
      out.push(...this.list(a.branch("body"), ctx));
      return true;
    }
    if (cmd === "\\rule") {
      out.push(this.boxNoad(this.ruleBox(a), a));
      return true;
    }
    if (cmd === "\\doubleprime" || cmd === "\\tripleprime" || cmd === "\\qprime") {
      const n = cmd === "\\doubleprime" ? 2 : cmd === "\\tripleprime" ? 3 : 4;
      for (let i = 0; i < n; i++) out.push(this.mathChar(0x0230, ctx, a));
      return true;
    }
    return false;
  }

  /** Control space "\ ": inter-word glue of the current text font. */
  private controlSpace(a: AtomLike): MItem[] {
    const f = this.textRoman();
    return [
      {
        type: GLUE_NODE,
        spec: { width: f.param(2), stretch: f.param(3), stretchOrder: NORMAL, shrink: f.param(4), shrinkOrder: NORMAL },
        subtype: NORMAL,
        tag: a,
      } as Node,
    ];
  }

  /** Kernel macro atoms: LaTeX's own definition when it has one, else the expansion. */
  private macro(a: AtomLike, ctx: Ctx): MItem[] {
    const cmd = a.command ?? "";
    if (cmd === "\\pmod" || cmd === "\\pod") return this.pmod(a, ctx, cmd === "\\pmod");
    if (cmd === "\\substack") {
      const arr = this.findAtom(a.branch("body"), (x) => x.type === "array");
      if (arr) return [this.subarray(arr, ctx, "c")];
    }
    const tpl = this.template(cmd, a);
    if (tpl) return tpl;
    return this.list(a.branch("body"), ctx);
  }

  /** amsmath \pmod / \pod: \allowbreak\mkern18mu or 8mu ( [mod\mkern6mu] #1 ) */
  private pmod(a: AtomLike, ctx: Ctx, withMod: boolean): MItem[] {
    const arg = (a.branch("body") ?? []).find((x) => x.type === "macro-argument");
    const out: MItem[] = [{ type: 12, penalty: 0, tag: a } as Node]; // \allowbreak
    out.push(this.muKern((ctx.display ? 18 : 8) * PT, a));
    out.push(this.mathChar(this.mathCodeOf("(", ctx)!, ctx, a));
    if (withMod) {
      // \operator@font mod: a group with the letters in the operator font
      const letters = [..."mod"].map((ch) => this.mathChar(this.data.meanings.mathcode[ch], { ...ctx, fam: this.role.rm }, a));
      out.push(this.groupNoad(letters, a));
      out.push(this.muKern(6 * PT, a));
    }
    out.push(...this.list(arg?.branch("body"), ctx));
    out.push(this.mathChar(this.mathCodeOf(")", ctx)!, ctx, a));
    return out;
  }

  // ------------------------------------------------------------- symbols --
  /** \mathcode / \mathchar value to a noad (tex.web 1155 set_math_char). */
  private mathChar(code: number, ctx: Ctx, tag: Tag): Noad {
    let cls = (code >> 12) & 7;
    let fam = (code >> 8) & 15;
    const ch = code & 255;
    if (cls === 7) {
      if (ctx.fam >= 0) fam = ctx.fam;
      cls = 0;
    }
    if (ctx.bold) {
      const b = this.boldFamily(fam);
      if (b === fam && this.boldEnv && this.data.profile?.boldmath?.[fam]) return this.boldmathChar(cls, fam, ch, tag);
      fam = b;
    }
    const n = newNoad(ORD_NOAD + cls, tag);
    n.nucleus = this.letterField(fam, ch) ?? charField(fam, ch);
    return n;
  }

  /**
   * bm.sty for a character whose family has no bold family: the character in
   * the bold math version, \mathchoice{\hbox{\boldmath$\displaystyle\m@th c$}}...
   * (\bm@mchoice). A choice item, not a noad: scripts start a new ord noad.
   */
  private boldmathChar(cls: number, fam: number, ch: number, tag: Tag): Noad {
    const box = (style: number): MItem[] => {
      const n = newNoad(ORD_NOAD + cls, tag);
      n.nucleus = charField(fam, ch);
      return [this.boxNoad(this.mathBox([{ type: STYLE_NODE, style, tag }, n], TEXT_STYLE, tag, this.boldEnv!), tag)];
    };
    const choice: ChoiceItem = {
      type: CHOICE_NODE,
      display: box(DISPLAY_STYLE),
      text: box(TEXT_STYLE),
      script: box(SCRIPT_STYLE),
      scriptScript: box(SCRIPT_SCRIPT_STYLE),
      tag,
    };
    return choice as unknown as Noad;
  }

  /**
   * A letter or digit in the editor's letter font, when one is chosen (each
   * layout fills the letter families): that family and the Unicode
   * character. Everything else keeps the TeX fonts.
   */
  private letterField(fam: number, ch: number): Field | null {
    const r = this.role;
    const to = fam === r.mi || fam === r.it ? FAM.letterIt : fam === r.rm ? FAM.letterUp : fam === r.bf ? FAM.letterBf : fam === r.bm ? FAM.letterBfIt : undefined;
    if (to === undefined || !this.fams.has(to)) return null;
    const mathItalic = fam === r.mi || fam === r.bm;
    let c: string | undefined;
    if ((ch >= 0x41 && ch <= 0x5a) || (ch >= 0x61 && ch <= 0x7a)) c = String.fromCharCode(ch);
    // the math italic fonts' digits are old-style figures
    else if (ch >= 0x30 && ch <= 0x39) c = mathItalic ? undefined : String.fromCharCode(ch);
    else if (ch <= 0x0a) c = UPPER_GREEK[ch];
    else if (mathItalic && ch <= 0x27) c = LOWER_GREEK[ch - 0x0b];
    return c === undefined ? null : charField(to, c.codePointAt(0)!);
  }

  /** bm: families with a bold version. */
  private boldFamily(fam: number): number {
    const bold = this.data.profile?.bold;
    if (bold) return bold[fam] ?? fam;
    const r = this.role;
    if (fam === r.mi) return r.bm;
    if (fam === r.sy) return r.bmsy;
    if (fam === r.rm) return r.bf;
    if (fam === r.ex && r.bmex !== undefined) return r.bmex;
    return fam;
  }

  private meaning(cmd: string): string | undefined {
    if (cmd === "\\|") cmd = "\\Vert"; // \| is LaTeX's \Vert
    const m = this.data.meanings.commands[cmd];
    if (m === undefined || m === "undefined") return undefined;
    return m;
  }

  /** The 27-bit \delimiter code in a meaning such as `\delimiter "426830A ` or `\delimiter 69645069 `. */
  private delimiterCodeIn(m: string): number | null {
    const r = /\\delimiter ?(?:"([0-9A-F]+)|(\d+))/.exec(m);
    if (!r) return null;
    return r[1] ? parseInt(r[1], 16) : parseInt(r[2], 10);
  }

  /** The \mathchar code of a command or character, if it is a simple one. */
  private mathCodeOf(key: string, ctx: Ctx): number | null {
    if (key.length === 1 && key.charCodeAt(0) < 128) {
      let mc = this.data.meanings.mathcode[key];
      if (ctx.operator) {
        // \newmcodes@ (amsopn): ' * . - / : behave as in text
        const nm: Record<string, number> = { "'": 39, "*": 42, ".": 0x613a, "-": 45, "/": 47, ":": 0x603a };
        if (nm[key] !== undefined) mc = nm[key];
      }
      if (mc === undefined || mc === 0x8000) return null;
      return mc;
    }
    const m = this.meaning(key);
    if (!m) return null;
    let r = /^\\mathchar"([0-9A-F]+)$/.exec(m);
    if (r) return parseInt(r[1], 16);
    if (/^(?:\\protected )?macro:->\\delimiter ?[0-9"]/.test(m) || /^\\delimiter"/.test(m)) {
      const code = this.delimiterCodeIn(m);
      if (code !== null) return code >> 12;
    }
    return null;
  }

  /** Whether LaTeX defines this command, so the engine renders it even if the kernel does not know it. */
  knowsCommand(cmd: string): boolean {
    return !!this.data.templates[cmd] || this.mathCodeOf(cmd, BASE_CTX) !== null;
  }

  private template(key: string, tag: Tag): MItem[] | null {
    const t = this.data.templates[key];
    if (!t) return null;
    return this.json.convert(t.mathlist, t.fontids, tag);
  }

  private symbol(a: AtomLike, ctx: Ctx): MItem[] {
    const keys: string[] = [];
    if (a.command) keys.push(a.command);
    if (a.value && a.value !== a.command) keys.push(a.value);
    for (const key of keys) {
      const mc = this.mathCodeOf(key, ctx);
      if (mc !== null) return [this.mathChar(mc, ctx, a)];
      const tpl = this.template(key, a);
      if (tpl) return tpl;
    }
    return this.unknownSymbol(a, ctx);
  }

  /** A symbol LaTeX does not define (e.g. a Chinese character): its character, sized like \text. */
  private unknownSymbol(a: AtomLike, _ctx: Ctx): MItem[] {
    const text = a.value ?? a.command ?? "?";
    const mk = (size: string): MItem[] => {
      const n = newNoad(ORD_NOAD, a);
      n.nucleus = boxField(this.textBox(text, this.textId(this.text.rm, "m", "n", size), a));
      return [n];
    };
    const [s0, s1, s2] = this.text.sizes;
    const choice: ChoiceItem = { type: CHOICE_NODE, display: mk(s0), text: mk(s0), script: mk(s1), scriptScript: mk(s2), tag: a };
    return [this.groupNoad([choice], a)];
  }

  // ------------------------------------------------------------- fractions --
  private genfrac(a: AtomLike, ctx: Ctx): MItem[] {
    const cmd = a.command ?? "\\frac";
    const num = this.listField(a.branch("above"), ctx);
    const den = this.listField(a.branch("below"), ctx);
    const frac = newNoad(FRACTION_NOAD, a);
    frac.num = num;
    frac.den = den;
    frac.thickness = a.hasBarLine === false ? 0 : DEFAULT_CODE;
    const ld = a.leftDelim, rd = a.rightDelim;
    if (ld || rd) {
      frac.leftDelim = this.delimiter(ld ?? ".");
      frac.rightDelim = this.delimiter(rd ?? ".");
    }
    if (cmd === "\\frac" || cmd === "\\over" || cmd === "\\atop" || cmd === "\\choose") {
      // amsmath \frac: {\begingroup#1\endgroup\@@over#2}
      return [this.groupNoad([frac], a)];
    }
    if (cmd === "\\cfrac") return this.cfrac(a, ctx);
    // \dfrac, \tfrac, \binom, \dbinom, \tbinom, \genfrac:
    // \@genfrac: {#1{\begingroup#4\endgroup#2#3\relax#5}} with #1 = style
    const style =
      a.mathstyleName === "displaystyle" || cmd === "\\dfrac" || cmd === "\\dbinom"
        ? DISPLAY_STYLE
        : a.mathstyleName === "textstyle" || cmd === "\\tfrac" || cmd === "\\tbinom"
          ? TEXT_STYLE
          : null;
    const inner = this.groupNoad([frac], a);
    const outer: MItem[] = style === null ? [inner] : [{ type: STYLE_NODE, style, tag: a }, inner];
    return [this.groupNoad(outer, a)];
  }

  /** amsmath \cfrac: {\displaystyle\frac{\strut#2}{#3}}\kern-\nulldelimiterspace */
  private cfrac(a: AtomLike, ctx: Ctx): MItem[] {
    const strut = newNullBox(a);
    // \strut: \vrule height .7\baselineskip depth .3\baselineskip (12pt)
    const sh = texScale(".7", 12 * PT), sd = texScale(".3", 12 * PT);
    strut.list = [{ type: RULE_NODE, width: 0, height: sh, depth: sd, tag: a }];
    strut.height = sh;
    strut.depth = sd;
    const strutNoad = newNoad(ORD_NOAD, a);
    strutNoad.nucleus = boxField(strut);
    const numList: MItem[] = [strutNoad];
    if (a.align === "r") numList.push({ type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 2, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a } as Node);
    numList.push(...this.list(a.branch("above"), ctx));
    if (a.align === "l") numList.push({ type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 2, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a } as Node);
    const frac = newNoad(FRACTION_NOAD, a);
    frac.num = this.listField(a.branch("above"), ctx, numList);
    frac.den = this.listField(a.branch("below"), ctx);
    frac.thickness = DEFAULT_CODE;
    const fracGroup = this.groupNoad([frac], a);
    const g = this.groupNoad([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, fracGroup], a);
    return [g, { type: KERN_NODE, width: -this.env.nullDelimiterSpace, subtype: NORMAL, tag: a } as Node];
  }

  // ---------------------------------------------------------------- roots --
  private surd(a: AtomLike, ctx: Ctx): MItem[] {
    const index = a.branch("above");
    const sqrtDelim = this.radicalDelimiter();
    const radical = (bodyCtx: Ctx): Noad => {
      const r = newNoad(RADICAL_NOAD, a);
      r.leftDelim = sqrtDelim;
      r.nucleus = this.field(a.branch("body"), bodyCtx);
      return r;
    };
    const visibleIndex = (index ?? []).filter((x) => x.type !== "first");
    if (visibleIndex.length === 0) return [radical(ctx)];
    // amsmath \root#1\of: \setbox\rootbox\hbox{$\m@th\scriptscriptstyle{#1}$}
    // then \mathpalette\r@@t{#2}
    const rootList: MItem[] = [{ type: STYLE_NODE, style: SCRIPT_SCRIPT_STYLE, tag: a }, this.groupNoad(this.list(index, ctx), a)];
    const rootbox = this.mathBox(rootList, TEXT_STYLE, a);
    const variant = (style: number): MItem[] => {
      // \setboxz@h{$\m@th#1\sqrtsign{#2}$}
      const box0 = this.mathBox([{ type: STYLE_NODE, style, tag: a }, radical(ctx)], TEXT_STYLE, a);
      let dimen = box0.height - box0.depth;
      // \setbox\@ne\hbox{$\m@th#1\mskip\uproot@ mu$}; \advance\dimen@ by1.667\wd\@ne
      const box1 = this.mathBox([{ type: STYLE_NODE, style, tag: a }, this.muGlue(0, a)], TEXT_STYLE, a);
      dimen += texScale("1.667", box1.width);
      const raised: BoxNode = { ...rootbox, list: rootbox.list, shift: -texScale(".6", dimen) };
      const rn = newNoad(ORD_NOAD, a);
      rn.nucleus = boxField(raised);
      const sq = newNoad(ORD_NOAD, a);
      sq.nucleus = boxField(box0);
      return [this.muKern(0, a), this.muKern(5 * PT, a), rn, this.muKern(-10 * PT, a), this.muKern(0, a), sq];
    };
    const choice: ChoiceItem = {
      type: CHOICE_NODE,
      display: variant(DISPLAY_STYLE),
      text: variant(TEXT_STYLE),
      script: variant(SCRIPT_STYLE),
      scriptScript: variant(SCRIPT_SCRIPT_STYLE),
      tag: a,
    };
    return [this.groupNoad([choice], a)];
  }

  private radicalDelimiter(): Delimiter {
    // \sqrtsign = \radical"270370
    return { smallFam: 2, smallChar: 0x70, largeFam: 3, largeChar: 0x70 };
  }

  // -------------------------------------------------------------- accents --
  private accent(a: AtomLike, ctx: Ctx): Noad {
    const cmd = a.command ?? "\\hat";
    const tpl = this.data.onearg[cmd];
    const acc = tpl?.mathlist?.length === 1 && tpl.mathlist[0].type === "accent" ? tpl.mathlist[0] : null;
    if (!acc) {
      // Not a TeX math accent (e.g. an over-arrow or brace): defer to fallback.
      return this.groupNoad(this.fallback(a, ctx), a);
    }
    const n = newNoad(ACCENT_NOAD, a);
    n.accentChr = charField(acc.fam, acc.code);
    n.nucleus = this.field(a.branch("body"), ctx);
    return n;
  }

  // ----------------------------------------------------------- delimiters --
  /** Delimiter code for "(", "\\langle", "." etc. (\delcode / \delimiter). */
  delimiter(name: string): Delimiter {
    if (!name || name === "." || name === "\\.") return NULL_DELIMITER;
    let code: number | undefined;
    // mathtools' \lparen and \rparen (the editor's fences use them too)
    if (name === "\\lparen" || name === "\\rparen") name = name === "\\lparen" ? "(" : ")";
    if (name.length === 1) {
      const dc = this.data.meanings.delcode[name];
      if (dc !== undefined && dc > 0) code = dc;
    } else {
      const alias: Record<string, string> = {
        "\\{": "\\lbrace",
        "\\}": "\\rbrace",
        "\\lbrace": "\\lbrace",
        "\\|": "\\|",
        "\\lVert": "\\lVert",
        "\\rVert": "\\rVert",
        "\\vert": "\\vert",
        "\\Vert": "\\Vert",
      };
      const key = alias[name] ?? name;
      const m = this.meaning(key) ?? "";
      const dc = this.delimiterCodeIn(m);
      if (dc !== null) code = dc & 0xffffff;
      else {
        // a macro for one character (\lbrack -> [) has that character's \delcode
        const one = /^macro:->(.)$/.exec(m)?.[1] ?? (key.length === 2 ? key[1] : undefined);
        const dc = one === undefined ? undefined : this.data.meanings.delcode[one];
        if (dc !== undefined && dc > 0) code = dc;
      }
    }
    if (code === undefined) return NULL_DELIMITER;
    const small = (code >> 12) & 0xfff;
    const large = code & 0xfff;
    return { smallFam: (small >> 8) & 15, smallChar: small & 255, largeFam: (large >> 8) & 15, largeChar: large & 255 };
  }

  /**
   * \left...\right. A side the editor still waits for ("?": the closing
   * bracket of a bracket just typed) is the matching delimiter, which the
   * editor draws faded.
   */
  private leftRight(a: AtomLike, ctx: Ctx): Noad {
    const left = newNoad(LEFT_NOAD, a);
    left.delim = this.delimiter(a.leftDelim === "?" ? LEFT_FENCE[a.rightDelim] ?? "." : a.leftDelim ?? ".");
    const right = newNoad(RIGHT_NOAD, a);
    right.delim = this.delimiter(a.rightDelim === "?" ? RIGHT_FENCE[a.leftDelim] ?? a.leftDelim ?? "." : a.rightDelim ?? ".");
    const inner = newNoad(INNER_NOAD, a);
    inner.nucleus = mlistField([left, ...this.list(a.branch("body"), ctx), right]);
    return inner;
  }

  private middle(a: AtomLike): Noad {
    const m = newNoad(RIGHT_NOAD, a);
    m.subtype = MIDDLE_NOAD;
    m.delim = this.delimiter(a.value ?? a.command ?? ".");
    return m;
  }

  /** amsmath \big etc.: an \hbox of \left<d>\vcenter to <f>\big@size{}\right. */
  private sizedDelim(a: AtomLike): Noad {
    const factors = ["1", "1.5", "2", "2.5"];
    const f = factors[Math.max(0, Math.min(3, (a.size ?? 1) - 1))];
    // \big@size = 1.2\ht\Mathstrutbox@ + 1.2\dp\Mathstrutbox@ (amsmath \resetMathstrut@:
    // the character of the \mathcode of "(" in its family's \textfont)
    const mc = this.data.meanings.mathcode["("];
    const pf = this.env.famFnt((mc >> 8) & 15, 0 /* text size */)!;
    const paren = pf.charInfo(mc & 255);
    const bigSize = texScale("1.2", pf.charHeight(paren)) + texScale("1.2", pf.charDepth(paren));
    const vc = newNoad(VCENTER_NOAD, a);
    const vbox = vpackage([], texScale(f, bigSize), EXACTLY, undefined, a);
    vc.nucleus = boxField(vbox);
    const left = newNoad(LEFT_NOAD, a);
    left.delim = this.delimiter(a.value ?? ".");
    const right = newNoad(RIGHT_NOAD, a);
    right.delim = NULL_DELIMITER;
    const env = { ...this.env, nullDelimiterSpace: 0 };
    const box = this.mathBox([left, vc, right], TEXT_STYLE, a, env);
    const cls: Record<string, number> = { mopen: OPEN_NOAD, mclose: CLOSE_NOAD, mrel: REL_NOAD, mbin: BIN_NOAD, mpunct: PUNCT_NOAD };
    const n = newNoad(cls[a.delimType] ?? ORD_NOAD, a);
    n.nucleus = boxField(box);
    return n;
  }

  // ------------------------------------------------------------ operators --
  private operator(a: AtomLike, ctx: Ctx): MItem[] {
    const cmd = a.command ?? "";
    let items: MItem[] | null = null;
    if (cmd === "\\operatorname" || cmd === "\\operatorname*" || cmd === "\\mathop") {
      items = [this.operatorName(a, ctx, cmd === "\\operatorname*")];
    } else if (this.pathOps.has(cmd)) {
      // like amsmath's integrals: \mathop with \nolimits (\ilimits@)
      const op = newNoad(OP_NOAD, a);
      op.nucleus = charField(FAM.paths, this.pathOps.get(cmd)!);
      op.subtype = NO_LIMITS;
      items = [op];
    } else if (cmd === "\\oiiint" && !this.data.templates[cmd]) {
      // no LaTeX package in the reference set defines it: amsmath's \iiint
      // with a contour oval drawn over the three signs (as the app's export does)
      items = this.template("\\iiint", a);
      if (items) this.ovals.add(a);
    } else if (a.explicitSubsupPlacement && a.subsupPlacement === "over-under" && this.data.templates[cmd + "\\limits"]) {
      // amsmath's multiple integrals: \limits makes the signs one \mathop (\ints@b)
      items = this.template(cmd + "\\limits", a);
    } else {
      const mc = this.mathCodeOf(cmd, ctx);
      if (mc !== null) items = [this.mathChar(mc, ctx, a)];
      else items = this.template(cmd, a);
    }
    if (!items || items.length === 0) items = this.unknownSymbol(a, ctx);
    // \limits / \nolimits typed explicitly by the user
    if (a.explicitSubsupPlacement) {
      const last = items[items.length - 1] as Noad;
      if (last.type === OP_NOAD) {
        if (a.subsupPlacement === "over-under") last.subtype = LIMITS;
        else if (a.subsupPlacement === "adjacent") last.subtype = NO_LIMITS;
        else last.subtype = NORMAL;
      }
    }
    return items;
  }

  /** \qopname: \mathop{#1\kern\z@\operator@font#3}\n<limits>limits@ */
  private operatorName(a: AtomLike, ctx: Ctx, star: boolean): Noad {
    const body = a.branch("body") ?? [];
    const inner: MItem[] = [newKern(0, a) as unknown as MItem];
    const opCtx: Ctx = { ...ctx, fam: this.role.rm, operator: true };
    if (body.length > 0) {
      for (const b of body) {
        if (b.type === "first") continue;
        if (b.command === "\\," || b.command === "\\:" || b.command === "\\;" || b.command === "\\!") {
          inner.push(...(this.template(b.command, b) ?? []));
          continue;
        }
        this.atom(b, opCtx, inner);
      }
    }
    else if (a.value) for (const ch of a.value) inner.push(...this.symbol({ type: "mord", value: ch, command: ch, branch: () => undefined }, opCtx));
    const op = newNoad(OP_NOAD, a);
    op.nucleus = mlistField(inner);
    op.subtype = star ? NORMAL : NO_LIMITS;
    return op;
  }

  // -------------------------------------------------------------- spacing --
  private spacing(a: AtomLike, ctx: Ctx): MItem[] {
    const cmd = a.command ?? "";
    const tpl = this.template(cmd, a);
    if (tpl) return tpl;
    if (cmd === "\\hspace" || cmd === "\\hspace*" || cmd === "\\kern" || cmd === "\\hskip") {
      const w = this.dimenFromAtom(a);
      return [{ type: GLUE_NODE, spec: { width: w, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a } as Node];
    }
    if (cmd === "\\mkern" || cmd === "\\mskip" || cmd === "\\mspace") {
      const w = this.dimenFromAtom(a, "mu");
      return [{ type: KERN_NODE, width: w, subtype: MU_GLUE, tag: a } as Node];
    }
    void ctx;
    return [];
  }

  private dimenFromAtom(a: AtomLike, unit = "pt"): number {
    const w = a.width;
    if (!w) return 0;
    const value = typeof w === "object" ? w.dimension ?? w.value ?? 0 : Number(w);
    const u = typeof w === "object" ? w.unit ?? unit : unit;
    // em and ex of the current text font; mu lengths come as points
    const rm = this.textRoman();
    return scanDimen(String(value), u === "mu" ? "pt" : u, { em: rm.param(6), ex: rm.param(5), px: 65781 });
  }

  /**
   * latex.ltx \@rule[#1]#2#3: \hbox{\vrule width #2 height #3+#1 depth -#1}.
   * The kernel gives the optional raise and the two sizes as args.
   */
  private ruleBox(a: AtomLike): BoxNode {
    const args = (a.args ?? []) as unknown[];
    const dimen = (d: unknown) => (d && typeof d === "object" ? this.dimenFromAtom({ type: "spacing", width: d, branch: () => undefined }) : 0);
    const raise = dimen(args[0]);
    const rule = newRule(a);
    rule.width = dimen(args[1]);
    rule.height = dimen(args[2]) + raise;
    rule.depth = -raise;
    return hpack([rule], 0, undefined, a);
  }

  private muKern(width: number, tag: Tag): MItem {
    return { type: KERN_NODE, width, subtype: MU_GLUE, tag } as Node;
  }

  private muGlue(width: number, tag: Tag): MItem {
    return { type: GLUE_NODE, spec: { width, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 }, subtype: MU_GLUE, tag } as Node;
  }

  // ----------------------------------------------------------------- text --
  /** amsmath \text: a \mathchoice of \hbox{text} at 10/10/7/5pt. */
  private textRun(run: AtomLike[]): Noad {
    const tag = runTag(run);
    const mk = (b: BoxNode): MItem[] => {
      const n = newNoad(ORD_NOAD, tag);
      n.nucleus = boxField(b);
      return [n];
    };
    const choice: ChoiceItem = {
      type: CHOICE_NODE,
      display: mk(this.textRunBox(run, 0)),
      text: mk(this.textRunBox(run, 0)),
      script: mk(this.textRunBox(run, 1)),
      scriptScript: mk(this.textRunBox(run, 2)),
      tag,
    };
    return this.groupNoad([choice], tag);
  }

  /** \hbox of a text run in its font at a size (0, 1, 2: the text, script and scriptscript sizes). */
  private textRunBox(run: AtomLike[], sizeIndex: number): BoxNode {
    const s = run[0].style ?? {};
    const t = this.text;
    const sans = s.fontFamily === "sans-serif" || s.fontFamily === "cmss";
    const mono = s.fontFamily === "monospace" || s.fontFamily === "cmtt";
    const family = sans ? t.sf : mono ? t.tt : t.rm;
    const series = s.fontSeries === "b" || s.fontSeries === "bx" ? t.bold : "m";
    const shape = s.fontShape === "it" ? "it" : s.fontShape === "sl" ? "sl" : s.fontShape === "sc" ? "sc" : "n";
    // Collapse consecutive spaces as TeX's input reader would.
    const tokens: TextToken[] = [];
    let lastSpace = false;
    for (const a of run) {
      const v = a.value ?? "";
      for (const ch of v) {
        if (ch === " " || ch === " ") {
          if (!lastSpace) tokens.push({ kind: "space", tag: a });
          lastSpace = true;
        } else {
          tokens.push({ kind: "char", code: ch.codePointAt(0)!, tag: a });
          lastSpace = false;
        }
      }
    }
    // roman text in the editor's letter font, when one is chosen
    const letters = !sans && !mono ? this.fams.get(series !== "m" ? (shape === "it" || shape === "sl" ? FAM.letterBfIt : FAM.letterBf) : shape === "it" || shape === "sl" ? FAM.letterIt : FAM.letterUp) : undefined;
    const f = letters?.[sizeIndex] ?? this.textFont(this.textId(family, series, shape, t.sizes[sizeIndex]));
    const nodes = typesetText(tokens, f, 1000, this.native?.fallback(f.size) ?? null);
    // \textit, \textsl: \check@icr adds the italic correction (\/) of the last
    // character of slanted text, unless a period or comma follows (\nocorrlist)
    const tail = nodes[nodes.length - 1];
    let next: AtomLike | undefined;
    try {
      next = (run[run.length - 1] as { rightSibling?: AtomLike }).rightSibling;
    } catch {
      next = undefined; // an atom outside a kernel tree (no parent)
    }
    if (!letters && f.param(1) > 0 && tail?.type === CHAR_NODE && next?.value !== "." && next?.value !== ",") {
      const italic = tail.font.charItalic(tail.font.charInfo(tail.char));
      if (italic !== 0) nodes.push({ type: KERN_NODE, width: italic, subtype: 1, tag: tail.tag } as Node);
    }
    return hpack(nodes, 0, undefined, runTag(run));
  }

  /**
   * The command being typed (\alph...): its raw characters in typewriter type,
   * at the size of the style it is typed in, as a \text is (10pt, 7pt in a
   * script, 5pt in a script's script).
   */
  private latexInput(a: AtomLike): MItem {
    const atoms = a.type === "latexgroup" ? a.branch("body") ?? [] : [a];
    const tokens: TextToken[] = [];
    for (const c of atoms) {
      if (c.type === "first") continue;
      for (const ch of c.value ?? "") tokens.push(ch === " " ? { kind: "space", tag: c } : { kind: "char", code: ch.codePointAt(0)!, tag: c });
    }
    const atSize = (size: number): MItem[] => {
      const f = this.textFont(this.textId(this.text.tt, "m", "n", this.text.sizes[size]));
      const b = hpack(typesetText(tokens, f, 1000, this.native?.fallback(f.size) ?? null), 0, undefined, a);
      if (atoms[0]?.type === "first") b.firstTag = atoms[0];
      return [this.boxNoad(b, a)];
    };
    return this.palette(a, (style) => atSize(style >= SCRIPT_SCRIPT_STYLE ? 2 : style >= SCRIPT_STYLE ? 1 : 0));
  }

  /** An NFSS font identifier of the profile's encoding ("OT1/cmr/m/n/10"). */
  private textId(family: string, series: string, shape: string, size: string): string {
    return `${this.text.encoding}/${family}/${series}/${shape}/${size}`;
  }

  private textFont(nfssId: string): TeXFont {
    const name = this.data.textFonts[nfssId] ?? this.data.textFonts[this.textId(this.text.rm, "m", "n", this.text.sizes[0])];
    return this.cache.byFontname(name)!;
  }

  /**
   * What the editor's caret is measured from (editor-geometry.ts caretMetrics):
   * the roman "d", amsmath's \Mathstrut@ parenthesis and the math italic "x".
   */
  caretFonts(): { dHeight: number; paren: { height: number; depth: number }; xWidth: number } {
    const rm = this.textRoman();
    const of = (ch: string) => {
      const mc = this.data.meanings.mathcode[ch];
      const f = this.env.famFnt((mc >> 8) & 15, 0)!;
      return { f, i: f.charInfo(mc & 255) };
    };
    const p = of("("), x = of("x");
    return {
      dHeight: rm.charHeight(rm.charInfo(0x64)),
      paren: { height: p.f.charHeight(p.i), depth: p.f.charDepth(p.i) },
      xWidth: x.f.charWidth(x.i),
    };
  }

  /** The document's text font at 10pt (\rmfamily, upright, medium): em, ex and inter-word glue. */
  private textRoman(): TeXFont {
    return this.textFont(this.textId(this.text.rm, "m", "n", this.text.sizes[0]));
  }

  /** An \hbox of plain text in one font (used for unknown symbols, errors). */
  private textBox(text: string, nfssId: string, tag: Tag): BoxNode {
    const tokens: TextToken[] = [...text].map((ch) => (ch === " " ? { kind: "space", tag } : { kind: "char", code: ch.codePointAt(0)!, tag }));
    const f = this.textFont(nfssId);
    return hpack(typesetText(tokens, f, 1000, this.native?.fallback(f.size) ?? null), 0, undefined, tag);
  }

  // -------------------------------------------------- phantoms and boxes --
  private styleList(style: number, a: AtomLike, items: MItem[]): MItem[] {
    return [{ type: STYLE_NODE, style, tag: a }, ...items];
  }

  /** \mathpalette: one variant per style, as a \mathchoice. */
  private palette(a: AtomLike, make: (style: number) => MItem[]): ChoiceItem {
    // every variant is typeset ahead, but only the one for the style TeX meets
    // is kept: the atom scales the editor records come from that one alone
    const scales = this.env.atomScales;
    this.env.atomScales = undefined;
    try {
      return {
        type: CHOICE_NODE,
        display: make(DISPLAY_STYLE),
        text: make(TEXT_STYLE),
        script: make(SCRIPT_STYLE),
        scriptScript: make(SCRIPT_SCRIPT_STYLE),
        tag: a,
      };
    } finally {
      this.env.atomScales = scales;
    }
  }

  private boxNoad(b: BoxNode, a: Tag, type = ORD_NOAD): Noad {
    const n = newNoad(type, a);
    n.nucleus = boxField(b);
    return n;
  }

  /** \phantom, \vphantom, \hphantom (latex.ltx \mathph@nt + \finph@nt), \smash. */
  private phantom(a: AtomLike, ctx: Ctx): MItem[] {
    const cmd = a.command ?? "\\phantom";
    const body = a.branch("body");
    if (cmd === "\\smash") {
      // amsmath \smash[tb]: \setbox\z@\hbox{$\m@th#1{#2}$} then zero ht and/or dp
      const which = a.smashHeight === false ? "b" : a.smashDepth === false ? "t" : "tb";
      return [
        this.palette(a, (st) => {
          const b = this.mathBox(this.styleList(st, a, [this.groupNoad(this.list(body, ctx), a)]), TEXT_STYLE, a);
          if (which !== "b") b.height = 0;
          if (which !== "t") b.depth = 0;
          return [this.boxNoad(b, a)];
        }),
      ];
    }
    const v = cmd !== "\\hphantom";
    const h = cmd !== "\\vphantom";
    return [
      this.palette(a, (st) => {
        const b0 = this.mathBox(this.styleList(st, a, [this.groupNoad(this.list(body, ctx), a)]), TEXT_STYLE, a);
        const b2 = newNullBox(a);
        b2.role = "phantom";
        if (v) {
          b2.height = b0.height;
          b2.depth = b0.depth;
        }
        if (h) b2.width = b0.width;
        return [this.boxNoad(b2, a)];
      }),
    ];
  }

  /**
   * \boxed{#1} = \fbox{\m@th$\displaystyle#1$}; xcolor \fcolorbox{f}{b}{#1} =
   * \color@b@x{\fboxsep\z@\color{f}\fbox}{\color{b}}{#1}.
   */
  private boxAtom(a: AtomLike, ctx: Ctx): MItem[] {
    const cmd = a.command ?? "\\boxed";
    if (cmd === "\\boxed" || cmd === "\\fbox") {
      const content = this.mathBox(this.styleList(DISPLAY_STYLE, a, this.list(a.branch("body"), ctx)), TEXT_STYLE, a);
      return [this.boxNoad(this.frameBox(content, 3 * PT, undefined, a), a)];
    }
    if (cmd === "\\fcolorbox") {
      const color = (v: unknown) => (v && typeof v === "object" ? (v as { string?: string }).string : typeof v === "string" ? v : undefined);
      const [block, back, noad] = this.colorBox(this.textModeBox(a.branch("body"), ctx, a), color(a.backgroundcolor) ?? "yellow", a);
      const inner = hpack([block as unknown as Node, back as unknown as Node, (noad as Noad).nucleus.box!], 0, undefined, a);
      return [this.boxNoad(this.frameBox(inner, 0, color(a.framecolor), a), a)];
    }
    return this.fallback(a, ctx);
  }

  /** latex.ltx \@frameb@x: the box framed by \fboxrule rules \fboxsep away. */
  private frameBox(content: BoxNode, fboxsep: number, color: string | undefined, a: AtomLike): BoxNode {
    const fboxrule = decimalToScaled(".4");
    const tempboxa = hpack([newKern(fboxsep, a), content, newKern(fboxsep, a)], 0, undefined, a);
    const tempdima = fboxrule + fboxsep + tempboxa.depth;
    const vglue = (w: number): Node => ({ type: GLUE_NODE, spec: { width: w, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a });
    const innerV = vpackage([vglue(fboxsep), tempboxa, vglue(fboxsep)], 0, undefined, undefined, a);
    const vrule = (): Node => ({ type: RULE_NODE, width: fboxrule, height: -(2 ** 30), depth: -(2 ** 30), tag: a, color });
    const hrule = (): Node => ({ type: RULE_NODE, width: -(2 ** 30), height: fboxrule, depth: 0, tag: a, color });
    const middle = hpack([vrule(), innerV, vrule()], 0, undefined, a);
    const frame = vpackage([hrule(), middle, hrule()], 0, undefined, undefined, a);
    const lowered = hpack([frame], 0, undefined, a);
    lowered.shift = tempdima;
    const outer = hpack([lowered], 0, undefined, a);
    outer.role = "boxed";
    return outer;
  }

  /** An \hbox of text-mode material: text runs at 10pt, $...$ in text style. */
  private textModeBox(atoms: AtomLike[] | undefined, ctx: Ctx, tag: Tag): BoxNode {
    const parts: Node[] = [];
    const visible = (atoms ?? []).filter((x) => x.type !== "first");
    for (const run of this.styleRuns(visible, ctx, true)) {
      if (run[0].mode === "text") parts.push(this.textRunBox(run, 0));
      else {
        const items: MItem[] = [];
        for (const r of this.styleRuns(run, ctx, false)) this.styledRun(r, ctx, items);
        parts.push(this.mathBox(items, TEXT_STYLE, run[0]));
      }
    }
    return hpack(parts, 0, undefined, tag);
  }

  // ------------------------------------------------------ over and under --
  /** amsmath \binrel@: the class (bin/rel/ord) of a math list, by measurement. */
  private binrelClass(make: () => MItem[], a: AtomLike): number {
    const env: MathEnv = {
      ...this.env,
      thinMuSkip: { width: 0, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 },
      medMuSkip: { width: -PT, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 },
      thickMuSkip: { width: PT, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 },
      // a measurement: TeX's own class rules, nothing recorded for the editor
      editorBinSpacing: false,
      atomScales: undefined,
    };
    const box2 = this.mathBox(make(), TEXT_STYLE, a, env);
    const empty = () => {
      const n = newNoad(ORD_NOAD, a);
      n.nucleus = mlistField([]);
      return n;
    };
    const box0 = this.mathBox([empty(), ...make(), empty()], TEXT_STYLE, a, env);
    const w = box0.width - box2.width;
    return w < 0 ? BIN_NOAD : w > 0 ? REL_NOAD : ORD_NOAD;
  }

  private overUnder(a: AtomLike, ctx: Ctx): MItem[] {
    const cmd = a.command ?? "";
    const body = a.branch("body");
    const above = a.branch("above");
    const below = a.branch("below");
    const nonEmpty = (x: AtomLike[] | undefined) => (x ?? []).some((y) => y.type !== "first");
    if (cmd === "\\stackrel") {
      // \stackrel#1#2 = \mathrel{\mathop{#2}\limits^{#1}}
      const op = newNoad(OP_NOAD, a);
      op.subtype = LIMITS;
      op.nucleus = this.field(body, ctx);
      op.sup = this.field(above, ctx);
      const rel = newNoad(REL_NOAD, a);
      rel.nucleus = this.simplify([op]);
      return [rel];
    }
    if (cmd === "\\overset" || cmd === "\\underset" || cmd === "\\overunderset") {
      // \binrel@{#2}\binrel@@{\mathop{\kern\z@#2}\limits^{#1}}
      const cls = this.binrelClass(() => this.list(body, ctx), a);
      const op = newNoad(OP_NOAD, a);
      op.subtype = LIMITS;
      op.nucleus = this.simplify([newKern(0, a) as unknown as MItem, ...this.list(body, ctx)]);
      if (cmd !== "\\underset" && nonEmpty(above)) op.sup = this.field(above, ctx);
      if (cmd !== "\\overset" && nonEmpty(below)) op.sub = this.field(below, ctx);
      if (cls === ORD_NOAD) return [this.groupNoad([op], a)];
      const w = newNoad(cls, a);
      w.nucleus = this.simplify([op]);
      return [w];
    }
    if (cmd === "\\overbrace" || cmd === "\\underbrace") return [this.brace(a, ctx, cmd === "\\overbrace")];
    const over = OVER_ARROWS[cmd];
    if (over) return [this.overArrow(a, ctx, over)];
    const ext = EXT_ARROWS[cmd];
    if (ext) return [this.extArrow(a, ctx, ext)];
    return this.fallback(a, ctx);
  }

  // ---------------------------------------------------------- cancel.sty --
  /** \cancel, \bcancel, \xcancel (the kernel's enclose notations); others keep their content. */
  private enclose(a: AtomLike, ctx: Ctx): MItem[] {
    const n = (a.notation ?? {}) as Record<string, boolean>;
    const keys = Object.keys(n).filter((k) => n[k]);
    const up = !!n.updiagonalstrike, down = !!n.downdiagonalstrike;
    if (keys.length === 0 || keys.some((k) => k !== "updiagonalstrike" && k !== "downdiagonalstrike")) return this.fallback(a, ctx);
    // \cancel = \@can@slash{}, \bcancel = \@can@slash{-}, \xcancel = \@can@slash{+}\@can@slash{-}
    const slashes = up && down ? [1, -1] : down ? [-1] : [1];
    return [this.palette(a, (st) => this.cancelBox(slashes, st, a, ctx))];
  }

  /**
   * cancel.sty \@cancel#1#2#3: the content in an \hbox, vcentred; the slashes
   * (\line in picture mode) and the content overlap in an \ooalign, raised
   * back to the content's baseline.
   */
  private cancelBox(slashes: number[], style: number, a: AtomLike, ctx: Ctx): MItem[] {
    // \@begin@tempboxa\hbox{\m@th$#2{#3}$}
    const box0 = this.mathBox(this.styleList(style, a, [this.groupNoad(this.list(a.branch("body"), ctx), a)]), TEXT_STYLE, a);
    const totalheight = box0.height + box0.depth;
    let dimen = box0.height;
    // \setbox\@tempboxa\hbox{$\m@th\vcenter{\box\@tempboxa}$}
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(vpack([box0], a));
    const tempboxa = this.mathBox([vc], TEXT_STYLE, a);
    dimen -= tempboxa.height;
    const hskip = (w: number, stretch = 0, order = NORMAL): Node => ({ type: GLUE_NODE, spec: { width: w, stretch, stretchOrder: order, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a });
    const hideskip = () => hskip(-1000 * PT, 65536, 2); // \hideskip = -1000pt plus 1fill
    const mathOn = (): Node => ({ type: MATH_NODE, width: 0, subtype: 0, tag: a } as Node);
    const rows: (AlignRow | AlignNoalign)[] = [];
    for (const sign of slashes) {
      const line = hpack(this.canSlashLine(sign, tempboxa.width, totalheight, a), 0, undefined, a);
      const lv = newNoad(VCENTER_NOAD, a);
      lv.nucleus = boxField(vpack([line], a));
      const ts = new MathTypesetter(this.env);
      const math = ts.mlistToHlist([lv], TEXT_STYLE, false);
      rows.push({ kind: "row", cells: [[hideskip(), mathOn(), ...math, { ...mathOn(), subtype: 1 } as Node, hideskip()]] });
    }
    rows.push({ kind: "row", cells: [[hskip(0, 65536, 1), tempboxa, hskip(0, 65536, 1)]] });
    // \ooalign = \lineskiplimit-\maxdimen\oalign: \vtop{\baselineskip\z@skip \lineskip.25ex \ialign{...}}
    const ex = this.textRoman().param(5);
    const v = new VList({ baselineskip: glue(0), lineskip: glue(texScale(".25", ex)), lineskiplimit: -MAX_DIMEN }, a);
    halign(rows, [glue(0)], v, a);
    const vtop = vpack(v.items, a);
    const first = vtop.list[0];
    const h = first && (first.type === HLIST_NODE || first.type === VLIST_NODE || first.type === RULE_NODE) ? (first as BoxNode).height : 0;
    vtop.depth = vtop.depth - h + vtop.height;
    vtop.height = h;
    // {\/\raise\dimen@\hbox{\ooalign{...}}}: in math \/ is a zero kern
    const raised = hpack([vtop], 0, undefined, a);
    raised.shift = -dimen;
    const inner = newNoad(ORD_NOAD, a);
    inner.nucleus = boxField(raised);
    return [this.groupNoad([newKern(0, a) as unknown as MItem, inner], a)];
  }

  /** cancel.sty \@can@slash: the slope and length of the slash, then \line. */
  private canSlashLine(sign: number, width: number, totalheight0: number, a: AtomLike): Node[] {
    const minPt = (d: number, p: number) => (d < p * PT ? p * PT : d);
    let dimen = minPt(width, 2);
    let totalheight = minPt(totalheight0, 6);
    let x: number, y: number;
    if (totalheight < dimen) {
      // wide
      dimen = minPt(dimen, 8);
      const c = Math.trunc((totalheight * 5) / dimen);
      dimen += 2 * PT;
      [x, y] = [[6, 1], [4, 1], [2, 1], [4, 3]][c] ?? [1, 1];
    } else {
      // tall
      totalheight = minPt(totalheight, 8) + 2 * PT;
      const c = Math.trunc((dimen * 5) / totalheight);
      dimen = texScale([".16", ".25", ".5", ".75"][c] ?? "1", totalheight);
      [x, y] = [[1, 6], [1, 4], [1, 2], [3, 4]][c] ?? [1, 1];
    }
    return this.pictureSline(x, sign * y, dimen, a);
  }

  /** latex.ltx \@sline: a slanted \line(x,y){len} of line10 characters (\thinlines). */
  private pictureSline(xarg: number, yarg: number, linelen: number, a: AtomLike): Node[] {
    const f = this.cache.get("line10");
    const out: Node[] = [];
    const hskip = (w: number): Node => ({ type: GLUE_NODE, spec: { width: w, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a });
    const negarg = xarg < 0;
    const yyarg = negarg ? -yarg : yarg;
    if (negarg) xarg = -xarg;
    // \@getlinechar(x,y)
    const c = xarg * 8 - 9 + (yyarg > 0 ? yyarg : -yyarg + 64);
    const info = f.charInfo(c);
    if (!info.exists) return out;
    const wd = f.charWidth(info), ht = f.charHeight(info);
    if (wd === 0) return out;
    const raise = yarg > 0; // \@upordown: \raise, else \lower
    const place = (dist: number) => {
      const b = hpack([newCharacter(f, c, a)], 0, undefined, a);
      b.shift = raise ? -dist : dist;
      out.push(b);
    };
    let clnht = raise ? 0 : ht;
    let clnwd = wd;
    if (negarg) out.push(hskip(-wd));
    while (clnwd < linelen) {
      place(clnht);
      if (negarg) out.push(hskip(-2 * wd));
      clnht += ht;
      clnwd += wd;
    }
    clnht -= ht;
    clnwd -= wd;
    let tempdima = linelen - clnwd;
    const tempdimb = tempdima - wd;
    out.push(hskip(negarg ? -tempdimb : tempdimb));
    tempdima *= 1000;
    const tempcnta = Math.trunc(tempdima / wd);
    clnht += Math.trunc((ht * tempcnta) / 1000);
    if (linelen < wd) out.push(hskip(wd));
    else place(clnht);
    return out;
  }

  // ---------------------------------------------------- braces and arrows --
  /** \ialign{##...}: rows of horizontal material in internal vertical mode. */
  private ialignBox(items: (AlignRow | AlignNoalign)[], a: AtomLike, vtop = false): BoxNode {
    const v = new VList({ baselineskip: glue(12 * PT), lineskip: glue(PT), lineskiplimit: 0 }, a);
    halign(items, [glue(0)], v, a);
    const box = vpack(v.items, a);
    if (vtop) {
      // \vtop: the reference point moves to the baseline of the first box
      const first = box.list[0];
      const h = first && (first.type === HLIST_NODE || first.type === VLIST_NODE || first.type === RULE_NODE) ? (first as BoxNode).height : 0;
      box.depth = box.depth - h + box.height;
      box.height = h;
    }
    return box;
  }

  private kernNoalign(w: number, a: AtomLike, nointerline = false): AlignNoalign {
    return {
      kind: "noalign",
      apply: (v) => {
        v.append(newKern(w, a));
        if (nointerline) v.noInterlineSkip();
      },
    };
  }

  /** fontmath.ltx \downbracefill / \upbracefill as a math formula. */
  private braceFill(down: boolean, a: AtomLike): Node[] {
    if (this.meaning("\\downbracefill")?.includes("\\br@cext")) return this.braceFillLeaders(down, a);
    const ex = (c: number) => this.mathChar(0x0300 + c, BASE_CTX, a); // \braceld etc: \mathord fam3
    const ex3 = this.env.famFnt(3, 0 /* text size */)!; // \textfont3: the height of \braceld
    const ht = ex3.charHeight(ex3.charInfo(0x7a));
    const leaders = (): MItem =>
      ({
        type: GLUE_NODE,
        spec: glue(0, PT, 2),
        subtype: 100,
        leader: { type: RULE_NODE, width: decimalToScaled(".4"), height: ht, depth: 0, tag: a },
        tag: a,
      }) as Node;
    const [l1, r1, l2, r2] = down ? [0x7a, 0x7d, 0x7c, 0x7b] : [0x7c, 0x7b, 0x7a, 0x7d];
    return this.cellFormula([ex(l1), leaders(), ex(r1), ex(l2), leaders(), ex(r2)], a);
  }

  /**
   * newtxmath's brace fills: the extension piece \br@cext as \cleaders,
   * $\m@th\braceld\mkern-1mu\cleaders\hbox{$\mkern-.5mu\br@cext\mkern-.5mu$}\hfill\mkern-1mu
   * \braceru\bracelu\mkern-1mu\cleaders...\hfill\mkern-1mu\bracerd$ (\upbracefill: lu rd ld ru).
   */
  private braceFillLeaders(down: boolean, a: AtomLike): Node[] {
    const piece = (cmd: string) => this.mathChar(this.mathCodeOf(cmd, BASE_CTX) ?? 0, BASE_CTX, a);
    const ext = this.mathBox([this.muKern(-PT / 2, a), piece("\\br@cext"), this.muKern(-PT / 2, a)], TEXT_STYLE, a);
    const leaders = (): MItem => ({ type: GLUE_NODE, spec: glue(0, PT, 2), subtype: 101, leader: ext, tag: a }) as Node;
    const [l1, r1, l2, r2] = down ? ["\\braceld", "\\braceru", "\\bracelu", "\\bracerd"] : ["\\bracelu", "\\bracerd", "\\braceld", "\\braceru"];
    const k = () => this.muKern(-PT, a);
    return this.cellFormula([piece(l1), k(), leaders(), k(), piece(r1), piece(l2), k(), leaders(), k(), piece(r2)], a);
  }

  /** \overbrace / \underbrace: \mathop{\vbox{\ialign{...}}}\limits */
  private brace(a: AtomLike, ctx: Ctx, over: boolean): Noad {
    const content = (): Node[] => [
      this.hfil(a),
      ...this.cellFormula([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, this.groupNoad(this.list(a.branch("body"), ctx), a)], a),
      this.hfil(a),
    ];
    const items: (AlignRow | AlignNoalign)[] = over
      ? [this.kernNoalign(3 * PT, a), { kind: "row", cells: [this.braceFill(true, a)] }, this.kernNoalign(3 * PT, a, true), { kind: "row", cells: [content()] }]
      : [{ kind: "row", cells: [content()] }, this.kernNoalign(3 * PT, a, true), { kind: "row", cells: [this.braceFill(false, a)] }, this.kernNoalign(3 * PT, a)];
    const box = this.ialignBox(items, a, !over);
    box.role = over ? "overbrace" : "underbrace";
    const op = newNoad(OP_NOAD, a);
    op.subtype = LIMITS;
    op.nucleus = boxField(box);
    return op;
  }

  /** amsmath \arrowfill@ in a given style: $\m@th ... #4#1\mkern-7mu\cleaders\hbox{...}\hfill\mkern-7mu#3$ */
  private arrowFill(kind: ArrowFill, style: number, a: AtomLike): Node[] {
    const env: MathEnv = {
      ...this.env,
      thickMuSkip: glue(0),
      medMuSkip: glue(0),
      thinMuSkip: glue(0),
    };
    const piece = (name: string): MItem[] => {
      if (name === "\\relbar" || name === "\\Relbar") {
        // amsmath: \relbar = \mathrel{\smash-}, \Relbar = \mathrel=
        if (name === "\\Relbar") return [this.mathChar(0x303d, BASE_CTX, a)];
        const minus = this.mathChar(this.data.meanings.mathcode["-"], BASE_CTX, a);
        const smashed = this.palette(a, (st) => {
          const b = this.mathBox([{ type: STYLE_NODE, style: st, tag: a }, this.groupNoad([this.mathChar(this.data.meanings.mathcode["-"], BASE_CTX, a)], a)], TEXT_STYLE, a);
          b.height = 0;
          b.depth = 0;
          return [this.boxNoad(b, a)];
        });
        void minus;
        const rel = newNoad(REL_NOAD, a);
        rel.nucleus = mlistField([smashed]);
        return [rel];
      }
      return [this.mathChar(this.mathCodeOf(name, BASE_CTX) ?? 0, BASE_CTX, a)];
    };
    const ts = (items: MItem[]): Node[] => {
      const t = new MathTypesetter(env);
      const h = t.mlistToHlist(items, TEXT_STYLE, false);
      const on: Node = { type: MATH_NODE, width: 0, subtype: 0, tag: a };
      return [on, ...h, { ...on }];
    };
    const fillBox = hpack(ts([{ type: STYLE_NODE, style, tag: a }, this.muKern(-2 * PT, a), ...piece(kind[1]), this.muKern(-2 * PT, a)]), 0, undefined, a);
    const cleaders: Node = { type: GLUE_NODE, spec: glue(0, PT, 2), subtype: 101, leader: fillBox, tag: a };
    return ts([
      { type: STYLE_NODE, style, tag: a },
      ...piece(kind[0]),
      this.muKern(-7 * PT, a),
      cleaders as MItem,
      this.muKern(-7 * PT, a),
      ...piece(kind[2]),
    ]);
  }

  /** amsmath \overrightarrow etc.: \mathpalette{\overarrow@\rightarrowfill@} */
  private overArrow(a: AtomLike, ctx: Ctx, spec: { fill: ArrowFill; under: boolean }): MItem {
    return this.palette(a, (st) => {
      // $\m@th\hfil#2#3\hfil$ : the \hfil are inside the formula
      const contentRow = this.cellFormula(
        [
          { type: GLUE_NODE, spec: glue(0, PT, 1), subtype: NORMAL, tag: a } as Node,
          { type: STYLE_NODE, style: st, tag: a },
          ...this.list(a.branch("body"), ctx),
          { type: GLUE_NODE, spec: glue(0, PT, 1), subtype: NORMAL, tag: a } as Node,
        ],
        a,
      );
      const fill = this.arrowFill(spec.fill, st, a);
      const items: (AlignRow | AlignNoalign)[] = spec.under
        ? [{ kind: "row", cells: [contentRow] }, this.kernNoalign(texScale("1.3", PT), a, true), { kind: "row", cells: [fill] }]
        : [{ kind: "row", cells: [fill] }, { kind: "noalign", apply: (v) => v.noInterlineSkip() }, { kind: "row", cells: [contentRow] }];
      const box = this.ialignBox(items, a, spec.under);
      box.role = "over-arrow";
      return [this.boxNoad(box, a)];
    });
  }

  /** amsmath \ext@arrow (\xrightarrow and friends). */
  private extArrow(a: AtomLike, ctx: Ctx, spec: { k: [number, number, number, number]; fill: ArrowFill }): Noad {
    const [k1, k2, k3, k4] = spec.k;
    const below = a.branch("below");
    const above = a.branch("above");
    const nonEmpty = (x: AtomLike[] | undefined) => (x ?? []).some((y) => y.type !== "first");
    const box0 = hpack(this.arrowFill(spec.fill, DISPLAY_STYLE, a), 0, undefined, a);
    const labelBox = (x: AtomLike[] | undefined) =>
      this.mathBox([{ type: STYLE_NODE, style: SCRIPT_STYLE, tag: a }, this.muKern(k3 * PT, a), this.groupNoad(this.list(x, ctx), a), this.muKern(k4 * PT, a)], TEXT_STYLE, a);
    const w = Math.max(labelBox(below).width, labelBox(above).width, box0.width);
    const arrow = hpack(box0.list, w, EXACTLY, a);
    arrow.role = "ext-arrow";
    const op = newNoad(OP_NOAD, a);
    op.subtype = LIMITS;
    op.nucleus = boxField(arrow);
    const script = (x: AtomLike[] | undefined): Field => {
      const items: MItem[] = [];
      if (k1 !== 0) items.push(this.muKern(k1 * PT, a));
      items.push(...this.list(x, ctx));
      if (k2 !== 0) items.push(this.muKern(k2 * PT, a));
      return this.simplify(items);
    };
    if (nonEmpty(above)) op.sup = script(above);
    if (nonEmpty(below)) op.sub = script(below);
    const rel = newNoad(REL_NOAD, a);
    rel.nucleus = mlistField([op]);
    return rel;
  }

  // --------------------------------------------------------- environments --
  /** LaTeX's \strutbox at 10pt: .7\baselineskip and .3\baselineskip. */
  private strutDims(): [number, number] {
    return [texScale(".7", 12 * PT), texScale(".3", 12 * PT)];
  }

  /** "$\relax <atoms>$" in an alignment cell: an independent text-style formula. */
  private cellFormula(items: MItem[], a: AtomLike): Node[] {
    const ts = new MathTypesetter(this.env);
    const h = ts.mlistToHlist(items, TEXT_STYLE, false);
    const on: Node = { type: MATH_NODE, width: 0, subtype: 0, tag: a };
    return [on, ...h, { ...on }];
  }

  private hfil(a: AtomLike): Node {
    return { type: GLUE_NODE, spec: glue(0, PT, 1), subtype: NORMAL, tag: a };
  }

  private hskip(w: number, a: AtomLike): Node {
    return { type: GLUE_NODE, spec: glue(w), subtype: NORMAL, tag: a };
  }

  /** Split a row's cells; leading \hline "error" atoms become rules above the row. */
  private rowCells(row: AtomLike[][]): { cells: AtomLike[][]; hlines: number } {
    let hlines = 0;
    const cells = row.map((cell, j) => {
      if (j !== 0) return cell;
      const out: AtomLike[] = [];
      for (const x of cell) {
        if (x.type === "error" && x.command === "\\hline" && out.every((y) => y.type === "first")) hlines++;
        else out.push(x);
      }
      return out;
    });
    return { cells, hlines };
  }

  private hlineNoalign(n: number, a: AtomLike): AlignNoalign {
    const rw = decimalToScaled(".4");
    return {
      kind: "noalign",
      apply: (v: VList) => {
        for (let i = 0; i < n; i++) {
          if (i > 0) {
            v.append(this.hskip(2 * PT, a)); // \vskip\doublerulesep (as glue in a vlist)
            v.append(this.hskip(-rw, a)); // \vskip-\arrayrulewidth
          }
          v.appendRule({ type: RULE_NODE, width: -(2 ** 30), height: rw, depth: 0, tag: a });
        }
      },
    };
  }

  /** Row gap of \\[len] (kernel rowGaps entry) in sp; 0 if none. */
  private rowGap(a: AtomLike, r: number): number {
    const g = a.rowGaps?.[r];
    if (!g) return 0;
    const v = typeof g === "object" ? g.dimension ?? 0 : Number(g);
    if (!v) return 0;
    const unit = typeof g === "object" ? g.unit ?? "pt" : "pt";
    return this.dimenFromAtom({ type: "spacing", width: { dimension: v, unit }, branch: () => undefined });
  }

  /**
   * LaTeX's \@mkpream for the array environment: column templates as pieces.
   */
  private mkpream(spec: string): { pre: string[]; align: "c" | "l" | "r"; post: string[] }[] {
    type Tok = { ch: string; arg?: string };
    const toks: Tok[] = [];
    for (let i = 0; i < spec.length; i++) {
      const ch = spec[i];
      if (ch === " ") continue;
      if (ch === "@" || ch === "p" || ch === "*") {
        // read a braced argument (or a single char)
        let j = i + 1;
        while (spec[j] === " ") j++;
        let arg = "";
        if (spec[j] === "{") {
          let depth = 0;
          let k = j;
          for (; k < spec.length; k++) {
            if (spec[k] === "{") depth++;
            else if (spec[k] === "}" && --depth === 0) break;
          }
          arg = spec.slice(j + 1, k);
          i = k;
        } else {
          arg = spec[j] ?? "";
          i = j;
        }
        if (ch === "*") {
          // *{n}{cols}
          let j2 = i + 1;
          while (spec[j2] === " ") j2++;
          let cols = "";
          if (spec[j2] === "{") {
            let depth = 0;
            let k = j2;
            for (; k < spec.length; k++) {
              if (spec[k] === "{") depth++;
              else if (spec[k] === "}" && --depth === 0) break;
            }
            cols = spec.slice(j2 + 1, k);
            i = k;
          }
          const n = parseInt(arg, 10) || 0;
          spec = spec.slice(0, i + 1) + cols.repeat(n) + spec.slice(i + 1);
          continue;
        }
        toks.push({ ch, arg });
        continue;
      }
      toks.push({ ch });
    }
    const out: string[] = [];
    let firstamp = true;
    let last = 6;
    const addamp = () => {
      if (firstamp) firstamp = false;
      else out.push("&");
    };
    const acol = () => out.push("acol");
    for (const t of toks) {
      let cls: number;
      if (t.ch === "@") cls = 2;
      else if (t.ch === "|") cls = 1;
      else if (t.ch === "p") cls = 3;
      else cls = 0;
      switch (cls) {
        case 0:
          switch (last) {
            case 0: acol(); addamp(); acol(); break;
            case 1: addamp(); acol(); break;
            case 4: addamp(); break;
            case 5: acol(); addamp(); acol(); break;
            case 6: firstamp = false; acol(); break;
          }
          out.push("cell:" + (t.ch === "l" ? "l" : t.ch === "r" ? "r" : "c"));
          break;
        case 1:
          switch (last) {
            case 0: acol(); out.push("rule"); break;
            case 1: out.push("dblsep", "rule"); break;
            case 4: out.push("rule"); break;
            case 5: acol(); out.push("rule"); break;
            case 6: out.push("rule"); break;
          }
          break;
        case 2:
          // \@classii, then the @-argument (\@arrayclassiv: $#$)
          if (last === 1) out.push("halfrule");
          out.push("math:" + (t.arg ?? ""));
          cls = 4;
          break;
      }
      last = cls;
    }
    if (last === 0 || last === 5) acol();
    const cols: { pre: string[]; align: "c" | "l" | "r"; post: string[] }[] = [];
    let cur: string[] = [];
    const flush = () => {
      const i = cur.findIndex((x) => x.startsWith("cell:"));
      if (i >= 0) cols.push({ pre: cur.slice(0, i), align: cur[i].slice(5) as "c" | "l" | "r", post: cur.slice(i + 1) });
      cur = [];
    };
    for (const x of out) {
      if (x === "&") flush();
      else cur.push(x);
    }
    flush();
    return cols;
  }

  private piece(x: string, a: AtomLike, ctx: Ctx): Node[] {
    const rw = decimalToScaled(".4");
    switch (x) {
      case "acol":
        return [this.hskip(scalePermille(5 * PT, this.arrayColSep), a)]; // \arraycolsep
      case "rule":
        return [this.hskip(-texScale(".5", rw), a), { type: RULE_NODE, width: rw, height: -(2 ** 30), depth: -(2 ** 30), tag: a }, this.hskip(-texScale(".5", rw), a)];
      case "halfrule":
        return [this.hskip(texScale(".5", rw), a)];
      case "dblsep":
        return [this.hskip(2 * PT, a)];
    }
    if (x.startsWith("math:")) {
      const src = x.slice(5).trim();
      // the @{...} material is typeset as a formula; we support the forms
      // LaTeX environments use ("" and "\quad")
      const items: MItem[] = src === "\\quad" ? [this.hskip(this.textRoman().param(6), a) as MItem] : [];
      void ctx;
      return this.cellFormula(items, a);
    }
    return [];
  }

  /**
   * A trailing \\ (an empty last row) adds no row in LaTeX; the editor keeps
   * it, since the caret of a line just started with Shift+Enter stands there.
   */
  private dropsRow(cells: AtomLike[][], r: number, count: number): boolean {
    return !this.editor && r === count - 1 && r > 0 && !cells.some((c) => c.some((x) => x.type !== "first"));
  }

  /** LaTeX \array (latex.ltx \@array): returns the \vcenter noad. */
  private latexArray(a: AtomLike, spec: string, arraystretch: string, ctx: Ctx, cellStyle: number | null = null): Noad {
    const cols = this.mkpream(spec);
    const [sh, sd] = this.strutDims();
    // \arraystretch, times the user's row spacing
    const arH = scalePermille(texScale(arraystretch, sh), this.arrayRowSep);
    const arD = scalePermille(texScale(arraystretch, sd), this.arrayRowSep);
    const items: (AlignRow | AlignNoalign)[] = [];
    const rows: AtomLike[][][] = a.rows ?? [];
    rows.forEach((row, r) => {
      const { cells, hlines } = this.rowCells(row);
      if (hlines > 0) items.push(this.hlineNoalign(hlines, a));
      if (this.dropsRow(cells, r, rows.length)) return;
      const cellLists: Node[][] = cells.slice(0, cols.length).map((cell, j) => {
        const col = cols[j];
        const list: Node[] = [];
        if (j === 0) list.push({ type: RULE_NODE, width: 0, height: arH, depth: arD, tag: a }); // \@arstrut
        for (const p of col.pre) list.push(...this.piece(p, a, ctx));
        if (col.align !== "l") list.push(this.hfil(a));
        const content = this.list(cell, ctx);
        list.push(...this.cellFormula(cellStyle === null ? content : [{ type: STYLE_NODE, style: cellStyle, tag: a }, ...content], a));
        if (j === cells.length - 1 && r < rows.length - 1) {
          const gap = this.rowGap(a, r);
          if (gap > 0) list.push({ type: RULE_NODE, width: 0, height: 0, depth: arD + gap, tag: a });
        }
        if (col.align !== "r") list.push(this.hfil(a));
        for (const p of col.post) list.push(...this.piece(p, a, ctx));
        return list;
      });
      items.push({ kind: "row", cells: cellLists, tag: a });
      const gap = this.rowGap(a, r);
      if (gap < 0 && r < rows.length - 1) items.push({ kind: "noalign", apply: (v) => v.append(this.hskip(gap, a)) });
    });
    const v = new VList({ baselineskip: glue(0), lineskip: glue(0), lineskiplimit: 0 }, a);
    halign(items, [glue(0)], v, a);
    const box = vpack(v.items, a);
    box.role = "array";
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(box);
    return vc;
  }

  private leftRightWrap(a: AtomLike, left: string, right: string, body: MItem[]): Noad {
    const l = newNoad(LEFT_NOAD, a);
    l.delim = this.delimiter(left);
    const rr = newNoad(RIGHT_NOAD, a);
    rr.delim = this.delimiter(right);
    const inner = newNoad(INNER_NOAD, a);
    inner.nucleus = mlistField([l, ...body, rr]);
    return inner;
  }

  /** \null: an empty \hbox (an ord noad in math). */
  private nullBox(a: AtomLike): Noad {
    return this.boxNoad(newNullBox(a), a);
  }

  private arrayEnv(a: AtomLike, ctx: Ctx): MItem[] {
    const env: string = a.environmentName ?? "matrix";
    const acs = scalePermille(5 * PT, this.arrayColSep); // \arraycolsep
    const delims: Record<string, [string, string]> = {
      pmatrix: ["(", ")"],
      bmatrix: ["[", "]"],
      Bmatrix: ["\\lbrace", "\\rbrace"],
      vmatrix: ["\\lvert", "\\rvert"],
      Vmatrix: ["\\lVert", "\\rVert"],
    };
    switch (env) {
      case "matrix":
      case "pmatrix":
      case "bmatrix":
      case "Bmatrix":
      case "vmatrix":
      case "Vmatrix": {
        // \env@matrix: \hskip-\arraycolsep \array{*\c@MaxMatrixCols c} ... \endarray\hskip-\arraycolsep
        const body: MItem[] = [this.hskip(-acs, a) as MItem, this.latexArray(a, "*{10}{c}", "1", ctx), this.hskip(-acs, a) as MItem];
        const d = delims[env];
        return d ? [this.leftRightWrap(a, d[0], d[1], body)] : body;
      }
      case "cases":
      case "dcases":
      case "rcases": {
        // \left\lbrace \def\arraystretch{1.2} \array{@{}l@{\quad}l@{}} ... \right.
        const arr = this.latexArray(a, "@{}l@{\\quad}l@{}", "1.2", ctx, env === "dcases" ? DISPLAY_STYLE : null);
        return [env === "rcases" ? this.leftRightWrap(a, ".", "\\rbrace", [arr]) : this.leftRightWrap(a, "\\lbrace", ".", [arr])];
      }
      case "array": {
        const spec = (a.colFormat ?? [])
          .map((c: any) => (c.align ? c.align : c.separator === "solid" ? "|" : c.separator === "dashed" ? "|" : ""))
          .join("");
        return [this.latexArray(a, spec || "c", "1", ctx)];
      }
      case "aligned":
      case "alignedat":
      case "split":
        return this.alignedEnv(a, ctx, env === "alignedat");
      case "gathered":
        return this.gatheredEnv(a, ctx);
      case "smallmatrix":
        return this.smallMatrixEnv(a, ctx);
      default: {
        const body: MItem[] = [this.hskip(-acs, a) as MItem, this.latexArray(a, "*{10}{c}", "1", ctx), this.hskip(-acs, a) as MItem];
        return body;
      }
    }
  }

  /** amsmath \strut@ box (\strutbox@ with \normallineskiplimit = 0). */
  private amsStrut(a: AtomLike): BoxNode {
    const [sh, sd] = this.strutDims();
    const strut = hpack([{ type: RULE_NODE, width: 0, height: sh, depth: sd, tag: a }], 0, undefined, a);
    const vb = vpack([newKern(0, a), strut], a);
    return hpack([vb], 0, undefined, a);
  }

  /** \spread@equation (\openup\jot) applied to LaTeX's 12/1/0pt (\jot = 3pt, times the user's row spacing). */
  private spreadParams(): VParams {
    const jot = scalePermille(3 * PT, this.alignRowSep);
    return { baselineskip: glue(12 * PT + jot), lineskip: glue(PT + jot), lineskiplimit: jot };
  }

  /** amsmath aligned / alignedat (\start@aligned). */
  private alignedEnv(a: AtomLike, ctx: Ctx, alignedat: boolean): MItem[] {
    const alignsep = alignedat ? 0 : 10 * PT; // \minalignsep
    const rows: AtomLike[][][] = a.rows ?? [];
    const items: (AlignRow | AlignNoalign)[] = [];
    rows.forEach((row, r) => {
      const { cells } = this.rowCells(row);
      if (this.dropsRow(cells, r, rows.length)) return;
      const lists: Node[][] = cells.map((cell, j) => {
        if (j % 2 === 0) {
          // &\column@plus \hfil \strut@ $\m@th\displaystyle{##}$ \tabskip\z@skip
          const f = this.cellFormula([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, this.groupNoad(this.list(cell, ctx), a)], a);
          return [this.hfil(a), this.amsStrut(a), ...f];
        }
        // &\column@plus $\m@th\displaystyle{{}##}$ \hfil \tabskip\alignsep@
        const empty = newNoad(ORD_NOAD, a);
        empty.nucleus = mlistField([]);
        const f = this.cellFormula([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, this.groupNoad([empty, ...this.list(cell, ctx)], a)], a);
        return [...f, this.hfil(a)];
      });
      // \math@cr@@@aligned: an even number of fields gets "&\kern-\alignsep@"
      if (!alignedat && r < rows.length - 1 && lists.length % 2 === 0) lists.push([newKern(-alignsep, a)]);
      items.push({ kind: "row", cells: lists, tag: a });
      const gap = this.rowGap(a, r);
      if (gap !== 0 && r < rows.length - 1) items.push({ kind: "noalign", apply: (v) => v.append(this.hskip(gap, a)) });
    });
    const tabskips: ReturnType<typeof glue>[] = [glue(0)];
    const maxCols = Math.max(1, ...items.map((it) => (it.kind === "row" ? it.cells.length : 0)));
    for (let j = 0; j < maxCols; j++) tabskips.push(j % 2 === 0 ? glue(0) : glue(alignsep));
    const v = new VList(this.spreadParams(), a);
    halign(items, tabskips, v, a);
    const box = vpack(v.items, a);
    box.role = "array";
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(box);
    return [this.nullBox(a), vc];
  }

  /** amsmath gathered */
  private gatheredEnv(a: AtomLike, ctx: Ctx): MItem[] {
    const rows: AtomLike[][][] = a.rows ?? [];
    const items: (AlignRow | AlignNoalign)[] = [];
    rows.forEach((row, r) => {
      const { cells } = this.rowCells(row);
      if (this.dropsRow(cells, r, rows.length)) return;
      const f = this.cellFormula([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, ...this.list(cells[0], ctx)], a);
      items.push({ kind: "row", cells: [[this.hfil(a), this.amsStrut(a), ...f, this.hfil(a)]], tag: a });
    });
    const v = new VList(this.spreadParams(), a);
    halign(items, [glue(0)], v, a);
    const box = vpack(v.items, a);
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(box);
    return [this.nullBox(a), vc];
  }

  private findAtom(atoms: AtomLike[] | undefined, pred: (x: AtomLike) => boolean): AtomLike | null {
    for (const x of atoms ?? []) {
      if (pred(x)) return x;
      for (const b of ["body", "above", "below"]) {
        const r = this.findAtom(x.branch?.(b), pred);
        if (r) return r;
      }
    }
    return null;
  }

  /** amsmath subarray (used by \\substack): script-style rows in a \\vcenter. */
  private subarray(a: AtomLike, ctx: Ctx, align: "c" | "l"): Noad {
    const sy = this.fams.get(2)![1]!; // \\scriptfont2
    const ex = this.fams.get(3)![1]!; // \\scriptfont3
    const bls = sy.param(10) + sy.param(12);
    const ls = 3 * ex.param(8);
    const rows: AtomLike[][][] = a.rows ?? [];
    const items: (AlignRow | AlignNoalign)[] = [];
    rows.forEach((row, r) => {
      const { cells } = this.rowCells(row);
      if (this.dropsRow(cells, r, rows.length)) return;
      const f = this.cellFormula([{ type: STYLE_NODE, style: SCRIPT_STYLE, tag: a }, ...this.list(cells[0], ctx)], a);
      items.push({ kind: "row", cells: [[...(align === "c" ? [this.hfil(a)] : []), ...f, this.hfil(a)]], tag: a });
    });
    const v = new VList({ baselineskip: glue(bls), lineskip: glue(ls), lineskiplimit: ls }, a);
    halign(items, [glue(0)], v, a);
    const box = vpack(v.items, a);
    box.role = "array";
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(box);
    return vc;
  }

  /** amsmath smallmatrix */
  private smallMatrixEnv(a: AtomLike, ctx: Ctx): MItem[] {
    const ex = PT; // \ex@ at 10pt
    const rows: AtomLike[][][] = a.rows ?? [];
    const items: (AlignRow | AlignNoalign)[] = [];
    const thick = texScale(".2777", this.textRoman().param(6));
    rows.forEach((row, r) => {
      const { cells } = this.rowCells(row);
      if (this.dropsRow(cells, r, rows.length)) return;
      const lists = cells.map((cell, j) => {
        const f = this.cellFormula([{ type: STYLE_NODE, style: SCRIPT_STYLE, tag: a }, ...this.list(cell, ctx)], a);
        return j === 0 ? [this.hfil(a), ...f, this.hfil(a)] : [newKern(thick, a), this.hfil(a), ...f, this.hfil(a)];
      });
      items.push({ kind: "row", cells: lists, tag: a });
    });
    const ls = texScale("1.5", ex);
    const v = new VList({ baselineskip: glue(6 * ex), lineskip: glue(ls), lineskiplimit: ls }, a);
    halign(items, [glue(0)], v, a);
    const box = vpack(v.items, a);
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(box);
    return [this.nullBox(a), this.muGlue(3 * PT, a), vc, this.muGlue(3 * PT, a)];
  }

  // --------------------------------------------------------- editor atoms --
  /** An empty slot: a box the size of a capital letter, drawn by the painter. */
  /**
   * A placeholder takes the place of an x, as in the VisualTeX kernel: the
   * formula is laid out as if an x were there, and the editor paints its
   * marker over that box instead of the glyph.
   */
  private placeholder(a: AtomLike, ctx: Ctx): Noad {
    return this.mathChar(this.mathCodeOf("x", ctx) ?? 0x7178, ctx, a);
  }

  private fallback(a: AtomLike, ctx: Ctx): MItem[] {
    // Unknown construct: render its body (if any) as a group, else its value.
    const body = a.branch("body");
    if (body) return [this.groupNoad(this.list(body, ctx), a)];
    if (a.value || a.command) return this.symbol(a, ctx);
    return [];
  }

  // -------------------------------------------------------------- helpers --
  /** \hbox{$\m@th ...$}: typeset a math list from a starting style. */
  mathBox(list: MItem[], style: number, tag: Tag, env?: MathEnv): BoxNode {
    const ts = new MathTypesetter(env ?? this.env);
    const h = ts.mlistToHlist(list, style, false);
    const on: Node = { type: MATH_NODE, width: 0, subtype: 0, tag };
    return hpack([on, ...h, { ...on }], 0, undefined, tag);
  }
}

export { HLIST_NODE, VLIST_NODE, SUB_BOX, SUB_MLIST, MATH_CHAR, UNDER_NOAD, OVER_NOAD, BIN_NOAD, REL_NOAD, xOverN, remainder, div };
export type { GlueSpec };
