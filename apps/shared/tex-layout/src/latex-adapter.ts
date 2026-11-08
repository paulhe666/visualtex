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
  A_LEADERS,
  FIL,
  FILL as FILL_ORDER,
  X_LEADERS,
  newCharacter,
  newKern,
  newNullBox,
  newRule,
  type BoxNode,
  type GlueSpec,
  type Node,
  type RuleNode,
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
  /** Greek letters and digits in fonts of their own, chosen apart from the letters' (the same four faces) */
  greekIt: 25,
  greekUp: 26,
  greekBf: 27,
  greekBfIt: 28,
  digitIt: 29,
  digitUp: 30,
  digitBf: 31,
  digitBfIt: 32,
  /** stmaryrd's symbol font (U/stmry: stmary10, 7, 5), when the profile's data has it; not a TeX family number */
  stmry: 33,
  /** upgreek's upright Greek (its default Euler option: U/eur, eurm10, 7, 5), when the data has it */
  eur: 34,
} as const;

/** Fonts of packages the reference preamble loads, set up as families when the profile's data has them. */
export const PACKAGE_FAMILIES: [number, [string, string, string]][] = [
  [FAM.stmry, ["stmary10", "stmary7", "stmary5"]],
  [FAM.eur, ["eurm10", "eurm7", "eurm5"]],
];

/** The families of a group of characters set in a chosen font: italic, upright, bold, bold italic. */
export const FACE_GROUP_FAMS = {
  letter: [FAM.letterIt, FAM.letterUp, FAM.letterBf, FAM.letterBfIt],
  greek: [FAM.greekIt, FAM.greekUp, FAM.greekBf, FAM.greekBfIt],
  digit: [FAM.digitIt, FAM.digitUp, FAM.digitBf, FAM.digitBfIt],
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
  /** inside bm's \boldmath box: characters with no bold family stay as they are (no poor man's bold) */
  boldmath?: boolean;
}

const BASE_CTX: Ctx = { fam: -1, bold: false, operator: false, display: false };

/** bm.sty's poor man's bold step: \mkern.4mu (mu glue widths are in sp of mu). */
const PMB_STEP = Math.round(0.4 * 65536);

/** amsmath arrow fills: [left piece, repeated piece, right piece]; a list of commands is a braced group */
type ArrowFill = [string | string[], string, string | string[]];
const FILL = {
  right: ["\\relbar", "\\relbar", "\\rightarrow"] as ArrowFill,
  left: ["\\leftarrow", "\\relbar", "\\relbar"] as ArrowFill,
  leftright: ["\\leftarrow", "\\relbar", "\\rightarrow"] as ArrowFill,
  Right: ["\\Relbar", "\\Relbar", "\\Rightarrow"] as ArrowFill,
  Left: ["\\Leftarrow", "\\Relbar", "\\Relbar"] as ArrowFill,
  LeftRight: ["\\Leftarrow", "\\Relbar", "\\Rightarrow"] as ArrowFill,
  // mathtools (\MT_..._fill:)
  rightharpoonup: ["\\relbar", "\\relbar", "\\rightharpoonup"] as ArrowFill,
  rightharpoondown: ["\\relbar", "\\relbar", "\\rightharpoondown"] as ArrowFill,
  leftharpoonup: ["\\leftharpoonup", "\\relbar", "\\relbar"] as ArrowFill,
  leftharpoondown: ["\\leftharpoondown", "\\relbar", "\\relbar"] as ArrowFill,
  longright: ["\\relbar", "\\relbar", "\\longrightarrow"] as ArrowFill,
  longleft: ["\\longleftarrow", "\\relbar", "\\relbar"] as ArrowFill,
  hookleft: ["\\leftarrow", "\\relbar", ["\\relbar", "\\joinrel", "\\rhook"]] as ArrowFill,
  hookright: [["\\lhook", "\\joinrel", "\\relbar"], "\\relbar", "\\rightarrow"] as ArrowFill,
  mapsto: [["\\mapstochar", "\\relbar"], "\\relbar", "\\rightarrow"] as ArrowFill,
  // extarrows (\x@arrow is \DOTSB\Relbar)
  equal: ["\\Relbar", "\\Relbar", "\\Relbar"] as ArrowFill,
  Longleftright: ["\\Longleftarrow", "\\Relbar", "\\Longrightarrow"] as ArrowFill,
  longleftright: ["\\longleftarrow", "\\relbar", "\\longrightarrow"] as ArrowFill,
  Longleft: ["\\Longleftarrow", "\\Relbar", "\\Relbar"] as ArrowFill,
  Longright: ["\\Relbar", "\\Relbar", "\\Longrightarrow"] as ArrowFill,
  // extpfeil
  twoheadright: ["\\relbar", "\\relbar", "\\twoheadrightarrow"] as ArrowFill,
  twoheadleft: ["\\twoheadleftarrow", "\\relbar", "\\relbar"] as ArrowFill,
};
const OVER_ARROWS: Record<string, { fill: ArrowFill; under: boolean }> = {
  "\\overrightarrow": { fill: FILL.right, under: false },
  "\\overleftarrow": { fill: FILL.left, under: false },
  "\\overleftrightarrow": { fill: FILL.leftright, under: false },
  "\\underrightarrow": { fill: FILL.right, under: true },
  "\\underleftarrow": { fill: FILL.left, under: true },
  "\\underleftrightarrow": { fill: FILL.leftright, under: true },
};
/**
 * \ext@arrow #1#2#3#4 (script kerns, measuring kerns) and the fill; `space`: mathtools passes the labels with a
 * control space before ("pre"), after ("post") or on both sides of them (\xRightarrow: {#1\ }{#2\ }).
 */
type ExtArrow = { k: [number, number, number, number]; fill: ArrowFill; space?: "pre" | "post" | "both" };
const EXT_ARROWS: Record<string, ExtArrow> = {
  "\\xrightarrow": { k: [0, 3, 5, 9], fill: FILL.right },
  "\\xleftarrow": { k: [3, 0, 9, 5], fill: FILL.left },
  // mathtools (loaded before extarrows, whose \providecommand then keeps these)
  "\\xleftrightarrow": { k: [3, 0, 9, 5], fill: FILL.leftright },
  "\\xRightarrow": { k: [0, 0, 5, 5], fill: FILL.Right, space: "post" },
  "\\xLeftarrow": { k: [0, 0, 5, 5], fill: FILL.Left, space: "pre" },
  "\\xLeftrightarrow": { k: [0, 0, 5, 5], fill: FILL.LeftRight, space: "both" },
  "\\xrightharpoondown": { k: [0, 3, 5, 9], fill: FILL.rightharpoondown },
  "\\xrightharpoonup": { k: [0, 3, 5, 9], fill: FILL.rightharpoonup },
  "\\xleftharpoondown": { k: [3, 0, 9, 5], fill: FILL.leftharpoondown },
  "\\xleftharpoonup": { k: [3, 0, 9, 5], fill: FILL.leftharpoonup },
  "\\xlongrightarrow": { k: [0, 3, 5, 9], fill: FILL.longright },
  "\\xlongleftarrow": { k: [3, 0, 9, 5], fill: FILL.longleft },
  "\\xhookleftarrow": { k: [3, 0, 9, 5], fill: FILL.hookleft },
  "\\xhookrightarrow": { k: [3, 0, 9, 5], fill: FILL.hookright },
  "\\xmapsto": { k: [0, 3, 9, 5], fill: FILL.mapsto },
  // extarrows
  "\\xlongequal": { k: [0, 0, 9, 9], fill: FILL.equal },
  "\\xLongleftrightarrow": { k: [0, 0, 9, 9], fill: FILL.Longleftright },
  "\\xlongleftrightarrow": { k: [0, 0, 9, 9], fill: FILL.longleftright },
  "\\xLongleftarrow": { k: [0, 0, 9, 9], fill: FILL.Longleft },
  "\\xLongrightarrow": { k: [0, 0, 9, 9], fill: FILL.Longright },
  // extpfeil
  "\\xtwoheadrightarrow": { k: [0, 5, 40, 0], fill: FILL.twoheadright },
  "\\xtwoheadleftarrow": { k: [5, 0, 0, 40], fill: FILL.twoheadleft },
};
/**
 * mathtools' harpoon pairs: \mathrel{\raise.22ex\hbox{$<upper>$}\setbox0=\hbox{$<lower>$}\kern-\wd0
 * \lower.22ex\box0}, each an \ext@arrow; the upper one has a \phantom of the label below, the lower one of the
 * label above.
 */
const EXT_HARPOONS: Record<string, [ExtArrow, ExtArrow]> = {
  "\\xrightleftharpoons": [EXT_ARROWS["\\xrightharpoonup"], EXT_ARROWS["\\xleftharpoondown"]],
  "\\xleftrightharpoons": [EXT_ARROWS["\\xleftharpoonup"], EXT_ARROWS["\\xrightharpoondown"]],
};

/**
 * stmaryrd.sty's symbols: \DeclareMathSymbol class and position in its font
 * (FAM.stmry), as \mathchar codes without the family. LaTeX's own definitions
 * (amssymb's \boxdot, the kernel's \bigtriangleup) stay as the app's export
 * loads stmaryrd only for the commands it alone defines.
 */
const STMARYRD: Record<string, number> = {
  "\\shortleftarrow": 0x300,
  "\\shortrightarrow": 0x301,
  "\\shortuparrow": 0x302,
  "\\shortdownarrow": 0x303,
  "\\Yup": 0x204,
  "\\Ydown": 0x205,
  "\\Yleft": 0x206,
  "\\Yright": 0x207,
  "\\varcurlyvee": 0x208,
  "\\varcurlywedge": 0x209,
  "\\minuso": 0x20a,
  "\\baro": 0x20b,
  "\\sslash": 0x20c,
  "\\bbslash": 0x20d,
  "\\moo": 0x20e,
  "\\varotimes": 0x20f,
  "\\varoast": 0x210,
  "\\varobar": 0x211,
  "\\varodot": 0x212,
  "\\varoslash": 0x213,
  "\\varobslash": 0x214,
  "\\varocircle": 0x215,
  "\\varoplus": 0x216,
  "\\varominus": 0x217,
  "\\boxast": 0x218,
  "\\boxbar": 0x219,
  "\\boxdot": 0x21a,
  "\\boxslash": 0x21b,
  "\\boxbslash": 0x21c,
  "\\boxcircle": 0x21d,
  "\\boxbox": 0x21e,
  "\\boxempty": 0x21f,
  "\\lightning": 0x020,
  "\\merge": 0x221,
  "\\vartimes": 0x222,
  "\\fatsemi": 0x223,
  "\\sswarrow": 0x324,
  "\\ssearrow": 0x325,
  "\\curlywedgeuparrow": 0x326,
  "\\curlywedgedownarrow": 0x327,
  "\\fatslash": 0x228,
  "\\fatbslash": 0x229,
  "\\lbag": 0x22a,
  "\\rbag": 0x22b,
  "\\varbigcirc": 0x22c,
  "\\leftrightarroweq": 0x32d,
  "\\curlyveedownarrow": 0x32e,
  "\\curlyveeuparrow": 0x32f,
  "\\nnwarrow": 0x330,
  "\\nnearrow": 0x331,
  "\\leftslice": 0x232,
  "\\rightslice": 0x233,
  "\\varolessthan": 0x234,
  "\\varogreaterthan": 0x235,
  "\\varovee": 0x236,
  "\\varowedge": 0x237,
  "\\talloblong": 0x238,
  "\\interleave": 0x239,
  "\\obar": 0x23a,
  "\\obslash": 0x23b,
  "\\olessthan": 0x23c,
  "\\ogreaterthan": 0x23d,
  "\\ovee": 0x23e,
  "\\owedge": 0x23f,
  "\\oblong": 0x240,
  "\\inplus": 0x341,
  "\\niplus": 0x342,
  "\\nplus": 0x243,
  "\\subsetplus": 0x344,
  "\\supsetplus": 0x345,
  "\\subsetpluseq": 0x346,
  "\\supsetpluseq": 0x347,
  "\\Lbag": 0x448,
  "\\Rbag": 0x549,
  "\\llparenthesis": 0x44c,
  "\\rrparenthesis": 0x54d,
  "\\binampersand": 0x44e,
  "\\bindnasrepma": 0x54f,
  "\\trianglelefteqslant": 0x350,
  "\\trianglerighteqslant": 0x351,
  "\\ntrianglelefteqslant": 0x352,
  "\\ntrianglerighteqslant": 0x353,
  "\\llfloor": 0x454,
  "\\rrfloor": 0x555,
  "\\llceil": 0x456,
  "\\rrceil": 0x557,
  "\\arrownot": 0x358,
  "\\Arrownot": 0x359,
  "\\Mapstochar": 0x35a,
  "\\mapsfromchar": 0x35b,
  "\\Mapsfromchar": 0x35c,
  "\\leftrightarrowtriangle": 0x35d,
  "\\leftarrowtriangle": 0x35e,
  "\\rightarrowtriangle": 0x35f,
  "\\bigtriangledown": 0x160,
  "\\bigtriangleup": 0x161,
  "\\bigcurlyvee": 0x162,
  "\\bigcurlywedge": 0x163,
  "\\bigsqcap": 0x164,
  "\\bigbox": 0x165,
  "\\bigparallel": 0x166,
  "\\biginterleave": 0x167,
  "\\bignplus": 0x170,
};
/** stmaryrd's \DeclareMathDelimiter: small and large positions in its font. */
const STMARYRD_DELIMS: Record<string, [number, number]> = { "\\llbracket": [0x4a, 0x71], "\\rrbracket": [0x4b, 0x79] };
/** stmaryrd's arrows made of two characters, as its \def\Mapsto{\Mapstochar\Rightarrow} etc. */
const STMARYRD_DEFS: Record<string, string[]> = {
  "\\Mapsto": ["\\Mapstochar", "\\Rightarrow"],
  "\\mapsfrom": ["\\leftarrow", "\\mapsfromchar"],
  "\\Mapsfrom": ["\\Leftarrow", "\\Mapsfromchar"],
  "\\Longmapsto": ["\\Mapstochar", "\\Longrightarrow"],
  "\\longmapsfrom": ["\\longleftarrow", "\\mapsfromchar"],
  "\\Longmapsfrom": ["\\Longleftarrow", "\\Mapsfromchar"],
};
/** upgreek.sty (its default Euler option): \mathord characters of Euler Roman (FAM.eur). */
const UPGREEK: Record<string, number> = {
  "\\upalpha": 0x0b,
  "\\upbeta": 0x0c,
  "\\upgamma": 0x0d,
  "\\updelta": 0x0e,
  "\\upepsilon": 0x0f,
  "\\upzeta": 0x10,
  "\\upeta": 0x11,
  "\\uptheta": 0x12,
  "\\upiota": 0x13,
  "\\upkappa": 0x14,
  "\\uplambda": 0x15,
  "\\upmu": 0x16,
  "\\upnu": 0x17,
  "\\upxi": 0x18,
  "\\uppi": 0x19,
  "\\uprho": 0x1a,
  "\\upsigma": 0x1b,
  "\\uptau": 0x1c,
  "\\upupsilon": 0x1d,
  "\\upphi": 0x1e,
  "\\upchi": 0x1f,
  "\\uppsi": 0x20,
  "\\upomega": 0x21,
  "\\upvarepsilon": 0x22,
  "\\upvartheta": 0x23,
  "\\upvarpi": 0x24,
  "\\upvarphi": 0x27,
  "\\Upgamma": 0x00,
  "\\Updelta": 0x01,
  "\\Uptheta": 0x02,
  "\\Uplambda": 0x03,
  "\\Upxi": 0x04,
  "\\Uppi": 0x05,
  "\\Upsigma": 0x06,
  "\\Upupsilon": 0x07,
  "\\Upphi": 0x08,
  "\\Uppsi": 0x09,
  "\\Upomega": 0x0a,
  "\\upvarrho": 0x1a,
  "\\upvarsigma": 0x1b,
};

/** ot1enc.def's accents: Unicode combining mark -> accent character (\`, \', \^, \~, \", \H, \r, \v, \u, \=, \., \c). */
const OT1_ACCENTS: Record<number, number> = {
  0x300: 0x12,
  0x301: 0x13,
  0x302: 0x5e,
  0x303: 0x7e,
  0x308: 0x7f,
  0x30b: 0x7d,
  0x30a: 0x17,
  0x30c: 0x14,
  0x306: 0x15,
  0x304: 0x16,
  0x307: 0x5f,
  0x327: 0x18,
};
/** ot1enc.def's letters: \ss, \ae, \oe, \o, \AE, \OE, \O, \i, \j. */
const OT1_LETTERS: Record<string, number> = { "ß": 0x19, "æ": 0x1a, "œ": 0x1b, "ø": 0x1c, "Æ": 0x1d, "Œ": 0x1e, "Ø": 0x1f, "ı": 0x10, "ȷ": 0x11 };

const PT = 65536;

/** amsmath's dots commands: what each one is before it looks at the next token */
const SMART_DOTS: Record<string, string> = {
  "\\dots": "dots",
  "\\dotsc": "dotsc",
  "\\dotso": "dotso",
  "\\cdots": "cdots",
  "\\dotsb": "cdots",
  "\\dotsm": "cdots",
  "\\dotsi": "dotsi",
};
/** amsmath \rightdelim@ */
const RIGHT_DELIM_TOKENS = new Set([")", "]", "\\rbrack", "\\}", "\\rbrace", "\\rangle", "\\rceil", "\\rfloor", "\\rgroup", "\\rmoustache", "\\right", "\\bigr", "\\biggr", "\\Bigr", "\\Biggr"]);

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
/** VisualTeX: amsmath's \tag / \tag* atom (kernel environments.ts). */
function isTagAtom(a: AtomLike): boolean {
  return a.command === "\\tag" || a.command === "\\tag*";
}

/** VisualTeX: amsmath's \intertext, mathtools' \shortintertext. */
function isIntertextAtom(a: AtomLike): boolean {
  return a.command === "\\intertext" || a.command === "\\shortintertext";
}

/**
 * Text without its leading spaces (TeX skips them at the start of a cell; \ignorespaces) and, with `end`, its
 * trailing ones (\unskip). A space before & or \\ stays in a cell, as in TeX.
 */
function trimText(atoms: AtomLike[], end = false): AtomLike[] {
  const space = (x: AtomLike) => x.type === "first" || (x.mode === "text" && (x.value ?? "").trim() === "");
  let i = 0;
  let j = atoms.length;
  while (i < j && space(atoms[i])) i++;
  while (end && j > i && space(atoms[j - 1])) j--;
  return [...atoms.slice(0, i).filter((x) => x.type === "first"), ...atoms.slice(i, j)];
}

/** A line of a display environment: a row (its cells, its tag, its source row) or an \intertext. */
type DisplayLine = { kind: "row"; cells: AtomLike[][]; tag: AtomLike | null; r: number } | { kind: "text"; atom: AtomLike };

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

/** array.sty's template pieces between the cells (\@mkpream) */
type ArrayPiece =
  | { k: "acol" } // \hskip\col@sep
  | { k: "rule"; pattern: string } // | (pattern "solid"), arydshln's : ("") and ;{dash/gap}
  | { k: "dbl" } // \hskip\doublerulesep between two rules
  | { k: "content"; src: string; index: number }; // @{...}, !{...}: the index-th of the specification's

/** One column's template: pieces, the cell's alignment (p, m, b: a paragraph `width` wide), >{...} and <{...} */
interface ArrayColumn {
  pre: ArrayPiece[];
  align: "l" | "c" | "r" | "p" | "m" | "b";
  width?: string;
  before: string[];
  after: string[];
  post: ArrayPiece[];
}

/** The material between an array's rows arydshln's dashed column rules cross or stop at */
interface ArrayMarks {
  connect: WeakSet<Node>;
  cut: WeakSet<Node>;
  rowsConnect: WeakSet<AlignRow>;
}

/** The array parameters a formula may set before an array */
interface ArrayParams {
  stretch?: string;
  colsep?: number;
  tabcolsep?: number;
  rulewidth?: number;
  doublerulesep?: number;
  extrarowheight?: number;
}

/** The kernel's cell wrappers (atoms/array.ts CellWrapper) */
type CellWrapperLike =
  | { command: "\\multicolumn"; columns: number; spec: string }
  | { command: "\\multirow"; rows: string; width: string; vpos?: string; bigstruts?: string; vmove?: string }
  | { command: "\\hdotsfor"; columns: number; spacing?: string };

/** A column specification's tokens with their arguments; *{n}{spec} expanded (array.sty's \NC@rewrite@*). */
function specTokens(spec: string): { ch: string; arg?: string }[] {
  const out: { ch: string; arg?: string }[] = [];
  const group = (i: number): [string, number] => {
    while (spec[i] === " ") i++;
    if (spec[i] !== "{") return [spec[i] ?? "", i];
    let depth = 0;
    for (let k = i; k < spec.length; k++) {
      if (spec[k] === "{") depth++;
      else if (spec[k] === "}" && --depth === 0) return [spec.slice(i + 1, k), k];
    }
    return [spec.slice(i + 1), spec.length];
  };
  for (let i = 0; i < spec.length; i++) {
    const ch = spec[i];
    if (ch === " ") continue;
    if (ch === "*") {
      const [n, i1] = group(i + 1);
      const [cols, i2] = group(i1 + 1);
      spec = spec.slice(0, i) + cols.repeat(Math.max(0, parseInt(n, 10) || 0)) + spec.slice(i2 + 1);
      i -= 1;
      continue;
    }
    if ("pmb@!<>;".includes(ch)) {
      const [arg, end] = group(i + 1);
      out.push({ ch, arg });
      i = end;
      continue;
    }
    out.push({ ch });
  }
  return out;
}

/** A tabular cell without its leading and trailing spaces (the template's \ignorespaces, \unskip). */
function trimSpaces(cell: AtomLike[]): AtomLike[] {
  const isSpace = (x: AtomLike) => x.type === "text" && /^\s*$/.test(x.value ?? "");
  let i = 0;
  let j = cell.length;
  while (i < j && (cell[i].type === "first" || isSpace(cell[i]))) i++;
  while (j > i && isSpace(cell[j - 1])) j--;
  return [...cell.slice(0, i).filter((x) => x.type === "first"), ...cell.slice(i, j)];
}

export class LatexAdapter {
  readonly data: LatexData;
  readonly cache: FontCache;
  readonly fams: FamilyTable;
  readonly env: MathEnv;
  /** the families of the math alphabets and symbol fonts */
  readonly role: FamilyRoles;
  /** families of the bold math version (ProfileData.boldmath) for bm's boxes */
  private readonly boldEnv: MathEnv;
  /** NFSS defaults of the profile (\text) */
  readonly text: ProfileData["text"] & { bold: string };
  native: NativeFonts | null = null;
  /**
   * The italic correction (em) of the character a \class{...} box stands for,
   * when it is the nucleus of \mathop (custom symbols drawn on a large
   * operator; LayoutOptions.classOperatorItalic).
   */
  classOperatorItalic: ((className: string) => number | null | undefined) | null = null;
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
  /** the array parameters set before arrays in the current group (\arraystretch, \arraycolsep ...) */
  private arrayParams: ArrayParams = {};
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
    // Computer Modern's \boldmath: operators cmbx, letters cmmib, symbols cmbsy (fontmath.ltx), the rest unchanged
    const cmBold = (): FamilyTable => new Map([...this.fams, [FAM.rm, this.fams.get(FAM.bf)!], [FAM.mi, this.fams.get(FAM.bm)!], [FAM.sy, this.fams.get(FAM.bmsy)!]]);
    this.boldEnv = makeEnv(boldmath ? new Map([...this.fams, ...familiesFrom(cache, boldmath)]) : cmBold(), LATEX_PARAMS);
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
    // VisualTeX: a display's own \tag goes to the right end of its line
    const tag = display ? atoms.find((x) => isTagAtom(x)) : undefined;
    if (tag) {
      out.push(this.taggedDisplay(atoms.filter((x) => !isTagAtom(x)), tag, { ...BASE_CTX, display }));
      return out;
    }
    out.push(...this.list(atoms, { ...BASE_CTX, display }));
    return out;
  }

  // ---------------------------------------------------------------- lists --
  /** A math list (the contents of a branch). */
  list(atoms: AtomLike[] | undefined, ctx: Ctx): MItem[] {
    // a group: array parameters set in it (\renewcommand{\arraystretch}...) end with it
    const arrayParams = this.arrayParams;
    const out: MItem[] = [];
    const visible = (atoms ?? []).filter((a) => a.type !== "first");
    // Background runs first (\colorbox), then style runs, as the kernel's
    // LaTeX serializer forms them.
    for (const run of this.styleRuns(visible, ctx, true)) {
      const bg = this.backgroundOf(run[0]);
      if (!bg) {
        this.colouredRuns(run, ctx, out);
        continue;
      }
      // xcolor \colorbox{c}{...}: its argument is set in text mode
      let content: BoxNode;
      if (run[0].mode === "text") content = this.textRunBox(run, 0);
      else {
        const items: MItem[] = [];
        this.colouredRuns(run, ctx, items);
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
    this.arrayParams = arrayParams;
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
      const bg = background ? this.backgroundOf(a) : null;
      const k = background ? (bg ? bg + "|" + (a.mode === "text" ? "text" : "math") : "") : this.styleKey(a, ctx);
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
    // A script carrier (the ^3 of c^3) styled like its base stays in the base's run, so the script attaches to the
    // base inside the style group, as the kernel serializes it (\textcolor{red}{\int_a^bc^3}); styled otherwise it
    // starts a run and attaches to the group before it (\textcolor{red}{c}^3).
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

  /**
   * Atoms of one background as the kernel serializes their styles: \textcolor groups (text and math together) around
   * the alphabet and weight groups, so a colour does not split a + b, and each script stays with its base.
   */
  private colouredRuns(atoms: AtomLike[], ctx: Ctx, out: MItem[]): void {
    for (const run of this.colourRuns(atoms, ctx)) {
      const first = run[0];
      const color = first.style?.color;
      // a colour its parent already has makes no group (the serializer writes no \textcolor for it)
      const inherited = !!color && color === first.parent?.style?.color;
      if (!color || inherited || ctx.operator || first.type === "latex" || first.mode === "latex") {
        for (const r of this.styleRuns(run, ctx, false)) this.styledRun(r, ctx, out, inherited);
        continue;
      }
      let rest = run;
      if (rest[0].type === "subsup") {
        // a script coloured apart from its base attaches to what is before it
        this.atom(rest[0], ctx, out);
        rest = rest.slice(1);
        if (rest.length === 0) continue;
      }
      const items: MItem[] = [];
      for (const r of this.styleRuns(rest, ctx, false)) this.styledRun(r, ctx, items, true);
      out.push(this.groupNoad(items, runTag(rest)));
    }
  }

  private colourRuns(atoms: AtomLike[], ctx: Ctx): AtomLike[][] {
    const runs: AtomLike[][] = [];
    let cur: AtomLike[] = [];
    let key = "";
    for (const a of atoms) {
      const k = ctx.operator ? "op" : a.style?.color ?? "";
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

  /** Convert a style run; wrap it in the group its style command makes (`colourDone`: inside its colour group). */
  private styledRun(run: AtomLike[], ctx: Ctx, out: MItem[], colourDone = false): void {
    const first = run[0];
    if (ctx.operator) {
      for (const a of run) this.atom(a, ctx, out);
      return;
    }
    if (first.type === "subsup") {
      // a script styled apart from its base attaches to what is before it; the rest of the run is styled as usual
      this.atom(first, ctx, out);
      if (run.length > 1) this.styledRun(run.slice(1), ctx, out, colourDone);
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
    const color = colourDone ? undefined : first.style?.color;
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
    // bm.sty sets each symbol of \bm{...} in its bold font with its own math
    // class, without a group: a bold run keeps the spacing of the same run in
    // normal weight (a + b stays a binary operation); the alphabet and colour
    // commands make a math group
    if (bold === "\\bm" && !alphabet && !color) {
      out.push(...items);
      return;
    }
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
      case "macro": {
        const items = this.macro(a, ctx);
        // \limits / \nolimits typed after a macro that ends in an operator (a custom symbol), as after \int
        const last = items[items.length - 1] as Noad | undefined;
        if (a.explicitSubsupPlacement && last?.type === OP_NOAD) {
          last.subtype = a.subsupPlacement === "over-under" ? LIMITS : a.subsupPlacement === "adjacent" ? NO_LIMITS : NORMAL;
        }
        out.push(...items);
        break;
      }
      case "macro-argument":
        out.push(...this.list(a.branch("body"), ctx));
        break;
      case "subsup":
        this.attachScripts(out, a, ctx, true);
        return;
      case "group":
        if (ctx.bold && !ctx.boldmath) out.push(this.bmBox(a, ctx, (c) => [this.groupNoad(this.list(a.branch("body"), c), a)]));
        else out.push(this.groupNoad(this.list(a.branch("body"), ctx), a));
        break;
      case "genfrac":
        // bm: \frac is a brace group ({\begingroup#1\endgroup\over#2})
        if (ctx.bold && !ctx.boldmath) out.push(this.bmBox(a, ctx, (c) => this.genfrac(a, c)));
        else out.push(...this.genfrac(a, ctx));
        break;
      case "surd":
        out.push(...this.surd(a, ctx));
        break;
      case "accent":
        out.push(this.accent(a, ctx));
        break;
      case "leftright":
        out.push(...this.leftRightOrFence(a, ctx));
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
      case "overlap":
        out.push(...this.overlap(a, ctx));
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
      case "text":
        // \renewcommand{\arraystretch}{1.5}, \setlength{\arraycolsep}{2pt} ...: kept by the kernel as an invisible atom
        if (!a.value && typeof a.verbatimLatex === "string" && this.arraySetting(a.verbatimLatex)) break;
        out.push(...this.fallback(a, ctx));
        break;
      case "composition":
        // text an input method is composing (pinyin letters): all of it, as
        // typed; the atom's command still holds the first letter
        if (a.value) out.push(...this.unknownSymbol(a, ctx));
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
      if (cmd === "\\mathop" && this.classOperatorItalic) {
        const body = (a.branch("body") ?? []).filter((x) => x.type !== "first");
        const className = body.length === 1 && /^\\(?:class|htmlClass)$/.test(body[0].command ?? "") ? this.classOf(body[0]) : null;
        const italic = className ? this.classOperatorItalic(className) : null;
        if (italic && italic > 0) n.italicEm = italic;
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
    if (cmd === "\\label" || cmd === "\\notag" || cmd === "\\nonumber") {
      // VisualTeX: they draw nothing (an empty glue keeps the atom's place for the caret)
      out.push(this.hskip(0, a) as MItem);
      return true;
    }
    if (isTagAtom(a)) {
      // a tag where no display line takes it (LaTeX: "\tag not allowed here"): after a \quad
      out.push(this.hskip(this.textRoman().param(6), a) as MItem, this.boxNoad(this.tagBox(a, ctx), a));
      return true;
    }
    if (isIntertextAtom(a)) {
      // an \intertext outside an alignment: its text in line
      out.push(this.boxNoad(this.textModeBox(a.branch("body"), ctx, a), a));
      return true;
    }
    if (cmd in SMART_DOTS) {
      out.push(...this.smartDots(a, cmd));
      return true;
    }
    if (cmd === "\\mbox" && a.branch("body") !== undefined) {
      // latex.ltx \mbox#1 = \leavevmode\hbox{#1}: text at the document size in every style
      out.push(this.boxNoad(this.textModeBox(a.branch("body"), ctx, a), a));
      return true;
    }
    if (this.packageCommand(a, ctx, out)) return true;
    if (cmd === "\\doubleprime" || cmd === "\\tripleprime" || cmd === "\\qprime") {
      const n = cmd === "\\doubleprime" ? 2 : cmd === "\\tripleprime" ? 3 : 4;
      for (let i = 0; i < n; i++) out.push(this.mathChar(0x0230, ctx, a));
      return true;
    }
    return false;
  }

  /**
   * amsmath's dots, which look at the token after them: \dots chooses \dotsc before a comma, \dotsb (\@cdots)
   * before a binary operator or relation (\keybin@, a \mathchar of class 2 or 3, a macro starting with \DOTSB,
   * \mathbin, \mathrel or \not), \dotsi before \DOTSI, else \dotso@; \cdots, \dotsc and \dotso add \, before
   * punctuation or a closing delimiter (\extrap@, \extra@).
   */
  private smartDots(a: AtomLike, cmd: string): MItem[] {
    const next = this.nextToken(a);
    const meaning = next.startsWith("\\") ? (this.meaning(next) ?? "").replace(/^\\long /, "") : "";
    const body = meaning.startsWith("macro:->") ? meaning.slice(8) : null;
    // \extra@: a closing delimiter, the end of the formula, or a macro starting with \DOTSX
    const extra = RIGHT_DELIM_TOKENS.has(next) || next === "$" || !!body?.startsWith("\\DOTSX");
    let kind: "l" | "c" | "i" = "l";
    let thin = false;
    let which = SMART_DOTS[cmd];
    if (which === "dots") {
      const mathchar = /^\\mathchar"([0-9A-F]+)$/.exec(meaning);
      const cls = mathchar ? Math.floor(parseInt(mathchar[1], 16) / 4096) : -1;
      if (next === ",") which = "dotsc";
      else if (next === "\\not" || "+=<>-*:".includes(next) && next.length === 1 || cls === 2 || cls === 3) which = "dotsb@";
      else if (body?.startsWith("\\not") || body?.startsWith("\\DOTSB") || /^\\math(?:bin|rel)\b/.test(body ?? "")) which = "dotsb@";
      else if (body?.startsWith("\\DOTSI")) which = "dotsi";
      else which = "dotso@";
    }
    if (which === "dotsb@") kind = "c";
    else if (which === "dotsi") kind = "i";
    else if (which === "dotso@") thin = extra;
    else if (which === "dotsc") thin = next === ";" || next === "." || extra;
    else {
      // \extrap@: \cdots (\dotsb, \dotsm) and \dotso
      kind = which === "cdots" ? "c" : "l";
      thin = next === "," || next === ";" || next === "." || extra;
    }
    const items = this.template(kind === "l" ? "\\ldots" : kind === "c" ? "\\cdots" : "\\dotsi", a) ?? [];
    if (thin) items.push(...(this.template("\\,", a) ?? []));
    return items;
  }

  /** The token TeX meets after an atom: "$" at the end of the formula, "}" at the end of another group. */
  private nextToken(a: AtomLike): string {
    if ((a.branch("superscript") ?? []).length > 0 || (a.branch("subscript") ?? []).length > 0) return "^";
    let next: AtomLike | undefined;
    try {
      next = (a as { rightSibling?: AtomLike }).rightSibling;
    } catch {
      next = undefined; // an atom outside a kernel tree
    }
    if (!next) {
      const parent = a.parent as AtomLike | undefined;
      if (parent?.type === "root") return "$";
      if (parent?.type === "leftright" && a.parentBranch === "body") return "\\right";
      return "}";
    }
    if (next.type === "subsup") return (next.branch("superscript") ?? []).length > 0 ? "^" : "_";
    return next.command ?? next.value ?? "";
  }

  /** The class name of a \class / \htmlClass atom. */
  private classOf(a: AtomLike): string | null {
    const c = (a as { args?: unknown[] }).args?.[0];
    return typeof c === "string" && c ? c : null;
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
    if (cmd === "\\mod") return this.mod(a, ctx);
    if (cmd === "\\sideset") {
      const items = this.sideset(a, ctx);
      if (items) return items;
    }
    if (cmd === "\\vcentcolon") {
      // mathtools \mathrel{\mathop\ordinarycolon}: the character itself is the operator's nucleus (centred on the axis)
      const op = newNoad(OP_NOAD, a);
      op.nucleus = this.mathChar(this.mathCodeOf(":", ctx) ?? 0x303a, ctx, a).nucleus;
      const rel = newNoad(REL_NOAD, a);
      rel.nucleus = mlistField([op]);
      return [rel];
    }
    if (cmd === "\\substack") {
      const arr = this.findAtom(a.branch("body"), (x) => x.type === "array");
      if (arr) return [this.subarray(arr, ctx, "c")];
    }
    // a package command read with its arguments (physics' \cos[2](x), ...) is its expansion
    const tpl = a.argumentTemplate === undefined ? this.template(cmd, a) : null;
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

  /**
   * amsmath \sideset#1#2#3: box0 = #3 measured in display style; box2 an empty \vbox of box0's height and depth;
   * box4 = {\copy2 #1} (the scripts before the operator), box6 = {#3\nolimits#2} (the operator with the scripts
   * after it); then \hbox to\dimen@{}\mathop{\kern-\dimen@\box4\box6} with \dimen@ = \wd4 + \wd0 - \wd6.
   * The operator keeps its limits: scripts after \sideset{..}{..}\sum attach to that \mathop.
   */
  private sideset(a: AtomLike, ctx: Ctx): MItem[] | null {
    const args = new Map<number, AtomLike>();
    for (const x of a.branch("body") ?? []) if (x.type === "macro-argument" && !args.has(x.argumentIndex)) args.set(x.argumentIndex, x);
    if (!args.has(1) || !args.has(2) || !args.has(3)) return null;
    const display = (items: MItem[]) => this.mathBox([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, this.groupNoad(items, a)], TEXT_STYLE, a);
    const operator = () => this.list(args.get(3)!.branch("body"), ctx);
    const withScripts = (base: MItem[], scripts: AtomLike) => {
      for (const x of (scripts.branch("body") ?? []).filter((y) => y.type !== "first")) this.atom(x, ctx, base);
      return base;
    };
    const box0 = display(operator());
    const box2 = newNullBox(a);
    box2.type = VLIST_NODE;
    box2.height = box0.height;
    box2.depth = box0.depth;
    const box4 = display(withScripts([this.boxNoad(box2, a)], args.get(1)!));
    const op = operator();
    const last = op[op.length - 1] as Noad | undefined;
    if (last?.type === OP_NOAD) last.subtype = NO_LIMITS; // \nolimits
    const box6 = display(withScripts(op, args.get(2)!));
    const dimen = box4.width + box0.width - box6.width;
    const spacer = newNullBox(a);
    spacer.width = dimen;
    const result = newNoad(OP_NOAD, a);
    result.nucleus = mlistField([{ type: KERN_NODE, width: -dimen, subtype: 1, tag: a } as Node, this.boxNoad(box4, a), this.boxNoad(box6, a)]);
    return [this.boxNoad(spacer, a), result];
  }

  /** amsmath \mod: \allowbreak\mkern18mu or 12mu {\operator@font mod}\,\,#1 */
  private mod(a: AtomLike, ctx: Ctx): MItem[] {
    const arg = (a.branch("body") ?? []).find((x) => x.type === "macro-argument");
    const out: MItem[] = [{ type: 12, penalty: 0, tag: a } as Node]; // \allowbreak
    out.push(this.muKern((ctx.display ? 18 : 12) * PT, a));
    const letters = [..."mod"].map((ch) => this.mathChar(this.data.meanings.mathcode[ch], { ...ctx, fam: this.role.rm }, a));
    out.push(this.groupNoad(letters, a));
    for (let i = 0; i < 2; i++) out.push(...(this.template("\\,", a) ?? []));
    out.push(...this.list(arg?.branch("body"), ctx));
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
      if (b === fam && !ctx.boldmath && this.data.profile?.boldmath?.[fam]) return this.boldmathChar(cls, fam, ch, tag);
      if (b === fam && !ctx.boldmath) return this.poorMansBoldChar(cls, fam, ch, tag);
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
      return [this.boxNoad(this.mathBox([{ type: STYLE_NODE, style, tag }, n], TEXT_STYLE, tag, this.boldEnv), tag)];
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
   * bm.sty for a character that has no bold font at all, not even in
   * \boldmath (the large symbols, the AMS symbols): poor man's bold
   * (\bm@pmb@@): the character three times, the second raised, each .4mu right
   * of the one before, in a \mathchoice so that each style sets its own size
   * (a large operator keeps its display size). The class stays: an operator
   * keeps its limits, a relation its spacing.
   */
  private poorMansBoldChar(cls: number, fam: number, ch: number, tag: Tag): Noad {
    const box = (style: number): MItem[] => {
      const k = this.mathBox([{ type: STYLE_NODE, style, tag }, this.muKern(PMB_STEP, tag)], TEXT_STYLE, tag).width;
      const copy = () => {
        const n = newNoad(ORD_NOAD + cls, tag);
        n.nucleus = charField(fam, ch);
        return this.mathBox([{ type: STYLE_NODE, style, tag }, n], TEXT_STYLE, tag);
      };
      const first = copy();
      const raised = copy();
      raised.shift = -Math.round(1.5 * k);
      const w = first.width;
      return [this.boxNoad(hpack([first, newKern(k - w, tag), raised, newKern(k - w, tag), copy()], 0, undefined, tag), tag)];
    };
    const choice: ChoiceItem = {
      type: CHOICE_NODE,
      display: box(DISPLAY_STYLE),
      text: box(TEXT_STYLE),
      script: box(SCRIPT_STYLE),
      scriptScript: box(SCRIPT_SCRIPT_STYLE),
      tag,
    };
    const n = newNoad(ORD_NOAD + (cls === 7 ? 0 : cls), tag);
    n.nucleus = mlistField([choice as unknown as MItem]);
    return n;
  }

  /**
   * A letter or digit in the editor's letter font, when one is chosen (each
   * layout fills the letter families): that family and the Unicode
   * character. Everything else keeps the TeX fonts.
   */
  private letterField(fam: number, ch: number): Field | null {
    const r = this.role;
    // italic, upright, bold, bold italic (FACE_GROUP_FAMS order)
    const face = fam === r.mi || fam === r.it ? 0 : fam === r.rm ? 1 : fam === r.bf ? 2 : fam === r.bm ? 3 : -1;
    if (face < 0) return null;
    const mathItalic = fam === r.mi || fam === r.bm;
    let c: string | undefined;
    let group: keyof typeof FACE_GROUP_FAMS = "letter";
    if ((ch >= 0x41 && ch <= 0x5a) || (ch >= 0x61 && ch <= 0x7a)) c = String.fromCharCode(ch);
    // the math italic fonts' digits are old-style figures
    else if (ch >= 0x30 && ch <= 0x39) {
      c = mathItalic ? undefined : String.fromCharCode(ch);
      group = "digit";
    } else if (ch <= 0x0a) {
      c = UPPER_GREEK[ch];
      group = "greek";
    } else if (mathItalic && ch <= 0x27) {
      c = LOWER_GREEK[ch - 0x0b];
      group = "greek";
    }
    const to = (FACE_GROUP_FAMS[group] as readonly number[])[face];
    if (c === undefined || !this.fams.has(to)) return null;
    return charField(to, c.codePointAt(0)!);
  }

  /**
   * bm.sty's group in \bm{...} (\bm@mchoice): \mathord{\mathchoice{\hbox{\boldmath$\m@th\displaystyle ...$}}...}, the
   * group set in the bold math version in each style (its fractions take their parameters from the bold symbol font).
   */
  private bmBox(a: AtomLike, ctx: Ctx, make: (ctx: Ctx) => MItem[]): Noad {
    const inner: Ctx = { ...ctx, boldmath: true };
    const choice = this.palette(a, (st) => [this.boxNoad(this.mathBox([{ type: STYLE_NODE, style: st, tag: a }, ...make(inner)], TEXT_STYLE, a, this.boldEnv), a)]);
    const n = newNoad(ORD_NOAD, a);
    n.nucleus = mlistField([choice]);
    return n;
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

  /**
   * A symbol LaTeX defines as a macro (\sum is \DOTSB\sum@\slimits@) set inside
   * \bm: bm.sty reaches the math characters it expands to, each in its bold
   * family, or in poor man's bold when it has none; the class and the limits stay.
   */
  private emboldened(items: MItem[] | null, ctx: Ctx): MItem[] | null {
    if (!items || !ctx.bold) return items;
    return items.map((item) => {
      const n = item as Noad;
      if (!(n.type >= ORD_NOAD && n.type < LEFT_NOAD) || n.nucleus?.mathType !== MATH_CHAR) return item;
      const { fam, char } = n.nucleus as { fam: number; char: number };
      const bold = this.boldFamily(fam);
      if (bold !== fam) {
        n.nucleus = this.letterField(bold, char) ?? charField(bold, char);
        return item;
      }
      if (ctx.boldmath) return item;
      const pmb = this.poorMansBoldChar(n.type - ORD_NOAD, fam, char, n.tag);
      pmb.subtype = n.subtype;
      return pmb as MItem;
    });
  }

  private template(key: string, tag: Tag): MItem[] | null {
    const t = this.data.templates[key];
    if (!t) return null;
    try {
      return this.json.convert(t.mathlist, t.fontids, tag);
    } catch (e) {
      // a recorded list the converter cannot rebuild: the caller's fallback draws the command, the formula stays
      console.warn("tex-layout: template " + key + " unusable:", e);
      return null;
    }
  }

  /** The recorded list of a one-argument command applied to x (latex-onearg.json). */
  private template1(key: string, tag: Tag): MItem[] | null {
    const t = this.data.onearg[key];
    if (!t) return null;
    try {
      return this.json.convert(t.mathlist, t.fontids, tag);
    } catch (e) {
      console.warn("tex-layout: template " + key + "{x} unusable:", e);
      return null;
    }
  }

  /**
   * A one-argument template with its argument in place: the one plain x (math italic, no scripts) in the list is
   * replaced by the argument's noads, as TeX substitutes #1. Null when the list has no such x or more than one.
   */
  private withArgument(items: MItem[] | null, arg: () => MItem[]): MItem[] | null {
    if (!items) return null;
    const spots: { list: MItem[]; i: number }[] = [];
    const isX = (n: Noad) =>
      n.type === ORD_NOAD && n.nucleus?.mathType === MATH_CHAR && n.nucleus.char === 0x78 && n.sub.mathType === EMPTY && n.sup.mathType === EMPTY;
    const walk = (list: MItem[]) => {
      list.forEach((item, i) => {
        const n = item as Noad;
        if (!(n.type >= ORD_NOAD && n.type < LEFT_NOAD)) return;
        if (isX(n)) spots.push({ list, i });
        for (const f of [n.nucleus, n.sup, n.sub, n.num, n.den]) if (f?.mathType === SUB_MLIST && f.list) walk(f.list);
      });
    };
    walk(items);
    if (spots.length !== 1) return null;
    const { list, i } = spots[0];
    list.splice(i, 1, ...arg());
    return items;
  }

  private symbol(a: AtomLike, ctx: Ctx): MItem[] {
    const keys: string[] = [];
    if (a.command) keys.push(a.command);
    if (a.value && a.value !== a.command) keys.push(a.value);
    for (const key of keys) {
      const mc = this.mathCodeOf(key, ctx);
      if (mc !== null) return [this.mathChar(mc, ctx, a)];
      const tpl = this.emboldened(this.template(key, a), ctx);
      if (tpl) return tpl;
    }
    const pkg = a.command ? this.packageSymbol(a.command, ctx, a) : null;
    if (pkg) return pkg;
    return this.unknownSymbol(a, ctx);
  }

  /**
   * A symbol of a package the reference preamble loads and LaTeX's own data
   * does not define (stmaryrd, upgreek), set from the package's font when the
   * profile has it (PACKAGE_FAMILIES).
   */
  private packageSymbol(cmd: string, ctx: Ctx, tag: Tag): MItem[] | null {
    const noad = (cls: number, fam: number, code: number): Noad => {
      const n = newNoad(ORD_NOAD + cls, tag);
      n.nucleus = charField(fam, code);
      return n;
    };
    if (UPGREEK[cmd] !== undefined && this.fams.has(FAM.eur)) return [noad(0, FAM.eur, UPGREEK[cmd])];
    if (!this.fams.has(FAM.stmry)) return null;
    const mc = STMARYRD[cmd] ?? (STMARYRD_DELIMS[cmd] ? (cmd === "\\llbracket" ? 0x400 : 0x500) + STMARYRD_DELIMS[cmd][0] : undefined);
    if (mc !== undefined) return [noad(mc >> 8, FAM.stmry, mc & 0xff)];
    const parts = STMARYRD_DEFS[cmd];
    if (!parts) return null;
    const out: MItem[] = [];
    for (const part of parts) {
      const code = this.mathCodeOf(part, ctx);
      if (code !== null) out.push(this.mathChar(code, ctx, tag));
      else if (STMARYRD[part] !== undefined) out.push(noad(STMARYRD[part] >> 8, FAM.stmry, STMARYRD[part] & 0xff));
      else out.push(...(this.template(part, tag) ?? []));
    }
    return out;
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
    // \above <dimen>, \genfrac's thickness argument
    if (a.barThickness) frac.thickness = this.dimenFromAtom({ type: "spacing", width: a.barThickness, branch: () => undefined });
    const ld = a.leftDelim, rd = a.rightDelim;
    if (ld || rd) {
      frac.leftDelim = this.delimiter(ld ?? ".");
      frac.rightDelim = this.delimiter(rd ?? ".");
    }
    if (cmd === "\\frac" || cmd === "\\over" || cmd === "\\atop" || cmd === "\\choose" || /^\\(?:above|abovewithdelims|overwithdelims|atopwithdelims)$/.test(cmd)) {
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
          : a.mathstyleName === "scriptstyle"
            ? SCRIPT_STYLE
            : a.mathstyleName === "scriptscriptstyle"
              ? SCRIPT_SCRIPT_STYLE
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
    // \cfrac[l], [r]: the kernel's align is "left" or "right"
    const align = a.align === "left" ? "l" : a.align === "right" ? "r" : a.align;
    if (align === "r") numList.push({ type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 2, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a } as Node);
    numList.push(...this.list(a.branch("above"), ctx));
    if (align === "l") numList.push({ type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 2, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a } as Node);
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
    // amsmath \leftroot and \uproot in the index set the counts \leftroot@ and \uproot@
    let leftroot = 0, uproot = 0;
    const indexList = (index ?? []).filter((x) => {
      if (x.command !== "\\leftroot" && x.command !== "\\uproot") return true;
      const v = (x.args ?? [])[0] as { number?: number; dimension?: number } | undefined;
      const n = Math.trunc(Number(v?.number ?? v?.dimension ?? 0)) || 0;
      if (x.command === "\\leftroot") leftroot = n;
      else uproot = n;
      return false;
    });
    const visibleIndex = indexList.filter((x) => x.type !== "first");
    if (visibleIndex.length === 0) return [radical(ctx)];
    // amsmath \root#1\of: \setbox\rootbox\hbox{$\m@th\scriptscriptstyle{#1}$}
    // then \mathpalette\r@@t{#2}
    const rootList: MItem[] = [{ type: STYLE_NODE, style: SCRIPT_SCRIPT_STYLE, tag: a }, this.groupNoad(this.list(indexList, ctx), a)];
    const rootbox = this.mathBox(rootList, TEXT_STYLE, a);
    const variant = (style: number): MItem[] => {
      // \setboxz@h{$\m@th#1\sqrtsign{#2}$}
      const box0 = this.mathBox([{ type: STYLE_NODE, style, tag: a }, radical(ctx)], TEXT_STYLE, a);
      let dimen = box0.height - box0.depth;
      // \setbox\@ne\hbox{$\m@th#1\mskip\uproot@ mu$}; \advance\dimen@ by1.667\wd\@ne
      const box1 = this.mathBox([{ type: STYLE_NODE, style, tag: a }, this.muGlue(uproot * PT, a)], TEXT_STYLE, a);
      dimen += texScale("1.667", box1.width);
      const raised: BoxNode = { ...rootbox, list: rootbox.list, shift: -texScale(".6", dimen) };
      const rn = newNoad(ORD_NOAD, a);
      rn.nucleus = boxField(raised);
      const sq = newNoad(ORD_NOAD, a);
      sq.nucleus = boxField(box0);
      // \mkern-\leftroot@ mu\mkern5mu\raise.6\dimen@\copy\rootbox\mkern-10mu\mkern\leftroot@ mu\box\z@
      return [this.muKern(-leftroot * PT, a), this.muKern(5 * PT, a), rn, this.muKern(-10 * PT, a), this.muKern(leftroot * PT, a), sq];
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
      // amsmath \dddot, \ddddot: LaTeX's list for \cmd{x}, with the argument in place of the x
      const items = tpl ? this.withArgument(this.template1(cmd, a), () => this.list(a.branch("body"), ctx)) : null;
      if (items) return this.groupNoad(items, a);
      // Not a TeX math accent (e.g. an over-arrow or brace): defer to fallback.
      return this.groupNoad(this.fallback(a, ctx), a);
    }
    const n = newNoad(ACCENT_NOAD, a);
    n.accentChr = charField(acc.fam, acc.code);
    // amsfonts \widehat, \widetilde: \@mathmeasure\z@\textstyle{#1}, then the AMS accent when \wd\z@ > 2em
    const wide = /^macro:#1->\\@mathmeasure \\z@ \\textstyle \{#1\}\\ifdim \\wd \\z@ >\\tw@ em\\mathaccent "([0-9A-F]+)\{#1\}\\else /.exec(this.meaning(cmd) ?? "");
    if (wide) {
      const box0 = this.mathBox([{ type: STYLE_NODE, style: TEXT_STYLE, tag: a }, this.groupNoad(this.list(a.branch("body"), ctx), a)], TEXT_STYLE, a);
      if (box0.width > 2 * this.textRoman().param(6)) {
        const code = parseInt(wide[1], 16);
        n.accentChr = charField((code >> 8) & 15, code & 255);
      }
    }
    n.nucleus = this.field(a.branch("body"), ctx);
    return n;
  }

  // ----------------------------------------------------------- delimiters --
  /** Delimiter code for "(", "\\langle", "." etc. (\delcode / \delimiter). */
  delimiter(name: string): Delimiter {
    if (!name || name === "." || name === "\\.") return NULL_DELIMITER;
    // stmaryrd's \DeclareMathDelimiter (\llbracket, \rrbracket), unless LaTeX's data defines them
    const stmry = STMARYRD_DELIMS[name];
    if (stmry && this.fams.has(FAM.stmry) && !this.meaning(name))
      return { smallFam: FAM.stmry, smallChar: stmry[0], largeFam: FAM.stmry, largeChar: stmry[1] };
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
      const sy = this.role.sy, ex = this.role.ex;
      // fontmath.ltx delimiters the recorded meanings lack: symbols and largesymbols characters
      const fontmath: Record<string, number> = {
        "\\arrowvert": (sy << 20) | (0x6a << 12) | (ex << 8) | 0x3c,
        "\\Arrowvert": (sy << 20) | (0x6b << 12) | (ex << 8) | 0x3d,
        "\\bracevert": (ex << 20) | (0x3e << 12) | (ex << 8) | 0x3e,
      };
      if (dc !== null) code = dc & 0xffffff;
      else if (!m && fontmath[key] !== undefined) code = fontmath[key];
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
  /**
   * The kernel reads `\exp(x)` as a smart fence (a leftright without \left/\right) and writes it back as `\exp(x)`:
   * in LaTeX those are the plain characters, an open and a close noad at their natural size.
   */
  private leftRightOrFence(a: AtomLike, ctx: Ctx): MItem[] {
    const plain = a.variant === "" && a.leftDelim && a.leftDelim !== "?" && a.rightDelim && a.rightDelim !== "?";
    const open = plain ? this.mathCodeOf(a.leftDelim, ctx) : null;
    const close = plain ? this.mathCodeOf(a.rightDelim, ctx) : null;
    if (open === null || close === null) return [this.leftRight(a, ctx)];
    return [this.mathChar(open, ctx, a), ...this.list(a.branch("body"), ctx), this.mathChar(close, ctx, a)];
  }

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
      else items = this.emboldened(this.template(cmd, a), ctx);
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
    let tokens: TextToken[] = [];
    let lastSpace = false;
    // a tabular in the text (\text{\begin{tabular}...}) is a box between its words
    const parts: (TextToken[] | BoxNode)[] = [tokens];
    for (const a of run) {
      if (a.type === "array") {
        parts.push(this.tabularBox(a), (tokens = []));
        lastSpace = false;
        continue;
      }
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
    const nodes = parts.flatMap((part) =>
      Array.isArray(part) ? typesetText(letters ? part : this.encodeText(part), f, 1000, this.native?.fallback(f.size) ?? null) : [part],
    );
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

  /**
   * Accented and special Latin letters in the OT1 encoding, as LaTeX's
   * ot1enc.def sets them: é is \'e (\accent19 e, on a dotless i for í), ß is
   * \ss (\char25). Letters the encoding has neither way keep the fallback font.
   */
  private encodeText(tokens: TextToken[]): TextToken[] {
    if (this.text.encoding !== "OT1") return tokens;
    return tokens.map((t) => {
      if (t.kind !== "char" || t.code < 0x80) return t;
      const ch = String.fromCodePoint(t.code);
      const letter = OT1_LETTERS[ch];
      if (letter !== undefined) return { kind: "char", code: letter, tag: t.tag };
      const parts = [...ch.normalize("NFD")];
      if (parts.length !== 2 || parts[0].codePointAt(0)! >= 0x80) return t;
      const accent = OT1_ACCENTS[parts[1].codePointAt(0)!];
      if (accent === undefined) return t;
      const base = parts[0] === "i" ? 0x10 : parts[0] === "j" ? 0x11 : parts[0].codePointAt(0)!;
      return { kind: "accent", accent, code: base, composed: t.code, below: accent === 0x18, tag: t.tag };
    });
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
      // the kernel resolves xcolor names and mixes (frameColorValue, backgroundColorValue)
      const [block, back, noad] = this.colorBox(this.textModeBox(a.branch("body"), ctx, a), a.backgroundColorValue ?? color(a.backgroundcolor) ?? "yellow", a);
      const inner = hpack([block as unknown as Node, back as unknown as Node, (noad as Noad).nucleus.box!], 0, undefined, a);
      return [this.boxNoad(this.frameBox(inner, 0, a.frameColorValue ?? color(a.framecolor), a), a)];
    }
    return this.fallback(a, ctx);
  }

  /**
   * latex.ltx \llap, \rlap, \clap: \hb@xt@\z@{\hss#1}, {#1\hss}, {\hss#1\hss}; mathtools \mathllap etc.:
   * {}\llap{$\m@th<style>{#1}$} in the current style. The kernel reads the argument of \llap in math mode, as
   * MathJax does and the kernel's mhchem relies on: that math keeps the current style (MathJax's \llap); text
   * in it is LaTeX's box of text.
   */
  private overlap(a: AtomLike, ctx: Ctx): MItem[] {
    const cmd = a.command ?? "";
    const align: Record<string, string> = {
      "\\llap": "l", "\\rlap": "r", "\\clap": "c", "\\mathllap": "l", "\\mathrlap": "r", "\\mathclap": "c",
    };
    const side = align[cmd];
    if (!side) return this.fallback(a, ctx);
    const hss = (): Node => ({ type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 1, shrink: PT, shrinkOrder: 1 }, subtype: NORMAL, tag: a });
    const lap = (content: BoxNode): BoxNode =>
      hpack([...(side !== "r" ? [hss()] : []), content, ...(side !== "l" ? [hss()] : [])], 0, EXACTLY, a);
    const body = a.branch("body");
    const inStyle = (): MItem => this.palette(a, (st) => [this.boxNoad(lap(this.mathBox(this.styleList(st, a, [this.groupNoad(this.list(body, ctx), a)]), TEXT_STYLE, a)), a)]);
    if (!cmd.startsWith("\\math")) {
      const text = (body ?? []).some((x) => x.mode === "text");
      return [text ? this.boxNoad(lap(this.textModeBox(body, ctx, a)), a) : inStyle()];
    }
    const empty = newNoad(ORD_NOAD, a);
    empty.nucleus = mlistField([]);
    return [empty, inStyle()];
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
    for (const styled of this.styleRuns(visible, ctx, true)) {
      // VisualTeX: a style run may hold text and $...$ math (\tag{if $x$})
      const runs: AtomLike[][] = [];
      for (const x of styled) {
        const last = runs[runs.length - 1];
        if (last && (last[0].mode === "text") === (x.mode === "text")) last.push(x);
        else runs.push([x]);
      }
      for (const run of runs) {
        if (run[0].mode === "text") parts.push(this.textRunBox(run, 0));
        else {
          const items: MItem[] = [];
          this.colouredRuns(run, ctx, items);
          parts.push(this.mathBox(items, TEXT_STYLE, run[0]));
        }
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
    const harpoons = EXT_HARPOONS[cmd];
    if (harpoons) return [this.extHarpoons(a, ctx, harpoons)];
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
    return this.pictureSlineState(xarg, yarg, linelen, a).out;
  }

  /** \@sline with the state \@svector goes on with: \@clnht, \wd\@linechar, \@upordown. */
  private pictureSlineState(xarg: number, yarg: number, linelen: number, a: AtomLike): { out: Node[]; clnht: number; wd: number; raise: boolean } {
    const f = this.cache.get("line10");
    const out: Node[] = [];
    const hskip = (w: number): Node => ({ type: GLUE_NODE, spec: { width: w, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a });
    const negarg = xarg < 0;
    const yyarg = negarg ? -yarg : yarg;
    if (negarg) xarg = -xarg;
    // \@getlinechar(x,y)
    const c = xarg * 8 - 9 + (yyarg > 0 ? yyarg : -yyarg + 64);
    const info = f.charInfo(c);
    if (!info.exists) return { out, clnht: 0, wd: 0, raise: yarg > 0 };
    const wd = f.charWidth(info), ht = f.charHeight(info);
    if (wd === 0) return { out, clnht: 0, wd: 0, raise: yarg > 0 };
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
    return { out, clnht, wd, raise };
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
      const mc = this.mathCodeOf(name, BASE_CTX);
      if (mc !== null) return [this.mathChar(mc, BASE_CTX, a)];
      // macros (\joinrel, \longrightarrow, ...): LaTeX's math list
      return this.template(name, a) ?? [];
    };
    const fillPiece = (p: string | string[]): MItem[] =>
      typeof p === "string" ? piece(p) : [this.groupNoad(p.flatMap((name) => piece(name)), a)];
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
      ...fillPiece(kind[0]),
      this.muKern(-7 * PT, a),
      cleaders as MItem,
      this.muKern(-7 * PT, a),
      ...fillPiece(kind[2]),
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

  /**
   * amsmath \ext@arrow (\xrightarrow and friends). `phantom`: the label above or below is a \phantom of itself
   * (mathtools' harpoon pairs).
   */
  private extArrow(a: AtomLike, ctx: Ctx, spec: ExtArrow, phantom?: "above" | "below"): Noad {
    const [k1, k2, k3, k4] = spec.k;
    const below = a.branch("below");
    const above = a.branch("above");
    const nonEmpty = (x: AtomLike[] | undefined) => (x ?? []).some((y) => y.type !== "first");
    // the label as the macro argument: \phantom{#}, and mathtools' control spaces around it
    const label = (x: AtomLike[] | undefined, which: "above" | "below"): MItem[] => {
      const items = phantom === which ? [this.phantomOf(() => this.list(x, ctx), a)] : this.list(x, ctx);
      const space = (): MItem[] => this.controlSpace(a);
      return [...(spec.space === "pre" || spec.space === "both" ? space() : []), ...items, ...(spec.space === "post" || spec.space === "both" ? space() : [])];
    };
    const box0 = hpack(this.arrowFill(spec.fill, DISPLAY_STYLE, a), 0, undefined, a);
    const labelBox = (x: AtomLike[] | undefined, which: "above" | "below") =>
      this.mathBox([{ type: STYLE_NODE, style: SCRIPT_STYLE, tag: a }, this.muKern(k3 * PT, a), this.groupNoad(label(x, which), a), this.muKern(k4 * PT, a)], TEXT_STYLE, a);
    const w = Math.max(labelBox(below, "below").width, labelBox(above, "above").width, box0.width);
    const arrow = hpack(box0.list, w, EXACTLY, a);
    arrow.role = "ext-arrow";
    const op = newNoad(OP_NOAD, a);
    op.subtype = LIMITS;
    op.nucleus = boxField(arrow);
    const script = (x: AtomLike[] | undefined, which: "above" | "below"): Field => {
      const items: MItem[] = [];
      if (k1 !== 0) items.push(this.muKern(k1 * PT, a));
      items.push(...label(x, which));
      if (k2 !== 0) items.push(this.muKern(k2 * PT, a));
      return this.simplify(items);
    };
    // \@ifnotempty: a label with mathtools' control space or a \phantom is never empty
    if (nonEmpty(above) || spec.space || phantom === "above") op.sup = script(above, "above");
    if (nonEmpty(below) || spec.space || phantom === "below") op.sub = script(below, "below");
    const rel = newNoad(REL_NOAD, a);
    rel.nucleus = mlistField([op]);
    return rel;
  }

  /** mathtools \xrightleftharpoons, \xleftrightharpoons: two \ext@arrow boxes, raised and lowered by .22ex. */
  private extHarpoons(a: AtomLike, ctx: Ctx, [upper, lower]: [ExtArrow, ExtArrow]): Noad {
    const inner: Ctx = { ...ctx, display: false };
    const shift = texScale(".22", this.textRoman().param(5));
    const top = this.mathBox([this.extArrow(a, inner, upper, "below")], TEXT_STYLE, a);
    top.shift = -shift;
    const box0 = this.mathBox([this.extArrow(a, inner, lower, "above")], TEXT_STYLE, a);
    box0.shift = shift;
    const rel = newNoad(REL_NOAD, a);
    rel.nucleus = mlistField([this.boxNoad(top, a), { type: KERN_NODE, width: -box0.width, subtype: NORMAL, tag: a } as Node, this.boxNoad(box0, a)]);
    return rel;
  }

  /** latex.ltx \phantom: a \mathpalette box of the material's size, empty. */
  private phantomOf(make: () => MItem[], a: AtomLike): MItem {
    return this.palette(a, (st) => {
      const b0 = this.mathBox(this.styleList(st, a, [this.groupNoad(make(), a)]), TEXT_STYLE, a);
      const b2 = newNullBox(a);
      b2.role = "phantom";
      b2.height = b0.height;
      b2.depth = b0.depth;
      b2.width = b0.width;
      return [this.boxNoad(b2, a)];
    });
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

  /** A row's cells (a column a \multicolumn covers has no cell: empty here). */
  private rowCells(row: (AtomLike[] | undefined)[]): { cells: AtomLike[][] } {
    return { cells: row.map((cell) => cell ?? []) };
  }

  /** \arrayrulewidth */
  private ruleWidth(): number {
    return this.arrayParams.rulewidth ?? decimalToScaled(".4");
  }

  /** \doublerulesep */
  private doubleRuleSep(): number {
    return this.arrayParams.doublerulesep ?? 2 * PT;
  }

  /** A length as written ("2pt", ".5em") in sp, em and ex of the text font; null if it is none. */
  private dimen(text: string | undefined): number | null {
    const m = /^([+-]?(?:\d+\.?\d*|\.\d+))\s*(pt|pc|in|bp|cm|mm|dd|cc|sp|em|ex|px)$/.exec((text ?? "").trim());
    if (!m) return null;
    const rm = this.textRoman();
    return scanDimen(m[1], m[2], { em: rm.param(6), ex: rm.param(5), px: 65781 });
  }

  /**
   * \renewcommand{\arraystretch}{1.5}, \def\arraystretch{1.5},
   * \setlength{\arraycolsep}{2pt} ... (an invisible kernel atom): the arrays
   * after it in its group use the value. False if the atom is none of them.
   */
  private arraySetting(latex: string): boolean {
    const stretch = /^\\(?:renewcommand|def|newcommand|providecommand)\s*\{?\s*\\arraystretch\s*\}?\s*\{\s*(\d*\.?\d+)\s*\}$/.exec(latex.trim());
    if (stretch) {
      this.arrayParams = { ...this.arrayParams, stretch: stretch[1] };
      return true;
    }
    const length = /^\\setlength\s*\{?\s*\\(arraycolsep|tabcolsep|arrayrulewidth|doublerulesep|extrarowheight)\s*\}?\s*\{([^{}]*)\}$/.exec(latex.trim());
    const value = length ? this.dimen(length[2]) : null;
    if (!length || value === null) return false;
    const key = ({ arraycolsep: "colsep", tabcolsep: "tabcolsep", arrayrulewidth: "rulewidth", doublerulesep: "doublerulesep", extrarowheight: "extrarowheight" } as const)[
      length[1] as "arraycolsep"
    ];
    this.arrayParams = { ...this.arrayParams, [key]: value };
    return true;
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
   * array.sty's \@mkpream (its rules as arydshln makes them): the column
   * templates of a specification, as pieces around each column's cell.
   */
  private mkpream(spec: string): ArrayColumn[] {
    const out: (ArrayPiece | "&" | ArrayColumn)[] = [];
    let firstamp = true;
    let last = 4; // \@lastchclass
    let before: string[] = [];
    let lastColumn: ArrayColumn | null = null;
    let contentIndex = 0;
    const addamp = () => {
      if (firstamp) firstamp = false;
      else out.push("&");
    };
    const acol = () => out.push({ k: "acol" });
    // \@classx: before a column (or > and p/m/b)
    const classx = () => {
      if (last === 0 || last === 2) {
        acol();
        addamp();
        acol();
      } else if (last === 1) {
        addamp();
        acol();
      } else if (last === 4) {
        acol();
        firstamp = false;
      } else if (last === 5) addamp();
    };
    // \@classvi: before a rule or !{...}
    const classvi = () => {
      if (last === 0 || last === 2) acol();
      else if (last === 1) out.push({ k: "dbl" });
    };
    for (const t of specTokens(spec)) {
      switch (t.ch) {
        case "|":
        case ":": // arydshln
        case ";":
          classvi();
          out.push({ k: "rule", pattern: t.ch === "|" ? "solid" : t.ch === ";" ? t.arg ?? "" : "" });
          last = 1;
          break;
        case "!":
          classvi();
          out.push({ k: "content", src: t.arg ?? "", index: contentIndex++ });
          last = 1;
          break;
        case "@":
          out.push({ k: "content", src: t.arg ?? "", index: contentIndex++ });
          last = 5;
          break;
        case ">":
          classx();
          before.unshift(t.arg ?? ""); // \save@decl puts a later declaration first
          last = 3;
          break;
        case "<":
          lastColumn?.after.push(t.arg ?? "");
          last = 2;
          break;
        default: {
          // c, l, r, p{..}, m{..}, b{..}; anything else is `c' (\@preamerr 0)
          classx();
          const align = "lrpmb".includes(t.ch) ? (t.ch as ArrayColumn["align"]) : "c";
          lastColumn = { pre: [], align, width: t.arg, before, after: [], post: [] };
          before = [];
          out.push(lastColumn);
          last = 0;
        }
      }
    }
    if (last === 0 || last === 2) acol();
    const cols: ArrayColumn[] = [];
    let pieces: ArrayPiece[] = [];
    let column: ArrayColumn | null = null;
    const flush = () => {
      if (column) {
        column.post = pieces;
        cols.push(column);
      }
      column = null;
      pieces = [];
    };
    for (const x of out) {
      if (x === "&") flush();
      else if ("align" in x) {
        column = x;
        column.pre = pieces;
        pieces = [];
      } else pieces.push(x);
    }
    flush();
    return cols;
  }

  /**
   * The pieces of a template (\hskip\col@sep, \vline ...). `gaps`: the
   * kernel's atoms of the specification's @{...} and !{...}, in order.
   */
  private piece(p: ArrayPiece, a: AtomLike, ctx: Ctx, colsep: number, gaps: AtomLike[][]): Node[] {
    const rw = this.ruleWidth();
    switch (p.k) {
      case "acol":
        return [this.hskip(colsep, a)]; // \hskip\col@sep
      case "rule":
        // arydshln: \vrule width\arrayrulewidth height0pt depth0pt; the rule is drawn over the rows (dashColumns)
        return [{ type: RULE_NODE, width: rw, height: 0, depth: 0, tag: a }];
      case "dbl":
        return [this.hskip(this.doubleRuleSep(), a)];
    }
    // @{...}, !{...}: \d@llarbegin ... \d@llarend
    const atoms = gaps[p.index];
    if (atoms && atoms.every((x) => typeof x?.branch === "function")) return this.cellFormula(this.list(atoms, ctx), a);
    // LaTeX's own specifications (cases: @{}l@{\quad}l@{})
    return this.cellFormula(p.src.trim() === "\\quad" ? [this.hskip(this.textRoman().param(6), a) as MItem] : [], a);
  }

  /** A declaration in a >{...} or <{...}: a math style switch (others have no effect here). */
  private declStyle(decl: string, a: AtomLike): MItem[] {
    const styles: Record<string, number> = {
      "\\displaystyle": DISPLAY_STYLE,
      "\\textstyle": TEXT_STYLE,
      "\\scriptstyle": SCRIPT_STYLE,
      "\\scriptscriptstyle": SCRIPT_SCRIPT_STYLE,
    };
    const out: MItem[] = [];
    for (const m of decl.matchAll(/\\[a-zA-Z]+/g)) if (styles[m[0]] !== undefined) out.push({ type: STYLE_NODE, style: styles[m[0]], tag: a });
    return out;
  }

  /**
   * A trailing \\ (an empty last row) adds no row in LaTeX; the editor keeps
   * it, since the caret of a line just started with Shift+Enter stands there.
   * Followed by rules (`\\ \hline` after it), it is a row in LaTeX too.
   */
  private dropsRow(a: AtomLike, cells: AtomLike[][], r: number, count: number): boolean {
    return (
      !this.editor &&
      !(a.bottomRules > 0) &&
      r === count - 1 &&
      r > 0 &&
      !cells.some((c, j) => c.some((x) => x.type !== "first") || a.getCellWrappers?.(r, j))
    );
  }

  /** arydshln's dash pattern ("4pt/2pt"; \dashlinedash/\dashlinegap, 4pt each, if empty; "solid": none). */
  private dashPattern(pattern: string): [number, number] {
    if (pattern === "solid") return [0, 0];
    const [d, g] = pattern.split("/");
    return [this.dimen(d) ?? 4 * PT, this.dimen(g) ?? 4 * PT];
  }

  /**
   * arydshln's \adl@drawi: `half` dash, half gap, \xleaders of (half gap,
   * dash, half gap) over the rest, half gap, half dash; along a vertical
   * (`vertical`: a vlist `length` long) or horizontal line, rules `rw` thick.
   * A dash or gap of 0 makes a solid line.
   */
  private dashLine(pattern: string, vertical: boolean, rw: number, a: AtomLike, length = 0): Node[] {
    const [d, g] = this.dashPattern(pattern);
    const rule = (x: number): Node =>
      vertical ? { type: RULE_NODE, width: rw, height: x, depth: 0, tag: a } : { type: RULE_NODE, width: x, height: rw, depth: 0, tag: a };
    if (d <= 0 || g <= 0) {
      if (vertical) return [rule(length)];
      return [{ type: GLUE_NODE, spec: glue(0, PT, FILL_ORDER), subtype: A_LEADERS, leader: { type: RULE_NODE, width: 0, height: rw, depth: 0, tag: a }, tag: a }];
    }
    const hd = texScale(".5", d);
    const hg = texScale(".5", g);
    const unitList = [this.hskip(hg, a), rule(d), this.hskip(hg, a)];
    const unit = vertical ? vpack(unitList, a) : hpack(unitList, 0, undefined, a);
    const leaders: Node = { type: GLUE_NODE, spec: glue(0, PT, FIL, PT, FIL), subtype: X_LEADERS, leader: unit, tag: a };
    return [rule(hd), this.hskip(hg, a), leaders, this.hskip(hg, a), rule(hd)];
  }

  /** booktabs' and LaTeX's rule commands: the name and arguments of `\cmidrule[w](lr){2-3}`. */
  private parseRule(src: string | undefined): { name: string; args: { kind: string; value: string }[] } {
    const m = /^\\[a-zA-Z]+/.exec(src ?? "");
    if (!m || src === undefined) return { name: "", args: [] };
    const args: { kind: string; value: string }[] = [];
    let i = m[0].length;
    while (i < src.length) {
      const open = src[i];
      const close = open === "[" ? "]" : open === "(" ? ")" : open === "{" ? "}" : "";
      if (!close) {
        i++;
        continue;
      }
      let depth = 0;
      let j = i + 1;
      for (; j < src.length; j++) {
        if (src[j] === "{") depth++;
        else if (src[j] === "}" && depth > 0) depth--;
        else if (src[j] === close && depth === 0) break;
      }
      args.push({ kind: open, value: src.slice(i + 1, j) });
      i = j + 1;
    }
    return { name: m[0], args };
  }

  /** "2-3" of \cline, \cdashline, \cmidrule: the first column and the number of columns, within `ncols`. */
  private ruleColumns(range: string | undefined, ncols: number): [number, number] {
    const m = /^\s*(\d+)\s*-\s*(\d+)\s*$/.exec(range ?? "");
    const from = Math.max(1, Math.min(m ? parseInt(m[1], 10) : 1, ncols));
    const to = Math.max(from, Math.min(m ? parseInt(m[2], 10) : ncols, ncols));
    return [from, to - from + 1];
  }

  /**
   * The rules between two rows, in order: \hline, \cline (latex.ltx with
   * array.sty's \doublerulesep), arydshln's \hdashline and \cdashline,
   * booktabs' rules and \addlinespace. Rows they add go to `marks.rows`, the
   * material arydshln's vertical lines cross to `marks.connect`, the gaps
   * they stop at to `marks.cut`.
   */
  private tableRules(rules: readonly string[], a: AtomLike, ncols: number, items: (AlignRow | AlignNoalign)[], marks: ArrayMarks): void {
    const rw = this.ruleWidth();
    const rm = this.textRoman();
    const em = (f: string) => texScale(f, rm.param(6));
    const ex = (f: string) => texScale(f, rm.param(5));
    const vskip = (h: number, kind: "cut" | "connect" | null): AlignNoalign => ({
      kind: "noalign",
      apply: (v) => {
        const n = this.hskip(h, a);
        if (kind) marks[kind].add(n);
        v.append(n);
      },
    });
    const hrule = (h: number, kind: "cut" | "connect"): AlignNoalign => ({
      kind: "noalign",
      apply: (v) => {
        const n: RuleNode = { type: RULE_NODE, width: -(2 ** 30), height: h, depth: 0, tag: a };
        marks[kind].add(n);
        v.appendRule(n);
      },
    });
    // a row of \omit'ed cells: the columns before `from` empty, then `list` over `span` columns
    const spanRow = (from: number, span: number, list: Node[], connect: boolean) => {
      const cells: Node[][] = from > 1 ? [[], list] : [list];
      const spans = from > 1 ? [from - 1, span] : [span];
      const row: AlignRow = { kind: "row", cells, spans, tag: a };
      if (connect) marks.rowsConnect.add(row);
      items.push(row);
    };
    const hfillRule = (h: number): Node => ({
      type: GLUE_NODE,
      spec: glue(0, PT, FILL_ORDER),
      subtype: A_LEADERS,
      leader: { type: RULE_NODE, width: 0, height: h, depth: 0, tag: a },
      tag: a,
    });
    const BOOKTABS = ["\\toprule", "\\midrule", "\\bottomrule", "\\cmidrule", "\\specialrule", "\\addlinespace"];
    let lastClass = 0; // booktabs' \@lastruleclass
    rules.forEach((src, i) => {
      const { name, args } = this.parseRule(src);
      const next = this.parseRule(rules[i + 1]).name;
      const opt = (kind: string, k = 0) => args.filter((x) => x.kind === kind)[k]?.value;
      switch (name) {
        case "\\hline":
        case "\\hdashline": {
          if (name === "\\hline") items.push(hrule(rw, "connect"));
          else spanRow(1, ncols, this.dashLine(opt("[") ?? "", false, rw, a), true);
          // array.sty's \@xhline (arydshln's \adl@ixhline): \vskip\doublerulesep
          // before a \hline or \hdashline right after it (a space after
          // \hdashline[...] is a token \futurelet sees)
          if ((next === "\\hline" || next === "\\hdashline") && !(name === "\\hdashline" && /\]\s$/.test(src)))
            items.push(vskip(this.doubleRuleSep(), "cut"));
          break;
        }
        case "\\cline":
        case "\\cdashline": {
          // \omit ... \leaders\hrule\@height\arrayrulewidth\hfill \cr \noalign{\vskip-\arrayrulewidth}
          const [from, span] = this.ruleColumns(opt("{"), ncols);
          spanRow(from, span, name === "\\cline" ? [hfillRule(rw)] : this.dashLine(opt("[") ?? "", false, rw, a), false);
          items.push(vskip(-rw, null));
          break;
        }
        case "\\toprule":
        case "\\midrule":
        case "\\bottomrule":
        case "\\specialrule":
        case "\\addlinespace": {
          const special = name === "\\specialrule";
          const width = special
            ? this.dimen(opt("{", 0)) ?? 0
            : this.dimen(opt("[")) ?? (name === "\\midrule" ? em(".05") : name === "\\addlinespace" ? 0 : em(".08"));
          const above = special ? this.dimen(opt("{", 1)) ?? 0 : name === "\\toprule" ? 0 : ex(".4");
          const below = special
            ? this.dimen(opt("{", 2)) ?? 0
            : name === "\\bottomrule"
              ? 0
              : name === "\\addlinespace"
                ? this.dimen(opt("[")) ?? em(".5")
                : ex(".65");
          const thisClass = special || name === "\\addlinespace" ? 2 : 1;
          if (name !== "\\addlinespace") {
            // \@BTrule
            if (thisClass === 2 || lastClass === 0) items.push(vskip(above, "cut"));
            else if (lastClass === 1) items.push(vskip(this.doubleRuleSep(), "cut"));
            items.push(hrule(width, "cut"));
          }
          // \@BTendrule
          lastClass = BOOKTABS.includes(next) ? thisClass : 0;
          if (lastClass !== 1) items.push(vskip(below, "cut"));
          break;
        }
        case "\\cmidrule": {
          const width = this.dimen(opt("[")) ?? em(".03");
          const [from, span] = this.ruleColumns(opt("{"), ncols);
          // \@setrulekerning: l and r trim by \cmidrulekern (.5em) or the length after them
          let kl = 0;
          let kr = 0;
          for (const m of (opt("(") ?? "").matchAll(/([lr])(?:\{([^}]*)\})?/g)) {
            const k = m[2] !== undefined ? this.dimen(m[2]) ?? 0 : em(".5");
            if (m[1] === "l") kl = k;
            else kr = k;
          }
          if (lastClass === 0) items.push(vskip(ex(".4"), "cut"));
          spanRow(from, span, [this.hskip(kl, a), hfillRule(width), newKern(0, a), this.hskip(kr, a)], false);
          // \@xcmidrule
          if (next === "\\cmidrule") {
            items.push(vskip(-width, "cut"));
            lastClass = 1;
          } else if (next === "\\morecmidrules") {
            items.push(vskip(this.doubleRuleSep(), "cut"));
            lastClass = 1;
          } else {
            items.push(vskip(ex(".65"), "cut"));
            lastClass = 0;
          }
          break;
        }
      }
    });
  }

  /**
   * multirow.sty's \multirow[vpos]{n}[bigstruts]{width}[vmove]{text}: an
   * \hbox{\strut text\strut} (or a paragraph `width` wide) in a box of no
   * height and depth, moved to the middle (top, bottom) of the n rows from
   * this one (n < 0: up to this one); `arH`, `arD`: \@arstrutbox.
   */
  private multirowBox(w: Extract<CellWrapperLike, { command: "\\multirow" }>, content: Node[], arH: number, arD: number, a: AtomLike): BoxNode {
    const [sh, sd] = this.strutDims();
    const strut = (): Node => ({ type: RULE_NODE, width: 0, height: sh, depth: sd, tag: a });
    const width = w.width.trim() === "*" || w.width.trim() === "=" ? null : this.dimen(w.width);
    const line = width === null ? hpack([strut(), ...content, strut()], 0, undefined, a) : hpack([strut(), ...content, strut(), this.hfil(a)], width, EXACTLY, a);
    const vpos = (w.vpos ?? "c").trim();
    const n = (w.rows ?? "1").trim();
    const negative = n.startsWith("-");
    // \bigstrutjot (\jot) for each big strut, t and b ones at the ends
    const bs = /^(t?)(b?)(\d*)$/.exec((w.bigstruts ?? "0").trim());
    const jot = 3 * PT;
    const struts = bs ? parseInt(bs[3] || "0", 10) : 0;
    const total = Math.abs(texScale(negative ? n.slice(1) : n, arH)) + Math.abs(texScale(negative ? n.slice(1) : n, arD)) + struts * jot;
    const h = line.height;
    const d = line.depth;
    let dima: number;
    if (!negative) {
      if (vpos === "t") dima = h;
      else {
        dima = arH + (bs?.[1] ? jot : 0);
        if (vpos === "b") dima += arD + (bs?.[2] ? jot : 0);
      }
    } else {
      dima = total;
      if (vpos !== "b") {
        dima -= arD + (bs?.[2] ? jot : 0);
        if (vpos === "t") dima += -arH - (bs?.[1] ? jot : 0) + h;
      }
    }
    dima += this.dimen(w.vmove) ?? 0;
    // \vtop to n rows {\vfill line \vfill}: t, the line first; b, the line last (its depth below)
    const margin = vpos === "t" ? 0 : vpos === "b" ? total - h : Math.round((total - h - d) / 2);
    line.shift = -dima + margin + h;
    const box = hpack([line], 0, undefined, a);
    box.height = 0;
    box.depth = 0;
    return box;
  }

  /** amsmath's \hdotsfor dots: \mkern-s\dotsspace@ \xleaders\hbox{$\mkern s\dotsspace@.\mkern s\dotsspace@$}\hfill \mkern-s\dotsspace@ */
  private hdots(spacing: string | undefined, a: AtomLike): MItem[] {
    const s = texScale((spacing ?? "1").trim() || "1", texScale("1.5", PT)); // \dotsspace@ = 1.5mu
    const dot = this.mathBox([this.muKern(s, a), this.mathChar(this.data.meanings.mathcode["."], BASE_CTX, a), this.muKern(s, a)], TEXT_STYLE, a);
    const leaders: Node = { type: GLUE_NODE, spec: glue(0, PT, FILL_ORDER), subtype: X_LEADERS, leader: dot, tag: a };
    return [this.muKern(-s, a), leaders as MItem, this.muKern(-s, a)];
  }

  /**
   * LaTeX's \array (array.sty's \@array; arydshln's dashed rules): returns the
   * \vcenter noad. `tabular`: its cells are text and \tabcolsep apart.
   */
  private latexArray(a: AtomLike, spec: string, arraystretch: string, ctx: Ctx, cellStyle: number | null = null, tabular = false): Noad {
    const cols = this.mkpream(spec);
    const params = this.arrayParams;
    const [sh, sd] = this.strutDims();
    // \arraystretch, times the user's row spacing
    const arH = scalePermille(texScale(arraystretch, sh + (params.extrarowheight ?? 0)), this.arrayRowSep);
    const arD = scalePermille(texScale(arraystretch, sd), this.arrayRowSep);
    const colsep = tabular ? params.tabcolsep ?? 6 * PT : scalePermille(params.colsep ?? 5 * PT, this.arrayColSep);
    const gaps: AtomLike[][] = (a.colFormat ?? []).filter((c: any) => Array.isArray(c.gap)).map((c: any) => c.gap);
    const rw = this.ruleWidth();
    const items: (AlignRow | AlignNoalign)[] = [];
    const marks: ArrayMarks = { connect: new WeakSet(), cut: new WeakSet(), rowsConnect: new WeakSet() };
    const dataRows = new WeakSet<AlignRow>();
    // arydshln's column rules: a box in each row that has one, filled once the rows are set
    const dashes: { holder: BoxNode; key: string; pattern: string; offset: number; row: AlignRow }[] = [];
    const rows: (AtomLike[] | undefined)[][] = a.rows ?? [];
    const rules: string[][] = a.rowRules ?? [];
    const strut = (): Node => ({ type: RULE_NODE, width: 0, height: arH, depth: arD, tag: a }); // \@arstrut
    rows.forEach((row, r) => {
      this.tableRules(rules[r] ?? [], a, cols.length, items, marks);
      const { cells } = this.rowCells(row);
      if (this.dropsRow(a, cells, r, rows.length)) return;
      const lists: Node[][] = [];
      const spans: number[] = [];
      const pending: { key: string; pattern: string; offset: number; holder: BoxNode }[] = [];
      let lastCell = -1;
      for (let j = 0; j < Math.min(row.length, cols.length); j++) if (row[j]) lastCell = j;
      for (let j = 0; j < Math.min(row.length, cols.length); j++) {
        const cell = row[j];
        if (!cell) continue;
        // a \multicolumn spans the columns after it it covers
        let span = 1;
        while (j + span < Math.min(row.length, cols.length) && !row[j + span]) span++;
        const wrappers: CellWrapperLike[] = a.getCellWrappers?.(r, j) ?? [];
        const multicolumn = wrappers.find((w) => w.command === "\\multicolumn" || w.command === "\\hdotsfor");
        const col = multicolumn ? this.mkpream(multicolumn.command === "\\multicolumn" ? multicolumn.spec : "c")[0] ?? cols[j] : cols[j];
        const list: Node[] = [];
        if (j === 0 || multicolumn) list.push(strut());
        // arydshln draws a column's rules from its edge, \doublerulesep apart:
        // the k-th rule from the edge k\arrayrulewidth closer to it than its place
        const pieces = (side: "L" | "R", ps: ArrayPiece[]) => {
          const count = ps.filter((p) => p.k === "rule").length;
          let index = 0;
          for (const p of ps) {
            if (p.k !== "rule") {
              list.push(...this.piece(p, a, ctx, colsep, gaps));
              continue;
            }
            const k = side === "L" ? index : count - 1 - index;
            index++;
            const holder: BoxNode = { type: HLIST_NODE, width: rw, height: 0, depth: 0, shift: 0, list: [], glueSet: 0, glueSign: 0, glueOrder: 0, tag: a };
            const key = `${side}${side === "L" ? j : j + span - 1}:${k}:${p.pattern}`;
            pending.push({ key, pattern: p.pattern, offset: side === "L" ? -k * rw : k * rw, holder });
            list.push(holder);
          }
        };
        pieces("L", col.pre);
        // the cell's formula ($...$; text in a tabular), with the >{...} and <{...} style switches
        let content: MItem[];
        const multirow = wrappers.find((w) => w.command === "\\multirow") as Extract<CellWrapperLike, { command: "\\multirow" }> | undefined;
        const visible = tabular ? trimSpaces(cell) : cell;
        if (multicolumn?.command === "\\hdotsfor") content = [...this.hdots(multicolumn.spacing, a), ...this.list(visible, ctx)];
        else if (multirow) content = [this.boxNoad(this.multirowBox(multirow, this.cellFormula(this.list(visible, ctx), a), arH, arD, a), a)];
        else content = this.list(visible, ctx);
        const styles = [...(cellStyle === null ? [] : [{ type: STYLE_NODE, style: cellStyle, tag: a } as MItem]), ...col.before.flatMap((x) => this.declStyle(x, a))];
        const formula = this.cellFormula([...styles, ...content, ...col.after.flatMap((x) => this.declStyle(x, a))], a);
        // \\[len]: \do@row@strut in the last cell
        const gap = this.rowGap(a, r);
        const rowStrut: Node[] =
          j === lastCell && (r < rows.length - 1 || a.bottomRules > 0) && gap > 0 ? [{ type: RULE_NODE, width: 0, height: 0, depth: arD + gap, tag: a }] : [];
        if (col.align === "p" || col.align === "m" || col.align === "b") {
          // \@startpbox{w} (\everypar strut, \parindent 0) ... \@finalstrut\@arstrutbox \par \@endpbox\hfil:
          // one line of the paragraph (no line breaking)
          const w = this.dimen(col.width) ?? 0;
          const lineBox = hpack(
            [{ type: RULE_NODE, width: 0, height: arH, depth: 0, tag: a }, ...formula, { type: RULE_NODE, width: 0, height: 0, depth: arD, tag: a }, this.hfil(a)],
            w,
            EXACTLY,
            a,
          );
          // m: arydshln's \adl@startmbox, $\vcenter{...}$
          let pbox: BoxNode = lineBox;
          if (col.align === "m") {
            const vc = newNoad(VCENTER_NOAD, a);
            vc.nucleus = boxField(vpack([lineBox], a));
            pbox = this.mathBox([vc], TEXT_STYLE, a);
          }
          list.push(pbox, ...rowStrut, this.hfil(a));
        } else {
          if (col.align !== "l") list.push(this.hfil(a));
          list.push(...formula, ...rowStrut);
          if (col.align !== "r") list.push(this.hfil(a));
        }
        pieces("R", col.post);
        lists.push(list);
        spans.push(span);
      }
      const item: AlignRow = { kind: "row", cells: lists, spans, tag: a };
      dataRows.add(item);
      items.push(item);
      for (const p of pending) dashes.push({ ...p, row: item });
      const gap = this.rowGap(a, r);
      if (gap < 0 && (r < rows.length - 1 || a.bottomRules > 0))
        items.push({
          kind: "noalign",
          apply: (v) => {
            const n = this.hskip(gap, a);
            marks.connect.add(n);
            v.append(n);
          },
        });
    });
    // the rules under the last row's \\ (kernel rowRules' last entry)
    this.tableRules(rules[rows.length] ?? [], a, cols.length, items, marks);
    const v = new VList({ baselineskip: glue(0), lineskip: glue(0), lineskiplimit: 0 }, a);
    const rowBoxes = halign(items, [glue(0)], v, a);
    if (dashes.length > 0) this.dashColumns(v.items, items, rowBoxes, dataRows, marks, dashes, rw, a);
    const box = vpack(v.items, a);
    box.role = "array";
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(box);
    return vc;
  }

  /**
   * arydshln's \adl@makevlr: each column rule runs down the rows that have
   * it, through the rules (\hline, \hdashline) between them, and stops at a
   * double rule's gap or a row without it; a dashed one's dashes are spread
   * over that length. Each line goes into the box of its first row.
   */
  private dashColumns(
    vitems: Node[],
    items: (AlignRow | AlignNoalign)[],
    rowBoxes: BoxNode[],
    dataRows: WeakSet<AlignRow>,
    marks: ArrayMarks,
    dashes: { holder: BoxNode; key: string; pattern: string; offset: number; row: AlignRow }[],
    rw: number,
    a: AtomLike,
  ): void {
    const rowOf = new Map<Node, AlignRow>();
    items.filter((x): x is AlignRow => x.kind === "row").forEach((row, i) => rowOf.set(rowBoxes[i], row));
    type Seg = { kind: "row" | "connect" | "cut" | "none"; top: number; h: number; row?: AlignRow; box?: BoxNode };
    const segs: Seg[] = [];
    let y = 0;
    for (const n of vitems) {
      let h = 0;
      if (n.type === HLIST_NODE || n.type === VLIST_NODE || n.type === RULE_NODE) h = n.height + n.depth;
      else if (n.type === GLUE_NODE) h = n.spec.width;
      else if (n.type === KERN_NODE) h = n.width;
      const row = rowOf.get(n);
      const kind = row && dataRows.has(row) ? "row" : (row && marks.rowsConnect.has(row)) || marks.connect.has(n) ? "connect" : marks.cut.has(n) ? "cut" : "none";
      segs.push({ kind, top: y, h, row, box: row ? (n as BoxNode) : undefined });
      y += h;
    }
    const keys = new Map<string, Map<AlignRow, { holder: BoxNode; pattern: string; offset: number }>>();
    for (const d of dashes) {
      if (!keys.has(d.key)) keys.set(d.key, new Map());
      keys.get(d.key)!.set(d.row, d);
    }
    type Start = { top: number; holder: BoxNode; pattern: string; offset: number; baseline: number };
    for (const byRow of keys.values()) {
      const open: Start[] = [];
      let lastconn = 0;
      const emit = (end: number) => {
        const start = open.pop();
        if (!start) return;
        const length = end - start.top;
        const line = vpackage(this.dashLine(start.pattern, true, rw, a, length), length, EXACTLY, MAX_DIMEN, a);
        line.shift = start.top - start.baseline + line.height;
        start.holder.list.push(newKern(start.offset, a), line, newKern(-start.offset - rw, a));
      };
      for (const s of segs) {
        if (s.kind === "row") {
          const d = byRow.get(s.row!);
          if (d && open.length === 0) open.push({ top: s.top - lastconn, holder: d.holder, pattern: d.pattern, offset: d.offset, baseline: s.top + s.box!.height });
          else if (!d) emit(s.top);
          lastconn = 0;
        } else if (s.kind === "cut") {
          emit(s.top);
          lastconn = 0;
        } else if (s.kind === "connect") lastconn = s.h;
      }
      emit(y);
    }
  }

  /** A tabular in text (\text{\begin{tabular}...}): \hbox{$\vcenter{...}$} at a text size. */
  private tabularBox(a: AtomLike): BoxNode {
    const spec = a.colSpec ?? "c";
    return this.mathBox([this.latexArray(a, spec, this.arrayParams.stretch ?? "1", BASE_CTX, null, true)], TEXT_STYLE, a);
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
    const acs = scalePermille(this.arrayParams.colsep ?? 5 * PT, this.arrayColSep); // \arraycolsep
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
        const body: MItem[] = [this.hskip(-acs, a) as MItem, this.latexArray(a, "*{10}{c}", this.arrayParams.stretch ?? "1", ctx), this.hskip(-acs, a) as MItem];
        const d = delims[env];
        return d ? [this.leftRightWrap(a, d[0], d[1], body)] : body;
      }
      case "cases": {
        // \left\lbrace \def\arraystretch{1.2} \array{@{}l@{\quad}l@{}} ... \right.
        const arr = this.latexArray(a, "@{}l@{\\quad}l@{}", "1.2", ctx);
        return [this.leftRightWrap(a, "\\lbrace", ".", [arr])];
      }
      case "dcases":
      case "rcases":
      case "drcases":
      case "cases*":
      case "dcases*":
      case "rcases*":
      case "drcases*":
        return this.mathtoolsCases(a, env, ctx);
      case "matrix*":
      case "pmatrix*":
      case "bmatrix*":
      case "Bmatrix*":
      case "vmatrix*":
      case "Vmatrix*": {
        // mathtools \MT_matrix_begin:N: \hskip-\arraycolsep \array{*\c@MaxMatrixCols #1} with one alignment
        const align = /^\[\s*([lcr])\s*\]/.exec(a.environmentArguments ?? "")?.[1] ?? "c";
        const body: MItem[] = [this.hskip(-acs, a) as MItem, this.latexArray(a, `*{10}{${align}}`, "1", ctx), this.hskip(-acs, a) as MItem];
        const d = delims[env.slice(0, -1)];
        return d ? [this.leftRightWrap(a, d[0], d[1], body)] : body;
      }
      case "subarray":
        return [this.subarray(a, ctx, (a.colFormat ?? []).find((c: any) => c.align)?.align === "l" ? "l" : "c")];
      case "align":
      case "align*":
      case "flalign":
      case "flalign*":
      case "alignat":
      case "alignat*":
      case "eqnarray":
      case "eqnarray*":
      case "gather":
      case "gather*":
      case "multline":
      case "multline*":
      case "equation":
      case "equation*":
      case "displaymath":
        return this.displayEnv(a, env.replace(/\*$/, ""), ctx);
      case "array": {
        // the column specification as written (kernel colSpec)
        const spec =
          a.colSpec ??
          (a.colFormat ?? []).map((c: any) => (c.align ? c.align : c.separator === "solid" ? "|" : c.separator === "dashed" ? ":" : "")).join("");
        return [this.latexArray(a, spec || "c", this.arrayParams.stretch ?? "1", ctx)];
      }
      case "tabular":
        return [this.boxNoad(this.tabularBox(a), a)];
      case "aligned":
      case "alignedat":
      case "split":
        return this.alignedEnv(a, ctx, env === "alignedat");
      case "gathered":
        return this.gatheredEnv(a, ctx);
      case "smallmatrix":
        return this.smallMatrixEnv(a, ctx);
      default: {
        const body: MItem[] = [this.hskip(-acs, a) as MItem, this.latexArray(a, "*{10}{c}", this.arrayParams.stretch ?? "1", ctx), this.hskip(-acs, a) as MItem];
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
      if (this.dropsRow(a, cells, r, rows.length)) return;
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
    return [this.nullBox(a), this.positioned(box, a)];
  }

  /** amsmath gathered */
  private gatheredEnv(a: AtomLike, ctx: Ctx): MItem[] {
    const rows: AtomLike[][][] = a.rows ?? [];
    const items: (AlignRow | AlignNoalign)[] = [];
    rows.forEach((row, r) => {
      const { cells } = this.rowCells(row);
      if (this.dropsRow(a, cells, r, rows.length)) return;
      const f = this.cellFormula([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, ...this.list(cells[0], ctx)], a);
      items.push({ kind: "row", cells: [[this.hfil(a), this.amsStrut(a), ...f, this.hfil(a)]], tag: a });
    });
    const v = new VList(this.spreadParams(), a);
    halign(items, [glue(0)], v, a);
    const box = vpack(v.items, a);
    return [this.nullBox(a), this.positioned(box, a)];
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
      if (this.dropsRow(a, cells, r, rows.length)) return;
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
      if (this.dropsRow(a, cells, r, rows.length)) return;
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

  // ------------------------------------------------- display environments --
  /**
   * VisualTeX: the editor has no line width. Where amsmath spreads a display
   * over \displaywidth (the space between align's column pairs, flalign,
   * multline, the place of an equation tag) the display is set as in an
   * article's \textwidth at 10pt (345pt); without tags the outer margins of
   * align and alignat are dropped, so the formula starts at its first column.
   */
  private displayWidth(): number {
    return 345 * PT;
  }

  /** amsmath \mintagsep: .5\fontdimen6\textfont2 */
  private minTagSep(): number {
    return texScale(".5", this.fams.get(2)![0]!.param(6));
  }

  /** \jot, times the user's row spacing */
  private jot(): number {
    return scalePermille(3 * PT, this.alignRowSep);
  }

  /** amsmath \tagform@ / \maketag@@@ in \setboxz@h{\@lign\strut@{...}}: \hbox{\normalfont (#1)}, or #1 for \tag*. */
  private tagBox(a: AtomLike, ctx: Ctx): BoxNode {
    // \ignorespaces #1\unskip
    const body = this.textModeBox(trimText(a.branch("body") ?? [], true), ctx, a);
    if (a.command === "\\tag*") return hpack([this.amsStrut(a), body], 0, undefined, a);
    const paren = (ch: string): BoxNode => {
      const b = this.textRunBox([{ type: "mord", mode: "text", value: ch, style: (a.branch("body") ?? []).find((x) => x.mode === "text")?.style, branch: () => undefined }], 0);
      for (const n of b.list) (n as { tag?: Tag }).tag = a;
      b.tag = a;
      return b;
    };
    return hpack([this.amsStrut(a), paren("("), body, paren(")")], 0, undefined, a);
  }

  /** A display's single line with its \tag (amsmath \calc@shift@gather, tags on the right). */
  private taggedDisplay(atoms: AtomLike[], tag: AtomLike, ctx: Ctx): Noad {
    const f = this.cellFormula([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: undefined }, ...this.list(atoms, ctx)], tag);
    const line = this.gatherLines([{ list: f, tag: this.tagBox(tag, ctx) }]);
    return this.boxNoad(line[0], undefined);
  }

  /**
   * gather / equation lines: each centred in the display width, its tag at the
   * right end (amsmath \calc@shift@gather, \place@tag@gather). Without tags
   * the lines are centred on the widest.
   */
  private gatherLines(lines: { list: Node[]; tag: BoxNode | null }[]): BoxNode[] {
    const hfil = (): Node => ({ type: GLUE_NODE, spec: glue(0, PT, 1), subtype: NORMAL });
    const widths = lines.map((l) => hpack(l.list).width);
    const tagged = lines.some((l) => l.tag);
    const sep = this.minTagSep();
    let d = Math.max(0, ...widths);
    if (tagged) d = Math.max(this.displayWidth(), ...lines.map((l, i) => widths[i] + (l.tag ? 2 * sep + l.tag.width : 0)));
    return lines.map((l, i) => {
      if (!tagged) return hpack([hfil(), ...l.list, hfil()], d, EXACTLY);
      let shift = d - widths[i];
      if (l.tag && shift < 4 * l.tag.width) shift -= l.tag.width;
      shift = Math.max(0, div(shift, 2));
      return hpack([newKern(shift), ...l.list, hfil(), ...(l.tag ? [l.tag] : [])], d, EXACTLY);
    });
  }

  /** The lines of a display environment: its rows split at \intertext, each row's \tag taken out. */
  private displayLines(a: AtomLike): DisplayLine[] {
    const out: DisplayLine[] = [];
    const rows: AtomLike[][][] = a.rows ?? [];
    const blank = (cells: AtomLike[][]) => cells.every((c) => c.every((x) => x.type === "first"));
    rows.forEach((row, r) => {
      const { cells } = this.rowCells(row);
      if (this.dropsRow(a, cells, r, rows.length)) return;
      let cur: AtomLike[][] = [];
      let tag: AtomLike | null = null;
      for (const cell of cells) {
        let part: AtomLike[] = [];
        for (const x of cell) {
          if (isIntertextAtom(x)) {
            // amsmath: \ifvmode\else\\\@empty\fi \noalign{...}: it ends the row it stands in
            if (!blank([...cur, part]) || tag) {
              out.push({ kind: "row", cells: [...cur, part], tag, r });
              cur = [];
              part = [];
              tag = null;
            }
            out.push({ kind: "text", atom: x });
          } else if (isTagAtom(x)) tag ??= x;
          else part.push(x);
        }
        cur.push(part);
      }
      out.push({ kind: "row", cells: cur, tag, r });
    });
    return out;
  }

  /** \intertext / \shortintertext (mathtools \MT_intertext:): the text between \abovedisplayskip-sized gaps. */
  private intertextNoalign(x: AtomLike, ctx: Ctx, lineskiplimit: number): AlignNoalign {
    const short = x.command === "\\shortintertext";
    // \abovedisplayskip = \belowdisplayskip = 10pt plus 2pt minus 5pt, \abovedisplayshortskip 0pt plus 3pt
    // (\belowdisplayshortskip in \MT_shortintertext: is not used); the short form adds 3pt on each side
    const skip = short ? 3 * PT : 10 * PT;
    return {
      kind: "noalign",
      apply: (v) => {
        v.append(this.hskip(skip - lineskiplimit, x));
        v.appendBox(hpack([this.textModeBox(x.branch("body"), ctx, x)], 0, undefined, x));
        v.append(this.hskip(skip - lineskiplimit, x));
      },
    };
  }

  /**
   * amsmath's top-level displays (align, flalign, alignat, gather, multline,
   * equation) and LaTeX's eqnarray, laid out as a \vcenter: the editor's
   * display is the formula itself.
   */
  private displayEnv(a: AtomLike, env: string, ctx: Ctx): MItem[] {
    const lines = this.displayLines(a);
    const items: (AlignRow | AlignNoalign)[] = [];
    const tags = new Map<AlignRow, BoxNode>();
    const jot = this.jot();
    const gapAfter = (r: number) => (r < (a.rows ?? []).length - 1 ? this.rowGap(a, r) : 0);
    const display = (list: MItem[]) => this.cellFormula([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, this.groupNoad(list, a)], a);
    const lastRow = [...lines].reverse().find((l) => l.kind === "row") as DisplayLine | undefined;
    const noalign = (apply: (v: VList) => void): AlignNoalign => ({ kind: "noalign", apply });
    if (env === "gather" || env === "equation" || env === "displaymath" || env === "multline") {
      // one column: each line set to the display's width
      const multline = env === "multline";
      const gap = 10 * PT; // \multlinegap, \multlinetaggap
      const rows = lines.filter((l) => l.kind === "row") as Extract<DisplayLine, { kind: "row" }>[];
      // multline: the tag goes on the last line, wherever it was given
      const multlineTag = multline ? rows.find((l) => l.tag)?.tag ?? null : null;
      const sets: { list: Node[]; tag: BoxNode | null }[] = rows.map((l, i) => {
        const cell = l.cells.flat();
        if (!multline) {
          const tag = l.tag ? this.tagBox(l.tag, ctx) : null;
          return { list: [this.amsStrut(a), ...display(this.list(cell, ctx))], tag };
        }
        // \hbox to\totwidth@{\hfil\strut@$\m@th\displaystyle{}##$\hfil}, the first line \hfilneg\hskip\multlinegap,
        // the last \hskip\multlinegap\hfilneg (\hskip\multlinetaggap and the tag when there is one)
        const empty = newNoad(ORD_NOAD, a);
        empty.nucleus = mlistField([]);
        const f = this.cellFormula([{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, empty, ...this.list(cell, ctx)], a);
        const first = i === 0;
        const last = i === rows.length - 1;
        const tag = last && multlineTag ? this.tagBox(multlineTag, ctx) : null;
        const list: Node[] = [first ? newKern(gap) : this.hfil(a), this.amsStrut(a), ...f];
        if (last) list.push(newKern(gap), ...(tag ? [tag] : []));
        else list.push(this.hfil(a));
        if (first && last) list.splice(list.length - (tag ? 2 : 1), 0, this.hfil(a));
        return { list, tag: null };
      });
      let boxes: BoxNode[];
      if (multline) {
        const d = Math.max(this.displayWidth(), ...sets.map((s) => hpack(s.list).width));
        boxes = sets.map((s) => hpack(s.list, d, EXACTLY, a));
      } else boxes = this.gatherLines(sets);
      const v = new VList(this.spreadParams(), a);
      let k = 0;
      for (const l of lines) {
        if (l.kind === "text") {
          this.intertextNoalign(l.atom, ctx, jot).apply(v);
          continue;
        }
        const b = boxes[k++];
        b.tag = undefined;
        v.appendBox(b);
        const g = l === lastRow ? 0 : gapAfter(l.r);
        if (g !== 0) v.append(this.hskip(g, a));
      }
      return [this.positioned(this.displayBox(v), a)];
    }
    // column pairs (align, flalign, alignat) or eqnarray's three columns
    const eqnarray = env === "eqnarray";
    const minalignsep = 10 * PT;
    const acs = scalePermille(5 * PT, this.arrayColSep); // \arraycolsep
    const rows: { line: Extract<DisplayLine, { kind: "row" }>; cells: Node[][] }[] = [];
    for (const l of lines) {
      if (l.kind !== "row") continue;
      const cells = l.cells.map((cell, j): Node[] => {
        if (eqnarray) {
          // \hskip\@centering$\displaystyle{##}$ & \hskip2\arraycolsep\hfil${##}$\hfil & \hskip2\arraycolsep$\displaystyle{##}$\hfil
          if (j === 0) return [this.hfil(a), ...display(this.list(cell, ctx))];
          const f = j === 1 ? this.cellFormula([this.groupNoad(this.list(cell, ctx), a)], a) : display(this.list(cell, ctx));
          return [this.hskip(2 * acs, a), ...(j === 1 ? [this.hfil(a)] : []), ...f, this.hfil(a)];
        }
        // \align@preamble: \hfil\strut@$\m@th\displaystyle{##}$ & $\m@th\displaystyle{{}##}$\hfil
        if (j % 2 === 0) return [this.hfil(a), this.amsStrut(a), ...display(this.list(cell, ctx))];
        const empty = newNoad(ORD_NOAD, a);
        empty.nucleus = mlistField([]);
        return [...display([empty, ...this.list(cell, ctx)]), this.hfil(a)];
      });
      rows.push({ line: l, cells });
    }
    let fields = Math.max(eqnarray ? 3 : 2, ...rows.map((r) => r.cells.length));
    if (!eqnarray && fields % 2 === 1) fields++;
    for (const r of rows) while (r.cells.length < fields) r.cells.push([]);
    const colW = new Array(fields).fill(0);
    for (const r of rows) r.cells.forEach((c, j) => (colW[j] = Math.max(colW[j], hpack(c).width)));
    const tot = colW.reduce((x, y) => x + y, 0);
    const tagBoxes = rows.map((r) => (r.line.tag ? this.tagBox(r.line.tag, ctx) : null));
    const tagged = tagBoxes.some((t) => t);
    const maxTag = Math.max(0, ...tagBoxes.map((t) => t?.width ?? 0));
    const sep = this.minTagSep();
    const w = this.displayWidth();
    // amsmath \measure@: \eqnshift@ and \alignsep@ from \displaywidth - \totwidth@
    const pairs = fields / 2;
    let xat = env === "alignat" ? 0 : env === "align" ? 1 : pairs < 2 ? 0 : 2;
    if (eqnarray) xat = 0;
    let eqnshift = 0;
    let alignsep = 0;
    const room = w - tot - (tagged && xat === 2 ? maxTag + sep : 0);
    if (xat === 0) eqnshift = div(room, 2);
    else if (xat === 1) eqnshift = alignsep = div(room, pairs + 1);
    else alignsep = div(room, pairs - 1);
    if (xat !== 0 && alignsep < minalignsep) {
      alignsep = minalignsep;
      if (eqnshift > 0) eqnshift = div(w - tot - (pairs - 1) * alignsep, 2);
    }
    if (eqnshift < 0 || !tagged) eqnshift = 0;
    const tabskips = [glue(eqnshift)];
    for (let j = 0; j < fields; j++) tabskips.push(glue(!eqnarray && j % 2 === 1 && j < fields - 1 ? alignsep : 0));
    if (tagged) {
      const body = eqnshift + tot + (eqnarray ? 0 : (pairs - 1) * alignsep);
      const d = Math.max(w, body + sep + maxTag + (eqnarray ? eqnshift : 0));
      tabskips[fields] = glue(d - body - maxTag);
      tabskips.push(glue(0));
    }
    let k = 0;
    for (const l of lines) {
      if (l.kind === "text") {
        items.push(this.intertextNoalign(l.atom, ctx, eqnarray ? 0 : jot));
        continue;
      }
      const r = rows[k];
      const tag = tagBoxes[k++];
      const row: AlignRow = { kind: "row", cells: tagged ? [...r.cells, tag ? [this.hfil(a), tag] : []] : r.cells, tag: a };
      if (tag) tags.set(row, tag);
      items.push(row);
      const g = l === lastRow ? 0 : gapAfter(l.r);
      // eqnarray: \cr\noalign{\penalty\@eqpen\vskip\jot ...}
      if (eqnarray && l !== lastRow) items.push(noalign((v) => v.append(this.hskip(jot + g, a))));
      else if (g !== 0) items.push(noalign((v) => v.append(this.hskip(g, a))));
    }
    const v = new VList(eqnarray ? { baselineskip: glue(12 * PT), lineskip: glue(PT), lineskiplimit: 0 } : this.spreadParams(), a);
    halign(items, tabskips, v, a);
    return [this.positioned(this.displayBox(v), a)];
  }

  private displayBox(v: VList): BoxNode {
    const box = vpack(v.items, v.tag);
    box.role = "array";
    return box;
  }

  /** amsmath \ams@start@box: \vtop for [t], \vbox for [b], else \vcenter. */
  private positioned(box: BoxNode, a: AtomLike): Noad {
    const pos = /^\[\s*([tb])\s*\]/.exec(a.environmentArguments ?? "")?.[1];
    if (pos === "t") {
      // \vtop: the reference point moves to the baseline of the first box
      const first = box.list[0];
      const h = first && (first.type === HLIST_NODE || first.type === VLIST_NODE || first.type === RULE_NODE) ? (first as BoxNode).height : 0;
      box.depth = box.depth - h + box.height;
      box.height = h;
    }
    if (pos) return this.boxNoad(box, a);
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(box);
    return vc;
  }

  /**
   * mathtools \MT_start_cases:nnnn (dcases, rcases, cases*, ...):
   * \left#4 \vcenter{\spread@equation \ialign{\strut@#2 & #1\strut@ #3}} \right#5
   */
  private mathtoolsCases(a: AtomLike, env: string, ctx: Ctx): MItem[] {
    const displayStyle = env.startsWith("d");
    const right = env.includes("rcases");
    const text = env.endsWith("*");
    const rows: AtomLike[][][] = a.rows ?? [];
    const items: (AlignRow | AlignNoalign)[] = [];
    const quad = this.textRoman().param(6);
    const formula = (cell: AtomLike[]) => this.cellFormula(displayStyle ? [{ type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, ...this.list(cell, ctx)] : this.list(cell, ctx), a);
    rows.forEach((row, r) => {
      const { cells } = this.rowCells(row);
      if (this.dropsRow(a, cells, r, rows.length)) return;
      const lists = cells.slice(0, 2).map((cell, j): Node[] => {
        if (j === 0) return [this.amsStrut(a), ...formula(cell), this.hfil(a)];
        return [this.hskip(quad, a), this.amsStrut(a), ...(text ? [this.textModeBox(trimText(cell), ctx, a)] : formula(cell)), this.hfil(a)];
      });
      items.push({ kind: "row", cells: lists, tag: a });
      const gap = this.rowGap(a, r);
      if (gap !== 0 && r < rows.length - 1) items.push({ kind: "noalign", apply: (v) => v.append(this.hskip(gap, a)) });
    });
    const v = new VList(this.spreadParams(), a);
    halign(items, [glue(0)], v, a);
    const box = vpack(v.items, a);
    box.role = "array";
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(box);
    return [right ? this.leftRightWrap(a, ".", "\\rbrace", [vc]) : this.leftRightWrap(a, "\\lbrace", ".", [vc])];
  }

  // ------------------------------------------- package commands (VisualTeX) --
  /**
   * Commands the kernel keeps as plain atoms named by their command (its
   * latex-commands/visualtex-commands.ts): mathtools' brackets, cancel's
   * \cancelto, LaTeX's \fbox, \framebox and \hbox, \nobreak, tensor's indices.
   */
  private packageCommand(a: AtomLike, ctx: Ctx, out: MItem[]): boolean {
    const cmd = a.command ?? "";
    switch (cmd) {
      case "\\leftroot":
      case "\\uproot":
        // only meaningful in a root's index (surd), where it is read
        return true;
      case "\\nobreak":
        out.push({ type: 12, penalty: 10000, tag: a } as Node);
        return true;
      case "\\overbracket":
      case "\\underbracket":
        if (a.branch("body") === undefined) return false;
        out.push(this.bracket(a, ctx, cmd === "\\overbracket"));
        return true;
      case "\\cancelto":
        if (a.branch("body") === undefined) return false;
        out.push(this.palette(a, (st) => this.cancelTo(a, ctx, st)));
        return true;
      case "\\fbox":
        if (a.branch("body") === undefined) return false;
        out.push(this.boxNoad(this.frameBox(this.textModeBox(a.branch("body"), ctx, a), 3 * PT, undefined, a), a));
        return true;
      case "\\framebox":
        if (a.branch("body") === undefined) return false;
        out.push(this.boxNoad(this.frameBoxTo(a, ctx), a));
        return true;
      case "\\hbox":
        if (a.branch("body") === undefined) return false;
        out.push(this.boxNoad(this.textModeBox(a.branch("body"), ctx, a), a));
        return true;
      case "\\indices":
      case "\\indices*":
        if (a.branch("body") === undefined) return false;
        this.indices(a, ctx, out, cmd === "\\indices*");
        return true;
      case "\\vtxcearrow":
        out.push(this.chemArrow(a, ctx));
        return true;
      case "\\vtxcellap":
        if (a.branch("body") === undefined) return false;
        out.push(this.chemLlap(a, ctx));
        return true;
      case "\\vtxceminispace":
        out.push(this.chemMinispace(a));
        return true;
      case "\\tensor":
      case "\\tensor*":
        if (a.branch("body") === undefined) return false;
        out.push(this.palette(a, (st) => this.tensor(a, ctx, st, cmd === "\\tensor*")));
        return true;
    }
    return false;
  }

  /** A kernel LatexValue argument (args[i]) in sp, or null when absent or not a dimension. */
  private argDimen(a: AtomLike, i: number): number | null {
    const v = (a.args ?? [])[i] as { dimension?: number; unit?: string } | null | undefined;
    if (!v || typeof v !== "object" || v.dimension === undefined) return null;
    return this.dimenFromAtom({ type: "spacing", width: v, branch: () => undefined });
  }

  /**
   * mathtools \overbracket[#1][#2]{#3}: \mathop{\vbox{\m@th\ialign{##\crcr
   * \downbracketfill{#1}{#2}\crcr\noalign{\kern.2\fontdimen5\textfont2\nointerlineskip}
   * $\hfil\displaystyle{#3}\hfil$\crcr}}}\limits; \underbracket the same
   * upside down in a \vtop. #1 defaults to the height of \braceld, #2 to
   * .7\fontdimen5\textfont2.
   */
  private bracket(a: AtomLike, ctx: Ctx, over: boolean): Noad {
    const xheight = this.env.famFnt(2, 0)!.param(5);
    const ld = this.mathCodeOf("\\braceld", BASE_CTX) ?? 0x37a;
    const ldFont = this.env.famFnt((ld >> 8) & 15, 0)!;
    const t = this.argDimen(a, 0) ?? ldFont.charHeight(ldFont.charInfo(ld & 0xff));
    const h = this.argDimen(a, 1) ?? texScale(".7", xheight);
    const RUNNING = -(2 ** 30);
    const rule = (width: number, height: number, depth: number): Node => ({ type: RULE_NODE, width, height, depth, tag: a });
    // \downbracketend: \vrule width #1 depth #2; \leaders\vrule height #1 depth 0pt\hfill
    // \upbracketend: \vrule height #2 width #1; \leaders\vrule height 0pt depth #1\hfill
    const end = () => (over ? rule(t, RUNNING, h) : rule(t, h, RUNNING));
    const leaders: Node = {
      type: GLUE_NODE,
      spec: glue(0, PT, 2),
      subtype: 100,
      leader: over ? rule(decimalToScaled(".4"), t, 0) : rule(decimalToScaled(".4"), 0, t),
      tag: a,
    } as Node;
    const fill: Node[] = [end(), leaders, end()];
    const content = (): Node[] =>
      this.cellFormula([this.hfil(a), { type: STYLE_NODE, style: DISPLAY_STYLE, tag: a }, this.groupNoad(this.list(a.branch("body"), ctx), a), this.hfil(a)], a);
    const kern = this.kernNoalign(texScale(".2", xheight), a, true);
    const items: (AlignRow | AlignNoalign)[] = over
      ? [{ kind: "row", cells: [fill] }, kern, { kind: "row", cells: [content()] }]
      : [{ kind: "row", cells: [content()] }, kern, { kind: "row", cells: [fill] }];
    const box = this.ialignBox(items, a, !over);
    box.role = over ? "overbrace" : "underbrace";
    const op = newNoad(OP_NOAD, a);
    op.subtype = LIMITS;
    op.nucleus = boxField(box);
    return op;
  }

  /**
   * cancel.sty \cancelto{#1}{#2} (default "smaller" option): \mathchoice of
   * \@cancelto{value style}{#1}{style}{#2}, the expression with an arrow
   * (\vector) from its lower left to the value, set past its upper right.
   */
  private cancelTo(a: AtomLike, ctx: Ctx, style: number): MItem[] {
    const valueStyle = style === DISPLAY_STYLE ? TEXT_STYLE : style === TEXT_STYLE ? SCRIPT_STYLE : SCRIPT_SCRIPT_STYLE;
    // \@begin@tempboxa\hbox{\m@th$#3{#4}$}
    const box0 = this.mathBox(this.styleList(style, a, [this.groupNoad(this.list(a.branch("body"), ctx), a)]), TEXT_STYLE, a);
    let dimen = Math.max(box0.width, 2 * PT);
    let totalheight = Math.max(box0.height + box0.depth, 4 * PT);
    let spec: string, dimb: number;
    if (totalheight < dimen) {
      // wide
      const cnta = Math.trunc((totalheight * 5) / dimen);
      dimb = 3 * PT;
      dimen += ([5, 5, 4, 3][cnta] ?? 2) * PT;
      dimen = Math.max(dimen, 9 * PT) + PT;
      spec = ["5441", "5441", "5421", "4443"][cnta] ?? "3611";
    } else {
      // tall
      totalheight += 3 * PT;
      const cnta = Math.trunc((dimen * 5) / totalheight);
      totalheight += 3 * PT;
      dimen = texScale([".25", ".25", ".5", ".75"][cnta] ?? "1", totalheight);
      dimb = texScale([".8", ".8", "1.2", "1.5"][cnta] ?? "2", PT);
      spec = ["0814", "0814", "1812", "2734"][cnta] ?? "3611";
    }
    // \dimen@\height \setbox\@tempboxa\hbox{$\m@th\vcenter{\box\@tempboxa}$} \advance\dimen@-\height
    const vc = newNoad(VCENTER_NOAD, a);
    vc.nucleus = boxField(vpack([box0], a));
    const tempboxa = this.mathBox([vc], TEXT_STYLE, a);
    const raise = box0.height - tempboxa.height;
    const [n1, n2, n3, n4] = [...spec].map(Number);
    // \canto@vector: \dimen@.5\p@ \setbox\z@\vbox{\boxmaxdepth.5\p@\hbox{\kern-1.2\p@\kern#1\dimen@$#7{#8}\m@th$}}
    const half = PT / 2;
    const valueBox = this.mathBox(this.styleList(valueStyle, a, [this.groupNoad(this.list(a.branch("above"), ctx), a)]), TEXT_STYLE, a);
    const zbox = vpackage([hpack([newKern(-texScale("1.2", PT), a), newKern(n1 * half, a), valueBox], 0, undefined, a)], 0, 1, half, a);
    zbox.width = 0; // \wd\z@\z@ (\canto@fil = \hidewidth)
    // \vcenter{\hbox{\dimen@#6\unitlength \kern\dimen@ \multiply\dimen@#4\divide\dimen@#3 \vrule\@depth\dimen@\@width\z@ \vector(#3,#4){#5}}}
    const arrowParts: Node[] = [newKern(dimb, a), { type: RULE_NODE, width: 0, height: -(2 ** 30), depth: Math.trunc((dimb * n4) / n3), tag: a } as Node];
    arrowParts.push(...this.pictureSvector(n3, n4, dimen, a));
    const arrowVc = newNoad(VCENTER_NOAD, a);
    arrowVc.nucleus = boxField(vpack([hpack(arrowParts, 0, undefined, a)], a));
    // ^{\raise#2\dimen@\copy\z@\kern-\scriptspace}
    const raised: BoxNode = { ...zbox, shift: -n2 * half };
    arrowVc.sup = mlistField([this.boxNoad(raised, a), newKern(-this.env.scriptSpace, a) as unknown as MItem]);
    const hskip = (w: number, stretch = 0, order = NORMAL): Node => ({ type: GLUE_NODE, spec: { width: w, stretch, stretchOrder: order, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a });
    const hideskip = () => hskip(-1000 * PT, 65536, 2);
    const rows: (AlignRow | AlignNoalign)[] = [
      { kind: "row", cells: [[hideskip(), ...this.cellFormula([arrowVc], a), hideskip()]] },
      { kind: "row", cells: [[hskip(0, 65536, 1), tempboxa, newKern(zbox.width, a), hskip(0, 65536, 1)]] },
    ];
    // \ooalign
    const ex = this.textRoman().param(5);
    const v = new VList({ baselineskip: glue(0), lineskip: glue(texScale(".25", ex)), lineskiplimit: -MAX_DIMEN }, a);
    halign(rows, [glue(0)], v, a);
    const vtop = vpack(v.items, a);
    const first = vtop.list[0];
    const h = first && (first.type === HLIST_NODE || first.type === VLIST_NODE || first.type === RULE_NODE) ? (first as BoxNode).height : 0;
    vtop.depth = vtop.depth - h + vtop.height;
    vtop.height = h;
    const outer = hpack([vtop], 0, undefined, a);
    outer.shift = -raise;
    const inner = newNoad(ORD_NOAD, a);
    inner.nucleus = boxField(outer);
    return [this.groupNoad([newKern(0, a) as unknown as MItem, inner], a)];
  }

  /** latex.ltx \vector(x,y){len} for x, y > 0 (\@svector): the \@sline and its arrowhead from line10. */
  private pictureSvector(xarg: number, yarg: number, linelen: number, a: AtomLike): Node[] {
    // \vector: |x| must be below 5, else \@badlinearg (nothing drawn); \@svector: |y| below 5
    if (Math.abs(xarg) >= 5 || Math.abs(yarg) >= 5) return [];
    const { out, clnht, wd, raise } = this.pictureSlineState(xarg, yarg, linelen, a);
    const f = this.cache.get("line10");
    // \@getrarrow(#1,#2)
    const y = Math.abs(yarg);
    let c = y === 0 ? 45 : y === 1 ? (xarg < 3 ? xarg * 24 - 6 : xarg === 3 ? 49 : 58) : y === 2 ? (xarg < 3 ? xarg * 24 - 3 : 51) : y === 3 ? xarg * 16 - 2 : xarg * 16 + 7;
    if (yarg < 0) c += 64;
    out.push({ type: GLUE_NODE, spec: glue(-wd), subtype: NORMAL, tag: a } as Node);
    const head = hpack([newCharacter(f, c, a)], 0, undefined, a);
    head.shift = raise ? -clnht : clnht;
    out.push(head);
    return out;
  }

  /** latex.ltx \@iframebox[#1][#2]{#3}: the text in an \hbox to #1 between \fboxsep kerns, framed. */
  private frameBoxTo(a: AtomLike, ctx: Ctx): BoxNode {
    const content = this.textModeBox(a.branch("body"), ctx, a);
    const width = this.argDimen(a, 0);
    if (width === null) return this.frameBox(content, 3 * PT, undefined, a);
    const pos = String((a.args ?? [])[1] ?? "c").trim() || "c";
    const fboxsep = 3 * PT, fboxrule = decimalToScaled(".4");
    const hss = (): Node => ({ type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 1, shrink: PT, shrinkOrder: 1 }, subtype: NORMAL, tag: a });
    // \bm@c, \bm@l (t), \bm@r (b), \bm@s: \unhbox\@tempboxa with \hss around
    const inner: Node[] = [...content.list];
    const middle = pos === "l" || pos === "t" ? [...inner, hss()] : pos === "r" || pos === "b" ? [hss(), ...inner] : pos === "s" ? inner : [hss(), ...inner, hss()];
    const tempboxa = hpack([newKern(fboxsep, a), ...middle, newKern(fboxsep, a)], width, EXACTLY, a);
    // \@frameb@x{\kern-\fboxrule}
    const tempdima = fboxrule + fboxsep + tempboxa.depth;
    const vglue = (w: number): Node => ({ type: GLUE_NODE, spec: { width: w, stretch: 0, stretchOrder: 0, shrink: 0, shrinkOrder: 0 }, subtype: NORMAL, tag: a });
    const innerV = vpackage([vglue(fboxsep), tempboxa, vglue(fboxsep)], 0, undefined, undefined, a);
    const vrule = (): Node => ({ type: RULE_NODE, width: fboxrule, height: -(2 ** 30), depth: -(2 ** 30), tag: a });
    const hrule = (): Node => ({ type: RULE_NODE, width: -(2 ** 30), height: fboxrule, depth: 0, tag: a });
    const mid = hpack([vrule(), newKern(-fboxrule, a), innerV, newKern(-fboxrule, a), vrule()], 0, undefined, a);
    const frame = vpackage([hrule(), mid, hrule()], 0, undefined, undefined, a);
    const lowered = hpack([frame], 0, undefined, a);
    lowered.shift = tempdima;
    const outer = hpack([lowered], 0, undefined, a);
    outer.role = "boxed";
    return outer;
  }

  /**
   * tensor.sty's index lists: each index joins its own list ({#4}) and, unless
   * starred, a phantom of its width joins the other (\tnsr@Hph); built afresh
   * at every call (the typesetter consumes noads).
   */
  private indexLists(items: AtomLike[] | undefined, ctx: Ctx, star: boolean): { sps: MItem[]; sbs: MItem[] } {
    const sps: MItem[] = [];
    const sbs: MItem[] = [];
    for (const it of items ?? []) {
      if (it.type === "first") continue;
      const sup = it.branch("superscript");
      const sub = it.branch("subscript");
      const atoms = sup ?? sub;
      if (atoms === undefined) continue;
      const [mine, other] = sup !== undefined ? [sps, sbs] : [sbs, sps];
      mine.push(this.groupNoad(this.list(atoms, ctx), it));
      if (!star) {
        // \mathpalette\tnsr@Mph{#4}: \makebox[width of $#1{#4}$][c]{$#1\indexmarker$}
        other.push(
          this.palette(it, (st) => {
            const w = this.mathBox(this.styleList(st, it, [this.groupNoad(this.list(atoms, ctx), it)]), TEXT_STYLE, it).width;
            const b = newNullBox(it);
            b.width = w;
            b.role = "phantom";
            return [this.boxNoad(b, it)];
          }) as unknown as MItem,
        );
      }
    }
    return { sps, sbs };
  }

  /** \tnsr@Fin: ^{sps}_{sbs} on the last noad of `out` (tex.web 1176), both always present. */
  private attachIndexScripts(out: MItem[], sps: MItem[], sbs: MItem[], tag: Tag): void {
    for (const [which, list] of [["sup", sps], ["sub", sbs]] as const) {
      let last = out[out.length - 1] as Noad | undefined;
      if (!last || !(last.type >= ORD_NOAD && last.type < LEFT_NOAD) || last[which].mathType !== EMPTY) {
        last = newNoad(ORD_NOAD, tag);
        out.push(last);
      }
      last[which] = mlistField(list);
    }
  }

  /** tensor.sty \indices{^a_b...}: the staggered indices as scripts of what comes before. */
  private indices(a: AtomLike, ctx: Ctx, out: MItem[], star: boolean): void {
    const { sps, sbs } = this.indexLists(a.branch("body"), ctx, star);
    this.attachIndexScripts(out, sps, sbs, a);
  }

  /**
   * tensor.sty \tensor[pre]{T}{post} in one style (\mathpalette\tnsr@Plt):
   * the prescripts' phantom {}^{..}_{..}, then T with the scripts, the
   * prescripts first, kerned back over T (\tnsr@Krn).
   */
  private tensor(a: AtomLike, ctx: Ctx, style: number, star: boolean): MItem[] {
    const out: MItem[] = [];
    const nucleus = () => this.list(a.branch("body"), ctx);
    const pre = (a.branch("above") ?? []).some((x) => x.type !== "first");
    let sps: MItem[] = [];
    let sbs: MItem[] = [];
    if (pre) {
      // \hphantom{{}\tnsr@Fin}
      const phantomList = (): MItem[] => {
        const l = this.indexLists(a.branch("above"), ctx, star);
        const empty = newNoad(ORD_NOAD, a);
        empty.nucleus = mlistField([]);
        const items: MItem[] = [empty];
        this.attachIndexScripts(items, l.sps, l.sbs, a);
        return items;
      };
      const pb = newNullBox(a);
      pb.width = this.mathBox(this.styleList(style, a, phantomList()), TEXT_STYLE, a).width;
      pb.role = "phantom";
      out.push(this.boxNoad(pb, a));
      // \tnsr@Krn: \kern-(width of $\tnsr@Obj^{#1}\mkern-1mu$) #1 \kern(width of $\tnsr@Obj$)
      const objWidth = this.mathBox(this.styleList(style, a, nucleus()), TEXT_STYLE, a).width;
      const krn = (which: "sps" | "sbs"): MItem[] => {
        const measure = nucleus();
        this.attachIndexScripts(measure, this.indexLists(a.branch("above"), ctx, star)[which], [], a);
        // only the superscript is measured (\tnsr@Obj^{#1}): drop the empty subscript again
        const last = measure[measure.length - 1] as Noad;
        last.sub = { mathType: EMPTY } as Field;
        const w = this.mathBox(this.styleList(style, a, [...measure, this.muKern(-PT, a)]), TEXT_STYLE, a).width;
        return [newKern(-w, a) as unknown as MItem, ...this.indexLists(a.branch("above"), ctx, star)[which], newKern(objWidth, a) as unknown as MItem];
      };
      sps = krn("sps");
      sbs = krn("sbs");
    }
    const post = this.indexLists(a.branch("below"), ctx, star);
    sps.push(...post.sps);
    sbs.push(...post.sbs);
    out.push(...nucleus());
    this.attachIndexScripts(out, sps, sbs, a);
    return out;
  }

  /** \llap{\ensuremath{<style>#1}} in each style (mhchem.sty's \__mhchem_output_coreFive_llap:n). */
  private chemLlap(a: AtomLike, ctx: Ctx): MItem {
    return this.palette(a, (st) => {
      const content = this.mathBox(this.styleList(st, a, this.list(a.branch("body"), ctx)), TEXT_STYLE, a);
      const hss: Node = { type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 1, shrink: PT, shrinkOrder: 1 }, subtype: NORMAL, tag: a };
      return [this.boxNoad(hpack([hss, content], 0, EXACTLY, a), a)];
    });
  }

  /**
   * \mhchem@minispace: \kern(0pt - width of ${}^8_8$ + width of $\text{C}^8_8$
   * - width of $\text{C}^{}_{}$), measured in text style.
   */
  private chemMinispace(a: AtomLike): MItem {
    const eight = () => mlistField([this.mathChar(this.data.meanings.mathcode["8"], BASE_CTX, a)]);
    const measure = (base: () => Noad, scripts: boolean): number => {
      const n = base();
      n.sup = scripts ? eight() : mlistField([]);
      n.sub = scripts ? eight() : mlistField([]);
      return this.mathBox([n], TEXT_STYLE, a).width;
    };
    const empty = () => {
      const n = newNoad(ORD_NOAD, a);
      n.nucleus = mlistField([]);
      return n;
    };
    const c = () => this.boxNoad(this.textBox("C", this.textId(this.text.rm, "m", "n", this.text.sizes[0]), a), a);
    return newKern(-measure(empty, true) + measure(c, true) - measure(c, false), a) as unknown as MItem;
  }

  /**
   * mhchem.sty v4's arrows (\__mhchem_arrow_output:nnn, the font arrows):
   * \__mhchem_arrow_ext:nnnnnnnn, amsmath's \ext@arrow with an arrow at least
   * 2em long, and the equilibrium arrows as two of them raised and lowered.
   */
  private chemArrow(a: AtomLike, ctx: Ctx): Noad {
    const name = String((a.args ?? [])[0] ?? "yields");
    const rm = this.textRoman();
    const em = rm.param(6), ex = rm.param(5);
    const minLength = 2 * em;
    type Label = (() => MItem[]) | null;
    const visible = (x: AtomLike[] | undefined) => (x ?? []).some((y) => y.type !== "first");
    const above: Label = visible(a.branch("above")) ? () => this.list(a.branch("above"), ctx) : null;
    const below: Label = visible(a.branch("below")) ? () => this.list(a.branch("below"), ctx) : null;
    // \hphantom{#} (\mathpalette\mathph@nt), followed by \kern<k>
    const phantom = (label: Label, k = 0): Label =>
      label &&
      (() => {
        const items: MItem[] = [
          this.palette(a, (st) => {
            const b = newNullBox(a);
            b.width = this.mathBox(this.styleList(st, a, label()), TEXT_STYLE, a).width;
            b.role = "phantom";
            return [this.boxNoad(b, a)];
          }),
        ];
        if (k !== 0) items.push(newKern(k, a) as unknown as MItem);
        return items;
      });
    const right: ArrowFill = ["\\relbar", "\\relbar", "\\rightarrow"];
    const left: ArrowFill = ["\\leftarrow", "\\relbar", "\\relbar"];
    const leftright: ArrowFill = ["\\leftarrow", "\\relbar", "\\rightarrow"];
    const harpoonUp: ArrowFill = ["\\relbar", "\\relbar", "\\rightharpoonup"];
    const harpoonDown: ArrowFill = ["\\leftharpoondown", "\\relbar", "\\relbar"];
    // \mathrel{\mathop{\makebox[#7]{#8\displaystyle}}\limits^{\mkern#1mu#3\mkern#2mu}
    //   _{\makebox{\raisebox{1.25ex-\heightof{$\scriptstyle\mkern#4mu#6\mkern#5mu$}}[0pt]{$\scriptstyle\mkern#4mu#6\mkern#5mu$}}}}
    const ext = (k1: number, k2: number, sup: Label, k4: number, k5: number, sub: Label, width: number, fill: ArrowFill): Noad => {
      const hss = (): Node => ({ type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 1, shrink: PT, shrinkOrder: 1 }, subtype: NORMAL, tag: a });
      const box = hpack([hss(), ...this.arrowFill(fill, DISPLAY_STYLE, a), hss()], width, EXACTLY, a);
      box.role = "ext-arrow";
      const op = newNoad(OP_NOAD, a);
      op.subtype = LIMITS;
      op.nucleus = boxField(box);
      if (sup) op.sup = mlistField([this.muKern(k1 * PT, a), ...sup(), this.muKern(k2 * PT, a)]);
      if (sub) {
        const x = this.mathBox([{ type: STYLE_NODE, style: SCRIPT_STYLE, tag: a }, this.muKern(k4 * PT, a), ...sub(), this.muKern(k5 * PT, a)], TEXT_STYLE, a);
        const raised = hpack([{ ...x, shift: -(texScale("1.25", ex) - x.height) }], 0, undefined, a);
        raised.height = 0;
        op.sub = boxField(hpack([raised], 0, undefined, a));
      }
      const rel = newNoad(REL_NOAD, a);
      rel.nucleus = mlistField([op]);
      return rel;
    };
    // \raisebox{<d>}{$...$} and \rlap{...}
    const raise = (factor: string, items: MItem[]): BoxNode => {
      const b = this.mathBox(items, TEXT_STYLE, a);
      return hpack([{ ...b, shift: -texScale(factor, ex) }], 0, undefined, a);
    };
    const lower = (factor: string, items: MItem[]): BoxNode => {
      const b = this.mathBox(items, TEXT_STYLE, a);
      return hpack([{ ...b, shift: texScale(factor, ex) }], 0, undefined, a);
    };
    const rlap = (b: BoxNode): BoxNode =>
      hpack([b, { type: GLUE_NODE, spec: { width: 0, stretch: PT, stretchOrder: 1, shrink: PT, shrinkOrder: 1 }, subtype: NORMAL, tag: a } as Node], 0, EXACTLY, a);
    const pair = (top: BoxNode, bottom: BoxNode): Noad => {
      const rel = newNoad(REL_NOAD, a);
      rel.nucleus = mlistField([this.boxNoad(rlap(top), a), this.boxNoad(bottom, a)]);
      return rel;
    };
    const kern = (w: number) => newKern(w, a) as unknown as MItem;
    switch (name) {
      case "yieldsLeft":
        return ext(9, 5, above, 9, 5, below, minLength, left);
      case "mesomerism":
        return ext(9, 9, above, 9, 9, below, minLength, leftright);
      case "yieldsLeftRight":
        return pair(
          raise(".44", [ext(9, 9, above, 9, 9, below ? phantom(below) : null, minLength, right)]),
          lower(".44", [ext(9, 9, above ? phantom(above) : null, 9, 9, below, minLength, left)]),
        );
      case "equilibrium":
        return pair(
          raise(".22", [below ? ext(9, 9, above, 9, 5, phantom(below), minLength, harpoonUp) : ext(5, 9, above, 9, 9, null, minLength, harpoonUp)]),
          lower(".22", [above ? ext(5, 9, phantom(above), 9, 9, below, minLength, harpoonDown) : ext(9, 9, null, 9, 5, below, minLength, harpoonDown)]),
        );
      case "equilibriumRight":
        if (!above)
          return pair(
            lower(".22", [kern(texScale(".5", em)), ext(0, 0, null, 9, 5, below, minLength - em, harpoonDown)]),
            raise(".22", [below ? ext(0, 0, null, 9, 5, phantom(below, em), minLength, harpoonUp) : ext(0, 0, null, 0, 0, null, minLength, harpoonUp)]),
          );
        return pair(
          lower(".22", [kern(texScale(".5", em)), ext(5, 9, phantom(above, -em), 9, 9, below, minLength - em, harpoonDown)]),
          raise(".22", [below ? ext(9, 9, above, 9, 9, phantom(below, em), minLength, harpoonUp) : ext(5, 9, above, 0, 0, null, minLength, harpoonUp)]),
        );
      case "equilibriumLeft":
        if (!below)
          return pair(
            raise(".22", [kern(texScale(".5", em)), ext(5, 9, above, 0, 0, null, minLength - em, harpoonUp)]),
            lower(".22", [above ? ext(5, 9, phantom(above, em), 0, 0, null, minLength, harpoonDown) : ext(0, 0, null, 0, 0, null, minLength, harpoonDown)]),
          );
        return pair(
          raise(".22", [kern(texScale(".5", em)), ext(9, 9, above, 5, 9, phantom(below, -em), minLength - em, harpoonUp)]),
          lower(".22", [above ? ext(9, 9, phantom(above, em), 9, 9, below, minLength, harpoonDown) : ext(0, 0, null, 9, 5, below, minLength, harpoonDown)]),
        );
      default:
        return ext(5, 9, above, 5, 9, below, minLength, right);
    }
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
