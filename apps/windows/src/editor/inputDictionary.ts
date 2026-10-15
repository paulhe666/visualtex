// The one dictionary behind the input judgment (inputJudgment.ts,
// docs/upright-typography-plan-20261015.md): every typed word or symbol
// sequence that turns into something else, and every fixed name that is set
// upright. Quick escapes (alpha -> \alpha) and upright typography (sin, km,
// e, d) are decided by the same routine from this one table.

// ---------------------------------------------------------------------------
// Upright: units
// ---------------------------------------------------------------------------

// SI prefixes typed as plain letters (micro is \mu and is not a letter run).
const SI_PREFIXES = [
  "Q", "R", "Y", "Z", "E", "P", "T", "G", "M", "k", "h", "da",
  "d", "c", "m", "n", "p", "f", "a", "z", "y", "r", "q",
] as const;

// Units that take SI prefixes: base and derived SI units with letter
// symbols, plus accepted units commonly prefixed (L, eV, Da, bar, Wh, cal).
const PREFIXABLE_UNITS = [
  "m", "g", "s", "A", "K", "mol", "cd",
  "Hz", "N", "Pa", "J", "W", "C", "V", "F", "S", "Wb", "T", "H",
  "lm", "lx", "Bq", "Gy", "Sv", "kat", "rad",
  "L", "l", "eV", "Da", "bar", "Wh", "cal",
] as const;

// Units used without prefixes: accepted non-SI units (SI Brochure table 8)
// and a few common ones (atm, mmHg, r for revolutions as in r/min).
const PLAIN_UNITS = [
  "sr", "min", "h", "d", "au", "ha", "t", "Np", "B", "dB",
  "atm", "mmHg", "r",
] as const;

// Two-letter unit spellings that are practically never a product of two
// variables. Every other two-letter (and every one-letter) unit is ambiguous.
const CERTAIN_SHORT_UNITS = new Set([
  "km", "kg", "cm", "mm", "Hz",
  "kJ", "kW", "kN", "kV", "kA",
  "MW", "MJ", "MV", "GW", "GJ", "mL",
]);

// Longer unit spellings that are also ordinary words or other fixed names.
const AMBIGUOUS_LONG_UNITS = new Set(["min", "bar", "kat", "dam", "cal"]);

function buildUnits() {
  const units = new Set<string>([...PREFIXABLE_UNITS, ...PLAIN_UNITS]);
  for (const prefix of SI_PREFIXES) {
    for (const unit of PREFIXABLE_UNITS) units.add(prefix + unit);
  }
  return units;
}

export const UNIT_SYMBOLS: ReadonlySet<string> = buildUnits();

export function isUnitSymbol(text: string) {
  return UNIT_SYMBOLS.has(text);
}

export function isCertainUnit(text: string) {
  if (!UNIT_SYMBOLS.has(text)) return false;
  // dB, dm, dmol: a leading d is also the differential (\frac{dB}{dt}).
  if (text.startsWith("d")) return false;
  if (text.length <= 2) return CERTAIN_SHORT_UNITS.has(text);
  return !AMBIGUOUS_LONG_UNITS.has(text);
}

// MathLive's former unit shortcuts typed these as \operatorname{km};
// documents from before the input judgment carry that form (sometimes
// wrapped as \operatorname{\mathrm{km}}). They are rewritten to \mathrm{km}.
export const LEGACY_OPERATOR_UNITS: ReadonlySet<string> = new Set([
  "mm", "cm", "km", "kg",
]);

// ---------------------------------------------------------------------------
// Upright: functions
// ---------------------------------------------------------------------------

export interface FixedFunction {
  latex: string;
  /** Ambiguous names (lg = l·g, Re = Reynolds number) need an argument after them. */
  certain: boolean;
}

const BUILT_IN_FUNCTIONS = [
  "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
  "sinh", "cosh", "tanh", "coth", "exp", "ln", "log", "lim", "limsup",
  "liminf", "max", "min", "sup", "inf", "det", "dim", "ker", "gcd", "arg",
  "deg", "hom",
] as const;

const OPERATORNAME_FUNCTIONS = [
  "lcm", "rank", "diag", "sgn", "erf", "erfc", "arccot", "sech", "csch",
  "arsinh", "arcosh", "artanh", "arcoth", "arcsec", "arccsc", "cosec",
] as const;

const AMBIGUOUS_FUNCTIONS: Record<string, string> = {
  lg: "\\lg",
  lb: "\\operatorname{lb}",
  tr: "\\operatorname{tr}",
  Tr: "\\operatorname{Tr}",
  Re: "\\operatorname{Re}",
  Im: "\\operatorname{Im}",
  Pr: "\\Pr",
};

export const FIXED_FUNCTIONS: ReadonlyMap<string, FixedFunction> = new Map<string, FixedFunction>([
  ...BUILT_IN_FUNCTIONS.map((name): [string, FixedFunction] => [name, { latex: `\\${name}`, certain: true }]),
  ...OPERATORNAME_FUNCTIONS.map(
    (name): [string, FixedFunction] => [name, { latex: `\\operatorname{${name}}`, certain: true }],
  ),
  ...Object.entries(AMBIGUOUS_FUNCTIONS).map(
    ([name, latex]): [string, FixedFunction] => [name, { latex, certain: false }],
  ),
  ["argmin", { latex: "\\operatorname*{arg\\,min}", certain: true }],
  ["argmax", { latex: "\\operatorname*{arg\\,max}", certain: true }],
]);

// ---------------------------------------------------------------------------
// Upright: labels and constants
// ---------------------------------------------------------------------------

// Words used as descriptive subscripts (v_max, E_kin, T_eff). Only a
// subscript that is exactly one of these words is a label; single letters
// stay italic (x_i is an index).
export const SUBSCRIPT_LABELS: ReadonlySet<string> = new Set([
  "max", "min", "eff", "tot", "avg", "mean", "in", "out", "ext", "int",
  "kin", "pot", "rel", "ref", "crit", "sat", "abs", "app", "obs", "exp",
  "th", "calc", "std", "rms", "eq", "ini", "fin", "net", "opt",
]);

// Labels that are fixed names anywhere (IUPAC prints pH upright).
export const CERTAIN_LABELS: ReadonlySet<string> = new Set(["pH", "pOH"]);

// Typed doubled letters for the constants (MathLive convention).
export const TYPED_CONSTANTS: Readonly<Record<string, string>> = {
  ee: "\\mathrm{e}",
  ii: "\\mathrm{i}",
  jj: "\\mathrm{j}",
};

// ---------------------------------------------------------------------------
// Quick escapes: typed words
// ---------------------------------------------------------------------------

const GREEK_NAMES = [
  "alpha", "beta", "gamma", "delta", "epsilon", "varepsilon", "zeta", "eta",
  "theta", "vartheta", "iota", "kappa", "lambda", "mu", "nu", "xi", "pi",
  "varpi", "rho", "varrho", "sigma", "varsigma", "tau", "upsilon", "phi",
  "varphi", "chi", "psi", "omega", "Gamma", "Delta", "Theta", "Lambda", "Xi",
  "Pi", "Sigma", "Upsilon", "Phi", "Psi", "Omega",
] as const;

export const GREEK_WORDS: Readonly<Record<string, string>> = Object.fromEntries(
  GREEK_NAMES.map((name) => [name, `\\${name}`]),
);

// Symbols, relations, logic, sets and arrows typed by name. Only whole
// words convert, so these never fire inside spin, point or for. Removed
// (2026-10-15): two-letter products and English words (ge le ne ll gg pp ss
// mm dd nn uu vv AA EE TT ch lt gt or and not of del oo sub) and MathLive's
// calculator names (mean median fft bessel randomReal randomInteger).
export const SYMBOL_WORDS: Readonly<Record<string, string>> = {
  infty: "\\infty", infinity: "\\infty", partial: "\\partial",
  nabla: "\\nabla", grad: "\\nabla", ell: "\\ell", hbar: "\\hbar",
  angle: "\\angle", degree: "\\degree", therefore: "\\therefore",
  because: "\\because", forall: "\\forall", exists: "\\exists",
  nexists: "\\nexists", emptyset: "\\emptyset", aleph: "\\aleph",
  land: "\\land", lor: "\\lor", neg: "\\neg", implies: "\\implies",
  iff: "\\iff", cup: "\\cup", cap: "\\cap", union: "\\cup",
  intersection: "\\cap", setminus: "\\setminus", ldots: "\\ldots",
  cdots: "\\cdots", vdots: "\\vdots", ddots: "\\ddots",
  NN: "\\mathbb{N}", ZZ: "\\mathbb{Z}", QQ: "\\mathbb{Q}", RR: "\\mathbb{R}",
  CC: "\\mathbb{C}",
  plus: "+", add: "+", minus: "-", subtract: "-", times: "\\times",
  mul: "\\times", multiply: "\\times", div: "\\div", divide: "\\div",
  eq: "=", equal: "=", pm: "\\pm", plusminus: "\\pm", mp: "\\mp",
  minusplus: "\\mp", cdot: "\\cdot", ast: "\\ast", asterisk: "\\ast",
  star: "\\star", circ: "\\circ", bullet: "\\bullet",
  geq: "\\geq", gte: "\\geq", leq: "\\leq", lte: "\\leq", neq: "\\neq",
  propto: "\\propto", prop: "\\propto", approx: "\\approx",
  equiv: "\\equiv", sim: "\\sim", simeq: "\\simeq", cong: "\\cong",
  doteq: "\\doteq", prec: "\\prec", preceq: "\\preceq", succ: "\\succ",
  succeq: "\\succeq", in: "\\in", notin: "\\notin", subset: "\\subset",
  subseteq: "\\subseteq", sube: "\\subseteq", supset: "\\supset",
  supseteq: "\\supseteq", supe: "\\supseteq", parallel: "\\parallel",
  perp: "\\perp", mid: "\\mid", nmid: "\\nmid",
  to: "\\to", approaches: "\\to", mapsto: "\\mapsto", uparrow: "\\uparrow",
  downarrow: "\\downarrow", rarr: "\\rightarrow", larr: "\\leftarrow",
  harr: "\\leftrightarrow", uarr: "\\uparrow", darr: "\\downarrow",
  rArr: "\\Rightarrow", lArr: "\\Leftarrow", hArr: "\\Leftrightarrow",
  diamond: "\\diamond", square: "\\square",
  // Big operators insert bare (2026-10-15): limits are typed with _ and ^,
  // the versions with limit boxes stay in the toolbar and candidate list.
  sum: "\\sum", prod: "\\prod", int: "\\int", iint: "\\iint",
  iiint: "\\iiint", oint: "\\oint", mod: "\\bmod",
};

// Templates: the caret goes into the first slot.
export const TEMPLATE_WORDS: Readonly<Record<string, string>> = {
  frac: "\\frac{\\placeholder{}}{\\placeholder{}}",
  dfrac: "\\dfrac{\\placeholder{}}{\\placeholder{}}",
  tfrac: "\\tfrac{\\placeholder{}}{\\placeholder{}}",
  binom: "\\binom{\\placeholder{}}{\\placeholder{}}",
  sqrt: "\\sqrt{\\placeholder{}}",
  cbrt: "\\sqrt[3]{\\placeholder{}}",
  nthroot: "\\sqrt[\\placeholder{}]{\\placeholder{}}",
  ceil: "\\left\\lceil\\placeholder{}\\right\\rceil",
  floor: "\\left\\lfloor\\placeholder{}\\right\\rfloor",
  defint: "\\int_{\\placeholder{}}^{\\placeholder{}}",
  ...Object.fromEntries(
    [
      "acute", "grave", "hat", "widehat", "bar", "overline", "vec", "tilde",
      "widetilde", "dot", "ddot", "dddot", "breve", "check", "mathring",
    ].map((accent) => [accent, `\\${accent}{\\placeholder{}}`]),
  ),
};

/** deg after a number: 30 deg -> 30°. */
export const AFTER_NUMBER_WORDS: Readonly<Record<string, string>> = {
  deg: "\\degree",
};

// ---------------------------------------------------------------------------
// Quick escapes: symbol sequences (no letters), converted as they are typed
// ---------------------------------------------------------------------------

export const SYMBOL_SEQUENCES: Readonly<Record<string, string>> = {
  "&": "\\&", "%": "\\%", "$": "\\$",
  "''": "^{\\doubleprime}", "'''": "^{\\prime\\prime\\prime}",
  "''''": "^{\\prime\\prime\\prime\\prime}",
  "∞": "\\infty", "∑": "\\sum", "∆": "\\differentialD",
  "∂": "\\differentialD",
  "&&": "\\land", "+-": "\\pm", "-+": "\\mp",
  "≠": "\\ne", "!=": "\\ne", "≥": "\\ge", ">=": "\\ge",
  "≤": "\\le", "<=": "\\le", "<<": "\\ll", ">>": "\\gg",
  "~~": "\\approx", "~=": "\\approx", "===": "\\equiv", "-=": "\\equiv",
  "≈": "\\approx", "?=": "\\questeq", "÷": "\\div", "¬": "\\neg",
  ":=": "\\coloneq", "::": "\\Colon", "(:": "\\langle", ":)": "\\rangle",
  "µ": "\\mu", "Ω": "\\omega", ":.": "\\therefore",
  "...": "\\ldots", "+...": "+\\cdots", "-...": "-\\cdots",
  "->...": "\\to\\cdots", "-->...": "\\longrightarrow\\cdots",
  "->": "\\to", "<-": "\\leftarrow", "<->": "\\leftrightarrow",
  "|->": "\\mapsto", "-->": "\\longrightarrow", "<--": "\\longleftarrow",
  "=>": "\\Rightarrow", "==>": "\\Longrightarrow", "<==": "\\Longleftarrow",
  "<=>": "\\iff", "<==>": "\\Longleftrightarrow",
  "(.)": "\\odot", "(+)": "\\oplus", "(/)": "\\oslash", "(*)": "\\otimes",
  "(-)": "\\ominus", "||": "\\Vert", "*": "\\cdot", "**": "\\star",
  "***": "\\ast", "//": "\\slash", "\\\\": "\\backslash",
  "|><": "\\ltimes", "><|": "\\rtimes", "|><|": "\\bowtie", "-:": "\\div",
  "@": "\\circ", "^^": "\\wedge", "^^^": "\\bigwedge",
  "-<": "\\prec", "-<=": "\\preceq", ">-=": "\\succeq",
  "_|_": "\\bot", "|--": "\\vdash", "|==": "\\models",
  "|__": "\\lfloor", "__|": "\\rfloor", "|~": "\\lceil", "~|": "\\rceil",
  ">->": "\\rightarrowtail", "->>": "\\twoheadrightarrow",
  ">->>": "\\twoheadrightarrowtail",
};

// ---------------------------------------------------------------------------
// The mappings page in the input settings lists the dictionary in groups.
// ---------------------------------------------------------------------------

export type InputShortcutDefinition = string | { value: string; after?: string };

export interface InputShortcutGroup {
  id: string;
  titleZh: string;
  titleEn: string;
  shortcuts: Readonly<Record<string, InputShortcutDefinition>>;
}

const FUNCTION_WORDS: Record<string, string> = Object.fromEntries(
  [...FIXED_FUNCTIONS].map(([name, entry]) => [name, entry.latex]),
);

export const INPUT_SHORTCUT_GROUPS: readonly InputShortcutGroup[] = [
  { id: "greek", titleZh: "希腊字母", titleEn: "Greek letters", shortcuts: GREEK_WORDS },
  { id: "functions", titleZh: "函数名（正体）", titleEn: "Function names (upright)", shortcuts: FUNCTION_WORDS },
  { id: "symbols", titleZh: "符号与关系", titleEn: "Symbols and relations", shortcuts: SYMBOL_WORDS },
  { id: "templates", titleZh: "结构", titleEn: "Structures", shortcuts: TEMPLATE_WORDS },
  {
    id: "constants",
    titleZh: "常数与单位",
    titleEn: "Constants and units",
    shortcuts: {
      ...TYPED_CONSTANTS,
      ...Object.fromEntries(
        Object.entries(AFTER_NUMBER_WORDS).map(([word, latex]) => [word, { value: latex, after: "digit+space" }]),
      ),
    },
  },
  { id: "sequences", titleZh: "符号组合", titleEn: "Symbol sequences", shortcuts: SYMBOL_SEQUENCES },
];
