import { normalizeExtendedIntegralLatexCommands } from "../math/extendedIntegralCompatibility.ts";

const chineseChar = /[\u3400-\u9fff\uf900-\ufaff，。；：！？、（）【】《》“”‘’]/;

const protectedTypographyCommands = new Set([
  "text",
  "textnormal",
  "textrm",
  "mathrm",
  "operatorname",
  "mbox",
]);

const uprightOperators: Record<string, string> = {
  arcsin: "\\arcsin",
  arccos: "\\arccos",
  arctan: "\\arctan",
  sinh: "\\sinh",
  cosh: "\\cosh",
  tanh: "\\tanh",
  coth: "\\coth",
  limsup: "\\limsup",
  liminf: "\\liminf",
  sin: "\\sin",
  cos: "\\cos",
  tan: "\\tan",
  cot: "\\cot",
  sec: "\\sec",
  csc: "\\csc",
  exp: "\\exp",
  log: "\\log",
  ln: "\\ln",
  lg: "\\lg",
  lim: "\\lim",
  max: "\\max",
  min: "\\min",
  sup: "\\sup",
  inf: "\\inf",
  det: "\\det",
  dim: "\\dim",
  ker: "\\ker",
  gcd: "\\gcd",
  lcm: "\\operatorname{lcm}",
  mod: "\\mod",
  rank: "\\operatorname{rank}",
  tr: "\\operatorname{tr}",
  diag: "\\operatorname{diag}",
  sgn: "\\operatorname{sgn}",
  erf: "\\operatorname{erf}",
  erfc: "\\operatorname{erfc}",
  Re: "\\operatorname{Re}",
  Im: "\\operatorname{Im}",
};

const uprightOperatorPattern = new RegExp(
  `(^|[^\\\\A-Za-z])(${Object.keys(uprightOperators)
    .sort((left, right) => right.length - left.length)
    .join("|")})(?=$|[^A-Za-z])`,
  "g",
);

const MATHLIVE_CANONICAL_UPRIGHT_COMMANDS: ReadonlyArray<
  readonly [command: string, standardLatex: string]
> = [
  ["capitalDifferentialD", "\\mathrm{D}"],
  ["differentialD", "\\mathrm{d}"],
  ["exponentialE", "\\mathrm{e}"],
  ["imaginaryI", "\\mathrm{i}"],
  ["imaginaryJ", "\\mathrm{j}"],
];

export type VisualTexInlineShortcutDefinition =
  | string
  | { value: string; after?: string };

export type VisualTexInlineShortcutDefinitions = Record<
  string,
  VisualTexInlineShortcutDefinition
>;

const UPRIGHT_SHORTCUT_AFTER =
  "nothing+function+frac+surd+binop+relop+punct+array+openfence+closefence+space+text";

const GREEK_DIFFERENTIAL_VARIABLES = {
  alpha: "\\alpha",
  beta: "\\beta",
  gamma: "\\gamma",
  delta: "\\delta",
  theta: "\\theta",
  phi: "\\phi",
  varphi: "\\varphi",
  psi: "\\psi",
  omega: "\\omega",
  rho: "\\rho",
  sigma: "\\sigma",
  tau: "\\tau",
  mu: "\\mu",
  nu: "\\nu",
  xi: "\\xi",
  eta: "\\eta",
  zeta: "\\zeta",
  kappa: "\\kappa",
  lambda: "\\lambda",
  chi: "\\chi",
} as const;

const LATIN_DIFFERENTIAL_VARIABLES = ["x", "y", "t"] as const;

export const visualTexUprightInlineShortcuts: VisualTexInlineShortcutDefinitions =
  Object.fromEntries([
    ...LATIN_DIFFERENTIAL_VARIABLES.map((variable) => [
      `d${variable}`,
      {
        after: UPRIGHT_SHORTCUT_AFTER,
        value: `\\mathrm{d}${variable}`,
      },
    ]),
    ...Object.entries(GREEK_DIFFERENTIAL_VARIABLES).map(
      ([name, variableLatex]) => [
        `d${name}`,
        {
          after: UPRIGHT_SHORTCUT_AFTER,
          value: `\\mathrm{d}${variableLatex}`,
        },
      ],
    ),
  ]);

export const GREEK_INLINE_SHORTCUTS: VisualTexInlineShortcutDefinitions = {
  alpha: "\\alpha",
  beta: "\\beta",
  gamma: "\\gamma",
  delta: "\\delta",
  epsilon: "\\epsilon",
  varepsilon: "\\varepsilon",
  zeta: "\\zeta",
  eta: "\\eta",
  theta: "\\theta",
  vartheta: "\\vartheta",
  iota: "\\iota",
  kappa: "\\kappa",
  lambda: "\\lambda",
  mu: "\\mu",
  nu: "\\nu",
  xi: "\\xi",
  pi: "\\pi",
  varpi: "\\varpi",
  rho: "\\rho",
  varrho: "\\varrho",
  sigma: "\\sigma",
  varsigma: "\\varsigma",
  tau: "\\tau",
  upsilon: "\\upsilon",
  phi: "\\phi",
  varphi: "\\varphi",
  chi: "\\chi",
  psi: "\\psi",
  omega: "\\omega",
  Gamma: "\\Gamma",
  Delta: "\\Delta",
  Theta: "\\Theta",
  Lambda: "\\Lambda",
  Xi: "\\Xi",
  Pi: "\\Pi",
  Sigma: "\\Sigma",
  Upsilon: "\\Upsilon",
  Phi: "\\Phi",
  Psi: "\\Psi",
  Omega: "\\Omega",
};

export const BASIC_OPERATOR_INLINE_SHORTCUTS: VisualTexInlineShortcutDefinitions = {
  pp: "+",
  plus: "+",
  add: "+",
  ss: "-",
  minus: "-",
  subtract: "-",
  mm: "\\times",
  mul: "\\times",
  multiply: "\\times",
  times: "\\times",
  dd: "\\div",
  div: "\\div",
  divide: "\\div",
  eq: "=",
  equal: "=",
  pm: "\\pm",
  plusminus: "\\pm",
  mp: "\\mp",
  minusplus: "\\mp",
  cdot: "\\cdot",
  ast: "\\ast",
  star: "\\star",
  circ: "\\circ",
  bullet: "\\bullet",
};

export const RELATION_INLINE_SHORTCUTS: VisualTexInlineShortcutDefinitions = {
  ">=": "\\ge",
  "<=": "\\le",
  "!=": "\\ne",
  "~=": "\\approx",
  "~~": "\\approx",
  "===": "\\equiv",
  ge: "\\ge",
  geq: "\\geq",
  gte: "\\geq",
  le: "\\le",
  leq: "\\leq",
  lte: "\\leq",
  ne: "\\ne",
  neq: "\\neq",
  gt: ">",
  lt: "<",
  gg: "\\gg",
  ll: "\\ll",
  propto: "\\propto",
  prop: "\\propto",
  approx: "\\approx",
  equiv: "\\equiv",
  sim: "\\sim",
  simeq: "\\simeq",
  cong: "\\cong",
  doteq: "\\doteq",
  prec: "\\prec",
  preceq: "\\preceq",
  succ: "\\succ",
  succeq: "\\succeq",
  notin: "\\notin",
  subset: "\\subset",
  subseteq: "\\subseteq",
  supset: "\\supset",
  supseteq: "\\supseteq",
  parallel: "\\parallel",
  perp: "\\perp",
  mid: "\\mid",
  nmid: "\\nmid",
};

export const ARROW_INLINE_SHORTCUTS: VisualTexInlineShortcutDefinitions = {
  "->": "\\to",
  "<-": "\\leftarrow",
  "<->": "\\leftrightarrow",
  "=>": "\\Rightarrow",
  "==>": "\\Longrightarrow",
  "<==": "\\Longleftarrow",
  "<==>": "\\Longleftrightarrow",
  mapsto: "\\mapsto",
  uparrow: "\\uparrow",
  downarrow: "\\downarrow",
};

export const ACCENT_INLINE_SHORTCUTS: VisualTexInlineShortcutDefinitions = {
  acute: "\\acute{#?}",
  grave: "\\grave{#?}",
  hat: "\\hat{#?}",
  widehat: "\\widehat{#?}",
  bar: "\\bar{#?}",
  overline: "\\overline{#?}",
  vec: "\\vec{#?}",
  tilde: "\\tilde{#?}",
  widetilde: "\\widetilde{#?}",
  dot: "\\dot{#?}",
  ddot: "\\ddot{#?}",
  dddot: "\\dddot{#?}",
  breve: "\\breve{#?}",
  check: "\\check{#?}",
  mathring: "\\mathring{#?}",
};

export const COMMON_COMMAND_INLINE_SHORTCUTS: VisualTexInlineShortcutDefinitions = {
  frac: "\\frac{#?}{#?}",
  dfrac: "\\dfrac{#?}{#?}",
  tfrac: "\\tfrac{#?}{#?}",
  sqrt: "\\sqrt{#?}",
  binom: "\\binom{#?}{#?}",
  sum: "\\sum_{#?}^{#?}",
  prod: "\\prod_{#?}^{#?}",
  int: "\\int_{#?}^{#?}",
  iint: "\\iint_{#?}^{#?}",
  iiint: "\\iiint_{#?}^{#?}",
  oint: "\\oint_{#?}^{#?}",
  lim: "\\lim_{#?\\to#?}",
  sin: "\\sin",
  cos: "\\cos",
  tan: "\\tan",
  cot: "\\cot",
  sec: "\\sec",
  csc: "\\csc",
  arcsin: "\\arcsin",
  arccos: "\\arccos",
  arctan: "\\arctan",
  sinh: "\\sinh",
  cosh: "\\cosh",
  tanh: "\\tanh",
  log: "\\log",
  ln: "\\ln",
  exp: "\\exp",
  max: "\\max",
  min: "\\min",
  det: "\\det",
  gcd: "\\gcd",
  infty: "\\infty",
  partial: "\\partial",
  nabla: "\\nabla",
  ell: "\\ell",
  hbar: "\\hbar",
  angle: "\\angle",
  degree: "\\degree",
  therefore: "\\therefore",
  because: "\\because",
  forall: "\\forall",
  exists: "\\exists",
  nexists: "\\nexists",
  land: "\\land",
  lor: "\\lor",
  neg: "\\neg",
  implies: "\\implies",
  iff: "\\iff",
  NN: "\\mathbb{N}",
  ZZ: "\\mathbb{Z}",
  QQ: "\\mathbb{Q}",
  RR: "\\mathbb{R}",
  CC: "\\mathbb{C}",
  "+-": "\\pm",
  "-+": "\\mp",
  cup: "\\cup",
  cap: "\\cap",
  union: "\\cup",
  intersection: "\\cap",
  emptyset: "\\emptyset",
  setminus: "\\setminus",
  ldots: "\\ldots",
  cdots: "\\cdots",
  vdots: "\\vdots",
  ddots: "\\ddots",
};

export interface VisualTexAutoEscapeShortcutGroup {
  id: string;
  titleZh: string;
  titleEn: string;
  shortcuts: VisualTexInlineShortcutDefinitions;
}

export const visualTexAutoEscapeShortcutGroups: readonly VisualTexAutoEscapeShortcutGroup[] = [
  {
    id: "greek",
    titleZh: "希腊字母",
    titleEn: "Greek letters",
    shortcuts: GREEK_INLINE_SHORTCUTS,
  },
  {
    id: "operators",
    titleZh: "基本运算",
    titleEn: "Basic operators",
    shortcuts: BASIC_OPERATOR_INLINE_SHORTCUTS,
  },
  {
    id: "relations",
    titleZh: "关系与集合",
    titleEn: "Relations and sets",
    shortcuts: RELATION_INLINE_SHORTCUTS,
  },
  {
    id: "arrows",
    titleZh: "箭头",
    titleEn: "Arrows",
    shortcuts: ARROW_INLINE_SHORTCUTS,
  },
  {
    id: "accents",
    titleZh: "重音结构",
    titleEn: "Accents",
    shortcuts: ACCENT_INLINE_SHORTCUTS,
  },
  {
    id: "commands",
    titleZh: "常用命令",
    titleEn: "Common commands",
    shortcuts: COMMON_COMMAND_INLINE_SHORTCUTS,
  },
  {
    id: "differentials",
    titleZh: "微分变量",
    titleEn: "Differentials",
    shortcuts: visualTexUprightInlineShortcuts,
  },
];

export const visualTexAutoEscapeInlineShortcuts: VisualTexInlineShortcutDefinitions = {
  ...GREEK_INLINE_SHORTCUTS,
  ...BASIC_OPERATOR_INLINE_SHORTCUTS,
  ...RELATION_INLINE_SHORTCUTS,
  ...ARROW_INLINE_SHORTCUTS,
  ...ACCENT_INLINE_SHORTCUTS,
  ...COMMON_COMMAND_INLINE_SHORTCUTS,
  ...visualTexUprightInlineShortcuts,
};

const DISABLED_AUTO_ESCAPE_SHORTCUT_KEYS = new Set([
  "xx",
  "th",
  "mathrm",
  "mathit",
  "mathbf",
  "mathnormal",
  "boldsymbol",
  "mathbb",
  "mathcal",
  "mathscr",
  "mathfrak",
  "mathsf",
  "mathtt",
  "mathbfit",
  "mathsfit",
  "mathsfbf",
  "mathsfbfit",
  "symbf",
  "symbb",
  "symcal",
  "symfrak",
  "symit",
  "symnormal",
  "symsf",
  "symtt",
]);

export function resolveVisualTexInlineShortcuts(
  mathLiveDefaults: Readonly<VisualTexInlineShortcutDefinitions>,
  enabled: boolean,
): VisualTexInlineShortcutDefinitions {
  if (!enabled) return {};
  const safeMathLiveDefaults = Object.fromEntries(
    Object.entries(mathLiveDefaults).filter(
      ([shortcut]) => !DISABLED_AUTO_ESCAPE_SHORTCUT_KEYS.has(shortcut),
    ),
  );
  return {
    ...safeMathLiveDefaults,
    ...visualTexAutoEscapeInlineShortcuts,
  };
}

function transformOutsideProtectedCommands(
  source: string,
  transform: (value: string) => string,
): string {
  let result = "";
  let chunkStart = 0;
  let index = 0;

  while (index < source.length) {
    if (source[index] !== "\\") {
      index += 1;
      continue;
    }

    const commandMatch = /^\\([A-Za-z]+)\{/.exec(source.slice(index));
    if (!commandMatch || !protectedTypographyCommands.has(commandMatch[1])) {
      index += 1;
      continue;
    }

    result += transform(source.slice(chunkStart, index));
    const end = readBracedCommand(source, index);
    result += source.slice(index, end);
    index = end;
    chunkStart = end;
  }

  return result + transform(source.slice(chunkStart));
}

function normalizeNamedOperators(source: string): string {
  return source.replace(
    uprightOperatorPattern,
    (_match, prefix: string, operator: string) =>
      `${prefix}${uprightOperators[operator]}`,
  );
}

function normalizeEulerConstant(source: string): string {
  let result = "";
  let index = 0;
  while (index < source.length) {
    if (source[index] !== "e") {
      result += source[index];
      index += 1;
      continue;
    }

    const previous = source[index - 1];
    if (previous === "\\" || /[A-Za-z]/.test(previous ?? "")) {
      result += source[index];
      index += 1;
      continue;
    }

    let next = index + 1;
    while (/\s/.test(source[next] ?? "")) next += 1;
    if (source[next] !== "^") {
      result += source[index];
      index += 1;
      continue;
    }

    result += "\\mathrm{e}";
    index += 1;
  }
  return result;
}

function normalizeImaginaryUnitInExponentials(source: string): string {
  return source
    .replace(
      /(\\mathrm\{e\}\s*\^\s*\{\s*)i(?=$|[^A-Za-z])/g,
      "$1\\mathrm{i}",
    )
    .replace(
      /(\\mathrm\{e\}\s*\^\s*)i(?=$|[^A-Za-z])/g,
      "$1\\mathrm{i}",
    )
    .replace(
      /(\\exp\s*[({]\s*)i(?=$|[^A-Za-z])/g,
      "$1\\mathrm{i}",
    );
}

function normalizeTypographyChunk(source: string): string {
  return normalizeImaginaryUnitInExponentials(
    normalizeEulerConstant(
      normalizeContextualDifferentialOperators(
        normalizeNamedOperators(source),
      ),
    ),
  );
}

export function normalizeContextualUprightSymbols(source: string): string {
  return transformOutsideProtectedCommands(source, normalizeTypographyChunk);
}

export function normalizeMathLiveCanonicalUprightCommands(
  source: string,
): string {
  let normalized = source;
  for (const [command, standardLatex] of MATHLIVE_CANONICAL_UPRIGHT_COMMANDS) {
    normalized = normalized.replace(
      new RegExp(`\\\\${command}(?![A-Za-z])[ \\t]*`, "g"),
      standardLatex,
    );
  }
  return normalized.replace(
    /\\(?:mathrm|textrm)\{d([A-Za-z])\}/g,
    "\\mathrm{d}$1",
  );
}

const differentialFractionCommands = ["\\dfrac", "\\tfrac", "\\frac"];
const integralCommandPattern =
  /\\(?:oiiint|oiint|oint|iiint|iint|int)(?![A-Za-z])/g;
const nonVariableCommands = new Set([
  "sin",
  "cos",
  "tan",
  "cot",
  "sec",
  "csc",
  "sinh",
  "cosh",
  "tanh",
  "log",
  "ln",
  "exp",
  "lim",
  "max",
  "min",
  "det",
  "gcd",
  "int",
  "iint",
  "iiint",
  "oint",
  "oiint",
  "oiiint",
  "sum",
  "prod",
  "frac",
  "dfrac",
  "tfrac",
  "sqrt",
  "left",
  "right",
  "cdot",
  "times",
  "div",
  "partial",
]);
const styledVariableCommands = new Set([
  "mathbf",
  "boldsymbol",
  "vec",
  "hat",
  "widehat",
  "bar",
  "overline",
  "tilde",
  "widetilde",
  "mathit",
  "mathrm",
  "mathsf",
  "mathtt",
  "mathbb",
  "mathcal",
  "mathscr",
  "mathfrak",
]);
const greekVariableCommands = new Set([
  "alpha",
  "beta",
  "gamma",
  "delta",
  "epsilon",
  "varepsilon",
  "zeta",
  "eta",
  "theta",
  "vartheta",
  "iota",
  "kappa",
  "lambda",
  "mu",
  "nu",
  "xi",
  "pi",
  "varpi",
  "rho",
  "varrho",
  "sigma",
  "varsigma",
  "tau",
  "upsilon",
  "phi",
  "varphi",
  "chi",
  "psi",
  "omega",
  "Gamma",
  "Delta",
  "Theta",
  "Lambda",
  "Xi",
  "Pi",
  "Sigma",
  "Upsilon",
  "Phi",
  "Psi",
  "Omega",
]);

interface BracedSpan {
  open: number;
  close: number;
  content: string;
}

interface DifferentialPrefix {
  dIndex: number;
  end: number;
  alreadyUpright: boolean;
}

function readBracedSpan(source: string, open: number): BracedSpan | null {
  if (source[open] !== "{") return null;
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    else if (source[index] === "}") {
      depth -= 1;
      if (depth === 0) {
        return {
          open,
          close: index,
          content: source.slice(open + 1, index),
        };
      }
    }
  }
  return null;
}

function readCommandEnd(source: string, start: number): number {
  if (source[start] !== "\\") return start;
  let end = start + 1;
  if (/[A-Za-z]/.test(source[end] ?? "")) {
    while (/[A-Za-z]/.test(source[end] ?? "")) end += 1;
    return end;
  }
  return Math.min(source.length, end + 1);
}

function skipDifferentialSpacing(source: string, start: number): number {
  let index = start;
  while (index < source.length) {
    if (/\s/.test(source[index])) {
      index += 1;
      continue;
    }
    if (
      source.startsWith("\\,", index) ||
      source.startsWith("\\!", index) ||
      source.startsWith("\\:", index) ||
      source.startsWith("\\;", index) ||
      source.startsWith("\\ ", index)
    ) {
      index += 2;
      continue;
    }
    const spacing = source.slice(index).match(
      /^\\(?:quad|qquad|enspace|thinspace|medspace|thickspace)(?![A-Za-z])/,
    );
    if (spacing) {
      index += spacing[0].length;
      continue;
    }
    break;
  }
  return index;
}

function readScriptEnd(source: string, start: number): number {
  let index = start;
  while (source[index] === "^" || source[index] === "_") {
    index += 1;
    index = skipDifferentialSpacing(source, index);
    if (source[index] === "{") {
      const group = readBracedSpan(source, index);
      if (!group) return source.length;
      index = group.close + 1;
    } else if (source[index] === "\\") {
      index = readCommandEnd(source, index);
    } else if (index < source.length) {
      index += 1;
    }
    index = skipDifferentialSpacing(source, index);
  }
  return index;
}

function readDifferentialVariableEnd(source: string, start: number): number {
  let index = skipDifferentialSpacing(source, start);
  const character = source[index];
  if (!character) return -1;

  if (/[A-Za-z]/.test(character)) {
    if (/[A-Za-z0-9]/.test(source[index + 1] ?? "")) return -1;
    return readScriptEnd(source, index + 1);
  }

  if (character === "{") {
    const group = readBracedSpan(source, index);
    return group ? readScriptEnd(source, group.close + 1) : -1;
  }

  if (character === "(") {
    let depth = 0;
    for (let cursor = index; cursor < source.length; cursor += 1) {
      if (source[cursor] === "(") depth += 1;
      else if (source[cursor] === ")") {
        depth -= 1;
        if (depth === 0) return readScriptEnd(source, cursor + 1);
      }
    }
    return -1;
  }

  if (character !== "\\") return -1;
  const commandEnd = readCommandEnd(source, index);
  const command = source.slice(index + 1, commandEnd);
  if (!command || nonVariableCommands.has(command)) return -1;

  let end = commandEnd;
  if (styledVariableCommands.has(command)) {
    end = skipDifferentialSpacing(source, end);
    const group = readBracedSpan(source, end);
    if (!group) return -1;
    end = group.close + 1;
  }
  return readScriptEnd(source, end);
}

function readDifferentialPrefix(
  source: string,
  allowBareOperator: boolean,
): DifferentialPrefix | null {
  let index = skipDifferentialSpacing(source, 0);
  if (source.startsWith("\\mathrm{d}", index)) {
    const end = index + "\\mathrm{d}".length;
    const variableEnd = readDifferentialVariableEnd(source, end);
    if (variableEnd >= 0 || allowBareOperator) {
      return { dIndex: index, end: variableEnd >= 0 ? variableEnd : end, alreadyUpright: true };
    }
    return null;
  }
  if (source.startsWith("\\differentialD", index)) {
    const end = index + "\\differentialD".length;
    const variableEnd = readDifferentialVariableEnd(source, end);
    if (variableEnd >= 0 || allowBareOperator) {
      return { dIndex: index, end: variableEnd >= 0 ? variableEnd : end, alreadyUpright: true };
    }
    return null;
  }
  if (source[index] !== "d") return null;

  const dIndex = index;
  index += 1;
  index = skipDifferentialSpacing(source, index);
  if (source[index] === "^") {
    index = readScriptEnd(source, index);
  }
  index = skipDifferentialSpacing(source, index);
  const variableEnd = readDifferentialVariableEnd(source, index);
  if (variableEnd >= 0) {
    return { dIndex, end: variableEnd, alreadyUpright: false };
  }
  if (allowBareOperator && index >= source.length) {
    return { dIndex, end: index, alreadyUpright: false };
  }
  return null;
}

function uprightDifferentialPrefix(
  source: string,
  allowBareOperator: boolean,
): string {
  const prefix = readDifferentialPrefix(source, allowBareOperator);
  if (!prefix || prefix.alreadyUpright) return source;
  return (
    source.slice(0, prefix.dIndex) +
    "\\mathrm{d}" +
    source.slice(prefix.dIndex + 1)
  );
}

function leadingDifferentialVariableCommand(source: string): string | null {
  let index = skipDifferentialSpacing(source, 0);
  if (source.startsWith("\\mathrm{d}", index)) return null;
  if (source.startsWith("\\differentialD", index)) return null;
  if (source[index] !== "d") return null;

  index += 1;
  index = skipDifferentialSpacing(source, index);
  if (source[index] === "^") {
    index = readScriptEnd(source, index);
  }
  index = skipDifferentialSpacing(source, index);

  if (source[index] === "{") {
    const group = readBracedSpan(source, index);
    if (!group) return null;
    const groupIndex = skipDifferentialSpacing(group.content, 0);
    if (group.content[groupIndex] !== "\\") return null;
    const commandEnd = readCommandEnd(group.content, groupIndex);
    return group.content.slice(groupIndex + 1, commandEnd);
  }

  if (source[index] !== "\\") return null;
  const commandEnd = readCommandEnd(source, index);
  return source.slice(index + 1, commandEnd);
}

function uprightLeadingGreekDifferential(source: string): string {
  const variableCommand = leadingDifferentialVariableCommand(source);
  if (!variableCommand || !greekVariableCommands.has(variableCommand)) {
    return source;
  }
  return uprightDifferentialPrefix(source, false);
}

function uprightDifferentialSequence(source: string): string {
  const sequence = integralMeasureTail(source, 0, source.length);
  if (!sequence) return uprightDifferentialPrefix(source, false);

  let normalized = source;
  for (const dIndex of [...sequence.dIndices].sort((a, b) => b - a)) {
    normalized =
      normalized.slice(0, dIndex) +
      "\\mathrm{d}" +
      normalized.slice(dIndex + 1);
  }
  return normalized;
}

function normalizeDerivativeFractions(source: string): string {
  let result = "";
  let index = 0;

  while (index < source.length) {
    const command = differentialFractionCommands.find((candidate) =>
      source.startsWith(candidate, index),
    );
    if (!command) {
      result += source[index];
      index += 1;
      continue;
    }

    const firstOpen = skipDifferentialSpacing(source, index + command.length);
    const numeratorGroup = readBracedSpan(source, firstOpen);
    if (!numeratorGroup) {
      result += source[index];
      index += 1;
      continue;
    }
    const secondOpen = skipDifferentialSpacing(source, numeratorGroup.close + 1);
    const denominatorGroup = readBracedSpan(source, secondOpen);
    if (!denominatorGroup) {
      result += source[index];
      index += 1;
      continue;
    }

    let numerator = uprightLeadingGreekDifferential(
      normalizeDerivativeFractions(numeratorGroup.content),
    );
    let denominator = uprightLeadingGreekDifferential(
      normalizeDerivativeFractions(denominatorGroup.content),
    );
    if (
      readDifferentialPrefix(numerator, true) &&
      readDifferentialPrefix(denominator, false)
    ) {
      numerator = uprightDifferentialPrefix(numerator, true);
      denominator = uprightDifferentialSequence(denominator);
    }

    result +=
      command +
      source.slice(index + command.length, firstOpen) +
      "{" +
      numerator +
      "}" +
      source.slice(numeratorGroup.close + 1, secondOpen) +
      "{" +
      denominator +
      "}";
    index = denominatorGroup.close + 1;
  }
  return result;
}

function topLevelBoundaryBefore(source: string, position: number): number {
  let braceDepth = 0;
  let fenceDepth = 0;
  for (let index = position - 1; index >= 0; index -= 1) {
    const character = source[index];
    if (character === "}") braceDepth += 1;
    else if (character === "{") braceDepth = Math.max(0, braceDepth - 1);
    else if (character === ")" || character === "]") fenceDepth += 1;
    else if (character === "(" || character === "[") {
      if (fenceDepth > 0) fenceDepth -= 1;
      else if (braceDepth === 0) return index + 1;
    } else if (
      braceDepth === 0 &&
      fenceDepth === 0 &&
      /[=+\-;,]/.test(character)
    ) {
      return index + 1;
    }
  }
  return 0;
}

function topLevelBoundaryAfter(source: string, position: number): number {
  let braceDepth = 0;
  let fenceDepth = 0;
  for (let index = position; index < source.length; index += 1) {
    const character = source[index];
    if (character === "{") braceDepth += 1;
    else if (character === "}") braceDepth = Math.max(0, braceDepth - 1);
    else if (character === "(" || character === "[") fenceDepth += 1;
    else if (character === ")" || character === "]") {
      if (fenceDepth > 0) fenceDepth -= 1;
      else if (braceDepth === 0) return index;
    } else if (
      braceDepth === 0 &&
      fenceDepth === 0 &&
      /[=+\-;,]/.test(character)
    ) {
      return index;
    }
  }
  return source.length;
}

function normalizeDirectDifferentialQuotients(source: string): string {
  const replacements = new Set<number>();
  let braceDepth = 0;
  for (let index = 0; index < source.length; index += 1) {
    if (source[index] === "{") braceDepth += 1;
    else if (source[index] === "}") braceDepth = Math.max(0, braceDepth - 1);
    else if (source[index] === "/" && source[index - 1] !== "\\") {
      const leftStart = topLevelBoundaryBefore(source, index);
      const rightEnd = topLevelBoundaryAfter(source, index + 1);
      const left = source.slice(leftStart, index);
      const right = source.slice(index + 1, rightEnd);
      const leftPrefix = readDifferentialPrefix(left, true);
      const rightPrefix = readDifferentialPrefix(right, false);
      if (leftPrefix && rightPrefix) {
        if (!leftPrefix.alreadyUpright) replacements.add(leftStart + leftPrefix.dIndex);
        if (!rightPrefix.alreadyUpright) replacements.add(index + 1 + rightPrefix.dIndex);
      }
    }
  }
  if (replacements.size === 0) return source;

  let result = source;
  for (const position of [...replacements].sort((a, b) => b - a)) {
    result = result.slice(0, position) + "\\mathrm{d}" + result.slice(position + 1);
  }
  return result;
}

function integralMeasureTail(
  source: string,
  start: number,
  boundary: number,
): { dIndices: number[]; end: number } | null {
  const dIndices: number[] = [];
  let index = skipDifferentialSpacing(source, start);
  while (index < boundary) {
    while (/[,.;]/.test(source[index] ?? "")) index += 1;
    if (source.startsWith("\\right", index)) {
      index = readCommandEnd(source, index);
      if (source[index] === "\\") index = readCommandEnd(source, index);
      else if (index < boundary) index += 1;
      index = skipDifferentialSpacing(source, index);
      continue;
    }
    const prefix = readDifferentialPrefix(source.slice(index, boundary), false);
    if (!prefix) break;
    if (!prefix.alreadyUpright) dIndices.push(index + prefix.dIndex);
    index += prefix.end;
    index = skipDifferentialSpacing(source, index);
  }
  index = skipDifferentialSpacing(source, index);
  while (/[,.;)]/.test(source[index] ?? "")) index += 1;
  return index >= boundary && dIndices.length > 0
    ? { dIndices, end: index }
    : null;
}

function normalizeIntegralDifferentials(source: string): string {
  const replacements = new Set<number>();
  const integralMatches = [...source.matchAll(integralCommandPattern)];
  for (const match of integralMatches) {
    const integralEnd = (match.index ?? 0) + match[0].length;
    const segmentEnd = topLevelBoundaryAfter(source, integralEnd);
    let braceDepth = 0;
    let fenceDepth = 0;
    for (let index = integralEnd; index < segmentEnd; index += 1) {
      const character = source[index];
      if (character === "{") {
        braceDepth += 1;
        continue;
      }
      if (character === "}") {
        braceDepth = Math.max(0, braceDepth - 1);
        continue;
      }
      if (character === "(" || character === "[") {
        fenceDepth += 1;
        continue;
      }
      if (character === ")" || character === "]") {
        fenceDepth = Math.max(0, fenceDepth - 1);
        continue;
      }
      if (character !== "d" || braceDepth > 0 || fenceDepth > 0) continue;
      if (index > 0 && /[A-Za-z\\]/.test(source[index - 1])) continue;
      const tail = integralMeasureTail(source, index, segmentEnd);
      if (!tail) continue;
      for (const dIndex of tail.dIndices) replacements.add(dIndex);
      break;
    }
  }
  if (replacements.size === 0) return source;

  let result = source;
  for (const position of [...replacements].sort((a, b) => b - a)) {
    result = result.slice(0, position) + "\\mathrm{d}" + result.slice(position + 1);
  }
  return result;
}

export function normalizeContextualDifferentialOperators(source: string): string {
  return normalizeIntegralDifferentials(
    normalizeDirectDifferentialQuotients(normalizeDerivativeFractions(source)),
  );
}

function readBracedCommand(source: string, start: number): number {
  const openingBrace = source.indexOf("{", start);
  if (openingBrace < 0) return start;
  let depth = 0;
  for (let index = openingBrace; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    if (source[index] === "}") {
      depth -= 1;
      if (depth === 0) return index + 1;
    }
  }
  return source.length;
}

export function normalizeChineseLatex(source: string): string {
  const normalizedTextCommands = normalizeContextualUprightSymbols(
    normalizeMathLiveCanonicalUprightCommands(
      normalizeExtendedIntegralLatexCommands(source),
    ),
  ).replace(
    /\\(?:mathrm|textrm)\{([\u3400-\u9fff\uf900-\ufaff，。；：！？、（）【】《》“”‘’\s]+)\}/g,
    "\\text{$1}",
  );

  let result = "";
  let index = 0;

  while (index < normalizedTextCommands.length) {
    if (normalizedTextCommands.startsWith("\\text{", index)) {
      const end = readBracedCommand(normalizedTextCommands, index);
      result += normalizedTextCommands.slice(index, end);
      index = end;
      continue;
    }

    if (chineseChar.test(normalizedTextCommands[index])) {
      let end = index + 1;
      while (
        end < normalizedTextCommands.length &&
        (chineseChar.test(normalizedTextCommands[end]) ||
          (normalizedTextCommands[end] === " " &&
            end + 1 < normalizedTextCommands.length &&
            chineseChar.test(normalizedTextCommands[end + 1])))
      ) {
        end += 1;
      }
      result += "\\text{" + normalizedTextCommands.slice(index, end) + "}";
      index = end;
      continue;
    }

    result += normalizedTextCommands[index];
    index += 1;
  }

  return result;
}

export function normalizeMultilineLatex(source: string): string {
  return source
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map(normalizeChineseLatex)
    .join("\n");
}
