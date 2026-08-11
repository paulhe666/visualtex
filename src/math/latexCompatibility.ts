import {
  EXTENDED_INTEGRAL_MATHML_MACROS,
  EXTENDED_INTEGRAL_SVG_MACROS,
} from "./extendedIntegralCompatibility.ts";

export type VisualTexMathJaxMacro = string | readonly [replacement: string, argumentCount: number];

/**
 * Rendering-only compatibility aliases shared by every VisualTeX export path.
 * The saved LaTeX is intentionally left untouched. In particular, \bm must
 * retain the mathematical bold-italic semantics of amsmath/bm rather than be
 * degraded to upright \mathbf text.
 */
const STANDARD_COMPATIBILITY_MACROS = {
  bm: ["\\boldsymbol{#1}", 1] as const,
  mathbbm: ["\\mathbb{#1}", 1] as const,

  // Common physics-package aliases. These are rendering-only expansions; the
  // original commands remain in formula metadata and are restored on reopen.
  ip: ["\\left\\langle #1, #2 \\right\\rangle", 2] as const,
  innerproduct: ["\\left\\langle #1, #2 \\right\\rangle", 2] as const,
  dv: ["\\frac{\\mathrm{d} #1}{\\mathrm{d} #2}", 2] as const,
  pdv: ["\\frac{\\partial #1}{\\partial #2}", 2] as const,
  fdv: ["\\frac{\\delta #1}{\\delta #2}", 2] as const,
  abs: ["\\left|#1\\right|", 1] as const,
  norm: ["\\left\\lVert#1\\right\\rVert", 1] as const,
  eval: ["\\left.#1\\right|", 1] as const,
  comm: ["\\left[#1,#2\\right]", 2] as const,
  acomm: ["\\left\\{#1,#2\\right\\}", 2] as const,
  poissonbracket: ["\\left\\{#1,#2\\right\\}", 2] as const,
  expval: ["\\left\\langle #1 \\right\\rangle", 1] as const,
  mel: ["\\left\\langle #1 \\middle| #2 \\middle| #3 \\right\\rangle", 3] as const,
  dd: ["\\mathrm{d}#1", 1] as const,
  vb: ["\\boldsymbol{#1}", 1] as const,
  va: ["\\vec{#1}", 1] as const,
  vu: ["\\hat{#1}", 1] as const,
  pb: ["\\left\\{#1,#2\\right\\}", 2] as const,
  order: ["\\mathcal{O}\\left(#1\\right)", 1] as const,
  tr: "\\operatorname{tr}",
  Tr: "\\operatorname{Tr}",
  rank: "\\operatorname{rank}",
  Res: "\\operatorname{Res}",
  mqty: ["\\begin{matrix}#1\\end{matrix}", 1] as const,
  pmqty: ["\\begin{pmatrix}#1\\end{pmatrix}", 1] as const,
  bqty: ["\\begin{bmatrix}#1\\end{bmatrix}", 1] as const,
  Bqty: ["\\begin{Bmatrix}#1\\end{Bmatrix}", 1] as const,
  vmqty: ["\\begin{vmatrix}#1\\end{vmatrix}", 1] as const,
  Vmqty: ["\\begin{Vmatrix}#1\\end{Vmatrix}", 1] as const,

  // Practical siunitx compatibility. Unit atoms expand to upright symbols so
  // \SI, \si and \unit can pass through the same MathJax/OMML pipeline.
  SI: ["#1\\,#2", 2] as const,
  si: ["#1", 1] as const,
  unit: ["#1", 1] as const,
  num: ["#1", 1] as const,
  numrange: ["#1\\text{--}#2", 2] as const,
  qtyrange: ["#1\\text{--}#2\\,#3", 3] as const,
  numlist: ["#1", 1] as const,
  ang: ["#1^{\\circ}", 1] as const,
  per: "\\,/\\,",
  squared: "^{2}",
  square: "^{2}",
  cubed: "^{3}",
  cubic: "^{3}",
  meter: "\\mathrm{m}",
  metre: "\\mathrm{m}",
  second: "\\mathrm{s}",
  gram: "\\mathrm{g}",
  kilogram: "\\mathrm{kg}",
  ampere: "\\mathrm{A}",
  kelvin: "\\mathrm{K}",
  mole: "\\mathrm{mol}",
  candela: "\\mathrm{cd}",
  hertz: "\\mathrm{Hz}",
  newton: "\\mathrm{N}",
  pascal: "\\mathrm{Pa}",
  joule: "\\mathrm{J}",
  watt: "\\mathrm{W}",
  coulomb: "\\mathrm{C}",
  volt: "\\mathrm{V}",
  farad: "\\mathrm{F}",
  ohm: "\\Omega",
  siemens: "\\mathrm{S}",
  weber: "\\mathrm{Wb}",
  tesla: "\\mathrm{T}",
  henry: "\\mathrm{H}",
  lumen: "\\mathrm{lm}",
  lux: "\\mathrm{lx}",
  becquerel: "\\mathrm{Bq}",
  gray: "\\mathrm{Gy}",
  sievert: "\\mathrm{Sv}",
  katal: "\\mathrm{kat}",
  radian: "\\mathrm{rad}",
  steradian: "\\mathrm{sr}",
  minute: "\\mathrm{min}",
  hour: "\\mathrm{h}",
  day: "\\mathrm{d}",
  litre: "\\mathrm{L}",
  liter: "\\mathrm{L}",
  tonne: "\\mathrm{t}",
  electronvolt: "\\mathrm{eV}",
  dalton: "\\mathrm{Da}",
  astronomicalunit: "\\mathrm{au}",
  parsec: "\\mathrm{pc}",
  bar: "\\mathrm{bar}",
  barn: "\\mathrm{b}",
  knot: "\\mathrm{kn}",
  hectare: "\\mathrm{ha}",
  decibel: "\\mathrm{dB}",
  degree: "{}^{\\circ}",
  arcminute: "{}^{\\prime}",
  arcsecond: "{}^{\\prime\\prime}",
  degreeCelsius: "{}^{\\circ}\\mathrm{C}",
  percent: "\\%",
  deci: "\\mathrm{d}",
  deca: "\\mathrm{da}",
  hecto: "\\mathrm{h}",
  kilo: "\\mathrm{k}",
  mega: "\\mathrm{M}",
  giga: "\\mathrm{G}",
  tera: "\\mathrm{T}",
  peta: "\\mathrm{P}",
  exa: "\\mathrm{E}",
  centi: "\\mathrm{c}",
  milli: "\\mathrm{m}",
  micro: "\\mu",
  nano: "\\mathrm{n}",
  pico: "\\mathrm{p}",
  femto: "\\mathrm{f}",
  atto: "\\mathrm{a}",
} satisfies Record<string, VisualTexMathJaxMacro>;

export const VISUALTEX_MATHML_MACROS = {
  ...EXTENDED_INTEGRAL_MATHML_MACROS,
  ...STANDARD_COMPATIBILITY_MACROS,
} satisfies Record<string, VisualTexMathJaxMacro>;

export const VISUALTEX_SVG_MACROS = {
  ...EXTENDED_INTEGRAL_SVG_MACROS,
  ...STANDARD_COMPATIBILITY_MACROS,
} satisfies Record<string, VisualTexMathJaxMacro>;

type BalancedArgument = {
  content: string;
  end: number;
};

function skipLatexWhitespace(source: string, index: number) {
  while (index < source.length && /\s/.test(source[index])) index += 1;
  return index;
}

function readBalancedArgument(
  source: string,
  start: number,
  open: string,
  close: string,
): BalancedArgument | null {
  if (source[start] !== open) return null;
  if (open === close) {
    for (let index = start + 1; index < source.length; index += 1) {
      if (source[index] === "\\") {
        index += 1;
        continue;
      }
      if (source[index] === close) {
        return { content: source.slice(start + 1, index), end: index + 1 };
      }
    }
    return null;
  }

  let depth = 0;
  for (let index = start; index < source.length; index += 1) {
    if (source[index] === "\\") {
      index += 1;
      continue;
    }
    if (source[index] === open) depth += 1;
    else if (source[index] === close) {
      depth -= 1;
      if (depth === 0) {
        return { content: source.slice(start + 1, index), end: index + 1 };
      }
    }
  }
  return null;
}

function readPackageMacroArgument(source: string, start: number) {
  const position = skipLatexWhitespace(source, start);
  if (source[position] === "{") {
    return readBalancedArgument(source, position, "{", "}");
  }
  if (source[position] === "\\") {
    const command = source.slice(position).match(/^\\(?:[A-Za-z@]+|.)/);
    return command
      ? { content: command[0], end: position + command[0].length }
      : null;
  }
  return position < source.length
    ? { content: source[position], end: position + 1 }
    : null;
}

function normalizeMatrixQuantityCommands(source: string) {
  const commandPattern = /\\(mqty|pmqty|bqty|Bqty|vmqty|Vmqty)(?![A-Za-z@])/g;
  const environments: Record<string, string> = {
    mqty: "matrix",
    pmqty: "pmatrix",
    bqty: "bmatrix",
    Bqty: "Bmatrix",
    vmqty: "vmatrix",
    Vmqty: "Vmatrix",
  };
  let output = "";
  let cursor = 0;
  while (cursor < source.length) {
    commandPattern.lastIndex = cursor;
    const match = commandPattern.exec(source);
    if (!match) {
      output += source.slice(cursor);
      break;
    }
    output += source.slice(cursor, match.index);
    let position = skipLatexWhitespace(source, commandPattern.lastIndex);
    if (source[position] === "*") {
      position = skipLatexWhitespace(source, position + 1);
    }
    const body = source[position] === "{"
      ? readBalancedArgument(source, position, "{", "}")
      : null;
    if (!body) {
      output += match[0];
      cursor = commandPattern.lastIndex;
      continue;
    }
    const environment = environments[match[1]];
    output += `\\begin{${environment}}${body.content}\\end{${environment}}`;
    cursor = body.end;
  }
  return output;
}

function normalizeDerivativeCommands(source: string) {
  const commandPattern = /\\(dv|pdv|fdv)(?![A-Za-z@])/g;
  let output = "";
  let cursor = 0;
  while (cursor < source.length) {
    commandPattern.lastIndex = cursor;
    const match = commandPattern.exec(source);
    if (!match) {
      output += source.slice(cursor);
      break;
    }
    output += source.slice(cursor, match.index);
    let position = commandPattern.lastIndex;
    if (source[position] === "*") position += 1;
    position = skipLatexWhitespace(source, position);
    let order = "";
    if (source[position] === "[") {
      const orderGroup = readBalancedArgument(source, position, "[", "]");
      if (!orderGroup) {
        output += match[0];
        cursor = commandPattern.lastIndex;
        continue;
      }
      order = orderGroup.content.trim();
      position = orderGroup.end;
    }

    const first = readPackageMacroArgument(source, position);
    if (!first) {
      output += match[0];
      cursor = commandPattern.lastIndex;
      continue;
    }
    const second = readPackageMacroArgument(source, first.end);
    const differential = match[1] === "dv"
      ? "\\mathrm{d}"
      : match[1] === "pdv"
        ? "\\partial"
        : "\\delta";
    const numerator = second ? first.content : "";
    const variable = second?.content ?? first.content;
    const power = order ? `^{${order}}` : "";
    output += `\\frac{${differential}${power}${numerator ? ` ${numerator}` : ""}}{${differential} ${variable}${power}}`;
    cursor = second?.end ?? first.end;
  }
  return output;
}

function stripSiunitxOptions(source: string) {
  return source.replace(
    /\\(SI|si|unit|num|numrange|qtyrange|numlist|ang)\s*\[[^\]]*\]/g,
    "\\$1",
  );
}

function normalizeQtyCommands(source: string) {
  let output = "";
  let cursor = 0;
  while (cursor < source.length) {
    const command = source.indexOf("\\qty", cursor);
    if (command < 0) {
      output += source.slice(cursor);
      break;
    }
    output += source.slice(cursor, command);
    const commandEnd = command + "\\qty".length;
    if (/[A-Za-z@]/.test(source[commandEnd] ?? "")) {
      output += "\\qty";
      cursor = commandEnd;
      continue;
    }

    let argumentStart = skipLatexWhitespace(source, commandEnd);
    if (source[argumentStart] === "*") {
      argumentStart = skipLatexWhitespace(source, argumentStart + 1);
    }
    const delimiter = source[argumentStart];
    if (delimiter === "{") {
      const first = readBalancedArgument(source, argumentStart, "{", "}");
      if (!first) {
        output += "\\qty";
        cursor = commandEnd;
        continue;
      }
      const secondStart = skipLatexWhitespace(source, first.end);
      const second = source[secondStart] === "{"
        ? readBalancedArgument(source, secondStart, "{", "}")
        : null;
      if (second) {
        // siunitx v3: \qty{number}{unit}
        output += `\\SI{${first.content}}{${second.content}}`;
        cursor = second.end;
      } else {
        // physics: \qty{expression}
        output += `\\left\\{${first.content}\\right\\}`;
        cursor = first.end;
      }
      continue;
    }

    if (delimiter === "[") {
      const options = readBalancedArgument(source, argumentStart, "[", "]");
      const firstStart = options ? skipLatexWhitespace(source, options.end) : argumentStart;
      const first = options && source[firstStart] === "{"
        ? readBalancedArgument(source, firstStart, "{", "}")
        : null;
      const secondStart = first ? skipLatexWhitespace(source, first.end) : firstStart;
      const second = first && source[secondStart] === "{"
        ? readBalancedArgument(source, secondStart, "{", "}")
        : null;
      if (first && second) {
        // siunitx v3 with options: \qty[...]{number}{unit}
        output += `\\SI{${first.content}}{${second.content}}`;
        cursor = second.end;
        continue;
      }
    }

    const pair = delimiter === "(" ? ["(", ")"] as const
      : delimiter === "[" ? ["[", "]"] as const
        : delimiter === "|" ? ["|", "|"] as const
          : null;
    if (!pair) {
      output += "\\qty";
      cursor = commandEnd;
      continue;
    }
    const argument = readBalancedArgument(source, argumentStart, pair[0], pair[1]);
    if (!argument) {
      output += "\\qty";
      cursor = commandEnd;
      continue;
    }
    const left = pair[0] === "|" ? "\\left|" : `\\left${pair[0]}`;
    const right = pair[1] === "|" ? "\\right|" : `\\right${pair[1]}`;
    output += `${left}${argument.content}${right}`;
    cursor = argument.end;
  }
  return output;
}

/** Normalize package syntax that cannot be represented by fixed MathJax macros. */
export function normalizePackageLatexCommands(source: string) {
  return normalizeDerivativeCommands(
    normalizeMatrixQuantityCommands(
      normalizeQtyCommands(stripSiunitxOptions(source)),
    ),
  );
}

const unresolvedCommand = /\\[A-Za-z@]+/;
const structuralPlaceholder = /\\placeholder\s*\{\s*\}/;
const trailingCommand = /\\[A-Za-z@]+$/;
const environmentToken = /\\(begin|end)\s*\{([^{}]+)\}/g;
const mathText = /<mtext\b([^>]*)>([\s\S]*?)<\/mtext>/gi;

function stripXmlMarkup(value: string) {
  return value.replace(/<[^>]*>/g, "").replace(/&#x5c;|&#92;|&bsol;/gi, "\\");
}

function hasUnbalancedLatexGroups(source: string) {
  let depth = 0;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (character === "%") {
      while (index + 1 < source.length && source[index + 1] !== "\n") index += 1;
      continue;
    }
    if (character === "\\") {
      if (source[index + 1] === "{" || source[index + 1] === "}") index += 1;
      continue;
    }
    if (character === "{") depth += 1;
    else if (character === "}") depth -= 1;
    if (depth < 0) return true;
  }
  return depth !== 0;
}

function hasUnclosedLatexEnvironment(source: string) {
  environmentToken.lastIndex = 0;
  const stack: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = environmentToken.exec(source))) {
    const [, kind, name] = match;
    if (kind === "begin") stack.push(name);
    else if (stack.pop() !== name) return true;
  }
  return stack.length > 0;
}

/**
 * Office drafts are saved on every keystroke. Structural placeholders and the
 * tail of a command/environment are valid editing states, not user-facing
 * export failures. Explicit Insert/Update still uses strict MathJax validation.
 */
export function isIncompleteLatexDraft(source: string, error?: unknown) {
  const normalized = source.replace(/\r\n?/g, "\n").trim();
  if (!normalized) return false;
  if (structuralPlaceholder.test(normalized)) return true;
  if (hasUnbalancedLatexGroups(normalized)) return true;
  if (hasUnclosedLatexEnvironment(normalized)) return true;
  if (
    (normalized.match(/\\left\b/g)?.length ?? 0) !==
    (normalized.match(/\\right\b/g)?.length ?? 0)
  ) {
    return true;
  }

  if (error !== undefined) {
    const trailing = normalized.match(trailingCommand)?.[0];
    const message = error instanceof Error ? error.message : String(error ?? "");
    if (
      trailing &&
      /unresolved|unknown|undefined|did not resolve|missing argument/i.test(message) &&
      message.includes(trailing)
    ) {
      return true;
    }
  }
  return false;
}

export function assertNoUnfilledStructuralPlaceholders(source: string) {
  if (structuralPlaceholder.test(source)) {
    throw new Error(
      "The formula still contains empty VisualTeX placeholders. Fill or remove them before inserting.",
    );
  }
}

/** Reject MathJax's permissive unknown-command fallback before it reaches Word. */
export function assertResolvedPresentationMathMl(mathMl: string) {
  mathText.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = mathText.exec(mathMl))) {
    const attributes = match[1] ?? "";
    const text = stripXmlMarkup(match[2] ?? "");
    if (unresolvedCommand.test(text) || /mathcolor=["']?red/i.test(attributes)) {
      const command = text.match(unresolvedCommand)?.[0] ?? (text.trim() || "unknown command");
      throw new Error(`MathJax did not resolve LaTeX command ${command}.`);
    }
  }
  if (/<merror\b/i.test(mathMl)) {
    throw new Error("MathJax produced an error node instead of semantic MathML.");
  }
}

/** The document-import preview uses SVG directly, so guard that branch too. */
export function assertResolvedMathJaxSvg(svg: string) {
  if (
    /<g\b[^>]*data-mml-node=["']mtext["'][^>]*(?:fill|stroke)=["']red["']/i.test(svg)
    || /<g\b[^>]*(?:fill|stroke)=["']red["'][^>]*data-mml-node=["']mtext["']/i.test(svg)
    || /data-mml-node=["']merror["']/i.test(svg)
  ) {
    const text = stripXmlMarkup(svg)
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&amp;/gi, "&");
    const command = text.match(unresolvedCommand)?.[0];
    throw new Error(
      command
        ? `MathJax did not resolve LaTeX command ${command}.`
        : "MathJax left an unresolved LaTeX command in the rendered formula.",
    );
  }
}
