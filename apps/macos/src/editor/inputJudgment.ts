import {
  EXTENDED_INTEGRAL_COMMANDS,
  EXTENDED_INTEGRAL_COMMAND_PATTERN_SOURCE,
} from "../math/extendedIntegralCompatibility.ts";
import {
  AFTER_NUMBER_WORDS,
  CERTAIN_LABELS,
  FIXED_FUNCTIONS,
  GREEK_WORDS,
  LEGACY_OPERATOR_UNITS,
  SUBSCRIPT_LABELS,
  SYMBOL_SEQUENCES,
  SYMBOL_WORDS,
  TEMPLATE_WORDS,
  TYPED_CONSTANTS,
  isCertainUnit,
  isUnitSymbol,
} from "./inputDictionary.ts";

// The input judgment (docs/upright-typography-plan-20261015.md): one routine,
// one dictionary (inputDictionary.ts) for both what typed words turn into
// (quick escapes: alpha -> \alpha) and what is set upright (units, function
// names, constants, labels, the differential d). The MathLive kernel calls it
// when a typed word is finished; formula normalization calls it for pasted,
// imported and stored LaTeX (upright only, never quick escapes).
//
// The upright rule: a symbol that names one fixed thing is set upright; a
// symbol that stands for a value stays italic. A spelling that could be
// either stays italic unless its surroundings show the fixed use.

export interface InputJudgmentOptions {
  /** Automatic upright typography. Defaults to the editor setting. */
  upright?: boolean;
  /** Quick escapes for typed words (alpha, frac, sintheta). */
  escapes?: boolean;
  /**
   * The LaTeX was just typed in the editor. Quick escapes and the typed-only
   * forms (dx, ii, 30 deg) apply only then, never to pasted or imported LaTeX.
   */
  typed?: boolean;
  /**
   * The caret sits right after a letter: that word may still grow
   * (cm -> cmax), so words wait until finished. Differentials, e and i are
   * decided at once.
   */
  deferLetterRuns?: boolean;
  /**
   * Produce the same judgment with every converted word restored as typed
   * and marked explicit (\mathit{...}), so it is never converted again.
   * Backspace right after a conversion applies this.
   */
  revert?: boolean;
}

let defaultUprightEnabled = true;

/** The editor's "automatic upright typography" setting. */
export function setUprightJudgmentEnabled(enabled: boolean) {
  defaultUprightEnabled = enabled;
}

export function isUprightJudgmentEnabled() {
  return defaultUprightEnabled;
}

/** Symbol sequences (>=, ->, ...) the kernel converts as they are typed. */
export const VISUALTEX_SYMBOL_SEQUENCES = SYMBOL_SEQUENCES;

interface RuleContext {
  topLevel: boolean;
  inSubscript: boolean;
  inScript: boolean;
  exponentOfE: boolean;
  upright: boolean;
  escapes: boolean;
  typed: boolean;
  deferLetterRuns: boolean;
  revert: boolean;
}

type Part = string | { content: string; child: Partial<RuleContext> };

interface Item {
  kind:
    | "ws" | "run" | "number" | "space" | "slash" | "cdot" | "open"
    | "script" | "group" | "command" | "styled" | "other";
  raw: string;
  /** Letters of a run, or of the argument of \mathrm / legacy \operatorname. */
  text?: string;
  /** Command name without the backslash. */
  name?: string;
  which?: "^" | "_";
  /** Script body without braces. */
  body?: string;
  parts?: Part[];
  legacyUnit?: boolean;
}

// Commands whose arguments are a style or text the user chose, or are not
// math at all (dimensions, colors, labels): copied verbatim, never rewritten.
const VERBATIM_ARGUMENT_COMMANDS = new Set([
  "text", "textrm", "textit", "textbf", "textsf", "texttt", "textnormal",
  "textup", "mbox", "hbox", "mathrm", "mathit", "mathbf", "mathsf",
  "mathtt", "mathbb", "mathcal", "mathscr", "mathfrak", "mathnormal",
  "mathbfit", "mathsfit", "boldsymbol", "bm", "pmb", "symrm", "symit",
  "symbf", "symsf", "symtt", "symbb", "symcal", "symfrak", "symnormal",
  "operatorname", "operatorname*", "ce", "pu", "color", "textcolor",
  "colorbox", "fcolorbox", "hspace", "vspace", "kern", "mkern", "mskip",
  "hskip", "rule", "raisebox", "label", "tag", "ref", "eqref", "href",
  "url", "begin", "end", "placeholder",
]);

const SPACE_COMMANDS = new Set([
  " ", ",", ":", ";", "quad", "qquad", "enspace", "thinspace", "medspace",
  "thickspace",
]);

const DELIMITER_SIZE_COMMANDS =
  /^(?:left|right|middle|bigl|bigr|Bigl|Bigr|biggl|biggr|Biggl|Biggr|big|Big|bigg|Bigg)$/;

// Big operators behave like functions in a chain of names: intx -> \int x.
const CHAIN_OPERATORS: Readonly<Record<string, string>> = {
  sum: "\\sum", prod: "\\prod", int: "\\int", iint: "\\iint",
  iiint: "\\iiint", oint: "\\oint",
};

function readBraced(source: string, open: number) {
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    const character = source[index];
    if (character === "\\") {
      index += 1;
      continue;
    }
    if (character === "{") depth += 1;
    else if (character === "}") {
      depth -= 1;
      if (depth === 0) return { content: source.slice(open + 1, index), end: index + 1 };
    }
  }
  return null;
}

function readControlSequence(source: string, start: number) {
  let end = start + 1;
  if (/[A-Za-z]/.test(source[end] ?? "")) {
    while (/[A-Za-z]/.test(source[end] ?? "")) end += 1;
  } else {
    end = Math.min(source.length, end + 1);
  }
  return end;
}

function skipWhitespace(source: string, start: number) {
  let index = start;
  while (/\s/.test(source[index] ?? "")) index += 1;
  return index;
}

function styledLetters(name: string, argument: string) {
  if (name === "mathrm" || name === "textrm") {
    return /^[A-Za-z]+$/.test(argument) ? argument : undefined;
  }
  if (name === "operatorname") {
    const match = /^(?:\\mathrm\{([A-Za-z]+)\}|([A-Za-z]+))$/.exec(argument.trim());
    return match ? match[1] ?? match[2] : undefined;
  }
  return undefined;
}

function tokenize(source: string): Item[] {
  const items: Item[] = [];
  let index = 0;
  while (index < source.length) {
    const character = source[index];
    const start = index;
    if (/\s/.test(character)) {
      index = skipWhitespace(source, index);
      items.push({ kind: "ws", raw: source.slice(start, index) });
      continue;
    }
    if (/[A-Za-z]/.test(character)) {
      while (/[A-Za-z]/.test(source[index] ?? "")) index += 1;
      const raw = source.slice(start, index);
      items.push({ kind: "run", raw, text: raw });
      continue;
    }
    if (/[0-9]/.test(character)) {
      const match = /^[0-9]+(?:\.[0-9]+)?/.exec(source.slice(index))!;
      index += match[0].length;
      items.push({ kind: "number", raw: match[0] });
      continue;
    }
    if (character === "/") {
      items.push({ kind: "slash", raw: "/" });
      index += 1;
      continue;
    }
    if (character === "·") {
      items.push({ kind: "cdot", raw: character });
      index += 1;
      continue;
    }
    if (character === "~") {
      items.push({ kind: "space", raw: "~" });
      index += 1;
      continue;
    }
    if (character === "(" || character === "[") {
      items.push({ kind: "open", raw: character });
      index += 1;
      continue;
    }
    if (character === "{") {
      const group = readBraced(source, index);
      if (!group) {
        items.push({ kind: "other", raw: source.slice(index) });
        break;
      }
      index = group.end;
      items.push({
        kind: "group",
        raw: source.slice(start, index),
        parts: ["{", { content: group.content, child: {} }, "}"],
      });
      continue;
    }
    if (character === "^" || character === "_") {
      const which = character;
      const bodyStart = skipWhitespace(source, index + 1);
      const parts: Part[] = [source.slice(index, bodyStart)];
      let body = "";
      if (source[bodyStart] === "{") {
        const group = readBraced(source, bodyStart);
        if (!group) {
          items.push({ kind: "other", raw: source.slice(index) });
          break;
        }
        body = group.content;
        parts.push("{", { content: body, child: {} }, "}");
        index = group.end;
      } else if (source[bodyStart] === "\\") {
        let end = readControlSequence(source, bodyStart);
        if (source.startsWith("{}", end)) end += 2;
        body = source.slice(bodyStart, end);
        parts.push(body);
        index = end;
      } else {
        body = source[bodyStart] ?? "";
        parts.push(body);
        index = Math.min(source.length, bodyStart + 1);
      }
      items.push({ kind: "script", raw: source.slice(start, index), which, body, parts });
      continue;
    }
    if (character !== "\\") {
      items.push({ kind: "other", raw: character });
      index += 1;
      continue;
    }

    let end = readControlSequence(source, index);
    let name = source.slice(index + 1, end);
    if (source[end] === "*" && /^[A-Za-z]+$/.test(name)) {
      name += "*";
      end += 1;
    }
    if (SPACE_COMMANDS.has(name)) {
      index = end;
      items.push({ kind: "space", raw: source.slice(start, index) });
      continue;
    }
    if (name === "cdot") {
      index = end;
      items.push({ kind: "cdot", raw: source.slice(start, index) });
      continue;
    }
    if (DELIMITER_SIZE_COMMANDS.test(name)) {
      const delimiterStart = skipWhitespace(source, end);
      const delimiterEnd = source[delimiterStart] === "\\"
        ? readControlSequence(source, delimiterStart)
        : Math.min(source.length, delimiterStart + 1);
      const delimiter = source.slice(delimiterStart, delimiterEnd);
      index = delimiterEnd;
      const opening = !/^(?:right|middle|bigr|Bigr|biggr|Biggr)$/.test(name);
      items.push({
        kind: opening && (delimiter === "(" || delimiter === "[") ? "open" : "other",
        raw: source.slice(start, index),
      });
      continue;
    }

    const verbatim = VERBATIM_ARGUMENT_COMMANDS.has(name);
    const parts: Part[] = [source.slice(index, end)];
    let firstArgument: string | undefined;
    index = end;
    for (;;) {
      const argumentStart = skipWhitespace(source, index);
      if (source[argumentStart] !== "{") break;
      const group = readBraced(source, argumentStart);
      if (!group) break;
      firstArgument ??= group.content;
      parts.push(source.slice(index, argumentStart));
      if (verbatim) parts.push(source.slice(argumentStart, group.end));
      else parts.push("{", { content: group.content, child: {} }, "}");
      index = group.end;
    }
    const raw = source.slice(start, index);
    if (verbatim) {
      const letters = firstArgument === undefined ? undefined : styledLetters(name, firstArgument);
      items.push({
        kind: "styled",
        raw,
        name,
        text: letters,
        legacyUnit: name === "operatorname" && letters !== undefined && LEGACY_OPERATOR_UNITS.has(letters),
      });
    } else {
      items.push({ kind: "command", raw, name, parts });
    }
  }
  return items;
}

function nextIndex(items: Item[], index: number) {
  for (let cursor = index + 1; cursor < items.length; cursor += 1) {
    if (items[cursor].kind !== "ws") return cursor;
  }
  return -1;
}

function previousIndex(items: Item[], index: number) {
  for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
    if (items[cursor].kind !== "ws") return cursor;
  }
  return -1;
}

interface UnitMember {
  text: string;
  certain: boolean;
  /** Already \mathrm{...}: confirmed, needs no rewrite. */
  upright: boolean;
  /** Old \operatorname{km}: confirmed, rewritten to \mathrm{km}. */
  legacy: boolean;
  /** \min after a number: the unit minute. */
  command: boolean;
}

function unitMember(items: Item[], index: number): UnitMember | null {
  const item = items[index];
  const after = items[nextIndex(items, index)];
  // m_1 or h(t) is a variable, not a unit.
  if (after && ((after.kind === "script" && after.which === "_") || after.kind === "open")) return null;
  if (item.kind === "run" && isUnitSymbol(item.text!)) {
    return { text: item.text!, certain: isCertainUnit(item.text!), upright: false, legacy: false, command: false };
  }
  if (item.kind === "styled" && item.text && isUnitSymbol(item.text)) {
    if (item.legacyUnit) {
      return { text: item.text, certain: true, upright: false, legacy: true, command: false };
    }
    if (item.name === "mathrm" || item.name === "textrm") {
      return { text: item.text, certain: false, upright: true, legacy: false, command: false };
    }
  }
  if (item.kind === "command" && item.name === "min" && item.parts!.length === 1) {
    return { text: "min", certain: false, upright: false, legacy: false, command: true };
  }
  return null;
}

/** A space after a number: 72 km/h, 3\times10^8 m/s. */
function followsNumberAndSpace(items: Item[], index: number, spaceRequired = true) {
  let cursor = previousIndex(items, index);
  let spaces = 0;
  while (cursor >= 0 && items[cursor].kind === "space") {
    spaces += 1;
    cursor = previousIndex(items, cursor);
  }
  while (cursor >= 0 && items[cursor].kind === "script" && items[cursor].which === "^") {
    cursor = previousIndex(items, cursor);
  }
  return (spaces > 0 || !spaceRequired) && cursor >= 0 && items[cursor].kind === "number";
}

/** lg(x), lg x, Re z: an argument follows. */
function argumentFollows(items: Item[], index: number) {
  let cursor = nextIndex(items, index);
  while (cursor >= 0 && items[cursor].kind === "script") cursor = nextIndex(items, cursor);
  if (cursor < 0) return false;
  if (items[cursor].kind === "open") return true;
  if (items[cursor].kind !== "space") return false;
  const argument = items[nextIndex(items, cursor)];
  return Boolean(argument && /^(?:run|number|command|styled|group|open)$/.test(argument.kind));
}

/** Exponent of e: a variable, i or a sign makes e^{...} the exponential. */
function exponentIsVariable(body: string) {
  const content = body.trim();
  if (!content || /^\\placeholder\{\}$/.test(content)) return false;
  return !/^[\s0-9.,{}]*$/.test(content);
}

function onlyItem(items: Item[], index: number) {
  return items.every((item, cursor) => cursor === index || item.kind === "ws" || item.kind === "space");
}

/** A typed differential word: dx, dt, dtheta at the start of a term. */
function typedDifferential(text: string, context: RuleContext) {
  const match = /^d([A-Za-z]+)$/.exec(text);
  // dd is two variables, not the differential of d.
  if (!match || match[1] === "d") return null;
  if (match[1].length === 1) return `\\mathrm{d}${match[1]}`;
  return context.escapes && GREEK_WORDS[match[1]] ? `\\mathrm{d}${GREEK_WORDS[match[1]]}` : null;
}

/** A name that may start or continue a chain (sintheta, pir, intx). */
function chainName(candidate: string, context: RuleContext) {
  if (context.escapes && GREEK_WORDS[candidate]) {
    return { latex: GREEK_WORDS[candidate], takesArgument: false };
  }
  if (context.escapes && CHAIN_OPERATORS[candidate]) {
    return { latex: CHAIN_OPERATORS[candidate], takesArgument: true };
  }
  const fixed = FIXED_FUNCTIONS.get(candidate);
  if (context.upright && fixed?.certain) return { latex: fixed.latex, takesArgument: true };
  return null;
}

/**
 * A typed word made of names from its start, each right after the previous
 * one, plus an argument: sintheta -> \sin\theta, pir -> \pi r. A name in the
 * middle of other letters (spin, point, number) is never recognized.
 */
function chainLatex(word: string, context: RuleContext) {
  const parts: string[] = [];
  let position = 0;
  let lastTakesArgument = false;
  while (position < word.length) {
    let match: { length: number; latex: string; takesArgument: boolean } | null = null;
    for (let end = word.length; end >= position + 2; end -= 1) {
      const name = chainName(word.slice(position, end), context);
      if (name) {
        match = { length: end - position, ...name };
        break;
      }
    }
    if (!match) break;
    parts.push(match.latex);
    lastTakesArgument = match.takesArgument;
    position += match.length;
  }
  const rest = word.slice(position);
  if (parts.length === 0 || (parts.length === 1 && !rest)) return null;
  if (rest && !lastTakesArgument && rest.length > 1) return null;
  if (rest) parts.push(rest);
  return parts.reduce((latex, part) =>
    /\\[A-Za-z]+$/.test(latex) && /^[A-Za-z]/.test(part) ? `${latex} ${part}` : latex + part);
}

function classify(items: Item[], context: RuleContext) {
  const replacements = new Map<number, string>();
  const exponentialScripts = new Set<number>();
  const decided = new Set<number>();
  // Revert mode restores what was typed, marked explicit so it stays.
  const convert = (index: number, latex: string) => {
    replacements.set(index, context.revert ? `\\mathit{${items[index].raw}}` : latex);
  };

  // Units: a group of unit symbols joined by /, \cdot or a space, each with
  // an optional exponent, is upright as a whole when one member is certain
  // (or already upright), when it follows a number and a space, or when it
  // is the whole formula (decision A1, 2026-10-15).
  if (context.upright && !context.deferLetterRuns) {
    for (let index = 0; index < items.length; index += 1) {
      const first = unitMember(items, index);
      if (!first) continue;
      const members = [{ index, member: first }];
      let joinedByOperator = false;
      let end = index;
      let cursor = nextIndex(items, index);
      for (;;) {
        while (cursor >= 0 && items[cursor].kind === "script" && items[cursor].which === "^") {
          end = cursor;
          cursor = nextIndex(items, cursor);
        }
        let joiner = cursor;
        let operators = 0;
        while (joiner >= 0 && /^(?:slash|cdot|space)$/.test(items[joiner].kind)) {
          if (items[joiner].kind !== "space") operators += 1;
          joiner = nextIndex(items, joiner);
        }
        const next = joiner !== cursor && joiner >= 0 && operators <= 1 ? unitMember(items, joiner) : null;
        if (!next) break;
        if (operators > 0) joinedByOperator = true;
        members.push({ index: joiner, member: next });
        end = joiner;
        cursor = nextIndex(items, joiner);
      }

      const confirmed = members.some(({ member }) =>
        member.certain || member.legacy || (member.upright && member.text !== "d"));
      const wholeFormula =
        context.topLevel &&
        members.length >= 2 &&
        joinedByOperator &&
        items.every((item, position) =>
          (position >= index && position <= end) || item.kind === "ws" || item.kind === "space");
      if (confirmed || followsNumberAndSpace(items, index) || wholeFormula) {
        for (const { index: position, member } of members) {
          decided.add(position);
          if (member.upright) continue;
          if (member.legacy && context.revert) continue;
          convert(position, `\\mathrm{${member.text}}`);
        }
      }
      index = end;
    }
  }

  for (let index = 0; index < items.length; index += 1) {
    if (decided.has(index)) continue;
    const item = items[index];
    const next = nextIndex(items, index);
    const previous = previousIndex(items, index);

    if (item.kind === "styled" && item.name === "mathrm" && item.text === "e" &&
        next >= 0 && items[next].kind === "script" && items[next].which === "^") {
      exponentialScripts.add(next);
      continue;
    }
    if (item.kind !== "run") continue;
    const text = item.text!;

    // Constants, decided at once like the differential d.
    if (context.upright) {
      if (text === "e" && next >= 0 && items[next].kind === "script" &&
          items[next].which === "^" && exponentIsVariable(items[next].body!)) {
        convert(index, "\\mathrm{e}");
        exponentialScripts.add(next);
        continue;
      }
      if (text === "i") {
        const afterExp = previous >= 0 && items[previous].kind === "open" &&
          items[previousIndex(items, previous)]?.kind === "command" &&
          items[previousIndex(items, previous)].name === "exp";
        const beforePrevious = previousIndex(items, previous);
        const complexNumber = !context.inScript && previous >= 0 &&
          items[previous].kind === "number" && previous === index - 1 &&
          (beforePrevious < 0 || /^[+-]$/.test(items[beforePrevious].raw)) &&
          (next < 0 || items[next].kind === "other" || items[next].kind === "space");
        if (context.exponentOfE || afterExp || complexNumber) convert(index, "\\mathrm{i}");
        continue;
      }
    }

    if (context.deferLetterRuns) continue;
    const previousKind = previous >= 0 ? items[previous].kind : "none";

    // Labels: a descriptive subscript word (v_max) or pH anywhere.
    if (context.upright && context.inSubscript && SUBSCRIPT_LABELS.has(text) && onlyItem(items, index)) {
      convert(index, `\\mathrm{${text}}`);
      continue;
    }
    if (context.upright && CERTAIN_LABELS.has(text)) {
      convert(index, `\\mathrm{${text}}`);
      continue;
    }

    if (context.typed) {
      // 30 deg -> 30°.
      if (context.escapes && AFTER_NUMBER_WORDS[text] && followsNumberAndSpace(items, index, false)) {
        convert(index, AFTER_NUMBER_WORDS[text]);
        continue;
      }
      // ee, ii, jj typed for the constants.
      if (context.upright && TYPED_CONSTANTS[text] && previousKind !== "run") {
        convert(index, TYPED_CONSTANTS[text]);
        continue;
      }
      // dx, dt, dtheta at the start of a term.
      const differential = context.upright && previousKind !== "run" ? typedDifferential(text, context) : null;
      if (differential) {
        convert(index, differential);
        continue;
      }
      // A whole word that is a name: alpha, infty, frac, in.
      const escape = context.escapes
        ? GREEK_WORDS[text] ?? SYMBOL_WORDS[text] ?? TEMPLATE_WORDS[text]
        : undefined;
      if (escape) {
        convert(index, escape);
        continue;
      }
    }

    // Functions: certain names always, ambiguous ones (lg, Re) with an argument.
    const fixedFunction = context.upright ? FIXED_FUNCTIONS.get(text) : undefined;
    if (fixedFunction && (fixedFunction.certain || argumentFollows(items, index))) {
      convert(index, fixedFunction.latex);
      continue;
    }

    // Names chained from the start of a typed word: sintheta, pir, intx.
    if (context.typed) {
      const chain = chainLatex(text, context);
      if (chain) convert(index, chain);
    }
  }
  return { replacements, exponentialScripts };
}

function processBranch(source: string, context: RuleContext): string {
  const items = tokenize(source);
  const { replacements, exponentialScripts } = classify(items, context);
  const render = (parts: Part[], child: Partial<RuleContext>) =>
    parts
      .map((part) =>
        typeof part === "string"
          ? part
          : processBranch(part.content, { ...context, topLevel: false, ...child }))
      .join("");

  return items
    .map((item, index) => {
      const replacement = replacements.get(index);
      if (replacement !== undefined) {
        // A control word must not swallow a following letter.
        const following = items[index + 1]?.raw ?? "";
        return /\\[A-Za-z]+$/.test(replacement) && /^[A-Za-z]/.test(following)
          ? replacement + " "
          : replacement;
      }
      if (item.kind === "script") {
        return render(item.parts!, {
          inScript: true,
          inSubscript: item.which === "_",
          exponentOfE: exponentialScripts.has(index) || (context.exponentOfE && item.which === "^"),
        });
      }
      if (item.kind === "group" || item.kind === "command") return render(item.parts!, {});
      return item.raw;
    })
    .join("");
}

// The replacement for a differential d; \mathit{d} in revert mode.
let differentialReplacement = "\\mathrm{d}";

/** The input judgment for a formula's LaTeX. Idempotent. */
export function applyInputJudgment(
  source: string,
  options: InputJudgmentOptions = {},
): string {
  const context: RuleContext = {
    topLevel: true,
    inSubscript: false,
    inScript: false,
    exponentOfE: false,
    upright: options.upright ?? defaultUprightEnabled,
    escapes: Boolean(options.escapes),
    typed: Boolean(options.typed),
    deferLetterRuns: Boolean(options.deferLetterRuns),
    revert: Boolean(options.revert),
  };
  if (!context.upright && !context.escapes) return source;
  const letters = processBranch(source, context);
  if (!context.upright) return letters;
  differentialReplacement = context.revert ? "\\mathit{d}" : "\\mathrm{d}";
  try {
    return transformOutsideProtectedCommands(letters, normalizeContextualDifferentialOperators);
  } finally {
    differentialReplacement = "\\mathrm{d}";
  }
}

/**
 * The quick escape of the confirm-escape key (Space by default): the longest
 * name the typed letters before the caret end with, and only that name
 * (thetaalpha -> theta\alpha). A function name wins over a shorter escape it
 * ends with (sin is not s\in): automatic upright typography sets it. deg is a
 * degree sign only right after a number (30 deg); otherwise it is the
 * function. Returns how many trailing letters it replaces, or null.
 */
export function escapeWordAtEnd(
  letters: string,
  afterNumber = false,
): { length: number; latex: string } | null {
  for (let start = 0; start < letters.length; start += 1) {
    const word = letters.slice(start);
    if (start === 0 && afterNumber && AFTER_NUMBER_WORDS[word]) {
      return { length: word.length, latex: AFTER_NUMBER_WORDS[word] };
    }
    const escape = GREEK_WORDS[word] ?? SYMBOL_WORDS[word] ?? TEMPLATE_WORDS[word];
    if (escape) return { length: word.length, latex: escape };
    if (FIXED_FUNCTIONS.has(word)) return null;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Differential d (operator kind): derivative fractions, quotients and
// integral measures. Moved from normalizeChineseLatex.ts.
// ---------------------------------------------------------------------------


const protectedTypographyCommands = new Set([
  "text",
  "textnormal",
  "textrm",
  "mathrm",
  "operatorname",
  "mbox",
]);

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

    const end = readBracedCommand(source, index);
    const semanticUprightSymbol =
      (commandMatch[1] === "mathrm" || commandMatch[1] === "textrm") &&
      /^\\(?:mathrm|textrm)\{[dDeij]\}$/.test(source.slice(index, end));
    if (semanticUprightSymbol) {
      // Keep canonical one-symbol upright atoms in the surrounding semantic
      // chunk. Incremental MathLive input can normalize the denominator first
      // (for example \frac{d^2y}{\mathrm{d}x^2}); splitting at that atom
      // would hide the still-italic numerator from derivative detection.
      index = end;
      continue;
    }

    result += transform(source.slice(chunkStart, index));
    result += source.slice(index, end);
    index = end;
    chunkStart = end;
  }

  return result + transform(source.slice(chunkStart));
}

const differentialFractionCommands = ["\\dfrac", "\\tfrac", "\\frac"];
// The extended list covers only rare/esint spellings; the standard integrals
// must be listed here or ordinary \int ... dx never receives an upright d.
const integralCommandPattern = new RegExp(
  `\\\\(?:${EXTENDED_INTEGRAL_COMMAND_PATTERN_SOURCE}|iiint|iint|oint|int)(?![A-Za-z])`,
  "g",
);
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
  ...EXTENDED_INTEGRAL_COMMANDS,
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
    differentialReplacement +
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
      differentialReplacement +
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
    result = result.slice(0, position) + differentialReplacement + result.slice(position + 1);
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
      if (insideCommandName(source, index)) continue;
      const tail = integralMeasureTail(source, index, segmentEnd);
      if (!tail) continue;
      for (const dIndex of tail.dIndices) replacements.add(dIndex);
      break;
    }
  }
  if (replacements.size === 0) return source;

  let result = source;
  for (const position of [...replacements].sort((a, b) => b - a)) {
    result = result.slice(0, position) + differentialReplacement + result.slice(position + 1);
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

// The letter at `index` belongs to a control word such as \pmod.
function insideCommandName(source: string, index: number) {
  let cursor = index - 1;
  while (cursor >= 0 && /[A-Za-z]/.test(source[cursor])) cursor -= 1;
  return cursor >= 0 && source[cursor] === "\\";
}
