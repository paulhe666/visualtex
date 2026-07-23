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

const mathLiveCanonicalUprightCommands: Record<string, string> = {
  differentialD: "\\mathrm{d}",
  capitalDifferentialD: "\\mathrm{D}",
  exponentialE: "\\mathrm{e}",
  imaginaryI: "\\mathrm{i}",
  imaginaryJ: "\\mathrm{j}",
};

const mathLiveCanonicalUprightPattern = new RegExp(
  `\\\\(${Object.keys(mathLiveCanonicalUprightCommands).join("|")})(?![A-Za-z])\\s*`,
  "g",
);

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

function isAsciiLetter(value: string | undefined): boolean {
  return Boolean(value && /[A-Za-z]/.test(value));
}

function readExponentEnd(source: string, start: number): number {
  if (source[start] !== "^") return start;
  let index = start + 1;
  if (source[index] === "{") {
    let depth = 0;
    for (; index < source.length; index += 1) {
      if (source[index] === "{") depth += 1;
      if (source[index] === "}") {
        depth -= 1;
        if (depth === 0) return index + 1;
      }
    }
    return source.length;
  }
  if (source[index] === "\\") {
    index += 1;
    while (isAsciiLetter(source[index])) index += 1;
    return index;
  }
  return Math.min(source.length, index + 1);
}

function readDifferentialVariableEnd(source: string, start: number): number {
  let index = start;
  while (/\s/.test(source[index] ?? "")) index += 1;

  if (isAsciiLetter(source[index])) {
    const end = index + 1;
    if (isAsciiLetter(source[end])) return -1;
    return end;
  }

  if (source[index] === "\\") {
    const match = /^\\([A-Za-z]+)/.exec(source.slice(index));
    if (!match) return -1;
    if (
      ["frac", "dfrac", "tfrac", "text", "mathrm", "operatorname"].includes(
        match[1],
      )
    ) {
      return -1;
    }
    return index + match[0].length;
  }

  return -1;
}

const latexSpacingCommands = [
  "\\qquad",
  "\\quad",
  "\\,",
  "\\;",
  "\\:",
  "\\!",
];

interface BracedArgumentRange {
  contentStart: number;
  contentEnd: number;
  nextIndex: number;
}

interface SourceDepth {
  braces: number;
  parentheses: number;
  brackets: number;
}

function skipWhitespaceAndSpacing(
  source: string,
  start: number,
  limit = source.length,
): number {
  let index = start;
  while (index < limit) {
    if (/\s/.test(source[index] ?? "")) {
      index += 1;
      continue;
    }
    const spacing = latexSpacingCommands.find((command) =>
      source.startsWith(command, index),
    );
    if (!spacing) break;
    index += spacing.length;
  }
  return index;
}

function readBracedArgument(
  source: string,
  start: number,
): BracedArgumentRange | null {
  const openingBrace = skipWhitespaceAndSpacing(source, start);
  if (source[openingBrace] !== "{") return null;

  let depth = 0;
  for (let index = openingBrace; index < source.length; index += 1) {
    if (source[index] === "{" && source[index - 1] !== "\\") depth += 1;
    if (source[index] === "}" && source[index - 1] !== "\\") {
      depth -= 1;
      if (depth === 0) {
        return {
          contentStart: openingBrace + 1,
          contentEnd: index,
          nextIndex: index + 1,
        };
      }
    }
  }
  return null;
}

function readBalancedGroupEnd(
  source: string,
  start: number,
  opening: string,
  closing: string,
  limit = source.length,
): number {
  if (source[start] !== opening) return -1;
  let depth = 0;
  for (let index = start; index < limit; index += 1) {
    if (source[index - 1] === "\\") continue;
    if (source[index] === opening) depth += 1;
    if (source[index] === closing) {
      depth -= 1;
      if (depth === 0) return index + 1;
    }
  }
  return -1;
}

function readScriptArgumentEnd(
  source: string,
  start: number,
  limit = source.length,
): number {
  const argumentStart = skipWhitespaceAndSpacing(source, start, limit);
  if (argumentStart >= limit) return -1;
  if (source[argumentStart] === "{") {
    return readBalancedGroupEnd(source, argumentStart, "{", "}", limit);
  }
  if (source[argumentStart] === "\\") {
    const command = /^\\(?:[A-Za-z]+|.)/.exec(source.slice(argumentStart));
    return command ? Math.min(limit, argumentStart + command[0].length) : -1;
  }
  return Math.min(limit, argumentStart + Array.from(source.slice(argumentStart))[0].length);
}

function readMathAtomEnd(
  source: string,
  start: number,
  limit = source.length,
): number {
  const atomStart = skipWhitespaceAndSpacing(source, start, limit);
  if (atomStart >= limit) return -1;

  let atomEnd = -1;
  const first = source[atomStart];
  if (first === "{") {
    atomEnd = readBalancedGroupEnd(source, atomStart, "{", "}", limit);
  } else if (first === "(") {
    atomEnd = readBalancedGroupEnd(source, atomStart, "(", ")", limit);
  } else if (first === "[") {
    atomEnd = readBalancedGroupEnd(source, atomStart, "[", "]", limit);
  } else if (first === "\\") {
    const command = /^\\(?:[A-Za-z]+|.)/.exec(source.slice(atomStart));
    if (!command) return -1;
    atomEnd = atomStart + command[0].length;
    while (true) {
      const argumentStart = skipWhitespaceAndSpacing(source, atomEnd, limit);
      if (source[argumentStart] !== "{") break;
      const argumentEnd = readBalancedGroupEnd(
        source,
        argumentStart,
        "{",
        "}",
        limit,
      );
      if (argumentEnd < 0) return -1;
      atomEnd = argumentEnd;
    }
  } else if (!/[=,+\-*/;&|]/.test(first ?? "")) {
    const character = Array.from(source.slice(atomStart))[0];
    atomEnd = atomStart + character.length;
  }

  if (atomEnd < 0) return -1;
  while (true) {
    const scriptStart = skipWhitespaceAndSpacing(source, atomEnd, limit);
    if (source[scriptStart] !== "^" && source[scriptStart] !== "_") break;
    const scriptEnd = readScriptArgumentEnd(source, scriptStart + 1, limit);
    if (scriptEnd < 0) return -1;
    atomEnd = scriptEnd;
  }
  return atomEnd;
}

function readDifferentialTokenEnd(
  source: string,
  dIndex: number,
  limit = source.length,
): number {
  let variableStart = skipWhitespaceAndSpacing(source, dIndex + 1, limit);
  if (source[variableStart] === "^") {
    variableStart = skipWhitespaceAndSpacing(
      source,
      readExponentEnd(source, variableStart),
      limit,
    );
  }

  const variableEnd = readMathAtomEnd(source, variableStart, limit);
  return variableEnd < 0 ? -1 : variableEnd;
}

function differentialOperatorAtArgumentStart(
  source: string,
  range: BracedArgumentRange,
  allowBareD: boolean,
): number | null {
  const dIndex = skipWhitespaceAndSpacing(
    source,
    range.contentStart,
    range.contentEnd,
  );
  if (source[dIndex] !== "d") return null;
  if (source[dIndex - 1] === "\\" || isAsciiLetter(source[dIndex - 1])) {
    return null;
  }

  let operandStart = skipWhitespaceAndSpacing(source, dIndex + 1, range.contentEnd);
  if (source[operandStart] === "^") {
    operandStart = skipWhitespaceAndSpacing(
      source,
      readExponentEnd(source, operandStart),
      range.contentEnd,
    );
  }

  return operandStart < range.contentEnd || allowBareD ? dIndex : null;
}

function denominatorDifferentialPositions(
  source: string,
  range: BracedArgumentRange,
): number[] {
  const positions: number[] = [];
  let index = skipWhitespaceAndSpacing(
    source,
    range.contentStart,
    range.contentEnd,
  );

  while (index < range.contentEnd) {
    if (
      source[index] !== "d" ||
      source[index - 1] === "\\" ||
      isAsciiLetter(source[index - 1])
    ) {
      return [];
    }

    const tokenEnd = readDifferentialTokenEnd(source, index, range.contentEnd);
    if (tokenEnd < 0 || tokenEnd > range.contentEnd) return [];
    positions.push(index);
    index = skipWhitespaceAndSpacing(source, tokenEnd, range.contentEnd);
  }

  return positions;
}

function collectFractionDifferentialPositions(source: string): Set<number> {
  const positions = new Set<number>();
  const fractionPattern = /\\(?:dfrac|tfrac|frac)(?![A-Za-z])/g;
  for (const match of source.matchAll(fractionPattern)) {
    const numerator = readBracedArgument(source, match.index! + match[0].length);
    if (!numerator) continue;
    const denominator = readBracedArgument(source, numerator.nextIndex);
    if (!denominator) continue;

    const numeratorD = differentialOperatorAtArgumentStart(source, numerator, true);
    const denominatorDs = denominatorDifferentialPositions(source, denominator);
    if (numeratorD === null || denominatorDs.length === 0) continue;
    positions.add(numeratorD);
    denominatorDs.forEach((position) => positions.add(position));
  }
  return positions;
}

function sourceDepthAt(source: string, end: number): SourceDepth {
  const depth: SourceDepth = {
    braces: 0,
    parentheses: 0,
    brackets: 0,
  };
  for (let index = 0; index < end; index += 1) {
    if (source[index - 1] === "\\") continue;
    if (source[index] === "{") depth.braces += 1;
    else if (source[index] === "}") depth.braces = Math.max(0, depth.braces - 1);
    else if (source[index] === "(") depth.parentheses += 1;
    else if (source[index] === ")") {
      depth.parentheses = Math.max(0, depth.parentheses - 1);
    } else if (source[index] === "[") depth.brackets += 1;
    else if (source[index] === "]") {
      depth.brackets = Math.max(0, depth.brackets - 1);
    }
  }
  return depth;
}

function depthsEqual(left: SourceDepth, right: SourceDepth): boolean {
  return (
    left.braces === right.braces &&
    left.parentheses === right.parentheses &&
    left.brackets === right.brackets
  );
}

function hasIntegralAtSameDepth(source: string, dIndex: number): boolean {
  const differentialDepth = sourceDepthAt(source, dIndex);
  const integralPattern = /\\(?:oiiint|oiint|iiint|iint|oint|int)(?![A-Za-z])/g;
  let found = false;
  for (const match of source.slice(0, dIndex).matchAll(integralPattern)) {
    if (depthsEqual(sourceDepthAt(source, match.index!), differentialDepth)) {
      found = true;
    }
  }
  return found;
}

function isTrailingDifferentialBoundary(source: string, tokenEnd: number): boolean {
  const next = skipWhitespaceAndSpacing(source, tokenEnd);
  if (next >= source.length) return true;
  if (source.startsWith("\\right", next) || source.startsWith("\\\\", next)) {
    return true;
  }
  if (source[next] === "d" && readDifferentialTokenEnd(source, next) >= 0) {
    return true;
  }
  return /[}\])=,+\-;&]/.test(source[next]);
}

function isIntegralTrailingDifferential(
  source: string,
  dIndex: number,
  tokenEnd: number,
): boolean {
  let variableStart = skipWhitespaceAndSpacing(source, dIndex + 1);
  if (source[variableStart] === "^") {
    variableStart = skipWhitespaceAndSpacing(
      source,
      readExponentEnd(source, variableStart),
    );
  }

  const beginsWithGroupedExpression = /[([{]/.test(source[variableStart] ?? "");
  const beginsWithLeftDelimiter = source.startsWith("\\left", variableStart);
  return (
    tokenEnd >= 0 &&
    !beginsWithGroupedExpression &&
    !beginsWithLeftDelimiter &&
    hasIntegralAtSameDepth(source, dIndex) &&
    isTrailingDifferentialBoundary(source, tokenEnd)
  );
}

function normalizeDifferentialD(source: string): string {
  const fractionDifferentials = collectFractionDifferentialPositions(source);
  let result = "";
  let index = 0;

  while (index < source.length) {
    if (
      source[index] !== "d" ||
      source[index - 1] === "\\" ||
      isAsciiLetter(source[index - 1])
    ) {
      result += source[index];
      index += 1;
      continue;
    }

    const tokenEnd = readDifferentialTokenEnd(source, index);
    const isExplicitStructure =
      fractionDifferentials.has(index) ||
      isIntegralTrailingDifferential(source, index, tokenEnd);
    if (!isExplicitStructure) {
      result += source[index];
      index += 1;
      continue;
    }

    result += "\\mathrm{d}";
    index += 1;
  }

  return result;
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
    if (previous === "\\" || isAsciiLetter(previous)) {
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

export function normalizeMathLiveCanonicalUprightCommands(source: string): string {
  return source.replace(
    mathLiveCanonicalUprightPattern,
    (_match, command: string) => mathLiveCanonicalUprightCommands[command],
  );
}

function normalizeTypographyChunk(source: string): string {
  return normalizeImaginaryUnitInExponentials(
    normalizeEulerConstant(
      normalizeDifferentialD(
        normalizeNamedOperators(
          normalizeMathLiveCanonicalUprightCommands(source),
        ),
      ),
    ),
  );
}

export function normalizeChineseTextLatex(source: string): string {
  const normalizedTextCommands = source.replace(
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

export function normalizeMathTypography(source: string): string {
  return transformOutsideProtectedCommands(source, normalizeTypographyChunk);
}

export function normalizeChineseLatex(source: string): string {
  return normalizeMathTypography(normalizeChineseTextLatex(source));
}

export function normalizeMultilineLatex(source: string): string {
  return source
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map(normalizeChineseLatex)
    .join("\n");
}
