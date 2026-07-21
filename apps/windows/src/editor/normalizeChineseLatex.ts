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

function hasIntegralBefore(source: string, index: number): boolean {
  const prefix = source.slice(0, index);
  return ["\\iiint", "\\iint", "\\oint", "\\oiint", "\\int"].some(
    (command) => prefix.lastIndexOf(command) >= 0,
  );
}

function isDifferentialContext(source: string, index: number): boolean {
  if (index === 0) return true;
  const previous = source[index - 1];
  if (previous === "\\" || isAsciiLetter(previous)) return false;
  if (/\s|[{}()[\]/=,+\-]/.test(previous)) return true;
  if (
    ["\\,", "\\;", "\\:", "\\!", "\\quad", "\\qquad"].some(
      (spacing) => source.slice(0, index).endsWith(spacing),
    )
  ) {
    return true;
  }
  return hasIntegralBefore(source, index);
}

function normalizeDifferentialD(source: string): string {
  let result = "";
  let index = 0;

  while (index < source.length) {
    if (source[index] !== "d" || !isDifferentialContext(source, index)) {
      result += source[index];
      index += 1;
      continue;
    }

    let tailStart = index + 1;
    while (/\s/.test(source[tailStart] ?? "")) tailStart += 1;
    if (source[tailStart] === "^") {
      tailStart = readExponentEnd(source, tailStart);
    }
    const variableEnd = readDifferentialVariableEnd(source, tailStart);
    if (variableEnd < 0) {
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

function normalizeMathLiveCanonicalUprightCommands(source: string): string {
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
