export type FormulaSelectionToggleKind = "bold" | "italic";

export interface FormulaSelectionStyleState {
  allBold: boolean;
  allBoldItalic: boolean;
  allItalic: boolean;
  allUpright: boolean;
}

interface OuterLatexCommand {
  command: string;
  body: string;
}

const boldCommands = new Set([
  "mathbf",
  "mathbfit",
  "boldsymbol",
  "bm",
  "bold",
]);

const uprightCommands = new Set(["mathrm", "textrm", "mathbf"]);
const italicCommands = new Set([
  "mathit",
  "mathbfit",
  "boldsymbol",
  "bm",
]);
const mathVariantCommands = new Set([
  ...boldCommands,
  ...uprightCommands,
  ...italicCommands,
  "mathnormal",
]);

function findMatchingBrace(source: string, openIndex: number) {
  let depth = 0;
  for (let index = openIndex; index < source.length; index += 1) {
    const token = source[index];
    if (token === "\\") {
      index += 1;
      continue;
    }
    if (token === "{") depth += 1;
    else if (token === "}") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}

export function parseOuterLatexCommand(source: string): OuterLatexCommand | null {
  const normalized = source.trim();
  const commandMatch = normalized.match(/^\\([A-Za-z]+)\s*\{/);
  if (!commandMatch) return null;
  const openIndex = normalized.indexOf("{", commandMatch[0].length - 1);
  if (openIndex < 0) return null;
  const closeIndex = findMatchingBrace(normalized, openIndex);
  if (closeIndex < 0 || normalized.slice(closeIndex + 1).trim()) return null;
  return {
    command: commandMatch[1],
    body: normalized.slice(openIndex + 1, closeIndex),
  };
}

function transformCommandWrappers(
  source: string,
  transform: (command: string, body: string) => string | null,
): string {
  let result = "";
  for (let index = 0; index < source.length; index += 1) {
    if (source[index] !== "\\") {
      result += source[index];
      continue;
    }
    const commandMatch = source.slice(index).match(/^\\([A-Za-z]+)\s*\{/);
    if (!commandMatch) {
      result += source[index];
      continue;
    }
    const openIndex = index + commandMatch[0].lastIndexOf("{");
    const closeIndex = findMatchingBrace(source, openIndex);
    if (closeIndex < 0) {
      result += source[index];
      continue;
    }
    const command = commandMatch[1];
    const body = source.slice(openIndex + 1, closeIndex);
    const transformedBody = transformCommandWrappers(body, transform);
    const replacement = transform(command, transformedBody);
    if (replacement === null) {
      result += source.slice(index, openIndex + 1);
      result += transformedBody;
      result += "}";
    } else {
      result += replacement;
    }
    index = closeIndex;
  }
  return result;
}

function stripCommands(source: string, commands: ReadonlySet<string>) {
  return transformCommandWrappers(source, (command, body) =>
    commands.has(command) ? body : null,
  );
}

function stripMathVariantCommands(source: string) {
  return stripCommands(source, mathVariantCommands);
}

function removeBold(source: string) {
  return stripCommands(source, boldCommands);
}

function removeItalicPreservingBold(source: string) {
  return transformCommandWrappers(source, (command, body) => {
    if (command === "mathbfit" || command === "boldsymbol" || command === "bm") {
      return `\\mathbf{${body}}`;
    }
    if (command === "mathit") return body;
    return null;
  });
}

function applyItalicPreservingBold(source: string) {
  return transformCommandWrappers(source, (command, body) => {
    if (command === "mathrm" || command === "textrm" || command === "mathnormal") {
      return body;
    }
    if (command === "mathbf") return `\\mathbfit{${body}}`;
    return null;
  });
}

function rootCommand(source: string) {
  return parseOuterLatexCommand(source)?.command ?? null;
}

function canonicalizeContextualVariantChain(
  source: string,
  queried: FormulaSelectionStyleState,
) {
  const chain: OuterLatexCommand[] = [];
  let current = source.trim();
  while (true) {
    const outer = parseOuterLatexCommand(current);
    if (!outer || !mathVariantCommands.has(outer.command)) break;
    chain.push(outer);
    current = outer.body.trim();
  }
  if (chain.length < 2) return source.trim();
  if (queried.allBoldItalic) return `\\mathbfit{${current}}`;
  if (queried.allBold) return `\\mathbf{${current}}`;
  if (queried.allItalic) return `\\mathit{${current}}`;
  if (queried.allUpright) return `\\mathrm{${current}}`;
  return source.trim();
}

function containsCommand(source: string, commands: ReadonlySet<string>) {
  const pattern = /\\([A-Za-z]+)/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(source))) {
    if (commands.has(match[1])) return true;
  }
  return false;
}

export function inferFormulaSelectionStyleState(
  source: string,
  queried: FormulaSelectionStyleState,
): FormulaSelectionStyleState {
  const command = rootCommand(source);
  const rootBold = command ? boldCommands.has(command) : false;
  const rootItalic = command ? italicCommands.has(command) : false;
  const rootUpright = command ? uprightCommands.has(command) : false;
  const containsUpright = containsCommand(source, uprightCommands);
  return {
    allBold: queried.allBold || rootBold,
    allBoldItalic:
      queried.allBoldItalic || command === "mathbfit" || command === "boldsymbol" || command === "bm",
    allItalic:
      queried.allItalic ||
      rootItalic ||
      (!queried.allUpright && !rootUpright && !rootBold && !containsUpright),
    allUpright: queried.allUpright || rootUpright,
  };
}

export function toggleFormulaSelectionLatex(
  source: string,
  kind: FormulaSelectionToggleKind,
  queriedState: FormulaSelectionStyleState,
) {
  const normalized = canonicalizeContextualVariantChain(source, queriedState);
  if (!normalized) return normalized;
  const state = inferFormulaSelectionStyleState(normalized, queriedState);
  const outer = parseOuterLatexCommand(normalized);

  if (kind === "bold") {
    if (state.allBold || state.allBoldItalic) {
      return removeBold(normalized);
    }
    const body = outer && mathVariantCommands.has(outer.command)
      ? stripMathVariantCommands(outer.body)
      : stripMathVariantCommands(normalized);
    return `\\mathbf{${body}}`;
  }

  if (state.allItalic || state.allBoldItalic) {
    if (outer?.command === "mathbfit"
      || outer?.command === "boldsymbol"
      || outer?.command === "bm") {
      return `\\mathbf{${outer.body}}`;
    }
    if (outer?.command === "mathit") {
      return `\\mathrm{${outer.body}}`;
    }
    return `\\mathrm{${removeItalicPreservingBold(normalized)}}`;
  }

  if (outer?.command === "mathrm"
    || outer?.command === "textrm"
    || outer?.command === "mathnormal") {
    return outer.body;
  }
  if (outer?.command === "mathbf") {
    return `\\mathbfit{${outer.body}}`;
  }
  return applyItalicPreservingBold(normalized);
}
