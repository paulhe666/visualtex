import type { MathfieldElement } from "mathlive";

const ESINT_REPLACEMENTS = new Map<string, string>([
  ["fint", "\\mathop{⨏}"],
  ["dashint", "\\mathop{⨍}"],
  ["ddashint", "\\mathop{⨎}"],
  ["oiint", "\\mathop{∯}"],
  ["oiiint", "\\mathop{∰}"],
  ["varointclockwise", "\\mathop{∲}"],
  ["ointctrclockwise", "\\mathop{∳}"],
  ["varointctrclockwise", "\\mathop{∳}"],
]);

function definition(def: string) {
  return {
    def,
    args: 0,
    expand: false,
    captureSelection: true,
  } as const;
}

export function esintMathLiveMacros() {
  return Object.fromEntries(
    [...ESINT_REPLACEMENTS].map(([command, replacement]) => [
      command,
      definition(replacement),
    ]),
  );
}

export function applyEsintMacrosToMathfield(field: MathfieldElement) {
  field.macros = {
    ...field.macros,
    ...esintMathLiveMacros(),
  };
}

export function isEsintCommand(command: string) {
  return ESINT_REPLACEMENTS.has(command.trim().replace(/^\\/, ""));
}

export function expandEsintCommands(source: string) {
  if (!source.includes("\\")) return source;
  let output = "";
  let index = 0;
  while (index < source.length) {
    if (source[index] !== "\\") {
      output += source[index];
      index += 1;
      continue;
    }
    if (source[index + 1] === "\\") {
      output += "\\\\";
      index += 2;
      continue;
    }
    let end = index + 1;
    while (/[A-Za-z]/.test(source[end] ?? "")) end += 1;
    if (end === index + 1) {
      output += source[index];
      index += 1;
      continue;
    }
    const command = source.slice(index + 1, end);
    const replacement = ESINT_REPLACEMENTS.get(command);
    if (replacement) output += replacement;
    else output += source.slice(index, end);
    index = end;
  }
  return output;
}
