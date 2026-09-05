import type { MathfieldElement } from "mathlive";

export function applyVisualTexBoldMacrosToMathfield(field: MathfieldElement) {
  field.macros = {
    ...field.macros,
    bm: {
      def: "\\boldsymbol{#1}",
      args: 1,
      expand: false,
      captureSelection: true,
    },
    mathbfit: {
      def: "\\boldsymbol{#1}",
      args: 1,
      expand: false,
      captureSelection: true,
    },
  };
}

export function expandVisualTexBoldCommands(source: string) {
  if (!source.includes("\\")) return source;
  return source
    .replace(/\\bm(?=\s*\{)/g, "\\boldsymbol")
    .replace(/\\mathbfit(?=\s*\{)/g, "\\boldsymbol");
}

export function isVisualTexBoldCompatibilityCommand(command: string) {
  const normalized = command.trim().replace(/^\\/, "");
  return normalized === "bm" || normalized === "mathbfit";
}
