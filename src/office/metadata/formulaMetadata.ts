export interface VisualTeXFormulaMetadata {
  schema: "visualtex-formula";
  schemaVersion: 1;

  formulaId: string;

  title: string;
  latex: string;

  lines: Array<{
    id: string;
    latex: string;
  }>;

  codeFormat: string;
  displayMode: "inline" | "block";

  createdWithVersion: string;
  updatedWithVersion: string;

  createdAt: string;
  updatedAt: string;
}

export const VISUALTEX_FORMULA_SCHEMA = "visualtex-formula" as const;
export const VISUALTEX_FORMULA_SCHEMA_VERSION = 1 as const;

export function isVisualTeXFormulaMetadata(
  value: unknown,
): value is VisualTeXFormulaMetadata {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<VisualTeXFormulaMetadata>;
  return (
    candidate.schema === VISUALTEX_FORMULA_SCHEMA &&
    candidate.schemaVersion === VISUALTEX_FORMULA_SCHEMA_VERSION &&
    typeof candidate.formulaId === "string" &&
    typeof candidate.latex === "string" &&
    Array.isArray(candidate.lines)
  );
}
