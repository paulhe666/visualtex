// Word's equation numbering defaults in a configuration file (the "word" block): the same format in both apps.
// Reading and applying them is the app's (src/platform/wordPreferences.ts).

export interface VisualTexConfigurationWordPreferences {
  defaultDisplayEquationNumbered?: boolean;
  defaultEquationNumberFormat?: string;
}

const validWordNumberFormats = new Set([
  "continuous",
  "heading1-dot",
  "heading1-dash",
  "heading2-dot",
  "heading2-dash",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

export function normalizeWordPreferences(value: unknown): VisualTexConfigurationWordPreferences | undefined {
  if (!isRecord(value)) return undefined;
  const result: VisualTexConfigurationWordPreferences = {};
  if (typeof value.defaultDisplayEquationNumbered === "boolean") {
    result.defaultDisplayEquationNumbered = value.defaultDisplayEquationNumbered;
  }
  if (
    typeof value.defaultEquationNumberFormat === "string" &&
    validWordNumberFormats.has(value.defaultEquationNumberFormat)
  ) {
    result.defaultEquationNumberFormat = value.defaultEquationNumberFormat;
  }
  return Object.keys(result).length ? result : undefined;
}
