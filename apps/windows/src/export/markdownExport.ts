export interface MarkdownExportOptions {
  includeTitle?: boolean;
}

function normalizeMarkdownTitle(title: string) {
  return title.replace(/\r?\n/g, " ").trim();
}

export function buildMarkdownDocument(
  title: string,
  formulas: readonly string[],
  options: MarkdownExportOptions = {},
) {
  const sections: string[] = [];
  const normalizedTitle = normalizeMarkdownTitle(title);

  if (options.includeTitle !== false && normalizedTitle) {
    sections.push(`# ${normalizedTitle}`);
  }

  for (const formula of formulas) {
    const normalizedFormula = formula.trim();
    if (!normalizedFormula) continue;
    sections.push(`$$\n${normalizedFormula}\n$$`);
  }

  return sections.length ? `${sections.join("\n\n")}\n` : "";
}
