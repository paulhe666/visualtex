export type CopyFormat = "plain" | "inline" | "display" | "equation";

export const copyFormatLabels: Record<CopyFormat, { title: string; hint: string }> = {
  display: { title: "独立公式（推荐）", hint: "$$ ... $$" },
  plain: { title: "纯公式源码", hint: "\\frac{x}{y}" },
  inline: { title: "行内公式", hint: "\\( ... \\)" },
  equation: { title: "equation 环境", hint: "\\begin{equation}" },
};

export function splitLatexLines(latex: string): string[] {
  const lines = latex.replace(/\r\n?/g, "\n").split("\n");
  return lines.length ? lines : [""];
}

export function formatLatex(latex: string, format: CopyFormat): string {
  const lines = splitLatexLines(latex).filter((line) => line.trim().length > 0);
  const safeLines = lines.length ? lines : [""];

  switch (format) {
    case "inline":
      return safeLines.map((line) => "\\(" + line + "\\)").join("\n");
    case "display":
      return safeLines.map((line) => "$$" + line + "$$").join("\n");
    case "equation":
      return safeLines
        .map((line) => "\\begin{equation}\n" + line + "\n\\end{equation}")
        .join("\n\n");
    default:
      return safeLines.join("\n");
  }
}

export function parseLatexSource(source: string): string[] {
  const normalized = source.replace(/\r\n?/g, "\n").trim();
  if (!normalized) return [""];

  const dollarBlocks = [...normalized.matchAll(/\$\$([\s\S]*?)\$\$/g)];
  if (dollarBlocks.length) {
    return dollarBlocks.map((match) => match[1].trim());
  }

  const bracketBlocks = [...normalized.matchAll(/\\\[([\s\S]*?)\\\]/g)];
  if (bracketBlocks.length) {
    return bracketBlocks.map((match) => match[1].trim());
  }

  const equationBlocks = [
    ...normalized.matchAll(/\\begin\{equation\*?\}([\s\S]*?)\\end\{equation\*?\}/g),
  ];
  if (equationBlocks.length) {
    return equationBlocks.map((match) => match[1].trim());
  }

  return normalized
    .split("\n")
    .map((line) =>
      line
        .trim()
        .replace(/^\\\(/, "")
        .replace(/\\\)$/, "")
        .replace(/^\$\$/, "")
        .replace(/\$\$$/, ""),
    );
}

export async function copyLatex(latex: string, format: CopyFormat = "display") {
  await navigator.clipboard.writeText(formatLatex(latex, format));
}
