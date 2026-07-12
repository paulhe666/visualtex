export interface OcrCandidate {
  latex: string;
  confidence: number;
  backend: string;
}

export interface FormulaOcrResult {
  candidates: OcrCandidate[];
  modelVersion: string | null;
  warnings: string[];
}

export interface OcrRegion {
  kind: string;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string | null;
  latex: string | null;
  confidence: number;
}

export interface DocumentOcrResult {
  imagePath: string | null;
  pageWidth: number;
  pageHeight: number;
  regions: OcrRegion[];
  readingOrder: number[];
  modelVersion: string | null;
  warnings: string[];
}

export function escapeLatexText(value: string): string {
  return value
    .replaceAll("\\", "\\textbackslash{}")
    .replaceAll("{", "\\{")
    .replaceAll("}", "\\}")
    .replaceAll("%", "\\%")
    .replaceAll("$", "\\$")
    .replaceAll("#", "\\#")
    .replaceAll("&", "\\&")
    .replaceAll("_", "\\_")
    .replaceAll("^", "\\textasciicircum{}")
    .replaceAll("~", "\\textasciitilde{}");
}

export function ocrRegionContent(region: OcrRegion): string {
  return region.latex ?? region.text ?? "";
}

export function documentOcrToLatex(result: DocumentOcrResult): string {
  const blocks: string[] = [];
  const seen = new Set<number>();
  for (const regionIndex of result.readingOrder) {
    if (!Number.isInteger(regionIndex) || seen.has(regionIndex)) continue;
    seen.add(regionIndex);
    const region = result.regions[regionIndex];
    if (!region) continue;
    const kind = normalizeKind(region.kind);
    const content = ocrRegionContent(region).trim();
    if (!content || kind === "ignore") continue;

    if (kind.includes("formula")) {
      blocks.push(`\\begin{equation}\n  ${content}\n\\end{equation}`);
    } else if (kind.includes("document_title")) {
      blocks.push(`\\section*{${escapeLatexText(content)}}`);
    } else if (kind.includes("paragraph_title") || kind === "title") {
      blocks.push(`\\section{${escapeLatexText(content)}}`);
    } else if (kind.includes("abstract")) {
      blocks.push(`\\begin{abstract}\n${escapeLatexText(content)}\n\\end{abstract}`);
    } else if (kind.includes("table")) {
      blocks.push(markdownTableToLatex(content) ?? `\\paragraph{OCR table}\n${escapeLatexText(content)}`);
    } else if (kind.includes("list")) {
      const items = content
        .split("\n")
        .map((line) => line.replace(/^\s*[-•·*]\s*/, "").trim())
        .filter(Boolean);
      blocks.push([
        "\\begin{itemize}",
        ...items.map((item) => `  \\item ${escapeLatexText(item)}`),
        "\\end{itemize}",
      ].join("\n"));
    } else if (kind.includes("reference")) {
      blocks.push(`\\section*{References}\n${escapeLatexText(content)}`);
    } else if (kind.includes("image") || kind.includes("figure") || kind.includes("chart")) {
      blocks.push(`% OCR ${region.kind} region: ${escapeLatexText(content)}`);
    } else if (!isPageDecoration(kind)) {
      blocks.push(escapeLatexText(content));
    }
  }
  return blocks.join("\n\n");
}

function markdownTableToLatex(content: string): string | null {
  const rows = content
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.includes("|"))
    .map((line) => line.replace(/^\||\|$/g, "").split("|").map((cell) => cell.trim()))
    .filter((row) => !row.every((cell) => /^:?-{3,}:?$/.test(cell)));
  if (rows.length === 0) return null;
  const columns = Math.max(...rows.map((row) => row.length));
  if (columns === 0 || columns > 20) return null;
  return [
    "\\begin{table}[htbp]",
    "  \\centering",
    `  \\begin{tabular}{${"l".repeat(columns)}}`,
    "    \\hline",
    ...rows.map((row) => {
      const normalized = [
        ...row.map(escapeLatexText),
        ...Array.from({ length: columns - row.length }, () => ""),
      ];
      return `    ${normalized.join(" & ")} \\\\`;
    }),
    "    \\hline",
    "  \\end{tabular}",
    "\\end{table}",
  ].join("\n");
}

function normalizeKind(kind: string): string {
  return kind.trim().toLowerCase().replaceAll("-", "_").replaceAll(" ", "_");
}

function isPageDecoration(kind: string): boolean {
  return kind.includes("header") || kind.includes("footer") || kind.includes("page_number");
}
