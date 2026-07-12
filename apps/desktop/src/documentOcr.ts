import type { DocumentOcrResult, OcrRegion } from "@visualtex/protocol";

export const OCR_REGION_KIND_OPTIONS = [
  { value: "text", label: "正文" },
  { value: "paragraph_title", label: "小标题" },
  { value: "document_title", label: "文档标题" },
  { value: "abstract", label: "摘要" },
  { value: "formula", label: "公式" },
  { value: "table", label: "表格" },
  { value: "figure", label: "图片/图表" },
  { value: "list", label: "列表" },
  { value: "reference", label: "参考文献" },
  { value: "header", label: "页眉" },
  { value: "footer", label: "页脚" },
  { value: "page_number", label: "页码" },
  { value: "ignore", label: "忽略" },
] as const;

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

export function markdownTableToLatex(content: string): string | null {
  const rows = content
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.includes("|"))
    .map((line) => line.replace(/^\||\|$/g, "").split("|").map((cell) => cell.trim()))
    .filter((row) => !row.every((cell) => /^:?-{3,}:?$/.test(cell)));
  if (rows.length === 0) return null;
  const columns = Math.max(...rows.map((row) => row.length));
  if (columns === 0 || columns > 20) return null;
  const normalized = rows.map((row) => [
    ...row.map(escapeLatexText),
    ...Array.from({ length: columns - row.length }, () => ""),
  ]);
  return [
    "\\begin{table}[htbp]",
    "  \\centering",
    `  \\begin{tabular}{${"l".repeat(columns)}}`,
    "    \\hline",
    ...normalized.map((row) => `    ${row.join(" & ")} \\\\`),
    "    \\hline",
    "  \\end{tabular}",
    "\\end{table}",
  ].join("\n");
}

export function ocrRegionContent(region: OcrRegion): string {
  return region.latex ?? region.text ?? "";
}

export function updateOcrRegionContent(region: OcrRegion, value: string): OcrRegion {
  return isFormulaKind(region.kind)
    ? { ...region, latex: value }
    : { ...region, text: value };
}

export function changeOcrRegionKind(region: OcrRegion, kind: string): OcrRegion {
  const content = ocrRegionContent(region);
  if (isFormulaKind(kind)) {
    return { ...region, kind, latex: content, text: null };
  }
  return { ...region, kind, text: content, latex: null };
}

export function moveReadingOrder(
  readingOrder: number[],
  regionIndex: number,
  delta: -1 | 1,
): number[] {
  const currentPosition = readingOrder.indexOf(regionIndex);
  if (currentPosition < 0) return readingOrder;
  const nextPosition = currentPosition + delta;
  if (nextPosition < 0 || nextPosition >= readingOrder.length) return readingOrder;
  const next = [...readingOrder];
  [next[currentPosition], next[nextPosition]] = [next[nextPosition]!, next[currentPosition]!];
  return next;
}

export function documentOcrToLatex(result: DocumentOcrResult): string {
  const blocks: string[] = [];
  for (const regionIndex of result.readingOrder) {
    const region = result.regions[regionIndex];
    if (!region) continue;
    const kind = normalizedKind(region.kind);
    const content = ocrRegionContent(region).trim();
    if (!content || kind === "ignore") continue;

    if (isFormulaKind(kind)) {
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
    } else if (!isSuppressedPageDecoration(kind)) {
      blocks.push(escapeLatexText(content));
    }
  }
  return blocks.join("\n\n");
}

function normalizedKind(kind: string): string {
  return kind.trim().toLowerCase().replaceAll("-", "_").replaceAll(" ", "_");
}

function isFormulaKind(kind: string): boolean {
  return normalizedKind(kind).includes("formula");
}

function isSuppressedPageDecoration(kind: string): boolean {
  return kind.includes("header") || kind.includes("footer") || kind.includes("page_number");
}
