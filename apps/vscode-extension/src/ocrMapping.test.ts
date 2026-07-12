import { describe, expect, it } from "vitest";
import { documentOcrToLatex, type DocumentOcrResult, type OcrRegion } from "./ocrMapping";

function region(kind: string, content: string, confidence = 0.9): OcrRegion {
  return {
    kind,
    x: 0,
    y: 0,
    width: 100,
    height: 40,
    text: kind.includes("formula") ? null : content,
    latex: kind.includes("formula") ? content : null,
    confidence,
  };
}

function document(regions: OcrRegion[], readingOrder: number[]): DocumentOcrResult {
  return {
    imagePath: null,
    pageWidth: 1000,
    pageHeight: 1400,
    regions,
    readingOrder,
    modelVersion: "test",
    warnings: [],
  };
}

describe("documentOcrToLatex", () => {
  it("uses the reviewed reading order and skips duplicate/ignored indices", () => {
    const result = document(
      [region("text", "正文_A"), region("paragraph_title", "结果"), region("ignore", "跳过")],
      [1, 0, 1, 2, 99],
    );
    expect(documentOcrToLatex(result)).toBe("\\section{结果}\n\n正文\\_A");
  });

  it("converts formula, list and markdown table blocks", () => {
    const result = document(
      [
        region("formula", "E=mc^2"),
        region("list", "- 第一项\n• 第二项"),
        region("table", "| A | B |\n|---|---|\n| 1 | 2 |"),
      ],
      [0, 1, 2],
    );
    const latex = documentOcrToLatex(result);
    expect(latex).toContain("\\begin{equation}\n  E=mc^2\n\\end{equation}");
    expect(latex).toContain("\\item 第一项");
    expect(latex).toContain("A & B \\\\");
  });
});
