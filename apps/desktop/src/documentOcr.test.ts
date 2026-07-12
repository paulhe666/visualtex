import { describe, expect, it } from "vitest";
import type { DocumentOcrResult, OcrRegion } from "@visualtex/protocol";
import {
  changeOcrRegionKind,
  documentOcrToLatex,
  markdownTableToLatex,
  moveReadingOrder,
  updateOcrRegionContent,
} from "./documentOcr";

function region(kind: string, content: string): OcrRegion {
  return {
    kind,
    x: 0,
    y: 0,
    width: 100,
    height: 40,
    text: kind.includes("formula") ? null : content,
    latex: kind.includes("formula") ? content : null,
    confidence: 0.9,
  };
}

function document(regions: OcrRegion[], readingOrder = regions.map((_, index) => index)): DocumentOcrResult {
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
  it("respects corrected reading order and skips ignored/page decorations", () => {
    const result = document(
      [
        region("text", "正文_A"),
        region("paragraph_title", "结果"),
        region("ignore", "不应出现"),
        region("page_number", "12"),
      ],
      [1, 0, 2, 3],
    );
    expect(documentOcrToLatex(result)).toBe(
      "\\section{结果}\n\n正文\\_A",
    );
  });

  it("emits formula, abstract, list and reference blocks", () => {
    const result = document([
      region("abstract", "摘要内容"),
      region("formula", "E=mc^2"),
      region("list", "- 第一项\n• 第二项"),
      region("reference", "[1] Example"),
    ]);
    const latex = documentOcrToLatex(result);
    expect(latex).toContain("\\begin{abstract}\n摘要内容\n\\end{abstract}");
    expect(latex).toContain("\\begin{equation}\n  E=mc^2\n\\end{equation}");
    expect(latex).toContain("\\item 第一项");
    expect(latex).toContain("\\section*{References}");
  });

  it("converts a simple markdown table", () => {
    expect(markdownTableToLatex("| A | B |\n|---|---|\n| 1 | 2 |")).toContain(
      "A & B \\\\",
    );
  });
});

describe("OCR region corrections", () => {
  it("moves content between text and latex when the kind changes", () => {
    const formula = changeOcrRegionKind(region("text", "x+y"), "formula");
    expect(formula.text).toBeNull();
    expect(formula.latex).toBe("x+y");
    const text = changeOcrRegionKind(formula, "text");
    expect(text.text).toBe("x+y");
    expect(text.latex).toBeNull();
  });

  it("updates the field appropriate to the corrected kind", () => {
    expect(updateOcrRegionContent(region("formula", "x"), "y").latex).toBe("y");
    expect(updateOcrRegionContent(region("text", "a"), "b").text).toBe("b");
  });

  it("moves a region without losing or duplicating indices", () => {
    expect(moveReadingOrder([4, 7, 9], 7, -1)).toEqual([7, 4, 9]);
    expect(moveReadingOrder([4, 7, 9], 7, 1)).toEqual([4, 9, 7]);
    expect(moveReadingOrder([4, 7, 9], 4, -1)).toEqual([4, 7, 9]);
    expect(moveReadingOrder([4, 7, 9], 99, 1)).toEqual([4, 7, 9]);
  });
});
