import { describe, expect, it } from "vitest";
import type {
  LayoutBox,
  NodeKind,
  PdfTextGlyph,
  PdfTextHit,
  PdfTextLine,
  SourceSpan,
  VisualNode,
} from "@visualtex/protocol";
import {
  findDirectEdit,
  findDirectEditFromSource,
  filterVisualLinesForParagraph,
  findDirectEditFromTextHit,
  isDirectlyEditable,
  sourceByteOffsetAtLineColumn,
  splitParagraphDraftByVisualLines,
} from "./index";

const defaultSource = { fileId: "file-1", startByte: 10, endByte: 20 };

function node(
  id: string,
  kind: NodeKind,
  text: string,
  source: SourceSpan = defaultSource,
): VisualNode {
  return {
    id,
    kind,
    support: "native",
    source,
    children: [],
    text,
    command: null,
    attributes: {
      placement: null,
      caption: null,
      label: null,
      imagePath: null,
      imageWidth: null,
      columnSpec: null,
      tableRows: [],
    },
  };
}

function glyph(index: number, text: string, x: number): PdfTextGlyph {
  return {
    index,
    text,
    rect: { page: 1, x, y: 120, width: 8, height: 12 },
    fontName: text.match(/[A-Za-z0-9=]/) ? "LatinModernMath" : "FandolSong",
    fontSizePoints: 10,
  };
}

function textHit(glyphs: PdfTextGlyph[], glyphIndex: number): PdfTextHit {
  const selected = glyphs.find((candidate) => candidate.index === glyphIndex)!;
  return {
    pageIndex: 0,
    glyphIndex,
    glyph: selected,
    lineGlyphs: glyphs,
  };
}

function textLine(glyphs: PdfTextGlyph[]): PdfTextLine {
  return {
    pageIndex: 0,
    text: glyphs.map((candidate) => candidate.text).join(""),
    rect: { page: 1, x: 100, y: 120, width: Math.max(8, glyphs.length * 9), height: 14 },
    glyphs,
  };
}

function layout(
  nodeId: string,
  rect = { page: 1, x: 100, y: 120, width: 180, height: 24 },
  confidence: LayoutBox["confidence"] = "high",
  source: SourceSpan = defaultSource,
): LayoutBox {
  return {
    nodeId,
    source,
    rects: [rect],
    startMarker: null,
    endMarker: null,
    confidence,
    method: "sync_tex",
  };
}

describe("compiled-page direct editing", () => {
  it("allows high-confidence inline formulas but rejects medium-confidence mappings", () => {
    const formula = node("formula", "inline_math", "E=mc^2");
    expect(isDirectlyEditable(formula, layout(formula.id, undefined, "high"))).toBe(true);
    expect(isDirectlyEditable(formula, layout(formula.id, undefined, "medium"))).toBe(false);
  });

  it("uses the smallest actual hit rectangle instead of an unconditional formula priority", () => {
    const paragraphSpan = { fileId: "file-1", startByte: 0, endByte: 60 };
    const formulaSpan = { fileId: "file-1", startByte: 20, endByte: 30 };
    const paragraph = node("paragraph", "paragraph", "Energy is $E=mc^2$ and more text.", paragraphSpan);
    const formula = node("formula", "inline_math", "E=mc^2", formulaSpan);
    const paragraphLayout = layout(
      paragraph.id,
      { page: 1, x: 80, y: 120, width: 280, height: 26 },
      "high",
      paragraphSpan,
    );
    const formulaLayout = layout(
      formula.id,
      { page: 1, x: 175, y: 120, width: 72, height: 26 },
      "high",
      formulaSpan,
    );

    expect(findDirectEdit(0, 110, 130, [paragraphLayout, formulaLayout], [paragraph, formula])?.node.id)
      .toBe(paragraph.id);
    expect(findDirectEdit(0, 200, 130, [paragraphLayout, formulaLayout], [paragraph, formula])?.node.id)
      .toBe(formula.id);
  });

  it("uses inverse SyncTeX source position to select prose outside a nested formula", () => {
    const text = "\\begin{document}\n正文开头 $E=mc^2$ 正文结尾。\n\\end{document}\n";
    const lineStart = sourceByteOffsetAtLineColumn(text, 2, 1)!;
    const paragraphSpan = {
      fileId: "file-1",
      startByte: lineStart,
      endByte: sourceByteOffsetAtLineColumn(text, 2, 27)!,
    };
    const formulaSpan = {
      fileId: "file-1",
      startByte: sourceByteOffsetAtLineColumn(text, 2, 6)!,
      endByte: sourceByteOffsetAtLineColumn(text, 2, 14)!,
    };
    const paragraph = node("paragraph", "paragraph", "正文开头 $E=mc^2$ 正文结尾。", paragraphSpan);
    const formula = node("formula", "inline_math", "E=mc^2", formulaSpan);
    const layouts = [
      layout(paragraph.id, { page: 1, x: 80, y: 120, width: 300, height: 24 }, "high", paragraphSpan),
      layout(formula.id, { page: 1, x: 170, y: 120, width: 80, height: 24 }, "high", formulaSpan),
    ];

    const lineGlyphs = Array.from("正文开头E=mc2正文结尾。", (character, index) =>
      glyph(index, character, 100 + index * 9),
    );
    const prose = findDirectEditFromSource(
      0,
      310,
      130,
      text,
      "main.tex",
      { sourcePath: "C:/paper/main.tex", line: 2, column: 18, offset: null },
      null,
      layouts,
      [paragraph, formula],
    );
    const math = findDirectEditFromSource(
      0,
      200,
      130,
      text,
      "main.tex",
      { sourcePath: "C:/paper/main.tex", line: 2, column: 8, offset: null },
      null,
      layouts,
      [paragraph, formula],
    );
    const mathWithoutColumn = findDirectEditFromSource(
      0,
      200,
      130,
      text,
      "main.tex",
      { sourcePath: "C:/paper/main.tex", line: 2, column: null, offset: null },
      textHit(lineGlyphs, 7),
      layouts,
      [paragraph, formula],
    );
    const proseWithoutColumn = findDirectEditFromSource(
      0,
      310,
      130,
      text,
      "main.tex",
      { sourcePath: "C:/paper/main.tex", line: 2, column: null, offset: null },
      textHit(lineGlyphs, 15),
      layouts,
      [paragraph, formula],
    );

    expect(prose?.node.id).toBe(paragraph.id);
    expect(math?.node.id).toBe(formula.id);
    expect(mathWithoutColumn?.node.id).toBe(formula.id);
    expect(proseWithoutColumn?.node.id).toBe(paragraph.id);
  });

  it("selects a medium-confidence paragraph from PDF text even when inverse SyncTeX returns the wrong line", () => {
    const text = "\\begin{document}\n这是正文开头 $E=mc^2$ 后面仍然是正文。\n\\end{document}\n";
    const paragraphSpan = {
      fileId: "file-1",
      startByte: sourceByteOffsetAtLineColumn(text, 2, 1)!,
      endByte: sourceByteOffsetAtLineColumn(text, 2, 30)!,
    };
    const formulaSpan = {
      fileId: "file-1",
      startByte: sourceByteOffsetAtLineColumn(text, 2, 8)!,
      endByte: sourceByteOffsetAtLineColumn(text, 2, 16)!,
    };
    const paragraph = node("paragraph", "paragraph", "这是正文开头 $E=mc^2$ 后面仍然是正文。", paragraphSpan);
    paragraph.support = "partial";
    const formula = node("formula", "inline_math", "E=mc^2", formulaSpan);
    const sharedRect = { page: 1, x: 80, y: 120, width: 340, height: 24 };
    const layouts = [
      layout(paragraph.id, sharedRect, "medium", paragraphSpan),
      layout(formula.id, sharedRect, "high", formulaSpan),
    ];
    const lineGlyphs = Array.from("这是正文开头E=mc2后面仍然是正文。", (character, index) =>
      glyph(index, character, 100 + index * 9),
    );
    const proseHit = textHit(lineGlyphs, 16);

    const fromWrongSourceLine = findDirectEditFromSource(
      0,
      300,
      130,
      text,
      "main.tex",
      { sourcePath: "C:/paper/main.tex", line: 3, column: null, offset: null },
      proseHit,
      layouts,
      [paragraph, formula],
    );
    const withoutSyncTex = findDirectEditFromTextHit(
      0,
      300,
      130,
      proseHit,
      layouts,
      [paragraph, formula],
    );

    expect(fromWrongSourceLine?.node.id).toBe(paragraph.id);
    expect(withoutSyncTex?.node.id).toBe(paragraph.id);
  });

  it("filters a drifting SyncTeX line that actually belongs to the following paragraph", () => {
    const source = "第一段正文在这里，并且会自动换行。";
    const first = textLine(Array.from("第一段正文在这里，", (character, index) =>
      glyph(index, character, 100 + index * 9),
    ));
    const second = textLine(Array.from("并且会自动换行。", (character, index) =>
      glyph(100 + index, character, 100 + index * 9),
    ));
    const drifted = textLine(Array.from("下一段正文不应显示。", (character, index) =>
      glyph(200 + index, character, 100 + index * 9),
    ));

    expect(filterVisualLinesForParagraph(source, [first, second, drifted]))
      .toEqual([first, second]);
  });

  it("splits a paragraph into the PDF visual lines while retaining inline formula source", () => {
    const source = "第一行正文比较长，第二部分包含 $E=mc^2$，最后还有一些正文。";
    const firstGlyphs = Array.from("第一行正文比较长，", (character, index) =>
      glyph(index, character, 100 + index * 9),
    );
    const secondGlyphs = Array.from("第二部分包含E=mc2，", (character, index) =>
      glyph(100 + index, character, 100 + index * 9),
    );
    const thirdGlyphs = Array.from("最后还有一些正文。", (character, index) =>
      glyph(200 + index, character, 100 + index * 9),
    );

    expect(splitParagraphDraftByVisualLines(source, [
      textLine(firstGlyphs),
      textLine(secondGlyphs),
      textLine(thirdGlyphs),
    ])).toEqual([
      "第一行正文比较长，",
      "第二部分包含 $E=mc^2$，",
      "最后还有一些正文。",
    ]);
  });

  it("converts Unicode line and character columns to UTF-8 byte offsets", () => {
    const text = "第一行\n甲乙ABC\n";
    expect(sourceByteOffsetAtLineColumn(text, 2, 1)).toBe(new TextEncoder().encode("第一行\n").length);
    expect(sourceByteOffsetAtLineColumn(text, 2, 3)).toBe(new TextEncoder().encode("第一行\n甲乙").length);
  });
});
