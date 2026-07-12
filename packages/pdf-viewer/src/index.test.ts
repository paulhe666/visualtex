import { describe, expect, it } from "vitest";
import type { LayoutBox, NodeKind, VisualNode } from "@visualtex/protocol";
import { findDirectEdit, isDirectlyEditable } from "./index";

const source = { fileId: "file-1", startByte: 10, endByte: 20 };

function node(id: string, kind: NodeKind, text: string): VisualNode {
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

function layout(nodeId: string, confidence: LayoutBox["confidence"] = "high"): LayoutBox {
  return {
    nodeId,
    source,
    rects: [{ page: 1, x: 100, y: 120, width: 180, height: 24 }],
    startMarker: null,
    endMarker: null,
    confidence,
    method: "sync_tex",
  };
}

describe("compiled-page direct editing", () => {
  it("allows high-confidence inline formulas but rejects medium-confidence mappings", () => {
    const formula = node("formula", "inline_math", "E=mc^2");
    expect(isDirectlyEditable(formula, layout(formula.id, "high"))).toBe(true);
    expect(isDirectlyEditable(formula, layout(formula.id, "medium"))).toBe(false);
  });

  it("prioritizes a formula when its SyncTeX rectangle overlaps the paragraph", () => {
    const paragraph = node("paragraph", "paragraph", "Energy is $E=mc^2$.");
    const formula = node("formula", "inline_math", "E=mc^2");
    const edit = findDirectEdit(
      0,
      140,
      130,
      [layout(paragraph.id), layout(formula.id)],
      [paragraph, formula],
    );

    expect(edit?.node.id).toBe(formula.id);
    expect(edit?.draft).toBe("E=mc^2");
  });
});
