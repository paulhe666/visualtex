import { describe, expect, it } from "vitest";
import { completionContextAt, computeSourceChange, sourcePositionAtUtf8Byte } from "./index";

describe("computeSourceChange", () => {
  it("uses UTF-8 byte offsets for Chinese text", () => {
    expect(computeSourceChange("你好 world", "您好 world")).toEqual({
      startByte: 0,
      endByte: 3,
      replacement: "您",
    });
  });

  it("does not split a surrogate pair", () => {
    expect(computeSourceChange("A😀B", "A😄B")).toEqual({
      startByte: 1,
      endByte: 5,
      replacement: "😄",
    });
  });

  it("returns null for identical strings", () => {
    expect(computeSourceChange("same", "same")).toBeNull();
  });
});

describe("sourcePositionAtUtf8Byte", () => {
  it("maps UTF-8 byte spans to one-based source positions without splitting Unicode", () => {
    const source = "第一行\nA😀中B";
    const byteAtChinese = new TextEncoder().encode("第一行\nA😀").length;
    expect(sourcePositionAtUtf8Byte(source, byteAtChinese)).toEqual({
      line: 2,
      column: 4,
      utf16Offset: 7,
    });
  });

  it("clamps byte offsets that land inside a scalar to its start", () => {
    expect(sourcePositionAtUtf8Byte("A中B", 2)).toEqual({
      line: 1,
      column: 2,
      utf16Offset: 1,
    });
  });
});

describe("completionContextAt", () => {
  it("detects label completion inside reference commands", () => {
    const prefix = "See \\cref{sec:int";
    expect(completionContextAt(prefix)).toEqual({
      kind: "label",
      fromOffset: prefix.indexOf("sec:int"),
    });
  });

  it("completes only the current citation key after a comma", () => {
    const prefix = "\\citep{einstein1905, kn";
    expect(completionContextAt(prefix)).toEqual({
      kind: "citation",
      fromOffset: prefix.length - 3,
    });
  });

  it("detects command completion from the backslash", () => {
    expect(completionContextAt("Text \\vec")).toEqual({
      kind: "command",
      fromOffset: 5,
    });
  });

  it("does not offer completions in ordinary text", () => {
    expect(completionContextAt("ordinary text")).toBeNull();
  });
});
