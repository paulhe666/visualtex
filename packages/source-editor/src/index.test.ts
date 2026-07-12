import { describe, expect, it } from "vitest";
import { completionContextAt, computeSourceChange } from "./index";

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
