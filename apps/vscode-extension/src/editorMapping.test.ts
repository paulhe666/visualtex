import { describe, expect, it } from "vitest";
import { byteToUtf16, visualTextReplacement } from "./editorMapping";

function byteOffset(source: string, needle: string): number {
  const utf16 = source.indexOf(needle);
  if (utf16 < 0) throw new Error(`missing ${needle}`);
  return Buffer.byteLength(source.slice(0, utf16), "utf8");
}

describe("byteToUtf16", () => {
  it("maps UTF-8 offsets across Chinese text and emoji", () => {
    const source = "A中😀B";
    expect(byteToUtf16(source, 0)).toBe(0);
    expect(byteToUtf16(source, 1)).toBe(1);
    expect(byteToUtf16(source, 4)).toBe(2);
    expect(byteToUtf16(source, 8)).toBe(4);
    expect(byteToUtf16(source, 9)).toBe(5);
  });

  it("rejects an offset inside a multibyte scalar", () => {
    expect(() => byteToUtf16("中", 1)).toThrow("not a UTF-8 boundary");
  });
});

describe("visualTextReplacement", () => {
  it("replaces only a section argument while preserving the command", () => {
    const source = "前文\n\\section{旧标题}\n正文";
    const startByte = byteOffset(source, "\\section");
    const fragment = "\\section{旧标题}";
    const replacement = visualTextReplacement(
      source,
      {
        kind: "section",
        support: "native",
        source: {
          startByte,
          endByte: startByte + Buffer.byteLength(fragment, "utf8"),
        },
      },
      "新标题😀",
    );
    expect(replacement).not.toBeNull();
    const next = source.slice(0, replacement!.startUtf16)
      + replacement!.text
      + source.slice(replacement!.endUtf16);
    expect(next).toBe("前文\n\\section{新标题😀}\n正文");
  });

  it("replaces only display-math content", () => {
    const source = "\\begin{equation}\na+b=c\n\\end{equation}";
    const replacement = visualTextReplacement(
      source,
      {
        kind: "display_math",
        support: "native",
        source: { startByte: 0, endByte: Buffer.byteLength(source, "utf8") },
      },
      "E=mc^2",
    );
    expect(replacement).not.toBeNull();
    const next = source.slice(0, replacement!.startUtf16)
      + replacement!.text
      + source.slice(replacement!.endUtf16);
    expect(next).toBe("\\begin{equation}E=mc^2\\end{equation}");
  });

  it("refuses opaque nodes", () => {
    expect(visualTextReplacement(
      "\\unknown{value}",
      {
        kind: "raw_latex",
        support: "opaque",
        source: { startByte: 0, endByte: 15 },
      },
      "replacement",
    )).toBeNull();
  });
});
