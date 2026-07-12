export interface VisualSourceNode {
  kind: string;
  support: "native" | "partial" | "opaque" | "unstable";
  source: { startByte: number; endByte: number };
}

export interface TextReplacement {
  startUtf16: number;
  endUtf16: number;
  text: string;
}

export function byteToUtf16(text: string, byteOffset: number): number {
  let bytes = 0;
  let utf16 = 0;
  for (const scalar of text) {
    if (bytes >= byteOffset) break;
    const scalarBytes = Buffer.byteLength(scalar, "utf8");
    if (bytes + scalarBytes > byteOffset) {
      throw new Error(`Byte offset ${byteOffset} is not a UTF-8 boundary`);
    }
    bytes += scalarBytes;
    utf16 += scalar.length;
  }
  if (bytes !== byteOffset) throw new Error(`Byte offset ${byteOffset} exceeds the document`);
  return utf16;
}

export function visualTextReplacement(
  source: string,
  node: VisualSourceNode,
  content: string,
): TextReplacement | null {
  if (node.support === "opaque" || node.support === "unstable") return null;
  const startUtf16 = byteToUtf16(source, node.source.startByte);
  const endUtf16 = byteToUtf16(source, node.source.endByte);
  const fragment = source.slice(startUtf16, endUtf16);
  let innerStart = 0;
  let innerEnd = fragment.length;

  if (["title", "author", "section", "subsection", "citation", "reference", "footnote", "bibliography"].includes(node.kind)) {
    innerStart = fragment.indexOf("{") + 1;
    innerEnd = fragment.lastIndexOf("}");
    if (innerStart <= 0 || innerEnd < innerStart) return null;
  } else if (node.kind === "inline_math" || node.kind === "display_math") {
    if (fragment.startsWith("\\[") && fragment.endsWith("\\]")) {
      innerStart = 2;
      innerEnd = fragment.length - 2;
    } else if (fragment.startsWith("$$") && fragment.endsWith("$$")) {
      innerStart = 2;
      innerEnd = fragment.length - 2;
    } else if (fragment.startsWith("$") && fragment.endsWith("$")) {
      innerStart = 1;
      innerEnd = fragment.length - 1;
    } else if (fragment.startsWith("\\begin{")) {
      innerStart = fragment.indexOf("}") + 1;
      innerEnd = fragment.lastIndexOf("\\end{");
    } else {
      return null;
    }
  } else if (["abstract", "figure", "table", "list", "theorem"].includes(node.kind)) {
    innerStart = fragment.indexOf("}") + 1;
    innerEnd = fragment.lastIndexOf("\\end{");
    if (innerStart <= 0 || innerEnd < innerStart) return null;
  } else if (node.kind !== "paragraph") {
    return null;
  }

  return {
    startUtf16: startUtf16 + innerStart,
    endUtf16: startUtf16 + innerEnd,
    text: content,
  };
}
