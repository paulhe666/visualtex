import assert from "node:assert/strict";
import {
  DOCUMENT_IMPORT_MAX_FILE_BYTES,
  decodeDocumentImportBytes,
  documentFormatFromFileName,
  readDocumentImportFile,
} from "../src/office/documentImport/documentImportFile.ts";

function utf16Le(value: string, bom = false) {
  const bytes = new Uint8Array(value.length * 2 + (bom ? 2 : 0));
  let offset = 0;
  if (bom) {
    bytes[0] = 0xff;
    bytes[1] = 0xfe;
    offset = 2;
  }
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    bytes[offset + index * 2] = code & 0xff;
    bytes[offset + index * 2 + 1] = code >> 8;
  }
  return bytes;
}

function utf16Be(value: string, bom = false) {
  const bytes = new Uint8Array(value.length * 2 + (bom ? 2 : 0));
  let offset = 0;
  if (bom) {
    bytes[0] = 0xfe;
    bytes[1] = 0xff;
    offset = 2;
  }
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    bytes[offset + index * 2] = code >> 8;
    bytes[offset + index * 2 + 1] = code & 0xff;
  }
  return bytes;
}

assert.equal(documentFormatFromFileName("paper.TEX"), "latex");
assert.equal(documentFormatFromFileName("notes.md"), "markdown");
assert.equal(documentFormatFromFileName("README.Markdown"), "markdown");
assert.throws(() => documentFormatFromFileName("paper.txt"), /只支持导入/);
assert.throws(() => documentFormatFromFileName("no-extension"), /只支持导入/);

const utf8 = decodeDocumentImportBytes(
  new TextEncoder().encode("# 标题\r\n\r\n正文 $x=1$\r\n"),
);
assert.equal(utf8.encoding, "UTF-8");
assert.equal(utf8.source, "# 标题\n\n正文 $x=1$\n");

const utf8Bom = decodeDocumentImportBytes(
  new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode("# BOM")]),
);
assert.equal(utf8Bom.encoding, "UTF-8 BOM");
assert.equal(utf8Bom.source, "# BOM");

for (const [bytes, encoding] of [
  [utf16Le("\\section{中文}\r\n正文", true), "UTF-16 LE"],
  [utf16Le("\\section{中文}\r\n正文"), "UTF-16 LE"],
  [utf16Be("# 中文\r\n正文", true), "UTF-16 BE"],
  [utf16Be("# 中文\r\n正文"), "UTF-16 BE"],
] as const) {
  const decoded = decodeDocumentImportBytes(bytes);
  assert.equal(decoded.encoding, encoding);
  assert.ok(decoded.source.includes("中文"));
  assert.ok(!decoded.source.includes("\r"));
}

const gb18030 = decodeDocumentImportBytes(
  new Uint8Array([0xd6, 0xd0, 0xce, 0xc4]),
);
assert.equal(gb18030.encoding, "GB18030");
assert.equal(gb18030.source, "中文");

assert.throws(() => decodeDocumentImportBytes(new Uint8Array()), /文件为空/);
assert.throws(
  () => decodeDocumentImportBytes(new Uint8Array([0xfe, 0xff, 0x00])),
  /不完整的字符数据/,
);
assert.equal(
  decodeDocumentImportBytes(new TextEncoder().encode("page\fbreak")).source,
  "page\fbreak",
);
assert.throws(
  () => decodeDocumentImportBytes(new Uint8Array([0, 1, 2, 3, 0, 4])),
  /二进制内容|不像有效的文本源码/,
);

const markdownFile = new File(
  [new TextEncoder().encode("# Imported\n\n$E=mc^2$")],
  "document.md",
  { type: "text/markdown" },
);
const importedMarkdown = await readDocumentImportFile(markdownFile);
assert.equal(importedMarkdown.name, "document.md");
assert.equal(importedMarkdown.format, "markdown");
assert.equal(importedMarkdown.encoding, "UTF-8");
assert.equal(importedMarkdown.source, "# Imported\n\n$E=mc^2$");

const latexFile = new File([utf16Le("\\section{导入}\n正文", true)], "paper.tex");
const importedLatex = await readDocumentImportFile(latexFile);
assert.equal(importedLatex.format, "latex");
assert.equal(importedLatex.encoding, "UTF-16 LE");
assert.ok(importedLatex.source.includes("\\section{导入}"));

await assert.rejects(
  () => readDocumentImportFile(new File(["text"], "paper.txt")),
  /只支持导入/,
);

const oversized = {
  name: "large.tex",
  size: DOCUMENT_IMPORT_MAX_FILE_BYTES + 1,
  arrayBuffer: async () => new ArrayBuffer(0),
} as File;
await assert.rejects(() => readDocumentImportFile(oversized), /超过 5 MB/);

console.log("Document import file decoding regression passed");
