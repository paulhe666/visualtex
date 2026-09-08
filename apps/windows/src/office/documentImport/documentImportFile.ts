import type { DocumentImportLanguage, DocumentSourceFormat } from "./documentImportParser";

function L(language: DocumentImportLanguage, chinese: string, english: string) {
  return language === "en" ? english : chinese;
}

export const DOCUMENT_IMPORT_MAX_FILE_BYTES = 5_000_000;

export type DocumentImportFileEncoding =
  | "UTF-8"
  | "UTF-8 BOM"
  | "UTF-16 LE"
  | "UTF-16 BE"
  | "GB18030";

export interface ImportedDocumentFile {
  name: string;
  source: string;
  format: Exclude<DocumentSourceFormat, "auto">;
  encoding: DocumentImportFileEncoding;
  size: number;
}

function extensionOf(name: string) {
  const dot = name.lastIndexOf(".");
  return dot >= 0 ? name.slice(dot).toLowerCase() : "";
}

export function documentFormatFromFileName(
  name: string,
  language: DocumentImportLanguage = "cn",
): Exclude<DocumentSourceFormat, "auto"> {
  const extension = extensionOf(name);
  if (extension === ".tex") return "latex";
  if (extension === ".md" || extension === ".markdown") return "markdown";
  throw new Error(L(language, "只支持导入 .tex、.md 或 .markdown 文件。", "Only .tex, .md and .markdown files are supported."));
}

function stripLeadingBom(value: string) {
  return value.charCodeAt(0) === 0xfeff ? value.slice(1) : value;
}

function normalizeSource(value: string) {
  return stripLeadingBom(value).replace(/\r\n?/g, "\n");
}

function decodeUtf16Be(bytes: Uint8Array, language: DocumentImportLanguage) {
  if (bytes.byteLength % 2 !== 0) {
    throw new Error(L(language, "UTF-16 BE 文件包含不完整的字符数据。", "The UTF-16 BE file contains incomplete character data."));
  }
  const swapped = new Uint8Array(bytes.byteLength);
  for (let index = 0; index < bytes.byteLength; index += 2) {
    swapped[index] = bytes[index + 1];
    swapped[index + 1] = bytes[index];
  }
  return new TextDecoder("utf-16le", { fatal: true }).decode(swapped);
}

function looksLikeUtf16WithoutBom(bytes: Uint8Array) {
  const sampleLength = Math.min(bytes.byteLength - (bytes.byteLength % 2), 1024);
  if (sampleLength < 4) return null;
  let evenZeros = 0;
  let oddZeros = 0;
  const pairs = sampleLength / 2;
  for (let index = 0; index < sampleLength; index += 2) {
    if (bytes[index] === 0) evenZeros += 1;
    if (bytes[index + 1] === 0) oddZeros += 1;
  }
  const evenRatio = evenZeros / pairs;
  const oddRatio = oddZeros / pairs;
  if (oddRatio >= 0.25 && evenRatio <= 0.08) return "le" as const;
  if (evenRatio >= 0.25 && oddRatio <= 0.08) return "be" as const;
  return null;
}

function assertLooksLikeText(value: string, language: DocumentImportLanguage) {
  if (!value.trim()) throw new Error(L(language, "所选文件为空，无法导入。", "The selected file is empty and cannot be imported."));
  let suspicious = 0;
  for (const character of value.slice(0, 8192)) {
    const code = character.charCodeAt(0);
    if (code === 0) throw new Error(L(language, "所选文件包含二进制内容，无法作为文档源码导入。", "The selected file contains binary data and cannot be imported as document source."));
    if (
      code < 32 &&
      character !== "\n" &&
      character !== "\r" &&
      character !== "\t" &&
      character !== "\f"
    ) {
      suspicious += 1;
    }
  }
  if (suspicious > 0) {
    throw new Error(L(language, "所选文件不像有效的文本源码，无法导入。", "The selected file does not appear to be valid text source and cannot be imported."));
  }
}

export function decodeDocumentImportBytes(bytes: Uint8Array, language: DocumentImportLanguage = "cn"): {
  source: string;
  encoding: DocumentImportFileEncoding;
} {
  if (bytes.byteLength === 0) throw new Error(L(language, "所选文件为空，无法导入。", "The selected file is empty and cannot be imported."));

  let source: string;
  let encoding: DocumentImportFileEncoding;

  if (
    bytes.byteLength >= 3 &&
    bytes[0] === 0xef &&
    bytes[1] === 0xbb &&
    bytes[2] === 0xbf
  ) {
    source = new TextDecoder("utf-8", { fatal: true }).decode(bytes.subarray(3));
    encoding = "UTF-8 BOM";
  } else if (bytes.byteLength >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    source = new TextDecoder("utf-16le", { fatal: true }).decode(bytes.subarray(2));
    encoding = "UTF-16 LE";
  } else if (bytes.byteLength >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    source = decodeUtf16Be(bytes.subarray(2), language);
    encoding = "UTF-16 BE";
  } else {
    const utf16 = looksLikeUtf16WithoutBom(bytes);
    if (utf16 === "le") {
      source = new TextDecoder("utf-16le", { fatal: true }).decode(bytes);
      encoding = "UTF-16 LE";
    } else if (utf16 === "be") {
      source = decodeUtf16Be(bytes, language);
      encoding = "UTF-16 BE";
    } else {
      try {
        source = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
        encoding = "UTF-8";
      } catch {
        try {
          source = new TextDecoder("gb18030", { fatal: true }).decode(bytes);
          encoding = "GB18030";
        } catch {
          throw new Error(L(language, "无法识别文件编码。请将文件保存为 UTF-8、UTF-16 或 GB18030 后重试。", "Unable to detect the file encoding. Save the file as UTF-8, UTF-16 or GB18030 and try again."));
        }
      }
    }
  }

  source = normalizeSource(source);
  assertLooksLikeText(source, language);
  return { source, encoding };
}

export async function readDocumentImportFile(
  file: File,
  language: DocumentImportLanguage = "cn",
): Promise<ImportedDocumentFile> {
  const format = documentFormatFromFileName(file.name, language);
  if (file.size > DOCUMENT_IMPORT_MAX_FILE_BYTES) {
    throw new Error(L(language, "文件超过 5 MB，无法批量导入。", "The file is larger than 5 MB and cannot be imported."));
  }
  const bytes = new Uint8Array(await file.arrayBuffer());
  const decoded = decodeDocumentImportBytes(bytes, language);
  return {
    name: file.name,
    source: decoded.source,
    format,
    encoding: decoded.encoding,
    size: file.size,
  };
}
