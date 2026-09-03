export interface ValidatedQuickOcrCapture {
  dataBase64: string;
  extension: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

const MAX_QUICK_OCR_BASE64_LENGTH = 28 * 1024 * 1024;

export function decodeQuickOcrCapture(
  value: unknown,
): ValidatedQuickOcrCapture | null {
  if (value === null) return null;
  if (!isRecord(value)) {
    throw new Error("VisualTeX Quick OCR returned invalid capture data.");
  }
  const dataBase64 = value.dataBase64;
  const extension = value.extension;
  if (
    typeof dataBase64 !== "string" ||
    dataBase64.length === 0 ||
    dataBase64.length > MAX_QUICK_OCR_BASE64_LENGTH
  ) {
    throw new Error("VisualTeX Quick OCR returned invalid image data.");
  }
  if (
    typeof extension !== "string" ||
    !/^[A-Za-z0-9]{1,10}$/.test(extension)
  ) {
    throw new Error("VisualTeX Quick OCR returned an invalid image extension.");
  }
  return {
    dataBase64,
    extension: extension.toLowerCase(),
  };
}
