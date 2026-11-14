import { UnfilledPlaceholderError } from "../../math/latexCompatibility";

// The reason a formula could not be converted or redrawn, as the Office
// add-ins show it: one specific sentence per cause, in the user's language.
// The converter window words its own failures here; the add-ins (C#,
// OfficeErrorText) word the transport and Office failures.

/** A failure whose message is already worded for the user (in their language). */
export class OfficeConversionError extends Error {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = "OfficeConversionError";
  }
}

export function conversionText(isEn: boolean, zh: string, en: string) {
  return isEn ? en : zh;
}

/** Messages of the shared export code (English) and their Chinese wording. */
const KNOWN: { pattern: RegExp; zh: (m: RegExpMatchArray) => string; en: (m: RegExpMatchArray) => string }[] = [
  {
    pattern: /^Unsupported LaTeX: (.+)$/,
    zh: (m) => `VisualTeX 无法排版这些 LaTeX 命令：${m[1]}`,
    en: (m) => `VisualTeX cannot typeset these LaTeX commands: ${m[1]}`,
  },
  { pattern: /^Cannot export an empty formula\.$/, zh: () => "公式为空。", en: () => "The formula is empty." },
  {
    pattern: /^VisualTeX has no outlines for the fonts (.+)\.$/,
    zh: (m) => `VisualTeX 缺少字体 ${m[1]} 的字形轮廓，无法绘制公式图片。`,
    en: (m) => `VisualTeX has no glyph outlines for the fonts ${m[1]}, so the formula image cannot be drawn.`,
  },
  {
    pattern: /^MathJax did not resolve LaTeX command (.+)\.$/,
    zh: (m) => `MathJax 无法识别 LaTeX 命令 ${m[1]}。`,
    en: (m) => `MathJax does not know the LaTeX command ${m[1]}.`,
  },
  {
    pattern: /^MathJax produced an error node: (.+)$/,
    zh: (m) => `MathJax 解析出错：${m[1]}`,
    en: (m) => `MathJax could not parse the formula: ${m[1]}`,
  },
  {
    pattern: /^Unable to rasterize the generated SVG\.$/,
    zh: () => "无法把公式图片（SVG）转换为 PNG。",
    en: () => "The formula image (SVG) could not be converted to PNG.",
  },
  {
    pattern: /^Unable to create a PNG canvas context\.$/,
    zh: () => "无法创建用于生成 PNG 的画布（图片可能过大）。",
    en: () => "No canvas could be created for the PNG (the image may be too large).",
  },
  {
    pattern: /^The PNG of this formula is empty/,
    zh: () => "生成的 PNG 为空（公式图片可能过大）。",
    en: () => "The generated PNG is empty (the formula image may be too large).",
  },
];

function messageOf(reason: unknown): string {
  if (reason instanceof Error) {
    const message = reason.message.trim();
    if (message) return message;
    return reason.name && reason.name !== "Error" ? reason.name : "";
  }
  if (typeof reason === "string") return reason.trim();
  if (reason && typeof reason === "object") {
    const nested = (reason as { message?: unknown; error?: unknown }).message ?? (reason as { error?: unknown }).error;
    if (typeof nested === "string" && nested.trim()) return nested.trim();
    try {
      return JSON.stringify(reason).slice(0, 500);
    } catch {
      return "";
    }
  }
  return reason === undefined || reason === null ? "" : String(reason);
}

/** One formula's failure, worded for the user. */
export function describeConversionFailure(reason: unknown, isEn: boolean): string {
  if (reason instanceof UnfilledPlaceholderError) {
    return conversionText(isEn, "公式中仍有空白占位框，请先填写或删除。", "The formula still has empty boxes. Fill or delete them first.");
  }
  if (reason instanceof OfficeConversionError) return reason.message;
  const message = messageOf(reason);
  if (!message) {
    return conversionText(
      isEn,
      `转换窗口在处理该公式时出错，但没有给出原因（${reason instanceof Error ? reason.name : typeof reason}）。`,
      `The conversion window failed on this formula without giving a reason (${reason instanceof Error ? reason.name : typeof reason}).`,
    );
  }
  for (const known of KNOWN) {
    const match = message.match(known.pattern);
    if (match) return isEn ? known.en(match) : known.zh(match);
  }
  return message;
}

/** Put a stage before a lower-level reason ("MathML could not be generated: …"). */
export function inConversionStage(isEn: boolean, zhStage: string, enStage: string, reason: unknown): OfficeConversionError {
  if (reason instanceof OfficeConversionError) return reason;
  const detail = describeConversionFailure(reason, isEn);
  return new OfficeConversionError(isEn ? `${enStage}: ${detail}` : `${zhStage}：${detail}`, { cause: reason });
}
