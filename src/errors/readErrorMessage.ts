const GENERIC_OBJECT_TEXT = "[object Object]";

function usefulString(value: unknown) {
  if (typeof value !== "string") return "";
  const normalized = value.trim();
  return normalized && normalized !== GENERIC_OBJECT_TEXT ? normalized : "";
}

export function readErrorMessage(error: unknown, fallback = "发生未知错误。"): string {
  if (error instanceof Error) {
    const message = usefulString(error.message);
    if (message) return message;
  }

  const direct = usefulString(error);
  if (direct) return direct;

  if (error && typeof error === "object") {
    const record = error as Record<string, unknown>;
    for (const key of [
      "message",
      "error",
      "detail",
      "reason",
      "description",
      "cause",
    ]) {
      const value = record[key];
      const message = usefulString(value);
      if (message) return message;
      if (value && typeof value === "object" && value !== error) {
        const nested: string = readErrorMessage(value, "");
        if (nested) return nested;
      }
    }

    const id = usefulString(record.id);
    const code = usefulString(record.code);
    const args = Array.isArray(record.args)
      ? record.args
          .map((value) => usefulString(value))
          .filter(Boolean)
      : [];
    const structured = [code || id, args.join(", ")].filter(Boolean).join(": ");
    if (structured) return structured;

    try {
      const serialized = JSON.stringify(error);
      if (serialized && serialized !== "{}") return serialized;
    } catch {
      // Cyclic and host objects fall through to the stable fallback below.
    }
  }

  return fallback;
}

export async function readResponseErrorMessage(
  response: Response,
  fallback = "VisualTeX 服务请求失败。",
): Promise<string> {
  const text = await response.text().catch(() => "");
  if (text.trim()) {
    try {
      const parsed = JSON.parse(text) as unknown;
      return readErrorMessage(parsed, text.trim());
    } catch {
      const message = usefulString(text);
      if (message) return message;
    }
  }
  return usefulString(response.statusText) || fallback;
}
