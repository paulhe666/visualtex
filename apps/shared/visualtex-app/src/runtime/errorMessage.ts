const ERROR_DETAIL_KEYS = [
  "message",
  "error",
  "description",
  "details",
  "reason",
  "cause",
] as const;

function nestedErrorMessage(reason: unknown, seen: Set<object>): string {
  if (typeof reason === "string") {
    const value = reason.trim();
    if (!value || /^\[object (?:Object|Error)\]$/i.test(value)) return "";
    if ((value.startsWith("{") && value.endsWith("}")) ||
        (value.startsWith("[") && value.endsWith("]"))) {
      try {
        const parsed = JSON.parse(value) as unknown;
        const parsedMessage = nestedErrorMessage(parsed, seen);
        if (parsedMessage) return parsedMessage;
      } catch {
        // Preserve ordinary non-JSON strings below.
      }
    }
    return value;
  }
  if (reason === null || reason === undefined) return "";
  if (typeof reason !== "object") return String(reason).trim();
  if (seen.has(reason)) return "";
  seen.add(reason);

  if (reason instanceof Error) {
    const directMessage = nestedErrorMessage(reason.message, seen);
    if (directMessage) return directMessage;
    const cause = (reason as Error & { cause?: unknown }).cause;
    const causeMessage = nestedErrorMessage(cause, seen);
    if (causeMessage) return causeMessage;
    for (const key of Object.getOwnPropertyNames(reason)) {
      if (["name", "message", "stack", "cause"].includes(key)) continue;
      const propertyMessage = nestedErrorMessage(
        (reason as unknown as Record<string, unknown>)[key],
        seen,
      );
      if (propertyMessage) return propertyMessage;
    }
    const name = reason.name.trim();
    if (name && name.toLowerCase() !== "error") return name;
  }

  const record = reason as Record<string, unknown>;
  for (const key of ERROR_DETAIL_KEYS) {
    const message = nestedErrorMessage(record[key], seen);
    if (message) return message;
  }

  // structured errors without a message: "code: arg, arg" (or the id)
  const code = typeof record.code === "string" ? record.code.trim() : "";
  const id = typeof record.id === "string" ? record.id.trim() : "";
  const args = Array.isArray(record.args)
    ? record.args.filter((value): value is string => typeof value === "string" && value.trim() !== "").map((value) => value.trim())
    : [];
  const structured = [code || id, args.join(", ")].filter(Boolean).join(": ");
  if (structured) return structured;

  try {
    const encoded = JSON.stringify(reason, (_key, value) =>
      typeof value === "bigint" ? value.toString() : value,
    );
    if (encoded && encoded !== "{}") return encoded;
  } catch {
    // Cyclic or host-provided objects fall through to the caller's fallback.
  }
  return "";
}

/** A readable message for anything thrown or rejected: an Error, a string (also JSON text), an object
 *  with a message-like field, or a structured {code, args} error; else the fallback. */
export function errorMessage(reason: unknown, fallback = "发生未知错误。"): string {
  return nestedErrorMessage(reason, new Set<object>()) || fallback;
}

export async function responseErrorMessage(
  response: Response,
  fallback = "VisualTeX service request failed.",
): Promise<string> {
  const text = await response.text().catch(() => "");
  if (text.trim()) {
    try {
      return errorMessage(JSON.parse(text) as unknown, text.trim());
    } catch {
      return errorMessage(text, fallback);
    }
  }
  return errorMessage(response.statusText, fallback);
}
