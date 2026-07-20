interface OfficeLikeError {
  name?: unknown;
  message?: unknown;
  code?: unknown;
  debugInfo?: unknown;
}

function readableString(value: unknown) {
  return typeof value === "string" && value.trim() ? value.trim() : null;
}

function debugLocation(debugInfo: unknown) {
  if (!debugInfo || typeof debugInfo !== "object") return null;
  const value = debugInfo as Record<string, unknown>;
  return (
    readableString(value.errorLocation) ??
    readableString(value.location) ??
    readableString(value.statement)
  );
}

export class OfficeIntegrationError extends Error {
  constructor(
    message: string,
    readonly code = "office_error",
    readonly retryable = false,
  ) {
    super(message);
    this.name = "OfficeIntegrationError";
  }
}

export function officeErrorMessage(error: unknown, fallback: string) {
  if (typeof error === "string" && error.trim()) return error.trim();

  if (error && typeof error === "object") {
    const value = error as OfficeLikeError;
    const message = readableString(value.message);
    const name = readableString(value.name);
    const code = readableString(value.code);
    const location = debugLocation(value.debugInfo);
    const details = [
      code ? `code=${code}` : null,
      location ? `location=${location}` : null,
    ].filter(Boolean);

    if (message) {
      return details.length ? `${message} (${details.join(", ")})` : message;
    }
    if (name) {
      return details.length ? `${name} (${details.join(", ")})` : name;
    }
    if (details.length) return `${fallback} (${details.join(", ")})`;
  }

  return fallback;
}

export function withTimeout<T>(
  operation: Promise<T>,
  milliseconds: number,
  message: string,
): Promise<T> {
  let timeout: ReturnType<typeof setTimeout> | undefined;
  return Promise.race([
    operation,
    new Promise<T>((_, reject) => {
      timeout = setTimeout(
        () => reject(new OfficeIntegrationError(message, "timeout", true)),
        milliseconds,
      );
    }),
  ]).finally(() => {
    if (timeout !== undefined) clearTimeout(timeout);
  });
}
