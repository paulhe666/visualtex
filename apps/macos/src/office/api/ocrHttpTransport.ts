import type { OcrTransport } from "../../ocr/ocrService";

export interface TransportEvent<T> {
  event: string;
  id: number;
  payload: T;
}

export type TransportEventHandler<T> = (event: TransportEvent<T>) => void;
export type TransportUnlistenFn = () => void;

interface OcrEventEnvelope {
  cursor: number;
  events: Array<{
    id: number;
    event: string;
    payload: unknown;
  }>;
}

const INSTALL_TOKEN_HEADER = "X-VisualTeX-Install-Token";
const POLL_INTERVAL_MS = 160;

function installToken() {
  return (
    window.__VISUALTEX_INSTALL_TOKEN__ ??
    document
      .querySelector<HTMLMetaElement>('meta[name="visualtex-install-token"]')
      ?.content ??
    ""
  );
}

function authenticatedHeaders(extra: HeadersInit = {}) {
  return {
    Accept: "application/json",
    [INSTALL_TOKEN_HEADER]: installToken(),
    ...extra,
  };
}

async function readJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(
      `VisualTeX OCR API failed (${response.status})${detail ? `: ${detail}` : ""}`,
    );
  }
  if (response.status === 204) return undefined as T;
  return (await response.json()) as T;
}

function requestPayload(args: Record<string, unknown> | undefined) {
  const nested = args?.request;
  return nested && typeof nested === "object"
    ? (nested as Record<string, unknown>)
    : args;
}

function bytePayload(args: Record<string, unknown> | undefined) {
  const payload = requestPayload(args);
  const candidate =
    payload?.imageBytes ??
    payload?.image_bytes ??
    payload?.bytes ??
    payload?.data;
  if (candidate instanceof Uint8Array) return candidate;
  if (candidate instanceof ArrayBuffer) return new Uint8Array(candidate);
  if (ArrayBuffer.isView(candidate)) {
    return new Uint8Array(
      candidate.buffer,
      candidate.byteOffset,
      candidate.byteLength,
    );
  }
  if (Array.isArray(candidate)) return Uint8Array.from(candidate);
  throw new Error("OCR recognition requires image bytes.");
}

function modelHeader(args: Record<string, unknown> | undefined) {
  const payload = requestPayload(args);
  const model = payload?.model ?? payload?.modelName ?? payload?.model_name;
  return typeof model === "string" && model.trim()
    ? model.trim()
    : "PP-FormulaNet_plus-M";
}

function extensionHeader(args: Record<string, unknown> | undefined) {
  const payload = requestPayload(args);
  const extension = payload?.extension;
  return typeof extension === "string" && extension.trim()
    ? extension.trim().toLowerCase()
    : "png";
}

export async function invoke<T>(
  command: string,
  args?: Record<string, unknown>,
): Promise<T> {
  switch (command) {
    case "get_ocr_runtime_status": {
      const forceRefresh = args?.forceRefresh === true;
      return readJson<T>(
        await fetch(
          `/api/v1/ocr/status?forceRefresh=${forceRefresh ? "true" : "false"}`,
          {
          cache: "no-store",
          credentials: "same-origin",
            headers: authenticatedHeaders(),
          },
        ),
      );
    }
    case "install_ocr_runtime":
      return readJson<T>(
        await fetch("/api/v1/ocr/install", {
          method: "POST",
          cache: "no-store",
          credentials: "same-origin",
          headers: authenticatedHeaders(),
        }),
      );
    case "recognize_formula_image": {
      const bytes = bytePayload(args);
      return readJson<T>(
        await fetch("/api/v1/ocr/recognize", {
          method: "POST",
          cache: "no-store",
          credentials: "same-origin",
          headers: authenticatedHeaders({
            "Content-Type": "application/octet-stream",
            "X-VisualTeX-Ocr-Model": modelHeader(args),
            "X-VisualTeX-Ocr-Extension": extensionHeader(args),
          }),
          body: bytes,
        }),
      );
    }
    case "prewarm_ocr_model":
      return readJson<T>(
        await fetch("/api/v1/ocr/prewarm", {
          method: "POST",
          cache: "no-store",
          credentials: "same-origin",
          headers: authenticatedHeaders({
            "X-VisualTeX-Ocr-Model": modelHeader(args),
          }),
        }),
      );
    case "cancel_ocr_recognition":
      return readJson<T>(
        await fetch("/api/v1/ocr/cancel", {
          method: "POST",
          cache: "no-store",
          credentials: "same-origin",
          headers: authenticatedHeaders(),
        }),
      );
    case "restart_ocr_worker":
      return readJson<T>(
        await fetch("/api/v1/ocr/restart", {
          method: "POST",
          cache: "no-store",
          credentials: "same-origin",
          headers: authenticatedHeaders(),
        }),
      );
    case "reset_ocr_runtime":
      return readJson<T>(
        await fetch("/api/v1/ocr/reset", {
          method: "POST",
          cache: "no-store",
          credentials: "same-origin",
          headers: authenticatedHeaders(),
        }),
      );
    default:
      throw new Error(`Unsupported Office OCR command: ${command}`);
  }
}

export async function listen<T>(
  eventName: string,
  handler: TransportEventHandler<T>,
): Promise<TransportUnlistenFn> {
  let stopped = false;
  let cursor = 0;
  let timer = 0;

  const baseline = await readJson<OcrEventEnvelope>(
    await fetch(
      `/api/v1/ocr/events?event=${encodeURIComponent(eventName)}`,
      {
        cache: "no-store",
        credentials: "same-origin",
        headers: authenticatedHeaders(),
      },
    ),
  );
  cursor = baseline.cursor;

  const poll = async () => {
    if (stopped) return;
    try {
      const response = await fetch(
        `/api/v1/ocr/events?cursor=${encodeURIComponent(cursor)}&event=${encodeURIComponent(eventName)}`,
        {
          cache: "no-store",
          credentials: "same-origin",
          headers: authenticatedHeaders(),
        },
      );
      const envelope = await readJson<OcrEventEnvelope>(response);
      cursor = Math.max(cursor, envelope.cursor);
      for (const event of envelope.events) {
        if (event.event !== eventName) continue;
        handler({
          event: event.event,
          id: event.id,
          payload: event.payload as T,
        });
      }
    } catch {
      // The next polling pass retries automatically. Command failures are
      // still surfaced through invoke(), while transient event polling errors
      // must not abort an OCR operation.
    } finally {
      if (!stopped) timer = window.setTimeout(poll, POLL_INTERVAL_MS);
    }
  };

  void poll();
  return () => {
    stopped = true;
    window.clearTimeout(timer);
  };
}

export const officeOcrTransport: OcrTransport = {
  environment: "office",
  invoke,
  listen,
};

export type Event<T> = TransportEvent<T>;
export type UnlistenFn = TransportUnlistenFn;
