import { readResponseErrorMessage } from "../../errors/readErrorMessage";
import type { FormulaDocument } from "../../types/formula";
import type {
  FormulaChineseFont,
  FormulaLetterFont,
} from "../../editor/formulaFontPreferences";
import type { CustomThemeState } from "../../themeCustomization";
import type { VisualTeXFormulaMetadata } from "./formulaMetadata";

export type OfficeSessionMode = "create" | "edit";
export type OfficeHost = "word" | "powerpoint";
export type OfficeObjectMode =
  | "nativeOle"
  | "mathTypeOle"
  | "wordOmml"
  | "crossPlatformPicture";

export type OfficeSessionStatus =
  | "created"
  | "editing"
  | "committing"
  | "completed"
  | "cancelled"
  | "failed";

export interface OfficeExportResult {
  svg: string;
  svgBase64: string;
  mathMl?: string;
  pngBase64?: string;
  width: number;
  height: number;
  baseline?: number;
  formulaLetterFont?: FormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
}

export interface NativePowerPointCommitSelection {
  shapeName: string;
  slideIndex: number;
  slideId?: number;
  presentationIdentity?: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface PreparedPowerPointCommit {
  session: OfficeFormulaSession;
  selection: NativePowerPointCommitSelection;
}

export interface OfficeThemeStatus {
  theme: string;
  editorLayout: "standard" | "classic";
}

export interface OfficePreferences {
  powerpointDefaultFontSizePt: number;
  editorPreferences?: {
    settings?: Partial<FormulaDocument["settings"]>;
    customTheme?: CustomThemeState;
  };
}

export interface OfficeBatchConversion {
  sessionIds: string[];
}

export interface OfficeFormulaSession {
  id: string;
  mode: OfficeSessionMode;
  host: OfficeHost;
  formulaId: string;
  sourceDocumentId: string | null;
  sourceObjectId: string | null;
  title: string;
  lines: Array<{ id: string; latex: string }>;
  activeLineId: string | null;
  codeFormat: string;
  displayMode: "inline" | "block";
  objectMode: OfficeObjectMode;
  numbered: boolean;
  fontSizePt: number;
  exportWidth: number;
  exportHeight: number;
  exportResult: OfficeExportResult | null;
  originalMetadata: VisualTeXFormulaMetadata | null;
  dirty: boolean;
  status: OfficeSessionStatus;
  autoCommitOnClose: boolean;
  explicitCancel: boolean;
  error: string | null;
  createdAt: number;
  updatedAt: number;
  expiresAt: number;
}

export interface CreateOfficeSessionInput {
  mode: OfficeSessionMode;
  host: OfficeHost;
  formulaId?: string;
  sourceDocumentId?: string | null;
  sourceObjectId?: string | null;
  title?: string;
  lines?: OfficeFormulaSession["lines"];
  activeLineId?: string | null;
  codeFormat?: string;
  displayMode?: "inline" | "block";
  objectMode?: OfficeObjectMode;
  numbered?: boolean;
  fontSizePt?: number;
  exportWidth?: number;
  exportHeight?: number;
  originalMetadata?: VisualTeXFormulaMetadata | null;
  autoCommitOnClose?: boolean;
}

export type UpdateOfficeSessionInput = Partial<
  Omit<OfficeFormulaSession, "id" | "createdAt">
>;

declare global {
  interface Window {
    __VISUALTEX_INSTALL_TOKEN__?: string;
  }
}

function installToken() {
  if (typeof window === "undefined") return "";
  return (
    window.__VISUALTEX_INSTALL_TOKEN__ ??
    document
      .querySelector<HTMLMetaElement>('meta[name="visualtex-install-token"]')
      ?.content ??
    ""
  );
}

async function requestJson<T>(path: string, init: RequestInit = {}): Promise<T> {
  const token = installToken();
  const headers = new Headers(init.headers);
  headers.set("Accept", "application/json");
  if (init.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  if (token) headers.set("X-VisualTeX-Install-Token", token);

  const timeoutController = new AbortController();
  const timeout = globalThis.setTimeout(() => timeoutController.abort(), 12_000);
  try {
    const response = await fetch(path, {
      ...init,
      credentials: "same-origin",
      cache: "no-store",
      headers,
      signal: init.signal ?? timeoutController.signal,
    });
    if (!response.ok) {
      const detail = await readResponseErrorMessage(
        response,
        "VisualTeX companion request failed.",
      );
      throw new Error(
        `VisualTeX companion request failed (${response.status}): ${detail}`,
      );
    }
    if (response.status === 204) return undefined as T;
    return (await response.json()) as T;
  } catch (reason) {
    if (timeoutController.signal.aborted && !init.signal?.aborted) {
      throw new Error(
        "VisualTeX Office Session 请求超时。请重试；若仍失败，请重启 VisualTeX。",
      );
    }
    throw reason;
  } finally {
    globalThis.clearTimeout(timeout);
  }
}

export function createOfficeSession(input: CreateOfficeSessionInput) {
  return requestJson<OfficeFormulaSession>("/api/v1/sessions", {
    method: "POST",
    body: JSON.stringify(input),
  });
}

export function getOfficeSession(sessionId: string) {
  return requestJson<OfficeFormulaSession>(
    `/api/v1/sessions/${encodeURIComponent(sessionId)}`,
  );
}

export function getOfficeTheme() {
  return requestJson<OfficeThemeStatus>("/api/v1/theme");
}

export function getOfficePreferences() {
  return requestJson<OfficePreferences>("/api/v1/preferences");
}

export function takeOfficeConverterBatch() {
  return requestJson<OfficeBatchConversion>(
    "/api/v1/app/converter/next-batch",
  );
}

export function updateOfficeSession(
  sessionId: string,
  update: UpdateOfficeSessionInput,
) {
  return requestJson<OfficeFormulaSession>(
    `/api/v1/sessions/${encodeURIComponent(sessionId)}`,
    { method: "PATCH", body: JSON.stringify(update) },
  );
}

export function commitNativePowerPointSession(sessionId: string) {
  return requestJson<PreparedPowerPointCommit>(
    `/api/v1/powerpoint/sessions/${encodeURIComponent(sessionId)}/commit`,
    { method: "POST", body: "{}" },
  );
}

export function confirmNativePowerPointSession(sessionId: string) {
  return requestJson<OfficeFormulaSession>(
    `/api/v1/powerpoint/sessions/${encodeURIComponent(sessionId)}/confirm`,
    { method: "POST", body: "{}" },
  );
}

export function commitWindowsOfficeSession(sessionId: string) {
  return requestJson<OfficeFormulaSession>(
    `/api/v1/windows/sessions/${encodeURIComponent(sessionId)}/commit`,
    { method: "POST", body: "{}" },
  );
}

export function commitNativePowerPointSessionKeepalive(
  sessionId: string,
  update: UpdateOfficeSessionInput,
) {
  const headers = new Headers({
    Accept: "application/json",
    "Content-Type": "application/json",
  });
  const token = installToken();
  if (token) headers.set("X-VisualTeX-Install-Token", token);
  return fetch(
    `/api/v1/powerpoint/sessions/${encodeURIComponent(sessionId)}/commit`,
    {
      method: "POST",
      credentials: "same-origin",
      cache: "no-store",
      keepalive: true,
      headers,
      body: JSON.stringify(update),
    },
  );
}

export function closeOfficeSessionWindow(sessionId: string) {
  return requestJson<void>(
    `/api/v1/app/sessions/${encodeURIComponent(sessionId)}/close`,
    { method: "POST", body: "{}" },
  );
}

export function deleteOfficeSession(sessionId: string) {
  return requestJson<void>(
    `/api/v1/sessions/${encodeURIComponent(sessionId)}`,
    { method: "DELETE" },
  );
}

export function saveOfficeSessionKeepalive(
  sessionId: string,
  update: UpdateOfficeSessionInput,
) {
  const headers = new Headers({
    Accept: "application/json",
    "Content-Type": "application/json",
  });
  const token = installToken();
  if (token) headers.set("X-VisualTeX-Install-Token", token);
  return fetch(`/api/v1/sessions/${encodeURIComponent(sessionId)}`, {
    method: "PATCH",
    credentials: "same-origin",
    cache: "no-store",
    keepalive: true,
    headers,
    body: JSON.stringify(update),
  });
}
