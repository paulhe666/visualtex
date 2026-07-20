import type { VisualTeXFormulaMetadata } from "./formulaMetadata";
import {
  invokeTauri,
  isTauriRuntimeAvailable,
} from "./tauriTransport";

export type OfficeSessionMode = "create" | "edit";
export type OfficeHost = "word" | "powerpoint";

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
  pngBase64?: string;
  ommlBase64?: string;
  ommlDocxBase64?: string;
  width: number;
  height: number;
  baseline?: number;
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
  numbered: boolean;
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
  numbered?: boolean;
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

export function isMacosOfflineTauriTransport() {
  if (typeof window === "undefined") return false;
  const transport = new URLSearchParams(window.location.search).get("transport");
  return transport === "tauri" && isTauriRuntimeAvailable();
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

  const response = await fetch(path, {
    ...init,
    credentials: "same-origin",
    cache: "no-store",
    headers,
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(
      `VisualTeX companion request failed (${response.status})${detail ? `: ${detail}` : ""}`,
    );
  }
  if (response.status === 204) return undefined as T;
  return (await response.json()) as T;
}

export function createOfficeSession(input: CreateOfficeSessionInput) {
  return requestJson<OfficeFormulaSession>("/api/v1/sessions", {
    method: "POST",
    body: JSON.stringify(input),
  });
}

export function getOfficeSession(sessionId: string) {
  if (isMacosOfflineTauriTransport()) {
    return invokeTauri<OfficeFormulaSession>(
      "get_macos_offline_office_session",
      { sessionId },
    );
  }
  return requestJson<OfficeFormulaSession>(
    `/api/v1/sessions/${encodeURIComponent(sessionId)}`,
  );
}

export function updateOfficeSession(
  sessionId: string,
  update: UpdateOfficeSessionInput,
) {
  if (isMacosOfflineTauriTransport()) {
    return invokeTauri<OfficeFormulaSession>(
      "update_macos_offline_office_session",
      { sessionId, patch: update },
    );
  }
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

export function deleteOfficeSession(sessionId: string) {
  if (isMacosOfflineTauriTransport()) {
    return invokeTauri<void>("delete_macos_offline_office_session", {
      sessionId,
    });
  }
  return requestJson<void>(
    `/api/v1/sessions/${encodeURIComponent(sessionId)}`,
    { method: "DELETE" },
  );
}

export function commitMacosOfflineOfficeSession(sessionId: string) {
  return invokeTauri<OfficeFormulaSession>(
    "commit_macos_offline_office_session",
    { sessionId },
  );
}

export function cancelMacosOfflineOfficeSession(sessionId: string) {
  return invokeTauri<OfficeFormulaSession>(
    "cancel_macos_offline_office_session",
    { sessionId },
  );
}

export function saveOfficeSessionKeepalive(
  sessionId: string,
  update: UpdateOfficeSessionInput,
) {
  if (isMacosOfflineTauriTransport()) {
    return invokeTauri<OfficeFormulaSession>(
      "update_macos_offline_office_session",
      { sessionId, patch: update },
    ).then(() => new Response(null, { status: 204 }));
  }
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
