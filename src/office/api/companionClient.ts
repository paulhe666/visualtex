import type { VisualTeXFormulaMetadata } from "../metadata/formulaMetadata";

export interface CompanionHealth {
  ok: boolean;
  appVersion: string;
  officeUiVersion: string;
  protocolVersion: number;
  ocrAvailable: boolean;
}

export interface NativePowerPointSelection {
  shapeName: string;
  slideIndex: number;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface PowerPointInteractionEvent {
  cursor: number;
  host: "word" | "powerpoint";
  kind: "edit-selected";
  formulaId: string;
  shapeName: string;
  createdAt: number;
}

export interface NativePowerPointSlideSnapshot {
  presentationIdentity: string;
  slideIndex: number;
  slideId: number;
  shapeCount: number;
  shapeNames: string[];
}

export async function getCompanionHealth(): Promise<CompanionHealth> {
  const response = await fetch("/health", {
    cache: "no-store",
    credentials: "same-origin",
    headers: { Accept: "application/json" },
  });
  if (!response.ok) {
    throw new Error(`VisualTeX companion health check failed (${response.status})`);
  }
  return (await response.json()) as CompanionHealth;
}

function installToken() {
  return (
    window.__VISUALTEX_INSTALL_TOKEN__ ??
    document
      .querySelector<HTMLMetaElement>('meta[name="visualtex-install-token"]')
      ?.content ??
    ""
  );
}

function metadataUrl(formulaId: string) {
  return `/api/v1/formulas/${encodeURIComponent(formulaId)}/metadata`;
}

export async function getCachedFormulaMetadata(formulaId: string) {
  const response = await fetch(metadataUrl(formulaId), {
    cache: "no-store",
    credentials: "same-origin",
    headers: {
      Accept: "application/json",
      "X-VisualTeX-Install-Token": installToken(),
    },
  });
  if (response.status === 404) return null;
  if (!response.ok) {
    throw new Error(`Unable to read VisualTeX formula cache (${response.status})`);
  }
  return (await response.json()) as VisualTeXFormulaMetadata;
}

export async function putCachedFormulaMetadata(
  metadata: VisualTeXFormulaMetadata,
) {
  const response = await fetch(metadataUrl(metadata.formulaId), {
    method: "PUT",
    cache: "no-store",
    credentials: "same-origin",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json",
      "X-VisualTeX-Install-Token": installToken(),
    },
    body: JSON.stringify(metadata),
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(
      `Unable to save VisualTeX formula cache (${response.status})${
        detail ? `: ${detail}` : ""
      }`,
    );
  }
  return (await response.json()) as VisualTeXFormulaMetadata;
}

async function nativePowerPointRequest<T>(
  path: string,
  init: RequestInit = {},
): Promise<T> {
  const response = await fetch(path, {
    ...init,
    cache: "no-store",
    credentials: "same-origin",
    headers: {
      Accept: "application/json",
      ...(init.body ? { "Content-Type": "application/json" } : {}),
      "X-VisualTeX-Install-Token": installToken(),
      ...init.headers,
    },
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => null) as
      | { error?: string }
      | null;
    throw new Error(
      payload?.error ?? `PowerPoint native integration failed (${response.status})`,
    );
  }
  if (response.status === 204) return undefined as T;
  return (await response.json()) as T;
}

export function getNativePowerPointSelection() {
  return nativePowerPointRequest<NativePowerPointSelection>(
    "/api/v1/powerpoint/selection",
  );
}

export function markNativePowerPointSelection(formulaId: string) {
  return nativePowerPointRequest<NativePowerPointSelection>(
    "/api/v1/powerpoint/selection/mark",
    {
      method: "POST",
      body: JSON.stringify({ formulaId }),
    },
  );
}

export function getNativePowerPointSlideSnapshot() {
  return nativePowerPointRequest<NativePowerPointSlideSnapshot>(
    "/api/v1/powerpoint/slide/snapshot",
  );
}

export function markLastNativePowerPointFormula(
  formulaId: string,
  previousShapeNames: string[],
) {
  return nativePowerPointRequest<NativePowerPointSelection>(
    "/api/v1/powerpoint/shape/mark-last",
    {
      method: "POST",
      body: JSON.stringify({ formulaId, previousShapeNames }),
    },
  );
}

export function replaceLastNativePowerPointFormula(
  formulaId: string,
  previousShapeNames: string[],
  originalShapeName: string,
  geometry: { left: number; top: number; width: number; height: number },
) {
  return nativePowerPointRequest<NativePowerPointSelection>(
    "/api/v1/powerpoint/shape/replace-last",
    {
      method: "POST",
      body: JSON.stringify({
        formulaId,
        previousShapeNames,
        originalShapeName,
        ...geometry,
      }),
    },
  );
}

export function deleteNativePowerPointShape(
  slideIndex: number,
  shapeName: string,
) {
  return nativePowerPointRequest<void>("/api/v1/powerpoint/shape/delete", {
    method: "POST",
    body: JSON.stringify({ slideIndex, shapeName }),
  });
}

export function getPowerPointInteractionEvents(cursor: number) {
  return nativePowerPointRequest<PowerPointInteractionEvent[]>(
    `/api/v1/powerpoint/events?cursor=${encodeURIComponent(String(cursor))}`,
  );
}

export async function revealDesktopApp() {
  const response = await fetch("/api/v1/app/reveal", {
    method: "POST",
    cache: "no-store",
    credentials: "same-origin",
    headers: {
      Accept: "application/json",
      "X-VisualTeX-Install-Token": installToken(),
    },
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(
      `无法打开 VisualTeX 桌面窗口 (${response.status})${detail ? `: ${detail}` : ""}`,
    );
  }
}

export async function ensureCompanionReady() {
  try {
    const health = await getCompanionHealth();
    if (!health.ok) throw new Error("VisualTeX companion reported an unhealthy state");
    return health;
  } catch (error) {
    throw new Error(
      "VisualTeX 本地伴侣服务未运行。请先启动 VisualTeX.app。",
      { cause: error },
    );
  }
}
