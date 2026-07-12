import type { VisualTeXFormulaMetadata } from "../metadata/formulaMetadata";

export interface CompanionHealth {
  ok: boolean;
  appVersion: string;
  officeUiVersion: string;
  protocolVersion: number;
  ocrAvailable: boolean;
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
