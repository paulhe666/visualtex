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
