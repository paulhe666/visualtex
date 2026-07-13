import {
  OFFICE_BRIDGE_PROTOCOL_VERSION,
  type OfficeBridgeEvent,
  type OfficeBridgeMethod,
  type OfficeBridgeRequest,
  type OfficeBridgeResponse,
} from "../shared/protocol";
import { OfficeIntegrationError, withTimeout } from "../shared/errors";

function installToken() {
  return (
    window.__VISUALTEX_INSTALL_TOKEN__ ??
    document
      .querySelector<HTMLMetaElement>('meta[name="visualtex-install-token"]')
      ?.content ??
    ""
  );
}

export async function callWindowsOle<TResult>(
  method: OfficeBridgeMethod,
  params: Record<string, unknown> = {},
  timeoutMs = 15_000,
): Promise<TResult> {
  const request: OfficeBridgeRequest = {
    protocolVersion: OFFICE_BRIDGE_PROTOCOL_VERSION,
    id: crypto.randomUUID(),
    method,
    params,
  };
  const response = await withTimeout(
    fetch("/api/v1/windows/bridge", {
      method: "POST",
      cache: "no-store",
      credentials: "same-origin",
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
        "X-VisualTeX-Install-Token": installToken(),
      },
      body: JSON.stringify(request),
    }),
    timeoutMs,
    `Windows Office Bridge 请求超时：${method}`,
  );
  const payload = (await response.json().catch(() => null)) as
    | OfficeBridgeResponse<TResult>
    | null;
  if (!response.ok || !payload?.ok) {
    throw new OfficeIntegrationError(
      payload?.error?.message ?? `Windows Office Bridge 调用失败：${method}`,
      payload?.error?.code ?? "windows_bridge_failed",
      payload?.error?.retryable ?? false,
    );
  }
  return payload.result as TResult;
}

export async function getWindowsOleEvents(cursor: number) {
  const response = await fetch(
    `/api/v1/windows/events?cursor=${encodeURIComponent(String(cursor))}`,
    {
      cache: "no-store",
      credentials: "same-origin",
      headers: {
        Accept: "application/json",
        "X-VisualTeX-Install-Token": installToken(),
      },
    },
  );
  if (!response.ok) return [];
  return (await response.json()) as Array<
    OfficeBridgeEvent & { cursor: number }
  >;
}
