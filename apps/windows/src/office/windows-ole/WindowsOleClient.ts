import {
  OFFICE_BRIDGE_PROTOCOL_VERSION,
  type OfficeBridgeMethod,
  type OfficeBridgeRequest,
  type OfficeBridgeResponse,
} from "../shared/protocol";
import { createUuid } from "../../runtime/browserCompatibility";
import { OfficeIntegrationError, withTimeout } from "../shared/errors";
import {
  decodeOfficeBridgeEvents,
  decodeOfficeBridgeResponseEnvelope,
} from "./windowsOlePayloadValidation";

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
  decodeResult?: (value: unknown) => TResult,
): Promise<TResult> {
  const request: OfficeBridgeRequest = {
    protocolVersion: OFFICE_BRIDGE_PROTOCOL_VERSION,
    id: createUuid(),
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
  let rawPayload: unknown;
  try {
    rawPayload = await response.json();
  } catch {
    throw new OfficeIntegrationError(
      `Windows Office Bridge 返回了无效 JSON：${method}`,
      "windows_bridge_invalid_response",
      false,
    );
  }

  let payload: OfficeBridgeResponse<TResult>;
  try {
    payload = decodeOfficeBridgeResponseEnvelope<TResult>(rawPayload, request.id);
  } catch (error) {
    throw new OfficeIntegrationError(
      error instanceof Error
        ? error.message
        : `Windows Office Bridge 返回了无效响应：${method}`,
      "windows_bridge_invalid_response",
      false,
    );
  }
  if (!response.ok || !payload.ok) {
    throw new OfficeIntegrationError(
      payload.error?.message ?? `Windows Office Bridge 调用失败：${method}`,
      payload.error?.code ?? "windows_bridge_failed",
      payload.error?.retryable ?? false,
    );
  }
  try {
    return decodeResult
      ? decodeResult(payload.result)
      : (payload.result as TResult);
  } catch (error) {
    throw new OfficeIntegrationError(
      error instanceof Error
        ? error.message
        : `Windows Office Bridge 返回了无效结果：${method}`,
      "windows_bridge_invalid_result",
      false,
    );
  }
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
  let payload: unknown;
  try {
    payload = await response.json();
  } catch {
    return [];
  }
  try {
    return decodeOfficeBridgeEvents(payload);
  } catch {
    return [];
  }
}
