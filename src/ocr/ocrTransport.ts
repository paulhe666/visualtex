import { invoke as tauriInvoke } from "@tauri-apps/api/core";
import {
  listen as tauriListen,
  type Event,
  type UnlistenFn,
} from "@tauri-apps/api/event";
import type { OcrTransport, OcrTransportEvent } from "./ocrService";

export const desktopOcrTransport: OcrTransport = {
  environment: "desktop",
  invoke<T>(command: string, args?: Record<string, unknown>) {
    return tauriInvoke<T>(command, args);
  },
  listen<T>(
    eventName: string,
    handler: (event: OcrTransportEvent<T>) => void,
  ): Promise<UnlistenFn> {
    return tauriListen<T>(eventName, (event: Event<T>) => {
      handler({
        event: event.event,
        id: event.id,
        payload: event.payload,
      });
    });
  },
};
