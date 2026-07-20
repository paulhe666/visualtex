import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";

export interface TauriCloseRequestEvent {
  preventDefault(): void;
}

export function isTauriRuntimeAvailable() {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

export function invokeTauri<T>(
  command: string,
  args?: Record<string, unknown>,
) {
  return invoke<T>(command, args);
}

export async function closeCurrentTauriWindow() {
  await getCurrentWindow().close();
}

export function onCurrentTauriWindowCloseRequested(
  handler: (event: TauriCloseRequestEvent) => void,
) {
  return getCurrentWindow().onCloseRequested(handler);
}
