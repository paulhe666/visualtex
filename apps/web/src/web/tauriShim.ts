// The web build has no Tauri host. Vite aliases the Tauri packages used by the
// synced macOS sources to this module, so their browser branches run unchanged.

export function isTauri(): boolean {
  return false;
}

export async function invoke<T>(command: string): Promise<T> {
  throw new Error(`"${command}" is only available in the VisualTeX desktop app`);
}

export function convertFileSrc(path: string): string {
  return path;
}

// ---- @tauri-apps/api/event: no other windows, nothing arrives ----

export type UnlistenFn = () => void;

export async function listen(): Promise<UnlistenFn> {
  return () => undefined;
}

export async function once(): Promise<UnlistenFn> {
  return () => undefined;
}

export async function emit(): Promise<void> {}

export async function emitTo(): Promise<void> {}

// ---- @tauri-apps/api/window and /webview ----

export function getCurrentWindow(): never {
  throw new Error("Window controls are only available in the VisualTeX desktop app");
}

export function getCurrentWebview(): never {
  throw new Error("Web view events are only available in the VisualTeX desktop app");
}

// ---- @tauri-apps/plugin-dialog: the browser's own dialogs ----

export async function ask(message: string): Promise<boolean> {
  return window.confirm(message);
}

export async function confirm(message: string): Promise<boolean> {
  return window.confirm(message);
}

export async function message(text: string): Promise<void> {
  window.alert(text);
}

export async function save(): Promise<string | null> {
  return null;
}

export async function open(): Promise<null> {
  return null;
}

// ---- @tauri-apps/plugin-opener ----

export async function openUrl(url: string): Promise<void> {
  window.open(url, "_blank", "noopener");
}
