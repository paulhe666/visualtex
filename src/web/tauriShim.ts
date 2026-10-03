// The web build has no Tauri host. Vite aliases the Tauri packages used by the
// synced macOS sources to this module, so their browser branches run unchanged.

export function isTauri(): boolean {
  return false;
}

export async function invoke<T>(command: string): Promise<T> {
  throw new Error(`"${command}" is only available in the VisualTeX desktop app`);
}

export function getCurrentWindow(): never {
  throw new Error("Window controls are only available in the VisualTeX desktop app");
}

export async function save(): Promise<string | null> {
  return null;
}
