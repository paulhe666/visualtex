interface TauriCloseRequestEvent {
  preventDefault(): void;
}

function unavailable(): never {
  throw new Error("Tauri transport is unavailable in the independent Office bundle.");
}

export function isTauriRuntimeAvailable() {
  return false;
}

export async function invokeTauri<T>(
  _command: string,
  _args?: Record<string, unknown>,
): Promise<T> {
  return unavailable();
}

export async function closeCurrentTauriWindow() {
  unavailable();
}

export async function onCurrentTauriWindowCloseRequested(
  _handler: (event: TauriCloseRequestEvent) => void,
): Promise<() => void> {
  unavailable();
}
