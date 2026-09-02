import type { MathfieldElement } from "mathlive";

type MathLiveInternalOptionAccess = {
  _getOptions?: (keys?: string | string[]) => Record<string, unknown> | null;
  _setOptions?: (options: Record<string, unknown>) => void;
};

const guardedFields = new WeakSet<MathfieldElement>();

function isTransientMissingModeTarget(error: unknown) {
  if (!(error instanceof TypeError)) return false;
  return /Cannot set properties of undefined \(setting ['"]mode['"]\)/.test(
    error.message,
  );
}

function internalOptionAccess(field: MathfieldElement) {
  return field as unknown as MathLiveInternalOptionAccess;
}

export function readMathLiveOptionBeforeMount<T>(
  field: MathfieldElement,
  key: string,
): T | undefined {
  const options = internalOptionAccess(field)._getOptions?.([key]);
  return options?.[key] as T | undefined;
}

export function setMathLiveOptionsBeforeMount(
  field: MathfieldElement,
  options: Record<string, unknown>,
) {
  if (field.isConnected) {
    throw new Error("MathLive deferred options must be configured before mount");
  }
  const setter = internalOptionAccess(field)._setOptions;
  if (typeof setter !== "function") {
    throw new Error("MathLive deferred option API is unavailable");
  }
  setter.call(field, options);
}

export function installMathLiveOptionMutationGuard(field: MathfieldElement) {
  if (guardedFields.has(field)) return;
  const access = internalOptionAccess(field);
  const original = access._setOptions;
  if (typeof original !== "function") return;

  const retry = (options: Record<string, unknown>, attempt: number) => {
    if (!field.isConnected) return;
    try {
      original.call(field, options);
    } catch (error) {
      if (!isTransientMissingModeTarget(error)) throw error;
      if (attempt >= 2) {
        console.warn(
          "VisualTeX deferred a MathLive option update because its model was temporarily incomplete.",
          error,
        );
        return;
      }
      window.requestAnimationFrame(() => retry(options, attempt + 1));
    }
  };

  access._setOptions = (options: Record<string, unknown>) => {
    try {
      original.call(field, options);
    } catch (error) {
      if (!isTransientMissingModeTarget(error)) throw error;
      window.queueMicrotask(() => retry(options, 0));
    }
  };
  guardedFields.add(field);
}
