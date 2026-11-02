import type { MathfieldElement } from "mathlive";

type MathLiveInternalModel = {
  root?: {
    firstChild?: unknown;
  };
};

type MathLiveInternalOptionAccess = {
  _getOptions?: (keys?: string | string[]) => Record<string, unknown> | null;
  _setOptions?: (options: Record<string, unknown>) => void;
  _mathfield?: {
    model?: MathLiveInternalModel;
    options?: Record<string, unknown>;
  } | null;
  reflectAttributes?: () => void;
};

const guardedFields = new WeakSet<MathfieldElement>();

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object";
}

function isTypeErrorLike(error: unknown) {
  return (
    error instanceof TypeError ||
    (isRecord(error) && error.name === "TypeError")
  );
}

function errorDetail(error: unknown) {
  if (!isRecord(error)) return String(error ?? "");
  return [error.name, error.message, error.stack]
    .filter((value): value is string => typeof value === "string")
    .join("\n");
}

function optionValueContains(
  actual: unknown,
  expected: unknown,
  visited = new WeakSet<object>(),
): boolean {
  if (Object.is(actual, expected)) return true;
  if (typeof expected === "string" && isRecord(actual)) {
    // MathLive normalizes string macro definitions to objects with a `def`.
    if (actual.def === expected) return true;
  }
  if (!isRecord(expected)) return false;
  if (visited.has(expected)) return true;
  visited.add(expected);

  if (Array.isArray(expected)) {
    return (
      Array.isArray(actual) &&
      actual.length === expected.length &&
      expected.every((value, index) =>
        optionValueContains(actual[index], value, visited),
      )
    );
  }
  if (!isRecord(actual) || Array.isArray(actual)) return false;
  return Object.keys(expected).every((key) =>
    optionValueContains(actual[key], expected[key], visited),
  );
}

function optionsWereCommitted(
  access: MathLiveInternalOptionAccess,
  expected: Record<string, unknown>,
) {
  const keys = Object.keys(expected);
  if (keys.length === 0) return true;
  let current: Record<string, unknown> | null | undefined;
  try {
    current = access._getOptions?.(keys) ?? access._mathfield?.options;
  } catch {
    return false;
  }
  if (!current) return false;
  return keys.every((key) =>
    optionValueContains(current?.[key], expected[key]),
  );
}

function isCommittedEmptyModelModeFailure(
  field: MathfieldElement,
  access: MathLiveInternalOptionAccess,
  options: Record<string, unknown>,
  error: unknown,
) {
  if (!isTypeErrorLike(error) || !field.isConnected) return false;
  const root = access._mathfield?.model?.root;
  if (!root || root.firstChild != null) return false;
  try {
    if (field.value !== "") return false;
  } catch {
    return false;
  }

  const detail = errorDetail(error).toLowerCase();
  const identifiesModeWrite =
    detail.includes("mode") &&
    (detail.includes("undefined") || detail.includes("null")) &&
    (detail.includes("setoptions") ||
      detail.includes("_setoptions") ||
      detail.includes("firstchild"));
  return identifiesModeWrite && optionsWereCommitted(access, options);
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

  access._setOptions = (options: Record<string, unknown>) => {
    try {
      original.call(field, options);
    } catch (error) {
      // MathLive 0.109.2 commits updated options before it dereferences the
      // absent first child of an empty model. Suppress only that fully verified
      // failure. All non-empty, uncommitted, or unrelated failures still throw.
      if (!isCommittedEmptyModelModeFailure(field, access, options, error)) {
        throw error;
      }
      try {
        access.reflectAttributes?.call(field);
      } catch (reflectionError) {
        console.warn(
          "VisualTeX could not reflect a committed MathLive option update.",
          reflectionError,
        );
      }
    }
  };
  guardedFields.add(field);
}
