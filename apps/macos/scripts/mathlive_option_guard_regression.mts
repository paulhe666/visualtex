import assert from "node:assert/strict";
import type { MathfieldElement } from "mathlive";
import {
  installMathLiveOptionMutationGuard,
  readMathLiveOptionBeforeMount,
  setMathLiveOptionsBeforeMount,
} from "../src/editor/mathLiveOptionCompatibility.ts";

type FakeOptions = Record<string, unknown>;

type FakeMathfield = {
  isConnected: boolean;
  value: string;
  calls: number;
  reflections: number;
  options: FakeOptions;
  _mathfield: {
    model: { root: { firstChild?: unknown } };
    options: FakeOptions;
  };
  _getOptions: (keys?: string | string[]) => FakeOptions;
  _setOptions: (options: FakeOptions) => void;
  reflectAttributes: () => void;
};

function normalizeFakeOption(value: unknown): unknown {
  if (typeof value === "string") return { def: value };
  if (Array.isArray(value)) return value.map(normalizeFakeOption);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [
        key,
        key === "macros" && entry && typeof entry === "object"
          ? Object.fromEntries(
              Object.entries(entry).map(([macro, definition]) => [
                macro,
                normalizeFakeOption(definition),
              ]),
            )
          : entry,
      ]),
    );
  }
  return value;
}

function modeWriteError(kind: "v8" | "safari") {
  const error = new TypeError(
    kind === "v8"
      ? "Cannot set properties of undefined (setting 'mode')"
      : "undefined is not an object (evaluating 'this.model.root.firstChild.mode = mode')",
  );
  error.stack = `${error.name}: ${error.message}\n    at kd.setOptions (mathlive.js:1:1)\n    at a6._setOptions (mathlive.js:1:2)`;
  return error;
}

function createFakeMathfield({
  connected = true,
  value = "",
  firstChild,
  commit = true,
  errorFactory = () => modeWriteError("v8"),
}: {
  connected?: boolean;
  value?: string;
  firstChild?: unknown;
  commit?: boolean;
  errorFactory?: () => Error;
} = {}): FakeMathfield {
  const fake: FakeMathfield = {
    isConnected: connected,
    value,
    calls: 0,
    reflections: 0,
    options: {},
    _mathfield: {
      model: { root: { firstChild } },
      options: {},
    },
    _getOptions(keys) {
      if (keys === undefined) return { ...this.options };
      const requested = typeof keys === "string" ? [keys] : keys;
      return Object.fromEntries(
        requested.map((key) => [key, this.options[key]]),
      );
    },
    _setOptions(options) {
      this.calls += 1;
      if (commit) {
        const normalized = Object.fromEntries(
          Object.entries(options).map(([key, entry]) => [
            key,
            key === "macros" && entry && typeof entry === "object"
              ? Object.fromEntries(
                  Object.entries(entry).map(([macro, definition]) => [
                    macro,
                    normalizeFakeOption(definition),
                  ]),
                )
              : entry,
          ]),
        );
        this.options = { ...this.options, ...normalized };
        this._mathfield.options = this.options;
      }
      throw errorFactory();
    },
    reflectAttributes() {
      this.reflections += 1;
    },
  };
  return fake;
}

function asMathfield(fake: FakeMathfield) {
  return fake as unknown as MathfieldElement;
}

for (const kind of ["v8", "safari"] as const) {
  const fake = createFakeMathfield({ errorFactory: () => modeWriteError(kind) });
  const field = asMathfield(fake);
  installMathLiveOptionMutationGuard(field);
  installMathLiveOptionMutationGuard(field);

  assert.doesNotThrow(() => {
    fake._setOptions({
      macros: { visualTexProbe: "\\mathrm{VT}" },
      smartMode: false,
    });
  });
  assert.equal(fake.calls, 1, `${kind}: guard is idempotent`);
  assert.equal(fake.reflections, 1, `${kind}: attributes are reflected once`);
  assert.deepEqual(fake.options.macros, {
    visualTexProbe: { def: "\\mathrm{VT}" },
  });
  assert.equal(fake.options.smartMode, false);
}

{
  const fake = createFakeMathfield({ value: "x" });
  installMathLiveOptionMutationGuard(asMathfield(fake));
  assert.throws(
    () => fake._setOptions({ macros: { probe: "x" } }),
    /setting 'mode'/,
    "a non-empty model failure must remain visible",
  );
}

{
  const fake = createFakeMathfield({ firstChild: { mode: "math" } });
  installMathLiveOptionMutationGuard(asMathfield(fake));
  assert.throws(
    () => fake._setOptions({ macros: { probe: "x" } }),
    /setting 'mode'/,
    "a failure with a live root child must remain visible",
  );
}

{
  const fake = createFakeMathfield({ commit: false });
  installMathLiveOptionMutationGuard(asMathfield(fake));
  assert.throws(
    () => fake._setOptions({ macros: { probe: "x" } }),
    /setting 'mode'/,
    "an uncommitted option mutation must remain visible",
  );
}

{
  const fake = createFakeMathfield({
    errorFactory: () => new TypeError("unrelated option normalization failure"),
  });
  installMathLiveOptionMutationGuard(asMathfield(fake));
  assert.throws(
    () => fake._setOptions({ macros: { probe: "x" } }),
    /unrelated option normalization failure/,
    "an unrelated TypeError must remain visible",
  );
}

{
  const fake = createFakeMathfield({ connected: false });
  fake._setOptions = function setDeferredOptions(options) {
    this.calls += 1;
    this.options = { ...this.options, ...options };
    this._mathfield.options = this.options;
  };
  const field = asMathfield(fake);
  setMathLiveOptionsBeforeMount(field, { smartFence: false });
  assert.equal(readMathLiveOptionBeforeMount<boolean>(field, "smartFence"), false);
  fake.isConnected = true;
  assert.throws(
    () => setMathLiveOptionsBeforeMount(field, { smartFence: true }),
    /before mount/,
  );
}

console.log("VisualTeX MathLive option mutation guard regression passed");
