import { invoke, isTauri } from "@tauri-apps/api/core";
import init, { EngineRuntime } from "../../../../engine/web/pkg/visualtex_engine_wasm.js";
import wasmUrl from "../../../../engine/web/pkg/visualtex_engine_wasm_bg.wasm?url";
import mathFontUrl from "../../src-windows/artifacts/office-fonts/latinmodern-math.otf?url";
import {
  EngineEditorSessionOwner,
  type EngineRuntimePort,
} from "../../../../engine/web/host/editor-session.ts";
import type { FontIdentity, LayoutOptions } from "../../../../engine/web/host/layout.ts";
import type { WorkspaceMode } from "../workspace/workspaceTypes";
import {
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "../editor/formulaFontPreferences";
import { loadBundledFormulaFont, loadBundledTexFontSet, loadBundledTextFamily } from "./bundledFormulaFonts";

const EXPECTED_CORE_WASM_SHA256 =
  "0497fb94613855117479d92761a6692771ff017906ebde61c06cbdd99ec72865";
const EXPECTED_MATH_FONT_SHA256 =
  "6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f";

export interface CoreEngineRuntimePort extends EngineRuntimePort {
  registerFont(bytes: Uint8Array, faceIndex: number): string;
  readonly activeSessions: number;
  free(): void;
}

export interface CoreEngineRuntimeContext {
  generation: number;
  runtime: CoreEngineRuntimePort;
  mathFont: FontIdentity;
  baseLayout: LayoutOptions;
}

export interface LockedCoreFormulaSession {
  owner: EngineEditorSessionOwner;
  layout: LayoutOptions;
}

export interface HeadlessCoreFormulaOwner {
  context: CoreEngineRuntimeContext;
  owner: EngineEditorSessionOwner;
}

let runtimePromise: Promise<CoreEngineRuntimeContext> | null = null;
let runtimeGeneration = 0;
let activeRuntimeGeneration = 0;
let pageHideInstalled = false;
const systemFontPromises = new Map<string, Promise<FontIdentity | null>>();

export class CoreRuntimeFaultError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "CoreRuntimeFaultError";
  }
}

export function isCoreRuntimeFault(error: unknown): error is CoreRuntimeFaultError {
  return error instanceof CoreRuntimeFaultError;
}

class FaultAwareCoreRuntime implements CoreEngineRuntimePort {
  private faulted = false;
  private freed = false;

  constructor(
    private readonly inner: EngineRuntime,
    private readonly generation: number,
  ) {}

  private fail(operation: string, reason: unknown): never {
    this.faulted = true;
    if (activeRuntimeGeneration === this.generation) {
      runtimePromise = null;
      systemFontPromises.clear();
    }
    try {
      if (!this.freed) this.inner.free();
    } catch {
      // A trapped instance may reject teardown. The next owner still receives a
      // fresh EngineRuntime object because the global context was invalidated.
    }
    this.freed = true;
    const detail = reason instanceof Error ? reason.message : String(reason);
    throw new CoreRuntimeFaultError(
      `VisualTeX core runtime failed during ${operation}; the last committed formula state was preserved and a fresh runtime will be created on retry. ${detail}`,
      reason instanceof Error ? { cause: reason } : undefined,
    );
  }

  private available() {
    if (this.faulted || this.freed) {
      throw new CoreRuntimeFaultError(
        "VisualTeX core runtime instance is no longer available; retry to rebuild it from the last committed formula state.",
      );
    }
  }

  dispatch(request: string) {
    this.available();
    try {
      return this.inner.dispatch(request);
    } catch (error) {
      return this.fail("dispatch", error);
    }
  }

  registerFont(bytes: Uint8Array, faceIndex: number) {
    this.available();
    try {
      return this.inner.registerFont(bytes, faceIndex);
    } catch (error) {
      return this.fail("font registration", error);
    }
  }

  get activeSessions() {
    this.available();
    try {
      return this.inner.activeSessions;
    } catch (error) {
      return this.fail("session inspection", error);
    }
  }

  free() {
    if (this.freed) return;
    this.freed = true;
    if (activeRuntimeGeneration === this.generation) {
      runtimePromise = null;
      systemFontPromises.clear();
    }
    this.inner.free();
  }
}

interface NativeSystemFontAsset {
  requestedFamily: string;
  resolvedFamily: string;
  faceIndex: number;
  sha256: string;
  byteLength: number;
  bytesBase64: string;
  hasMathTable: boolean;
  italic: boolean;
  weight: number;
}

function bytesToHex(bytes: ArrayBuffer) {
  return Array.from(new Uint8Array(bytes), (value) =>
    value.toString(16).padStart(2, "0"),
  ).join("");
}

async function readVerifiedAsset(
  url: string,
  expectedSha256: string,
  label: "WASM" | "font",
) {
  const response = await fetch(url, { cache: "force-cache" });
  if (!response.ok) {
    throw new Error(`VisualTeX core ${label} resource failed to load (${response.status}).`);
  }
  const bytes = await response.arrayBuffer();
  const actual = bytesToHex(await crypto.subtle.digest("SHA-256", bytes));
  if (actual !== expectedSha256) {
    throw new Error(`VisualTeX core ${label} resource hash mismatch.`);
  }
  return new Uint8Array(bytes);
}

function registerFont(
  runtime: CoreEngineRuntimePort,
  bytes: Uint8Array,
  faceIndex: number,
  expectedSha256: string,
) {
  const response = JSON.parse(runtime.registerFont(bytes, faceIndex)) as {
    status: "ok" | "error";
    result?: { kind: string; font: FontIdentity };
    error?: { code: string; message: string };
  };
  if (response.status !== "ok" || response.result?.kind !== "fontRegistered") {
    throw new Error(
      response.error
        ? `${response.error.code}: ${response.error.message}`
        : "VisualTeX core rejected its verified font resource.",
    );
  }
  if (
    response.result.font.sha256 !== expectedSha256 ||
    response.result.font.faceIndex !== faceIndex
  ) {
    throw new Error("VisualTeX core registered an unexpected font identity.");
  }
  return response.result.font;
}

async function createRuntime(): Promise<CoreEngineRuntimeContext> {
  const wasmBytes = await readVerifiedAsset(
    wasmUrl,
    EXPECTED_CORE_WASM_SHA256,
    "WASM",
  );
  await init({ module_or_path: wasmBytes });
  const generation = ++runtimeGeneration;
  activeRuntimeGeneration = generation;
  const runtime = new FaultAwareCoreRuntime(new EngineRuntime(), generation);
  try {
    const bytes = await readVerifiedAsset(
      mathFontUrl,
      EXPECTED_MATH_FONT_SHA256,
      "font",
    );
    const mathFont = registerFont(runtime, bytes, 0, EXPECTED_MATH_FONT_SHA256);
    const baseLayout: LayoutOptions = {
      mathFont: mathFont.id,
      // Text is resolved from the user's actual verified Windows font preference
      // when running inside Tauri. A plain browser candidate keeps this null.
      textFont: null,
      mathFallbacks: [],
      customSymbolTextFonts: [],
      style: "display",
      fontSizePx: 28,
      paddingEm: 0.125,
    };
    if (!pageHideInstalled && typeof window !== "undefined") {
      pageHideInstalled = true;
      window.addEventListener(
        "pagehide",
        () => {
          try {
            runtime.free();
          } catch {
            // The page is already being destroyed.
          }
          runtimePromise = null;
          systemFontPromises.clear();
        },
        { once: true },
      );
    }
    return { generation, runtime, mathFont, baseLayout };
  } catch (error) {
    runtime.free();
    throw error;
  }
}

export function coreEditorExperimentEnabled(mode: WorkspaceMode) {
  return (
    mode === "desktop" &&
    import.meta.env.VITE_VISUALTEX_CORE_EDITOR === "1"
  );
}

function decodeSystemFontAsset(value: unknown): NativeSystemFontAsset {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("VisualTeX returned an invalid system font asset.");
  }
  const asset = value as Partial<NativeSystemFontAsset>;
  if (
    typeof asset.requestedFamily !== "string" || !asset.requestedFamily || asset.requestedFamily.length > 128 ||
    typeof asset.resolvedFamily !== "string" || !asset.resolvedFamily || asset.resolvedFamily.length > 128 ||
    !Number.isSafeInteger(asset.faceIndex) || Number(asset.faceIndex) < 0 || Number(asset.faceIndex) > 65535 ||
    typeof asset.sha256 !== "string" || !/^[0-9a-f]{64}$/u.test(asset.sha256) ||
    !Number.isSafeInteger(asset.byteLength) || Number(asset.byteLength) <= 0 || Number(asset.byteLength) > 128 * 1024 * 1024 ||
    typeof asset.bytesBase64 !== "string" || asset.bytesBase64.length > 180 * 1024 * 1024 ||
    typeof asset.hasMathTable !== "boolean" ||
    typeof asset.italic !== "boolean" ||
    !Number.isSafeInteger(asset.weight) || Number(asset.weight) < 1 || Number(asset.weight) > 1000
  ) {
    throw new Error("VisualTeX returned malformed system font metadata.");
  }
  return asset as NativeSystemFontAsset;
}

function decodeBase64Bytes(value: string) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

async function resolveSystemFont(
  context: CoreEngineRuntimeContext,
  families: readonly string[],
  requireMathTable = false,
): Promise<FontIdentity | null> {
  const normalized = Array.from(new Set(families.map((family) => family.trim()).filter(Boolean))).slice(0, 16);
  if (!normalized.length || !isTauri()) return null;
  const key = `${context.generation}:${requireMathTable ? "math" : "text"}:${normalized
    .map((family) => family.toLocaleLowerCase())
    .join("\u001f")}`;
  const existing = systemFontPromises.get(key);
  if (existing) return existing;
  const pending = (async () => {
    const candidates = requireMathTable ? normalized.map((family) => [family]) : [normalized];
    let lastError: unknown = null;
    for (const candidate of candidates) {
      try {
        const asset = decodeSystemFontAsset(
          await invoke<unknown>("read_system_font_asset", { fontFamilies: candidate }),
        );
        if (requireMathTable && !asset.hasMathTable) continue;
        const bytes = decodeBase64Bytes(asset.bytesBase64);
        if (bytes.byteLength !== asset.byteLength) {
          throw new Error("VisualTeX system font byte length mismatch.");
        }
        const sha256 = bytesToHex(await crypto.subtle.digest("SHA-256", bytes));
        if (sha256 !== asset.sha256) {
          throw new Error("VisualTeX system font hash mismatch.");
        }
        return registerFont(context.runtime, bytes, asset.faceIndex, asset.sha256);
      } catch (error) {
        if (isCoreRuntimeFault(error)) throw error;
        lastError = error;
      }
    }
    if (lastError) throw lastError;
    return null;
  })().catch((error) => {
    systemFontPromises.delete(key);
    if (isCoreRuntimeFault(error)) throw error;
    console.warn("VisualTeX core could not resolve a verified system font:", error);
    return null;
  });
  systemFontPromises.set(key, pending);
  return pending;
}

async function resolveSystemFontVariant(
  context: CoreEngineRuntimeContext,
  families: readonly string[],
  italic: boolean,
  weight = 400,
): Promise<FontIdentity | null> {
  const normalized = Array.from(new Set(families.map((family) => family.trim()).filter(Boolean))).slice(0, 16);
  if (!normalized.length || !isTauri()) return null;
  const key = `${context.generation}:variant:${italic ? "italic" : "upright"}:${weight}:${normalized
    .map((family) => family.toLocaleLowerCase())
    .join("\u001f")}`;
  const existing = systemFontPromises.get(key);
  if (existing) return existing;
  const pending = (async () => {
    const asset = decodeSystemFontAsset(
      await invoke<unknown>("read_system_font_variant_asset", {
        fontFamilies: normalized,
        italic,
        weight,
      }),
    );
    if (asset.italic !== italic || ((asset.weight >= 600) !== (weight >= 600))) {
      throw new Error("VisualTeX system font variant style/weight mismatch.");
    }
    const bytes = decodeBase64Bytes(asset.bytesBase64);
    if (bytes.byteLength !== asset.byteLength) {
      throw new Error("VisualTeX system font variant byte length mismatch.");
    }
    const sha256 = bytesToHex(await crypto.subtle.digest("SHA-256", bytes));
    if (sha256 !== asset.sha256) {
      throw new Error("VisualTeX system font variant hash mismatch.");
    }
    return registerFont(context.runtime, bytes, asset.faceIndex, asset.sha256);
  })().catch((error) => {
    systemFontPromises.delete(key);
    if (isCoreRuntimeFault(error)) throw error;
    console.warn("VisualTeX core could not resolve a verified system font variant:", error);
    return null;
  });
  systemFontPromises.set(key, pending);
  return pending;
}

export interface CoreFontResolution {
  layout: LayoutOptions;
  requestedLetterFont: FormulaLetterFont;
  requestedChineseFont: FormulaChineseFont;
  actualLetterFamily: string;
  actualChineseFamily: string;
  letterSource: "bundled" | "system";
  chineseSource: "bundled" | "system";
  exactLetterFamily: boolean;
  notices: string[];
}

const fontResolutions = new WeakMap<CoreEngineRuntimePort, Map<string, Promise<CoreFontResolution>>>();

async function resolveRealTextFamily(context: CoreEngineRuntimeContext, families: readonly string[]) {
  if (!isTauri()) return null;
  for (const family of families) {
    // Resolve all four faces from ONE family. Never combine regular Times with
    // italic Arial merely because a CSS-style fallback list matched differently.
    const [regular, italic, bold, boldItalic] = await Promise.all([
      resolveSystemFontVariant(context, [family], false, 400),
      resolveSystemFontVariant(context, [family], true, 400),
      resolveSystemFontVariant(context, [family], false, 700),
      resolveSystemFontVariant(context, [family], true, 700),
    ]);
    if (regular && italic && bold && boldItalic) return { family, regular, italic, bold, boldItalic };
  }
  return null;
}

/** One font plan for editor, candidate, palette and graphics export. Optional
 * system fonts can improve an explicit selection; they cannot gate core startup.
 * Every offline fallback is a named real bundled face, never an invented family. */
export async function resolveCoreFontPreferences(
  context: CoreEngineRuntimeContext,
  letterFont: FormulaLetterFont,
  chineseFont: FormulaChineseFont,
  style: "display" | "text",
): Promise<CoreFontResolution> {
  let cache = fontResolutions.get(context.runtime);
  if (!cache) { cache = new Map(); fontResolutions.set(context.runtime, cache); }
  const key = `${letterFont}:${chineseFont}:${style}`;
  const existing = cache.get(key);
  if (existing) return structuredClone(await existing);
  const task = (async (): Promise<CoreFontResolution> => {
    const layout: LayoutOptions = { ...context.baseLayout, style, texFontSet: {} };
    const notices: string[] = [];
    let actualLetterFamily = "KaTeX";
    let letterSource: "bundled" | "system" = "bundled";
    let exactLetterFamily = true;
    const stix = () => loadBundledFormulaFont(context.runtime, "STIXTwoMath-Regular");
    if (letterFont === "katex") {
      const tex = await loadBundledTexFontSet(context.runtime);
      layout.texFontSet = tex;
      layout.alphabetUprightFont = tex.main;
      layout.alphabetItalicFont = tex.mathItalic;
      layout.alphabetBoldFont = tex.mainBold;
      layout.alphabetBoldItalicFont = tex.mathBoldItalic;
      // Rare characters absent from the 20 KaTeX faces have an explicit,
      // verified bundled mathematical supplement; no operating-system fallback.
      layout.mathFallbacks = [(await stix()).id];
    } else if (letterFont === "stix" || letterFont === "cambria") {
      const native = letterFont === "cambria" ? await resolveSystemFont(context, ["Cambria Math"], true) : null;
      const math = native ?? await stix();
      layout.mathFont = math.id;
      layout.mathFallbacks = [context.mathFont.id];
      actualLetterFamily = native ? "Cambria Math" : "STIX Two Math";
      letterSource = native ? "system" : "bundled";
      exactLetterFamily = letterFont === "stix" || Boolean(native);
      if (!exactLetterFamily) notices.push("Cambria Math 不可用，当前使用随软件提供的 STIX Two Math；未改写保存的字体偏好。");
    } else {
      const profiles = {
        times: { families: ["Times New Roman"], bundled: "termes", name: "TeX Gyre Termes" },
        palatino: { families: ["Palatino Linotype", "Palatino"], bundled: "pagella", name: "TeX Gyre Pagella" },
        helvetica: { families: ["Arial", "Helvetica Neue", "Helvetica"], bundled: "heros", name: "TeX Gyre Heros" },
      } as const;
      const profile = profiles[letterFont];
      const native = await resolveRealTextFamily(context, profile.families);
      const faces = native ?? await loadBundledTextFamily(context.runtime, profile.bundled);
      layout.alphabetUprightFont = faces.regular.id;
      layout.alphabetItalicFont = faces.italic.id;
      layout.alphabetBoldFont = faces.bold.id;
      layout.alphabetBoldItalicFont = faces.boldItalic.id;
      layout.mathFallbacks = [(await stix()).id];
      actualLetterFamily = native?.family ?? profile.name;
      letterSource = native ? "system" : "bundled";
      exactLetterFamily = Boolean(native);
      if (!native) notices.push(`${profile.families[0]} 不可用，当前使用随软件提供的 ${profile.name}；两者是不同字体。`);
    }
    const chineseProfiles = {
      system: { system: null, bundled: "NotoSansCJKsc-Regular", name: "Noto Sans CJK SC" },
      pingfang: { system: "PingFang SC", bundled: "NotoSansCJKsc-Regular", name: "Noto Sans CJK SC" },
      songti: { system: "SimSun", bundled: "NotoSerifCJKsc-Regular", name: "Noto Serif CJK SC" },
      kaiti: { system: "KaiTi", bundled: "LXGWWenKai-Regular", name: "LXGW WenKai" },
      heiti: { system: "SimHei", bundled: "NotoSansCJKsc-Regular", name: "Noto Sans CJK SC" },
    } as const;
    const chinese = chineseProfiles[chineseFont];
    const nativeChinese = chinese.system ? await resolveSystemFontVariant(context, [chinese.system], false, 400) : null;
    const text = nativeChinese ?? await loadBundledFormulaFont(context.runtime, chinese.bundled);
    layout.textFont = text.id;
    if (chinese.system && !nativeChinese) notices.push(`${chinese.system} 不可用，中文使用随软件提供的 ${chinese.name}。`);
    return { layout, requestedLetterFont: letterFont, requestedChineseFont: chineseFont,
      actualLetterFamily, actualChineseFamily: nativeChinese ? chinese.system! : chinese.name,
      letterSource, chineseSource: nativeChinese ? "system" : "bundled", exactLetterFamily, notices };
  })();
  cache.set(key, task);
  void task.catch(() => { if (cache?.get(key) === task) cache.delete(key); });
  return structuredClone(await task);
}

export async function resolveCoreLayoutForPreferences(
  context: CoreEngineRuntimeContext,
  letterFont: FormulaLetterFont,
  chineseFont: FormulaChineseFont,
  style: "display" | "text",
): Promise<LayoutOptions | null> {
  try {
    return (await resolveCoreFontPreferences(context, letterFont, chineseFont, style)).layout;
  } catch (error) {
    if (isCoreRuntimeFault(error)) throw error;
    // Font-resource resolution is never allowed to decide whether the editor
    // exists. Keep the self-owned core alive with its verified bundled MATH
    // face, and add the bundled CJK text face when it is available. The exact
    // requested plan remains retryable because failed font promises are evicted.
    console.warn("VisualTeX formula font plan fell back to the bundled safety layout:", error);
    let textFont: FontIdentity | null = null;
    try {
      textFont = await loadBundledFormulaFont(context.runtime, "NotoSansCJKsc-Regular");
    } catch (textError) {
      console.warn("VisualTeX bundled CJK safety font is unavailable:", textError);
    }
    return {
      ...context.baseLayout,
      textFont: textFont?.id ?? null,
      texFontSet: {},
      alphabetUprightFont: null,
      alphabetItalicFont: null,
      alphabetBoldFont: null,
      alphabetBoldItalicFont: null,
      mathFallbacks: [],
      style,
    };
  }
}

export function getCoreEngineRuntime() {
  runtimePromise ??= createRuntime().catch((error) => {
    runtimePromise = null;
    throw error;
  });
  return runtimePromise;
}

/**
 * Create a document-scoped core owner without mounting any editor view or
 * requesting layout. Hidden Office conversion uses this entry point so the
 * Word OMML fast path can export Presentation MathML without paying for SVG,
 * hit testing, an input host, or IME/candidate UI.
 */
export async function createHeadlessCoreFormulaOwner(
  documentId: string,
  source: string,
  symbols: readonly unknown[] = [],
): Promise<HeadlessCoreFormulaOwner> {
  const context = await getCoreEngineRuntime();
  return {
    context,
    owner: new EngineEditorSessionOwner(
      context.runtime,
      documentId,
      source,
      symbols,
    ),
  };
}

/**
 * The engine choice is made once for the complete formula owner. A failed
 * create/layout attempt is destroyed and the caller must keep that formula on
 * the legacy editor for the lifetime of the mounted session. There is no
 * subtree or mid-edit fallback.
 */
export async function tryCreateLockedCoreFormulaSession(
  documentId: string,
  source: string,
  style: "display" | "text",
  symbols: readonly unknown[] = [],
  layoutOverride?: LayoutOptions | null,
): Promise<LockedCoreFormulaSession | null> {
  const { context, owner } = await createHeadlessCoreFormulaOwner(
    documentId,
    source,
    symbols,
  );
  const layout: LayoutOptions = layoutOverride ?? {
    ...context.baseLayout,
    style,
  };
  try {
    if (!source.trim()) return { owner, layout };
    owner.layout(layout, false);
    return { owner, layout };
  } catch {
    owner.destroy();
    return null;
  }
}
