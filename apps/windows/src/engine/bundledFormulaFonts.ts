import manifest from "../../src-tauri/resources/formula-fonts/manifest.json";
import type { FontIdentity, LayoutOptions } from "../../../../engine/web/host/layout.ts";

/** Only immutable app assets are included. No operating-system fonts or oracle
 * runtime modules are imported. Vite emits the same bytes for desktop and Office. */
const urls = import.meta.glob<string>(
  "../../src-tauri/resources/formula-fonts/*.{otf,ttf}",
  { eager: true, query: "?url", import: "default" },
);
export interface FontResourceRuntime {
  registerFont(bytes: Uint8Array, faceIndex: number): string;
}
export type BundledFontRecord = (typeof manifest.fonts)[number];
const records = new Map(manifest.fonts.map((record) => [record.id, record]));
const pendingByRuntime = new WeakMap<FontResourceRuntime, Map<string, Promise<FontIdentity>>>();
export const BUNDLED_FORMULA_FONT_CATALOG = manifest;

export async function loadBundledFormulaFont(runtime: FontResourceRuntime, id: string): Promise<FontIdentity> {
  let pending = pendingByRuntime.get(runtime);
  if (!pending) { pending = new Map(); pendingByRuntime.set(runtime, pending); }
  const existing = pending.get(id);
  if (existing) return existing;
  const task = (async () => {
    const record = records.get(id);
    if (!record) throw new Error(`Missing bundled formula font: ${id}`);
    const url = urls[`../../src-tauri/resources/formula-fonts/${record.file}`];
    if (!url) throw new Error(`Formula font was omitted from this application build: ${record.file}`);
    const response = await fetch(url, { cache: "force-cache" });
    if (!response.ok) throw new Error(`Formula font could not be loaded: ${record.file} (${response.status})`);
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength !== record.bytes) throw new Error(`Formula font length mismatch: ${record.file}`);
    const digest = await crypto.subtle.digest("SHA-256", buffer);
    const actual = Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
    if (actual !== record.sha256) throw new Error(`Formula font integrity mismatch: ${record.file}`);
    const result = JSON.parse(runtime.registerFont(new Uint8Array(buffer), record.faceIndex));
    if (result.status !== "ok" || result.result?.kind !== "fontRegistered") {
      throw new Error(`Formula font registration failed (${record.file}): ${result.error?.message ?? "invalid response"}`);
    }
    const font = result.result.font as FontIdentity;
    if (font.sha256 !== actual || font.faceIndex !== record.faceIndex) {
      throw new Error(`Formula font identity mismatch: ${record.file}`);
    }
    return font;
  })();
  pending.set(id, task);
  // A failed load is retryable; it must not poison every palette/editor forever.
  void task.catch(() => { if (pending?.get(id) === task) pending.delete(id); });
  return task;
}

const TEX_ROLES = {
  main: "KaTeX_Main-Regular", mainItalic: "KaTeX_Main-Italic",
  mainBold: "KaTeX_Main-Bold", mainBoldItalic: "KaTeX_Main-BoldItalic",
  mathItalic: "KaTeX_Math-Italic", mathBoldItalic: "KaTeX_Math-BoldItalic",
  ams: "KaTeX_AMS-Regular", calligraphic: "KaTeX_Caligraphic-Regular",
  calligraphicBold: "KaTeX_Caligraphic-Bold", script: "KaTeX_Script-Regular",
  fraktur: "KaTeX_Fraktur-Regular", frakturBold: "KaTeX_Fraktur-Bold",
  sans: "KaTeX_SansSerif-Regular", sansBold: "KaTeX_SansSerif-Bold",
  sansItalic: "KaTeX_SansSerif-Italic", monospace: "KaTeX_Typewriter-Regular",
  size1: "KaTeX_Size1-Regular", size2: "KaTeX_Size2-Regular",
  size3: "KaTeX_Size3-Regular", size4: "KaTeX_Size4-Regular",
} as const;

export async function loadBundledTexFontSet(runtime: FontResourceRuntime): Promise<NonNullable<LayoutOptions["texFontSet"]>> {
  return Object.fromEntries(await Promise.all(Object.entries(TEX_ROLES).map(async ([role, id]) =>
    [role, (await loadBundledFormulaFont(runtime, id)).id],
  )));
}

export async function loadBundledTextFamily(runtime: FontResourceRuntime, family: "termes" | "pagella" | "heros") {
  const [regular, italic, bold, boldItalic] = await Promise.all(
    ["regular", "italic", "bold", "bolditalic"].map((face) => loadBundledFormulaFont(runtime, `texgyre${family}-${face}`)),
  );
  return { regular, italic, bold, boldItalic };
}
