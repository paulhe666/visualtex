import { EngineEditorSessionOwner } from "../../../../engine/web/host/editor-session.ts";
import { SvgSceneView } from "../../../../engine/web/host/svg-view.ts";
import {
  getCoreEngineRuntime,
  resolveCoreLayoutForPreferences,
} from "./coreEngineRuntime";
import type {
  FormulaChineseFont,
  FormulaLetterFont,
} from "../editor/formulaFontPreferences";

export interface CoreStaticPreviewRequest {
  latex: string;
  style?: "display" | "text";
  formulaLetterFont: FormulaLetterFont;
  formulaChineseFont: FormulaChineseFont;
  customSymbolRevision: number;
  symbols: readonly unknown[];
}

export const CORE_PREVIEW_RETRY_EVENT = "visualtex-core-preview-retry";
const cache = new Map<string, Promise<string | null>>();
const CACHE_LIMIT = 512;
let nextPreviewOwner = 1;

function cacheKey(request: CoreStaticPreviewRequest) {
  return [
    request.formulaLetterFont,
    request.formulaChineseFont,
    request.style ?? "text",
    request.customSymbolRevision,
    request.latex,
  ].join("\u0000");
}

function remember(key: string, value: Promise<string | null>) {
  if (cache.size >= CACHE_LIMIT) {
    const oldest = cache.keys().next().value;
    if (typeof oldest === "string") cache.delete(oldest);
  }
  cache.set(key, value);
  void value.then((markup) => {
    // A failed/unsupported preview must not become a permanently blank cache
    // entry after resource recovery. Never evict a newer in-flight replacement.
    if (!markup && cache.get(key) === value) cache.delete(key);
  });
  return value;
}

/**
 * Pure static preview path: no BrowserInputBridge, command state or editor host
 * is created. A short-lived core document owner exists only for one cache miss,
 * exports validated SVG, and is destroyed immediately afterwards.
 */
export function renderCoreStaticPreview(
  request: CoreStaticPreviewRequest,
): Promise<string | null> {
  const latex = request.latex.trim();
  if (!latex || latex.length > 16_384) return Promise.resolve(null);
  const key = cacheKey({ ...request, latex });
  const existing = cache.get(key);
  if (existing) return existing;
  return remember(
    key,
    (async () => {
      const context = await getCoreEngineRuntime();
      const layout = await resolveCoreLayoutForPreferences(
        context,
        request.formulaLetterFont,
        request.formulaChineseFont,
        request.style ?? "text",
      );
      if (!layout) return null;
      const activeBefore = context.runtime.activeSessions;
      const owner = new EngineEditorSessionOwner(
        context.runtime,
        `static-preview:${nextPreviewOwner++}`,
        latex,
        request.symbols,
      );
      try {
        try {
          const result = owner.layout(layout, true);
          if (result.svg) return result.svg;
        } catch {
          // Incomplete-but-owned source such as a toolbar template with
          // \\placeholder{} is still previewable from the core layout snapshot.
        }
        const result = owner.layout(layout, false);
        if (!result.snapshot.documentId) return null;
        const detached = document.createElement("span");
        const view = new SvgSceneView(detached, { accessibleLabel: "" });
        try {
          view.setSnapshot(result.snapshot);
          for (const region of result.snapshot.regions) {
            if (region.role !== "emptySlot" && region.role !== "emptyRow") continue;
            const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
            const width = Math.max(0.22, region.logical.xMax - region.logical.xMin);
            const height = Math.max(0.5, region.logical.yMax - region.logical.yMin);
            rect.setAttribute("x", String(region.logical.xMin));
            rect.setAttribute("y", String(region.logical.yMin));
            rect.setAttribute("width", String(width));
            rect.setAttribute("height", String(height));
            rect.setAttribute("rx", "0.06");
            rect.setAttribute("fill", "none");
            rect.setAttribute("stroke", "currentColor");
            rect.setAttribute("stroke-width", "0.055");
            rect.setAttribute("stroke-dasharray", "0.11 0.08");
            rect.setAttribute("opacity", "0.72");
            rect.setAttribute("data-core-placeholder", "true");
            view.svg.append(rect);
          }
          return view.svg.outerHTML;
        } finally {
          view.dispose();
        }
      } catch {
        return null;
      } finally {
        owner.destroy();
        if (context.runtime.activeSessions !== activeBefore) {
          throw new Error("Static preview leaked a core editor session.");
        }
      }
    })().catch(() => null),
  );
}

export function clearCoreStaticPreviewCache() {
  cache.clear();
}

export function retryCoreStaticPreviews() {
  clearCoreStaticPreviewCache();
  if (typeof window !== "undefined") window.dispatchEvent(new Event(CORE_PREVIEW_RETRY_EVENT));
}
