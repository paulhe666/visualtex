import type {
  CustomSymbolDefinition,
  CustomSymbolDesignerSourceArchive,
  CustomSymbolDesignerSourceAsset,
  CustomSymbolDesignerSourceLayer,
} from "./customSymbolTypes";
import type {
  CustomSymbolDesignerDocument,
  CustomSymbolDesignerGeometryLayer,
  CustomSymbolDesignerGlyphLayer,
  CustomSymbolDesignerLayer,
} from "./customSymbolDesignerTypes";

function deepClone<T>(value: T): T {
  if (typeof structuredClone === "function") return structuredClone(value);
  return JSON.parse(JSON.stringify(value)) as T;
}

function assetKey(layer: CustomSymbolDesignerGlyphLayer) {
  return JSON.stringify({
    sourceLatex: layer.asset.sourceLatex,
    metrics: layer.asset.metrics,
    shapes: layer.asset.shapes,
  });
}

/**
 * Compact the editable designer document into a source archive. Identical glyph
 * assets are stored once, so three non-destructive slices of the same \int do
 * not duplicate the original MathJax path three times in localStorage.
 */
export function createCustomSymbolDesignerSourceArchive(
  document: CustomSymbolDesignerDocument,
): CustomSymbolDesignerSourceArchive {
  const assets: CustomSymbolDesignerSourceAsset[] = [];
  const assetIdByKey = new Map<string, string>();
  const layers: CustomSymbolDesignerSourceLayer[] = [];

  for (const layer of document.layers) {
    const base = {
      id: layer.id,
      name: layer.name,
      visible: layer.visible,
      locked: layer.locked,
      transform: deepClone(layer.transform),
      ...(layer.effects ? { effects: deepClone(layer.effects) } : {}),
      ...(layer.clipRect ? { clipRect: deepClone(layer.clipRect) } : {}),
    };
    if (layer.kind === "glyph") {
      const key = assetKey(layer);
      let assetId = assetIdByKey.get(key);
      if (!assetId) {
        assetId = `asset-${assets.length + 1}`;
        assetIdByKey.set(key, assetId);
        assets.push({
          id: assetId,
          sourceLatex: layer.asset.sourceLatex,
          metrics: deepClone(layer.asset.metrics),
          shapes: deepClone(layer.asset.shapes),
        });
      }
      layers.push({ ...base, kind: "glyph", assetId });
      continue;
    }
    layers.push({
      ...base,
      kind: "geometry",
      ...(layer.geometryPreset ? { geometryPreset: layer.geometryPreset } : {}),
      shape: deepClone(layer.shape),
      bounds: deepClone(layer.bounds),
    });
  }
  return { version: 1, metrics: deepClone(document.metrics), assets, layers };
}

export function restoreCustomSymbolDesignerDocument(
  symbol: CustomSymbolDefinition,
): {
  document: CustomSymbolDesignerDocument;
  sourceMode: "editable" | "flattened-legacy";
} {
  const archive = symbol.designerSource;
  if (archive?.version === 1) {
    const assets = new Map(archive.assets.map((asset) => [asset.id, asset]));
    const layers: CustomSymbolDesignerLayer[] = [];
    for (const layer of archive.layers) {
      const base = {
        id: layer.id,
        name: layer.name,
        visible: layer.visible,
        locked: layer.locked,
        transform: deepClone(layer.transform),
        effects: layer.effects ? deepClone(layer.effects) : undefined,
        clipRect: layer.clipRect ? deepClone(layer.clipRect) : null,
      };
      if (layer.kind === "glyph") {
        const asset = assets.get(layer.assetId);
        if (!asset) continue;
        layers.push({
          ...base,
          kind: "glyph",
          asset: {
            sourceLatex: asset.sourceLatex,
            metrics: deepClone(asset.metrics),
            shapes: deepClone(asset.shapes),
          },
        } satisfies CustomSymbolDesignerGlyphLayer);
        continue;
      }
      layers.push({
        ...base,
        kind: "geometry",
        ...(layer.geometryPreset ? { geometryPreset: layer.geometryPreset } : {}),
        shape: deepClone(layer.shape),
        bounds: deepClone(layer.bounds),
      } satisfies CustomSymbolDesignerGeometryLayer);
    }
    return {
      document: {
        version: 1,
        symbolId: symbol.id,
        name: symbol.name,
        command: symbol.command,
        role: symbol.role,
        limitsBehavior: symbol.limitsBehavior,
        metrics: deepClone(archive.metrics ?? symbol.metrics),
        ommlFallback: symbol.ommlFallback ?? null,
        layers,
      },
      sourceMode: "editable",
    };
  }

  const canvasWidth = symbol.metrics.widthEm * 1000;
  const canvasHeight =
    (symbol.metrics.ascentEm + symbol.metrics.descentEm) * 1000;
  return {
    document: {
      version: 1,
      symbolId: symbol.id,
      name: symbol.name,
      command: symbol.command,
      role: symbol.role,
      limitsBehavior: symbol.limitsBehavior,
      metrics: deepClone(symbol.metrics),
      ommlFallback: symbol.ommlFallback ?? null,
      layers: symbol.artwork.shapes.map(
        (shape, index): CustomSymbolDesignerGeometryLayer => ({
          id: `legacy-${index + 1}`,
          name: `Compiled shape ${index + 1}`,
          kind: "geometry",
          visible: true,
          locked: false,
          transform: {},
          clipRect: null,
          shape: deepClone(shape),
          bounds: { x: 0, y: 0, width: canvasWidth, height: canvasHeight },
        }),
      ),
    },
    sourceMode: "flattened-legacy",
  };
}
