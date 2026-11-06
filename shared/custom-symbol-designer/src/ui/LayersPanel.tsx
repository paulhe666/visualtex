import { Eye, EyeOff, Lock, Unlock } from "lucide-react";
import type { CustomSymbolDesignerDocument, CustomSymbolDesignerLayer } from "../designerTypes.ts";
import type { CustomSymbolDesignerPlatform } from "../platform.ts";
import { shapeLabel } from "./MaterialsPanel.tsx";

interface Props {
  document: CustomSymbolDesignerDocument;
  selectedId: string | null;
  isEn: boolean;
  platform: CustomSymbolDesignerPlatform;
  onSelect: (id: string) => void;
  onChange: (id: string, update: (layer: CustomSymbolDesignerLayer) => CustomSymbolDesignerLayer) => void;
}

/** The layers, topmost first. */
export function LayersPanel({ document, selectedId, isEn, platform, onSelect, onChange }: Props) {
  const { MathPreview } = platform;
  const layers = [...document.layers].reverse();
  if (!layers.length) {
    return <p className="csd-empty">{isEn ? "Add a symbol or a shape from the left." : "从左侧加入符号或图形。"}</p>;
  }
  return (
    <ul className="csd-layer-list" data-csd-layer-list>
      {layers.map((layer) => (
        <li
          key={layer.id}
          className={`${selectedId === layer.id ? "is-selected" : ""}${layer.visible ? "" : " is-hidden"}`}
          data-csd-layer-row={layer.id}
          onClick={() => onSelect(layer.id)}
        >
          <span className="csd-layer-thumb">
            {layer.kind === "glyph" && /^[\x20-\x7e]+$/.test(layer.asset.sourceLatex) ? (
              <MathPreview latex={layer.asset.sourceLatex} fit />
            ) : layer.kind === "glyph" ? (
              <span>{layer.asset.sourceLatex}</span>
            ) : (
              <span className={`csd-shape-icon is-${layer.geometryPreset ?? "line"}`} aria-hidden="true" />
            )}
          </span>
          <span className="csd-layer-name">
            <strong>{layer.name}</strong>
            <small>{layer.kind === "glyph" ? layer.asset.sourceLatex : shapeLabel(layer.geometryPreset ?? "line", isEn)}</small>
          </span>
          <button
            type="button"
            title={layer.visible ? (isEn ? "Hide" : "隐藏") : isEn ? "Show" : "显示"}
            data-csd-layer-visibility
            onClick={(event) => {
              event.stopPropagation();
              onChange(layer.id, (item) => ({ ...item, visible: !item.visible }));
            }}
          >
            {layer.visible ? <Eye size={13} /> : <EyeOff size={13} />}
          </button>
          <button
            type="button"
            title={layer.locked ? (isEn ? "Unlock" : "解锁") : isEn ? "Lock" : "锁定"}
            data-csd-layer-lock
            onClick={(event) => {
              event.stopPropagation();
              onChange(layer.id, (item) => ({ ...item, locked: !item.locked }));
            }}
          >
            {layer.locked ? <Lock size={13} /> : <Unlock size={13} />}
          </button>
        </li>
      ))}
    </ul>
  );
}
