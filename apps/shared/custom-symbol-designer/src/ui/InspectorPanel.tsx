import { useState, type ReactNode } from "react";
import {
  AlignCenterHorizontal,
  AlignCenterVertical,
  ArrowDownToLine,
  BringToFront,
  ChevronRight,
  Copy,
  FlipHorizontal2,
  FlipVertical2,
  Link2,
  Link2Off,
  RotateCcw,
  RotateCw,
  SendToBack,
  Trash2,
} from "lucide-react";
import type { CustomSymbolDesignerLayer } from "../designerTypes.ts";
import { customSymbolGeometryProperties, glyphLayerFullClip, updateCustomSymbolGeometryLayer } from "../geometry.ts";
import type { CustomSymbolLayerEffects } from "../symbolTypes.ts";
import { anchorPosition, withAnchorPosition, withCenteredOrigin } from "./layerMath.ts";
import { shapeLabel } from "./MaterialsPanel.tsx";

export type LayerAction =
  | "duplicate"
  | "delete"
  | "forward"
  | "backward"
  | "baseline"
  | "center"
  | "axis"
  | "split-horizontal"
  | "split-vertical";

interface Props {
  layer: CustomSymbolDesignerLayer;
  baseline: number;
  isEn: boolean;
  onChange: (update: (layer: CustomSymbolDesignerLayer) => CustomSymbolDesignerLayer, gesture?: string) => void;
  onAction: (action: LayerAction) => void;
}

function NumberField({
  label,
  value,
  step,
  digits = 3,
  min,
  max,
  suffix,
  field,
  onChange,
}: {
  label: string;
  value: number;
  step: number;
  digits?: number;
  min?: number;
  max?: number;
  suffix?: string;
  field: string;
  onChange: (value: number) => void;
}) {
  const [draft, setDraft] = useState<string | null>(null);
  const shown = draft ?? String(Number(value.toFixed(digits)));
  return (
    <label className="csd-field">
      <span>{label}</span>
      <span className="csd-field-input">
        <input
          type="number"
          value={shown}
          step={step}
          min={min}
          max={max}
          data-csd-field={field}
          onFocus={() => setDraft(String(Number(value.toFixed(digits))))}
          onBlur={() => setDraft(null)}
          onChange={(event) => {
            const text = event.currentTarget.value;
            setDraft(text);
            const next = Number(text);
            if (text.trim() !== "" && Number.isFinite(next)) {
              onChange(Math.min(max ?? Infinity, Math.max(min ?? -Infinity, next)));
            }
          }}
        />
        {suffix ? <em>{suffix}</em> : null}
      </span>
    </label>
  );
}

function Section({ title, children, defaultOpen = true }: { title: string; children: ReactNode; defaultOpen?: boolean }) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <section className={`csd-section${open ? " is-open" : ""}`}>
      <button type="button" className="csd-section-title" aria-expanded={open} onClick={() => setOpen(!open)}>
        <ChevronRight size={13} />
        {title}
      </button>
      {open ? <div className="csd-section-body">{children}</div> : null}
    </section>
  );
}

export function InspectorPanel({ layer, baseline, isEn, onChange, onAction }: Props) {
  const [lockRatio, setLockRatio] = useState(true);
  const t = (zh: string, en: string) => (isEn ? en : zh);
  const position = anchorPosition(layer, baseline);
  const scaleX = layer.transform.scaleX ?? 1;
  const scaleY = layer.transform.scaleY ?? 1;
  const isEraser = layer.kind === "geometry" && layer.geometryPreset === "eraser";
  const geometry = layer.kind === "geometry" ? customSymbolGeometryProperties(layer) : null;
  const gesture = (field: string) => `field:${layer.id}:${field}`;

  /** A change of rotation, slant or scale: about the box centre, so the layer stays in place. */
  const changeTransform = (patch: Partial<CustomSymbolDesignerLayer["transform"]>, field: string) =>
    onChange((current) => {
      const centered = withCenteredOrigin(current);
      return { ...centered, transform: { ...centered.transform, ...patch } };
    }, gesture(field));

  const setScale = (axis: "x" | "y", percent: number) => {
    const magnitude = Math.max(0.02, percent / 100);
    const signX = Math.sign(scaleX) || 1;
    const signY = Math.sign(scaleY) || 1;
    if (lockRatio) {
      const ratio = axis === "x" ? magnitude / Math.abs(scaleX) : magnitude / Math.abs(scaleY);
      changeTransform({ scaleX: signX * Math.abs(scaleX) * ratio, scaleY: signY * Math.abs(scaleY) * ratio }, "scale");
    } else if (axis === "x") {
      changeTransform({ scaleX: signX * magnitude }, "scale-x");
    } else {
      changeTransform({ scaleY: signY * magnitude }, "scale-y");
    }
  };

  const setEffect = (kind: keyof CustomSymbolLayerEffects, patch: Record<string, number | boolean>, field: string) =>
    onChange((current) => {
      const defaults = kind === "outline" ? { enabled: false, width: 60 } : { enabled: false, depth: 240, angleDeg: 35, steps: 8 };
      return {
        ...current,
        effects: { ...current.effects, [kind]: { ...defaults, ...(current.effects?.[kind] ?? {}), ...patch } },
      };
    }, gesture(field));

  const setClip = (preset: "full" | "top" | "middle" | "bottom" | "left" | "center" | "right") =>
    onChange((current) => {
      if (current.kind !== "glyph") return current;
      const full = glyphLayerFullClip(current);
      const w = full.width / 3;
      const h = full.height / 3;
      const clips = {
        full: null,
        top: { ...full, height: h },
        middle: { ...full, y: h, height: h },
        bottom: { ...full, y: h * 2, height: h },
        left: { ...full, width: w },
        center: { ...full, x: w, width: w },
        right: { ...full, x: w * 2, width: w },
      };
      return { ...current, clipRect: clips[preset] };
    });

  return (
    <div className="csd-inspector" data-csd-inspector={layer.id}>
      <div className="csd-layer-title">
        <input
          value={layer.name}
          aria-label={t("图层名称", "Layer name")}
          data-csd-layer-name
          onChange={(event) => {
            const name = event.currentTarget.value;
            onChange((current) => ({ ...current, name }), gesture("name"));
          }}
        />
        <span className="csd-tag">
          {layer.kind === "glyph" ? t("字符", "Glyph") : shapeLabel(layer.geometryPreset ?? "line", isEn)}
        </span>
      </div>

      <Section title={t("位置与大小", "Position & size")}>
        <div className="csd-field-grid">
          <NumberField label={t("水平位置", "X")} suffix="em" step={0.01} value={position.xEm} field="x"
            onChange={(value) => onChange((current) => withAnchorPosition(current, baseline, { xEm: value }), gesture("x"))} />
          <NumberField label={t("基线上升", "Raise")} suffix="em" step={0.01} value={position.raiseEm} field="raise"
            onChange={(value) => onChange((current) => withAnchorPosition(current, baseline, { raiseEm: value }), gesture("raise"))} />
          <NumberField label={t("横向缩放", "Width")} suffix="%" step={1} digits={1} min={2} value={Math.abs(scaleX) * 100} field="scale-x"
            onChange={(value) => setScale("x", value)} />
          <NumberField label={t("纵向缩放", "Height")} suffix="%" step={1} digits={1} min={2} value={Math.abs(scaleY) * 100} field="scale-y"
            onChange={(value) => setScale("y", value)} />
          <NumberField label={t("旋转", "Rotation")} suffix="°" step={1} digits={2} value={layer.transform.rotateDeg ?? 0} field="rotate"
            onChange={(value) => changeTransform({ rotateDeg: value }, "rotate")} />
          <NumberField label={t("倾斜", "Slant")} suffix="°" step={1} digits={2} min={-60} max={60} value={-(layer.transform.skewXDeg ?? 0)} field="slant"
            onChange={(value) => changeTransform({ skewXDeg: -value }, "slant")} />
        </div>
        <div className="csd-icon-row">
          <button type="button" className={lockRatio ? "is-active" : ""} aria-pressed={lockRatio} title={t("锁定宽高比", "Keep proportions")} onClick={() => setLockRatio(!lockRatio)}>
            {lockRatio ? <Link2 size={14} /> : <Link2Off size={14} />}
          </button>
          <button type="button" title={t("水平翻转", "Flip horizontally")} data-csd-flip="horizontal" onClick={() => changeTransform({ scaleX: -scaleX }, "flip-x")}>
            <FlipHorizontal2 size={14} />
          </button>
          <button type="button" title={t("垂直翻转", "Flip vertically")} data-csd-flip="vertical" onClick={() => changeTransform({ scaleY: -scaleY }, "flip-y")}>
            <FlipVertical2 size={14} />
          </button>
          <button type="button" title={t("逆时针 90°", "Rotate −90°")} onClick={() => changeTransform({ rotateDeg: (layer.transform.rotateDeg ?? 0) - 90 }, "rot-90")}>
            <RotateCcw size={14} />
          </button>
          <button type="button" title={t("顺时针 90°", "Rotate +90°")} onClick={() => changeTransform({ rotateDeg: (layer.transform.rotateDeg ?? 0) + 90 }, "rot+90")}>
            <RotateCw size={14} />
          </button>
          <button type="button" title={t("恢复原始大小与角度", "Reset size and angle")} data-csd-reset-transform
            onClick={() => changeTransform({ scaleX: 1, scaleY: 1, rotateDeg: 0, skewXDeg: 0, skewYDeg: 0 }, "reset")}>
            1:1
          </button>
        </div>
      </Section>

      <Section title={t("对齐与排列", "Align & arrange")}>
        <div className="csd-icon-row">
          <button type="button" title={t("落在基线上", "Sit on the baseline")} data-csd-action="baseline" onClick={() => onAction("baseline")}>
            <ArrowDownToLine size={14} />
          </button>
          <button type="button" title={t("与其他图层水平居中", "Centre over the other layers")} data-csd-action="center" onClick={() => onAction("center")}>
            <AlignCenterVertical size={14} />
          </button>
          <button type="button" title={t("以数学轴为中心", "Centre on the math axis")} data-csd-action="axis" onClick={() => onAction("axis")}>
            <AlignCenterHorizontal size={14} />
          </button>
          <button type="button" title={t("上移一层", "Bring forward")} onClick={() => onAction("forward")}>
            <BringToFront size={14} />
          </button>
          <button type="button" title={t("下移一层", "Send backward")} onClick={() => onAction("backward")}>
            <SendToBack size={14} />
          </button>
          <button type="button" title={t("复制图层", "Duplicate")} data-csd-action="duplicate" onClick={() => onAction("duplicate")}>
            <Copy size={14} />
          </button>
          <button type="button" className="is-danger" title={t("删除图层", "Delete")} data-csd-action="delete" onClick={() => onAction("delete")}>
            <Trash2 size={14} />
          </button>
        </div>
      </Section>

      {geometry ? (
        <Section title={isEraser ? t("橡皮擦轨迹", "Eraser stroke") : t("图形", "Shape")}>
          <div className="csd-field-grid">
            {!isEraser ? (
              <NumberField label={layer.kind === "geometry" && layer.geometryPreset === "line" ? t("长度", "Length") : t("宽度", "Width")}
                suffix="em" step={0.01} min={0.01} value={geometry.width / 1000} field="shape-width"
                onChange={(value) => onChange((current) => current.kind === "geometry" ? updateCustomSymbolGeometryLayer(current, { width: value * 1000 }) : current, gesture("shape-width"))} />
            ) : null}
            {!isEraser && layer.kind === "geometry" && layer.geometryPreset !== "line" && layer.geometryPreset !== "circle" ? (
              <NumberField label={t("高度", "Height")} suffix="em" step={0.01} min={0.01} value={geometry.height / 1000} field="shape-height"
                onChange={(value) => onChange((current) => current.kind === "geometry" ? updateCustomSymbolGeometryLayer(current, { height: value * 1000 }) : current, gesture("shape-height"))} />
            ) : null}
            <NumberField label={isEraser ? t("擦除宽度", "Eraser width") : t("线宽", "Stroke")} suffix="em" step={0.005} min={0} value={geometry.strokeWidth / 1000} field="shape-stroke"
              onChange={(value) => onChange((current) => current.kind === "geometry" ? updateCustomSymbolGeometryLayer(current, { strokeWidth: value * 1000 }) : current, gesture("shape-stroke"))} />
            {layer.kind === "geometry" && layer.geometryPreset === "rect" ? (
              <NumberField label={t("圆角", "Corner")} suffix="em" step={0.005} min={0} value={geometry.cornerRadius / 1000} field="shape-corner"
                onChange={(value) => onChange((current) => current.kind === "geometry" ? updateCustomSymbolGeometryLayer(current, { cornerRadius: value * 1000 }) : current, gesture("shape-corner"))} />
            ) : null}
          </div>
          {layer.kind === "geometry" && !isEraser && !["line", "arrow", "arc"].includes(layer.geometryPreset ?? "") ? (
            <label className="csd-check">
              <input type="checkbox" checked={geometry.fill} data-csd-shape-fill
                onChange={(event) => {
                  const fill = event.currentTarget.checked;
                  onChange((current) => current.kind === "geometry" ? updateCustomSymbolGeometryLayer(current, { fill }) : current);
                }} />
              <span>{t("填充", "Filled")}</span>
            </label>
          ) : null}
        </Section>
      ) : null}

      {layer.kind === "glyph" ? (
        <Section title={t("裁剪", "Crop")} defaultOpen={Boolean(layer.clipRect)}>
          <div className="csd-chips is-tight">
            {([
              ["full", t("完整", "Full")],
              ["top", t("上", "Top")],
              ["middle", t("中", "Middle")],
              ["bottom", t("下", "Bottom")],
              ["left", t("左", "Left")],
              ["center", t("中列", "Centre")],
              ["right", t("右", "Right")],
            ] as const).map(([id, label]) => (
              <button key={id} type="button" data-csd-crop={id} onClick={() => setClip(id)}>{label}</button>
            ))}
          </div>
          {layer.clipRect ? (
            <div className="csd-field-grid">
              {(["x", "y", "width", "height"] as const).map((key) => (
                <NumberField key={key} label={{ x: "X", y: "Y", width: t("宽", "W"), height: t("高", "H") }[key]} suffix="em" step={0.01}
                  min={key === "width" || key === "height" ? 0.001 : undefined}
                  value={(layer.clipRect?.[key] ?? 0) / 1000} field={`crop-${key}`}
                  onChange={(value) => onChange((current) => current.kind === "glyph" && current.clipRect
                    ? { ...current, clipRect: { ...current.clipRect, [key]: value * 1000 } }
                    : current, gesture(`crop-${key}`))} />
              ))}
            </div>
          ) : null}
          <div className="csd-chips is-tight">
            <button type="button" data-csd-action="split-horizontal" onClick={() => onAction("split-horizontal")}>{t("上中下三分", "Split in 3 rows")}</button>
            <button type="button" data-csd-action="split-vertical" onClick={() => onAction("split-vertical")}>{t("左中右三分", "Split in 3 columns")}</button>
          </div>
        </Section>
      ) : null}

      {!isEraser ? (
        <Section title={t("效果", "Effects")} defaultOpen={Boolean(layer.effects?.outline?.enabled || layer.effects?.perspective?.enabled)}>
          <label className="csd-check">
            <input type="checkbox" checked={layer.effects?.outline?.enabled ?? false}
              onChange={(event) => setEffect("outline", { enabled: event.currentTarget.checked }, "outline")} />
            <span>{t("空心描边", "Hollow outline")}</span>
          </label>
          {layer.effects?.outline?.enabled ? (
            <div className="csd-field-grid">
              <NumberField label={t("描边粗细", "Outline")} suffix="em" step={0.005} min={0.001} value={layer.effects.outline.width / 1000} field="outline-width"
                onChange={(value) => setEffect("outline", { width: value * 1000 }, "outline-width")} />
            </div>
          ) : null}
          <label className="csd-check">
            <input type="checkbox" checked={layer.effects?.perspective?.enabled ?? false}
              onChange={(event) => setEffect("perspective", { enabled: event.currentTarget.checked }, "perspective")} />
            <span>{t("立体挤出", "Extrusion")}</span>
          </label>
          {layer.effects?.perspective?.enabled ? (
            <div className="csd-field-grid">
              <NumberField label={t("深度", "Depth")} suffix="em" step={0.01} min={0} value={layer.effects.perspective.depth / 1000} field="extrude-depth"
                onChange={(value) => setEffect("perspective", { depth: value * 1000 }, "extrude-depth")} />
              <NumberField label={t("方向", "Direction")} suffix="°" step={1} value={layer.effects.perspective.angleDeg} field="extrude-angle"
                onChange={(value) => setEffect("perspective", { angleDeg: value }, "extrude-angle")} />
              <NumberField label={t("层数", "Steps")} step={1} digits={0} min={1} max={24} value={layer.effects.perspective.steps} field="extrude-steps"
                onChange={(value) => setEffect("perspective", { steps: Math.round(value) }, "extrude-steps")} />
            </div>
          ) : null}
        </Section>
      ) : null}
    </div>
  );
}
