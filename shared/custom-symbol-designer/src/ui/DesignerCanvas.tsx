import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type PointerEvent as ReactPointerEvent,
  type ReactNode,
  type WheelEvent as ReactWheelEvent,
} from "react";
import { Maximize, Minus, Plus } from "lucide-react";
import type { CustomSymbolBox } from "../symbolBox.ts";
import { customSymbolDesignerLayerLocalShapes } from "../compiler.ts";
import type { CustomSymbolDesignerDocument, CustomSymbolDesignerLayer, CustomSymbolGlyphAsset } from "../designerTypes.ts";
import { buildSmoothCustomSymbolEraserPath } from "../geometry.ts";
import { apply, invert, svgTransformAttribute, transformMatrix, type Affine } from "../matrix.ts";
import { shapePaint, shapePathData } from "../shapes.ts";
import type { CustomSymbolVectorShape, CustomSymbolVectorTransform } from "../symbolTypes.ts";
import {
  CAP_HEIGHT,
  MATH_AXIS,
  X_HEIGHT,
  baselineOf,
  contentBounds,
  layerBox,
  layerMatrix,
  withCenteredOrigin,
} from "./layerMath.ts";

export type DesignerTool = "select" | "pan" | "eraser";

interface Props {
  document: CustomSymbolDesignerDocument;
  selectedId: string | null;
  tool: DesignerTool;
  eraserSize: number;
  reference: { asset: CustomSymbolGlyphAsset; label: string } | null;
  /** the registered symbol's box (its base character's, or the ink's) */
  crop: CustomSymbolBox | null;
  /** Changes when the view should frame the symbol again. */
  fitKey: string;
  isEn: boolean;
  onSelect: (id: string | null) => void;
  /** A layer's new transform during a drag; `gesture` is the same for one drag. */
  onTransform: (id: string, transform: CustomSymbolVectorTransform, gesture: string) => void;
  onGestureEnd: () => void;
  onEraserStroke: (points: Array<{ x: number; y: number }>) => void;
}

interface View {
  x: number;
  y: number;
  width: number;
  height: number;
}

type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

type Gesture =
  | { kind: "move"; pointerId: number; layerId: string; start: { x: number; y: number }; transform: CustomSymbolVectorTransform; key: string }
  | {
      kind: "resize";
      pointerId: number;
      layerId: string;
      handle: Handle;
      inverse: Affine;
      box: { x: number; y: number; width: number; height: number };
      transform: CustomSymbolVectorTransform;
      key: string;
    }
  | { kind: "rotate"; pointerId: number; layerId: string; center: { x: number; y: number }; startAngle: number; transform: CustomSymbolVectorTransform; key: string }
  | { kind: "pan"; pointerId: number; startClient: { x: number; y: number }; startView: View; moved: boolean; deselect: boolean }
  | { kind: "erase"; pointerId: number; points: Array<{ x: number; y: number }> };

/** 100 % shows one em (1000 units) as 100 CSS pixels. */
const PX_PER_EM_AT_100 = 100;
const MIN_PX_PER_UNIT = 0.004;
const MAX_PX_PER_UNIT = 4;

function handleGeometry(handle: Handle, box: { x: number; y: number; width: number; height: number }) {
  const left = box.x;
  const right = box.x + box.width;
  const top = box.y;
  const bottom = box.y + box.height;
  const cx = (left + right) / 2;
  const cy = (top + bottom) / 2;
  const positions: Record<Handle, [number, number, number, number]> = {
    nw: [left, top, right, bottom],
    n: [cx, top, cx, bottom],
    ne: [right, top, left, bottom],
    e: [right, cy, left, cy],
    se: [right, bottom, left, top],
    s: [cx, bottom, cx, top],
    sw: [left, bottom, right, top],
    w: [left, cy, right, cy],
  };
  const [hx, hy, ax, ay] = positions[handle];
  return { hx, hy, ax, ay };
}

function ShapeElement({ shape, clipId, paint }: { shape: CustomSymbolVectorShape; clipId: string; paint: string }) {
  const { fill, strokeWidth } = shapePaint(shape);
  const common = {
    fill: fill ? paint : "none",
    stroke: strokeWidth > 0 ? paint : "none",
    strokeWidth,
    strokeLinecap: shape.lineCap,
    strokeLinejoin: shape.lineJoin,
  } as const;
  // The user transform outside the crop, the compiler's matrix inside, as the symbol is drawn.
  const outer = svgTransformAttribute(shape.transform ? { ...shape.transform, matrix: undefined } : undefined);
  const matrix = shape.transform?.matrix ? `matrix(${shape.transform.matrix.join(" ")})` : undefined;
  let element: ReactNode;
  if (shape.kind === "text") {
    element = (
      <text
        x={shape.x}
        y={shape.y}
        fontFamily={shape.fontFamily}
        fontSize={shape.fontSize}
        fontStyle={shape.fontStyle}
        fontWeight={shape.fontWeight}
        transform={matrix}
        {...common}
      >
        {shape.text}
      </text>
    );
  } else {
    element = <path d={shapePathData(shape) ?? ""} transform={matrix} {...common} />;
  }
  if (shape.clipRect) {
    element = (
      <>
        <clipPath id={clipId} clipPathUnits="userSpaceOnUse">
          <rect {...shape.clipRect} />
        </clipPath>
        <g clipPath={`url(#${clipId})`}>{element}</g>
      </>
    );
  }
  return outer ? <g transform={outer}>{element}</g> : <>{element}</>;
}

function isErase(layer: CustomSymbolDesignerLayer) {
  return layer.kind === "geometry" && layer.shape.operation === "erase";
}

export function DesignerCanvas({
  document,
  selectedId,
  tool,
  eraserSize,
  reference,
  crop,
  fitKey,
  isEn,
  onSelect,
  onTransform,
  onGestureEnd,
  onEraserStroke,
}: Props) {
  const shellRef = useRef<HTMLDivElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);
  const [size, setSize] = useState({ width: 800, height: 600 });
  const [view, setView] = useState<View>({ x: -1500, y: 0, width: 4000, height: 3000 });
  const [gesture, setGesture] = useState<Gesture | null>(null);
  const [cursor, setCursor] = useState<{ x: number; y: number } | null>(null);
  const [spaceHeld, setSpaceHeld] = useState(false);
  const baseline = baselineOf(document);
  const pxPerUnit = size.width / Math.max(view.width, 1);
  const unitsPerPx = 1 / pxPerUnit;

  // The view keeps the element's aspect ratio, so a unit is square on screen.
  useLayoutEffect(() => {
    const shell = shellRef.current;
    if (!shell) return;
    const observer = new ResizeObserver(() => {
      const rect = shell.getBoundingClientRect();
      if (rect.width < 2 || rect.height < 2) return;
      setSize({ width: rect.width, height: rect.height });
      setView((current) => {
        const height = current.width * (rect.height / rect.width);
        return { ...current, y: current.y + (current.height - height) / 2, height };
      });
    });
    observer.observe(shell);
    return () => observer.disconnect();
  }, []);

  const referencePlacement = useCallback(() => {
    if (!reference) return null;
    const content = crop
      ? { x: crop.left }
      : contentBounds(document) ?? { x: 0 };
    const width = reference.asset.metrics.widthEm * 1000;
    return {
      x: content.x - width - 320,
      y: baseline - reference.asset.metrics.ascentEm * 1000,
      width,
      height: (reference.asset.metrics.ascentEm + reference.asset.metrics.descentEm) * 1000,
    };
  }, [baseline, crop, document, reference]);

  const fit = useCallback(() => {
    const boxes: Array<{ x: number; y: number; width: number; height: number }> = [];
    if (crop) {
      boxes.push({
        x: crop.left,
        y: crop.top,
        width: crop.metrics.widthEm * 1000,
        height: (crop.metrics.ascentEm + crop.metrics.descentEm) * 1000,
      });
    }
    const content = contentBounds(document);
    if (content) boxes.push(content);
    const ref = referencePlacement();
    if (ref) boxes.push(ref);
    if (!crop && !content) boxes.push({ x: 0, y: baseline - 800, width: 800, height: 1100 });
    const x = Math.min(...boxes.map((box) => box.x));
    const y = Math.min(...boxes.map((box) => box.y));
    const right = Math.max(...boxes.map((box) => box.x + box.width));
    const bottom = Math.max(...boxes.map((box) => box.y + box.height));
    const pad = Math.max(300, Math.max(right - x, bottom - y) * 0.35);
    const aspect = size.height / Math.max(size.width, 1);
    let width = right - x + pad * 2;
    let height = bottom - y + pad * 2;
    if (height / width > aspect) width = height / aspect;
    else height = width * aspect;
    const cx = (x + right) / 2;
    const cy = (y + bottom) / 2;
    setView({ x: cx - width / 2, y: cy - height / 2, width, height });
  }, [baseline, crop, document, referencePlacement, size]);

  // Frame the symbol when asked (a document opened, the first material added, the window resized at first).
  const fitRef = useRef(fit);
  fitRef.current = fit;
  useEffect(() => {
    const frame = requestAnimationFrame(() => fitRef.current());
    return () => cancelAnimationFrame(frame);
  }, [fitKey, size.width > 2 && size.height > 2]);

  useEffect(() => {
    const down = (event: KeyboardEvent) => {
      if (event.code !== "Space") return;
      const target = event.target as HTMLElement | null;
      if (target?.closest("input, textarea, select, [contenteditable='true']")) return;
      setSpaceHeld(true);
    };
    const up = (event: KeyboardEvent) => {
      if (event.code === "Space") setSpaceHeld(false);
    };
    window.addEventListener("keydown", down);
    window.addEventListener("keyup", up);
    return () => {
      window.removeEventListener("keydown", down);
      window.removeEventListener("keyup", up);
    };
  }, []);

  const worldPoint = (clientX: number, clientY: number) => {
    const rect = svgRef.current?.getBoundingClientRect();
    if (!rect) return { x: 0, y: 0 };
    return {
      x: view.x + ((clientX - rect.left) / rect.width) * view.width,
      y: view.y + ((clientY - rect.top) / rect.height) * view.height,
    };
  };

  const zoomBy = (factor: number, pivot?: { x: number; y: number }) => {
    setView((current) => {
      const nextPxPerUnit = Math.min(MAX_PX_PER_UNIT, Math.max(MIN_PX_PER_UNIT, (size.width / current.width) * factor));
      const width = size.width / nextPxPerUnit;
      const applied = width / current.width;
      const center = pivot ?? { x: current.x + current.width / 2, y: current.y + current.height / 2 };
      return {
        x: center.x - (center.x - current.x) * applied,
        y: center.y - (center.y - current.y) * applied,
        width,
        height: current.height * applied,
      };
    });
  };

  const onWheel = (event: ReactWheelEvent<SVGSVGElement>) => {
    event.preventDefault();
    if (event.ctrlKey || event.metaKey) {
      // pinch (trackpads report it as ctrl+wheel) or Ctrl/Cmd + wheel: zoom at the pointer
      zoomBy(Math.exp(-event.deltaY * 0.0025), worldPoint(event.clientX, event.clientY));
      return;
    }
    const dx = (event.shiftKey && !event.deltaX ? event.deltaY : event.deltaX) * unitsPerPx;
    const dy = (event.shiftKey && !event.deltaX ? 0 : event.deltaY) * unitsPerPx;
    setView((current) => ({ ...current, x: current.x + dx, y: current.y + dy }));
  };

  const capture = (event: ReactPointerEvent) => {
    svgRef.current?.setPointerCapture?.(event.pointerId);
  };

  const startPan = (event: ReactPointerEvent, deselect: boolean) => {
    capture(event);
    setGesture({
      kind: "pan",
      pointerId: event.pointerId,
      startClient: { x: event.clientX, y: event.clientY },
      startView: view,
      moved: false,
      deselect,
    });
  };

  const onBackgroundPointerDown = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (event.button !== 0 && event.button !== 1) return;
    svgRef.current?.focus({ preventScroll: true });
    if (event.button === 1 || tool === "pan" || spaceHeld) {
      startPan(event, false);
      event.preventDefault();
      return;
    }
    if (tool === "eraser") {
      const point = worldPoint(event.clientX, event.clientY);
      capture(event);
      setGesture({ kind: "erase", pointerId: event.pointerId, points: [point] });
      event.preventDefault();
      return;
    }
    startPan(event, true);
    event.preventDefault();
  };

  const onLayerPointerDown = (event: ReactPointerEvent, layer: CustomSymbolDesignerLayer) => {
    if (event.button !== 0 || tool !== "select" || spaceHeld) return;
    event.stopPropagation();
    event.preventDefault();
    svgRef.current?.focus({ preventScroll: true });
    onSelect(layer.id);
    if (layer.locked) return;
    capture(event);
    setGesture({
      kind: "move",
      pointerId: event.pointerId,
      layerId: layer.id,
      start: worldPoint(event.clientX, event.clientY),
      transform: { ...layer.transform },
      key: `move:${layer.id}:${event.timeStamp}`,
    });
  };

  const onHandlePointerDown = (event: ReactPointerEvent, layer: CustomSymbolDesignerLayer, handle: Handle) => {
    if (event.button !== 0 || layer.locked) return;
    event.stopPropagation();
    event.preventDefault();
    capture(event);
    setGesture({
      kind: "resize",
      pointerId: event.pointerId,
      layerId: layer.id,
      handle,
      inverse: invert(layerMatrix(layer)),
      box: layerBox(layer),
      transform: { ...layer.transform },
      key: `resize:${layer.id}:${event.timeStamp}`,
    });
  };

  const onRotatePointerDown = (event: ReactPointerEvent, layer: CustomSymbolDesignerLayer) => {
    if (event.button !== 0 || layer.locked) return;
    event.stopPropagation();
    event.preventDefault();
    const centered = withCenteredOrigin(layer);
    const box = layerBox(centered);
    const center = apply(layerMatrix(centered), box.x + box.width / 2, box.y + box.height / 2);
    const point = worldPoint(event.clientX, event.clientY);
    capture(event);
    setGesture({
      kind: "rotate",
      pointerId: event.pointerId,
      layerId: layer.id,
      center,
      startAngle: Math.atan2(point.y - center.y, point.x - center.x),
      transform: { ...centered.transform },
      key: `rotate:${layer.id}:${event.timeStamp}`,
    });
  };

  const onPointerMove = (event: ReactPointerEvent<SVGSVGElement>) => {
    const point = worldPoint(event.clientX, event.clientY);
    if (tool === "eraser") setCursor(point);
    if (!gesture || gesture.pointerId !== event.pointerId) return;
    switch (gesture.kind) {
      case "pan": {
        const dx = (event.clientX - gesture.startClient.x) * unitsPerPx;
        const dy = (event.clientY - gesture.startClient.y) * unitsPerPx;
        if (!gesture.moved && Math.hypot(dx, dy) * pxPerUnit > 3) setGesture({ ...gesture, moved: true });
        setView({ ...gesture.startView, x: gesture.startView.x - dx, y: gesture.startView.y - dy });
        return;
      }
      case "erase": {
        const last = gesture.points[gesture.points.length - 1];
        if (last && Math.hypot(point.x - last.x, point.y - last.y) < unitsPerPx * 1.5) return;
        setGesture({ ...gesture, points: [...gesture.points, point] });
        return;
      }
      case "move": {
        let dx = point.x - gesture.start.x;
        let dy = point.y - gesture.start.y;
        if (event.shiftKey) {
          // constrain to the dominant axis
          if (Math.abs(dx) > Math.abs(dy)) dy = 0;
          else dx = 0;
        }
        onTransform(
          gesture.layerId,
          {
            ...gesture.transform,
            translateX: (gesture.transform.translateX ?? 0) + dx,
            translateY: (gesture.transform.translateY ?? 0) + dy,
          },
          gesture.key,
        );
        return;
      }
      case "resize": {
        const local = apply(gesture.inverse, point.x, point.y);
        const { hx, hy, ax, ay } = handleGeometry(gesture.handle, gesture.box);
        let fx = hx !== ax ? (local.x - ax) / (hx - ax) : 1;
        let fy = hy !== ay ? (local.y - ay) / (hy - ay) : 1;
        const horizontal = gesture.handle === "e" || gesture.handle === "w";
        const vertical = gesture.handle === "n" || gesture.handle === "s";
        if (!horizontal && !vertical && !event.shiftKey) {
          // corners keep the proportions (Shift: free)
          const uniform = Math.abs(fx - 1) >= Math.abs(fy - 1) ? fx : fy;
          fx = uniform;
          fy = uniform;
        }
        const sx0 = gesture.transform.scaleX ?? 1;
        const sy0 = gesture.transform.scaleY ?? 1;
        const scaleX = vertical ? sx0 : Math.sign(sx0 || 1) * Math.max(0.02, Math.abs(sx0) * Math.max(0.02, fx));
        const scaleY = horizontal ? sy0 : Math.sign(sy0 || 1) * Math.max(0.02, Math.abs(sy0) * Math.max(0.02, fy));
        // keep the opposite handle where it is
        const before = apply(transformMatrix(gesture.transform, false), ax, ay);
        const next = { ...gesture.transform, scaleX, scaleY };
        const after = apply(transformMatrix(next, false), ax, ay);
        onTransform(
          gesture.layerId,
          {
            ...next,
            translateX: (gesture.transform.translateX ?? 0) + before.x - after.x,
            translateY: (gesture.transform.translateY ?? 0) + before.y - after.y,
          },
          gesture.key,
        );
        return;
      }
      case "rotate": {
        const angle = Math.atan2(point.y - gesture.center.y, point.x - gesture.center.x);
        let delta = ((angle - gesture.startAngle) * 180) / Math.PI;
        while (delta > 180) delta -= 360;
        while (delta < -180) delta += 360;
        let rotation = (gesture.transform.rotateDeg ?? 0) + delta;
        if (event.shiftKey) rotation = Math.round(rotation / 15) * 15;
        else {
          const right = Math.round(rotation / 90) * 90;
          if (Math.abs(rotation - right) < 4) rotation = right;
        }
        onTransform(gesture.layerId, { ...gesture.transform, rotateDeg: Number(rotation.toFixed(2)) }, gesture.key);
        return;
      }
    }
  };

  const onPointerUp = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (!gesture || gesture.pointerId !== event.pointerId) return;
    if (gesture.kind === "erase" && gesture.points.length) onEraserStroke(gesture.points);
    if (gesture.kind === "pan" && !gesture.moved && gesture.deselect) onSelect(null);
    if (gesture.kind === "move" || gesture.kind === "resize" || gesture.kind === "rotate") onGestureEnd();
    setGesture(null);
    svgRef.current?.releasePointerCapture?.(event.pointerId);
  };

  const visible = document.layers.filter((layer) => layer.visible);
  const paintLayers = visible.filter((layer) => !isErase(layer));
  const eraseLayers = visible.filter(isErase);
  const liveErase = gesture?.kind === "erase" ? buildSmoothCustomSymbolEraserPath(gesture.points) : "";
  const masked = eraseLayers.length > 0 || Boolean(liveErase);
  const ref = referencePlacement();
  const selected = visible.find((layer) => layer.id === selectedId) ?? null;
  const zoomPercent = Math.round(pxPerUnit * 1000 * (100 / PX_PER_EM_AT_100));
  const px = (value: number) => value * unitsPerPx;
  const viewRight = view.x + view.width;
  const guides = [
    { y: baseline, className: "is-baseline", label: isEn ? "baseline" : "基线" },
    { y: baseline - MATH_AXIS, className: "is-axis", label: isEn ? "math axis" : "数学轴" },
    { y: baseline - X_HEIGHT, className: "is-minor", label: isEn ? "x-height" : "x 高度" },
    { y: baseline - CAP_HEIGHT, className: "is-minor", label: isEn ? "cap height" : "大写高度" },
  ];
  const cursorClass =
    gesture?.kind === "pan" ? " is-panning" : tool === "pan" || spaceHeld ? " is-pan-tool" : tool === "eraser" ? " is-erasing" : "";

  const renderSelection = (layer: CustomSymbolDesignerLayer) => {
    const box = layerBox(layer);
    const matrix = layerMatrix(layer);
    const corner = (x: number, y: number) => apply(matrix, x, y);
    const polygon = [
      corner(box.x, box.y),
      corner(box.x + box.width, box.y),
      corner(box.x + box.width, box.y + box.height),
      corner(box.x, box.y + box.height),
    ];
    const outline = polygon.map((point) => `${point.x},${point.y}`).join(" ");
    if (isErase(layer)) {
      return <polygon points={outline} className="csd-selection is-eraser" pointerEvents="none" />;
    }
    const topMid = corner(box.x + box.width / 2, box.y);
    const bottomMid = corner(box.x + box.width / 2, box.y + box.height);
    const up = { x: topMid.x - bottomMid.x, y: topMid.y - bottomMid.y };
    const length = Math.hypot(up.x, up.y) || 1;
    const stem = px(26);
    const knob = { x: topMid.x + (up.x / length) * stem, y: topMid.y + (up.y / length) * stem };
    const handleSize = px(9);
    return (
      <g className="csd-selection-group" data-csd-selection={layer.id}>
        <polygon points={outline} className="csd-selection" pointerEvents="none" />
        {!layer.locked ? (
          <>
            <line x1={topMid.x} y1={topMid.y} x2={knob.x} y2={knob.y} className="csd-rotate-stem" pointerEvents="none" />
            <circle
              cx={knob.x}
              cy={knob.y}
              r={px(5.5)}
              className="csd-rotate-handle"
              data-csd-rotate-handle
              onPointerDown={(event) => onRotatePointerDown(event, layer)}
            />
            {HANDLES.map((handle) => {
              const { hx, hy } = handleGeometry(handle, box);
              const point = corner(hx, hy);
              return (
                <rect
                  key={handle}
                  x={point.x - handleSize / 2}
                  y={point.y - handleSize / 2}
                  width={handleSize}
                  height={handleSize}
                  rx={px(1.5)}
                  className={`csd-resize-handle is-${handle}`}
                  data-csd-resize-handle={handle}
                  onPointerDown={(event) => onHandlePointerDown(event, layer, handle)}
                />
              );
            })}
          </>
        ) : null}
      </g>
    );
  };

  return (
    <div className="csd-canvas-shell" ref={shellRef} data-csd-canvas>
      <svg
        ref={svgRef}
        className={`csd-canvas${cursorClass}`}
        viewBox={`${view.x} ${view.y} ${view.width} ${view.height}`}
        tabIndex={-1}
        data-csd-view={`${Math.round(view.x)} ${Math.round(view.y)} ${Math.round(view.width)} ${Math.round(view.height)}`}
        onWheel={onWheel}
        onPointerDown={onBackgroundPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
        onPointerLeave={() => setCursor(null)}
      >
        <defs>
          <pattern id="csd-grid-minor" width="100" height="100" patternUnits="userSpaceOnUse" y={baseline}>
            <path d="M100 0H0V100" className="csd-gridline is-minor" />
          </pattern>
          <pattern id="csd-grid-major" width="500" height="500" patternUnits="userSpaceOnUse" y={baseline}>
            <path d="M500 0H0V500" className="csd-gridline is-major" />
          </pattern>
          {masked ? (
            <mask id="csd-erase-mask" maskUnits="userSpaceOnUse" x={view.x - view.width} y={view.y - view.height} width={view.width * 3} height={view.height * 3}>
              <rect x={view.x - view.width} y={view.y - view.height} width={view.width * 3} height={view.height * 3} fill="white" />
              {eraseLayers.map((layer) =>
                layer.kind === "geometry" ? (
                  <g key={layer.id} transform={svgTransformAttribute(layer.transform)}>
                    <ShapeElement shape={{ ...layer.shape, operation: undefined }} clipId={`csd-erase-${layer.id}`} paint="black" />
                  </g>
                ) : null,
              )}
              {liveErase ? (
                <path d={liveErase} fill="none" stroke="black" strokeWidth={eraserSize} strokeLinecap="round" strokeLinejoin="round" />
              ) : null}
            </mask>
          ) : null}
        </defs>

        {(() => {
          // A background three views wide stays under the pointer while panning.
          const background = { x: view.x - view.width, y: view.y - view.height, width: view.width * 3, height: view.height * 3 };
          return (
            <>
              <rect {...background} className="csd-paper" />
              {pxPerUnit * 100 >= 6 ? <rect {...background} fill="url(#csd-grid-minor)" pointerEvents="none" /> : null}
              <rect {...background} fill="url(#csd-grid-major)" pointerEvents="none" />
            </>
          );
        })()}

        {guides.map((guide) => (
          <g key={guide.className + guide.y} className={`csd-guide ${guide.className}`} pointerEvents="none">
            <line x1={view.x - view.width} x2={viewRight + view.width} y1={guide.y} y2={guide.y} />
            <text x={view.x + px(8)} y={guide.y - px(4)} fontSize={px(10)}>{guide.label}</text>
          </g>
        ))}
        <line
          className="csd-guide is-origin"
          x1={0}
          x2={0}
          y1={view.y - view.height}
          y2={view.y + view.height * 2}
          pointerEvents="none"
        />

        {crop ? (
          <g className="csd-crop" data-csd-crop pointerEvents="none">
            <rect
              x={crop.left}
              y={crop.top}
              width={crop.metrics.widthEm * 1000}
              height={(crop.metrics.ascentEm + crop.metrics.descentEm) * 1000}
            />
            <text
              x={crop.left}
              y={crop.top + (crop.metrics.ascentEm + crop.metrics.descentEm) * 1000 + px(14)}
              fontSize={px(10.5)}
            >
              {`${crop.metrics.widthEm.toFixed(2)} × ${crop.metrics.ascentEm.toFixed(2)} + ${crop.metrics.descentEm.toFixed(2)} em`}
            </text>
          </g>
        ) : null}

        {reference && ref ? (
          <g className="csd-reference" transform={`translate(${ref.x} ${ref.y})`} pointerEvents="none" data-csd-reference={reference.label}>
            {reference.asset.shapes.map((shape, index) => (
              <ShapeElement key={index} shape={shape} clipId={`csd-ref-${index}`} paint="currentColor" />
            ))}
          </g>
        ) : null}

        <g className="csd-artwork" mask={masked ? "url(#csd-erase-mask)" : undefined}>
          {paintLayers.map((layer) => (
            <g key={layer.id} transform={svgTransformAttribute(layer.transform)} data-csd-artwork={layer.id}>
              {customSymbolDesignerLayerLocalShapes(layer).map((shape, index) => (
                <ShapeElement
                  key={index}
                  shape={layer.kind === "glyph" && layer.clipRect ? { ...shape, clipRect: layer.clipRect } : shape}
                  clipId={`csd-clip-${layer.id}-${index}`}
                  paint="currentColor"
                />
              ))}
            </g>
          ))}
        </g>

        {tool === "select"
          ? [...paintLayers, ...eraseLayers].map((layer) => {
              const box = layerBox(layer);
              return (
                <rect
                  key={layer.id}
                  {...box}
                  transform={svgTransformAttribute(layer.transform)}
                  className="csd-hit"
                  data-csd-layer={layer.id}
                  onPointerDown={(event) => onLayerPointerDown(event, layer)}
                />
              );
            })
          : null}

        {selected && tool === "select" ? renderSelection(selected) : null}

        {tool === "eraser" && cursor ? (
          <circle cx={cursor.x} cy={cursor.y} r={Math.max(px(2), eraserSize / 2)} className="csd-eraser-cursor" pointerEvents="none" />
        ) : null}
      </svg>
      <div className="csd-zoom" aria-label={isEn ? "Zoom" : "缩放"}>
        <button type="button" onClick={() => zoomBy(1 / 1.25)} title={isEn ? "Zoom out" : "缩小"} data-csd-zoom-out>
          <Minus size={13} />
        </button>
        <span data-csd-zoom-percent>{zoomPercent}%</span>
        <button type="button" onClick={() => zoomBy(1.25)} title={isEn ? "Zoom in" : "放大"} data-csd-zoom-in>
          <Plus size={13} />
        </button>
        <button type="button" onClick={fit} title={isEn ? "Fit the symbol" : "适合字符"} data-csd-fit>
          <Maximize size={13} />
        </button>
      </div>
    </div>
  );
}
