import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type PointerEvent as ReactPointerEvent,
  type ReactNode,
  type WheelEvent as ReactWheelEvent,
} from "react";
import type {
  CustomSymbolDesignerDocument,
  CustomSymbolDesignerLayer,
  CustomSymbolGlyphAsset,
} from "../math/customSymbolDesignerTypes";
import { customSymbolDesignerLayerLocalShapes } from "../math/customSymbolDesignerCompiler";
import { buildSmoothCustomSymbolEraserPath } from "../math/customSymbolDesignerGeometry";
import type {
  CustomSymbolVectorShape,
  CustomSymbolVectorTransform,
} from "../math/customSymbolTypes";

interface Props {
  documentState: CustomSymbolDesignerDocument;
  selectedLayerId: string | null;
  referenceAsset: CustomSymbolGlyphAsset | null;
  showReference: boolean;
  referenceLabel: string;
  eraserMode: boolean;
  eraserSize: number;
  isEn: boolean;
  onSelectLayer: (layerId: string | null) => void;
  onMoveLayer: (layerId: string, x: number, y: number) => void;
  onResizeLayer: (
    layerId: string,
    scaleX: number,
    scaleY: number,
    translateX: number,
    translateY: number,
  ) => void;
  onRotateLayer: (layerId: string, rotateDeg: number) => void;
  onDeleteLayer: (layerId: string) => void;
  onAddEraserStroke: (points: Array<{ x: number; y: number }>) => void;
}

interface ViewBoxState {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface DragState {
  pointerId: number;
  layerId: string;
  startX: number;
  startY: number;
  translateX: number;
  translateY: number;
}

interface PanState {
  pointerId: number;
  startClientX: number;
  startClientY: number;
  startViewBox: ViewBoxState;
  moved: boolean;
}

interface EraseState {
  pointerId: number;
  points: Array<{ x: number; y: number }>;
}

type ResizeHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

const infiniteCanvasExtent = 1_000_000;

interface ResizeState {
  pointerId: number;
  layerId: string;
  handle: ResizeHandle;
  inverseMatrix: DOMMatrix;
  bounds: { x: number; y: number; width: number; height: number };
  anchorX: number;
  anchorY: number;
  originalTransform: CustomSymbolVectorTransform;
}

interface RotateState {
  pointerId: number;
  layerId: string;
  centerClientX: number;
  centerClientY: number;
  startPointerAngle: number;
  originalRotateDeg: number;
}

interface AffineMatrix {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}

function svgTransform(transform?: CustomSymbolVectorTransform) {
  if (!transform) return undefined;
  const parts: string[] = [];
  const tx = transform.translateX ?? 0;
  const ty = transform.translateY ?? 0;
  const sx = transform.scaleX ?? 1;
  const sy = transform.scaleY ?? 1;
  const skewX = transform.skewXDeg ?? 0;
  const skewY = transform.skewYDeg ?? 0;
  const angle = transform.rotateDeg ?? 0;
  const ox = transform.originX ?? 0;
  const oy = transform.originY ?? 0;
  if (tx || ty) parts.push(`translate(${tx} ${ty})`);
  if (ox || oy) parts.push(`translate(${ox} ${oy})`);
  if (angle) parts.push(`rotate(${angle})`);
  if (skewX) parts.push(`skewX(${skewX})`);
  if (skewY) parts.push(`skewY(${skewY})`);
  if (sx !== 1 || sy !== 1) parts.push(`scale(${sx} ${sy})`);
  if (ox || oy) parts.push(`translate(${-ox} ${-oy})`);
  if (transform.matrix) parts.push(`matrix(${transform.matrix.join(" ")})`);
  return parts.length ? parts.join(" ") : undefined;
}

function multiplyAffine(left: AffineMatrix, right: AffineMatrix): AffineMatrix {
  return {
    a: left.a * right.a + left.c * right.b,
    b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d,
    d: left.b * right.c + left.d * right.d,
    e: left.a * right.e + left.c * right.f + left.e,
    f: left.b * right.e + left.d * right.f + left.f,
  };
}

function translationAffine(x: number, y: number): AffineMatrix {
  return { a: 1, b: 0, c: 0, d: 1, e: x, f: y };
}

function rotationAffine(degrees: number): AffineMatrix {
  const radians = (degrees * Math.PI) / 180;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  return { a: cosine, b: sine, c: -sine, d: cosine, e: 0, f: 0 };
}

function scaleAffine(x: number, y: number): AffineMatrix {
  return { a: x, b: 0, c: 0, d: y, e: 0, f: 0 };
}

function skewXAffine(degrees: number): AffineMatrix {
  return { a: 1, b: 0, c: Math.tan((degrees * Math.PI) / 180), d: 1, e: 0, f: 0 };
}

function skewYAffine(degrees: number): AffineMatrix {
  return { a: 1, b: Math.tan((degrees * Math.PI) / 180), c: 0, d: 1, e: 0, f: 0 };
}

function layerAffineMatrix(transform: CustomSymbolVectorTransform): AffineMatrix {
  const tx = transform.translateX ?? 0;
  const ty = transform.translateY ?? 0;
  const ox = transform.originX ?? 0;
  const oy = transform.originY ?? 0;
  const matrices = [
    translationAffine(tx, ty),
    translationAffine(ox, oy),
    rotationAffine(transform.rotateDeg ?? 0),
    skewXAffine(transform.skewXDeg ?? 0),
    skewYAffine(transform.skewYDeg ?? 0),
    scaleAffine(transform.scaleX ?? 1, transform.scaleY ?? 1),
    translationAffine(-ox, -oy),
  ];
  return matrices.reduce(multiplyAffine, {
    a: 1,
    b: 0,
    c: 0,
    d: 1,
    e: 0,
    f: 0,
  });
}

function affinePoint(matrix: AffineMatrix, x: number, y: number) {
  return {
    x: matrix.a * x + matrix.c * y + matrix.e,
    y: matrix.b * x + matrix.d * y + matrix.f,
  };
}

function normalizedAngleDelta(value: number) {
  let result = value;
  while (result > 180) result -= 360;
  while (result < -180) result += 360;
  return result;
}

const rightAngleSnapThresholdDeg = 7;

function snapRotationToRightAngle(value: number) {
  const nearestRightAngle = Math.round(value / 90) * 90;
  return Math.abs(value - nearestRightAngle) <= rightAngleSnapThresholdDeg
    ? nearestRightAngle
    : value;
}

function Shape({
  shape,
  clipId,
  paint = "currentColor",
}: {
  shape: CustomSymbolVectorShape;
  clipId: string;
  paint?: string;
}) {
  const strokeWidth = shape.strokeWidth ?? (shape.kind === "line" ? 50 : 0);
  const fillEnabled = shape.fill ?? shape.kind !== "line";
  const common = {
    transform: svgTransform(shape.transform),
    fill: fillEnabled ? paint : "none",
    stroke: strokeWidth > 0 ? paint : "none",
    strokeWidth,
    strokeLinecap: shape.lineCap,
    strokeLinejoin: shape.lineJoin,
  } as const;
  let geometry: ReactNode;
  switch (shape.kind) {
    case "path":
      geometry = <path d={shape.d} {...common} />;
      break;
    case "circle":
      geometry = <circle cx={shape.cx} cy={shape.cy} r={shape.r} {...common} />;
      break;
    case "ellipse":
      geometry = <ellipse cx={shape.cx} cy={shape.cy} rx={shape.rx} ry={shape.ry} {...common} />;
      break;
    case "line":
      geometry = <line x1={shape.x1} y1={shape.y1} x2={shape.x2} y2={shape.y2} {...common} />;
      break;
    case "rect":
      geometry = (
        <rect
          x={shape.x}
          y={shape.y}
          width={shape.width}
          height={shape.height}
          rx={shape.rx}
          ry={shape.ry}
          {...common}
        />
      );
      break;
    case "polygon":
      geometry = (
        <polygon
          points={shape.points.map(([x, y]) => `${x},${y}`).join(" ")}
          {...common}
        />
      );
      break;
    case "text":
      geometry = (
        <text
          x={shape.x}
          y={shape.y}
          fontFamily={shape.fontFamily}
          fontSize={shape.fontSize}
          fontStyle={shape.fontStyle}
          fontWeight={shape.fontWeight}
          {...common}
        >
          {shape.text}
        </text>
      );
      break;
  }
  if (!shape.clipRect) return geometry;
  return (
    <>
      <defs>
        <clipPath id={clipId} clipPathUnits="userSpaceOnUse">
          <rect {...shape.clipRect} />
        </clipPath>
      </defs>
      <g clipPath={`url(#${clipId})`}>{geometry}</g>
    </>
  );
}

function layerBounds(layer: CustomSymbolDesignerLayer) {
  if (layer.kind === "glyph") {
    return (
      layer.clipRect ?? {
        x: 0,
        y: 0,
        width: layer.asset.metrics.widthEm * 1000,
        height:
          (layer.asset.metrics.ascentEm + layer.asset.metrics.descentEm) * 1000,
      }
    );
  }
  return layer.bounds;
}

function paddedViewBox(x: number, y: number, width: number, height: number): ViewBoxState {
  const normalizedWidth = Math.max(40, width);
  const normalizedHeight = Math.max(40, height);
  const padding = Math.max(
    650,
    Math.min(1_800, Math.max(normalizedWidth, normalizedHeight) * 0.65),
  );
  return {
    x: x - padding,
    y: y - padding,
    width: normalizedWidth + padding * 2,
    height: normalizedHeight + padding * 2,
  };
}

function resizeGeometry(
  handle: ResizeHandle,
  bounds: { x: number; y: number; width: number; height: number },
) {
  const left = bounds.x;
  const top = bounds.y;
  const right = bounds.x + bounds.width;
  const bottom = bounds.y + bounds.height;
  const centerX = (left + right) / 2;
  const centerY = (top + bottom) / 2;
  switch (handle) {
    case "nw":
      return { handleX: left, handleY: top, anchorX: right, anchorY: bottom };
    case "n":
      return { handleX: centerX, handleY: top, anchorX: centerX, anchorY: bottom };
    case "ne":
      return { handleX: right, handleY: top, anchorX: left, anchorY: bottom };
    case "e":
      return { handleX: right, handleY: centerY, anchorX: left, anchorY: centerY };
    case "se":
      return { handleX: right, handleY: bottom, anchorX: left, anchorY: top };
    case "s":
      return { handleX: centerX, handleY: bottom, anchorX: centerX, anchorY: top };
    case "sw":
      return { handleX: left, handleY: bottom, anchorX: right, anchorY: top };
    case "w":
      return { handleX: left, handleY: centerY, anchorX: right, anchorY: centerY };
  }
}

function pointWithMatrix(clientX: number, clientY: number, matrix: DOMMatrix) {
  return new DOMPoint(clientX, clientY).matrixTransform(matrix);
}

function eraseLayer(layer: CustomSymbolDesignerLayer) {
  return layer.kind === "geometry" && layer.shape.operation === "erase";
}

export function CustomSymbolDesignerCanvas({
  documentState,
  selectedLayerId,
  referenceAsset,
  showReference,
  referenceLabel,
  eraserMode,
  eraserSize,
  isEn,
  onSelectLayer,
  onMoveLayer,
  onResizeLayer,
  onRotateLayer,
  onDeleteLayer,
  onAddEraserStroke,
}: Props) {
  const svgRef = useRef<SVGSVGElement>(null);
  const contentRef = useRef<SVGGElement>(null);
  const fitWidthRef = useRef(1);
  const [drag, setDrag] = useState<DragState | null>(null);
  const [pan, setPan] = useState<PanState | null>(null);
  const [resize, setResize] = useState<ResizeState | null>(null);
  const [rotate, setRotate] = useState<RotateState | null>(null);
  const [erase, setErase] = useState<EraseState | null>(null);
  const [eraserCursor, setEraserCursor] = useState<{ x: number; y: number } | null>(null);
  const [inkBounds, setInkBounds] = useState<{
    x: number;
    y: number;
    width: number;
    height: number;
  } | null>(null);
  const width = Math.max(20, documentState.metrics.widthEm * 1000);
  const height = Math.max(
    20,
    (documentState.metrics.ascentEm + documentState.metrics.descentEm) * 1000,
  );
  const baseline = documentState.metrics.ascentEm * 1000;
  const workspace = useMemo(
    () => ({
      x: -infiniteCanvasExtent,
      y: -infiniteCanvasExtent,
      width: infiniteCanvasExtent * 2,
      height: infiniteCanvasExtent * 2,
    }),
    [],
  );
  const initialViewBox = useMemo(() => paddedViewBox(0, 0, width, height), [width, height]);
  const [viewBox, setViewBox] = useState<ViewBoxState>(initialViewBox);
  const paintLayers = documentState.layers.filter((layer) => layer.visible && !eraseLayer(layer));
  const eraseLayers = documentState.layers.filter((layer) => layer.visible && eraseLayer(layer));
  const eraseMaskId = "visualtex-custom-symbol-designer-erase-mask";

  const measureInkBounds = useCallback(() => {
    const content = contentRef.current;
    if (!content) return null;
    try {
      const graphics = content as SVGGElement & {
        getBBox(options?: {
          fill?: boolean;
          stroke?: boolean;
          markers?: boolean;
          clipped?: boolean;
        }): DOMRect;
      };
      let bounds: DOMRect;
      try {
        bounds = graphics.getBBox({
          fill: true,
          stroke: true,
          markers: true,
          clipped: true,
        });
      } catch {
        bounds = content.getBBox();
      }
      if (
        !Number.isFinite(bounds.x) ||
        !Number.isFinite(bounds.y) ||
        !Number.isFinite(bounds.width) ||
        !Number.isFinite(bounds.height) ||
        bounds.width <= 0 ||
        bounds.height <= 0
      ) {
        return null;
      }
      return {
        x: bounds.x,
        y: bounds.y,
        width: bounds.width,
        height: bounds.height,
      };
    } catch {
      return null;
    }
  }, []);

  const fitViewport = useCallback(() => {
    const bounds = measureInkBounds();
    const next = bounds
      ? paddedViewBox(bounds.x, bounds.y, bounds.width, bounds.height)
      : initialViewBox;
    setInkBounds(bounds);
    fitWidthRef.current = next.width;
    setViewBox(next);
  }, [initialViewBox, measureInkBounds]);

  useEffect(() => {
    // Reframe only when switching to a different registered symbol. Inserting
    // or resizing layers must never move the user's viewport underneath them.
    const frame = requestAnimationFrame(() => {
      const bounds = measureInkBounds();
      const next = bounds
        ? paddedViewBox(bounds.x, bounds.y, bounds.width, bounds.height)
        : initialViewBox;
      setInkBounds(bounds);
      fitWidthRef.current = next.width;
      setViewBox(next);
    });
    return () => cancelAnimationFrame(frame);
    // Intentionally keyed only by document identity; evolving layer metrics
    // are measured without implicitly reframing the viewport.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [documentState.symbolId]);

  useEffect(() => {
    const frame = requestAnimationFrame(() => {
      setInkBounds(measureInkBounds());
    });
    return () => cancelAnimationFrame(frame);
  }, [documentState.layers, measureInkBounds]);

  const localPoint = (
    event: ReactPointerEvent<SVGSVGElement> | ReactWheelEvent<SVGSVGElement>,
  ) => {
    const svg = svgRef.current;
    const matrix = svg?.getScreenCTM();
    if (!svg || !matrix) return null;
    const point = svg.createSVGPoint();
    point.x = event.clientX;
    point.y = event.clientY;
    return point.matrixTransform(matrix.inverse());
  };

  const zoomAround = (factor: number, center?: { x: number; y: number }) => {
    setViewBox((current) => {
      const nextWidth = Math.min(
        Math.max(current.width * factor, 40),
        infiniteCanvasExtent * 2,
      );
      const appliedFactor = nextWidth / current.width;
      const nextHeight = current.height * appliedFactor;
      const pivot = center ?? {
        x: current.x + current.width / 2,
        y: current.y + current.height / 2,
      };
      return {
        x: pivot.x - (pivot.x - current.x) * appliedFactor,
        y: pivot.y - (pivot.y - current.y) * appliedFactor,
        width: nextWidth,
        height: nextHeight,
      };
    });
  };

  const beginDrag = (
    event: ReactPointerEvent<SVGGElement>,
    layer: CustomSymbolDesignerLayer,
  ) => {
    if (eraserMode) return;
    onSelectLayer(layer.id);
    if (layer.locked) return;
    const svg = svgRef.current;
    const matrix = svg?.getScreenCTM();
    if (!svg || !matrix) return;
    svg.focus({ preventScroll: true });
    const start = pointWithMatrix(event.clientX, event.clientY, matrix.inverse());
    svg.setPointerCapture?.(event.pointerId);
    setDrag({
      pointerId: event.pointerId,
      layerId: layer.id,
      startX: start.x,
      startY: start.y,
      translateX: layer.transform.translateX ?? 0,
      translateY: layer.transform.translateY ?? 0,
    });
    event.preventDefault();
    event.stopPropagation();
  };

  const beginResize = (
    event: ReactPointerEvent<SVGRectElement>,
    layer: CustomSymbolDesignerLayer,
    handle: ResizeHandle,
  ) => {
    if (layer.locked || eraserMode) return;
    const group = event.currentTarget.closest(
      ".custom-symbol-designer-layer",
    ) as SVGGElement | null;
    const matrix = group?.getScreenCTM();
    const svg = svgRef.current;
    if (!group || !matrix || !svg) return;
    const bounds = layerBounds(layer);
    const geometry = resizeGeometry(handle, bounds);
    onSelectLayer(layer.id);
    svg.setPointerCapture?.(event.pointerId);
    setResize({
      pointerId: event.pointerId,
      layerId: layer.id,
      handle,
      inverseMatrix: matrix.inverse(),
      bounds,
      anchorX: geometry.anchorX,
      anchorY: geometry.anchorY,
      originalTransform: { ...layer.transform },
    });
    event.preventDefault();
    event.stopPropagation();
  };

  const beginResizeFromHitTarget = (
    event: ReactPointerEvent<SVGRectElement>,
    layer: CustomSymbolDesignerLayer,
  ) => {
    if (layer.locked || eraserMode) return;
    const group = event.currentTarget.closest(
      ".custom-symbol-designer-layer",
    ) as SVGGElement | null;
    const matrix = group?.getScreenCTM();
    if (!group || !matrix) return;

    const bounds = layerBounds(layer);
    let nearestHandle: ResizeHandle | null = null;
    let nearestDistance = Number.POSITIVE_INFINITY;
    for (const handle of resizeHandles) {
      const geometry = resizeGeometry(handle, bounds);
      const screenPoint = new DOMPoint(
        geometry.handleX,
        geometry.handleY,
      ).matrixTransform(matrix);
      const distance = Math.hypot(
        event.clientX - screenPoint.x,
        event.clientY - screenPoint.y,
      );
      if (distance < nearestDistance) {
        nearestDistance = distance;
        nearestHandle = handle;
      }
    }
    if (nearestHandle) beginResize(event, layer, nearestHandle);
  };

  const beginRotate = (
    event: ReactPointerEvent<SVGCircleElement>,
    layer: CustomSymbolDesignerLayer,
  ) => {
    if (layer.locked || eraserMode) return;
    const group = event.currentTarget.closest(
      ".custom-symbol-designer-layer",
    ) as SVGGElement | null;
    const matrix = group?.getScreenCTM();
    const svg = svgRef.current;
    if (!group || !matrix || !svg) return;
    const bounds = layerBounds(layer);
    const center = new DOMPoint(
      bounds.x + bounds.width / 2,
      bounds.y + bounds.height / 2,
    ).matrixTransform(matrix);
    onSelectLayer(layer.id);
    svg.setPointerCapture?.(event.pointerId);
    setRotate({
      pointerId: event.pointerId,
      layerId: layer.id,
      centerClientX: center.x,
      centerClientY: center.y,
      startPointerAngle: Math.atan2(
        event.clientY - center.y,
        event.clientX - center.x,
      ),
      originalRotateDeg: layer.transform.rotateDeg ?? 0,
    });
    event.preventDefault();
    event.stopPropagation();
  };

  const beginCanvasPointer = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (event.button !== 0 && event.button !== 1) return;
    event.currentTarget.focus({ preventScroll: true });
    if (eraserMode && event.button === 0) {
      const point = localPoint(event);
      if (!point) return;
      setEraserCursor({ x: point.x, y: point.y });
      event.currentTarget.setPointerCapture?.(event.pointerId);
      setErase({
        pointerId: event.pointerId,
        points: [{ x: point.x, y: point.y }],
      });
      event.preventDefault();
      return;
    }
    const target = event.target as Element;
    if (target.closest?.(".custom-symbol-designer-layer")) return;
    event.currentTarget.setPointerCapture?.(event.pointerId);
    setPan({
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startViewBox: viewBox,
      moved: false,
    });
    event.preventDefault();
  };

  const handleCanvasKeyDown = (event: ReactKeyboardEvent<SVGSVGElement>) => {
    if (
      !selectedLayerId ||
      (event.key !== "Backspace" && event.key !== "Delete") ||
      event.metaKey ||
      event.ctrlKey ||
      event.altKey
    ) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    onDeleteLayer(selectedLayerId);
  };

  const movePointer = (event: ReactPointerEvent<SVGSVGElement>) => {
    const pointerPoint = eraserMode ? localPoint(event) : null;
    if (eraserMode && pointerPoint) {
      setEraserCursor({ x: pointerPoint.x, y: pointerPoint.y });
    }
    if (erase && event.pointerId === erase.pointerId) {
      const point = pointerPoint ?? localPoint(event);
      if (!point) return;
      const svg = svgRef.current;
      const minimumDistance = Math.max(
        0.5,
        (viewBox.width / Math.max(svg?.clientWidth ?? 1, 1)) * 1.25,
      );
      setErase((current) => {
        if (!current) return current;
        const last = current.points[current.points.length - 1];
        if (last && Math.hypot(point.x - last.x, point.y - last.y) < minimumDistance) {
          return current;
        }
        return {
          ...current,
          points: [...current.points, { x: point.x, y: point.y }],
        };
      });
      return;
    }

    if (rotate && event.pointerId === rotate.pointerId) {
      const pointerAngle = Math.atan2(
        event.clientY - rotate.centerClientY,
        event.clientX - rotate.centerClientX,
      );
      const deltaDeg = normalizedAngleDelta(
        ((pointerAngle - rotate.startPointerAngle) * 180) / Math.PI,
      );
      let nextRotation = rotate.originalRotateDeg + deltaDeg;
      nextRotation = event.shiftKey
        ? Math.round(nextRotation / 15) * 15
        : snapRotationToRightAngle(nextRotation);
      onRotateLayer(rotate.layerId, nextRotation);
      return;
    }

    if (resize && event.pointerId === resize.pointerId) {
      const point = pointWithMatrix(event.clientX, event.clientY, resize.inverseMatrix);
      const geometry = resizeGeometry(resize.handle, resize.bounds);
      const denominatorX = geometry.handleX - geometry.anchorX;
      const denominatorY = geometry.handleY - geometry.anchorY;
      let factorX = denominatorX ? (point.x - geometry.anchorX) / denominatorX : 1;
      let factorY = denominatorY ? (point.y - geometry.anchorY) / denominatorY : 1;
      const originalScaleX = resize.originalTransform.scaleX ?? 1;
      const originalScaleY = resize.originalTransform.scaleY ?? 1;
      factorX = Math.max(
        0.02 / Math.max(Math.abs(originalScaleX), 0.02),
        factorX,
      );
      factorY = Math.max(
        0.02 / Math.max(Math.abs(originalScaleY), 0.02),
        factorY,
      );
      const horizontalOnly = resize.handle === "e" || resize.handle === "w";
      const verticalOnly = resize.handle === "n" || resize.handle === "s";
      const cornerHandle = !horizontalOnly && !verticalOnly;
      if (cornerHandle && !event.shiftKey) {
        const uniform =
          Math.abs(factorX - 1) >= Math.abs(factorY - 1) ? factorX : factorY;
        factorX = uniform;
        factorY = uniform;
      }
      const nextScaleX = verticalOnly
        ? originalScaleX
        : Math.sign(originalScaleX || 1) *
          Math.max(0.02, Math.abs(originalScaleX) * factorX);
      const nextScaleY = horizontalOnly
        ? originalScaleY
        : Math.sign(originalScaleY || 1) *
          Math.max(0.02, Math.abs(originalScaleY) * factorY);
      const oldAnchor = affinePoint(
        layerAffineMatrix(resize.originalTransform),
        resize.anchorX,
        resize.anchorY,
      );
      const nextTransform = {
        ...resize.originalTransform,
        scaleX: nextScaleX,
        scaleY: nextScaleY,
      };
      const shiftedAnchor = affinePoint(
        layerAffineMatrix(nextTransform),
        resize.anchorX,
        resize.anchorY,
      );
      const nextTranslateX =
        (resize.originalTransform.translateX ?? 0) +
        oldAnchor.x -
        shiftedAnchor.x;
      const nextTranslateY =
        (resize.originalTransform.translateY ?? 0) +
        oldAnchor.y -
        shiftedAnchor.y;
      onResizeLayer(
        resize.layerId,
        nextScaleX,
        nextScaleY,
        nextTranslateX,
        nextTranslateY,
      );
      return;
    }

    if (drag && event.pointerId === drag.pointerId) {
      const point = localPoint(event);
      if (!point) return;
      onMoveLayer(
        drag.layerId,
        drag.translateX + point.x - drag.startX,
        drag.translateY + point.y - drag.startY,
      );
      return;
    }

    if (pan && event.pointerId === pan.pointerId) {
      const svg = svgRef.current;
      if (!svg) return;
      const dx = event.clientX - pan.startClientX;
      const dy = event.clientY - pan.startClientY;
      const unitX = pan.startViewBox.width / Math.max(svg.clientWidth, 1);
      const unitY = pan.startViewBox.height / Math.max(svg.clientHeight, 1);
      if (!pan.moved && Math.hypot(dx, dy) > 3) {
        setPan((current) => (current ? { ...current, moved: true } : current));
      }
      setViewBox({
        ...pan.startViewBox,
        x: pan.startViewBox.x - dx * unitX,
        y: pan.startViewBox.y - dy * unitY,
      });
    }
  };

  const endPointer = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (erase?.pointerId === event.pointerId) {
      if (erase.points.length >= 1) onAddEraserStroke(erase.points);
      setErase(null);
    }
    if (resize?.pointerId === event.pointerId) setResize(null);
    if (rotate?.pointerId === event.pointerId) setRotate(null);
    if (drag?.pointerId === event.pointerId) setDrag(null);
    if (pan?.pointerId === event.pointerId) {
      if (!pan.moved) onSelectLayer(null);
      setPan(null);
    }
    event.currentTarget.releasePointerCapture?.(event.pointerId);
  };

  const worldPerPixel =
    viewBox.width / Math.max(svgRef.current?.clientWidth ?? 800, 1);
  const handleSizeWorld = Math.max(7, worldPerPixel * 8);
  const rotationDistanceWorld = Math.max(handleSizeWorld * 3.6, worldPerPixel * 34);
  const zoomPercent = Math.max(
    8,
    Math.min(1200, Math.round((fitWidthRef.current / Math.max(viewBox.width, 1)) * 100)),
  );
  const referenceWidth = (referenceAsset?.metrics.widthEm ?? 0) * 1000;
  const referenceAscent = (referenceAsset?.metrics.ascentEm ?? 0) * 1000;
  const referenceX = (width - referenceWidth) / 2;
  const referenceY = baseline - referenceAscent;
  const resizeHandles: ResizeHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
  const liveErasePath = erase ? buildSmoothCustomSymbolEraserPath(erase.points) : "";
  const backgroundBounds = {
    x: viewBox.x - viewBox.width,
    y: viewBox.y - viewBox.height,
    width: viewBox.width * 3,
    height: viewBox.height * 3,
  };
  const inkSizeLabel = inkBounds
    ? `${(inkBounds.width / 1000).toFixed(2)} × ${(inkBounds.height / 1000).toFixed(2)} em`
    : "—";

  const renderArtworkLayer = (layer: CustomSymbolDesignerLayer) => {
    const shapes = customSymbolDesignerLayerLocalShapes(layer);
    return (
      <g
        key={`artwork-${layer.id}`}
        transform={svgTransform(layer.transform)}
        data-custom-symbol-artwork-layer={layer.id}
        data-custom-symbol-artwork-kind={layer.kind}
        data-custom-symbol-artwork-preset={
          layer.kind === "geometry" ? layer.geometryPreset ?? "" : ""
        }
        data-custom-symbol-artwork-operation={eraseLayer(layer) ? "erase" : "paint"}
      >
        {shapes.map((shape, index) => (
          <Shape
            key={`${layer.id}-${index}`}
            shape={{
              ...shape,
              ...(layer.clipRect ? { clipRect: layer.clipRect } : {}),
            }}
            clipId={`designer-artwork-${layer.id}-${index}`}
          />
        ))}
      </g>
    );
  };

  const renderInteractionLayer = (layer: CustomSymbolDesignerLayer) => {
    const bounds = layerBounds(layer);
    const selected = selectedLayerId === layer.id;
    const scaleX = Math.max(Math.abs(layer.transform.scaleX ?? 1), 0.02);
    const scaleY = Math.max(Math.abs(layer.transform.scaleY ?? 1), 0.02);
    const handleWidth = handleSizeWorld / scaleX;
    const handleHeight = handleSizeWorld / scaleY;
    const screenWidth = Math.max(1, (bounds.width * scaleX) / worldPerPixel);
    const screenHeight = Math.max(1, (bounds.height * scaleY) / worldPerPixel);
    const shortestScreenSide = Math.min(screenWidth, screenHeight);
    const resizeHitPixels = Math.min(
      14,
      Math.max(9, shortestScreenSide * 0.24),
    );
    const resizeHitSizeWorld = worldPerPixel * resizeHitPixels;
    const hitWidth = resizeHitSizeWorld / scaleX;
    const hitHeight = resizeHitSizeWorld / scaleY;
    const averageScale = Math.sqrt(scaleX * scaleY);
    const rotationAnchorX = bounds.x + bounds.width / 2;
    const verticallyFlipped = (layer.transform.scaleY ?? 1) < 0;
    const rotationAnchorY = verticallyFlipped
      ? bounds.y + bounds.height
      : bounds.y;
    const rotationDirection = verticallyFlipped ? 1 : -1;
    const rotationHandleY =
      rotationAnchorY +
      rotationDirection * (rotationDistanceWorld / scaleY);
    const rotationRadius = Math.max(4, handleSizeWorld * 0.58) / averageScale;
    const rotationHitRadius = Math.max(rotationRadius * 3, resizeHitSizeWorld / 2 / averageScale);
    const isErase = eraseLayer(layer);
    return (
      <g
        key={`interaction-${layer.id}`}
        transform={svgTransform(layer.transform)}
        className={`custom-symbol-designer-layer${selected ? " is-selected" : ""}${layer.locked ? " is-locked" : ""}${isErase ? " is-eraser-layer" : ""}`}
        data-custom-symbol-canvas-layer={layer.id}
        data-custom-symbol-layer-kind={layer.kind}
        data-custom-symbol-layer-preset={
          layer.kind === "geometry" ? layer.geometryPreset ?? "" : ""
        }
        data-custom-symbol-layer-operation={isErase ? "erase" : "paint"}
        onPointerDown={(event) => beginDrag(event, layer)}
      >
        <rect
          x={bounds.x}
          y={bounds.y}
          width={bounds.width}
          height={bounds.height}
          fill="transparent"
          pointerEvents={eraserMode ? "none" : "all"}
        />
        {selected && !eraserMode && isErase && layer.kind === "geometry" && layer.shape.kind === "path" ? (
          <path
            d={layer.shape.d}
            className="custom-symbol-designer-eraser-centerline"
            fill="none"
            pointerEvents="none"
            data-custom-symbol-eraser-centerline
          />
        ) : null}
        {selected && !eraserMode && !isErase ? (
          <>
            <rect
              x={bounds.x}
              y={bounds.y}
              width={bounds.width}
              height={bounds.height}
              className="custom-symbol-designer-selection-box"
              pointerEvents="none"
            />
            {!layer.locked ? (
              <>
                <g className="custom-symbol-designer-resize-handles" data-custom-symbol-resize-handles>
                  {resizeHandles.map((handle) => {
                    const geometry = resizeGeometry(handle, bounds);
                    return (
                      <g key={handle}>
                        <rect
                          x={geometry.handleX - hitWidth / 2}
                          y={geometry.handleY - hitHeight / 2}
                          width={hitWidth}
                          height={hitHeight}
                          className={`custom-symbol-designer-resize-hit-target is-${handle}`}
                          data-custom-symbol-resize-hit-target={handle}
                          onPointerDown={(event) => beginResizeFromHitTarget(event, layer)}
                        />
                        <rect
                          x={geometry.handleX - handleWidth / 2}
                          y={geometry.handleY - handleHeight / 2}
                          width={handleWidth}
                          height={handleHeight}
                          rx={Math.min(handleWidth, handleHeight) * 0.12}
                          className={`custom-symbol-designer-resize-handle is-${handle}`}
                          data-custom-symbol-resize-handle={handle}
                          pointerEvents="none"
                        />
                      </g>
                    );
                  })}
                </g>
                <g
                  className="custom-symbol-designer-rotation-control"
                  data-custom-symbol-rotation-control
                >
                  <line
                    x1={rotationAnchorX}
                    y1={rotationAnchorY}
                    x2={rotationAnchorX}
                    y2={rotationHandleY}
                    className="custom-symbol-designer-rotation-stem"
                    pointerEvents="none"
                  />
                  <circle
                    cx={rotationAnchorX}
                    cy={rotationHandleY}
                    r={rotationHitRadius}
                    className="custom-symbol-designer-rotation-hit-target"
                    data-custom-symbol-rotation-hit-target
                    onPointerDown={(event) => beginRotate(event, layer)}
                  />
                  <circle
                    cx={rotationAnchorX}
                    cy={rotationHandleY}
                    r={rotationRadius}
                    className="custom-symbol-designer-rotation-handle"
                    data-custom-symbol-rotation-handle
                    pointerEvents="none"
                  />
                </g>
              </>
            ) : null}
          </>
        ) : null}
      </g>
    );
  };

  return (
    <div
      className="custom-symbol-designer-canvas-shell"
      data-custom-symbol-canvas-shell
      data-custom-symbol-infinite-canvas="true"
      data-custom-symbol-auto-ink-bounds="true"
    >
      <div className="custom-symbol-designer-viewport-controls" aria-label={isEn ? "Canvas zoom" : "画布缩放"}>
        <span
          className="custom-symbol-designer-ink-size"
          data-custom-symbol-ink-size
          title={isEn ? "Ink bounds" : "墨迹尺寸"}
        >
          {inkSizeLabel}
        </span>
        <button type="button" data-custom-symbol-zoom-out onClick={() => zoomAround(1.18)} title={isEn ? "Zoom out" : "缩小"}>−</button>
        <span data-custom-symbol-zoom-percent>{zoomPercent}%</span>
        <button type="button" data-custom-symbol-fit-view onClick={fitViewport}>{isEn ? "Locate" : "定位"}</button>
        <button type="button" data-custom-symbol-zoom-in onClick={() => zoomAround(0.84)} title={isEn ? "Zoom in" : "放大"}>+</button>
      </div>
      <svg
        ref={svgRef}
        className={`custom-symbol-designer-canvas${pan?.moved ? " is-panning" : ""}${rotate ? " is-rotating" : ""}${eraserMode ? " is-erasing" : ""}`}
        viewBox={`${viewBox.x} ${viewBox.y} ${viewBox.width} ${viewBox.height}`}
        preserveAspectRatio="xMidYMid meet"
        data-custom-symbol-canvas
        data-custom-symbol-infinite-canvas="true"
        data-custom-symbol-viewbox={`${viewBox.x} ${viewBox.y} ${viewBox.width} ${viewBox.height}`}
        data-custom-symbol-workspace={`${workspace.x} ${workspace.y} ${workspace.width} ${workspace.height}`}
        tabIndex={-1}
        onKeyDown={handleCanvasKeyDown}
        onWheel={(event) => {
          event.preventDefault();
          const point = localPoint(event);
          zoomAround(event.deltaY > 0 ? 1.12 : 0.89, point ?? undefined);
        }}
        onPointerMove={movePointer}
        onPointerUp={endPointer}
        onPointerCancel={endPointer}
        onPointerDown={beginCanvasPointer}
        onPointerLeave={() => {
          if (!erase) setEraserCursor(null);
        }}
      >
        <defs>
          <filter id="visualtex-custom-symbol-grid-soften" x="-20%" y="-20%" width="140%" height="140%">
            <feGaussianBlur stdDeviation="0.32" />
          </filter>
          <pattern id="visualtex-custom-symbol-grid-minor" width="80" height="80" patternUnits="userSpaceOnUse">
            <path
              d="M80 0H0V80"
              className="custom-symbol-designer-grid-line is-minor"
              filter="url(#visualtex-custom-symbol-grid-soften)"
            />
          </pattern>
          <pattern id="visualtex-custom-symbol-grid-major" width="400" height="400" patternUnits="userSpaceOnUse">
            <path
              d="M400 0H0V400"
              className="custom-symbol-designer-grid-line is-major"
              filter="url(#visualtex-custom-symbol-grid-soften)"
            />
          </pattern>
          {(eraseLayers.length || liveErasePath) ? (
            <mask
              id={eraseMaskId}
              maskUnits="userSpaceOnUse"
              x={workspace.x}
              y={workspace.y}
              width={workspace.width}
              height={workspace.height}
            >
              <rect
                x={workspace.x}
                y={workspace.y}
                width={workspace.width}
                height={workspace.height}
                fill="white"
              />
              {eraseLayers.map((layer) =>
                layer.kind === "geometry" ? (
                  <g key={`erase-mask-${layer.id}`} transform={svgTransform(layer.transform)}>
                    <Shape
                      shape={layer.shape}
                      clipId={`designer-erase-mask-${layer.id}`}
                      paint="black"
                    />
                  </g>
                ) : null,
              )}
              {liveErasePath ? (
                <path
                  d={liveErasePath}
                  fill="none"
                  stroke="black"
                  strokeWidth={eraserSize}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              ) : null}
            </mask>
          ) : null}
        </defs>

        <rect
          x={backgroundBounds.x}
          y={backgroundBounds.y}
          width={backgroundBounds.width}
          height={backgroundBounds.height}
          className="custom-symbol-designer-workspace"
          data-custom-symbol-workspace-paper
        />
        <rect
          x={backgroundBounds.x}
          y={backgroundBounds.y}
          width={backgroundBounds.width}
          height={backgroundBounds.height}
          fill="url(#visualtex-custom-symbol-grid-minor)"
          className="custom-symbol-designer-grid-layer is-minor"
          pointerEvents="none"
        />
        <rect
          x={backgroundBounds.x}
          y={backgroundBounds.y}
          width={backgroundBounds.width}
          height={backgroundBounds.height}
          fill="url(#visualtex-custom-symbol-grid-major)"
          className="custom-symbol-designer-grid-layer is-major"
          pointerEvents="none"
        />
        <line
          x1={backgroundBounds.x}
          y1={baseline}
          x2={backgroundBounds.x + backgroundBounds.width}
          y2={baseline}
          className="custom-symbol-designer-baseline"
          data-custom-symbol-baseline
        />

        {showReference && referenceAsset ? (
            <g
              className="custom-symbol-designer-reference-alpha custom-symbol-designer-reference-glyph"
              transform={`translate(${referenceX} ${referenceY})`}
              data-custom-symbol-reference
              data-custom-symbol-reference-alpha={referenceLabel === "α" ? "true" : undefined}
              data-custom-symbol-reference-label={referenceLabel}
              pointerEvents="none"
            >
              {referenceAsset.shapes.map((shape, index) => (
                <Shape
                  key={`reference-${index}`}
                  shape={shape}
                  clipId={`designer-reference-${index}`}
                />
              ))}
            </g>
          ) : null}

        <g
          ref={contentRef}
          data-custom-symbol-canvas-content
          data-custom-symbol-ink-content
          mask={(eraseLayers.length || liveErasePath) ? `url(#${eraseMaskId})` : undefined}
        >
          {paintLayers.map(renderArtworkLayer)}
        </g>
        {paintLayers.map(renderInteractionLayer)}
        {eraseLayers.map(renderInteractionLayer)}

        {eraserMode && eraserCursor ? (
          <g className="custom-symbol-designer-eraser-cursor" pointerEvents="none" data-custom-symbol-eraser-cursor>
            <circle
              cx={eraserCursor.x}
              cy={eraserCursor.y}
              r={Math.max(2, eraserSize / 2)}
            />
            <line
              x1={eraserCursor.x - Math.min(eraserSize * 0.16, 12)}
              y1={eraserCursor.y}
              x2={eraserCursor.x + Math.min(eraserSize * 0.16, 12)}
              y2={eraserCursor.y}
            />
            <line
              x1={eraserCursor.x}
              y1={eraserCursor.y - Math.min(eraserSize * 0.16, 12)}
              x2={eraserCursor.x}
              y2={eraserCursor.y + Math.min(eraserSize * 0.16, 12)}
            />
          </g>
        ) : null}
      </svg>
    </div>
  );
}
