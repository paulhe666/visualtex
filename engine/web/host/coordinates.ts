// The SVG view reports CSS client pixels; the core owns natural layout pixels.
// Scroll offsets are actual CSS pixels measured at the viewport, not document
// em units. devicePixelRatio is deliberately absent from this mapping.
export interface Point { readonly x: number; readonly y: number }
export interface Rect extends Point { readonly width: number; readonly height: number }
export interface ViewportTransform {
  readonly left: number;
  readonly top: number;
  readonly scrollLeft: number;
  readonly scrollTop: number;
  readonly zoom: number;
}
export interface GeometryIdentity {
  readonly documentId: string;
  readonly revision: number;
  readonly layoutGeneration: number;
}
export class GeometryError extends Error {
  constructor(readonly code: 'invalidCoordinate' | 'staleGeometry', message: string) {
    super(message);
    this.name = 'GeometryError';
  }
}
function finite(...values: number[]): void {
  if (!values.every(Number.isFinite)) {
    throw new GeometryError('invalidCoordinate', 'Coordinates must be finite CSS/layout values');
  }
}
function checkedPoint(x: number, y: number): Point {
  finite(x, y);
  return { x, y };
}
function checkedRect(rect: Rect): Rect {
  finite(rect.x, rect.y, rect.width, rect.height);
  if (rect.width < 0 || rect.height < 0) {
    throw new GeometryError('invalidCoordinate', 'Rectangle dimensions cannot be negative');
  }
  return rect;
}

/** An immutable mapping; a scroll/zoom change creates a new mapping, not a new
 * document revision. No DOM measurements are used to infer mathematical layout. */
export class ViewportMapping {
  private readonly view: ViewportTransform;
  constructor(view: ViewportTransform) {
    finite(view.left, view.top, view.scrollLeft, view.scrollTop, view.zoom);
    if (view.zoom <= 0) {
      throw new GeometryError('invalidCoordinate', 'Zoom must be positive');
    }
    this.view = Object.freeze({ ...view });
  }
  clientToLayout(point: Point): Point {
    finite(point.x, point.y);
    const v = this.view;
    return checkedPoint((point.x - v.left + v.scrollLeft) / v.zoom,
      (point.y - v.top + v.scrollTop) / v.zoom);
  }
  layoutToClient(point: Point): Point {
    finite(point.x, point.y);
    const v = this.view;
    return checkedPoint(point.x * v.zoom + v.left - v.scrollLeft,
      point.y * v.zoom + v.top - v.scrollTop);
  }
  layoutRectToClient(rect: Rect): Rect {
    checkedRect(rect);
    const origin = this.layoutToClient(rect);
    return checkedRect({ ...origin, width: rect.width * this.view.zoom, height: rect.height * this.view.zoom });
  }
  clientRectToLayout(rect: Rect): Rect {
    checkedRect(rect);
    const origin = this.clientToLayout(rect);
    return checkedRect({ ...origin, width: rect.width / this.view.zoom, height: rect.height / this.view.zoom });
  }
}

/** A font/context update may change geometry without editing the source. Both
 * content revision and layout generation must match before pointer hit-testing. */
export function requireCurrentGeometry(expected: GeometryIdentity, actual: GeometryIdentity): void {
  for (const identity of [expected, actual]) {
    if (!identity.documentId || !Number.isSafeInteger(identity.revision) || identity.revision < 0
      || !Number.isSafeInteger(identity.layoutGeneration) || identity.layoutGeneration < 0) {
      throw new GeometryError('invalidCoordinate', 'Invalid geometry identity');
    }
  }
  if (expected.documentId !== actual.documentId || expected.revision !== actual.revision
    || expected.layoutGeneration !== actual.layoutGeneration) {
    throw new GeometryError('staleGeometry', 'Geometry does not belong to the current document/layout revision');
  }
}
