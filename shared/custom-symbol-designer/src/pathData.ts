// SVG path data reduced to absolute moves, lines and cubic curves (what
// TikZ's `--` and `.. controls ..` draw), with the bounding box of the curves.

export type PathSegment =
  | { op: "M"; x: number; y: number }
  | { op: "L"; x: number; y: number }
  | { op: "C"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: "Z" };

const argumentCounts: Record<string, number> = {
  M: 2, L: 2, H: 1, V: 1, C: 6, S: 4, Q: 4, T: 2, A: 7, Z: 0,
};

function tokenize(d: string) {
  const tokens: Array<string | number> = [];
  const pattern = /([MLHVCSQTAZmlhvcsqtaz])|([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(d))) {
    if (match[1]) tokens.push(match[1]);
    else tokens.push(Number(match[2]));
  }
  return tokens;
}

function arcToCubics(
  x1: number, y1: number, rxIn: number, ryIn: number, angleDeg: number,
  largeArc: boolean, sweep: boolean, x2: number, y2: number,
): PathSegment[] {
  if (x1 === x2 && y1 === y2) return [];
  let rx = Math.abs(rxIn);
  let ry = Math.abs(ryIn);
  if (!rx || !ry) return [{ op: "L", x: x2, y: y2 }];
  const phi = (angleDeg * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy;
  const y1p = -sin * dx + cos * dy;
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const numerator = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const denominator = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let factor = Math.sqrt(Math.max(0, numerator / denominator));
  if (largeArc === sweep) factor = -factor;
  const cxp = (factor * rx * y1p) / ry;
  const cyp = (-factor * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const angle = (ux: number, uy: number, vx: number, vy: number) => {
    const sign = ux * vy - uy * vx < 0 ? -1 : 1;
    const dot = (ux * vx + uy * vy) / (Math.hypot(ux, uy) * Math.hypot(vx, vy));
    return sign * Math.acos(Math.max(-1, Math.min(1, dot)));
  };
  const theta1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let delta = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;
  const pieces = Math.max(1, Math.ceil(Math.abs(delta) / (Math.PI / 2)));
  const step = delta / pieces;
  const kappa = (4 / 3) * Math.tan(step / 4);
  const point = (t: number) => ({
    x: cx + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin,
    y: cy + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos,
  });
  const derivative = (t: number) => ({
    x: -rx * Math.sin(t) * cos - ry * Math.cos(t) * sin,
    y: -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos,
  });
  const segments: PathSegment[] = [];
  for (let piece = 0; piece < pieces; piece += 1) {
    const t1 = theta1 + piece * step;
    const t2 = t1 + step;
    const p1 = point(t1);
    const p2 = point(t2);
    const d1 = derivative(t1);
    const d2 = derivative(t2);
    segments.push({
      op: "C",
      x1: p1.x + kappa * d1.x,
      y1: p1.y + kappa * d1.y,
      x2: p2.x - kappa * d2.x,
      y2: p2.y - kappa * d2.y,
      x: piece === pieces - 1 ? x2 : p2.x,
      y: piece === pieces - 1 ? y2 : p2.y,
    });
  }
  return segments;
}

/** `d` as absolute M / L / C / Z segments. Throws on malformed data. */
export function normalizePathData(d: string): PathSegment[] {
  const tokens = tokenize(d);
  const segments: PathSegment[] = [];
  let index = 0;
  let command = "";
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  let lastControlX = 0;
  let lastControlY = 0;
  let lastWasCubic = false;
  let lastWasQuadratic = false;
  const take = () => {
    const value = tokens[index++];
    if (typeof value !== "number") throw new Error("Malformed SVG path data.");
    return value;
  };
  while (index < tokens.length) {
    const token = tokens[index];
    if (typeof token === "string") {
      command = token;
      index += 1;
      if (command === "Z" || command === "z") {
        segments.push({ op: "Z" });
        x = startX;
        y = startY;
        lastWasCubic = lastWasQuadratic = false;
        continue;
      }
    } else if (!command) {
      throw new Error("SVG path data must start with a command.");
    }
    const upper = command.toUpperCase();
    const relative = command !== upper;
    if (!(upper in argumentCounts) || upper === "Z") throw new Error("Malformed SVG path data.");
    const ox = relative ? x : 0;
    const oy = relative ? y : 0;
    switch (upper) {
      case "M": {
        x = ox + take();
        y = oy + take();
        startX = x;
        startY = y;
        segments.push({ op: "M", x, y });
        command = relative ? "l" : "L";
        lastWasCubic = lastWasQuadratic = false;
        break;
      }
      case "L":
        x = ox + take();
        y = oy + take();
        segments.push({ op: "L", x, y });
        lastWasCubic = lastWasQuadratic = false;
        break;
      case "H":
        x = ox + take();
        segments.push({ op: "L", x, y });
        lastWasCubic = lastWasQuadratic = false;
        break;
      case "V":
        y = oy + take();
        segments.push({ op: "L", x, y });
        lastWasCubic = lastWasQuadratic = false;
        break;
      case "C": {
        const x1 = ox + take(), y1 = oy + take(), x2 = ox + take(), y2 = oy + take();
        x = ox + take();
        y = oy + take();
        segments.push({ op: "C", x1, y1, x2, y2, x, y });
        lastControlX = x2;
        lastControlY = y2;
        lastWasCubic = true;
        lastWasQuadratic = false;
        break;
      }
      case "S": {
        const x1 = lastWasCubic ? 2 * x - lastControlX : x;
        const y1 = lastWasCubic ? 2 * y - lastControlY : y;
        const x2 = ox + take(), y2 = oy + take();
        x = ox + take();
        y = oy + take();
        segments.push({ op: "C", x1, y1, x2, y2, x, y });
        lastControlX = x2;
        lastControlY = y2;
        lastWasCubic = true;
        lastWasQuadratic = false;
        break;
      }
      case "Q":
      case "T": {
        let qx: number;
        let qy: number;
        if (upper === "Q") {
          qx = ox + take();
          qy = oy + take();
        } else {
          qx = lastWasQuadratic ? 2 * x - lastControlX : x;
          qy = lastWasQuadratic ? 2 * y - lastControlY : y;
        }
        const endX = ox + take();
        const endY = oy + take();
        segments.push({
          op: "C",
          x1: x + (2 / 3) * (qx - x),
          y1: y + (2 / 3) * (qy - y),
          x2: endX + (2 / 3) * (qx - endX),
          y2: endY + (2 / 3) * (qy - endY),
          x: endX,
          y: endY,
        });
        x = endX;
        y = endY;
        lastControlX = qx;
        lastControlY = qy;
        lastWasQuadratic = true;
        lastWasCubic = false;
        break;
      }
      case "A": {
        const rx = take(), ry = take(), rotation = take(), large = take(), sweep = take();
        const endX = ox + take();
        const endY = oy + take();
        segments.push(...arcToCubics(x, y, rx, ry, rotation, large !== 0, sweep !== 0, endX, endY));
        x = endX;
        y = endY;
        lastWasCubic = lastWasQuadratic = false;
        break;
      }
    }
  }
  return segments;
}

/** The bounding box of the curves themselves (not their control points). */
export function pathSegmentsBounds(segments: readonly PathSegment[]) {
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  const include = (px: number, py: number) => {
    if (px < minX) minX = px;
    if (px > maxX) maxX = px;
    if (py < minY) minY = py;
    if (py > maxY) maxY = py;
  };
  let x = 0;
  let y = 0;
  for (const segment of segments) {
    if (segment.op === "Z") continue;
    if (segment.op === "C") {
      for (const axis of ["x", "y"] as const) {
        const p0 = axis === "x" ? x : y;
        const p1 = axis === "x" ? segment.x1 : segment.y1;
        const p2 = axis === "x" ? segment.x2 : segment.y2;
        const p3 = segment[axis];
        // roots of the derivative of the cubic Bezier along this axis
        const a = -p0 + 3 * p1 - 3 * p2 + p3;
        const b = 2 * (p0 - 2 * p1 + p2);
        const c = p1 - p0;
        const roots: number[] = [];
        if (Math.abs(a) < 1e-12) {
          if (Math.abs(b) > 1e-12) roots.push(-c / b);
        } else {
          const discriminant = b * b - 4 * a * c;
          if (discriminant >= 0) {
            const root = Math.sqrt(discriminant);
            roots.push((-b + root) / (2 * a), (-b - root) / (2 * a));
          }
        }
        for (const t of roots) {
          if (t <= 0 || t >= 1) continue;
          const mt = 1 - t;
          const value = mt * mt * mt * p0 + 3 * mt * mt * t * p1 + 3 * mt * t * t * p2 + t * t * t * p3;
          if (axis === "x") include(value, y);
          else include(x, value);
        }
      }
    }
    include(segment.x, segment.y);
    x = segment.x;
    y = segment.y;
  }
  if (!Number.isFinite(minX)) return null;
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}
