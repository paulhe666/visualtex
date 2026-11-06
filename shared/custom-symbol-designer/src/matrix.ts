import type { CustomSymbolVectorTransform } from "./symbolTypes.ts";

/** An SVG affine matrix [a b c d e f]: (x, y) -> (a x + c y + e, b x + d y + f). */
export type Affine = [number, number, number, number, number, number];

export const IDENTITY: Affine = [1, 0, 0, 1, 0, 0];

export function multiply(left: Affine, right: Affine): Affine {
  const [a1, b1, c1, d1, e1, f1] = left;
  const [a2, b2, c2, d2, e2, f2] = right;
  return [
    a1 * a2 + c1 * b2,
    b1 * a2 + d1 * b2,
    a1 * c2 + c1 * d2,
    b1 * c2 + d1 * d2,
    a1 * e2 + c1 * f2 + e1,
    b1 * e2 + d1 * f2 + f1,
  ];
}

export function translate(x: number, y: number): Affine {
  return [1, 0, 0, 1, x, y];
}

export function scale(x: number, y: number): Affine {
  return [x, 0, 0, y, 0, 0];
}

export function rotate(degrees: number): Affine {
  const radians = (degrees * Math.PI) / 180;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  return [cosine, sine, -sine, cosine, 0, 0];
}

export function skewX(degrees: number): Affine {
  return [1, 0, Math.tan((degrees * Math.PI) / 180), 1, 0, 0];
}

export function skewY(degrees: number): Affine {
  return [1, Math.tan((degrees * Math.PI) / 180), 0, 1, 0, 0];
}

export function apply(matrix: Affine, x: number, y: number) {
  return {
    x: matrix[0] * x + matrix[2] * y + matrix[4],
    y: matrix[1] * x + matrix[3] * y + matrix[5],
  };
}

export function invert(matrix: Affine): Affine {
  const [a, b, c, d, e, f] = matrix;
  const determinant = a * d - b * c || 1e-12;
  return [
    d / determinant,
    -b / determinant,
    -c / determinant,
    a / determinant,
    (c * f - d * e) / determinant,
    (b * e - a * f) / determinant,
  ];
}

/**
 * The matrix of a designer transform, in the order the SVG `transform`
 * attribute of the canvas writes it: translate, then about the origin
 * rotate, skewX, skewY and scale, then (innermost) the compiler's matrix.
 */
export function transformMatrix(
  transform: CustomSymbolVectorTransform | undefined,
  includeMatrix = true,
): Affine {
  if (!transform) return IDENTITY;
  const ox = transform.originX ?? 0;
  const oy = transform.originY ?? 0;
  const parts: Affine[] = [
    translate(transform.translateX ?? 0, transform.translateY ?? 0),
    translate(ox, oy),
    rotate(transform.rotateDeg ?? 0),
    skewX(transform.skewXDeg ?? 0),
    skewY(transform.skewYDeg ?? 0),
    scale(transform.scaleX ?? 1, transform.scaleY ?? 1),
    translate(-ox, -oy),
  ];
  if (includeMatrix && transform.matrix) parts.push([...transform.matrix] as Affine);
  return parts.reduce(multiply, IDENTITY);
}

/** The SVG `transform` attribute of a designer transform (undefined: none). */
export function svgTransformAttribute(transform: CustomSymbolVectorTransform | undefined) {
  if (!transform) return undefined;
  const parts: string[] = [];
  const tx = transform.translateX ?? 0;
  const ty = transform.translateY ?? 0;
  const sx = transform.scaleX ?? 1;
  const sy = transform.scaleY ?? 1;
  const kx = transform.skewXDeg ?? 0;
  const ky = transform.skewYDeg ?? 0;
  const angle = transform.rotateDeg ?? 0;
  const ox = transform.originX ?? 0;
  const oy = transform.originY ?? 0;
  if (tx || ty) parts.push(`translate(${tx} ${ty})`);
  if (ox || oy) parts.push(`translate(${ox} ${oy})`);
  if (angle) parts.push(`rotate(${angle})`);
  if (kx) parts.push(`skewX(${kx})`);
  if (ky) parts.push(`skewY(${ky})`);
  if (sx !== 1 || sy !== 1) parts.push(`scale(${sx} ${sy})`);
  if (ox || oy) parts.push(`translate(${-ox} ${-oy})`);
  if (transform.matrix) parts.push(`matrix(${transform.matrix.join(" ")})`);
  return parts.length ? parts.join(" ") : undefined;
}

/** How much a matrix scales lengths on average (sqrt |det|). */
export function lengthScale(matrix: Affine) {
  return Math.sqrt(Math.abs(matrix[0] * matrix[3] - matrix[1] * matrix[2])) || 0;
}
