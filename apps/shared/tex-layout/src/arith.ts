// TeX's arithmetic on scaled dimensions (tex.web part 7), ported literally.
// A "scaled" value is an integer number of sp (1pt = 65536sp). All results are
// integers; JavaScript doubles hold them exactly in TeX's range (< 2^30).

export const UNITY = 65536;
export const TWO = 131072;
export const MAX_DIMEN = 0o7777777777; // 2^30 - 1
export const NULL_FLAG = -0o10000000000; // -2^30: "running" rule dimension

/** Pascal `div`: truncates toward zero. */
export function div(a: number, b: number): number {
  return Math.trunc(a / b);
}

/** Pascal `mod`: sign follows the dividend. */
export function mod(a: number, b: number): number {
  return a - b * Math.trunc(a / b);
}

export function half(x: number): number {
  return (x & 1) !== 0 ? div(x + 1, 2) : div(x, 2);
}

/** Pascal round(): halves away from zero. */
export function texRound(x: number): number {
  return x >= 0 ? Math.floor(x + 0.5) : -Math.floor(-x + 0.5);
}

export let arithError = false;
let remainderValue = 0;

export function remainder(): number {
  return remainderValue;
}

export function resetArithError(): void {
  arithError = false;
}

export function multAndAdd(n: number, x: number, y: number, maxAnswer: number): number {
  if (n < 0) {
    x = -x;
    n = -n;
  }
  if (n === 0) return y;
  if (x <= div(maxAnswer - y, n) && -x <= div(maxAnswer + y, n)) return n * x + y;
  arithError = true;
  return 0;
}

export function nxPlusY(n: number, x: number, y: number): number {
  return multAndAdd(n, x, y, 0o7777777777);
}

export function xOverN(x: number, n: number): number {
  let negative = false;
  let result: number;
  if (n === 0) {
    arithError = true;
    remainderValue = x;
    return 0;
  }
  if (n < 0) {
    x = -x;
    n = -n;
    negative = true;
  }
  if (x >= 0) {
    result = div(x, n);
    remainderValue = mod(x, n);
  } else {
    result = -div(-x, n);
    remainderValue = -mod(-x, n);
  }
  if (negative) remainderValue = -remainderValue;
  return result;
}

export function xnOverD(x: number, n: number, d: number): number {
  let positive = true;
  if (x < 0) {
    x = -x;
    positive = false;
  }
  const t = mod(x, 0o100000) * n;
  let u = div(x, 0o100000) * n + div(t, 0o100000);
  const v = mod(u, d) * 0o100000 + mod(t, 0o100000);
  if (div(u, d) >= 0o100000) arithError = true;
  else u = 0o100000 * div(u, d) + div(v, d);
  if (positive) {
    remainderValue = mod(v, d);
    return u;
  }
  remainderValue = -mod(v, d);
  return -u;
}

/** TeX's round_decimals for k decimal digits (as read by the scanner). */
export function roundDecimals(digits: number[]): number {
  let a = 0;
  for (let k = digits.length - 1; k >= 0; k--) {
    a = div(a + digits[k] * TWO, 10);
  }
  return div(a + 1, 2);
}

/** Parse a decimal dimension the way TeX's scanner does ("-1.25" -> sp). */
export function decimalToScaled(text: string): number {
  let s = text.trim();
  let negative = false;
  while (s.startsWith("-") || s.startsWith("+")) {
    if (s[0] === "-") negative = !negative;
    s = s.slice(1);
  }
  const [intPart, fracPart = ""] = s.split(".");
  const integer = intPart === "" ? 0 : parseInt(intPart, 10);
  const digits = fracPart.slice(0, 17).split("").map((d) => d.charCodeAt(0) - 48);
  const value = integer * UNITY + roundDecimals(digits);
  return negative ? -value : value;
}

/** Physical units as TeX converts them to points (tex.web 458). */
const TRUE_UNITS: Record<string, [number, number]> = {
  pt: [1, 1],
  in: [7227, 100],
  pc: [12, 1],
  cm: [7227, 254],
  mm: [7227, 2540],
  bp: [7227, 7200],
  dd: [1238, 1157],
  cc: [14856, 1157],
};

/**
 * TeX's scan_dimen once the number is read (tex.web 448-461): a decimal
 * constant in a unit. `em` and `ex` are the current font's quad and x-height,
 * `px` is pdfTeX's \pdfpxdimen (all in sp).
 */
export function scanDimen(text: string, unit: string, units: { em: number; ex: number; px: number }): number {
  let s = text.trim();
  let negative = false;
  while (s.startsWith("-") || s.startsWith("+")) {
    if (s[0] === "-") negative = !negative;
    s = s.slice(1);
  }
  const [intPart, fracPart = ""] = s.split(".");
  let curVal = intPart === "" ? 0 : parseInt(intPart, 10);
  let f = roundDecimals(fracPart.slice(0, 17).split("").map((d) => d.charCodeAt(0) - 48));
  const v = unit === "em" ? units.em : unit === "ex" ? units.ex : unit === "px" ? units.px : 0;
  if (v !== 0) {
    curVal = nxPlusY(curVal, v, xnOverD(v, f, UNITY));
  } else if (unit === "sp") {
    // the fraction is ignored
  } else {
    const [num, denom] = TRUE_UNITS[unit] ?? [1, 1];
    if (num !== 1 || denom !== 1) {
      curVal = xnOverD(curVal, num, denom);
      f = div(num * f + UNITY * remainder(), denom);
      curVal = curVal + div(f, UNITY);
      f = mod(f, UNITY);
    }
    curVal = curVal * UNITY + f;
  }
  return negative ? -curVal : curVal;
}

/**
 * TeX's store_scaled (tex.web 571-572): fix_word `fw` times font size `z` (sp).
 */
export function storeScaled(fw: number, z: number): number {
  let alpha = 16;
  while (z >= 0o40000000) {
    z = div(z, 2);
    alpha += alpha;
  }
  const beta = div(256, alpha);
  alpha = alpha * z;
  const u = fw >>> 0;
  const a = (u >>> 24) & 255;
  const b = (u >>> 16) & 255;
  const c = (u >>> 8) & 255;
  const d = u & 255;
  const sw = div(div(div(d * z, 0o400) + c * z, 0o400) + b * z, beta);
  if (a === 0) return sw;
  if (a === 255) return sw - alpha;
  throw new Error("bad fix_word");
}
