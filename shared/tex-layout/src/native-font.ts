// Fonts that are not TeX fonts (a Chinese font for \text, the editor's letter
// font): metrics are measured by the host (canvas in the browser) and drawn
// with the real Unicode character. A NativeFont behaves like a TeXFont without
// ligatures or kerns; its spacing parameters come from its measured space.
import { NULL_CHARACTER, TeXFont, type CharInfo, type TfmData } from "./tfm.ts";

/** Ink box of one character at 1em: advance width, ascent, descent. */
export interface NativeMetrics {
  width: number;
  ascent: number;
  descent: number;
  /** how far the ink reaches right of the advance (a slanted letter), used as its italic correction */
  italic?: number;
}

/** Measures text at 1em in a CSS font: family list and style words ("italic", "bold", "italic bold" or none). */
export type NativeMeasure = (family: string, text: string, style?: string) => NativeMetrics | null;

export interface NativeFontOptions {
  /** CSS font-style and font-weight words ("italic", "bold", "italic bold") */
  style?: string;
  /** the characters this font is for (the others go to a fallback font) */
  covers?: (c: number) => boolean;
}

const EMPTY_TFM: TfmData = {
  name: "native",
  checksum: 0,
  designSize: 10 << 20,
  codingScheme: "native",
  // any Unicode character (charInfo says which ones the font has)
  bc: 0,
  ec: 0x10ffff,
  widths: [0],
  heights: [0],
  depths: [0],
  italics: [0],
  chars: {},
  ligKern: [],
  kerns: [],
  exten: [],
  params: [0, 0, 0, 0, 0, 0, 0],
};

export class NativeFont extends TeXFont {
  readonly native = true;
  readonly family: string;
  /** CSS font-style and font-weight words ("" for the regular face) */
  readonly style: string;
  private readonly covers: ((c: number) => boolean) | undefined;
  private readonly measure: NativeMeasure;
  private readonly cache = new Map<number, CharInfo>();
  private readonly w: number[] = [0];
  private readonly h: number[] = [0];
  private readonly d: number[] = [0];
  private readonly ic: number[] = [0];
  private measuredSpace: number | undefined;

  constructor(family: string, size: number, measure: NativeMeasure, options: NativeFontOptions = {}) {
    super({ ...EMPTY_TFM, name: "native:" + family + (options.style ? ":" + options.style : "") }, size);
    this.family = family;
    this.style = options.style ?? "";
    this.covers = options.covers;
    this.measure = measure;
  }

  override charInfo(c: number): CharInfo {
    let i = this.cache.get(c);
    if (i) return i;
    const m = this.covers && !this.covers(c) ? null : this.measure(this.family, String.fromCodePoint(c), this.style);
    if (!m) {
      this.cache.set(c, NULL_CHARACTER);
      return NULL_CHARACTER;
    }
    const k = this.w.length;
    this.w.push(Math.round(m.width * this.size));
    this.h.push(Math.round(m.ascent * this.size));
    this.d.push(Math.round(m.descent * this.size));
    this.ic.push(Math.round(Math.max(0, m.italic ?? 0) * this.size));
    i = { exists: true, wi: k, hi: k, di: k, ii: k, tag: 0, rem: 0 };
    this.cache.set(c, i);
    return i;
  }

  /**
   * TeX's spacing parameters from the measured space: 2 space, 3 stretch
   * (1/2 of it), 4 shrink (1/3), 5 x-height, 6 quad, 7 extra space (1/3).
   */
  override param(k: number): number {
    if (k === 5) {
      const x = this.charInfo(0x78);
      return x.exists ? this.charHeight(x) : 0;
    }
    if (k === 6) return this.size;
    if (k < 2 || k > 7) return 0;
    if (this.measuredSpace === undefined) this.measuredSpace = Math.round((this.measure(this.family, " ", this.style)?.width ?? 0.25) * this.size);
    return k === 2 ? this.measuredSpace : k === 3 ? Math.round(this.measuredSpace / 2) : Math.round(this.measuredSpace / 3);
  }

  override charWidth(i: CharInfo): number {
    return this.w[i.wi];
  }

  override charHeight(i: CharInfo): number {
    return this.h[i.hi];
  }

  override charDepth(i: CharInfo): number {
    return this.d[i.di];
  }

  override charItalic(i: CharInfo): number {
    return this.ic[i.ii] ?? 0;
  }
}

export function isNativeFont(f: TeXFont): f is NativeFont {
  return (f as NativeFont).native === true;
}
