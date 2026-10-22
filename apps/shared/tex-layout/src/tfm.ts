// TeX font metric data (tex.web part 30) for a font loaded at a given size.
import { div, storeScaled } from "./arith.ts";

/** Raw TFM data as produced by tools/tfm2json.py (fix_words kept raw). */
export interface TfmData {
  name: string;
  checksum: number;
  designSize: number; // fix_word, points * 2^20
  codingScheme: string;
  bc: number;
  ec: number;
  widths: number[];
  heights: number[];
  depths: number[];
  italics: number[];
  chars: Record<string, [number, number, number, number, number, number]>;
  ligKern: [number, number, number, number][];
  kerns: number[];
  exten: [number, number, number, number][];
  params: number[];
  /**
   * Not TFM data: how far a character's outline reaches past its box, left of
   * 0, right of its width, above its height, below its depth (thousandths of
   * the design size, tools/inkbounds.py); only characters where it matters.
   */
  ink?: Record<string, [number, number, number, number]>;
}

export const NO_TAG = 0;
export const LIG_TAG = 1;
export const LIST_TAG = 2;
export const EXT_TAG = 3;
export const STOP_FLAG = 128;
export const KERN_FLAG = 128;

/** char_info four_quarters for one character (or the null character). */
export interface CharInfo {
  exists: boolean;
  wi: number;
  hi: number;
  di: number;
  ii: number;
  tag: number;
  rem: number;
}

export const NULL_CHARACTER: CharInfo = { exists: false, wi: 0, hi: 0, di: 0, ii: 0, tag: 0, rem: 0 };

/** A lig/kern instruction (skip_byte, next_char, op_byte, rem_byte). */
export interface LigKern {
  skip: number;
  next: number;
  op: number;
  rem: number;
}

let fontSerial = 0;

/**
 * An internal font: one TFM file loaded at one size (TeX's internal_font_number).
 */
export class TeXFont {
  readonly id: number;
  readonly name: string;
  readonly size: number; // sp
  readonly designSize: number; // sp
  readonly bc: number;
  readonly ec: number;
  readonly widths: number[];
  readonly heights: number[];
  readonly depths: number[];
  readonly italics: number[];
  readonly kerns: number[];
  readonly params: number[]; // params[1..n]; params[1] = slant (scaled number)
  readonly ligKern: LigKern[];
  readonly exten: [number, number, number, number][];
  skewChar = -1;
  private readonly info: Map<number, CharInfo>;
  private readonly ink: Record<string, [number, number, number, number]>;

  constructor(data: TfmData, size: number | null = null) {
    this.id = ++fontSerial;
    this.name = data.name;
    // The design size is stored as fix_word points; TeX converts it to sp by
    // shifting (tex.web 568): design_size := z div 16 ... with z in 2^20 units.
    this.designSize = div(data.designSize, 16);
    this.size = size ?? this.designSize;
    const z = this.size;
    this.bc = data.bc;
    this.ec = data.ec;
    this.widths = data.widths.map((fw) => storeScaled(fw, z));
    this.heights = data.heights.map((fw) => storeScaled(fw, z));
    this.depths = data.depths.map((fw) => storeScaled(fw, z));
    this.italics = data.italics.map((fw) => storeScaled(fw, z));
    this.kerns = data.kerns.map((fw) => storeScaled(fw, z));
    this.params = [0];
    data.params.forEach((fw, k) => {
      if (k === 0) this.params.push(fw >> 4); // slant: a pure number (tex.web 575)
      else this.params.push(storeScaled(fw, z));
    });
    while (this.params.length <= 7) this.params.push(0); // at least 7 params (tex.web 576)
    this.ligKern = data.ligKern.map(([skip, next, op, rem]) => ({ skip, next, op, rem }));
    this.exten = data.exten;
    this.info = new Map();
    for (const [code, [wi, hi, di, ii, tag, rem]] of Object.entries(data.chars)) {
      this.info.set(Number(code), { exists: wi > 0, wi, hi, di, ii, tag, rem });
    }
    this.ink = data.ink ?? {};
  }

  /**
   * How far a character's ink reaches past its box (sp at this size): left of
   * its origin, right of its width, above its height, below its depth; null
   * when it stays inside.
   */
  inkOverhang(c: number): [number, number, number, number] | null {
    const v = this.ink[c];
    if (!v) return null;
    const sp = (thousandths: number) => Math.ceil((thousandths * this.size) / 1000);
    return [sp(v[0]), sp(v[1]), sp(v[2]), sp(v[3])];
  }

  charInfo(c: number): CharInfo {
    if (c < this.bc || c > this.ec) return NULL_CHARACTER;
    return this.info.get(c) ?? NULL_CHARACTER;
  }

  charWidth(i: CharInfo): number {
    return this.widths[i.wi];
  }

  charHeight(i: CharInfo): number {
    return this.heights[i.hi];
  }

  charDepth(i: CharInfo): number {
    return this.depths[i.di];
  }

  charItalic(i: CharInfo): number {
    return this.italics[i.ii];
  }

  /** Index of the first lig/kern instruction for a character with LIG_TAG. */
  ligKernStart(i: CharInfo): number {
    return i.rem;
  }

  /** tex.web 557: restart address when the first instruction has skip > 128. */
  ligKernRestart(k: LigKern): number {
    return 256 * k.op + k.rem;
  }

  /** tex.web 557: char_kern for a kern instruction. */
  charKern(k: LigKern): number {
    return this.kerns[256 * (k.op - KERN_FLAG) + k.rem];
  }

  param(k: number): number {
    return this.params[k] ?? 0;
  }

  get slant(): number {
    return this.param(1);
  }

  get space(): number {
    return this.param(2);
  }

  get xHeight(): number {
    return this.param(5);
  }

  get quad(): number {
    return this.param(6);
  }
}
