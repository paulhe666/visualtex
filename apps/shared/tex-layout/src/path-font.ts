// Large operators drawn from outline data instead of a TeX font (the rare
// integrals the app takes from STIX Two Math: \sumint, \intbar, ...).
//
// Each command has a text-size and a display-size glyph. As in a TeX math
// extension font, the small one's character tag points to the large one, so
// TeX's make_op picks the large glyph in display style and centres either on
// the axis. Metrics are the outlines' own (font units, y up).
import { NULL_CHARACTER, TeXFont, type CharInfo, type TfmData } from "./tfm.ts";

export interface PathGlyph {
  path: string;
  advanceWidth: number;
  italicCorrection: number;
  height: number;
  depth: number;
}

export interface PathGlyphSet {
  name: string;
  unitsPerEm: number;
  glyphs: { command: string; aliases: string[]; small: PathGlyph; large: PathGlyph }[];
}

const LIST_TAG = 2;

const EMPTY_TFM: TfmData = {
  name: "paths",
  checksum: 0,
  designSize: 10 << 20,
  codingScheme: "paths",
  bc: 0,
  ec: 0,
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

export class PathFont extends TeXFont {
  readonly paths = true;
  readonly unitsPerEm: number;
  /** outline of each character code */
  readonly outlines: string[] = [];
  /** character code of each command's text-size glyph */
  readonly codes = new Map<string, number>();
  private readonly infos: CharInfo[] = [];
  private readonly w: number[] = [];
  private readonly h: number[] = [];
  private readonly d: number[] = [];
  private readonly ic: number[] = [];

  constructor(set: PathGlyphSet, size: number) {
    super({ ...EMPTY_TFM, name: "paths:" + set.name }, size);
    this.unitsPerEm = set.unitsPerEm;
    const scaled = (v: number) => Math.round((v / set.unitsPerEm) * size);
    set.glyphs.forEach((g, i) => {
      const small = 2 * i, large = small + 1;
      for (const [code, glyph] of [[small, g.small], [large, g.large]] as const) {
        this.outlines[code] = glyph.path;
        this.w[code] = scaled(glyph.advanceWidth);
        this.h[code] = scaled(glyph.height);
        this.d[code] = scaled(glyph.depth);
        this.ic[code] = scaled(glyph.italicCorrection);
        this.infos[code] = { exists: true, wi: code, hi: code, di: code, ii: code, tag: code === small ? LIST_TAG : 0, rem: code === small ? large : 0 };
      }
      for (const name of [g.command, ...g.aliases]) this.codes.set("\\" + name.replace(/^\\/, ""), small);
    });
  }

  override charInfo(c: number): CharInfo {
    return this.infos[c] ?? NULL_CHARACTER;
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

export function isPathFont(f: TeXFont): f is PathFont {
  return (f as PathFont).paths === true;
}
