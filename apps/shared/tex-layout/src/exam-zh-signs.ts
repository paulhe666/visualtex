// The congruence and similarity signs of Chinese geometry as exam-zh draws them.
//
// exam-zh-symbols.sty (the exam-zh class's symbol module) draws \cong (全等, ≌), \ncong, \sim (相似, ∽) and \nsim as TikZ
// pictures; their outlines are generated from its coordinates by scripts/gen-exam-zh-signs.py (exam-zh-signs.generated.ts).
// The engine stores the meanings under the names that mean them everywhere: ≌ is \backcong (U+224C), ∽ is \backsim
// (U+223D), the negations \not\backcong and \not\backsim; exam-zh's own spellings (\cong, \sim ...) are the app's to map
// at its edges. TikZ strokes with an absolute width at every size, so each sign has its own outline for 10, 7 and 5 pt,
// and the font at a size takes the outline drawn for that size, as a TeX font's sizes have their own designs.
import { EXAM_ZH_SIGNS } from "./exam-zh-signs.generated.ts";
import { NULL_CHARACTER, TeXFont, type CharInfo, type TfmData } from "./tfm.ts";

/** Character codes of the signs in the font. */
export const EXAM_ZH_SIGN = { backcong: 0, notbackcong: 1, backsim: 2, notbacksim: 3 } as const;

const ORDER = ["backcong", "notbackcong", "backsim", "notbacksim"] as const;

/** The sign a command or character stands for (negated: after \not), or null. */
export function examZhSignCode(name: string, negated = false): number | null {
  switch (name) {
    case "\\backcong":
    case "≌":
      return negated ? EXAM_ZH_SIGN.notbackcong : EXAM_ZH_SIGN.backcong;
    case "\\backsim":
    case "∽":
      return negated ? EXAM_ZH_SIGN.notbacksim : EXAM_ZH_SIGN.backsim;
    default:
      return null;
  }
}

const EMPTY_TFM: TfmData = {
  name: "paths:exam-zh-signs",
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

/** The four signs at one size (in sp): the outlines exam-zh draws at the nearest of 10, 7 and 5 pt. */
export class ExamZhSignFont extends TeXFont {
  readonly paths = true;
  readonly unitsPerEm = EXAM_ZH_SIGNS.unitsPerEm;
  /** outline of each character code, in thousandths of this size */
  readonly outlines: string[] = [];
  private readonly infos: CharInfo[] = [];
  private readonly w: number[] = [];
  private readonly h: number[] = [];
  private readonly d: number[] = [];

  constructor(size: number) {
    super(EMPTY_TFM, size);
    const pt = size / 65536;
    const index = EXAM_ZH_SIGNS.sizes.reduce((best, s, i) => (Math.abs(s - pt) < Math.abs(EXAM_ZH_SIGNS.sizes[best] - pt) ? i : best), 0);
    const scaled = (v: number) => Math.round((v / this.unitsPerEm) * size);
    ORDER.forEach((name, code) => {
      const g = EXAM_ZH_SIGNS.glyphs[name][index];
      this.outlines[code] = g.path;
      this.w[code] = scaled(g.advanceWidth);
      this.h[code] = scaled(g.height);
      this.d[code] = scaled(g.depth);
      this.infos[code] = { exists: true, wi: code, hi: code, di: code, ii: code, tag: 0, rem: 0 };
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

  override charItalic(): number {
    return 0;
  }
}
