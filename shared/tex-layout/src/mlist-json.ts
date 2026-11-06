// Convert math lists in the JSON form produced by tools/texoracle.py (TeX's
// \showlists, parsed) into engine noads. Used for LaTeX templates and tests.
import { decimalToScaled, MAX_DIMEN, NULL_FLAG } from "./arith.ts";
import type { FontCache } from "./latex.ts";
import {
  ACCENT_NOAD,
  BIN_NOAD,
  CHOICE_NODE,
  CLOSE_NOAD,
  DEFAULT_CODE,
  FRACTION_NOAD,
  INNER_NOAD,
  LEFT_NOAD,
  LIMITS,
  MIDDLE_NOAD,
  NO_LIMITS,
  OP_NOAD,
  OPEN_NOAD,
  ORD_NOAD,
  OVER_NOAD,
  PUNCT_NOAD,
  RADICAL_NOAD,
  REL_NOAD,
  RIGHT_NOAD,
  STYLE_NODE,
  UNDER_NOAD,
  VCENTER_NOAD,
  boxField,
  charField,
  emptyField,
  mlistField,
  newNoad,
  type Delimiter,
  type Field,
  type MItem,
} from "./mlist.ts";
import {
  CHAR_NODE,
  COND_MATH_GLUE,
  FIL,
  FILL,
  FILLL,
  GLUE_NODE,
  KERN_NODE,
  MATH_NODE,
  MU_GLUE,
  NORMAL,
  PENALTY_NODE,
  RULE_NODE,
  ZERO_GLUE,
  type BoxNode,
  type Node,
  type Tag,
} from "./nodes.ts";
import { EXACTLY, hpack, vpackage } from "./pack.ts";

export const NOAD_TYPES: Record<string, number> = {
  mathord: ORD_NOAD,
  mathop: OP_NOAD,
  mathbin: BIN_NOAD,
  mathrel: REL_NOAD,
  mathopen: OPEN_NOAD,
  mathclose: CLOSE_NOAD,
  mathpunct: PUNCT_NOAD,
  mathinner: INNER_NOAD,
  overline: OVER_NOAD,
  underline: UNDER_NOAD,
  vcenter: VCENTER_NOAD,
};

const STYLES: Record<string, number> = { display: 0, text: 2, script: 4, scriptscript: 6 };

const sp = decimalToScaled;

function order(unit: string): number {
  return unit === "fil" ? FIL : unit === "fill" ? FILL : unit === "filll" ? FILLL : NORMAL;
}

export class MlistJson {
  private fontIds: Record<string, string> = {};
  private readonly cache: FontCache;
  /** tag given to every node and noad created by the conversion */
  tag: Tag = undefined;
  /** the field of a character (the adapter puts letters in the editor's letter font) */
  charField: (fam: number, code: number) => Field = charField;

  constructor(cache: FontCache) {
    this.cache = cache;
  }

  convert(list: any[], fontIds: Record<string, string> = {}, tag?: Tag): MItem[] {
    this.fontIds = fontIds;
    this.tag = tag;
    return list.map((it) => this.item(it));
  }

  private delim(d: any): Delimiter {
    return { smallFam: d.smallFam, smallChar: d.smallCode, largeFam: d.largeFam, largeChar: d.largeCode };
  }

  field(f: any): Field {
    if (!f) return emptyField();
    if (f.kind === "char") return this.charField(f.fam, f.code);
    if (f.kind === "list") return mlistField(f.list.map((it: any) => this.item(it)));
    if (f.kind === "box") return boxField(this.box(f.box));
    throw new Error("unsupported field kind " + f.kind);
  }

  item(it: any): MItem {
    const tag = this.tag;
    const t = NOAD_TYPES[it.type];
    if (t !== undefined) {
      const n = newNoad(t, tag);
      if (it.limits === "limits") n.subtype = LIMITS;
      if (it.limits === "nolimits") n.subtype = NO_LIMITS;
      n.nucleus = this.field(it.nucleus);
      n.sup = this.field(it.sup);
      n.sub = this.field(it.sub);
      return n;
    }
    switch (it.type) {
      case "radical": {
        const n = newNoad(RADICAL_NOAD, tag);
        n.leftDelim = this.delim(it.delim);
        n.nucleus = this.field(it.nucleus);
        n.sup = this.field(it.sup);
        n.sub = this.field(it.sub);
        return n;
      }
      case "accent": {
        const n = newNoad(ACCENT_NOAD, tag);
        n.accentChr = charField(it.fam, it.code);
        n.nucleus = this.field(it.nucleus);
        n.sup = this.field(it.sup);
        n.sub = this.field(it.sub);
        return n;
      }
      case "left":
      case "right":
      case "middle": {
        const n = newNoad(it.type === "left" ? LEFT_NOAD : RIGHT_NOAD, tag);
        if (it.type === "middle") n.subtype = MIDDLE_NOAD;
        n.delim = this.delim(it.delim);
        return n;
      }
      case "fraction": {
        const n = newNoad(FRACTION_NOAD, tag);
        n.thickness = it.thickness === "default" ? DEFAULT_CODE : sp(it.thickness);
        n.num = this.field(it.num);
        n.den = this.field(it.den);
        if (it.left) n.leftDelim = this.delim(it.left);
        if (it.right) n.rightDelim = this.delim(it.right);
        return n;
      }
      case "style":
        return { type: STYLE_NODE, style: STYLES[it.style], tag };
      case "choice":
        return {
          type: CHOICE_NODE,
          display: it.display.map((x: any) => this.item(x)),
          text: it.text.map((x: any) => this.item(x)),
          script: it.script.map((x: any) => this.item(x)),
          scriptScript: it.scriptscript.map((x: any) => this.item(x)),
          tag,
        };
      case "glue": {
        if (it.kind === "nonscript") return { type: GLUE_NODE, spec: ZERO_GLUE, subtype: COND_MATH_GLUE, tag } as Node;
        const s = it.spec;
        const spec = {
          width: sp(s.width),
          stretch: sp(s.stretch),
          stretchOrder: order(s.stretchUnit),
          shrink: sp(s.shrink),
          shrinkOrder: order(s.shrinkUnit),
        };
        return { type: GLUE_NODE, spec, subtype: it.kind === "mu" ? MU_GLUE : NORMAL, tag } as Node;
      }
      case "kern":
        return { type: KERN_NODE, width: sp(it.width), subtype: it.mu ? MU_GLUE : NORMAL, tag } as Node;
      case "penalty":
        return { type: PENALTY_NODE, penalty: it.value, tag } as Node;
      default:
        throw new Error("unsupported mlist item " + it.type + " " + JSON.stringify(it).slice(0, 120));
    }
  }

  private dim(text: string): number {
    return text === "*" ? NULL_FLAG : sp(text);
  }

  hitem(it: any): Node {
    const tag = this.tag;
    switch (it.type) {
      case "char": {
        const name = this.fontIds[it.fontid];
        if (!name) throw new Error("unknown font id " + it.fontid);
        return { type: CHAR_NODE, font: this.cache.byFontname(name)!, char: it.code, tag };
      }
      case "math":
        return { type: MATH_NODE, width: sp(it.width), subtype: 0, tag };
      case "kern":
        return { type: KERN_NODE, width: sp(it.width), subtype: it.explicit ? 1 : NORMAL, tag };
      case "rule":
        return { type: RULE_NODE, height: this.dim(it.height), depth: this.dim(it.depth), width: this.dim(it.width), tag };
      case "glue": {
        const s = it.spec;
        return {
          type: GLUE_NODE,
          subtype: it.leaders ? (it.leaders === "c" ? 101 : it.leaders === "x" ? 102 : 100) : NORMAL,
          leader: it.leader ? (this.hitem(it.leader) as any) : undefined,
          spec: {
            width: sp(s.width),
            stretch: sp(s.stretch),
            stretchOrder: order(s.stretchUnit),
            shrink: sp(s.shrink),
            shrinkOrder: order(s.shrinkUnit),
          },
          tag,
        };
      }
      case "penalty":
        return { type: PENALTY_NODE, penalty: it.value, tag };
      case "hbox":
      case "vbox":
        return this.box(it);
      default:
        throw new Error("unsupported box item " + JSON.stringify(it).slice(0, 100));
    }
  }

  /** Rebuild a finished box; glue is re-set from the exact size, as TeX packed it. */
  box(b: any): BoxNode {
    const items = (b.content ?? []).map((x: any) => this.hitem(x));
    const [h, d, w] = b.raw.map(sp);
    const r = b.type === "hbox" ? hpack(items, w, EXACTLY, this.tag) : vpackage(items, h, EXACTLY, MAX_DIMEN, this.tag);
    r.width = w;
    r.height = h;
    r.depth = d;
    r.shift = b.shift ? sp(b.shift) : 0;
    return r;
  }
}

/** Serialize noads back to the oracle JSON shape (for structural diffs). */
export function toOracleJson(list: MItem[]): any[] {
  const inv: Record<number, string> = {};
  for (const [k, v] of Object.entries(NOAD_TYPES)) inv[v] = k;
  const field = (f: Field): any => {
    if (!f || f.mathType === 0) return null;
    if (f.mathType === 1 || f.mathType === 4) return { kind: "char", fam: f.fam, code: f.char };
    if (f.mathType === 3) return { kind: "list", list: toOracleJson(f.list!) };
    if (f.mathType === 2) return { kind: "box", box: { w: f.box!.width, h: f.box!.height, d: f.box!.depth } };
    return null;
  };
  const delim = (d: any) => d && { smallFam: d.smallFam, smallCode: d.smallChar, largeFam: d.largeFam, largeCode: d.largeChar };
  return list.map((q: any) => {
    if (inv[q.type]) {
      const o: any = { type: inv[q.type] };
      if (q.subtype === LIMITS) o.limits = "limits";
      if (q.subtype === NO_LIMITS) o.limits = "nolimits";
      o.nucleus = field(q.nucleus);
      o.sup = field(q.sup);
      o.sub = field(q.sub);
      return o;
    }
    switch (q.type) {
      case RADICAL_NOAD:
        return { type: "radical", delim: delim(q.leftDelim), nucleus: field(q.nucleus), sup: field(q.sup), sub: field(q.sub) };
      case ACCENT_NOAD:
        return { type: "accent", fam: q.accentChr.fam, code: q.accentChr.char, nucleus: field(q.nucleus), sup: field(q.sup), sub: field(q.sub) };
      case LEFT_NOAD:
        return { type: "left", delim: delim(q.delim) };
      case RIGHT_NOAD:
        return { type: q.subtype === MIDDLE_NOAD ? "middle" : "right", delim: delim(q.delim) };
      case FRACTION_NOAD:
        return { type: "fraction", thickness: q.thickness === DEFAULT_CODE ? "default" : q.thickness, num: field(q.num), den: field(q.den) };
      case STYLE_NODE:
        return { type: "style", style: ["display", "", "text", "", "script", "", "scriptscript"][q.style] };
      case CHOICE_NODE:
        return { type: "choice" };
      case GLUE_NODE:
        return { type: "glue", subtype: q.subtype, width: q.spec.width };
      case KERN_NODE:
        return { type: "kern", mu: q.subtype === MU_GLUE, width: q.width };
      default:
        return { type: "node" + q.type };
    }
  });
}
