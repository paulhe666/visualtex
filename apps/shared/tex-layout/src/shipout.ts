// Absolute positions for a finished box, following hlist_out/vlist_out
// (tex.web 619-637) exactly, including TeX's cumulative glue rounding.
// Coordinates are sp; x grows to the right, y grows downward, and the
// reference point of the outermost box is (0, 0).
import { NULL_FLAG, texRound } from "./arith.ts";
import {
  A_LEADERS,
  CHAR_NODE,
  C_LEADERS,
  GLUE_NODE,
  HLIST_NODE,
  KERN_NODE,
  MATH_NODE,
  NORMAL,
  RULE_NODE,
  STRETCHING,
  VLIST_NODE,
  type BoxNode,
  type RuleNode,
  type Tag,
} from "./nodes.ts";
import type { TeXFont } from "./tfm.ts";

export interface PlacedGlyph {
  font: TeXFont;
  char: number;
  x: number;
  y: number; // baseline
  tag?: Tag;
}

export interface PlacedRule {
  x: number;
  y: number; // top edge
  width: number;
  height: number;
  tag?: Tag;
  color?: string;
}

export interface PlacedBox {
  box: BoxNode;
  x: number; // left edge
  y: number; // baseline
}

export interface PlacedSpan {
  x: number;
  width: number;
  y: number; // baseline
  tag?: Tag;
  /** a character's italic correction */
  italic?: boolean;
}

export interface ShipResult {
  glyphs: PlacedGlyph[];
  rules: PlacedRule[];
  /** horizontal extent of tagged glue and kerns (for editor regions) */
  spans: PlacedSpan[];
  /** every box with its position, outermost first (for editor regions) */
  boxes: PlacedBox[];
  /**
   * The baseline of the line on which each tag's outermost material sits,
   * before any shift of that material (a big operator centred on the axis
   * stays on its line; a script's characters are on the script's line).
   */
  lineOf: Map<Tag, number>;
}

const BILLION = 1000000000;

function vetGlue(g: number): number {
  if (g > BILLION) return BILLION;
  if (g < -BILLION) return -BILLION;
  return g;
}

export function shipOut(box: BoxNode): ShipResult {
  const out: ShipResult = { glyphs: [], rules: [], spans: [], boxes: [], lineOf: new Map() };
  if (box.tag !== undefined) out.lineOf.set(box.tag, 0);
  if (box.type === VLIST_NODE) vlistOut(box, 0, 0, out);
  else hlistOut(box, 0, 0, out);
  return out;
}

function hlistOut(thisBox: BoxNode, leftEdge: number, baseLine: number, out: ShipResult): void {
  out.boxes.push({ box: thisBox, x: leftEdge, y: baseLine });
  let curH = leftEdge;
  let curG = 0;
  let curGlue = 0;
  const gOrder = thisBox.glueOrder;
  const gSign = thisBox.glueSign;
  for (const p of thisBox.list) {
    const tag = (p as { tag?: Tag }).tag;
    if (tag !== undefined && !out.lineOf.has(tag)) out.lineOf.set(tag, baseLine);
    switch (p.type) {
      case CHAR_NODE:
        out.glyphs.push({ font: p.font, char: p.char, x: curH, y: baseLine, tag: p.tag });
        curH += p.font.charWidth(p.font.charInfo(p.char));
        break;
      case HLIST_NODE:
      case VLIST_NODE:
        if (p.list.length === 0) {
          out.boxes.push({ box: p, x: curH, y: baseLine + p.shift });
          curH += p.width;
        } else {
          const edge = curH;
          if (p.type === VLIST_NODE) vlistOut(p, curH, baseLine + p.shift, out);
          else hlistOut(p, curH, baseLine + p.shift, out);
          curH = edge + p.width;
        }
        break;
      case RULE_NODE: {
        let ruleHt = p.height;
        let ruleDp = p.depth;
        const ruleWd = p.width;
        if (ruleHt === NULL_FLAG) ruleHt = thisBox.height;
        if (ruleDp === NULL_FLAG) ruleDp = thisBox.depth;
        ruleHt = ruleHt + ruleDp;
        if (ruleHt > 0 && ruleWd > 0) {
          out.rules.push({ x: curH, y: baseLine + ruleDp - ruleHt, width: ruleWd, height: ruleHt, tag: p.tag, color: (p as RuleNode).color });
        }
        curH += ruleWd;
        break;
      }
      case GLUE_NODE: {
        const g = p.spec;
        let ruleWd = g.width - curG;
        if (gSign !== NORMAL) {
          if (gSign === STRETCHING) {
            if (g.stretchOrder === gOrder) {
              curGlue += g.stretch;
              curG = texRound(vetGlue(thisBox.glueSet * curGlue));
            }
          } else if (g.shrinkOrder === gOrder) {
            curGlue -= g.shrink;
            curG = texRound(vetGlue(thisBox.glueSet * curGlue));
          }
        }
        ruleWd += curG;
        if (p.tag !== undefined && p.subtype < A_LEADERS) out.spans.push({ x: curH, width: ruleWd, y: baseLine, tag: p.tag });
        if (p.subtype >= A_LEADERS && p.leader) {
          // tex.web 626-628: output leaders in an hlist
          const leaderBox = p.leader;
          if (leaderBox.type === RULE_NODE) {
            let ruleHt = leaderBox.height;
            let ruleDp = leaderBox.depth;
            if (ruleHt === NULL_FLAG) ruleHt = thisBox.height;
            if (ruleDp === NULL_FLAG) ruleDp = thisBox.depth;
            ruleHt += ruleDp;
            if (ruleHt > 0 && ruleWd > 0) {
              out.rules.push({ x: curH, y: baseLine + ruleDp - ruleHt, width: ruleWd, height: ruleHt, tag: p.tag, color: leaderBox.color });
            }
            curH += ruleWd;
            break;
          }
          const leaderWd = leaderBox.width;
          if (leaderWd > 0 && ruleWd > 0) {
            ruleWd += 10; // compensate for floating-point rounding
            const edge = curH + ruleWd;
            let lx = 0;
            if (p.subtype === A_LEADERS) {
              const saveH = curH;
              curH = leftEdge + leaderWd * Math.trunc((curH - leftEdge) / leaderWd);
              if (curH < saveH) curH += leaderWd;
            } else {
              const lq = Math.trunc(ruleWd / leaderWd);
              const lr = ruleWd % leaderWd;
              if (p.subtype === C_LEADERS) curH += Math.trunc(lr / 2);
              else {
                lx = Math.trunc(lr / (lq + 1));
                curH += Math.trunc((lr - (lq - 1) * lx) / 2);
              }
            }
            while (curH + leaderWd <= edge) {
              const saveH = curH;
              if (leaderBox.type === VLIST_NODE) vlistOut(leaderBox, curH, baseLine + leaderBox.shift, out);
              else hlistOut(leaderBox, curH, baseLine + leaderBox.shift, out);
              curH = saveH + leaderWd + lx;
            }
            curH = edge - 10;
            break;
          }
        }
        curH += ruleWd;
        break;
      }
      case KERN_NODE:
      case MATH_NODE:
        if (p.type === KERN_NODE && p.tag !== undefined && p.width > 0)
          out.spans.push({ x: curH, width: p.width, y: baseLine, tag: p.tag, italic: p.italic });
        curH += p.width;
        break;
      default:
        break;
    }
  }
}

function vlistOut(thisBox: BoxNode, leftEdge: number, baseLine: number, out: ShipResult): void {
  out.boxes.push({ box: thisBox, x: leftEdge, y: baseLine });
  let curV = baseLine - thisBox.height;
  let curG = 0;
  let curGlue = 0;
  const gOrder = thisBox.glueOrder;
  const gSign = thisBox.glueSign;
  for (const p of thisBox.list) {
    switch (p.type) {
      case CHAR_NODE:
        throw new Error("confusion: vlistout");
      case HLIST_NODE:
      case VLIST_NODE:
        if (p.tag !== undefined && !out.lineOf.has(p.tag)) out.lineOf.set(p.tag, curV + p.height);
        if (p.list.length === 0) {
          out.boxes.push({ box: p, x: leftEdge + p.shift, y: curV + p.height });
          curV += p.height + p.depth;
        } else {
          curV += p.height;
          const saveV = curV;
          if (p.type === VLIST_NODE) vlistOut(p, leftEdge + p.shift, curV, out);
          else hlistOut(p, leftEdge + p.shift, curV, out);
          curV = saveV + p.depth;
        }
        break;
      case RULE_NODE: {
        const ruleHt = p.height + p.depth;
        let ruleWd = p.width;
        if (ruleWd === NULL_FLAG) ruleWd = thisBox.width;
        curV += ruleHt;
        if (ruleHt > 0 && ruleWd > 0) {
          out.rules.push({ x: leftEdge, y: curV - ruleHt, width: ruleWd, height: ruleHt, tag: p.tag, color: (p as RuleNode).color });
        }
        break;
      }
      case GLUE_NODE: {
        const g = p.spec;
        let ruleHt = g.width - curG;
        if (gSign !== NORMAL) {
          if (gSign === STRETCHING) {
            if (g.stretchOrder === gOrder) {
              curGlue += g.stretch;
              curG = texRound(vetGlue(thisBox.glueSet * curGlue));
            }
          } else if (g.shrinkOrder === gOrder) {
            curGlue -= g.shrink;
            curG = texRound(vetGlue(thisBox.glueSet * curGlue));
          }
        }
        ruleHt += curG;
        if (p.subtype >= A_LEADERS && p.leader) {
          // tex.web 635-637: output leaders in a vlist
          const leaderBox = p.leader;
          if (leaderBox.type === RULE_NODE) {
            let ruleWd = leaderBox.width;
            if (ruleWd === NULL_FLAG) ruleWd = thisBox.width;
            curV += ruleHt;
            if (ruleHt > 0 && ruleWd > 0) out.rules.push({ x: leftEdge, y: curV - ruleHt, width: ruleWd, height: ruleHt, tag: p.tag, color: leaderBox.color });
            break;
          }
          const leaderHt = leaderBox.height + leaderBox.depth;
          if (leaderHt > 0 && ruleHt > 0) {
            ruleHt += 10;
            const edge = curV + ruleHt;
            let lx = 0;
            const topEdge = baseLine - thisBox.height;
            if (p.subtype === A_LEADERS) {
              const saveV = curV;
              curV = topEdge + leaderHt * Math.trunc((curV - topEdge) / leaderHt);
              if (curV < saveV) curV += leaderHt;
            } else {
              const lq = Math.trunc(ruleHt / leaderHt);
              const lr = ruleHt % leaderHt;
              if (p.subtype === C_LEADERS) curV += Math.trunc(lr / 2);
              else {
                lx = Math.trunc(lr / (lq + 1));
                curV += Math.trunc((lr - (lq - 1) * lx) / 2);
              }
            }
            while (curV + leaderHt <= edge) {
              const h = leftEdge + leaderBox.shift;
              curV += leaderBox.height;
              const saveV = curV;
              if (leaderBox.type === VLIST_NODE) vlistOut(leaderBox, h, curV, out);
              else hlistOut(leaderBox, h, curV, out);
              curV = saveV - leaderBox.height + leaderHt + lx;
            }
            curV = edge - 10;
            break;
          }
        }
        curV += ruleHt;
        break;
      }
      case KERN_NODE:
        curV += p.width;
        break;
      default:
        break;
    }
  }
}

/** TeX's print_scaled (tex.web 103): the decimal TeX shows for a dimension. */
export function printScaled(s: number): string {
  let out = "";
  if (s < 0) {
    out += "-";
    s = -s;
  }
  out += String(Math.trunc(s / 65536));
  out += ".";
  s = 10 * (s % 65536) + 5;
  let delta = 10;
  do {
    if (delta > 65536) s = s + 0o100000 - 50000; // round the last digit
    out += String.fromCharCode(48 + Math.trunc(s / 65536));
    s = 10 * (s % 65536);
    delta *= 10;
  } while (s > delta);
  return out;
}
