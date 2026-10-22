// Packaging (tex.web part 33): hpack and vpackage, without the diagnostics.
import { MAX_DIMEN } from "./arith.ts";
import {
  A_LEADERS,
  CHAR_NODE,
  FIL,
  FILL,
  FILLL,
  GLUE_NODE,
  HLIST_NODE,
  KERN_NODE,
  MATH_NODE,
  NORMAL,
  RULE_NODE,
  SHRINKING,
  STRETCHING,
  VLIST_NODE,
  type BoxNode,
  type Node,
  type Tag,
} from "./nodes.ts";

export const EXACTLY = 0;
export const ADDITIONAL = 1;

/** hpack(p, w, m): with m=ADDITIONAL and w=0 this is hpack(p, natural). */
export function hpack(list: Node[], w = 0, m = ADDITIONAL, tag?: Tag): BoxNode {
  let h = 0;
  let d = 0;
  let x = 0;
  const totalStretch = [0, 0, 0, 0];
  const totalShrink = [0, 0, 0, 0];
  for (const p of list) {
    switch (p.type) {
      case CHAR_NODE: {
        const i = p.font.charInfo(p.char);
        x += p.font.charWidth(i);
        const s = p.font.charHeight(i);
        if (s > h) h = s;
        const dd = p.font.charDepth(i);
        if (dd > d) d = dd;
        break;
      }
      case HLIST_NODE:
      case VLIST_NODE:
      case RULE_NODE: {
        x += p.width;
        const s = p.type === RULE_NODE ? 0 : p.shift;
        if (p.height - s > h) h = p.height - s;
        if (p.depth + s > d) d = p.depth + s;
        break;
      }
      case GLUE_NODE: {
        const g = p.spec;
        x += g.width;
        totalStretch[g.stretchOrder] += g.stretch;
        totalShrink[g.shrinkOrder] += g.shrink;
        if (p.subtype >= A_LEADERS && p.leader) {
          if (p.leader.height > h) h = p.leader.height;
          if (p.leader.depth > d) d = p.leader.depth;
        }
        break;
      }
      case KERN_NODE:
      case MATH_NODE:
        x += p.width;
        break;
      default:
        break;
    }
  }
  const r: BoxNode = {
    type: HLIST_NODE,
    width: 0,
    height: h,
    depth: d,
    shift: 0,
    list,
    glueSet: 0,
    glueSign: NORMAL,
    glueOrder: NORMAL,
    tag,
  };
  if (m === ADDITIONAL) w = x + w;
  r.width = w;
  x = w - x;
  if (x === 0) return r;
  if (x > 0) {
    const o = totalStretch[FILLL] !== 0 ? FILLL : totalStretch[FILL] !== 0 ? FILL : totalStretch[FIL] !== 0 ? FIL : NORMAL;
    r.glueOrder = o;
    r.glueSign = STRETCHING;
    if (totalStretch[o] !== 0) r.glueSet = x / totalStretch[o];
    else {
      r.glueSign = NORMAL;
      r.glueSet = 0;
    }
    return r;
  }
  const o = totalShrink[FILLL] !== 0 ? FILLL : totalShrink[FILL] !== 0 ? FILL : totalShrink[FIL] !== 0 ? FIL : NORMAL;
  r.glueOrder = o;
  r.glueSign = SHRINKING;
  if (totalShrink[o] !== 0) r.glueSet = -x / totalShrink[o];
  else {
    r.glueSign = NORMAL;
    r.glueSet = 0;
  }
  if (totalShrink[o] < -x && o === NORMAL && list.length > 0) r.glueSet = 1;
  return r;
}

/** vpackage(p, h, m, l); vpack(p, natural) = vpackage(p, 0, ADDITIONAL, MAX_DIMEN). */
export function vpackage(list: Node[], h = 0, m = ADDITIONAL, l = MAX_DIMEN, tag?: Tag): BoxNode {
  let w = 0;
  let d = 0;
  let x = 0;
  const totalStretch = [0, 0, 0, 0];
  const totalShrink = [0, 0, 0, 0];
  for (const p of list) {
    switch (p.type) {
      case CHAR_NODE:
        throw new Error("vpack: character node in a vertical list");
      case HLIST_NODE:
      case VLIST_NODE:
      case RULE_NODE: {
        x += d + p.height;
        d = p.depth;
        const s = p.type === RULE_NODE ? 0 : p.shift;
        if (p.width + s > w) w = p.width + s;
        break;
      }
      case GLUE_NODE: {
        x += d;
        d = 0;
        const g = p.spec;
        x += g.width;
        totalStretch[g.stretchOrder] += g.stretch;
        totalShrink[g.shrinkOrder] += g.shrink;
        if (p.subtype >= A_LEADERS && p.leader && p.leader.width > w) w = p.leader.width;
        break;
      }
      case KERN_NODE:
        x += d + p.width;
        d = 0;
        break;
      default:
        break;
    }
  }
  const r: BoxNode = {
    type: VLIST_NODE,
    width: w,
    height: 0,
    depth: 0,
    shift: 0,
    list,
    glueSet: 0,
    glueSign: NORMAL,
    glueOrder: NORMAL,
    tag,
  };
  if (d > l) {
    x = x + d - l;
    r.depth = l;
  } else r.depth = d;
  if (m === ADDITIONAL) h = x + h;
  r.height = h;
  x = h - x;
  if (x === 0) return r;
  if (x > 0) {
    const o = totalStretch[FILLL] !== 0 ? FILLL : totalStretch[FILL] !== 0 ? FILL : totalStretch[FIL] !== 0 ? FIL : NORMAL;
    r.glueOrder = o;
    r.glueSign = STRETCHING;
    if (totalStretch[o] !== 0) r.glueSet = x / totalStretch[o];
    else {
      r.glueSign = NORMAL;
      r.glueSet = 0;
    }
    return r;
  }
  const o = totalShrink[FILLL] !== 0 ? FILLL : totalShrink[FILL] !== 0 ? FILL : totalShrink[FIL] !== 0 ? FIL : NORMAL;
  r.glueOrder = o;
  r.glueSign = SHRINKING;
  if (totalShrink[o] !== 0) r.glueSet = -x / totalShrink[o];
  else {
    r.glueSign = NORMAL;
    r.glueSet = 0;
  }
  if (totalShrink[o] < -x && o === NORMAL && list.length > 0) r.glueSet = 1;
  return r;
}

export function vpack(list: Node[], tag?: Tag): BoxNode {
  return vpackage(list, 0, ADDITIONAL, MAX_DIMEN, tag);
}
