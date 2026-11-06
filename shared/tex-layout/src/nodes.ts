// Box, glue, kern, rule and character nodes (tex.web part 10), as JS objects.
// Lists are arrays. Every node may carry `tag`, the id of the formula atom that
// produced it; tags do not affect any dimension.
import type { TeXFont } from "./tfm.ts";
import { NULL_FLAG } from "./arith.ts";

export const HLIST_NODE = 0;
export const VLIST_NODE = 1;
export const RULE_NODE = 2;
export const MATH_NODE = 9;
export const GLUE_NODE = 10;
export const KERN_NODE = 11;
export const PENALTY_NODE = 12;
export const CHAR_NODE = 100; // not a TeX type code: char nodes are "is_char_node"

// glue orders and signs
export const NORMAL = 0;
export const FIL = 1;
export const FILL = 2;
export const FILLL = 3;
export const STRETCHING = 1;
export const SHRINKING = 2;

// glue subtypes
export const COND_MATH_GLUE = 98;
export const MU_GLUE = 99;
export const A_LEADERS = 100;
export const C_LEADERS = 101;
export const X_LEADERS = 102;

// kern subtypes
export const EXPLICIT = 1;
export const ACC_KERN = 2;

export type Tag = unknown;

export interface CharNode {
  type: typeof CHAR_NODE;
  font: TeXFont;
  char: number;
  tag?: Tag;
}

export interface BoxNode {
  type: typeof HLIST_NODE | typeof VLIST_NODE;
  width: number;
  height: number;
  depth: number;
  shift: number;
  list: Node[];
  glueSet: number;
  glueSign: number;
  glueOrder: number;
  tag?: Tag;
  /** Editor metadata: role of this box (e.g. "numerator"); never affects layout. */
  role?: string;
  /** Editor metadata: the branch-start atom of the field this box was made from. */
  firstTag?: Tag;
}

export interface RuleNode {
  type: typeof RULE_NODE;
  width: number;
  height: number;
  depth: number;
  tag?: Tag;
  /** colour of the rule (\colorbox backgrounds, \fcolorbox frames); default: the atom's */
  color?: string;
}

export interface GlueSpec {
  width: number;
  stretch: number;
  stretchOrder: number;
  shrink: number;
  shrinkOrder: number;
}

export interface GlueNode {
  type: typeof GLUE_NODE;
  spec: GlueSpec;
  subtype: number;
  /** leader box or rule when subtype >= A_LEADERS */
  leader?: BoxNode | RuleNode;
  tag?: Tag;
}

export interface KernNode {
  type: typeof KERN_NODE;
  width: number;
  subtype: number;
  tag?: Tag;
  /** Editor metadata: the italic correction after a character (tex.web 755); never affects layout. */
  italic?: boolean;
}

export interface PenaltyNode {
  type: typeof PENALTY_NODE;
  penalty: number;
  tag?: Tag;
}

export interface MathNode {
  type: typeof MATH_NODE;
  width: number;
  subtype: number;
  tag?: Tag;
}

export type Node = CharNode | BoxNode | RuleNode | GlueNode | KernNode | PenaltyNode | MathNode;

export const ZERO_GLUE: GlueSpec = { width: 0, stretch: 0, stretchOrder: NORMAL, shrink: 0, shrinkOrder: NORMAL };
export const SS_GLUE: GlueSpec = { width: 0, stretch: 65536, stretchOrder: FIL, shrink: 65536, shrinkOrder: FIL };

export function newNullBox(tag?: Tag): BoxNode {
  return {
    type: HLIST_NODE,
    width: 0,
    height: 0,
    depth: 0,
    shift: 0,
    list: [],
    glueSet: 0,
    glueSign: NORMAL,
    glueOrder: NORMAL,
    tag,
  };
}

export function newRule(tag?: Tag): RuleNode {
  return { type: RULE_NODE, width: NULL_FLAG, height: NULL_FLAG, depth: NULL_FLAG, tag };
}

export function newKern(width: number, tag?: Tag): KernNode {
  return { type: KERN_NODE, width, subtype: NORMAL, tag };
}

export function newGlue(spec: GlueSpec, tag?: Tag): GlueNode {
  return { type: GLUE_NODE, spec, subtype: NORMAL, tag };
}

export function newPenalty(penalty: number, tag?: Tag): PenaltyNode {
  return { type: PENALTY_NODE, penalty, tag };
}

export function newCharacter(font: TeXFont, char: number, tag?: Tag): CharNode {
  return { type: CHAR_NODE, font, char, tag };
}

export function isCharNode(n: Node | undefined): n is CharNode {
  return n !== undefined && n.type === CHAR_NODE;
}

export function isBox(n: Node | undefined): n is BoxNode {
  return n !== undefined && (n.type === HLIST_NODE || n.type === VLIST_NODE);
}
