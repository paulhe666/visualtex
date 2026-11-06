// User spacing adjustments: how far scripts, fractions, limits, radicals,
// operators and arrays stand from what they belong to, as a percentage of
// what LaTeX does (100 = LaTeX exactly). Each adjustment scales the TeX
// parameters that place that part, so every construct built from them
// follows, and the editor's caret and hit-testing (which read the layout)
// follow too.

/** One adjustment: which TeX values it scales. */
export interface SpacingDefinition {
  key: SpacingKey;
  group: "scripts" | "fractions" | "limits" | "spaces" | "arrays";
  /** \fontdimen numbers of the symbol font (family 2) it scales */
  sy?: number[];
  /** \fontdimen numbers of the extension font (family 3) it scales */
  ex?: number[];
  min: number;
  max: number;
}

export type SpacingKey =
  | "supRaise"
  | "subDrop"
  | "scriptAttach"
  | "scriptSpace"
  | "numRaise"
  | "denomDrop"
  | "ruleThickness"
  | "radicalGap"
  | "limitsGap"
  | "limitsDistance"
  | "limitsPadding"
  | "delimiterSize"
  | "thinSpace"
  | "medSpace"
  | "thickSpace"
  | "arrayColSep"
  | "arrayRowSep"
  | "alignRowSep";

/** Percentages by key; a missing key is 100. */
export type Spacing = Partial<Record<SpacingKey, number>>;

export const SPACING_DEFINITIONS: readonly SpacingDefinition[] = [
  // superscripts: sup1, sup2, sup3 (raise in display, text, cramped styles)
  { key: "supRaise", group: "scripts", sy: [13, 14, 15], min: 0, max: 250 },
  // subscripts: sub1, sub2 (drop alone, with a superscript)
  { key: "subDrop", group: "scripts", sy: [16, 17], min: 0, max: 250 },
  // sup_drop, sub_drop: scripts of a tall nucleus hang from its top and bottom
  { key: "scriptAttach", group: "scripts", sy: [18, 19], min: 0, max: 250 },
  // \scriptspace: the space after a script
  { key: "scriptSpace", group: "scripts", min: 0, max: 400 },
  // num1, num2, num3: numerators (display, text, \atop)
  { key: "numRaise", group: "fractions", sy: [8, 9, 10], min: 0, max: 250 },
  // denom1, denom2: denominators
  { key: "denomDrop", group: "fractions", sy: [11, 12], min: 0, max: 250 },
  // default_rule_thickness: fraction bars, radical and over/underlines, and the gaps measured by it
  { key: "ruleThickness", group: "fractions", ex: [8], min: 25, max: 300 },
  // the clearance between a radical's content and its overbar
  { key: "radicalGap", group: "fractions", min: 0, max: 300 },
  // big_op_spacing1, 2: least gap between a large operator and its limits
  { key: "limitsGap", group: "limits", ex: [9, 10], min: 0, max: 300 },
  // big_op_spacing3, 4: least distance of the limits' baselines from the operator
  { key: "limitsDistance", group: "limits", ex: [11, 12], min: 0, max: 300 },
  // big_op_spacing5: extra space above and below the limits
  { key: "limitsPadding", group: "limits", ex: [13], min: 0, max: 400 },
  // \delimiterfactor: how much of the content \left and \right cover
  { key: "delimiterSize", group: "limits", min: 50, max: 130 },
  // \thinmuskip (operators, punctuation), \medmuskip (binary operators), \thickmuskip (relations)
  { key: "thinSpace", group: "spaces", min: 0, max: 400 },
  { key: "medSpace", group: "spaces", min: 0, max: 400 },
  { key: "thickSpace", group: "spaces", min: 0, max: 400 },
  // \arraycolsep (matrices, arrays, cases), \arraystretch, \jot (aligned, gathered)
  { key: "arrayColSep", group: "arrays", min: 0, max: 400 },
  { key: "arrayRowSep", group: "arrays", min: 50, max: 300 },
  { key: "alignRowSep", group: "arrays", min: 0, max: 500 },
];

const BY_KEY = new Map(SPACING_DEFINITIONS.map((d) => [d.key, d]));

/** The percentage of an adjustment, within its range (100 when unset or invalid). */
export function spacingPercent(spacing: Spacing | null | undefined, key: SpacingKey): number {
  const v = spacing?.[key];
  const d = BY_KEY.get(key)!;
  if (typeof v !== "number" || !Number.isFinite(v)) return 100;
  return Math.min(d.max, Math.max(d.min, Math.round(v)));
}

/** Per-mille factors of the font parameters: family 2 and family 3 \fontdimen number -> factor. */
export function spacingParamScales(spacing: Spacing | null | undefined): { sy: Record<number, number>; ex: Record<number, number> } {
  const sy: Record<number, number> = {}, ex: Record<number, number> = {};
  for (const d of SPACING_DEFINITIONS) {
    const p = spacingPercent(spacing, d.key);
    if (p === 100) continue;
    for (const k of d.sy ?? []) sy[k] = p * 10;
    for (const k of d.ex ?? []) ex[k] = p * 10;
  }
  return { sy, ex };
}

/** `v` (sp) times `permille`/1000, rounded to sp; exact integers, so every computer gets the same. */
export function scalePermille(v: number, permille: number): number {
  if (permille === 1000) return v;
  const r = Math.round((Math.abs(v) * permille) / 1000);
  return v < 0 ? -r : r;
}

/** A stable key of the adjustments that are not 100 (to see whether they changed). */
export function spacingKey(spacing: Spacing | null | undefined): string {
  return SPACING_DEFINITIONS.map((d) => spacingPercent(spacing, d.key))
    .map((p, i) => (p === 100 ? "" : `${SPACING_DEFINITIONS[i].key}=${p}`))
    .filter(Boolean)
    .join(",");
}
