// VisualTeX layout engine: LaTeX-exact typesetting for the editor kernel.
//
// The kernel keeps its formula model, editing, caret logic and selection;
// this module replaces how formulas are laid out and drawn. It returns markup
// in which every atom has an element with data-atom-id spanning its TeX box,
// so the kernel's caret, selection and hit-testing code work unchanged.
import { FontCache } from "./latex.ts";
import { FAM, LatexAdapter, type AtomLike, type LatexData } from "./latex-adapter.ts";
import { MathTypesetter, TEXT_STYLE } from "./mlist.ts";
import { MATH_NODE, type BoxNode, type Node, type Tag } from "./nodes.ts";
import { hpack } from "./pack.ts";
import { alignShift, caretBarHtml, computeRegions, formulaHtml, type CaretBar, type FormulaElement, type Region } from "./painter.ts";
import { buildEditorGeometry, caretMetrics, type EditorGeometry, type GeometryAtom } from "./editor-geometry.ts";
import { shipOut, type ShipResult } from "./shipout.ts";
import type { TeXFont, TfmData } from "./tfm.ts";
import { NativeFont, type NativeMeasure } from "./native-font.ts";
import { PathFont, type PathGlyphSet } from "./path-font.ts";
import type { Spacing } from "./spacing.ts";
import { formulaSvgBody, type GlyphTable, type SvgFormula, type SvgOverlay } from "./svg.ts";

export type { GlyphTable, SvgFormula } from "./svg.ts";
export { SPACING_DEFINITIONS, spacingPercent, type Spacing, type SpacingKey, type SpacingDefinition } from "./spacing.ts";

export interface LayoutData extends LatexData {
  tfm: Record<string, TfmData>;
  fonts: { tfm: string; family: string }[];
}

export interface RenderedFormula {
  html: string;
  box: BoxNode;
  ship: ShipResult;
  regions: Map<Tag, Region>;
}

type KernelAtom = AtomLike & {
  id?: string;
  caret?: string;
  isSelected?: boolean;
  containsCaret?: boolean;
  isSuggestion?: boolean;
  isError?: boolean;
  branches?: string[];
  rows?: AtomLike[][][];
  parent?: KernelAtom;
};

interface Rect {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

const TEN_PT = 10 * 65536;
/** thin rules (fraction bars, radical overbars) get a hit area at least this tall */
const MIN_HIT = Math.round(0.12 * TEN_PT);

function union(rs: Rect[]): Rect {
  return {
    x0: Math.min(...rs.map((r) => r.x0)),
    y0: Math.min(...rs.map((r) => r.y0)),
    x1: Math.max(...rs.map((r) => r.x1)),
    y1: Math.max(...rs.map((r) => r.y1)),
  };
}

function padRule(r: Rect): Rect {
  const h = r.y1 - r.y0;
  if (h >= MIN_HIT) return r;
  const d = Math.ceil((MIN_HIT - h) / 2);
  return { x0: r.x0, y0: r.y0 - d, x1: r.x1, y1: r.y1 + d };
}

/** Each atom's own ink (glyphs and rules tagged with it, not with its children). */
function ownInk(ship: ShipResult): Map<Tag, { glyphs: Rect[]; rules: Rect[] }> {
  const ink = new Map<Tag, { glyphs: Rect[]; rules: Rect[] }>();
  const of = (t: Tag) => {
    let v = ink.get(t);
    if (!v) ink.set(t, (v = { glyphs: [], rules: [] }));
    return v;
  };
  for (const g of ship.glyphs) {
    const i = g.font.charInfo(g.char);
    of(g.tag).glyphs.push({ x0: g.x, y0: g.y - g.font.charHeight(i), x1: g.x + g.font.charWidth(i), y1: g.y + g.font.charDepth(i) });
  }
  for (const r of ship.rules) of(r.tag).rules.push({ x0: r.x, y0: r.y, x1: r.x + r.width, y1: r.y + r.height });
  return ink;
}

/** Opening and closing fences: the atom's own glyphs left and right of its middle. */
function fences(glyphs: Rect[], r: Region): FormulaElement[] {
  const mid = (r.x0 + r.x1) / 2;
  const open = glyphs.filter((g) => (g.x0 + g.x1) / 2 < mid);
  const close = glyphs.filter((g) => (g.x0 + g.x1) / 2 >= mid);
  const out: FormulaElement[] = [];
  if (open.length) out.push({ ...union(open), classes: "ML__open" });
  if (close.length) out.push({ ...union(close), classes: "ML__close" });
  return out;
}

/** Elements over the structural ink of an atom, named as in MathLive's markup. */
function inkParts(a: KernelAtom, ink: ReturnType<typeof ownInk>, r: Region): FormulaElement[] {
  const own = ink.get(a);
  if (!own) return [];
  switch (a.type) {
    case "genfrac":
      return [...own.rules.map((x) => ({ ...padRule(x), classes: "ML__frac-line" })), ...fences(own.glyphs, r)];
    case "surd":
      return [
        ...(own.glyphs.length ? [{ ...union(own.glyphs), classes: "ML__sqrt-sign" }] : []),
        ...own.rules.map((x) => ({ ...padRule(x), classes: "ML__sqrt-line" })),
      ];
    case "accent":
      return own.glyphs.length ? [{ ...union(own.glyphs), classes: "ML__visualtex-accent-mark" }] : [];
    case "leftright":
      return fences(own.glyphs, r);
  }
  return [];
}

/** MathLive's class for the box of an atom of this kind. */
function structuralClass(a: KernelAtom, adapter: LatexAdapter): string {
  switch (a.type) {
    case "root":
      return "ML__base";
    case "genfrac":
      return "ML__mfrac";
    case "surd":
      return "ML__sqrt";
    case "accent":
      return "ML__accent-body";
    case "subsup":
      return "ML__msubsup";
    case "array":
      return "ML__mtable";
    case "mop":
    case "operator":
    case "extensible-symbol":
      return "ML__op-group";
    case "placeholder": {
      let c = "ML__placeholder visualtex-structural-placeholder";
      const p = a.parent;
      if (p?.type === "accent" && a.parentBranch === "body" && (p.branch("body") ?? []).length === 2) c += " visualtex-accent-structural-placeholder";
      if (a.isSelected) c += " ML__placeholder-selected";
      return c;
    }
    case "latex":
      return "ML__raw-latex" + (a.isSuggestion ? " ML__suggestion" : "") + (a.isError ? " ML__error" : "");
    case "error":
      // a LaTeX command the kernel does not know but the engine renders
      return a.command && adapter.knowsCommand(a.command) ? "" : "ML__error";
  }
  if (a.command === "\\rule") return "ML__rule";
  return htmlClassOf(a) ?? "";
}

/** The class of a \class{...}{...} / \htmlClass{...}{...} annotation. */
function htmlClassOf(a: KernelAtom): string | null {
  if (a.command !== "\\class" && a.command !== "\\htmlClass") return null;
  const c = (a.args as unknown[] | undefined)?.[0];
  return typeof c === "string" && c ? c : null;
}

/** Atoms with an element in static markup (previews): what CSS styles. */
function hasStaticElement(a: KernelAtom, adapter: LatexAdapter): boolean {
  return a.type === "placeholder" || (a.type === "error" && structuralClass(a, adapter) !== "") || a.command === "\\rule" || htmlClassOf(a) !== null;
}

/** An atom's box contains its children's (as nested boxes do in MathLive). */
function propagateRegions(regions: Map<Tag, Region>, stop: KernelAtom | null): void {
  for (const [tag, r] of [...regions]) {
    for (let p = (tag as KernelAtom)?.parent; p && p !== stop && p.type !== "root"; p = p.parent) {
      const q = regions.get(p);
      if (!q) regions.set(p, { ...r, tag: p, hasGlyph: false });
      else {
        q.x0 = Math.min(q.x0, r.x0);
        q.y0 = Math.min(q.y0, r.y0);
        q.x1 = Math.max(q.x1, r.x1);
        q.y1 = Math.max(q.y1, r.y1);
      }
    }
  }
}

export interface LayoutOptions {
  /** measures characters of non-TeX fonts (canvas in the browser) */
  measureNative?: NativeMeasure;
  /** CSS font-family list for characters the TeX fonts lack (e.g. Chinese) */
  fallbackFamily?: string | (() => string);
  /** large operators drawn from outlines (rare integrals), see PathFont */
  pathGlyphs?: PathGlyphSet;
  /** where the editor shows the caret: the editing core's rule (apps/shared/visualtex-core) */
  caretState?: (mathfield: any) => CaretState;
  /**
   * The letter font chosen for formulas (CSS family lists of its upright and
   * italic faces), or null for LaTeX's Computer Modern: letters, digits and
   * roman text are then set in it, measured; symbols keep the TeX fonts.
   */
  letterFamilies?: () => { upright: string; italic: string } | null;
  /** the user's spacing adjustments (spacing.ts), read at every layout; null or absent: LaTeX's */
  spacing?: () => Spacing | null;
}

/** Chinese, Japanese and Korean characters: they stay in the formula's Chinese font. */
const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]|[\u{20000}-\u{2FA1F}]/u;

/** Where a mathfield shows its caret: at a stop, in an empty slot, or nowhere (a range is selected). */
export type CaretState =
  | { kind: "stop"; offset: number }
  | { kind: "empty"; placeholder: GeometryAtom }
  | { kind: "none" };

export type { EditorGeometry, CaretGeometry, Rect as GeometryRect } from "./editor-geometry.ts";

export function createVisualTexLayout(data: LayoutData, options: LayoutOptions = {}) {
  const cache = new FontCache((name) => {
    const d = data.tfm[name];
    if (!d) throw new Error("VisualTeX layout: missing font metrics " + name);
    return d;
  }, data.profile);
  const adapter = new LatexAdapter(data, cache);
  const nativeCache = new Map<string, NativeFont>();
  if (options.measureNative) {
    const measure = options.measureNative;
    adapter.native = {
      fallback(size: number) {
        const fam = typeof options.fallbackFamily === "function" ? options.fallbackFamily() : options.fallbackFamily ?? "serif";
        const key = fam + "@" + size;
        let f = nativeCache.get(key);
        if (!f) nativeCache.set(key, (f = new NativeFont(fam, size, measure)));
        return f;
      },
    };
  }
  if (options.pathGlyphs) {
    // text, script and scriptscript sizes, as for the TeX fonts
    const sizes = [10, 7, 5].map((pt) => new PathFont(options.pathGlyphs!, pt * 65536));
    adapter.fams.set(FAM.paths, sizes as [PathFont, PathFont, PathFont]);
    // LaTeX's own definitions (amsmath, esint) take precedence
    adapter.pathOps = new Map([...sizes[0].codes].filter(([cmd]) => !data.templates[cmd]));
  }
  // the letter font (LayoutOptions.letterFamilies) fills the adapter's letter
  // families at 10, 7 and 5pt; they are rebuilt when the choice changes
  let letterChoice: string | null = null;
  const syncLetterFonts = () => {
    const families = options.measureNative ? options.letterFamilies?.() ?? null : null;
    const choice = families ? families.upright + "\n" + families.italic : null;
    if (choice === letterChoice) return;
    letterChoice = choice;
    for (const fam of [FAM.letterIt, FAM.letterUp, FAM.letterBf, FAM.letterBfIt]) adapter.fams.delete(fam);
    if (!families) return;
    const measure = options.measureNative!;
    const covers = (c: number) => !CJK.test(String.fromCodePoint(c));
    const sizes = (family: string, style: string): [TeXFont, TeXFont, TeXFont] => {
      const at = (pt: number) => new NativeFont(family, pt * 65536, measure, { style, covers });
      return [at(10), at(7), at(5)];
    };
    adapter.fams.set(FAM.letterIt, sizes(families.italic, "italic"));
    adapter.fams.set(FAM.letterUp, sizes(families.upright, ""));
    adapter.fams.set(FAM.letterBf, sizes(families.upright, "bold"));
    adapter.fams.set(FAM.letterBfIt, sizes(families.italic, "italic bold"));
  };
  const familyOf = new Map(data.fonts.map((f) => [f.tfm, f.family]));
  // the caret's size and the strut of empty formulas, from the profile's fonts
  const caret = caretMetrics(adapter.caretFonts());
  let serial = 0;
  const idOf = (atom: KernelAtom): string => {
    if (!atom.id) atom.id = "vtx" + (++serial).toString(36);
    return atom.id;
  };

  /**
   * `editor`: a formula being edited, where an operator just typed keeps its
   * place (MathEnv.editorBinSpacing), the scale every atom is set at is
   * recorded for the caret (MathEnv.atomScales) and empty slots and rows get
   * a place for their caret (LatexAdapter.editor).
   */
  function layout(atoms: AtomLike[], display: boolean, editor = false) {
    syncLetterFonts();
    adapter.setSpacing(options.spacing?.() ?? null);
    const scales = editor ? new Map<Tag, number>() : undefined;
    adapter.editor = editor;
    adapter.env.editorBinSpacing = editor;
    adapter.env.atomScales = scales;
    try {
      const mlist = adapter.formula(atoms, display);
      const ts = new MathTypesetter(adapter.env);
      const hl = ts.mlistToHlist(mlist, TEXT_STYLE, false);
      const on: Node = { type: MATH_NODE, width: 0, subtype: 0 };
      const box = hpack([on, ...hl, { ...on }]);
      const ship = shipOut(box);
      return { box, ship, regions: computeRegions(ship, caret.strut), scales };
    } finally {
      adapter.editor = false;
      adapter.env.editorBinSpacing = false;
      adapter.env.atomScales = undefined;
    }
  }

  function fontFamily(f: { name: string }): string {
    return familyOf.get(f.name) ?? "serif";
  }

  /** The element whose children include the element of `atom` (the box of its branch). */
  function containerOf(elements: FormulaElement[], atom: KernelAtom): FormulaElement | null {
    const visit = (e: FormulaElement): FormulaElement | null => {
      for (const c of e.children ?? []) {
        if (c.tag === atom) return e;
        const found = visit(c);
        if (found) return found;
      }
      return null;
    };
    for (const e of elements) {
      const found = visit(e);
      if (found) return found;
    }
    return null;
  }

  /** Contour ovals over the integral signs of \oiiint atoms (their own glyphs). */
  function contourOvals(ship: ShipResult) {
    const out: NonNullable<Parameters<typeof formulaHtml>[2]["ovals"]> = [];
    for (const tag of adapter.ovals) {
      const signs = ship.glyphs.filter((g) => g.tag === tag);
      if (signs.length === 0) continue;
      const box = union(
        signs.map((g) => {
          const i = g.font.charInfo(g.char);
          return { x0: g.x, y0: g.y - g.font.charHeight(i), x1: g.x + g.font.charWidth(i), y1: g.y + g.font.charDepth(i) };
        }),
      );
      const cx = (box.x0 + box.x1) / 2, cy = (box.y0 + box.y1) / 2;
      const rx = (box.x1 - box.x0) * 0.36, ry = (box.y1 - box.y0) * 0.16;
      out.push({ x0: cx - rx, x1: cx + rx, y0: cy - ry, y1: cy + ry, thickness: Math.round(signs[0].font.size * 0.04), tag });
    }
    return out;
  }

  const isPlaceholder = (tag: Tag) => !!tag && typeof tag === "object" && (tag as KernelAtom).type === "placeholder";

  /** Scripts set beside their nucleus (not limits above or below an operator). */
  function besideNucleus(own: Rect[], others: FormulaElement[], scripts: FormulaElement[]): boolean {
    const nucleus = own.length ? union(own) : others.length ? union(others) : null;
    if (!nucleus) return true;
    return scripts.every((s) => s.x0 >= nucleus.x1 - 2 * 65536);
  }

  /**
   * Atom elements, nested as the atoms are. Editor markup has one per atom
   * (with data-atom-id); static markup only those that CSS styles.
   */
  function atomElements(roots: KernelAtom[], all: Map<Tag, Region>, ship: ShipResult, editor: boolean): FormulaElement[] {
    const ink = ownInk(ship);
    const seen = new Set<KernelAtom>();
    const build = (a: KernelAtom): FormulaElement[] => {
      if (seen.has(a)) return [];
      seen.add(a);
      const kids: FormulaElement[] = [];
      const scripts: FormulaElement[] = [];
      // each branch and each array cell has a box of its own, as in MathLive's
      // markup (the caret goes into the box of its branch)
      const branchBox = (list: FormulaElement[]): FormulaElement[] =>
        editor && list.length > 0 ? [{ ...union(list), classes: "ML__vtx-branch", children: list }] : list;
      for (const name of a.branches ?? []) {
        const list: FormulaElement[] = [];
        for (const c of (a.branch(name) ?? []) as KernelAtom[]) list.push(...build(c));
        (name === "superscript" || name === "subscript" ? scripts : kids).push(...branchBox(list));
      }
      if (a.rows)
        for (const row of a.rows)
          for (const cell of row) {
            const list: FormulaElement[] = [];
            for (const c of cell as KernelAtom[]) list.push(...build(c));
            kids.push(...branchBox(list));
          }
      // a rule inside a phantom draws nothing: its box is the phantom's
      const r = all.get(a) ?? (a.command === "\\rule" && a.parent?.type === "phantom" ? all.get(a.parent) : undefined);
      if (scripts.length > 0) {
        // MathLive wraps scripts beside the nucleus in an ML__msubsup box
        // (a subsup atom is that box itself)
        if (a.type !== "subsup" && r && besideNucleus(ink.get(a)?.glyphs ?? [], kids, scripts))
          kids.push({ ...union(scripts), classes: "ML__msubsup", children: scripts });
        else kids.push(...scripts);
      }
      if (!r || (!editor && !hasStaticElement(a, adapter))) return kids; // the children go to the enclosing atom
      const el: FormulaElement = {
        x0: r.x0,
        y0: r.y0,
        x1: r.x1,
        y1: r.y1,
        id: editor ? idOf(a) : null,
        tag: a,
        classes: [structuralClass(a, adapter), a.isSelected ? "ML__selected" : ""].filter(Boolean).join(" "),
        children: [...inkParts(a, ink, r), ...kids],
      };
      if (a.type === "placeholder") el.scale = r.scale;
      if (editor && a.type === "latex") el.text = a.value ?? "";
      // the symbol itself, as hidden text (textContent as in MathLive's markup)
      else if (editor && !a.branches?.length && !a.rows && a.type !== "placeholder" && a.type !== "first" && a.value && a.value.length <= 8)
        el.text = a.value;
      return [el];
    };
    return roots.flatMap(build);
  }

  /** Static markup (previews, tiles, palettes): no caret, no atom ids. */
  function renderAtoms(atoms: AtomLike[], opts: { display?: boolean } = {}): string {
    const { box, ship, regions } = layout(atoms, opts.display !== false);
    propagateRegions(regions, null);
    return formulaHtml(box, ship, {
      fontFamily,
      ovals: contourOvals(ship),
      atomColor: (tag) => (tag as KernelAtom | undefined)?.style?.color,
      hideGlyph: isPlaceholder,
      elements: atomElements(atoms as KernelAtom[], regions, ship, false),
    });
  }

  /** Editor markup for a mathfield (kernel flags isSelected/containsCaret are already set). */
  function renderMathfield(mathfield: any, _renderOptions?: unknown): string {
    const model = mathfield.model;
    const root = model.root as KernelAtom;
    const atoms = root.branch("body") ?? [];
    const display = mathfield.options?.defaultMode !== "inline-math";
    const { box, ship, regions, scales } = layout(atoms, display, true);
    propagateRegions(regions, root);
    // the root spans the whole formula (the kernel's hit-testing starts there)
    const empty = box.width === 0 && box.height === 0 && box.depth === 0;
    regions.set(root, {
      tag: root,
      x0: 0,
      x1: box.width,
      y0: empty ? -7.5 * 65536 : -box.height,
      y1: empty ? 2.5 * 65536 : box.depth,
      baseline: 0,
      scale: 1,
      hasGlyph: false,
    });
    // explicit alignment markers on the main row (the editor aligns rows on them)
    const markerAtoms = atoms.filter((a) => htmlClassOf(a as KernelAtom) === "visualtex-align-marker") as KernelAtom[];
    const alignMarkers = markerAtoms.map((a) => regions.get(a)?.x0 ?? 0);
    const geometry = buildEditorGeometry(model, regions, ship, box, alignMarkers, scales, caret);
    mathfield.visualTexGeometry = geometry;
    const state: CaretState = options.caretState?.(mathfield) ?? { kind: "none" };

    const elements = atomElements([root], regions, ship, true);
    const rootEl = elements[0];
    const body = rootEl?.children?.find((c) => c.classes === "ML__vtx-branch") ?? rootEl;
    if (alignMarkers.length > 0) {
      rootEl.extraWidth = body.extraWidth = alignShift(alignMarkers, Infinity);
      for (const child of body.children ?? []) {
        const k = markerAtoms.findIndex((m) => m.id !== undefined && m.id === child.id);
        child.shift = alignShift(alignMarkers, child.x0, k >= 0 ? k : undefined);
      }
    }
    // the caret goes into the element of the slot that holds it
    const stop = state.kind === "stop" ? geometry.stops[state.offset] : null;
    const emptySlot = state.kind === "empty" ? geometry.markers.get(state.placeholder) : undefined;
    const bar: CaretBar | null = stop
      ? { ...stop, mode: model.mode === "text" || model.mode === "latex" ? model.mode : "math" }
      : emptySlot && state.kind === "empty"
        ? { x: emptySlot.x0, top: emptySlot.y0, bottom: emptySlot.y1, scale: regions.get(state.placeholder)?.scale ?? 1, side: "left", mode: "math" }
        : null;
    let rootCaret: CaretBar | null = null;
    if (bar) {
      let anchor = state.kind === "stop" ? (model.at(state.offset) as KernelAtom) : state.kind === "empty" ? (state.placeholder as KernelAtom) : null;
      // a slot's start has no element of its own: the caret goes into the box
      // of its slot, found through the slot's first item
      if (anchor?.type === "first") {
        const items = (anchor.parent?.branch(anchor.parentBranch as never) ?? []) as KernelAtom[];
        anchor = items.find((x) => x.type !== "first" && regions.has(x)) ?? anchor;
      }
      const container = anchor ? containerOf(elements, anchor) : null;
      if (container) {
        const shift = (container === rootEl || container === body) && alignMarkers.length > 0 ? alignShift(alignMarkers, bar.x) : null;
        container.inner = (container.inner ?? "") + caretBarHtml(bar, container.x0, container.y0, shift);
      } else rootCaret = bar;
    }
    const markers = [...geometry.markers].map(([tag, r]) => ({ ...r, selected: state.kind === "empty" && state.placeholder === tag }));
    return formulaHtml(box, ship, {
      fontFamily,
      alignMarkers,
      ovals: contourOvals(ship),
      markers,
      atomClasses: (tag) => ((tag as KernelAtom | undefined)?.isSelected ? "ML__selected" : ""),
      // the side of a bracket pair the editor still waits for ("?": the
      // closing bracket of a bracket just typed) is faded
      glyphClasses: (g) => {
        const a = g.tag as KernelAtom | undefined;
        if (a?.type !== "leftright" || (a.leftDelim !== "?" && a.rightDelim !== "?")) return "";
        const r = regions.get(a);
        if (!r) return "";
        const right = g.x >= (r.x0 + r.x1) / 2;
        return (right ? a.rightDelim : a.leftDelim) === "?" ? "ML__smart-fence__close" : "";
      },
      atomColor: (tag) => {
        const a = tag as KernelAtom | undefined;
        if (!a || typeof a !== "object") return undefined;
        if (a.type === "latex" || a.mode === "latex") return a.isSuggestion ? "var(--neutral-500, #9e9e9e)" : "var(--_latex-color)";
        // the fences and the radical of the structure holding the caret
        if (a.containsCaret && (a.type === "leftright" || a.type === "surd")) return "var(--_contains-highlight-color)";
        return a.style?.color;
      },
      hideGlyph: isPlaceholder,
      elements,
      caret: rootCaret,
      extent: { height: geometry.height, depth: geometry.depth, right: geometry.right },
      rootClass: "ML__latex ML__vtx" + (mathfield.hasEditablePrompts ? " ML__prompting" : ""),
    });
  }

  /** Every atom of the formula, parents before their children. */
  function eachAtom(roots: KernelAtom[], visit: (a: KernelAtom) => void): void {
    const seen = new Set<KernelAtom>();
    const walk = (a: KernelAtom) => {
      if (seen.has(a)) return;
      seen.add(a);
      visit(a);
      for (const name of a.branches ?? []) for (const c of (a.branch(name) ?? []) as KernelAtom[]) walk(c);
      if (a.rows) for (const row of a.rows) for (const cell of row) for (const c of cell as KernelAtom[]) walk(c);
    };
    roots.forEach(walk);
  }

  /**
   * The formula as SVG (exported, copied and Office images): the same layout
   * as the editor's, drawn with the fonts' outlines. `classArtwork` draws
   * over the box of a \class{...}{...} annotation (custom symbols); `unknown`
   * lists what the engine could not typeset (unknown commands, raw LaTeX).
   */
  function renderSvg(
    atoms: AtomLike[],
    opts: {
      display?: boolean;
      pxPerEm: number;
      glyphs(tfm: string): GlyphTable | undefined;
      color?: string;
      idPrefix?: string;
      classArtwork?(className: string): ((width: number, height: number) => string) | null;
    },
  ): SvgFormula & { unknown: string[]; placeholders: number } {
    const { box, ship, regions } = layout(atoms, opts.display !== false);
    propagateRegions(regions, null);
    const overlays: SvgOverlay[] = [];
    const unknown: string[] = [];
    let placeholders = 0;
    eachAtom(atoms as KernelAtom[], (a) => {
      if (a.type === "placeholder") placeholders++;
      else if (a.type === "latex") unknown.push(a.value ?? "");
      else if (a.type === "error" && !(a.command && adapter.knowsCommand(a.command))) unknown.push(a.command ?? a.value ?? "");
      const cls = htmlClassOf(a);
      const art = cls ? opts.classArtwork?.(cls) : null;
      if (!art) return;
      // a rule inside a phantom (a custom symbol's metric box) has the phantom's box
      let r = regions.get(a);
      if (!r) eachAtom([a], (c) => (r ??= regions.get(c)));
      if (!r) return;
      let color: string | undefined;
      eachAtom([a], (c) => (color ??= c.style?.color));
      overlays.push({ x0: r.x0, y0: r.y0, x1: r.x1, y1: r.y1, color, markup: art });
    });
    const svg = formulaSvgBody(box, ship, {
      pxPerEm: opts.pxPerEm,
      glyphs: opts.glyphs,
      color: opts.color,
      idPrefix: opts.idPrefix,
      atomColor: (tag) => (tag as KernelAtom | undefined)?.style?.color,
      hideGlyph: isPlaceholder,
      ovals: contourOvals(ship),
      overlays,
    });
    return { ...svg, unknown, placeholders };
  }

  return { layout, renderAtoms, renderMathfield, renderSvg, adapter };
}

export type VisualTexLayout = ReturnType<typeof createVisualTexLayout>;
