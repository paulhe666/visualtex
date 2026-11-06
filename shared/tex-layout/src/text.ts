// Text in restricted horizontal mode: TeX's main loop for characters
// (tex.web 1034-1044) with ligatures, kerns, boundary characters, inter-word
// glue and space factors. No discretionaries (restricted horizontal mode).
import { xnOverD } from "./arith.ts";
import { GLUE_NODE, KERN_NODE, NORMAL, newCharacter, newKern, type GlueSpec, type Node, type Tag } from "./nodes.ts";
import { KERN_FLAG, LIG_TAG, STOP_FLAG, type LigKern, type TeXFont } from "./tfm.ts";

const NON_CHAR = 256;

/** \sfcode values under LaTeX's \nonfrenchspacing. */
export function latexSfCode(c: number): number {
  if (c >= 65 && c <= 90) return 999; // uppercase letters
  switch (String.fromCharCode(c)) {
    case ".":
    case "?":
    case "!":
      return 3000;
    case ":":
      return 2000;
    case ";":
      return 1500;
    case ",":
      return 1250;
    case ")":
    case "'":
    case "]":
      return 0;
    default:
      return 1000;
  }
}

interface BoundaryInfo {
  bchar: number;
  falseBchar: number;
  bcharLabel: number; // index into ligKern or -1
}

const boundaryCache = new WeakMap<TeXFont, BoundaryInfo>();

/** font_bchar, font_false_bchar, bchar_label (tex.web 573, 576). */
function boundary(f: TeXFont): BoundaryInfo {
  let b = boundaryCache.get(f);
  if (b) return b;
  let bchar = 256;
  let bchLabel = 0o77777;
  const lk = f.ligKern;
  if (lk.length > 0) {
    if (lk[0].skip === 255) bchar = lk[0].next;
    const last = lk[lk.length - 1];
    if (last.skip === 255) bchLabel = 256 * last.op + last.rem;
  }
  let falseBchar = bchar;
  if (bchar <= f.ec && bchar >= f.bc && f.charInfo(bchar).exists) falseBchar = NON_CHAR;
  b = { bchar, falseBchar, bcharLabel: bchLabel < lk.length ? bchLabel : -1 };
  boundaryCache.set(f, b);
  return b;
}

/** A character of input text, or a space. */
export type TextToken = { kind: "char"; code: number; tag?: Tag } | { kind: "space"; tag?: Tag };

/**
 * Typeset text tokens in one font, as TeX's main loop would in restricted
 * horizontal mode. Returns the horizontal list (chars, ligature chars, kerns, glue).
 * A ligature is represented by its ligature character (a char node), which is
 * how it appears in the output.
 */
export function typesetText(tokens: TextToken[], f: TeXFont, spaceFactorStart = 1000, fallback?: TeXFont | null): Node[] {
  const out: Node[] = [];
  let spaceFactor = spaceFactorStart;
  let i = 0;
  const adjustSpaceFactor = (c: number) => {
    const s = latexSfCode(c);
    if (s === 1000) spaceFactor = 1000;
    else if (s < 1000) {
      if (s > 0) spaceFactor = s;
    } else if (spaceFactor < 1000) spaceFactor = 1000;
    else spaceFactor = s;
  };
  while (i < tokens.length) {
    const t = tokens[i];
    if (t.kind === "space") {
      out.push({ type: GLUE_NODE, spec: spaceGlue(f, spaceFactor), subtype: NORMAL, tag: t.tag });
      i++;
      continue;
    }
    // A run of characters in one font: the ligature/kern main loop; characters
    // the TeX font lacks go to the fallback (native) font, without ligatures.
    const inFont = (c: number) => c >= f.bc && c <= f.ec && f.charInfo(c).exists;
    const t0 = tokens[i] as { kind: "char"; code: number; tag?: Tag };
    const native = !inFont(t0.code) && !!fallback;
    const run: TextToken[] = [];
    while (i < tokens.length && tokens[i].kind === "char" && (!fallback || inFont((tokens[i] as { code: number }).code) !== native)) run.push(tokens[i++]);
    if (native) {
      for (const t of run as { kind: "char"; code: number; tag?: Tag }[]) {
        adjustSpaceFactor(t.code);
        if (fallback!.charInfo(t.code).exists) out.push(newCharacter(fallback!, t.code, t.tag));
      }
      continue;
    }
    spaceFactor = ligKernRun(run as { kind: "char"; code: number; tag?: Tag }[], f, out, spaceFactor, adjustSpaceFactor);
  }
  return out;
}

/** tex.web 1042-1044: inter-word glue for the current space factor. */
function spaceGlue(f: TeXFont, spaceFactor: number): GlueSpec {
  const g: GlueSpec = { width: f.param(2), stretch: f.param(3), stretchOrder: NORMAL, shrink: f.param(4), shrinkOrder: NORMAL };
  if (spaceFactor === 1000) return g;
  if (spaceFactor >= 2000) g.width += f.param(7);
  g.stretch = xnOverD(g.stretch, spaceFactor, 1000);
  g.shrink = xnOverD(g.shrink, 1000, spaceFactor);
  return g;
}

/** An entry of TeX's lig_stack: a character node or a lig item (pseudo-ligature). */
interface LigItem {
  isChar: boolean;
  char: number;
  /** lig_ptr of a lig item: the original character it stands for */
  ligPtr?: { char: number; tag?: Tag } | null;
  tag?: Tag;
}

/**
 * The inner ligature/kern loop for one run of characters (tex.web 1034-1040),
 * as a literal state machine over TeX's labels. `out` is the current list;
 * a ligature replaces the characters after `curQ` by one char node.
 */
function ligKernRun(
  run: { code: number; tag?: Tag }[],
  f: TeXFont,
  out: Node[],
  spaceFactor: number,
  adjust: (c: number) => void,
): number {
  const { bchar: fontBchar, falseBchar, bcharLabel } = boundary(f);
  let bchar = fontBchar;
  let pos = 0;
  let curL = 0;
  let curR = 0;
  let curQ = out.length;
  let ligStack: LigItem[] = [];
  let ligaturePresent = false;
  let lftHit = false;
  let rtHit = false;
  let mainK = -1;
  let mainI = f.charInfo(0);
  let mainJ: LigKern = f.ligKern[0];
  let ligTag: Tag;

  const packLig = (z: boolean) => {
    // new_ligature(main_f, cur_l, link(cur_q)): replaces the chars after cur_q
    const first = out[curQ] as { tag?: Tag } | undefined;
    out.splice(curQ, out.length - curQ, newCharacter(f, curL, first?.tag ?? ligTag));
    if (lftHit) lftHit = false;
    if (z && ligStack.length === 0) rtHit = false;
    ligaturePresent = false;
  };
  const wrapup = (z: boolean) => {
    if (curL < NON_CHAR && ligaturePresent) packLig(z);
  };

  // <Append character cur_chr and the following characters (if any)...>
  const first = run[pos++];
  adjust(first.code);
  ligStack = [{ isChar: true, char: first.code, tag: first.tag }];
  curL = first.code;
  ligTag = first.tag;
  curQ = out.length;
  mainK = bcharLabel;
  type Label = "wrapup" | "move" | "move1" | "move2" | "lookahead" | "lig" | "lig1" | "lig2" | "movelig";
  let label: Label;
  if (mainK < 0) label = "move2";
  else {
    curR = curL;
    curL = NON_CHAR;
    label = "lig1";
  }
  for (;;) {
    switch (label) {
      case "wrapup":
        wrapup(rtHit);
        label = "move";
        break;
      case "move":
        if (ligStack.length === 0) return spaceFactor; // goto reswitch
        curQ = out.length;
        curL = ligStack[0].char;
        label = "move1";
        break;
      case "move1":
        label = ligStack[0].isChar ? "move2" : "movelig";
        break;
      case "move2": {
        if (curL < f.bc || curL > f.ec || !f.charInfo(curL).exists) {
          ligStack = []; // char_warning; the character is dropped
          return spaceFactor;
        }
        mainI = f.charInfo(curL);
        const top = ligStack.shift()!;
        out.push(newCharacter(f, curL, top.tag)); // link(tail):=lig_stack
        ligTag = top.tag;
        label = "lookahead";
        break;
      }
      case "lookahead":
        if (pos < run.length) {
          const c = run[pos++];
          adjust(c.code);
          ligStack = [{ isChar: true, char: c.code, tag: c.tag }];
          curR = c.code;
          if (curR === falseBchar) curR = NON_CHAR;
        } else {
          curR = bchar;
          ligStack = [];
        }
        label = "lig";
        break;
      case "lig":
        if (mainI.tag !== LIG_TAG || curR === NON_CHAR) {
          label = "wrapup";
          break;
        }
        mainK = f.ligKernStart(mainI);
        mainJ = f.ligKern[mainK];
        if (mainJ.skip <= STOP_FLAG) {
          label = "lig2";
          break;
        }
        mainK = f.ligKernRestart(mainJ);
        label = "lig1";
        break;
      case "lig1":
        mainJ = f.ligKern[mainK];
        label = "lig2";
        break;
      case "lig2": {
        const j = mainJ;
        if (j.next === curR && j.skip <= STOP_FLAG) {
          // <Do ligature or kern command> (1040)
          if (j.op >= KERN_FLAG) {
            wrapup(rtHit);
            out.push(newKern(f.charKern(j), ligTag));
            label = "move";
            break;
          }
          if (curL === NON_CHAR) lftHit = true;
          else if (ligStack.length === 0) rtHit = true;
          let next: Label | null = null;
          switch (j.op) {
            case 1:
            case 5:
              curL = j.rem;
              mainI = f.charInfo(curL);
              ligaturePresent = true;
              break;
            case 2:
            case 6:
              curR = j.rem;
              if (ligStack.length === 0) {
                ligStack = [{ isChar: false, char: curR, ligPtr: null }];
                bchar = NON_CHAR;
              } else if (ligStack[0].isChar) {
                const p = ligStack[0];
                ligStack = [{ isChar: false, char: curR, ligPtr: { char: p.char, tag: p.tag } }];
              } else ligStack[0].char = curR;
              break;
            case 3:
              curR = j.rem;
              ligStack.unshift({ isChar: false, char: curR, ligPtr: null });
              break;
            case 7:
            case 11:
              wrapup(false);
              curQ = out.length;
              curL = j.rem;
              mainI = f.charInfo(curL);
              ligaturePresent = true;
              break;
            default:
              curL = j.rem;
              ligaturePresent = true;
              next = ligStack.length === 0 ? "wrapup" : "move1";
              break;
          }
          if (next) {
            label = next;
            break;
          }
          if (j.op > 4 && j.op !== 7) {
            label = "wrapup";
            break;
          }
          if (curL < NON_CHAR) {
            label = "lig";
            break;
          }
          mainK = bcharLabel;
          label = "lig1";
          break;
        }
        if (j.skip === 0) mainK++;
        else {
          if (j.skip >= STOP_FLAG) {
            label = "wrapup";
            break;
          }
          mainK = mainK + j.skip + 1;
        }
        label = "lig1";
        break;
      }
      case "movelig": {
        // <Move the cursor past a pseudo-ligature> (1037)
        const top = ligStack.shift()!;
        const mainP = top.ligPtr ?? null;
        if (mainP) out.push(newCharacter(f, mainP.char, mainP.tag));
        mainI = f.charInfo(curL);
        ligaturePresent = true;
        if (ligStack.length === 0) {
          if (mainP) {
            label = "lookahead";
            break;
          }
          curR = bchar;
        } else curR = ligStack[0].char;
        label = "lig";
        break;
      }
    }
  }
}

export { KERN_NODE };
