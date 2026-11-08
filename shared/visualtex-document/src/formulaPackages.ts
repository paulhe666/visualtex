// The LaTeX packages for the editor's physics, braket and siunitx commands in
// an exported document, and the rewrites that keep each formula's meaning when
// packages that define the same commands are loaded together:
// - braket.sty, physics and siunitx are loaded in that order (the preamble
//   sorts its packages): physics redefines braket.sty's \bra, \ket, \braket,
//   and siunitx leaves \qty to physics (it says so in a warning);
// - with physics loaded, siunitx's \qty{<number>}{<unit>} is written \SI
//   (the same command in siunitx 3), braket.sty's \braket{\phi|\psi} as its
//   definition, and the plain LaTeX \div, \Re, \Im the editor draws (÷, ℜ, ℑ)
//   by physics' names for them (\divisionsymbol, \real, \imaginary).

/** physics.sty's commands that plain LaTeX and amsmath do not have. */
const PHYSICS_COMMANDS = [
  "quantity", "pqty", "bqty", "Bqty", "vqty", "absolutevalue", "abs", "norm", "order", "evaluated", "eval",
  "commutator", "comm", "anticommutator", "acommutator", "acomm", "poissonbracket", "pb",
  "vectorbold", "vb", "vectorarrow", "va", "vectorunit", "vu", "dotproduct", "vdot", "crossproduct", "cross", "cp",
  "gradient", "grad", "divergence", "curl", "laplacian", "vnabla",
  "trace", "Trace", "rank", "erf", "Residue", "principalvalue", "pv", "PV", "tr", "Tr", "Res",
  "arccsc", "arcsec", "arccot", "asin", "acos", "atan", "acsc", "asec", "acot", "csch", "sech",
  "qqtext", "qq", "qcomma", "qc", "qif", "qthen", "qelse", "qotherwise", "qunless", "qgiven", "qusing", "qassume",
  "qsince", "qlet", "qfor", "qall", "qeven", "qodd", "qinteger", "qand", "qor", "qas", "qin", "qcc",
  "differential", "dd", "derivative", "dv", "partialderivative", "pderivative", "pdv", "variation", "var",
  "functionalderivative", "fderivative", "fdv", "bra", "ket", "innerproduct", "ip", "outerproduct", "dyad", "op",
  "ketbra", "expectationvalue", "expval", "ev", "vev", "matrixelement", "matrixel", "mel",
  "matrixquantity", "mqty", "pmqty", "Pmqty", "bmqty", "vmqty", "smallmatrixquantity", "smqty", "spmqty", "sPmqty",
  "sbmqty", "svmqty", "matrixdeterminant", "mdet", "smallmatrixdeterminant", "smdet", "identitymatrix", "imat",
  "xmatrix", "xmat", "zeromatrix", "zmat", "paulimatrix", "pmat", "diagonalmatrix", "dmat", "antidiagonalmatrix",
  "admat", "flatfrac", "varE", "ordersymbol", "divisionsymbol",
];
const PHYSICS = new RegExp(`\\\\(?:${PHYSICS_COMMANDS.join("|")})(?![A-Za-z])`);
const SIUNITX = /\\(?:SI|si|num|unit|ang|numlist|numrange|numproduct|qtylist|qtyrange|qtyproduct|SIlist|SIrange|complexnum|complexqty)(?![A-Za-z])/;
const BRAKET_STY = /\\(?:Braket|Set|set|Bra|Ket)(?![A-Za-z])/;

/** The text of the brace group at `start` and the index after it. */
function braceGroup(source: string, start: number): { content: string; end: number } | null {
  let index = start;
  while (/\s/.test(source[index] ?? "")) index += 1;
  if (source[index] !== "{") return null;
  let depth = 0;
  for (let i = index; i < source.length; i += 1) {
    if (source[i] === "\\") {
      i += 1;
      continue;
    }
    if (source[i] === "{") depth += 1;
    else if (source[i] === "}" && --depth === 0) return { content: source.slice(index + 1, i), end: i + 1 };
  }
  return null;
}

function hasTopLevelBar(text: string): boolean {
  let depth = 0;
  for (let i = 0; i < text.length; i += 1) {
    if (text[i] === "\\") {
      i += 1;
      continue;
    }
    if (text[i] === "{") depth += 1;
    else if (text[i] === "}") depth -= 1;
    else if (text[i] === "|" && depth === 0) return true;
  }
  return false;
}

const SIUNITX_NUMBER = /^(?:\\pm|\+-|[-+])?\d*(?:[.,]\d*)?(?:\(\d+(?:[.,]\d+)?\))?(?:(?:\\pm|\+-)\d*(?:[.,]\d*)?)?(?:[eEdD][-+]?\d+)?$/;

/**
 * Each \qty and \braket, with what it is (siunitx's \qty, braket.sty's
 * \braket, or physics' forms), and the formula with siunitx's \qty written
 * \SI (`qtyAsSI`) and braket.sty's \braket written as its definition.
 */
type Scan = { siunitx: boolean; braketSty: boolean; physicsForms: boolean; out: string };

function scanQtyAndBraket(latex: string, qtyAsSI: boolean): Scan {
  let siunitx = false;
  let braketSty = false;
  let physicsForms = false;
  let out = "";
  let copied = 0;
  const pattern = /\\(qty|braket)(?![A-Za-z])/g;
  for (let m = pattern.exec(latex); m; m = pattern.exec(latex)) {
    let cursor = pattern.lastIndex;
    if (m[1] === "qty") {
      const options = /^\s*\[[^\]]*\]/.exec(latex.slice(cursor));
      if (options) cursor += options[0].length;
      const number = braceGroup(latex, cursor);
      const unit = number && braceGroup(latex, number.end);
      if (number && unit && /\d/.test(number.content) && SIUNITX_NUMBER.test(number.content.replace(/\s+/g, ""))) {
        siunitx = true;
        if (qtyAsSI) {
          out += latex.slice(copied, m.index) + "\\SI";
          copied = pattern.lastIndex;
        }
      } else physicsForms = true;
    } else {
      const first = braceGroup(latex, cursor);
      if (first && !braceGroup(latex, first.end) && hasTopLevelBar(first.content)) {
        braketSty = true;
        const inner = scanQtyAndBraket(first.content, qtyAsSI);
        physicsForms ||= inner.physicsForms;
        siunitx ||= inner.siunitx;
        out += latex.slice(copied, m.index) + `\\mathinner{\\langle{${inner.out}}\\rangle}`;
        copied = first.end;
        pattern.lastIndex = first.end;
      } else physicsForms = true;
    }
  }
  out += latex.slice(copied);
  return { siunitx, braketSty, physicsForms, out };
}

/** The packages a formula's physics, braket and siunitx commands need. */
export function formulaPackageNeeds(latex: string): string[] {
  const { siunitx, braketSty, physicsForms } = scanQtyAndBraket(latex, false);
  const packages: string[] = [];
  if (physicsForms || PHYSICS.test(latex)) packages.push("physics");
  if (braketSty || BRAKET_STY.test(latex)) packages.push("braket");
  if (siunitx || SIUNITX.test(latex)) packages.push("siunitx");
  return packages;
}

/** A formula as it keeps its meaning with the document's packages (see the top of this file). */
export function formulaForPackages(latex: string, packages: ReadonlySet<string>): string {
  if (!packages.has("physics")) return latex;
  return scanQtyAndBraket(latex, packages.has("siunitx"))
    .out.replace(/\\div(?![A-Za-z])/g, "\\divisionsymbol")
    .replace(/\\Re(?![A-Za-z])/g, "\\real")
    .replace(/\\Im(?![A-Za-z])/g, "\\imaginary");
}
