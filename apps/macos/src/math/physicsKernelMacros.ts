/**
 * Non-conflicting physics.sty 1.3 commands for MathLive's default parser.
 *
 * These definitions cover the literal braced signatures shown here. The
 * package's starred, optional-order, dynamic-size and generated-matrix forms
 * need a variable-arity parser and are deliberately not approximated by a
 * fixed-arity macro. In particular, never redefine TeX's division \div.
 */
export interface VisualTexPhysicsKernelMacro {
  name: string;
  def: string;
  args: number;
}

const entry = (name: string, def: string, args = 0): VisualTexPhysicsKernelMacro => ({
  name,
  def,
  args,
});

export const VISUALTEX_PHYSICS_KERNEL_MACROS: readonly VisualTexPhysicsKernelMacro[] =
  Object.freeze([
    // Symbols and the package's explicit spelling for the original \div.
    entry("varE", "\\mathcal{E}"),
    entry("ordersymbol", "\\mathcal{O}"),
    entry("divisionsymbol", "\\div"),

    // Fixed braced delimiter and matrix forms. \qty is shared with siunitx,
    // so only physics' unambiguous long name is included.
    entry("quantity", "\\left\\{#1\\right\\}", 1),
    entry("pqty", "\\left(#1\\right)", 1),
    entry("bqty", "\\left[#1\\right]", 1),
    entry("Bqty", "\\left\\{#1\\right\\}", 1),
    entry("vqty", "\\left\\lvert#1\\right\\rvert", 1),
    entry("pmqty", "\\begin{pmatrix}#1\\end{pmatrix}", 1),
    entry("Pmqty", "\\left\\lgroup\\begin{matrix}#1\\end{matrix}\\right\\rgroup", 1),
    entry("bmqty", "\\begin{bmatrix}#1\\end{bmatrix}", 1),
    entry("vmqty", "\\begin{vmatrix}#1\\end{vmatrix}", 1),
    entry("matrixquantity", "\\begin{matrix}#1\\end{matrix}", 1),
    entry("mqty", "\\begin{matrix}#1\\end{matrix}", 1),
    entry("matrixdeterminant", "\\begin{vmatrix}#1\\end{vmatrix}", 1),
    entry("mdet", "\\begin{vmatrix}#1\\end{vmatrix}", 1),
    entry("spmqty", "\\left(\\begin{smallmatrix}#1\\end{smallmatrix}\\right)", 1),
    entry("sPmqty", "\\left\\lgroup\\begin{smallmatrix}#1\\end{smallmatrix}\\right\\rgroup", 1),
    entry("sbmqty", "\\left[\\begin{smallmatrix}#1\\end{smallmatrix}\\right]", 1),
    entry("svmqty", "\\left\\lvert\\begin{smallmatrix}#1\\end{smallmatrix}\\right\\rvert", 1),
    entry("smallmatrixquantity", "\\begin{smallmatrix}#1\\end{smallmatrix}", 1),
    entry("smqty", "\\begin{smallmatrix}#1\\end{smallmatrix}", 1),
    entry("smallmatrixdeterminant", "\\left\\lvert\\begin{smallmatrix}#1\\end{smallmatrix}\\right\\rvert", 1),
    entry("smdet", "\\left\\lvert\\begin{smallmatrix}#1\\end{smallmatrix}\\right\\rvert", 1),
    entry("absolutevalue", "\\left\\lvert#1\\right\\rvert", 1),
    entry("order", "\\mathcal{O}\\left(#1\\right)", 1),
    entry("evaluated", "\\left.#1\\right\\rvert", 1),
    entry("eval", "\\left.#1\\right\\rvert", 1),
    entry("poissonbracket", "\\left\\{#1,#2\\right\\}", 2),
    entry("commutator", "\\left[#1,#2\\right]", 2),
    entry("anticommutator", "\\left\\{#1,#2\\right\\}", 2),
    entry("acommutator", "\\left\\{#1,#2\\right\\}", 2),

    // Vector and differential operators. The existing \div retains ÷.
    entry("vectorbold", "\\mathbf{#1}", 1),
    entry("vectorarrow", "\\vec{\\mathbf{#1}}", 1),
    entry("vectorunit", "\\mathbf{\\hat{#1}}", 1),
    entry("dotproduct", "\\boldsymbol{\\cdot}"),
    entry("vdot", "\\boldsymbol{\\cdot}"),
    entry("crossproduct", "\\boldsymbol{\\times}"),
    entry("cross", "\\boldsymbol{\\times}"),
    entry("cp", "\\boldsymbol{\\times}"),
    entry("vnabla", "\\boldsymbol{\\nabla}"),
    entry("gradient", "\\boldsymbol{\\nabla}"),
    entry("grad", "\\boldsymbol{\\nabla}"),
    entry("divergence", "\\boldsymbol{\\nabla}\\boldsymbol{\\cdot}"),
    entry("curl", "\\boldsymbol{\\nabla}\\boldsymbol{\\times}"),
    entry("laplacian", "\\nabla^{2}"),

    // Operators not already defined by TeX/MathLive.
    entry("trace", "\\operatorname{tr}"),
    entry("Trace", "\\operatorname{Tr}"),
    entry("rank", "\\operatorname{rank}"),
    entry("erf", "\\operatorname{erf}"),
    entry("Residue", "\\operatorname{Res}"),
    entry("tr", "\\operatorname{tr}"),
    entry("Tr", "\\operatorname{Tr}"),
    entry("Res", "\\operatorname{Res}"),
    entry("principalvalue", "\\mathcal{P}"),
    entry("pv", "\\mathcal{P}"),
    entry("PV", "\\mathrm{P.V.}"),
    entry("imaginary", "\\Im"),
    entry("arccot", "\\operatorname{arccot}"),
    entry("arccosecant", "\\operatorname{arccsc}"),
    entry("arcsecant", "\\operatorname{arcsec}"),
    entry("arccotangent", "\\operatorname{arccot}"),
    entry("asin", "\\operatorname{asin}"),
    entry("asine", "\\operatorname{asin}"),
    entry("acos", "\\operatorname{acos}"),
    entry("acosine", "\\operatorname{acos}"),
    entry("atan", "\\operatorname{atan}"),
    entry("atangent", "\\operatorname{atan}"),
    entry("acsc", "\\operatorname{acsc}"),
    entry("acosecant", "\\operatorname{acsc}"),
    entry("asec", "\\operatorname{asec}"),
    entry("asecant", "\\operatorname{asec}"),
    entry("acot", "\\operatorname{acot}"),
    entry("acotangent", "\\operatorname{acot}"),
    entry("csch", "\\operatorname{csch}"),
    entry("hypcosecant", "\\operatorname{csch}"),
    entry("sech", "\\operatorname{sech}"),
    entry("hypsecant", "\\operatorname{sech}"),

    // Math-mode text with physics' quad spacing.
    entry("qqtext", "\\quad\\text{#1}\\quad", 1),
    entry("qq", "\\quad\\text{#1}\\quad", 1),
    entry("qcomma", ",\\quad"),
    entry("qc", ",\\quad"),
    ...[
      "if", "then", "else", "otherwise", "unless", "given", "using",
      "assume", "since", "let", "for", "all", "even", "odd",
      "integer", "and", "or", "as", "in", "c.c.",
    ].map((word) => entry(
      word === "c.c." ? "qcc" : `q${word}`,
      `\\quad\\text{${word}}\\quad`,
    )),

    // physics.sty uses function-first, variable-second in its two-argument
    // derivative forms. Existing VisualTeX \dv and \pdv stay untouched.
    entry("differential", "\\mathrm{d}"),
    entry("derivative", "\\frac{\\mathrm{d}#1}{\\mathrm{d}#2}", 2),
    entry("partialderivative", "\\frac{\\partial #1}{\\partial #2}", 2),
    entry("pderivative", "\\frac{\\partial #1}{\\partial #2}", 2),
    entry("variation", "\\delta"),
    entry("var", "\\delta"),
    entry("functionalderivative", "\\frac{\\delta #1}{\\delta #2}", 2),
    entry("fderivative", "\\frac{\\delta #1}{\\delta #2}", 2),
    entry("fdv", "\\frac{\\delta #1}{\\delta #2}", 2),

    // Dirac notation variants not already supplied by VisualTeX.
    entry("innerproduct", "\\left\\langle#1\\middle\\vert#2\\right\\rangle", 2),
    entry("ip", "\\left\\langle#1\\middle\\vert#2\\right\\rangle", 2),
    entry("outerproduct", "\\left\\lvert#1\\right\\rangle\\!\\left\\langle#2\\right\\rvert", 2),
    entry("dyad", "\\left\\lvert#1\\right\\rangle\\!\\left\\langle#2\\right\\rvert", 2),
    entry("op", "\\left\\lvert#1\\right\\rangle\\!\\left\\langle#2\\right\\rvert", 2),
    entry("expectationvalue", "\\left\\langle#2\\middle\\vert#1\\middle\\vert#2\\right\\rangle", 2),
    entry("ev", "\\left\\langle#1\\right\\rangle", 1),
    entry("vev", "\\left\\langle0\\middle\\vert#1\\middle\\vert0\\right\\rangle", 1),
    entry("matrixelement", "\\left\\langle#1\\middle\\vert#2\\middle\\vert#3\\right\\rangle", 3),
    entry("matrixel", "\\left\\langle#1\\middle\\vert#2\\middle\\vert#3\\right\\rangle", 3),
    entry("flatfrac", "\\left.#1\\middle\\slash#2\\right.", 2),
  ]);
