// The physics package's matrix commands (\pmqty{...}, \mdet{...}, ...) and the
// environment each one wraps its argument in. Data only: the parser and the
// environment definitions both read it.
export const VISUALTEX_PHYSICS_MATRIX_WRAPPERS: Record<string, [open: string, close: string]> = {
  pmqty: ["\\begin{pmatrix}", "\\end{pmatrix}"],
  Pmqty: ["\\left\\lgroup\\begin{matrix}", "\\end{matrix}\\right\\rgroup"],
  bmqty: ["\\begin{bmatrix}", "\\end{bmatrix}"],
  vmqty: ["\\begin{vmatrix}", "\\end{vmatrix}"],
  matrixquantity: ["\\begin{matrix}", "\\end{matrix}"],
  mqty: ["\\begin{matrix}", "\\end{matrix}"],
  matrixdeterminant: ["\\begin{vmatrix}", "\\end{vmatrix}"],
  mdet: ["\\begin{vmatrix}", "\\end{vmatrix}"],
  spmqty: ["\\left(\\begin{smallmatrix}", "\\end{smallmatrix}\\right)"],
  sPmqty: ["\\left\\lgroup\\begin{smallmatrix}", "\\end{smallmatrix}\\right\\rgroup"],
  sbmqty: ["\\left[\\begin{smallmatrix}", "\\end{smallmatrix}\\right]"],
  svmqty: ["\\left\\lvert\\begin{smallmatrix}", "\\end{smallmatrix}\\right\\rvert"],
  smallmatrixquantity: ["\\begin{smallmatrix}", "\\end{smallmatrix}"],
  smqty: ["\\begin{smallmatrix}", "\\end{smallmatrix}"],
  smallmatrixdeterminant: ["\\left\\lvert\\begin{smallmatrix}", "\\end{smallmatrix}\\right\\rvert"],
  smdet: ["\\left\\lvert\\begin{smallmatrix}", "\\end{smallmatrix}\\right\\rvert"]
};
