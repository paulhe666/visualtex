import type { LatexCommand } from "../types/command";
import { VISUALTEX_PHYSICS_KERNEL_MACROS } from "../math/physicsKernelMacros";

type PhysicsTile = readonly [
  name: string,
  previewLatex: string,
  labelZh: string,
  labelEn: string,
  priority: number,
  insertTemplate?: string,
];

// Only expose distinct, non-conflicting physics.sty forms. Existing toolbar
// tiles already cover bra/ket, commutators, vector shorthands and ∇ operators.
const tiles: readonly PhysicsTile[] = [
  ["quantity", "\\quantity{x}", "自适应花括号", "Physics braces", 82],
  ["pqty", "\\pqty{x}", "物理圆括号", "Physics parentheses", 82],
  ["bqty", "\\bqty{x}", "物理方括号", "Physics brackets", 81],
  ["vqty", "\\vqty{x}", "物理绝对值", "Physics absolute value", 80],
  ["pmqty", "\\pmqty{a&b\\\\c&d}", "圆括号矩阵", "Physics parenthesized matrix", 79,
    "\\pmqty{\\placeholder{}&\\placeholder{}\\\\\\placeholder{}&\\placeholder{}}"],
  ["bmqty", "\\bmqty{a&b\\\\c&d}", "方括号矩阵", "Physics bracketed matrix", 78,
    "\\bmqty{\\placeholder{}&\\placeholder{}\\\\\\placeholder{}&\\placeholder{}}"],
  ["vmqty", "\\vmqty{a&b\\\\c&d}", "行列式矩阵", "Physics determinant matrix", 77,
    "\\vmqty{\\placeholder{}&\\placeholder{}\\\\\\placeholder{}&\\placeholder{}}"],
  ["spmqty", "\\spmqty{a&b}", "小圆括号矩阵", "Small physics matrix", 76,
    "\\spmqty{\\placeholder{}&\\placeholder{}}"],
  ["sbmqty", "\\sbmqty{a&b}", "小方括号矩阵", "Small bracketed matrix", 75,
    "\\sbmqty{\\placeholder{}&\\placeholder{}}"],
  ["order", "\\order{x}", "大 O 记号", "Order notation", 85],
  ["eval", "\\eval{x}_0", "代入求值", "Evaluate at", 84,
    "\\eval{\\placeholder{}}_{\\placeholder{}}"],
  ["dotproduct", "a\\dotproduct b", "向量点积", "Vector dot product", 88],
  ["crossproduct", "a\\crossproduct b", "向量叉积", "Vector cross product", 87],
  ["vnabla", "\\vnabla f", "粗体 Nabla", "Bold nabla", 83],
  ["principalvalue", "\\principalvalue\\int", "柯西主值", "Principal value", 77],
  ["Trace", "\\Trace", "大写迹", "Capital trace", 73],
  ["erf", "\\erf", "误差函数", "Error function", 72],
  ["Residue", "\\Residue", "留数算子", "Residue operator", 71],
  ["derivative", "\\derivative{f}{x}", "physics 导数", "Physics derivative", 86],
  ["partialderivative", "\\partialderivative{f}{x}", "physics 偏导", "Physics partial derivative", 85],
  ["functionalderivative", "\\functionalderivative{F}{f}", "泛函导数", "Functional derivative", 82],
  ["variation", "\\variation x", "变分符号", "Variation", 78],
  ["vev", "\\vev{A}", "真空期望值", "Vacuum expectation", 83],
  ["flatfrac", "\\flatfrac{a}{b}", "斜线分式", "Physics flat fraction", 74],
  ["varE", "\\varE", "花体电场 E", "Calligraphic E", 74],
  ["ordersymbol", "\\ordersymbol", "大 O 符号", "Order symbol", 70],
];

const kernelMacros = new Map(
  VISUALTEX_PHYSICS_KERNEL_MACROS.map((macro) => [macro.name, macro]),
);

export const physicsToolbarCommands: LatexCommand[] = tiles.map(
  ([name, previewLatex, labelZh, labelEn, defaultPriority, insertTemplate]) => {
    const macro = kernelMacros.get(name);
    if (!macro) throw new Error(`Physics toolbar command \\${name} is not supported by the kernel`);
    return {
      id: `physics-${name}`,
      command: `\\${name}`,
      insertTemplate: insertTemplate ??
        `\\${name}${"{\\placeholder{}}".repeat(macro.args)}`,
      previewLatex,
      labelZh,
      labelEn,
      aliases: [name],
      keywords: ["physics", "物理", name],
      category: "physics",
      defaultPriority,
      supportedInMathMode: true,
    };
  },
);
