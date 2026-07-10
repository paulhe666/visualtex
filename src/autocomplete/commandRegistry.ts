import type { LatexCommand } from "../types/command";
import { additionalCommands } from "./additionalCommands";

const baseCommandRegistry: LatexCommand[] = [
  { id: "frac", command: "\\frac", insertTemplate: "\\frac{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\frac{a}{b}", labelZh: "分式", labelEn: "Fraction", aliases: ["divide", "fraction"], keywords: ["分数", "除法"], category: "structure", defaultPriority: 100, supportedInMathMode: true },
  { id: "sqrt", command: "\\sqrt", insertTemplate: "\\sqrt{\\placeholder{}}", previewLatex: "\\sqrt{x}", labelZh: "平方根", labelEn: "Square root", aliases: ["root", "square root"], keywords: ["根号", "开方"], category: "structure", defaultPriority: 98, supportedInMathMode: true },
  { id: "nthroot", command: "\\sqrt", insertTemplate: "\\sqrt[\\placeholder{}]{\\placeholder{}}", previewLatex: "\\sqrt[n]{x}", labelZh: "n 次根", labelEn: "Nth root", aliases: ["nth root"], keywords: ["根式"], category: "structure", defaultPriority: 82, supportedInMathMode: true },
  { id: "power", command: "^", insertTemplate: "^{\\placeholder{}}", previewLatex: "x^{n}", labelZh: "上标", labelEn: "Superscript", aliases: ["power", "superscript"], keywords: ["次方", "指数"], category: "structure", defaultPriority: 90, supportedInMathMode: true },
  { id: "subscript", command: "_", insertTemplate: "_{\\placeholder{}}", previewLatex: "x_{i}", labelZh: "下标", labelEn: "Subscript", aliases: ["index", "subscript"], keywords: ["角标"], category: "structure", defaultPriority: 88, supportedInMathMode: true },
  { id: "parentheses", command: "\\left", insertTemplate: "\\left(\\placeholder{}\\right)", previewLatex: "\\left(x\\right)", labelZh: "自适应圆括号", labelEn: "Parentheses", aliases: ["bracket", "parentheses"], keywords: ["括号", "圆括号"], category: "structure", defaultPriority: 84, supportedInMathMode: true },
  { id: "absolute", command: "\\left|", insertTemplate: "\\left|\\placeholder{}\\right|", previewLatex: "\\left|x\\right|", labelZh: "绝对值", labelEn: "Absolute value", aliases: ["abs", "absolute"], keywords: ["模", "绝对值"], category: "structure", defaultPriority: 83, supportedInMathMode: true },
  { id: "binom", command: "\\binom", insertTemplate: "\\binom{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\binom{n}{k}", labelZh: "二项式", labelEn: "Binomial", aliases: ["choose", "binomial"], keywords: ["组合数"], category: "structure", defaultPriority: 72, supportedInMathMode: true },

  { id: "int", command: "\\int", insertTemplate: "\\int_{\\placeholder{}}^{\\placeholder{}} \\placeholder{}\\,\\mathrm{d}\\placeholder{}", previewLatex: "\\int_a^b f(x)\\,\\mathrm{d}x", labelZh: "定积分", labelEn: "Integral", aliases: ["integral"], keywords: ["积分", "定积分"], category: "calculus", defaultPriority: 100, supportedInMathMode: true },
  { id: "iint", command: "\\iint", insertTemplate: "\\iint_{\\placeholder{}} \\placeholder{}\\,\\mathrm{d}A", previewLatex: "\\iint_D f\\,\\mathrm{d}A", labelZh: "二重积分", labelEn: "Double integral", aliases: ["double integral"], keywords: ["二重积分"], category: "calculus", defaultPriority: 78, supportedInMathMode: true },
  { id: "oint", command: "\\oint", insertTemplate: "\\oint_{\\placeholder{}} \\placeholder{}\\,\\mathrm{d}\\placeholder{}", previewLatex: "\\oint_C \\mathbf{F}\\cdot\\mathrm{d}\\mathbf{r}", labelZh: "环路积分", labelEn: "Contour integral", aliases: ["contour integral"], keywords: ["闭合积分", "环积分"], category: "calculus", defaultPriority: 76, supportedInMathMode: true },
  { id: "sum", command: "\\sum", insertTemplate: "\\sum_{\\placeholder{}}^{\\placeholder{}} \\placeholder{}", previewLatex: "\\sum_{i=1}^{n} a_i", labelZh: "求和", labelEn: "Summation", aliases: ["summation"], keywords: ["求和", "西格玛"], category: "calculus", defaultPriority: 96, supportedInMathMode: true },
  { id: "prod", command: "\\prod", insertTemplate: "\\prod_{\\placeholder{}}^{\\placeholder{}} \\placeholder{}", previewLatex: "\\prod_{i=1}^{n} a_i", labelZh: "连乘", labelEn: "Product", aliases: ["product"], keywords: ["乘积", "连乘"], category: "calculus", defaultPriority: 79, supportedInMathMode: true },
  { id: "lim", command: "\\lim", insertTemplate: "\\lim_{\\placeholder{}\\to\\placeholder{}} \\placeholder{}", previewLatex: "\\lim_{x\\to 0} f(x)", labelZh: "极限", labelEn: "Limit", aliases: ["limit"], keywords: ["极限"], category: "calculus", defaultPriority: 92, supportedInMathMode: true },
  { id: "partial", command: "\\partial", insertTemplate: "\\frac{\\partial \\placeholder{}}{\\partial \\placeholder{}}", previewLatex: "\\frac{\\partial f}{\\partial x}", labelZh: "偏导数", labelEn: "Partial derivative", aliases: ["partial", "derivative"], keywords: ["偏导", "偏微分"], category: "calculus", defaultPriority: 86, supportedInMathMode: true },
  { id: "nabla", command: "\\nabla", insertTemplate: "\\nabla", previewLatex: "\\nabla f", labelZh: "Nabla 算子", labelEn: "Nabla", aliases: ["gradient", "del"], keywords: ["梯度", "哈密顿算子"], category: "calculus", defaultPriority: 74, supportedInMathMode: true },
  { id: "infty", command: "\\infty", insertTemplate: "\\infty", previewLatex: "\\infty", labelZh: "无穷", labelEn: "Infinity", aliases: ["infinity"], keywords: ["无穷大"], category: "calculus", defaultPriority: 88, supportedInMathMode: true },

  { id: "matrix2", command: "\\begin{bmatrix}", insertTemplate: "\\begin{bmatrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{bmatrix}", previewLatex: "\\begin{bmatrix}a&b\\\\c&d\\end{bmatrix}", labelZh: "2×2 方括号矩阵", labelEn: "2×2 matrix", aliases: ["matrix", "bmatrix"], keywords: ["矩阵", "方阵"], category: "matrix", defaultPriority: 100, supportedInMathMode: true },
  { id: "matrix3", command: "\\begin{bmatrix}", insertTemplate: "\\begin{bmatrix}\\placeholder{} & \\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{} & \\placeholder{}\\end{bmatrix}", previewLatex: "\\begin{bmatrix}a&b&c\\\\d&e&f\\\\g&h&i\\end{bmatrix}", labelZh: "3×3 方括号矩阵", labelEn: "3×3 matrix", aliases: ["matrix", "bmatrix"], keywords: ["矩阵", "三阶矩阵"], category: "matrix", defaultPriority: 90, supportedInMathMode: true },
  { id: "pmatrix2", command: "\\begin{pmatrix}", insertTemplate: "\\begin{pmatrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{pmatrix}", previewLatex: "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}", labelZh: "圆括号矩阵", labelEn: "Parenthesized matrix", aliases: ["pmatrix"], keywords: ["矩阵", "圆括号矩阵"], category: "matrix", defaultPriority: 88, supportedInMathMode: true },
  { id: "determinant", command: "\\begin{vmatrix}", insertTemplate: "\\begin{vmatrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{vmatrix}", previewLatex: "\\begin{vmatrix}a&b\\\\c&d\\end{vmatrix}", labelZh: "行列式", labelEn: "Determinant", aliases: ["det", "determinant"], keywords: ["行列式"], category: "matrix", defaultPriority: 87, supportedInMathMode: true },
  { id: "vector", command: "\\vec", insertTemplate: "\\vec{\\placeholder{}}", previewLatex: "\\vec{v}", labelZh: "向量", labelEn: "Vector", aliases: ["vector"], keywords: ["矢量", "向量"], category: "matrix", defaultPriority: 84, supportedInMathMode: true },
  { id: "boldsymbol", command: "\\mathbf", insertTemplate: "\\mathbf{\\placeholder{}}", previewLatex: "\\mathbf{A}", labelZh: "粗体", labelEn: "Bold math", aliases: ["bold", "mathbf"], keywords: ["黑体", "矩阵粗体"], category: "matrix", defaultPriority: 72, supportedInMathMode: true },

  { id: "alpha", command: "\\alpha", insertTemplate: "\\alpha", previewLatex: "\\alpha", labelZh: "阿尔法", labelEn: "Alpha", aliases: ["alpha"], keywords: ["希腊字母"], category: "greek", defaultPriority: 100, supportedInMathMode: true },
  { id: "beta", command: "\\beta", insertTemplate: "\\beta", previewLatex: "\\beta", labelZh: "贝塔", labelEn: "Beta", aliases: ["beta"], keywords: ["希腊字母"], category: "greek", defaultPriority: 98, supportedInMathMode: true },
  { id: "gamma", command: "\\gamma", insertTemplate: "\\gamma", previewLatex: "\\gamma", labelZh: "伽马", labelEn: "Gamma", aliases: ["gamma"], keywords: ["希腊字母"], category: "greek", defaultPriority: 96, supportedInMathMode: true },
  { id: "delta", command: "\\delta", insertTemplate: "\\delta", previewLatex: "\\delta", labelZh: "德尔塔", labelEn: "Delta", aliases: ["delta"], keywords: ["希腊字母"], category: "greek", defaultPriority: 94, supportedInMathMode: true },
  { id: "theta", command: "\\theta", insertTemplate: "\\theta", previewLatex: "\\theta", labelZh: "西塔", labelEn: "Theta", aliases: ["theta"], keywords: ["角度", "希腊字母"], category: "greek", defaultPriority: 92, supportedInMathMode: true },
  { id: "lambda", command: "\\lambda", insertTemplate: "\\lambda", previewLatex: "\\lambda", labelZh: "拉姆达", labelEn: "Lambda", aliases: ["lambda"], keywords: ["波长", "本征值"], category: "greek", defaultPriority: 90, supportedInMathMode: true },
  { id: "mu", command: "\\mu", insertTemplate: "\\mu", previewLatex: "\\mu", labelZh: "缪", labelEn: "Mu", aliases: ["mu"], keywords: ["希腊字母"], category: "greek", defaultPriority: 88, supportedInMathMode: true },
  { id: "pi", command: "\\pi", insertTemplate: "\\pi", previewLatex: "\\pi", labelZh: "圆周率", labelEn: "Pi", aliases: ["pi"], keywords: ["派", "圆周率"], category: "greek", defaultPriority: 100, supportedInMathMode: true },
  { id: "sigma", command: "\\sigma", insertTemplate: "\\sigma", previewLatex: "\\sigma", labelZh: "西格玛", labelEn: "Sigma", aliases: ["sigma"], keywords: ["标准差", "希腊字母"], category: "greek", defaultPriority: 86, supportedInMathMode: true },
  { id: "omega", command: "\\omega", insertTemplate: "\\omega", previewLatex: "\\omega", labelZh: "欧米伽", labelEn: "Omega", aliases: ["omega"], keywords: ["角频率", "希腊字母"], category: "greek", defaultPriority: 84, supportedInMathMode: true },

  { id: "equal", command: "=", insertTemplate: "=", previewLatex: "=", labelZh: "等于", labelEn: "Equals", aliases: ["equal"], keywords: ["等于"], category: "relation", defaultPriority: 100, supportedInMathMode: true },
  { id: "neq", command: "\\neq", insertTemplate: "\\neq", previewLatex: "\\neq", labelZh: "不等于", labelEn: "Not equal", aliases: ["not equal"], keywords: ["不等于"], category: "relation", defaultPriority: 96, supportedInMathMode: true },
  { id: "approx", command: "\\approx", insertTemplate: "\\approx", previewLatex: "\\approx", labelZh: "约等于", labelEn: "Approximately", aliases: ["approximately"], keywords: ["约等于", "近似"], category: "relation", defaultPriority: 92, supportedInMathMode: true },
  { id: "leq", command: "\\leq", insertTemplate: "\\leq", previewLatex: "\\leq", labelZh: "小于等于", labelEn: "Less or equal", aliases: ["less equal"], keywords: ["小于等于"], category: "relation", defaultPriority: 90, supportedInMathMode: true },
  { id: "geq", command: "\\geq", insertTemplate: "\\geq", previewLatex: "\\geq", labelZh: "大于等于", labelEn: "Greater or equal", aliases: ["greater equal"], keywords: ["大于等于"], category: "relation", defaultPriority: 90, supportedInMathMode: true },
  { id: "propto", command: "\\propto", insertTemplate: "\\propto", previewLatex: "\\propto", labelZh: "正比于", labelEn: "Proportional to", aliases: ["proportional"], keywords: ["正比"], category: "relation", defaultPriority: 82, supportedInMathMode: true },

  { id: "in", command: "\\in", insertTemplate: "\\in", previewLatex: "\\in", labelZh: "属于", labelEn: "Element of", aliases: ["element"], keywords: ["属于"], category: "set", defaultPriority: 100, supportedInMathMode: true },
  { id: "notin", command: "\\notin", insertTemplate: "\\notin", previewLatex: "\\notin", labelZh: "不属于", labelEn: "Not an element", aliases: ["not element"], keywords: ["不属于"], category: "set", defaultPriority: 92, supportedInMathMode: true },
  { id: "subset", command: "\\subset", insertTemplate: "\\subset", previewLatex: "\\subset", labelZh: "真子集", labelEn: "Subset", aliases: ["subset"], keywords: ["子集"], category: "set", defaultPriority: 90, supportedInMathMode: true },
  { id: "cup", command: "\\cup", insertTemplate: "\\cup", previewLatex: "\\cup", labelZh: "并集", labelEn: "Union", aliases: ["union"], keywords: ["并集"], category: "set", defaultPriority: 88, supportedInMathMode: true },
  { id: "cap", command: "\\cap", insertTemplate: "\\cap", previewLatex: "\\cap", labelZh: "交集", labelEn: "Intersection", aliases: ["intersection"], keywords: ["交集"], category: "set", defaultPriority: 88, supportedInMathMode: true },
  { id: "forall", command: "\\forall", insertTemplate: "\\forall", previewLatex: "\\forall", labelZh: "任意", labelEn: "For all", aliases: ["for all"], keywords: ["任意", "所有"], category: "set", defaultPriority: 86, supportedInMathMode: true },
  { id: "exists", command: "\\exists", insertTemplate: "\\exists", previewLatex: "\\exists", labelZh: "存在", labelEn: "Exists", aliases: ["exists"], keywords: ["存在"], category: "set", defaultPriority: 84, supportedInMathMode: true },

  { id: "to", command: "\\to", insertTemplate: "\\to", previewLatex: "\\to", labelZh: "趋于", labelEn: "To", aliases: ["to", "right arrow"], keywords: ["趋于", "箭头"], category: "arrow", defaultPriority: 100, supportedInMathMode: true },
  { id: "rightarrow", command: "\\rightarrow", insertTemplate: "\\rightarrow", previewLatex: "\\rightarrow", labelZh: "右箭头", labelEn: "Right arrow", aliases: ["right arrow"], keywords: ["右箭头"], category: "arrow", defaultPriority: 96, supportedInMathMode: true },
  { id: "leftarrow", command: "\\leftarrow", insertTemplate: "\\leftarrow", previewLatex: "\\leftarrow", labelZh: "左箭头", labelEn: "Left arrow", aliases: ["left arrow"], keywords: ["左箭头"], category: "arrow", defaultPriority: 94, supportedInMathMode: true },
  { id: "leftrightarrow", command: "\\leftrightarrow", insertTemplate: "\\leftrightarrow", previewLatex: "\\leftrightarrow", labelZh: "双向箭头", labelEn: "Both ways", aliases: ["both arrow"], keywords: ["双向箭头"], category: "arrow", defaultPriority: 88, supportedInMathMode: true },
  { id: "Rightarrow", command: "\\Rightarrow", insertTemplate: "\\Rightarrow", previewLatex: "\\Rightarrow", labelZh: "推出", labelEn: "Implies", aliases: ["implies"], keywords: ["推出", "蕴含"], category: "arrow", defaultPriority: 90, supportedInMathMode: true },

  { id: "hbar", command: "\\hbar", insertTemplate: "\\hbar", previewLatex: "\\hbar", labelZh: "约化普朗克常数", labelEn: "Reduced Planck constant", aliases: ["hbar", "planck"], keywords: ["普朗克常数", "量子"], category: "physics", defaultPriority: 100, supportedInMathMode: true },
  { id: "dd", command: "\\mathrm{d}", insertTemplate: "\\mathrm{d}\\placeholder{}", previewLatex: "\\mathrm{d}x", labelZh: "微分元", labelEn: "Differential", aliases: ["differential"], keywords: ["微分", "微分元"], category: "physics", defaultPriority: 96, supportedInMathMode: true },
  { id: "bra", command: "\\langle", insertTemplate: "\\langle\\placeholder{}|", previewLatex: "\\langle\\psi|", labelZh: "左矢", labelEn: "Bra", aliases: ["bra"], keywords: ["狄拉克", "量子态"], category: "physics", defaultPriority: 92, supportedInMathMode: true },
  { id: "ket", command: "\\rangle", insertTemplate: "|\\placeholder{}\\rangle", previewLatex: "|\\psi\\rangle", labelZh: "右矢", labelEn: "Ket", aliases: ["ket"], keywords: ["狄拉克", "量子态"], category: "physics", defaultPriority: 92, supportedInMathMode: true },
  { id: "degree", command: "^\\circ", insertTemplate: "^{\\circ}", previewLatex: "30^{\\circ}", labelZh: "角度", labelEn: "Degree", aliases: ["degree"], keywords: ["度", "角度"], category: "physics", defaultPriority: 86, supportedInMathMode: true },
];

export const commandRegistry: LatexCommand[] = [
  ...baseCommandRegistry,
  ...additionalCommands,
];

export const categoryLabels: Record<string, string> = {
  common: "常用",
  structure: "结构",
  calculus: "微积分",
  matrix: "矩阵",
  greek: "希腊字母",
  relation: "关系",
  set: "集合与逻辑",
  arrow: "箭头",
  physics: "物理常用",
};

export const categoryLabelsEn: Record<string, string> = {
  common: "Common",
  structure: "Structures",
  calculus: "Calculus",
  matrix: "Matrices",
  greek: "Greek",
  relation: "Relations",
  set: "Sets & Logic",
  arrow: "Arrows",
  physics: "Physics",
};

export const commonCommandIds = [
  "frac",
  "sqrt",
  "power",
  "subscript",
  "int",
  "sum",
  "lim",
  "matrix2",
  "cases",
  "derivative",
  "pi",
  "infty",
  "vector",
  "Rset",
];
