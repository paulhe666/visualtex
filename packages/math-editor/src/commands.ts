export type CommandCategory =
  | "common"
  | "structure"
  | "calculus"
  | "matrix"
  | "greek"
  | "relation"
  | "set"
  | "arrow"
  | "physics";

export interface LatexCommand {
  id: string;
  command: string;
  insertTemplate: string;
  previewLatex: string;
  labelZh: string;
  labelEn: string;
  aliases: string[];
  keywords: string[];
  category: CommandCategory;
  defaultPriority: number;
}

const command = (
  id: string,
  commandName: string,
  insertTemplate: string,
  previewLatex: string,
  labelZh: string,
  labelEn: string,
  category: CommandCategory,
  defaultPriority: number,
  aliases: string[] = [],
  keywords: string[] = [],
): LatexCommand => ({
  id,
  command: commandName,
  insertTemplate,
  previewLatex,
  labelZh,
  labelEn,
  aliases,
  keywords,
  category,
  defaultPriority,
});

// Based on the VisualTeX 1.0.6 command registry. The desktop paper editor
// keeps the same templates and selection semantics, but exposes them through
// a controlled single-formula component instead of the original document store.
export const commandRegistry: LatexCommand[] = [
  command("frac", "\\frac", "\\frac{\\placeholder{}}{\\placeholder{}}", "\\frac{a}{b}", "分式", "Fraction", "structure", 100, ["divide", "fraction"], ["分数", "除法"]),
  command("smallfrac", "\\tfrac", "\\tfrac{\\placeholder{}}{\\placeholder{}}", "\\tfrac{a}{b}", "行内分式", "Text fraction", "structure", 79),
  command("displayfrac", "\\dfrac", "\\dfrac{\\placeholder{}}{\\placeholder{}}", "\\dfrac{a}{b}", "大型分式", "Display fraction", "structure", 78),
  command("sqrt", "\\sqrt", "\\sqrt{\\placeholder{}}", "\\sqrt{x}", "平方根", "Square root", "structure", 98, ["root"], ["根号", "开方"]),
  command("nthroot", "\\sqrt", "\\sqrt[\\placeholder{}]{\\placeholder{}}", "\\sqrt[n]{x}", "n 次根", "Nth root", "structure", 82),
  command("scripts", "_{}^{}", "\\placeholder{}_{\\placeholder{}}^{\\placeholder{}}", "X_a^b", "添加上下标", "Upper/lower limits", "structure", 96, ["limits", "scripts"], ["上下标", "上下限"]),
  command("lower-script", "_{}", "\\placeholder{}_{\\placeholder{}}", "X_a", "添加下标", "Subscript", "structure", 92),
  command("upper-script", "^{}", "\\placeholder{}^{\\placeholder{}}", "X^b", "添加上标", "Superscript", "structure", 92),
  command("parentheses", "\\left(", "\\left(\\placeholder{}\\right)", "\\left(x\\right)", "自适应圆括号", "Parentheses", "structure", 84),
  command("brackets", "\\left[", "\\left[\\placeholder{}\\right]", "\\left[x\\right]", "方括号", "Brackets", "structure", 81),
  command("braces", "\\left\\{", "\\left\\{\\placeholder{}\\right\\}", "\\left\\{x\\right\\}", "花括号", "Braces", "structure", 80),
  command("absolute", "\\left|", "\\left|\\placeholder{}\\right|", "\\left|x\\right|", "绝对值", "Absolute value", "structure", 83, ["abs"]),
  command("cases", "\\begin{cases}", "\\begin{cases}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{cases}", "f(x)=\\begin{cases}x&x>0\\\\0&x\\le0\\end{cases}", "分段函数", "Cases", "structure", 85, ["piecewise"]),
  command("overline", "\\overline", "\\overline{\\placeholder{}}", "\\overline{x}", "上划线", "Overline", "structure", 75),
  command("underbrace", "\\underbrace", "\\underbrace{\\placeholder{}}_{\\placeholder{}}", "\\underbrace{x+\\cdots+x}_n", "下花括号", "Underbrace", "structure", 72),

  command("intplain", "\\int", "\\int \\placeholder{}\\,\\mathrm{d}\\placeholder{}", "\\int f(x)\\,\\mathrm{d}x", "不定积分", "Indefinite integral", "calculus", 99),
  command("int", "\\int", "\\int_{\\placeholder{}}^{\\placeholder{}} \\placeholder{}\\,\\mathrm{d}\\placeholder{}", "\\int_a^b f(x)\\,\\mathrm{d}x", "定积分", "Integral", "calculus", 100),
  command("iint", "\\iint", "\\iint_{\\placeholder{}} \\placeholder{}\\,\\mathrm{d}A", "\\iint_D f\\,\\mathrm{d}A", "二重积分", "Double integral", "calculus", 94),
  command("iiint", "\\iiint", "\\iiint_{\\placeholder{}} \\placeholder{}\\,\\mathrm{d}V", "\\iiint_V f\\,\\mathrm{d}V", "三重积分", "Triple integral", "calculus", 93),
  command("oint", "\\oint", "\\oint_{\\placeholder{}} \\placeholder{}\\,\\mathrm{d}\\placeholder{}", "\\oint_C \\mathbf{F}\\cdot\\mathrm{d}\\mathbf{r}", "环路积分", "Contour integral", "calculus", 92),
  command("sum", "\\sum", "\\sum_{\\placeholder{}}^{\\placeholder{}} \\placeholder{}", "\\sum_{i=1}^{n}a_i", "求和", "Summation", "calculus", 96),
  command("prod", "\\prod", "\\prod_{\\placeholder{}}^{\\placeholder{}} \\placeholder{}", "\\prod_{i=1}^{n}a_i", "连乘", "Product", "calculus", 90),
  command("lim", "\\lim", "\\lim_{\\placeholder{}\\to\\placeholder{}} \\placeholder{}", "\\lim_{x\\to0}f(x)", "极限", "Limit", "calculus", 92),
  command("derivative", "\\frac{d}{dx}", "\\frac{\\mathrm{d}\\placeholder{}}{\\mathrm{d}\\placeholder{}}", "\\frac{\\mathrm{d}f}{\\mathrm{d}x}", "导数", "Derivative", "calculus", 91),
  command("partial", "\\partial", "\\frac{\\partial \\placeholder{}}{\\partial \\placeholder{}}", "\\frac{\\partial f}{\\partial x}", "偏导数", "Partial derivative", "calculus", 90),
  command("nabla", "\\nabla", "\\nabla", "\\nabla f", "Nabla 算子", "Nabla", "calculus", 80),
  command("infty", "\\infty", "\\infty", "\\infty", "无穷", "Infinity", "calculus", 88),
  command("sin", "\\sin", "\\sin\\left(\\placeholder{}\\right)", "\\sin x", "正弦", "Sine", "calculus", 84),
  command("cos", "\\cos", "\\cos\\left(\\placeholder{}\\right)", "\\cos x", "余弦", "Cosine", "calculus", 84),
  command("ln", "\\ln", "\\ln\\left(\\placeholder{}\\right)", "\\ln x", "自然对数", "Natural log", "calculus", 82),

  command("matrix2", "\\begin{bmatrix}", "\\begin{bmatrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{bmatrix}", "\\begin{bmatrix}a&b\\\\c&d\\end{bmatrix}", "2×2 方括号矩阵", "2×2 matrix", "matrix", 100),
  command("matrix3", "\\begin{bmatrix}", "\\begin{bmatrix}\\placeholder{} & \\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{} & \\placeholder{}\\end{bmatrix}", "\\begin{bmatrix}a&b&c\\\\d&e&f\\\\g&h&i\\end{bmatrix}", "3×3 方括号矩阵", "3×3 matrix", "matrix", 90),
  command("pmatrix2", "\\begin{pmatrix}", "\\begin{pmatrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{pmatrix}", "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}", "圆括号矩阵", "Parenthesized matrix", "matrix", 88),
  command("determinant", "\\begin{vmatrix}", "\\begin{vmatrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{vmatrix}", "\\begin{vmatrix}a&b\\\\c&d\\end{vmatrix}", "行列式", "Determinant", "matrix", 87),
  command("vector", "\\vec", "\\vec{\\placeholder{}}", "\\vec{v}", "向量", "Vector", "matrix", 84),
  command("norm", "\\lVert", "\\left\\lVert\\placeholder{}\\right\\rVert", "\\lVert\\mathbf{x}\\rVert", "范数", "Norm", "matrix", 81),
  command("transpose", "^{\\mathsf{T}}", "^{\\mathsf{T}}", "A^{\\mathsf{T}}", "转置", "Transpose", "matrix", 84),

  ...[
    ["alpha", "\\alpha", "阿尔法", "Alpha"],
    ["beta", "\\beta", "贝塔", "Beta"],
    ["gamma", "\\gamma", "伽马", "Gamma"],
    ["delta", "\\delta", "德尔塔", "Delta"],
    ["theta", "\\theta", "西塔", "Theta"],
    ["lambda", "\\lambda", "拉姆达", "Lambda"],
    ["mu", "\\mu", "缪", "Mu"],
    ["pi", "\\pi", "圆周率", "Pi"],
    ["sigma", "\\sigma", "西格玛", "Sigma"],
    ["phi", "\\phi", "斐", "Phi"],
    ["psi", "\\psi", "普赛", "Psi"],
    ["omega", "\\omega", "欧米伽", "Omega"],
    ["Delta", "\\Delta", "大写德尔塔", "Delta"],
    ["Sigma", "\\Sigma", "大写西格玛", "Sigma"],
    ["Omega", "\\Omega", "大写欧米伽", "Omega"],
  ].map(([id, latex, zh, en], index) => command(id, latex, latex, latex, zh, en, "greek", 100 - index)),

  command("equal", "=", "=", "=", "等于", "Equals", "relation", 100),
  command("neq", "\\neq", "\\neq", "\\neq", "不等于", "Not equal", "relation", 96),
  command("approx", "\\approx", "\\approx", "\\approx", "约等于", "Approximately", "relation", 92),
  command("leq", "\\leq", "\\leq", "\\leq", "小于等于", "Less or equal", "relation", 90),
  command("geq", "\\geq", "\\geq", "\\geq", "大于等于", "Greater or equal", "relation", 90),
  command("propto", "\\propto", "\\propto", "\\propto", "正比于", "Proportional", "relation", 82),
  command("perp", "\\perp", "\\perp", "\\perp", "垂直", "Perpendicular", "relation", 81),

  command("in", "\\in", "\\in", "\\in", "属于", "Element of", "set", 100),
  command("notin", "\\notin", "\\notin", "\\notin", "不属于", "Not an element", "set", 92),
  command("subseteq", "\\subseteq", "\\subseteq", "\\subseteq", "子集或相等", "Subset or equal", "set", 88),
  command("cup", "\\cup", "\\cup", "\\cup", "并集", "Union", "set", 88),
  command("cap", "\\cap", "\\cap", "\\cap", "交集", "Intersection", "set", 88),
  command("forall", "\\forall", "\\forall", "\\forall", "任意", "For all", "set", 86),
  command("exists", "\\exists", "\\exists", "\\exists", "存在", "Exists", "set", 84),
  command("Rset", "\\mathbb{R}", "\\mathbb{R}", "\\mathbb{R}", "实数集", "Real numbers", "set", 92),

  command("to", "\\to", "\\to", "\\to", "趋于", "To", "arrow", 100),
  command("rightarrow", "\\rightarrow", "\\rightarrow", "\\rightarrow", "右箭头", "Right arrow", "arrow", 96),
  command("leftarrow", "\\leftarrow", "\\leftarrow", "\\leftarrow", "左箭头", "Left arrow", "arrow", 94),
  command("leftrightarrow", "\\leftrightarrow", "\\leftrightarrow", "\\leftrightarrow", "双向箭头", "Both ways", "arrow", 88),
  command("Rightarrow", "\\Rightarrow", "\\Rightarrow", "\\Rightarrow", "推出", "Implies", "arrow", 90),

  command("hbar", "\\hbar", "\\hbar", "\\hbar", "约化普朗克常数", "Reduced Planck constant", "physics", 100),
  command("dd", "\\mathrm{d}", "\\mathrm{d}\\placeholder{}", "\\mathrm{d}x", "微分元", "Differential", "physics", 96),
  command("bra", "\\langle", "\\langle\\placeholder{}|", "\\langle\\psi|", "左矢", "Bra", "physics", 92),
  command("ket", "\\rangle", "|\\placeholder{}\\rangle", "|\\psi\\rangle", "右矢", "Ket", "physics", 92),
  command("commutator", "[", "\\left[\\placeholder{},\\placeholder{}\\right]", "[A,B]", "对易子", "Commutator", "physics", 88),
  command("laplacian", "\\nabla^{2}", "\\nabla^{2}", "\\nabla^{2}", "拉普拉斯算子", "Laplacian", "physics", 88),
];

export const categoryLabels: Record<CommandCategory, string> = {
  common: "常用",
  structure: "结构",
  calculus: "微积分",
  matrix: "矩阵",
  greek: "希腊字母",
  relation: "关系",
  set: "集合与逻辑",
  arrow: "箭头",
  physics: "物理",
};

export const categories: CommandCategory[] = [
  "common",
  "structure",
  "calculus",
  "matrix",
  "greek",
  "relation",
  "set",
  "arrow",
  "physics",
];

export const commonCommandIds = [
  "frac",
  "sqrt",
  "scripts",
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

const normalize = (value: string) => value.trim().replace(/^\\/, "").toLocaleLowerCase();

function editDistance(left: string, right: string): number {
  const matrix = Array.from({ length: left.length + 1 }, () =>
    new Array<number>(right.length + 1).fill(0),
  );
  for (let row = 0; row <= left.length; row += 1) matrix[row]![0] = row;
  for (let column = 0; column <= right.length; column += 1) matrix[0]![column] = column;
  for (let row = 1; row <= left.length; row += 1) {
    for (let column = 1; column <= right.length; column += 1) {
      const cost = left[row - 1] === right[column - 1] ? 0 : 1;
      matrix[row]![column] = Math.min(
        matrix[row - 1]![column]! + 1,
        matrix[row]![column - 1]! + 1,
        matrix[row - 1]![column - 1]! + cost,
      );
    }
  }
  return matrix[left.length]![right.length]!;
}

function matchScore(query: string, candidate: LatexCommand): number {
  if (!query) return candidate.defaultPriority / 3;
  const values = [
    candidate.command,
    ...candidate.aliases,
    candidate.labelZh,
    candidate.labelEn,
    ...candidate.keywords,
  ].map(normalize);
  let score = Number.NEGATIVE_INFINITY;
  for (const value of values) {
    if (value === query) score = Math.max(score, 420);
    else if (value.startsWith(query)) score = Math.max(score, 320 - (value.length - query.length) * 2);
    else if (value.includes(query)) score = Math.max(score, 220 - value.indexOf(query) * 4);
    else if (query.length >= 3) {
      const distance = editDistance(query, value.slice(0, Math.max(query.length, Math.min(value.length, query.length + 2))));
      if (distance <= Math.max(1, Math.floor(query.length / 3))) {
        score = Math.max(score, 145 - distance * 25);
      }
    }
  }
  return score;
}

export function searchCommands(rawQuery: string, limit = 8): LatexCommand[] {
  const query = normalize(rawQuery);
  return commandRegistry
    .map((candidate) => ({
      candidate,
      score: matchScore(query, candidate) + candidate.defaultPriority / 5,
    }))
    .filter((entry) => Number.isFinite(entry.score))
    .sort((left, right) => right.score - left.score)
    .slice(0, limit)
    .map((entry) => entry.candidate);
}

export function commandsForCategory(category: CommandCategory): LatexCommand[] {
  if (category === "common") {
    return commonCommandIds
      .map((id) => commandRegistry.find((candidate) => candidate.id === id))
      .filter((candidate): candidate is LatexCommand => Boolean(candidate));
  }
  return commandRegistry.filter((candidate) => candidate.category === category);
}

export function createMatrixCommand(rows: number, columns: number, delimiter: "bmatrix" | "pmatrix" | "vmatrix"): LatexCommand {
  const body = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => "\\placeholder{}").join(" & "),
  ).join(" \\\\ ");
  return command(
    `custom-${delimiter}-${rows}x${columns}`,
    `\\begin{${delimiter}}`,
    `\\begin{${delimiter}}${body}\\end{${delimiter}}`,
    `\\begin{${delimiter}}a&b\\\\c&d\\end{${delimiter}}`,
    `${rows}×${columns} 矩阵`,
    `${rows}×${columns} matrix`,
    "matrix",
    120,
  );
}

export function templateForSelection(candidate: LatexCommand, selectedLatex: string): string {
  if (!selectedLatex) return candidate.insertTemplate;
  switch (candidate.id) {
    case "scripts":
      return `${selectedLatex}_{\\placeholder{}}^{\\placeholder{}}`;
    case "lower-script":
      return `${selectedLatex}_{\\placeholder{}}`;
    case "upper-script":
      return `${selectedLatex}^{\\placeholder{}}`;
    case "sum":
      return `\\sum_{\\placeholder{}}^{\\placeholder{}} ${selectedLatex}`;
    case "prod":
      return `\\prod_{\\placeholder{}}^{\\placeholder{}} ${selectedLatex}`;
    case "int":
      return `\\int_{\\placeholder{}}^{\\placeholder{}} ${selectedLatex}\\,\\mathrm{d}\\placeholder{}`;
    case "intplain":
      return `\\int ${selectedLatex}\\,\\mathrm{d}\\placeholder{}`;
    case "frac":
    case "smallfrac":
    case "displayfrac":
      return `${candidate.command}{${selectedLatex}}{\\placeholder{}}`;
    case "sqrt":
      return `\\sqrt{${selectedLatex}}`;
    case "parentheses":
      return `\\left(${selectedLatex}\\right)`;
    case "brackets":
      return `\\left[${selectedLatex}\\right]`;
    case "braces":
      return `\\left\\{${selectedLatex}\\right\\}`;
    case "absolute":
      return `\\left|${selectedLatex}\\right|`;
    default:
      return candidate.insertTemplate.replace("\\placeholder{}", selectedLatex);
  }
}
