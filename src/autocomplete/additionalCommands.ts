import type { CommandCategory, LatexCommand } from "../types/command";

const makeCommand = (
  id: string,
  command: string,
  insertTemplate: string,
  previewLatex: string,
  labelZh: string,
  labelEn: string,
  category: CommandCategory,
  defaultPriority = 72,
  aliases: string[] = [],
  keywords: string[] = [],
): LatexCommand => ({
  id,
  command,
  insertTemplate,
  previewLatex,
  labelZh,
  labelEn,
  aliases,
  keywords,
  category,
  defaultPriority,
  supportedInMathMode: true,
});

const complexCommands: LatexCommand[] = [
  // 对现有选区添加上标、下标或上下限
  makeCommand("scripts", "_{}^{}", "\\placeholder{}_{\\placeholder{}}^{\\placeholder{}}", "X_{a}^{b}", "添加上下标", "Add upper/lower limits", "structure", 96, ["limits", "scripts"], ["上下限", "上下标"]),
  makeCommand("lower-script", "_{}", "\\placeholder{}_{\\placeholder{}}", "X_{a}", "添加下标", "Add lower limit", "structure", 92, ["lower limit", "subscript"], ["下限", "下标"]),
  makeCommand("upper-script", "^{}", "\\placeholder{}^{\\placeholder{}}", "X^{b}", "添加上标", "Add upper limit", "structure", 92, ["upper limit", "superscript"], ["上限", "上标"]),

  // 基础结构
  makeCommand("smallfrac", "\\tfrac", "\\tfrac{\\placeholder{}}{\\placeholder{}}", "\\tfrac{a}{b}", "行内分式", "Text fraction", "structure", 79, ["tfrac"], ["小分式"]),
  makeCommand("displayfrac", "\\dfrac", "\\dfrac{\\placeholder{}}{\\placeholder{}}", "\\dfrac{a}{b}", "大型分式", "Display fraction", "structure", 78, ["dfrac"], ["大分式"]),
  makeCommand("brackets", "\\left[", "\\left[\\placeholder{}\\right]", "\\left[x\\right]", "方括号", "Brackets", "structure", 81, ["square brackets"], ["方括号"]),
  makeCommand("braces", "\\left\\{", "\\left\\{\\placeholder{}\\right\\}", "\\left\\{x\\right\\}", "花括号", "Braces", "structure", 80, ["curly braces"], ["大括号", "花括号"]),
  makeCommand("anglebrackets", "\\langle", "\\left\\langle\\placeholder{}\\right\\rangle", "\\langle x\\rangle", "尖括号", "Angle brackets", "structure", 77, ["angle brackets"], ["尖括号"]),
  makeCommand("floor", "\\lfloor", "\\left\\lfloor\\placeholder{}\\right\\rfloor", "\\lfloor x\\rfloor", "下取整", "Floor", "structure", 76, ["floor"], ["向下取整"]),
  makeCommand("ceil", "\\lceil", "\\left\\lceil\\placeholder{}\\right\\rceil", "\\lceil x\\rceil", "上取整", "Ceiling", "structure", 76, ["ceil", "ceiling"], ["向上取整"]),
  makeCommand("overline", "\\overline", "\\overline{\\placeholder{}}", "\\overline{x}", "上划线", "Overline", "structure", 75, ["bar", "overline"], ["平均值", "上划线"]),
  makeCommand("underline", "\\underline", "\\underline{\\placeholder{}}", "\\underline{x}", "下划线", "Underline", "structure", 73, ["underline"], ["下划线"]),
  makeCommand("overbrace", "\\overbrace", "\\overbrace{\\placeholder{}}^{\\placeholder{}}", "\\overbrace{a+\\cdots+a}^{n}", "上花括号", "Overbrace", "structure", 72, ["overbrace"], ["上括注"]),
  makeCommand("underbrace", "\\underbrace", "\\underbrace{\\placeholder{}}_{\\placeholder{}}", "\\underbrace{a+\\cdots+a}_{n}", "下花括号", "Underbrace", "structure", 72, ["underbrace"], ["下括注"]),
  makeCommand("cases", "\\begin{cases}", "\\begin{cases}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{cases}", "f(x)=\\begin{cases}x&x>0\\\\0&x\\leq0\\end{cases}", "分段函数", "Cases", "structure", 85, ["cases", "piecewise"], ["分段函数"]),
  makeCommand("boxed", "\\boxed", "\\boxed{\\placeholder{}}", "\\boxed{x}", "方框公式", "Boxed", "structure", 70, ["boxed", "box"], ["方框"]),

  // 微积分与常用函数
  makeCommand("iiint", "\\iiint", "\\iiint_{\\placeholder{}}\\placeholder{}\\,\\mathrm{d}V", "\\iiint_V f\\,\\mathrm{d}V", "三重积分", "Triple integral", "calculus", 77, ["triple integral"], ["三重积分"]),
  makeCommand("intplain", "\\int", "\\int \\placeholder{}\\,\\mathrm{d}\\placeholder{}", "\\int f(x)\\,\\mathrm{d}x", "不定积分", "Indefinite integral", "calculus", 93, ["indefinite integral"], ["不定积分"]),
  makeCommand("derivative", "\\frac{d}{dx}", "\\frac{\\mathrm{d}\\placeholder{}}{\\mathrm{d}\\placeholder{}}", "\\frac{\\mathrm{d}f}{\\mathrm{d}x}", "导数", "Derivative", "calculus", 91, ["derivative"], ["导数", "微分"]),
  makeCommand("secondderivative", "\\frac{d^2}{dx^2}", "\\frac{\\mathrm{d}^{2}\\placeholder{}}{\\mathrm{d}\\placeholder{}^{2}}", "\\frac{\\mathrm{d}^{2}f}{\\mathrm{d}x^{2}}", "二阶导数", "Second derivative", "calculus", 83, ["second derivative"], ["二阶导数"]),
  makeCommand("partialsecond", "\\partial^2", "\\frac{\\partial^{2}\\placeholder{}}{\\partial\\placeholder{}^{2}}", "\\frac{\\partial^{2}f}{\\partial x^{2}}", "二阶偏导", "Second partial", "calculus", 82, ["second partial"], ["二阶偏导"]),
  makeCommand("mixedpartial", "\\partial^2", "\\frac{\\partial^{2}\\placeholder{}}{\\partial\\placeholder{}\\partial\\placeholder{}}", "\\frac{\\partial^{2}f}{\\partial x\\partial y}", "混合偏导", "Mixed partial", "calculus", 81, ["mixed partial"], ["混合偏导"]),
  makeCommand("evalbar", "\\left.", "\\left.\\placeholder{}\\right|_{\\placeholder{}}^{\\placeholder{}}", "\\left.F(x)\\right|_a^b", "代入上下限", "Evaluation", "calculus", 80, ["evaluate"], ["代入", "上下限"]),
  makeCommand("series", "\\sum", "\\sum_{\\placeholder{}=0}^{\\infty}\\placeholder{}", "\\sum_{n=0}^{\\infty}a_n", "无穷级数", "Infinite series", "calculus", 88, ["series"], ["级数"]),
  makeCommand("productseries", "\\prod", "\\prod_{\\placeholder{}=1}^{\\infty}\\placeholder{}", "\\prod_{n=1}^{\\infty}a_n", "无穷乘积", "Infinite product", "calculus", 75, ["infinite product"], ["无穷乘积"]),
  makeCommand("log", "\\log", "\\log_{\\placeholder{}}\\left(\\placeholder{}\\right)", "\\log_a x", "对数", "Logarithm", "calculus", 86, ["logarithm"], ["对数"]),
  makeCommand("ln", "\\ln", "\\ln\\left(\\placeholder{}\\right)", "\\ln x", "自然对数", "Natural log", "calculus", 87, ["natural log"], ["自然对数"]),
  makeCommand("exp", "\\exp", "\\exp\\left(\\placeholder{}\\right)", "\\exp(x)", "指数函数", "Exponential", "calculus", 85, ["exponential"], ["指数函数"]),
  makeCommand("sin", "\\sin", "\\sin\\left(\\placeholder{}\\right)", "\\sin x", "正弦", "Sine", "calculus", 89, ["sine"], ["正弦"]),
  makeCommand("cos", "\\cos", "\\cos\\left(\\placeholder{}\\right)", "\\cos x", "余弦", "Cosine", "calculus", 89, ["cosine"], ["余弦"]),
  makeCommand("tan", "\\tan", "\\tan\\left(\\placeholder{}\\right)", "\\tan x", "正切", "Tangent", "calculus", 84, ["tangent"], ["正切"]),
  makeCommand("arcsin", "\\arcsin", "\\arcsin\\left(\\placeholder{}\\right)", "\\arcsin x", "反正弦", "Arcsine", "calculus", 74, ["arcsine"], ["反正弦"]),
  makeCommand("min", "\\min", "\\min_{\\placeholder{}}\\placeholder{}", "\\min_x f(x)", "最小值", "Minimum", "calculus", 78, ["minimum"], ["最小值"]),
  makeCommand("max", "\\max", "\\max_{\\placeholder{}}\\placeholder{}", "\\max_x f(x)", "最大值", "Maximum", "calculus", 78, ["maximum"], ["最大值"]),

  // 矩阵与线性代数
  makeCommand("matrixplain2", "\\begin{matrix}", "\\begin{matrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{matrix}", "\\begin{matrix}a&b\\\\c&d\\end{matrix}", "无括号矩阵", "Plain matrix", "matrix", 82, ["matrix"], ["无括号矩阵"]),
  makeCommand("Bmatrix2", "\\begin{Bmatrix}", "\\begin{Bmatrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{Bmatrix}", "\\begin{Bmatrix}a&b\\\\c&d\\end{Bmatrix}", "花括号矩阵", "Brace matrix", "matrix", 80, ["Bmatrix"], ["花括号矩阵"]),
  makeCommand("Vmatrix2", "\\begin{Vmatrix}", "\\begin{Vmatrix}\\placeholder{} & \\placeholder{} \\\\ \\placeholder{} & \\placeholder{}\\end{Vmatrix}", "\\begin{Vmatrix}a&b\\\\c&d\\end{Vmatrix}", "双竖线矩阵", "Double-bar matrix", "matrix", 78, ["Vmatrix"], ["范数矩阵"]),
  makeCommand("rowvector", "\\begin{bmatrix}", "\\begin{bmatrix}\\placeholder{} & \\placeholder{} & \\placeholder{}\\end{bmatrix}", "\\begin{bmatrix}a&b&c\\end{bmatrix}", "行向量", "Row vector", "matrix", 86, ["row vector"], ["行向量"]),
  makeCommand("colvector", "\\begin{bmatrix}", "\\begin{bmatrix}\\placeholder{} \\\\ \\placeholder{} \\\\ \\placeholder{}\\end{bmatrix}", "\\begin{bmatrix}a\\\\b\\\\c\\end{bmatrix}", "列向量", "Column vector", "matrix", 86, ["column vector"], ["列向量"]),
  makeCommand("det", "\\det", "\\det\\left(\\placeholder{}\\right)", "\\det(A)", "行列式算子", "Determinant operator", "matrix", 82, ["det"], ["行列式算子"]),
  makeCommand("trace", "\\operatorname{tr}", "\\operatorname{tr}\\left(\\placeholder{}\\right)", "\\operatorname{tr}(A)", "迹", "Trace", "matrix", 78, ["trace", "tr"], ["矩阵的迹"]),
  makeCommand("rank", "\\operatorname{rank}", "\\operatorname{rank}\\left(\\placeholder{}\\right)", "\\operatorname{rank}(A)", "秩", "Rank", "matrix", 78, ["rank"], ["矩阵的秩"]),
  makeCommand("norm", "\\lVert", "\\left\\lVert\\placeholder{}\\right\\rVert", "\\lVert\\mathbf{x}\\rVert", "范数", "Norm", "matrix", 81, ["norm"], ["范数"]),
  makeCommand("unitvector", "\\hat", "\\hat{\\mathbf{\\placeholder{}}}", "\\hat{\\mathbf{e}}", "单位矢量", "Unit vector", "matrix", 84, ["unit vector"], ["单位矢量"]),

  // 物理常用结构
  makeCommand("expectation", "\\langle", "\\left\\langle\\placeholder{}\\right\\rangle", "\\langle A\\rangle", "期望值", "Expectation value", "physics", 90, ["expectation"], ["期望值"]),
  makeCommand("braket", "\\langle", "\\left\\langle\\placeholder{}\\middle|\\placeholder{}\\right\\rangle", "\\langle\\phi|\\psi\\rangle", "内积", "Bra-ket", "physics", 92, ["braket", "inner product"], ["内积", "狄拉克"]),
  makeCommand("commutator", "[", "\\left[\\placeholder{},\\placeholder{}\\right]", "[A,B]", "对易子", "Commutator", "physics", 88, ["commutator"], ["对易子"]),
  makeCommand("anticommutator", "\\{", "\\left\\{\\placeholder{},\\placeholder{}\\right\\}", "\\{A,B\\}", "反对易子", "Anticommutator", "physics", 82, ["anticommutator"], ["反对易子"]),
  makeCommand("timederivative", "\\dot", "\\dot{\\placeholder{}}", "\\dot{x}", "时间一阶导", "Time derivative", "physics", 86, ["dot derivative"], ["时间导数"]),
  makeCommand("timesecond", "\\ddot", "\\ddot{\\placeholder{}}", "\\ddot{x}", "时间二阶导", "Second time derivative", "physics", 84, ["double dot"], ["二阶时间导数"]),
];

type SimpleDefinition = [
  id: string,
  command: string,
  labelZh: string,
  labelEn: string,
  category: CommandCategory,
  priority?: number,
  aliases?: string[],
  keywords?: string[],
];

const simpleDefinitions: SimpleDefinition[] = [
  // 矩阵
  ["transpose", "^{\\mathsf{T}}", "转置", "Transpose", "matrix", 84, ["transpose"], ["转置"]],
  ["inverse", "^{-1}", "逆矩阵", "Inverse", "matrix", 85, ["inverse"], ["逆矩阵"]],
  ["dotproduct", "\\cdot", "点积", "Dot product", "matrix", 83, ["dot product"], ["点积", "内积"]],
  ["crossproduct", "\\times", "叉积", "Cross product", "matrix", 83, ["cross product"], ["叉积", "外积"]],

  // 希腊字母
  ["epsilon", "\\epsilon", "艾普西隆", "Epsilon", "greek", 82],
  ["varepsilon", "\\varepsilon", "变体艾普西隆", "Variant epsilon", "greek", 80],
  ["zeta", "\\zeta", "泽塔", "Zeta", "greek", 78],
  ["eta", "\\eta", "伊塔", "Eta", "greek", 78],
  ["vartheta", "\\vartheta", "变体西塔", "Variant theta", "greek", 78],
  ["iota", "\\iota", "约塔", "Iota", "greek", 76],
  ["kappa", "\\kappa", "卡帕", "Kappa", "greek", 80],
  ["nu", "\\nu", "纽", "Nu", "greek", 79],
  ["xi", "\\xi", "克西", "Xi", "greek", 76],
  ["rho", "\\rho", "柔", "Rho", "greek", 82],
  ["varrho", "\\varrho", "变体柔", "Variant rho", "greek", 72],
  ["tau", "\\tau", "陶", "Tau", "greek", 81],
  ["upsilon", "\\upsilon", "宇普西隆", "Upsilon", "greek", 70],
  ["phi", "\\phi", "斐", "Phi", "greek", 86],
  ["varphi", "\\varphi", "变体斐", "Variant phi", "greek", 84],
  ["chi", "\\chi", "希", "Chi", "greek", 74],
  ["psi", "\\psi", "普赛", "Psi", "greek", 84],
  ["Gamma", "\\Gamma", "大写伽马", "Gamma", "greek", 76],
  ["Delta", "\\Delta", "大写德尔塔", "Delta", "greek", 84],
  ["Theta", "\\Theta", "大写西塔", "Theta", "greek", 72],
  ["Lambda", "\\Lambda", "大写拉姆达", "Lambda", "greek", 76],
  ["Xi", "\\Xi", "大写克西", "Xi", "greek", 68],
  ["Pi", "\\Pi", "大写派", "Pi", "greek", 74],
  ["Sigma", "\\Sigma", "大写西格玛", "Sigma", "greek", 78],
  ["Phi", "\\Phi", "大写斐", "Phi", "greek", 76],
  ["Psi", "\\Psi", "大写普赛", "Psi", "greek", 78],
  ["Omega", "\\Omega", "大写欧米伽", "Omega", "greek", 82],

  // 关系符号
  ["lt", "<", "小于", "Less than", "relation", 92],
  ["gt", ">", "大于", "Greater than", "relation", 92],
  ["equiv", "\\equiv", "恒等于", "Equivalent", "relation", 88],
  ["sim", "\\sim", "相似", "Similar", "relation", 82],
  ["simeq", "\\simeq", "渐近相等", "Asymptotically equal", "relation", 80],
  ["cong", "\\cong", "全等", "Congruent", "relation", 78],
  ["ll", "\\ll", "远小于", "Much less", "relation", 82],
  ["gg", "\\gg", "远大于", "Much greater", "relation", 82],
  ["parallel", "\\parallel", "平行", "Parallel", "relation", 79],
  ["perp", "\\perp", "垂直", "Perpendicular", "relation", 81],
  ["mid", "\\mid", "整除", "Divides", "relation", 74],
  ["nmid", "\\nmid", "不整除", "Does not divide", "relation", 72],
  ["prec", "\\prec", "先于", "Precedes", "relation", 68],
  ["succ", "\\succ", "后于", "Succeeds", "relation", 68],

  // 集合与逻辑
  ["subseteq", "\\subseteq", "子集或相等", "Subset or equal", "set", 88],
  ["supset", "\\supset", "真超集", "Superset", "set", 86],
  ["supseteq", "\\supseteq", "超集或相等", "Superset or equal", "set", 84],
  ["emptyset", "\\emptyset", "空集", "Empty set", "set", 90],
  ["setminus", "\\setminus", "差集", "Set difference", "set", 82],
  ["land", "\\land", "逻辑与", "Logical and", "set", 86],
  ["lor", "\\lor", "逻辑或", "Logical or", "set", 86],
  ["neg", "\\neg", "逻辑非", "Logical not", "set", 84],
  ["Nset", "\\mathbb{N}", "自然数集", "Natural numbers", "set", 88],
  ["Zset", "\\mathbb{Z}", "整数集", "Integers", "set", 88],
  ["Qset", "\\mathbb{Q}", "有理数集", "Rationals", "set", 86],
  ["Rset", "\\mathbb{R}", "实数集", "Real numbers", "set", 92],
  ["Cset", "\\mathbb{C}", "复数集", "Complex numbers", "set", 90],

  // 箭头
  ["uparrow", "\\uparrow", "上箭头", "Up arrow", "arrow", 84],
  ["downarrow", "\\downarrow", "下箭头", "Down arrow", "arrow", 84],
  ["updownarrow", "\\updownarrow", "上下箭头", "Up-down arrow", "arrow", 76],
  ["Leftarrow", "\\Leftarrow", "左双线箭头", "Left double arrow", "arrow", 82],
  ["Leftrightarrow", "\\Leftrightarrow", "等价箭头", "If and only if", "arrow", 88],
  ["mapsto", "\\mapsto", "映射到", "Maps to", "arrow", 86],
  ["longrightarrow", "\\longrightarrow", "长右箭头", "Long right arrow", "arrow", 78],
  ["longleftarrow", "\\longleftarrow", "长左箭头", "Long left arrow", "arrow", 76],
  ["hookrightarrow", "\\hookrightarrow", "右钩箭头", "Right hook arrow", "arrow", 72],
  ["rightharpoonup", "\\rightharpoonup", "右鱼叉箭头", "Right harpoon", "arrow", 70],
  ["leftharpoonup", "\\leftharpoonup", "左鱼叉箭头", "Left harpoon", "arrow", 70],
  ["rightleftarrows", "\\rightleftarrows", "可逆反应箭头", "Equilibrium arrows", "arrow", 78],

  // 物理常用
  ["planck", "h", "普朗克常数", "Planck constant", "physics", 88],
  ["kb", "k_{\\mathrm{B}}", "玻尔兹曼常数", "Boltzmann constant", "physics", 91],
  ["epsilon0", "\\varepsilon_{0}", "真空介电常数", "Vacuum permittivity", "physics", 89],
  ["mu0", "\\mu_{0}", "真空磁导率", "Vacuum permeability", "physics", 89],
  ["lightspeed", "c", "光速", "Speed of light", "physics", 86],
  ["electroncharge", "e", "元电荷", "Elementary charge", "physics", 84],
  ["laplacian", "\\nabla^{2}", "拉普拉斯算子", "Laplacian", "physics", 88],
  ["divergence", "\\nabla\\cdot", "散度", "Divergence", "physics", 86],
  ["curl", "\\nabla\\times", "旋度", "Curl", "physics", 86],
];

const simpleCommands = simpleDefinitions.map(
  ([id, command, labelZh, labelEn, category, priority = 72, aliases = [], keywords = []]) =>
    makeCommand(
      id,
      command,
      command,
      command,
      labelZh,
      labelEn,
      category,
      priority,
      aliases,
      keywords,
    ),
);

export const additionalCommands: LatexCommand[] = [
  ...complexCommands,
  ...simpleCommands,
];
