import type { CommandCategory } from "../types/command";

export interface VisualTexCoreLatexAlias {
  id: string;
  name: string;
  replacement: string;
  previewLatex: string;
  labelZh: string;
  labelEn: string;
  aliases: string[];
  keywords: string[];
  category: CommandCategory;
  priority: number;
}

/**
 * Common LaTeX/AMS spellings that MathLive 0.109.2 does not parse natively.
 *
 * Keep these aliases atomic: they should preserve the user's source command
 * while rendering/editing through an equivalent primitive already supported by
 * MathLive. This registry is shared by the MathLive macro layer and native
 * semantic completion so support cannot drift between parsing and discovery.
 */
export interface VisualTexCoreMathLiveMacro {
  name: string;
  def: string;
  args: number;
}

export const VISUALTEX_CORE_LATEX_ALIASES: readonly VisualTexCoreLatexAlias[] =
  Object.freeze([
    {
      id: "dots",
      name: "dots",
      replacement: "\\ldots",
      previewLatex: "\\ldots",
      labelZh: "省略号",
      labelEn: "Contextual ellipsis",
      aliases: ["dots", "ellipsis"],
      keywords: ["省略号", "ellipsis", "amsmath"],
      category: "common",
      priority: 96,
    },
    {
      id: "dotsc",
      name: "dotsc",
      replacement: "\\ldots",
      previewLatex: "\\ldots",
      labelZh: "逗号省略号",
      labelEn: "Comma ellipsis",
      aliases: ["dotsc"],
      keywords: ["省略号", "comma ellipsis", "amsmath"],
      category: "common",
      priority: 82,
    },
    {
      id: "dotso",
      name: "dotso",
      replacement: "\\ldots",
      previewLatex: "\\ldots",
      labelZh: "普通省略号",
      labelEn: "Other ellipsis",
      aliases: ["dotso"],
      keywords: ["省略号", "ellipsis", "amsmath"],
      category: "common",
      priority: 80,
    },
    {
      id: "dotsb",
      name: "dotsb",
      replacement: "\\cdots",
      previewLatex: "\\cdots",
      labelZh: "二元运算省略号",
      labelEn: "Binary-operator ellipsis",
      aliases: ["dotsb"],
      keywords: ["省略号", "binary ellipsis", "amsmath"],
      category: "common",
      priority: 82,
    },
    {
      id: "dotsm",
      name: "dotsm",
      replacement: "\\cdots",
      previewLatex: "\\cdots",
      labelZh: "乘法省略号",
      labelEn: "Multiplication ellipsis",
      aliases: ["dotsm"],
      keywords: ["省略号", "multiplication ellipsis", "amsmath"],
      category: "common",
      priority: 81,
    },
    {
      id: "dotsi",
      name: "dotsi",
      replacement: "\\cdots",
      previewLatex: "\\cdots",
      labelZh: "积分省略号",
      labelEn: "Integral ellipsis",
      aliases: ["dotsi"],
      keywords: ["省略号", "integral ellipsis", "amsmath"],
      category: "common",
      priority: 81,
    },
    {
      id: "qed",
      name: "qed",
      replacement: "\\square",
      previewLatex: "\\square",
      labelZh: "证毕符号",
      labelEn: "QED symbol",
      aliases: ["qed"],
      keywords: ["证毕", "proof", "amsthm"],
      category: "common",
      priority: 72,
    },
    {
      id: "qedsymbol",
      name: "qedsymbol",
      replacement: "\\square",
      previewLatex: "\\square",
      labelZh: "证毕符号",
      labelEn: "QED symbol",
      aliases: ["qedsymbol"],
      keywords: ["证毕", "proof", "amsthm"],
      category: "common",
      priority: 70,
    },
    {
      id: "diagdown",
      name: "diagdown",
      replacement: "\\backslash",
      previewLatex: "\\backslash",
      labelZh: "下降对角线",
      labelEn: "Diagonal down",
      aliases: ["diagdown"],
      keywords: ["对角线", "diagonal", "amssymb"],
      category: "common",
      priority: 68,
    },
    {
      id: "thinspace",
      name: "thinspace",
      replacement: "\\,",
      previewLatex: "A\\,B",
      labelZh: "细空格",
      labelEn: "Thin space",
      aliases: ["thinspace"],
      keywords: ["空格", "spacing", "amsmath"],
      category: "common",
      priority: 64,
    },
    {
      id: "negthinspace",
      name: "negthinspace",
      replacement: "\\!",
      previewLatex: "A\\!B",
      labelZh: "负细空格",
      labelEn: "Negative thin space",
      aliases: ["negthinspace"],
      keywords: ["空格", "spacing", "amsmath"],
      category: "common",
      priority: 62,
    },
    {
      id: "medspace",
      name: "medspace",
      replacement: "\\:",
      previewLatex: "A\\:B",
      labelZh: "中空格",
      labelEn: "Medium space",
      aliases: ["medspace"],
      keywords: ["空格", "spacing", "amsmath"],
      category: "common",
      priority: 62,
    },
    {
      id: "negmedspace",
      name: "negmedspace",
      replacement: "\\mkern-4mu",
      previewLatex: "A\\mkern-4mu B",
      labelZh: "负中空格",
      labelEn: "Negative medium space",
      aliases: ["negmedspace"],
      keywords: ["空格", "spacing", "amsmath"],
      category: "common",
      priority: 60,
    },
    {
      id: "thickspace",
      name: "thickspace",
      replacement: "\\;",
      previewLatex: "A\\;B",
      labelZh: "宽空格",
      labelEn: "Thick space",
      aliases: ["thickspace"],
      keywords: ["空格", "spacing", "amsmath"],
      category: "common",
      priority: 62,
    },
    {
      id: "negthickspace",
      name: "negthickspace",
      replacement: "\\mkern-5mu",
      previewLatex: "A\\mkern-5mu B",
      labelZh: "负宽空格",
      labelEn: "Negative thick space",
      aliases: ["negthickspace"],
      keywords: ["空格", "spacing", "amsmath"],
      category: "common",
      priority: 60,
    },
  ]);

export const VISUALTEX_CORE_MATHLIVE_MACROS: readonly VisualTexCoreMathLiveMacro[] =
  Object.freeze([
    ...VISUALTEX_CORE_LATEX_ALIASES.map((alias) => ({
      name: alias.name,
      def: alias.replacement,
      args: 0,
    })),
    { name: "pod", def: "\\quad(#1)", args: 1 },
    { name: "substack", def: "\\begin{array}{c}#1\\end{array}", args: 1 },
    { name: "prescript", def: "{}^{#1}_{#2}#3", args: 3 },
    { name: "sideset", def: "{}#1#3#2", args: 3 },
  ]);
