import type { CommandCategory, LatexCommand } from "../types/command";

interface CompatibilityCommandSpec {
  id: string;
  command: string;
  insertTemplate: string;
  previewLatex: string;
  labelZh: string;
  labelEn: string;
  aliases?: string[];
  keywords?: string[];
  category?: CommandCategory;
  priority?: number;
  wrapper?: boolean;
  canonicalWrapperCommand?: string;
  rawPlaceholderTemplate?: string;
  sourceSupported?: boolean;
}

const specs: CompatibilityCommandSpec[] = [
  // unicode-math style math alphabet commands. VisualTeX keeps these spellings
  // in source while supplying rendering-only compatibility macros.
  { id: "sym-up", command: "\\symup", insertTemplate: "\\symup{\\placeholder{}}", previewLatex: "\\symup{ABC}", labelZh: "Unicode 数学正体", labelEn: "Unicode math upright", aliases: ["symup"], keywords: ["正体", "unicode math", "字体"], wrapper: true },
  { id: "sym-it", command: "\\symit", insertTemplate: "\\symit{\\placeholder{}}", previewLatex: "\\symit{ABC}", labelZh: "Unicode 数学斜体", labelEn: "Unicode math italic", aliases: ["symit"], keywords: ["斜体", "unicode math", "字体"], wrapper: true },
  { id: "sym-bf", command: "\\symbf", insertTemplate: "\\symbf{\\placeholder{}}", previewLatex: "\\symbf{ABC}", labelZh: "Unicode 数学粗体", labelEn: "Unicode math bold", aliases: ["symbf"], keywords: ["粗体", "unicode math", "字体"], wrapper: true },
  { id: "sym-bfup", command: "\\symbfup", insertTemplate: "\\symbfup{\\placeholder{}}", previewLatex: "\\symbfup{ABC}", labelZh: "Unicode 粗正体", labelEn: "Unicode bold upright", aliases: ["symbfup"], keywords: ["粗正体", "unicode math", "字体"], wrapper: true },
  { id: "sym-bfit", command: "\\symbfit", insertTemplate: "\\symbfit{\\placeholder{}}", previewLatex: "\\symbfit{ABC\\alpha}", labelZh: "Unicode 数学粗斜体", labelEn: "Unicode bold italic", aliases: ["symbfit", "simbfit"], keywords: ["粗斜体", "unicode math", "字体"], wrapper: true },
  { id: "sym-bfit-typo", command: "\\simbfit", insertTemplate: "\\symbfit{\\placeholder{}}", previewLatex: "\\symbfit{ABC\\alpha}", labelZh: "粗斜体输入别名", labelEn: "Bold italic input alias", aliases: ["simbfit"], keywords: ["粗斜体", "输入别名", "字体"], wrapper: true, canonicalWrapperCommand: "\\symbfit", sourceSupported: false },
  { id: "sym-bb", command: "\\symbb", insertTemplate: "\\symbb{\\placeholder{}}", previewLatex: "\\symbb{ABC}", labelZh: "Unicode 黑板粗体", labelEn: "Unicode double struck", aliases: ["symbb"], keywords: ["黑板粗体", "双线体", "unicode math"], wrapper: true },
  { id: "sym-cal", command: "\\symcal", insertTemplate: "\\symcal{\\placeholder{}}", previewLatex: "\\symcal{ABC}", labelZh: "Unicode 花体", labelEn: "Unicode calligraphic", aliases: ["symcal"], keywords: ["花体", "unicode math"], wrapper: true },
  { id: "sym-bfcal", command: "\\symbfcal", insertTemplate: "\\symbfcal{\\placeholder{}}", previewLatex: "\\symbfcal{ABC}", labelZh: "Unicode 粗花体", labelEn: "Unicode bold calligraphic", aliases: ["symbfcal"], keywords: ["粗花体", "unicode math"], wrapper: true },
  { id: "sym-scr", command: "\\symscr", insertTemplate: "\\symscr{\\placeholder{}}", previewLatex: "\\symscr{gG}", labelZh: "Unicode 手写体", labelEn: "Unicode script", aliases: ["symscr"], keywords: ["手写体", "unicode math"], wrapper: true },
  { id: "sym-bfscr", command: "\\symbfscr", insertTemplate: "\\symbfscr{\\placeholder{}}", previewLatex: "\\symbfscr{gG}", labelZh: "Unicode 粗手写体", labelEn: "Unicode bold script", aliases: ["symbfscr"], keywords: ["粗手写体", "unicode math"], wrapper: true },
  { id: "sym-frak", command: "\\symfrak", insertTemplate: "\\symfrak{\\placeholder{}}", previewLatex: "\\symfrak{ABC}", labelZh: "Unicode 哥特体", labelEn: "Unicode Fraktur", aliases: ["symfrak"], keywords: ["哥特体", "unicode math"], wrapper: true },
  { id: "sym-bffrak", command: "\\symbffrak", insertTemplate: "\\symbffrak{\\placeholder{}}", previewLatex: "\\symbffrak{ABC}", labelZh: "Unicode 粗哥特体", labelEn: "Unicode bold Fraktur", aliases: ["symbffrak"], keywords: ["粗哥特体", "unicode math"], wrapper: true },
  { id: "sym-sfup", command: "\\symsfup", insertTemplate: "\\symsfup{\\placeholder{}}", previewLatex: "\\symsfup{ABC}", labelZh: "Unicode 无衬线正体", labelEn: "Unicode sans upright", aliases: ["symsfup"], keywords: ["无衬线", "unicode math"], wrapper: true },
  { id: "sym-sfit", command: "\\symsfit", insertTemplate: "\\symsfit{\\placeholder{}}", previewLatex: "\\symsfit{ABC}", labelZh: "Unicode 无衬线斜体", labelEn: "Unicode sans italic", aliases: ["symsfit"], keywords: ["无衬线斜体", "unicode math"], wrapper: true },
  { id: "sym-bfsfup", command: "\\symbfsfup", insertTemplate: "\\symbfsfup{\\placeholder{}}", previewLatex: "\\symbfsfup{ABC}", labelZh: "Unicode 无衬线粗正体", labelEn: "Unicode sans bold upright", aliases: ["symbfsfup"], keywords: ["无衬线粗体", "unicode math"], wrapper: true },
  { id: "sym-bfsfit", command: "\\symbfsfit", insertTemplate: "\\symbfsfit{\\placeholder{}}", previewLatex: "\\symbfsfit{ABC}", labelZh: "Unicode 无衬线粗斜体", labelEn: "Unicode sans bold italic", aliases: ["symbfsfit"], keywords: ["无衬线粗斜体", "unicode math"], wrapper: true },
  { id: "sym-tt", command: "\\symtt", insertTemplate: "\\symtt{\\placeholder{}}", previewLatex: "\\symtt{ABC}", labelZh: "Unicode 等宽体", labelEn: "Unicode typewriter", aliases: ["symtt"], keywords: ["等宽体", "unicode math"], wrapper: true },

  // Legacy/declaration spellings are accepted as input aliases but immediately
  // canonicalized to braced math commands, so copied LaTeX remains well scoped.
  { id: "legacy-bf", command: "\\bf", insertTemplate: "\\mathbf{\\placeholder{}}", previewLatex: "\\mathbf{ABC}", labelZh: "旧式粗正体", labelEn: "Legacy bold upright", aliases: ["bf"], keywords: ["旧式", "粗体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathbf" },
  { id: "legacy-bfseries", command: "\\bfseries", insertTemplate: "\\mathbf{\\placeholder{}}", previewLatex: "\\mathbf{ABC}", labelZh: "粗体声明", labelEn: "Bold series", aliases: ["bfseries"], keywords: ["粗体声明", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathbf" },
  { id: "legacy-it", command: "\\it", insertTemplate: "\\mathit{\\placeholder{}}", previewLatex: "\\mathit{ABC}", labelZh: "旧式斜体", labelEn: "Legacy italic", aliases: ["it"], keywords: ["旧式", "斜体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathit" },
  { id: "legacy-rm", command: "\\rm", insertTemplate: "\\mathrm{\\placeholder{}}", previewLatex: "\\mathrm{ABC}", labelZh: "旧式正体", labelEn: "Legacy roman", aliases: ["rm"], keywords: ["旧式", "正体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathrm" },
  { id: "legacy-sf", command: "\\sf", insertTemplate: "\\mathsf{\\placeholder{}}", previewLatex: "\\mathsf{ABC}", labelZh: "旧式无衬线体", labelEn: "Legacy sans", aliases: ["sf"], keywords: ["旧式", "无衬线", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathsf" },
  { id: "legacy-tt", command: "\\tt", insertTemplate: "\\mathtt{\\placeholder{}}", previewLatex: "\\mathtt{ABC}", labelZh: "旧式等宽体", labelEn: "Legacy typewriter", aliases: ["tt"], keywords: ["旧式", "等宽体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathtt" },
  { id: "legacy-cal", command: "\\cal", insertTemplate: "\\mathcal{\\placeholder{}}", previewLatex: "\\mathcal{ABC}", labelZh: "旧式花体", labelEn: "Legacy calligraphic", aliases: ["cal"], keywords: ["旧式", "花体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathcal" },
  { id: "legacy-Bbb", command: "\\Bbb", insertTemplate: "\\mathbb{\\placeholder{}}", previewLatex: "\\mathbb{ABC}", labelZh: "黑板粗体简称", labelEn: "Blackboard bold alias", aliases: ["Bbb"], keywords: ["简称", "黑板粗体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathbb" },
  { id: "legacy-frak", command: "\\frak", insertTemplate: "\\mathfrak{\\placeholder{}}", previewLatex: "\\mathfrak{ABC}", labelZh: "哥特体简称", labelEn: "Fraktur alias", aliases: ["frak"], keywords: ["简称", "哥特体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathfrak" },
  { id: "legacy-bold", command: "\\bold", insertTemplate: "\\mathbfit{\\placeholder{}}", previewLatex: "\\mathbfit{A\\alpha}", labelZh: "数学粗体简称", labelEn: "Bold math alias", aliases: ["bold"], keywords: ["简称", "粗斜体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathbfit" },
  { id: "legacy-boldmath", command: "\\boldmath", insertTemplate: "\\mathbfit{\\placeholder{}}", previewLatex: "\\mathbfit{A\\alpha}", labelZh: "粗数学声明", labelEn: "Bold math declaration", aliases: ["boldmath"], keywords: ["粗斜体", "数学版本", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathbfit" },
  { id: "poor-mans-bold", command: "\\pmb", insertTemplate: "\\mathbfit{\\placeholder{}}", previewLatex: "\\mathbfit{A\\alpha}", labelZh: "模拟数学粗体", labelEn: "Poor man's bold", aliases: ["pmb"], keywords: ["粗斜体", "粗体", "字体"], wrapper: true, canonicalWrapperCommand: "\\mathbfit" },

  // MathLive 0.109 does not register the amsmath modulo family even though the
  // MathJax/OMML export path accepts it. Expose all four source spellings so
  // command completion, raw-source validation and live rendering agree.
  { id: "operator-bmod", command: "\\bmod", insertTemplate: "\\bmod", previewLatex: "a\\bmod b", labelZh: "二元模运算", labelEn: "Binary modulo operator", aliases: ["bmod", "modulo"], keywords: ["模", "取模", "余数", "modulo"], category: "relation", priority: 90 },
  { id: "operator-mod", command: "\\mod", insertTemplate: "\\mod", previewLatex: "a\\equiv b\\mod n", labelZh: "模同余后缀", labelEn: "Modulo congruence suffix", aliases: ["mod"], keywords: ["模", "同余", "modulo"], category: "relation", priority: 84 },
  { id: "operator-pmod", command: "\\pmod", insertTemplate: "\\pmod{\\placeholder{}}", previewLatex: "a\\equiv b\\pmod n", labelZh: "括号模同余", labelEn: "Parenthesized modulo", aliases: ["pmod"], keywords: ["模", "同余", "括号", "modulo"], category: "relation", priority: 88, wrapper: true },
  { id: "operator-pod", command: "\\pod", insertTemplate: "\\pod{\\placeholder{}}", previewLatex: "a\\equiv b\\pod n", labelZh: "括号同余参数", labelEn: "Parenthesized congruence argument", aliases: ["pod"], keywords: ["同余", "括号", "modulo"], category: "relation", priority: 76, wrapper: true },

  // Common package shorthand commands already supported by the compatibility
  // macro layer. Registering them here makes raw-command completion and Space
  // confirmation consistent with \mathbb and \bm.
  { id: "shortcut-abs", command: "\\abs", insertTemplate: "\\abs{\\placeholder{}}", previewLatex: "\\abs{x}", labelZh: "绝对值简称", labelEn: "Absolute value shorthand", aliases: ["abs"], keywords: ["physics", "简称", "绝对值"], category: "structure", priority: 86, wrapper: true },
  { id: "shortcut-norm", command: "\\norm", insertTemplate: "\\norm{\\placeholder{}}", previewLatex: "\\norm{\\mathbf{x}}", labelZh: "范数简称", labelEn: "Norm shorthand", aliases: ["norm"], keywords: ["physics", "简称", "范数"], category: "matrix", priority: 86, wrapper: true },
  { id: "shortcut-dd", command: "\\dd", insertTemplate: "\\dd{\\placeholder{}}", previewLatex: "\\dd{x}", labelZh: "微分元简称", labelEn: "Differential shorthand", aliases: ["dd"], keywords: ["physics", "简称", "微分"], category: "calculus", priority: 88, wrapper: true },
  { id: "shortcut-bra", command: "\\bra", insertTemplate: "\\bra{\\placeholder{}}", previewLatex: "\\bra{\\psi}", labelZh: "Bra 简称", labelEn: "Bra shorthand", aliases: ["bra"], keywords: ["physics", "狄拉克"], category: "physics", priority: 90, wrapper: true },
  { id: "shortcut-ket", command: "\\ket", insertTemplate: "\\ket{\\placeholder{}}", previewLatex: "\\ket{\\psi}", labelZh: "Ket 简称", labelEn: "Ket shorthand", aliases: ["ket"], keywords: ["physics", "狄拉克"], category: "physics", priority: 90, wrapper: true },
  { id: "shortcut-expval", command: "\\expval", insertTemplate: "\\expval{\\placeholder{}}", previewLatex: "\\expval{A}", labelZh: "期望值简称", labelEn: "Expectation shorthand", aliases: ["expval"], keywords: ["physics", "期望值"], category: "physics", priority: 84, wrapper: true },
  { id: "shortcut-vb", command: "\\vb", insertTemplate: "\\vb{\\placeholder{}}", previewLatex: "\\vb{v}", labelZh: "粗向量简称", labelEn: "Bold vector shorthand", aliases: ["vb"], keywords: ["physics", "向量", "简称"], category: "physics", priority: 84, wrapper: true },
  { id: "shortcut-va", command: "\\va", insertTemplate: "\\va{\\placeholder{}}", previewLatex: "\\va{v}", labelZh: "箭头粗向量简称", labelEn: "Arrow vector shorthand", aliases: ["va"], keywords: ["physics", "向量", "简称"], category: "physics", priority: 82, wrapper: true },
  { id: "shortcut-vu", command: "\\vu", insertTemplate: "\\vu{\\placeholder{}}", previewLatex: "\\vu{e}", labelZh: "单位向量简称", labelEn: "Unit vector shorthand", aliases: ["vu"], keywords: ["physics", "单位向量", "简称"], category: "physics", priority: 82, wrapper: true },

  { id: "shortcut-comm", command: "\\comm", insertTemplate: "\\comm{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\comm{A}{B}", labelZh: "对易子简称", labelEn: "Commutator shorthand", aliases: ["comm"], keywords: ["physics", "对易子"], category: "physics", priority: 84, rawPlaceholderTemplate: "\\comm{\\placeholder{}}{\\placeholder{}}" },
  { id: "shortcut-acomm", command: "\\acomm", insertTemplate: "\\acomm{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\acomm{A}{B}", labelZh: "反对易子简称", labelEn: "Anticommutator shorthand", aliases: ["acomm"], keywords: ["physics", "反对易子"], category: "physics", priority: 80, rawPlaceholderTemplate: "\\acomm{\\placeholder{}}{\\placeholder{}}" },
  { id: "shortcut-pb", command: "\\pb", insertTemplate: "\\pb{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\pb{f}{g}", labelZh: "泊松括号简称", labelEn: "Poisson bracket shorthand", aliases: ["pb"], keywords: ["physics", "泊松括号"], category: "physics", priority: 78, rawPlaceholderTemplate: "\\pb{\\placeholder{}}{\\placeholder{}}" },
  { id: "shortcut-dv", command: "\\dv", insertTemplate: "\\dv{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\dv{f}{x}", labelZh: "导数简称", labelEn: "Derivative shorthand", aliases: ["dv"], keywords: ["physics", "导数"], category: "calculus", priority: 88, rawPlaceholderTemplate: "\\dv{\\placeholder{}}{\\placeholder{}}" },
  { id: "shortcut-pdv", command: "\\pdv", insertTemplate: "\\pdv{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\pdv{f}{x}", labelZh: "偏导简称", labelEn: "Partial derivative shorthand", aliases: ["pdv"], keywords: ["physics", "偏导"], category: "calculus", priority: 88, rawPlaceholderTemplate: "\\pdv{\\placeholder{}}{\\placeholder{}}" },
  { id: "shortcut-braket", command: "\\braket", insertTemplate: "\\braket{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\braket{\\phi}{\\psi}", labelZh: "内积简称", labelEn: "Bra-ket shorthand", aliases: ["braket"], keywords: ["physics", "内积"], category: "physics", priority: 88, rawPlaceholderTemplate: "\\braket{\\placeholder{}}{\\placeholder{}}" },
  { id: "shortcut-ketbra", command: "\\ketbra", insertTemplate: "\\ketbra{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\ketbra{\\psi}{\\phi}", labelZh: "外积简称", labelEn: "Ket-bra shorthand", aliases: ["ketbra"], keywords: ["physics", "外积"], category: "physics", priority: 86, rawPlaceholderTemplate: "\\ketbra{\\placeholder{}}{\\placeholder{}}" },
  { id: "shortcut-mel", command: "\\mel", insertTemplate: "\\mel{\\placeholder{}}{\\placeholder{}}{\\placeholder{}}", previewLatex: "\\mel{\\phi}{A}{\\psi}", labelZh: "矩阵元简称", labelEn: "Matrix element shorthand", aliases: ["mel"], keywords: ["physics", "矩阵元"], category: "physics", priority: 86, rawPlaceholderTemplate: "\\mel{\\placeholder{}}{\\placeholder{}}{\\placeholder{}}" },
];

function toLatexCommand(spec: CompatibilityCommandSpec): LatexCommand {
  return {
    id: spec.id,
    command: spec.command,
    insertTemplate: spec.insertTemplate,
    previewLatex: spec.previewLatex,
    labelZh: spec.labelZh,
    labelEn: spec.labelEn,
    aliases: spec.aliases ?? [spec.command.replace(/^\\/, "")],
    keywords: spec.keywords ?? ["字体"],
    category: spec.category ?? "matrix",
    defaultPriority: spec.priority ?? 72,
    supportedInMathMode: true,
  };
}

export const compatibilityCommands: LatexCommand[] = specs.map(toLatexCommand);

export const compatibilityCommandNames = new Set(
  specs
    .filter((spec) => spec.sourceSupported !== false)
    .map((spec) => spec.command.replace(/^\\/, "")),
);

export const compatibilityRequiredArgumentCounts = new Map<string, number>([
  ["bm", 1],
  ["mathbfit", 1],
  ...specs.flatMap((spec) => {
    if (spec.sourceSupported === false) return [];
    const count = spec.rawPlaceholderTemplate
      ? spec.rawPlaceholderTemplate.match(/\\placeholder\{\}/g)?.length ?? 0
      : spec.wrapper
        ? 1
        : 0;
    return count > 0
      ? [[spec.command.replace(/^\\/, ""), count] as const]
      : [];
  }),
]);

export const compatibilityWrapperPreviews = new Map<string, string>(
  specs
    .filter((spec) => spec.wrapper)
    .map((spec) => [spec.command, spec.previewLatex]),
);

export const compatibilityWrapperCanonicalTargets = new Map<string, string>(
  specs
    .filter((spec) => spec.wrapper && spec.canonicalWrapperCommand)
    .map((spec) => [spec.command, spec.canonicalWrapperCommand!]),
);

export const compatibilityRawPlaceholderTemplates = new Map<string, string>(
  specs
    .filter((spec) => spec.rawPlaceholderTemplate)
    .map((spec) => [spec.command, spec.rawPlaceholderTemplate!]),
);
