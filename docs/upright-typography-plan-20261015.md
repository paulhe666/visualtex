# 输入判定：快捷转义与正体/斜体的统一规则（2026-10-15）

范围：Windows 可视化公式编辑器中，键入的字母词变成什么（快捷转义：alpha→α、frac→分式），以及哪些符号排成正体（单位、函数名、常数、标签、微分 d）。两者由**同一个判定程序、同一部词典**决定，并内建在 MathLive 内核里。
依据：SI 手册（BIPM 第 9 版）、ISO 80000-1/-2、GB 3101/3102（与 ISO 一致）、IUPAC 绿皮书（pH）。

> **决定（2026-10-15）**
> - A1：整条公式只有一个单位组（`m/s`、`N·m`）时按单位处理（正体）。
> - B1：π 保持常见的斜体。
> - 时机：字母词**写完后**才判定（键入非字母、移动光标、离开公式），不再逐字母触发。
> - 删除：`ge le ne ll gg pp ss mm dd nn uu vv AA EE TT ch lt gt or and not of del oo sub`，以及 MathLive 计算器用的 `mean median fft bessel randomReal randomInteger`；`sup` 改为上确界函数（不再是 ⊃）。
> - `int sum prod lim` 作为单词插入裸算子，上下限用 `_`、`^` 输入；带上下限框的版本仍在工具栏和候选框（`defint` 为带上下限的积分）。
> - 两个开关：“常用数学快捷转义”与“自动正体排版”，都送入同一判定程序；工具栏和候选框插入的内容不受影响。

---

## 1. 一条总规则

**符号代表一个固定的东西 → 正体；代表一个可变的值 → 斜体。同一拼写两者皆可时，默认斜体，除非上下文表明它是固定的东西。** 快捷转义是同一件事的另一面：一个**完整的词**若是词典里的名字（alpha、infty、frac），就换成它所代表的符号。

## 2. 一个判定过程

对每个写完的字母词，按同样顺序：

1. 用户明确指定了样式（`\mathit`、`\mathrm`、`\text`、`\operatorname` 等）→ 保留，永不改写。
2. 查词典（`inputDictionary.ts`）：单位、函数、常数、标签、希腊字母、符号词、结构模板、符号组合，每条记录类别与“确定/歧义”。
3. 单位组（`/`、`·`、空格、指数连起来）：有确定成员（km、kg）、或前面是数字+空格、或整条公式只有这一组（A1）→ 整组正体。
4. 下标里整个是一个标签词（`v_max`）→ 正体；`pH` 任何位置都正体。
5. 仅键入时：`30 deg`→30°；`ee ii jj`→常数；项首的 `dx`、`dt`、`dθ`→正体 d。
6. 完整的词是名字（键入时）：alpha、infty、in、frac……
7. 函数名：确定的总是正体；歧义的（lg、Re、Im、tr）只在后面跟参数时（`lg(x)`、`lg x`）。
8. 名字链（键入时）：从词首开始的名字，一个接一个，再加参数：`sintheta`→sin θ、`pir`→π r、`intx`→∫ x。词中间的名字永远不识别：`spin`、`point`、`for`、`number` 保持原样。
9. 常数：e 作为指数函数的底且指数含变量/i/负号（`e^{-t}`；`e^2` 保持斜体）；i 在 e 的指数里、`\exp(i…)`、`3+4i`。
10. 微分 d：导数分式、商、积分测度（积分里 `xdx` 的 d 也识别）。

快捷转义（第 5、6、8 步）只用于在编辑器里键入的内容；粘贴、OCR、Word 导入和保存的 LaTeX 只做正体判定。

## 3. 何时判定

- 字母词写完时：键入任何非字母键（运算符、括号、数字、`^`、`_`、空格、Tab、方向键）、点击移动光标、离开公式。
- 结束一个**被转换**的词的空格不插入（`pi r` → πr）；单位后的空格保留（`72 km h`）。
- 转换成模板时光标进入第一个空位（`frac` 空格 → 分子）。
- 转换后立即按 Backspace：恢复为键入的字母并标成显式斜体（`\mathit{pi}`），之后不再转换。
- 符号组合（`>=`、`->`、`!=`、`+-`、`...`）键入即转换。

## 4. 实现

| 位置 | 内容 |
|---|---|
| `apps/windows/src/editor/inputDictionary.ts` | 唯一词典；设置页“映射”列表也由它生成 |
| `apps/windows/src/editor/inputJudgment.ts` | 唯一判定程序 `applyInputJudgment()`（含原微分 d 代码） |
| `vite.mathliveIntegralCompatibility.ts` | 内核加载时 `import` 判定模块 |
| kernel `part-007` | MathLive 自带 `INLINE_SHORTCUTS` 整表换成判定模块的符号组合表 |
| kernel `part-020` `onKeystroke` | 结束词的按键先运行判定；用掉的空格/Backspace 在此吞掉 |
| kernel `part-021` | `visualTexJudgeInput`、`visualTexJudgeKeystroke`、撤销转换、命令 `visualTexJudgeInput` |
| kernel `part-026` `setValue` | 清空快捷键缓冲（修复 hc 之后键入 hnu 拼出 h·ch·ν） |
| kernel `part-027` | `onBlur` 时判定；`MathfieldElement.visualTexJudgeKeystroke()` 供宿主先处理的键（Space、Backspace）调用 |
| `MathEditor.tsx` | 把两个开关传给内核（`visualTexJudgmentSettings`）；点击移动光标时调用内核判定 |
| `normalizeChineseLatex.ts` | 保存/导入路径调用同一判定程序（只做正体） |
| 设置 `autoUprightTypography` | 新开关“自动正体排版”（默认开）；关闭后不做任何正体判定 |

## 5. 回归

- `scripts/input_judgment_regression.mjs`：真实浏览器键入 84 例（审计语料 + 空格/模板/Backspace + 四种开关组合）。
- `scripts/targeted_upright_typography_acceptance.mjs`：保存/导入路径 48 例（原有 23 例 + 新增 25 例）。
- 已更新的旧测试：`targeted_editor_regression.mjs direct-shortcut-placeholder`（单词写完后转换；sum/int 裸算子；defint 带上下限）、`toolbar-placeholder-overflow`（`frac ` 加空格）、`input_behavior_regression.mjs`（设置项 7→8）。

## 6. 暂不做（下一步）

- `\si{km/h}`、`\SI{72}{km/h}` 的参数按同一单位语法解析。
- μ 前缀（`\mu m`）、Ω（`k\Omega`）、°C、化学式、单字母描述性下标（`T_c`）。
- 转换后用 `+` 等字符结束的词，Backspace 删除该字符而不是撤销转换（只有用空格结束的转换可以 Backspace 撤销）。
- macOS 版同步。
