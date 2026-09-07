import type { UpdateCheckResult } from "./updateService";
import { CURRENT_VERSION } from "./updateService";

export const RELEASE_WELCOME_VERSION = "1.2.6";
export const RELEASE_WELCOME_STORAGE_KEY =
  `visualtex.release-welcome.${RELEASE_WELCOME_VERSION}.seen`;

const RELEASE_WELCOME_NOTES = `## 中文

### 新增功能
- 新增并完善 Office 公式复制工作流：Word / PowerPoint 中复制出的 VisualTeX 公式会获得独立身份，可分别移动、双击编辑和再次保存；在 PowerPoint 中移动复制件后编辑，不再影响原公式的位置或内容。
- LaTeX 源码区升级：加入语法高亮、环境缩进与折叠、Tab / Shift+Tab、稳定的环境自动闭合，并加强多行、行内 / 行间混合源码的双向实时同步。
- OCR 扩展 PaddleOCR-VL 1.6 与 SimpleTex；完善多行识别、align 结构、快捷 OCR、静默 OCR 和本地识别流程。
- 公式字体扩展 KaTeX、Times New Roman、Cambria Math、STIX、Palatino、Helvetica 等选择，并支持中文字体独立设置。
- 自定义符号设计器扩展系统数学字形、运行时注册、渲染和导出能力；文档导入与复杂 LaTeX 结构兼容继续增强。
- 公式工具栏新增可在箭头上方和下方分别输入条件的化学反应箭头与可逆反应箭头结构。

### Office 与稳定性改进
- PowerPoint 新建、编辑、Apply、重新打开和复制件编辑链路进一步稳定；复制公式即使与原对象曾共享名称，也会在 fork 后立即获得独立公式身份。
- Word 编号公式复制、图片公式 / OMML 互转、行内基线、编号、引用、批量重绘和恢复 LaTeX 的稳定性继续改进。
- Office 常驻进程启动时不再提前访问 Keychain；只有真正进入 OCR / 相关设置路径时才读取对应凭据，减少启动和 Office 测试时的系统密码弹窗。
- 加固 Office Session、插件更新、异常恢复和前后台窗口唤起流程，并继续保持连续编辑性能。

### 编辑器问题修复
- 修复多层成对括号中按 Space 时 MathLive 偶发把光标从内层括号跳到外层的问题；普通输入和命令确认后继续输入时保持当前括号层级。
- 恢复“重音内容输入后跳出”的预期行为：手打 \\vec、\\hat、\\bar、\\dot 等重音命令时，输入第一个内容后会按设置自动离开重音作用域，同时不影响嵌套括号中的正常输入。
- 将错误 / 兼容候选 \\dag 统一规范为标准 \\dagger；旧的 dag 搜索习惯仍可找到该命令，但保存源码使用 \\dagger。
- 修复通过公式工具栏插入 \\ket{}、\\bra{} 等 package shorthand 后，在参数 placeholder 中输入内容时 MathLive 不接受字符、导致可视公式和第三行 LaTeX 源码不实时更新的问题。
- 修复 \\cdot 等运算符在部分字体下的间距，以及 \\sqint / \\sqiint、原生候选预览、align / aligned 对齐、环境自动闭合等 MathLive 兼容问题。

### 其他
- 更新 / 首次启动提示提供微信、支付宝和 QQ 群三个二维码；有经济能力并且觉得产品不错的可以支持一下作者呀！打赏完全自愿，不影响 VisualTeX 的任何功能和正常使用。
- 本次 GitHub Release 提供 macOS Apple Silicon 安装包。

## English

### New features
- Added and hardened Office formula-copy workflows. Copied VisualTeX formulas in Word and PowerPoint receive independent identities and can be moved, double-click edited, and saved independently; moving and editing a copied PowerPoint formula no longer changes the original formula's geometry or content.
- Upgraded the LaTeX source editor with syntax highlighting, environment-aware indentation and folding, Tab / Shift+Tab, reliable environment auto-closing, and stronger live synchronization for multiline and mixed inline / display source.
- Expanded OCR with PaddleOCR-VL 1.6 and SimpleTex, including improved multiline recognition, align output, Quick OCR, Silent OCR, and local workflows.
- Added formula font choices including KaTeX, Times New Roman, Cambria Math, STIX, Palatino, and Helvetica, with independent Chinese font settings.
- Expanded the Custom Symbol Designer, runtime glyph registration/rendering, export support, document import, and complex LaTeX compatibility.
- Added chemistry reaction-arrow and reversible-reaction-arrow structures with independently editable labels above and below the arrow.

### Office and stability improvements
- Further stabilized PowerPoint create, edit, Apply, reopen, and copied-formula editing. Copied formulas now fork to a unique formula identity even when PowerPoint initially duplicates the same shape name.
- Improved Word numbered-formula copying, picture / OMML round trips, inline baselines, numbering, references, bulk redraw, and LaTeX restoration.
- The Office resident no longer touches Keychain on startup; credential access is deferred until an OCR or related settings path actually needs it.
- Hardened Office sessions, add-in updates, recovery, foreground activation, and repeated-edit performance.

### Editor bug fixes
- Fixed Space occasionally moving the caret out of the innermost paired delimiter while typing inside nested parentheses or brackets.
- Restored the expected “exit accent after input” behavior: when typing accents such as \\vec, \\hat, \\bar, or \\dot, the caret now leaves the accent scope after the first content input when that option is enabled, without breaking nested-delimiter input.
- Canonicalized the legacy / incorrect \\dag candidate to standard \\dagger. Searching for dag remains supported, while saved source uses \\dagger.
- Fixed toolbar-inserted package shorthands such as \\ket{} and \\bra{} swallowing input inside their parameter placeholder and leaving the corresponding LaTeX source row stale.
- Fixed operator spacing such as \\cdot with some fonts, special-integral rendering such as \\sqint / \\sqiint, native suggestion previews, align / aligned behavior, environment auto-closing, and other MathLive compatibility issues.

### Other
- Update and first-launch dialogs include WeChat Pay, Alipay, and QQ community QR codes. Users who have the means and enjoy the product are welcome to support the author; tipping is completely optional and never affects any VisualTeX feature or normal use.
- This GitHub Release provides the macOS Apple Silicon installer.
`;

export function shouldShowReleaseWelcome(): boolean {
  return CURRENT_VERSION === RELEASE_WELCOME_VERSION;
}

export function releaseWelcomeResult(): UpdateCheckResult {
  return {
    currentVersion: RELEASE_WELCOME_VERSION,
    latestVersion: RELEASE_WELCOME_VERSION,
    releaseUrl: `https://github.com/paulhe666/visualtex/releases/tag/v${RELEASE_WELCOME_VERSION}`,
    releaseName: `VisualTeX ${RELEASE_WELCOME_VERSION}`,
    releaseNotes: RELEASE_WELCOME_NOTES,
    publishedAt: "",
    updateAvailable: false,
  };
}
