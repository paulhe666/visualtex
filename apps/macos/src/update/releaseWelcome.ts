import type { UpdateCheckResult } from "./updateService";
import { CURRENT_VERSION } from "./updateService";

export const RELEASE_WELCOME_VERSION = "1.2.6";
export const RELEASE_WELCOME_STORAGE_KEY =
  `visualtex.release-welcome.${RELEASE_WELCOME_VERSION}.seen`;

const RELEASE_WELCOME_NOTES = `## 中文

### 新增功能
- LaTeX 源码区升级：语法高亮、环境缩进、折叠、Tab / Shift+Tab，以及更稳定的自动闭合。
- OCR 新增 PaddleOCR-VL 1.6 与 SimpleTex；多行识别可直接生成 align 等结构，快捷 OCR 与静默 OCR 同步完善。
- 公式字体新增 KaTeX、Times New Roman、Cambria Math、STIX、Palatino、Helvetica 等选择，中文字体可独立设置。
- Word 公式引用支持直接选择；文档导入、批量重绘、编号、图片公式与 OMML 工作流继续完善。
- 自定义符号设计器扩展系统数学字形与运行时渲染能力。

### 问题修复
- 修复 \\cdot 等运算符在部分公式字体下间距异常。
- 修复 \\sqint、\\sqiint 等特殊积分及原生候选预览的字形、尺寸和快捷键排版。
- 修复 align / aligned 对齐、环境自动闭合及多项 MathLive 输入兼容问题。
- 修复 Word 图片 / OMML 公式基线、编辑性能、编号与引用稳定性，以及恢复 LaTeX 时的编号残留。
- 加固 Office 会话、插件更新和异常恢复流程。

### 其他
- 更新提示加入微信、支付宝和 QQ 群二维码。
- 本次 Release 提供 macOS Apple Silicon 安装包。

## English

### New features
- Upgraded the LaTeX source editor with syntax highlighting, environment-aware indentation, folding, Tab / Shift+Tab, and more reliable auto-closing.
- Added PaddleOCR-VL 1.6 and SimpleTex OCR, with improved multiline, Quick OCR, and Silent OCR workflows.
- Added formula font choices including KaTeX, Times New Roman, Cambria Math, STIX, Palatino, and Helvetica, with independent Chinese font settings.
- Improved Word equation references, document import, batch redraw, numbering, picture formulas, and OMML workflows.
- Expanded the Custom Symbol Designer with system math glyphs and runtime rendering improvements.

### Bug fixes
- Fixed spacing around operators such as \\cdot with some formula fonts.
- Fixed glyph, sizing, and shortcut layout issues for special integrals such as \\sqint and \\sqiint.
- Fixed align / aligned alignment, environment auto-closing, and several MathLive input compatibility issues.
- Fixed Word picture / OMML baselines, edit performance, numbering and reference stability, and equation-number remnants when restoring LaTeX.
- Hardened Office sessions, add-in updates, and recovery paths.

### Other
- Added WeChat Pay, Alipay, and QQ group QR codes to the update experience.
- This release provides the macOS Apple Silicon installer.
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
