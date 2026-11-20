# VisualTeX 1.2.5

## 中文

### 1. 编辑器与交互

- 新增 **小键盘模式**，为紧凑公式输入重新设计布局；普通模式与小键盘模式分别记忆窗口尺寸，互不影响。
- 新增 **显式对齐点编辑**，`align` / `aligned` 等多行公式可直接使用 `&` 指定对齐位置。
- 加强 **`cases` / 分段函数编辑**，完善多行结构、对齐点、大括号与 Enter 行为。
- 扩展公式工具栏和命令兼容，补充更多扩展积分符号、宏包命令及对应的渲染、复制和导出支持。
- 统一主应用与 Office 编辑窗口的默认参数、输入行为、公式工具栏和格式状态。
- 完善 **自定义公式快捷键**、常用工具和 **希腊字母快捷输入**；快捷键配置统一管理并支持持久化。
- 工作区布局、面板状态、公式颜色、字号、工具栏状态、编辑模式等更多参数支持跨重启记忆。
- 改进紧凑窗口和浮层布局，降低菜单、工具栏、颜色面板等在小尺寸窗口中的裁切和遮挡。

### 2. 自定义符号系统

- 新增完整的 **自定义符号设计器**，可在画布中绘制并组合自定义数学符号。
- 自定义符号可注册为 VisualTeX / LaTeX 字符，并进入公式工具栏、自动补全、MathLive 渲染以及 SVG / PNG 导出链路。
- 新增自定义符号存档、运行时注册、字形编译和隔离机制，避免影响内置字符。
- 优化有效可视区域裁剪、工具栏尺寸适配、清晰度和基线定位，避免按整张画布渲染造成过宽、过小或位置偏移。
- 改进绘制、缩放、旋转、层级、参考字符、橡皮擦等编辑体验。

### 3. 主题与界面自定义

- 重构主题系统，由基础明暗主题扩展为多套配色与完整的 **主题颜色自定义**。
- 主背景、纸张、抬升层、凹陷层、选中状态、强调色、公式光标等统一纳入主题体系。
- 设置页、引导页、编辑器、公式工具栏、帮助窗口等界面统一跟随当前主题。
- 进一步压缩设置页和编辑窗口中的非必要说明文字，优化信息层级和可用空间。

### 4. OCR

- 新增 **快捷 OCR**，优化截图、识别、返回编辑器和结果回填流程。
- 新增 **静默 OCR**，可在后台完成截图识别，并将 LaTeX 结果直接写入剪贴板。
- 新增 OCR 状态 / HUD，显示识别过程与完成状态。
- 统一普通 OCR、快捷 OCR、静默 OCR 的模型、设置和主界面入口。
- 优化 OCR 模型切换、窗口切换、截图触发和识别结果回填的稳定性。

### 5. Office 编辑器启动与性能

- Word / PowerPoint 公式编辑器采用更成熟的 **常驻 / 预热式生命周期**，减少重复冷启动等待。
- 修复 Office 打开编辑器时主应用抢前台、窗口顺序错误、图标异常等问题。
- 编辑窗口可直接浮在 Word / PowerPoint 上方，编辑完成后焦点正确返回 Office。
- 针对真正冷启动场景进一步优化启动稳定性。
- 加入 Office 性能预算与回归测试，降低公式数量增长后出现线性性能退化的风险。

### 6. Word Ribbon 与 Office 集成

- 重整 Word Ribbon，并统一 VisualTeX 图标与功能分组。
- 主要入口覆盖：图片行内、图片行间、OMML 行内、OMML 行间、编辑所选、重绘所选、重绘全文、图片 / OMML 转换、公式交叉引用、更新编号、编号格式和文档导入。
- Office 对话框与主应用进一步统一公式工具栏、参数、字体、字号、主题和输入行为。
- 减少公式连续打开与编辑过程中的重复初始化，提高连续编辑效率。

### 7. Office 插件安装、更新与修复

- 新增更完整的 **DOTM / PPAM 版本检测机制**。
- 可识别插件未安装、版本过旧、文件不完整，以及 PowerPoint 插件未注册等状态。
- 旧版 DOTM / PPAM 可由 VisualTeX 直接覆盖升级，无需用户手工删除。
- Office 插件入口区分 **安装 / 更新 / 修复** 状态。
- 更新插件时可请求退出 Word / PowerPoint，并在 Office 正常退出后继续安装。
- 插件更新选择和相关用户偏好支持持久化。

### 8. Word 重绘与文档导入

- 继续加强 **重绘所选 / 重绘全文**，提升大量公式和复杂文档下的稳定性。
- 修复重绘后段落布局、图片公式位置和双击编辑目标发生变化的问题。
- 加强正文、行内公式、行间公式、OMML 等混合内容的导入与回归测试。
- 通过更严格的公式 identity / metadata 保持公式身份和原始结构，减少重绘后公式与文本错位。
- 优化批量处理中的目标定位和恢复逻辑，不再仅依赖公式当前位置。

### 9. Word 公式编号与交叉引用

- 编号设置由临时状态改为 **持久化配置**。
- 支持并稳定化顺序编号、按章节编号、按节编号，以及 `.` / `-` 等不同分隔形式。
- Word Ribbon 的编号格式与实际文档状态进一步同步。
- **更新编号**升级为全局修复入口，可按公式在文档中的物理顺序重新整理编号。
- 优化大量公式下的编号更新、连续插入和编辑性能；常规尾部插入优先使用局部快速路径，避免每次全文扫描。
- 强化图片公式与 OMML 公式混排、中间插入公式后的编号顺序。
- 修复旧编号结构、孤立 `SEQ`、损坏 helper 和重复编号等历史遗留状态。
- 加强失败事务回滚，避免编号操作失败后留下不完整结构。
- 加强 **交叉引用与编号联动**：中间插入、插入引用、再次编辑和更新编号后，公式数量、`SEQ` 与正文 `REF` 可保持一致。

### 10. 帮助手册

- 新增软件内 **帮助手册** 入口，直接在 VisualTeX 中查看，无需跳转浏览器。
- 手册覆盖编辑器、快捷键、常用工具、磁贴、`align`、`cases`、矩阵、小键盘、OCR、自定义符号、Word 和 PowerPoint 等主要功能。
- 帮助界面跟随当前主题，并针对正文、目录、标题、代码块和表格优化字号与可读性。

### Windows 端补充

- **PowerPoint 格式转换**：新增图片、OLE、OMML 公式之间的格式转换入口，便于在演示文稿中按编辑性、兼容性和输出需求切换公式格式。

---

## English

### 1. Editor and interaction

- Added **Keypad Mode**, with a compact layout designed for focused formula entry. Normal Mode and Keypad Mode now keep independent window-size memory.
- Added **explicit alignment-point editing** for multiline environments such as `align` / `aligned`; use `&` to choose the visual alignment position directly.
- Improved **`cases` / piecewise-function editing**, including multiline structure, alignment points, braces, and Enter behavior.
- Expanded the formula toolbar and command compatibility with more extended integral symbols, package commands, and matching rendering, copy, and export support.
- Unified default parameters, input behavior, toolbar state, and formatting behavior between the main app and Office editor windows.
- Improved **custom formula hotkeys**, Common tools, and **Greek-letter quick input**, with centralized persistent configuration.
- More workspace settings now persist across restarts, including layout, panel state, formula colors, font size, toolbar state, and editor mode.
- Improved compact-window and floating-layer layout to reduce clipping and overlap in menus, toolbars, and color panels.

### 2. Custom symbol system

- Added a complete **Custom Symbol Designer** for drawing and composing mathematical symbols on a canvas.
- Custom symbols can be registered as VisualTeX / LaTeX characters and used throughout the formula toolbar, autocomplete, MathLive rendering, and SVG / PNG export pipeline.
- Added symbol archives, runtime registration, glyph compilation, and isolation from built-in symbols.
- Improved visible-bounds cropping, toolbar sizing, sharpness, and baseline placement to avoid oversized, undersized, or offset results caused by full-canvas rendering.
- Improved drawing, scaling, rotation, layer management, reference glyphs, and eraser behavior.

### 3. Themes and interface customization

- Reworked the theme system from basic light/dark modes into multiple palettes with full **theme-color customization**.
- Main background, paper, elevated and recessed surfaces, selection states, accents, and formula-caret colors now share one theme system.
- Settings, onboarding, editor, formula toolbar, and Help Manual now follow the active theme consistently.
- Reduced unnecessary explanatory text in settings and editor windows to improve information density and usable space.

### 4. OCR

- Added **Quick OCR** with a streamlined screenshot, recognition, editor-return, and result-insertion workflow.
- Added **Silent OCR**, which can recognize a screenshot in the background and place the resulting LaTeX directly on the clipboard.
- Added an OCR status / HUD for recognition progress and completion feedback.
- Unified model selection, settings, and entry points for standard OCR, Quick OCR, and Silent OCR.
- Improved model switching, window transitions, screenshot triggering, and result delivery stability.

### 5. Office editor startup and performance

- Word / PowerPoint formula editors now use a more mature **persistent / prewarmed lifecycle**, reducing repeated cold-start delays.
- Fixed the main app stealing focus, incorrect window ordering, and icon issues when opening the Office editor.
- The editor can appear directly above Word / PowerPoint, with focus returning correctly to Office after editing.
- Further hardened true cold-start behavior.
- Added Office performance budgets and regression coverage to reduce the risk of formula-count-dependent performance regressions.

### 6. Word Ribbon and Office integration

- Reorganized the Word Ribbon and introduced a consistent VisualTeX icon set and command grouping.
- Main commands now cover inline/display picture formulas, inline/display OMML formulas, Edit Selection, Redraw Selection, Redraw Document, Picture / OMML conversion, cross-references, Update Numbering, numbering format, and document import.
- Office dialogs now more closely match the main app for formula toolbar behavior, parameters, fonts, font size, theme, and input behavior.
- Reduced repeated initialization during consecutive formula editing sessions.

### 7. Office add-in installation, update, and repair

- Added a more complete **DOTM / PPAM version-detection system**.
- Detects missing add-ins, outdated versions, incomplete files, and PowerPoint add-ins that have not been registered.
- Old DOTM / PPAM files can be overwritten directly by VisualTeX without manual removal.
- Office add-in status now distinguishes **Install / Update / Repair**.
- Updates can request Word / PowerPoint to quit and continue after Office exits normally.
- Update choices and related preferences are persisted.

### 8. Word redraw and document import

- Further improved **Redraw Selection / Redraw Document** for large formula sets and complex documents.
- Fixed paragraph layout, picture-formula position, and double-click edit targets changing after redraw.
- Expanded mixed-content import and regression coverage for prose, inline formulas, display formulas, and OMML.
- Uses stricter formula identity / metadata to preserve formula identity and original structure and reduce formula/text displacement after redraw.
- Improved target lookup and recovery in bulk operations instead of relying only on current positions.

### 9. Word equation numbering and cross-references

- Numbering settings are now **persistent** instead of session-only.
- Stabilized sequential, chapter-based, and section-based numbering, including `.` and `-` separator styles.
- Improved synchronization between Ribbon numbering controls and actual document state.
- **Update Numbering** now acts as a global repair command that can replay numbering in the formulas' physical document order.
- Improved numbering updates, sequential insertion, and editing performance for large documents; normal tail insertion prefers a local fast path instead of rescanning the full document.
- Improved ordering when picture and OMML formulas are mixed or formulas are inserted in the middle of a document.
- Repairs legacy numbering structures, orphaned `SEQ` fields, damaged helpers, and duplicate numbering states.
- Strengthened transactional rollback so failed numbering operations do not leave partial structures behind.
- Improved **cross-reference / numbering coordination** so middle insertion, reference insertion, subsequent editing, and numbering updates keep formula count, `SEQ`, and body `REF` fields consistent.

### 10. Help Manual

- Added an in-app **Help Manual** that opens directly inside VisualTeX without launching a browser.
- Covers the editor, hotkeys, Common tools, tiles, `align`, `cases`, matrices, Keypad Mode, OCR, custom symbols, Word, and PowerPoint workflows.
- The Help UI follows the active theme and uses improved typography and contrast for body text, navigation, headings, code blocks, and tables.

### Windows addition

- **PowerPoint format conversion**: added conversion between picture, OLE, and OMML formulas so presentation formulas can be switched according to editability, compatibility, and output requirements.
