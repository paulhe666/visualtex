<div align="center">
  <h1>VisualTeX</h1>
  <p><strong>为科研写作打造的可视化公式编辑器与 Microsoft Office 原生公式工具</strong></p>
  <p>流畅结构化输入 · LaTeX 双向编辑 · MathType 兼容 · Word / PowerPoint 深度集成</p>
  <p>
    <a href="https://visualtex.pauljianliao.com/">官方网站</a> ·
    <a href="https://github.com/paulhe666/visualtex/releases">下载与更新</a> ·
    <a href="#english">English</a>
  </p>
</div>

---

VisualTeX 面向数学、物理、工程和科研写作，把可视化公式输入、LaTeX 源码、图片识别与 Office 排版放在同一套工作流里。你可以像使用专业公式编辑器一样直接搭建复杂结构，也可以随时切换到源码精确调整；完成后的公式能够进入 Word 或 PowerPoint，继续编辑、转换、编号和引用，而不是停留在一次性的截图。

<div align="center">
  <img src="docs/images/application.png" width="92%" alt="VisualTeX 应用界面" />
</div>

## 为什么选择 VisualTeX

- **编辑手感优先**：基于 MathLive 的结构化编辑与 CodeMirror 源码区实时同步，兼顾鼠标操作、键盘连续输入和精确源码控制。
- **Office 不是导出终点**：Word 与 PowerPoint 提供原生插件入口，公式可以新建、二次编辑、批量重绘、格式转换和继续管理。
- **面向现有论文资产**：Windows 端重点适配 MathType 原生 OLE，可在 VisualTeX、MathType 与 Word OMML 之间迁移公式，降低旧文档改造成本。
- **从图片到可编辑公式**：既支持本地离线 PP-FormulaNet，也支持多种可选 OCR API；识别结果直接回到编辑器继续校正。
- **围绕长文档设计**：公式编号、交叉引用、章节编号、批量导入和全文重绘都以论文、报告和教材中的真实文档为目标。

## 公式编辑：可视化与源码并重

VisualTeX 的主编辑区适合连续录入，而不是只提供一个公式预览框。分式、根式、积分、求和、极限、上下标、重音、定界符、集合、关系、箭头、物理符号等常用结构可以直接插入；矩阵支持自定义尺寸，多行公式可以按行管理。

对于复杂推导，VisualTeX 对以下结构提供了专门的编辑行为：

- `align` / `aligned` 可显式设置 `&` 对齐点；
- `cases` 支持稳定的多行分段函数编辑；
- 支持 `equation`、`gather`、`multline`、`split` 及常用矩阵环境；
- 每一行可独立设置行内或行间模式，文字与公式混排可使用 `$...$`；
- Enter、Tab、选区包裹、结构跳出和命令候选针对数学输入做了适配。

LaTeX 源码区采用 CodeMirror，提供行号、分层语法高亮、环境缩进、缩进引导线、当前行提示、代码折叠、括号匹配、Tab / Shift+Tab 与自动缩进。可视化内容和源码保持双向同步，既能快速输入，也能直接处理较长的 `align`、矩阵或多行公式。

编辑器还包含：

- 可自定义的公式工具栏、常用工具和公式磁贴；
- 公式快捷键、冲突检查与 `⌘G` / 对应 Windows 快捷方式的希腊字母输入；
- 粗体、数学斜体、正体等选区格式切换；
- 多公式历史记录、撤销/重做和 JSON 文档保存；
- 普通模式与小键盘模式，以及分别记忆的窗口布局；
- 中英文界面、多套主题、完整配色自定义、公式字号与界面密度设置；
- 自定义字符设计器，可组合矢量字形、注册 LaTeX 命令，并进入编辑、自动补全与导出链路。

## MathType 兼容：让旧公式继续可用

VisualTeX 的 Windows Office 插件针对 MathType 7 原生公式进行了系统适配。它不是把 MathType 公式简单截图，而是识别并处理 Word 中的 `Equation.DSMT4` / MathType OLE 对象及其原生编号结构。

支持的主要工作流包括：

- 在 VisualTeX 编辑器中打开并修改所选 MathType 公式；
- VisualTeX OLE 与 MathType OLE 双向转换；
- Word OMML 与 MathType OLE 双向转换；
- 将选区或全文中的 LaTeX 重绘为 MathType、VisualTeX OLE 或 Word OMML；
- 将选区或全文中的 MathType、VisualTeX OLE、OMML 还原为 LaTeX；
- 批量转换时保留行内/行间语义、公式位置、字号、编号与引用关系；
- 兼容 MathType 左/右编号，以及 `MTDisplayEquation`、`MTPlaceRef`、`MTEqn`、`MTChap`、`MTSec` 和 `ZEqnNum` 等原生 Word 结构；
- 更新 VisualTeX 与 MathType 公式编号及对应引用，而不把两套体系粗暴混为一种格式。

需要 MathType 原生编辑器或渲染器的操作仍以本机正确安装 MathType 为前提。macOS 版本采用自身的图片公式与 OMML 工作流；直接的 MathType OLE 转换能力以 Windows 版本为主。

## Word：从单个公式到整篇论文

Word 插件直接出现在 Ribbon 中，覆盖公式从插入到维护的完整过程。

### 插入与编辑

- 插入行内或行间公式；
- macOS 可选择图片公式或 Word 原生 OMML；
- Windows 可选择 VisualTeX OLE、Word OMML，并处理 MathType OLE；
- 支持按钮编辑与原生双击唤起；
- 二次编辑时恢复 LaTeX、显示模式、字号和公式元数据；
- 图片、OMML、VisualTeX OLE 与 MathType 的可用转换按平台提供。

### 编号与交叉引用

VisualTeX 支持全文连续编号、按章编号和按节编号，并提供 `.` / `-` 等常用分隔形式。带编号公式可在文档中间插入，引用通过 Word 原生字段与书签保持联动；“更新公式编号”可以按文档物理顺序刷新和修复编号、引用及兼容结构。

### 批量导入与重绘

- 导入 LaTeX 或 Markdown，把普通文字保留为 Word 原生段落；
- 将其中每个行内、行间公式生成为独立可编辑对象；
- 扫描选区或全文中的 `$...$`、`\(...\)`、`$$...$$` 与 `\[...\]`；
- 原位重绘为 VisualTeX 公式、Word OMML，Windows 端还可选择 MathType；
- 公式字号按周围正文继承，尽量保持段落布局和公式位置；
- 批量处理使用预检、缓存、稳定目标定位与事务式替换，失败时优先保留原内容。

## PowerPoint：公式保持可编辑

VisualTeX 为 PowerPoint 提供原生加载项入口，可新建公式、编辑或替换所选公式、删除公式，并从插件直接打开主应用。macOS 使用 PPAM 与本地矢量/图片公式流程；Windows 使用原生 VSTO 与 OLE，并提供图片、OLE、OMML 等格式之间的转换能力，便于在可编辑性、兼容性和演示效果之间切换。

## OCR：本地优先，也可连接现有服务

公式图片可以通过文件选择、拖放、剪贴板粘贴或区域截图送入 OCR。识别结果会回填到当前光标位置，多行识别结果可以继续作为多行公式编辑。

本地模式使用 PaddleOCR PP-FormulaNet：

- **plus-S**：体积较小、速度优先；
- **plus-M**：兼顾中文、复杂公式与速度，作为默认推荐；
- **plus-L**：精度优先，适合更复杂的图片。

除完整 OCR 窗口外，还提供快捷 OCR 和静默 OCR：前者缩短“截图—识别—回填”的路径，后者可在后台识别并把 LaTeX 直接写入剪贴板。深色背景、透明图片、进度提示和取消操作也包含在统一流程中。

如果希望复用已有服务，还可选择 OpenAI 兼容接口、Ollama、Mathpix、PaddleOCR-VL 1.6 或 SimpleTex（标准/极速）。密钥与提供器设置由本机配置管理；只有选择在线提供器时，图片才会发送到相应服务。

## 导出、复制与个性化

VisualTeX 不依赖 TeX Live 即可完成常用公式输出：

- 复制 LaTeX 或带透明背景的 PNG；
- 导出 Markdown、SVG 和 PNG；
- 记忆导出目录、PNG 背景色与常用输出选项；
- 支持公式颜色、字号、布局、工具栏尺寸、面板状态和主题持久化；
- 配置可备份和迁移，方便在新设备恢复快捷键、磁贴、外观与编辑习惯。

## 两个平台，各自采用原生实现

| | macOS | Windows |
|---|---|---|
| 桌面应用 | React + TypeScript + Tauri | React + TypeScript + Tauri |
| Word | 原生 DOTM、VBA、AppleScriptTask 与本地 Session | 原生 VSTO Ribbon、Office 事件与 COM |
| PowerPoint | 原生 PPAM 与本地 Session | 原生 VSTO Ribbon 与 OLE |
| 公式对象 | 图片公式、Word OMML | VisualTeX OLE、Word OMML、图片公式、MathType OLE |
| MathType | 与现有文档共存；直接转换不是主要路线 | 原生 OLE 编辑、插入、批量转换、编号与引用适配 |
| 安装维护 | 应用内检测并安装/更新/修复 DOTM、PPAM | 安装器部署与 Office 位数匹配的原生组件 |

两个平台拥有独立源码、依赖、Office 插件和构建流程，避免用大量平台判断维持一套脆弱实现；同时通过一致的公式元数据约定，让图片公式和 OMML 在跨平台文档中尽可能保留 LaTeX、公式 ID 与显示模式。

## 下载与使用

- [VisualTeX 官方网站](https://visualtex.pauljianliao.com/)
- [GitHub Releases](https://github.com/paulhe666/visualtex/releases)
- [macOS 帮助手册](apps/macos/docs/help/VisualTeX_帮助手册.md)

macOS 构建要求 macOS 11 或更高版本。Office 功能需要本机安装受支持的 Microsoft Word / PowerPoint；Windows 的 MathType 原生功能需要相应的 MathType 安装环境。应用、可选 OCR 模型和 Office 插件的具体版本请以下载页说明为准。

## 开发

仓库将 macOS 与 Windows 实现完全隔离：

```text
visualtex/
├── apps/
│   ├── macos/       # macOS 应用、Tauri、DOTM/PPAM、OCR 与测试
│   └── windows/     # Windows 应用、Tauri、VSTO/OLE、OCR 与测试
├── docs/            # 架构与 README 资源
├── tools/           # 仓库结构检查
└── README.md
```

常用顶层命令：

```bash
npm run bootstrap
npm run build:macos
npm run build:windows
npm run test:repository
npm run check
```

平台原生打包与 Office 验收请进入对应应用目录执行。架构说明见 [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)，各平台开发入口见 [apps/macos/README.md](apps/macos/README.md) 与 [apps/windows/README.md](apps/windows/README.md)。

## License

VisualTeX 以 [MIT License](LICENSE) 发布。

---

# English

VisualTeX is a desktop formula editor for mathematics, physics, engineering, and scientific writing. It combines fast structured input with two-way LaTeX source editing, formula OCR, and native Microsoft Word / PowerPoint workflows.

Its main strengths are:

- MathLive visual editing paired with an IDE-like CodeMirror LaTeX editor;
- dedicated editing for multiline structures such as `align`, `aligned`, `cases`, matrices, and mixed inline/display content;
- customizable toolbars, hotkeys, formula tiles, themes, history, and a compact keypad mode;
- Word insertion, re-editing, numbering, cross-references, document import, and selection/document-wide redraw;
- native PowerPoint add-ins for creating and re-editing presentation formulas;
- local PP-FormulaNet OCR plus optional OpenAI-compatible, Ollama, Mathpix, PaddleOCR-VL, and SimpleTex providers;
- SVG, PNG, Markdown, and LaTeX output without requiring a TeX Live installation;
- Windows-native MathType 7 OLE workflows, including VisualTeX / MathType / OMML conversion, LaTeX restoration, batch processing, and numbering/reference preservation.

The macOS implementation uses native offline DOTM / PPAM add-ins with VBA, AppleScriptTask, and local Tauri sessions. The Windows implementation uses native VSTO Ribbons, COM/OLE, and real `VisualTeX.Formula.1` objects. See [Releases](https://github.com/paulhe666/visualtex/releases) for installers and [the project architecture](docs/ARCHITECTURE.md) for technical details.

## 支持 VisualTeX

如果 VisualTeX 改善了你的论文、报告或课件公式工作流，欢迎支持项目继续开发。使用问题、功能建议和 Office 插件交流可加入 QQ 群：`1045801770`。

<div align="center">
  <img src="docs/images/wechat-pay.jpg" width="28%" alt="微信收款" />
  <img src="docs/images/alipay.jpg" width="28%" alt="支付宝收款" />
  <img src="docs/images/qq-group.png" width="28%" alt="VisualTeX QQ 交流群" />
</div>
