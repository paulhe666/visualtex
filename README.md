<div align="center">
  <img src="apps/macos/src-tauri/app-icon.svg" width="160" alt="VisualTeX" />
  <h1>VisualTeX</h1>
  <p><strong>面向科研写作的跨平台公式编辑器</strong></p>
  <p>可视化公式编辑 · LaTeX 源码编辑 · MathType OLE · Word / PowerPoint 加载项 · 公式 OCR</p>
  <p>
    <a href="#中文">中文</a> · <a href="#english">English</a> ·
    <a href="https://visualtex.pauljianliao.com/">官方网站</a> ·
    <a href="https://github.com/paulhe666/visualtex/releases">下载安装</a> ·
    <a href="apps/macos/docs/help/VisualTeX_帮助手册.md">帮助手册</a>
  </p>
  <p>
    <a href="https://github.com/paulhe666/visualtex/stargazers"><img src="https://img.shields.io/github/stars/paulhe666/visualtex?style=for-the-badge&logo=github&label=STARS" alt="GitHub Stars" /></a>
    <a href="https://github.com/paulhe666/visualtex/releases/latest"><img src="https://img.shields.io/github/v/release/paulhe666/visualtex?style=for-the-badge&logo=github&label=RELEASE" alt="Latest Release" /></a>
    <a href="https://github.com/paulhe666/visualtex/releases"><img src="https://img.shields.io/github/downloads/paulhe666/visualtex/total?style=for-the-badge&logo=github&label=DOWNLOADS" alt="Total Downloads" /></a>
    <a href="https://visualtex.pauljianliao.com/"><img src="https://img.shields.io/badge/WEBSITE-visualtex.pauljianliao.com-0A84FF?style=for-the-badge&logo=googlechrome&logoColor=white" alt="VisualTeX Website" /></a>
  </p>
  <p>
    <img src="https://img.shields.io/badge/WORD-SUPPORTED-0099E5?style=for-the-badge&labelColor=555555" alt="Word" />
    <img src="https://img.shields.io/badge/POWERPOINT-SUPPORTED-0099E5?style=for-the-badge&labelColor=555555" alt="PowerPoint" />
    <img src="https://img.shields.io/badge/LaTeX-INPUT-0099E5?style=for-the-badge&labelColor=555555" alt="LaTeX" />
    <img src="https://img.shields.io/badge/LICENSE-MIT-55B800?style=for-the-badge&labelColor=555555" alt="MIT License" />
  </p>
</div>

---

# 中文

## 产品概述

VisualTeX 是适用于 macOS 和 Windows 的桌面公式编辑器，提供可视化公式输入、LaTeX 源码编辑、公式 OCR，以及 Microsoft Word 和 PowerPoint 加载项。公式可插入 Office 文档并保留后续编辑、格式转换、编号和引用所需的元数据。

<div align="center">
  <img src="docs/images/application.png" width="92%" alt="VisualTeX 应用界面" />
</div>

## 支持 VisualTeX 

VisualTeX 的持续开发由用户支持。使用问题、功能建议和 Office 加载项相关交流可加入 QQ 群：`1045801770`。

<table align="center">
  <tr>
    <td align="center"><img src="docs/images/wechat-pay.jpg" height="320" alt="微信收款 / WeChat Pay" /></td>
    <td align="center"><img src="docs/images/alipay.jpg" height="320" alt="支付宝收款 / Alipay" /></td>
    <td align="center"><img src="docs/images/qq-group.png" height="320" alt="VisualTeX QQ group" /></td>
  </tr>
</table>

## 公式编辑器

可视化编辑区与 LaTeX 源码区实时同步。两种编辑方式可在同一份公式内容上切换。

### 可视化编辑

- 支持分式、根式、上下标、积分、求和、极限、重音、定界符、希腊字母、集合、关系、箭头和常用物理符号。
- 支持多行公式、自定义尺寸矩阵，以及 `equation`、`align`、`aligned`、`gather`、`multline`、`split`、`cases` 等环境。
- `align` 和 `aligned` 支持显式设置 `&` 对齐点。
- 每一行可独立设置为行内公式或行间公式。文字与公式混排使用 `$...$`。
- Enter、Tab、结构跳出、选区包裹和命令补全针对结构化公式输入进行处理。
- 支持粗体、数学斜体和正体格式切换。
- 支持自定义快捷键、希腊字母快捷输入、常用工具和自定义公式磁贴。
- 普通模式与小键盘模式分别保存窗口尺寸。
- 撤销、重做和历史记录保存公式内容、活动行、光标和选区状态。

### LaTeX 源码编辑

源码区基于 CodeMirror，提供以下功能：

- LaTeX 语法高亮和命令分类配色。
- 行号、当前行高亮、括号匹配和代码折叠。
- 环境层级缩进、缩进引导线、自动缩进和 Tab / Shift+Tab。
- 多行环境的结构化显示与编辑。
- 纯 LaTeX、`$...$`、`\(...\)`、`\[...\]` 等复制格式。

常用复制和导出功能不依赖 TeX Live。

## Microsoft Word 加载项

加载项在 Word Ribbon 中提供公式插入、编辑、转换、编号、引用、导入和批量重绘功能。

### 公式插入与编辑

- macOS 支持图片公式和 Word 原生 OMML。
- Windows 支持 VisualTeX OLE、Word OMML 和原生 MathType OLE。
- 支持行内公式和行间公式。
- 支持从 Ribbon 编辑所选公式，以及通过双击重新打开公式。
- 编辑已有公式时恢复 LaTeX、字号、显示模式和公式元数据。
- 各平台按支持范围提供图片公式、OMML、VisualTeX OLE 和 MathType OLE 之间的格式转换。

### 原生 MathType OLE（Windows）

Windows 版可将 VisualTeX 中的公式直接插入为 `Equation.DSMT4` 原生 MathType OLE 对象，覆盖以下类型：

- 行内公式。
- 无编号行间公式。
- 带 MathType 原生左编号的行间公式。
- 带 MathType 原生右编号的行间公式。

VisualTeX 对 MathType OLE 的创建、插入、渲染和编辑不依赖 MathType 安装。使用 MathType 自带编辑器打开对象时需要安装 MathType。

相关转换和文档维护功能包括：

- VisualTeX OLE 与 MathType OLE 双向转换。
- Word OMML 与 MathType OLE 双向转换。
- LaTeX 重绘为 VisualTeX OLE、Word OMML 或 MathType OLE。
- VisualTeX OLE、Word OMML 和 MathType OLE 恢复为 LaTeX。
- 对选区或全文执行批量转换。
- 保留 `MTDisplayEquation`、`MTPlaceRef`、`MTEqn`、`MTChap`、`MTSec`、`ZEqnNum` 等 MathType 原生编号与引用结构。
- 同步更新文档中的 VisualTeX 编号、MathType 编号和相关引用。

原生 MathType OLE 功能由 Windows 版提供。macOS 版使用图片公式和 Word OMML。

### 公式编号与引用

- 支持连续编号、按章编号和按节编号。
- 支持点号和短横线等编号分隔符。
- 支持在文档中间插入带编号公式。
- 引用随公式编号更新。
- “更新公式编号”用于刷新编号、引用及旧版文档中的兼容结构。

### 文档导入与批量重绘

- 导入 LaTeX 或 Markdown，普通文本保留为 Word 原生段落。
- 行内公式和行间公式分别生成独立的可编辑对象。
- 扫描选区或全文中的 `$...$`、`\(...\)`、`$$...$$` 和 `\[...\]`。
- 将选区或全文中的公式原位重绘为 VisualTeX 公式或 Word OMML；Windows 版增加 MathType OLE 输出。
- 行内公式参考周围正文的字号，行间公式参考相邻段落的字号。
- 批量处理包含预检、缓存和失败回退。

## Microsoft PowerPoint 加载项

macOS 使用 PPAM 加载项，Windows 使用 VSTO 加载项。两端均支持新建公式、编辑或替换所选公式、删除公式，以及从 PowerPoint 打开 VisualTeX。

Windows 版使用 VisualTeX OLE 对象，并提供图片、OLE 和 OMML 格式转换。macOS 版使用本地矢量和图片公式。

## 公式 OCR

OCR 支持文件选择、拖放、剪贴板粘贴和区域截图。识别结果写入当前编辑位置，多行结果按多行公式处理。

### 本地 OCR

本地 OCR 使用 PaddleOCR PP-FormulaNet，图片在本机完成处理。

- plus-S：较小模型，侧重识别速度。
- plus-M：兼顾中文、复杂公式和处理速度，作为默认模型。
- plus-L：侧重复杂公式的识别精度。

OCR 提供完整窗口、快捷 OCR 和静默 OCR 三种入口。静默 OCR 在后台完成识别并将 LaTeX 写入剪贴板。

### 在线与自建服务

支持 OpenAI 兼容接口、Ollama、Mathpix、PaddleOCR-VL 1.6 和 SimpleTex。选择在线服务后，图片发送至对应接口。API 地址、模型和密钥保存在本机配置中。

## 导出与配置

- 复制 LaTeX 或 PNG。
- 导出 Markdown、SVG 和 PNG。
- 设置透明背景或自定义 PNG 背景色。
- 保存导出目录和常用输出选项。
- 支持中英文界面、内置主题和完整配色设置。
- 保存公式字号、工具栏尺寸、面板状态和编辑模式。
- 支持配置备份与迁移。
- 自定义字符设计器支持组合矢量字形、注册 LaTeX 命令，并将自定义字符用于编辑器、命令补全、SVG、PNG 和 Office 图片公式。

## 平台说明

| 功能 | macOS | Windows |
|---|---|---|
| 桌面应用 | React、TypeScript、Tauri | React、TypeScript、Tauri |
| Word 加载项 | DOTM、VBA、AppleScriptTask、本地 Session | VSTO Ribbon、Office 事件、COM |
| PowerPoint 加载项 | PPAM、本地 Session | VSTO Ribbon、OLE |
| 主要公式格式 | 图片公式、Word OMML | VisualTeX OLE、Word OMML、图片公式、MathType OLE |
| MathType OLE | 不提供原生创建与转换 | 创建、插入、渲染、编辑、转换、编号和引用 |
| 加载项维护 | 应用内安装、更新和修复 DOTM / PPAM | 安装器部署与 Office 位数匹配的原生组件 |

macOS 和 Windows 使用独立的源码、依赖、Office 加载项和构建流程。两端采用一致的公式元数据约定，用于在跨平台文档中保留 LaTeX、公式 ID 和显示模式。

## 下载与运行要求

- [VisualTeX 官方网站](https://visualtex.pauljianliao.com/)
- [GitHub Releases](https://github.com/paulhe666/visualtex/releases)
- [macOS 帮助手册](apps/macos/docs/help/VisualTeX_帮助手册.md)

macOS 版本要求 macOS 11 或更高版本。Office 功能要求本机安装 Microsoft Word 或 PowerPoint。Windows 版的 MathType OLE 创建、插入、渲染和 VisualTeX 编辑功能不依赖 MathType 安装。

## 开发

仓库中的 macOS 与 Windows 实现相互独立：

```text
visualtex/
├── apps/
│   ├── macos/       # macOS 应用、Tauri、DOTM/PPAM、OCR 与测试
│   └── windows/     # Windows 应用、Tauri、VSTO/OLE、OCR 与测试
├── docs/            # 架构文档与 README 图片
├── tools/           # 仓库结构检查
└── README.md
```

常用命令：

```bash
npm run bootstrap
npm run build:macos
npm run build:windows
npm run test:repository
npm run check
```

架构说明见 [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)。平台构建和测试说明见 [apps/macos/README.md](apps/macos/README.md) 与 [apps/windows/README.md](apps/windows/README.md)。

## 开源协议

VisualTeX 使用 [MIT License](LICENSE)。

---

# English

## Product overview

VisualTeX is a desktop formula editor for macOS and Windows. It provides visual formula editing, LaTeX source editing, formula OCR, and Microsoft Word and PowerPoint add-ins. Formulas inserted into Office documents retain the metadata required for later editing, format conversion, numbering, and references.

## Formula editor

The visual editor and LaTeX source editor operate on synchronized formula content.

### Visual editing

- Fractions, roots, scripts, integrals, sums, limits, accents, delimiters, Greek letters, sets, relations, arrows, and common physics symbols.
- Multiline formulas, custom-size matrices, and environments including `equation`, `align`, `aligned`, `gather`, `multline`, `split`, and `cases`.
- Explicit `&` alignment points in `align` and `aligned`.
- Independent inline or display mode for each line. Mixed text and mathematics use `$...$`.
- Formula-aware Enter, Tab, structure exit, selection wrapping, and command completion.
- Bold, math italic, and upright formatting.
- Custom shortcuts, quick Greek-letter input, favorite tools, and custom formula tiles.
- Separate window dimensions for standard and compact keypad modes.
- Undo, redo, and history records containing formula content, active line, cursor, and selection state.

### LaTeX source editing

The CodeMirror-based source editor provides:

- LaTeX syntax highlighting and command-group colors.
- Line numbers, active-line highlighting, bracket matching, and code folding.
- Environment-aware indentation, indentation guides, automatic indentation, and Tab / Shift+Tab.
- Structured display and editing for multiline environments.
- Copy formats for plain LaTeX, `$...$`, `\(...\)`, and `\[...\]`.

TeX Live is not required for standard copy and export operations.

## Microsoft Word add-in

The Word Ribbon add-in provides formula insertion, editing, conversion, numbering, references, document import, and batch redraw.

### Formula insertion and editing

- macOS supports picture formulas and native Word OMML.
- Windows supports VisualTeX OLE, Word OMML, and native MathType OLE.
- Inline and display formulas are supported.
- Formulas can be edited from the Ribbon or reopened by double-clicking.
- Reopening an existing formula restores its LaTeX, font size, display mode, and metadata.
- Format conversion is available between picture formulas, OMML, VisualTeX OLE, and MathType OLE according to platform support.

### Native MathType OLE on Windows

VisualTeX for Windows inserts formula content as native `Equation.DSMT4` MathType OLE objects in the following forms:

- Inline formulas.
- Unnumbered display formulas.
- Display formulas with native MathType numbering on the left.
- Display formulas with native MathType numbering on the right.

Creation, insertion, rendering, and editing through VisualTeX do not require a MathType installation. Opening an object in the MathType editor requires MathType.

Related conversion and document maintenance functions include:

- Two-way conversion between VisualTeX OLE and MathType OLE.
- Two-way conversion between Word OMML and MathType OLE.
- Redraw from LaTeX to VisualTeX OLE, Word OMML, or MathType OLE.
- Restore VisualTeX OLE, Word OMML, and MathType OLE to LaTeX.
- Batch conversion for the current selection or the full document.
- Preservation of native MathType numbering and reference structures, including `MTDisplayEquation`, `MTPlaceRef`, `MTEqn`, `MTChap`, `MTSec`, and `ZEqnNum`.
- Synchronized updates for VisualTeX equation numbers, MathType equation numbers, and their references.

Native MathType OLE support is provided by the Windows version. The macOS version uses picture formulas and Word OMML.

### Equation numbers and references

- Continuous, chapter-based, and section-based numbering.
- Dot and hyphen number separators.
- Insertion of numbered formulas in the middle of a document.
- Reference updates after equation numbers change.
- Number and reference maintenance for current and legacy document structures.

### Document import and batch redraw

- Import LaTeX or Markdown while retaining ordinary text as native Word paragraphs.
- Create separate editable objects for inline and display formulas.
- Scan the selection or full document for `$...$`, `\(...\)`, `$$...$$`, and `\[...\]`.
- Redraw formulas in place as VisualTeX objects or Word OMML, with MathType OLE output on Windows.
- Derive inline formula size from surrounding text and display formula size from adjacent paragraphs.
- Apply preflight checks, caching, and failure fallback during batch processing.

## Microsoft PowerPoint add-in

macOS uses a PPAM add-in. Windows uses a VSTO add-in. Both versions support formula creation, editing or replacement of the selected formula, formula deletion, and opening VisualTeX from PowerPoint.

The Windows version uses VisualTeX OLE objects and provides picture, OLE, and OMML conversion. The macOS version uses local vector and picture formulas.

## Formula OCR

OCR accepts files, drag-and-drop input, clipboard images, and selected screen regions. Recognition results are inserted at the current editing position. Multiline results are handled as multiline formulas.

### Local OCR

Local OCR uses PaddleOCR PP-FormulaNet and processes images on the local computer.

- plus-S: smaller model with emphasis on recognition speed.
- plus-M: balanced support for Chinese text, complex formulas, and processing speed; used as the default model.
- plus-L: emphasis on recognition accuracy for complex formulas.

OCR is available through the full OCR window, Quick OCR, and Silent OCR. Silent OCR processes the image in the background and writes the resulting LaTeX to the clipboard.

### Online and self-hosted services

Supported services include OpenAI-compatible APIs, Ollama, Mathpix, PaddleOCR-VL 1.6, and SimpleTex. Images are sent to the selected endpoint when an online service is used. API URLs, models, and keys are stored in local configuration.

## Export and configuration

- Copy LaTeX or PNG.
- Export Markdown, SVG, and PNG.
- Configure transparent or custom PNG backgrounds.
- Store export directories and common output options.
- Chinese and English interfaces, built-in themes, and complete color configuration.
- Persistent formula size, toolbar dimensions, panel state, and editing mode.
- Configuration backup and migration.
- A custom character designer for composing vector glyphs, registering LaTeX commands, and using custom characters in editing, completion, SVG, PNG, and Office picture formulas.

## Platform overview

| Feature | macOS | Windows |
|---|---|---|
| Desktop application | React, TypeScript, Tauri | React, TypeScript, Tauri |
| Word add-in | DOTM, VBA, AppleScriptTask, local session | VSTO Ribbon, Office events, COM |
| PowerPoint add-in | PPAM, local session | VSTO Ribbon, OLE |
| Main formula formats | Picture formulas, Word OMML | VisualTeX OLE, Word OMML, picture formulas, MathType OLE |
| MathType OLE | Native creation and conversion unavailable | Creation, insertion, rendering, editing, conversion, numbering, and references |
| Add-in maintenance | In-app installation, update, and repair for DOTM / PPAM | Installer-deployed native components matched to Office bitness |

The macOS and Windows versions use separate source trees, dependencies, Office add-ins, and build processes. Both versions use a common formula metadata convention to retain LaTeX, formula IDs, and display mode in cross-platform documents.

## Download and requirements

- [VisualTeX website](https://visualtex.pauljianliao.com/)
- [GitHub Releases](https://github.com/paulhe666/visualtex/releases)
- [macOS user guide](apps/macos/docs/help/VisualTeX_帮助手册.md)

The macOS version requires macOS 11 or later. Office functions require Microsoft Word or PowerPoint. MathType OLE creation, insertion, rendering, and VisualTeX editing on Windows do not require a MathType installation.

## Development

The macOS and Windows implementations are maintained separately:

```text
visualtex/
├── apps/
│   ├── macos/       # macOS application, Tauri, DOTM/PPAM, OCR, and tests
│   └── windows/     # Windows application, Tauri, VSTO/OLE, OCR, and tests
├── docs/            # Architecture documents and README images
├── tools/           # Repository structure checks
└── README.md
```

Common commands:

```bash
npm run bootstrap
npm run build:macos
npm run build:windows
npm run test:repository
npm run check
```

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the repository architecture. Platform build and test instructions are available in [apps/macos/README.md](apps/macos/README.md) and [apps/windows/README.md](apps/windows/README.md).

## License

VisualTeX is distributed under the [MIT License](LICENSE).


