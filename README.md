<div align="center">
  <img src="apps/macos/src-tauri/app-icon.svg" width="112" alt="VisualTeX" />
  <h1>VisualTeX</h1>
  <p><strong>可视化公式编辑器与 Microsoft Office 原生公式插件</strong></p>
  <p><strong>Visual formula editor with native Microsoft Office integration</strong></p>
  <p>
    <a href="#中文">中文</a> · <a href="#english">English</a> ·
    <a href="https://github.com/paulhe666/visualtex/releases">Releases</a>
  </p>
</div>

---

# 中文

VisualTeX 是一款面向数学、物理、工程与科研写作的桌面公式编辑器。它提供结构化可视化输入、LaTeX 源码编辑、本地图片公式识别，以及 Word 和 PowerPoint 原生公式工作流。

当前仓库采用**平台完全分离**的结构：macOS 与 Windows 拥有各自独立的编辑器、Tauri 后端、Office 插件、构建脚本和依赖锁文件。修改一个平台不会编译或影响另一个平台。

## 实际界面

下图由当前仓库中的两个平台应用分别完成生产构建后，通过真实页面渲染截取，不是设计稿或模拟图。

### macOS 编辑器

![VisualTeX macOS 实际运行界面](docs/images/visualtex-macos-editor.png)

### Windows 编辑器

![VisualTeX Windows 实际运行界面](docs/images/visualtex-windows-editor.png)

## 主要功能

### 可视化公式编辑

- 结构化输入分式、根式、积分、求和、极限、上下标、希腊字母、集合与关系符号；
- 支持多公式行、1×1 至 10×10 矩阵、定界符和基于选区的结构插入；
- 文档级撤销与重做会恢复公式内容、活动行、光标和选区；
- 支持公式缩放、浅色/深色主题、中文/英文界面、本地历史和 JSON 文档。

### LaTeX 源码

- CodeMirror 源码编辑区与可视化公式双向同步；
- 支持纯 LaTeX、`$...$`、`\(...\)`、`\[...\]`、`equation`、`align`、`gather`、`multline`、`split` 等格式；
- 多行环境自动处理顶层对齐符，同时保护矩阵内部的 `&` 与换行结构；
- 复制公式不要求安装 TeX Live。

### 本地图片公式识别

- 选择、拖入或直接粘贴公式图片，识别结果可插回原光标位置；
- 使用 PaddleOCR PP-FormulaNet plus-S、plus-M 和 plus-L；
- 支持深色背景、透明图片、进度显示和取消；
- 图片只在本机处理，不上传第三方服务。

## macOS 版本

macOS 版本位于 [`apps/macos`](apps/macos)，使用完全离线的原生 Office 集成：

- Word 通过 DOTM 全局模板加载 VisualTeX 标签页；
- PowerPoint 通过固定路径 PPAM 加载项工作；
- VBA、AppleScriptTask、Office Group Container 与 Tauri 本地窗口组成离线 Session 流程；
- Word 支持图片公式和原生 OMML 行内/行间公式；
- 支持图片公式转换为 Word 原生公式、公式编号、交叉引用、按钮编辑和双击编辑；
- PowerPoint 支持新建、替换、删除和双击编辑 VisualTeX 公式；
- 不依赖 Office.js、本地 HTTPS、证书或网络请求。

## Windows 版本

Windows 版本位于 [`apps/windows`](apps/windows)，使用 VSTO 与真正的 COM/OLE 公式对象：

- Word 和 PowerPoint 使用原生 VSTO Ribbon 与 Office 事件；
- 专业模式插入真实的 `VisualTeX.Formula.1` OLE 对象；
- OLE 对象保存公式元数据、EMF 矢量预览与 PNG 兼容预览；
- Word 支持 OLE 与 OMML 行内/行间公式、格式转换、编号和引用；
- PowerPoint 支持 OLE 公式的新建、编辑、删除与图片导出；
- Office 原生双击可重新打开 VisualTeX 编辑器；
- 兼容图片模式用于跨平台文档和旧公式迁移。

## 仓库结构

```text
visualtex/
├── apps/
│   ├── macos/       # 独立 macOS 应用、Tauri、DOTM/PPAM、OCR 与测试
│   └── windows/     # 独立 Windows 应用、Tauri、VSTO/OLE、OCR 与测试
├── docs/            # 仓库架构与真实界面截图
├── tools/           # 顶层结构检查
├── .github/         # 按平台独立运行的 CI
├── package.json     # 只负责调度两个子项目
└── README.md
```

两个子项目不共享 `src`、`src-tauri`、编辑器组件、Office Session 实现、`package-lock.json` 或 `Cargo.lock`。

## 本地开发

```bash
# 安装两个平台各自的依赖
npm run bootstrap

# 分别构建
npm run build:macos
npm run build:windows

# 检查仓库隔离结构并构建两个前端
npm run check
```

平台原生打包和 Office 验收请进入对应目录执行：

```bash
cd apps/macos
npm run tauri:build

cd ../windows
npm run tauri:build
```

详细设计见 [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)。

---

# English

VisualTeX is a desktop formula editor for mathematics, physics, engineering, and scientific writing. It combines structured visual input, editable LaTeX source, local formula-image recognition, and native Word and PowerPoint workflows.

This repository uses a **fully platform-separated architecture**. macOS and Windows have independent editors, Tauri backends, Office add-ins, build scripts, and lockfiles. A change to one platform does not compile or modify the other.

## Real application renders

The following images were captured from production builds of the two applications in this repository. They are real rendered interfaces, not mockups.

### macOS editor

![VisualTeX macOS real application render](docs/images/visualtex-macos-editor.png)

### Windows editor

![VisualTeX Windows real application render](docs/images/visualtex-windows-editor.png)

## Core features

### Visual formula editing

- Structured fractions, roots, integrals, sums, limits, scripts, Greek letters, sets, and relations;
- Multiple formula rows, 1×1 to 10×10 matrices, delimiters, and selection-aware structure insertion;
- Document-level undo and redo with active-row, caret, and selection restoration;
- Formula zoom, light and dark themes, Chinese and English UI, local history, and JSON documents.

### LaTeX source

- Two-way synchronization between the visual editor and CodeMirror source editor;
- Raw LaTeX plus `$...$`, `\(...\)`, `\[...\]`, `equation`, `align`, `gather`, `multline`, `split`, and related formats;
- Top-level alignment handling that preserves matrix-internal separators and line breaks;
- Formula copying without a TeX Live installation.

### Local formula OCR

- Select, drag, or paste formula images and insert recognized LaTeX at the saved caret;
- PaddleOCR PP-FormulaNet plus-S, plus-M, and plus-L models;
- Dark-background and transparent-image preprocessing, progress reporting, and cancellation;
- Images remain on the local device.

## macOS application

The macOS application lives in [`apps/macos`](apps/macos) and uses a fully offline native Office route:

- A Word DOTM global template and a fixed-path PowerPoint PPAM add-in;
- VBA, AppleScriptTask, the Office Group Container, and local Tauri Session windows;
- Word picture formulas and native OMML inline or display formulas;
- Picture-to-OMML conversion, equation numbering, cross-references, Ribbon editing, and double-click editing;
- PowerPoint formula creation, replacement, deletion, and double-click editing;
- No Office.js, local HTTPS certificate, or network dependency in the native route.

## Windows application

The Windows application lives in [`apps/windows`](apps/windows) and uses VSTO with real COM/OLE formula objects:

- Native VSTO Ribbons and Office events for Word and PowerPoint;
- Real `VisualTeX.Formula.1` OLE objects in professional mode;
- Embedded formula metadata, EMF vector previews, and PNG compatibility previews;
- Word OLE and OMML inline/display formulas, conversion, numbering, and references;
- PowerPoint OLE creation, editing, deletion, and picture export;
- Native Office double-click activation;
- A picture mode for cross-platform documents and legacy migration.

## Repository layout

```text
visualtex/
├── apps/
│   ├── macos/       # Independent macOS editor, Tauri app, DOTM/PPAM, OCR, tests
│   └── windows/     # Independent Windows editor, Tauri app, VSTO/OLE, OCR, tests
├── docs/            # Repository architecture and real screenshots
├── tools/           # Top-level structure verification
├── .github/         # Platform-specific CI workflows
├── package.json     # Dispatches commands to the two applications only
└── README.md
```

The two applications do not share `src`, `src-tauri`, editor components, Office Session implementations, `package-lock.json`, or `Cargo.lock`.

## Development

```bash
npm run bootstrap
npm run build:macos
npm run build:windows
npm run check
```

Run native packaging and Office acceptance from the corresponding application directory:

```bash
cd apps/macos
npm run tauri:build

cd ../windows
npm run tauri:build
```

See [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for the repository design.
