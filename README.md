<div align="center">
  <img src="src-tauri/app-icon.svg" width="128" alt="VisualTeX logo" />
  <h1>VisualTeX</h1>
  <p><strong>可视化 LaTeX 公式编辑器 · Visual LaTeX Formula Editor</strong></p>
  <p>
    <a href="https://github.com/paulhe666/visualtex/releases/tag/v1.0.3">下载 v1.0.3 / Download v1.0.3</a>
    ·
    <a href="#中文">中文</a>
    ·
    <a href="#english">English</a>
  </p>
</div>

---

# 中文

VisualTeX 是一款面向数学、物理、工程和科研写作场景的桌面 LaTeX 公式编辑器。它把结构化可视化编辑、LaTeX 源码、命令候选和本地公式 OCR 放在同一个工作区中，让用户无需安装 TeX Live，也能快速创建、修改、复制和整理数学公式。

当前版本：**1.0.3**

## 下载

| 平台 | 安装包 | 说明 |
| --- | --- | --- |
| macOS Apple Silicon | [VisualTeX_1.0.3_aarch64.dmg](https://github.com/paulhe666/visualtex/releases/download/v1.0.3/VisualTeX_1.0.3_aarch64.dmg) | 适用于 M 系列 Apple 芯片 |
| Windows x64 | [VisualTeX_1.0.3_x64-setup.exe](https://github.com/paulhe666/visualtex/releases/download/v1.0.3/VisualTeX_1.0.3_x64-setup.exe) | NSIS 安装程序，包含 Python 前置检测 |
| Linux x64 | [VisualTeX_1.0.3_amd64.AppImage](https://github.com/paulhe666/visualtex/releases/download/v1.0.3/VisualTeX_1.0.3_amd64.AppImage) | 通用 AppImage |
| Debian / Ubuntu x64 | [VisualTeX_1.0.3_amd64.deb](https://github.com/paulhe666/visualtex/releases/download/v1.0.3/VisualTeX_1.0.3_amd64.deb) | Debian 软件包 |

当前安装包尚未进行商业代码签名，因此 macOS Gatekeeper、Windows SmartScreen 或部分 Linux 桌面环境可能显示安全提醒。请确认安装包来自本仓库的正式 Release 页面。

## 主要功能

### 可视化公式编辑

- 基于 MathLive 的结构化数学公式输入；
- 支持不限行数的多公式编辑；
- 按 Enter 新建下一行，空行可快速删除；
- 工具栏插入分式、根式、积分、求和、连乘、上下标、极限、矩阵、希腊字母、集合符号和常用关系符；
- 在已有选区上应用分式、根号、括号、上下标等结构；
- 公式缩放、撤销和重做；
- 中文内容自动规范为适合数学模式的 `\text{...}` 表达。

### 命令候选与快速输入

- 输入反斜杠后自动显示 LaTeX 命令候选；
- 支持命令前缀、英文别名、中文关键词和模糊检索；
- 使用方向键选择候选，Enter 或 Tab 插入；
- 根据使用频率和最近使用情况进行个性化排序；
- 插入后自动定位到结构占位符，便于连续键盘输入。

### LaTeX 源码双向同步

- 可随时展开 CodeMirror 6 源码编辑区；
- 可视化编辑器与 LaTeX 源码保持双向同步；
- 支持复制纯 LaTeX、行内公式、独立公式和 `equation` 环境；
- 每一行公式可独立整理为显示公式源码；
- 不依赖 TeX Live，不执行本地 PDF 编译。

### 本地公式 OCR

VisualTeX 集成 PaddleOCR PP-FormulaNet，可把公式截图转换为可编辑 LaTeX。

- 支持选择、拖入或粘贴公式图片；
- 可直接在公式输入框内粘贴图片，并把识别结果插入粘贴时的原光标位置；
- 支持 PP-FormulaNet S、M、L 三档模型；
- 自动检测黑底白字、深色背景和透明背景，并统一为适合模型的输入；
- 显示预处理、模型加载和识别进度；
- 支持取消正在运行的 OCR 任务；
- OCR 在本机运行，公式图片不会上传到第三方服务。

OCR 为可选功能。第一次使用时需要安装独立 Python 运行环境并下载模型，因此需要网络连接和额外磁盘空间。编辑器的其他功能不依赖 OCR 环境。

### 文档与工作区

- 本地公式历史记录；
- VisualTeX JSON 文档导入与导出；
- 自动保存当前编辑状态；
- 深色与浅色主题；
- 中文与英文界面；
- 新手教程，可从主菜单随时重新打开；
- 响应式侧边公式工具栏与紧凑桌面布局。

## 安装

### macOS

1. 下载 `.dmg`；
2. 打开镜像并把 VisualTeX 拖入 Applications；
3. 如果 macOS 阻止首次启动，请在 Finder 中右键应用并选择“打开”。

当前 macOS Release 面向 Apple Silicon。Intel Mac 安装包尚未提供。

### Windows

使用 `VisualTeX_1.0.3_x64-setup.exe`。安装程序会在复制文件前检测 OCR 所需的 64 位 Python 3.9–3.13；若环境不兼容，会明确提示，但仍允许继续安装不依赖 OCR 的公式编辑功能。若 SmartScreen 显示未知发布者，请确认文件来自本仓库 Release 后再继续。

### Linux

AppImage：

```bash
chmod +x VisualTeX_1.0.3_amd64.AppImage
./VisualTeX_1.0.3_amd64.AppImage
```

Debian / Ubuntu：

```bash
sudo apt install ./VisualTeX_1.0.3_amd64.deb
```

## OCR 环境说明

本地 OCR 安装器会创建 VisualTeX 专属虚拟环境。建议准备：

- 64 位 Python 3.9–3.13；
- 首次安装和模型下载所需的稳定网络；
- 至少约 2 GB 可用磁盘空间；
- 使用 M、L 模型时预留数 GB 可用内存。

模型大致下载量：

| 模型 | 定位 | 下载量 |
| --- | --- | ---: |
| PP-FormulaNet plus-S | 速度优先 | 约 260 MB |
| PP-FormulaNet plus-M | 均衡，默认推荐 | 约 621 MB |
| PP-FormulaNet plus-L | 精度优先 | 约 732 MB |

## 技术架构

- [Tauri 2](https://tauri.app/)：跨平台桌面容器；
- [React](https://react.dev/) + TypeScript：应用界面与状态管理；
- [MathLive](https://mathlive.io/)：结构化公式编辑；
- [CodeMirror 6](https://codemirror.net/)：LaTeX 源码编辑；
- PaddleOCR PP-FormulaNet：本地公式图片识别；
- Rust：本地运行时、进程管理、文件处理与 OCR 侧车通信。

## 本地开发

需要 Node.js、Rust 和目标平台对应的 Tauri 系统依赖。

```bash
npm install
npm run tauri:dev
```

只运行前端：

```bash
npm run dev
```

构建检查：

```bash
npm run build
cargo test --manifest-path src-tauri/Cargo.toml
```

构建桌面安装包：

```bash
npm run tauri:build
```

Windows 和 Linux 的自动构建说明见 [`docs/WINDOWS_LINUX_RELEASE.md`](docs/WINDOWS_LINUX_RELEASE.md)。

## 核心设计原则

VisualTeX 始终以 LaTeX 字符串作为公式的单一数据源。可视化编辑、工具栏插入、命令候选、OCR 结果和源码编辑最终都作用于同一份 LaTeX 内容，从而避免界面状态与源码状态分离。

---

# English

VisualTeX is a desktop LaTeX formula editor for mathematics, physics, engineering, education, and scientific writing. It combines structured visual editing, editable LaTeX source, command suggestions, and local formula OCR in one workspace. No TeX Live installation is required for editing and copying formulas.

Current version: **1.0.3**

## Downloads

| Platform | Package | Notes |
| --- | --- | --- |
| macOS Apple Silicon | [VisualTeX_1.0.3_aarch64.dmg](https://github.com/paulhe666/visualtex/releases/download/v1.0.3/VisualTeX_1.0.3_aarch64.dmg) | For Apple M-series Macs |
| Windows x64 | [VisualTeX_1.0.3_x64-setup.exe](https://github.com/paulhe666/visualtex/releases/download/v1.0.3/VisualTeX_1.0.3_x64-setup.exe) | NSIS installer with Python prerequisite detection |
| Linux x64 | [VisualTeX_1.0.3_amd64.AppImage](https://github.com/paulhe666/visualtex/releases/download/v1.0.3/VisualTeX_1.0.3_amd64.AppImage) | Portable AppImage |
| Debian / Ubuntu x64 | [VisualTeX_1.0.3_amd64.deb](https://github.com/paulhe666/visualtex/releases/download/v1.0.3/VisualTeX_1.0.3_amd64.deb) | Debian package |

The current packages are not commercially code-signed. macOS Gatekeeper, Windows SmartScreen, or some Linux desktop environments may display a warning. Verify that the package was downloaded from the official Release page of this repository.

## Features

### Visual formula editing

- Structured mathematical input powered by MathLive;
- Unlimited multi-line formula editing;
- Press Enter to create a new formula line;
- Insert fractions, roots, integrals, sums, products, scripts, limits, matrices, Greek letters, sets, relations, and common symbols from the formula toolbar;
- Apply structures such as fractions, roots, brackets, and scripts to an existing selection;
- Formula zoom, undo, and redo;
- Automatic normalization of Chinese text into math-compatible `\text{...}` expressions.

### Command suggestions and fast input

- LaTeX command suggestions appear after typing a backslash;
- Prefix search, English aliases, Chinese keywords, and fuzzy matching;
- Arrow-key navigation with Enter or Tab to insert;
- Personalized ordering based on frequency and recent usage;
- Automatic navigation to structural placeholders after insertion.

### Two-way LaTeX source editing

- Expand an integrated CodeMirror 6 source editor at any time;
- Two-way synchronization between the visual editor and LaTeX source;
- Copy raw LaTeX, inline math, display math, or an `equation` environment;
- Export each formula line as clean display-math source;
- No TeX Live dependency and no local PDF compilation.

### Local formula OCR

VisualTeX integrates PaddleOCR PP-FormulaNet to convert formula images into editable LaTeX.

- Select, drag, or paste a formula image;
- Paste an image directly into a formula field and insert the result at the original caret position;
- Choose between PP-FormulaNet S, M, and L models;
- Automatically handle dark backgrounds, white-on-black formulas, and transparent images;
- Display preprocessing, model loading, and inference progress;
- Cancel a running OCR task;
- Process images locally without uploading them to a third-party service.

OCR is optional. The first use requires installing an isolated Python runtime and downloading model files, so an internet connection and additional disk space are required. All non-OCR editor features work without the OCR runtime.

### Documents and workspace

- Local formula history;
- Import and export VisualTeX JSON documents;
- Automatic restoration of the current editing state;
- Light and dark themes;
- Chinese and English interfaces;
- Built-in onboarding tour, available again from the main menu;
- Responsive formula sidebar and compact desktop workspace.

## Installation

### macOS

1. Download the `.dmg` file;
2. Open it and drag VisualTeX into Applications;
3. If macOS blocks the first launch, right-click the app in Finder and choose **Open**.

The current macOS package targets Apple Silicon. An Intel Mac build is not included yet.

### Windows

Use `VisualTeX_1.0.3_x64-setup.exe`. Before copying files, the installer checks for the 64-bit Python 3.9–3.13 runtime required by OCR. An incompatible environment produces a clear warning while still allowing installation of all non-OCR editor features. If SmartScreen reports an unknown publisher, verify the file against the official Release page before continuing.

### Linux

AppImage:

```bash
chmod +x VisualTeX_1.0.3_amd64.AppImage
./VisualTeX_1.0.3_amd64.AppImage
```

Debian / Ubuntu:

```bash
sudo apt install ./VisualTeX_1.0.3_amd64.deb
```

## OCR runtime requirements

The local OCR installer creates a dedicated VisualTeX virtual environment. Recommended requirements:

- 64-bit Python 3.9–3.13;
- A stable network connection for the first installation and model download;
- At least about 2 GB of free disk space;
- Several GB of available memory when using the M or L model.

Approximate model download sizes:

| Model | Profile | Download |
| --- | --- | ---: |
| PP-FormulaNet plus-S | Speed | ~260 MB |
| PP-FormulaNet plus-M | Balanced, recommended | ~621 MB |
| PP-FormulaNet plus-L | Accuracy | ~732 MB |

## Technology

- [Tauri 2](https://tauri.app/) for the cross-platform desktop shell;
- [React](https://react.dev/) and TypeScript for the application UI;
- [MathLive](https://mathlive.io/) for structured formula editing;
- [CodeMirror 6](https://codemirror.net/) for LaTeX source editing;
- PaddleOCR PP-FormulaNet for local image-to-LaTeX recognition;
- Rust for the local runtime, process lifecycle, file handling, and OCR sidecar communication.

## Development

Node.js, Rust, and the Tauri system dependencies for the target platform are required.

```bash
npm install
npm run tauri:dev
```

Run the frontend only:

```bash
npm run dev
```

Build checks:

```bash
npm run build
cargo test --manifest-path src-tauri/Cargo.toml
```

Build desktop packages:

```bash
npm run tauri:build
```

See [`docs/WINDOWS_LINUX_RELEASE.md`](docs/WINDOWS_LINUX_RELEASE.md) for the Windows and Linux GitHub Actions workflow.

## Core design principle

LaTeX text is the single source of truth in VisualTeX. Visual editing, toolbar insertion, command suggestions, OCR results, and source editing all update the same LaTeX content, preventing the rendered editor and source representation from diverging.
