<div align="center">
  <img src="src-tauri/app-icon.svg" width="128" alt="VisualTeX logo" />
  <h1>VisualTeX</h1>
  <p><strong>可视化 LaTeX 公式编辑器与 Office 公式插件 · Visual LaTeX Editor and Office Formula Add-in</strong></p>
  <p>
    <a href="https://github.com/paulhe666/visualtex/releases/tag/v1.1.0">下载 v1.1.0 / Download v1.1.0</a>
    · <a href="#中文">中文</a>
    · <a href="#english">English</a>
  </p>
</div>

---

# 中文

VisualTeX 是一款面向数学、物理、工程、教学与科研写作的桌面公式编辑器。它把结构化可视化编辑、LaTeX 源码、本地公式 OCR、文档历史以及 Microsoft Word/PowerPoint 公式编辑放在同一个工作流中。编辑和复制公式不需要安装 TeX Live。

当前版本：**1.1.0**

## 下载

| 平台 | 安装包 | 说明 |
| --- | --- | --- |
| macOS Apple Silicon | [VisualTeX_1.1.0_aarch64.dmg](https://github.com/paulhe666/visualtex/releases/download/v1.1.0/VisualTeX_1.1.0_aarch64.dmg) | 适用于 Apple M 系列芯片，包含离线 OCR 运行包与 macOS Office 集成 |
| Windows x64 | [VisualTeX_1.1.0_x64-setup.exe](https://github.com/paulhe666/visualtex/releases/download/v1.1.0/VisualTeX_1.1.0_x64-setup.exe) | NSIS 安装程序，可一键启用 Word/PowerPoint OLE 集成 |
| Linux x64 | [VisualTeX_1.1.0_amd64.AppImage](https://github.com/paulhe666/visualtex/releases/download/v1.1.0/VisualTeX_1.1.0_amd64.AppImage) | 通用 AppImage |
| Debian / Ubuntu x64 | [VisualTeX_1.1.0_amd64.deb](https://github.com/paulhe666/visualtex/releases/download/v1.1.0/VisualTeX_1.1.0_amd64.deb) | Debian 软件包 |

发布页同时提供 `SHA256SUMS.txt`。当前安装包若未配置商业代码签名，macOS Gatekeeper 或 Windows SmartScreen 可能显示安全提醒，请确认文件来自本仓库正式 Release。

## 1.1.0 重点功能

### Word 与 PowerPoint 公式插件

- 在 Word 和 PowerPoint 中新建、插入、更新和再次编辑 VisualTeX 公式；
- Word 支持行内公式与行间公式，行内公式会根据真实导出基线自动校准，不再上下漂移；
- Word 支持公式编号刷新，并可通过按钮或双击重新打开可视化编辑器；
- PowerPoint 支持按钮编辑和双击编辑，公式对象会保存可恢复的 VisualTeX metadata；
- PowerPoint 连续把短公式改成长公式时保持原有字形大小，以扩展宽度为主，不会反复压缩；
- 插入成功后公式窗口自动关闭，失败时保留窗口并显示明确错误；
- Office 插件与桌面端共享公式渲染、OCR、缓存和编辑 Session。

### macOS Office 集成

- 首次启动 VisualTeX 时，在新手教程之前提供 Office 集成安装选项；
- 自动安装 Word/PowerPoint Manifest、本地 HTTPS 证书和 LaunchAgent 后台服务；
- Word 或 PowerPoint 中通过“开始 → 加载项 → 我的加载项/开发人员加载项 → VisualTeX”启用插件；
- 设置页可修复、卸载集成、重新生成证书、停止伴侣服务或单独关闭开机启动；
- macOS Office Manifest 版本与应用版本统一生成，1.1.0 对应当前 Manifest 版本 `1.1.0.3`。

### Windows OLE Office 集成

- Windows 安装器默认提供“VisualTeX + OLE Office 集成”选项；
- 勾选后自动完成证书、可信 Office Catalog、Word/PowerPoint Manifest、Ribbon 缓存、后台启动和旧 VSTO 禁用；
- 正常安装成功后无需额外手动注册；安装时应先关闭 Word 和 PowerPoint；
- 设置页可修复 OLE、停止当前伴侣服务、关闭开机启动或移除 OLE Manifest。

### 可视化公式编辑

- 基于 MathLive 的结构化公式输入，支持不限行数的多公式文档；
- 工具栏支持分式、根式、积分、求和、连乘、极限、上下标、希腊字母、集合和关系符；
- 支持 1×1 至 10×10 的自定义矩阵及多种定界符；
- 支持对已有选区应用分式、根号、括号和上下标结构；
- 公式缩放范围 20%–160%，简单公式紧凑显示，高公式自动扩展行高；
- 文档级撤销/重做覆盖输入、增删行、结构插入、OCR、源码应用、历史恢复和文件操作，并恢复光标与选区；
- 自动把中文内容规范为数学模式下的 `\text{...}`。

### LaTeX 源码与代码格式

- CodeMirror 6 源码编辑区与可视化公式双向同步；
- 支持纯源码、`$...$`、`\(...\)`、`$$...$$`、`\[...\]`、`equation`、`align`、`gather`、`multline`、`split` 等 16 种格式；
- 对齐环境自动在顶层关系符前加入 `&`，同时保护矩阵内部的 `&` 与 `\\`；
- 切换格式前自动保存尚未同步的源码草稿；
- VisualTeX JSON 文档会保存公式行、代码格式、缩放和编辑状态。

### 本地公式 OCR

VisualTeX 使用 PaddleOCR PP-FormulaNet，把公式截图转换为可编辑 LaTeX。

- 支持选择、拖入或直接粘贴公式图片；
- 可把识别结果插回粘贴时的原光标位置；
- 支持 PP-FormulaNet plus-S、plus-M、plus-L；
- 自动处理黑底白字、深色背景和透明背景；
- 显示预处理、模型加载和推理进度，并支持取消；
- 图片仅在本机处理，不上传到第三方服务；
- macOS 完整包内置离线 Python、PaddleOCR 和默认 M 模型；Windows 需要兼容的 64 位 Python 3.9–3.13。

### 平台化新手教程与运行管理

- macOS 教程说明如何在 Office 中添加 VisualTeX，以及如何卸载集成和关闭开机启动；
- Windows 教程说明 OLE 安装完成后的状态，以及如何停止服务、关闭开机启动和移除 Manifest；
- macOS 使用 LaunchAgent，Windows 使用当前用户 Run 项，两端均默认支持登录时启动，并可在设置中关闭；
- 支持浅色/深色主题、中文/英文界面、本地历史、更新检查和新手教程重复打开。

## 安装与 Office 使用

### macOS

1. 下载并打开 `VisualTeX_1.1.0_aarch64.dmg`；
2. 将 VisualTeX 拖入 Applications；
3. 首次打开 VisualTeX，在新手教程前选择“安装 Office 集成”；
4. 完成后打开 Word 或 PowerPoint，进入“开始 → 加载项 → 我的加载项/开发人员加载项”，选择 VisualTeX；
5. 若重启 Office 后标签页未显示，按相同路径再次选择 VisualTeX。

若 Gatekeeper 阻止启动，请在 Finder 中右键 VisualTeX 并选择“打开”。当前 macOS Release 仅提供 Apple Silicon 构建。

### Windows

1. 关闭 Word 和 PowerPoint；
2. 运行 `VisualTeX_1.1.0_x64-setup.exe`；
3. 保持默认的“VisualTeX + OLE Office 集成”选项；
4. 安装器会自动完成 Office 集成，不需要额外手动注册；
5. 若安装器提示 OLE 配置失败，请打开 VisualTeX 设置并点击“修复 OLE 集成”。

安装器会检测 OCR 所需的 64 位 Python 3.9–3.13。缺少兼容 Python 不影响公式编辑和 Office 插件，但 OCR 将不可用。

### Linux

```bash
chmod +x VisualTeX_1.1.0_amd64.AppImage
./VisualTeX_1.1.0_amd64.AppImage
```

Debian / Ubuntu：

```bash
sudo apt install ./VisualTeX_1.1.0_amd64.deb
```

Linux 版本提供桌面公式编辑和 OCR 工作流，不包含 Microsoft Office 桌面插件。

## 本地开发

需要 Node.js、Rust 和目标平台对应的 Tauri 系统依赖。

```bash
npm ci
npm run tauri:dev
```

常用检查：

```bash
npm run build
npm run test:platform-onboarding
npm run verify:office-manifest
npm run test:office-bridge
npm run test:word-adapter
npm run test:powerpoint-adapter
cargo test --manifest-path src-tauri/Cargo.toml
```

构建安装包：

```bash
npm run tauri:build
```

## 技术架构

- Tauri 2 + Rust：桌面容器、本地 HTTPS、Office 生命周期、文件和进程管理；
- React + TypeScript：桌面界面、Office Bridge 与可视化编辑窗口；
- MathLive：结构化数学输入；
- CodeMirror 6：LaTeX 源码编辑；
- MathJax：SVG 公式导出；
- Microsoft Office.js + macOS AppleScript + Windows OLE：Word/PowerPoint 集成；
- PaddleOCR PP-FormulaNet：本地图片公式识别。

---

# English

VisualTeX is a desktop formula editor for mathematics, physics, engineering, education, and scientific writing. It combines structured visual editing, editable LaTeX source, local formula OCR, document history, and Microsoft Word/PowerPoint formula editing in one workflow. Editing and copying formulas does not require TeX Live.

Current version: **1.1.0**

## Downloads

| Platform | Package | Notes |
| --- | --- | --- |
| macOS Apple Silicon | [VisualTeX_1.1.0_aarch64.dmg](https://github.com/paulhe666/visualtex/releases/download/v1.1.0/VisualTeX_1.1.0_aarch64.dmg) | Includes the offline OCR runtime and macOS Office integration |
| Windows x64 | [VisualTeX_1.1.0_x64-setup.exe](https://github.com/paulhe666/visualtex/releases/download/v1.1.0/VisualTeX_1.1.0_x64-setup.exe) | NSIS installer with one-step Word/PowerPoint OLE setup |
| Linux x64 | [VisualTeX_1.1.0_amd64.AppImage](https://github.com/paulhe666/visualtex/releases/download/v1.1.0/VisualTeX_1.1.0_amd64.AppImage) | Portable AppImage |
| Debian / Ubuntu x64 | [VisualTeX_1.1.0_amd64.deb](https://github.com/paulhe666/visualtex/releases/download/v1.1.0/VisualTeX_1.1.0_amd64.deb) | Debian package |

The release also includes `SHA256SUMS.txt`. Unsigned packages may trigger macOS Gatekeeper or Windows SmartScreen; verify that the files come from the official repository Release page.

## Highlights in 1.1.0

### Word and PowerPoint add-ins

- Create, insert, update, and reopen VisualTeX formulas in Word and PowerPoint;
- Word supports inline and display formulas with native baseline correction for inline images;
- Word equation-number refresh and reopening through the Ribbon button or double-click;
- PowerPoint button and double-click editing with durable VisualTeX metadata;
- Repeatedly extending a PowerPoint formula preserves its visual glyph scale and expands width instead of progressively shrinking;
- The formula editor closes automatically after a successful insertion and remains open with a clear error after failure;
- Desktop and Office workflows share rendering, OCR, cache, and editing sessions.

### macOS Office integration

- The first VisualTeX launch offers Office integration before the onboarding tour;
- Installs Word/PowerPoint manifests, a local HTTPS certificate, and a LaunchAgent companion service;
- Enable VisualTeX in Word or PowerPoint through Home → Add-ins → My Add-ins/Developer Add-ins;
- Settings can repair or uninstall the integration, regenerate the certificate, stop the companion, or disable startup independently;
- Office manifest versions are generated from the product version. VisualTeX 1.1.0 currently uses manifest version `1.1.0.3`.

### Windows OLE Office integration

- The Windows installer offers VisualTeX + OLE Office integration by default;
- Setup automatically configures the certificate, trusted catalog, Word/PowerPoint manifests, Ribbon command cache, startup entry, and disables legacy VSTO buttons;
- No extra manual registration is required after a successful setup; Word and PowerPoint should be closed during installation;
- Settings can repair OLE, stop the current companion, disable startup, or remove the OLE manifest.

### Formula editing and LaTeX source

- Structured MathLive editing with unlimited formula rows, matrices from 1×1 to 10×10, selection-aware structures, and 20%–160% zoom;
- One document-level undo/redo timeline restoring the active row, caret, and selection;
- CodeMirror 6 two-way source editing with 16 raw, inline, display, equation, alignment, gather, multline, and split formats;
- Automatic top-level relation alignment while preserving matrix-internal markers;
- Persistent VisualTeX JSON documents, local history, bilingual UI, themes, and update checks.

### Local formula OCR

- Select, drag, or paste formula images and insert recognized LaTeX at the saved caret;
- PP-FormulaNet plus-S, plus-M, and plus-L support;
- Dark-background and transparent-image preprocessing, progress reporting, and cancellation;
- Images stay on the local device;
- The complete macOS package bundles offline Python, PaddleOCR, and the default M model. Windows uses a compatible 64-bit Python 3.9–3.13 runtime.

## Installation and Office setup

### macOS

1. Open `VisualTeX_1.1.0_aarch64.dmg` and drag VisualTeX into Applications;
2. On first launch, choose **Install Office integration** before the onboarding tour;
3. Open Word or PowerPoint and choose Home → Add-ins → My Add-ins/Developer Add-ins → VisualTeX;
4. Repeat the Add-ins selection after an Office restart if the sideloaded tab is hidden.

The macOS package currently targets Apple Silicon only. If Gatekeeper blocks the first launch, right-click the app in Finder and choose **Open**.

### Windows

1. Close Word and PowerPoint;
2. Run `VisualTeX_1.1.0_x64-setup.exe`;
3. Keep the default **VisualTeX + OLE Office integration** option;
4. Setup completes the Office registration automatically;
5. If setup reports an OLE failure, open VisualTeX Settings and choose **Repair OLE integration**.

A missing compatible Python runtime affects OCR only; formula editing and the Office add-in remain available.

### Linux

```bash
chmod +x VisualTeX_1.1.0_amd64.AppImage
./VisualTeX_1.1.0_amd64.AppImage
```

Debian / Ubuntu:

```bash
sudo apt install ./VisualTeX_1.1.0_amd64.deb
```

The Linux build includes the desktop formula workflow but not the Microsoft Office desktop add-in.

## Development

```bash
npm ci
npm run tauri:dev
```

Useful checks:

```bash
npm run build
npm run test:platform-onboarding
npm run verify:office-manifest
npm run test:office-bridge
npm run test:word-adapter
npm run test:powerpoint-adapter
cargo test --manifest-path src-tauri/Cargo.toml
```

Build installer packages:

```bash
npm run tauri:build
```

## License

VisualTeX is released under the [MIT License](LICENSE).
