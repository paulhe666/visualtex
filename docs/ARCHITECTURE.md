# VisualTeX repository architecture / 仓库架构

## 中文

### 目标

`main` 保存两个完整而独立的桌面应用，而不是一套通过大量平台判断运行的共享应用。

```text
apps/macos    = macOS 编辑器 + macOS Tauri + DOTM/PPAM + VBA/AppleScript + macOS OCR
apps/windows  = Windows 编辑器 + Windows Tauri + VSTO/OLE + Windows OCR
```

两边拥有独立的：

- `src/`
- `src-tauri/`
- `package.json` 与 `package-lock.json`
- `Cargo.toml` 与 `Cargo.lock`
- Vite、TypeScript、Tauri 配置
- Office 插件、安装器、脚本和测试
- 编辑器组件、Session 协议实现和平台运行时

仓库顶层不再保留第三套 `src`、`src-tauri`、`office` 或 `src-windows`。顶层只负责说明、CI 和命令调度。

### 平台边界

#### macOS

macOS 使用 DOTM、PPAM、VBA、AppleScriptTask、Office Group Container 和 Tauri 本地编辑窗口。Word 可插入图片或 OMML 公式；PowerPoint 使用本地图片/矢量公式流程。该实现不注册 Windows COM/OLE 类。

#### Windows

Windows 使用 VSTO Ribbon、Office 事件和 ATL COM LocalServer。专业模式保存真实的 `VisualTeX.Formula.1` OLE 对象，同时保留 OMML 与跨平台图片模式。该实现不依赖 macOS DOTM、PPAM 或 AppleScript。

### 文档兼容与源码解耦

源码完全分开不等于文档必须互不识别。两个应用可以各自实现相同的公式元数据字段，以便图片公式或 OMML 在不同平台打开时保留 LaTeX、公式 ID 和显示模式。兼容规则是文件格式约定，不是共享代码包。

### 分支策略

以下两个旧分支仅作为本次导入来源：

- `feat/macos-offline-office-native`
- `dev-windows-native-office-v1.2.0`

导入完成后，新功能分支应从 `main` 创建，并只修改对应目录：

```text
feat/macos-*    → apps/macos/**
feat/windows-*  → apps/windows/**
```

不要再从旧的“仓库根目录就是单个平台应用”的分支继续长期开发，否则后续仍需要重复做目录前缀迁移。

### CI

- `.github/workflows/macos.yml` 只在 macOS 代码变化时运行；
- `.github/workflows/windows.yml` 只在 Windows 代码变化时运行；
- `.github/workflows/repository.yml` 检查顶层结构和平台隔离；
- 一个平台的依赖、构建失败或 Office 工具链不会阻断另一个平台的日常开发任务。

### 顶层命令

顶层 `package.json` 不安装 React、Tauri 或 Office 依赖，只把命令转发到子项目：

```bash
npm run bootstrap
npm run build:macos
npm run build:windows
npm run test:repository
npm run check
```

`tools/verify_repository_structure.mjs` 会检查：

- 两个应用目录和关键平台代码是否存在；
- 仓库根目录是否重新出现旧应用源码；
- 两个子项目名称和版本是否合理；
- README 使用的真实截图是否存在；
- 平台目录是否为真实目录而非符号链接；
- 一个平台的源码是否直接引用另一个平台目录。

---

## English

### Goal

`main` contains two complete, independent desktop applications. It does not contain one shared application controlled by a growing set of platform conditionals.

```text
apps/macos    = macOS editor + macOS Tauri + DOTM/PPAM + VBA/AppleScript + macOS OCR
apps/windows  = Windows editor + Windows Tauri + VSTO/OLE + Windows OCR
```

Each application owns its own:

- `src/`
- `src-tauri/`
- `package.json` and `package-lock.json`
- `Cargo.toml` and `Cargo.lock`
- Vite, TypeScript, and Tauri configuration
- Office add-ins, installers, scripts, and tests
- Editor components, Session implementation, and platform runtime

There is no third application at the repository root. The root only provides documentation, CI, and command dispatching.

### Platform boundaries

#### macOS

macOS uses DOTM, PPAM, VBA, AppleScriptTask, the Office Group Container, and local Tauri editor windows. Word supports picture and OMML formulas; PowerPoint uses the local picture/vector workflow. This application never registers the Windows COM/OLE class.

#### Windows

Windows uses VSTO Ribbons, Office events, and an ATL COM LocalServer. Professional mode stores real `VisualTeX.Formula.1` OLE objects while retaining OMML and cross-platform picture modes. This application does not depend on macOS DOTM, PPAM, or AppleScript components.

### Document compatibility without source coupling

Fully separate source trees do not require incompatible documents. Each application may independently implement the same metadata fields so picture formulas or OMML can preserve LaTeX, formula IDs, and display mode across platforms. That is a file-format agreement, not a shared code package.

### Branch strategy

The original platform branches are import sources for this migration:

- `feat/macos-offline-office-native`
- `dev-windows-native-office-v1.2.0`

After migration, new work should branch from `main` and modify only its platform directory:

```text
feat/macos-*    → apps/macos/**
feat/windows-*  → apps/windows/**
```

Continuing long-term development from the old root-shaped branches would recreate the same prefix migration problem.

### CI

- `.github/workflows/macos.yml` runs only for macOS changes;
- `.github/workflows/windows.yml` runs only for Windows changes;
- `.github/workflows/repository.yml` verifies the repository boundary;
- One platform's dependency or native Office toolchain does not block ordinary development of the other.

### Root commands

The root `package.json` has no React, Tauri, or Office dependencies. It only dispatches commands:

```bash
npm run bootstrap
npm run build:macos
npm run build:windows
npm run test:repository
npm run check
```

`tools/verify_repository_structure.mjs` verifies required platform files, the absence of legacy root application directories, package identity, screenshot presence, real directories instead of symlinks, and accidental cross-platform path references.
