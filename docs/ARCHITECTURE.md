# VisualTeX repository architecture / 仓库架构

## 中文

### 结构

一个仓库、一条历史（`main`），Windows 和 Mac 两台机器都在这条历史上工作。仓库里是两个桌面应用、网页版和它们共用的代码：

```text
apps/shared     所有应用原样共用的代码
  visualtex-core          编辑内核（src/kernel 是 VisualTeX 自己的 MathLive TypeScript 分支）、结构规则、光标与选区
  tex-layout              TeX 排版引擎、字体与度量
  visualtex-document      文档编辑器（Markdown 行、公式、表格、引用）
  custom-symbol-designer  自定义符号设计器
  visualtex-app           两个应用共用的前端代码：主窗口、设置、文档模式、OCR、VisualTeX Anywhere、导出、样式表
  visualtex-backend       两个 Tauri 后端都编译的 Rust 模块（离线 OCR 模型、文档库），用 #[path] 引入：Tauri
                          命令必须在应用自己的 crate 里
apps/windows    Windows 应用：编辑器外壳（src）、Tauri 后端（src-tauri）、C# Office 加载项（src-windows：
                VSTO、OLE 服务器、MathType，WPS 通过 COM 使用同一套加载项）、输入法（input-method）、OCR
apps/macos      macOS 应用：编辑器外壳（src）、Tauri 后端（src-tauri）、Office 加载项（office/macos-offline：
                DOTM/PPAM、VBA 与 AppleScript，另有 WPS 加载项）、输入法（input-method）、OCR
apps/web        网页版（visualtex.pauljianliao.com）：首页、教程页、浏览器中的编辑器（共享前端加网页版的
                src/platform）、在线 OCR 的转发 Worker（worker/index.js）；构建到仓库根目录的 dist/
docs            本文件、设计文档（editing-core、document-editor、office-mcp）、equation-format-v2.md、HISTORY.md（移出的材料在哪里）
tools           仓库结构检查、网页版发布（deploy_web.mjs）
```

每个应用各自拥有 `package.json` 与 `package-lock.json`、`Cargo.toml` 与 `Cargo.lock`、Vite/TypeScript/Tauri 配置、
脚本和测试。`apps/shared` 没有自己的 `node_modules`：应用以相对路径引用共享文件，由应用的 Vite 配置从应用自己的
`node_modules` 解析依赖（`apps/shared/visualtex-app/vite/sharedPackages.ts`）。内核通过构建期模块
`virtual:visualtex-kernel-host`（`apps/shared/visualtex-app/vite/visualtexKernelHost.ts`）取得共享的命令表和输入规则。
Node 脚本（测试、生成器）用 `tsx --import ../shared/visualtex-app/testing/node-loader.mjs` 运行，共享代码的依赖和
`@platform/…` 从当前应用解析；两个应用都运行的测试在 `apps/shared/visualtex-app/tests`。浏览器测试用
`testing/browser_test_runtime.mjs`：按系统找到 Chrome/Edge，`withAppBrowser` = 应用构建好的 dist（`vite preview`）
或 Vite 开发服务器加经 CDP 驱动的无头浏览器（端口由系统分配，可并行），`withBrowserTest` = 开发服务器的空白页，
用来在页面里引入应用模块（`testing/sharedSourceUrl.mjs`）。编辑器场景在 `tests/editor-scenarios`：一个运行器，
每个场景一个文件，只适用于一个平台的场景用 `platforms` 标明（`npm run test:editor-scenarios` 全部运行）。

### 共享规则

1. `apps/shared` 不引用 `apps/windows`、`apps/macos` 或 `apps/web`。（`npm run test:repository` 检查）
2. 一个应用不引用另一个应用。（同上）
3. 两个应用都需要、且内容相同的文件放在 `apps/shared`，不在两个应用里各放一份。因平台不同而不同的文件留在各自的应用里。
4. 共享代码里的平台差异是一个模块：每个应用在 `src/platform/<模块>` 提供导出相同的版本，共享代码以
   `@platform/<模块>` 引用（见下表）；内核的差异由 `virtual:visualtex-kernel-host` 提供。两个系统叫法不同的约定
   （快捷键写法 Ctrl+B / ⌘B、主修饰键）用 `runtime/platformConventions.ts` 的 `onPlatform`、`primaryModifierOnly`。
   一个应用没有的部分（Office、系统级输入、截图 OCR、本地 OCR、更新、小键盘窗口）由 `APP_FEATURES` 说明，共享代码
   据此不显示对应的按钮、设置和快捷键。不保留改过的副本。
5. 不提交构建产物、日志、截图或测试文档；有名字的测试素材目录除外。

留在应用里的是两个平台做法不同的部分：入口和它们的 HTML 页面、Office 编辑窗口与文档导入（后端协议不同）、
各自的 Office 集成设置、快捷 OCR 与静默 OCR 窗口、macOS 的 Office 快捷键、供 AI 智能体使用的 Office MCP 后端
（`src-tauri/src/office/automation`）及其设置（两个系统上的加载项连接方式不同），以及 `src/platform`。

| `src/platform` 模块 | 内容 |
|---|---|
| `appInfo` | `APP_PLATFORM`（共享代码按哪个系统的方式工作）、`KEYBOARD_PLATFORM`（快捷键约定）、`APP_LABEL`、版本、`APP_FEATURES`（应用有哪些部分） |
| `styles.css` | 只有这个应用有的组件和窗口的样式，在共享样式表之后加载 |
| `AppOfficeIntegration` | 主窗口在 Office 集成中的部分（macOS：加载项安装、PowerPoint 注册、Office 快捷键、为 AI 智能体绘制公式；Windows：把编辑器偏好、布局和 PowerPoint 字号交给加载项） |
| `OfficeIntegrationSettings`、`OfficeShortcutsTab`、`officeHotkeyOwner`、`wordPreferences` | 设置里的 Office 部分 |
| `quickOcr` | 快捷 OCR 的截图方式与文字、静默 OCR 快捷键 |
| `ocrBackend` | OCR 后端提供什么（安装恢复、存储位置、模型下载、预热命令、可用的在线服务） |
| `OcrDialog` | OCR 窗口（桌面版：共享窗口；网页版：在线服务窗口） |
| `anywhereKeyRules`、`anywhereWindow` | VisualTeX Anywhere 的快捷键规则、提示与快捷键转交 |
| `exportFile` | 写出导出文件 |
| `tutorials`、`helpManual` | 文档模式的教程卡片（文字，各平台的按键与文件夹）、帮助手册 |

网页版（`apps/web`）是第三个提供 `src/platform` 的应用：它没有 Office、系统级输入和本地 OCR（`APP_FEATURES` 全为
false），OCR 通过 `ocrService` 的 `web` 传输走在线服务（`src/ocr/webOcrTransport.ts`），文档存在浏览器中
（`documentLibraryApi` 的浏览器实现），PDF 用浏览器打印（`printDocumentInBrowser`）。共享代码不为网页版打补丁。

### 在两台机器上工作

- 两台机器都提交到 `main`。在一台机器上开始工作前，先合入另一台机器的新提交；提交之后，把提交带到另一台机器。
  在有共同的远程仓库之前，提交以 git bundle 传递：`git bundle create`，在另一台机器上 `git fetch` 这个 bundle 再合并。
- 对 `apps/shared` 的修改只做一次（在哪台机器上都可以），经 git 到达另一台机器，不再复制文件。
- 对一个应用的修改在它自己的目录里完成。若另一个应用也要同样的行为，那是另一个应用目录里的第二个修改，
  或者是把这段代码移入 `apps/shared` 的理由。

### 平台边界

#### macOS

macOS 使用 DOTM、PPAM（VBA、AppleScriptTask）、Office Group Container 和 Tauri 本地编辑窗口；WPS 有自己的加载项。
Word 可插入图片公式或原生公式（OMML）；PowerPoint 使用本地图片/矢量公式流程。不包含 Office.js 和 Windows COM/OLE 代码。

#### Windows

Windows 使用 VSTO Ribbon、Office 事件和 ATL COM LocalServer。专业模式保存真实的 `VisualTeX.Formula.1` OLE 对象，
同时保留 OMML 与跨平台图片模式；WPS 文字与 WPS 演示通过 COM 使用同一套 C# 加载项。不包含 Office.js、DOTM、PPAM 或
AppleScript。

### 跨平台文档

两个应用写入相同的公式元数据（LaTeX、公式 ID、显示模式），图片公式和 OMML 公式在另一平台打开时仍保留这些信息；
格式见 `docs/equation-format-v2.md`。Word 公式在两个平台上应得到相同的结果：Mac 应用的 OMML 写入器是 Windows C#
写入器的 Rust 移植，它的输入 MathML 来自 Windows Office 编辑器所用的同一份共享 LaTeX → MathML 代码
（`apps/shared/visualtex-app/src/export/runtime.ts`）。

### 分支

- `main`：产品：两个桌面应用和网页版。
- `web`：网站的部署分支（visualtex.pauljianliao.com），Cloudflare 构建它（根目录 `npm run build` → `dist/`，
  `wrangler.jsonc`）。`npm run deploy:web`（`tools/deploy_web.mjs`）把它设为 GitHub 上 `main` 的内容：一个以旧 `web`
  和该 `main` 提交为父提交的新提交，快进推送，不改写历史。

### CI

- `.github/workflows/macos.yml` 在 `apps/macos/**` 或 `apps/shared/**` 变化时运行；
- `.github/workflows/windows.yml` 在 `apps/windows/**` 或 `apps/shared/**` 变化时运行；
- `.github/workflows/web.yml` 在 `apps/web/**` 或 `apps/shared/**` 变化时运行（与 Cloudflare 相同的构建、首页与 OCR 转发检查）；
- `.github/workflows/repository.yml` 检查仓库结构。

### 顶层命令

顶层 `package.json` 不安装 React、Tauri 或 Office 依赖，只把命令转发到子项目：

```bash
npm run bootstrap
npm run build:macos
npm run build:windows
npm run build:web
npm run build            # 网页版（Cloudflare 的构建命令：安装 apps/web 的依赖并构建到 dist/）
npm run deploy:web       # 发布网页版（web 分支 = GitHub 上的 main）
npm run test:repository
npm run check            # 仓库结构 + 本机应用的全部检查
npm run check:windows    # 或 check:macos
```

`tools/check_app.mjs` 先构建一次，再逐个运行应用 `package.json` 里的 `test:*`、`verify:*`、`check:*`；不运行的
（需要 Office、安装或发布版、OCR 模型，或会用当前会话的剪贴板）列在该 `package.json` 的 `checks.skip` 并写明原因。
日志和结果在 `apps/<app>/.check-logs`。

`tools/verify_repository_structure.mjs` 会检查：

- 两个应用目录、`apps/shared` 的各个包和关键平台代码是否存在；
- 已退役的根目录应用源码和 Office.js 路径是否重新出现；
- 两个子项目的名称，以及 package、Tauri、Cargo 版本是否一致；
- README 使用的截图是否存在；
- 应用目录与 `apps/shared` 是否为真实目录而非符号链接；
- 上面的共享规则 1 和 2。

---

## English

### Layout

One repository and one history (`main`), used on both the Windows and the Mac machine. It holds two desktop
applications, the web edition and the code they share:

```text
apps/shared     code every application runs unchanged
  visualtex-core          editing kernel (src/kernel is VisualTeX's own TypeScript fork of MathLive), structure rules, caret and selection
  tex-layout              TeX layout engine, fonts and metrics
  visualtex-document      document editor (Markdown lines, formulas, tables, references)
  custom-symbol-designer  custom symbol designer
  visualtex-app           frontend code common to both applications: the main window, settings, document mode, OCR,
                          VisualTeX Anywhere, export, the stylesheets
  visualtex-backend       Rust modules both Tauri backends compile (offline OCR models, the document library), included
                          with #[path] because Tauri commands must live in the application's own crate
apps/windows    Windows application: editor shell (src), Tauri backend (src-tauri), Office add-ins in C# (src-windows:
                VSTO, OLE server, MathType; WPS uses the same add-ins through COM), input method (input-method), OCR
apps/macos      macOS application: editor shell (src), Tauri backend (src-tauri), Office add-ins (office/macos-offline:
                DOTM/PPAM with VBA and AppleScript, and a WPS add-in), input method (input-method), OCR
apps/web        web edition (visualtex.pauljianliao.com): landing page, tutorial page, the editor in a browser (the
                shared frontend with the web's src/platform), the online OCR relay worker (worker/index.js); it
                builds to the repository's dist/
docs            this file, the design documents (editing-core, document-editor, office-mcp), equation-format-v2.md,
                HISTORY.md (where removed material went)
tools           repository checks, publishing the web edition (deploy_web.mjs)
```

Each application has its own `package.json` and `package-lock.json`, `Cargo.toml` and `Cargo.lock`,
Vite/TypeScript/Tauri configuration, scripts and tests. `apps/shared` has no `node_modules`: an application imports
shared files by relative path, and its Vite configuration resolves their packages from the application's own
`node_modules` (`apps/shared/visualtex-app/vite/sharedPackages.ts`). The kernel receives the shared command tables
and input rules through the build-time module `virtual:visualtex-kernel-host`
(`apps/shared/visualtex-app/vite/visualtexKernelHost.ts`). Node scripts (tests, generators) run with
`tsx --import ../shared/visualtex-app/testing/node-loader.mjs`, which resolves the packages and `@platform/…` of
shared code from the application they run in; the tests both applications run are in
`apps/shared/visualtex-app/tests`. Browser tests use `testing/browser_test_runtime.mjs`: Chrome or Edge located per
system; `withAppBrowser` = the application's built dist (`vite preview`) or its Vite dev server plus a headless
browser driven over CDP (ports chosen by the system, so tests can run in parallel); `withBrowserTest` = a blank page
of the dev server for importing application modules in the page (`testing/sharedSourceUrl.mjs`). Editor scenarios
are in `tests/editor-scenarios`: one runner, one file per scenario, `platforms` marks a scenario that applies to one
platform only (`npm run test:editor-scenarios` runs them all).

### Sharing rules

1. `apps/shared` never imports from `apps/windows`, `apps/macos` or `apps/web`. (checked by `npm run test:repository`)
2. An application never imports from the other application. (checked as well)
3. A file both applications need with the same content belongs in `apps/shared`, not in both applications. A file
   that differs because the platforms differ stays in its application.
4. A platform difference inside shared code is a module each application supplies with the same exports in
   `src/platform/<module>`, imported by shared code as `@platform/<module>` (table below); the kernel's differences
   come through `virtual:visualtex-kernel-host`. Conventions the systems name differently (shortcut labels Ctrl+B /
   ⌘B, the primary modifier key) go through `onPlatform` and `primaryModifierOnly` in
   `runtime/platformConventions.ts`. What an application lacks (Office, system-wide input, screen OCR, local OCR,
   updates, the keypad window) its `APP_FEATURES` says, and shared code leaves out those controls, settings and
   shortcuts. Never an edited copy.
5. No build output, logs, screenshots or test documents are committed, except in a named fixtures folder.

What stays in the applications is what the platforms do differently: the entry points and their HTML pages, the
Office editor window and document import (their backends differ), each Office integration's settings, the quick and
silent OCR windows, the Office hotkeys of macOS, the Office MCP backend for AI agents (`src-tauri/src/office/automation`)
with its settings (the add-ins connect to it differently on each system), and `src/platform`.

| `src/platform` module | What it holds |
|---|---|
| `appInfo` | `APP_PLATFORM` (whose way the shared code works), `KEYBOARD_PLATFORM` (shortcut conventions), `APP_LABEL`, version, `APP_FEATURES` (what the application has) |
| `styles.css` | styles of the components and windows only this application has, loaded after the shared stylesheets |
| `AppOfficeIntegration` | the main window's part in the Office integration (macOS: add-in installation, PowerPoint registration, Office hotkeys, formulas drawn for AI agents; Windows: the editor preferences, layout and PowerPoint size handed to the add-ins) |
| `OfficeIntegrationSettings`, `OfficeShortcutsTab`, `officeHotkeyOwner`, `wordPreferences` | the Office parts of Settings |
| `quickOcr` | quick OCR's capture modes and their texts, the silent OCR shortcut |
| `ocrBackend` | what the OCR backend offers (install recovery, storage folder, model downloads, the warm-up command, the online services it reaches) |
| `OcrDialog` | the OCR window (desktop: the shared one; web: the online services' window) |
| `anywhereKeyRules`, `anywhereWindow` | VisualTeX Anywhere's shortcut rules, notes and shortcut hand-over |
| `exportFile` | writing an export file |
| `tutorials`, `helpManual` | document mode's tutorial cards (their text, each platform's keys and folder), the help manual |

The web edition (`apps/web`) is the third application that supplies `src/platform`: it has no Office, system-wide
input or local OCR (`APP_FEATURES` all false); OCR goes to online services through `ocrService`'s `web` transport
(`src/ocr/webOcrTransport.ts`); documents stay in the browser (the browser side of `documentLibraryApi`); PDF goes
through the browser's printing (`printDocumentInBrowser`). Shared code carries no web patches.

### Working on both machines

- Both machines commit to `main`. Before starting work on one machine, bring in the commits made on the other; after
  committing, carry them over. Until there is a shared remote, commits travel as git bundles: `git bundle create`,
  then `git fetch` the bundle on the other machine and merge.
- A change to `apps/shared` is made once, on either machine, and reaches the other machine through git, not by
  copying files.
- A change to one application is made in its own directory. When the other application should behave the same, that
  is a second change in the other application's directory, or a reason to move the code to `apps/shared`.

### Platform boundaries

#### macOS

macOS uses DOTM and PPAM add-ins (VBA, AppleScriptTask), the Office group container and local Tauri editor windows;
WPS has its own add-in. Word gets picture formulas or native equations (OMML); PowerPoint uses the local
picture/vector workflow. The application contains no Office.js and no Windows COM/OLE code.

#### Windows

Windows uses VSTO Ribbons, Office events and an ATL COM LocalServer. Professional mode stores real
`VisualTeX.Formula.1` OLE objects while keeping OMML and the cross-platform picture mode; WPS Writer and WPS
Presentation use the same C# add-ins through COM. The application contains no Office.js, DOTM, PPAM or AppleScript
components.

### Documents across platforms

Both applications write the same formula metadata (LaTeX, formula ID, display mode), so picture and OMML formulas
keep it when opened on the other platform; `docs/equation-format-v2.md` describes the format. Word equations are
meant to come out the same on both platforms: the Mac application's OMML writer is a Rust port of the Windows C#
writer, and its input MathML comes from the shared LaTeX → MathML code the Windows Office editor uses
(`apps/shared/visualtex-app/src/export/runtime.ts`).

### Branches

- `main`: the product: both desktop applications and the web edition.
- `web`: the website's deployment branch (visualtex.pauljianliao.com), which Cloudflare builds (`npm run build` at the
  root -> `dist/`, `wrangler.jsonc`). `npm run deploy:web` (`tools/deploy_web.mjs`) sets it to GitHub's `main`: a new
  commit whose parents are the previous `web` and that `main` commit, pushed as a fast-forward, nothing rewritten.

### CI

- `.github/workflows/macos.yml` runs when `apps/macos/**` or `apps/shared/**` changes;
- `.github/workflows/windows.yml` runs when `apps/windows/**` or `apps/shared/**` changes;
- `.github/workflows/web.yml` runs when `apps/web/**` or `apps/shared/**` changes (Cloudflare's build, the landing and OCR relay checks);
- `.github/workflows/repository.yml` checks the repository structure.

### Root commands

The root `package.json` has no React, Tauri or Office dependencies. It only dispatches commands:

```bash
npm run bootstrap
npm run build:macos
npm run build:windows
npm run build:web
npm run build            # the web edition (Cloudflare's build command: apps/web's packages, then dist/)
npm run deploy:web       # publish the web edition (the web branch = GitHub's main)
npm run test:repository
npm run check            # repository structure + every check of this system's application
npm run check:windows    # or check:macos
```

`tools/check_app.mjs` builds once, then runs each `test:*`, `verify:*` and `check:*` script of the application's
`package.json`; the ones it leaves out (they need Office, an installed or release build, the OCR models, or use the
session's clipboard) are listed under `checks.skip` in that `package.json` with the reason. Logs and results go to
`apps/<app>/.check-logs`.

`tools/verify_repository_structure.mjs` checks that the application directories, the `apps/shared` packages and key
platform files exist; that retired root-level application sources and Office.js paths have not come back; that the
package names and the package, Tauri and Cargo versions agree; that the README screenshots exist; that the
application directories and `apps/shared` are real directories, not symbolic links; and sharing rules 1 and 2.
