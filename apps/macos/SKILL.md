---
name: visualtex-macos-development
description: 在 VisualTeX 仓库中开发、测试和打包 macOS 应用，严格保持平台隔离与原生离线 Office 路线。
---

# VisualTeX macOS 开发 Skill

## 适用范围

本 Skill 只适用于 `apps/macos`。除非用户明确要求，不修改：

- `apps/windows`
- Windows VBA、OLE、NSIS 或 Windows 帮助手册
- 顶层跨平台架构与发布内容
- 远程分支、标签和发行附件

## 接手工作区

开始前必须：

1. 读取 `git status --short --branch`。
2. 保留已有未跟踪文件和未提交修改。
3. 不执行 `reset`、`clean`、`stash`、切分支或远程推送。
4. 先阅读本文件、`README.md`、`LaTeX_Visual_Formula_Editor_Project_Spec.md`，以及与当前功能直接相关的架构文档。
5. 用户提供测试文档、PPT 或截图时，以其中逐条反馈为验收依据，不擅自合并或跳过条目。

## 当前工程结构

- React + TypeScript：主界面、MathLive 编辑器、公式工具栏、自定义字符设计器。
- Tauri + Rust：桌面窗口、文件与系统能力、OCR 和 Office Session 桥接。
- macOS Office：原生离线 `VisualTeX.dotm` / `VisualTeX.ppam`，通过 AppleScriptTask 与本地 Session 通信。
- OCR：本地 PaddleOCR PP-FormulaNet sidecar，不依赖在线识别服务。
- 自定义字符：可编辑矢量图层、LaTeX 命令注册、MathLive/MathJax/SVG/PNG/OMML 兼容链路。

## 当前产品方向

优先级以 `LaTeX_Visual_Formula_Editor_Project_Spec.md` 为准。当前重点是完善自定义字符系统：

1. 可扩展字符库、可靠持久化、公式复制粘贴与磁贴复用。
2. 设计器素材区只保留可独立使用的裸字符和裸符号，不照搬复合公式模板。
3. 组合字符使用视觉无边界画布，并按实际墨迹自动计算运行时边界。
4. 保持数学字形原有斜度。
5. 水平翻转、垂直翻转及围绕自身中心旋转。
6. 高质量空心描边和后续透视效果。
7. 扩展 LaTeX、Unicode 与系统 Cambria Math 可用字符范围。

## 修改原则

- 先稳定复现，再定位根因，再修改。
- 不用只针对单一示例的字符串补丁代替模型层修复。
- MathLive 输入、IME、Office 会话和自定义字符运行时属于高回归风险区域，修改时必须检查相邻链路。
- 新增的用户数据格式必须可迁移、可校验；不能静默截断用户已有数据。
- 自定义字符在主编辑器、源码模式、SVG、PNG、Office 图片、OMML fallback、历史记录和磁贴中必须保持一致。
- 不将对话记录、一次性构建哈希、个人绝对路径或某一轮交接说明写入长期文档。
- 不声称真实 Word、PowerPoint、DMG 或视觉验收已通过，除非实际执行并保留结果。

## 最低验证

每次改动至少执行与改动直接相关的回归；提交前通常执行：

```bash
cd apps/macos
npm run build:desktop
cargo test --manifest-path src-tauri/Cargo.toml --lib
```

自定义字符相关改动执行统一回归：

```bash
npm run test:custom-symbols
```

该命令同时覆盖可扩展持久化与故障恢复、普通公式非干扰、LaTeX/系统字体字形、输出边界、翻转/旋转/倾斜、空心/透视、设计器 UI、MathLive/Office 同步和 macOS CoreText 轮廓提取。

Office 相关改动使用：

```bash
npm run test:macos-offline-office
```

需要完整验收时再运行：

```bash
npm run test:macos-offline-office:full
```

## 打包

用户要求 DMG 时，覆盖当前 macOS Tauri 的既有 bundle 输出，不在仓库其他位置散落新安装包：

```text
apps/macos/src-tauri/target/release/bundle/dmg/
```

构建后报告实际文件名、大小、SHA-256 和执行过的验证；只有用户明确要求时才创建本地提交。
