# VisualTeX macOS 可复现构建与测试说明

> 本文不保存某一次构建的分支、绝对路径、哈希或“已通过”结论。每次交付应根据当前 `HEAD` 重新执行命令，并把实际结果写入被忽略的 `build-logs/` 或 `test-results/`。

## 1. 构建前记录

在 `apps/macos` 下记录：

```bash
git status --short --branch
git rev-parse HEAD
node --version
npm --version
rustc --version
cargo --version
sw_vers
uname -m
```

保留所有用户已有未跟踪文件和未提交修改。不要为构建执行 `reset`、`clean`、`stash` 或切换分支。

## 2. 安装依赖

```bash
cd apps/macos
npm ci
```

OCR 离线运行时只在需要重新打包或验证时准备：

```bash
npm run prepare:ocr-offline
npm run verify:ocr-offline
```

## 3. 基础验证

```bash
npm run build:desktop
cargo test --manifest-path src-tauri/Cargo.toml --lib
```

编辑器输入、Office 或导出相关改动还应选择相应回归：

```bash
npm run test:ime-enter
npm run test:input-behavior
npm run test:latex-format
npm run test:svg-export
npm run test:word-omml
npm run test:office-formula-editor
```

## 4. 自定义字符验证

自定义字符相关修改至少执行统一回归：

```bash
npm run test:custom-symbols
```

该命令包括桌面构建、可扩展字符库持久化、LaTeX/系统字体字形、输出范围适配、翻转/倾斜/旋转、空心/透视、设计器 UI、MathLive/Office 运行时同步、普通公式非干扰，以及 macOS CoreText 字形轮廓提取测试。

按改动范围补充：

```bash
npx tsx scripts/custom_symbol_prototype_export_regression.mts
node scripts/custom_symbol_prototype_regression.mjs
node scripts/custom_symbol_prototype_png_regression.mjs
```

测试必须覆盖持久化、命令冲突、设计器源档恢复、MathLive 运行时刷新、SVG/PNG 输出和普通公式不受影响。

## 5. Office 验证

普通源码与本地回归：

```bash
npm run test:macos-offline-office
```

需要完整宿主验收时：

```bash
npm run test:macos-offline-office:full
```

完整宿主验收不等于只运行脚本。涉及 DOTM、PPAM、Word、PowerPoint、双击编辑、编号或真实页面布局时，还必须按 `VisualTeX_验收清单.md` 做人工检查，并明确区分：

- 自动化通过；
- 真实 Office 宿主通过；
- 人工视觉通过；
- 尚未执行。

## 6. Tauri 与 DMG

发布构建：

```bash
npm run tauri:build
```

DMG 使用 Tauri 现有输出目录：

```text
apps/macos/src-tauri/target/release/bundle/dmg/
```

不要在仓库根目录或其他临时工作区重复散落 DMG。需要覆盖旧测试包时，只覆盖该目录内对应输出。

构建后执行：

```bash
npm run verify:mac-dmg
```

并记录实际产物：

```bash
ls -lh src-tauri/target/release/bundle/dmg/
shasum -a 256 src-tauri/target/release/bundle/dmg/*.dmg
```

## 7. Office 加载项

原生加载项源码与资源位于：

```text
office/macos-offline/
```

加载项构建、注入和验证以 `office/macos-offline/BUILD_ADDINS.md` 为准。不得使用空白 OOXML、仅改扩展名的文件或未经真实 Office 编译的 `vbaProject.bin` 冒充可用 DOTM/PPAM。

## 8. 交付报告格式

每次交付只报告实际发生的内容：

```text
HEAD:
修改范围:
执行的测试:
未执行的测试:
DMG 路径:
DMG SHA-256:
已知限制:
```

一次性日志留在本地构建目录，不把个人路径、旧版本哈希或某轮对话结论继续写回本文。
