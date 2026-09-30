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

Word 中文版式图片对齐专项：

```bash
npx tsx scripts/word_chinese_image_alignment_regression.mts
```

脚本用生产 SVG 字体固化代码在 Office Scratch 目录生成 45 张图片、10 页的 `word-chinese-image-alignment-fixture.docx`。安装当前重新编译的加载项与客户端后执行：

```bash
npx tsx scripts/word_chinese_image_alignment_regression.mts --run
npx tsx scripts/word_chinese_image_alignment_regression.mts --verify
npx tsx scripts/word_chinese_image_alignment_regression.mts --live
node scripts/word_image_baseline_visual_probe.mjs --document=word-chinese-image-alignment-fixture.docx --no-snapshot
```

专项覆盖顶部、居中、基线、底端、自动五种版式，以及用户参考图中的分式、积分、求和、上下标、行内 display style 和有无编号的行间图片。检查版式保留、字号往返、正文中的光标触发整段修复、多段选择及保存后的尺寸与位置。段落版式始终保留；按各自的 Word 锚点、正文实际字体边界和公式的数学轴/主体字母/墨迹中心计算图片位置。顶部短图片需要负位移时，仅扩展透明顶部画布并用正位移补偿；切换版式时恢复原始字形尺寸，避免 Word 撤销裁剪时累积取整误差。

另需对 Word 原生截图与 Word 导出的 144 dpi 页面逐像素验收，不能用宏或 OOXML 检查代替视觉检查。导出 PDF 后执行 `pdftoppm -r 144 -png PDF_PATH RENDER_DIRECTORY/after-page`，再运行 `python3 scripts/word_chinese_image_alignment_pixels.py RENDER_DIRECTORY`；原生窗口执行同一脚本加 `--native`。按正文墨迹高度归一化到参考图，两种渲染器的阈值均为 3 参考像素，并分别记录实测最大误差。Word 整磅位移、屏幕字体留白和栅格取整存在差异：顶部求和公式无法用同一个整磅位置使窗口与 PDF 同时小于 2.5 参考像素，当前选择两者最大误差较小的位置。主体字母基线差不超过 1 个实际像素，行间公式与编号墨迹中心差不超过 1.5 个实际像素。

`--live` 从全新 Word 进程打开文档，不直接调用初始化或定位宏，逐一改变整份文档的五种版式，等待真实后台回调修复全部 45 张图片。Word 的四类空闲任务共享一个调度入口，避免 OnTime 单一槽位相互覆盖；AutoOpen 在文档创建后启用事件与监听。

真实主题字体文档还需单独验收：保持正文主题字体、原公式字体、字号、行内 text/display style 与换行不变。在 Word 打开可写的验收副本后，逐一执行以下命令，将 `MODE` 替换为 `bottom`、`baseline`、`center`、`top`、`auto`，每次等待自动位置检查通过后截取原生窗口：

```bash
npx tsx scripts/word_chinese_theme_alignment_regression.mts --mode=MODE --output=BUILD_DIRECTORY/MODE-live.json
python3 scripts/word_chinese_theme_alignment_pixels.py BUILD_DIRECTORY
```

第二条命令测量五张 `MODE-native.png` 中两行正文和五个公式的数学轴，按参考图的正文墨迹高度归一化，阈值保持为 3 参考像素。不能修改验收副本的字体、字号或行宽以匹配固定测试样例。

主题字体先通过 Word 单个正文字符解析为实际字体，再在客户端进程注册 Word 私有字体并测量笔画边界。顶部、居中补偿还须考虑实际字体上下留白的不对称程度。主体字母标志由 LaTeX 内容解析，不能用固定 3.5 pt 数学轴判断；KaTeX 缩放会改变轴距，几何判定会将普通分式、积分等误判为主体字母。公式身份恢复沿用保存的元数据和文档变量。

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
