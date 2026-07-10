# VisualTeX

VisualTeX 是一款基于 Tauri 2、React、TypeScript、MathLive 和 CodeMirror 6 的桌面可视化 LaTeX 公式编辑器。

## 当前版本

第一版 MVP 已包含：

- MathLive 结构化公式编辑器；
- 分式、根式、积分、求和、上下标、矩阵、希腊字母等公式工具栏；
- 输入反斜杠后出现的命令候选；
- 前缀、模糊、中文关键词和英文别名检索；
- 键盘上下选择、Enter/Tab 确认、Esc 关闭；
- 基于频率、前缀和最近使用时间的个性化排序；
- 不限行数的多公式编辑，按 Enter 新建公式行；
- 每行公式独立生成 `$$...$$` 源码并与 CodeMirror 6 双向同步；
- 公式中的中文自动转换为 `\\text{中文}`；
- 撤销、重做、缩放、深浅色模式；
- 纯公式、行内公式、独立公式和 equation 环境复制；
- 本地公式历史；
- VisualTeX JSON 文档导入和导出；
- macOS / Windows 桌面打包基础。

## 开发

\`\`\`bash
npm install
npm run tauri:dev
\`\`\`

只运行前端：

\`\`\`bash
npm run dev
\`\`\`

构建检查：

\`\`\`bash
npm run build
cargo check --manifest-path src-tauri/Cargo.toml
\`\`\`

桌面应用打包：

\`\`\`bash
npm run tauri:build
\`\`\`

## 核心原则

VisualTeX 以 LaTeX 字符串作为单一数据源。可视化编辑、命令候选、工具栏插入和源码编辑始终作用于同一个公式，不依赖 TeX Live，也不进行 PDF 编译。
