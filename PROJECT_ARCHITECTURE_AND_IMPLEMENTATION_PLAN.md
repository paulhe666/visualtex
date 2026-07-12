# VisualStudio / VisualTeX Next：离线实时双向论文编辑器项目设计与实施文档

> 文档状态：架构基线  
> 项目性质：全新项目，从零设计与实现  
> 参考项目：`paulhe666/visualtex`，仅借鉴可复用思路与局部实现，不在原项目上继续扩建  
> 目标平台：macOS、Windows、Linux；后续接入 VS Code、TeXstudio 及其他编辑器  
> 核心要求：离线、源码与可视化双向实时编辑、真实 PDF 排版一致性、OCR、完整 LaTeX 项目编译、插件与外部编辑器适配

---

## 1. 项目最终目标

本项目要实现一套完整的离线学术论文编辑系统。用户可以同时打开：

- LaTeX 源码编辑区；
- 与最终 PDF 排版一致的页面编辑区；
- 项目文件树、编译日志、文献、图片、公式和结构导航面板。

源码区和可视化区均可直接编辑，任何一侧的修改都会以增量操作同步到另一侧。系统不通过 HTML/CSS 重新模仿 LaTeX 的最终排版，而是始终以真实 TeX 编译生成的 PDF 作为权威排版结果，并在 PDF 页面上叠加结构化编辑层。

最终产品应同时具有三种身份：

1. **独立桌面论文编辑器**  
   可以创建、打开、编辑、编译和导出完整 LaTeX 项目。

2. **可复用编辑核心**  
   源码解析、文档模型、双向同步、编译、SyncTeX、PDF 映射、OCR 等能力独立于桌面界面。

3. **外部编辑器扩展能力**  
   通过 VS Code 扩展、TeXstudio Bridge 和通用本地协议，将双向可视化编辑能力接入其他成熟编辑器。

---

## 2. 必须明确的产品边界

### 2.1 本项目承诺的“排版一致”

“可视化页面与最终 PDF 一致”定义为：

- 非编辑状态下显示的页面，直接来源于当前 LaTeX 项目真实编译生成的 PDF；
- 字体、字距、断行、分页、单双栏、浮动体、页眉页脚、公式编号和参考文献均以 TeX 编译结果为准；
- 导出的 PDF 与预览所显示的 PDF 是同一个构建产物或同一构建流程的产物。

### 2.2 本项目承诺的“实时编辑”

“实时”分为两条通道：

1. **交互实时通道**  
   键盘输入、光标、选择、段落文本、公式内容和图片属性在本地编辑覆盖层中立即更新，不等待整篇 TeX 编译。

2. **权威排版通道**  
   用户停止连续输入后触发增量调度的后台 TeX 编译。编译完成后，PDF 页面、分页、引用、编号和浮动体位置更新为权威结果。

因此，正在输入的局部区域可以先由编辑覆盖层显示；页面完整排版在编译成功后确认。项目不虚假承诺“任意复杂论文每次按键后都能在一帧内完成完整 TeX 编译”。

### 2.3 不承诺自动理解任意宏

LaTeX 允许用户定义任意宏和环境。系统无法可靠推断每个自定义宏的视觉语义，因此文档节点分为：

- `native`：完全可视化编辑；
- `partial`：可修改部分属性；
- `opaque`：显示真实编译结果，但内部使用源码编辑。

未知结构必须无损保留，不得为了可视化而改写或丢失用户源码。

### 2.4 第三方插件兼容边界

独立桌面版第一阶段不承诺直接运行任意 VS Code `.vsix` 插件。兼容分为：

- 声明式资源：Snippet、主题、TextMate Grammar；
- 标准服务：LSP、Formatter、Linter、CLI；
- 编辑器专属扩展：通过 VS Code 适配器使用；
- 完整 VS Code Extension Host：仅作为未来基于 Theia/Code-OSS 的独立产品线评估，不作为核心桌面版的前置条件。

---

## 3. 总体技术决策

最终推荐技术组合如下：

| 子系统 | 选择 |
|---|---|
| 核心语言 | Rust |
| 桌面容器 | Tauri 2 |
| 前端语言 | TypeScript，开启 strict |
| 前端框架 | React |
| 源码编辑器 | CodeMirror 6 |
| 结构化文档编辑 | ProseMirror 直接集成 |
| 数学公式编辑 | MathLive，封装为独立 MathNodeEditor |
| 增量语法解析 | Tree-sitter + 自定义容错 LaTeX 语义层 |
| 编译主后端 | 用户本机 TeX Live / MacTeX / MiKTeX + latexmk |
| 可选编译后端 | Tectonic |
| PDF 渲染 | PDFium，由 Rust 核心统一调用 |
| 源码/PDF 映射 | SyncTeX + 影子编译插桩 |
| OCR 协调层 | Rust |
| OCR 推理层 | Python + PaddleOCR，作为可选离线 Sidecar |
| 进程与异步 | Tokio |
| 本地通信 | JSON-RPC 2.0；Unix Domain Socket / Named Pipe / stdio |
| 项目元数据 | JSON；缓存和索引可使用 SQLite |
| 文件监听 | Rust `notify` |
| 日志 | Rust `tracing` |
| Rust 测试 | cargo test / nextest / proptest / insta / cargo-fuzz |
| 前端测试 | Vitest + Testing Library + Playwright |
| Monorepo | Cargo Workspace + pnpm Workspace |
| CI | GitHub Actions 多平台矩阵 |

---

## 4. 为什么选择这些语言和架构

## 4.1 核心语言筛选

### 候选语言比较

| 语言 | 优点 | 主要问题 | 结论 |
|---|---|---|---|
| Rust | 高性能、内存安全、适合解析器/文件/进程/PDF/IPC；跨平台；与 Tauri 原生结合 | 学习和开发复杂度高于 Go/TypeScript | **核心首选** |
| C++ | PDF、TeX、Qt 等生态成熟；性能最高 | 内存安全、构建和跨平台依赖复杂；业务迭代成本高 | 只通过 C API 使用 PDFium 等底层库 |
| Go | 并发和服务开发简单；跨平台打包方便 | GUI、PDF 原生库、复杂文本编辑和细粒度内存控制不如 Rust | 不作为核心 |
| TypeScript/Node.js | UI 和插件生态强；开发快 | 长生命周期编译进程、文件安全、原生 PDF 和高性能解析不适合作为唯一核心 | 用于界面和编辑器扩展 |
| Python | OCR/机器学习生态最强 | 启动、分发、内存、类型和桌面核心可靠性不足 | 仅作为 OCR Sidecar |
| Kotlin/JVM | 跨平台、类型系统好 | JVM 体积和桌面生态不适合本项目核心分发 | 不选 |
| C#/.NET | Windows 体验好；Avalonia 可跨平台 | macOS/Linux 原生体验和 LaTeX/PDF 底层整合成本较高 | 不选 |
| Swift | macOS 原生体验好 | 无法作为 Windows/Linux 主核心 | 不选 |
| Dart/Flutter | 跨平台 UI 一致 | 富文本、源码编辑器、VS Code Webview 代码复用弱；原生库桥接复杂 | 不选为主 UI |
| Zig | 低层控制和跨平台潜力 | 编辑器、异步、生态成熟度不足 | 暂不采用 |

### 选择结论

Rust 负责所有不能丢数据、需要高并发、跨平台原生能力或需要长期稳定运行的部分：

- 项目和文件系统；
- 文本缓冲与修订版本；
- LaTeX 增量解析；
- 双向同步调度；
- TeX 编译；
- PDFium；
- SyncTeX；
- 日志诊断；
- OCR 进程管理；
- 本地 RPC；
- 安全边界；
- 缓存和恢复。

TypeScript 只负责用户界面、浏览器编辑组件和 VS Code 扩展层，不作为论文内容的最终权威存储。

Python 只负责当前最成熟的 OCR 推理流程。OCR 失败不得导致主编辑器崩溃。

---

## 4.2 桌面框架筛选

| 框架 | 评价 |
|---|---|
| Tauri 2 | 可复用 Web 编辑器生态，Rust 后端自然，包体相对轻，支持 Sidecar 和权限控制 |
| Electron | 插件和 Node 生态强，但内存和包体较大；若目标是完整 VS Code Extension Host 才更有优势 |
| Qt 6 | 原生能力强，但需要 C++/QML，难以直接复用 ProseMirror、CodeMirror、MathLive 和 VS Code Webview 代码 |
| Flutter | 普通业务 UI 跨平台优秀，但不适合复用成熟 Web 富文本/源码编辑器 |
| Theia/Code-OSS | 适合完整 IDE 和 VS Code 插件生态，但整体很重，不适合作为第一版独立编辑器外壳 |

**最终选择：Tauri 2。**

理由：

- 核心服务本来就选 Rust；
- CodeMirror、ProseMirror、MathLive 可直接运行；
- 相同前端组件可复用于 VS Code Webview；
- 可以打包 OCR 和辅助进程；
- 可以对文件、Shell 和 Sidecar 权限做显式限制。

未来若商业目标明确要求在独立应用中运行大量 VS Code 扩展，再建立 Theia/Code-OSS 产品线，核心服务仍可复用。

---

## 4.3 前端框架筛选

### React

选择 React 的原因：

- ProseMirror、MathLive、PDF Canvas、虚拟列表和复杂编辑状态有大量成熟案例；
- VS Code Webview 中易复用；
- 招聘、维护和第三方组件生态最大；
- 当前参考项目已有 React 经验，但新项目不复制原组件结构。

### 为什么不选择 Vue/Svelte/Solid

这些框架都能实现界面，但本项目的难点不在普通 UI 性能，而在：

- 编辑器事务；
- DOM Selection；
- IME；
- PDF 覆盖层；
- 光标位置映射；
- 大量成熟库的生命周期适配。

React 的生态和可维护性综合风险最低。

---

## 4.4 源码编辑器选择

**选择 CodeMirror 6。**

原因：

- 增量 `ChangeSet` 和事务模型适合把源码变化转换成统一的 `TextEdit`；
- 模块化，适合 Tauri 和 VS Code Webview；
- 比 Monaco 更轻；
- 独立桌面版不需要完整模拟 VS Code 工作台；
- 对多实例、装饰、诊断和自定义语言支持可控。

VS Code 扩展模式下，不强制嵌入 CodeMirror；可以继续使用 VS Code 原生文本编辑器，并通过扩展 API 与可视化页面同步。

---

## 4.5 结构化可视化编辑器选择

**选择 ProseMirror，不直接依赖高层封装作为核心。**

原因：

- 文档由严格 Schema 控制；
- 每次编辑产生 Transaction；
- 支持映射、撤销、插件、节点视图；
- 适合建立段落、标题、公式、图片、表格和 Raw LaTeX 节点；
- 可将 ProseMirror Transaction 转换成核心 `EditOperation`；
- 可视化编辑器不会退化成不可控的 `contenteditable` HTML。

Tiptap 可用于原型，但核心层直接基于 ProseMirror，避免高层 API 限制源码位置映射。

---

## 4.6 LaTeX 解析器选择

LaTeX 不是普通上下文无关语言，宏定义可能改变后续语义。单一解析器无法同时实现：

- 编辑时容错；
- 完整宏展开；
- 快速增量；
- 与真实 TeX 完全一致。

因此采用双层方案：

### 第一层：Tree-sitter 容错语法树

负责：

- 增量更新；
- 花括号、命令、环境、注释和数学区域识别；
- 在源码不完整时仍生成可用 CST；
- 快速确定受影响范围。

### 第二层：VisualTeX 自定义语义层

负责识别可视化支持结构：

- `documentclass`；
- `usepackage`；
- `section` 系列；
- 普通段落；
- 行内和行间公式；
- `figure`、`table`、列表；
- `label`、`ref`、`cite`；
- `input`、`include`；
- 常见定理和算法环境；
- 已注册模板宏。

无法可靠理解的结构生成 `OpaqueNode`，源码原样保存。

### 不采用 Pandoc 作为主解析器

Pandoc 适合文档转换，但不适合作为每次按键都运行的无损 LaTeX 增量编辑核心。Pandoc 可以在导入、导出或格式转换功能中作为可选工具。

---

## 4.7 PDF 引擎选择

**选择 PDFium。**

原因：

- Chromium 使用的 PDF 引擎；
- 适合跨平台页面渲染、文字信息、坐标和位图输出；
- 许可比 AGPL 的 MuPDF 更适合闭源或商业扩展；
- Rust 核心可通过稳定封装调用；
- 桌面版和编辑器插件均可通过统一核心服务获取页面瓦片和布局数据。

PDF 页面不能交给系统浏览器内置查看器作为核心，因为需要：

- 精确页坐标；
- 点击命中；
- 页面瓦片缓存；
- 高 DPI；
- 编辑覆盖层；
- 自动化像素回归测试。

---

## 4.8 编译器选择

### 主路径

优先使用用户本机已有环境：

- macOS：MacTeX / TeX Live；
- Windows：MiKTeX / TeX Live；
- Linux：TeX Live；
- 调度器：`latexmk`；
- 引擎：pdfLaTeX、XeLaTeX、LuaLaTeX；
- 文献：BibTeX、Biber；
- 索引：makeindex 等。

这是复杂模板兼容性最高的方案。

### 可选路径

Tectonic 作为：

- 快速安装模式；
- 受控模板；
- 简单项目；
- CI 或便携模式。

不得把 Tectonic 作为唯一后端，否则复杂期刊模板、shell-escape、外部程序和特殊工作流兼容性不足。

---

## 4.9 OCR 技术选择

### 第一阶段

Python Sidecar + PaddleOCR：

- 公式识别；
- 文本 OCR；
- 版面检测；
- 多栏阅读顺序；
- 表格和图注识别；
- 结果置信度；
- 可取消和进度上报。

选择 Python 是因为模型官方推理、预处理和部署资料最成熟。

### 后续优化

对启动体积或推理性能有明确需求后，再评估：

- Paddle C++ 推理；
- ONNX Runtime；
- CoreML；
- DirectML；
- Metal；
- CUDA。

不得在第一阶段为了消除 Python 而重写全部预后处理，从而拖慢核心编辑器开发。

---

## 5. 总体架构

```text
┌──────────────────────────────────────────────────────────┐
│                     应用与适配层                          │
│  Tauri Desktop │ VS Code Extension │ TeXstudio Bridge    │
│  CLI            │ Future Theia      │ Other IDE Adapter   │
└─────────────────────────────┬────────────────────────────┘
                              │ JSON-RPC / Tauri IPC
┌─────────────────────────────▼────────────────────────────┐
│                    VisualTeX Core (Rust)                  │
│                                                          │
│ Project │ Buffer │ Revision │ Parser │ Semantic Model    │
│ Sync Engine │ Undo Log │ Compiler │ Diagnostics          │
│ PDFium │ SyncTeX │ Layout Map │ Cache │ Recovery          │
│ OCR Coordinator │ Plugin Adapter │ Security              │
└───────────────┬───────────────────────┬──────────────────┘
                │                       │
       ┌────────▼────────┐     ┌────────▼──────────┐
       │ TeX Toolchains  │     │ OCR Sidecar       │
       │ latexmk/xelatex │     │ Python/PaddleOCR  │
       │ biber/bibtex    │     │ optional package  │
       └─────────────────┘     └───────────────────┘

┌──────────────────────────────────────────────────────────┐
│                  TypeScript 编辑界面                     │
│ Source Editor │ ProseMirror │ MathLive │ PDF Canvas      │
│ Overlay Editor │ Project Tree │ Diagnostics │ OCR Review │
└──────────────────────────────────────────────────────────┘
```

---

## 6. 数据权威关系

项目必须严格遵守以下关系：

```text
.tex / .bib / project files
            │
            │ 唯一权威持久内容
            ▼
      Source Buffer + Revision
            │
            ├── 派生：容错语法树
            ├── 派生：语义文档模型
            ├── 派生：可视化编辑节点
            ├── 派生：编译中间文件
            ├── 派生：PDF
            └── 派生：位置映射与索引
```

禁止：

- 把 ProseMirror JSON 作为论文唯一保存格式；
- 把 Zustand/React State 作为权威文档；
- 每次视觉编辑后重新生成整篇 LaTeX；
- 解析失败时丢弃未知源码；
- 用格式化器静默改写用户全部文件。

---

## 7. 核心数据模型

## 7.1 文本缓冲

```rust
pub struct DocumentBuffer {
    pub file_id: FileId,
    pub path: PathBuf,
    pub revision: Revision,
    pub text: Rope,
    pub dirty: bool,
}
```

推荐使用 Rope 或 Piece Table，避免大文件每次修改都复制完整字符串。

## 7.2 增量文本操作

```rust
pub struct TextEdit {
    pub operation_id: OperationId,
    pub origin: EditOrigin,
    pub file_id: FileId,
    pub base_revision: Revision,
    pub start_byte: usize,
    pub end_byte: usize,
    pub replacement: String,
}
```

`EditOrigin` 至少包含：

- `SourceEditor`
- `VisualEditor`
- `MathEditor`
- `Ocr`
- `ExternalFileChange`
- `Formatter`
- `UndoRedo`
- `Plugin`

## 7.3 语义节点

```rust
pub struct VisualNode {
    pub id: NodeId,
    pub kind: NodeKind,
    pub support: SupportLevel,
    pub source: SourceSpan,
    pub children: Vec<NodeId>,
    pub attributes: NodeAttributes,
}
```

`NodeKind` 包含：

- document；
- preamble；
- title；
- author；
- abstract；
- section/subsection；
- paragraph；
- text；
- inline_math；
- display_math；
- figure；
- table；
- list；
- theorem；
- citation；
- reference；
- footnote；
- bibliography；
- raw_latex。

## 7.4 PDF 布局节点

```rust
pub struct LayoutBox {
    pub node_id: NodeId,
    pub page: u32,
    pub rect: PdfRect,
    pub baseline: Option<f32>,
    pub source: SourceSpan,
    pub confidence: MappingConfidence,
}
```

## 7.5 编译结果

```rust
pub struct CompileArtifact {
    pub build_id: BuildId,
    pub source_revision: ProjectRevision,
    pub pdf_path: PathBuf,
    pub synctex_path: Option<PathBuf>,
    pub diagnostics: Vec<Diagnostic>,
    pub status: CompileStatus,
}
```

只允许显示“与当前修订匹配”或明确标注为旧修订的 PDF。不得把过时编译结果误认为当前内容。

---

## 8. 双向同步算法

## 8.1 源码到可视化

1. CodeMirror 产生增量 ChangeSet；
2. 转换为一个或多个 `TextEdit`；
3. Rust 核心验证 `base_revision`；
4. 更新 Rope/Piece Table；
5. 对 Tree-sitter 树执行增量 edit；
6. 重新解析最小受影响区域；
7. 语义层比较旧节点和新节点；
8. 尽量保持未改变节点的 `NodeId`；
9. 生成 `VisualPatch`；
10. ProseMirror 只更新受影响节点；
11. 编译调度器接收新修订。

## 8.2 可视化到源码

1. ProseMirror 或 MathLive 产生 Transaction；
2. 根据 `NodeId` 找到源范围；
3. 使用节点专用 Serializer 生成最小 LaTeX 片段；
4. 生成 `TextEdit`；
5. 核心验证范围和修订；
6. 应用到权威文本缓冲；
7. 向 CodeMirror 或 VS Code `TextDocument` 派发变化；
8. 回传新修订；
9. UI 不得再次把同一 `operation_id` 回传，防止循环。

## 8.3 不完整语法

用户可能暂时输入：

```latex
\begin{figure
```

处理规则：

- 源码编辑永远优先接受；
- 保留上一次有效语义节点；
- 当前受影响区变成 `UnstableNode`；
- 可视化区显示源码占位块或上一次有效预览；
- 语法恢复后重新结构化；
- 编译失败时保留上一版成功 PDF，并清楚显示“预览落后于源码修订”。

## 8.4 撤销与重做

撤销栈位于 Rust 核心，不分别依赖 CodeMirror、ProseMirror 和 MathLive 的私有历史。

每个事务记录：

- 前后修订；
- 原始操作；
- 逆操作；
- 选择状态；
- 来源；
- 时间；
- 可合并组。

连续输入可按语义分组。源码区和可视化区的操作必须处在同一个全局顺序中。

---

## 9. PDF 级可视化编辑实现

## 9.1 页面层级

```text
PDF page bitmap/vector layer
        +
selection/highlight layer
        +
editable overlay layer
        +
caret/IME layer
        +
diagnostic and mapping layer
```

### 非编辑状态

只显示真实 PDF 渲染页面。

### 点击可编辑节点

1. PDF 坐标转页面坐标；
2. 命中 `LayoutBox`；
3. 找到 `NodeId` 和 `SourceSpan`；
4. 打开对应节点编辑覆盖层；
5. 隐藏或遮盖该节点原 PDF 图像区域；
6. 在覆盖层中立即编辑；
7. 同步源码；
8. 后台编译；
9. 新 PDF 到达后重新计算映射；
10. 光标重新锚定；
11. 关闭覆盖层或继续编辑。

## 9.2 为什么需要 SyncTeX 和自有插桩同时存在

SyncTeX 可提供通用的源文件行列与 PDF 区域映射，但不足以稳定表达每个结构化节点的边界。

因此：

- 第一层使用标准 SyncTeX，保证通用正向和反向搜索；
- 第二层生成影子编译源，在不修改用户文件的情况下插入位置标记；
- 影子源与真实源保存映射表；
- 编译后读取位置标记，生成 `NodeId → PageRect`。

不得直接往用户 `.tex` 中永久插入 VisualTeX 私有命令。

## 9.3 PDF 缓存

- 页面按缩放级别分瓦片渲染；
- 只渲染视口附近页面；
- 缩放时先复用低分辨率缓存，再替换高分辨率瓦片；
- 新构建只使变化页面失效；
- 缓存键包含 PDF 哈希、页码、缩放、DPI 和颜色模式。

---

## 10. 编译系统

## 10.1 工具链探测

启动或打开项目时检测：

- `latexmk`
- `pdflatex`
- `xelatex`
- `lualatex`
- `bibtex`
- `biber`
- `makeindex`
- `tectonic`

检测结果记录版本、路径和可用性，不修改用户全局环境。

## 10.2 项目编译配置

项目配置只保存 VisualTeX 自身信息：

```json
{
  "rootFile": "main.tex",
  "engine": "xelatex",
  "builder": "latexmk",
  "outputDirectory": ".visualtex/build",
  "shellEscape": false
}
```

真实源码仍为标准 LaTeX 项目。

## 10.3 编译调度

1. 文本修订变化；
2. 延迟调度，连续输入合并；
3. 若旧构建尚未开始，直接替换；
4. 若旧构建正在执行，按策略取消或允许快速完成；
5. 在隔离构建目录运行；
6. 实时解析 stdout/stderr；
7. 将日志转换为结构化诊断；
8. 校验输出 PDF；
9. 读取 SyncTeX；
10. 原子替换当前成功构建。

## 10.4 安全

默认：

- 禁止 `--shell-escape`；
- 限制工作目录；
- 外部命令白名单；
- 设置 CPU 时间和总超时；
- 限制递归进程；
- 不自动执行项目中的任意脚本；
- 不信任下载模板；
- 打开陌生项目时进入 Restricted Mode；
- OCR、TeX 和插件进程彼此隔离。

---

## 11. OCR 系统

## 11.1 公式 OCR

流程：

1. 粘贴、拖入、截图或选择图片；
2. 检测透明背景和深色背景；
3. 方向、裁边和清晰度预处理；
4. 调用公式模型；
5. 返回候选 LaTeX 和置信度；
6. 在 MathLive 中预览；
7. 用户确认后生成 `TextEdit`；
8. OCR 操作进入统一撤销栈。

## 11.2 整页 OCR

流程：

1. PDF 或图片页输入；
2. 页面方向识别；
3. 透视和弯曲校正；
4. 版面区域检测；
5. 多栏阅读顺序恢复；
6. 文本、公式、表格、图片、图注分别识别；
7. 生成中间 `OcrDocument`；
8. 用户在校对界面确认；
9. 转换为 VisualNode；
10. 序列化为 LaTeX 项目。

OCR 不得直接生成不可审查的大段源码。

## 11.3 OCR 结果模型

每个区域包含：

- bbox；
- 类型；
- 文本或 LaTeX；
- 置信度；
- 模型版本；
- 原图裁剪；
- 用户修改记录。

## 11.4 离线分发

提供：

- 主程序；
- 可选 OCR Runtime；
- 可选模型包；
- 模型校验和；
- 离线安装入口；
- 不联网模式。

主程序在没有 OCR 包时仍应完整运行。

---

## 12. 插件和外部编辑器

## 12.1 VS Code 扩展

实现 `CustomTextEditorProvider` 或侧边 Webview 模式。

核心要求：

- 使用 VS Code `TextDocument` 作为源码权威；
- 所有可视化修改通过 `WorkspaceEdit` 写回；
- 监听 `onDidChangeTextDocument`；
- 支持 VS Code Undo/Redo；
- 支持多视图同一文档；
- 支持工作区信任；
- 复用 VisualTeX 前端页面组件；
- 通过本地 Core 服务进行编译、PDF、OCR 和映射。

VS Code 用户仍可继续使用 LaTeX Workshop、Git、拼写检查和其他扩展。

## 12.2 TeXstudio Bridge

第一阶段不修改 TeXstudio 源码，采用外部桥接：

- 启动 VisualTeX 并传入当前项目；
- 文件变化监听；
- 光标/选区位置通信；
- SyncTeX 正反向跳转；
- 可选择复用 TeXstudio 编译命令；
- 本地 Socket/JSON-RPC；
- 提供 TeXstudio 宏或脚本安装包。

完整嵌入 TeXstudio 主窗口仅在后续获得稳定插件 API或决定维护 TeXstudio 分支时实施。

## 12.3 通用编辑器协议

提供：

- `visualtex open <project>`
- `visualtex compile`
- `visualtex inverse-search`
- `visualtex forward-search`
- 本地 JSON-RPC；
- LSP 兼容接口；
- URI Scheme；
- 编辑器适配 SDK。

---

## 13. 新项目目录结构

```text
visualstudio/
├── Cargo.toml
├── package.json
├── pnpm-workspace.yaml
├── rust-toolchain.toml
├── README.md
├── docs/
│   ├── PROJECT_ARCHITECTURE_AND_IMPLEMENTATION_PLAN.md
│   ├── protocol.md
│   ├── security.md
│   ├── testing.md
│   └── adr/
├── apps/
│   ├── desktop/
│   │   ├── src/
│   │   └── src-tauri/
│   ├── vscode-extension/
│   └── cli/
├── crates/
│   ├── vt-core/
│   ├── vt-protocol/
│   ├── vt-buffer/
│   ├── vt-latex-syntax/
│   ├── vt-latex-semantic/
│   ├── vt-sync/
│   ├── vt-project/
│   ├── vt-compiler/
│   ├── vt-diagnostics/
│   ├── vt-pdf/
│   ├── vt-synctex/
│   ├── vt-layout-map/
│   ├── vt-ocr/
│   ├── vt-rpc/
│   └── vt-security/
├── packages/
│   ├── protocol-ts/
│   ├── ui-shell/
│   ├── source-editor/
│   ├── visual-editor/
│   ├── math-editor/
│   ├── pdf-viewer/
│   └── test-utils/
├── services/
│   └── ocr-python/
├── fixtures/
│   ├── latex-corpus/
│   ├── templates/
│   ├── ocr/
│   └── visual-golden/
├── tests/
│   ├── end-to-end/
│   ├── compatibility/
│   ├── performance/
│   └── security/
└── scripts/
```

初期可以减少 crate 数量，但模块边界必须按上述职责设计，避免所有逻辑再次集中到一个前端 `App.tsx`。

---

## 14. 从参考 VisualTeX 项目借鉴什么

现有 VisualTeX 仅作为实现经验库。

### 可借鉴

- MathLive 与 React 生命周期的封装方式；
- 公式选区、光标恢复和结构化插入经验；
- 公式命令目录和候选搜索；
- OCR 图片预处理；
- Rust 持久 Python Worker 和逐行 JSON 协议；
- OCR 进度、取消和模型缓存；
- 文档级历史事务思路；
- Tauri 多平台打包和安装器经验；
- 现有编辑回归测试中的问题样例。

### 不直接沿用

- `FormulaLine[]` 文档模型；
- Zustand 作为文档内容权威；
- localStorage 保存正文；
- 源码草稿点击“应用”的同步模式；
- 整段 LaTeX 替换；
- 公式格式化器承担整篇文档解析；
- 单体 App 组件；
- 只面向单公式/多公式行的文件格式。

### 代码引用规则

若复制局部代码：

- 单独列出来源文件；
- 记录原提交；
- 重新编写适配测试；
- 不把旧状态模型带入新核心；
- 优先复制算法测试，不复制强耦合 UI。

---

# 15. 分阶段实施计划

以下阶段必须按依赖顺序推进。每个阶段都有明确退出标准，未通过验收不得把后续 UI 功能建立在不稳定核心上。

---

## 阶段 0：需求冻结与工程基线

### 目标

建立全新仓库、工程规范、协议和最小可运行骨架。

### 实施步骤

1. 初始化 Cargo Workspace 和 pnpm Workspace；
2. 创建 Tauri Desktop 空应用；
3. 创建 Rust Core 进程或库；
4. 建立 Rust/TypeScript 共享协议生成方式；
5. 配置 formatter、lint、typecheck；
6. 配置单元测试和 GitHub Actions；
7. 建立 ADR；
8. 定义支持平台和 CPU 架构；
9. 定义文件编码策略，统一 UTF-8；
10. 定义错误码、日志字段和崩溃报告格式；
11. 写明 Restricted Mode 和默认安全策略。

### 测试

- 三个平台构建最小应用；
- Rust、TypeScript 测试可运行；
- IPC 往返测试；
- CI 对格式错误、类型错误和测试失败进行阻断；
- 创建带中文、日文、数学符号路径的临时项目。

### 退出标准

- 新项目不依赖旧 VisualTeX Store；
- 桌面应用可启动；
- Rust 与 TypeScript 可交换版本化消息；
- CI 基线稳定；
- 所有架构决策有 ADR。

---

## 阶段 1：项目系统与权威文本缓冲

### 目标

安全打开、编辑、保存多文件 LaTeX 项目。

### 实施步骤

1. 项目根目录和主文件检测；
2. 支持 `main.tex`、子目录和 `.bib`；
3. 实现 Rope/Piece Table 文本缓冲；
4. 实现 Revision；
5. 实现增量 `TextEdit`；
6. 实现脏文件、保存、另存为；
7. 实现外部文件变化检测；
8. 实现冲突对话和三种处理策略；
9. 实现自动恢复日志；
10. 实现崩溃后恢复；
11. 实现统一 Undo/Redo 基础；
12. CodeMirror 接入真实文本缓冲。

### 测试

- UTF-8 边界和多字节字符；
- CRLF/LF；
- 大文件随机编辑属性测试；
- 保存后字节完全一致；
- 外部编辑冲突；
- 进程强制退出后恢复；
- Undo/Redo 1000 次随机操作；
- 文件重命名、移动和删除。

### 退出标准

- 可稳定编辑标准 `.tex` 文件；
- 不使用 localStorage 保存正文；
- 任意编辑操作都有修订号；
- 崩溃和外部修改不会静默丢数据。

---

## 阶段 2：容错增量 LaTeX 语法树

### 目标

源码每次变化后快速生成稳定、容错、可定位的语法树。

### 实施步骤

1. 选定或建立 Tree-sitter LaTeX Grammar 分支；
2. 覆盖命令、环境、参数、注释、数学模式；
3. 建立字节范围到行列映射；
4. 实现 Tree-sitter 增量 edit；
5. 对不完整命令和环境生成 Error Node；
6. 实现 include/input 文件依赖图；
7. 处理 verbatim、minted 等特殊区域；
8. 建立解析 Corpus；
9. 建立 AST 快照测试；
10. 对解析器进行 fuzz。

### 测试

- 括号未闭合；
- 环境缺失；
- 嵌套数学环境；
- 中文正文；
- 注释中的命令；
- verbatim；
- 自定义宏；
- 100 个常见模板样例；
- 随机删除和插入字符后不崩溃；
- 增量树与全量解析结果等价。

### 退出标准

- 源码暂时非法时仍可继续编辑；
- 增量解析不会阻塞 UI；
- 解析器对未知宏保持原文；
- Fuzz 不出现崩溃和越界。

---

## 阶段 3：语义文档模型

### 目标

把语法树转换为稳定的论文结构节点。

### 实施步骤

1. 建立 `VisualNode`；
2. 建立稳定 NodeId 算法；
3. 识别章节、段落、公式、图片、表格、列表；
4. 识别 label/ref/cite；
5. 识别文档类和宏包；
6. 建立 `native/partial/opaque`；
7. 建立节点 Serializer；
8. 建立源码范围映射；
9. 支持多文件节点；
10. 建立模板扩展注册表；
11. 实现旧树和新树 diff；
12. 输出最小 `VisualPatch`。

### 测试

- Parse → Semantic → Serialize 局部往返；
- 未修改节点序列化字节不变；
- Opaque 节点完全无损；
- NodeId 在相邻文本编辑后保持稳定；
- 跨文件 section；
- 引用和标签索引；
- 复杂模板不丢命令。

### 退出标准

- 常见论文结构可被识别；
- 不认识的结构不会丢失；
- 最小局部编辑不触发整篇重建；
- 语义节点具有稳定源码范围。

---

## 阶段 4：源码与结构化编辑双向同步

### 目标

实现不依赖 PDF 的第一版真正双向编辑。

### 实施步骤

1. 建立 ProseMirror Schema；
2. 实现段落、标题、公式、图片和 Raw LaTeX NodeView；
3. 接入 MathLive；
4. ProseMirror Transaction 转 `TextEdit`；
5. Source `TextEdit` 转 `VisualPatch`；
6. 防止 operation 回环；
7. 实现选择和光标映射；
8. 实现统一 Undo/Redo；
9. 实现 IME 和中文输入；
10. 实现剪贴板；
11. 实现源码暂时非法时的降级；
12. 实现多视图一致性。

### 测试

- 源码输入后可视化自动更新；
- 可视化输入后源码自动更新；
- 连续删除公式结构；
- 中文 IME；
- 复制带公式段落；
- 在两侧交替撤销；
- 旧 revision 操作被拒绝或重定基；
- 10 万次随机双向操作后内容一致；
- 不产生无限同步循环。

### 退出标准

- 两侧无“点击应用”按钮；
- 文本模型更新不依赖编译；
- 两侧交替编辑仍只有一份权威内容；
- 统一 Undo/Redo 顺序正确。

---

## 阶段 5：本地 TeX 编译系统

### 目标

完整编译真实 LaTeX 项目并输出结构化诊断。

### 实施步骤

1. 工具链探测；
2. latexmk Adapter；
3. XeLaTeX/pdfLaTeX/LuaLaTeX；
4. BibTeX/Biber；
5. 输出目录隔离；
6. 编译任务状态机；
7. 合并和取消旧任务；
8. 日志解析；
9. 诊断定位；
10. 构建缓存；
11. shell-escape 权限；
12. Tectonic Adapter；
13. 导出 PDF；
14. 清理构建产物。

### 测试

- article；
- ctexart；
- IEEEtran；
- revtex；
- elsarticle；
- beamer；
- BibTeX；
- Biber；
- 多次交叉引用；
- 图片路径包含空格和中文；
- 编译失败保留旧 PDF；
- 僵尸进程清理；
- 恶意命令受限；
- 用户取消编译。

### 退出标准

- 独立应用可以完成论文编译和 PDF 导出；
- 日志错误可跳到源文件；
- 旧任务不能覆盖新修订；
- 默认不执行危险外部命令。

---

## 阶段 6：PDFium 页面预览与 SyncTeX

### 目标

显示真实 PDF，并实现源码和页面的双向定位。

### 实施步骤

1. PDFium 跨平台封装；
2. 页面信息和缩略图；
3. 高 DPI 渲染；
4. 虚拟滚动；
5. 瓦片缓存；
6. SyncTeX 解析；
7. 源码到 PDF 高亮；
8. PDF 点击到源码；
9. 保留滚动位置；
10. 新旧 PDF 原子切换；
11. 页面变化 diff；
12. 自动化截图接口。

### 测试

- 单栏、双栏；
- 旋转页；
- 大页面；
- 100 页文档；
- 高分屏；
- 缩放；
- SyncTeX 正向/反向；
- 新 PDF 切换不闪白；
- 页面缓存正确失效；
- PDF 损坏时安全失败。

### 退出标准

- 预览就是实际导出 PDF；
- 点击源码可定位页面；
- 点击页面可定位源码；
- 大文档滚动流畅；
- PDF 解析错误不影响源码保存。

---

## 阶段 7：影子编译插桩与节点布局映射

### 目标

获得结构化节点级别的 PDF 坐标。

### 实施步骤

1. 生成影子源；
2. 保留真实源到影子源映射；
3. 在安全位置插入节点标记；
4. 编译影子项目；
5. 提取标记坐标；
6. 组合 SyncTeX；
7. 计算节点多行、多栏和跨页区域；
8. 维护 Mapping Confidence；
9. 未映射节点降级到 SyncTeX；
10. 对模板和宏包冲突建立黑名单/适配器。

### 测试

- 普通段落；
- 双栏段落；
- 跨页段落；
- 公式编号；
- figure；
- table；
- footnote；
- 标题跨栏；
- 自定义模板；
- 插桩前后 PDF 像素差异必须为零或在严格容差内。

### 退出标准

- 插桩不改变排版；
- 常见节点可以可靠命中；
- 低置信度映射不会错误开放可视化编辑；
- 用户源文件不出现私有标记。

---

## 阶段 8：PDF 页面上的直接编辑

### 目标

在真实 PDF 页面位置上编辑常见论文节点。

### 实施步骤

1. PDF 命中测试；
2. 编辑覆盖层；
3. 段落 Editor Overlay；
4. 公式 MathLive Overlay；
5. 标题 Overlay；
6. 图片属性 Overlay；
7. 简单表格 Overlay；
8. PDF 区域遮盖；
9. 光标锚定；
10. 新 PDF 到达后重定位；
11. 页面重排提示；
12. 编辑失败回退源码；
13. 低置信度节点只允许源码编辑。

### 测试

- 点击任意可编辑段落；
- 中文和英文输入；
- 行内公式；
- 多行公式；
- 删除导致分页变化；
- 双栏之间重排；
- 图片尺寸变化；
- 编译失败时继续保留输入；
- 光标在重新编译后尽量维持语义位置；
- 编辑覆盖层与 PDF 基线误差检查。

### 退出标准

- 用户能在页面上直接修改正文和公式；
- 非编辑状态完全显示真实 PDF；
- 编译失败不会丢失覆盖层输入；
- 页面重排后不会把编辑内容写到错误节点。

---

## 阶段 9：论文完整编辑能力

### 目标

达到可独立写完整论文的功能范围。

### 实施步骤

1. 项目模板；
2. 文档结构导航；
3. 图片管理；
4. 表格编辑；
5. 参考文献浏览和插入；
6. label/ref/cite 自动补全；
7. 章节拆分和合并；
8. 多文件重命名引用更新；
9. 用户宏和宏包索引；
10. LSP/texlab；
11. 格式化和检查工具；
12. Git 状态显示；
13. 搜索替换；
14. 拼写检查接口；
15. 项目设置和导出。

### 测试

使用一篇完整示例论文从空项目完成：

- 标题；
- 摘要；
- 多章节；
- 行内/行间公式；
- 图；
- 表；
- 引用；
- 参考文献；
- 附录；
- 双栏模板；
- 最终 PDF 导出。

### 退出标准

- 不依赖其他 IDE 也能完成完整论文；
- 项目可被普通 LaTeX 编辑器继续打开；
- VisualTeX 私有元数据删除后仍能编译；
- 不支持的宏仍可在源码中正常使用。

---

## 阶段 10：公式 OCR

### 目标

将现有公式 OCR 思路迁移到新架构。

### 实施步骤

1. 独立 OCR 协议；
2. Rust Worker Manager；
3. Python 环境管理；
4. 模型下载/离线导入；
5. 图片预处理；
6. 公式模型；
7. 候选和置信度；
8. 结果校对；
9. 插入当前节点；
10. 取消、超时和进度；
11. Worker 崩溃恢复；
12. 模型版本记录。

### 测试

- 白底黑字；
- 黑底白字；
- 透明背景；
- 手写公式；
- 多行公式；
- 模糊图；
- 大图；
- 取消；
- Worker 崩溃；
- 离线模型；
- OCR 插入后的 LaTeX 可编译率。

### 退出标准

- 公式 OCR 完全本地运行；
- OCR 失败不影响编辑器；
- 结果先校对再写入；
- OCR 操作可撤销。

---

## 阶段 11：整页论文 OCR

### 目标

将论文 PDF/扫描图片转换成可校对的结构化 LaTeX 项目。

### 实施步骤

1. 页面预处理；
2. 版面检测；
3. 多栏阅读顺序；
4. 标题/正文/公式/图片/表格分类；
5. 分模块识别；
6. OcrDocument；
7. 校对界面；
8. 转 VisualNode；
9. 生成项目；
10. 保留原页面对照；
11. 置信度和人工复核工作流；
12. 导入结果回归。

### 测试

建立 OCR Benchmark：

- 单栏中文；
- 双栏英文；
- 中英混排；
- 多公式；
- 复杂表格；
- 图注；
- 扫描倾斜；
- 低分辨率；
- 页眉页脚；
- 参考文献。

指标至少包括：

- 字符错误率；
- 公式可编译率；
- 公式结构准确率；
- 阅读顺序准确率；
- 表格结构准确率；
- 区域分类 F1；
- 人工修正次数。

### 退出标准

- OCR 输出不是黑盒大段文本；
- 用户可逐区域校对；
- 双栏阅读顺序可靠；
- 低置信度区域明显标识。

---

## 阶段 12：VS Code 扩展

### 目标

在 VS Code 中使用同一双向可视化编辑核心。

### 实施步骤

1. 创建扩展 Manifest；
2. Webview；
3. CustomTextEditorProvider；
4. TextDocument 同步；
5. WorkspaceEdit；
6. Undo/Redo；
7. 多视图；
8. Core 服务启动和连接；
9. PDF 页面和 OCR；
10. 选择同步；
11. 设置和命令；
12. 工作区信任；
13. 扩展打包和测试。

### 测试

- 原生文本编辑器和 VisualTeX 并排；
- 外部扩展修改文档；
- Git checkout；
- Undo/Redo；
- 多窗口；
- Remote/WSL 场景明确支持或明确拒绝；
- Core 断开重连；
- Webview reload；
- VS Code 重启恢复。

### 退出标准

- VS Code 仍以 TextDocument 为权威；
- 与 LaTeX Workshop 等扩展不冲突；
- 不形成文档修改循环；
- 独立桌面和 VS Code 使用同一协议。

---

## 阶段 13：TeXstudio 和通用 Bridge

### 目标

让其他编辑器调用 VisualTeX 页面、PDF 和同步能力。

### 实施步骤

1. JSON-RPC 服务；
2. 本地 Socket；
3. CLI；
4. URI Scheme；
5. TeXstudio 宏；
6. 打开当前项目；
7. 光标和 SyncTeX；
8. 文件变化；
9. 编译策略选择；
10. 通用 Adapter SDK。

### 测试

- TeXstudio 修改文件后 VisualTeX 更新；
- VisualTeX 修改后 TeXstudio 更新；
- 光标跳转；
- 多项目；
- 连接断开；
- 非法客户端；
- 版本不匹配；
- 路径包含中文和空格。

### 退出标准

- 不需要修改 TeXstudio 源码即可完成主要联动；
- 协议有版本和能力协商；
- Bridge 不暴露任意本地文件访问权限。

---

## 阶段 14：插件能力

### 目标

建立受控、可维护的扩展体系。

### 第一层

- Snippet；
- 主题；
- 命令包；
- 模板；
- 自定义宏语义声明。

### 第二层

- LSP；
- Formatter；
- Linter；
- CLI Tool；
- 文献工具；
- Git 工具。

### 第三层

VisualTeX 自有插件 API：

- WASI 优先；
- 能力权限；
- 文件范围；
- 命令注册；
- 语义节点扩展；
- OCR 后处理；
- 模板适配。

### 测试

- 权限拒绝；
- 插件崩溃；
- 超时；
- 插件版本升级；
- 恶意路径；
- 插件禁用；
- 项目级和全局配置隔离。

### 退出标准

- 插件不能绕过 Core 安全边界；
- 插件失败不破坏文档；
- 插件 API 有版本控制；
- 不声称兼容所有 VS Code 插件。

---

## 阶段 15：发布、性能和稳定性

### 目标

达到可长期使用的桌面产品质量。

### 实施步骤

1. macOS 签名和公证；
2. Windows 签名；
3. Linux AppImage/deb/rpm；
4. 自动更新；
5. OCR 可选包；
6. TeX 环境检测向导；
7. 崩溃恢复；
8. 诊断导出；
9. 性能分析；
10. 内存泄漏检查；
11. 辅助功能；
12. 国际化；
13. 隐私和离线模式；
14. Release 回归。

### 退出标准

- 三平台安装、卸载和更新可靠；
- 断网可编辑和编译已安装环境；
- OCR 包安装后可断网使用；
- 崩溃恢复不丢最近操作；
- 发布构建通过完整回归矩阵。

---

# 16. 测试体系

## 16.1 单元测试

覆盖：

- TextEdit；
- Revision；
- Rope；
- 语法节点；
- Serializer；
- NodeId；
- 操作去重；
- 编译状态机；
- 日志解析；
- PDF 坐标；
- SyncTeX；
- OCR 协议。

## 16.2 属性测试

必须验证：

- Edit + InverseEdit 恢复原文；
- 增量解析和全量解析等价；
- 语义未改变时序列化无损；
- 任意 Unicode 编辑不越界；
- 操作重放结果确定；
- 过时 Revision 不会覆盖新 Revision。

## 16.3 Fuzz

目标：

- LaTeX Parser；
- 日志 Parser；
- SyncTeX Parser；
- PDF 输入边界；
- JSON-RPC；
- OCR 返回数据；
- 项目配置。

## 16.4 Golden 测试

保存标准项目及预期：

- AST；
- Semantic Tree；
- Diagnostics；
- PDF 页数；
- 页面截图；
- 节点布局框；
- 导出源码。

## 16.5 视觉回归

对同一 TeX 工具链和固定字体环境：

- 渲染 PDF；
- 截图；
- 像素差；
- SSIM；
- 基线漂移；
- Overlay 基线误差；
- 插桩前后排版差异。

## 16.6 兼容性项目库

至少包含：

- article；
- report；
- book；
- ctexart；
- IEEEtran；
- revtex4-2；
- elsarticle；
- beamer；
- amsmath；
- cleveref；
- hyperref；
- biblatex；
- longtable；
- algorithm；
- minted；
- TikZ；
- 自定义 class；
- 多文件项目。

## 16.7 端到端测试

关键流程：

1. 新建项目；
2. 源码编辑；
3. 页面编辑；
4. 插入公式；
5. 插入图片；
6. 编译；
7. 修复错误；
8. 引用；
9. OCR；
10. 导出；
11. 关闭重开；
12. 崩溃恢复；
13. VS Code 联动；
14. TeXstudio 联动。

---

# 17. 性能验收指标

所有指标必须在固定参考机器和固定样例集上记录 P50/P95，不只记录最好结果。

建议基线：

| 项目 | 目标 |
|---|---|
| 普通按键到本地文本缓冲 | P95 小于一帧 |
| 文本变化到可视节点更新 | 常见局部编辑 P95 小于 50 ms |
| 页面 Overlay 本地响应 | P95 小于 50 ms |
| 200 KB 文档增量解析 | 常见局部修改 P95 小于 50 ms |
| 打开 1 MB 多文件项目 | 不阻塞 UI，逐步完成索引 |
| PDF 滚动 | 可见区域持续流畅 |
| 100 页文档 | 只渲染视口附近页，内存有上限 |
| 连续输入 | 不为每个按键启动独立 TeX 进程 |
| OCR Worker | 模型加载后复用，不重复初始化 |
| Core 崩溃恢复 | 不丢失已进入操作日志的编辑 |

编译耗时受模板、TeX 环境、图片和硬件影响，不使用单一绝对时间作为所有项目的承诺。必须记录：

- 冷启动编译；
- 热编译；
- 无引用变化；
- 引用变化；
- TikZ/大图项目；
- 任务取消延迟。

---

# 18. 数据安全验收

必须通过以下场景：

1. 编辑后未保存，应用被强制结束；
2. 保存过程中断电模拟；
3. 外部编辑器同时修改；
4. 磁盘空间不足；
5. 文件权限变化；
6. 编译器异常退出；
7. OCR Worker 异常退出；
8. PDFium 打开损坏文件；
9. 插件发送非法操作；
10. 旧 revision 重放；
11. 项目路径被移动；
12. 自动更新中断。

原则：

- 用户源文件永远优先；
- 写文件采用临时文件 + fsync + 原子替换；
- 恢复日志与源文件分离；
- 构建产物不能覆盖源文件；
- 不静默格式化整篇文档。

---

# 19. 最终验收场景

项目最终完成时，必须用以下完整流程验收：

1. 在断网环境启动应用；
2. 创建一个双栏期刊模板项目；
3. 在源码区输入标题和正文；
4. 可视化页面立即出现结构变化；
5. 点击 PDF 页面正文并直接修改；
6. 源码同步变化；
7. 插入行内公式和多行公式；
8. 从图片 OCR 公式并校对；
9. 插入图片，调整宽度和图注；
10. 创建表格；
11. 插入 label、ref 和 cite；
12. 使用 BibTeX/Biber 生成参考文献；
13. 故意制造编译错误并定位；
14. 修复后更新真实 PDF；
15. 切换单栏/双栏模板，页面由真实编译结果变化；
16. 关闭并重新打开项目；
17. 导出 PDF；
18. 使用普通命令行 TeX 编译同一项目；
19. 两者内容和排版一致；
20. 在 VS Code 中打开同一项目；
21. VS Code 源码与 VisualTeX 页面双向同步；
22. 在 TeXstudio 中修改文件；
23. VisualTeX 正确接收外部变化；
24. 全程不需要上传论文内容。

---

# 20. 完成定义

项目只有同时满足以下条件，才算达到最终目标：

- 可以从空项目写完一篇真实论文；
- 源码区和可视化区都能编辑；
- 双向同步是增量、实时且不会循环；
- 真实 PDF 是权威排版；
- 单栏、双栏和模板不由 CSS 猜测；
- 未知 LaTeX 结构无损保留；
- 公式和整页 OCR 本地运行；
- 可以独立编译和导出 PDF；
- VS Code 扩展可用；
- TeXstudio Bridge 可用；
- 三个平台可安装；
- 崩溃和外部修改不会静默丢数据；
- 核心服务与 UI 解耦；
- 测试覆盖解析、同步、编译、PDF、OCR、安全和端到端流程。

---

# 21. 开始开发时的第一批任务

正式编码的第一批 Issue 应严格限定为：

1. `ADR-0001：源码是唯一权威数据`
2. `ADR-0002：Rust Core + TypeScript UI`
3. `ADR-0003：Tauri 2 桌面容器`
4. `ADR-0004：PDF 权威页面 + Overlay 编辑`
5. `ADR-0005：Tree-sitter + 语义层`
6. `ADR-0006：JSON-RPC 协议`
7. 初始化 Cargo/pnpm Monorepo；
8. 建立 `vt-protocol`；
9. 建立 `vt-buffer`；
10. 实现 Revision/TextEdit；
11. CodeMirror 发送真实增量操作；
12. 属性测试：随机编辑与撤销；
13. 建立最小 LaTeX Corpus；
14. 建立首个三平台 CI；
15. 创建旧 VisualTeX 可借鉴代码清单，但不迁移旧 Store。

第一阶段不得提前开发：

- 漂亮的最终 UI；
- 整页 OCR；
- TeXstudio 内嵌；
- 任意 VS Code 插件兼容；
- 复杂表格；
- TikZ 可视化编辑；
- 云同步。

先保证权威文本、修订、解析和同步正确，否则后续所有功能都会重复返工。

---

# 22. 官方技术资料

- Tauri 2 Sidecar：<https://v2.tauri.app/develop/sidecar/>
- VS Code Custom Editor：<https://code.visualstudio.com/api/extension-guides/custom-editors>
- VS Code Extension Host：<https://code.visualstudio.com/api/advanced-topics/extension-host>
- Tree-sitter：<https://tree-sitter.github.io/tree-sitter/>
- ProseMirror：<https://prosemirror.net/docs/guide/>
- CodeMirror 6：<https://codemirror.net/docs/guide/>
- SyncTeX：<https://github.com/jlaurens/synctex>
- TeX Live：<https://www.tug.org/texlive/doc/texlive-en/texlive-en.html>
- Tectonic：<https://github.com/tectonic-typesetting/tectonic>
- PDFium：<https://github.com/chromium/pdfium>
- texlab：<https://github.com/latex-lsp/texlab>
- PaddleOCR PP-StructureV3：<https://www.paddleocr.ai/main/en/version3.x/pipeline_usage/PP-StructureV3.html>

---

## 23. 最终架构结论

本项目应从零开始，旧 VisualTeX 只作为公式交互、OCR Worker 和历史管理经验来源。最终架构固定为：

```text
标准 LaTeX 源文件作为唯一权威内容
                +
Rust 增量文本、解析、同步和编译核心
                +
TypeScript/React 编辑界面
                +
CodeMirror + ProseMirror + MathLive
                +
真实 TeX 编译 PDF
                +
PDFium 页面渲染
                +
SyncTeX 与影子插桩的位置映射
                +
PDF 页面编辑覆盖层
                +
Python/PaddleOCR 可选离线 Sidecar
                +
VS Code / TeXstudio / 通用协议适配
```

这一路线能够同时满足：

- 离线；
- 完整论文；
- 双向编辑；
- 实际模板；
- 单双栏；
- 公式、文字、图片混排；
- PDF 一致性；
- OCR；
- 独立编译；
- 外部编辑器接入；
- 后续插件扩展。

它也明确避开了两个最危险的错误方向：

1. 用 HTML/CSS 重新实现 TeX 排版；
2. 为兼容所有 VS Code 插件而在第一阶段重做一个完整 VS Code。
