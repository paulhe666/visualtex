# VisualTeX 自研数学可视化引擎：项目设计与逐阶段实施规范

版本：1.0 · 编写日期：2026-09-11  
适用对象：接手 VisualTeX 的开发 agent、维护者和验收人员  
状态：设计与实施计划；本文件不表示新内核已经实现或通过验收。

## 0. 必须先读：任务目标与不可突破的边界

用户要建设完全自主研发的数学可视化引擎，最终自己掌握公式表示、编辑事务、光标选区、数学排版与绘制。把 MathLive 包一层、长期维护其 fork，或者改用另一个现成公式编辑器，都不算完成目标。

同时，现有 Office 功能必须保持正常。内核改造的主要影响范围只能是主应用、公式编辑窗口及它们内部的渲染/导出适配。Word/PowerPoint 中已经稳定的插入、编辑、转换、重绘、导入、编号、引用、复制粘贴、文档持久化和性能不能成为这次重构的牺牲品。

执行原则：

1. **先证明接口与旧功能受保护，再接入新内核。** 不得先切默认引擎，再靠修改 Office 后端修补回归。
2. **新内核按最终架构建设。** 范围可以从小到大，每一个已支持结构必须贯通解析、编辑、排版、导出和测试，不建设以后要整块丢弃的演示内核。
3. **Office 契约默认冻结。** DTO 名称、字段、单位、状态语义、路由、COM ABI、OLE 存储、文档编号结构保持兼容。类型检查通过不等于行为兼容。
4. **主应用、交互编辑窗口、自动转换运行窗口分开接入。** 当前它们共享部分文件，但不是同一条业务链。
5. **保留旧文件读取能力。** 新内核内部模型不直接写进 Office 文档，不强制迁移旧文档，不在打开时全量重绘。
6. **最终真实 Office 验收必须经过真实 Ribbon/UI。** COM 可以准备测试文档、选择目标、只读检查结构；直接调用回调、反射或服务方法只能算辅助测试，不能替代 UI 验收。
7. **不得保存、关闭用户原有 Word/PowerPoint 文档，不得批量结束 Office 进程。** 只允许在可证明属于本轮测试的文档/进程上进行事先明确的测试操作。
8. **保留工作区历史 dirty/untracked。** 禁止 reset、clean、stash、切分支、删除未知文件、pull/rebase/merge；不得用 `git add .` 或 `git add -A` 混入历史修改。
9. 当前提交与推送许可只对应用户明确要求的本轮操作。后续 agent 的提交、推送、安装覆盖、上传安装包、发布 release，须按后续明确指令执行；本计划不是持续发布授权。
10. 本计划先在 Windows 工作区落地。核心设计跨平台，但未经许可不得修改 macOS/Web 工作区。未测的平台必须写“未验收”。

“Office 不受影响”是发布门槛，不能用“理论上接口没变”代替证据。遇到冲突时保留原功能，停止该接入阶段，不降低验收标准。

## 1. 当前真实基线与本轮 Git 操作

### 1.1 工作区与发布关系

- 工作区：`C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7`。
- 当前分支：`reference-3d207d7`。
- 本会话有效 Devspace workspace：`ws_0149a503d8`。后续优先复用；只有服务明确报告失效时，才重新打开同一实际目录。
- 前序本地提交：`2459c2573cff72c1a3e425fbd397c10df80d7d6c`，包含私有 MathType runtime、转换性能与安装器 guard。
- 本轮新本地提交：`c4690c2e69e266e6273eb0eaf5803ce74a054d1f`，`feat(windows): add standalone no-OCR installer flavor`，只包含本轮 9 个无 OCR 打包文件。
- 本轮推送后的远端 `main`：`6bb4fcee229db2f8cbc22330b54e3d4d0f5235f6`。
- 推送前远端 `main`：`861761744fa59ed17395e79765cba8bed280dcea`。

本地与远端存在不同历史。已验证远端的 Windows 基线与本地 `6bc5fdb` 相同，因此本轮使用独立临时 Git index，把 `6bc5fdb..c4690c2` 的 Windows 与相关文档增量放到远端原 main 之上，生成单父提交并普通快进推送。没有 force push，没有切换当前工作分支，没有覆盖远端 macOS 更新。发布提交相对原 main 涉及 37 个文件，含前序 runtime 修复及本轮无 OCR 修改。

**本地 HEAD 与远端 main 的 SHA 不同是有意的，不是未推送成功。** 后续不得因此强推当前分支到 main。每次接手都必须重新核对实际分支、HEAD、远端及变更范围。

### 1.2 安装包与二进制来源不能混为一谈

以下是在编写本文时只读核对的文件，不是新一轮完整 Office 验收结果。路径相对于 `apps/windows`。

| 对象 | 路径或来源 | 大小 / SHA256 |
|---|---|---|
| 保留的完整安装包 | `src-tauri/target/release/bundle/nsis/VisualTeX_1.2.6_x64-setup.exe` | 346823610 bytes；`6B97AE0B95089C435BBA1BA3A5C1754CE99594AA525BE472EC4EC6762E8C6A65` |
| 无 OCR 安装包 | 同目录 `VisualTeX_1.2.6_x64-no-ocr-setup.exe` | 119541342 bytes；`683EB5829B7A2A979E334F0187AF5B0A6EB6290B26A1F6BA3FD9387934DABAF4` |
| 当前资源目录 x64 MSI | `src-tauri/resources/windows-office/VisualTeX-WindowsOffice-VSTO-x64.msi` | 1945600 bytes；`C9A73E0333E63E3DDA8A5C51C985708114736B5343F78C2B68E2E9E86239A1A1` |
| 当前资源目录 x86 MSI | 同目录 `VisualTeX-WindowsOffice-VSTO-x86.msi` | 1925120 bytes；`3E8F7C944C055CD56D8F963C03A42B7680180FC096D511029CBB11B9B0FE7E9F` |

交接记录中的此前已验收 x64 MSI 为 `D6D6774C43D854BA0D3D4BCB46EF4B5C882AC1C523AB2872328C0A1393A99103`，x86 为 `DF30C97280300CB139496C3517C5F13EBB903AF96C8997B81455A169FBC87A6B`，Word DLL 为 `2B0BFC73BE468822113C94B080E2AC0F7F678A3795352F55189127879AFD76AB`。

**当前资源目录的 MSI 与前述历史验收 MSI 不同。** P00 必须建立“源码版本—构建输入—实际安装 DLL—最终包哈希—验收报告”的对应关系。不得把当前目录资源视为历史验收二进制；不得从报告中的同名安装包路径推断报告测试的就是现在这份字节。无 OCR 最终包与上轮 smoke 的对应关系也需要按哈希核对。

已知历史参考：`VisualTeX.WindowsOffice.Tests` 曾报告 503/503；100 个带编号公式 MathType→OMML 约 44.2s，VisualTeX→OMML 约 106.6s，两者 100/100、failed=0。这里引用用户交接记录，不宣称本轮重跑过这些测试。

本轮实际重新执行的是打包脚本 JavaScript 语法检查、相关 JSON 解析、无 OCR 安装包静态配置检查与文件哈希检查。静态检查不能单独证明包内所有资源、全部 Office 功能或完整 UI 行为正确。

### 1.3 必须保留的工作区状态

仍有历史未提交修改，例如 `src/office/dialog/OfficeDialogApp.tsx`、`src-tauri/src/office/lifecycle.rs`、`sessions.rs`、若干 CSS、测试与 UI 工具文件。还有历史 untracked 文件。它们未混入本轮无 OCR 提交，也不得在后续清理掉。

验收基线需要同时记录 HEAD 文件和实际工作副本差异。当前运行产品可能包含历史 dirty 产物，不能只用一个 Git SHA 代替完整来源记录。

## 2. 最终交付标准与非目标

### 2.1 必须自主拥有的部分

- 面向编辑与排版的数学表示树，以及保留源码细节的语法表示。
- 结构命令、位置映射、选择、导航、结构删除、撤销重做。
- 未完成命令、占位槽、输入法组合事务的数学编辑规则。
- 字体选择策略、数学布局算法、增量失效范围与布局缓存。
- 排版结果中的字形坐标、光标停靠点、命中区和辅助技术结构。
- 自有绘制指令及自身 SVG/PNG 等图形输出。
- 支持语法范围内的 LaTeX 解析/序列化和 Presentation MathML 编码。

可复用标准工具、字体解析与文本塑形库、浏览器/操作系统输入法、SVG 绘制能力、CodeMirror、React、Tauri。这些基础设施不代替公式编辑和数学排版规则。

最终新引擎及正式可视编辑、预览和自身导出不能依靠 MathLive 或 MathJax 运行时兜底。测试工具可保留旧版本作对照，但不得进入正式运行依赖图。OMML 的 Word 排版和 MathType 的原生运行时属于目标格式宿主，不属于必须重写的 VisualTeX 内核。

### 2.2 本项目不做什么

不重写 Word/PowerPoint 插件；不改 OLE 格式、COM 身份、SEQ/REF 编号架构；不重新尝试“正式 MathType 优先于私有 runtime”的已撤回实验；不顺便重做 OCR、更新器、安装器或云同步；不把主应用整体迁移到 Qt；不承诺任意 TeX 程序均能结构化编辑；不在本轮解决跨软件字体绝对一致/WPS 全兼容问题。

新引擎能够统一自己的显示与图形导出，但导出为可编辑 OMML/MTEF 后仍由 Word/MathType 自行处理。不得宣称这些外部排版结果天然与 VisualTeX 逐像素一致。

## 3. 架构决策及理由

### 3.1 采用：Rust 共享核心 + 同进程 WASM 交互绑定 + 自有 SVG View

核心采用模块化单体设计，不依赖 Tauri、React、DOM、COM 或系统文件路径。它在桌面/Web 编辑区域编译为 WebAssembly，在测试和无界面工具中编译为原生库/程序。同一份代码负责结构、事务和数学几何。

React 保留工具栏、设置、面板和文档视口；一个独立 View 按核心给出的坐标执行绘制。SVG 路径是正式显示后端，不必把它视为以后必须淘汰的过渡实现。Canvas/原生 View 可以扩展，但没有证据前不增加第二套正式绘制路径。

这是内核自主、可原生运行、界面保留混合外壳的方案。它不等于整个应用是纯原生 GUI。Tauri 本身使用 Rust 与 WebView，见 [R03]。

### 3.2 为什么不选另外几条路线

| 方案 | 判断 |
|---|---|
| 纯 TypeScript 自研核心 | 技术上完全可行，不因语言而不专业；但本项目已有 Rust 字体基础，且需要原生无界面复用，故不作为首选。不能无测量宣称 Rust 必然更快。 |
| C++/Qt 全应用重建 | 扩大到 UI 与宿主迁移，远超内核替换范围。C++ 可以跨平台，排除理由是本项目的迁移成本，不是“C++ 不能做 Web”。 |
| MathLive fork 或永久 adapter | 可用于迁移边界，不能成为最终内核。其结构与布局仍受原实现控制。 |
| 用 MathJax/MathML 渲染器加一层光标 | 渲染树不等于可编辑、保真、可撤销的文档。只能作参考或旧输出链，不能充当终态。 |
| 把每次按键发给 Tauri/Rust 后端 | 在高频交互上引入 IPC 排队和状态同步，缺乏必要性。采用编辑 View 内的 WASM 调用。 |
| 一开始多进程、插件化、GPU 重写 | 先增加同步和资源风险，并不直接解决结构编辑。核心模块可扩展，不追求无限插件能力。 |

原生核心可用于未来批量导出，但本项目不要求改现有 Office 通信路径，不向 WINWORD 进程加载一个新 Rust DLL。先让现有消费者继续收到相同协议的数据。

### 3.3 AxMath 的证据边界

公开操作文档 [R08] 可确认：可视编辑与 LaTeX 输入有区分；Mixed LaTeX 保持临时命令输入状态；复制支持复合数据与 LaTeX 等不同用途。借鉴这些可观察的交互，不复制其资源、代码或格式，不把未核实的 MFC、内部类树、排版器实现作为设计事实。

AxMath 对照应使用新建的无敏感测试公式，逐项记录真实操作。若未安装或无法操作，只记录文档证据，不编造体验测试。优先保持 VisualTeX 已验收的用户习惯；参考产品行为发生冲突时，必须形成明确产品决策。

## 4. 代码边界与建议目录

以下目录是待新增设计，不表示现有仓库已提供这些模块。

```text
engine/                           独立 workspace，不改 Tauri 根构建
  Cargo.toml
  Cargo.lock
  crates/
    core/                         单一核心 crate，内部按职责分模块
      src/model/                  数学表示、身份、源码细节
      src/syntax/                 lexer、parser、serializer、宏范围
      src/edit/                   事务、选区、导航、删除、历史
      src/fonts/                  字体资源、塑形、MATH 数据
      src/layout/                 盒式排版、增量失效、对齐约束
      src/scene/                  绘制指令、光标区、命中区
      src/export/                 SVG、MathML、安全检查
    wasm/                         wasm-bindgen 边界，不放业务规则
    test-cli/                     无界面验证、基准与场景回放
  web/
    host/                         View、输入法桥、资源加载、辅助技术
    sandbox/                      与正式主应用/Office 隔离的验收入口
  fixtures/                       公式、事件、语义和几何样本
  specs/                          支持范围、输入行为、数值约定

apps/windows/src/editor/engine/   主应用的新内核接入层
apps/windows/src/office/compat/   Office 输出兼容适配，不改 wire DTO

docs/engine-rebuild/              本计划、阶段报告、ADR、基线清单
```

先验证目录不与现有文件冲突。不要为了共享核心把三个平台应用搬家，也不要直接复制三份内核。核心保持统一来源，平台接入分别授权、分别验收。

建议只设 core、wasm、test-cli 三个 Rust crate。不要每个数学节点建一个 crate。为核心固定依赖和工具链，但避免根目录 rust-toolchain 全局影响原有 Tauri/Office 构建。

### 4.1 单向依赖

```text
主应用 / Office 编辑窗口 / CodeMirror / OCR 结果 / 工具栏
                       ↓ 命令与输入
                 EngineSession
        文档 + 源码草稿 + 事务 + 选区 + 版本
                       ↓
              FontProvider / Shaper
                       ↓
                  MathLayout
                       ↓
        同一 revision 的 LayoutSnapshot
           ├─ DisplayList → SVG View / 图形导出
           ├─ CaretStops / HitRegions
           └─ AccessibilityProjection

OfficeCompatibilityAdapter → 原 OfficeExportResult / Session 更新入口
```

核心不回调 Word、不抓取 DOM、不读应用 store。宿主只传递容器约束、输入事件和经过验证的资源。View 不得通过测量分子 DOM 决定分母位置。

## 5. Office 保护区：文件、职责和准入规则

以下路径相对于 `apps/windows`。保护是对进入阶段时的实际工作副本做增量比较，不是把历史 dirty 强行还原为 HEAD。

| 等级 | 范围 | 规则 |
|---|---|---|
| 冻结 | `src-windows/VisualTeX.WordVsto/**`、`VisualTeX.PowerPointVsto/**` | 不修改插入/替换、编号、复制粘贴、性能查找和 Ribbon 业务 |
| 冻结 | `VisualTeX.WindowsOffice.Contracts/**`、`VisualTeX.FormulaOleServer/**`、`VisualTeX.WindowsOffice.VstoShared/**`、`VisualTeX.WindowsOleBridge/**` | 不改 COM/存储/消费者协议，不为适配新 SVG 放宽原安全校验 |
| 冻结 | `src-tauri/src/office/{sessions,server,lifecycle,state,windows_backend,windows_pipe}.rs` 及平台桥 | 不改 API、状态转换、会话寿命、启动/关闭流程、路由鉴权 |
| 冻结 | `src/office/{shared,api}/` 中 Session DTO、协议、元数据、持久化及校验模块 | 不改字段、编码、超时语义、保存队列和提交规则 |
| 冻结 | `src/office/dialog/useOfficeSession.ts` 与既有提交/取消/关闭机制 | 不另起一套 Session 保存机制 |
| 冻结 | Office MSI、私有 MathType runtime、注册表、安装器 hooks | 不因内核开发重新生成或覆盖已验收资源 |
| 受控改动 | `OfficeDialogApp.tsx` 的编辑表面接入与导出提供者注入 | UI 与导出变化分开提交；不得顺手改 Session effects、提交排队或自动转换批处理 |
| 受控改动 | `src/export/*`、`MathPreview`、CodeMirror 桥、Zustand 公式状态桥、Vite 资源配置 | 外部输入输出保持契约，独立证据覆盖全部调用点 |
| 新增 | `engine/**`、新 View/适配器、测试和设计文档 | 按阶段开放，不得主动接管 Office |

静态资源加载如果确实需要修复 `.wasm` MIME 或映射，应先证明问题并单独申请最小例外；不因此开放 Office API 重构。系统字体若需新增宿主读取接口，也必须放在独立字体资源边界，不扩展 Session/COM 协议。

保护区已有问题应单列原缺陷，不得悄悄混入本项目修改。新内核触发的回归优先修新内核或兼容层。

## 6. 必须保持的实际 Office 契约

### 6.1 Session 与元数据

实现依据：`src/office/shared/sessionClient.ts`、`formulaMetadata.ts`、`src-tauri/src/office/sessions.rs`、C# Contracts。

| 类别 | 必须保持 |
|---|---|
| host / mode | `word`、`powerpoint`；`create`、`edit` |
| objectMode | `nativeOle`、`mathTypeOle`、`wordOmml`、`crossPlatformPicture` |
| 显示模式 | Office 使用 `inline` / `block`；主应用行使用 `inline` / `display`，适配器显式映射 |
| Session 身份 | `id`、`formulaId`、`sourceDocumentId`、`sourceObjectId` 不被内部 NodeId 替代 |
| 公式数据 | `lines[{id,latex}]`、`activeLineId`、`codeFormat`、`title` 的含义不变 |
| 编号 | `numbered`、`mathTypeNumberPosition`、`equationTag` 仍由原业务语义控制 |
| 状态 | `created/editing/committing/completed/cancelled/failed`；允许的转换以现有代码为准，不重定义 |
| 保存与取消 | `dirty`、`autoCommitOnClose`、`explicitCancel`、error、时间与过期机制保持原规则 |
| 元数据 | `schema="visualtex-formula"`、`schemaVersion=1`、前缀 `visualtex:v1:deflate:`、XML namespace `urn:visualtex:formula:1` 不改 |
| 尺寸持久化 | `fontSizePt`、`renderFontSizePt`、`renderWidthPx/HeightPx`、`baseline`、`wordInlineOleWidthPt/HeightPt` 不混用 |
| 原生来源 | `nativeOmmlFingerprint` 和 originalMetadata 保留；不以自研结构哈希替代原 OMML 指纹 |

旧 schema 不具备的引擎版本字段不能偷偷追加。自研文档/缓存可独立版本化，但 Office v1 元数据继续按原协议输出。不得声称旧 Office 文档已经保存了本来不存在的布局版本。

### 6.2 输出数据保持原名、原单位

当前实际输出接口：

```ts
interface OfficeExportResult {
  svg: string;
  svgBase64: string;
  mathMl?: string;
  pngBase64?: string;
  width: number;
  height: number;
  baseline?: number;
  formulaLetterFont?: FormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
}
```

这不是新接口草案，不得改成 `mathML`、把 base64 data URI 塞入原纯 base64 字段，或把 width/height 改为物理屏幕像素。

尺寸约定：自然输出采用 96 dpi 的 CSS px；Office 磅是 72 dpi，`pt = px × 0.75`。`baseline` 从 SVG 顶边向下计，单位也是自然 CSS px。屏幕 zoom、devicePixelRatio、PNG 二倍采样都不改变语义字号或自然宽高。SVG viewBox、padding、导出宽高和 baseline 必须来自同一个布局快照。依据：`OfficeFormulaSizing.cs`、`src/export/exportTypes.ts`。

`SvgExportResult.base64` 到 `OfficeExportResult.svgBase64` 是现有边界映射，不要混淆。

### 6.3 路由、鉴权与 COM ABI

不得改变 `/api/v1/sessions` 的 POST、`/api/v1/sessions/{id}` 的 GET/PATCH/DELETE，Windows commit、应用 session close、converter next-batch、PowerPoint commit/confirm 的现有路由。保留 `X-VisualTeX-Install-Token`、同源凭据及现有响应解析。

代码中的完整路径与客户端函数必须纳入 P01 机器可检验快照；不要只锁本文列出的例子。

`FormulaOleContract.cs` 当前 `ProtocolVersion=2`、`StorageSchemaVersion=1`；这是两个不同版本号。`VisualTeX.Formula.1`、现有 CLSID/IID/AppID/TypeLib、方法 DispId 均冻结。

OLE 流名保持 `VisualTeX.Formula.json`、`VisualTeX.Preview.emf`、`VisualTeX.Preview.png`。MathType 继续保留真正的 Equation Native/MTEF，不用图片伪装可编辑 MathType 对象。

### 6.4 四个不能漏掉的隐藏依赖

**依赖 A：自动转换与交互编辑共享 OfficeDialogApp。** 文件中 `IS_VSTO_CONVERT_RUNTIME` / `vsto-convert` 有独立批处理行为。替换可视编辑区不得令隐藏转换窗口创建交互内核、抢焦点或初始化输入法。P12 和 P13 分别开关、分别验收。

**依赖 B：OMML 批转换快路径。** `generateSessionExportResult()` 对 `wordOmml` 可以只生成 MathML，返回空 SVG 字段并沿用尺寸元数据。不得为了接口统一强制栅格化所有公式，更不能让它逐个打开编辑窗口。

**依赖 C：MathType 对齐几何。** `mathTypeAlignmentGeometry.ts` 当前从 MathJax SVG 表格几何给 MathML 写入 `data-visualtex-mtef-ruler-stops`。换成自有 SVG 后不能假定仍有 `data-mml-node`/`translate` 的原组织。最终适配器从自有布局结果读取对齐列锚点，输出同名属性、相同单位：CSS px 到 MTEF 1/32pt 单位为乘 24，保持范围与严格递增检查。用单/多组 aligned、嵌套表格和不同字号验收，不能默默丢标记后宣称接口没变。

**依赖 D：保存和宿主完成不是同一事件。** `useOfficeSession.ts` 有 PATCH 排队；后台 Session 防止旧 autosave 覆盖 committing/completed。新内核的 revision/异步渲染结果必须先经过原更新入口，不能绕过队列、提前标 completed 或把老窗口结果提交给新会话。

此外，MathType 输出的 Times 字体策略、原生 Full Size、私有 runtime 路径继续保留；公式编号在 Word 段落/域层，不能混入新 AST 当数学内容。

## 7. 内核的数据、状态与编辑模型

### 7.1 一个 EngineSession 是一个权威写入点

编辑状态包含：数学表示、源码细节、源码草稿、选区、输入法状态、历史、revision。所有输入都成为事务；任何视图或 store 不得直接改内部树。

主模型是**面向排版与编辑的数学表示**，不是会自动化简的计算 AST。`a/b` 与分式、显式括号、手工间距、上下标放置、颜色和作者宏调用必须可区分。计算语义可以派生，不能覆盖作者表达。

可视编辑与源码编辑共用同一事务时间线。源码未完成时保留真实草稿和诊断，显示最后有效结构或标记的容错预览；禁止拿自动补全的预览反写用户源码。修复草稿后由同一个事务更新结构和位置映射。

Zustand 最终保存应用配置和受控投影，不与 core 各维护一份可独立编辑的公式真源。

### 7.2 分开文档、结构、显示与宿主身份

- FormulaDocumentId / FormulaId：公式文档/现有宿主公式身份。
- NodeId：内核可编辑节点身份。
- LayoutFragmentId：同一个节点的一段显示实例身份。
- LineId：现有外层逻辑行身份。
- SessionId：Office 会话身份。

这些 ID 不互相借用。复制子树生成新的 NodeId；撤销恢复被删除节点的身份；重新排版不重建编辑身份。宏展开或一字多 glyph 时，多个显示片段可映射回同一源码/编辑节点，不伪造一一对应。

### 7.3 行与容器

主应用的逻辑行、公式内部 Row、矩阵行、AlignedGroup 行、物理源码换行、布局软换行都必须区分。首轮不增加自动公式换行功能；已有多行与对齐功能不能因此丢失。

核心结构至少覆盖 Row、Symbol、Number、TextRun、Fraction、Radical、Scripts、Fence、Accent、LargeOperator、Matrix、Cases、AlignedGroup、Style、Space/Phantom、CustomSymbol、MacroCall、EmptySlot、UnparsedSource。具体数据类型在 P01/P03 固定，不在本计划里伪装成已经可直接编译的完整 API。

### 7.4 源码保真与支持范围

保留 token、注释、控制词拼写、参数边界和源范围。有效且未编辑的支持范围源码必须能够保真往返；只在明确修改的范围进行必要序列化。规范化导出与保真保存分开。

例如 `\pdv{f}{x}` 可展开成分式显示，但保留宏调用身份。改参数优先输出 `\pdv{g}{x}`；需要破坏宏结构的操作必须按定义展开并能撤销。未知宏保存原文与诊断，不执行任意代码，不算完整可视化支持。

每项能力分别登记 parse/edit/layout/LaTeX/SVG/MathML/Office-target 状态。源码保留、能画占位框、MathJax 兜底都不能冒充完整支持。

### 7.5 位置、选区与事务

位置表示容器内边界或 TextRun 内 grapheme 边界，并包含插入偏向/前后依附关系。需支持连续选区、整个结构选择和矩阵矩形选区。

事务返回新 revision、逆操作/历史记录、位置映射、失效范围、诊断及派生事件。节点删除、源码重解析、宏展开、复制粘贴、跨行合并都必须映射旧位置。仅有“稳定 NodeId”仍然不足以解决这些问题。

浏览器 UTF-16、Rust UTF-8 字节和用户感知字符索引显式区分。不得通过 LaTeX 字符数推算数学光标位置；Unicode 字符边界参考 [R04]。

撤销记录操作/结构变化及选区，不依靠前后 LaTeX 字符串差异猜操作。CodeMirror 的文本事务通过版本化桥接进入同一历史；不能让两套 undo 都响应同一个 Ctrl+Z。源码草稿也必须可撤销。

### 7.6 输入法与异步结果

CompositionSession 保存开始 revision、目标锚点、预编辑文本和原选区；确认只生成一次事务，取消不改原公式。Enter 先按当前组合状态处理，不能同时确认候选并分行。真实输入法测试从最小闭环开始，不延后为附加功能。

OCR、字体加载、导出和 worker 结果必须携带文档/会话身份和 revision。目标已变化时映射、明确拒绝或提示重新定位；不得落到“当前活动公式”。WASM trap/OOM 后宿主保留最后检查点并可重建实例，不能假设 `catch_unwind` 能捕获所有 WASM 故障。

## 8. 字体、数学布局与绘制设计

### 8.1 字体是布局输入，不是最后换皮

先确定具体字体资源、face index、版本/哈希、variation axes 和数学样式，再获取字形度量并排版。字体缓存键包含这些身份，不只包含 family name。数学字体与中文/普通文本字体分开管理。

当前 `system_math_glyphs.rs` 的 ttf-parser 轮廓读取可参考，但它主要是单字形/Windows 资源查询，不能直接充当完整跨平台 FontProvider。新核心先使用经过许可审查的固定字体资源，经宿主一次加载；不要按每个 glyph 做 Tauri IPC。

核心可复用 ttf-parser 的字体表/轮廓访问和 rustybuzz 等文本塑形基础库 [R05][R06]。数学 MATH 常量、可伸展构件、斜体修正、重音与二维放置由 VisualTeX 自己决定。字体文件不是算法，[R01] 也不提供一个可直接替代自研的完整排版实现。

缺字体不得静默冒充选定字体。产品 fallback 策略需显示可诊断信息，并把实际资源纳入布局键。禁止从用户系统随意打包字体文件；依赖与字体许可分别核验，不向报告附带系统字体二进制。

### 8.2 基础排版约定

每个盒至少有 advance、ink bounds、ascent/descent、数学轴线、子盒变换和布局版本。墨迹边界与前进宽度不同；斜体边缘、重音和深下标不能因选择框宽高而被裁掉。

实现 display/text/script/scriptscript 与必要的 cramped 状态；统一数学符号类别和间距；分别实现分式、根式、定界符拼接、上下标、上下限、重音、矩阵和对齐约束。参考 [R01][R02]，不得用一组经验百分比代替全部 MATH 规则。

内部数值约定在 ADR 中固定：推荐以 1em=1 的 f64 几何运算为基础，字体设计单位只在入口换算，确定性遍历，最终输出边界统一量化；不做每一步整数四舍五入。Native/WASM 用固定数据和容差比较几何，不要求不同平台抗锯齿逐像素相同。

增量失效沿编辑节点父链与相关约束传播。矩阵一格可能使同列宽度变化，对齐组可能更新整个组；不能承诺所有操作 O(1)，但不能扫描无关文档全部公式。

### 8.3 一份 LayoutSnapshot 服务所有视觉能力

DisplayList、光标停靠点、命中区域、可访问结构属于同一 revision。DOM 的职责是执行绘制与上报视口，不反向推断公式结构。

View 只把 CSS 视口坐标转换到布局坐标。缩放/滚动不改变文档内容。光标闪烁只更新叠加层；移动选区不重新 parse 或全量 typeset。排版未就绪时暂缓依赖几何的输入，不使用旧版本命中区。

绘制指令可以包含 GlyphRun、Path、Rule、Transform、Paint 与 Clip；Office SVG 输出必须降低为当前转换器可接受的安全子集。编辑屏幕可以有选择背景与槽位框，导出不包含这些交互叠加层。

### 8.4 Office SVG 兼容子集

当前 `OfficeOlePreview.cs` 支持部分 svg/g/use/path/rect/line/polyline/polygon/circle/ellipse/text，并有严格校验。新导出优先输出路径及基础图元；`defs/use` 仅引用文件内部，必要时展开。

不得引入外链字体、远程资源、script、foreignObject、未经现有消费者支持的 filter/mask/透明混合，不能把 PNG 塞进 EMF 冒充向量。复杂图形效果要降低为明确路径；不能降低时阻止该目标格式导出并保留原文档，不修改消费者放松安全约束。

不要直接把“对所有浏览器有效的 SVG”当成“对当前 Office 转换器有效的 SVG”。须用原转换器的真实测试验证。详见 P11。

### 8.5 自主显示与历史兼容

新内核的标准排版配置与 `office-v1` 兼容配置分开。它们共用算法，兼容配置只表达经证据确认的字号、边距、尺寸和字体策略，不能成为每个公式一条特例表。

未改动的旧公式不开启自动重绘；能确定缓存与内容对应时按现有规则复用缓存。新算法会造成视觉变化时必须记录差异并经用户验收，不能假装只换代码就不会改变布局。新旧比较先分清“bug 修复”“审美变化”“文档回归”，不自动刷新 golden 消除失败。

## 9. 实施总顺序与准入机制

全部阶段初始状态均为 `NOT_STARTED`。编写本文不代表执行了任一新内核阶段。

```text
P00 基线保护 → P01 契约/ADR/样本 → P02 独立核心与双目标构建
  → P03 最小保真表示 → P04 事务与位置 → P05 字体/布局/绘制
  → P06 真实输入最小闭环
  → P07 数学结构扩展 → P08 矩阵/多行/对齐
  → P09 源码/宏/剪贴板/自定义字符完整接入
  → P10 主应用接入
  → P11 Office 输出兼容证明 → P12 Office 交互编辑窗口接入
  → P13 自动转换链接入
  → P14 跨宿主与故障/性能加固
  → P15 删除第三方公式引擎运行依赖
  → P16 打包与最终真实 Office 验收
```

P03—P06 是第一个贯通的自主内核切片。后续每个新结构都按“表示—编辑—布局—导出—测试”纵向推进，不先堆一个巨大解析器再补光标。

每阶段报告必须写：开始版本、允许改动路径、实际改动、测试命令和环境、结果、未覆盖项、保护区差异、下一阶段准入结论。失败用 `BLOCKED`，没有执行用 `NOT_RUN`，不得把二者记为成功。

### P00：冻结真实基线与保护用户环境

**前置：** 阅读本文和当前工作区指令；仅进行保护性检查及批准的测试准备。

**实施：**

1. 核对目录、分支、HEAD、远端 main；记录 tracked dirty、untracked 和 index 状态，不清理。
2. 给保护区和受控区生成来源清单：Git blob、工作副本哈希、编译输入、实际 DLL/EXE/MSI 与前端资源哈希。历史 dirty 单独列出。
3. 区分三个来源：此前已验收完整包、当前 no-OCR 包、当前资源目录。核对各测试报告实际测试的哈希；不匹配的报告只作历史参考。
4. 只读记录正在运行的 Word/PowerPoint/VisualTeX PID、应用版本、文档身份与保存状态。不得以验证为由调用 Save、Close、Quit。
5. 建立专用测试文档命名和身份清单。测试文档的创建、写入、保存/重开须限定于本轮拥有的对象；普通用户文档永远不在自动关闭列表。
6. 使用实际认可的产品版本记录主应用、编辑窗口、Office 操作和性能基线。当前版本已有失败项单独记录，禁止重新定义为正常。
7. 记录标准版/no-OCR 的文件来源、Office 子包来源、私有 runtime 路径、Word 插件实际加载路径。

**验收与产物：** `BASELINE.md`、机器可读 provenance manifest、保护区路径表、原始测试结果和测试文档身份表。P00 尚未建立可靠二进制对应关系时，不允许开始 Office 接入。

**停止条件：** 无法辨认文档所有权、资源与报告混用、出现未解释的新 dirty、需要关闭用户 Office 才能继续。不得自动“修复”这些状态。

### P01：固定契约、设计决策和可执行测试标准

**前置：** P00 基线有效。允许新增规范、测试和最小观测设施，不替换生产引擎。

**实施：**

1. 写 ADR：核心运行位置、单写者文档、源码草稿/保真规则、ID/选区、字体与数值约定、Office 兼容边界。
2. 从实际 TS/Rust/C# 实现生成 Office 契约快照；固定字段名、默认值、optional/null/空字符串语义、大小写、单位和错误行为。
3. 为 Office 导出建立输入与输出 fixture。只归一化明确无业务意义的时间/临时路径；不得把尺寸、baseline、编号、源对象 ID、指纹等一并抹掉。
4. 建立受保护文件增量检查。白名单为空为默认；受控区域的例外必须关联阶段和证据，不能只靠文件名放行整文件。
5. 建立语言能力矩阵和现有交互行为清单：核心符号、全部当前支持模板、宏、中文、上下标跳出、cases Enter、多行 split/merge、复制和撤销。
6. 固定 screen→layout 坐标换算、序列化模式、字体标识和错误类型。新内核内部 API 必须与 Office wire API 区分。
7. 规定“旧行为参考”不等于“旧 bug 永远正确”；改变已验收行为必须有独立决定，不靠更新测试期望掩盖。

**验收与产物：** 契约测试先在旧实现上通过；修改一个测试副本中的尺寸单位或字段名时测试必须失败。形成 ADR、接口快照、行为规格、能力清单和测试样本清单。

**停止条件：** 只做 TypeScript 类型检查，没有协议/行为测试；保护区变更无法解释。

### P02：建立独立 Rust 核心、WASM 绑定和原生验证入口

**前置：** P01 接口草案固定。生产应用继续使用原引擎。

**实施：**

1. 建立 `engine/` 独立 Cargo workspace，core 不依赖 Tauri、React、Word、文件系统或网络。
2. 建立 wasm 与 test-cli 薄边界，生成或校验 TypeScript 类型。绑定版本独立于 Office ProtocolVersion，不改后者。
3. 明确宿主注入：字体字节、时间/随机身份需求、日志回调、资源限制。内部节点可用会话命名空间与单调计数，不把系统随机接口散入核心。
4. 在原生与 WASM 目标运行同一纯数据测试，检查确定性排序、数值、错误和反序列化行为。
5. 按 handle + revision 操作 EngineSession，禁止向 JS 暴露可长期持有且随内存增长失效的裸指针。批量返回变化，不为每个 glyph 做一次 JS/WASM 调用。
6. 固定构建工具链和依赖，使用独立 target。错误通过 Result/结构化诊断返回；不依赖 WASM 上并不等价的操作系统服务 [R07]。
7. 建立空会话创建/销毁/重复创建、资源释放和错误恢复测试。

**验收与产物：** Native 与 WASM 的基础协议/数据测试一致；无 MathLive/MathJax/Tauri 运行依赖；生产构建与 Office 文件没有改变。

**停止条件：** 需要改全局 Cargo/Node 设置才能构建、要求把核心放进 WINWORD、绑定初始化拖慢现有应用启动。先修独立构建边界。

### P03：最小保真数学表示与解析/序列化

**前置：** P02 双目标基础通过。

**实施：**

1. 最小支持集 K0：Row、普通符号/数字、TextRun、Fraction、上下标、EmptySlot 和未解析源码。
2. 同时建立 token/源码范围和数学表示，保留宏名及明确分组；明确源范围的 UTF-8/UTF-16 转换。
3. 增量解析先保证正确性：优先重新解析最小可独立子树，必要时扩大到公式；不得为了“增量”制造不正确上下文。
4. 建立保真序列化与规范化导出两种入口。未改部分保留原拼写和可保留的注释/空白；格式美化只在用户显式请求时执行。
5. 支持未完成分母、未配对括号/控制词的编辑诊断，不把自动修复预览写回源码。
6. 设定宏展开、嵌套、节点、文本和资源预算；任何失败保留原输入，不执行 TeX 外部命令。
7. 用重复子式和局部修改测试身份保留。不能用整段文本哈希作为 NodeId，否则相同内容和修改后身份会混乱。

**验收与产物：** K0 的 parse→serialize 保真样本、无关范围不变样本、错误源码不丢失样本、Native/WASM 等价样本。输出语言支持矩阵 K0，不报告全 LaTeX 支持。

**停止条件：** 仍调用 MathLive/MathJax 完成核心解析；通过简单字符串切片假装嵌套结构解析；未知命令被吞掉。

### P04：结构事务、光标位置、删除与撤销内核

**前置：** P03 K0 数据与序列化稳定。

**实施：**

1. 定义 InsertText、InsertStructure、ReplaceRange、DeleteBackward/Forward、Move/ExtendSelection、SetStyle、Split/Merge 等命令与原子性规则。
2. 定义容器边界、文本 grapheme 位置、方向偏向，以及连续/节点/单元格选区。
3. 每次事务产生 PositionMap；节点删除、拆分、合并、复制和 undo/redo 后的锚点都须有效。
4. 左右导航先由结构规则决定，上下导航保留目标横坐标请求供布局解决，不写成线性整数加减。
5. 结构删除按产品规格执行，一次真实按键只对应一个决定；禁止“没变化就再调用一次删除”的试探循环。
6. 用事务操作与逆操作管理历史，支持输入分组、粘贴整体撤销、结构展开撤销。移动光标不错误增加内容历史。
7. 建立 CompositionSession 的纯状态模型和异步结果锚点检查，为 P06 真实 IME 接入做准备。

**验收与产物：** 操作回放和性质测试：撤销恢复内容+选区+身份；重复复制节点身份不同；事件重放无重复提交；每个位置引用有效。小随机测试保留 seed 和最小失败案例。

**停止条件：** 新选区仍然存 MathLive offset；历史通过 LaTeX 长度猜光标；把相同内容当成相同对象。

### P05：字体服务、K0 数学排版与自有绘制结果

**前置：** P03/P04 核心通过。此阶段仍不挂载正式 Office 编辑器。

**实施：**

1. 选择并锁定测试数学字体与中文/文本字体资源；记录许可与哈希。只把可用字节传入 core，文件发现属于宿主。
2. 建立 glyph 选择、文本塑形、MATH 常量访问和 fallback 诊断；复用基础库但自己实现二维数学布局。
3. K0 支持数学轴线、不同 math style、分式、上下标与文本基线。区分 advance、ink bounds、logical bounds。
4. 生成 DisplayList、CaretStops、HitRegions 和基础辅助技术投影，不使用浏览器数学排版获取公式几何。
5. 实现纯 SVG 路径输出和无界面几何 JSON；用同一快照绘制与导出。任何 UI 占位框不进入普通发布导出。
6. 实现坐标单位检查、非有限数值拒绝、盒边界检查与深层结构预算。
7. Native/WASM 比较字形 ID、位置、baseline 和盒几何，容差在 P01 固定；跨平台字体资源不同的结果不能混入该等价测试。

**验收与产物：** K0 图像与结构几何黄金样本、固定字体下的双目标一致性、无 DOM/MathLive/MathJax 排版依赖。

**停止条件：** 排完再替换字体；只靠浏览器 getBoundingClientRect 反算公式；把错误盒高用固定像素补齐。

### P06：第一条真正自主的交互闭环

**前置：** P02—P05 通过。仅在独立 sandbox 启用，不改变正式默认引擎。

**实施：**

1. 建立新 SVG View 和真实键盘/指针输入桥；React 不管理每个 glyph 的独立业务状态。
2. 实现点击定位、拖选、左右/上下导航、模板插入、删除和 undo/redo，使用 core 的几何和事务。
3. 接入真实中文 IME。文本输入宿主不能使用 `display:none` 破坏输入法；候选窗口根据当前光标矩形定位。EditContext 仅作可选平台通路 [R09]，不可假设所有宿主可用。
4. 输入 `\frac{x_1+\text{速度}}{y^2}`，通过真实输入逐项修改分子/分母/上下标/中文，导出 SVG，再重新加载源码。
5. 测试未确认中文候选时 Backspace/Escape、候选 Enter、输入反斜杠、选区替换与输入法切换。
6. 改变 zoom、DPI、窗口尺寸与滚动后重复命中和候选定位测试。
7. 加入无障碍结构与位置描述的最小验证；不能只有“这是一条公式”的 aria-label。

**验收与产物：** 可运行的原生/WASM 自主最小内核、真实事件回放、截图、几何与导出一致性、IME 录像/日志。此闭环中所有数学编辑与布局均由 core 完成。

**停止条件：** 只演示预填公式，未真正输入；让 MathLive 接管光标/IME；用旧渲染器代画未实现的分式后报告自主闭环完成。

### P07：按结构纵向扩展数学能力

**前置：** P06 输入闭环稳定。

**实施顺序：**

1. 根式与带指标根式：数据、槽位、伸展构件、导航、序列化同时完成。
2. 成对/单侧/隐形定界符、嵌套括号：明确自动配对和跳出规则，保证未知右界不被误删。
3. 重音与上下标注：区分重音字符、可伸展重音、over/under、装饰和占位槽，覆盖嵌套自动跳出。
4. 大运算符：符号类别、limits/nolimits、display/inline、下限→上限→正文的产品导航；保留用户配置而非硬编码某一家编辑器习惯。
5. 字体样式、数学字母、正体函数、空白/phantom、颜色与边框；数学语义和显示样式分离。
6. 希腊字母、罕见积分和现有兼容命令：复用现有合法 glyph 数据，但重建新引擎自己的注册与布局。
7. 每类结构都扩充行为、源往返、屏幕/导出、边界与随机测试。已有 K0 测试不得回退。

**验收与产物：** 每类结构独立能力记录和通过证据。没有编辑/删除/序列化支持的“能显示结构”不得记为完成。

**停止条件：** 采用每个公式一个特例、按截图数值微调整个规则、为达到覆盖率吞并未支持命令。

### P08：矩阵、cases、多行文档与对齐

**前置：** P07 基础数学结构通过。

**实施：**

1. 建立 Matrix/Cases/AlignedGroup 与外层文档行的区别，明确 cell/row/column ID。
2. 实现增删行列、单元格移动、矩形选区、单元格粘贴与完整撤销；空格子也是有效编辑位置。
3. 对齐组共同计算列宽和锚点，支持现有 `&`、多组配对及嵌套环境；不用多份独立 Mathfield 加 CSS margin 对齐。
4. 保留 cases 内 Enter 新增行与外层 Enter 分行的区别；未完成命令/输入法组合优先处理。
5. 实现多行选择、分行、合并、混合 inline/display；物理源码换行不错误拆开完整环境。
6. 在长文档视口中虚拟化显示，不为所有离屏公式创建可输入 DOM。核心文档不因视图卸载丢失状态。
7. 给布局输出增加显式对齐锚点数据，供 P11 MathType 兼容层使用。

**验收与产物：** 2×2/3×3/大矩阵、分段函数、多个 `&&`、嵌套分式矩阵、多行跨选、分合行、复制粘贴几何与历史样本。

**停止条件：** 矩阵被当作一坨不可编辑字符串；不同层级 Enter 混用；为了对齐重写所有行的 LaTeX。

### P09：源码、命令、宏、剪贴板与自定义符号闭环

**前置：** P08 完成；仍在 sandbox 验证，不提前影响 Office。

**实施：**

1. 保留 CodeMirror，接入版本化源码事务。桥接更新带 origin，避免视图→store→视图循环。源码编辑历史与结构历史有唯一协调者。
2. 保留命令数据、别名、频率与搜索；命令 schema 升级为结构动作。把 `\partial` 符号和“偏导分式模板”分成不同动作。
3. 自有 CommandInputState 负责 `\` 临时输入、候选、取消、确认；确认生成结构事务。删除 MathLive 原生 popover 克隆机制时保留现有交互习惯。
4. 支持现有宏和模板参数的原文保留、受限展开、参数内编辑；建立未支持语法诊断，不靠旧解析器过滤错误。
5. 新建带版本的内部结构剪贴板格式，同时提供纯 LaTeX/文本回退。验证大小、节点类型、深度、样式和自定义资产引用；不执行剪贴板中的任意脚本或 HTML。
6. 复用自定义符号定义、geometry、metrics、role、limitsBehavior、设计器与用户库；替换其 MathLive 宏/CSS 注入。嵌入新文档时保留必要资产快照，避免同名用户库后来改变已存公式。
7. OCR 仍产出 LaTeX；只换插入目标适配。异步 OCR 返回时校验原会话/公式/revision，映射目标失败不得插入别人的公式。
8. 已有 `.vtx`、历史记录、字体偏好和快捷键配置继续可读。首轮不把内部 core snapshot 设为唯一持久化格式。

**验收与产物：** 现有 runtime command registry 的能力覆盖清单、宏往返、未知输入保存、源码未完成恢复、结构/纯文本复制、OCR 目标过期和自定义符号测试。

**停止条件：** CodeMirror 被重写；源码编辑自动美化全部文档；两个历史管理器重复撤销；用户符号库格式被无必要升级。

### P10：接入主应用，保持 Office 路径不动

**前置：** P06—P09 核心与产品行为通过。只有主应用入口可以启用新引擎。

**实施：**

1. 建立主应用 EngineHost，保留外层布局、工具栏和配置 UI。只替换中央数学编辑表面与其状态桥。
2. 重新定义内部编辑选区 token，区分 legacy-engine 与新 core selection。禁止把 MathLive offsets 直接当成 NodeId/源码偏移；这类内部 TS 接口可受控调整，但 Office wire DTO 不改。
3. 新旧引擎在完整文档/公式会话边界选择；一个公式内部不混用两个可编辑子树。选择后锁定到会话结束，避免输入中途换引擎。
4. 启动失败恢复最后文档，迁移期间可明确返回旧入口；不得把“自动返回 MathLive”算成新引擎通过。最终移除该正式兜底路径。
5. 保留所有现有 inputBehavior 设置，字体/主题变化只影响 View 或布局上下文，不触发无关源码改写。
6. MathPreview/工具栏静态预览逐步使用同一核心的无交互渲染；量化并缓存，不为每个缩略图创建输入会话。
7. 检查主应用启动、关闭、二次打开、窗口布局、历史恢复、文件打开保存与 no-OCR 安装行为。新内核不启动本地 OCR 环境。

**验收与产物：** 主应用完整编辑回归与新旧对照；Office 保护区和默认路径无增量变化。主应用默认切换必须经过明确验收决定，不凭 sandbox 成功自动启用。

**停止条件：** 改动 App/全局 CSS 后遮挡 Office 候选或改变其窗口行为；需要修改 Session 才能保存主应用文件。

### P11：建立 Office 输出兼容证明，不接管真实 Session

**前置：** P10 主应用可验收，P01 Office 样本与消费者已固定。

**实施：**

1. 新建 OfficeCompatibilityAdapter，输出原 `OfficeExportResult`。core 不直接产生或修改 Session 状态。
2. 按消费目的请求结果：MathML-only、完整向量/PNG 预览、静态主应用预览。保留 OMML 无 SVG 快路径，不对每个目标无条件执行完整图形导出。
3. 实现 `office-v1` 字号、padding、baseline、line/display 映射；不把 zoom 或 PNG 像素放大系数带进几何。
4. MathType MathML 对齐标记从新布局列锚点生成；保留旧解析器接受的属性和单位，不依赖新 SVG 恰好长得像 MathJax。
5. 用原 Office SVG→EMF 消费者测试安全子集和向量属性，验证 PNG 栅格结果与 SVG 自然尺寸对应；不得放宽原消费者校验。
6. 自有 Presentation MathML 编码需覆盖原转换器需要的语义：结构、符号、样式、limits、矩阵/对齐、普通文本和自定义符号 fallback。MathML Core 的支持子集不等于当前 Office 转换器完整输入契约。
7. 测试 metadata 原值保留、无修改输出策略与新旧图形差异。缓存复用必须绑定原内容和原渲染设置，不能仅凭 formulaId 命中。
8. 验证每一种原已支持的 objectMode 对输入的要求。缺必要数据时错误要在宿主替换前发生，不允许产生“成功 0 个”或半写入公式。

**验收与产物：** 新输出喂入原 TS/Rust/C# 校验与转换器的离线契约报告、几何差异报告、MathType 对齐回读样本。这里的消费者测试仍不等于真实 UI 验收。

**停止条件：** 需要改 C# WordFormulaService/编号/COM 才能接受新结果；同名字段换了单位；新旧 MathML 语义不一致。留在适配层解决或登记架构阻塞，不自动突破保护区。

### P12：接入 Office 交互编辑窗口，保持提交机制原样

**前置：** P11 通过，P00 已确认实际运行二进制。自动转换路径仍关闭新内核接入。

**实施：**

1. 只对交互编辑 runtime 注入 EngineHost/导出提供者；保留 useOfficeSession、save 队列、鉴权、窗口复用、原 apply/cancel/close 和快捷键处理。
2. 首次 Session→core 导入不标 dirty，不产生编辑事务，不自动提交、不重绘 Word 对象。
3. 每个加载/导出异步操作绑定 `sessionId + formulaId + localRevision`。会话切换/关闭后旧结果不得提交。
4. 显式 Apply 先提交当前已确认输入，再导出同一个 revision；未确认输入法按既定策略完成或阻止 Apply，不静默丢掉预编辑文本。
5. 导出完成后走原 PATCH/commit 流程。不要由 core 决定 completed，不绕过现有防陈旧 autosave 机制。
6. 用专用 Word 文档通过真实 Ribbon 插入/编辑 VisualTeX OLE、OMML、MathType；行内/行间、有编号/无编号均覆盖。PowerPoint 只测试其当前真实支持的格式集合，不把 Word 功能凭空扩展到 PPT。
7. 验证取消、X 关闭、自动提交开关、连续 Apply、窗口复用、主题切换、加载失败/重试及未改动打开关闭。
8. 只对测试文档验证保存重开和离线缓存；用户文档与 Word 进程保持原状态。

**验收与产物：** 真实 UI 操作记录、Session 时间线、源文档前后结构/对象数量/编号/格式检查、实际加载 DLL 与前端资源哈希。明确这次改的是编辑窗口，没有改 Office 文档算法。

**停止条件：** 重复插入、重复 Session、取消改变文档、正文默认字体/段落格式污染、编号进入数学对象、粘贴副本不独立、渲染正确但应用错误。禁止用“编辑器本身没问题”跳过这些回归。

### P13：接入隐藏转换/批量导入/重绘链

**前置：** P12 交互窗口真实验收通过。此阶段单独切换自动转换出口。

**实施：**

1. 识别 `vsto-convert` 与 converter batch 生命周期，使用 core 的无交互实例，不创建 SVG 输入 View、输入法宿主或候选面板。
2. 保留队列顺序、逐项结果、失败继续/停止规则、取消规则与现有用户提示。规则以 P00 实际行为为准，不在本阶段重新设计。
3. MathML-only 任务不加载不必要的图形/PNG流程；图形输出使用批量字体/glyph 缓存，不逐公式重新初始化整个内核。
4. 每个结果绑定原请求身份；结果顺序变化不能使第 N 个结果覆盖第 N+1 个公式。
5. 按原能力矩阵测试各格式之间的所选/全文转换，以及 LaTeX 批量导入、重绘、还原源码。不能只测试一条 VisualTeX→MathType 方向。
6. 分别运行 100 和 1000 公式样本、长单公式、混合正文/复杂矩阵/轻微错误源码；保留错误源和未处理对象，检查正文完整性。
7. 保留私有 MathType 原生运行时和其 WMF 路径，不恢复已撤回的正式 MathType 优先实验，不更改 COM/App Paths 注册。
8. 与相同文档/相同二进制基线比较总时间、单项分段与内存。不能以批处理一次总成功掩盖大文档的非线性增长。

**验收与产物：** 各方向 UI 结果与 performance 表；对象数量、失败数、源码/语义、编号和引用对照；转换前后原始文档结构证据。要求成功数/失败项均可追溯。

**停止条件：** 批量流程弹出交互窗口、长时间抢 UI、公式错配、1000 公式失败或异常增长、为性能关闭验证/编号/取消能力。

### P14：跨宿主、资源、性能与故障加固

**前置：** P13 Windows 路径稳定。macOS/Web 实际代码变更仍需相应许可。

**实施：**

1. 同一 core 源码和字体样本在原生测试、Web 浏览器、Windows WebView2 中验证；获得许可后在 macOS WKWebView 验证，不凭浏览器测试推断宿主输入法完全相同。
2. 核对主应用嵌入资源和 Office companion 静态资源两种加载方式。验证 WASM/字体文件内容哈希、URL、MIME、缓存、离线安装及路径含空格/非 ASCII 字符。
3. 防止旧 JS 绑定加载新 WASM，或旧 Office UI assets 对应新 core ABI。资源版本校验失败要给清晰错误，不清空用户 WebKit 数据。
4. 做多窗口、多 Session、快速切换、异步 OCR、字体加载迟到、重复提交、关闭期间导出等竞态测试。
5. 做大结构、恶意嵌套、超大粘贴、未知宏、字体损坏、内存预算、WASM trap 的恢复测试。失败保留源数据和旧 Office 对象。
6. 测试真实中文输入、缩放命中、系统快捷键、复制到其他软件及读屏/辅助技术；不能只测拉丁字符。
7. 优化前后都记录预算和热点；光标移动不 parse，blink 不 layout，批量输出不初始化交互窗口，普通局部编辑不遍历无关公式。
8. 原生库/无界面 CLI 是同源能力，不引入新的 Office 进程内加载方案。

**验收与产物：** 宿主矩阵、冷/热性能、资源一致性报告、内存与错误恢复报告。未授权/未测试平台保持未验收状态，不影响对 Windows 已验证范围的准确描述。

**停止条件：** 为了跨端统一修改 macOS Office 代码而无许可；假定 EditContext/SIMD/线程在所有宿主可用；恢复方法是删除用户配置或关闭 Word。

### P15：移除 MathLive/MathJax 正式运行依赖

**前置：** 承诺支持的现有能力在新内核和全部已接入出口达到完整验收；有旧引擎运行时 fallback 的项目不满足本阶段条件。

**实施：**

1. 扫描运行依赖图、动态 import、构建 alias、Vite transform、静态 CSS、预览、SSR 校验、命令注册与自定义符号桥。
2. 删除正式 MathLive 入口、`.ML__*` 专用补丁、shadowRoot 修正、私有 model 操作、popover 克隆和 MathLive verifier。
3. 替换 `LatexCopyService` 中的 `validateLatex`、`MathPreview` 的 convertLatexToMarkup、字体/希腊快捷键与自定义字符中的 Mathfield 类型依赖。
4. 移除正式 MathJax SVG/MathML 排版依赖；保留可独立维护的输入别名/兼容数据，不保留无意义的 MathLive 命名语义。
5. 标准版和 no-OCR 版分别检查构建产物，确认旧引擎没有通过打包副本、动态加载或测试资源进入正式包。
6. 旧依赖作为测试 oracle 必须隔离为明确开发用途，正式产品在没有它们时仍可 build、编辑、预览、导出。
7. 更新依赖/字体/数据许可证清单。保留设计经验和产品测试，不删除失败样本掩盖覆盖缺口。

**验收与产物：** 依赖审计、全量回归、纯新内核构建报告。若只移除 npm 包但仍复制了其内核进项目，不得称自主替换完成。

**停止条件：** 仍有任何承诺功能依赖旧引擎；靠把功能藏起来通过扫描；旧文档不能安全读取。

### P16：安装包、固定 Office 子包和最终验收

**前置：** P15 完成；本阶段打包/安装按用户明确许可执行，不自动发布。

**实施：**

1. 审查实际构建脚本再执行。当前 `tauri_build.mjs → build:bundle` 会构建 Office 输入，不能在本项目中无条件运行后声称 Office 二进制未变。
2. 为候选包使用独立输出目录和固定的已验收 Office 子包输入。若必须改打包路径以复用它们，做最小单独改动并验收；不得覆盖原完整安装包或 no-OCR 包。
3. 核对 app exe、JS/CSS/WASM/字体、bridge、x64/x86 MSI、Word DLL、私有 runtime 与证据清单。不能只核对前端 index.html 两个静态引用而漏掉动态 WASM。
4. 真正检查安装包内容与安装结果。no-OCR 版不得夹带 Python/wheelhouse/模型目录；核心字体/WASM 属于引擎资源，不要错误当 OCR 删除。
5. 在有许可的专用环境测试安装器 runtime guard：否则不得结束私有进程；外部同名 MathType 与 Word 保持不受影响。保持已通过的二次检测/竞态防护。
6. 使用最终安装包、实际加载的正式 Office 插件和真实 Ribbon 完整复测，不拿开发服务器版本替代安装版。
7. 验证旧文档打开编辑、关闭不改、测试副本保存重开、离线预览、各种格式、编号/引用/复制与大文档性能。
8. 完整版与 no-OCR 分别记录最终大小、SHA256、子包 SHA、测试时间及报告。未经用户明确许可不上传、不发 release、不推送额外分支。

**最终交付：** 自研 core 源码、固定版本绑定、兼容适配器、源码/布局规格、全量测试样本、每阶段证据、两种候选包的来源记录和未覆盖项。满足第 14 节完成定义后才能报告整体完成。

**停止条件：** 最终包与测试包哈希不一致、Office 子包来源不明、no-OCR 资源泄漏、任一既有功能出现未解决回归、需要关闭用户文档才能安装或验收。停止交付该候选，不覆盖原包，不自行发布。

## 10. 测试设计：什么才算通过

### 10.1 四级证据不能互相替代

| 证据层 | 证明内容 | 不能证明 |
|---|---|---|
| 内核单元/性质测试 | 结构、位置、布局、序列化规则 | 真实输入法、Office 按钮链 |
| 浏览器/独立 View 测试 | WASM 加载、交互、显示、源同步 | 真实安装包与 Word 文档行为 |
| 旧消费者/协议测试 | TS/Rust/C# 接受输出、语义和几何契约 | Ribbon 到最终文档的全链路 |
| 最终包真实 Office UI | 实际主应用/编辑窗口/插件/文档行为 | 没测过的平台或格式 |

P00—P09 的隔离开发主要跑前三类并验证生产路径未改变，不必每次解析器修改都运行 1000 公式 Word 测试。进入 P12、P13、P16 必须达到相应真实 UI 门槛，不能因前三级成功跳过。

### 10.2 核心不变量

1. 保真入口中，支持且未修改的有效源码往返保持约定的原文；规范化只在明确出口发生。
2. 编辑事务后所有节点引用、选区、源范围和宏参数关系有效；非法状态不能被提交成“正常公式”。
3. undo/redo 恢复内容、结构、选区及身份关系；连续输入合并不吞掉独立工具栏/粘贴操作。
4. 复制公式内部节点与原节点独立；Office 的原 formulaId/副本身份规则仍由原流程处理。
5. 一次 IME 确认只有一次文本插入；取消未确认候选不删除已确认公式。
6. 同一 revision 的场景、光标、命中区与导出一致；旧异步结果不能覆盖较新编辑。
7. 缩放/DPI/PNG 采样率变化不改变语义字号、源码或自然导出尺寸。
8. 数学符号区分字形、语义和输入命令；不能因为外观相似把不同符号替换。
9. core 不依赖 DOM/React/Tauri/Office；Native/WASM 同源规则在固定资源下保持几何一致。
10. 用户主动未完成的公式可保存草稿，但不以假的完整公式提交到 Office。

### 10.3 最低样本清单

以下是必须新增或整理的测试规格，不是已经存在且通过的测试文件。

| ID | 样本/操作 | 重点 |
|---|---|---|
| E01 | `x+y=1`、希腊字母、关系符 | 字符与数学类别、方向移动 |
| E02 | `\frac{x_1+\text{速度}}{y^2}` | 第一条自主闭环、中文、脚本、分式 |
| E03 | 分母空槽、`\fr`、未配对定界符 | 草稿与提交合法性分离 |
| E04 | 多层分式、带指标根式 | math style、cramped、伸展与边界 |
| E05 | `\left(\frac{a}{b}\right)^2` | 结构边界命中、括号跳出 |
| E06 | `\hat{\vec{x}}`、长 overline/brace | 重音锚点、空槽删除、嵌套退出 |
| E07 | 求和、积分、limits/nolimits | 导航顺序、大小运算符、上下限 |
| E08 | `\mathbf`、`\mathit`、`\bm`、函数正体 | 样式保持与不同数学语义 |
| E09 | 2×2/3×3 与大矩阵增删行列 | 单元格选区、删除/粘贴/撤销 |
| E10 | cases 内回车、外层回车 | 环境行与文档行区分 |
| E11 | aligned 多组 `&&`、嵌套矩阵 | 统一列几何与 MathType ruler |
| E12 | 多行跨选、分行、合并、混合模式 | 外层文档身份和完整环境保存 |
| E13 | 宏参数编辑、未知宏、不完整参数 | 原文保留、受限展开、能力报告 |
| E14 | 中文候选确认/取消与反斜杠连续输入 | IME 无重复事件、无误分行 |
| E15 | Unicode 扩展字符/组合字符 | UTF-8/UTF-16/grapheme 映射 |
| E16 | 原公式复制后只编辑副本 | 内核和宿主两层身份独立 |
| E17 | OCR 返回前切换行/删除目标/换 Session | 异步锚点有效性 |
| E18 | 自定义符号更名、库变更、离线打开 | 资产引用、快照与旧数据兼容 |
| E19 | 50%/100%/200% zoom 与不同 DPI | 命中、候选框和导出字号不混用 |
| E20 | 100/1000 公式、重复结构、长单公式 | 增量失效、缓存、视口与内存 |
| E21 | 字体缺失、WASM 失败、超大/恶意输入 | 无数据丢失、恢复与预算 |
| E22 | 旧文档打开不编辑→关闭/再打开 | 不变脏、不重排、不丢缓存 |

测试不能仅比较最终 LaTeX 文本，还要包含结构、当前位置、历史和几何。属性测试须记录随机种子和重现输入；失败不得用跳过当前用例收尾。

### 10.4 Office 真实验收矩阵

沿用 P00 查明的真实支持集合；不存在的格式组合记为“不适用”，不是“通过”。

| 场景 | 必测项目 |
|---|---|
| Word 单公式 | VisualTeX OLE / OMML / MathType 的行内、行间，编号开关，字体字号，创建与编辑 |
| Word 多文档 | 两个测试文档之间切换、窗口复用、切换期间导出完成，结果不得写错对象 |
| 编辑收尾 | Apply、连续 Apply、显式取消、X 关闭、自动提交配置、快捷键、未编辑关闭 |
| 全局操作 | 所选/全文格式互转、重绘、导入、还原 LaTeX；各已支持方向分别验收 |
| 编号与引用 | 添加/取消编号、编号格式切换、SEQ/REF 更新、交叉引用、复制后编号策略 |
| 复制/剪切/粘贴 | 同文档与跨文档，副本独立，原编号保持/显式刷新符合基线，不泄漏域代码 |
| 正文保护 | 默认中西文字体、字号、段落样式、行距、制表位、正文内容、相邻段落不污染 |
| 大文档 | 100 与 1000 公式、混合正文、复杂公式、含错误源码、取消/失败后的未处理部分 |
| 持久化 | 只对测试副本保存重开，公式 ID/原生来源/预览缓存/尺寸/引用仍有效 |
| PowerPoint | 当前已存在的图片/OLE/OMML 能力、编辑、复制独立、尺寸比例、slide 身份、保存重开 |
| Runtime 共存 | 专用测试环境中的私有 runtime、正式 MathType 共存与未安装正式 MathType 场景 |
| 安装包 | 完整/no-OCR、安装后真实启动与编辑窗口，Office 子包哈希及实际加载来源对应 |

所有新式编号布局以 P00 实际代码/样本为基线，不参考可能过时的帮助文档重新设计。尤其禁止把真正 display OMML 改成 inline 加字号补偿，也不能随意改为另一种表格/制表位编号方式。

### 10.5 真实 UI 工具的正确用法

现有工具入口包括 `docs/remediation-3d207d7/word-ui.ps1`、`ui-flow.mjs`。它们自身可能有历史 dirty，执行前先读当前内容与参数，不直接照搬旧对话的固定坐标。

按顺序确认目标 PID、活动窗口、测试文档身份，读取 Ribbon/UI 树，再定位并触发真实控件。实际输入必须到可见编辑窗口；通过 COM 准备源码或读取结果不等于执行了 UI 操作。

记录按钮动作、开始/结束时间、最终对象计数、元数据、段落和编号结构，以及截图。应用完成要以真实宿主文档结果确认，不能以脚本返回 0 或关闭了编辑窗口作为唯一成功标准。

交接中的环境为 Microsoft Office Professional Plus 2021 x64、Word `16.0.14334.20848`；P00 重新确认实际版本。不要因为某个 Interop 类型库版本是 15/16 就把用户 Office 错报成另一产品年代。

不得使用按进程名停止 Word、按未保存状态猜测文档归属、无条件 `Application.Quit` 等脚本。对测试文档进行保存重开前必须确认它是本轮创建或批准的副本，用户原文档始终保持原状态。

## 11. 性能与观测

### 11.1 两组指标分开

新内核指标覆盖事件处理、事务、布局、绘制和导出；Office 指标覆盖 Ribbon/编辑器操作到文档真正更新的全程。不能用 Rust 核心耗时替代用户等待时长。

初始内核工程目标如下，属于待测目标，不是现有成绩；P01 固定测试机器、字体、数据与统计方法：

| 操作 | 初始目标/约束 |
|---|---|
| 光标/选区移动 | 不触发解析和完整布局；常用小公式核心操作 p95 ≤ 8ms |
| 常用局部编辑 | 约 200 节点以内、资源已热的典型公式，事件到可见更新 p95 ≤ 32ms |
| 光标闪烁 | 不重建数学场景，仅刷新叠加层 |
| 单节点修改 | 不扫描文档无关公式；记录受影响节点/盒数量 |
| 批量导出 | 字体与 glyph 缓存复用，不创建交互 View；冷启动成本单列 |
| 1000 公式文档 | 视图只挂载必要内容，不能因为离屏公式增加而使单公式编辑近似线性变慢 |

大矩阵、深层嵌套、复杂中文塑形另列分组，不把它们藏在平均值里。不得为达到目标删掉正确性校验、关闭编号或取消功能。

Office 的用户既有性能要求继续有效。新旧用相同样本、模式、编号与冷/热条件比较；内部初始报警线可设为基线加 `max(10%, 50ms)`，但这不是自动放宽用户已认可的绝对上限。批量分组比较 100/1000 的每项时间与增长趋势，异常增长必须定位，不能只汇报平均单公式耗时。

历史 44.2s/106.6s 仅作交接参照；新阶段必须记录实际采用的二进制与原始运行数据，不能把旧成绩写成候选成绩。

### 11.2 必须记录的观测字段

建议每个测试记录：runId、core/build/profile/字体哈希、宿主版本、公式/Session 匿名标识、revision、节点数量、事件类型、解析/事务/布局/绘制/导出耗时、缓存命中、invalidatedNodes、峰值内存、最终成功/失败。

业务数据和 API token 不进入公开日志。性能日志不能在每次光标移动时同步扫描全文或写大块文件。UI 自动化等待时间与真实算法耗时分别报告。

## 12. 现有代码保留/替换清单

下表基于当前 Windows 实现。路径相对于 `apps/windows`，函数名比行号更适合后续定位。

| 代码/资产 | 处理方式 | 主要阶段 |
|---|---|---|
| `src/types/formula.ts` 的旧文档/行格式 | 保留兼容边界，不当最终内核唯一模型 | P03/P09/P10 |
| `src/stores/editorStore.ts` | 保留设置与 UI 数据；公式写入改为受控 core 投影 | P10 |
| `src/history/HistoryManager.ts`、`historyTypes.ts` | 保留产品分组/检查点要求；结构事务与位置映射重建 | P04/P09 |
| `src/autocomplete/{commandRegistry,runtimeCommandRegistry,CommandSearchEngine}.ts` | 保留名称/别名/分类/频率；增加结构动作与能力表 | P07/P09 |
| `src/toolbar/FormulaToolbar.tsx` | 保留 UI/产品逻辑，替换执行与预览入口 | P09/P10 |
| `src/source-editor/LatexSourceEditor.tsx` | 保留 CodeMirror；接统一文档/草稿/历史桥 | P09 |
| `src/workspace/EditorWorkspace.tsx` | 保留面板与布局；替换 applySource 和编辑表面状态桥 | P10 |
| `src/editor/MathEditor.tsx` | 迁移产品行为，替换模型/输入/光标/几何实现；不先机械包装一万行 | P04—P10/P15 |
| `resetMathLiveModelRootForExternalSync` 等私有 model 操作 | 随新会话模型退出，不搬进新内核 | P10/P15 |
| `structuralBoundaryOffsetFromPoint`、caret repaint、占位符 DOM 修正 | 由 LayoutSnapshot/编辑状态原生承担 | P05—P08/P15 |
| `src/editor/imeCompositionGuard.ts` | 保留失败样本和需求，事件桥按新组合事务实现 | P04/P06 |
| `src/editor/formulaFontPreferences.ts` | 保留用户设置 ID 和字体偏好语义；实际度量由 FontProvider 管理 | P05/P10/P11 |
| `formulaFontRuntime.ts`、MathLive 专用字体样式 | 移除其 `.ML__*`/shadow 注入，改新 View 样式 | P10/P15 |
| `selectionStyleToggle.ts` | 保留行为样本；结构成熟后改为样式事务，不长期 regex 重写选中源码 | P07/P09 |
| 自定义字符 registry/types/designer/geometry/资产 | 保留有效数据与校验；拆掉 Mathfield 宏桥 | P07/P09 |
| `customSymbolRendering.ts` | 按 MathLive、MathJax、通用资产职责拆分，复用资产而非整文件照搬 | P09/P15 |
| `latexCompatibility.ts`、package/罕见积分数据 | 保留已验证输入别名和合法数据；迁移到独立语言/字形规则 | P07/P09 |
| `src/components/MathPreview.tsx` | 保留组件用途与缓存需求，改用 core 静态输出，不产生输入会话 | P10/P15 |
| `src/clipboard/LatexCopyService.ts` | 保留封装/复制格式，解析校验与结构操作改新核心 | P09/P15 |
| `src/export/runtime.ts`、`latexToSvg.ts` 等出口 | 迁移期保留；最终同名边界调用自主输出，不能一直依赖 MathJax | P11/P13/P15 |
| `src/office/shared/mathTypeAlignmentGeometry.ts` | 冻结输出属性语义；新 compat 从布局锚点生成，不改下游 MTEF 消费者 | P11 |
| `src/office/dialog/OfficeDialogApp.tsx` | 只换编辑表面/导出提供者；交互/转换分阶段 | P12/P13 |
| VSTO/COM/OLE/编号/引用/转换算法、私有 runtime | 保留，不参与内核重写 | 全阶段保护 |
| OCR 识别、API、模型安装 | 保留原机制；只接插入动作。no-OCR 仍不装本地 OCR 资源 | P09/P16 |
| `vite.mathliveIntegralCompatibility.ts`、runtimeSafety、alias、static.css | 迁移完成后删除正式依赖；保留相应产品回归测试 | P15 |

不能把“保留模块用途”误解为“内部实现完全不动”，也不能把“替换内部实现”误解为允许改变用户功能。

## 13. 构建、现有回归与工作纪律

### 13.1 现有命令与待新增命令分开

以下脚本在当前 `apps/windows/package.json` 中已有入口。此列表不是批量执行指令；agent 先读脚本，确认不会操作非测试文档、覆盖已验收文件或修改全局环境，再按阶段执行：

```text
npm run test:history
npm run test:editor
npm run test:editor-history
npm run test:input-behavior
npm run test:formula-hotkeys
npm run test:toolbar-completion
npm run test:source-structural
npm run test:source-editor-ux
npm run test:office-metadata
npm run test:office-formula-document
npm run test:office-session-safety
npm run test:office-apply-shortcut
npm run test:mathtype-align-geometry
npm run test:windows-office-architecture
npm run test:windows-ocr-editor-parity
```

原有 .NET 测试、真实 Word 工具和安装验收也要审查当前配置后使用。历史总数 503 不是不可变目标；不得为保留通过数而删除测试或降低断言。

以下是 P02 建立 workspace 后的**预期命令形态**，目前不能假设已经可运行。crate 名称、锁文件及工具链须先在该阶段创建并记录：

```text
cargo test --manifest-path engine/Cargo.toml -p visualtex-engine-core --locked
cargo build --manifest-path engine/Cargo.toml -p visualtex-engine-wasm --target wasm32-unknown-unknown --locked
```

为其配置独立构建输出，不清空已有 target。WASM 测试需要真实执行而不只是成功编译；配套 runner 在 P02 实现。

### 13.2 包装脚本风险

当前完整版入口是 `npm run tauri:build`，no-OCR 是 `npm run tauri:build:no-ocr`。两者不是单纯编译前端，会触及 bundle 输入准备与 Office 构建。开发核心阶段不能直接用它们当日常测试命令。

不要因构建锁定而关闭用户 Word，不运行不明“清理脚本”，不清除用户 WebKit/localStorage 配置。隔离候选的输出目录、资源清单和证据目录；是否生成 MSI、替换安装资源必须显式决定。

### 13.3 状态报告模板

每阶段至少保留下面这些字段：

```text
阶段：Pxx / 名称
开始基线：本地 HEAD、实际 dirty 清单标识、运行二进制哈希
本次允许路径：……
实际改动：……
原功能保留情况：……
验证：命令 / 实际 UI 动作 / 宿主 / 样本 / 结果
性能：before / after / 冷热状态 / 原始报告位置
保护区增量：无；或列出已批准例外及原因
未覆盖/失败：……（明确 NOT_RUN / FAILED / BLOCKED）
准入结论：是否可以进入下一阶段
下一步：具体阶段和第一个动作
```

不要把“正在定位”汇报为“已经找到根因”，不要把静态检查说成真实 UI，通过哪一层就只报告哪一层。交接时列出仍打开的测试窗口和用户文档保护状态，不要求下一 agent 猜测。

本计划内阶段会改变目标内容与测试范围，但不会自动给予 Git/安装/发布权限。后续提交只 stage 明确文件，禁止顺带提交历史 dirty、运行数据、安装包、字体或密钥。

## 14. 最终完成定义

只有同时满足下列条件，才能报告“VisualTeX 已拥有自主可视化引擎”：

- 公式表示、编辑、光标、命中、数学布局和自身图形输出由自有核心实现。
- 新核心在 Native/WASM 使用同一套规则，纯核心不依赖 Tauri/DOM/Office。
- 承诺支持的现有公式、命令、中文输入、源码往返、多行/矩阵、自定义符号与历史功能均通过能力矩阵。
- 正式主应用、Office 编辑窗口和转换/导出路径不再运行 MathLive/MathJax；未实现语法如实报告，旧文档不丢失。
- Office DTO、API、COM/OLE、元数据、单位和状态语义保持既有兼容；没有用重写 Word 后端掩盖内核不兼容。
- Word/PowerPoint 当前已有功能经过实际最终包、真实 UI 验收，保护区改动为零或仅存在用户明确批准且单独验收的例外。
- 100/1000 公式和单公式性能不出现未解释回归；no-OCR/完整版资源与已验收 Office 子包可追溯。
- 未改动打开旧公式不会自动改写；取消/故障不破坏原对象；复制身份独立。
- 文档、样本、日志与构建来源完整；没有“过渡 fallback 仍在，却宣布全自研”的情况。

“测试都绿”若只是旧引擎回退成功，不满足以上完成定义。跨平台只对实际验收的平台负责；未授权的 macOS/Web 接入不能被计入完成。

## 15. 主要风险与应对

| 风险 | 最早发现阶段 | 处理 |
|---|---|---|
| 源码、store、core 三个独立真源 | P03/P04 | 单写者 EngineSession，事务桥、origin 和 revision |
| MathLive offset 被伪装为通用位置 | P01/P04 | 新位置模型，旧 token 隔离；不做无法证明的直接映射 |
| 同名字体不同度量，屏幕和导出不一致 | P05 | 字体实际资源身份先于布局，统一场景与单位 |
| 只完成渲染，没有真正编辑 | P06 | 真实键盘/IME/删除/选择的纵向闭环 |
| 用户数据包含未知 TeX 或新符号 | P03/P09 | 保留原文与资产；明确能力，不静默降级为不同公式 |
| MathType aligned 标记丢失 | P11 | 独立 ruler 契约和 MTEF 回读，布局直接产出锚点 |
| 改编辑窗口却破坏隐藏转换 | P12/P13 | 两条 runtime 独立开关与测试，资源初始化按用途 |
| 新导出无法被旧 EMF 消费者接受 | P11 | 安全子集降低，不放宽 Office 转换器 |
| 遗留 autosave 或旧导出覆盖新 Session | P12 | 原保存队列保持、异步身份绑定、revision 过期检查 |
| 编译顺带重建已验收 Office 子包 | P00/P16 | 固定来源、独立输出、按哈希而非文件名验收 |
| 性能被全量序列化/DOM测量/多次 IPC 拖垮 | P05/P10/P14 | 增量事务/几何、同 View WASM、批量资源传输 |
| 架构越来越像万能平台 | P01/P02 | 限制 crate/插件/后端数量，优先完成可验收纵向功能 |
| 为了跨平台擅自改另一端 | 全阶段 | 核心可移植，平台改动另行授权与验收 |

## 16. Agent 接手时的第一轮执行要求

第一轮从 P00 开始，不先改 `MathEditor.tsx`，不运行完整安装构建。先确认真实基线、历史 dirty、运行二进制与测试证据，并完成 P01 的保护区与接口/行为规格。

完成基线后，下一轮实现 P02—P06 的最小自研闭环。第一块实际引擎代码应成为最终核心的一部分，而不是先把全部 MathLive 调用包装一遍。后续按依赖关系扩展，在主应用证明稳定后才接 Office 输出和窗口。

可直接交给 agent 的任务说明：

> 读取本工作区 `docs/engine-rebuild/VISUALTEX_ENGINE_PROJECT_PLAN.md`，按 P00→P16 推进 VisualTeX 自研数学可视化引擎。用户的硬约束是现有 Office 功能与协议不受影响，主要改变只能落在主应用、编辑窗口和内部输出适配层。先完成来源基线、保护区与契约测试，不要直接替换 MathEditor，不要修改 VSTO/COM/编号/Session 状态机，不要关闭或保存用户 Office 文档。保留历史 dirty/untracked，不切分支、不做 reset/clean/stash/pull/rebase/merge，不擅自提交、推送、安装覆盖或发布。当前本地 HEAD 为 c4690c2，远端 main 为 6bb4fce，二者历史不同但 Windows 已提交内容已同步，禁止据此强推。第一块新内核实现必须在独立 sandbox 中贯通自有结构、事务、中文输入、数学布局和 SVG 输出；随后分别验收主应用、Office 交互窗口和隐藏批转换。每阶段给出实际文件、测试层级、性能与未覆盖项；需要真实 Office 验收时必须走真实 Ribbon/UI，脚本回调成功不能替代。任何回归先修新内核/适配层，不拿修改 Office 后端或关闭功能作为捷径。

本说明不要求未经批准跨平台改动，也不要求 agent 在证据不足时宣布整个项目完成。

## 17. 参考资料与证据说明

### 17.1 仓库实现证据

本文的 Office 契约与当前耦合分析以实际读取的 Windows 工作副本为依据，主要入口：

- `src/editor/MathEditor.tsx`：Mathfield 创建、私有模型恢复、DOM 命中修正、结构删除、历史与多行状态。
- `src/types/formula.ts`、`src/history/historyTypes.ts`、`HistoryManager.ts`：现有持久化与数值选区。
- `src/source-editor/LatexSourceEditor.tsx`、`src/workspace/EditorWorkspace.tsx`：CodeMirror 与草稿预览。
- `src/export/runtime.ts`：MathJax SVG/MathML、字体后处理、padding/baseline 和栅格化。
- `src-tauri/src/system_math_glyphs.rs`：Windows 字体扫描与轮廓读取。
- `src/office/shared/sessionClient.ts`、`formulaMetadata.ts`、`mathTypeAlignmentGeometry.ts`：实际 DTO、元数据与对齐标记。
- `src/office/dialog/OfficeDialogApp.tsx`、`useOfficeSession.ts`：交互/自动转换入口、OMML 快路径与保存队列。
- `src-windows/VisualTeX.WindowsOffice.Contracts/FormulaOleContract.cs`、`OfficeFormulaSizing.cs`、`FormulaFontSize.cs`：OLE ABI、单位与字号语义。
- `src-windows/VisualTeX.WindowsOffice.VstoShared/OfficeOlePreview.cs`：当前 SVG 子集与向量校验。
- `vite.mathliveIntegralCompatibility.ts`、`vite.mathliveRuntimeSafety.ts`：构建期 MathLive 内部修改。
- `scripts/tauri_build.mjs`、`build_platform_bundle.mjs`、打包验证脚本：现有标准/no-OCR 构建链。

后续源码变化时按函数重新定位，不将本文行号或历史帮助文档当作比当前实现更高的证据。

### 17.2 外部一手参考（访问日期 2026-09-11）

这些资料支持基础机制与标准，不证明 VisualTeX 已实现相应功能。本文的阶段划分、边界与验收目标属于本项目设计决策。

- [R01] Microsoft，[OpenType MATH — mathematical typesetting table](https://learn.microsoft.com/en-us/typography/opentype/spec/math)。用于字体数学常量、字形信息和伸展构件设计；不是完整排版引擎。
- [R02] W3C，[MathML Core](https://www.w3.org/TR/mathml-core/)。用于基础数学布局规则与参考测试；本次访问页为 2025-06-24 Candidate Recommendation Snapshot，不误写为覆盖全部 TeX/Office 功能的最终标准。
- [R03] Tauri，[Architecture](https://v2.tauri.app/concept/architecture/) 与 [IPC](https://v2.tauri.app/concept/inter-process-communication/)。用于区分 WebView、Rust 核心和现有异步通信，不据此推导具体性能成绩。
- [R04] Unicode，[UAX #29 Text Segmentation](https://www.unicode.org/reports/tr29/)。用于用户感知字符边界；实现需固定所用 Unicode 数据版本。
- [R05] ttf-parser 官方 API，[math module](https://docs.rs/ttf-parser/latest/ttf_parser/math/index.html)。可复用字体表读取，具体依赖版本在 P02/P05 锁定。
- [R06] HarfBuzz 项目，[rustybuzz](https://github.com/harfbuzz/rustybuzz)。作为文本塑形基础库候选，不代替二维公式布局。
- [R07] Rust 官方，[wasm32-unknown-unknown](https://doc.rust-lang.org/rustc/platform-support/wasm32-unknown-unknown.html)。用于核心宿主依赖与工具链设计，不能假定本机 std 文件/线程能力在 WASM 完全可用。
- [R08] AxMath，[Equation Editor](https://axmath.gitbooks.io/axmath-docs-en/content/3._equation_editor.html)。只作为公开交互行为参考，不作为内部源码/技术栈证据。
- [R09] W3C，[EditContext API](https://www.w3.org/TR/edit-context/)。作为自绘输入区与平台输入服务的参考；实际宿主支持需验证并保留合适通路。
- [R10] CodeMirror，[System Guide](https://codemirror.net/docs/guide/) 与 [Reference](https://codemirror.net/docs/ref/)。用于事务、位置变化与源码视图接入，不将它替换为数学排版核心。
- [R11] W3C，[SVG 2 Paths](https://www.w3.org/TR/SVG2/paths.html)。用于路径绘制后端；Office 可接受子集仍以现有消费者为准。

---

维护约定：每完成一个阶段更新独立阶段报告，不把本计划中的待实现项改写成已经实现的事实。任何 Office 契约例外必须先记录原因、最小范围、回归风险与批准依据。

