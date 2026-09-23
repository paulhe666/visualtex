# ADR-001 — 自研核心第一切片与冻结边界

日期：2026-09-13。状态：隔离实现采用；不是生产默认切换批准。
依据：VISUALTEX_ENGINE_PROJECT_PLAN.md、当前 P00/P01 检查结果。

## 运行边界

单一 Rust workspace `engine/`，仅 core / wasm / test-cli 三个 crate。core 不依赖 Tauri、DOM、Office、网络和系统文件接口；WASM 绑定同进程调用，CLI 负责原生测试的 stdin/stdout 与资源读取。不得逐按键经过 Tauri IPC，不向 Word 注入新 DLL。

Native 与 WASM 使用同一 EngineSession/操作代码。绑定协议 `visualtex-engine-v1` 与 Office ProtocolVersion=2 无关系。WASM 实例通过单调 handle 管理会话；销毁后旧 handle 必须报错，不能重新借给另一个会话。边界复制输入/输出，JS 不长期保存线性内存指针。

## 状态与身份

EngineSession 是唯一写入点，拥有数学结构、源码草稿、选区、事务历史和 revision。documentId 由宿主注入；NodeId 为文档内单调 ID，不使用文本 hash，不复用已经发出的 ID。Office formulaId、Session id 与 NodeId 类型和空间分离。

所有内容修改须提供 expectedRevision；失败原子回滚，过期结果不能落到当前活动文档。位置以 Row 边界或 TextRun 内 grapheme 边界表示，不使用 MathLive offset，也不能把 LaTeX 字符数当数学位置。位置变换独立返回；撤销恢复原节点身份、结构、草稿及选区。复制节点重新分配身份。

历史记录结构事务及前后结构快照/差量，不通过前后 LaTeX 字符串猜操作。光标移动不解析，不增加内容撤销记录；布局缓存只因内容/资源变化失效。草稿进入同一历史。新内核内部状态不写入 Office v1 metadata。

## 源码与最小语言 K0

K0 范围：Row、普通符号/数字、TextRun、Fraction、上下标、EmptySlot、UnparsedSource。保真保存和规范化导出分开。注释、空白、控制词拼写、参数边界和未改子树保留；修改某子树不美化相邻子树。

未完成输入保留原始草稿及诊断；可显示最后有效结构，绝不自动补齐后反写源码。未知命令保留原文并报告 unsupported，不执行 TeX 程序，不调用 MathLive/MathJax。能保留源码或显示占位不等于完整可编辑支持。

UTF-8 byte offset、JS UTF-16 code units、Unicode extended grapheme indices 是不同概念。转换显式检查边界，拒绝落在 UTF-8 字符/UTF-16 surrogate 中间的位置。文本编辑按固定版本 Unicode segmentation 数据处理。

CompositionSession 记录起始 revision、目标位置和原选区；预编辑不改正式树，确认恰好一笔历史，取消零次内容修改。移动目标/收到旧组合结果时拒绝，不盲插入新目标。真实系统 IME 是 P06 单独门槛，纯状态测试不能替代。

## 数值、字体与场景

内部几何以 1em=1 的 f64 表示，不在每步取整；输出边界量化到 1e-6 CSS px，禁止 NaN/Infinity。Native/WASM 几何绝对比较容差 1e-6，自主位图抗锯齿不要求跨宿主逐像素相同。

字体字节、face index、资源 SHA256 和样式进入布局输入与缓存键。字体轮廓/MATH 数据读取和文本塑形可用基础库；二维数学布局、结构光标与绘制指令必须由本核心实现。不能排好后改字体或用 DOM 测量拼公式。

LayoutSnapshot 统一提供 DisplayList、CaretStops、HitRegions、baseline 和辅助技术投影。视口映射为 `layoutPoint=(clientPoint-viewportOrigin+scrollOffset)/zoom`，DPR 不进入自然布局。光标闪烁只改叠加层。

资源限额必须覆盖输入字节数、节点数、嵌套深度、会话数、历史长度和字体输入；超过限制返回结构化错误，不保留半次编辑。WASM trap 后由宿主保留检查点并重建实例，不声称能捕获一切 OOM。

## Office 冻结与验证层级

保护基线：`checkpoints/20260913-P00-provenance.json`。
契约快照：`checkpoints/20260913-P01-contracts.json`。
冻结的是工作副本而非 HEAD。所有生产路径默认禁止修改；例外必须单独记录阶段、理由、授权、实际验证，不使用目录级宽泛豁免。

输出仍为 svg/svgBase64/mathMl/pngBase64/width/height/baseline；96dpi CSS px → 72dpi pt 乘0.75；baseline 从 SVG 顶边向下。保留 MathML-only OMML 快路径、MathType ruler 乘24单位和会话保存队列。新核心不能借改消费者安全策略获得通过。

P00 当前缺少安装版 UI/性能和完整包内来源证明，P12/P16 门槛关闭。仅批准隔离实现，不凭核心测试宣布 Office 已验收。macOS/Web 保持未验收。

## 验收和性能证据

首先在旧实现运行契约测试，再注入字段改名、单位更改、鉴权改变的负向样本，必须失败。能力矩阵逐项区分 parse/edit/layout/LaTeX/SVG/MathML/Office；E01—E22 沿用总计划完整清单，不删未知命令以提高覆盖率。

参考目标：200节点热资源局部输入到可见更新 p95≤32ms，光标操作 p95≤8ms；这不是已达到的成绩。内核和 Office 全链路计时分开。P02 使用同一请求样本在 Native CLI 和真实 WASM 执行，不能以编译成功冒充运行通过。

## 已核验的一手参考

Rust 官方 wasm32-unknown-unknown 平台文档（文件/线程等系统服务限制）、wasm-bindgen 官方运行测试指南、Unicode UAX #29、Microsoft OpenType MATH。2026-09-13 已重新读取；实际依赖/Unicode数据版本随 Cargo.lock 固定。参考标准不代表本实现已经具备对应功能。
