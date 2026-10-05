# 三种公式六向转换性能与优化（2026-10-05）

**后续CP287–290已完成20/40扩展性复测，见[独立扩展性报告](conversion-scaling-stage-20261005.md)。** 该轮36次/1080条转换，固定吞吐20（替换原3个失败围栏）重复为40；五向约线性，OMML→MT为2.258倍。样式复用实验只有局部收益、整批未证实，已仅撤回该实验；原CP282编号共享优化及CP284修复保持。CP290最终FreezeR3 Word与CP284同SHA0DE80AE7…B36E6C，不是新增全保真验收或默认脏树构建。下述CP280–283历史数据保持原范围。

状态：本轮性能基线、共享编号读取优化及对比测试已完成；17条子集六向各旧/新3次，共36次转换，另有完整20的前后各4向尝试。不是全20六向保真验收。用户明确授权六向性能测试与优化，原性能冻结在此范围解除。先完成本地提交`a57e568d5295dc67490e30f4169cd5d05c7ce7c5`（31文件，未push）；此前残余脚本/产物/现场证据保留，不是净树。随后优化未提交。

以OMML、VisualTeX OLE、MathType OLE六向为范围。固定representative20原20条LaTeX，记录最新导出MathML及来源哈希；已知F02/F09/间距/固定组重音等失败不排除、不改比较器。性能成功与保真成功分开报告。

旧word-format-conversion-render-matrix入口已由CP264封锁；WordPerformanceHost还会切换COMAddIn.Connect，不能复用该副作用。新入口必须明确/x /Automation自有PID与文档，加载源DLL且不修改用户加载项设置。held R4继续冻结，C#由Build-Candidate/FreezeR3构建。Word/Companion启动、源创建、实际转换、保存/重开与离线比较分项；不将Word预读后的耗时称冷启动，不把本地跨进程回调称用户输入延迟。

证据：`artifacts/conversion-performance-20261005/`。本阶段尚无六向性能数字或优化收益结论；先建可复现基线，再定位共享调用链瓶颈，优化后同条件复测并保留失败项。禁止安装/部署/注册表/安全机制/用户现有Office操作。

## CP280 基线环境建立

本地提交后已新增registry-free模式`word-owned-six-way-performance`；WordPerformanceHost要求显式自有根目录，并删除COMAddIn.Connect读写/恢复分支。沿用/x /Automation+PID/自有空文档核对工厂；每个case前后只读核对Word/PPT加载项LoadBehavior。新入口仍走生产ThisAddIn六向命令、实际Companion/Word，不使用合成预览代替目标渲染。

已构建C# `build-cp280-performance-baseline`，Word SHA256 `3A31648D862F127CD5C76B02127F641E5791DF749D62FA5AB0C89698DF81CE2F`，FreezeR3/heldR4保持。Companion由当前Rust源码新编译，使用rustc显式-o写入`runtime-baseline/visualtex.exe`，原有被四个旧实例使用的target/debug/visualtex.exe保持`81184AB7…5891`。测试专用UI-root覆盖只在isolated_acceptance_runtime生效，读取复制到本证据目录的CP278最终前端，不替换旧dist或影响已有实例。

专用Companion PID59760、端口43136健康启动；runtime.json在`.cache/companion-source-acceptance/perf-cp280-baseline/`。原20条LaTeX经当前实际latexToMathMl全部导出成功，输入旧manifest哈希写入`corpus-current20.json`。新性能入口把源已损失信息单列sourceDifferences，把每项内容/样式失败收集并明确置fidelityPassed=false；没有放宽既有比较器。定时仅包含生产转换调用与完成等待，源预读、保存、重开、验收另计，不能称冷启动。

正在创建首份20条VisualTeX源文档；此处仍无方向性能结果或优化收益。

## CP281 完整20条首轮基线与当前失败边界

同一cp280候选、Companion59760/43136、固定UI下，VisualTeX与OMML各20条自有源创建成功。VT→OMML实际转换14,141.62ms，20条写入；转换后及重开各F02/F05失败（显式斜体运算符），不得计保真通过。OMML→VT转换14,714.57ms，20条写入及逐项持久LaTeX重开通过；原OMML相对当前输入有8项sourceDifferences，不能算恢复原始保真。

VT→MT与OMML→MT完整20条均在原生MTEF预检F07停止，0目标/20源保留；分别1,469.62ms及结果JSON所记失败延迟，不是成功吞吐。当前F07/F08/F09伸缩结构严格失败早已在CP224/225记录，旧CP217的“排除F09即19条”不适用。口头进展一度沿用旧编号，已核对原公式后更正；草拟corpus-performance19.json未运行且保留NOT-RUN说明。后续性能可执行子集明确为17条（排除F07/F08/F09），全20失败不关闭，两个MT源方向尚未建立完整20源前提。

证据为本目录baseline-{vt-omml,vt-mt,omml-vt,omml-mt}-r1内performance-result/operation-metrics/conversion-trace及外侧provenance。VT→OMML主要指标：替换5.08s、准备4.80s（其中局部编号2.58s）、编号刷新2.16s。OMML→VT替换7.02s、准备2.79s、源捕获2.56s；MathML投影仅13ms。指标含嵌套，不能相加。当前测量是源DLL在验收进程中调用隔离Word的COM跨进程实际链，尚不能等同安装版进程内最终用户延迟，也无像素验收。

下一步固定17条六向可执行基线，审查共享COM调用与重复读取，按证据做小范围优化后同条件复测。未改保真比较器或生产转换规则；用户原进程与LoadBehavior前后核对保持。

## CP282 共享编号读取优化与首轮复测

17条基线已齐（毫秒）：VT→MT12020.74、VT→OMML11169.97、MT→VT10706.45、MT→OMML13476.06、OMML→VT12295.74、OMML→MT14439.71。每向17写入；两个OMML目标各F02/F05在转换后/重开严格失败，其他四向逐项通过。OMML17源使用baseline-vt-omml17-r1实际保存输出，源2项差异保留，不与直接创建OMML20的8项差异混同。

实测OMML编号Apply约1.6–2.2s；代码证据是IsCurrentField内部已经读取并筛选Code/Text，而native编号立即重新读取一次；MathType编号和native REF消费者也丢弃WordNumberingFieldCollection已有Entry再读取。修改4个生产文件：WordNumberingFieldCollection公开内部拥有式TryCaptureCurrentCode、HasVisualTeXNumbering消费Entry；MathTypeEquationNumbering.UpdateNumberFields消费同一Entry；WordNativeOmmlNumbering识别/刷新复用当前代码并将结构读写核心接受已经持有的OMath；WordNativeOmmlFieldRefresh使用保留OMath及REF Entry。保持全部修订过滤、唯一SEQ/STYLEREF、数学结构/范围/字段结果边界检查；没有跨异步缓存、全文新索引、符号补丁或比较器更改。

frozen-R3候选build-cp282-performance-fields-r1完整C#构建通过，Word SHA256 BD567E89984B859384EF29F78FFFB3F5BC054D36ADAC351CCD85045ECCF20157；heldR4不变。编号相关离线111/111（0skip）通过；新增安全入口word-owned-field-read-check复用实际Word修订探针，通过9项原始/已删/新插字段、快照和保留OMath刷新断言，fields-revision-r1保存原始COM字段记录。离线测试不能替代这个实机边界。

优化六向首轮复测进行中。离线Inspect-Results.py独立读取保存DOCX的实际字段结果与普通段落文字，已完成的17条方向编号1..9及全部前后普通文字保持。初次速度变化较小且部分方向有噪声，尚不宣称稳定收益；接下来交替旧/新候选重复测量，保留所有样本而非选最好一次。产物仍非发布包，未部署或提交优化改动。

## CP283 三轮配对结果及最终范围

每个方向固定同一已保存源文档、同一Companion/前端/进程外Word调用方式，旧/新各3次；第2轮旧→新，第3轮新→旧。每次是新自有Word进程，源预读发生在计时前。下表单位秒，正百分比表示中位耗时减少，负值表示增加；这是小样本观察，不是总体置信区间。

| 方向 | 基线中位数（最小–最大） | 优化中位数（最小–最大） | 中位耗时减少 | 本轮17项保真范围 |
|---|---:|---:|---:|---|
| VisualTeX→MathType | 12.021（11.977–12.089） | 12.101（11.703–12.431） | -0.67% | 转换/重开逐项通过 |
| VisualTeX→OMML | 11.435（11.170–11.573） | 11.374（11.313–11.441） | 0.53% | F02/F05斜体严格失败 |
| MathType→VisualTeX | 10.280（10.240–10.706） | 10.299（10.254–10.821） | -0.19% | 转换/重开逐项通过 |
| MathType→OMML | 13.464（13.134–13.476） | 12.918（12.907–13.138） | 4.06% | F02/F05斜体严格失败 |
| OMML→VisualTeX | 12.296（11.853–12.474） | 12.054（11.601–12.107） | 1.96% | 持久LaTeX通过，源2项损失仍在 |
| OMML→MathType | 14.440（14.301–14.919） | 14.069（13.914–14.454） | 2.57% | 读回通过，源2项损失仍在 |

实际优化集中在编号层：VT→OMML编号Apply中位1,558.5→1,347.5ms（13.54%减少），MT→OMML为1,570.9→1,332.8ms（15.16%减少）。其他大段Word准备/插入成本保留；整体只有小幅变化，两个方向略慢且在各自波动范围附近，不能称六向全面提速。保留这次共享去重是因为它减少确定的重复COM访问及重复定位，未引入额外缓存；不继续通过删内容校验追数字。

所有36次子集转换各17写入；每次转换后和独立重开比较均执行。4个native目标方向每轮34份读回MathML，旧/新3轮共408份逐字相同；两个VT目标仍按持久LaTeX逐项核对。36份DOCX独立解析编号1..9、普通正文和段落数全部保持。失败集始终仅两个OMML目标的F02/F05，没有把源已有损失判为恢复，没有新增公式身份/保真通过主张，也没有像素验收。

证据：performance-summary.json保留每个原始样本及指标（嵌套指标不可求和）；matrix-with-docx-checks.json、native-mathml-baseline-vs-optimized-all.json及全部case目录。最初独立DOCX读数工具只认英文SEQ标签，直接OMML20源使用Word本地化标签，造成假阴性；已改为标准SEQ语法+固定完整序列断言，未改生产比较器或预期，旧工具输出matrix-with-docx-checks-before-localized-sequence-reader.json保留。第一轮串行脚本曾用PowerShell5，失败stderr使两个provenance RunnerExit为空；实际EXE acceptance.log、结果及上层exit=1保留，后续pwsh正确记录，未丢弃性能或失败样本。

最终验收工具候选build-cp282-performance-harness-final完整构建通过，与计时候选Word DLL逐字相同BD567E89…0157；只多了测试SaveAs2显式AddToRecentFiles=false，未重测不同产品二进制。早期两个VT种子的旧Save方法未指定该参数；没有MRU之前快照，不能据LoadBehavior前后相同宣称所有Office设置均未写入。只读观察未获取到MRU分支，未修改用户MRU。全部源码/构建/24个UI文件哈希在pinned-sources-and-builds.json。

完整20最终复核，使用同一BD567E89…0157产品DLL，保留原始输入和源文档：

| 方向 | 最新结果 | 本次转换计时 | 验收界限 |
|---|---|---:|---|
| VisualTeX→OMML | 20写入，转换/重开各F02/F05失败 | 13.324s | 严格斜体差异未修；编号/正文/段落数保持 |
| OMML→VisualTeX | 20写入及持久LaTeX重开通过 | 15.019s | 源OMML8项已有差异；非原始全保真/像素通过 |
| VisualTeX→MathType | F07 MTEF预检失败，0目标、20源保留 | 1.493s失败延迟 | 不计成功吞吐；F08/F09本批未走到，既有缺口未关闭 |
| OMML→MathType | F07 MTEF预检失败，0目标、20源保留 | 3.027s失败延迟 | 同上；源8项已有差异单列 |
| MathType→VisualTeX | 完整20源前提未建立；未运行完整20 | — | 仅上述17子集完成三轮 |
| MathType→OMML | 完整20源前提未建立；未运行完整20 | — | 仅上述17子集完成三轮且F02/F05失败 |

两份20条成功写入输出及两份失败后保留原文，独立DOCX编号1..12/普通文字/段落数均保持。完整20是各1次最终回归，不能据其与较早基线的单次差值断言普遍收益。证据optimized-*20-final/（如optimized-vt-omml20-final/）及final20-summary.json；不覆盖baseline失败。

本轮所有自有Word正常退出；Companion PID59760在核对PID/路径/创建时间/二进制哈希后关闭。environment-after.json确认原Word93828/95356/88900、原Companion46904/42504/70768/78484创建身份保持，4个Word/PPT LoadBehavior分支前后相同，且无本轮Word/Companion残留。未安装/部署/推送/关闭安全机制，未操作用户文档或正常运行的Companion；HEAD仍a57e568。旧target/debug/visualtex.exe未替换；本轮独立Companion SHA256 3C0B358B014C459F8563970947D2470111B18B588B4DC812FE1712875CDA6B90。

结论：共享编号层确定的冗余读取已精简，局部刷新有约14%–15%的中位耗时减少；六向整批仅小幅变化且两向基本持平，不能宣传全面显著加速。保留正确性保护，不继续做缺乏证据的大范围性能重构。未运行：安装版进程内端到端时延、100/1000公式扩展性、不同机器/Office版本、PowerPoint性能、物理操作/Word像素验收。CP255、特殊符号、任意间距/复杂上下标及既有源损失仍未因性能测试关闭；不适合标注为全六向保真发布验收包。下一阶段优先处理F07/F08/F09原生结构和F02/F05样式权威，之后再做完整20六向及安装版真实用户延迟；不要为性能测试绕过严格预检。
