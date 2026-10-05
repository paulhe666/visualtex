# 近期公式与 Office 修改架构审查（2026-10-05）

后续状态：用户已授权根因修改。CP269–CP270修复与最终限定验证见 [根因修复报告](architecture-root-remediation-20261005.md)；下文是CP268审查时证据，未回写成当时已修复。PPT单次Undo、CP255和全量视觉仍未关闭。


状态：本轮只读审查完成；缺陷修复和下列实验尚未执行。结论：近期多项修复具有共享根因依据，但当前工作树不适合直接打成“已验收通过”的包，也不建议原样作为面向普通用户的验收候选。

## 范围与约束

- 基线：`fed066192373fc12364e822f5478acca9967f7b0`。
- 已核实 HEAD：`67e57fef62a4abcfa9577fac8849c53817f83893`，分支 `reference-3d207d7`。
- 审查两层：基线到 HEAD 的已提交源码，以及 HEAD 后的当前修改/新增源码。
- 只读产品与测试代码；仅写审查报告及历史检查点。不得启动 Office、改注册表/模板/绑定、部署、安装或执行 Git 变更。
- CP267 特殊探针和已收尾的文字卡顿调查不重开；性能仅分析调用链与复杂度，不重构。
- 旧失败、未建立的测试前提和环境阻塞分别保留，不能合并为通过。

## 实际覆盖与构建身份

已读取历史检查点 CP194–CP267 共 74 个编号阶段及五份指定阶段报告，结合 Git 差异、关键调用方/被调用方、原始 TRX/JSON、源码和 DLL 哈希核对。Git 基线到 HEAD 只有一条提交 `67e57fef`（截至 CP260 的快照）；当前 PowerPoint、Word 边界、快捷键、验收入口及新增 partial/测试文件另外审查，未因 untracked 而遗漏。二进制产物只核对身份、日志和有关编译方法证据，不声称审计每个二进制字节或供应商全部源码。此为有边界的深度审查，不是全项目不存在其他缺陷的证明。

工作区实际根为 `C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7`。本报告源码行号按本轮结束时文件；下文未写绝对路径的代码均相对于此根。C# 共享规则目录为 `apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared`。

| 身份 | 实际核对结果 |
| --- | --- |
| CP266 Word | `artifacts/omml-edit-audit-20261001/build-cp266-word-shortcuts-r3/VisualTeX.WordVsto.dll` SHA256 `827AAE47EE26B13BE515C72E24B4ED226CE54900C88090DAEEE538953D151CEF` |
| CP267 Word | `build-cp267-reported-omml-r3/VisualTeX.WordVsto.dll` 与上述字节一致 |
| CP263 PPT | `build-cp263-native-final-r2/VisualTeX.PowerPointVsto.dll` SHA256 `5857E55597B9D5352120F8EA80EFE7294EF5CBEBF651DAC237211493EA459BC3` |
| CP263 Word 依赖 | `FB031CF2A98E8C2266C90BF20730606985956220A7610A4CC88EC26763885EF2`，不能与 CP266 后续 Word 修改混称同一验收组合 |
| frozen R3 | `artifacts/mathtype-install-20260927/coexistence-20260928/release-inputs-r3/FreezeR3.targets` 替换一个 WordFormulaHostResolver Compile 项并验证 SHA256 `FCE1BB7534EC0C9BF518C06818FA7EDA6816EBA5A0266D203CD84B8A84EBD045` |
| held R4 当前源码 | SHA256 `6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`，未改，默认树构建会使用它 |

`artifacts/omml-edit-audit-20261001/Build-Candidate.ps1` 通过 CustomAfterMicrosoftCommonTargets 注入上述 targets；CP266 build-state 为 x64/net472、VS18 MSBuild、SDK 8.0.423。本轮未重建。CP263 Kernel、Formatting、NativeStyleAcceptance 和 basic regression 的当前哈希与 cp263-final-state 一致；Program.cs 当前为 `2EBD9DA6…B0B93`，已不是 CP263 当时 `314AC2AB…E5F`，与后续验收入口变更相符，不能把旧 EXE 验收自动归于当前工具。OCR/macOS 两个脚本虽在 status 显示 M，内容 diff 为空，未据此虚构行为变化。缓存、安装调试产物、用户证据及临时文件均保留。

## 已确认的审查前提

- CP255 候选未隔离，现有汇总保留失败；此事实阻止把整个当前树当成已验收发布版。
- CP266 的实际绑定测试不等于物理键盘和编辑器完整生命周期测试，双键前缀的真实宿主前提未建立。
- CP263 的 7 项样式和 10 项基础操作是限定服务/宿主验证；实际 Ribbon 点击为零，任意字体和混合字号不成立。
- CP264 初始启动器曾写 LoadBehavior；最终隔离检查不能倒推初始 Word 值已正确恢复。
- 下列阶段发现仍需与完整覆盖结论一并阅读，不表示其余路径已通过。

## 按严重程度处理的发现

编号沿阶段记录保留。优先级顺序为 **R03 / R01 → R04 → R07 → R02 / R05 / R08 → R06 与冗余项**。P1 表示内容/用户设置风险或交付阻断；静态风险与已经发生的事件分别标注，不把推断写成实测事故。

### R01：P1，PowerPoint 文本框失败路径缺乏原内容恢复（继承风险）

`PowerPointNativeOmmlKernel.InsertIntoText` 207 行先用段落分隔符覆盖选中正文，随后才调用 `Insert`；后者写入线性文本后还会检查命令可用性、选区及样式对应关系并可能抛出。`PowerPointFormulaService.InsertOmml` 854–857 行只删除新建独立 shape；文本框路径中该变量为空，没有恢复 `targetTextShape`。`StartNewUndoEntry` 只建立 Undo 边界，不是自动回滚，见 [Microsoft API 定义](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.startnewundoentry)。触发条件是既有文本框中选区替换/插入后转换或宿主命令失败；可能遗留正文删除、空段落或 UnicodeMath 半成品。这是静态贯通结果，未启动 PowerPoint 或注入宿主失败，不能称已实测数据丢失。建议尽量前移无副作用准备，并为文本框写入提供保留格式的受控失败恢复；不能仅赋回纯文本而丢掉原格式。成功路径的 10 项基础操作不能覆盖此失败原子性。独立新建 shape 的失败清理有效，应保留；CP263 严格样式映射抛错本身也是必要保护，需要补全其上层失败事务，而非改成吞错。

### R02：P2，继承的边界清理删除合法 token 内容，且依赖 XML 词法形式

`WordOmmlConverter.RemoveVisualTeXBoundaryArtifactsFromMathMl`（3135 行）移除所有 XText 中的 U+200B/U+200C/U+2060/U+FEFF，而不是限定已证明属于自有宿主边界的节点；正文 ZWNJ 也会被删。已核对 fed066：该逻辑与未接入的 `CreateSafeXmlReader` 均已存在，不能归为此次新增回归。对已核验 CP266 DLL 仅反射调用此纯函数，未实例化 Office：`<mtext>a[U+200C]b</mtext>` 返回 `ab`，等价 `<mtext>a&#8204;b</mtext>` 保留实体。前两次探针分别因 PowerShell 参数包装与数组表达式构造失败，未建立产品测试前提；修正传参后上述两个输入均成功执行。最小建议：清理归于已证明的宿主边界所有权，保留任意 MathML 文本；用字面/十进制/十六进制等价 XML 的独立内容断言覆盖。此问题不用于解释已搁置的 CP267 栈溢出。

### 性能与保护边界

PowerPoint 原生源水合通过 SaveCopyAs 解包，成本随整个演示文稿增大，但位于显式编辑/转换/字号命令，不能描述为每次被动选区都全量保存。保存前原生 XML 比较、独立对象限制、宿主身份检查应保留。Word 共享转换约有多段 XML 解析/归一化；其中规则是必要的，潜在精简点是重复表示转换，不是删掉内容边界。性能未计时、未重构。

### R03：P1，公式缓存已停用，但旧文本插入范围仍能成为写入目标（继承风险未一并关闭）

`PowerPoint.ThisAddIn.OnWindowSelectionChange:434` 对 `ppSelectionNone` 保留 `_lastTextInsertionTarget`；`ResolveOmmlTextInsertionTarget:553–563` 在无法捕获当前文本选区时无条件返回它。`BeginContextAwareOmmlSession:497` 与 `BeginOmmlTextInsertion:530` 均使用这个结果。CP262 已因空白取消选中与 Ribbon 临时失焦无法区分而删除公式选区回退，但同样的判断还存在于文本目标中。基线核对表明文本缓存逻辑已存在，不能称为 CP262 新引入。

第二个相关窗口是编辑器开启后正文改变：捕获的 `TextInsertionReference` 只存 slideId/shapeId/start/length；`PowerPointFormulaService.ResolveTextInsertionTarget:1994` 只检查对象还存在且有文本框；内核检查偏移仍在当前长度内，不验证原选中文本/上下文仍相同。改变同长度文本仍可通过。触发时可能在已取消的旧选区插入，或覆盖编辑器开启后新写入的文字。当前结论为静态贯通，未启动 PowerPoint 复现损失。

`PowerPointBasicFunctionsAcceptance:186–214` 的取消选中用例仅检查 `ResolveFormulaSelection` 和 `_lastFormulaSelection`，不能覆盖文本创建入口。最小修复方向：只允许有明确来源和选区代际的 Ribbon 交接，不把任意 None 当作交接证明；编辑会话保存可验证的源内容/格式边界，提交前发现变化则拒绝写入。不能只检查长度，也不能为此完全禁用正常 Ribbon 插入。后续最小实验为“选中文本→取消选择→插入”和“开启编辑器→同长度修改源文本→保存”，各自独立断言前后正文及格式。

### R04：P1，旧验收默认入口仍会修改用户级加载注册表，恢复失败被吞掉

`VisualTeX.VstoFlowAcceptance/Program.cs` 的 `UserOfficeAddInAutoLoadSuppression.Create:135` 创建 HKCU add-in 键并把 LoadBehavior 写为 0，`Dispose:171` 尝试恢复但捕获后静默；进程异常终止也不会执行 Dispose。主入口的普通分支仍会创建这些 scope。CP263/CP264 后来的专用入口、CP267 直接入口规避了它，但没有消除其他模式的默认副作用。CP264 的实际误用及初始 Word 值缺少快照，是已发生事件；最终一致性检查不能补出这段历史。

最小建议：把纯自有实例入口作为默认，任何用户级状态修改须单独显式授权；有授权的工具也必须先落盘完整快照并显式报告恢复失败，不把失败吞成成功。该项针对验收工具及后续使用风险，不是证明产品每次启动都写这些值。本轮未调用 Program.Main、未创建 Office 实例、未触碰注册表。

同类工具边界：新增进快照的 `docs/remediation-3d207d7/start-real-candidate.ps1:31` 临时写 HKCU CLSID/InprocServer32 CodeBase，作用域是用户而非自有 Word 进程；finally 删除 override 仍不能隔离并发启动或异常退出。`probe-inline-mathtype-alignment.ps1:29` 的 Inventory 分支会 Save，文件名不等于无写入；`capture-real-word-window.ps1` 只有旧文档路径/名称判断，截图前没有维护脚本那样的前台遮挡检查。未提交 `.devspace-capture-formula-20261002.ps1:24` 在证明前台前发送全局 Alt，不能复用于当前“不碰用户进程”验收。以上均只读审查，未运行；不能从其留存在仓库推出本轮发生了相应副作用。最低限度应把它们标为历史专用入口、避免通用自动重跑，并统一到显式所有权入口。

### R05：P2，解析深度/元素上限只在无调用者的 helper 内，公开转换链没有相同保护

`WordOmmlConverter.CreateSafeXmlReader:24` 和 `MathMlToLatexConverter.ParseMathMlSafely:13` 定义了深度/元素上限，调用搜索各仅有定义。实际 `MathMlToLatexConverter.Convert` 和 `WordOmmlConverter.ValidateMathMlForOmml` 使用另一套 reader；字符数/DTD 限制存在，不能误称完全不限制，但没有相同的树深度限制。部分清理还发生在 Validate 之前，后续转换有递归。

在已核实的 CP266 DLL 上，仅反射调用纯解析/校验方法：260 层 mrow 的小输入被 `ValidateMathMlForOmml` 接受，私有 `ParseMathMlSafely` 则按 256 限制拒绝。一次 PowerShell 语法错误未执行探针，修正后结果有效；没有运行完整深递归转换、没有试图制造崩溃。建议统一公开入口的受限解析，放在递归及重写前，并复用解析结果；不要新增多次全量预扫描。测试应通过真实公开边界验证，而不是仅验证未接入的 helper。该发现不解释 CP267 两次未定原因的外部运行器栈溢出，后者继续搁置。

### R06：P2，被动 PowerPoint 读取仍写元数据，已失去消费者的缓存读放大了职责混杂

`ThisAddIn.OnWindowSelectionChange:443` 为已无读取者的 `_lastFormulaSelection` 调用 `PowerPointFormulaService.ReadSelection`。后者 `129` 行会对自有元数据公式执行 `EnsureUniqueFormulaIdentity:1328`；即使 owner 已匹配也调用 `Configure:1887`，写 shape.Name、Tags、AlternativeText 并设置 LockAspectRatio。缺少旧 owner 时还可能扫描全演示文稿查重。因此“所有被动 PPT 读取均无修改/无全局扫描”不成立；原生 OMML 的延迟水合改善不覆盖这条元数据公式路径。

复制后身份唯一性保护有价值，但不宜藏在纯读取 API 中，更不宜为一个不再使用的缓存执行。建议后续授权时移除该无消费者读，并把必要的身份修复放在明确、可报告、可撤销的操作边界；不得简单删除防重复身份逻辑。未实测这里导致 Dirty 标志、Undo 变化或用户卡顿，因此这些影响列为待验证，未写成已发生故障。

### R07：P1，交付阻断：CP255 仍活跃，且已验收候选与默认树并非同一源码组合

`MathMlAccentGeometry.cs`、`MathMlToLatexConverter.ConvertOver`、`MathTypeMtefCodec.EmitOver:3815` / `EmitGlyphAccentTemplate:4287` 的宽度规则仍直接参与转换，没有候选隔离。真实 TRX 显示 CP255 的 2325 项为 2324 过/1 失败；CP257 的 2338 项为 2337 过/1 失败，均是 `EverySingleCharacterDecorationRetainsItsNativeEncodingBytes`。这是规则改变与旧字节契约的未决冲突，不能擅自判旧测试必然错误而改绿。固定组重音垂直位置问题另有原生证据；CP253/CP255 对应载荷已有相同结果，不能把整个旧位置缺陷归为 CP255 新增。

40 个定向检查、8 个源码检查及生成出 6 个原生结果，不能覆盖这两个问题。held R4 虽已在快照中，但候选一直替换为 frozen R3。最小处理：先明确哪个重音原生表示有完整字体/尺寸/模板和宿主视觉依据，再决定修复或有范围的候选隔离；不能回退整个含已通过颜色/样式修复的文件。打包必须声明并验证精确 resolver/Word/PPT/前端/Companion 组合，而不是仅记录 Git HEAD 或 bin 文件。当前无新安装包验收，本轮也不构建。

### R08：P2，已知原生能力边界未成为完整保真保证：字形可显示性和间距降级

`MathTypeMtefCodec.MatchesExplicitFontSpec:3250` 对显式 mi 样式的匹配没有普遍的字体槽位覆盖检查。CP246 的 script 小写 l 已表明：原生字符/样式可由同一 codec 正常读回，选用字体却未提供可见字形。共享读写闭环不能独立证明原生渲染正确。最小方向是按已验证的字体/编码/槽位建立能力边界，缺失时保留为不支持或明确降级，不能把 script-l 改成 ell 或统一放松签名。

`MathTypeMtefCodec.EmitMathSpace:2819` 对未证实宽度沿用 ASCII 空格，读回成为 0.2em；`MathTypeNativeSpaceContractTests.UncharacterizedWidthsRetainExistingFallbackWithoutClaimingIdentity` 明确测试这个旧回退，不是任意/负间距通过。OMML 的 `ProjectMathMlSpacesForOmml:3351` 返回 Capabilities/AuditLimits，但 `TransformMathMlToOmml:3101` 只取 MathMl，未将能力结果交给调用方。已有 OMML 测试明确记录某些不支持维度被 stock XSL 忽略。不能以“有 capability 测试”声称产品会阻止或报告每次损失；也不能据底层 MTEF 回退直接断言 Word 所有写入都会接受，因为写入前严格签名可能拒绝。

建议让受支持集合与调用方可见的转换能力一致，保留现有最佳努力行为的明确边界，优先补独立原生输出/样式证据；是否要改变不支持输入的产品交互需另定，不能本轮擅自改成一律拒绝。F02/F09、U+2AFD、单边上取整、group/arrowvert、其他括号/变体、复杂脚本/基线仍按阶段报告的具体用例状态处理，没有全量通过结论。

## 关键反证与原始结果核对

- 实际编译清单：`VisualTeX.WordVsto.csproj:21` 排除 `WordNativeOmmlSemanticComparer.cs`，验收 EXE 和测试项目才链接它。它的“生成期望 OMML + 反向比较”不是日常 Word DLL 运行成本。产品中的 MathType 写入前 readback/签名验证和宿主范围/身份检查则是真实运行保护，不能混为一谈。
- `WordOmmlConverter.ComputeOmmlFingerprint` 的蓝色/红色纯 OMML 探针对 CP266 DLL 得到同一 `omml2:46c08e0421e16bb81dcccc94326d1d5e6312fa09058009b086dc1f0218c87ba9`，与源码删除 w:rPr/ctrlPr 的内容指纹定义一致。进一步调用搜索发现旧 `RefreshForVisualTeX*` 快捷返回无当前 Word 产品主链调用者；测试及 `WordLatexRedrawAcceptance:1020/1347` 仍调用它，不能称其全仓无调用者。实际 `WordFormulaHostSemanticReader.ReadOmml:35` 从原生 XML 重新转换。因此**不能**据旧 helper 推断当前改色重开会返回旧缓存。此反证保留；使用旧 helper 的验收也不能证明当前主链或颜色保真，内容指纹不得代替颜色/重音控制样式断言。
- `PowerPoint.ThisAddIn._lastFormulaSelection` 当前只有赋值和清空，没有读取；移除回退后留下了事件读选区及缓存维护。属于确切冗余，不是确认的卡顿根因。应在后续授权精简中一并审查事件处理器其他职责，不能直接取消所有选区订阅。
- `MathTypeMtefCodec.MaterializeExplicitCharacterStyles:3090` 在每个 token 内重复 `ReadNativeCharacterStyleBindings(sourceMtef)`，同一函数 3037 行已经读取该源；外层 2900 行也读取。可复用一次操作内不可变定义表，不必加跨操作缓存。这是 O(样式 token 数 × 定义记录数) 的重复解析。
- 原始 TRX 只读核验：CP217 首轮 1983/9、修正后 1992/0；CP255 retained 2324/1；CP257 aggregate 2337/1，两个后者均为 `EverySingleCharacterDecorationRetainsItsNativeEncodingBytes`。本轮未运行汇总，不能把 CP266 构建成功改写为当前树全绿。
- 原始 CP263 `native-styles.json` 明确将任意逐字符字体/混合字号标为 passed=null、前提未建立；其 7 个受支持检查及基础 10 组不能扩大解释。CP266 final-r8 summary 的 9 项实机绑定检查逐项保存 ContextRestored=true，同时 UsesPhysicalInput=false；模拟路由结果 UsesRealWord=false。CP267 独立 exact audit 的 8 项均通过，但 SyntheticPreviews=true、PhysicalUi=false；两项外部运行器崩溃保持原因未定并搁置。

## 共享根因、局部补丁与值得保留的设计

| 范围/关键文件 | 判断及保留理由 | 仍需约束 |
| --- | --- | --- |
| MathMlTokenStyleContract、MTEF style materialization | mi/mo/mtext 身份、显式与默认样式、原生 typeface/style 分开处理；普通 mo 不被统一当斜体。字体变化后使不再适用的原生编码失效，比追加符号码位替换合理 | MathML 意图、原生可编码、宿主实际表现是三层，R08 未关闭 |
| NativeGlyphAuthorities、编码恢复 | 字体名/编码/CHAR 槽位/Unicode/角色共同判断，已有 WMF/原生字节及反例依据。增加真实编码表数据不等于逐案例分支补丁 | 不能由若干字形权威行扩展为全字体或全部变体通过 |
| MathSpaceContract | 原始宽度精确解析，区分源码表达和不同后端能力，避免浮点近似把 1.000…001em 合并到 1em；已知原生 1em/2em 身份保留 | 降级通道见 R08；不得称任意宽度已修好 |
| MathMlExplicitFenceContract 与各后端投影 | 源 INNER/OPEN/CLOSE 结构识别共享，后端模板能力分离；CP223 的回归由 CP224 收窄，端点有不可保留样式时不盲目物化 | 单边括号、group/arrowvert 与原生崩溃仍需独立证据，不靠改变 token 身份规避 |
| WordOmmlConverter 固定 run / BuildUp | 解析好的原生 OMML 不再被 BuildUp 二次重解析；线性输入仍保留必要 BuildUp。固定 run 与结构化函数识别依赖真实 m:r/m:t / 角色，而非公式字符串特判 | 不可在所有入口无差别删除 BuildUp；旧源已丢失样式时不能重建成虚假的完整往返 |
| Word host resolver/reader/mutation kernel | 局部宿主定位、纯语义读取、写入事务分层。范围/故事/唯一宿主、源版本、编号/段落/书签保护有明确目的；先准备后删除、失败恢复并验证应保留 | held R4 扩展完整本地 OMath 的方向可理解，但未替代 R3 验收；不能因读次数多删边界证明 |
| CP264 WordFormulaService.NormalizeTypingCaretAfterInlineFormula | 先证实精确尾边界再读自有 OLE payload；普通正文/前边界不再为无用 descriptor 请求 OLE/布局；原生 OMML、第三方 OLE 不进入被动修复 | 自有尾边界仍可能读 payload 和所属段落，不是零成本；未证明所有用户文字卡顿已解决 |
| PPT NativeOmml partial/package/bridge | 显式命令内 SaveCopyAs 取原生包；被动发现只返回 slide/shape 描述。原生目标不借旧 FormulaId 找别的对象；提交前比较原 XML；混排/多公式整对象替换拒绝，保留邻接文字与 alt/title | R01/R03 是文本创建另一链，不能被原生替换的源版本检查覆盖 |
| CP263 PPT NativeOmmlFormatting / Kernel | DrawingML 显式不透明 RGB 只在 PPT→Word 边界适配；raw snapshot 不改。保留 UnicodeMath 数学字母 scalar，颜色与 upright 映射按准确区间和一处 indexed MathZone 操作；整个序列不符则抛错 | 任意字体/混字号、主题色/透明度、结构笔画颜色、补充字符/组合字符 upright 未完整验收；不得全框统一 Font 来“修好” |
| CP266 WordFormulaShortcutPolicy + shared hook | Word 原生绑定优先，只接受 Nil；未知/禁用/前缀/读失败均让行。查 document/attached/Normal/global context 后 finally 恢复。焦点/前台/同线程约束、AltGr/Win/重复/keyup/peek 过滤有依据；回调失败不再吞键 | 35 模拟和 9 绑定不证明物理键盘；前缀历史只是保守推断，非 Word 当前前缀状态；异步 launch 接受不等于编辑器最终成功 |
| explicit-math-font、role-preserving-mml-visitor、vite.mathliveFontScope | 显式字体作用域覆盖真正 token；保留 intrinsic 数学字母身份与 OP 角色；boldsymbol 在延迟参数解析后组合，Set 去重避免重复应用。移除旧大写希腊特例，属于共同作用域修复 | 精确源码锚点的 MathLive 构建适配需锁版本和升级契约；不能把所有依赖源码替换当成稳定公共 API |
| EditorWorkspace.tsx 颜色弹层 | 复用 PortalOrInline，桌面 body 挂载及相应锚点，Office 保持原挂载；只在打开时订阅 resize/scroll；外点识别 portal；Esc 关闭阻止透传删选区。这些是布局/事件根因修正，无需再引入全局弹层状态机 | 浏览器构建验收不等于原生 WebView2/Office 生命周期；原生系统取色器未测 |
| Rust core symbol_catalog / layout / presentation_mathml | alternate glyph 的字体槽位与语义 Unicode 分离，mathrel/mathbin 进入 AST 原子类别；AMS-only 字体不会切换普通字母配置，缺字明确失败 | data-mjx-texclass/alternate 是传递输入语义的注解，只有后端实际原生表示并独立读回才可计保真；不能凭注解证明 MathType U+2AFD 有可靠字形 |
| vendored tauri-runtime-wry Context | Arc 使跨线程 AppHandle clone 不再递归 clone 内部 Tao Rc；去掉 DispatcherMainThreadContext.Clone，改所有权而非符号/公式路径。崩溃栈、旧版失败及新版三轮克隆回归支持共享根因 | unsafe Send/Sync 仍依赖派发契约；退出后最后一个 Arc 在哪个线程 drop 尚无专项证据，见下面实验。不能把任意未来崩溃都归此处 |

私有 OMML projection 属性只在转换中使用并被消费，不是把丢失藏在最终文档 metadata 的授权。MTEF provenance 可用于未编辑字段的保字节回写，但不能替代实际字符、字体、角色、颜色和位置证明。本轮没有发现新比较器通过统一 mi/mo、丢弃所有斜体或 script-l/ell 等价化来掩盖新失败；相反相关差异收紧了这些比较。存在的能力缺口和旧宽松 helper 单独记录，不能据此泛化所有测试可信或不可信。

## 真实调用频率、复杂度与可精简项

以下是源码复杂度和调用位置，不是新性能测量。n 为本次公式节点/字符数，d 为 MTEF 定义记录数，t 为模板数，s 为演示文稿 shape 数。

| 调用链 | 频率/成本 | 审查决定 |
| --- | --- | --- |
| Word selection event → NormalizeTypingCaretAfterInlineFormula | 每次相关选区通知；先 ±2 字符局部候选。CP264 前边界多余的激活/布局请求已移除；尾部 RestoreInlineCaret 仍可查所属段落 | 保护精确尾边界和第三方隔离；不把它描述成每次键盘输入全文扫描 |
| PPT selection event → ReadSelection → EnsureUniqueFormulaIdentity | 元数据公式被动选区可写 Configure；旧 owner 缺失时 O(s) 查重。缓存结果无消费者 | R06 是最明确的可删调用；身份修复仍需迁到明确命令，不能只删唯一性规则 |
| PPT 显式原生 edit/convert/font → NativeOmmlPackage | 每次显式命令可能保存整个 PPTX 并解包，随整个文件大小增长；不在普通原生选区事件中执行 | 当前捕获版本正确性优先；仅测清楚后考虑操作内快照复用，不加长期可陈旧缓存 |
| Word MathType prepare → codec 写 → compound storage → codec 读 → signature | 显式写入前真实运行时路径；宿主写后还有 materialized 内容验证。redraw preflight/实际写入各有入口 | 两次保护针对生成物与宿主物化，不能因看似重复删除；若复用必须证明相同输入/对象/版本 |
| WordNativeOmmlSemanticComparer | 产品 csproj 排除，离线/验收链接；可含前向构造及双向转换 | 不计入生产每次转换的成本。离线大量测试不说明产品慢 |
| MTEF MaterializeExplicitCharacterStyles | 同一 sourceMtef 外层 2900、函数 3037 及 token 循环 3090 重读原生 style 表，O(n×d) | 值得后续复用一次操作内不可变表；不需要跨会话缓存，当前性能冻结未改 |
| WordOmmlConverter.Transform* | 多个字符串→XDocument→字符串阶段；不同规则各自 O(n)，部分递归/脚本子树搜索有重复遍历 | 优先统一一次受限解析与阶段类型表示，保留独立规则和顺序契约；不因阶段多直接宣称指数复杂度 |
| PPT WordLinearizePreparedOmml | 每个文本元素跨 COM 读 Range/字体/颜色，O(n) 调用；写入需要样式映射 | 明确命令成本，不能算被动文字延迟。批量读取是否等价待实验 |
| PPT ApplyNativeUprightLetters:451 | 按元素循环 List.Contains，最坏 O(n²) | 未来操作内 HashSet 即可，勿加全局状态；目前无计时证明它是主瓶颈 |
| shared keyboard HookCallback:358 | 初始 Ctrl/Alt keydown 读配置；普通无修饰键不读配置。命中后 Word context 检查 O(t)；修饰键过滤晚于配置读取 | 配置读取与重复 context 可计量，绑定保护必须保留；不回退硬编码 Ctrl+D/F 黑名单，不凭模拟次数推断卡顿 |

确切遗留代码：PPT `_lastFormulaSelection` 写而不读；Word `WaitForMathTypeOleMaterialization:7412` 无调用者，其唯一使用的 `MathTypeOleRoundTripMatches:7505` 会去空格和花括号，但当前活动写入使用严格 MathML signature，因此**不能把该旧函数报成现行比较器放宽**。安全解析 helper 未接入见 R05，应统一而非继续并存。旧 Refresh helper 有测试/验收调用，不能当无调用者直接删除。清理这些项均留待授权，未在审查中实施。

异常策略应按语义区分：选区通知、可选 UI 清理失败让 Office 继续运行是合理防御；快捷键查询失败拒绝 VisualTeX 而让 Word 处理也合理。身份写入、用户设置恢复、内容失败回滚错误则必须可见，不能套同一 catch{}。PPT 原生类型/结构不满足时拒绝，优于跌到旧对象或整框字体回退。

## 测试独立性与证据分级

| 阶段 | 证据与结果 | 正确解释 |
| --- | --- | --- |
| CP217 | `artifacts/conversion-symbol-audit-20261002/shared-contract-aggregate/cp217-stage-aggregate.trx` 1983过/9失败；`cp217-stage-aggregate-r2.trx` 1992过/0失败 | 9 个旧角色/alternate 预期按原生字体/槽位证据最小更新，原失败保留。是当时精确构建汇总，不是本轮全树测试 |
| CP239/240/246/252 | 原前端实际导出、native 字节/字体证据、source audit、前后 TRX | normal/italic 源码及 mi-limit 的预期调整有输入角色依据；CP252 限定 7 项后经批准。仍应保留普通 mo/其他样式反例；不能靠实现的同一转换器自证 |
| CP255 | `shared-contract-aggregate/cp255-accent-width-aggregate-retained.trx` 2324过/1失败 | 固定宽度契约未关闭；固定组垂直位置另未通过。不是环境跳过 |
| CP257 | `cp257-color/aggregate.trx` 2337过/1失败；颜色专项后续16/16，相邻1016/1016；`cp257-color/word-edit/color-lifecycle.json` 5例 | 有实际 Word、逐字符及控制结构 RGB、Undo/Redo、保存后新 Word 重开证据，强于忽略 w:rPr 的内容指纹。没有物理点击/全六向视觉证明 |
| CP260 | `cp260-classic-color-popup/final-results.json`、`final-state.json` 及 after-* | 经典75、标准交互24、几何24按限定浏览器生产构建通过；旧按钮遮挡、Esc删选区失败证据保留。原生取色器与 Office WebView 生命周期未运行 |
| CP263 | `artifacts/powerpoint-basic-functions-20261004/styles-final-r2/native-styles.json`、`basic-final-cp263-r2/basic-functions.json`、`contract-final-cp263-r2/converter-contract.json` | 限定样式7/7、基础10/10；一组任意字体/混字号 passed=null，是前提未建立而非通过或已修复。实际 Ribbon 点击0，候选加载/交互环境受限；slide PNG像素不等于桌面截图。无 x86/MathType完整矩阵 |
| CP264/265 | `artifacts/word-prose-latency-20261004`，installed-code/summary、installed20、installed100-summary | 9个被动替身检查针对额外激活/布局；Ribbon替身11/15的4项接口未建模、idle fixture前提失败均保留。0/20/100 未见数量增长，跨进程回调/COM计时不代表物理输入延迟；用户已同意收尾 |
| CP266 | `build-cp266-word-shortcuts-r3/shortcut-routing-final.json`、`artifacts/word-shortcut-conflicts-20261004/owned-binding-final-r8/summary.json` | 35/35是模拟；9/9实际Word绑定逐项 ContextRestored=true、UsesPhysicalInput=false。早期 document-only binding失败r5/r6保留。双键原生fixture创建失败，只有模拟覆盖；attached/global定制模板独立验证未完成 |
| CP267 | `artifacts/reported-omml-conversions-20261004/basic-exact-audit.json`、current-basic-r3/results | 基础8/8精确内容/空白/颜色与Undo/Redo/同进程保存重开；SyntheticPreviews=true、PhysicalUi=false。历史两图在此集合未复现，不证明旧报告无效。两个特殊运行器栈溢出未出结果，原因未定、继续搁置 |
| 当前 CP268 | 只读差异/调用分析、原日志统计、已有产物 SHA256、纯函数有限反射探针 | 没有新完整构建、单测汇总、Office/物理UI/安装包/性能运行。R01/R03/R06 的宿主影响还需实验，不能写为已实测失败 |

CP217–220 的六向执行与后续 CP225/243/250/251 的不同 20 项集合不是同一次最新完整六向验收：原 19 项集合排除 F09；F02 在 OMML 写入方向保留严格样式失败；后来的源 OMML 有已丢失信息。不能把不同构建/不同源/不同断言的绿项拼成“当前六向20/20”。后续真实 Word 截图已有成功记录，并非一直被 Defender 阻塞；成功旧截图也不能证明 CP266 当前组合的视觉表现。新的截图或视觉验收本轮未做。

独立性薄弱处：新 `powerpoint_basic_functions_regression.mjs` 用源字符串检查方法/保护语句存在，适合阻止误删，不证明行为；编译后的 converter contract 实际反射调用能发现六参数签名不符，但也不等于宿主操作。MTEF encode/decode、OMML forward/reverse 和共用 canonicalizer 有共同错误风险，须用独立原生字节/原样保存 XML/真实渲染补足。CP257 RGB独立断言、CP263原生PPTX/导出PNG与CP267精确保空白审计是值得保留的补充。测试筛选排除的原生组不能因 TRX skipped=0 就算运行过。

## Companion、前端和工具还需的限定实验

- Companion：当前 Arc 修复证据支持运行中 handle clone 根因；`runtime_handle_acceptance.rs` 八工作线程各 20,000 次、三次运行和24回调不覆盖所有退出时序。Arc 只在线程安全地管理引用计数，最后一个引用会销毁内部值，并不自动使内部值线程安全，见 [Rust Arc 文档](https://doc.rust-lang.org/std/sync/struct.Arc.html)。将来可做无 Office、无窗口的“事件循环退出时仍有 worker持有最后handle”析构线程测试；目前是验证缺口，不是确认新崩溃。
- 快捷键：只在自有隔离 Word 中物理验证普通字母、Ctrl+D/F、未分配组合、菜单/查找框焦点、真实双键前缀及忙碌编辑器；先建立原生前提再计通过。当前不应再次扩大已收尾的普通文字数量测量。
- 前端：保持标准/经典、desktop/Office挂载交互对照，补原生WebView/系统取色器生命周期；源码锚点适配在依赖升级时必须 fail closed 并跑实际编辑后导出，不只字符串存在断言。
- 维护截图脚本 `artifacts/conversion-symbol-audit-20261002/cp198-VisualTeX-Omml/word-visual-evidence/Capture-CP198R1Style.ps1` 已有 manifest/hash、全新输出目录、自有PID/文档/HWND、截图前后前台检查及源文件哈希复核，失败不强杀未知进程；不改 Defender/AMSI 设置。它用于保存文档视觉证据，不执行转换，也不是所有权可自动转授给旧临时脚本的证明。九点遮挡采样不是像素级无遮挡保证，仍需看实际图；激活失败应单独记录，不能生成离屏图冒充 Word。
- `tools/chatgpt-auto-continue.user.js` / test 的近期变化是工具DOM回合定位、已观察活动的静默计时和折叠面板；不进入公式产品调用链。测试使用DOM替身，未在本轮运行或证明实际网站当前DOM。保留其独立范围，不把该工具测试计入公式验收。

## 后续最小处理顺序与交付条件

1. **内容安全先行**：修 R03 的目标来源/提交前变化检查，再补 R01 的文本框原子失败恢复。每项用两三个自有文本框用例及分阶段故障注入，独立比较正文、格式、邻接内容与 Undo/Redo，不先扩大到20条符号矩阵。
2. **把验收入口变为可安全复用**：封住 R04 的默认用户级设置副作用，保留 CP264 事件证据及不可恢复的初始快照缺口；旧脚本不得自动充当通用入口。
3. **明确候选内容**：解决/显式隔离 CP255、固定 resolver 选择与所有依赖，记录从源码清单到打包 payload 的哈希对应。单测旧失败需按原生证据裁定，不能只改数字。
4. **共享保真边界**：处理 R02/R05，在统一入口中修复清理所有权和受限解析；继续 R08 的字体/槽位与间距能力，避免逐符号码位补丁。先做公开入口和真实原生证据，再做日常20条。
5. **补齐限定验收后再扩六向**：Word/PPT物理命令、PPT混排拒绝与邻接保护、真实源样式基准、同一最终组合的转换/Undo/Redo/保存新实例重开/视觉。源已丢失信息的往返另列，不算完整通过。CP267特殊探针保持搁置。
6. **最后才讨论精简和性能**：R06、操作内重复解析、List.Contains、旧helper与配置读取按实际计量裁减；当前冻结不变，不以删边界检查换速度。

当前可以保留并共享明确标注范围的现有候选及这份审查报告；**不建议从默认脏树直接打面向用户的验收包，更不能宣布发布验收通过**。若后续只要专项内部诊断包，也必须显式声明 CP255/heldR4 与 R01/R03 风险、核验冻结构建及包内真实载荷，并另行得到构建任务授权。本轮未提交、推送、安装、部署、重构或修改产品/测试代码；仅此报告与 CP268 进度落盘。

已完成审查覆盖：历史检查点及五份阶段报告、两层Git差异、共享规则、运行时校验、Word/PPT宿主/范围/事件、快捷键、前端、Companion、测试独立性和工具副作用。未完成工作明确为上述修复/宿主实验/整体验收，而非继续读同一批代码或重复已搁置调查。
