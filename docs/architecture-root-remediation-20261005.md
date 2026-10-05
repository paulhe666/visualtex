# 架构审查后的根因修复（2026-10-05）

状态：进行中。用户在 CP268 审查后明确授权根源性修改。性能优化继续冻结，held R4 不动；不提交、部署、安装或改用户注册表/模板/绑定，保留所有旧修改与失败证据。

CP271范围调整：用户明确“不用纠结一次undo”。单次Undo停止作为当前阻塞项，保留失败证据及现有多步行为，不删除失败恢复保护。后续先处理CP255原生重音依据。

CP272最新：共享MTEF读入口补齐完整终止、原生递归及生成树深度边界；37项定向通过，离线汇总2371/2372（另4项原生服务筛选排除），旧CP255失败保留。精确net472 DLL的归档19载荷读回逐字不变。新候选与范围见文末，不能把CP270旧宿主验收自动算给本轮。

起点 HEAD `67e57fef62a4abcfa9577fac8849c53817f83893`，工作树与 CP268 审查相符。旧候选 CP266 Word `827AAE47…151CEF`、CP263 PPT `5857E555…459BC3` 只作基准，不冒充新修改的验收结果。

## 处理顺序

1. R03/R01：文本插入目标必须来自当前选择；操作内保存并验证真实源内容/格式；所有可失败的纯转换先准备，宿主写入须有可验证且不越过用户历史的恢复边界。不能靠取消正常插入或恢复纯文本来规避。
2. R04：封住旧验收默认入口的用户级设置写入，新测试仅走有所有权证明的独立入口。
3. R02/R05：统一受限 XML 解析并保留合法文本，测试实际公共入口而非未接入 helper。
4. R06 与其他保真能力：按正确性需要继续；不借本轮进行性能重构。CP255 及原生能力缺口必须有原生依据，不能修改旧预期来制造全绿。

## 验证边界

先小量根因用例、实际编译和独立内容断言；重要阶段再用既有代表集合。构建继续显式采用 `Build-Candidate.ps1` + `FreezeR3.targets`。自有 Office 启动必须确认新进程/空实例；出现已有 PowerPoint 则不创建或接管它。旧已搁置 CP267 特殊探针、文字数量卡顿调查不重开。每一阶段分别记录通过、失败、未运行和未建立前提。

起始阶段仅检查源码，随后按下列记录建立原生恢复前提并接入产品；不使用没有历史证明的盲目 Undo 或整对象替换作为文本恢复。

## CP269 原生恢复能力前提已建立

新增验收专用 `PowerPointTextSafetyAcceptance.cs` 和 Main 的前置 `powerpoint-text-safety` 分支，绕开全部注册表 suppression。frozen-R3 `build-cp269-text-undo-probe` 构建成功。自有 PPT48412（2026-10-04T18:37:18.3134995Z）中读取原生 Undo 列表、分组修改文本/字号、仅撤销新增前缀：原历史1项→2项→1项，正文与两侧粗/斜体恢复。独立解包 before/restored.pptx，整个 txBody XML 628字符逐字一致。此时产品代码尚未修改，这只是恢复能力前提，不是产品失败恢复验收。

证据 `artifacts/architecture-root-20261005/cp269-undo-probe/`。实际用户HKCU Word/PPT加载键前后不变，原Word88900/93828/95356及启动时间不变。PPT在Quit后异步退出，稍后只读复核PID48412已不存在，没有强杀。初始默认Shell是codexsandboxoffline，那里HKCU无同名键；该读失败属于上下文前提，随后经工具批准在实际桌面用户上下文只读快照并运行自有探针，未据错误hive推断注册表值。

## CP269 产品接入与保留的失败

修改 PowerPoint.ThisAddIn.cs 移除旧文本目标回退及事件内无消费者读取；PowerPointFormulaService.cs 的瞬时目标携带文档/slide/shape/范围和原生源指纹，提交验证来源；PowerPointNativeOmmlKernel.cs 将线性化与段落规划前移，实际写入交给新增 PowerPointTextEditTransaction.cs。事务只撤销原历史上方可证明的新增项，验证原生内容/格式及历史恢复；异常不吞掉，恢复无法证明时保留双重错误。PowerPointNativeOmmlPackage.cs 共享原生文本框读取。DirectOmmlStress 的会话夹具补齐真实 SourceDocumentId；未运行压力测试。

候选 r2–r4 构建成功，但未编辑源误报过期，测试未进入注入写入，失败全部保留。r4 保存原始源和提交前实际源，差异仅为后台校对新增的四个 dirty="0" 属性。依据 Microsoft OOXML Primer 的 dirty/err/smtClean 机器校对缓存定义，源指纹仅排除这三个状态属性，保留 noProof、语言、字体、颜色、间距、正文、几何和链接；没有放宽公式保真比较器。证据见 `cp269-text-safety-r4/*source.xml`，官方依据 https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf 。原始快照不改写。

frozen-R3 `build-cp269-text-safety-r5` Word SHA256仍827AAE47…151CEF，PPT SHA256 `89FECA7A3F9F91C0D5FB4D76ACEF7FBFE598F639C70BC03FD011603F2A7453C1`。8个定向检查通过：作者属性保留、取消选区、三个失败阶段完整恢复、同长度修改拒绝、仅格式修改拒绝、非法 MathML 显示插入无写入。成功插入后的包读取失败：a14 Choice 与 Fallback 均为相同ID的 p:sp，被新共享读取器当成两个对象；Undo/Redo和display成功测试未运行。已有CP263原生PPTX独立解包确认结构。正按已支持a14兼容分支解析修正，仍拒绝无关重复ID/未知分支，不用First随意取对象。r5环境核对KeysUnchanged/InstalledFilesUnchanged/ProcessesUnchanged均true。所有日志在 `artifacts/architecture-root-20261005/cp269-text-safety-r2` 至 `r5`，不得把部分通过写成整轮通过。

## CP269 完整失败恢复已通过，单次 Undo 限制仍失败

r6/r7 PPT DLL `E55932C806E1FFDA0C3394B79C67EE6318D545502D9F5E19E8E0B500D382D386`，Word仍827AAE47…151CEF。r7 13项定向检查通过（另有原生能力前提），包括行内和独立行原生转换完成后再注入异常，整个作者内容/格式XML及原Undo历史均恢复，成功插入保留shape身份和两侧正文。两种模式的单次Undo均失败：原生命令建立多个记录，一次只退回线性文本；Redo断言因前置失败未运行。仅为继续独立测试而验证并撤销测试自己的剩余前缀，不能算作单次Undo通过。全部XML/PPTX/失败保留于cp269-text-safety-r6/r7。环境三项核对全部true。未用整对象替换混排文本或吞掉宿主异常来规避限制。

## CP270 共享XML/文字角色、旧入口隔离与被动选区修复

- 新增 `FormulaXml.cs`，统一4,000,000字符/256层/250,000元素约束；代理XmlReader在实际装载过程中计数，没有先扫一遍再重读。Word OMML正向验证与反向Extract入口、MathML→LaTeX、MTEF创建/改写/语义签名/interop的现有parse接入；删除两个从未调用的16MiB helper。未把全项目任意XML入口都宣称覆盖，MTEF二进制深度另属边界。三个csproj显式链接同一规则。
- 删除 `WordOmmlConverter` 无所有权依据的全XText零宽清理。首轮新测试20/21：U+200B仍丢失，独立定位到本机Office `MML2OMML.XSL` OutputText的translate，随后按mtext/ms字面文本角色适配已加载XSL，保留正文字符和空白；数学运算符处理仍在数学角色分支。没有编辑Office安装XSL、逐码位补输出或隐藏元数据。r2定向22/22通过，含字面/十进制/十六进制、单零宽token、空白/NBSP/运算符形文字、实际公开入口深度/DTD/节点/长度边界。
- 保留r1构建失败（旧manual partial还引用已删除注册表scope）、r1测试20/1、r2汇总2363/4原始证据。汇总两项SVG受sandbox用户写实际用户临时目录权限阻塞；一项是已知CP255字节230→241；一项旧测试仅凭普通ORD mo U+200C就假定自有anchor。该夹具无宿主所有权证据，改为严格逐字断言c2+ZWNJ+1（不是删断言或放宽比较器）；原始失败保留，需实际用户上下文汇总复测。
- Program.cs删除HKCU LoadBehavior写入/静默恢复类；旧默认与clone/manual入口在宿主创建前报告隔离前提未建立并停止，已迁移前置独占入口可用。旧artifact脚本中的COM CodeBase/SendKeys证据未改，不得复用；本轮没有宣称旧所有运行器均已迁移。
- R06：ReadSelection/ReadFormulaAtScreenPoint不再Configure或查重；返回精确slide/shape引用，目标消失不按共享FormulaId替换另一副本。显式编辑取得操作锁后才PrepareManagedSelectionForEdit修复复制身份；owner已一致不再Configure。清除无消费者公式缓存及仅清缓存的关闭事件。保留真正的选区UI失效订阅。新测试覆盖被动Saved/Undo/名称/标签/锁纵横比不变、重名复制体显式修复、已删除复制体不能替换原件；正在实机运行。

当前frozen-R3候选 `build-cp270-root-r3` 编译成功；产品文件还包括 `PowerPointFormulaService.ManagedSelection.cs`、ThisAddIn/Service、共享转换器；测试为FormulaXmlBoundaryTests、TextSafety和BasicFunctions最小夹具更新。held R4未动。候选哈希与本轮通过/失败/未运行范围将在运行后补齐；CP255及原生表示/六向视觉缺口没有关闭。
## CP270 最终阶段结果（2026-10-05）

最终候选为 `artifacts/omml-edit-audit-20261001/build-cp270-root-final-r2`，net472/x64、VS18 MSBuild、SDK8.0.423、明确导入FreezeR3。Word DLL SHA256 `C1F790B2AE252360225E17C21E36954BD57B1516C116E6FD47CC4E5BBEC58FB3`；PPT DLL SHA256 `30957E9D7BC8CBD710C9A9BCE825408504D83087B7CD13ACAC3E506B0886D406`。held R4仍 `6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`，HEAD仍67e57fef。精确输入文件哈希、产物与环境清单：`artifacts/architecture-root-20261005/cp270-final-state.json`。

| 验证 | 最终结果与边界 | 证据 |
| --- | --- | --- |
| C#候选编译 | 成功；冻结R3，不是默认脏树构建 | build-cp270-root-final-r2/build.log、build-state.json |
| 离线源码汇总 | 2367通过 / 1失败 / 0跳过，共2368；失败仍为CP255原字节契约，未改它的预期 | cp270-aggregate-final/results.trx、run.log |
| XML新增定向 | 23/23包含于上述汇总；公开入口限制、DTD、字面/实体、mtext/ms角色及空白精确保留 | FormulaXmlBoundaryTests、上述TRX |
| PPT文本安全/选区身份 | 15项产品定向检查通过，另1项原生Undo能力前提通过；2项单次Undo失败，因前置失败Redo未运行 | cp270-final-r2-text/text-safety.json、原始source/target/one-undo.xml、PPTX |
| PPT基础操作 | 10/10限定操作通过，包括SVG/OLE、原生编辑和保存重开 | cp270-final-r2-basic/basic-functions.json |
| PPT限定样式 | 7/7通过；任意逐字符字体/混合字号前提未建立，实际Ribbon点击0 | cp270-final-r2-styles/native-styles.json |
| 旧运行器安全门 | all、exact clone、manual三项均按预期在宿主创建前拒绝；是隔离门通过，不是旧功能验收通过 | cp270-guard-r4-*/environment-checks.json、acceptance.log；此后Program源码未改 |
| 原用户环境 | 每个最终宿主运行的注册表键、安装DLL、原Word PID/启动时间均前后一致；自有Office正常退出 | 各最终目录environment-before/after/checks.json |

离线测试程序集是net8.0-windows，直接链接共享源码，并不包含完整Word宿主resolver；不要拿2367项当net472产品宿主全验收。测试入口使用冻结配置，实际产品DLL另由上述明确R3构建与宿主测试核对。首轮实际用户汇总r3因复用了sandbox NuGet资产路径而未建立编译前提；r4显式使用实际用户已有离线缓存后完成。所有失败/权限/依赖前提日志均保留。

最终R06还将三处延迟OLE几何恢复的目标改为精确slide/shape引用，调度代次以文档+对象为键，防止同FormulaId复制体互相替代/取消。没有更改调度时序或做性能优化。共享定位器的删除副本反例已实机通过；具体延迟回调的用户并发拖动场景未专项验收。兼容分支读取只接受已验证的单个a14 Choice，未知/多个Choice明确拒绝，不猜当前Office支持的备用分支。

保留的修改前证据：r2–r4源校对状态误拒；r5兼容分支重复；r6/r7两类Undo失败；CP270 r1构建缺引用及20/21新测试；r2汇总2363/4；r3旧目标测试原本仅捕获InvalidOperationException，实际严格解析器抛InvalidDataException，随后只为该断言指定精确异常类型并仍验证零写入，没有放宽生产判断。最终所有产品安全/角色断言保留完整内容和格式比较。

### 本轮结论及后续顺序

R01失败恢复、R03文本目标/来源、R02已知零宽/文本角色损失、R05所列公开转换入口、R06被动写入和精确对象定位已落实源码及相应定向证据。R04删除默认注册表suppression，未迁移旧模式被明确阻止；旧CodeBase/全局输入工具仍作为历史代码/证据留存，不能绕开Main或复用旧artifact脚本执行。

仍不适合标为完整验收通过的包。首要剩余问题是PPT文本插入单次Undo事务；应研究宿主支持的组合事务/单命令写入合同，不能拦截用户Ctrl+Z、盲撤多项或整框替换混排正文来造单次通过。其次是CP255固定组重音位置与原字节契约，需要真实MathType原生固定组样本的模板/CHAR字体槽位/尺寸/渲染依据，本轮没有猜测或回退整个共享文件。R08的script-l可显示性、U+2AFD、任意/负间距及其他符号/括号/上下标缺口依旧，不能用同codec读回冒充独立原生视觉。最后再扩展最新构建的六向矩阵、实际键盘/Ribbon/Companion和基线视觉；CP267特殊探针与文字数量卡顿调查继续搁置。

未提交、推送、安装、部署、改用户注册表/模板/快捷键，未触碰用户原有文档或关闭安全机制；全部原有修改保留。阶段报告和检查点已更新。

## CP271 原生命令撤销能力与范围调整

新增验收专用PowerPointUndoGroupingAcceptance.cs，Program.cs与Test-PowerPointTextSafety.ps1增加专用入口。frozen-R3 build-cp271-undo-capability-r1编译成功，产品Word/PPT DLL仍与CP270-final-r2一致。四种原生组合中，文本/格式合一项，插入math zone与Professional分别增加记录；直接和临时CommandBar事件回调均不能合并。每例一次Undo只退回线性x/y，没有恢复SELECT，故四项均未满足单次恢复，进程正常退出不能算产品验收通过。原始分阶段列表和PPTX见artifacts/architecture-root-20261005/cp271-undo-capability-r1；环境三项一致且自有Office退出。

用户随后明确“不用纠结一次undo”，停止该方向探索，未增加产品命令回调、剪贴板、VBA或快捷键接管。CP270单次Undo失败保持事实记录，但不再列为当前交付阻塞条件；其余内容保护、固定组重音与未运行保真范围不变。下一阶段审查是否能从MathType自身输入生成固定组重音证据，避免继续基于CP122人工构造的HatBox推断正确位置。

## CP272 共享原生结构读取修复

### 根因与修改

`MathTypeMtefCodec.MtefStructureReader.ReadRoot`的普通调用只返回首个根，没有使用已存在的`HasOnlyEquationEndRemaining`；后者仅被旧恢复探针使用。额外CHAR/另一个根被静默丢弃，缺少最后公式END也被接受。四个递归容器没有深度界限，且平铺的trailing-script可反复包装前一节点，最终树深度不能仅以原生调用栈衡量。

修改只在共享codec：ReadRoot检查完整终止；ReadLine/ReadPile/ReadMatrix/ReadTemplate在进入时执行256层限制，scope离开时恢复计数；已有namespace转换改为有界非递归遍历，同时限制生成MathML的256层/250k元素并清理同一内部run属性。没有额外XML序列化/解析往返，没有修改语义比较器、符号规则或用户文档。二进制字节大小及其他解析入口不因此被宣称全部有界；本修复也不用于归因搁置的CP267崩溃。

### 结果与精确身份

- 新增`MathTypeNativeStructureBoundaryTests.cs`独立拼结构字节，只复用有效前缀，修改前2过/6失败（六项均未抛异常）；修改后8/8，连同XML/扩展prefix共37/37。
- 新离线汇总2371通过/1失败，0跳过，另4项原生WhenEnabled测试明确筛选排除。CP270全计数2368包含4个未启用即返回的测试，因此本轮总数2372=2368+8−4；不能称四项原生服务已运行。旧CP255字节失败未修改。
- 完整net472/x64冻结R3候选`build-cp271-native-structure-final-r1`编译成功；阶段编号CP272与既有候选label的cp271前缀分别记录，不重命名产物。Word SHA256 `41F42D5F3CE327E39D7371D0F80748B1A0A91487BA9AEF39857496EE7AA14F87`，PPT `8F9E779477E9443524934C127978505548E747E81CEE17718F5D2660FB3215AC`。
- 新增证据脚本`artifacts/architecture-root-20261005/Probe-NativeRead.ps1`只读解包CP217保存的19条MathType OLE，分别加载CP270/新net472 DLL。两版MathML逐字19/19一致，归档DOCX哈希不变。没有Office实例或renderer参与，不能冒充宿主写入/Undo/重开/视觉通过。F09原本排除；历史源损失也未恢复。
- 证据为同目录`cp271-native-structure-before`、`after-r1`、`aggregate`、`cp271-native-read-*.json`；源码/构建/证据清单`cp272-final-state.json`。held R4与HEAD未变。

### 原生重音仍需的依据

只读核对本机厂商`Presentation MathML input.tdl`：817–831字符EMBELL要求单字符基底，851–864组帽/波浪使用扩展模板；562–564忽略accent属性。这意味着“经厂商导入器生成”仍可能丢失固定宽度意图，不能据其输出宣布CP255固定组问题解决。现有独立renderer输入是我们生成的MTEF，不能反过来成为原生作者意图证据。缺少真实MathType固定组的模板/CHAR/字体编码/尺寸及对应视觉样本，故没有加位置补丁或改旧预期；本轮也没有调用可能连接交互MathType的Word bridge、改服务或绕过宿主隔离。

当前继续保留CP255、R08及同一构建六向/物理/视觉缺口。已修复的内容安全与共享输入边界可以作为有限候选审阅，整体仍不能标记发布或全部验收通过；用户要求的单次Undo放行已记录，无需继续该实验。
