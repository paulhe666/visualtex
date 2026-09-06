# 3d207d7 整改记录（进行中）

本目录记录整改，不改写 `docs/audit-3d207d7` 的原始证据。基线 HEAD 为 `3d207d7a81ff40014fd861b242cff96ea77060a2`；没有提交、切分支、reset、clean、stash 或发布。接手时 Cargo.toml 已有行尾状态，保留。

## 当前状态（2026-09-06，尚未完成）

阶段记录按实际发生保留；下文较早的“尚未安装”描述为当时状态，不代表最新状态。尚无最终 NSIS，尚未进行最终同源安装后的完整矩阵验收。

- 阶段01b文档5的行级编辑、关闭自动应用、编号开关、引用恢复与真实鼠标跳转通过所记录路径；并非整个阶段完成。
- 阶段01d文档10的中间行增大/减小及后续编辑保持3个OMath/3个SEQ。移除了在目标文档末尾临时插入公式测量高度的路径；真实截图与COM/XML位于`stage01d-*`。字号操作带来的编号域MERGEFORMAT变化已记录，阶段01e继续改为复用健康编号宿主，但该项复用尚待新版本重新验收。
- 阶段01e文档12/13从新空文档走真实批量导入A/B：空文档为6个OMath/5个SEQ；已有1个编号OMML后导入为7个OMath/6个SEQ。新公式内容一致，原公式和域保持，公式字体均为Latin Modern Math 10.5。`stage01e-bulk-comparison.json`记录实际结果；完整引用、正文格式和字体选择矩阵尚未完成。
- MathType兼容引用别名与cases单侧定界符已有公共数据机制变更及补充测试；真实Word验收未完成，不能宣布修复通过。
- 邻接转换及失败恢复**未通过**。阶段02文档15在转换源范围删除时报错，Undo后COM结构及OLE/EMF数据恢复；其OOXML关系/对象ID重建导致旧签名误报，已修正签名对这些易变标识的规范化并保留实际载荷哈希验证。
- **阶段02b文档17发生真实数据丢失**：新空文档真实插入两个编号VisualTeX OLE和末尾一个行内OLE后全文转OMML，收尾指纹身份校验报错；错误恢复后从3个OLE/2个编号变为2个OMath/0个OLE/0个编号。不能用阶段02的签名误差解释这次恢复失败。`stage02b-vt-tail-before/after-__17.*`、`stage02b-word-hook.log`及`stage02b-live-error-observed*`保留证据。2026-09-06用户指出弹窗后，实际再次查看弹窗并只读检查现场，点击确定关闭提示，保留文档17内容。此前没有及时观察和汇报弹窗是执行疏漏。
- 当前优先追查撤销记录何时被截断。阶段02c只增加公共Undo状态和恢复组成项诊断，并取消公共Undo入口吞异常；不以辅助测试代替真实恢复验收。后续转换验收、重绘、性能矩阵和最终安装交付继续待办。

阶段02c文档19重新从空白走相同三OLE路径，09:40实际转换再次失败并立即停止脚本、查看弹窗、COM/XML取证后点击确定。日志确认首个行内OMML插入期间外层CustomRecord从true/1变为false/0；后两个插入各自建立独立记录，所以一次Undo只撤回最后一次插入。只读枚举真实Word的Standard工具栏内置Undo下拉控件（ID128、type6），可取得全部125项有序历史，而在CommandBars集合上FindControl返回的type1按钮只有单项标题。`stage02c-readonly-undo-list.json`保存这一证据。

阶段02d新增公共`WordUndoHistorySnapshot`，在任何目标修改前捕获完整已有撤销历史；失败时要求当前历史是“新增前缀+完整原历史”，只撤销这个已经证明属于当前操作的前缀，并核对撤销后历史完全回到原边界。历史被截断或不一致时拒绝猜测撤销次数。整批转换与局部OMML编辑/字号恢复共用此机制，保留内容、编号、书签、格式、字体及元数据的复原校验。该版本正在真实安装验证，不以构建通过宣布恢复成功。

`stage02b-original-preservation-comparison.json`记录用户指出错误后再次核对的八份原始文档：全部COM字段与接手一致、Saved=false。

阶段02d文档21的同一真实转换失败中，完整历史从84项增加到126项，公共机制仅撤销新增42项并核对回到原84项。实际恢复后为3个OLE、2个SEQ（含可见编号及其他域共7项）、文档长度353；前后完整COM清单逐字段一致。恢复签名唯一差异是两个同位置的编号/题注书签起始节点交换了XML顺序，范围与内容不变，`stage02d-recovery-diff-*`给出精确差异。阶段02e只在只读签名中排序连续、同种、同位置书签标记，仍保留每个名字、起止与所有介入内容；重新比较实际Word前后证据完全相等（`stage02d-recovery-final.json`）。文档21错误提示已查看、记录后关闭；这不是阶段02b/02c真实丢内容情况，后两者仍明确失败。

阶段02e同时整改实际转换：只对当前准备的新公式用导入内容签名验证Word规范化前后等价，再保存实际Word OMML的严格原生指纹。进入新公式收尾必须证明元数据中的暂存指纹匹配本批准备源；编号归属、物理OMath一一对应和后续原生指纹检查均保留。导入内容签名处理空属性容器、命名空间拼写、布尔值词法、默认积分符号及非字面数学减号；算子、上下限、指数、变量、根式属性和字面文本差异继续被拒。实际失败会话MathML与真实Word残留OMML的补充数据对照证明欧拉/高斯等价；旧原生指纹确实不等。此为转换根因证据，不代替真实Word成功验收。新增边界/导入内容测试和既有身份测试合计17项通过（`tests-stage02e/import-and-undo.trx`）；阶段02e正在新Word中验证正常转换。

阶段02e文档24的正常转换仍未通过，首轮导入签名改动尚未覆盖实际校验差异；脚本在真实错误弹窗处停止，立即查看截图、取COM/XML后点击确定。恢复后3个OLE/2个SEQ/7个域/长度353，实际外部前后XML规范化比较一致（`stage02e-recovery-diff.json`），但插件即时恢复校验仍报差异。阶段02f新增显式可选`VISUALTEX_VSTO_TRACE_RECOVERY_XML`诊断：只记录同一真实调用链内现有XML读取的值，既不注入故障也不修改检查结果；用于定位即时校验与事后取证的差异。正常转换与即时恢复验证仍为未完成项。

## 原始文档保护

原始 Word 进程为 27664，文档1、22、26、28、29、33、35、36保持打开、未保存。`takeover-original-inventory.json` 与审计最终快照逐字段一致。阶段安装后再次按进程27664只读取证；`stage01-original-preservation.json` 显示八份文档完整 COM 清单均未变化，Saved=false。未编辑、保存或关闭这些文档。

## 真实基线

独立新建文档71，通过已安装 Ribbon 插入勾股定理、欧拉恒等式、高斯积分三个编号 VisualTeX OLE。全文 VisualTeX→OMML 后为三个真 Display OMath、三个 SEQ、一个3×3宿主。

选择第二个 OMath，真实点击“编辑所选公式”，点击勾股预设后“更新公式”：变为四个 OMath、两个编号。真实会话 `a2159415-db1b-42d3-b3c5-6255d1930985` 以 failed 结束，原生错误为编号宿主不是 managed 1x3。对应 `baseline-three-ole`、`baseline-three-omml`、`baseline-middle-edit` 证据及 `ui-actions.ndjson`。失败编辑器随后点击取消关闭，文档71保留。

## 阶段一实现

- 共享编号行定位同时验证三列结构、列归属、行号、Story及Range包含关系。编辑收尾用编号身份定位同一行中心OMath，取消重复的1×3解析器与事务外收尾分支。
- 中心段落、身份锚、行高、编号刷新使用目标行；编号宿主范围返回所属行。取消多行编号时先用Word原生Split隔离已验证的目标行，保留其余行。
- 结构编辑要求身份和内容指纹均匹配；漂移恢复保留唯一性检查，不能用旧会话坐标覆盖身份失败。规范OMML锚还必须为折叠范围。
- OMML编辑使用独立Word自定义撤销记录，失败时撤销该记录并恢复元数据，校验局部行/段落XML、文档长度和数学字体。删除原先局部删对象、重插原公式并吞恢复异常的分支。编号异常向上传播。
- 插件连接日志记录真实进程、加载路径和模块MVID，用于安装与运行来源核验。

以上尚需真实阶段矩阵确认，不能据编译或辅助测试宣布修复完成。删除、引用、多行编号开关、失败恢复及性能均继续验证。

阶段01b的第二份新文档8仍经三OLE→OMML形成3×3。首行更新与末行直接关闭自动应用成功，保持3公式/3编号/3×3；但随后选中中间行点击Ribbon“增大”，出现4公式/2编号和文末副本（`stage01b-nx3-font-increase`）。因此阶段一尚未通过，不以此前编辑成功覆盖字号失败。根因继续定位到公共`EnsureDirectTableSequenceNumber`固定写`Cell(1,3)`；该处已改为经公式范围验证的所属行。字号事务也开始复用编辑的局部Word Undo与元数据恢复，并验证公式数量及编号行归属；尚待安装验证。

## 阶段安装与验证方式

现有生产MSI脚本要求关闭全部Office，不能在保护原件的条件下直接执行。因此阶段验证把实际构建DLL及依赖复制到用户目录按SHA-256命名的不可变目录，仅覆盖当前用户的对应COM类注册。保留旧机器安装文件和旧Word进程；新Word通过正常COMAddIn加载同一ProgID和Ribbon。没有创建测试服务、手动调用插件回调或写Word XML完成测试。

`install-stage.ps1`、`start-stage-word.ps1`、`stage01-installed.json`、`stage01-word-process.json`记录该过程。HKCU对应类此前不存在；阶段脚本拒绝覆盖不属于本轮独立目录的注册。后续最终安装需明确处理这个注册覆盖，不能留下遮蔽正式MSI的未知开发路径。

新Word进程59000日志确认加载 `69E6A5F755466870142149BA706C938BB8AB40D0EF62B179827D3FF992E58EEF` 目录的DLL，MVID `8aea5012-f70a-4023-ba02-e82a43c76bb6`。`VISUALTEX_VSTO_ACCEPTANCE` 未启用。真实UI操作仍使用Word Ribbon、VisualTeX编辑器与导入窗口；COM只用于新建/选择/输入源码及只读取证。多Word进程用原生Word窗口的COM对象连接指定进程，防止ROT默认返回旧进程。

`stage01-build.log`记录构建成功。辅助测试354项全部通过（`tests-stage01/stage01-authorized.trx`）。首次沙箱运行有两项产品临时SVG写入被拒，授权环境重跑通过，未跳过用例。这些测试不代替真实Word验收。

## 未完成

### 阶段一真实中间结果

进程59000的文档3由空白文档经真实Ribbon建立三OLE，再全文转为三个OMath、三个SEQ、一个3×3表格。中间行编辑成功，仍为三个OMath和三个编号；首尾公式内容、编号域代码及身份保持不变，只有中间内容变化。`stage01-middle-edit-comparison.json` 与对应COM/XML/截图互证，实际插件核心编辑耗时380ms。

关闭中间编号后为三个OMath、首尾两个1×3编号宿主，中间独立Display；重开后三个OMath、三个1×3编号宿主，编号1/2/3正确，核心耗时分别629/594ms。但中间引用在重建同值编号后仍显示未找到引用源，真实Ribbon更新编号也未恢复，因此该项未通过。阶段01b把引用结果健康与编号值变化分开检查，并在编号状态变化后刷新目标引用；重新构建并从新空文档验证中。

阶段01b正常COM加载进程48952，DLL SHA-256 `490AC3FA9D37B565F58C911A8A191587D89B423D7B726E108792299A7D3AC116`，MVID `1e0a4ee8-4736-4454-92f4-894cdf9db1d3`。阶段01证据与旧阶段文档保持可比，没有覆盖正在加载的DLL。

文档5重新从空白经三OLE→OMML、真实插入中间引用、关闭/重建编号：仍是三个真Display OMath及三个编号，原引用恢复0.0.2（`stage01b-number-on`）。随后首行、末行更新，中间行连续三轮（第三轮直接Alt+F4关闭编辑器自动应用）全部完成；`stage01b-row-edit-comparisons.json`比较各次真实XML，只有目标公式改变，非目标OMath XML和域代码/结果均不变。相应会话均completed。`stage01b-final-row-edit.png`可见三个内容和编号；`stage01b-reference-jump.json`记录真实鼠标双击正文引用后选区准确落到中间编号书签182:243，文本0.0.2。GetPoint对嵌套REF返回零宽，最终使用已观察截图坐标并以RangeFromPoint验证点击确属引用后再双击；没有调用Field.DoClick或Goto替代鼠标操作。

引用对话框列表的键盘聚焦存在不稳定行为，失败已保留在UI日志，脚本在失败处停止；现使用原生LB_GETITEMRECT读取真实列表项矩形、实际鼠标单击并核对LB_GETCURSEL。未改变任何产品对象完成列表选择。

### 阶段二实现进行中（尚未安装）

普通OMML范围解析开始共用严格的只读身份解析器，去掉按邻近内容猜测并在读取中重绑的重复分支；新插入/重绘尚未保存元数据的编号公式按已验证的物理编号行收尾。需要继续核对原生Word改写内容后的编辑入口，避免严格结构编辑校验误挡原有内容刷新功能。

OLE书签重绑改为同名Add重定位并验证精确跨度，不删除名称；转换收尾扫描实际内嵌身份，仅修复范围漂移的既有VTO书签，歧义会报错。批量OMML临时文档补传目标OMathFontName。局部撤销检查点补齐COM释放，并允许普通用户表格中的段落作为局部范围。`stage02-in-progress-build.log`编译通过，但这些变更尚未作为真实Word修复通过。指纹批量配对、转换完整撤销机制和其余整改仍在实施中。

`WordOmmlNativeSource.IndexConversionEquationIdentities`已替换单个pendingFormulaId扫描：按编号所属行、段落内折叠锚及完整公式指纹索引验证一一对应，跨段落漂移仅允许全局唯一内容恢复；同内容不同编号行、行身份与内容冲突、重复身份和非唯一恢复六项数据测试通过（`tests-stage02/identity-index.trx`）。这是机制测试，仍不是实际Word验收。后续必须解决转换外层遗留全篇FormattedText回滚和临时指纹异常吞掉路径后，再将该阶段视为完成。

阶段一真实矩阵进行中；邻接书签稳定与内容/编号对应、完整批量编号集合、MathType引用别名、批量字体、重绘边界和正文样式、cases单侧围栏机制及最终NSIS/安装全矩阵/性能对照尚未完成。最终交付前必须逐项填入真实证据。

### 阶段02f/02g即时证据与公共XML读取

02f文档27的插件即时前后XML明确显示恢复校验差异为OLE显示宽高取整，外部最终COM/XML已一致；不能把这次差异等同于02b/02c的丢内容。02g在最终验证和恢复前关闭批量临时源并恢复用户视图，文档31真实失败后插件日志 `body=True mathFont=True metadata=True variables=True`，外部全量COM逐字段相等（`stage02g-recovery-comparison.json`）。弹窗立即查看截图、COM/XML取证并确认关闭。此为该路径实际失败恢复通过，正常转换仍失败。

02g在同一真实调用链捕获直接对照：文档31的 `Content` 范围为0:63、COM有3个OMath，但其WordOpenXML正文为空；立即取得相同坐标的 `document.Range(0,63).WordOpenXML` 包含完整3个OMath及真实页设置。两份证据为 `stage02g-word-hook.log.export-mismatch.xml` 和 `export-explicit-range.xml`。02h增加公共 `WordDocumentXml.Read`，始终以Content确定界限并重新绑定主正文Range导出；转换身份、编号XML快速路径、恢复快照共用。COM/XML数量检查提前执行、保留；没有通过放宽指纹来容忍空导出。02h尚待真实转换结果。

### 阶段02i正常转换通过所记录路径

阶段02h文档34证明单次重新绑定Range不足：首导出仍为空、立即第二次导出3OMath；错误恢复即时四项通过。02i公共WordDocumentXml先读实际主正文OMath/OLE清单；首次导出与实际不符时只读重新绑定一次，第二次仍不符则失败，同时验证读取过程文档范围、公式/形状数量未变化。没有重试任何内容修改，没有关闭身份/内容检查。

02i文档39从真实新空白经Ribbon插入两个编号VisualTeX OLE和末尾行内OLE、全文转OMML。日志正常完成3/3，实际3个OMath/0个OLE/2个SEQ/一个2×3表格，可见编号(0.0.1)/(0.0.2)，勾股/欧拉/高斯顺序及内容正确；字体Latin Modern Math 10.5。`stage02i-vt-tail-after-__39.json/xml/png`保存证据。Word apply 10:25:16.158至10:25:20.007，约3.85秒；全流程计时另见ui-actions。

末尾独占一行的公式实际为Display。最初将此误判为回归；用户明确指出Word会自动把没有正文的行内公式显示为行间，故纠正判断：该独占一行路径通过，保留Word原生行为。刚增加的强制inline代码和相应失败判断已全部撤回，未安装；stage02j-build.log记录该未安装尝试的编译错误，仅保留追踪。下一项用新空白文档在末尾行内公式后加入真实正文验证。

当前完整辅助测试379通过、0失败、0跳过（tests-stage02i/stage02i.trx）；仍不代替真实Word完整矩阵。阶段02g/02h恢复与02i正常转换有分别证据，最终NSIS、后续MathType/重绘及完整性能矩阵仍未完成。

### 文档41及阶段03引用链

文档41在两编号OLE前加入宋体12磅正文，末尾行内OLE后加入宋体12磅正文，再经Ribbon插入第二公式引用、全文转OMML。实际3OMath/2SEQ/一个2×3宿主，末尾OMath.Type=1，编号(0.0.1)/(0.0.2)，REF结果0.0.2。前后正文run字体/字号XML一致（仅后方run增加Word rsidRPr修订标识）；`stage02i-prose-run-format-comparison.json`给出实际XML。引用实际鼠标双击跳转至第二编号书签117:178（`stage02i-prose-reference-jump-native-input.json`）。首次PowerShell逐次发鼠标消息时两次按下实际间隔558ms，未形成双击；失败证据保留，改为SendInput一次提交4个真实鼠标事件后通过。没有使用Goto/Field.DoClick。

阶段03统一MathType编号别名：VTEqNum_GUID仅覆盖数字，ZEqnNum覆盖含括号可见编号；同名Add并精确验证，不先删除引用书签。改编号重写与格式转换共用；直接REF/嵌套REF更新使用公共WordEquationReferenceFields，移除MathType重复刷新与更新吞异常，且不重复更新Document.Fields暴露的嵌套REF。

阶段03文档46从空白新插入编号OLE和引用后转MathType，再真实改为全文连续编号，原审计的兼容别名拒绝未再发生，MTPlaceRef及REF从0.0.1变1。该完整引用链仍未通过：截图为“(1”且REF在GOTOBUTTON外。转换前快照已经出现同样结构，确定是引用插入问题，不是编号重写才造成。阶段03b统一所有格式的引用创建，以明确的内部文本槽替代在Code.End折叠边界插入，并先写好前后括号；Ribbon恢复插入前选区，引用创建失败走公共Word撤销快照。该版本构建通过，正在新Word验收。

`stage03-original-preservation-comparison.json`：八份原始文档全部完整COM逐字段相同、Saved=false。

### 阶段03b–03g 引用创建的真实失败与修正

03b文档60、03c文档67出现引用插入空对象异常。03d文档73增加步骤诊断，明确为 Word 在 Field.Code 内 Fields.Add 后返回 null；公共恢复即时核验 body/mathFont/metadata/variables 全部 true，原公式和编号保留，没有部分引用。这些失败截图均实际查看，状态栏错误没有被“脚本退出码0”当作成功。

03e文档75改为从真实外层域重新获取唯一嵌套 REF，并校验类型、目标和范围；虽然不再抛异常，但截图仍显示内部 VTREF 占位文本，判为失败。03f文档77先删除严格匹配的本次私有文本槽，再在域代码内部的折叠范围插入；编号显示1，但 GOTOBUTTON 附带的 MERGEFORMAT 开关被显示出来，仍判为失败。

03g关闭 Fields.Add 自动附加的 MERGEFORMAT，由已有公共字符格式机制保持格式，并验证嵌套 REF 结果等于目标、无占位文本、引用尾部无额外文字。没有重试内容写入或删除用户书签/边界。此版本同时删除未被调用且固定 Cell(1,3) 的 RestoreReferenceBookmarkAliases，VisualTeX/MathType 别名重定位均使用同名 Add 后精确范围校验。正在真实新空白文档验收，尚不宣告完整引用链通过。

### 阶段03h正常引用、cases与阶段03i邻接转换

文档81新空白经Ribbon插入VisualTeX编号公式与引用，截图为完整(0.0.1)、无多余空格，实际双击选择书签111:157。转MathType后改连续编号，公式与原REF均为(1)，原兼容名称不变。实际双击截图高亮右侧正确编号；Word将位于MTPlaceRef代码中的目标报告为外层宏按钮起点28:28（Selection.Type=1），其目标书签100:143位于同一MTPlaceRef。`inspect-navigation.ps1`只读保存宿主、选择和XML，不调用Goto/DoClick。跳转检查对这种原生表示单独记录同一编号宿主归属，不能把正常Word范围表示当作错误。文档79曾因过时屏幕坐标点击空白导致增加3个空段，属于测试输入错误，证据保留；跳转脚本新增实际GetPoint行高范围校验，坐标越界不发送点击。

文档83原拟MathType导入，但提交前UIAutomation已实际显示Word原生OMML，真实结果也为6OMath/5SEQ；此项只记录为OMML对照，总耗时24.5秒。修正辅助流程为Home/Down/Down选择MathType，并在提交前强制读回格式与编号勾选状态，防止用脚本退出码冒充产品验收。

文档84从新空白真实Ribbon打开导入，提交前确认MathType OLE和全局编号On；完整import-source.tex含6公式，真实结果6个Equation.DSMT4、5个MTPlaceRef、可见编号1–5。主会话44743ecf及六个公式会话全部completed，空右定界符cases没有语义失败。两页实际截图已查看，正文顺序及标题/列表可见，完整COM/XML保留。总耗时97.5秒；尚需最终性能对照。

阶段03i将插入引用、编号格式和显式更新编号共用ExecuteDocumentEdit事务；删除MathType格式重写的局部XML重插回滚，失败由公共Word撤销前缀恢复并验证。当前文档成功后才发布未来文档默认编号格式。辅助完整测试379通过、0失败/跳过。八份原件COM与阶段03完全相等、仍Saved=false（stage03i-original-preservation-comparison）。

文档85新空白插入左编号MathType勾股、右编号MathType欧拉、末尾VisualTeX行内高斯并加入宋体12正文，真实插入第二公式MathType引用ZEqnNum100244后全文MathType→OMML。得到2OMath/1保留OLE/2SEQ/一个2×3宿主；尾部OLE元数据、宽高完全相同，正文段落文本/行距/对齐/段前后保持一致，VTO范围精确161:195，未跨界。核心Word apply约2.34秒。此正常转换通过所记录内容和非目标保护。

随后真实按章横线编号变为(0-1)/(0-2)，ZEqnNum目标内容已为(0-2)，但REF仍为(2)；该引用更新项判失败，现场已查看。根因是OMML编号快速路径只收集VTEqNum引用、无剩余MathType对象时另一路提前返回。03j在公共编号入口收尾统一更新所有实际目标的直接/嵌套公式引用，支持转换继承的ZEqnNum名称；尚待新空文档验证。03j还记录转换失败会话为failed（原来可能停在committing），状态写入失败会并入原异常报告，尚待真实异常验证。

### 阶段03j/03k实际结果与阶段04进行中

文档87从空白建立左编号MathType勾股、右编号MathType欧拉、尾部MathType高斯及宋体12正文，再真实插入第二编号引用。全文MathType→OMML成功为3OMath/2SEQ/一个2×3宿主；末尾OMath.Type=1，原引用ZEqnNum100289保留。连续编号后实际公式(1)/(2)，REF从(0-2)更新为(2)。实际鼠标双击选择98:157，与ZEqnNum目标精确相等。全部截图已查看，证据stage03j-mathtype-tail-*及stage03j-inherited-reference-jump。

03k把显式打开公式与旧会话提交分清：仅当原生锚及编号所属行完整时允许读取Word当前内容并生成新的会话指纹；漂移恢复和提交仍严格检查指纹。读取不再保存CustomXML，也不吞保存异常。FormulaIds移除按最近距离选择重复锚的实现，重复物理归属明确拒绝；字号读取使用已捕获的会话元数据。构建通过，完整辅助测试379通过、0失败/跳过。

03k文档89准备单条编号OMML时，实际已安装主程序PID20944在11:34:04发生0xc0000374堆损坏崩溃，编辑器消失，Word尚未写入仍为空。会话1f4f37ba停在editing，无错误文本。Windows Application事件1000/1001和原主程序哈希已保存stage03k-main-crash-*，不计通过，也未归因于VSTO本次读取改动。再次经Ribbon打开由主程序新进程76424承接，正常完成单条插入；该主程序仍是此前安装的版本，最终需新NSIS同源验收。

文档89随后以Selection选择原生公式中的a并输入x（prepare-native-edit.ps1，仅等价用户操作，不写XML），真实编辑器正确显示x²+b²=c²。取消前后完整COM与CustomXML相等（stage03k-native-cancel-comparison）。再次编辑点击欧拉预设，编辑器按自身正常行为在当前源码后追加欧拉表达式，真实提交会话aea87ac1的内容与Word最终单个OMath一致，仍1OMath/1SEQ/1×3。不能把预设追加误称为替换成只有欧拉公式。

文档90重绘手动换行源码仍真实失败，2OMath只有1编号，错误为“must occupy its own paragraph”。截图已查看、COM/XML取证后确认关闭提示，失败文档保留。阶段04a将重绘和批量导入接入公共ExecuteDocumentEdit；移除批量按当前Selection删除大范围及InsertPreparedFormula的局部吞异常删除回滚。移除只剩一个调用者的TryReconcileFormula吞异常包装。

阶段04a重绘写入前按真实CR/BR分隔公式两端，邻接手动BR仅在精确范围无对象/书签/Frame时提升为段落分隔，不删除引用或未知边界。三种格式共用隔离后的精确宿主，MathType收到精确插入参数，已有正文不走新打字段落重置。已安装DLL373001531AD44EA951C26242D05D3C027385D4D6632A9B05D49D7614C64F57A6，新Word PID78632，正在从空白文档进行真实重绘验证，尚未通过。

阶段04a文档92/93/94各从新空白输入原审计两公式源码和手动换行，真实Ribbon分别重绘OMML/VisualTeX OLE/MathType。结果分别2OMath+2SEQ、2VT OLE+2SEQ、2Equation.DSMT4+2MTPlaceRef，实际可见编号均1/2，正文—公式—正文—公式顺序正确。所有结果截图已查看，`stage04a-redraw-prose-xml-comparison.json`验证三种格式的两段正文run与paragraph格式XML均与输入精确相等；COM为宋体12磅。八份原件完整COM与03i相同、Saved=false（stage04a-original-preservation-comparison）。

阶段04a文档95新增正常多段LaTeX源码对照（不是修改产品XML或故障注入）：第一式的$$和内容跨三段，第二式与正文同行。后方公式已物化后，前式因源码不止一段而失败。现场提示已查看，公共原生撤销核验本轮79条Word撤销历史前缀后恢复，body/mathFont/metadata/variables全true；外部前后完整COM也完全相等，0OMath/0OLE/0编号、end60，原源码和全部正文恢复。证据stage04a-multiline-*。错误确认关闭，文档95保留。

阶段04b将“源码可跨段”与“目标宿主必须独占段落”分开：先校验源范围首尾段边界，删除本次完整源码后以向前一字符探针取得唯一空段，再插入。新Word PID69024加载AD48C2DA4301A3C7FE6FFA009D8F400DAF7D98DBB640B9CB7023AAA94ECF84C5。文档97相同多段源码真实重绘成功为2OMath/2SEQ/两1×3，x²+y²=z²与E=mc²内容正确，四段正文依次保留且宋体12。stage04b-multiline-*截图已查看。辅助完整测试379通过，0失败/跳过。普通CR三格式、局部富文本、其他非目标对象、批量公共恢复后的A/B、后续转换和最终同源NSIS矩阵仍待继续。


### 阶段04c：用户指出的引用字体不一致

文档97只读完整字体检查确认：正文和引用外侧括号为宋体12磅，嵌套REF的代码与结果却为Cambria Math 12磅，两个编号为等线10.5磅。此前只核对公式字体和段落汇总未捕获这一差异，不能把该次引用验收记为格式通过。stage04b-reference-font-before.json保存逐域、书签、正文和段落标记的Name/Ascii/Other/FarEast/Bi、字号、粗斜体、位置等真实COM数据，原WordOpenXML也显示REF的Cambria Math节点。

公共修复：编号入口只把物理编号范围内的REF交给编号字形处理；正文直接/嵌套REF统一保留自己的字符格式，删除旧的正文引用强制Cambria路径和局部吞异常重试。新WordCharacterFormatting在插入、更新、MathType转换前后以及新OMML表格字符格式继承共用，补齐西文/东亚/其他/复杂文字字体与粗斜体、上下标、间距等属性，读写失败向公共事务传播。移除MathType重复的部分字符格式实现。新OMML编号表格仍重置所需段落几何，但在公式物化前保留源段落标记的字符格式，不再重置成Normal默认的等线10.5。

stage04c安装实际DLL 8B8C06C485671A2C6BC5F7110FA94D2C69BB9EE564E95B0484A49EBFC8E37EBA，正常COMAddIn新Word PID59088，MVID0788429a-eec9-4015-8bc4-a6defa7ca5ac。完整辅助测试379通过、0失败/跳过。

文档99由新空白经真实Ribbon重绘多段源码，2OMath/2SEQ/两1×3，随后插入第二引用。正文、两个编号、引用括号及数字均为宋体12，公式内容仍Latin Modern Math12。改连续编号后可见(1)/(2)、REF为2，上述字体仍相同；真实鼠标双击准确选中目标127:169。stage04c-reference-created/renumber及对应-font.json和jump保存证据，截图全部查看。

文档99随后真实全文OMML→MathType成功2OLE/2MTPlaceRef，正文和编号、REF仍宋体12。实际再改按章横线后编号(0-1)/(0-2)、REF0-2，字体保持；双击落到第二MTPlaceRef宿主起点247:247，目标319:407位于同一宏按钮，截图高亮正确编号。stage04c-reference-to-mt、mt-renumber、mt-reference-jump保存结果。完整方向矩阵、富文本、批量和最终同源NSIS仍未完成。


阶段04c文档99回转MathType→OMML后，引用仍宋体12，但两个编号退回等线10.5；随后OMML→VisualTeX保留了这一已有编号格式。两次截图均已查看，stage04c-reference-back-omml-font/to-vt-font记录，编号字体验收失败。根因是MathType完整宿主替换所用的干净临时段落丢弃了原段落字符格式，并非引用再次受污染。

用户明确验收规则：编号是普通Word域，默认继承当前正文的西文字体及字号，不取OMath/Cambria Math/OLE内部字体。04d在完整段落替换及多公式连续段落替换中按源段落逐一捕获并恢复字符格式；空宿主清理和公式语义校验继续严格执行。source/default/font三者分开。新阶段Word81992实际DLL AF14308984C8B6D4EFE35D32063FC71F8E8ECF8723159185A295FE70580E7398，完整辅助测试379通过、0失败/跳过。

文档101新空白准备中西文字体分开：中文宋体、西文Times New Roman，12磅。原始源码XML明确ascii/hAnsi=Times New Roman、eastAsia=宋体、sz=24。真实OMML重绘2公式/2编号/两1×3后，两个SEQ结果Name/NameAscii=Times New Roman、NameFarEast=宋体、Size=12，正文同样设置；公式仍Latin Modern Math12。stage04d-western-redraw截图已查看。引用改号及完整往返仍在继续。


文档101真实插入第二引用并改连续编号后，两个编号、引用代码/结果及括号的NameAscii仍Times New Roman12，中文宋体；OMML→MathType→OMML后同样保持，2公式/2编号及正文顺序正确，stage04d-western-*截图均查看。八份原始文档完整COM与04a相等、Saved=false，stage04d-original-preservation-comparison.json保存结果。

后续字体仍正确，但连续引用链失败：OMML→VisualTeX时，最初引用的VTEqNum_8efa55faebf749b787030fb2c0e35778已丢失，REF缓存仍为2，不能据表面值判通过；下一次VisualTeX→MathType刷新后才显示未找到引用源。stage04d-western-to-vt与vt-to-mt的书签/域数据确定首次失效位置，截图已查看。失败文档101保持原状。

04e把VisualTeX/OMML和MathType的转换别名捕获合并到CaptureNumberAliases：按可见编号及数字本身的实际Range、Story、完整文本和外部引用捕获两种受支持名称，继承的旧VTEqNum不再因GUID不是当前FormulaId而丢失。不删除别名后重造引用、不改REF指向。结束校验显式检查每个引用目标存在、引用数量精确相等、显示值等于目标；格式快照恢复不再用Math.Min忽略缺失项，目标名称解析不再使用前缀匹配。新Word PID75288实际加载DLL F00AB7641AEC82EBEFDDF31D25FC0F6947DE463900FB1FCB5CCFAA9D3BF9ED90。完整辅助测试379通过、0失败/跳过；真实连续链正在新空白文档重新验证，尚未通过。

阶段04e文档103从新空白以中文宋体/西文Times New Roman12准备多段源码，真实OMML重绘2式2编号，插入第二式引用。实际OMML→MathType→OMML→VisualTeX→MathType→VisualTeX均通过对应内容/数量/实际编号/正文顺序及可见数字字体检查；期间在MathType和VisualTeX阶段各真实新增第二式引用，因此现有3种来源共3条正文引用。原VTEqNum_6a319116fb6a4df0af12bee279bcc8ee、MathType的ZEqnNum100273和后来VisualTeX的VTEqNum_138a8895f65840fbb469682042e57e47均保持原名与正确范围。此前04d最先失效的OMML→VisualTeX及随后REF刷新路径已正常，不能仅靠缓存的检查由COM实际书签与域文本相互校验补足。

在MathType阶段真实切换按节横线后，公式显示(0.0-1)/(0.0-2)，3条引用均同步为第二式编号且Times New Roman12；真实双击第一条继承引用，Word Selection291:291落到第二MTPlaceRef宏按钮起点，书签363:495、目标0.0-2位于同一宿主，截图高亮正确编号。stage04e-mt-inherited-jump保存实际鼠标记录。上述截图均实际查看；可见VisualTeX编号来自REF，内部计数SEQ的独立字幕样式单独记录，不混同可见编号。运行模块MVID da22363b-a23c-4493-91da-81d25d01f5aa。完整六方向最后一步、保存重开和正文/公式不同字号对照仍在继续。

04e六方向连续转换完成后，文档103在OMML中真实切换连续编号出现新的失败：MathType来源的第二条引用由(0.2)变为无括号的2，其REF目标ZEqnNum100273被改为VTEqNum_138a8895f65840fbb469682042e57e47，而外层GOTOBUTTON目标仍为ZEqnNum。stage04e-final-continuous截图已查看，不计通过，文档103失败现场保留，尚未执行保存重开。根因为UpdateNativeCrossReferences在检查物理生成编号归属之前改写所有可解析别名。04f将归属检查移到任何改写之前，只处理物理生成编号，正文引用统一公共更新；编号格式/更新编号两个公共事务还捕获并验证既有正文引用名称、数量和实际目标显示，任何失败进入原生撤销恢复。此修改尚未构建安装。

阶段04e不同字号对照：准备脚本新建106后，Ribbon使用窗口枚举顺序误激活旧测试文档103，24字符正文和回车误输入103，随后的空编辑器已取消；未写入新公式。此为测试准备错误，原失败证据stage04e-final-continuous已在误输入前完整保存，103现状不能再称未改动。word-ui新建后显式激活窗口并把文档名称记入target-word.json，后续UI窗口精确匹配，任何COM活动文档不同立即拒绝输入；snapshot同样读取指定文档名。

从重新新建的107准备正文Arial14/宋体14，经Ribbon打开OMML编辑器，实际读回公式字号10.5、编号On，完成插入后1OMath（Latin Modern Math10.5）/1SEQ（Arial14）/1×3。实际引用插入又发现后续正文空段落早在插入公式后被重置等线10.5，引用因此为默认正文10.5，字号/字体完整链失败，stage04e-distinct-size-*截图已查看。04f原编号别名更新修复构建/379测试通过，实际安装6E9D017EDCCD71A5D0C2E25EF821A32C8D9CA2468C314257FFAB88762D52E365启动Word74892，但尚未在其验收产品路径。

04g收敛编号公式续写段落的重复处理：从实际可见编号宿主段落标记捕获原正文字符格式，供OMML表格的创建/拆分/复用及OLE编号后的正文段落共用。移除默认字体Reset及两份重复处理，字体失败不再吞掉，空段落必须通过无对象/引用书签的真实边界检查才可应用；保留段落几何和内部表格分隔规则。构建通过、379辅助测试通过，实际DLL6B73B94D2CA44C1E4F630A47A08AE3CB3A758C9BB649C9789B31E8CB5EFF020B，新Word14912，文档108正重跑不同字号对照。八份原始文档完整COM均与04a逐字段相等、Saved=false，stage04g-original-preservation-comparison保存核对。

04g文档108实际从空白插入正文Arial14、OMML公式10.5，公式后续写空段落和编号均Arial14；再真实插入引用，代码/结果/括号也Arial14，1OMath/1SEQ/1×3，截图均查看。XML后续段落明确ascii/hAnsi=Arial、eastAsia=宋体、sz28；旧107此处无rPr而退回Normal，stage04g-continuation-xml-comparison记录。实际加载MVID9d9857fd-1b87-4926-b422-de3ae5ad396f。

随后OMML→MathType实际失败，错误“The Word OMML source formula moved before conversion started.”，已查看stage04g-distinct-to-mt-result截图并取证/关闭提示。失败前后完整COM相等（stage04g-distinct-size-reference与conversion-failure），1公式1编号1引用保留，尚未开始Word变更。日志显示第一次插入时旧final-fingerprint-transient-fallback曾用输入源指纹替代Word最终指纹；转换准备已捕获当前Word指纹，但ResolveSimpleOmmlSourceRange仍调用重新读取存储元数据的旧GetEquationRange并catch{return null}，于是有效内容被报为“移动”。04h将提交统一到GetEquationRangeVerifiedForStructuralEdit(...target.Metadata)，去掉吞异常；只读OMML导出与最终指纹共用精确范围/COM数量校验，只有真实COM证明是空导出时允许一次重新绑定读取，仍然必须得到唯一OMath；删除输入源指纹回退。原生OMML→LaTeX计数的只读分支同步使用当前内容读取，变更分支仍严格验证捕获指纹。04h构建通过，安装后继续真实对照，编号别名04f后完整链尚未重跑通过。


阶段04h文档110真实创建OMML10.5、正文Arial14，编号/续写段落/新增引用Arial14，1OMath/1SEQ/1×3。日志仅在COM证明空XML导出时重新绑定一次，随后严格捕获Word最终指纹成功；原OMML→MathType误报移动路径已真实完成，得到1OLE/1MTPlaceRef，引用目标存在且显示1。该次MathType编号为默认10.5，随后依用户要求先暂停字体实现并作原生对照。

原生MathType对照详见NATIVE_MATHTYPE_NUMBERING.md：本机MathType.exe 2025.9.16.454，文档111用MathType自己的右/左编号和编辑器，首次正文Arial14创建含该字体的MTDisplayEquation；第二段正文改Times New Roman12，编号仍复用Arial14样式。不能将MathType编号定义为逐条跟随当前正文，也不能将缺少样式初始化造成的默认10.5视作原生标准。用户明确要求遵循原生MathType字体格式，04i据此只在首次创建样式时设置正文字符格式、Normal基准及后续段落，已有样式不覆盖；移除未安装的逐条编号字体覆盖尝试。

04i还保留编号OMML拆分/完整替换后的正文空段落字符格式，避免原MathType样式创建前就丢失字体；新增一字符向前探针明确插入段落归属。删除952行未被调用的两套MathType插入/重新物化旧实现，其余实际插入链继续使用同一配置入口。构建通过，实际DLL2232F21B9C0CC4D3A2E209B3AD086BD0F457A87B7D338FDC7A73003E60BCA492，新Word20812。辅助测试首次377通过/2因沙箱无法写受控Office临时SVG而失败，授权同一测试临时写入后379通过/0失败或跳过，275ms，两份TRX均保留。

文档113从新空白经真实VisualTeX编辑器先插右编号MathType，正文Arial14，编号(1)Arial14；下一段改Times New Roman12，再插左编号，得到2OLE/2MTPlaceRef，实际(1)/(2)均Arial14，两段正文各自字体/字号保留，MTDisplayEquation被复用。stage04i-mt-right/left的COM、XML、逐域字体与已查看截图记录这一原生兼容对照。字体契约不硬编码Arial或Times New Roman。既有引用别名重编号、OMML往返、保存重开及整个后续矩阵尚未完成，最终同源NSIS仍未构建。


04i后续文档113真实插入第二式MathType引用后，ZEqnNum100293范围/实际值正确，但引用新段落为默认正文10.5，不能计完整格式通过；与原生111保留源正文Times New Roman12的续写段落不同。该问题属于续写正文，不应通过把REF强制设为MTDisplayEquation字体处理，仍待修。

真实按节横线更新编号后两个MT编号均Arial14、实际(0.0-1)/(0.0-2)，ZEqnNum引用同步(0.0-2)，两段原正文格式保持。随后全文MathType→OMML发生真实失败并恢复失败：stage04i-back-omml-result截图已立即查看，COM由2OLE/2MTPlaceRef/end620变为2OMath/0OLE/无编号/end193，原引用只剩缓存；不能视为通过。错误为相同勾股公式的两个OMML身份无法唯一恢复；Undo记录原144/当前144/owned0，body=False、mathFont/metadata/variables=True。两份真实create会话completed，转换走source-MathML-bypass没有新会话，不能伪称有failed会话。真实错误来自已加载VSTO日志。提示已确定关闭，文档113保持失败现场，不保存、不继续转换。

初步定位：实际COM零宽VTOMML书签位于OMath起点38/97，但Word XML将该零宽书签对写为w:body的直接子节点、位于公式段落之前，纯XML段内查找丢失身份，两个相同语义公式不能用唯一指纹兜底。撤销方面FinishBatchPresentation在EndCustomRecord之前关闭OMML临时源文档，应用级CustomRecord消失。04j先单独将源窗口关闭前结束当前记录、关闭后为编号收尾建立新记录，保留严格身份错误来通过原真实失败验证恢复，再处理同内容不同身份的正常转换。04j尚未构建验收。

04j实际安装DLL89DED2412BAE52F125EF4CEB1501AAA970813C3027B57B30FFFBF19CB01FCA16，Word87080/MVID3f2219ce-a2f9-48d9-b472-1069fad7686f，新文档115两条同内容MathType分别右/左编号，(0.0-1)/(0.0-2)、2OLE/2MTPlaceRef，正文Arial14与Times New Roman12，引用ZEqnNum100364为(0.0-2)，end602；stage04j-before保存完整COM/XML/字体/截图及只读撤销列表，开始实际失败恢复对照。

04j文档115在未修改身份错误的同一真实失败链中完整恢复：2OLE/2MTPlaceRef/14域/end602，全部COM字段、逐域字体、段落字符格式、书签格式和撤销列表均与转换前精确相等；原生Undo仅撤销本轮新增1条，body/mathFont/metadata/variables均true。stage04j-recovery-comparison保存结果。恢复后真实双击引用，跳到第二MTPlaceRef起点293并高亮(0.0-2)。首次脚本误将书签End=498与Code.End=498的合法相等边界排除，原失败记录保留；修正为同Story的包含关系后再次真实双击通过，stage04j-restored-jump-verified记录。已查看两次截图。

04k将完整XML身份索引与精确COM零宽锚(Story/Start/End)交叉核验，解决Word把合法锚导出成w:body子节点的表示差异。未取消指纹、编号所属行或一物一身份检查；无COM证明的段落外锚仍不能按“下一公式”认领，歧义指纹仍拒绝。增加5个辅助反例/正常对照，共384测试通过。实际DLL002C1EEDB0AFFA7894CD49ACFBACDD13AB8103658D653CA1A8E88102194A186E，Word79048。

04k文档117新空白重建两条同内容MathType、两段不同正文格式与第二式引用。准备动作一次因明确窗口模式未建立COM句柄而报错，尚无输入，截图确认空白后修正准备脚本连接模式。真实全文MathType→OMML的初始身份索引通过，但编号收尾的另一公共ResolveEquationIdentity仍要求已存在编号，把尚待编号的公式当成漂移而失败；提示已立即查看和取证。公共撤销再次恢复所有COM和逐域/段落/书签字体，stage04k-recovery-comparison全部true，2OLE/14域/end602。该转换不计通过。

04l在公共只读/编辑/编号解析器区分“编号意图已记载、宿主待建立”：仅精确零宽锚、非空且匹配的既有指纹、独占段落且无表格/其他公式/OLE/域/Frame/ContentControl/前后正文、无已有同名编号时允许此状态；已有编号仍检查实际所属行。去掉这条身份路径将COM异常吞成false的旧catch，并校验Story。实际DLL61065499E8FBBCFAF2488C7EA2107C22B78AC910F4F06BBA9AE071A22819FE3D，Word88240，384辅助测试通过。

04l文档119从新空白重跑相同两条MathType：全文转OMML已真实成功，2OMath/0OLE/2SEQ/两个1×3表格。实际编号(0.0-1)/(0.0-2)、Arial14，正文Arial14及Times New Roman12保留，ZEqnNum100364别名与引用值保留。真实改全文连续编号后(1)/(2)、引用(2)，双击实际跳165:224。只保存/关闭/重开此自建文档为stage04l-identical.docx，再次实际双击跳转通过，截图均查看。保存重开前后对象/域/书签及可见字体保持；COM不完全相等：两个表格行尾虚拟标记从Arial14变为Normal等线10.5，逐域报告preceding随之变化，stage04l-reopen-comparison明确保留。两个OMML公式均14磅，后来发现MathType字号读取取到了编号宿主字体，不能计入完整公式字号通过。原八份完整COM仍与04a相等、Saved=false，见stage04l-original-preservation-comparison。

04m将MathType直接插入与批量插入续写统一到EnsureBodyTypingParagraphAfterDisplay/NormalizeBodyTypingParagraph：按完整公式段落末尾CR插入/检查正文段落，并恢复插入前捕获的正文字符格式，避免右侧编号前错误断行；表格内限制在同一单元格。移除重复的批量MathType续写实现。384辅助测试通过，安装DLL02C3E3FA3ED74E86002B6C6110300446B230A5743E9DC5E5BEFE50B4BE3E457C，Word90792。新文档120两次真实插入后，续写分别Normal Arial14和Normal Times New Roman12；两条MT编号仍沿用首次建立样式Arial14，实际(1)/(2)。从Ribbon插第二式引用ZEqnNum100279，其代码/结果/括号Times New Roman12；实际改按节横线，编号/引用同步(0.0-2)，编号Arial14、引用Times New Roman12，真实双击高亮第二MTPlaceRef通过。截图与COM/XML/字体证据全部保留。此阶段直接插入字体对照通过，批量、重绘及表格内续写尚未通过验收。

04m新文档122真实批量MathType+全局编号出现E_FAIL，错误截图已立即查看。COM/XML仍为空白0对象/end1，七个会话在错误框开启时committing；关闭已取证错误框后，七个真实会话全部failed并保存原错误文本（stage04m-import-closed-runtime）。旧代码先等待模态错误框关闭再排队会话收尾，不能把开启期间的committing当作永久服务失联。具体E_FAIL调用栈尚缺，正在相同已安装版本的新空文档启用现有日志重跑，不算通过。

字号补充诊断：两个真实创建会话请求14/12磅，但从stage04m-second-__120.xml的实际OLE二进制只读提取MTEF，其EQN_PREFS Full均12磅、初始FULL记录。inspect-mtef-size.ps1只读已取证XML并调用现有codec解析，不写Word、不创建服务对象、不替代真实复现。创建路径使用固定12磅MTEF前缀，读取路径使用shape.Range.Font.Size（编号宿主14）而非内部字号；两条公共机制均尚待整改。字号与原生MathType编号样式是不同契约。

04m-debug在相同已安装DLL的新Word75316/文档123，通过原有真实Ribbon重跑并开启原有bulk日志。准备时FindEl误定位同名Text而非ComboBox，格式仍为OMML，脚本校验中止且未点击导入；已修正辅助UI定位为确切ComboBox并验证键盘焦点，再只读确认MathType后点击导入。真实E_FAIL调用栈证实发生在WordUndoHistorySnapshot.ReadRecords的原生_CommandBarComboBox.get_ListCount；六个MathType原生预览已成功，文档仍0对象/end1。实际关闭错误框后只读Undo状态：内置Undo和Standard原生历史控件均disabled。04n据此先交叉检查两个原生控件状态，只有都禁用才返回空历史；启用却无列表、两者矛盾、读取异常仍拒绝。

04n同时将批量失败和转换失败统一为有5秒超时且可等待的实际会话收尾，错误框之前保存完整错误；批量父会话创建后立即记录所有权，取消/解析错误不会漏掉父会话。成功先释放Word操作门，再等待会话完成；完成状态记录失败明确提示文档已完成但状态未记录，不能将已成功写入误报失败。移除旧批量fire-and-forget吞错实现及其接受测试延迟入口，bulk完整错误也写入既有Word Hook诊断。384辅助测试通过/0失败或跳过，290ms；实际DLL6B8F9C76BA7CE6E6C0B06794EDE79F5F209E993361A760816613E0DEFF9CD816，Word84748。

04n文档125从空白实际批量导入成功：6MathTypeOLE、5MTPlaceRef、35域、17段/end1476，原生预览896ms、Word写入2102ms，含实际用户窗口等待总15462ms。七个真实会话全部completed，stage04n-import-runtime保存。正文顺序和已查看屏幕公式正常，公式内部字号仍是已知待修问题。插第五式引用后真实切连续编号，发现章/节隐藏计数重启未同步：实际序列1/2/1/2/3，ZEqnNum101299和REF均(3)，不是应有的(5)；错误结果COM/XML及截图保留，不能计通过。双击辅助脚本因GetPoint返回无有效矩形停止，尚未发送点击，不计跳转通过。

04o在公共MathType编号格式计划中检查原生MTEditEquationSection2及其三条子SEQ精确归属，按新格式的计数范围修改仅MTEqn子域的restart/current开关，保留整个原生章/节域、段落、MTChap/MTSec及全部书签。全文连续不重启，按章仅章变化重启，按节章或节变化重启；未知或歧义所有权拒绝。调用方已有严格本地Undo恢复，未新增全篇XML替换。构建并安装DLL15D63C756CBF79B3CB794D23B7DFACE035C469FD23D8E51ABCCF605A56663274，Word82668，真实验收正在进行，尚未计通过。

04o新文档127真实MathType导入仍为6OLE/5MTPlaceRef/1原生章节域/end1476；7会话completed。插第五式引用后连续编号实际1至5，REF和ZEqnNum101299均(5)，真实鼠标双击跳第五MTPlaceRef起点872、目标943:988，已查看高亮截图。切回按节后REF恢复(2.0-3)，原生章节域仍1、全部6OLE二进制及6张预览二进制逐哈希完全不变，正文范围/首字符/段落标记格式相等，书签名保持；stage04o-counter-format-comparison保存。只保存并重开此自建文档为stage04o-mt-import.docx，6OLE/37域/1书签/end1551不变、Saved=true、引用实际(2.0-3)，已查看截图。此次确认计数规则切换及持续性；按章遇多个节的独立对照尚未跑，导入后编辑/所有格式转换和内部字号仍待办。实际发现源chapter/section均被映射为一级标题、首个列表项丢项目符号的正文问题，需与OMML/VT导入公共段落路径继续对照，不能据正文文本存在即称完整样式通过。

04o原始保护复核：八份完整COM均与04a逐字段一致，Saved=false，stage04o-original-preservation-comparison记录。MathType导入Word写入2007ms、原生预览949ms；整个操作106842ms包含人工检查导入窗口的等待，与04n总15462ms不能直接作为产品性能快慢比较。

用户再次明确三格式批量导入均须验证，并且每种导入后的编辑、关闭、转其他格式和往返也必须完成；后续矩阵按空文档/已有编号两组、OMML/VisualTeX OLE/MathType三种原始格式执行，单种成功不代表其余通过。当前04o新文档131真实OMML空白导入6OMath/5SEQ、一个2×3及三个1×3，全部公式Latin Modern Math10.5，正在插引用及编辑第二行验证。VT OLE当前版本批量仍待跑，不能宣布整批完成。

04o文档131的第二行编辑：关闭编辑器后完整COM相等；实际预设按钮会追加源码，追加勾股定理后仍6OMath/5SEQ，其余五式不变。随后在真实可见源码框Ctrl+A替换为a^2+b^2=c^2并更新，实际第二行被整式替换；引用VTEqNum_24ea394543294f6d9a9a2256bec17893仍2。全文OMML→VT成功6OLE/5SEQ，真实双击引用跳358:410。VT→MathType成功6OLE/5MTPlaceRef，真实打开第三OLE编辑器读回勾股定理，源码替换E=mc^2并更新成功，6OLE/5编号保持。截图已逐次查看。OMML批量的五条编号为Cambria Math10.5，与正文西文字体不一致；VT转换后可见REF为正文10.5，但内部SEQ使用题注样式11。该字体问题尚未修复，不计整体通过。

纠正列表证据的范围：04o的MathType原始批量首条列表项确实缺失numPr；同阶段OMML原始批量的两个列表项都有相同numId=1/ilvl=0及列表段落样式，不能称OMML也漏首项。chapter和section均一级标题的问题两者都存在。

04o上述文档MathType→OMML往返失败，已查看并保存真实错误框。原始错误为OMML身份8901cf2e…内容签名不符，恢复报“could not be verified”。stage04o-chain-recovery-com-diff完整COM差异0；插件自己保存的规范化XML仅六个VML宽高不同（例如134×35→134×34.95），后者与转换前真实COM宽高完全一致，二进制/文本/域/书签/其余XML没有差异。不能以数量相同单独证明恢复，也不删除尺寸检查。

04p新增公共WordInlineObjectGeometry：只读捕获作用范围内每个行内对象的精确范围、类型、ProgID和实际宽高；与XML内OLE数量、顺序、类型及ShapeID交叉校验。仅MathType已观察的整数点预览尺寸与对应实际COM尺寸，允许在仅用于哈希的证据中采用同一实际值；实际几何另加入签名，任意亚点尺寸变化仍不同，其他XML尺寸/布局/对象字节严格保留。文档级及行级恢复共用，不写XML回Word。新增正反例验证尺寸变化、错配、未知尺寸不得通过。

转换公共签名补充ISO默认分数类型bar（省略type或val均为bar），与noBar/lin/skw仍区分，分子分母内容检查保留；依据[Microsoft FractionType/ISO说明](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.fractiontype?view=openxml-3.0.1)。04p安装DLL EA326DB89BD8666901B5A1C8459D94B570984496A6B20D22C030E97683FF06E5，Word70992，新文档127（与其他进程同名，操作始终PID+活动名称守卫），真实MathType批量6OLE/5MTPlaceRef，Word写入1388ms，原生预览858ms，含窗口等待总14758ms。

04p真实插第二式引用后转OMML仍被剩余内容签名拒绝，完整错误与截图保留。此次真实恢复日志body/mathFont/metadata/variables均True；stage04p-recovery-comparison逐字段差异0，6OLE/23域/end946均恢复。该次转换不计成功。397辅助测试通过/0失败跳过，299ms。

04q继续修同一签名：仅一个m:e参数的定界符不使用sepChr，即使内部是多行eqArr也无参数间分隔符；始末定界符以及多参数sepChr仍严格区分。数学运行中的Word星号编码复用既有Word星号恢复规则，literal/nor文本不折叠。399辅助测试通过/0失败跳过，290ms；probe-import-identities只读两次真实失败证据，全部六式各自唯一匹配；它不替代真实验收。

04q实际安装DLL28106D86412CC434B3EB613269CF87B26D2ABA9B2B488793F1659D6DBEA5807E，Word84536，新空文档129。真实MathType批量6OLE/5MTPlaceRef，预览900ms、Word写入1353ms、含窗口等待14377ms；关闭第三OLE编辑器后完整COM差异0，重新编辑实际替换为E=mc^2。真实插第二式引用ZEqnNum100297，再全文MathType→OMML成功6OMath/5SEQ/4表格，首两式2×3，其余1×3。实际编号格式改按节横线，引用更新(0.0-2)，真实双击跳234:297；逐域读数5SEQ为正文西文字体10.5。所有相关截图已查看。

04q随后OMML→VT得到6OLE/5SEQ，但实际编号第二式从(0.0-2)变成(1.0-1)，后段从2.0-*变3.0-*，因此不计完整转换通过。COM/XML和截图stage04q-mt-omml-to-vt留存；XML明确第二式宿主新增pStyle=1（标题1），原两个正文标题仍在。转换后真实编辑第三OLE改勾股定理成功，6OLE/18域/end1024保持，stage04q-converted-vt-edited截图已查看。引用目标身份与章节数变化必须分别验证，不能用缓存文本证明跳转正确。

04r统一普通Word编号字体入口，移除表格/非正文story刷新分支硬设Cambria Math；创建、SEQ/REF更新均从自身story的唯一段落标记捕获不可变正文字符格式，再用于编号。TextFrame编号先从公式正文段落初始化标记；不使用主文档坐标读取其他story。原生MathType的MTDisplayEquation首次初始化与已有样式复用逻辑保持。另新增WordParagraphFormatting，在已确认OMML编号宿主替换期间保留原段落样式及直接段落格式，防止表格末行转文本时继承后方Heading1而增加章节；不重置后方标题、不写XML回Word。

04r构建通过，399辅助测试通过/0失败跳过，303ms；安装DLL879DDA6713EA333CB42B6B58C26BD3CBF55C43BDED721B988784FA5BBC1E8F05，新Word84580、真实空文档132开始OMML导入复验。此记录时04r真实编辑/转换链尚未通过，三原始格式×空白/已有编号×编辑关闭/更新/两个目标及往返矩阵仍在进行；最终同源NSIS及安装后全矩阵尚未完成。

04r真实空白OMML批量6OMath/5SEQ，5个普通编号均为正文西文字体10.5；关闭第三OMath编辑器完整COM差异0，修改第二式E=mc²成功。OMML→VT后6OLE/5SEQ，两个原标题及实际0.0-1/0.0-2、2.0-1至3保持，04q额外标题污染已修复。VT再编辑成功，但VT→MathType在NormalizeFormerNumberedFormulaParagraph的Range.Delete真实报0x800A1710；截图已查看，失败前后完整COM差异0，插件恢复body/mathFont/metadata/variables全True（stage04r-vt-mt-recovery-comparison）。旧引用双击仅落隐藏题注范围，屏幕没有跳到公式，明确不计可见跳转通过。

04s公共编号VT段落替换：验证唯一OLE/EMBED及无正文的已知公式宿主，移除该公式自有编号后，一次替换已证明的完整段落为LaTeX桥接内容，保留原段落/字符格式，收敛转换与还原LaTeX路径。文档134从VT原始批量6OLE/5SEQ，第二式编辑后→MT6OLE/5MTPlaceRef，再编辑勾股定理→OMML6OMath/5SEQ，再编辑E=mc²→VT6OLE/5SEQ。截图逐步查看，后两格式关闭编辑器COM差异0；这不是整体验收通过：原始VT导入书签有范围漂移，首次关闭编辑器改变了一条VTO范围，首项列表缺符号，引用插入因可见编号带Tab而失败，MathType内部字号仍待修。

04t安装DLL9B1E83F90D737B4FE6DC9D752F1C25E9BED91321747523071F2B59169A35D3BF，Word100440/文档136。批量完成统一按内嵌FormulaId与唯一物理OLE核对VTO，6个书签均精确覆盖各自对象；真实关闭编辑器后完整COM差异0。段落边界以所属story的向前一字符定位，两个列表符号已正常。普通REF取值目标与GOTOBUTTON可见导航目标分开，兼容既有VTEq范围的前导Tab；真实插入第二式引用成功，双击屏幕跳至第二式并高亮可见编号，stage04t-vt-jump-observed/destination保存双层证据。编号/引用常规错误增加Word所属错误框与日志，成功引用也记录完成，UI脚本不再只等待固定时间就把长转换算成功。

04t文档136第二式编辑E=mc²后VT→OMML真实失败，完整错误框已查看。六式内容各自唯一匹配，但E=mc²零宽锚漂移到相邻式；不能关闭指纹检查。恢复另有真实VTEqCap末式题注书签Start887变933（仅剩CR），完整COM差异2。插件恢复快照又因在草稿视图采集而误将OLE尺寸记录为半磅舍入值，恢复在打印视图比对，出现额外尺寸差异。失败现场保留，stage04t-vt-omml-recovery-comparison与recovery-xml-comparison记录，不能称恢复通过。

04u只先修恢复：在任何草稿视图切换之前采集原视图快照；删除自有题注前先解除原书签再删除Frame，避免Word撤销重建已缩短范围。Frame处理错误向上传播，重新绑定并校验原范围文本未变。405辅助测试通过，安装DLLFA6A32E23000D34F828C840F615576BE78E6C1C462BFD97FC77BC2E22085F31D，新Word101980/文档138正以相同真实失败检验恢复。编辑完成后一次准备光标调用RPC_E_CALL_REJECTED，已查看真实窗口、核对插件编辑完成记录后，才重新执行尚未发生的引用插入；没有重复编辑或忽视错误。

04v新增LiveInsertionOwners，保留InsertOmml已验证返回的真实OMath对象，在批量最终化将其当前COM范围与完整XML及原准备内容签名交叉核验；只在这些新插入对象证明归属时纠正零宽锚，不放宽已有编号行/旧锚的错配拒绝。408辅助测试通过/0失败跳过，289ms。此处04v尚未安装和真实验证。

前端文档导入会把解析后的blocks发给VSTO，旧已装前端仍把chapter和section都发为一级标题，因此04t只修C#源码解析不足。当前源代码同步前端层级规则，与C#共用6项JSON数据对照，代码块/注释不参与层级计算；文章section=1，书籍chapter=1/section=2。前端90项既有覆盖与6项共同对照通过。两个旧断言的初次失败记录保留：旧层级要求全部最高命令=1，以及Windows CRLF使8743逻辑字符变9010磁盘字符；后者改为只在字符数断言中规范换行，真实输入文件未改。前端尚未重建安装，标题真实复验仍待完成。

04u最终真实结果：打印视图快照已消除OLE尺寸伪差异，但最后题注书签仍缩短；除该书签标记之外的规范化XML完全一致，见stage04u-recovery-owned-span-proof.json（只读诊断，不是验收）。原始八文档04u完整COM均与04a相等且Saved=false。

04w合入真实OMath对象归属和公共WordBookmarkRecoverySnapshot。只捕获本次目标及明确引用别名的既存范围；原生Undo之后，先证明除已知漂移书签标记外全部正文、域、OLE载荷/尺寸、格式、其他书签完全一致，逐个核对原坐标文本，才通过COM同名Add恢复范围。最终仍强制完整含书签签名及元数据/变量/数学字体校验，未写XML回Word、未关闭检查。文档级转换和行级OMML编辑共用该机制。415辅助测试通过，负例覆盖正文变化、样式、域、非目标书签及未捕获名字。安装DLL B960833B214D910B872EBEDBA9614517EE1C752C99EA9CBD4C7231463A86CDB9，Word101568/真实新文档140。

04w真实结果：VT批量6OLE/5SEQ；关闭第三式编辑器完整COM差异0；第三式更新E=mc²成功；第二编号引用真实插入。VT→OMML通过六式内容/对象归属校验，随后旧全篇编号回退提前刷新尚未恢复的旧引用别名，报Reference number bookmark is missing。错误框和截图已立即查看。失败后原生Undo1334条（约122秒），公共恢复精确补回末题注VTEqCap 887:934；日志body/mathFont/metadata/variables全True，完整COM差异0，见stage04w-vt-recovery-comparison.json。此项可以记失败恢复通过，不能记转换通过。

04x删除转换编号失败后进入旧Shape最终化/全篇Reconcile的冲突回退，已知转换目标仅走公共1×3/N×3行宿主，局部构建失败保留原异常交由事务恢复；引用在全部目标书签恢复后再刷新。415辅助测试通过/0失败跳过，316ms，真实验证待运行。UI编排等待真实编辑会话释放并检查失败状态；光标和回车采用已授权的COM等价准备操作，减少前台竞争。前端/MathType字号/完整矩阵/同源NSIS交付仍未完成。

04x实际VT导入链：Word100856/新文档142，DLL BE4527937BF1A1FAC6A18BB8B7D50E14D15A00038CB5ACE8A5E051BD2E572216。带编辑、带引用VT→OMML成功：6OMath/5SEQ/1嵌套REF，2×3+三个1×3表格；六个真实会话成功，未进入全篇编号重建。第二行关闭编辑器完整COM差异0，E=mc²改勾股定理成功，只有第三OMath内容变化，其余五式及域代码/结果不变。Ribbon切连续编号后1至5正确，原引用0.0-2更新为2，真实鼠标双击跳到第二行可见编号，屏幕已查看；五个SEQ和REF的NameAscii均为正文西文主题字体、Size10.5。

04x不能记整体验收通过：首项列表符号在转换前快照已缺失（04w同样阶段正常，需按导入/编辑/引用分别定位）；旧已装前端标题层级仍未更新。随后OMML→MathType在第四OMath清理旧编号时Range.Delete真实报0x800A1710，命中RemoveVisibleEquationNumberLegacy；错误框已查看。失败后完整COM差异0，恢复日志四项全True，见stage04x-mt-recovery-comparison.json。阶段原始八文档只读复查均与04a完整COM相等且Saved=false，见stage04x-original-preservation-comparison.json。

04y公共RemoveFormulaNumberingArtifacts接收可选完整替换宿主，验证公式身份、独占1×3行及准确完整范围后，由后续原生整行替换统一移除公式和普通编号；不再对该行调用遗留编号Range.Delete，也不吞掉OMML宿主解析异常。兼容引用仍由转换原有捕获/恢复/验证机制维护。415辅助测试通过，304ms，安装DLL BB8E5E3C52B0A12FDE6AB066896B5B5410FF6C4FFA138B83190CA7823CC0F66C，Word105892/新文档144从OMML批量来源验证中。前端桌面与Office页面已build:all成功，但主程序尚未重新嵌入/安装，不把生成dist当已安装验收。

04y最终真实结果：导入、关闭编辑、更新第三OMath及引用后均6OMath/5SEQ，两个列表符号保持；OMML→MathType第二个替换行被完整宿主检查拒绝。完整COM恢复差异0，插件四项恢复均True；失败会话stage04y-failed-recovered保存，错误框已查看并关闭。04z增加行健康检查具体原因日志；文档146从空白直接导入再转MathType成功6OLE/5MTPlaceRef，尚不代表带编辑/引用链通过。另建文档147再次OMML导入，旧Reconcile的EstimateHeightPoints经过直接Range.WordOpenXML读取取得无OMath的瞬时导出而失败。原生Undo948条约90秒，实际恢复为空白0公式/0域/end1，插件四项恢复True，错误截图已查看。

04aa统一身份解析与编辑的完整OMML读取，保留COM计数/范围和XML双重检查；当前表格行先验证物理行身份，旧#SEQ宿主再做XML兼容检测。文档149带编辑和引用的OMML→MathType已完成末三式，到编辑过的第三OMath因存储指纹不符而拒绝；完整COM恢复差异0。只读实际CustomXML及Word公式诊断stage04aa-stale-fingerprint证明，只有E=mc²存储062780fa…与实际25e52097…不同，其余五式各自匹配。根因是ReplaceOmml最终化的性能分支把准备XML指纹直接保存，绕过Word落地后的最终指纹；没有放宽身份检查。

04ab移除上述编辑分支，始终读取最终活OMath，准备内容经ComputeImportedOmmlContentSignature等价验证后保存实际ComputeOmmlFingerprint。新增3项测试：Word空属性规范化后使用实际指纹，变量/指数变化仍抛错，合计418辅助测试通过。批量导入和转换收敛到BuildNewOmmlNumberingBatch，仅构建本批新行，编号规划包含已存编号；删除GetFreshConvertedOmmlRange重复捕获后回退同一解析的实现。安装DLL0A045B487B65E138CC2400BCC1BAB004532AB574491C760A57D1EBEEC3600FE7，Word27936/新文档151。导入6OMath/5SEQ，Word写入12634ms，相同空白OMML来源04aa为19633ms。带编辑和引用转MathType全部六式已落地，但原生编号刷新过早读取未恢复别名而失败；完整COM恢复差异0，四项恢复True。

04ac统一最终化顺序：先目标实际编号，再恢复兼容引用别名，最后公共引用更新；MathType拆分UpdateNumberFields和UpdateEquationNumbers，VT健康字段批次同样延后外部REF。移除编号已发生变更之后转全篇Reconcile的异常回退，保留异常给原有事务恢复。真实Word109548/新文档153从单个正常编号OMML再批量导入，7公式/6实际编号1至6，Word写入15047ms。第三编号式编辑E=mc²、插入引用3后→MathType成功7OLE/6MTPlaceRef；真实双击引用，已核对截图确实跳到E=mc²可见(3)，无隐藏Frame。MathType关闭编辑器COM差异0，改勾股定理成功；→VT成功7OLE/6SEQ，关闭编辑器COM差异0，改E=mc²成功。最后VT→OMML仍失败，不计整链通过：公共新编号接口误用调用方准备元数据，其指纹早于实际Word落地指纹。原始错误随后被缺引用书签异常遮盖。已保存真实错误框、失败会话、完整恢复COM差异0及插件四项恢复True，见stage04ac-recovery-comparison.json。

04ad/04ae补齐MathType内部Full字号，依据原生MTEF EQN_PREFS读取公式自身设置；创建和更新只修改Full尺寸值，保留其他尺寸、间距及字体偏好，实际编辑不再把OLE宿主Word字体当内部字号。8项正反例覆盖创建/读取、同字号保留完整前缀、变化字号及无效值拒绝，426辅助测试通过。此项需要真实安装后的导入/编辑/互转验证，辅助二进制测试不代表验收。

04ae将BuildNewOmmlNumberingBatch及转换健康检查接口收敛为公式ID列表，统一读取落地后已校验并持久化的身份；不允许调用方传递旧准备指纹。目标编号或别名最终化发生异常立即保留原始异常并恢复，避免继续执行遮盖根因。构建成功、426测试通过，安装DLL457699EFDFA3E7D4794070DE8A19084C99D103416E81C7266D41B79BC3A7D8B3，Word108116/新文档155正在从已有编号VT来源验证。八份原始文档04ae只读完整COM均与04a相同且Saved=false。最终NSIS同源安装及全矩阵仍未完成。

04ae真实复现原审查的VT已有编号漏号问题：元数据Numbered=true、真实导入页面全局编号勾选，但仅原先1SEQ，新五式没有实际编号。04af将VT导入与格式转换共用BuildNewVisualTeXNumberingBatch，从内嵌身份核对全部新对象后局部构建编号；完成阶段要求全部本批编号ID出现在健康实际域清单，不能仅检查已有题注。OMML与VT批量均不再调用可能全篇Reconcile的最终化回退。426辅助测试通过，DLL762489A7600333153C9E64DEC06D9805C62904CEFCDFE410FAE7B2AC22430E8F，Word81452/新文档157。先真实插一个编号VT再批量导入：7OLE/6SEQ、实际1至6，原式/三个原域/原公式段落均不变，两个列表项均有numPr；Word写入1711ms。第四OLE关闭编辑器完整COM差异0，源码改E=mc²及第三编号引用成功。

04af带编辑和引用的VT→OMML仍失败：新七式内容与身份配对通过，新六编号行完成创建后首式范围2:22的共享行归属检查拒绝，真实错误为The converted OMML row identities or number fields are incomplete。错误框已立即查看，Word原生Undo710条约99秒后四项恢复True；末题注书签由公共恢复机制还原1052:1105，外部完整COM差异0（stage04af-recovery-comparison.json）。不能记互转通过。下一步只补充TryGetManagedNumberTableRowIndex具体失败条件日志，尚未构建；未删除检查或猜测行号。

主程序04ae已重建，SHA256 0EFB2D2F00195DCF94E9878DA7BA2417D2C1FE36BE41841FA94EF7D3DE0A9585，嵌入index.html及当前JS/CSS表校验通过，尚未安装。2026-09-06 18:24用户要求关闭已验证旧测试文档以解除卡顿；按本轮真实创建记录及进程启动时间限制，用Document.Close(0)不保存关闭，进程无文档才Quit(0)。保留27664原始进程、81452当前157及历史恢复未确认现场；动作逐条记录stage04ag-close-verified.ndjson。未强杀进程，未保存或关闭原始八文档。

04ag收尾清理累计不保存关闭115份已完成测试及空白文档，正常退出50进程；最后清理时剩8个Word进程/3352MB。stage04ag-cleanup-final-summary.json与逐条关闭记录可核对；原始八文档stage04ag-original-preservation-comparison全部相等、Saved=false。旧文档编号会在新Word进程中复用，必须以PID及启动时间共同定位，不能仅按文档号码关闭。

04ag DLL8F63B42A61757E3A47ABDBE25BE459D963E3852364D5FD24017474A086BCA506，只在04af上增加行归属失败的具体条件日志。PID111272/新文档4：真实VT已有编号+六式批量得到7OLE/6SEQ，编辑第四式和插引用3后→OMML7/6→MathType7/6→VT7/6完成；实际编号1至6、REF3见stage04ag-chain-number-results.json。OMML及MT关闭编辑器完整COM差异0，二者实际更新编辑成功；OMML引用真实双击已查看截图，命中可见第三编号。全部MathType内部Full为10.5磅，真实编辑器五号。04af间歇行归属失败未复发，但没有证据证明根因解决，仍保留未解决项，不能归因于内存或仅凭这次成功关掉检查。

04ag文档6从空白真实MathType批量导入6OLE/5MTPlaceRef；直接cases屏幕与实际MTEF中msqrt均正常。先前文档4OMML→MT旧截图根号疑点不能算文档6失败；未据此修改产品。文档6关闭MT编辑器COM差异0、第三式改E=mc²、插第二编号引用、真实改按节横线后引用(0.0-2)，实际双击高亮可见MTPlaceRef通过。→OMML6OMath/5SEQ，cases实际显示正常，关闭OMML编辑器COM差异0，再更新第三式为勾股定理、→VT6OLE/5SEQ成功。对应stage04ag-mt-import-*证据。当前仍使用旧已安装前端，chapter/section两个标题均Heading1，不能算最终正文格式通过。

04ag发现编号格式公共路径遗漏：初始连续编号的MT文档切换按节后仅重写MTPlaceRef模板，未初始化Word标题所对应的原生章节状态，第三至五式为0.0-3/4/5；随后转OMML时公共Word标题计划产生2.0-1/2/3，属于同一格式设置下章节来源不一致，不能记完整互转通过。04ah在已验证MT编号段落位置上批量捕获公共Heading scopes，按范围倒序、每范围一次复用EnsureMathTypeHeadingScopeState；插入/转换/编号格式因此使用同一机制，已有原生MT章节标记仍保留。原生状态创建后仅按新插入位置定位MTEditEquationSection2，避免多章文档误操作全文第一条原生状态。保留实际段落增量检查、引用检查及原生Undo恢复，不做全篇重建。04ah构建/候选安装已完成，DLL6EF02523AE3D2905EB83E155D14F50F4ED648E0D4D32BEB0451A2209D5955988，真实验收进行中。

原生MathType章节语义补充：Wiris官方说明其章节数字由前方原生Chapter/Section Break决定，与Word分节不自动关联；VisualTeX自己的“按章/节”命令已在插入/转换时按Word标题初始化这些状态，本次补齐同一命令对已有公式的初始化，不改MathType原生Ribbon命令。来源：https://docs.wiris.com/en_US/using-mathtype-7/main-tutorials （Chapter/Section Breaks及Advanced equation numbering）。

04ah真实Word71532/新文档7：连续编号MathType批量6OLE/5MTPlaceRef，切按节横线后实际0.0-1、0.0-2、2.0-1至3，第三编号引用ZEqnNum100835=(2.0-1)。MT→OMML6OMath/5SEQ→编辑第四式E=mc²→VT6OLE/5SEQ→编辑第四式勾股定理→MT6OLE/5MTPlaceRef，编号和引用值保持。截图逐次查看，cases根号正常。此处旧前端chapter/section仍均一级标题，未计正文层级通过，引用可见跳转亦待本阶段实测。另新文档3先实际插入编号MathType再导入同源六式：7OLE/6MTPlaceRef，Word写入1479ms；该已有编号分支的完整编辑/互转仍待跑。

04ai发现未修改关闭的独立问题：04ah第四个复杂OMML实际打开后直接Alt+F4，真实会话b5c4271d-913d-4190-8d63-274e98644d36却dirty=true、autoCommitOnClose=true并completed。源码未改，原生来源没有formulaLetterFont元数据，旧前端用katex作为原指纹字体而界面已采用全局times，因此误提交并改变括号结构和后续位置。stage04ah-omml-close-comparison记录完整COM差异95；不能称此操作无变化，也不能称为用户取消失败。

04ai前端只保留一个不可变的完整初始界面指纹，等待偏好加载后以实际呈现状态建立基线；源码、编号、对象格式、字号和两类字体仍全部参与之后的改动比较。旧/native公式缺少字体元数据不再被“只打开”自动迁移；用户实际更改字体/内容仍须提交。canonical document回归通过、桌面及Office前端构建通过、主程序构建37.21s通过（62条既有Rust警告），候选主程序SHA256 1A6BB6016CB33F22E19E3DEE7836E23DB13F775DC3D0B0B78E5B8EF0CA980FDD，嵌入前端资产校验通过。build:all首轮沙箱写权限失败保留，授权后通过，不作为代码失败。

用户指出尚有工作不应提前打包；暂停NSIS。04ai仅备份并更新已安装主程序使用的23个Office页面资源（dialog-HL1bHzsS.js），主程序每次打开页面从安装目录读取文件，不需退出Word或重启OLE服务器。stage04ai-office-ui-installed.json记录原/新哈希和备份。此为开发候选真实Ribbon验证，不是最终NSIS或主程序/VSTO/OLE同源安装验收。04ai新文档10正在从空白批量MathType验证关闭/编辑/转换及前端标题层级。

04ai后续纠正：仅替换页面资源仍会复用已存在的编辑器WebView，第一次MT关闭dirty=true，不能算通过。确认无活动编辑会话后，仅备份/更新并重启空闲主程序为1A6BB601…（PID108292），未退出Word/OLE。新会话MT 0f4d7e17…、OMML 9ba9b0a8…均dirty=false，关闭后完整COM差异0；OMML原始XML仅100个rsid属性和1个settings rsid变化，不称原始XML字节相同。文档10实际MT→OMML→VT→MT→OMML并逐格式更新第四式成功，始终6公式/5编号；最后OMML与首次E=mc²编辑后的maths/paras/fields数组完全相等，引用ZEqnNum100835仍精确540:603、(1.1-1)，见stage04ai-omml-roundtrip-comparison。真实双击OMML引用已查看可见目标截图。原始八文档04ai再次完整COM相等、Saved=false。

04aj PID71532文档13由新空白VT编号公式加六式VT批量：7OLE/6SEQ，实际1至6。关闭第四式COM差异0，更新第四式并插第三编号引用后→OMML成功7/6，2×3的第二行关闭COM差异0，再编辑仅第四式内容变化。按节编号后引用(0.0-3)，OMML→MT7/6→编辑第四式→VT7/6完成；VT引用真实双击高亮可见E=mc²(0.0-3)，0隐藏Frame。04af同链行归属失败未复发，但未证明根因消除。MT阶段cases根号显示不完整，不能将整个转换记为通过。

04ak PID71532文档20新空白A/B：相同cases源码先OMML批量→MT，再直接MT批量。两个实际编号1/2、2OLE、8域。直接MT根号正常，经OMML转换的根号只剩细小痕迹；用户指出不是完全没有根号，记录采用“根号显示不完整”。两份真实Equation Native都含msqrt，内部Full同10.5，根号模板及尾部MTEF字节完全相同。stage04ak-cases-ab-visible截图已查看，原始COM/XML及纯数据只读提取完整保留。

04al纯数据补充诊断：从04ak真实Word XML提取两张WMF，异常图含完整根号折线/填充多边形，但多边形前无任何填充画刷定义；正常图有黑色画刷。共同MTEF写入器在matrix/pile槽引用COLOR 1，standalone前缀却没有COLOR_DEF；给异常数据仅增加RGB黑色定义后，通过已安装的MathPage渲染器得到完整根号，尺寸仍98×61磅、baseline不变，stage04al-native-request/response及brush-probe记录，重绘图片已查看。该诊断不连接Word、不构造服务对象，不替代真实验收。MTEF颜色定义编号规则参考Wiris官方MTEF v5文档：https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later 。

04al公共BuildRootStructure统一检查实际保留的前缀和根级颜色定义，仅在完全无定义时补充颜色1的RGB黑色定义；普通LINE及顶层对齐PILE共用，Create和Rewrite共用。已有全局/根级颜色表不重编号、不覆盖，公式结构及语义检查不改，不修改WMF遮掩问题。432项辅助回归通过（431ms，0失败/跳过），Word候选编译并安装CEC11D1427D70158CC31F6FFB0520B5206BC67A6023FA04783FCE4D3E7B912B6，新Word94656正在真实Ribbon复验。NSIS仍暂停。

04al PID94656/文档8真实A/B通过根号显示检查：新空白→OMML cases导入→Ribbon OMML→MathType，根号完整；继续直接MathType导入同源cases正常，2OLE/8域、实际1/2，两张截图已查看。未改动关闭第一式完整COM差异0；真实编辑第一式根号内m→n成功，第二式OLE及预览二进制逐哈希相同。插第一式引用ZEqnNum100100，真实按节重编号→MT→OMML成功2OMath/2SEQ、一个2×3、end332，根号分别n/2与m/2均可见。SEQ与REF均正文西文字体+西文正文10.5；引用实际(0.0-1)，真实双击高亮63:126可见第一编号、0Frames，已查看目标截图。这是该缺陷局部真实验收，不代表全量批量矩阵或最终安装验收。

04am PID94656/新文档11完整原始OMML批量：6OMath/5SEQ，一个2×3及三个1×3，全部Latin Modern Math10.5。第三式（第二行）关闭完整COM差异0，更新E=mc²成功；插第二编号引用VTEqNum_b73f44ea22c446eea52412013c0586ef，按节后(0.0-2)。→MT6/5，关闭COM差异0，更新第三式勾股定理；真实双击引用高亮可见第二编号，根号完整截图已查看。→VT6/5，关闭COM差异0，再编辑E=mc²→OMML6/5。该次返回后的六式文本/字体/字号/type与首次OMML E=mc²状态全部相等；9个纯正文非空无域段落的XML段落属性及运行格式完全相等，混合行内公式段落另行检查。→VT→MT仍6/5，前五个不同方向通过操作，但最后MT→OMML失败，不计完整六方向通过。

04am最后失败复现了公共行归属拒绝：20:34:29.733，managed-row-ownership-rejected check=initial reason=table-shape-or-range-membership，第三式范围226:237；这次不是先前04af的文首公式。调用方刚通过相同范围的表格成员与3列/非零行检查，公共入口重复读取时拒绝，具体哪项值变化尚待确认。真实原生Undo恢复559条操作（42.6秒，前置历史读取23秒），body/mathFont/metadata/variables全True；外部完整COM差异0、6OLE/37域/end1616恢复，stage04am-recovery-comparison。真实错误截图已立即查看/保留，取证后关闭。此MathType→OMML使用现有SourceMathML，不创建渲染会话，实际错误以插件自身完整日志为准；runtime-evidence查到的是之前编辑会话，不冒称它们为该次失败会话。

04an仅细化公共行归属入口日志为columns/rows/withinTable实际读值，并在已经启用的失败取证模式下保存Undo之前的完整真实XML；所有原有检查及恢复逻辑保留，诊断写失败会明确日志且不能阻止恢复。候选33B8D8C68842F83D9A7D6B8744C62EFEFFD8590657A9EFC206F276101FFDC97A编译安装，新Word112840通过实际MathType已有左编号/引用+完整批量路径定位；尚未证明行身份根因已修复。04al原始八文档再次完整COM相等、Saved=false。

04ao修复表格创建后新增空段落的归属。此前只有原后续段落含文字时才清理Tables.Add留下的新增锚点；遇到用户原有空段落，每次回OMML都多出一行。新机制捕获原段落末尾活Range及完整文本，插入后从原末尾重新定位，要求原文保持、相邻新增段落只有CR且无公式/OLE/域/表格/Frame/内容控件/书签，才删除这一确定由操作新增的段落；原空段落保留。异常交给原生Undo事务，不再吞掉异常。Word67232文档15往返两次完整COM仅公式身份名变化；文档20已有编号完整六方向互转，原空段落数始终1，公式7、实际编号6，三个格式的关闭完整COM差异0，原兼容REF可重编号并可见跳转。

04ap修复OMML公共字体处理额外修改混合正文段落网格设置的问题：移除StabilizeInlineOmmlFractionLineGrid及其编辑/字体入口调用，数学字体仍只作用于公式，段落网格和行距保留用户设置。同时把现有写入保护前移到SelectionChange的Ribbon宿主预读之前，字号同步读取和旧排队回调也遵守写入深度/代数；写入结束后在Word调度器刷新一次控件。此处隔离遗漏来自代码审查，尚不能认定它就是04am偶发行身份失败根因，身份/指纹/语义检查全部保留。

04ap实际候选6861688D7AECF188B7A7B4C79DABBFBE6D35763FD54ECCBDC8C37025F58054AC。Word77768文档21从空白经Ribbon完整OMML导入6公式/5编号；第二行仅改12磅，未编辑五式OMML字号保持，实际MT内部Full分别10.5/10.5/12/10.5/10.5/10.5；真实Ribbon显示小四。OMML/MT/VT关闭编辑器完整COM均差异0，三格式实际内容更新成功。首轮OMML→MT→VT→OMML仍6/5，包含混合行内公式段落的全部10个正文文本/运行格式/段落属性相等，旧04ao的同一证据比较为False，见stage04ap-roundtrip-prose-comparison与stage04ao-all-prose-comparison。此项是新正文修复的真实验证，不重新把已通过的字号/互转列为未修改。尚无NSIS及最终同源安装结论。
