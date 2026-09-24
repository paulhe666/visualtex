# VisualTeX Windows：真实 Word 回归审查与整改任务书

基线：`3d207d7a81ff40014fd861b242cff96ea77060a2`  
工作区：`C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7`  
审查日期：2026-09-06  
对象：该基线构建并实际安装的 Windows 1.2.6、Word VSTO、VisualTeX OLE 与原生 OMML/MathType 工作流。

本文件是交给实现 agent 的审查结论和整改要求，不是已经完成修复的声明。审查期间未修改产品实现。取证脚本、测试输入、XML/JSON、真实界面截图在本目录；不要把新增的审查脚本当作产品补丁。

## 1. 结论与优先级

问题不是一条“格式转换函数”出了错，也不是用户的 LaTeX 普遍不合法。当前版本存在几组明确的跨模块契约冲突：

1. **OMML 宿主不一致，且编辑失败恢复不完整。** 批量转换主动合并为 N×3 表格，后续编辑仍有强制 1×3、固定中心单元格的处理；真实编辑失败后从 3 个公式/3 个编号变为 4 个公式/2 个编号。直接关闭编辑器也能触发。优先级 P0。
2. **公式定位与邻接内容身份不能随结构修改保持稳定。** 转换前面的 MathType 段落时，后面的 VisualTeX OLE 身份书签扩张到新表格前；另一组顺序转换使行间公式与后面的行内公式发生编号宿主错配。非空边界和指纹检查捕获了部分问题，不能删除这些检查掩盖错误。优先级 P0。
3. **编号快速路径的检查集合不完整。** 批量插入跳过逐条编号物化，最后调用全局更新；全局快速更新只检查已经存在的编号产物，把“旧编号都健康”误当成“文档中所有应编号公式都已具有编号”。有一个旧编号时漏建新编号；空文档反而能正常导入。优先级 P0。
4. **MathType 引用迁移与编号重写的规则互相矛盾。** 转换代码主动把旧 `VTEqNum_` 引用别名绑定到新的 MathType 编号，编号格式重写却只接受 `ZEqnNum` 书签，拒绝自己刚生成的合法兼容状态。优先级 P1。
5. **批量 OMML 临时文档漏传数学字体。** 正常插入传递目标数学字体；一次性批量导入没有传入已有的可选字体参数，临时 Word 文档按 Cambria Math 物化，再以 FormattedText 复制，形成与直接插入不同的字体。优先级 P1。
6. **重绘混淆视觉换行与真实段落边界。** 同一段落内的手动换行、公式、后续正文没有被统一分割为受控范围；OMML 编号器拒绝混有正文的段落，生产包装层吞掉异常；MathType 的精确替换选项没有沿重绘入口传递，退回在整个段落末尾插入；新建公式后的正文样式重置还会污染已有正文。优先级 P0/P1。
7. **MathType cases 失败主要是当前语义往返校验的表示不一致。** 对真实渲染器输出做补充诊断，原式与 MTEF 解码结果的签名仅相差末尾空定界符 `o()`；微分、分式、矩阵行内容都保留。不能先归罪于 `\frac`、`\mathrm` 或缺少 `&`，也不能关闭整个语义检查。优先级 P1。

根本整改目标是统一“公式身份—内容范围—布局宿主—编号产物—引用依赖—字体上下文—事务恢复”的契约。只改报错字符串、给特定公式补丁、强删边界、让失败返回成功，都不属于修复。

## 2. 取证方法、边界与可信度

### 2.1 真实操作路径

所有主要复现均从真实可见 Word 的新建空文档开始。插入、编辑、格式转换、引用、编号格式、批量导入、重绘使用已安装插件的真实 Ribbon/编辑器/导入窗口。自动化通过 UI Automation Invoke/Expand、原生对话框按钮消息和实际键盘输入操作这些控件。没有直接写入伪造公式 XML 来冒充插入成功，也没有只构造一个新的服务实例代替已安装插件完成主复现。

COM 用于新建测试文档、定位所选对象、输入测试源码及只读结构取证。后续另有明确标注的“补充诊断”：读取实际转换结果的边界、书签，以及从真实会话取得 MathML 后调用已构建 codec 做纯数据往返。补充诊断不替代 UI 复现。

产品入口已查明：

- `VisualTeX.WordVsto/ThisAddIn.cs:226`：批量导入 Ribbon 控件绑定 `OnBulkImport`。
- `ThisAddIn.cs:524`：`OnBulkImport` 调用 `BulkImportAsync`；实现约在 2568 行开始，最后调度 `WordFormulaService.InsertBulkDocument`（约 2813 行）。
- `ThisAddIn.cs:525–536`：重绘回调；`RedrawLatexAsync` 约在 2095 行，最终进入 `ApplyLatexRedrawPlan`。
- `ThisAddIn.FormatConversion.cs:8–83`：选中/全文的六组双向转换回调，统一进入 `ConvertFormulaFormatAsync`。

因此，“批量导入没找到”属于早期自动化定位错误，不是产品缺少功能。导入窗口在独立 VisualTeX 进程，不能只在 Word 子树中查找它的选项。多窗口时，Word 的 COM 活动文档和桌面前台窗口也不一定一致。`ui-actions.ndjson` 中的定位异常和工具超时必须与产品异常分开；取证必须同时核对文档名称、真实对象结果、会话或原生报错。

### 2.2 原始现场保护

原始文档：文档1、22、26、28、29、33、35、36。首次读取时均打开且未保存。没有对它们执行公式编辑、转换、保存、关闭操作。结构资料以 `evidence/original-__编号.xml/.json` 保存，聚合资料为 `original-inventory.json`。

不要把文档35、36中已被用户多次操作的结果当作干净复现输入。不要根据截图的视觉重复直接断言存在几个独立 OMath；应结合真实对象、域、段落和元数据检查。

### 2.3 完成范围与限制

用户给出的每类操作链都已经实际执行；主要错误都有独立新文档证据，并补了关键对照。不是每个变体都会显示同一条异常。例如尾部为 MathType 行内公式时，本次左、右编号都触发指纹缺失；尾部为 VisualTeX 行内公式时触发非空边界。不要把编号方向当成唯一原因。

“转换后公式编号改变、正文引用仍停在旧编号”这个单独表现，没有在干净新文档中隔离成独立成功更新案例；本次 OMML/VisualTeX 两条来源路径均能复现更早的编号格式重写拒绝。整改必须将旧引用动态更新加入验收，不能声称本次已经验证该子变体的独立根因。

部分早期截图/脚本步骤存在窗口定位失败；最终证据以以下矩阵指定的文件为准。尤其 `audit-clean-ref-error-confirmed.png` 实际显示文档66，不作为文档62编号报错证据；采用干净文档70的 `audit-ref-final-error.png`。不得删除错误日志来让测试记录显得完整。

## 3. 原始活动文档的实际状态

| 原始文档 | 读取到的事实 |
|---|---|
| 文档22 | 一个 3×3 表格、4 个 OMML、2 个编号域；中间行公式单元格含两个 OMML，中间编号空缺。符合编辑失败后重复与丢编号。 |
| 文档26、28 | OLE 已是 Equation.DSMT4；MTPlaceRef 包内存在 VTEqNum 兼容书签，正文 GOTOBUTTON/REF 仍指向该名称。这说明存在显式别名保留，不是简单“所有引用都没迁移”。 |
| 文档29 | 两个左编号 MathType，后接 VisualTeX OLE 行内公式，当前没有 OMML 表格。此状态与失败恢复后的源文档相符。 |
| 文档1 | 两个右编号 MathType，后接 MathType 行内公式，当前仍是源对象。 |
| 文档33 | 多数未编辑导入 OMML 为 Cambria Math 10.5 磅；后面直接插入的对应公式为 Latin Modern Math 10.5 磅。用户已编辑过的一条导入公式不能再代表初始字体。 |
| 文档35 | 四个 OMML，仅两个拥有编号表格和编号；第一、第三个与后续正文处于同一段落，含手动换行。 |
| 文档36 | MathType 重绘区域的正文已经集中到两个公式之前，两个公式相邻；VisualTeX 区域存在正文和公式段落字号/字体不同。 |

以上是现场现状，不自动证明历史上每一个操作细节；下面的新建文档测试才用于因果对照。

## 4. 真实 Word 复现矩阵

所有证据路径相对本目录的 `evidence/`。测试编号是当前 Word 的临时文档名称，后续保存/关闭后可能变化，应优先按证据标签识别。

| ID | 新建文档、动作 | 实测结果 | 主要证据 |
|---|---|---|---|
| R1-Apply | 文档37：三个带编号 VT OLE → 全文 OMML → 编辑第二个 → 更新 | 转换后 3 OMML/3 编号/一个 3×3；编辑报 managed 1x3，留下 4 OMML/2 编号 | `repro01-three-ole-*`、`repro01-converted-*`、`repro01-edit-failed-*`、对应失败会话 |
| R1-Close | 文档59：同样新建三条 → OMML → 编辑第二条 → 不点更新，直接关闭 | 同样从 3/3 变成 4/2，失败会话错误一致 | `audit-close-before-*`、`audit-close-after-*`；会话 `942f7d4e-2789-4280-b04b-261be2511efe` |
| R1-Control | 文档69：直接插入单个带编号 OMML，再编辑更新 | 正常保留 1 OMML/1 编号/1×3 表格 | `audit-single-omml-control-after-*` |
| R2-VT | 文档40：VT 编号公式 + Ribbon 公式引用 → MathType → 改编号格式 | 状态栏明确拒绝 MTPlaceRef 内的 VTEqNum 书签 | `repro02-format-error.png`、`repro02-after-format-*` |
| R2-OMML | 文档70：OMML 编号公式 + Ribbon 引用 → MathType → 按节编号 | 同样报 non-MathType bookmarks，公式和引用仍显示旧格式 | `audit-ref-final-before-*`、`audit-ref-final-after-*`、`audit-ref-final-error.png` |
| R3-VT-inline | 文档43：两个左编号 MathType + 尾部 VT 行内 → 全文 OMML | non-empty boundary；失败后仍是 3 OLE/0 OMML | `repro03-left-vt-before-*`、`repro03-left-vt-after-*`、`repro03-left-vt-result-native.json` |
| R3-MT-right | 文档45：两个右编号 MathType + 尾部 MathType 行内 → 全文 OMML | fingerprint 2/3；失败恢复为源对象 | `repro03-right-mt-*`，结果 PNG 已核对 |
| R3-MT-left | 文档66：两个左编号 MathType + 尾部 MathType 行内 → 全文 OMML | 也是 fingerprint 2/3，并非仅右编号会触发 | `audit-left-mt-before-*`、`audit-left-mt-after-*`、`audit-left-mt-result-native.json` |
| R3-Control | 文档47：只有两个左编号 MathType → 全文 OMML | 转换成功，2 OMML/2 编号 | `repro03-control-*` |
| R4-Font | 文档49：完整源码，不编号，批量 OMML；再直接插入对应 OMML | 六条导入均 Cambria Math；直接插入 Latin Modern Math，字号同为 10.5 | `repro04-omml-import-unumbered-*`、`repro04-direct-omml-compare-*` |
| R4-Empty-OMML | 文档50：完全空文档，完整源码，批量 OMML 全部显示公式编号 | 正常 1 行内 + 5 行间，5 个编号 | `repro04b-*` |
| R4-Seeded-OMML | 文档57：从空文档先插入一个正常编号 OMML，再批量同源码并编号 | 精确触发 A numbered OMML table disappeared before row grouping；原先公式保留 | `repro04e-before-import-*`、`repro04e-seeded-omml-*` |
| R4-Empty-VT | 文档51：完全空文档，完整源码，批量 VT 并编号 | 六个公式，其中五个显示公式正常有编号 | `repro04c-*` |
| R4-Seeded-VT | 文档58：先建一个正常编号 VT，再批量同源码并编号 | 六条新公式插入，但没有新编号；总计仅旧编号。编辑导入的显示公式可见编号框为 On，应用后该条才获得编号 | `repro04f-before-import-*`、`repro04f-seeded-native-after-*`、`repro04f-unumbered-editor-uia.json`、`repro04f-after-edit-*` |
| R4-MTEF | 文档52：完整源码，MathType 批量导入 | 原 cases 触发 invalid standalone MathType MTEF，文档未插入部分结果 | `repro04d-mathtype-cases-*`，真实渲染会话 `a64f98c5-4cc6-467b-a6da-3915d5d9af3b` |
| R5-CR-Control | 文档53：两条源码以真实段落回车分隔，OMML 重绘并编号 | 两条均编号 | `repro05-before-*`、`repro05-omml-after-*` |
| R5-OMML | 文档63：同源码，匹配原现场的手动换行，宋体 12 磅，OMML 重绘并编号 | 两条 OMML 只有第二条编号；第一条编辑器编号框仍为 On | `audit-omml-soft-clean-after-*`、`audit-omml-soft-clean-editor-uia.json` |
| R5-VT | 文档64：同样手动换行输入，VT 重绘 | 两条均有编号，但中间正文从宋体 12 磅变成等线 10.5 磅，后段含混合格式 | `audit-vt-soft-clean-after-*` |
| R5-MT | 文档65：同样手动换行输入，MathType 重绘 | 两段正文被集中到前面，两个公式相邻 | `audit-mt-soft-clean-after-*` |

说明：早期文档54–56的自动化源码字体曾因 PowerShell 编码产生乱码字体名；它们的结构结果可参考，字号/字体结论以重新执行的文档63–65为准。

这组对照已经排除几种错误归因：不是所有 OMML 单条编辑都失败；不是带编号批量导入在任何空文档都失败；不是 cases 缺少第二列就必然无法编码；不是只要把左编号改为右编号就能解决转换。

## 5. 根因 A：N×3 生产者与 1×3 消费者，外加不完整失败恢复

以下源文件路径均相对 `apps/windows/src-windows/`，行号对应本基线；实现修改后行号会移动。

### 5.1 直接冲突

`VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.TableDisplay.cs:1535–1640` 的 `MergeAdjacentManagedNativeOmmlNumberTableRows` 主动删除可合并表格之间的分隔段落，把多个编号宿主合并为多行三列表格。`IsManagedNativeOmmlDirectTable` 在约 1676 行允许 `Rows.Count >= 1`，逐行检查中心公式和右侧 SEQ。全局转换在 `WordFormulaService.FormatConversion.cs` 约 2990 行使用这个分组合并流程；批量导入也在 `WordFormulaService.cs:6576–6590` 调用它。

但是 `WordFormulaService.cs:12861–12960` 中：

- `TryResolveNumberedOmmlFromKnownTable` 把多行表格排除。
- `ResolveNumberedOmmlFromNumberingOwner` 明确要求 `Rows.Count == 1 && Columns.Count == 3`，否则抛出用户截图原文；随后固定读取 `(1,2)`。

因此转换完成时的结构被分组合并器视为合法，却被编辑提交解析器视为非法。这是已确认的内部契约冲突，不需要假设 Word 随机破坏表格。

### 5.2 为什么报错后仍有重复与编号缺失

`WordFormulaService.cs:10430–11490` 的 `ReplaceOmmlCore` 有多个“健康 1×3”“健康独立显示公式”“旧式宿主”分支。识别不了多行宿主时不能安全使用只替换该行中心公式的路径。公式及编号可能先被改动，到后段重新按编号 owner 解析时才抛出 1×3 错误。

其 catch（约 11372 行）只针对部分分支删除 replacement 或 equationRange；原公式恢复使用 `originalOmmlStart` 和只捕获的原公式 WordOpenXML，再重新包装、保存、尝试编号协调。恢复过程的异常在约 11454 行被空 catch 吞掉。这不是对“原行公式、编号域、书签、段落及关联元数据”的完整事务恢复。`BeginUndoRecord/EndUndoRecord` 只把操作放在一个撤销记录里，代码中没有由这个记录自动还原全部失败变更的逻辑。

真实证据是最终多一条公式、少一个编号。不能以“catch 里有 rollback”或恢复函数的命名证明恢复有效。批量转换存在另外的整批恢复逻辑，不能将其与编辑恢复混为一谈。

### 5.3 实现要求

必须先决定一个一致的 OMML 宿主契约：保留 N×3 时，每条公式的身份是“表格 + 具体行 + 中心公式 + 该行编号”，不得把整张表作为一个公式；或者让生成路径保持单行独立宿主，并为已生成的 N×3 提供安全读取/迁移。不能只删除 `Rows.Count != 1` 检查，因为固定 `(1,2)` 会把编辑重定向到第一条。

所有插入、编辑、编号开关、字号调整、删除、引用、单个和全文转换都必须通过同一个宿主解析规则。提交前先验证目标身份及完整所有权；失败恢复覆盖实际会修改的全部局部结构和元数据，恢复失败须有独立错误记录。关闭自动应用和显式更新必须走相同事务。

## 6. 根因 B：邻接转换没有保护非目标对象的定位身份

### 6.1 不是简单“下面有正文，所以不能转换”

非空边界异常来自 `WordEquationNumbering.TableDisplay.cs:1593–1611`。代码先判断分隔文本是否只是结构空白；如果是空白但其范围关联书签、域、对象、表格或 Frame，则拒绝删除。保留这道保护是必要的。

在审查新建文档43中，原始 VT 行内对象的 `VTO_d57820ec06d4401d95ea3a356510a913` 范围是 `380:414`，仅覆盖该对象。随后用真实“转换选中部分”，先后转换前面两个 MathType：

- 第一次后书签变为 `276:310`，仍只覆盖原行内对象。
- 第二次后实际行内对象仍在后面，但同名 VTO 书签变为 `86:199`，包含前面的 CR、新表格及对象。
- 两张表格之间的 Range 为 `86:87`，文本严格只有一个 CR（字符13），没有公式、域、表格或 Frame；它却含上述跨越到后面对象的书签。

证据：`repro03-left-vt-before-__43.json`、`audit-separator-single1-__43.json`、`audit-separator-single2-__43.json`、`audit-vt-inline-separator.json`。

这把“非空”的实际来源定位到扩张的身份书签。修复若只是允许删除这个书签，后面的公式识别/编辑仍会错误，且可能破坏真实引用。

### 6.2 指纹失败也必须结合身份与宿主检查

`WordOmmlNativeSource.cs:165–230` 的 `RefreshFingerprintsFromDocumentOpenXml` 以单个 `pendingFormulaId` 扫描 bookmarkStart，再把后面遇到的 m:oMath 归给它。这个算法依赖“书签仍在正确公式前且其他书签不会覆盖待配对身份”的前提。它不是用公式所属行和实际范围来验证关联。2/3 只说明这套匹配没有得到三条正确关联，不等价于 Word 渲染器少生成了一条公式。

补充诊断中，将文档66尾部行内 MathType 单独转成 OMML可以成功；接着单独转换前面的第二条显示 MathType后，Euler 内容留在表格外，而原行内勾股公式进入编号表格，产生了宿主/内容错配。见 `audit-inline-mt-confirmed-*` 与 `audit-inline-anchor-drift-*`。这些是顺序单选操作的辅助定位证据，不冒称已经抓到全文批处理内部每一个 COM 中间步骤。

需要审查的生产链：

- `WordFormulaService.FormatConversion.cs:2060–2165`：转换顺序与源定位策略。
- 同文件约 2360–2590：删源、插入目标和目标存活验证。
- 同文件约 5737–5777：MathType 段落 body 删除，保留段落终止符。
- `WordFormulaService.InsertOmml`：插入后重新解析/包装公式。
- `WordEquationNumbering.cs:2320–2540`：批量编号阶段重新获取 converted OMath 并创建宿主。
- `WordEquationNumbering.TableDisplay.cs`：搬移 OMath、建立表格、删除旧段落及分组。
- `WordOmmlNativeSource.cs:165–230`：最后指纹配对。

这里已确认“邻接结构变更后身份所有权失效”及可导致内容/编号错配；单条 Word COM 语句的首次失效时刻仍需 agent 在上述链中记录变更前后范围。不要把未追踪到的精确时刻写成已确定的 Word 缺陷。

### 6.3 实现要求

变更一个公式时，不仅验证该目标数量，还要保护相邻的非目标公式和书签；每个后续步骤都使用经过验证的身份重新绑定对象，不能仅凭旧坐标、首个 OMath 或下一节点。跨段落迁移时必须处理起止边界的插入亲和性。全局流程应使目标内容、公式 ID、编号行、引用别名一一对应；指纹只能校验语义，不能代替所有权证明。

允许不合并的合法边界可以作为布局边界保留，但必须先证明不是自身范围漂移；不得为了消除空行强删正文、行内公式、书签或域。批量失败恢复后比较所有源对象及非目标内容，不只看返回数量 0。

## 7. 根因 C：引用别名迁移与 MathType 编号重写互相否定

`MathTypeEquationReferences.cs:671–763` 的 `RestoreFormatConversionAliasesToMathType` 会寻找目标段落的 MTPlaceRef 可见编号范围，并把捕获的别名重新绑定到完整编号或不带括号的数字范围。这个机制用来让原有 `VTEqNum_...` REF/GOTOBUTTON 继续有效，不能当作无用残留删除。

全局转换最终处理（`WordFormulaService.FormatConversion.cs` 约 3048–3120）确实调用别名恢复和引用刷新。

但 `MathTypeWordOpenXml.cs:126–205` 的 `RewriteMathTypePlaceRefFieldFlatOpc` 扫描该包书签后，拒绝任何不以 `ZEqnNum` 开头的名称。`MathTypeEquationNumbering.cs` 约 391 行在编号格式重写中使用它。这样，“转换产出的兼容引用别名”被“编号重写消费者”直接拒绝。

整改必须统一别名的类型、所有者、跨度与引用更新方式。可以保留旧名称并明确支持经过归属验证的兼容别名，也可以做真正完整的原子迁移：目标书签、全部 REF/GOTOBUTTON 依赖、括号范围、字符格式、跳转均一起改写。两种方案都不得粗暴放开任意陌生书签，也不得删除尚被正文引用使用的名称。

格式改变后，公式编号、所有引用的显示内容、点击跳转、章节/节前缀要一致。不是调用一次普通 Fields.Update 就算完成；需要检查嵌套 REF、旧格式缓存和格式保留开关。用户提到的“公式改变但引用不变”应作为独立回归用例，当前审查不把它冒充已隔离的另一个根因。

## 8. 根因 D：已有编号掩盖新公式的编号缺失

### 8.1 因果链已用空文档/已有编号对照确认

`WordFormulaService.InsertBulkDocument`（6322 行起）为显示公式设置 Numbered=true。但两条优化路径不立即建立编号：

- `InsertBulkOleDocumentTwoPhase`（7238 行起）调用 `InsertPreparedFormula(... bulkImport:true)`；native OLE 分支约 7645 行跳过 `TryReconcileShape`。
- `InsertBulkOmmlDocumentOneShot`（6649 行起）一次性物化、定位并保存公式/元数据，没有逐条创建最终编号宿主。

它们依赖约 6569 行的 `WordEquationNumbering.UpdateEquationNumbers(document)` 补齐。

`WordEquationNumbering.cs:361–490` 优先使用 `TryRefreshHealthyEquationNumbersInPlace`。其 `TryReadHealthyEquationNumberArtifactsFromOpenXml`（1191 行起）从已有 VTEq/VTEqCap/VTEqNum、SEQ 等产物构造集合，检查这些集合的一致性，却不将“此次新增的 Numbered=true 公式集合”纳入待满足条件。

结果是：

- 没有旧编号：快速路径无法成立，进入完整协调，于是批量编号正确。
- 有一个健康旧编号：旧产物自洽，快速路径成立，只更新它，跳过新公式编号物化。
- VT 路径保留缺编号的新 OLE，编辑器从元数据读到 On；编辑提交再局部物化，于是编号出现。
- OMML 路径随后按新 metadata 的编号 ID 找表格并分组，找不到未创建的表格，报“disappeared before row grouping”。在这组测试中，更准确的描述是“表格未创建”，不是已证明 Word 把它删除。

### 8.2 必须怎样修

把编号物化与健康编号刷新分开：批量插入必须向编号服务提供明确的 affected/expected formula ID 集合，并对这些公式建立/核实最终产物，再执行高效增量编号。快速路径必须证明预期集合完整，而不是只证明已存在产物完整。

不能简单让每次编辑都全篇 Reconcile、全篇刷新全部域来修复，避免把用户此前的性能收益全部撤销。对新增、开关编号、转换、删除，分别定义失效/刷新范围；对内容编辑保持局部处理。

完成标准必须检查真实可见编号数、编号身份唯一性及引用。元数据=true、函数没有异常、对象总数正确，都不是充分条件。

## 9. 根因 E：批量 OMML 临时文档丢失字体上下文

`WordFormulaService.cs:6773` 附近的一次性批量导入调用 `WordOmmlConverter.CreateWholeDocumentSource(_application, documentXml)`，没有把已经确定的数学字体传进去。

`WordOmmlConverter.cs:519–543` 的该函数本来具有 `string? mathFontName = null` 参数；`NormalizeMathFontName` 在 4138–4141 行把空参数默认成 `Cambria Math`。临时文档先按该字体创建，`WholeDocumentSource.Insert`（460–500 行）再以 `FormattedText` 将物化后的内容复制进用户文档。

与之对照，正常单条插入在 `WordFormulaService.cs` 约 3718、7531 行传入 `document.OMathFontName`。因此用户猜测的“批量路径没有跟上字体更新”在这里有明确遗漏点。

`ApplyBulkOmmlTypographyXml`（7106 行起）有意清除部分数学 run 的显式字体以使用文档级数学字体；不能因此直接给所有 run 硬塞 Latin Modern Math。应让目标文档、临时文档、准备好的 OMML、正常文本 run 与数学结构控制 run 使用同一字体上下文，并保留真正中文/正体文本的独立设置。

验收至少比较同一式子的直接插入、批量导入、重绘、转换、再次编辑，检查实际 run 字体、数学字体、字号及保存重开结果；不只比较工具栏里的字号值。

## 10. 根因 F：重绘复用“新插入”流程，未保护源文段

### 10.1 OMML 半数无编号

`ApplyLatexRedrawPlan`（`WordFormulaService.cs:4343–4548`）从后向前替换公式。`InsertPreparedFormula` 再复用单条插入路径。

`EnsureBlankDisplayParagraph`（8884 行起）遇到正文时在插入点打一个新段落，但这不保证公式后面的手动换行及正文被隔离到另一段落。第一条公式的后缀仍可能是 BR + 正文。

`WordEquationNumbering.TableDisplay.cs:603–620` 明确要求编号 OMML 自己占一个段落，否则抛出 `A numbered display OMML formula must occupy its own paragraph.`。`WordEquationNumbering.cs:2264–2301` 的 `TryReconcileFormula` 在普通生产模式 catch 后不继续抛异常；仅在 `VISUALTEX_VSTO_ACCEPTANCE=1` 时抛出。于是保留 OMML 与 Numbered=true 元数据，但编号没有物化。末尾全局更新又可能被前述健康快速路径挡住。

这解释了正常段落回车成功、手动换行版本只有最后一条编号的对照。不能靠把第一条元数据 Numbered 改为 false 来掩盖。

### 10.2 MathType 正文前移、公式相邻

`InsertMathTypeOle` 支持 `preserveExistingDisplayParagraphBoundary`（约 2205–2220 行）；它能够向下传 `replaceAtExactInsertion`。但 `InsertPreparedFormula` 的 MathType 分支（约 7414–7438 行）没有把自身持有的精确边界选项传入这条创建调用，而只在之后尝试清理多余段落。

最终 `ResolveDisplayInsertionRange`（8775–8861 行）在非精确模式看到当前段落已有正文时选择整个 `paragraphRange.End`。对于一个 Word 段落内含“正文—公式—BR—正文—公式”的输入，这会在删源码后把公式附到整段文字之后。公式本身与数量都可以正确，但顺序错误。

仅把一个 bool 传进去还不够：精确模式也不能允许把带编号显示宿主直接嵌进正文段落。必须先把目标公式的前缀、公式、后缀按实际 Word 边界分开，保留原文字格式，再以精确替换方式物化。

### 10.3 VisualTeX OLE 重绘污染正文格式

干净文档64已观察到中间正文从宋体12磅变为等线10.5磅。复用的新插入尾部流程 `MoveSelectionAfterDisplayFormula`（9099–9113 行）会在公式后 TypeParagraph 并将选区所在段落设为 Normal，再做格式重置；带编号路径还会调用 `EnsureNormalTypingParagraphAfterNumberedDisplay`（`WordEquationNumbering.cs:3487` 起），其中也包含恢复 Normal/Reset 的逻辑。

这些操作用于“为用户接下来打字创建新段落”，不能无条件用于“源码后面已经有用户正文”的重绘。需要以明确的生成段落所有权约束每个 reset，只清理本次新建空段；已有正文的字体、字号、加粗/斜体、段间距、段落样式必须保留。后置删除空行不能修复已经被改写的正文格式。

## 11. 根因 G：cases 的空右定界符导致往返校验误判

### 11.1 补充诊断采用真实渲染输出

先在文档52真实批量导入复现错误，然后从会话 `a64f98c5-4cc6-467b-a6da-3915d5d9af3b` 读取 `exportResult.mathMl`。`probe-codec.ps1` 加载本基线构建的 `VisualTeX.WordVsto.dll`，调用已有 CreateEquationNative、ReadEquationNativeMathMl、SemanticSignature，仅作纯数据诊断，不写 Word。

结果保存在 `evidence/codec-roundtrip.json`：

- 原真实 MathML 末尾有 `<mo ... fence="true" stretchy="true" ...></mo>`，表示 cases 的空右定界符。
- 输入签名以 `...o()` 结尾，MTEF 解码签名没有这最后一个 `o()`；其他数学结构签名一致。
- 最小数据对照 `{` + 两行单列 mtable + 空右 mo 同样不相等。
- 仅在补充诊断数据中去掉空右 mo，原实际公式签名即相等。
- 给每行增加一个空第二列仍不相等；因此不是简单缺 `&`/缺第二列。

### 11.2 源码原因与修复边界

`WordFormulaService.cs:2246–2257` 和准备批量 MathType 产物的对应路径执行编码—解码—语义签名对照，不相等则报 invalid standalone MTEF。

`MathTypeMtefCodec.cs` 约 689–723 行，`CanonicalizeMathMl` 的 mo 分支对空 Value 仍返回 `o()`；mtext 分支已经对纯空内容返回空签名。MTEF 输出不保留空字符 mo，二者表示规则不一致。这次不能据报错名称推断二进制损坏或所有微分命令不受支持。

修复需明确空定界符的数学/排版语义：对没有可见字符的定界符作结构化规范化，同时确保左括号的伸展高度、单侧围栏、矩阵行列都能正确生成和重读。不要用字符串全局删所有空节点，更不能跳过语义校验或把所有签名差异都忽略。

需补覆盖 `cases`、`\left\{...\right.`、`\left....\right\}`、单侧括号/绝对值、嵌套围栏、空条件列、真实空单元格、aligned 与矩阵。用户输入不用人为改写为某个特例才允许导入。

## 12. 事务与验收架构问题

当前不同入口的完成条件不一致：批量转换有整批恢复，编辑使用局部 XML 恢复，重绘的 Apply 末尾只有 finally 收尾，部分编号协调又吞掉异常。某条路径返回成功不代表编号/引用已经满足。

建议 agent 用明确的分阶段约束，而不是再叠一层“最后遍历修修补补”：

1. 读取并验证输入、对象身份、文档/故事范围、目标宿主、前后正文、关联编号和引用；捕获必须保留的状态。
2. 预先完成渲染、转译、字体准备和可验证目标数据，不先破坏原公式。
3. 以受控范围写入；每步维护身份、编号和边界，不让邻接对象的书签扩张成另一对象的宿主。
4. 验证内容对应、可见编号、引用依赖、正文顺序与格式、非目标内容不变，然后提交。
5. 任一步失败，恢复到进入该操作之前的用户可见与语义状态；恢复失败单独上报，禁止吞掉后仍声称已回滚。

性能优化必须带有明确前置条件和 affected ID 集合。避免全篇扫描不等于可以省略正确性验证；临时 Word 文档或内存 XML 的测试也不代替真实 UI 用户链。

## 13. 给实现 agent 的任务顺序

### 阶段一：冻结证据，先收敛契约

读取本报告及指定证据，确认 HEAD、当前修改与已安装二进制。不要 reset/clean/stash/切分支/删除未知文件，不得操作八个原始未保存文档。把当前失败案例、正常对照、字段/书签/正文不变量登记为可重复用例。先不要批量改写所有元数据 schema，也不要立即重写整个 Word 插件。

梳理并统一三种对象的身份解析和布局宿主描述。将同一行为的多套互相矛盾分支收敛；删除旧分支必须有调用证据，不能以减少行数为由取消保护。优先堵住编辑留中间态、转换错配与失败恢复。

### 阶段二：修复公共机制

按可独立审查的小步骤处理：OMML 行级宿主和编辑事务；邻接身份稳定与指纹匹配；编号物化/刷新集合；引用别名生命周期；临时文档字体上下文；重绘前后文隔离与样式保护；单侧围栏的语义规范化。每一步都运行受影响的真实 Word 链，再扩大回归矩阵。

不能通过调用安装在机器上的 MathType VBA 宏或原生命令绕开 VisualTeX 自主实现。MathType 可以作为兼容性对照，但产品功能仍应满足未安装 MathType 的既有目标。

### 阶段三：真实安装与最终验收

完成实现后构建正式 Windows NSIS，核对主程序、前端、VSTO/OLE、x64/x86 资源来自同一构建，安装后重新在真实 Word 空文档运行矩阵。旧版本安装包、工作区 DLL 反射测试或旧服务实例通过不能代替已安装版本验收。

本轮未授权提交、push 或发布 Release；agent 不得擅自执行。是否提交由用户另行授权。

## 14. 必须通过的验收矩阵

| 范围 | 最低要求 |
|---|---|
| OMML 编辑 | 直接单条、三个相邻转入、分散正文中转入；第一/中间/最后条；更新与直接关闭；连续编辑至少三轮；公式和编号数量不变，内容只改目标，无文末副本。 |
| OMML 宿主 | 单行和已有多行宿主均可读取；编辑、编号开关、字号、删除、再次转换不能错行；真 Display，不用行内字号补偿伪装。 |
| 转换与边界 | MathType 左/右编号×尾部无内容/正文/VT 行内/MT 行内/OMML 行内；单选和全文；核对每条内容—ID—编号对应，非目标对象与书签范围保持正确。 |
| 引用 | VT、OMML、MathType 三种起点先插引用，再走全部支持转换；连续/章/节及不同分隔符切换；已有引用、新引用、多个引用、保存重开、跳转均正确。 |
| 批量编号 | 真空文档、有旧编号、只有无编号公式、混合来源；本用户六公式源码应有六个公式、五个显示编号；新编号紧接既有序列，不从1重复。 |
| 批量字体 | 直接/导入/转换/重绘使用相同选择时真实数学字体和字号一致；支持既有字体选项，不把任何字体硬编码为唯一答案；中文/正体不污染。 |
| 重绘 | 原源码普通 CR、手动 BR、同段落前后正文、正文带局部格式；三种格式均保持正文—公式顺序，显示公式编号全部物化，正文样式/字号不变。 |
| MTEF | 原始 cases 无须改写即可导入；单侧围栏/嵌套/空条件列/矩阵/align；编码、解码、再编辑和保存重开内容及几何一致；保留严格语义校验。 |
| 失败恢复 | 渲染失败、编号失败、目标失效、指纹失败；检查公式、编号、引用、正文、书签、元数据都回到前态；不能只有对象数量相同。 |
| 性能 | 记录已安装版本的相同公式集时延；不新增全篇扫描到每次内容编辑，不以重绘/编号完整性换速度；与稳定对照比较，禁止明显回归。 |

每个失败用例修复后还应运行此前正常的对照，特别是普通段落回车重绘、空文档导入、单个1×3编辑和只有两条 MathType 的转换。禁止把某一条失败测试改成跳过或放宽断言后宣布通过。

## 15. 证据目录、运行方式与已知脚本陷阱

- `import-source.tex`：用户完整导入原文。
- `redraw-source.tex`：用户两条公式和夹杂正文。
- `word-ui.ps1`：已安装窗口/UI 控件操作辅助。
- `ui-flow.mjs`：真实操作组合；日志 `evidence/ui-actions.ndjson`。
- `inspect-word.ps1`：只读对象/表格/字体/域/书签/XML 取证。
- `inspect-boundaries.ps1`：只读相邻表格间边界、关联对象和身份书签取证。
- `runtime-evidence.ps1`：读取真实会话状态；参数为 DocumentFilter，不是 DocumentNumber。
- `probe-codec.ps1`：从真实会话输出做补充纯数据 codec 往返，结果 `codec-roundtrip.json`。

示例（从仓库根目录执行）：

```text
node docs/audit-3d207d7/ui-flow.mjs new agent-repro
node docs/audit-3d207d7/ui-flow.mjs insert native block right 勾股定理
node docs/audit-3d207d7/ui-flow.mjs next
node docs/audit-3d207d7/ui-flow.mjs convert vt-omml full agent-convert
node docs/audit-3d207d7/ui-flow.mjs snapshot agent-after
```

以上只是辅助操作，运行前须核对预期文档、编辑器会话及前台窗口。COM Activate 与 UI 前台不是天然相等；切换已有文档后的失败调用不计产品复现。部分 `snapshot` 原始输出包含 Word 数学线性文本的换行，不能把它们全部当真实段落；应同时检查 `Paragraphs` 与 WordOpenXML。Font.Size 的 9999999 表示混合格式读值，不是实际九百万磅字体。

屏幕截图包含用户桌面背景，仅作为本地审查证据，不要上传公开仓库、Issue 或 Release。审查文件未替用户做产品提交或发布。

## 16. 本轮交付的边界

已完成原始活动文档取证、全部主要操作链的真实新文档复现、关键正常对照、源码契约审查和整改/验收方案。详细的修复实现及修复后的验收由 agent 执行。

已确认的代码冲突和真实结果在上文给出；没有把“所有源码行已经逐条动态跟踪”“所有引用变体均已独立复现”“现有所有文档已经修好”写成事实。尤其相邻对象错配的首次 COM 边界失效点和用户曾出现的仅引用滞后变体，应在实施时进一步记录，但不能因此忽略已确证的公共身份契约问题。

## 17. 交付前核验

最后再次只读采集了八个原始文档，结果为 `evidence/final-original-inventory.json`。与首次快照逐字段比较公式、表格、域、书签、段落文本、字体、字号和间距：八份均无变化，且仍未保存。这个结论是已记录结构字段一致，不把它夸大为所有 Word 内部二进制状态完全相同。

`verify-audit-evidence.mjs` 对证据之间的计数、已记录原生报错、codec 对照、原文档保护、HEAD 与安装身份进行了43项一致性核验，全部通过；结果是 `evidence/final-audit-verification.json`。这43项是报告证据复核，绝不是“产品修复后43项验收通过”。575条 UI 动作记录中含自动化定位失败，已明确保留，不当作575个产品测试或575次成功操作。

已安装 `VisualTeX.WordVsto.dll` 与该工作区 x64 Release 构建 SHA-256 完全一致：`97841ac0f5ab7edf557c833d7182c71b27ed137b04846ab0776f658cf091a01a`。实际运行的 `visualtex.exe` 与原 NSIS 内提取到内存的 `visualtex.exe` 完全一致：`3029a2fc92f646fabeae17b4b7a2cf8820da63de926089f294a1086ad167f729`。没有通过重新安装来做这项核验，也没有改动正在运行的程序。

注意原始 `target/release/visualtex.exe` 的哈希不同，但逐字节对比只差偏移18350656–18350658的3字节，即 bundle-type 标记的 `UNK` 与包内/安装后的 `NSS`；正确安装身份应与 NSIS 的实际 payload 比对，不能仅凭原始目标 EXE 哈希不同误判安装了旧版。两份 EXE 大小均为22891008字节。

HEAD仍为3d207d7，`git diff --numstat`为空。Cargo.toml仍有工作树状态标记，但规范化 Git blob 与HEAD同为 `e65d7fc9e29748ad9b6f85a15c9147a2c2a0bc2d`，没有逻辑修改。新增文件仅位于本审查目录；没有提交、推送、发布或修复产品实现。
