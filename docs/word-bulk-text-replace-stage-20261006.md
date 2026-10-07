# Word 批量删除正文空格调查 — 2026-10-06

## CP314：最终结论与复现材料

**未稳定复现 VisualTeX 文档原生“全部替换”明显慢于 MathType/OMML；确认了更高的打开/保存开销，预览控制实验支持其中相当部分来自显示预览的处理。用户现场卡顿根因仍未建立，不能将这两个结论混为一谈。** 没有据此修改产品。

共90次真实Word操作：插件inert/隐藏Range.Find 24次、手动跨进程插件/可见Selection.Find 24次、真实进程内安装插件/可见Selection.Find 24次、200公式预览对照插件inert9次、进程内安装插件重复9次。每次是一个原生Find.Execute(wdReplaceAll)，不在计时区逐字符调用COM。每组3轮，含冷首轮、原固定顺序、不删异常值；均建立自有/x /Automation进程和文档，绝非用户文件。手动跨进程的24次内容结果有效，但回调耗时受跨线程与文档关闭重叠污染，不能用于断言安装性能。

真实进程内安装插件为**CP304**：SHA256 `279CA2F19998DEF308B69AE9BB8B8D121CF557B41CC484B44F8F564D02645C3E`，自身trace记录`C:\Program Files\VisualTeX\WindowsOffice\VSTO\VisualTeX.WordVsto.dll`、Word PID、MVID `02f45c8f-e575-4067-b549-4120121a744d`及正常hook启动；没有把刚交付尚未安装的CP309冒充安装版。本轮源码候选是**CP309相同字节**Word DLL `0BD8CCE72979B50FD7B49BF824AFBAAE7B60AC24AE560C70D41F8AA9161272C2`，仅测试程序集有变化。不能把两个托管方式的数值直接相减来计算插件开销。

### 原生全部替换耗时

下表为可见Word、真实进程内安装插件、每组三次中位数，单位毫秒：

| 公式/正文空格数 | VisualTeX | MathType | OMML |
| --- | ---: | ---: | ---: |
| 20 / 1120 | 32.956 | 31.355 | 33.515 |
| 40 / 2240 | 60.125 | 60.933 | 60.877 |
| 200 / 11200 首轮 | 443.754 | 304.041 | 未运行 |
| 200 / 11200 重复 | 287.287 | 292.533 | 未运行 |

首轮200的1.46倍差异确实存在，保留全部480.8/443.8/371.0ms样本；重复时VT287.3/292.7/281.4ms，不支持稳定格式退化结论。inert可见Word中200 VT286.470ms、MT295.555ms，同样未复现。真实安装插件与替换重叠的选区回调首轮最多6.378ms，重复最多1.925ms；每次替换0或1个选区事件，没有出现每个空格一次的事件风暴。未见转换、MathML解析或MTEF保真比较路径进入此操作。

### 已证实的打开/保存差异及定位边界

200公式反复对照中，原VT打开3308.875ms、保存6901.206ms；MT打开2859.197ms、保存5523.182ms。VT输入zip约3.33MB、引用EMF预览2671700字节、嵌入OLE4585472字节；MT输入zip约0.356MB、引用WMF预览327624字节、嵌入OLE718848字节（精确值见payload-sizes.json）。

诊断控制`200-VisualTeX-MTPreview.docx`保留所有200个VT原生OLE文件逐字节及形状尺寸，只将DOCX外部显示预览关系换成同序MathType公式预览。进程内插件下打开降至1163.600ms、保存4571.988ms，但替换300.238ms并没有比原VT287.287ms更快。inert对照也得到打开1206.418ms、保存4439.870ms、替换298.293ms。这支持**外部预览格式/记录/体积对打开保存成本的因果贡献**，不能进一步断言具体Word内部函数、仅字节大小、VisualTeX原生Draw实现或某个隐藏缓存是根因。控制文件只用于诊断，不是可发布的预览替换方案，更不代表视觉保真通过。

代码对应：`ThisAddIn.OnWindowSelectionChange`被事件驱动，而非替换命令自身的循环；`WordFormulaService.NormalizeTypingCaretAfterInlineFormulaCore`只有精确公式尾边界才读VT元数据，普通正文是局部±2字符探测；`ThisAddIn.OnDocumentBeforeSave`为空，不遍历修复公式。`FormulaOleObject.cpp`的Load/Save确实读写metadata/EMF/PNG，DrawEmf建立/播放EMF，但本轮没有证明那些函数在批量替换阶段被调用；不能仅凭源码给它们定责。`WordOleObjectAccessor.GetRunningObject`有有界激活等待，同样没有在本次原生ReplaceAll建立触发证据。

### 结果校验、失败与未运行范围

- 原生调用后的精确Word正文去ASCII空格、公式对象数：90/90通过；保存后独立ZIP/XML正文对照：90/90通过。
- 含OLE的72个保存样本，全部嵌入OLE载荷逐字节保持；这包括原生VT、MathType及预览控制，不依赖产品读写双方的signature。
- 18个OMML保存样本字符内容保持，但严格OMML XML/样式树等值全部报差异：Word将`a∥b`等run拆开并复制样式属性。保留`omml-save-diff-*.txt`与严格校验退出码1，不忽略或放宽该检查，不把本次性能操作算作OMML完整样式/视觉保真验收；尚未建立无操作保存对照来全面判定这些差异。
- 未物理点击Ctrl+H/“全部替换”，测试是原生Find COM API；未跑宏逐匹配移动选区、通配符、其他插件、自动保存叠加、修订跟踪、长单段或用户异常文档。20/40是统一正文的历史同组公式，200VT/MT来自另一同组原200公式，不能把组间变化全部当纯数量效应。
- 生成预览控制时首次读取回执遇到Python ZipInfo被writestr修改内存偏移，文件已生成；重新打开两个独立ZIP句柄确认校验和/压缩包有效、200载荷完全相等，原文件未覆盖，脚本改为复制ZipInfo。错误及恢复说明见fixtures/preview-control.json。其余测试前提失败和手动插件RCW错误见CP313，均未抹去。

### 可重复运行与用户侧复验

测试输入保存在`artifacts/word-bulk-replace-20261006/fixtures/`：020/040各有VisualTeX、MathType、OMML；200有VisualTeX、MathType和仅诊断用MTPreview。`manifest.json`记录原文档和输入哈希，`preview-control.json`记录控制文件身份。生成后的测试副本可直接用于用户侧比较：

1. 分别打开`200-VisualTeX.docx`与`200-MathType.docx`的副本，从正文开头执行Ctrl+H。
2. 查找一个半角空格U+0020，替换内容留空，清除查找格式条件、关闭通配符，执行全文“全部替换”。每份预期去除11200个正文空格，保留200个公式。
3. 分开记录“全部替换”完成时间、界面恢复时间和保存时间，不把打开/保存算进替换。每轮重新打开同一原始测试副本，保留冷首轮。
4. 若这些受控文件也慢，记录安装DLL版本、准确入口/选项/发生阶段；若只在原文档慢，需在获授权的原文档副本保留布局、修订/域、图片、其他插件等实际条件定位。当前没有用户原文档或宏，不能编造其根因。

自动化入口：冻结R3候选`build-cp312-bulk-replace-r4/VisualTeX.VstoFlowAcceptance.exe --mode=word-owned-bulk-text-replace`，包装脚本`Run-Probe.ps1`管理自有文档/快照，`-Selection`使用可见Selection.Find；`-Installed`正常加载现有插件，必须提供新建隔离Companion RuntimeManifest；无此开关时安装插件仅在该子进程被acceptance抑制。示例（只运行自有文件；RunLabel需使用未存在的目录）：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File artifacts/word-bulk-replace-20261006/Run-Probe.ps1 -BuildLabel cp312-bulk-replace-r4 -RunLabel local-repeat-01 -Selection -Pattern '200-*.docx'
```

该示例为插件inert对照，不冒充正常安装验收。正常插件对照必须用审阅过的`start_isolated_source_companion.ps1`创建新运行实例，再向Run-Probe传`-Installed -RuntimeManifest`；本轮旧runtime已停止，不能复用为健康服务。

### 产物与收尾

- `summary.json`：全部时序/加载身份/替换回调重叠与中位数；`verification.json`：逐保存文件的独立检查及18个严格OMML差异；所有90个样本和原始日志保留。
- `environment-final.json`：自有Companion104292已按PID/启动时间/路径/哈希关闭；全部自有Word退出；原61948/69908/97020身份保持，8处LoadBehavior保持，无新增Office/Companion/OLE服务进程。读取不到的受保护服务路径如实保留null，没有假称核对其二进制模块。
- `cp309-input-recheck.json`：CP309已有清单中只有测试Program.cs改变；产品文件保持，r4产品DLL同SHA。已安装CP304 DLL保持，HEAD仍66a8c709，冻结R3/heldR4保持，未安装/提交/修改注册表或安全设置。
- 本轮修改仅WordBulkTextReplaceAcceptance.cs、新测试分支Program.cs、OwnedAutomationWordAcceptance.cs显式子进程加载测试选项，以及脚本/报告。没有优化产品，没有重跑转换全矩阵或覆盖既有CP255/保真缺口。

下一步优先获取实际异常副本及入口，复现严重卡顿再按频率和阶段定位。若独立开展打开/保存优化，先量化实际预览导入/序列化开销并验证视觉保真，不能为了小样本耗时改掉原生公式、内容保护或比较规则。

## CP313：原生替换对照完成，继续确认进程内插件

native-r3：8个文档各3轮，共24次；20/40公式三种格式的同一历史公式组、200公式VisualTeX/MathType同一原200组，正文统一为每公式56个ASCII空格。单次 Range.Find.Execute 全部替换，1120/2240/11200空格。Word隐藏窗口、ScreenUpdating=true，Repaginate/ScreenRefresh/Save单独计时。精确正文和公式数24/24通过；独立包检查将另存verification.json。此为原生COM调用，未物理点击Ctrl+H。

替换中位数（ms）：20 MathType54.577、OMML51.623、VisualTeX61.892；40 MathType75.834、OMML73.890、VisualTeX83.409；200 MathType293.222、VisualTeX316.560。原生替换期间选区事件均0，尚未复现严重倍数退化。200保存中位数MathType6496.259ms、VisualTeX7429.231ms；这是单独保存开销，不能归入ReplaceAll。样本固定顺序、每组3次、包含冷首轮，不能当统计性能承诺。

addin-selection-r1：同24次，在可见自有Word中Selection.Find+WrapContinue，CP309字节插件由运行器手动跨进程托管。正文/数量通过，但出现跨线程选择回调与后续关闭/重开重叠、对象已被删除0x800A16C1及长回调；不能把这些回调时延当实际安装插件的性能根因。原日志全部保留。正在用独立子进程正常加载已安装CP304插件重测（安装SHA279CA2F1…645C3E，未擅自安装CP309）；真实加载路径/PID/MVID已由Word自己的trace证实，正常键盘hook也已加载。独立Companion PID104292/43136，安装服务69908未接入或停止。

候选仅测试更改：build-cp312-bulk-replace-r2/r3/r4均保持Word SHA0BD8CCE72979B50FD7B49BF824AFBAAE7B60AC24AE560C70D41F8AA9161272C2、冻结R3和heldR4。首轮编译缺using失败后仅修测试导入；native-r1/r2因服务StartTime不可读取而未启动Word，改只读CIM快照后native-r3成功。这些前提失败不是产品失败。没有改产品、注册表或安装文件。

## CP312：建立独立问题和对照

用户报告含 VisualTeX 公式的文档批量替换慢于 MathType/OMML。先按 Word 原生全部替换 ASCII 空格为空建立对照；用户实际入口待确认。CP264/265 是光标事件和少量插入，不是本问题证据，不能用其未复现数量增长来否定这次报告。

当前 HEAD 66a8c70994cbb80316c65283b0e7034b4931b324；产品工作树继承 CP309，冻结 R3/held R4 保持。只新增诊断测试和证据，不修改产品或性能策略。拟用既有真实 OLE/OMML 文档中的原生公式，统一正文段落与空格数量，测20/40公式三种类型，并用原200公式 VisualTeX/MathType 对照扩大数量。不把预览大小、公式结构或历史版本差异冒充插件回调根因。

测试分离打开、单次原生 Find.Execute(wdReplaceAll)、消息回调、Repaginate、刷新和保存；记录替换前后精确正文、公式数量、事件计数。先测试不手动加载插件的自有 /x /Automation Word（安装插件由子进程 acceptance 环境保持 inert），必要时再独立 Companion + 手动插件对比。手动插件为跨进程回调，不能等同安装版输入时延。所有文档均为自有复制品；不连接用户 Office、不改注册表/模板/设置、不安装、不提交。

证据目录：artifacts/word-bulk-replace-20261006。当前尚无性能结论，待完成真实 Word 运行和包内容检查。
