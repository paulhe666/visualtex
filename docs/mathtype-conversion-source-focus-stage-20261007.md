# MathType 出向转换与扫描前台窗口问题（2026-10-07）

## CP318：接手与复现准备

用户确认输入仍为 latex_100_inline_100_display.docx；扫描时切到前台的是浏览器、文件管理器等无关应用。归档源为 artifacts/acceptance-feedback-20261005/source-latex-100-inline-100-display.docx。错误为 “The MathType source changed while its conversion was being prepared.”，两个目标均受影响。

当前 HEAD 2476775786badb732e68a3e6a806bb85970b2ac0，比 CP317 新增中文字体前端提交；保留其修改及全部未提交文件。候选继续 Build-Candidate.ps1 + FreezeR3.targets；不以默认工作树构建替代。CP317 安装包与现场当前加载版本需分别记录，不能推定同版本。

已读共享 Capture/Apply/MathType 适配器：报错来自 FormatConversionCore 的源 MathML、display、编号、MTEF 语义字号、Word 呈现字号五项严格比较。尚未证明哪项误报。源读取正常走 WordOpenXML，不使用 MathType 原生编辑器激活。进度窗口单独 STA，当前是普通可激活 Form，异步关闭；是否引起无关应用前台切换仍待实测。

新增独立自有 Word 诊断入口 word-owned-mathtype-outgoing，保存五项前后证据，运行实际命令；尚未构建或运行，无产品修复结论。下一步旧候选复现、新鲜全文重绘链验证、焦点生命周期测量，然后最小共享修复与正反例回归。证据统一 artifacts/mathtype-outgoing-focus-20261007。仅自有进程/副本文档，不安装、不注册、不提交，不改比较器或安全设置。

## CP319：焦点缺陷已复现；转换触发条件仍在定位

现场已安装 Word DLL SHA256 为 0BD8CCE72979B50FD7B49BF824AFBAAE7B60AC24AE560C70D41F8AA9161272C2，与 CP315 内嵌一致。原 Word PID10496 及 Companion56088/NativeRender84392/Worker93292 受保护。隔离 Companion108056/43136 来自当前磁盘 Release 构建，独立 app-data、正常 TLS；新前端源码提交不能冒充该旧 Release 的载荷。所有测试用显式 /x /Automation 自有 Word，未接入原 PID10496。

基线候选 build-cp318-outgoing-baseline-r2（Word C5E11B23270C766BA8890C7AFB3592E407E907373415D36967E2F96671676CCE，冻结R3）对旧已保存200 MathType副本：source-audit 五项200/200一致，转换200/200对象，无残留MathType，命令约160.78s。只证明对象转换/计数，未做全符号保真声明。baseline-fresh200-omml 使用仅加错误字段日志的候选，从原200LaTeX副本重绘后不保存、不预扫描，立即执行实际出向命令：重绘70.05s，转换158.88s，无源改变报错；尚未覆盖正常确认窗口/可见Word。新鲜200 VisualTeX方向继续运行。用户确认重绘默认选项。

焦点基线 focus-baseline 的两个自有窗口试验均失败：初始宿主→进度→宿主；初始其他应用替身→进度→其他替身。根因是独立STA普通Form激活改写前台窗口历史，关闭时恢复之前的窗口。共享 WordOperationProgress 改为 ShowWithoutActivation + WS_EX_NOACTIVATE，无任务栏入口，仍保持Word所有权、独立STA刷新；不调用强行恢复旧前台窗口。修复后首两次测试未建立“自有窗口获准成为前台”前提，原日志保留，不能记为通过或产品失败。增加记录实际前台的后台生命周期对照，待运行；原前台两场景也需如实注明前提。

新测试文件 WordMathTypeOutgoingAcceptance.cs / WordProgressFocusAcceptance.cs 和 Program.cs 隔离入口；FormatConversionCore 仅新增五项失败诊断日志，未放宽或更改源校验。初次测试编译错误（不存在的计数帮助方法）、沙箱CIM拒绝、并行测试被既有互斥锁拒绝均保留，随后修正测试入口/运行方式。重要缺口：现场转换错误尚未复现，不能称已根治；正常确认、选区范围、真实改变源的拒绝保护、保存重开和焦点修复验证待继续。

## CP320：真实安装入口对照、焦点修复验证

- focus-fixed-background 两项自有前台前提实际均成立，2/2保持原前台，没有中途进度激活；基线2/2失败。候选Word SHA B45C279405884DC79FE8430B0F3AB2305B7381A9CD1F605E050243D8DB44D30B，冻结R3/heldR4不变。尚需实际Word的进度刷新/确认界面回归。
- `fresh200-visualtex` 名称与实际方向不符：脚本漏传TargetMode环境变量，日志证明又运行了OMML；仅算OMML对照。修正脚本后每次保存request.json，原结果不覆盖。`normal200-visualtex-r2` 才是真实MT→VT：可见Word、默认重绘选项、正常两次确认，200写入成功，MT残留0。此测试仅数量/完成，没有全符号保真或像素验收。
- `WordInstalledOutgoingAcceptance.cs` 新入口以COMAddIns.Object/IDispatch调用已安装插件，只有自有Word子进程取消acceptance suppression；不改Connect/LoadBehavior/安全设置。installed200-omml-r2是实际0BD8CCE7…1272C2/MVID dbe7e9f1-32dc-4a84-ae58-105ca6bb5445运行，完成默认200重绘→OMML，未复现报错。源Downloads磁盘内容只读SHA仍5726E650…C91，归档没有过期。
- 首次跨进程自定义接口强转E_NOINTERFACE，改为标准IDispatch；选区首轮读取正在追加的trace时FileShare模式不正确，修正测试日志读取为ReadWrite共享。错误退出后自有Word42104残留，已按PID/创建时间/路径/seed命令行核实清理，失败日志保留。时间Kind解析首次不匹配时没有关闭任何进程。用户原Word在调查期间自行关闭/新建，逐轮快照如实记录，不能宣称最初PID始终存在。
- 用户已再次提供同错误截图，正在准备失败MathType副本路径。转换错误仍未定位/修复，不得据对照通过宣布现场无问题。全选入口继续测试；收到原生失败副本后优先对照，不再扩大无关案例。

## CP321：失败现场已按用户要求另存，源文件未覆盖

用户明确授权当前文档另存副本。唯一当前Word35724的指定窗口Document路径核实后，仅成功执行一次SaveAs2到新的user-failure-mathtype-20261007.docx；原Downloads LaTeX前后SHA均5726E650…C91。新副本及独立分析副本SHA均E8B21FCF9D092189B22AB023920631C95C70C25E091C0A6747EAB5DF34218D96；200OLE、0OMML、兼容模式14、未修订。用户当前窗口指向新副本，不再接入此进程。证据user-failure-before-save.json / user-failure-after-save.json。脚本早期类型/代理/COM参数失败没有写文件；成功另存后的hash共享读取错误通过只读共享校验解决，没有重复保存。

actual-user-omml-diagnostic 在自有Word读保存后的现场副本，五项200/200一致，已进入OMML写入。保存/重开可能改变现场条件，不能据此关闭问题。安装插件全选200→OMML也通过，完整范围见completed-controls.json；关闭自有文档后的copy-snapshot“对象已删除”日志保留，未归因。焦点修改真实Word的STA阻塞进度刷新及正常确认20重绘/读回检查均通过，证据redraw-progress-20261006/cp320-word-focus-responsive及cp320-word-focus-confirm20。下一步独立普通交互Word与/x /Automation差异对照，测试启动器新增显式Interactive开关，仍按PID/seed隔离；不改产品校验。

## CP322：转换后测试代理删除与真实 Ribbon 现场前提

- `actual-user-direct-visualtex` 正常确认点 Yes 后，产品 trace 明确完成 MathType→VisualTeX 200/200；验收程序随后在 `CountMathTypeOleShapes` 读取已删除的 Document COM 代理时报 `0x800A16C1`。因此这次运行是“产品转换完成、测试后置计数失败”，既不是成功验收，也不是用户的“源改变”复现。输出包只读检查为200 OLE/0 OMML，SHA `4BCC90906FBA7359BF9BC58279FCF77E2AE6D3B2AACE4A88015C50EC448CF0C7`。
- 用户指出真实步骤必须直接点击当前文档的 `MathType → OMML`，等待扫描结束，再在全文确认框点击“是”。已回到仓库现有 `Ui-MathTypeCycle.ps1` 真实 Ribbon 路径，增加全文确认等待/点击和终态对话框采集；转换范围通过打开真实 split menu 只读核对按下状态，不写注册表。该脚本尚未执行转换。
- 20:15实际进程/窗口核对只有 WINWORD PID35724，启动时间 `2026-10-07T11:39:49.2663758Z`，`MainWindowHandle=0` 且无任何顶层窗口；既没有当前可点击的 Word 窗口，也无法用 `Connect-RunningWord` 绑定原生文档窗口。没有启动新Word替代现场，因为保存/重开可能消除触发状态；没有操作原Downloads文档。
- 此阶段新增的 live-rebind 旁路在绑定前提处失败，未接触文档；收到用户纠正后已撤回其源码入口/文件，保留失败证据。当前仍未知 MathML、display、编号、MTEF语义字号、Word呈现字号哪项发生变化，未放宽比较器、未删除保护边界、未安装/注册/提交。

### CP322 续：在用户一直打开的真实窗口完成了一次真实 Ribbon 转换

此前“PID35724无窗口”是沙箱桌面枚举的错误结论；切换到现有 `OpenInputDesktop` 连接脚本后，精确取得 PID35724/HWND17583344 和活动 `user-failure-mathtype-20261007.docx`。操作前Saved=true、200 OLE/0 OMML、磁盘SHA E8B21FCF…D96。真实UI脚本只读核对 `默认：转换全文` pressed=true，于20:27:49点击 Ribbon `MathType → OMML`；原生确认采集到“全文200个MathType公式重新绘制为OMML”及按钮IDYES=6，于20:29:29点击“是”。之后同一Word UI出现 `撤消VisualTeX Convert Formula Format`，只读COM计数为0 OLE/200 OMML、Saved=false，没有源改变错误。未保存转换结果，磁盘SHA保持E8B21FCF…D96。

结论仅是“本次当前状态真实Ribbon成功，用户原报错未复现”，不是根因修复。此前Save As及之后较长驻留/扫描可能改变瞬态状态，但现有证据不能区分保存、时间、布局刷新或其他因素，也不能猜哪一项五字段变化。旧安装版没有新增的失败字段trace。外层UI脚本对确认框的UIA句柄发生竞态，未识别后最终超时；确认框原生正文/IDYES点击和转换后0/200独立证明产品操作已完成，脚本超时不归为产品错误。证据 `artifacts/mathtype-outgoing-focus-20261007/cp322-live-real-ribbon-omml`。
