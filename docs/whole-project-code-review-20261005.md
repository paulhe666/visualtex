# 全项目架构与代码复审（2026-10-05）

> CP278后续状态：用户授权的W01–W07修复删减已完成限定代码/构建验证，详见[修复报告](whole-project-remediation-20261005.md)最终表和CP276–278。下文保留CP275原始审查证据与当时行号；不应再把已修代码视为仍原样存在，也不能把限定验证当实际Office或全项目验收。HEAD仍49cf3978，本輪修复未提交。

状态：本轮复审完成，修复与宿主验证待后续执行。用户要求先本地提交，再审视全项目，并明确不要做防御性编程。本轮只读产品代码，先证明真实调用链和可触发问题，再提出最小根因建议；未加入校验、兜底、重试，未开展性能重构。

审查快照：`49cf397818ee02b7bb18216630846eda9ade18e6`，父提交 `67e57fef62a4abcfa9577fac8849c53817f83893`。本地提交保存48个源码、测试、脚本和文档文件，未推送；生成产物和历史现场证据保留原地，未混入提交。提交前后清单见 `artifacts/local-commit-20261005/`。旧空 index.lock 在核实时间及无活动 Git 后移到该证据目录，没有删除用户修改。

## 范围与证据规则

- 覆盖检查：Windows/macOS前端编辑与持久化、Rust会话/Companion/OCR/更新、引擎、Office/OLE读写、构建和测试入口。跨平台应用有独立部署设计，不仅凭文件重复或体积大判定错误。
- CP268及CP269–272已有根因修复纳入复审；CP255原生重音及六向/实际视觉缺口保持。单次Undo不再追查；搁置的CP267探针、已收尾的数量卡顿调查不重开。
- held R4及冻结R3构建规则保持；不启动Office、不改变用户设置、不安装部署。新执行的离线检查单独记录，旧候选证据不自动归于本次快照。
- 每条问题记录具体方法/行、触发、影响、证据强度与最小建议。源代码推导、独立复现、宿主验收明确区分；不把未运行算通过。

## 进度

已完成本地提交、模块清点和关键链复审。下面保留第一轮证据，再给出最终问题排序、架构判断和覆盖边界。本报告覆盖主要子系统入口及关键调用链，并非声称逐行审计全仓所有文件。

## 第一轮已验证问题（暂定优先级）

1. **P1 会话串台**：Windows `useOfficeSession.ts` 的 `save()`完成后无条件`setSession(next)`，切换事件虽重置队列和GET代数，却未取消已发PATCH的UI写回。用未改动模块及延迟HTTP/hook生命周期模拟，A保存发出→切B并成功GET→A完成，得到`sessionId=B`但`session.id=A`。`OfficeDialogApp.tsx`的加载effect按`session.id`将旧行重新写入编辑store。macOS同名hook已有活动会话判断，Windows缺失。最小建议：把GET/PATCH完成统一归于当前会话所有者，不再仅重置Promise队列。证据`artifacts/whole-project-review-20261005/{probe-frontend.cjs,frontend-probes.json}`；非真实WebView/Office验收，尚未实测宿主误写。
2. **P1 安全启动覆盖唯一原文备份**：`editorCrashRecovery.ts:prepareEditorCrashSafeRestart`每次覆盖同一备份键。实际模块连续调用两次：首轮保存`x+1`原文，第二轮把备份覆盖为仅剩theme的清理状态，活动键也不含公式。入口是`main.tsx`启动失败和`VisualTexErrorBoundary.tsx`的安全启动按钮；底层故障未解除、再次使用按钮即可触发。最小建议：现有待恢复备份必须保留到明确恢复/导出后，不引入新的自动清理层。原有回归只测单次，未覆盖重复恢复。证据同上。
3. **P2 存储回退合同相互矛盾**：`safeStorage.ts`写入失败吞错后仍更新memory，但下一次成功读取旧localStorage又覆盖memory；`setItemStrict`在localStorage访问抛错时反而走无错误的内存成功分支。实际模块分别复现旧值复活和严格保存无持久化却成功。用户自建符号记录/索引及崩溃备份依赖strict。建议明确区分可丢偏好和必须持久的用户内容；严格路径不假报成功，普通回退只保留一个一致语义，不追加重试/多层兜底。
4. **P2 CI入口已失效**：原样执行`tools/verify_repository_structure.mjs`退出1，`git ls-files`输出超过默认缓冲区导致ENOBUFS。独立`-z`清点两平台20963个跟踪文件/2296229字节；检查还强制要求已不在Windows package.json中的Office.js依赖。`.github/workflows/repository.yml`直接调用此入口。最小建议：按当前VSTO产品边界重写过时断言并正确读取实际Git路径清单；不为过时检查重加废弃依赖。原始日志`repository-check-original.log`，清点`repository-probes.json`。后续检查尚未据此称通过。

## 当前新执行验证

- Windows当前默认前端TypeScript检查：退出0，`windows-typescript.log`。其tsconfig明确排除Core/compat/旧Office桥，不能覆盖这些代码。
- Rust engine core库单元测试：`cargo test --offline --locked -p visualtex-engine-core --lib`，23/23通过，`engine-core-lib.log`；未运行全部集成/布局/浏览器测试。
- 前述前端探针刻意断言并记录实际错误行为，脚本退出0表示复现成功，不表示产品合同通过。全程未启动Office、浏览器或Companion。

## 最终按严重程度排列

### W01 / P1：OCR“清理”会选中不属于本应用的进程

位置：`apps/windows/src-tauri/src/ocr_install.rs:543` `cleanup_runtime_processes`，尤其548的PowerShell筛选和574的`terminate_process_tree(pid)`；入口为`lib.rs:628` shutdown、767 reset、943/963切换存储、2180删除重试、2652安装。

当前规则为指定进程名且`(ExecutablePath + CommandLine).ToLowerInvariant().Contains(runtimeRoot)`，随后`taskkill /PID /T /F`。路径只是字符串前缀时，`ocr-runtime-backup\python\python.exe`也命中；独立项目的python若仅用命令行读取`ocr-runtime\logs\example.txt`也命中。生产使用的筛选表达式在三个合成进程记录上执行：自有worker正确命中，另外两个非自有记录均错误命中。没有执行Get-CimInstance或taskkill，也没有声称真实进程已被误杀。

根因是把“引用路径”当“拥有进程”，并通过额外全机清理扩大副作用。这里不是要继续叠校验，而是**删除宽泛进程扫描/强杀兜底，统一回收实际启动时持有的子进程/作业；遗留记录只能依据真实所有权回收**。现有更严格`process_belongs_to_runtime`仅在批量terminate失败后使用，不能保护已经成功误杀的记录。需优先修复，影响不局限于OCR自身，可能中断其他工作并丢失未保存结果。

证据：`artifacts/whole-project-review-20261005/probe-ocr-selection.ps1`及`ocr-selection.json`；后者两项`ContractSatisfied=false`是缺陷复现。建议修复验收：保留自有worker退出，两个无关合成/自有测试子进程必须存活；不使用用户真实进程测试。

### W02 / P1：会话切换只隔离GET，没有隔离PATCH完成

位置：`apps/windows/src/office/dialog/useOfficeSession.ts:102–123`；消费端`OfficeDialogApp.tsx:699–727`；真实切换事件由`src-tauri/src/office/server.rs:524`注入复用的编辑器。

触发和复现见第一轮第1项。切换重置Promise队列不是撤销旧请求；GET使用`loadRunIdRef`，PATCH完成却无条件写全局会话。消费者会用旧A的行重建当前编辑store。自动保存有`loadedSessionIdRef !== sessionId`检查，只能暂停部分保存，不能撤销已经发生的UI回退；显式`saveCurrentSession`仍使用当前hook的目标ID，因此存在把旧内容提交到新目标的调用链风险。**已动态证明状态串台；宿主误写尚未在Office实测**，两者不混称。

最小修复：一个活动会话/代数所有者统管GET和PATCH结果，旧请求可完成自身持久化，但不能修改新会话UI；复用macOS已有同类所有权原则，不再增加独立队列或补第二层watchdog。补实际hook中延迟完成、切换、关闭及再次打开的时序用例。payload字段校验测试与该问题无关。

### W03 / P1：第二次安全启动抹掉第一次备份中的原文

位置：`apps/windows/src/runtime/editorCrashRecovery.ts:22–55`，主因27行无条件覆盖固定备份键；入口`main.tsx:58`及`runtime/VisualTexErrorBoundary.tsx:99`。

一次恢复后原文仅在backup key中；启动故障若未解除，第二次安全启动把已清理的active envelope覆盖进同一个backup。实际模块探针前后记录证明原文`x+1`在active/backup两处都消失。这影响尚未另存文件且无其他历史快照的恢复场景，不能声称会删除已经另存的文件。

最小修复：把待恢复备份视为尚未消费的用户数据，保留到明确恢复/导出/丢弃；正常重试不重新覆盖。不要再设计“多级清理直到能启动”。现有`editor_crash_recovery_regression.mts`本轮原样通过，但每个用例先清空backup且只执行一次恢复，未验证重复启动；这是明确测试盲点，不能据其绿色宣称恢复安全。还需把解析失败与持久化失败的处理分开，避免同一catch把存储写入故障当成坏JSON后删active key。

### W04 / P2：内存回退与“严格持久化”给出错误成功语义

位置：Windows/macOS `src/runtime/safeStorage.ts:13–60`，两边关键逻辑相同；使用处为`editorPersistenceStorage.ts`及`customSymbolRegistry.ts:858`的记录/索引发布，还有W03备份。

实际Windows模块验证：localStorage写抛QuotaExceeded、读仍成功时，新memory值立即被旧磁盘值覆盖；访问localStorage抛SecurityError时，`setItemStrict`直接内存写后返回成功。自建符号因此可显示保存成功，重启后消失；普通偏好回退也不满足“本次会话保留新值”的自述合同。macOS本轮为同逻辑静态核对，未在WebKit复现。

最小修复：区分明确的可丢偏好接口和必须落盘的用户内容接口；严格接口无后端就返回失败，不伪造持久化成功；普通接口选择一致的失败后读写语义。不要每个调用点再加一组catch/重试，更不要把该基础层问题补进每个符号注册逻辑。

### W05 / P2：旧MathType剪贴板回退在构造失败时没有释放路径

位置：`apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/MathTypeOleStorage.cs:42–59,93–112`；生产入口Word `MathTypeOleStorage.cs:33–43` → `WordMathTypeHostAdapter.cs:178`等。

当旧Word容器无法通过WordOpenXML读取载荷，真实代码会进入`new ClipboardTransaction`。构造器先OleInitialize、保存旧IDataObject、Range.Copy，再读取/验证载荷。如果Copy后的读取或验证抛异常，构造没有返回，调用者的using不会执行Dispose；此时剪贴板已改变，显式COM引用/OLE初始化也没有配对释放。成功Dispose又无条件恢复旧剪贴板，可能覆盖操作间另一应用的新复制；恢复失败被吞掉。不是死测试helper，静态调用链已经贯通，但本轮没有触碰真实剪贴板/Word做故障注入。

最小修复：将构造阶段的资源获取放入可确定释放的操作作用域，按本次操作的剪贴板所有权恢复；不要在各转换调用者补catch。PPT新建MathType使用的`StandaloneClipboardTransaction`有不同的原生重入约束及当前所有者检查，不能把旧previous-IDataObject恢复策略照搬过去，也不能把此问题归因为已知Companion访问冲突。

### W06 / P2：关闭sidecar反而走“确保启动”

位置：`apps/windows/src-tauri/src/office/windows_pipe.rs:107–136,155–197,223–225`；调用链`lib.rs:3929 shutdown_runtime` → backend.shutdown → pipe.shutdown。

`shutdown()`先`request_method("shutdown")`，任何request都会`ensure_started()`。因此sidecar从未用过或已经退出时，关闭应用仍启动一个新sidecar来关闭；成功shutdown取走child后，后续Drop再次调用shutdown也可重复启动。这是源代码直接成立的生命周期错误；本轮未启动真实sidecar验证进程现象。连接有5秒等待，但管道`read_line`没有对应读期限，已有sidecar卡住时退出调用也可能等待不结束，具体时长需隔离实验。

最小修复：把启动请求与向现有自有进程发送请求分开；关闭只消费已经持有的child并保持幂等，Drop只释放现有资源。不要加“关闭失败再启动/重试”的新层。需验证未启动、已退出、正常运行三种退出场景，不把status探测成功当退出测试。

### W07 / P2：CI断言固定内部源码形状，已脱离实际项目

位置：`tools/verify_repository_structure.mjs:90–104`、`.github/workflows/repository.yml`；`apps/windows/scripts/windows_office_architecture_smoke.mjs:589`、`.github/workflows/windows.yml`。

原始仓库检查因Git清单超默认maxBuffer退出1，连积累的失败都未输出；进一步静态核对发现它强制已删除的Office.js依赖。另一个工作流的architecture smoke本轮原样退出1：断言主文件必须包含已无此定义的`InsertBulkOleDocumentTwoPhase`。当前实际入口为`InsertBulkDocumentHostCore`/`InsertMathTypeBulkSkeleton`，该失败不能证明插入功能坏，也不能被忽略成“测试都通过”。这是不可用的CI门槛与源码字符串耦合；不要为了恢复绿色把旧方法/依赖加回产品。

最小修复：仓库检查只表达当前部署结构和真实文件身份；架构约束测试验证公开职责/可观测行为，删除对私有方法名称、文件拆分和无关CSS字符串的偶然依赖。检查器自身要能处理仓库当前文件清单，不能通过筛掉产品失败来解决。

## 值得保留与该收缩的架构

应保留：共享token/style、原生字体/编码/槽位规则、MathML/OMML颜色作用域、内容和格式的独立读回、CP269/270有源身份和Undo历史证明的PPT失败恢复。它们保护实际文档内容，不因用户“不做防御性编程”就删除。CP272实际MTEF入口的终止/树界限与一次遍历也保留。离线语义比较器不在生产调用链中，不计为运行时成本。

Windows/macOS是独立应用，不能为了消除目录重复强行合并。但纯存储合同和会话所有权语义应有共同测试/同步依据，W02就是修正在macOS而Windows继续漂移。`office/api/sessionClient.ts`只是shared模块转导出，不是两份协议实现；不应为“重复文件”删除正确适配层。

前端弹层存在两套定位职责：`EditorWorkspace.tsx:368`局部portal/锚点与`runtime/floatingLayerAutoAvoidance.ts:167–213`全局修补器。后者监听整个document的childList变化，每次观察回调安排当前RAF及48/160/260ms三个定时任务，未在新一轮取消前一轮settle timers；所有候选层又会读样式/矩形/祖先并最多四轮写translate再读布局。弹层是否变化与正文DOM变化没有分离。这是明确可收缩的全局补丁结构，**不是本轮测得的输入卡顿**。后续应由弹层拥有者承担锚点和避让，减少整页观察与重复定时，不再叠第五次定位。性能冻结，本轮没有重构或实测延迟。

Word/PPT大型service及OfficeDialogApp确实职责过多，但仅按行数拆文件无法消除耦合。优先消除W02/W06的两个状态权威、W01的扩大清理和W04的错误回退，再沿准备数据/宿主写入/独立读回的已有职责分开。不要先造新的manager、全局缓存、万能retry或通用“安全执行器”。

已有审查的同一次MTEF定义反复解析、局部List.Contains、PPT显式命令整PPTX保存成本继续保留为有调用依据的性能候选，不在本轮擅自优化。CP270移除的被动选区无消费者缓存和隐式身份写入没有被恢复。

## 仍需实验的结论（不是已证明的新故障）

- Companion `lifecycle.rs:333 start`到`server.rs:3140`登记handle间没有启动任务所有权；`stop()`只关闭已有handle。快速start/stop和双start可能出现晚启动或状态被旧任务覆盖。静态时序风险已识别，需要独立端口/自有运行时验证；不把它用于解释CP217旧0xc0000005。
- 原生OLE `FormulaOleObject.cpp`的SetFormulaJson/InitializeOrUpdate在保存失败后恢复内存，但多stream写入或容器回调后的磁盘状态需要IStorage故障注入确认。当前只做了Load/Save/Update/extent与已有smoke用例的源审查，未宣布发现真实文档损坏。
- 自定义符号库采用记录先写、最后发布索引和上一代备份的设计有价值；跨WebView并发写及发布后GC是否冲突没有做双窗口实验。不要在未证明前添加第三层副本。
- macOS本机没有独立node_modules，类型检查退出2主要为缺React/Tauri等模块，前提未建立；未安装依赖、未运行macOS宿主，不把级联类型错误当产品缺陷数量。

## 覆盖和最终验证边界

| 子系统 | 实際检查 | 本轮结果/限制 |
| --- | --- | --- |
| Windows前端 | 持久化、恢复、历史、Office会话、弹层、更新URL/版本入口 | TS默认配置通过；独立缺陷探针成立；未跑浏览器/UI；Core及compat被配置排除 |
| macOS前端 | 同名存储/会话对照、独立应用边界 | 静态对照；类型检查缺依赖，不算通过或产品失败 |
| Rust engine | session单写者/revision、history、协议边界、产品接入 | core库23/23；未做完整集成/字体布局/WASM视觉；默认Windows未启用Core，不冒充默认编辑路径 |
| Companion/Rust | 会话原子写、服务启停、sidecar IPC、OCR安装/存储/取消/退出 | W01/W06；会话存储互斥和原子替换保留；没有启动服务 |
| Python OCR | 协议、模型缓存、响应错误/日志和LaTeX规范化入口 | 静态；未装/下载/运行模型 |
| Word/PPT C# | CP268–272修复调用链、原生读取/剪贴板、当前构建隔离与证据 | W05静态贯通；无新Office/新C#汇总，不扩大旧限定通过 |
| ATL OLE | Load/Save/SetFormulaJson/InitializeOrUpdate/extent及smoke断言 | 静态，IStorage故障注入未做 |
| 构建和工具 | 两个平台tsconfig/包、CI、结构/架构检查、候选脚本和冻结输入 | 两项真实检查失败；旧失败和证据保留 |

新增执行证据均在`artifacts/whole-project-review-20261005/`：`frontend-probes.json`（会话/回退/strict/重复恢复缺陷）、`ocr-selection.json`（两个非自有记录误匹配）、`repository-check-original.log`、`repository-probes.json`、`windows-office-architecture.log`、`windows-typescript.log`、`macos-typescript.log`、`existing-crash-recovery.log`（旧单次测试通过）、`engine-core-lib.log`。源码、候选DLL、HEAD和heldR4的核对清单为`review-state.json`。

重新核实旧CP272候选`build-cp271-native-structure-final-r1`：Word `41F42D5F3CE327E39D7371D0F80748B1A0A91487BA9AEF39857496EE7AA14F87`，PPT `8F9E779477E9443524934C127978505548E747E81CEE17718F5D2660FB3215AC`；heldR4 `6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`未变。Build-Candidate仍以FreezeR3替换resolver输入，冻结R3合同`FCE1BB7534EC0C9BF518C06818FA7EDA6816EBA5A0266D203CD84B8A84EBD045`；这不是当前默认树构建。本轮没有生成新的Office DLL，也没有重新跑2372项汇总；CP272的2371通过/CP255唯一失败/另外4项排除保持原范围。

## 优先处理顺序与打包结论

先W01（副作用越界）、W02/W03（编辑/恢复内容），再W04（存储真值）、W05/W06（资源生命周期），同时修W07使CI恢复表达真实合同。修复应收缩责任与错误回退，不添加逐用例条件、更多缓存或万能防护。每项先用最小独立用例，重要阶段再回到代表公式及同一候选的宿主验收。

**当前不适合对外宣称验收通过或直接打发布验收包。** 除本轮问题，CP255固定组重音/旧字节合同、F02/F09、特殊符号和间距、同一最新构建六向视觉、物理Ribbon/快捷键及部分PPT混排样式仍未关闭；单次Undo依用户指示不再阻塞。专项诊断快照需标明范围，不能拼接不同旧版本的通过项。

本轮在提交之后只新增审查文档、独立证据脚本/日志及运行生成的缓存；产品源码与已有测试预期未改，未改比较器，未提交第二次、推送、部署、注册表/Defender或用户Office状态。检查点CP274/CP275记录此结论。后续待办是修复和限定实验，而不是把“审查完成”当作“项目已验收”。
