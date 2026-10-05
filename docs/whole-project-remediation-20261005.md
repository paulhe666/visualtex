# 全项目复审后的修复删减（2026-10-05）

状态：本轮W01–W07源码修复、删减及限定验证已完成；不等于全项目/宿主验收通过。用户在CP275审查后明确要求“进行修复删减”。起点HEAD为`49cf397818ee02b7bb18216630846eda9ade18e6`。保留所有现有修改、旧失败和产物；本轮不提交/推送/安装/部署，不改用户Office、注册表、Defender或held R4。

按W01–W07修复：删除OCR宽泛进程清理；统一会话完成归属；修复恢复备份和严格存储语义；修复旧剪贴板资源作用域；关闭sidecar不得启动它；使CI检查表达当前结构和真实合同。先独立小用例，再适当构建和邻接回归。性能重构、单次Undo、CP267特殊探针不扩大。

证据目录：`artifacts/whole-project-remediation-20261005/`。每阶段更新此报告和历史检查点；未完成/未运行范围不算通过。

## CP276 第一阶段

- W01：删掉`ocr_install.rs`按ExecutablePath/CommandLine子串枚举全机并taskkill的函数和安装前best-effort包装，以及lib.rs退出/重置/存储切换/删除重试/安装的调用。自有worker和记录进程回收保留。删除只验证该旧包装吞错的测试。OCR运行器r1失败揭示既有重复终止：`InstallControl.cancel`和拥有Child的runner同时taskkill；已改为cancel只递增代数，runner的100ms循环统一回收。运行器自建孙进程补`WindowStyle Hidden`，没有操作用户进程。r2待结果。
- W02：Windows hook以每次激活的对象身份归属GET和PATCH，关闭/重开同ID也隔离；消费端按激活重置会话引用。删除与12秒HTTP期限重复的15秒watchdog。实际hook+替身React/HTTP的6个时序用例通过，覆盖切换、关闭、同ID重开、GET乱序、失败后的串行保存及卸载；非WebView/Office验收。
- W03/W04：两平台safeStorage改为明确的失败偏好覆盖层，成功后磁盘为准；strict无持久后端或失败直接报错，不把内存当落盘。恢复仅捕获JSON解析错误，备份未消费时不覆盖新内容；重复清理状态幂等，写失败不删除active。新增存储合同、扩展恢复回归均通过。符号库测试改为显式内存Storage依赖；旧Windows测试未等待async导出，r1在Promise.svg失败，现仅补await/assert.rejects而不改输出预期，r2待结果。
- W05：旧ClipboardTransaction构造失败也释放已取得资源；仅剪贴板序列仍属于本操作才恢复旧内容，资源释放在finally；恢复失败记录Trace。Standalone路径保持。frozen-R3候选`build-cp276-ownership-r1`完整C#构建成功；资源生命周期故障用例还在补。
- W06：sidecar进程状态收束为未启动/运行/已停止；shutdown只消费持有的Child，结束后不可由Drop/请求重新启动；原有RPC搬为无启动职责的transport。关闭RPC限时两秒后只回收自有进程。Rust `cargo check --offline --locked --tests`通过；独立生命周期执行待做。
- W07：Git清单按NUL分隔流式读取；独立应用各自package/Tauri/Cargo版本一致，不再要求两平台版本同步；检查真实声明的跨平台源依赖，不扫描文档文字/证据提及。1457行非空的Windows架构字符串金样缩为项目/部署合同（10项目、11共享源码、平台/框架及离线比较器隔离）；原始脚本保存在证据中。当前两入口均通过，仅代表结构合同；没有删除/放宽公式内容或样式比较器、C#保真测试及原生验收。私有方法、CSS和实现排布不再伪装成架构或行为验收。

截至本阶段Windows类型检查通过，无新Office启动、安装部署或提交；旧CP255及六向/视觉缺口不变。原始r1失败日志保留，具体范围见各日志，不能因结构入口绿色算全项目通过。

## CP277 第二阶段：限定执行验证

- OCR r2运行器通过（`ocr-ownership-r2.log`），覆盖自有子进程/孙进程取消、错误、step/idle deadline及日志/文件活动。不进行真正安装，不枚举/终止用户进程。r1双重taskkill失败原件保留。
- Windows符号库r3通过（`symbol-registry-r3.log`）。r1为Promise当SVG，r2为旧拒绝文案与当前MathJax不符；实际入口convert先抛`Undefined control sequence \selfdefb`，所以拒绝断言从旧精确文案改为该精确命令错误，没有改变SVG/MathML内容、符号角色/间距等输出预期。macOS注册表完整回归尚未执行。
- 剪贴板真实嵌套类+替身原生接口11/11通过（`clipboard-r1.log`及`clipboard-r1/Program.cs`）。补充修正：copyObject已改剪贴板后抛错也要在finally记录序列，未改变剪贴板则不重复恢复。覆盖previous/source采集、copy/read/validate失败、初始化失败、重复Dispose、替换、外部更新及恢复失败；不调用系统剪贴板，不算Word/OLE实机验证。该新增finally在C#候选r1之后，最终候选需重建。
- 新增隔离sidecar运行器：从真实源文件提取完整client，仅省略无关Tauri路径定位；唯一随机管道、仅运行器自建进程。正在验证未启动/已退出/shutdown/Drop不复活，以及超时只回收自有进程。结果未出前不算通过。
- CI实际workflow仍引用已删除的Office.js命令`verify:office-manifest`/`build:office:windows-ole`，已替换为现存`build:office:windows-native`及新`test:runtime-ownership`。没有运行远程CI/部署。
- 恢复失败UI显示已捕获的实际原因，便于辨认未消费备份，未增设新的备份管理状态。held R4未改。本阶段代码仍未提交。

### CP277 验证纠正（以后续证据为准）

新增cancel_and_wait限时验证使OCR r3失败，进而发现旧r2“passed”不能证明实际终止：沙箱taskkill对单独自建的3秒PowerShell进程返回`ERROR: Access denied`/exit1；旧运行器忽略终止错误并等自然退出，timeout测试的350ms期限实际耗时约5.17秒。原r1失败不能直接归因双重taskkill；双入口争抢是代码静态存在的竞态，但本机失败现有证据首先指向沙箱权限。保留上述旧记录为过程证据，撤回OCR进程回收已验证结论，直至严格用例通过。

已加cancel_and_wait并复用reset原先15秒等待逻辑，shutdown等回收完成再处理残留；新timeout用例要求5秒子进程在4秒内已回收，取消要求5秒内释放安装lease。运行器r4正以授权的工作区外执行权限运行，仍仅自建隔离子进程，无真实安装、用户进程/Office或安全设置操作。r3日志以及独立taskkill Access denied诊断保留。

邻接`ocr_install_regression.ts`首次执行在旧配置断言失败：OCR资源现由`tauri.ocr-bundle.conf.json`按构建脚本合并，原测试仍读基础tauri.conf。改为读取真实OCR配置；同时删除两个要求存在已删除宽泛全机清理函数的字符串断言，由实际owned runner替代。其他前端/协议/安装边界断言保持，结果待执行。不会把旧测试绿色当真实进程回收成功。

## CP278 最终状态（本节优先于前面的阶段中间结果）

起止HEAD仍为`49cf397818ee02b7bb18216630846eda9ade18e6`，未另提交/推送/部署。28个本轮源码/测试/CI文件清单、SHA-256及精确差异见`artifacts/whole-project-remediation-20261005/remediation-state.json`和`remediation-tracked.patch`。已跟踪文件增加521行、删除1955行，新独立测试375行，合计净减1059行；不含文档、产物或其他人的现场修改。没有为追求行数删除内容保护、独立保真比较或原生验收。

### 最终根因与修改范围

| 项目 | 最终修改和责任边界 | 源码/入口 |
|---|---|---|
| W01 | 删除全机路径子串清理及其调用/吞错包装；取消入口只发信号，运行器统一回收自己持有的进程。reset复用原15秒等待、shutdown等待5秒；失败保留进程记录并返回taskkill实际诊断，不假称已终止，也不继续无期限等待自然退出。正常错误/取消/超时的重复终止收束到一处。 | `src-tauri/src/ocr_install.rs`、`lib.rs`、`bin/ocr_install_acceptance.rs` |
| W02 | 每次激活的唯一对象统一归属GET/PATCH及提交/隐藏转换的后续UI操作；切换、关闭、同ID重开使旧完成失去UI归属。已发出的旧请求仍可完成其原Session持久化。删除重复15秒watchdog及未读取的activeSessionIdentityRef。 | `src/office/dialog/useOfficeSession.ts`、`OfficeDialogApp.tsx` |
| W03 | 备份未消费不覆盖；已清理状态重复调用幂等；active写失败后可复用逐字一致的备份重试；不同新内容与旧备份均保留并明确报错。JSON解析失败与存储失败分离。 | `src/runtime/editorCrashRecovery.ts`、`main.tsx`、`VisualTexErrorBoundary.tsx` |
| W04 | 成功存储以持久后端为准，普通偏好仅记录失败覆盖；strict没有持久后端/写失败即抛错，不伪称内存已落盘。两平台同规则。 | Windows/macOS `src/runtime/safeStorage.ts` |
| W05 | 旧ClipboardTransaction构造失败释放资源；copy后抛错也记录序列；Dispose只恢复仍属于本次复制/替换的剪贴板，并保证资源释放。保留Standalone已有原生重入边界。 | `src-windows/VisualTeX.WindowsOffice.VstoShared/MathTypeOleStorage.cs` |
| W06 | 进程状态只有NotStarted/Running/Stopped；shutdown/Drop只消费已有Child，不走ensure_started；2秒后只终止自己持有的进程。进程状态查询失败直接报错，不丢掉旧Child并再开一个。 | `src-tauri/src/office/windows_pipe.rs` |
| W07 | Git清单流式读取；每平台自身版本一致；核对实际源码依赖。架构脚本只保留项目/编译/部署合同，删私有实现排布与CSS字符串金样；workflow改用当前native入口，接入行为/资源所有权测试。旧OCR配置断言读实际合并配置，删除要求宽泛清理的断言和一条只检查注释原文的断言。 | `tools/verify_repository_structure.mjs`、`scripts/windows_office_architecture_smoke.mjs`、`ocr_install_regression.ts`、`package.json`、`.github/workflows/windows.yml` |

上述表中未写平台前缀的路径均位于`apps/windows/`。新测试fixture和入口的完整文件名见28文件manifest。存储/符号测试显式注入Storage，不再借产品strict的旧错误回退建立测试前提。

### 通过、失败、未运行与阻塞

| 验证 | 最终状态 | 证据/限制 |
|---|---|---|
| 实际hook + 替身React/HTTP | 6/6通过，旧回调isActive立即变false | `frontend-final-r2.log`；切换/关闭/同ID重开、GET乱序、队列失败恢复、卸载；非WebView |
| 恢复及两平台存储合同 | 通过 | 同日志；重复启动、不同新内容保留、落盘失败与同备份重试；macOS只执行实际TS模块合同 |
| Windows符号库完整脚本 | 通过 | `symbol-registry-r3.log`；await与精确错误诊断修复，不改公式输出预期 |
| ClipboardTransaction资源故障 | 11/11通过 | `clipboard-r1.log`；实际类原文编译、原生API替身；没有读取或修改系统剪贴板 |
| sidecar生命周期 | 4/4通过，另1个ignored为专供自建子进程的sleep fixture | `pipe-r2.log`；随机不可用管道、自建进程；无真实Office sidecar握手/优雅关闭验收 |
| OCR运行器严格取消/超时 | 最终通过 | `ocr-ownership-final.log`与`ocr-owned-final/visualtex-ocr-install-acceptance-103008-1791176722130/logs/ocr-install.log`；超时741ms、idle753ms、取消582ms，未等5秒/30秒自然退出；授权运行器仅操作自有子进程 |
| OCR前端/协议脚本 | 通过 | `ocr-frontend-r3.log`；包含行为断言及仍保留的源码合同，不能代替进程运行 |
| 默认Windows TS、Office TS | 通过 | `frontend-final-r2.log`；最终会话调用端修改后重跑 |
| 最终Rust源码含tests | 编译检查通过 | `rust-check-final.log`；64/35 warnings，未声称完整Rust libtest执行 |
| 桌面/Office前端打包 | 通过，Office最后修改后重打 | `frontend-bundles-r2.log`、`frontend-office-final.log`；独立outDir，大chunk警告保留，性能冻结未拆包 |
| 仓库结构/Office构建合同 | 通过 | `structure-final.log`；10项目、11共享源码；远程CI未运行 |
| C#候选 | frozen-R3 x64完整构建通过 | `artifacts/omml-edit-audit-20261001/build-cp277-ownership-final-r1/build.log`及build-state；不是默认脏树 |
| Office真实链、物理按键、截图、六向公式/视觉矩阵 | 本轮未运行 | 没有启动Word/PPT/Companion/WebView，不据mock或离线构建宣称实机通过 |
| macOS完整类型/界面/符号回归、x86 C#候选、C#完整公式汇总、安装包/部署 | 本轮未运行 | 本次未建立这些测试环境；不扩大通过范围 |

保留的失败及解释：symbol-r1异步测试未await、r2错误文案过时；repository-r1旧跨平台版本相等要求；architecture-r1新脚本误写入口文件名；ocr-frontend-r1旧基础配置断言、r2只匹配旧注释；OCR r1/r3及独立taskkill沙箱Access denied；frontend-bundles首轮esbuild读取父目录被沙箱拒绝。独立taskkill诊断显示Access denied，r2旧OCR passed为弱断言，只代表等到退出，**明确不计作回收通过**。最终授权重试没有改Defender、注册表或运行权限设置，仅对本次自有运行器/构建提高工具执行权限。

### 构建身份与遗留项

最终候选目录：`artifacts/omml-edit-audit-20261001/build-cp277-ownership-final-r1`。

- Word DLL SHA-256：`96A13F88B345FCF19FADC9C095C71434A87C72977B767B134F682F9012B4781D`
- PowerPoint DLL SHA-256：`3531DA2E5F1F70B05133773A3C7140882DB257AACC256E86D3DEE24660B1A6A9`
- held R4源：`6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`，未修改。
- 冻结R3源：`FCE1BB7534EC0C9BF518C06818FA7EDA6816EBA5A0266D203CD84B8A84EBD045`，Build-Candidate/FreezeR3继续生效。

本轮可交付的是修复源码和限定构建/测试证据，**仍不适合直接标记全项目发布验收通过**。CP255宽重音/旧固定宽度失败，F02/F09、特殊身份、任意/负间距、上下标和六向视觉缺口不变；PowerPoint既有限定样式/操作范围不扩大。全局弹层性能重构、一次Undo、CP267特殊探针和已收尾卡顿调查没有重开。

下一步若进行验收：使用上列同版本候选，在自有Office实例补会话切换后迟到请求/提交及真实OLE剪贴板生命周期，再按20条代表公式记录失败/未运行项；macOS用自身依赖建立完整验证；真正发布前另行关闭上述既有保真缺口。没有自动安装、用户Office操作、新提交或推送。
