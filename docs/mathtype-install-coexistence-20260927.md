# MathType 7.9.1 中文安装与 VisualTeX 共存审查

## 结论边界

两个问题需要分别处理，不能以静默安装成功代替普通安装验收：

1. **已定位的 VisualTeX 缺陷**：精简 runtime 遗漏 `Language/MT7chs.dll`。私有 MathType 子进程冷启动时找不到用户选择的中文资源，回退英文并改写公共 `HKCU/Software/Design Science/DSMT7/Config/AppLang`。
2. **仍在定位的安装交互异常**：下载包解压后外层退出，内层 `setup.exe` 仍等待安装对话框；不是已证实的崩溃，也尚未证实由 VisualTeX 导致。

## 语言问题的根因与最小修复

现场 A/B 使用同一个原始预览 bridge、同一份 MTEF：

- 原精简 runtime 冷启动：比较的 1,395 项公共设置/文件哈希中，只有公共 `AppLang` 的两个注册表视图发生变化，由 `0x0004,CHS` 变为 `0x0409,ENU`。
- 加入原安装包的 `MT7chs.dll`，不增加注册表补丁：冷启动渲染成功，1,395 项比较结果没有变化。
- 暖启动原 runtime 不再改写语言，进一步把变化定位到原生子进程启动。
- 在父 sidecar 增加 HKCU 重定向的实验无效，已从产品代码撤回。进程局部的 `RegOverridePredefKey` 不会自动成为子进程的注册表视图。

产品修改：

- runtime 打包保留原始载荷的 `MT7xxx.dll` 应用语言资源，中文 DLL 纳入必需文件和哈希清单。
- 不再接受遗漏中文依赖的旧缓存 runtime。
- 启动前检查已保存的语言；缺少对应资源时拒绝启动精简 runtime，不通过转换后恢复注册表来掩盖问题。
- 专项验收增加公共 MathType 设置的前后哈希对比，不记录许可证明文。

原始中文 DLL SHA-256：`3709C93CAB61487EE967AD089FB8A33E67C322AE41BCAF6ADA1865E537A4E46C`。

已完成：bridge 编译、runtime 暂存及缓存校验、源码边界检查、缺失语言包负向测试。负向测试返回明确错误且不启动原生 MathType 进程。最终候选版本的冷/暖启动复测等待交互安装结束，避免相互干扰。

## 普通安装异常的证据

原下载包：`MathType-win-zh-7.9.1.454.exe`，签名验证有效，签名方 MATHS FOR MORE SL。

SHA-256：`F32BC10DEC1892ABF63999846B6BC82B6CF3B29FC9E2A18DA1A172BAA89641E3`。

- 正常启动，无静默参数；内层命令行也没有静默/隐藏参数。
- 失败时没有 MathType runtime、独立 MathType 或 Office 进程在运行。
- 内层启动标志未要求隐藏，桌面为 `WinSta0/Default`。
- 32 位转储的调用栈停在 `DialogBoxParamW` / `NtUserWaitMessage`，对应主安装对话框资源 1200。
- 用户确认 Alt+Tab 没有安装窗口。
- 安装日志中的 `section "startup" not found` 在成功安装日志中同样存在，不能据此认定失败根因。
- 临时复制完整载荷，将 `setup.inf` 的 `no_background=1` 改为 `0` 后，无参数启动，用户确认出现“安装／移除”按钮。**此对照同时改变了路径，不能单凭这一结果认定唯一原因。** 原下载包保持不变。

临时配置不进入 VisualTeX 产品。尚未完成同路径单变量复核及两种安装顺序的干净环境验收，因此不宣称普通安装已根治。

## 现场操作异常

19:44 通过电脑操作入口指定已安装的 `Setup.exe` 路径，未显式传参，日志却记录 `-Q -RS` 并执行静默卸载。已向用户说明并停止使用该启动入口；参数如何被添加尚未确认。该事件不能归因于 VisualTeX。随后通过原始载荷的显式无参数启动恢复交互安装。后续检查已重新存在版本 7.9.1.454 的原生 MathType，公共 COM 注册的 32/64 位视图均指向 `Program Files (x86)/MathType/MathType.exe`，用户语言为 CHS。

## 后续原生 Word 插入崩溃与结构复审

用户确认复现路径：新建空白 Word 文档，点击原生 MathType 选项卡的“显示”。

Windows 记录的两次崩溃为 20:08:31 和 20:08:43：

- 进程 `WINWORD.EXE` 16.0.14334.20918。
- 故障模块 `wwlib.dll` 同版本，访问违规 `0xc0000005`，偏移 `0x2ee4e8`。
- 加载的 MathPage、MT6、中文 Word UI 资源均来自原生安装。Office 目录的 MathPage.WLL 与 MathType 安装目录中的文件哈希一致。
- 单凭故障模块名不能区分 Word、自身模板、MathType 回调或 VisualTeX 事件重入的责任。
- 调用原生 Ribbon 宏 `MTCommand_OnInsertInlineEqn` 和 `MTCommand_OnInsertDispEqn` 的空白文档自动化探针均返回且得到一个 InlineShape；未复现用户的物理点击闪退，不能视为问题已修复。
- 20:22 在单独的新建空白文档中进行实际点击跟踪，用户确认“显示”正常打开 MathType。记录中选择事件均返回。没有打开用户已有文档，没有改动全局 Word 插件设置。该次通过不解释或推翻 20:08 的两次崩溃。

源代码确认的结构性耦合：

1. `ThisAddIn.OnConnection` 自动 `AcquireSharedSession`，该入口立即预热 MathType，即使用户没有使用任何 VisualTeX 公式功能。
2. `MathTypeNativePreviewRenderer` 请求优先解析已安装 MathType；因此不是一个始终私有的渲染服务。
3. 清理代码扫描所有 MathType 进程，用 PID 基线、无主窗口和 `-mtrpc` 判定归属后终止进程。这不能证明进程由 VisualTeX 独占，更不能证明其未被原生 MathType 接管使用。
4. 父 sidecar 的注册表重定向不是子进程/COM/用户配置的完整隔离。Microsoft 明确规定 `RegOverridePredefKey` 只影响调用进程。

这些是确认存在的设计问题，不等于已证实它们导致这两次崩溃。

### 已确认的验收要求与隔离验证

推荐以**文档中的标准数据**作为两个产品之间唯一的互通边界：

`VisualTeX 编辑器 → 私有原生 MathType 渲染工作进程 + MTEF/OLE 编码 → Office 文档 ← 原生 MathType 应用`

- VisualTeX 使用自带原生 runtime 保持排版一致，但不得复用或终止原生应用的进程；进程所有权必须由创建时的句柄和独占作业证明，而非扫描同名进程推断。
- 不重定向或接管原生 MathType 的 COM 类、语言、偏好、宏和安装位置。
- 原生 MathType 操作保持原生路径；VisualTeX 编辑第三方公式应由明确动作或用户选择触发。
- 安装/卸载只管理 VisualTeX 自有组件，不修补第三方安装器。

现有 MTEF/OLE 编码和 Word 矢量后备路径可复用，但 PowerPoint 目前强依赖原生预览；不能只关掉预热或返回“不可用”就宣称迁移完成。

用户要求完全独立且保持 MathType 原生排版；最新澄清允许替换 runtime 乃至整个结构，保留 runtime 本身不是约束。runtime 的目的只是让未安装 MathType 应用的用户也能完成转换。自有排版引擎若未证明原生排版一致性，不能直接替代。需要验证整个原生子进程树的配置、COM、命名对象与生命周期隔离；不能用隐藏窗口或事后恢复配置冒充隔离。

2026-09-27 临时验证进展：

- 已编译两个不加载 MathType/Office 的临时探针包，比较 full-trust MSIX 与 AppContainer 的注册表、命名互斥量、COM 工厂及子进程边界。探针只使用随机 GUID 自有测试对象。
- 用户仅批准将测试证书加入 **CurrentUser/TrustedPeople**。实际安装返回 `0x80073CF0 / 0x800B0109`，尚未启动探针，因此没有隔离结果。
- 已移除本次创建的 CurrentUser/TrustedPeople、CurrentUser/My 测试证书和私钥。没有安装成功的测试组件，没有加入受信任根，没有修改开发者模式或安全策略。
- Microsoft 的 MSIX 文档要求测试签名证书置于 **LocalMachine/TrustedPeople**；此范围超出本次授权，未擅自执行。继续该安装验证需要另行取得明确许可。

MSIX 签名信任说明：<https://learn.microsoft.com/en-us/windows/msix/msix-troubleshooting-guide>。

### 本机信任获批后的实际隔离实验（20:45–21:01）

用户进一步批准本机 TrustedPeople 信任。本轮使用新测试证书，未加入受信任根，未修改开发者模式或公共 MathType 注册。

| 测试边界 | 现场结果 | 结论限制 |
| --- | --- | --- |
| 普通未打包对照进程 | 可读取外部哨兵互斥量和 COM 工厂 | 对照成立；IClassFactory.LockServer 的远程返回值不能用来当 PID，已更正探针日志含义 |
| Full-trust MSIX 及其子进程 | 继承包身份；HKCU 哨兵写入没有外泄；仍可打开外部互斥量 | 不是完整的单实例隔离 |
| 独立桌面 | 未打包和 full-trust 子进程都可创建；外部 COM 工厂查找返回 CLASSNOTREG | 仅证明该测试 COM/桌面边界，不证明所有 COM 注册和内核对象隔离 |
| AppContainer 及其子进程 | 外部互斥量访问被拒绝，HKCU 写入被拒绝；自注册 COM 工厂的查找也被拒绝 | 不能直接承诺现有 legacy runtime 无改动可运行；未把它作为正式方案 |
| 私有窗口站 | 未打包/full-trust/AppContainer 均返回 error 5 | 本轮不使用需要额外权限的窗口站路线 |
| Windows 10 19045 + uap16 BaseNamedObjectsIsolation=package | 组件可装；哨兵互斥量仍可访问 | 声明在本机未产生要求的效果，不能仅看安装成功 |
| 自有进程 SetTokenInformation(TokenBnoIsolation) | error 87 | 未继续采用未验证的 token 路线 |

真实 runtime 检查：

1. 将现有语言完整 runtime 与已有 bridge 原样装入临时 full-trust MSIX，在独立桌面、创建时绑定的 KILL_ON_JOB_CLOSE 作业内运行。
2. 原生预览未成功，响应中 `Success=false`、状态 `-6`；不以 bridge 进程 exit 0 冒充渲染通过。
3. 为排除自动发现，在同一独立桌面显式启动该包中的 MathType.exe `-mtrpc`，该进程立即返回 exit 0，未继续提供渲染。没有修改原生 EXE/DLL/DSC。
4. 两次分别比较 1,376 项公共 MathType/COM 配置，均无修改、增加或删除；原有 MathType/MathTypeLib/Word PID 集合均保持 63024/16988/75204。未结束这些用户进程。

进一步只读检查原生启动逻辑与系统对象：

- 本版 MathType.exe 启动函数在 VA `0x006DE580` 展开 `MainMTWin.<AppKey>`，先 FindWindow，再 CreateMutexA；在 `0x006DE638` 检查 ERROR_ALREADY_EXISTS (183)，已有对象则返回“已有实例”。没有修改这段程序。
- 当前用户会话中的 `MainMTWin.DSMT7` 互斥量确实存在（只读打开后关闭句柄，未等待或释放所有权）。
- 因而独立桌面未解决的**会话级同名单实例机制**是明确的结构冲突。不能据此宣称已经解释 20:08 的 Word 访问违规，后者仍没有充分崩溃现场证据。

本阶段结束时的下一步候选是**专用低权限身份、非交互会话中的原生渲染宿主**。随后用户明确批准临时低权限服务安装与验证，实际结果见下节；不再停留在等待授权状态。

本轮清理已执行：删除两个临时 MSIX 组件；删除本机 TrustedPeople 中新证书 `5D854D2AA06FCECF38E89D2F51144B69FE0935C0`，以及当前用户 My 中对应证书/私钥。工作区保留脚本、哈希报告、构建载荷和日志，可重新构建。没有将这些验证代码合入产品、没有重新打 NSIS。

会话命名空间依据：<https://learn.microsoft.com/en-us/windows/win32/termserv/kernel-object-namespaces>。

### 低权限服务原型与数据接口候选

用户批准后安装临时服务 `VisualTeXRuntimeProbe_2076b2747016`。使用专用虚拟服务账户，不使用 LocalSystem，不加载 Word/PowerPoint，确认：

- 实际运行于 Session 0 和非交互窗口站；HKCU 为该服务 SID 的独占配置单元，不是 `.DEFAULT` 或 LocalService。
- 新建服务账户的首次冷启动成功渲染；自有运行目录中的 MathType/MathTypeLib 均位于服务会话。
- 子进程暂停创建，加入独占的 `KILL_ON_JOB_CLOSE` 作业后才执行。结束时只清理自有作业，不按进程名、无窗口或 `-mtrpc` 推测所有权。
- 5 个专项样本渲染成功。与既有原生预览在相同 Windows GDI 栅格化条件下逐像素比较，各 160,000 像素均为零差异。WMF 字节并非完全相同，差异记录为字体创建记录；不能宣称所有公式或所有机器已证明一致。
- 1,376 项公共 MathType/COM 值的哈希无修改、增加、删除。原有交互会话 MathType/MathTypeLib/Word 的 PID 与创建时间未变。

随后新增**未切换正式 Office 入口**的候选代码：

- `VisualTeX.WindowsOffice.Contracts/NativeRenderProtocol.cs`：有界数据协议，只接受 MTEF、ID，返回 WMF 字节、尺寸和基线；没有客户端指定的服务端路径或执行命令。
- `NativeRenderClient.cs`：仅本机管道，发送数据前核对管道服务端 PID 与 SCM 已注册服务 PID；连接仅请求 Identification，不允许服务冒用调用者身份。
- `VisualTeX.NativeRenderService`：虚拟账户及独占 HKCU 检查、本机管道 ACL、拒绝远程客户端、独占首实例、串行作业、读写期限和进程作业归属。
- `VisualTeX.NativeRenderWorker`：精简的原生工作进程，不带 Office 自动化；拒绝在交互会话启动。严格私有 runtime 模式不再回退到已安装 MathType。
- `VisualTeX.NativeRender.Tests`：协议边界与可显式启用的真实服务验收。旧 WindowsOffice.Tests 总项目本身存在重复成员及缺少类型的编译问题；未为了运行本次专项测试修改其无关业务代码。

候选服务以 `VisualTeX.NativeRender` 临时安装测试：两个独立数据客户端并发提交，每个 5 个公式，全部成功；再次比较 1,376 项公共配置无变化，交互会话的三个既有进程未变化。5 个候选 WMF 与原生基线的像素差异均为零。首次 IPC 测试暴露 Windows 管道断开丢弃未读缓冲的问题，已将完成条件改为客户端确认完整响应，而非仅服务端写入返回；复测通过。没有通过放宽测试期望解决。

这些是后端可行性与接口证据，**不是完整产品修复验收**：

1. Word/PPT 的旧 `MathTypeNativePreviewRenderer`、预热、同名进程扫描清理尚未切换/删除。新候选不被现有安装版自动使用。
2. 当前机器安装有原生 MathType 和字体；精简 runtime 暂未封闭字体依赖。服务账户独立不等于 HKLM/系统字体完全不存在，不能据此宣称“从未安装 MathType”场景通过。
3. 还需干净 Windows 验证字体、运行依赖、厂商授权条件，以及两种安装顺序；未绕过许可或修改原生二进制。
4. 还需生产级安装/更新/卸载管理和 Word/PPT 真实插入再编辑验收；20:08 Word 崩溃仍未充分归因。没有新 NSIS，也没有把候选称为正式修复包。

后续结构决策：Office 只消费标准公式数据；MTEF/OLE 编码与排版后端分离。若隔离原生引擎在干净环境通过，可作为保持原生排版的后端；若不通过，需要另选有完整依赖与授权支持的引擎，不把保留旧 runtime 当作前提。正式迁移时删除旧入口与旧进程清理测试，不通过保留无效兼容方法让旧回归继续通过。

清理状态：两个临时服务均已删除，两个专属 ProgramData 运行目录均已删除，所有自有工作进程已结束；用户 MathType/Word 未关闭。Windows 仍保持以下两个**新建临时服务账户**的 profile/hive 加载，标准 Win32_UserProfile 删除返回“文件正被使用”，因此尚未删除这两个配置目录：

- `C:\Windows\ServiceProfiles\VisualTeXRuntimeProbe_2076b2747016`
- `C:\Windows\ServiceProfiles\VisualTeX.NativeRender`

保留精确 SID/路径的清理状态文件，待配置单元自然卸载或用户安排重启后再通过 Windows profile provider 清理；不强制卸载 hive、不结束系统服务、不自行重启。工作区保留可复现脚本和证据，临时服务已不再运行。

本机最初未发现 WindowsSandbox.exe、VMware/VirtualBox/Hyper-V 管理命令。用户随后明确批准“允许启用 Windows 沙盒，重启前再问我”。已使用 `Enable-WindowsOptionalFeature -Online -FeatureName Containers-DisposableClientVM -All -NoRestart` 完成启用，前态 Disabled，后态 Enabled，返回 `RestartNeeded=true`。**尚未重启，也没有关闭用户文档。** 待用户保存工作并安排/批准重启后再启动干净环境验证。结果记录：`artifacts/mathtype-install-20260927/isolation-probe/windows-sandbox-enable-result.json`。

重启后续接顺序：

1. 只读确认启动时间和 Windows 沙盒可用；复核两个临时服务不存在。
2. 使用忽略目录内保存的 `temporary-service-state.json` / `candidate-service-state.json` 及其管理脚本 Remove，验证精确 SID/路径后通过 Windows profile provider 清理两个残留服务账户配置。目录已经删除时脚本允许仅清理配置，不触碰用户账户。
3. 在 Windows 沙盒中映射仅包含候选程序、合法原始依赖与固定公式样本的只读输入目录；仅单独的结果目录可写回宿主。禁用不必要网络和剪贴板，不映射用户文档或整个工作区。不得以宿主测试冒充沙盒测试。
4. 先验证从未安装 MathType 的冷启动、字体/排版、所需授权条件，再验证安装顺序和正式 Word/PPT 行为。Windows 沙盒默认没有 Office，不能用渲染通过代替真实 Office 插入验收。
5. 验证通过后迁移正式 Office 入口和安装器、删除旧预热/进程猜测清理，再编译唯一明确的 NSIS 验收包。目前未生成该包。

最后专项检查：9 项协议/严格私有依赖边界单测通过；真实服务用例为显式启用，服务删除后的默认运行正确跳过该项。此前真实 IPC 并发用例已通过。直接在交互会话运行候选 worker 返回拒绝码 10，没有启动原生引擎。

非交互服务边界：<https://learn.microsoft.com/en-us/windows/win32/services/interactive-services>。

### 本机卸载/重装验证与入口迁移（22:17 起，更新前述阶段状态）

用户选择本机方案，并明确允许关闭 Word/PowerPoint/MathType 及丢弃未保存文档；没有要求继续使用沙盒。本轮没有重启系统。

- Word 测试窗口已消失后确认 Office 进程为空，仅关闭路径、PID、创建时间都核对一致的两个原生 MathType 后台进程。使用官方 `Setup.exe -Q -R` 卸载，exit 0，主应用和卸载注册项移除。
- 卸载后仍有 24 项系统字体注册和少量公共 COM/MathPage 残留。没有删除共享字体或用户设置。因此这是“无主应用”测试，不是“从未安装过 MathType”的完整干净机测试。
- 重新安装临时隔离服务（复用本轮自有服务 profile），主应用不存在时：10 项协议/真实服务专项测试通过；5 个样本各 160,000 像素与原生基线零差异；48 项公共配置哈希不变；结束后没有交互会话 MathType 进程。
- 随后无参数启动原始签名下载包。用户点击正常下一步后再次出现解压窗口消失；内层 PID 71752 仍在默认交互桌面运行，命令行无隐藏/静默参数，启动标志为 0。未把这一现象称为已证明的崩溃。

安装器单变量对照：

1. 相同临时路径、工作目录、EXE 哈希、无参数、Normal 启动，原始 `setup.inf` 的 `no_background=1`：PID 78964 没有可见安装窗口。
2. 保持以上条件和隔离服务状态，只把临时配置改为 `no_background=0`：PID 43552 出现背景窗口和安装对话框。没有更换安装器二进制或更改许可证检查。
3. 用户按普通交互流程完成安装，22:34 安装日志完成；主程序和 `Language/MT7chs.dll` 恢复。
4. 安装完成且安装器退出后，把临时及安装目录中受该实验影响的 `setup.inf` 都恢复为原始 SHA-256 `71D452B209775DA2E85C1C5239CA049715808BFA28F352B9012BC3B5FC0D1A67`。原始下载包从未修改。
5. 重装后再次运行 10 项隔离服务专项测试：全部通过，214 项公共配置无修改/增加，结束后无遗留交互进程。

这证明本机窗口消失与该安装器的背景显示分支直接相关，**不等于原始下载包的普通安装已修好，也不证明所有系统状态因素已排除**。临时配置实验不作为产品补丁发布。原生 Word 点击“显示”的崩溃仍需实际共存验收。

正式入口源代码迁移已开始：

- 共享 `MathTypeNativePreviewRenderer` 改成数据管道客户端，移除已安装 MathType 自动发现、启动/清理进程和共享预热。
- Word OnConnection/OnDisconnection 不再取得/释放 MathType 共享会话。
- Word 插入、批量转换、重绘和编辑不再静默回退到另一种排版；原生排版未准备完整时在修改文档前报错。
- 删除旧预热/按 `-mtrpc` 猜测进程归属的验收约定，迁移为私有渲染不改变交互进程身份的约定。没有为旧回归恢复空方法。
- .NET 8 SDK 下 x64 Word、PowerPoint 与 VstoFlowAcceptance 编译成功（仍有既存警告）；没有运行无关总回归。

**当前仍未完成**：字体依赖封闭、正式服务安装/更新/卸载集成、真实 Word/PPT 插入再编辑及安装顺序最终验收。安装版 Office 插件尚未切换，未打新 NSIS。候选 `VisualTeX.NativeRender` 此刻重新安装并运行，先前“服务均已删除”仅指上一阶段；原型 profile 清理仍待正常卸载。本轮证据位于 `local-machine-*`、`same-path-*`、`candidate-public-check-*` 以及 `service-candidate/`。

后续专项结果（22:44–22:57，仍非最终发布验收）：

- 正常打开刚重装的原生 MathType，出现中文“无标题 1”编辑窗口。新 Office 数据适配器连续三次真实渲染成功，原生 MathType/MathTypeLib 的 PID 和创建时间未变化。
- 新代码的 `word-mathtype-headless` 通过：MathType OLE 插入、保存重开、OMML 转 MathType；没有新增原生编辑窗口，原有交互 MathType 进程身份未变。此模式使用临时测试宿主，不是已安装插件的最终验收。证据在本机 `%LOCALAPPDATA%/Temp/VisualTeX/acceptance/vsto-flow/artifacts/mathtype-install-20260927/service-candidate/word-private-render/`。
- 新代码的 `powerpoint-mathtype` 通过：插入、编辑、位置/标签/动画保留、保存重开、独立 X/Y 缩放保留、过期目标拒绝。第三方存储用例因本次未指定真实源文件明确跳过，没有计为通过。证据在 `service-candidate/powerpoint-private-render/`。
- 将原生排版实现从通用 WindowsOleBridge 移入 `VisualTeX.NativeRenderWorker`，移除通用桥的 `--mathtype-preview-manifest` 入口及引擎里的已安装路径发现/环境变量后备。对应测试改走数据服务，不保留旧共享预热/同名进程清理契约。
- 两项源码边界检查通过；9 项协议/私有依赖单测通过，默认运行跳过需显式启用的 live 用例。独立 worker 重新发布并更新临时服务，新入口又完成三次实际渲染。Word/PPT/验收项目再次编译成功。

现场待确认：专用 `word-mathtype-manual-coexistence` 已在新空白 Word 中临时加载新代码，等待用户物理点击原生 MathType“显示”、返回公式并关闭测试文档。状态及结果在 `service-candidate/native-display-manual-new-client/`。该宿主最多等待十分钟，仅管理自有测试文档并在结束时恢复临时自动加载设置；不能把普通宏调用成功当作此前物理点击崩溃已经排除。字体依赖和正式部署/NSIS 仍未完成。

### 22:59 物理点击复现失败

用户在新代码测试宿主的空白 Word 中点击原生 MathType“显示”，Word PID 90576 于 22:59:19 崩溃。WER 为 `wwlib.dll 16.0.14334.20918`、`c0000005`、偏移 `0x2ee4e8`，与 20:08 的故障位置一致。测试宿主随后访问 Documents 得到 `0x800706BE`，这是 Word 已崩溃后的 RPC 失败，不是已确定的根因。此前后台渲染和自动化插入通过不能替代这一失败。原 `manual-state.json` 的等待状态未在异常分支更新，不能据此判断仍在运行。

本次证明 runtime 隔离尚未解决物理点击故障；下一步须排除测试宿主 COM 轮询和 VisualTeX 选区事件重入，并采集原生调用栈。未宣称最终修复，未发布安装包。

后续诊断准备：已保留该 WER，并把本次 `manual-state.json` 更正为 Failed。手工验收宿主移除 100ms 的 Documents/InlineShapes 轮询，只在操作员明确发出返回完成信号后读取和保存；异常必须写 Failed。另纠正验收宿主退出时无条件连接安装版插件的错误，现在仅恢复原本已连接的实例。这两项是测试可靠性修正，不当作用户闪退的根因修复，专项编译通过。

23:07 启动 PID 57716“MathType 原生独立对照”：启动期间临时抑制 VisualTeX 自动加载，核对 Connect=false，然后恢复注册表原值并释放全部自动化对象；没有订阅事件、没有 COM 轮询，未打开用户文档。等待用户在该空白实例物理点击“显示”的结果。恢复列表保持“以后查看”，没有删除恢复文件。已编译仅面向指定测试 Word PID 的本机崩溃抓取助手，但尚未附加或运行，不影响这一轮对照。临时原生服务仍在运行；此对照排除加载项，不排除所有机器全局因素。

23:10 用户确认原生独立对照正常打开 MathType；后续 UI 观察到测试文档内 MathType 对象。保存辅助脚本挂起后仅终止该辅助脚本，使用正常窗口关闭结束测试，没有新增 Word 崩溃记录。此对照没有证明所有安装路径无问题。

23:14 新代码无轮询测试 PID 56796：用户随后表示“正常打开并返回”的回答是误按，要求重测，因此该回答作废。宿主完成信号后的检查实际未发现返回的 MathType 对象，状态为 Failed；调试助手记录进程正常 exit 0，无 AV。该轮既不是通过，也不是重现崩溃。已重新启动 `native-display-no-poll-retry1`，待真实物理点击。

23:18 重测 PID 55000：用户确认“正常打开”。日志确有 Word 前台物理点击和随后的选区事件，各事件返回；崩溃抓取暂未记录 AV，Windows 应用程序日志未见该轮新 Word 崩溃。只确认“显示”打开阶段通过，尚未收到输入公式并返回 Word 的明确确认，未创建完成信号、未标为完整插入通过。这不能单独证明旧崩溃由轮询造成；此轮还启用了交互提醒和调试器，后续结论需控制这些因素。

23:20 用户进一步明确确认公式正常出现。操作员随后发出完成信号，宿主检测到 1 个 MathType OLE 对象并保存 `native-display-no-poll-retry1/manual-native-insertion.docx`，状态 Passed、exit 0；抓取助手记录 Word exit 0 后正常脱离，无 AV。这是本轮“物理显示插入→编辑→返回→保存→退出”通过，不能归因为轮询已经证明是旧崩溃唯一原因。已安装 VSTO 及 NSIS 服务部署仍未切换，不把这一轮作为正式安装版最终验收。

### 正式安装集成（23:38 起）

沿用通过验证的虚拟账户 / Session 0 / 私有 HKCU / 自有 Job 架构，不更换排版引擎。新增独立内部 MSI `VisualTeX-NativeRender-x64.msi`，由 Windows Installer 负责服务安装、停止、升级和卸载；NSIS 内嵌该 MSI，不再分发可由交互会话直接启动的裸 runtime。Office 安装脚本先验证并安装服务，再更换插件。

- 安装位置 `C:\Program Files\VisualTeX\NativeRender`；服务身份仍为 `NT SERVICE\VisualTeX.NativeRender`，绝不使用 LocalSystem 或当前用户。
- MSI 的公式临时目录只允许服务账户、SYSTEM、管理员访问。初次打包验证发现 Windows Installer 创建 jobs 子目录时保留了原 Program Files ACL，产生 UnauthorizedAccessException；已在 MSI 中为 data 与 jobs 分别声明保护 ACL，不通过放宽权限或运行时补丁处理。
- 23:41 正式 MSI 升级后，10 项专项测试全部通过（含两个并发客户端、各 5 个原生公式）。1,374 项公共注册配置无新增/修改，原生 MathType.exe 与 MathTypeLib.exe 的 PID、启动时间、路径均未变化。
- 临时 candidate 服务已由它自己的受限卸载脚本移除，数据日志留存；正式服务取代该临时组件。包校验新增实际 MSI ServiceInstall、Registry、MsiLockPermissionsEx 表检查，避免仅靠源码字符串判断。
- 仍不把原始 MathType 安装器的隐藏界面问题、从未安装过字体的纯净环境、安装版 Word 的完整验收标为已解决。它们必须有独立证据。

部署复核：安装目录 Word 插件 SHA-256 为 `36F25EFED9EAF6659B744DDE5BA1788E1E6B32DB69DC665D0E0916132C28AEFB`，本轮测试插件为 `8F85DB98059CB38D90839B1CB5763C372F795722A3425673172652A9D9307AF4`，确实不是同一版本。现有 bundle 仍只打包旧 `mathtype-runtime/`，没有新服务的正式部署链；不能把当前输出直接打包交付。另尝试按实际 wwlib RSDS 标识 `03732562-5A0F-4337-BF61-C5B5D1382B7C` age 2 查询微软符号，公开 PDB 地址返回 404，没有得到故障函数名，不以无符号的偏移猜测根因。未上传任何文档或转储。

诊断原始材料位于本机忽略目录 `artifacts/mathtype-install-20260927/`，没有上传第三方。

厂商安装界面说明：<https://docs.wiris.com/installing-mathtype-7-for-windows-and-mac>。

注册表重定向作用范围：<https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regoverridepredefkey>。

### 2026-09-28 实际安装版崩溃转储与事件边界候选

最新接手证据与后续状态以 `docs/mathtype-word-event-reentrancy-20260928.md` 为准。00:14:29 的共存崩溃没有外部 COM 轮询；00:34 再次在实际安装版 Word PID 27508 物理操作原生 MathType Ribbon 后捕获同偏移完整转储。托管栈确认是 `OnWindowSelectionChange → NormalizeTypingCaretAfterInlineFormulaCore → WordMathTypeHostAdapter.ResolveLocal → ReadMetadata → MathTypeWordOpenXml.Read → Range.WordOpenXML` 进入 `wwlib+0x2ee4e8`，不再仅凭猜测归因。

候选只修改 WordFormulaService/ThisAddIn 的被动观察与显式命令边界：被动光标规范化不导入第三方 MathType 内容；Ribbon 被动字号查询不导出其 OLE，实际字号操作仍读取权威 MTEF。已编译并部署到实际安装路径，备份旧 DLL；候选 SHA256 为 `C1A9BCD0539F8778CBAE000185FD9C89E0D175874C290D763473047BC7562118`。没有重启/更改原生服务或关闭原有 Word PID 93104。

PID 77076 的候选测试完成九次正常关闭公式的原生 Ribbon 插入（显示/左编号/右编号各三次），抓取期间无 AV。另存文件共十一份原生 MathType OLE，全部 MTEF 能正确解码，七份保留原生编号结构，无 VisualTeX 书签。首次测试脚本使用 Alt+F4 退出整个 MathType，后续已纠正为厂商规定的 Ctrl+F4 关闭公式；含退出操作、预置警告和初始化对话框的早期序列未当作正常共存验收。

该 00:55 检查点尚未最终通过；后续 01:09–01:11 进展如下，以最新记录为准。

01:09 更严格旧版对照已完成：实际安装旧 DLL 的 Word PID 30020，原生“显示→输入 x=1→Ctrl+F4 正常关闭公式→右编号”再次崩溃。原生 MathType PID 38260 保持运行；第二份完整转储仍为 `wwlib+0x2ee4e8`、读取 `0x3ac`，托管栈与前一份完全同路，证据在 `coexistence-20260928/baseline-ctrlf4-20260928-010917/`。因此故障不依赖 Alt+F4 或退出整个原生 MathType，也不是外部文档轮询造成的唯一现象。

R2 同一所有权边界补全：发现 `IsMathTypeOle(InlineShape)` 的缺失/未知 ProgID 识别回退本身也会读 WordOpenXML，因此被动字号查询改用无内容读取的宿主标识检查；显式命令仍保留完整 CFB/MTEF 识别。新增的定向编译 DLL 测试能检出旧版六项被动导出及 R1 两项遗漏；R2 六项被动观察均为零导出/激活/剪贴板复制/选区写入，显式命令仍可进入内容读取，第七项也通过。未修改旧测试的正确行为，没有跑全部历史回归。

01:10:11 已将 R2 部署回实际安装路径，当前 SHA256 为 `8E63D124CBA192D0055C692DCDC60D91794DCEDA29741B8F7B06CE224FE48FB5`，不是重现用的旧 DLL。原生 MathType 38260 / MathTypeLib 22688、原有 Word 93104 和渲染服务保持。R2 无调试器测试 Word PID 68756 已确认加载实际安装 DLL（MVID `c2e6692a-3dd0-4fbc-a875-10f9eb002d4e`），但物理插入测试在前台检查处停止：另一份 Word 文档“保存此文件”对话框遮挡当前测试窗口，未向那份文档发送输入、保存、取消或关闭。该次测试尚未进行插入，不能当作 R2 UI 通过。

最新状态：R1 九次实际原生循环与保存后的十一份真实 MTEF 通过；R2 七项定向编译代码检查通过、已安装，完整 UI 验收待继续。需用户先处理当前已有保存对话框，再验证 R2 无调试器原生/VisualTeX 交替插入、编辑、保存重开。最终新 NSIS、两种安装顺序、原安装器退出及完全无 MathType 字体残留环境仍未完成。全部证据及精确继续步骤已写入 `docs/mathtype-word-event-reentrancy-20260928.md`，未进行 Git 提交或推送。
