# LaTeX 重绘进度反馈（2026-10-06）

## CP311 交付状态

修复已包含在新的 `artifacts/nsis-acceptance-cp309-20261006/VisualTeX_1.2.8_x64_CP309_Acceptance_20261006.exe`（376945197字节），SHA256 `10AE3A80112BF804CF88F86E421CDF637A87A4808662AAAAC25D8345B5CDF237`。NSIS和实际内嵌MSI已解出并逐项核验，x64 Word精确同下述已测候选 `0BD8CCE7…1272C2`。源码仍未提交，安装包未执行，已安装DLL/8处LoadBehavior/原有进程保持。详见[打包报告](nsis-acceptance-stage-20261006.md)。

结论：已复现并修复“扫描完成进入确认时，旧扫描进度仍显示”的生命周期错误，以及准备过程不同阶段被混同的提示问题。正常确认后20条真实重绘、原200副本准备与取消、阻塞Office时刷新均通过。用户现场40–50秒没有分段trace，不能声称所有实际耗时已根除；原200新测约8.3秒包含约6.1秒读取字号，未删除其内容/格式保护。未重跑200全转换或六向矩阵，既有保真缺口不变。

## CP308 状态与初步根因

用户安装CP304后反馈读取/扫描公式范围停留40–50秒，认为扫描几秒已完成。现场只读DLL SHA确为 `279CA2F19998DEF308B69AE9BB8B8D121CF557B41CC484B44F8F564D02645C3E`。暂不把总停留全部认作扫描性能或窗口遮挡：需区分实际capture、字号准备、等待确认和渲染。

已确认代码缺陷：`CaptureLatexRedrawPlanCore` 中几个不同阶段均未上报，沿用初始“正在读取公式范围”；`RedrawLatexAsync` capture返回后进入`LatexRedrawDialog.ShowDialog()`，进度窗尚未退出，仍显示旧阶段并作为Word owned窗口浮动。确认框本身没有显式Word owner。此前正常确认被acceptance环境跳过，独立进度窗计时测试不能证明这段交互正确。

新增 `WordRedrawProgressAcceptance.cs` / Program的隔离入口：仅创建自有Word文档，20条简单LaTeX的真实capture，临时只在运行器内恢复正常确认分支并自动取消。记录窗口/owner/实际可见标签，检查取消后原文与对象数不变、operation gate释放；Word子进程保持原acceptance环境，安装插件仍被抑制。测试不会调用Companion渲染或修改用户文档。构建冻结R3基线 `build-cp308-progress-baseline` 与已安装CP304同DLL字节。

全部证据在 `artifacts/redraw-progress-20261006/`，入口 `Run-OwnedProgress.ps1` 在前后读取8处LoadBehavior和受保护进程身份。基线/修复结果待写入；尚无新修复验收或新安装包声明。

## CP309 复现与修复

`baseline-r2` 在CP304同字节DLL上明确复现：真实20 capture约0.26秒，800ms确认等待后仍可见“正在读取公式范围”和计时，确认owner不同于Word。没有阶段上报、确认期残留进度、owner不正确三个断言失败。首次baseline-r1源文本不变断言用了Word规范化前字符串；改为保存操作前/后实际Range.Text，两文本SHA均 `06A5D3EE4F0AE467D30A713966D52585FC2AFEE514E40C5E4100C2D2216191B9`，原测试失败保留。

`WordOperationProgress.cs` 提供原Word owner和可指定初始阶段。`ThisAddIn.cs` / `ThisAddIn.FormatConversion.cs` 在确认前dispose扫描进度，并将重绘/全文格式转换/公式转LaTeX确认框显式归属Word；接受后创建执行进度。重绘完成/错误通知前也结束进度。`WordFormulaService.LatexRedrawCore.cs` 按真实扫描、坐标定位、位置核对、字号索引和逐公式字号处理报告阶段与计数，保留全部原坐标/正文保护。重绘渲染补计数，连接/预热/MathType预检明确分阶段；本轮没有性能算法或codec改动。

`fixed-r1` 在冻结R3 Word `0BD8CCE72979B50FD7B49BF824AFBAAE7B60AC24AE560C70D41F8AA9161272C2` 上通过相同测试：确认框可见时没有进度窗口，其native owner等于Word，取消原文/对象数保持，gate释放。后续补正常接受后的完整20重绘及原200 capture/取消。

测试工具审计补充：最初baseline/fixed-r1在自有Word中手动OnConnection仍执行默认Companion后台预热，故此前“不会调用Companion”范围表述需收窄；这几次没有转换请求或用户文档操作，但不是零预热调用。现有Companion与8处LoadBehavior/原进程身份保持。后续 `Run-OwnedProgress.ps1` 要求显式runtime manifest，使用自有PID80352、43136端口与独立数据目录；预热也一并隔离。此历史不抹去。

## CP310 最终真实路径验证

候选 `build-cp309-progress-final` Word SHA仍为 `0BD8CCE72979B50FD7B49BF824AFBAAE7B60AC24AE560C70D41F8AA9161272C2`，冻结R3/heldR4不变；同r1产品字节。测试均只用自有Word，后两正常路径和窗口测试明确绑定独立Companion。

| 运行 | 结果与边界 |
| --- | --- |
| baseline-r2 | 旧CP304三个预期失败：无细阶段、确认时扫描进度残留、确认owner不正确；原文不变 |
| fixed-r1 | 相同20准备/正常确认/取消检查通过；此轮后台预热隔离限制见上 |
| confirm20-final | 正常接受确认后真实20 MathType重绘、逐项表达式读回、写入20/20进度、新执行进度生命周期、保存、gate释放通过；约9.963s含800ms确认观察，非性能基准 |
| capture200-final | 原200归档文档副本的直接capture及正常命令capture/确认/取消通过；取消原文不变、确认时无进度遮挡、owner正确；未执行200转换 |
| window-final | 实际自有进度窗在Office STA阻塞1300ms时仍显示阶段、37/200与00:01，释放回调通过 |

原200命令capture共8.285s：snapshot56ms、坐标1855ms、fontIndex58ms、逐目标准备6314ms（其中字体fallback6113ms，计200次）。这些嵌套时间不求和。此前字号索引重建23117字符与source23113不同，mismatch228，因此走已有精确字体fallback；本轮如实报告该阶段，没有为了进度修改算法、缩减校验或编造扫描百分比。直接capture独立计时7.972s。用户现场40–50秒没有原始分段trace，不能把全部等待归因已复现的确认窗口问题或宣称消除全部耗时。

修改范围为进度上报/确认生命周期及上述验收工具，codec未变。全文格式转换/公式转LaTeX的相同确认逻辑已修正并编译，但其真实点击未在本轮逐一运行。完整离线/六向视觉/全200转换没有重复，历史结果和CP255/保真缺口不变。

`Finish-OwnedRuntime.ps1` 核对PID、路径、SHA与创建时间后结束仅自有Companion80352；`environment-final.json` 原3进程身份、8处LoadBehavior均保持，无新增Office/Companion，已安装Word仍CP304。`final-inputs.json` 留存最终构建/源码哈希。准备新的CP309 NSIS并解包核验，只编译不安装或提交。
