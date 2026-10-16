# 公式格式转换：Word 每行双按钮与 PowerPoint 批量转换（2026-10-15）

范围：Windows Office 插件。Word 的“公式格式转换”改为每个方向一行、行内直接放“所选/全文”两个按钮；PowerPoint 新增 MathType、OMML（原生公式）、VisualTeX 三种格式之间的批量互转，范围可以是所选、本页或整个演示文稿，并有“一键统一”、进度和停止。

> **决定（2026-10-15）**
> - Word：去掉原来的分割按钮和“默认范围”开关，每行两个小按钮“所选 / 全文”，行为与原菜单项完全相同。
> - PowerPoint：每行三个按钮“所选 / 本页 / 全部”；“全部”开始前确认（提示先保存）。
> - 转换的标准是“观众看到的不变”：位置、显示大小、颜色、动画、触发器、超链接、名称、可见性、叠放次序都保持。做不到的情况一律报告，绝不静默丢失。
> - “→ VisualTeX”的结果形式由“VisualTeX 形式”菜单决定（OLE 对象或 SVG 图片），默认 OLE。

---

## 1. 界面

**Word**（“公式格式转换”组，两块，每块是“名称列 + 所选列 + 全文列”）：

| 行 | 所选 | 全文 |
|---|---|---|
| VisualTeX → MathType / MathType → VisualTeX / OMML → MathType | `OnConvert…Selection` | `OnConvert…Document` |
| MathType → OMML / VisualTeX → OMML / OMML → VisualTeX | 同上 | 同上 |

**PowerPoint**（“公式格式转换”组）：

- 六行方向 × “所选 / 本页 / 全部”，Ribbon tag 为 `源|目标|范围`（如 `mathTypeOle|wordOmml|slide`）。
- “一键 统一”菜单：本页或全部 → MathType / OMML / VisualTeX（tag 源为 `any`，已是目标格式的不动）。
- “VisualTeX 形式”菜单：OLE 对象 / SVG 图片，保存在 `HKCU\Software\VisualTeX\PowerPointVsto\FormatConversion\VisualTeXTargetMode`。
- 中英文 Ribbon 的控件 ID 集合相同，由 `test_windows_vsto_ribbon_dispatch.ps1` 检查。

## 2. 范围

| 范围 | 规则 |
|---|---|
| 所选 | 选中的对象（光标在文本框里时取该文本框；选中组合内的对象时按组合处理）；在缩略图中选中幻灯片时为这些幻灯片上的全部公式 |
| 本页 | 普通视图的当前幻灯片；幻灯片浏览视图下取选中的第一张 |
| 全部 | 演示文稿中的全部幻灯片（不含母版、版式和备注） |

## 3. 识别

一次 `SaveCopyAs` 快照，读取幻灯片 XML 建立顶层形状索引，再用 COM 确认：

- **MathType**：嵌入的 `Equation.DSMT*` OLE；从 Equation Native 读 MTEF → MathML → LaTeX。链接对象只报告。
- **OMML**：文本框中的 `a14:m` 公式；只含公式（可有多个段落）的文本框才转换，每个公式各成一个对象。
- **VisualTeX**：VisualTeX OLE 或带 VisualTeX 元数据的图片；读元数据中的 LaTeX。

## 4. 不转换、结束时报告的情况

| 原因 | 说明 |
|---|---|
| 与文字混排的 OMML | 只能整框替换，混排时会破坏文字 |
| 组合内 / 表格内 | 不拆组合、不改表格结构 |
| 旋转或翻转 | 目标能表示时保留（旋转：OMML、SVG；翻转：SVG），否则报告 |
| 一个文本框多个公式且带动画 | 拆开后动画无法一一对应 |
| 链接对象、无法读取的公式 | 无可转换的源 |
| 开始后被改动 | 写入前用新快照核对，变化的跳过 |

## 5. 流程

1. 取得操作锁，显示进度窗口（独立线程，Office 忙时也能刷新和点“停止”）。
2. 生成计划：范围 → 快照 → 识别 → 跳过项。
3. “全部”范围确认。
4. 进入输入保护：禁用本进程所有 PowerPoint 主窗口，结束后恢复。
5. 新快照核对源未变。
6. 准备：VisualTeX 目标用伴随程序批量渲染（每批 20，相同公式只渲染一次）；MathType/OMML 目标直接使用源 MathML；MathType 目标编码 MTEF 并生成排版预览（每组 24 个。MathType 渲染器在长批次中、以及服务刚启动时会因自身状态偶尔崩溃：缺失的公式再按 4 个一组、再逐个重试；上一轮有成功（渲染器在恢复）时最多再逐个重试 3 轮；90 秒内没有任何公式渲染成功即停止。仍失败的公式单独报告，不连累整批；规则在 `PowerPointPreviewRetry`）。
7. 写入：OMML 目标先预热内核；MathType 目标一次性在隐藏的暂存演示文稿中生成全部对象，每 12 个一组粘贴、整组用一次快照核对、核对通过后才删除源。
8. 每个公式一个撤销项；连续 3 个失败即停止；停止只在两个公式之间生效。
9. 收尾：恢复视图和当前幻灯片；“所选”范围结束后选中新对象；删除临时文件；报告结果和跳过项。

批量转换在工作线程上运行，每一次 PowerPoint 调用都通过 UI 调度器执行，PowerPoint 在两个公式之间正常刷新，长批次不会出现“未响应”。

## 6. 位置与大小

- 以公式本身的视觉中心对齐（OMML 用公式区边界，不用可能更宽的文本框）。
- 显示字号：VisualTeX 由尺寸反推；MathType = MTEF 字号 × 缩放；OMML = 公式区字号；渲染限制在 5–200 pt，超出部分用放置缩放补足。
- 新的 OMML 文本框自动收缩到公式大小，再按公式中心定位。

## 7. 颜色（显示颜色不变）

显式颜色通过 MathML `mathcolor` 在三种格式之间传递。未着色部分的显示颜色因格式而异：OMML 继承幻灯片文字颜色（深色主题为白色），MathType 与 VisualTeX 为黑色。规则：

- 源为 OMML：读取各 run 实际显示的颜色（主题已解析），排除显式颜色后取用得最多的作为“基色”；目标为 MathType/VisualTeX 且基色不是黑色时，在 MathML 根上写入该颜色（内部显式颜色优先）。
- 源为 MathType/VisualTeX：基色为黑色；新建 OMML 文本框的默认文字颜色不是黑色时，同样写入黑色。
- 浅色主题下两者都不写入任何颜色，原生公式继续跟随主题。

## 8. 保留的属性

| 属性 | 做法 |
|---|---|
| 主序列动画 | 动画刷（`PickupAnimation`/`ApplyAnimation`）复制全部设置，再 `MoveTo` 到原效果之后 |
| 触发器序列中以公式为对象的动画 | 在原序列中用 `AddTriggerEffect` 重建（退出、时长、延迟、重复、回卷） |
| 以公式为触发器的动画 | `TriggerShape` 改指新对象 |
| 超链接/动作、标签、标题、替代文字 | 复制（VisualTeX 元数据不复制到其他格式） |
| 名称 | 第一个对象沿用源名称；拆分出的其他对象为“名称 (2)”…；`VisualTeX_<id>` 与内部暂存名不带入其他格式 |
| 可见性、叠放次序、连接线 | 复制 / 原位置插入 / 端点改连 |
| 占位符 | 删除填充的占位符后 PowerPoint 会补一个空占位符，转换后删除（标题除外） |

## 9. 发现并修复的既有问题

- **动画丢失**：原替换路径用 `Effect.Shape = target` 改指动画，源被删除时 PowerPoint 仍把该效果一起删掉（持有的效果对象仍显示新形状，保存后的文件里已没有）。所有替换路径（含单个 MathType 编辑）改为第 8 节的做法；`powerpoint-mathtype` 验收改为从幻灯片重新读取效果，并检查保存后重新打开的文件。
- **取消被当作伴随程序故障**：传入已取消的令牌时，`VisualTeXSessionClient.EnsureHealthyAsync` 会误判伴随程序未运行、启动一个 VisualTeX 进程并报告“companion validation failed”。现在先检查取消，取消一律以取消上报。
- **深色主题颜色**：见第 7 节，转换前白色公式会变成黑色（在深色背景上看不见）。

## 10. 实现位置

| 文件 | 内容 |
|---|---|
| `VisualTeX.WordVsto/ThisAddIn.cs`、`WordRibbonLocalization.cs` | Word 两块按钮（中/英） |
| `VisualTeX.WordVsto/ThisAddIn.FormatConversion.cs` | 删除默认范围代码，回调不变 |
| `VisualTeX.PowerPointVsto/PowerPointFormulaConversionModel.cs` | 纯规则：tag 解析、范围、尺寸、跳过说明、幻灯片 XML 分析、颜色规则（链接进单元测试） |
| `VisualTeX.PowerPointVsto/PowerPointFormulaService.FormatConversion.cs` | 计划、核对、写入、MathType 暂存、收尾、命名 |
| `VisualTeX.PowerPointVsto/ThisAddIn.FormatConversion.cs` | Ribbon 回调与流程编排、进度、报告 |
| `VisualTeX.PowerPointVsto/PowerPointFormulaService.MathType.cs` | 动画/触发器转移、可见性、占位符清理 |
| `VisualTeX.PowerPointVsto/PowerPointNativeOmmlKernel.cs` | `CreateStandalone` 可选基色 |
| `VisualTeX.WindowsOffice.VstoShared/OfficeOperationProgress.cs` | 共享进度窗口（停止、完成后短暂显示结果） |
| `VisualTeX.WindowsOffice.Contracts/VisualTeXSessionClient.cs` | 取消处理 |

## 11. 测试

| 测试 | 内容 |
|---|---|
| `PowerPointFormulaConversionModelTests` | tag、范围、尺寸、跳过说明、幻灯片 XML、颜色规则、预览重试（崩溃后分组重试、只留下真正失败的公式、刚启动的服务恢复、坏掉的服务 90 秒放弃、字体缺失立即停止） |
| `VisualTeXSessionClientBatchingTests` | 已取消的健康检查只报告取消 |
| `test_windows_vsto_ribbon_dispatch.ps1` | 两个插件中英文 Ribbon：每个回调有 DispId、ID 唯一、中英文 ID 集合一致；Word 每行两个按钮；PowerPoint 24 个转换按钮 |
| 验收 `word-owned-ribbon-scope` | 独立 Word 中六行 × 所选/全文：“所选”只转换选中的那一个，“全文”转换其余全部 |
| 验收 `powerpoint-format-conversion` | 本页、所选、SVG 形式、旋转保留、一键统一、停止、空选择；动画、触发器（两个方向）、名称、占位符；保存后重新打开仍在 |
| 验收 `powerpoint-format-conversion-color` | 深色主题：继承白色、显式红色、部分着色在六个方向上保持；黑色公式不变白 |
| 验收 `powerpoint-format-conversion-scale` | 6 张 × 10 个公式整份统一为 MathType、VisualTeX OLE、OMML：全部转换、位置不变、Ribbon 回调立即返回（2026-10-15：约 22–32 s / 49–54 s / 33–37 s） |
| 验收 `powerpoint-mathtype` | 单个 MathType 编辑的位置、标签、超链接、动画（含保存后） |

`powerpoint-basic-functions`、`powerpoint-undo-grouping`、`powerpoint-text-safety`、`powerpoint-native-edit-entry`、`powerpoint-native-styles` 要求没有正在运行的 PowerPoint，需在关闭 PowerPoint 后运行。

## 12. 不在本次范围

- 与文字混排的 OMML、组合和表格中的公式：只报告。
- 母版、版式、备注页中的公式。
- OLE ↔ SVG 的批量互转（单个公式在编辑窗口的“保存为”中改；Ribbon 上旧的“转为 OMML / 转为原生 OLE / 转为 SVG 图片 / 删除所选公式 / 打开 VisualTeX”五个按钮已删去）。
- Word 的“一键统一”菜单；Word 进度窗口改用共享进度类。
- Word 转为 MathType 仍把全部预览放在一个请求里，渲染器崩溃时整次转换失败（文档不变）；可套用 PowerPoint 的分组重试。
- 用主题色（`schemeClr`）显式着色的局部 OMML：按基色处理，不保留各自的主题色。
- 首次渲染时伴随程序的转换器预热约 10 秒，之后同一会话内几乎为零。
