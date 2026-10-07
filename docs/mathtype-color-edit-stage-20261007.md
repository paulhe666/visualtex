# VisualTeX 编辑 MathType 公式颜色丢失（2026-10-07）

## 结论

问题已在共享 MathType MTEF 边界复现并修复。前端颜色操作本身有效，编辑结果中的 MathML 含 `mathcolor`；旧 `MathTypeMtefCodec` 没有把这些颜色写入 MathType 的 COLOR_DEF/COLOR 记录，读回时也丢弃颜色，因此 Word 中的 `Equation.DSMT4` 仍显示原颜色。

本轮不仅验证内存读回，也用脚本新建隔离 Word 文档，执行真实产品插入与编辑写回，检查 Word 中的 MathType OLE 预览，保存并重开后再次检查。最终结果为通过。

## 根因边界

- `EditorWorkspace.tsx` 调用 `applySelectionStyle({ kind, value })`，`MathEditor.tsx` 最终调用 MathLive `applyStyle({ color })`；这条路径会产生正确的 LaTeX/MathML 颜色。
- `MathMlColorContract` 已覆盖 OMML 的 `w:color` 写入、读取和继承，不是本问题缺口。
- `MathTypeMtefCodec` 旧实现只为若干原生模板补固定黑色定义；任意 `mathcolor` 不会生成 MTEF 调色板或颜色状态。MTEF 读取器遇到 COLOR/COLOR_DEF 只跳过，导致编辑后的颜色既不显示，也无法在下次编辑中恢复。
- 旧 `SemanticSignature` 不含颜色，因此单靠原语义签名会把有色与无色公式视为相同。

MTEF v5 的 COLOR 是持续到下一条 COLOR 的流状态，RGB COLOR_DEF 的三个分量使用0–1000整数。实现遵循该状态模型，而不是在预览层重新着色。参考：<https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later>。

## 修改

生产代码：

- `apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/MathTypeMtefCodec.cs`
  - 收集显式 MathML 颜色，复用已有RGB定义或追加新的 COLOR_DEF。
  - 用作用域写入 COLOR，并在嵌套颜色退出时恢复父颜色。
  - 原生模板要求黑色时，若当前有显式颜色则尊重显式颜色。
  - 读取RGB/CMYK COLOR_DEF和COLOR状态，把非黑颜色投影为 `mathcolor`。
  - 数字和文本合并增加颜色相等条件，防止颜色边界被吞并。
  - 带显式颜色重写时不复用源首对象的root-leading COLOR，防止颜色泄漏到后续未着色字符。
  - 非黑颜色进入语义签名；无色/黑色公式保持旧签名，避免无关历史数据变化。

测试与验收：

- `apps/windows/src-windows/VisualTeX.WindowsOffice.Tests/MathTypeColorDefinitionTests.cs`
  - 红色行、蓝色根分式、绿色矩阵、红色重音，以及红→蓝→红→默认黑的嵌套编辑。
  - 验证首次写入、读回、再次重写及未着色尾字符不串色。
- `apps/windows/src-windows/VisualTeX.VstoFlowAcceptance/WordOwnedMathTypeColorEditAcceptance.cs`
  - 新建自有 Word 文档，不接入用户 Word。
  - 插入黑色 MathType OLE，再走 `ReadSelection` + `ReplaceMathTypeOle` 写入选择颜色。
  - 检查 MTEF 读回、Word 增强图元预览、DOCX 保存及重开后的OLE显示缓存。
- `Program.cs` 增加隔离入口 `word-owned-mathtype-color-edit`。

## 复现与验证

修复前新增的四类颜色测试全部失败，读回颜色均为 null。修复后专项颜色11/11通过。

最终冻结R3构建：

- 目录：`artifacts/omml-edit-audit-20261001/build-cp322-mathtype-color-edit-r4`
- Word DLL SHA256：`353C20B4B1EFA31E3A53FB1E51E392F6970DAF5DEFDE9912A8FE4251BE58A18C`
- 验收EXE SHA256：`334AE895142862266B6F80BF574F935580D3BF9BF43D6AFCCE6F5894C24374E4`
- heldR4 SHA256：`6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`

最终独立 Word 运行：

- 证据目录：`artifacts/mathtype-color-edit-20261007/cp322-owned-r3-final`
- 自有 Word PID：106552，运行后已退出。
- 文档：`VisualTeX-MathType-Color-Edit.docx`
- 文档 SHA256：`71BE2526689603B7A1F2A2ADA74294CD87634E18DE226233950C8E84E74D32FE`
- 对象：1个，ProgID `Equation.DSMT4`。
- 应用后读回：x=`#DC2626`、y=`#2563EB`、+=`#DC2626`、z=`#DC2626`、q=默认黑。
- 应用后预览像素：红4152、蓝1676、黑3442。
- 保存重开后预览像素：红4622、蓝1848、黑3772。
- `applied-preview.png` 来自 Word 剪贴板增强图元；重开时剪贴板持续忙，`reopened-preview.png` 使用同一DOCX/OLE的持久WMF显示缓存。黑色占位SVG不能产生红/蓝像素。
- 两张图均已查看；字符、颜色和未着色q一致。

第一次独立运行在重开后的剪贴板取图失败，应用后的MTEF颜色读回及实际预览已经通过，但该轮仍记录为验收失败。增加持久OLE预览备用取证后，第二轮和最终R4轮完整通过；旧失败目录保留。

完整已编译测试DLL在正常用户环境直接运行：2459项中2458通过、1失败。唯一失败是当前HEAD已有的CP255宽重音契约冲突：测试要求未声明stretchy的 `~` 为固定单字符EMBELL，共享 `MathMlAccentGeometry` 则按字典默认使用可伸缩HatBox。本轮未修改该旧测试或几何规则，也不把它写成颜色回归。沙箱内完整运行另外两个失败是受控AppData目录写入被拒；正常用户运行时已通过。一次正常权限重新构建因宿主NuGet路径指向不存在的 `C:\Users\CodexSandboxOffline` 而未启动测试，也单独记录为构建环境失败。

## CP325 NSIS 验收包

用户随后要求编译供实际安装验收。已有CP324安装包完成于颜色修复源码写入之前，不能代表本修复，因此没有交付旧包；重新以当前颜色源码、冻结R3和heldR4构建CP325：

- 安装包：`artifacts/nsis-acceptance-cp325-20261007/VisualTeX_1.2.8_x64_CP325_MathTypeColor_Acceptance_20261007.exe`
- 大小：377000590字节。
- SHA256：`A2908B2629FFA055CA3A267DFCD69274027C000E8A527718ED966E2CDA19C99F`。
- 内嵌x64 Word：`353C20B4B1EFA31E3A53FB1E51E392F6970DAF5DEFDE9912A8FE4251BE58A18C`，与上面的真实Word颜色验收候选逐字节一致。
- 内嵌x86 Word：`D95B24CA60B308A6163D14A211AA3D8BD4F9607C44292881B1187CBC2AB8C1C5`。
- 内嵌Office MSI：x64 `58648D317E0C5D4F8F12AE06D55598FCFFCA9C851BF94DBD3986AB3A925DBAD3`，x86 `85C2390455DA9364471E5E47F63865B7C80E8B0733D901B43FE72D4DE722B1D2`。

NSIS仅反向解包，未执行安装。两种Office架构各19个实际文件、Office前端23、OCR74、NativeRender、Bridge及桌面嵌入资源全部与构建输入一致；最终状态为 `VerifiedEmbeddedPayloadNotInstalled`。注册表LoadBehavior、已安装DLL和原进程身份均未变化，没有新增受保护进程。完整打包报告见 `docs/nsis-acceptance-stage-20261007.md`。

## 范围

已完成共享代码修复、专项测试、冻结R3候选构建、新建隔离Word真实OLE写回、保存重开、视觉像素验收、x64/x86 Office及NSIS打包和实际载荷反向核验。未安装候选、未提交Git、未修改注册表、安全设置或用户文档。当前HEAD为 `dc501c9c6fe607b1aef0ef1404981b62d1668027`，本轮源修改仍未提交。
