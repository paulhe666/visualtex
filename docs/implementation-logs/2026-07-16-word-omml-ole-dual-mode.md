# Word OMML / OLE 双模式公式支持

日期：2026-07-16

## 完成功能

- Word 公式主体新增 `wordOmml` 模式，与现有 `nativeOle` 模式并存。
- Ribbon 新增：
  - 插入行内 OMML 公式
  - 插入行间 OMML 公式
  - 将所选公式转换为 OMML
  - 原有“转为原生 OLE”继续用于 OMML → OLE
- “编辑所选公式”同时支持 OLE 和 OMML。
- OMML 双击不被 VSTO 拦截，保留 Word 原生公式光标和公式工具栏编辑。
- OLE 双击继续进入 VisualTeX 编辑器。
- OMML 与 OLE 支持双向转换，并保留同一 `formulaId`、编号和交叉引用关系。

## 实现结构

- MathJax 将 LaTeX 转为 Presentation MathML。
- Windows 端使用 Office 自带 `MML2OMML.XSL` 转换为真正的 OMML。
- 通过临时 DOCX + Word `FormattedText` 写入原生 `OMath`，不使用剪贴板。
- OMML 公式使用 Word 原生零长度 Bookmark（`VTOMML_<uuid>`）作为稳定锚点。
- LaTeX、公式行、显示模式、编号状态等元数据保存在 DOCX Custom XML Part：
  - Namespace: `urn:visualtex:word-omml:1`
- Bookmark 位于公式外侧；Word 执行 `Linearize / BuildUp` 重建 OMath 后锚点仍能保留。
- Custom XML 更新使用“先添加新 Part，再删除旧 Part”的事务顺序，避免 `LoadXML` 在 OMath 重建后被 Word 拒绝。
- 行间 OMML 继续复用原有 Word 原生编号体系：
  - `SEQ` 题注域
  - `VTEqNum_*` 原生编号 Bookmark
  - `VTEq_*` 可见编号 Bookmark
  - `REF` 交叉引用域
  - Word 自带“引用 → 交叉引用 → 公式”目标列表

## 关键兼容处理

- 显示型 OMath 前存在 Word 数学段分隔符 `0x0B`，编号布局会越过该字符识别前导 Tab。
- 原生编号 Bookmark 只覆盖域结果文本，不能依赖 `bookmark.Range.Fields`；改为按 Bookmark 位置扫描文档原生 `SEQ` 域。
- OMML 更新和 OLE → OMML 使用 `FormattedText` 原位替换，避免 Word Range 自动扩展误删新公式。
- OMML 原生编辑与 VisualTeX 编辑均可用。每个 OMML 公式保存内容指纹；再次用 VisualTeX 编辑时，如检测到 Word 原生公式发生变化，则通过 Office `OMML2MML.XSL` 转换当前 OMath，再生成可编辑 LaTeX。常见分式、根式、上下标、积分求和、希腊字母、矩阵和定界符均已覆盖；反向生成的 LaTeX 允许规范化，不保证与最初输入逐字符一致。

## 已通过测试

- TypeScript `tsc --noEmit`：通过。
- Windows Office 前端生产构建：通过。
- SVG / MathML 导出 smoke：8 类公式通过。
- Office metadata smoke：通过。
- Windows Office architecture smoke：通过。
- C# 单元测试：74 / 74 通过。
- Word VSTO x64：编译通过。
- Word VSTO x86：编译通过。
- PowerPoint VSTO x64 / x86：编译通过。
- ATL Formula OLE Server x64 / Win32：编译通过。
- x64 / x86 MSI：生成通过。
- VSTO Ribbon dispatch：x64 / x86 通过。
- VSTO dependency loading：x64 / x86 通过。
- 真实 x64 Word / PowerPoint 验收：通过，包括：
  - OMML 插入与识别
  - Word 原生 Linearize / BuildUp 编辑循环
  - 原生 SEQ 编号
  - 原生交叉引用目标与 REF 域
  - OMML → OLE → OMML 往返
  - DOCX 保存并重新打开
  - OLE 离线缓存预览、重新注册后更新、再次离线打开

真实验收产物目录：

`src-windows/artifacts/test-logs/native-office-ole-20260716-234531`

安装包：

- `src-windows/VisualTeX.WindowsOffice.Installer/bin/x64/Release/VisualTeX-WindowsOffice-VSTO-x64.msi`
- `src-windows/VisualTeX.WindowsOffice.Installer/bin/x86/Release/VisualTeX-WindowsOffice-VSTO-x86.msi`

## 2026-07-17 真实使用问题修复与回归验收

针对实际 Word / PowerPoint 截图中发现的问题，完成以下修复：

- Word 带编号的 OLE 与 OMML 公式：
  - 可见编号不再写入 `m:oMath` 或 `m:oMathPara` 内部。
  - 不再使用会重排公式段落的 `Range.InsertCaption`；改为在独立隐藏段落中创建原生 `SEQ` 域。
  - 可见编号继续使用公式段落右对齐 Tab + `REF` 域，保持 Word 原生交叉引用能力。
  - 清除公式段落继承的项目符号、分页、与下段同页、段中不分页等属性，避免“显示格式标记”时出现黑色方块。
  - 修正隐藏原生编号的白色 1 pt 字体被传播到可见 `REF` 结果的问题。
- OMML → OLE：
  - OLE 在原 OMath 结束位置、右侧编号锚点之前生成，旧 OMML 范围和 Custom XML 元数据完整删除。
  - 转换后选择结果确认为 `nativeOle`，不再保留 OMML 选择框。
  - 双击路由改为使用 Word 双击事件传入的真实 Selection；OLE 进入 VisualTeX 公式编辑会话，OMML 留在 Word 原生公式编辑器。
- PowerPoint 图片 → OLE：
  - 尺寸换算优先保持物理高度，避免公式字体视觉上被压扁。
  - 按新公式自然宽高比重算宽度。
  - OLE 初始化及缓存刷新后再次恢复 Left / Top，避免 PowerPoint 将对象移动到幻灯片外。
  - 转换前后保持对象中心位置。

新增真实使用场景自动化验收：

- Word 同一文档同时存在编号 OMML 与编号 OLE。
- 检查公式段落无项目符号、无分页黑方块相关属性。
- 正文分别插入两个公式的原生交叉引用。
- 在文档顶部新增公式后，原公式编号从 1/2 自动更新为 2/3，正文 REF 同步更新。
- 删除顶部公式后，编号和 REF 自动恢复为 1/2。
- OMML → OLE → OMML 往返后检查对象类型、锚点、编号和双击路由。
- Word 原生 `Linearize / BuildUp` 编辑循环。
- DOCX 保存、关闭并重新打开。
- PowerPoint 使用被非等比调整过的图片公式转换为长 OLE 公式；检查高度、自然宽高比、中心位置以及保存重开后的几何尺寸。
- Word / PowerPoint OLE 离线缓存预览、重新注册、更新以及再次离线打开。

最终测试结果：

- C# 单元测试：84 / 84 通过。
- SVG / MathML 导出 smoke：8 类公式通过。
- Office metadata smoke：通过。
- Windows Office architecture smoke：通过。
- Windows Office 前端生产构建：通过。
- 真实 x64 Word / PowerPoint 九阶段验收：通过。
- 编号行间公式真实页面几何检查：公式中心与正文中心误差约 0.2 pt，编号末端与正文右边界误差约 0.2 pt，公式和编号纵向基线坐标一致。
- 同一编号公式连续 5 轮 OMML → OLE → OMML：段落位置、公式中心、宽高比、编号、书签、交叉引用和编辑路由全部保持有效。
- Word 原生 OMath 增加 `+z^3` 后，调用实际 VSTO“编辑所选公式”回调；VisualTeX 编辑 Session 的 `lines` 成功包含 `z^{3}`。
- 完整 VSTO 页面会话验收：Word 21 阶段、PowerPoint 10 阶段通过。
- Word / PowerPoint VSTO x64、x86：构建通过。
- Formula OLE Server x64、Win32：构建通过。
- Office MSI x64、x86：构建通过。
- Tauri / NSIS x64 EXE：构建通过。

最终真实验收产物：

- 底层 Word / PowerPoint 验收：`src-windows/artifacts/test-logs/native-office-ole-20260717-022401`
- 完整 VSTO 编辑页面会话验收：`src-windows/artifacts/test-logs/vsto-flow-20260717-023026`

最终安装包：

`src-tauri/target/release/bundle/nsis/VisualTeX_1.1.7_x64-setup.exe`

SHA-256：

`BFE45D76BCCF8889D2880166802181A8B2FB1451FAD8CB62D487669991431EFE`
