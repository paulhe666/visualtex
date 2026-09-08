# Word 公式复制粘贴：2026-09-08 验收记录

工作区保留原有修改；本轮未提交、未推送。用户文档1（本轮 PID 47836）未修改、保存或关闭。

## 当前实现

复制由 Word 原生命令完成。前台 Word 的剪贴板变化只触发源公式快照；粘贴完成后的 STA 回调检查对象数量变化和局部宿主，随后补齐副本身份及编号。无全局键盘拦截。VisualTeX OLE/OMML 的 FormulaId 与行 UUID 重新生成；MathType 保持原生 Equation.DSMT4 数据，复制其原生编号模板和左右位置。

本轮追加修复：相邻 OLE 粘贴导致旧 VTO 书签扩张覆盖两个对象；空段落粘贴将行内 OMML 自动升级为 display；整行粘贴后光标位于下一段，局部定位找不到复制宿主；整行 MathType 已有 MTPlaceRef 时不可再次追加；正常 Ribbon 插入/编辑必须刷新复制快照计数，不能作为粘贴处理。行内粘贴不覆盖目标正文段落标记的字体格式。

## 实机候选

Stage：`copy-complete-host-v4`。独立 Word `/x` PID 106044。
候选目录：`%LOCALAPPDATA%\VisualTeX\office\remediation-builds\FB80BC533C77788EBC22C408599F9339CC93CCC1CC8E17B247E317D2911AA0C0`。

所有功能操作从真实 Word Ribbon / VisualTeX Office 编辑器发起；复制粘贴使用真实 Ctrl+C / Ctrl+V。COM 仅准备正文、移动 Selection 和读取结果。源码修改验收启用 `VISUALTEX_PERF_REQUIRE_CHANGE=1`，等待实际会话完成。

## 已通过场景

- 三种行内公式：Arial 14 pt 正文、Normal 10.5 pt。各复制一次、粘贴两次；VisualTeX 与 MathType 包含第二次粘到第一个副本前方的相邻场景，OMML 包含正文和空段落两个目标。9 个公式保留，三种副本均可分别从 Ribbon 编辑；原公式和另一副本不变。
- 行内格式：OLE 尺寸与源公式相同；OMML 字体、14 pt、实际 Type=inline 均保留；VisualTeX 副本 FormulaId / 行 UUID 不重复，VTO 严格绑定单个物理 OLE。
- 三种带编号行间公式：完整 VisualTeX 段落、完整 OMML 1x3 表格、完整左编号 MathType 段落分别复制粘贴，均形成独立编号；没有重复添加可见编号或 MTPlaceRef。
- 分别编辑上述三个行间副本后，原始 OLE 二进制和原始 OMML 内容不变。公式数量、两个 1x3 表格和两个 MTPlaceRef 均保持。
- 复制 MathType 后通过 Ribbon 正常插入一个 VisualTeX OLE：原复制格式未覆盖新插入公式；MTPlaceRef 数量保持 2。
- 行内、行间测试文档保存重开。读取实际 CFB 内嵌 `VisualTeX.Formula.json`，持久化 FormulaId 唯一；EMF 预览与 MathType `Equation Native` 数据保留。已检查真实 Word 截图。

## 证据及复核

`evidence/copy-inline-v4-{sources,copies,edited,reopened}-inventory.json` 及其对应详细 JSON / Flat OPC。
`evidence/copy-inline-v4-{edited,reopened}-embedded-storage.json`。
`evidence/copy-inline-ui-validation.json`：最新四阶段校验通过。
`evidence/copy-block-v4-{sources,all-hosts,explicit-insert,edited,reopened}-inventory.json`。
`evidence/copy-block-v4-reopened-embedded-storage.json`。
`evidence/copy-complete-host-v4-word-hook.log`：本轮成功操作，无 word-operation-failed / paste-repair-failed / paste-repair-timeout。
`evidence/copy-inline-v4-final-visible.png`、`copy-block-v4-final-visible.png`：真实 Word 截图。

复核命令：

```powershell
python docs/remediation-3d207d7/validate-copy-inline-evidence.py copy-inline-v4-sources copy-inline-v4-copies copy-inline-v4-edited copy-inline-v4-reopened
```

说明：Word 在 Save 时将运行中的 VisualTeX OLE 元数据写入 IStorage，保存前后的 CFB 整包哈希可以不同。因此保存重开验证比较内嵌身份、语义数据和预览流，不将正常持久化误判成损坏。MathType 比较 Equation Native 流。

Word VSTO 编译和 WindowsOffice 462/462 已通过；OCR/editor parity、Office OCR HTTP transport、多行 OCR 归一化通过。本轮没有重新跑完整 100 公式性能矩阵；也没有把跨 Word 进程剪贴板、多公式混合选区或所有自定义样式宣称为已验收。

## 最终正式包及相邻 OMML 补验

正式包补验发现 Word 会把相邻的行内 OMML 合并，因此新增 `WordFormulaService.CopyPaste.OmmlBoundary.cs`。仅当复制部分与剪贴板快照的内容指纹一致、另一部分与已有公式的持久化指纹一致时，才将已合并的局部 OMath 分开。使用普通零宽 Word 字符隔开两个 OMath，不添加可见空格或新段落，已有公式保留原身份。无法证明完整复制的 Word 内部子表达式输入，不套用整份复制元数据。

最终 Stage：`copy-release-v6-binary`，独立 Word PID 77092。从正式打包输出注册 DLL，并已比对 SHA256 完全一致：
`37FD776521F61559B238F55766489427296083B32DB6A0C6BD81596411648902`。

真实 Ribbon 将源 OMML 编辑为分式 `\frac{a+b}{c+d}`，实际复制一次，在已有 `u+v=8` 前连续粘贴两次，OMath 数量由 3 变 5，两个新副本有独立身份、14 pt 行内格式，段落数量不变。再从 Ribbon 只编辑一个副本为 `s+t=10`，源式、另一副本及紧邻的 `u+v=8` 均不变。保存重开后内容、身份和格式保留。证据：`copy-adjacent-v6-{before,two-copies,edited,reopened}-inventory.json`；严格校验结果：`copy-adjacent-omml-validation.json`。

同一正式 DLL 还复测了 VisualTeX / MathType 行内复制（`copy-release-v6-all-types`），以及三种完整编号宿主复制（`copy-release-v6-block-before`、`copy-release-v6-block-all-hosts`）。后者最终 7 OLE / 3 OMML / 3 个编号表格 / 3 MTPlaceRef，原有 5 个 OLE 数据不变，副本尺寸与原式一致，身份不重复。最终 hook 无 word-operation-failed / paste-repair-failed / paste-repair-timeout。

NSIS 构建采用既有 `npm run tauri:build`，仅跳过可能影响打开 Office 的安装烟测。最终构建以 `evidence/copy-paste-v6-release-result.json` 为准（exitCode=0）；静态 Release / 前端嵌入校验通过。

产物：`apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.6_x64-setup.exe`。
大小：323865158 bytes。
SHA256：`774F73B31C5DA94E1221B52DC4F7A1F02889D1619B113E072A7758D337F05655`。
早期 v4/v5 安装包已被此最终产物替代，不作为本轮交付。
