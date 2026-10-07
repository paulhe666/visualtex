# NSIS 验收包（2026-10-06）

2026-10-07已由包含MathType颜色修复的CP325包取代；见[当前打包报告](nsis-acceptance-stage-20261007.md)。以下保留CP315及更早包的证据与限制。

## 当前包：CP315 重编译验收包（CP317交付）

[VisualTeX_1.2.8_x64_CP315_Acceptance_20261006.exe](../artifacts/nsis-acceptance-cp315-20261006/VisualTeX_1.2.8_x64_CP315_Acceptance_20261006.exe)，**376939702字节（约359.5MiB）**，SHA256 **21660726A9233105C7C630E7753EF12530D3EF13EAA2A573D3F60B329765AAA9**。版本1.2.8，Office MSI1.0.43.0，含两种Office架构和标准离线OCR运行时。此次实际重编译，产品行为同CP309；上一轮批量替换调查只新增测试/报告，没有新增产品性能优化。旧CP309/CP304包原位保留。

Office/OLE/MSI、离线加载探针、前端、Rust Release、NSIS及静态检查通过。实际解出NSIS和三MSI，核对两架构各19个Office文件、Office前端23、NativeRender83、OCR74、Bridge/VSTO和实际桌面嵌入资源；最终状态`VerifiedEmbeddedPayloadNotInstalled`，见`artifacts/nsis-acceptance-cp315-20261006/verified-embedded/report.json`。

- x64 Word `0BD8CCE72979B50FD7B49BF824AFBAAE7B60AC24AE560C70D41F8AA9161272C2`、PPT `938849565B2CFBB1A8B2602CF540EF3525E1ACE86ECBFC15ACA81462BE3C3325`，与已验候选相同。
- x86 Word `E17144B9A903725607334F88B89CE6B0985607D1B8E27A8CC5387D8B72151046`、PPT `E8193FAE6A0BD68A18816270BE526B7B5D748329BCFC3CA9A869F3832A6680AE`，与CP309相同。
- 新MSI x64 `549B51902DA9D575F16D4FD0E96C3F42CFD9C3359F2BE449CAADF206E365A00B`，x86 `07773E7C88961DBFDC891874F798F3EFDECD60D8CA6461AA22F0B282AE45B59E`。
- 内嵌Companion `B86B7295910E36759F1AF19279D5DDC0A493588F97A7610C3370277F7CFC9F4E`，构建输出 `EF4F8C8949D4EFDCEEC38C272A0B77A2C2122056057F6C951F290686584F67A9`，仅Tauri唯一3字节bundle标记不同，其余严格相等。

`environment-final.json`确认8处LoadBehavior、原进程身份、已安装CP304 DLL未变，无新增受保护进程；`source-recheck.json`确认锁定输入未变。未安装/提交/推送，冻结R3/heldR4和既有未提交修改保持。构建验证不扩大功能验收结论：CP255、完整六向保真/视觉缺口、CP314严格OMML保存XML差异与异常文档性能问题保持原状态。

## CP315 构建过程记录（已完成，见上文）

CP316进展：两架构Office/OLE/MSI和离线加载/COM探针、两前端TypeScript/Vite及引用核对通过，Rust Release进行中；产品DLL四项哈希均与CP309相等。新MSI哈希见本轮build/build-state.json，旧包保留。NSIS和实际解包尚未完成。

按用户要求从当前工作树实际重新编译，输出`artifacts/nsis-acceptance-cp315-20261006/VisualTeX_1.2.8_x64_CP315_Acceptance_20261006.exe`。上一轮CP312–314只新增性能诊断，产品预期仍同CP309；不声称增加性能修复。冻结R3/heldR4、标准OCR及x64/x86 Office组件保持；沿用已审阅分阶段构建、实际解包校验和环境快照入口，旧包保留，不安装/提交。

当前新包尚待构建和实际载荷核验；最终哈希与结果将在完成后记录。`verified-inputs.json`明确记录原CP309清单只有测试Program.cs变化，产品DLL预期严格保持。既有测试失败/未验收范围不因打包转为通过。

## 上一包：CP309 进度修复（CP311 交付）

路径：`artifacts/nsis-acceptance-cp309-20261006/VisualTeX_1.2.8_x64_CP309_Acceptance_20261006.exe`，版本1.2.8，376945197字节，SHA256 **10AE3A80112BF804CF88F86E421CDF637A87A4808662AAAAC25D8345B5CDF237**。包含最新重绘进度/确认窗口生命周期修复和CP304转换规则，取代下文旧包供本问题验收。旧包原位保留。

x64 Word `0BD8CCE72979B50FD7B49BF824AFBAAE7B60AC24AE560C70D41F8AA9161272C2` 精确同真实进度测试候选；PPT `938849565B2CFBB1A8B2602CF540EF3525E1ACE86ECBFC15ACA81462BE3C3325` 未变。x86 Word `E17144B9A903725607334F88B89CE6B0985607D1B8E27A8CC5387D8B72151046`，PPT `E8193FAE6A0BD68A18816270BE526B7B5D748329BCFC3CA9A869F3832A6680AE`。Office MSI x64 `A63A790D9512CA4E58AB7B8257E6EC65CB27421DF69218ECEFCE4B4E32DAAA94`，x86 `8FE9976896AAA2203BECDEA7897520D0A79AF577F48221B3CB9BFB812E291CE0`。

两架构Office/加载探针、前端、Rust、NSIS及静态验证通过；实际解包核对Office各19文件、前端23、NativeRender83、OCR74和Bridge/VSTO。内嵌Companion `DF9436E9A9F9481B4956C7C7DAACAB1E383DD5389313CFD1F75ABA6FA6D6E841`，与本轮构建输出 `D73111E762E6111E42DC85AA332328DBF9219BD4D1B2BB0A5681784AA2855CF5` 仅3字节UNK→NSS标记不同，其余字节严格相同。证据 `artifacts/nsis-acceptance-cp309-20261006/{build/build-state.json,verified-embedded/report.json,environment-final.json}`，最终 `VerifiedEmbeddedPayloadNotInstalled`。

仍是冻结R3/未提交工作树构建，heldR4保持。打包前后已安装CP304 DLL、8处LoadBehavior和原有进程身份保持，无新增Office/Companion，不安装/部署/提交。实际进度测试见[问题报告](redraw-progress-stage-20261006.md)：正常确认后的20条重绘、原200副本准备/取消及阻塞STA窗口刷新通过；本轮没有200全转换或六向完整保真新结论。其他既有缺口不变。

## 以下保留 CP304 打包历史

用户授权编译供其验收，本轮不安装、不修改注册表或安全设置、不提交。输入为HEAD `66a8c70994cbb80316c65283b0e7034b4931b324` 和现有未提交修改；使用冻结R3，保留heldR4及全部旧产物/失败证据。进度及构建/核验脚本位于 `artifacts/nsis-acceptance-cp304-20261006/`。

## 已完成：CP307 交付

交付文件：`artifacts/nsis-acceptance-cp304-20261006/VisualTeX_1.2.8_x64_CP304_Acceptance_20261006.exe`。

- 版本：VisualTeX 1.2.8 x64，Office MSI 1.0.43.0，包含x64/x86 Office组件和标准离线OCR运行时。
- 大小：376,950,662字节（约359.49 MiB）。
- SHA256：`C9D071B67CF53E818405B91504DC387C28B72C32C6D0FE364C314AE3AE6D6838`，邻接 `.exe.sha256` 文件同值。
- 状态：已编译、静态检查、实际反向解包核验；未执行安装包。

| 实际内嵌组件 | SHA256 |
| --- | --- |
| Word x64 | `279CA2F19998DEF308B69AE9BB8B8D121CF557B41CC484B44F8F564D02645C3E` |
| PowerPoint x64 | `938849565B2CFBB1A8B2602CF540EF3525E1ACE86ECBFC15ACA81462BE3C3325` |
| Word x86 | `14D37116142A687EB24A456A53178855AB30571E57EDE85169752B574424A6C5` |
| PowerPoint x86 | `E8193FAE6A0BD68A18816270BE526B7B5D748329BCFC3CA9A869F3832A6680AE` |
| Office MSI x64 | `F99668E53E23370CB1D2A960141C6592A83A7D7222C33EACEEDB1FBF8A4435F3` |
| Office MSI x86 | `9A88BB98D22EA4DE0E21171501AF859FC2F492CA302A7C2CE84B4C75DA5AF4A3` |
| Companion | `2D2F170013E926AAA241878E43AB1C874727FCBBE07CD3136D0B9E6D9CBBAAC3` |

x64两DLL与CP304已验候选逐字节一致。实际NSIS与三MSI已解包，逐项核对两架构各19个Office文件、Office前端23、NativeRender83、OCR74、Bridge和VSTO runtime；实际桌面前端嵌入索引、2个当前JS/CSS及19字体检查通过。Companion与本轮构建输出仅Tauri唯一3字节UNK→NSS标记不同，其余字节严格相同；没有修改DLL或放宽比较器。离线OCR Python3.12.10、71wheel闭包、OpenMP及模型排除按既有静态检查通过。原编译警告完整保留，没有错误豁免。

最终证据：`verified-embedded/report.json` 为 `VerifiedEmbeddedPayloadNotInstalled`，`build/build-state.json` 为构建完成时 `BuiltStaticVerifiedPendingExtraction`；`build-console.log` 和 `verify-console.log` 保存完整输出。`environment-final.json` 确认8处LoadBehavior、已安装CP295 Word/PPT DLL以及原3个Companion/NativeRender进程身份未变，无新增受保护进程。CP304源码与冻结输入验证哈希未变，Git HEAD未变；旧包及既有失败保留，没有安装、部署、用户文档操作、注册表/安全设置修改或提交。

本包包含CP304通用规则修复和此前CP300的进度/PPT原生编辑修复。打包没有重跑公式或宿主验收；本包与已验x64候选同字节，但x86仅构建/离线加载探针。CP255宽重音、混合文本/中西文样式及此前特殊符号/间距/上下标和六向视觉缺口仍开放。由用户安装此确切包复验原文档和PPT编辑；不能据本轮打包认定全项目发布验收通过。

## 构建中（CP305）

目标 `VisualTeX_1.2.8_x64_CP304_Acceptance_20261006.exe`，标准含离线OCR运行时版本，两种Office架构。重建x64/x86 OLE/VSTO/MSI、Office/桌面前端和Rust Release，生成NSIS。已有Bridge、NativeRender源码相对HEAD无差异且不包含本次共享codec；沿用已重建同源码产物并核对哈希。OCR沿用锁定资源，常规S/M/L模型包不内置。

最终x64 Word必须与CP304已测候选同SHA `279CA2F19998DEF308B69AE9BB8B8D121CF557B41CC484B44F8F564D02645C3E`，PPT同 `938849565B2CFBB1A8B2602CF540EF3525E1ACE86ECBFC15ACA81462BE3C3325`。构建入口核对 `final-verification.json` 中源码和冻结输入哈希。默认顶层的清理/安装冒烟入口不调用；只运行已审阅分阶段命令。

包生成后将解出实际NSIS与内嵌Office/NativeRender MSI，验证DLL、前端、Bridge、OCR及桌面嵌入资源。开始/结束只读记录当前注册表LoadBehavior、已安装DLL和受保护进程身份；不操作用户文档或现有进程。

本轮没有新的公式功能验收声明。沿用CP304的377扩展+原200静态preflight/OLE读回、PPT codec377、native离屏20、完整离线2453/2454；唯一CP255失败、混合文本/中西文样式及此前符号/间距/六向视觉缺口保持。构建与包内容核验不能替代用户实际安装和Office验收。

## Office 及前端已完成，桌面程序构建中（CP306）

两架构OLE、VSTO、MSI和离线Ribbon COM分发/依赖加载探针均通过。新x64 Word/PPT与CP304上述预期SHA逐字相等，x86 Word为 `14D37116142A687EB24A456A53178855AB30571E57EDE85169752B574424A6C5`、PPT为 `E8193FAE6A0BD68A18816270BE526B7B5D748329BCFC3CA9A869F3832A6680AE`。Office/桌面TypeScript与Vite构建、前端引用核对通过；现进入Rust Release。现有警告保留，无错误豁免。证据 `build/build-state.json`、`build-console.log`。NSIS生成和逆向解包尚未完成，不能提前宣称交付成功。
