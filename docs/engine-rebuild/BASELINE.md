# P00 — 当前工作副本与用户环境基线

采集：2026-09-13 03:48–03:49 UTC（本机 +08:00 为 11:48–11:49）。
开始版本：`c4690c2e69e266e6273eb0eaf5803ce74a054d1f`，分支 `reference-3d207d7`。
实际 workspace：`ws_7e5a0b84a3`。继续复用，不使用原计划中的旧 workspace id。

## 状态与准入

源码和现有资源冻结已建立，可进行隔离规范/测试建设。**P00 尚未达到正式 Office 接入门槛**：未重跑当前安装版的完整 UI/性能基线；安装 EXE 与构建目录 EXE 哈希不同，尚未完成包内逐文件核对；安装 DLL 不等于已经证明 Word 实际加载了该 DLL。上述缺口禁止在 P12/P16 被自动视为通过。

本轮允许新增路径：`docs/engine-rebuild/**`、后续 `engine/**`。不改生产源码、默认编辑器、安装文件、现有产物、根构建配置；不触及 macOS/Web。

## 机器清单

不可覆盖的初始清单：`checkpoints/20260913-P00-provenance.json`。
其中记录 65 个 tracked dirty 文件的路径和 SHA256、11,650 个历史 untracked 路径集合的 SHA256（不含本目录）、Git index SHA256、12 个保护分组、12 个二进制/安装清单。
保护分组覆盖 4,479 个 Windows 工作副本文件；包括未提交修改。基线绝不是 HEAD 文件的替代。

检查命令：

```text
node docs/engine-rebuild/tools/inspect-baseline.mjs --check docs/engine-rebuild/checkpoints/20260913-P00-provenance.json
powershell -NoProfile -ExecutionPolicy Bypass -File docs/engine-rebuild/tools/inspect-office.ps1
```

脚本只读且输出到 stdout；不修改工程、不 Save/Close/Quit Office。
2026-09-13 03:49:49 UTC 第一次保护检查：12 分组、12 产物，PASS。

## Git 与构建输入

本轮 `git ls-remote origin refs/heads/main`：`242d82767cfb4eb22f5b266c09bfbe3db6473649`。
与原计划的 `6bb4fce` 不同。没有 fetch、pull、合并、切分支或推送。
没有 staged 变更；没有 commit。存档点是独立证据文件，不声称是 Git 提交。
现有历史修改保留原样，包括 Office 编号/转换和前端源码；不可混入本项目提交。

本机 rustc/cargo 1.96.0；Node v25.4.0；.NET SDK 9.0.316；初始 Rust 仅安装 `x86_64-pc-windows-msvc` target；未发现 wasm-bindgen/wasm-pack 命令。`rg` 不在 PowerShell PATH，改用受控 Node/PowerShell 只读检索，不更改全局 PATH。

## 当前产物对应关系

完整包：`VisualTeX_1.2.7_x64-setup.exe`，346885204 bytes，SHA256 `324d9acbb13dccd7cc31962f85f9f3a021843784d4e4c1a9627d6db452665cd0`。
无 OCR 包：`VisualTeX_1.2.7_x64-no-ocr-setup.exe`，119602423 bytes，SHA256 `6bfcde5b4e19899fb1510965b5657ab8b960312ec1ff6db698279efa891dfdeb`。

当前 x64 MSI SHA256 `5b873c79a4e4445c39fdd17032efc84143101dcfe78378eefb664a668c05a898`。
当前 x86 MSI SHA256 `e85fb08dba928838a66e49582c40ad5d23f2cf7fde98e15d9151c063f8d984eb`。
x64 子包清单声明的 Word DLL SHA256 与安装目录实际文件相同：`cac453292a1dded4c4a2b62baa79966003f9edd8ba0ec8c07f2ec938d3c0b287`。
同一清单声明的 FormulaOleServer SHA256 也与实际安装及运行文件相同：`bfae55d0e26472da0e2c7d8270ef43ddc03b4daed0121ccde9e054532e66512d`。
这证明清单与已安装两个文件对应，不证明 NSIS 内部所有资源或全部 UI 已验收。

安装 EXE：`d7dde16cdb0aa1e211e9e8acd2bc111b6a3b62bf22228d3706054fdd1512ae3d`。
构建目录 EXE：`3a91ffb887121a6d92f65d9363ccccb842b590a2206a52f460f63447939e7444`。
两者不同，原因待核查。不得推断为相同代码或安装错误；不得为消除差异覆盖安装。

原计划中的 1.2.6 包、503/503 和批转换耗时仅为历史参考。本轮没有复现这些历史成绩。

## 用户文档与运行进程

只读采集：Word PID 43084，文件版本 `16.0.14334.20848`；VisualTeX PID 75320，版本 1.2.7，来自 `%LOCALAPPDATA%/VisualTeX/visualtex.exe`；FormulaOleServer PID 118232，来自 Program Files 的 VSTO 目录；没有运行 PowerPoint。

Word 正打开两个用户文档：`文档1` 和 `C:/Users/pojian_liao/Downloads/latex_1000chars_100formulas.docx`。两者 Saved=false、ReadOnly=false、ownedByThisRun=false。
本轮拥有的测试文档列表为空。任何后续测试文档必须采用 `VT_ENGINE_20260913_<runId>` 命名，并登记创建证据和身份；不能仅凭名字或未保存状态判断所有权。

Word 进程模块列表可见 Office16/MathPage.WLL，SHA256 `3056f7ad528cc0abec12abb4383626a048288310c9fb665841848d99cdddf7475`。该次模块枚举未显示托管 WordVsto 程序集；不据此推断其未加载，也不声称已核验加载路径。

## 已执行验证与未覆盖

`npm --prefix apps/windows run test:office-session-safety`：PASS，使用当前真实 TS payload decoder，涵盖默认值、非法枚举、非有限几何和错误 payload；不操作真实 Office。
保护区检查：PASS。
真实 Ribbon/UI 回归、当前安装版性能、包内全量对应关系、PowerPoint、macOS、Web：NOT_RUN。

下一步：P01 固定可执行契约和 ADR；所有 Office 接入门槛继续关闭。首次新增内核代码只允许在独立 workspace/sandbox 中验证。
