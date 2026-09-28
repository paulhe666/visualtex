# MathType 转换与插入弹出原生窗口修复

## 根因

现场独立 MathType 主程序已不存在，弹窗进程实际为
`%LOCALAPPDATA%\VisualTeX\mathtype-runtime\MathType.exe -mtrpc`。
同一进程出现 `MathType - Untitled 5`、`MathType - Untitled 6`。

私有 runtime 的 `PrivateRuntimeScope.TryCreate` 在每个预览 sidecar 中无条件
`Process.Start(MathType.exe, -mtrpc)`，随后又调用支持按需启动的 `MTInitAPI(0, 8)`。
已有 RPC 服务时，再启动 MathType 会进入其单实例编辑器处理，创建 Untitled 窗口。
`CreateNoWindow` 和隐藏启动不是 GUI 不出现的保证。
卸载独立 MathType 使转换改走这条内置 runtime 分支，暴露了重复启动问题。

## 修改

- 删除多余的显式启动、进程字段及 Dispose；统一由 MathPage 的 launch-as-needed 初始化连接或启动服务。
- 保留只作用于 sidecar 的 HKLM/HKCR 虚拟视图；不改 MathType 公共注册、语言或用户首选项。
- 未添加窗口隐藏轮询、强制关闭编辑器或转换后杀进程的产品补丁。
- 源检查不再要求有缺陷的显式启动；新增冷启动/连续预览和 Word 无窗口定向验收。

正常后台渲染仍可存在无窗口的 MathType RPC 进程；本修复禁止的是转换/插入意外打开原生编辑器，不改变用户明确调用原生编辑器的行为。

## 验证

证据目录：`docs/remediation-3d207d7/evidence/mathtype-headless-20260927/`。

- `warm/result.json`：连续三次预览，全部成功，未检测到编辑窗口，公共 COM 注册保持不变。
- `cold/result.json`：零 MathType 进程开始，冷启动与后续两次预览成功；首轮约 2.1 秒，后续约 0.11 秒；复用单个后台进程，无编辑窗口。
- `word-focused/acceptance.log`：实际 Word 插入 MathType、保存/重开、校验预览与源码、OMML 转换通过；25ms 周期检查未检测到原生窗口。使用从最终 NSIS 提取且哈希一致的 sidecar。
- `powerpoint/acceptance.log`：实际 PowerPoint 新建、编辑、保存/重开、缩放保持和过期目标拒绝通过；没有独立 MathType 安装。第三方公式用例本次未运行。
- `test_windows_mathtype_runtime_isolation_source.ps1` 通过。

此前调用现有 `word-omml-mathtype-single-performance`，行内/行间转换通过，但其后续 managed OMML 夹具缺失 VTOMML 书签断言失败。没有修改产品去迁就此断言，也不将该整项宣称通过；新专项测试只覆盖此次启动问题并完整通过。

未跑全量旧回归，未自动安装，也未人工逐项点击安装后的 Ribbon。电脑操作技能用于只读确认窗口列表；窗口截图接口返回错误，未将它当作截图验收通过。

## 交付

`apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.8_x64-setup.exe`

- 大小：347,977,017 字节。
- SHA-256：`EAF207DB8A6C2A990CC9A1E1E277767B7330B4C046D277CFD85CDBD3DD0736EB`。
- 内含 sidecar SHA-256：`66193DF814EEDCA6D149A02289C2CA328FD1B045E7FFA41EE9B237E343A75485`，已从安装包提取核验。
- 只重编发布受影响的 sidecar，再使用既有前端、主程序、Office MSI 和完整 OCR 资源重新打包 NSIS；未运行安装流程。
- 同内容副本：`VisualTeX_1.2.8_x64-mathtype-headless-setup.exe`，便于与之前同版本安装包区分。

关闭 Word、PowerPoint、VisualTeX 后安装此包即可更新。源码尚未提交 Git。
