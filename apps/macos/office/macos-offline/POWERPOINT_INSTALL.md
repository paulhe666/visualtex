# 在 PowerPoint 中登记 VisualTeX.ppam

VisualTeX 安装插件后，第一次打开 PowerPoint 时会通过 PowerPoint 自带的 AppleScript 命令 `register add in` 自动登记并加载 `VisualTeX.ppam`，并停用手动添加的其他 VisualTeX 加载项。它不会模拟菜单点击，也不会为此启动 PowerPoint。macOS 可能询问是否允许 VisualTeX 控制 PowerPoint，PowerPoint 可能询问是否启用宏。之后更新会始终覆盖同一路径，不需要重新登记。

以下手动步骤只在自动登记失败时使用（例如拒绝了自动化权限）：

1. 在 VisualTeX 的“macOS Office 集成”设置中点击“在 Finder 中显示 PPAM”。
2. 打开 Microsoft PowerPoint。
3. 选择“工具 → PowerPoint 加载项”。英文界面对应 “Tools → PowerPoint Add-ins”。
4. 点击添加按钮，选择 Finder 中显示的固定文件：

   `~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam`

5. 确认 VisualTeX 加载项已勾选，然后关闭对话框。
6. 退出并重新打开 PowerPoint。VisualTeX Ribbon 应自动出现。
7. 回到 VisualTeX 设置页刷新状态。只有 `powerpoint.json` 中出现由 `Auto_Open` 写入的 `loaded=true`，才表示运行时加载成功。

## 更新

VisualTeX 始终原路径覆盖 `VisualTeX.ppam`。文件名不会附加版本号，因此 PowerPoint 保存的登记路径不会失效。

## 修复

“修复原生加载项”会重复执行安全的文件安装、重新编译 AppleScriptTask，并重建占位图片，不会打开 PowerPoint，也不会操作任何演示文稿。

## 卸载

先在 PowerPoint 的加载项对话框中取消勾选或移除 VisualTeX，再从 VisualTeX 设置页点击“卸载原生加载项”。卸载不会删除文档中的公式图片和缓存预览。
