# 在 PowerPoint 中登记 VisualTeX.ppam

VisualTeX 不会模拟菜单点击，也不会根据界面语言或窗口坐标操作 PowerPoint。首次安装后需要手动登记一次，之后更新会始终覆盖同一路径，不需要重新登记。

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
