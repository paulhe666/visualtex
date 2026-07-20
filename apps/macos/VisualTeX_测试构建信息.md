# VisualTeX 测试构建信息

构建日期：2026-07-18

分支：`feat/macos-offline-office-native`  
构建基准 HEAD：`2b1cf2d0054e436895a8540e4d2be8092619b9f5`

本轮仅修改和验收 Word 端。PowerPoint 源码、PPAM、PPTM 和既有功能均保留；`VisualTeX.ppam` 的 SHA-256 仍为 `b2812aacf16650d0375a285d34cd55b2af3059a3a0e009cc8466f42e3611aba4`。

## 最终文件

- Word 加载项：`/Users/lpj/devspace/workspaces/visualtex-macos-offline/VisualTeX.dotm`
- Word 验收加载项：`/Users/lpj/devspace/workspaces/visualtex-macos-offline/artifacts/macos-offline-word-final/VisualTeX.dotm`
- macOS 验收安装包：`/Users/lpj/devspace/workspaces/visualtex-macos-offline/artifacts/macos-offline-word-final/VisualTeX_1.1.0_aarch64.dmg`
- 构建清单：`/Users/lpj/devspace/workspaces/visualtex-macos-offline/artifacts/macos-offline-word-final/BUILD_MANIFEST.json`
- 已安装 App：`/Applications/VisualTeX.app`

## 已安装路径

- Word Startup：`~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/VisualTeX.dotm`
- Word AppleScriptTask：`~/Library/Application Scripts/com.microsoft.Word/VisualTeXWord.scpt`
- VisualTeX 后台程序：`/Applications/VisualTeX.app/Contents/MacOS/visualtex --office-background`

## SHA-256

```text
VisualTeX.dotm                        81ee899450a63de1a8bea9147a39dbfacc15928685f9fcb6a869b455671e619e
VisualTeX_1.1.0_aarch64.dmg          83f45837a97af870998343898bf6d4d912ba43cb22fd17814521073b27b851e8
/Applications/VisualTeX.app executable 1946225e8202bec0475a6a9c4f79cde7565750de3404470dcb27af24617ade07
VisualTeX.ppam（本轮未改）             b2812aacf16650d0375a285d34cd55b2af3059a3a0e009cc8466f42e3611aba4
```

以下五份 DOTM 已核对为完全相同的 SHA-256：

- 根目录 `VisualTeX.dotm`
- `office/macos-offline/resources/VisualTeX.dotm`
- `artifacts/macos-offline-word-final/VisualTeX.dotm`
- Word Startup 中已安装的 `VisualTeX.dotm`
- `/Applications/VisualTeX.app` 内置的 `VisualTeX.dotm`

## 本轮解决的问题

1. **OMML 行间公式异常插入**：行间 OMML 统一先按安全的 inline OMath 转移，删除图片或占位符后重新定位 OMath，再提升为 Display，最后建立原生公式 Bookmark，避免旧 Range 吸收段落边界或产生异常内容。
2. **空白行/行首行内 OMML 居中**：Word 在删除相邻图片后会自动把空段落中的 OMath 提升为 Display。本轮在源对象删除后重新定位公式，再强制恢复 `wdOMathInline`，仅在公式前后没有正文时把段落恢复为左对齐。
3. **行首图片转 OMML 后居中**：图片转 OMML 使用相同的最终 inline 规范化流程，并增加源图片 FormattedText 回滚副本；不会破坏已有正文段落格式。
4. **失败窗口劫持后续编辑**：每个新 Word 请求均按当前公式创建独立 Session；打开新 Session 前会销毁同宿主的旧公式编辑窗口。真实连续 Session 测试中两个请求的 Session ID 不同，第二次打开后 VisualTeX 始终只有一个公式窗口。
5. **图片公式编号不进入交叉引用列表**：编号改为通过 Word 内置 Equation Caption API 注册，再保持现有中心公式、右侧编号布局。真实 Word 回归中 `GetCrossReferenceItems(wdCaptionEquation)` 返回 1 项。

## 自动化与真实宿主测试

### 完整验收

`npm run test:macos-offline-office:full` 最终结果：

```text
PASS 01-source-smoke
PASS 02-rust-regression
PASS 03-desktop-build
PASS 04-platform-boundaries
PASS 05-windows-architecture
PASS 06-office-manifests
PASS 07-platform-onboarding
PASS 08-native-office-package-smoke
VisualTeX macOS offline Office acceptance: PASS
```

完成时间：`2026-07-17T20:28:03.207Z`（日本时间 2026-07-18 05:28:03）。

### 详细结果

- IME Enter 回归：PASS；
- 分式、根式、积分、求和、上下标、矩阵和定界符结构化 OMML 回归：PASS；
- macOS Office source smoke：PASS；
- AppleScriptTask 编译：PASS；
- Rust 库测试：82 项全部 PASS；
- TypeScript 检查与 Vite 桌面构建：PASS；
- Tauri debug App 与 DMG 构建：PASS；
- 已安装 App 与构建 App 的深度签名校验：PASS；
- Word VBA 工程在 Microsoft VBE 中真实编译：PASS；
- 关闭 `VBAObjectModelIsTrusted` 后，最终安装版 Word 宿主自检：PASS；
- 真实 Word 16.89.1 原生公式回归：PASS；
- 空白行图片删除后行内 OMath 仍为 inline 且左对齐：PASS；
- 行间 OMath 提升为 Display 且居中：PASS；
- 原生公式 Bookmark 持久化：PASS；
- 图片 Equation 编号进入交叉引用列表：PASS，`crossReferenceItems=1`；
- 连续两个不同 Word Session 的窗口替换：PASS；
- Word Startup、App 内置资源、根目录和验收产物哈希一致：PASS；
- `/Applications/VisualTeX.app` 已更新，LaunchAgent 正常恢复到后台运行。

## 仍需人工视觉验收的项目

自动测试已经覆盖事务、对象类型、段落对齐、Bookmark、交叉引用注册和 Session 路由。以下项目依赖实际页面观感或真实鼠标操作，仍建议按验收清单检查：

- 多种复杂公式在正文中的基线、字号和行高是否符合个人视觉预期；
- 行间公式、矩阵、积分等高公式的上下间距和编号视觉中心；
- 解锁桌面后真实鼠标双击图片公式和 OMML 公式的手感；
- 保存文档、完全退出 Word/VisualTeX、重新打开后的页面视觉一致性；
- Word 原生“引用 → 交叉引用”对话框中的显示文字和用户操作流程。

这些项目不影响本轮已通过的自动化逻辑验收，但需要用户以最终文档页面为准进行视觉确认。
