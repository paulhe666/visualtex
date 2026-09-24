# 当前修改的 NSIS 构建交付

2026-09-06，按用户最后指令只构建当前修改的 NSIS，检查完成后结束本轮。未执行安装、安装后 Word 验收、提交或发布；原始八份未保存 Word 文档未操作。

- 安装包：`apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.6_x64-setup.exe`
- 文件大小：323777910 字节（约 308.78 MiB）
- SHA256：`A324ACB61252C15D9EC06FE118DA76A17339ABA6B30B94447421879413181F19`
- 源码：HEAD `3d207d7a81ff40014fd861b242cff96ea77060a2` 加当前未提交修改；版本 1.2.6，Office MSI 1.0.43.0。源码文件哈希快照在 `evidence/stage06-package-verification.json`。

构建和检查均成功：

1. `npm run build:all`：桌面和 Office 前端构建。
2. `build-office-candidate.ps1 -Label stage06`：项目 x64/x86 VSTO、OLE、MSI 构建及 Ribbon/依赖加载检查。
3. `tauri_build.mjs --prepare-nsis-template-only`、前端产物校验、`tauri build --no-bundle`：主程序 Release 构建。
4. `tauri bundle --bundles nsis`：新 NSIS 生成完成。
5. 项目 `verify_windows_release_artifacts.ps1` 静态检查通过；构建主程序和解包主程序均通过当前嵌入前端检查。
6. 使用 7-Zip 只解包读取：37 项一致性检查全部通过，包含 Office 前端、桥接程序、x64/x86 MSI，以及从这两个包内 MSI 取得的 Word VSTO、PowerPoint VSTO、共享契约 DLL 和 OLE 服务器。它们与本次构建输出的 SHA256 一致。

Tauri 在 NSIS 中把主程序唯一的 `__TAURI_BUNDLE_TYPE_VAR_UNK` 改为 `__TAURI_BUNDLE_TYPE_VAR_NSS`，与裸构建文件仅这三个字节不同；按该明确标记比较后，其余所有字节一致。包内主程序 SHA256 为 `0ABA884B4D02EAB5F34C9B3BF2C9618673F95FD37F8C8AC4668B7CE5198D9A4C`，裸构建文件为 `06B3B9AD5951EC4BD2AAAF6673AC48D2BDFA4666F15549E4DF3283DEDBB4A476`。

构建没有错误。前端保留资源体积提示，Rust 保留 62 条未使用代码等编译警告，未在本轮扩展修改。

主要记录：

- `evidence/stage06-frontend-build.log`
- `evidence/stage06-office-build-desktop.log`
- `evidence/stage06-main-build.log`
- `evidence/stage06-nsis-build.log`
- `evidence/stage06-release-verify.log`
- `evidence/stage06-packed-frontend-verify.log`
- `evidence/stage06-package-verification.json`
- `evidence/stage06-nsis-contents.log`

以上证明编译和包内产物一致性，不代表安装后完整产品矩阵通过。其余验收已按用户要求暂停，当前状态见 `CURRENT_PROGRESS.md`。
