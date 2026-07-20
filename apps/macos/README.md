# VisualTeX for macOS

本目录只包含 VisualTeX 的 macOS 应用。

## Office 集成

macOS 使用完全离线的原生 Office 路线：

- Word：`VisualTeX.dotm` 全局模板；
- PowerPoint：`VisualTeX.ppam` 加载项；
- 通信：VBA、AppleScriptTask、Office Group Container 与 Tauri 本地 Session；
- Word 支持图片公式、原生 OMML、编号、交叉引用与双击编辑；
- PowerPoint 支持公式新建、替换、删除与双击编辑。

本应用不包含 Office.js、Office XML Manifest、可信目录或本地 HTTPS 证书安装流程。

## 开发

```bash
npm ci
npm run build:desktop
npm run test:ime-enter
npm run test:word-omml
npm run test:platform-onboarding
cargo test --manifest-path src-tauri/Cargo.toml --lib
```

原生加载项源码和已验收资源位于 `office/macos-offline/`。

---

This directory contains only the VisualTeX macOS application.

Its Office integration uses a native offline Word DOTM and PowerPoint PPAM workflow with VBA, AppleScriptTask, Office Group Container files, and local Tauri Sessions. It does not ship Office.js, Office XML manifests, a trusted catalog, or a local HTTPS certificate installer.
