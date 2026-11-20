# VisualTeX for Windows

本目录只包含 VisualTeX 的 Windows 应用。

## Office 集成

Windows Office 集成统一使用原生 Ribbon COM 加载项与 OLE LocalServer：

- Word 和 PowerPoint 使用 `VisualTeX.WordVsto`、`VisualTeX.PowerPointVsto` 原生 Ribbon 与 Office 事件；
- 专业模式插入 `VisualTeX.Formula.1` OLE 对象；
- Word 支持 OLE/OMML 行内与行间公式、转换、编号和引用；
- PowerPoint 支持 OLE 公式的新建、编辑、删除与图片导出；
- Office 原生双击可重新打开 VisualTeX 编辑器；
- 不再安装或注册 Office.js XML 清单、WEF Trusted Catalog 或 `%LOCALAPPDATA%\VisualTeX\OfficeCatalog`。

安装流程先准备当前用户证书，再安装与 Office 位数匹配的 MSI。文件、注册表和 OLE LocalServer 验证通过后，仍需运行 `scripts/test_windows_office_runtime.ps1`，只有 Word 与 PowerPoint 的 `COMAddIn.Connect=True` 才算完整验证成功。

## 开发

```bash
npm ci
npm run build:desktop
npm run build:office:windows-native
npm run test:platform-onboarding
npm run test:windows-office-architecture
```

Windows 原生源码位于 `src-windows/`。

---

This directory contains only the VisualTeX Windows application. Windows Office integration uses native Word/PowerPoint Ribbon COM add-ins and the `VisualTeX.Formula.1` OLE LocalServer. Office.js manifests, WEF Trusted Catalog registration, and OfficeCatalog installation are retired. Runtime verification is complete only after both Word and PowerPoint report `COMAddIn.Connect=True`.
