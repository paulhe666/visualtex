# VisualTeX for Windows

本目录只包含 VisualTeX 的 Windows 应用。

## Office 集成

Windows 主要使用原生 VSTO、COM 与 OLE：

- Word 和 PowerPoint 使用原生 VSTO Ribbon 与 Office 事件；
- 专业模式插入 `VisualTeX.Formula.1` OLE 对象；
- Word 支持 OLE/OMML 行内与行间公式、转换、编号和引用；
- PowerPoint 支持 OLE 公式的新建、编辑、删除与图片导出；
- Office 原生双击可重新打开 VisualTeX 编辑器。

`office/windows/ole/` 中保留的 Office.js Manifest 和桥接页面仅用于当前仍受支持的 Windows OLE 兼容/清理路径。构建时从依赖生成 Office.js 静态资源，仓库不再提交通用 vendor 副本或任何 macOS Office.js 文件。

## 开发

```bash
npm ci
npm run build:desktop
npm run test:platform-onboarding
npm run test:windows-office-architecture
npm run build:office:windows-ole
```

Windows 原生源码位于 `src-windows/`。

---

This directory contains only the VisualTeX Windows application. Its primary Office integration uses VSTO, COM, and real OLE objects. The Windows-only Office.js manifests and bridge remain solely for the currently supported OLE compatibility and cleanup route; generated Office.js assets are not committed, and no macOS Office.js implementation is included.
