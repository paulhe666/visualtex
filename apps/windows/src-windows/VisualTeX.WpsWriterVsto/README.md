# VisualTeX.WpsWriterVsto

WPS 文字（Writer）COM 加载项，v1 范围：

- 插入 MathType OLE（`Equation.DSMT4`）行内 / 行间公式
- 双击 / Ribbon 重新编辑所选 MathType 公式
- 公式编号（连续 / 按章 / 按节 × 点 / 横线）与更新编号

技术路线与设计依据见 `~/.commandcode/plans/wps-mathtype-ole-addin.md` 和
`apps/windows/docs/WINDOWS_NATIVE_OFFICE_OLE_ARCHITECTURE.md`。

## 构建

要求：Visual Studio 2022 / MSBuild + NuGet 还原。**不需要安装 Office 或 WPS**：
Word PIA 来自 NuGet 包，`IDTExtensibility2` 与 `IRibbonExtensibility` 定义在
`OfficeInteropShims.cs`，项目中没有 `COMReference`。

GitHub Actions 也会自动构建：`.github/workflows/wps-addin.yml` 在
`windows-latest` 上分别构建 x64/x86，并上传完整输出目录作为 artifact。

```powershell
msbuild apps\windows\src-windows\VisualTeX.WpsWriterVsto\VisualTeX.WpsWriterVsto.csproj `
    /p:Configuration=Release /p:Platform=x64 /restore
```

WPS 与 Office 位数必须匹配（32 位 WPS 用 `Platform=x86`）。输出目录中的**全部文件**
都要一起部署，包括：

- `VisualTeX.WpsWriterVsto.dll`
- `VisualTeX.WindowsOffice.Contracts.dll`
- `System.Text.Json.dll` 及其依赖链
- `Microsoft.Office.Interop.Word.dll`

`VstoDependencyResolver` 从加载项目录解析上述程序集。

## 注册

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
    -File apps\windows\scripts\install_wps_addin.ps1 `
    -AssemblyPath <输出目录>\VisualTeX.WpsWriterVsto.dll
```

注册内容：

- `HKCU\Software\Classes\VisualTeX.WpsWriterVsto` → CLSID `{5D6E7F80-91A2-4B3C-8D4E-5F60718293A4}`
- `HKCU\Software\Classes\CLSID\{5D6E7F80-...}\InprocServer32`（`mscoree.dll`，AnyCPU/IL）
- `HKCU\Software\Microsoft\Office\Word\Addins\VisualTeX.WpsWriterVsto`（`LoadBehavior=3`）
- `HKCU\Software\kingsoft\office\WPS\AddinsWL\VisualTeX.WpsWriterVsto`（WPS 白名单）

32/64 位注册表视图都会写入。卸载：加 `-Unregister`。

## 先跑兼容性探针

在依赖任何降级路径之前，先运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
    -File apps\windows\scripts\wps_compat_probe.ps1
```

默认不执行 `Equation.DSMT4` 的 `AddOLEObject` 测试（OLE 服务器可能弹出编辑器窗口
并阻塞）。确认可以交互关闭 MathType 窗口后，再加 `-IncludeOleInsert` 运行完整探针。

报告写入 `%LOCALAPPDATA%\VisualTeX\office\logs\wps-compat-probe.json`。重点看：

- `range.insertXml`：决定插入主路径是否可用
- `inlineshapes.addOleObject.classType` / `inlineshapes.oleFormat.object`
- `event.windowBeforeDoubleClick.subscription`：决定双击路径
- `fields.seqCreateAndUpdate`：决定编号字段是否可用
- `saveReopen.roundTrip`：决定保存重开后能否识别

探针失败时的降级顺序已经实现在 `WpsWriterFormulaService`：

1. `Range.InsertXML(Flat OPC)`（主路径）
2. 剪贴板 `BeginStandaloneClipboardTransaction` + `Selection.Paste()`（备用）

## 已知限制（v1）

- 编辑写回采用"清除原对象 + 原位重新插入"的策略，不依赖 `OLEFormat.Object`
  和运行中的 MathType 服务器。这样在未安装 MathType 的机器上也能编辑。
- 编号使用标准 `SEQ`/`STYLEREF` 字段（WPS/Word 都支持），尚未采用 MathType
  原生 `MTPlaceRef`/`MTDisplayEquation` 结构。跨端打开时 MathType 对象本身可识别，
  编号字段由 Word/WPS 各自维护。
- 双击编辑依赖 `WindowBeforeDoubleClick` COM 事件；若 WPS 未触发该事件，
  下一步需要移植 `WordDoubleClickHook`（`WH_MOUSE_LL` + 选区命中检测）。
- WPS 演示（Presentation）加载项尚未创建，后续复用 `PowerPointFormulaService.MathType`
  的 staging-package 流程。

## 日志

`%LOCALAPPDATA%\VisualTeX\office\logs\wps-writer-vsto.log`
