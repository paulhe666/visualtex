# NSIS 验收包（2026-10-07）

## 当前包：CP325 MathType 颜色修复验收包

[VisualTeX_1.2.8_x64_CP325_MathTypeColor_Acceptance_20261007.exe](../artifacts/nsis-acceptance-cp325-20261007/VisualTeX_1.2.8_x64_CP325_MathTypeColor_Acceptance_20261007.exe)，377000590字节，SHA256 **A2908B2629FFA055CA3A267DFCD69274027C000E8A527718ED966E2CDA19C99F**。版本1.2.8，Office MSI ProductVersion 1.0.43.0，包含x64/x86 Office组件和标准离线OCR运行时。

现有CP324包完成于20:45，颜色修复源码在21:09–21:16才写入，所以CP324没有作为本次验收包交付。CP325锁定当前颜色源码、HEAD `dc501c9c6fe607b1aef0ef1404981b62d1668027`、冻结R3 `FreezeR3.targets` 及heldR4；构建前后均核对输入哈希。没有用默认工作树构建替代冻结R3，也没有覆盖旧包。

## 构建与实际载荷

x64/x86 Formula OLE、Word/PPT VSTO、Ribbon调度、离线依赖加载、两个Office MSI、Office/桌面前端、Rust Release、NSIS和静态发布检查全部通过。编译保留既有可空性、Office互操作和Rust未使用代码警告，没有编译错误或错误豁免。

实际解出NSIS和三个MSI，没有执行安装程序。两架构Office各19个文件、Office前端23个文件、OCR74个文件、NativeRender完整清单、Bridge/VSTO运行时、桌面入口/两个JS-CSS资源/19字体均逐项核对。内嵌Companion `38EC7E05C0225C5A5E20BB965C0316B5D936D188BAE0C356053E8F3DC5C6AB98` 与构建输出 `C3D4048F25689DA0F25BF5C33471E6F74FE26CA68828D3A5C06D792CF0A04042` 只存在Tauri预期的唯一3字节 `UNK`→`NSS` 标记差异，其余字节严格相等。最终报告状态为 `VerifiedEmbeddedPayloadNotInstalled`。

| 实际内嵌组件 | SHA256 |
| --- | --- |
| Word x64，冻结R3 | `353C20B4B1EFA31E3A53FB1E51E392F6970DAF5DEFDE9912A8FE4251BE58A18C` |
| PowerPoint x64 | `C8E90D01C88820DC5AFC03586A60B45DBF5B95CBF12476157D89228957A38798` |
| Word x86，冻结R3 | `D95B24CA60B308A6163D14A211AA3D8BD4F9607C44292881B1187CBC2AB8C1C5` |
| PowerPoint x86 | `C97E814DA421558F84142CC0942958FE7ABC7D17817E27F7842D5F7AB271E0EC` |
| Office MSI x64 | `58648D317E0C5D4F8F12AE06D55598FCFFCA9C851BF94DBD3986AB3A925DBAD3` |
| Office MSI x86 | `85C2390455DA9364471E5E47F63865B7C80E8B0733D901B43FE72D4DE722B1D2` |
| NativeRender MSI | `D2CBBBD985BD82B734D396E33A4F37549D6D3B64DACD74EDF0A7321E077221EA` |
| Office Bridge | `10C0142ECB31C2EC5BA8EE3F215AF479BCE1B8268E5488E778F0C17B4AA4F37C` |

x64 Word DLL与CP323真实自有Word颜色写回/保存重开验收所用候选逐字节一致。这证明打包载荷包含已测颜色修复，不把打包完整性等同用户安装后的功能验收。

## 环境与证据

`environment-final.json`确认八处Office LoadBehavior、当前已安装Word/PPT DLL、构建前已有受保护进程身份均未变化，新增受保护进程为0。未安装、未改注册表或安全设置、未操作用户文档、未提交或推送。

证据目录：

- `artifacts/nsis-acceptance-cp325-20261007/build/build-state.json`：完整构建状态和组件哈希。
- `artifacts/nsis-acceptance-cp325-20261007/verified-embedded/report.json`：实际解包与逐文件核验。
- `artifacts/nsis-acceptance-cp325-20261007/environment-final.json`：未安装环境复核。
- `artifacts/nsis-acceptance-cp325-20261007/verified-inputs.json`：当前颜色源码和冻结输入清单。

专项颜色11/11及自有Word真实OLE验收结果见 `docs/mathtype-color-edit-stage-20261007.md`。完整测试仍保留一个与本轮无关的既有CP255宽重音契约失败；NSIS构建没有改变该结论，也不代表所有历史保真问题均已关闭。
