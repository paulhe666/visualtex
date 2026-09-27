# PowerPoint MathType OLE 与快捷键页面收尾

## 本次修改

- 删除快捷键页面及设置入口的常驻解释文案；保留标题、动作名称、必要的输入错误反馈。
- PowerPoint Ribbon 新增“新建 MathType”（中英文），选中已有 MathType 后可用“编辑所选公式”进入 VisualTeX。
- PowerPoint 编辑器增加 SVG、VisualTeX OLE、MathType OLE、OMML 格式选项。
- MathType MTEF 编解码、CFB 存储、剪贴板协议、隔离预览改为 Word/PowerPoint 共用源文件；Word 的 COM、段落、行内和编号逻辑留在 Word 工程。
- PowerPoint 原生 MathType 按独立幻灯片对象处理，不引入 Word 的行内/行间布局逻辑。
- 从 PPTX 嵌入存储读取真实源码，不依赖 VisualTeX Tags，不通过激活 MathType 编辑器获取源码。
- 使用原生 MathType 矢量预览。修复 WMF placeable header 的逻辑坐标在转 EMF 时丢失导致空白的问题；不以猜测的图片尺寸冒充原生预览，预览失败不替换源对象。
- 精确绑定演示文稿、SlideID、ShapeID；编辑后保持位置和独立横纵缩放，迁移动画、触发器、连接线和点击动作。先验证新对象的嵌入源码，再删除旧对象。
- 原生 MathType 双击仍交由 MathType；VisualTeX 使用明确的 Ribbon 编辑入口。
- 未修改 MathType 公共注册、语言资源或用户已有演示文稿。

## 定向验证

实际 PowerPoint（x64）隔离夹具测试通过：

1. 新建 Equation.DSMT4 原生 OLE，读取 `x=1`。
2. 编辑为分式，保留位置、标签、超链接与动画目标/时长。
3. 保存 PPTX、关闭测试文稿、重新打开并读取公式；导出 PNG，检查不是空白。
4. 再编辑回原式，横纵独立缩放不累计漂移。
5. 写入前演示文稿身份变化时拒绝操作，原对象未删除。
6. 从已有真实数学试卷的 MathType 嵌入对象读取 `\forall x>1`，在 PowerPoint 编辑为根式。
7. 独立新建三次根号，导出幻灯片并人工查看图像，根指数、根号和被开方数正常。

证据：`docs/remediation-3d207d7/evidence/ppt-mathtype-20260927/`。
运行入口：`VisualTeX.VstoFlowAcceptance.exe --mode=powerpoint-mathtype <artifact-dir>`；第三方夹具路径由 `VISUALTEX_ACCEPTANCE_MATHTYPE_SOURCE_DOCX` 提供，只读使用，不提交该用户文档。

其他通过项：TypeScript 类型检查、Office 快捷键定向测试、PowerPoint 双击安全测试、MathType 运行库隔离与兼容性源检查、Rust 的 PowerPoint MathType 关闭提交条件测试。

没有运行全量旧回归。现有 WindowsOffice.Tests 工程有两个既有编译问题（缺少 WordFormulaHostDescriptor、重复的 ComputeVerifiedMaterializedOmmlFingerprint）；改用已有独立 codec 夹具检查字号、重音与颜色，共 23 项中 22 项通过。未通过项是既有三次根号“必须显式定义黑色调色板”的断言；本次 codec 与 HEAD 比较仅移动位置、替换日志入口和开放原有 MTEF 提取方法，未改变编解码算法。上述真实 PowerPoint 三次根号显示验收通过，没有为了该旧断言加入额外补丁。

## 范围与验收入口

- 本次支持独立、嵌入式 MathType 对象；不承诺链接对象、组合内部对象或任意历史 MTEF 版本。不能安全解析时失败退出，不猜测替换内容。
- PowerPoint 本身不允许旋转/翻转 MathType OLE；转换旋转或翻转的 SVG 时明确拒绝，原图保留。
- 真实 Office 验证使用代码通过 COM 创建并关闭测试文稿。原生窗口截图接口不可用，未声称完成安装后 Ribbon 的人工逐项点击验收。
- 安装包构建不自动安装；关闭 Office 后安装新包，再在 PowerPoint 的 VisualTeX 选项卡使用“新建 MathType”或“编辑所选公式”验收。

## 安装包

`apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.8_x64-setup.exe`

- 大小：347,983,769 字节。
- SHA-256：`5F76EE04F0C0FECC437F1DE48BB887279A082BA83534642A15EE37984097E0A1`。
- 包含本次重编的 x64/x86 Office 插件和完整内置运行库。复用了原构建流程，没有把其它已有源文件改动混入本次 Git 提交。

## 原生行为依据

- [MathType 官方 Office 文档](https://docs.wiris.com/mathtype-7-with-microsoft-office-2016-or-later?kb_language=en_US)：PowerPoint MathType 是浮动对象，而非 Word 行内公式。
- [Microsoft PowerPoint Shapes.AddOLEObject](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addoleobject)：PowerPoint 的嵌入 OLE 容器模型。
- [Microsoft Effect.Shape](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.effect.shape)：迁移动画目标时保留现有时间线。
