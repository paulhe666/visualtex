# MathType LaTeX 重绘：MTEF 误报与固定范数符号修复

## 工作基线与范围

- 工作区：C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7
- 分支：reference-3d207d7；HEAD：632d92b4f3133fe6a2d64fbab1579abf356e4e34，未提交、未推送。
- 原活动 Word：PID 68232，C:\Users\pojian_liao\Downloads\latex_100_inline_100_display.docx。
- 原文档为 303 段，100 个行内和 100 个行间 LaTeX 源码公式，OMaths=0、InlineShapes=0、Saved=True。仅通过已有只读 COM 连接读取，没有执行重绘、写入、保存或关闭原文档。
- 产品代码仅修改 apps/windows/src-windows/VisualTeX.WordVsto/MathTypeMtefCodec.cs；另增回归测试与取证脚本。未改动其他历史 dirty/untracked。

## 已确认根因

### 1. 固定定界符的校验规则与生成规则不一致

截图对应第 15 个公式（第 8 个行内公式）：`\lvert x_{8}-\mu\rvert<\varepsilon`。真实导出的 MathML 在竖线 mo 上同时具有 data-mjx-texclass=OPEN/CLOSE、fence=false、stretchy=false。

旧校验根据 OPEN/CLOSE 生成 `fence(|,|,...)`，忽略固定字符属性；实际 MTEF 正确保存为两个 CHAR，读回为 `o(|)...o(|)`。MathJax 的间距分类不会持久化到 MTEF。因而 expectedLen=50、actualLen=47 是内部签名长度差异，不是少了三个公式字符。

同类不一致还涉及普通括号、连乘号主槽带来的 mrow 分组，以及可伸缩括号内的普通括号。旧路径中有的对子节点逐个拼接，有的对整个子序列做成对识别，导致同一数学内容的签名随容器边界变化。

修复：固定属性优先于 MathJax 间距分类；生成端整行识别和签名端散列序列识别共用 IsExplicitlyFixedFence；定界符内部统一使用 CanonicalizeElementSequence。固定字符保持 CHAR，可伸缩定界符仍按其模板结构处理。

### 2. 帽号、向量箭头的默认 accent 属性被误判

输入 MathML 的 mover 可以省略 accent，核心 mo 的算符字典默认值决定其重音性质。旧签名代码将省略直接视作 false，产生 `sup(...)`；MTEF 的实际重音读回带 accent=true，产生 `accent(...)`。

修复：ResolveOverAccent 按父 mover 显式属性、核心 mo 显式属性、受支持算符的默认重音属性依次解析。显式 false 优先，普通 msup、缺失重音、改变符号等仍保持可区分。

规范依据：MathML 3 §3.4.5，https://www.w3.org/TR/2023/SPSD-MathML-20230307/chap3_4.html 。此处实现的是当前 codec 支持的重音算符集合，不声称实现整个 MathML 算符字典。

### 3. 扩展测试发现真实生成错误：固定范数双竖线

`\lVert x\rVert<1` 的 U+2016 曾进入 Unicode Text 回退。自解码结果可保留字符，但 MathType 7 原生输出实际是缺字符号方块，不能靠放宽签名比较处理。

已改为 MT Extra 固定双线字形（encoded8=0x50），保留 MTCode U+2016，与并行关系 U+2225 区分。读回该 MT Extra 字符为数学 mo。原生渲染图已确认从缺字方块恢复为双竖线。没有把任意 mtext 双竖线一概归一成数学符号；测试仍区分错误 Text 回退。

## 结果与证据

取证根目录：docs/remediation-3d207d7/evidence/mtef-redraw-20260923/

- 原文档修复前：200 个公式中 180 个预检通过、20 个被拦截。五类各四次：绝对值、连乘括号、分部积分括号、帽号、向量箭头。
- 修复后：原文档 200/200 通过正式 PrepareStandaloneMathTypeOleData，包括独立复合存储创建及读回。
- 原文档全部 200 个公式生成的 Equation Native 字节哈希修复前后完全一致。结合原生图像核对，说明此次原文档的拦截是校验误报，并非靠改变这 200 个公式的 MTEF 内容使测试通过。
- 扩展语料共 252 个，其中 250 个成功导出 MathML；这 250 个 MTEF 预检全部通过。余下 2 个是同一 `\bra{x}A\ket{y}` 的行内/行间前端导出错误，发生于进入 MTEF 之前，保留在 final/results.json 中，未删除或伪装为通过。本轮未改此前已 dirty 的前端兼容文件。
- 独立调用已安装 MathType 7/MathPage 的 sidecar：250/250 原生渲染请求成功，native-final/response.json 与 render-results.json 保存完整结果。实际查看了绝对值、连乘、分部积分、帽号、向量及范数的原生图像。请求成功不等于对 250 张图逐张人工完成了版面验收。
- 数据层定向回归：128/128，覆盖固定/可伸缩定界符、行内容器、矩阵单元格、上下标、创建/重写/读回再编辑、重音默认值、缺失/重复符号、分子分母颠倒、运算符及指数改变、字体和 MTEF 字体编码。证据：mtef-redraw-tests/TestResults/final.trx。
- 直接运行原 VisualTeX.WindowsOffice.Tests.csproj 时，在修改产品代码前已经存在 CS0246（WordFormulaHostDescriptor）和 CS0111（ComputeVerifiedMaterializedOmmlFingerprint 重复定义）编译错误。因此本轮用 mtef-redraw-tests/MathTypeRedraw.Tests.csproj 链接同一真实产品 codec 和已有 MathType 测试，不声称整个历史测试项目通过。
- 原文档正文哈希在恢复任务前后保持 A0EE1ABC85BF0A22CDCB08288A61FB0DF9D207374793DEFA1D7D658B4AD618A3，公式/段落计数及 Saved=True 保持不变。WordOpenXML 导出哈希在两次读取间不同，未据此声称完整 XML 字节恒等；未对原文档执行写入。

## 构建与当前安装状态

- Word VSTO x64、x86 Release 编译均通过（有历史 warning）。
- x64 DLL SHA256：A6E796FB8E468F01A95AD1FDE2FC7D71BC42CC71B8AE91E1C74C8D71B0209B2D
- x86 DLL SHA256：351623B804C258C0D35559411E7784AD836AF7BBA5E5090937C22092518FA2B9
- 正式 x64 DLL 反射：HAS_COMPARER=False；HAS_VALIDATE_OMML_RESULT=False；HAS_VALIDATE_MATERIALIZED_OMML=False；HAS_MTEF_PREFLIGHT=True。
- 未恢复 OMML runtime semantic validation，未修改纯 Word OMML 架构。
- 当前已安装 Word DLL 仍是旧 SHA256 6C11376E9DB243888F4E47ABA4981CD0D9DFAA7293FC2C00B566753BFF28B768。
- 本轮未打包新的 NSIS、未安装或替换用户运行中的插件；未做新安装版 Word Ribbon 全文重绘验收。下一步应在用户要求打包后使用新 NSIS 安装，再通过真实 Ribbon 在文档副本验收，不能把本轮数据层和原生渲染测试写成安装版 UI 已通过。

## 可复用入口

- capture-mtef-redraw-source.ps1：带 PID/文档名守卫的只读活动源码取证。
- apps/windows/scripts/mathtype_redraw_fence_corpus.mts：原文档与定界符扩展语料，保留前端失败。
- probe-mtef-redraw-corpus.ps1：反射调用真实产品 DLL 的生成及独立 OLE 预检。
- probe-mtef-native-render.ps1：直接调用原生 MathType sidecar 并保留响应和图像；-AllExportable 运行全部可导出案例。
- mtef-redraw-tests/MathTypeRedraw.Tests.csproj：隔离于历史 OMML 测试项目编译问题的数据层回归。
