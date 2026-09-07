# 原生 MathType 编号字体对照

2026-09-06，在本机真实 Word 16.0.14334 中新建文档111，使用 **MathType 自己的 Ribbon**，并非 VisualTeX 按钮或直接调用宏。COM 仅准备正文和只读取证；公式通过真正的 MathType 编辑器输入。

| 操作 | 当前正文 | 原生编号结果 |
| --- | --- | --- |
| 新文档首次“右编号”，默认章/节初始化，输入 a+b=c | 西文 Arial，中文宋体，14磅 | (1.1)，Arial 14磅 |
| 同文档下一段改字体，再“左编号”，输入 x=y | 西文 Times New Roman，中文宋体，12磅 | (1.2)，仍为 Arial 14磅 |

COM 与 WordOpenXML 交叉确认：第一次创建的 MTDisplayEquation 样式以 Normal 为基准、后续段落为 Normal，样式本身保存 ascii/hAnsi=Arial、eastAsia=宋体、sz=28。原生编号段落引用该样式，编号 run 没有逐条覆盖成当前正文的字体。第二条复用原样式，保留 Arial 14磅；并非每个编号都重新跟随当前正文，也不是固定使用 Cambria Math 或某个恒定字号。

实现要求因此调整为：缺失样式时按原生方式从首次插入的正文上下文初始化；已有 MTDisplayEquation 时尊重它，不因后续正文不同而重写编号字体。OMML/VisualTeX 自有编号仍采用它们的正文继承规则。没有把 Arial 或 Times New Roman 硬编码进产品。

证据：

- `evidence/stage04h-native-mt-source-__111.json/xml`：新空文档准备后的真实正文。
- `evidence/stage04h-native-mt-right-__111.json/xml`、`-font.json`、同名 PNG：首次右编号，1 OLE、原生 MTPlaceRef/SEQ，截图已查看。
- `evidence/stage04h-native-mt-left-complete-__111.json/xml`、`-font.json`、同名 PNG：左/右两公式实际完成、2 OLE，两组编号均 Arial14，第二段正文仍 Times New Roman12，截图已查看。
- `evidence/stage04h-native-mt-style-comparison.json`：原生创建的样式含字体/字号；VisualTeX 04h创建的样式缺少相应 rPr，解释为何后者落为默认10.5。

自动化准备事故单独记录，不作为产品失败：第一次直接调用准备脚本漏传PID，30字符测试正文与回车输入此前基线测试文档71；该文档不属于八份原始文档，但不能再称失败现场未改动。其旧完整证据保留。脚本现在省略PID时也只读取已记录测试进程。第二条原生编辑时焦点未切换成功，x=y误输入文档111的后续正文，Alt+F4引发保存提示，已点击取消，没有保存或关闭；随后验证前台窗口后完成真正的MathType公式，额外正文 X=y 留在证据中，未以删除掩盖。键盘操作已增加实际前台句柄校验。

八份原始文档（1/22/26/28/29/33/35/36）在 `stage04h-original-preservation-comparison.json` 中与04a完整COM字段均相等，Saved=false。

04i代码已按上述样式初始化/复用规则调整并构建安装。新文档113经真实VisualTeX按钮重跑同样两种字体：首个右编号Arial14，下一段正文Times New Roman12，第二个左编号仍Arial14；共2OLE/2MTPlaceRef、实际编号(1)/(2)，正文各自格式保留，stage04i-mt-right/left的COM/XML/-font.json和已查看截图证明。辅助测试379通过。引用往返、完整矩阵和最终同源NSIS尚未完成。

04m文档120进一步验证续写和引用：两次实际插入后的空白正文段落分别继承Arial14、Times New Roman12，使用Normal样式。MathType编号仍复用首次建立的Arial14；第二式新增引用的域代码、结果及括号均继承插入位置的Times New Roman12。编号改为(0.0-1)/(0.0-2)后，引用同步(0.0-2)，两类字体各自保持，实际双击正确高亮第二MTPlaceRef。见stage04m-first/second/reference/renumber的COM/XML、-font.json及已查看PNG，stage04m-renumber-jump记录真实跳转。由此确认遵循原生MathType编号样式，同时引用遵循插入位置正文格式；并未把引用强制改成编号字体。
