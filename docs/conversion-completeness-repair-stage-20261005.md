# 完整转换缺口修复（2026-10-05）

## CP284：重新建立根因边界

用户要求继续解决完整转换问题。当前 HEAD 为 a57e568d5295dc67490e30f4169cd5d05c7ce7c5；CP283 性能改动和其他原有未提交文件保留。本轮尚无生产修复、构建或新 Office 验收，不沿用旧 19/20 结论。当前原 20 条的 F07/F08/F09 MathType 围栏结构缺口、两个 OMML 目标 F02/F05 显式斜体缺口仍成立。

重新阅读 MTEF v5 官方模板规范和本机 MathType 的 Presentation MathML input.tdl：标准 ParBox 模板不提供通用任意 Unicode 伸缩围栏；现有特殊围栏字体槽位不足以证明伸缩结构。F07–F09 不会通过借用普通括号选择器、修改比较器或隐藏元数据宣称修复。单边上取整追加右 CHAR 的旧实验仍失败，原证据不覆盖。

新发现的实际共享缺陷：当前 codec 没有实现官方 selector 9（tmINTERVAL），使原生支持的混合/反向区间括号也降为普通字符；EmitFenced 直接串接 children，不写 MathML 默认逗号或显式 separators。修复前先建立独立原始记录/分隔符断言，并保留失败日志。显式 MathJax 行转 mfenced 同时必须声明空 separators，否则实现正确分隔符规则后会误插逗号。

Word 样式边界：微软 OfficeMath 文档说明数学模式中的 bold/italic 只作用于数学字母数字，普通标点/运算符保持直立；普通文本模式具有不同语义。Literal 文档也明确改变二元运算间距。结合 CP222 的标准 m:sty/w:i 真实插入及重开失败，本轮不采用角色替换作为修复；这不是对所有可能表示方式的不可能性证明。

来源：
- https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later
- https://devblogs.microsoft.com/math-in-office/officemath/
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.literal?view=openxml-3.0.1

下一步：补齐 tmINTERVAL 的共享读写和 mfenced 操作数展开，独立原始 MTEF 测试、受影响回归、冻结 R3 候选构建，再运行小量自有原生服务/Word 验证。F07–F09、F02/F05、CP255 及全量视觉保真保持未解决；不把新功能定向通过扩大为六向完整验收。

## CP285：共享修复及第一条真实 Word 方向通过

生产代码仅改 MathTypeMtefCodec.cs：新增官方 tmINTERVAL 选择器及两端独立 variation 编解码；mfenced 在继承字体规则之前展开为含真实分隔符的单操作数；MathJax 显式围栏归一化声明空 separators，避免引入默认逗号。没有改比较器、既有测试预期或格式身份。新增 MathTypeIntervalAndSeparatorTests.cs 独立构造全部16种官方区间记录，不依赖编码器生成读回输入，并核对编码记录、分隔符字面量、嵌套上下标和分隔符继承样式。

复现37项中33失败/4通过，修复后37/37；追加嵌套/继承2项后最终39/39。相关回归948项报1个旧CP255字节预期失败，其中4个native方法因未启用仅提前返回，不计实际原生通过。正式聚合明确排除这4项：2408通过/1失败/0跳过；初次沙箱运行另有2个受控临时SVG写权限失败，正常权限重跑后消失，全部原日志保留。未更新CP255预期。

完整冻结R3候选build-cp284-interval-separators-r1成功，Word SHA256 0DE80AE7800F1B2042AEB934CEA9F78637EC8644BBF616DFB53885F294B36E6C，heldR4不变。原20每项MTEF和OMML与CP282字节哈希完全一致；MTEF仍17/20，F07/F08/F09原失败未变。这是隔离修复影响的证据，不是解决原20缺口。

实际新建自有Companion PID64148/43136（前轮固定EXE 3C0B358B…DA6B90、前轮固定Office UI），只读身份/正常TLS通过；现有3个Word、4个Companion及4处LoadBehavior有前快照。首个6条小矩阵VT→MT：6转换、6保存重开严格通过，sourceDifferences为空，实际原生模板已被生产路径排版并写入OLE。控制包含混合区间、反向区间、分数/上下标和标点；不是原20或像素验收，其余5方向随后验证。

离线工具边界：原20及显式围栏6均正常完成。另一个以标准隐式mfenced替代I5/I6的离线外部进程在I5进入OMML投影前后异常终止（只落盘I1–I4，终端“Exception.ToString()失败”）；目前不能将其归因于本轮codec，也不算通过。该输入、部分输出保留在corpus-fence6-mfenced.json / fence6-pure，未扩大此前用户搁置的外部运行器栈问题。独立MTEF分隔符规则已有真实字符断言，Word小矩阵使用实际生产MathJax显式源，不能偷换成隐式mfenced端到端通过。

证据根目录artifacts/conversion-completeness-20261005；构建在artifacts/omml-edit-audit-20261001。所有编译路径错误、沙箱错误、旧失败TRX保留。本轮无提交/部署/注册表写入。

## CP286：六条控制六向闭环完成；保留完整转换边界

同一CP284候选、同一固定前端/Companion完成六向，每向6条转换及6条保存重开检查，共36次转换、72项阶段比较；全部sourceDifferences为空。VT源由实际MathJax/Companion生成，MT源来自本轮VT→MT保存结果，OMML源来自本轮VT→OMML保存结果。本次没有Undo/Redo、物理编辑操作或实际Word截图验收。

| 方向 | 转换 | 保存重开 | 证据子目录 |
|---|---:|---:|---|
| VisualTeX → MathType | 6/6 | 6/6 | vt-mt-r1 |
| VisualTeX → OMML | 6/6 | 6/6 | vt-omml-r1 |
| MathType → VisualTeX | 6/6 | 6/6 | mt-vt-r1 |
| MathType → OMML | 6/6 | 6/6 | mt-omml-r1 |
| OMML → VisualTeX | 6/6 | 6/6 | omml-vt-r1 |
| OMML → MathType | 6/6 | 6/6 | omml-mt-r1 |

从保存DOCX提取并实际检查6张原生WMF缓存预览，半开/反向括号身份和分数外侧伸缩方向正确，逗号/分号可见；这是离屏原生缓存检查，不是Word窗口视觉验收，也没有证明像素/所有字号几何完全相同。implicit-separator2另将两个真实隐式mfenced直接送入新MTEF编码器及既有核验原生服务，2/2排版完成，实际查看PNG确认默认逗号和显式分号存在。该路径不调用异常的OMML外部探针，不能算作隐式mfenced的OMML端到端通过。

原20未重跑完整Word矩阵，因为本轮该20项MTEF/OMML全部字节不变且原缺口没有修复：F07/F08/F09仍缺可靠原生伸缩表示，两个OMML目标F02/F05仍有明确斜体丢失，不能声明20/20。其他符号身份、任意/负间距、单边上取整崩溃、CP255固定重音、复杂上下标与完整视觉验收仍开着。下一步应寻找真实可编辑MathType特殊围栏样本/官方装配契约，和保持Word运算符角色的显式样式表示证据；没有新证据前不借模板、改角色或猜测恢复。新的隐式mfenced离线异常独立记录，用户已搁置的外部运行器栈调查未展开。

最终39/39独立定向测试；正式聚合2408/2409的唯一CP255失败保持（该聚合在最后追加两条测试前执行，两条均已在39/39中单独通过）。早期4个未启用原生测试的提前返回没有计入正式聚合或原生通过。生产改动只有codec 47增/2删；新测试独立文件，无既有比较器/测试预期修改。完整C#候选SHA仍0DE80AE7800F1B2042AEB934CEA9F78637EC8644BBF616DFB53885F294B36E6C，源码/产物/FreezeR3清单source-build-hashes.json；heldR4保持6692FDE3…AC。仍非默认脏工作区构建。

全部自有Word已退出，本轮Companion64148已按精确路径/哈希/UTC创建时间核对后停止；原3个Word和4个Companion身份不变，4处LoadBehavior存在性和值不变（environment-after.json）。收尾首次日期比较对已反序列化DateTime再次Parse导致时区误判，未停止任何进程，核对精确UTC身份后正确收尾。没有改安全设置、安装/部署、提交/推送。HEAD仍a57e568，原所有未提交修改保留。汇总matrix-summary.json，所有失败/未执行边界留在本报告及原证据中；当前工作树仍不适合宣称完整保真发布验收包。
