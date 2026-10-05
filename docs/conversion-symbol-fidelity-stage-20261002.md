# 公式转换符号保真：阶段报告（2026-10-02）

## CP261 当前代码本地快照（2026-10-04）

用户明确要求将当前代码（含另一任务推进到CP260的修改）本地提交。快照保留共享转换、字体/重音/颜色、编辑器弹层、Rust及运行时修复、测试工具和文档；维护中的截图脚本按原路径单独纳入，构建产物、缓存、图片、Word文件和原始日志留在本地。没有部署、安装或推送，也没有为提交改动生产行为或测试预期。

这是当前源码快照，包含此前held R4解析器（6692FDE3…3BAC）及CP255未验收候选。既有候选使用FreezeR3构建，不能把本快照默认构建等同于那些验收产物。核对CP257记录的10个源码哈希及CP260编辑器哈希均一致；截图脚本已是后续11EC275E…68B1D6版本，语法检查通过。既有汇总2337过/1失败及重音固定组位置缺陷保持，不宣称全树或最新六向视觉通过。本轮只做提交前检查，不重新运行产品测试。完整提交身份和文件清单见artifacts/local-commit-20261004/，细节见历史记录CP261。

## CP260 经典桌面颜色弹层修复完成（2026-10-04 08:00 UTC）

用户批准后，仅改EditorWorkspace.tsx的弹层挂载/定位：经典桌面也挂到body，固定在工具条上方并保留7px间隔；短窗按上方空间限高、内部滚动，不再被dock挤下来遮住按钮。通用避让、Office挂载和颜色/转换逻辑未改。

经典专项75/75通过，覆盖四种窗口、原按钮关闭、字体/背景双向切换、Esc/外点/重开、局部改色、去背景、自定义增用删、开着缩窗/恢复/滚动。标准实际交互24/24、桌面和Office双布局几何24/24通过；18个未改标准桌面/Office对照与CP258位置完全相同。760×360短窗背景弹层截图已实际查看。自定义值用事件输入，未验原生取色器或原生WebView2/Word生命周期。

桌面/Office类型检查和隔离Vite构建通过，最终桌面main-DskIrg8C.js SHA25609EF066C76EE0DA7AA18C73760FE5AF0F5ACFBC4F4BA05A04F461C19141B7086；Office dialog-_2GGEuf8.js SHA25681445AE2A333CACBF99D624D16E5F74776CE629F9B26C1CBF7998105581C1A51。证据在artifacts/conversion-symbol-audit-20261002/cp260-classic-color-popup，含final-results、after-interactions、after-{desktop-geometry,geometry}、unchanged-controls及final-state。最终检查首次调用返回取消，经用户明确同意原样重试一次后通过。未部署、安装或改普通dist，CP255等未验收边界不变。

## CP259 经典布局专项核查：确认按钮被遮挡（2026-10-04 07:45 UTC）

本轮仅验证，复用CP258最终隔离桌面构建；源与构建哈希一致，没有改生产代码、重建、部署、安装或触碰Word。

1280×760、560×560、760×360、440×360均复现：颜色弹层本身可见，但遮住原字体色和背景色按钮，无法再点原按钮关闭，也不能直接切换。根因是经典桌面底部dock的overflow:hidden被通用浮层避让当成边界，把本应向上展开的弹层强制下移到dock内部，覆盖按钮行。1280宽截图已实际查看。CP258经典对照只测弹层边界及内部命中，因此没有覆盖这个交互缺陷。

Esc、外点关闭、重新打开、局部字体/背景色、去背景、自定义字体/背景颜色增/用/删在四种窗口均通过；选中x改色后+y保持。短窗自定义项可滚动后使用。原生系统取色器未测，输入值使用事件模拟。两类自定义颜色列表本来独立，首次探针跨列表复用的错误假设已纠正并保留旧结果。

final-results.json原始59项为48通过、11标记：8项是确定的两个按钮×四窗口遮挡；另3项是150px滚动弹层里未滚动时控件中心不可见，滚动后的应用/删除成功，不计为新增功能缺陷。开着缩窗/恢复/滚轮检查没有出视口，但仍存在覆盖按钮问题。证据目录artifacts/conversion-symbol-audit-20261002/cp259-classic-color-popup，脚本cp259_classic_color_popup_qa.mjs。拟只修经典桌面弹层挂载与锚点，等待用户批准；未声称原生WebView2/Word端到端或发布验收，CP255等既有边界不变。

## CP258 标准布局颜色弹窗裁切修复（2026-10-04 06:48 UTC）

独立Windows编辑器标准布局在1280×760、760×560、560×420均复现：字体色和背景色弹层被34px高横向滚动工具栏及表头裁切，通用避让随后压成96px并推到屏幕上方。前后截图已实际检查；经典布局和Office对话框为正常对照。

生产仅改EditorWorkspace.tsx：标准桌面布局的颜色弹层通过现有PortalOrInline挂到body，固定定位跟随工具栏，并保留原视口避让和横向滚动；外点关闭同时识别弹层。开关回归还发现Esc关闭会透传MathLive删掉选区，现只在颜色弹层开启时消费Esc。原始+y证据保留，最终断言为关闭后仍是x+y。

最终桌面/Office类型检查与隔离Vite生产构建通过。24/24布局与命中检查（双端×双布局×3窗口×2颜色）；24/24实际指针/键盘交互断言覆盖开关、Esc/外点关闭、重开、局部字体/背景改色、去背景、自定义增/用/删、切换、开着缩窗和工具栏滚动。自定义颜色输入用事件模拟，未验系统原生取色器；Office为构建后的React界面，未声称原生WebView2/Word端到端。

证据目录artifacts/conversion-symbol-audit-20261002/cp258-color-popup，最终after-r2-{geometry,desktop-geometry,interactions}.json及截图。最终桌面main-DSX507lG.js SHA256 BFF94B618F0573FE72F41179EF768A9B4B308C39727F82FE09C6D35D9060A9B5；Office dialog-C65l998i.js SHA256 8581FBAD5341CD460ABAB77643470EC30A033752DE204B217448F6122B6D60D9。源码检查通过，构建只写独立验收目录；普通dist、运行实例和安装包未替换。未跑转换大矩阵、C#/Word或性能；CP255未验收状态、CP257 DLL、heldR4和其他脏改动保留。

## CP257 OMML 编辑器改色基线（2026-10-04 05:52 UTC）

本轮只处理用户新报的OMML字体颜色。实际生产MathLive选择改色及MathJax输出5例（整式、局部、粗斜体、分式、重音）都保留textcolor/mathcolor；精确CP255-r3 DLL随后生成的5份OMML全部缺失w:color，证明颜色在写入Word之前已丢失。MathML恢复LaTeX也完全遗漏颜色，需同时覆盖重开。原失败证据保留于cp257-color/before.json及cp257-editor-color.json、cp257-color-payload.json。尚未改生产代码或运行C#/Word；现有语义签名忽略w:rPr，不能冒充颜色验收。本轮保留CP255未验收候选和所有既有脏改动，未改变运行实例/安装包。

CP257 06:06 UTC：共享颜色作用域已写入run及分式/根号/重音ctrlPr，读回MathML与源码保色。专项13由12失败/1通过到13全过；相邻OMML/MathML/样式1016全过，无旧断言修改。原失败TRX保留。FreezeR3完整候选cp257-omml-color构建成功，精确哈希见build-state.json。未动比较器、指纹或运行时校验器；前端生产源码无改动。5例真实Word改色/Undo/Redo/重开/不改内容再应用已准备，尚未运行。当前树仍含CP255候选，不能因此宣称全汇总或安装包已验收。

CP257 06:14 UTC限定验收完成：颜色候选DLL为9D94D119601529AF7B6C6517F3CD3A7AEF651CEAA0F4D15F0D5E0C4ADCC68AF7（FreezeR3，heldR4不变）。5个真实Word案例分别完成整式/局部/粗斜体/分式/重音改色、Undo/Redo、保存后新Word重开、原生读回后不改内容再应用；逐字符RGB与控制字符RGB独立验过，原数学样式和邻接文字保留。证据cp257-color/word-edit/color-lifecycle.json及各阶段XML/MathML、5份docx。10个自有Word全部退出，原3个用户Word身份保持。此为生产编辑器API/导出+真实Word共享编辑内核验收，没有声称物理点击对话框或截图通过。

全managed汇总2337通过/1失败/0跳过，唯一失败仍是CP255原固定宽度字节测试，未改旧测试。上/下标、函数补充后颜色专项16/16；相邻1016/1016。保存源码再载入/编辑原内容颜色5/5保留（editor-cursor-audit.json）。代码核实普通重开后在蓝色末尾输入+z应继承光标颜色，输入内核明确保留普通LaTeX样式继承；单次工具栏选区改色后则执行color:none，新增字符不继承，该另一个入口也5/5。首探针误把普通重开追加字符期望为无色，得到1/5，原报告保留；现已按两个实际产品入口分别解释与验证，没有修改前端或放宽产品断言。候选包含CP255未验收内容，因此不宣称全树/发布通过。运行Companion、前端和安装包均未改，未运行性能/六向矩阵或接触用户文档。最终再次FreezeR3构建cp257-omml-color-final与真实Word验收DLL字节一致（9D94D119…CC68AF7）；cp257-color/final-state.json保存最终源码、测试、旧fixture与held resolver逐文件SHA256及最终进程清单。真实Word验收后没有生产代码变化。

## CP256 收敛与交付边界（2026-10-04 03:40 UTC）

保留CP253/254已验证的控制样式、源码与编辑器字体作用域修复。CP255宽度改动保留为未验收候选，不据此更新运行实例或安装包；源码8/8、40专项通过，完整汇总2324过1失败，不能声称不影响全部既有功能。原字节测试输入/预期未改，等待具体批准。固定组重音的原有垂直间距失败单独保留，不能用测试改绿代替修复。

本轮仅核对候选边界并记录，不再增加规则、重复矩阵或调整性能。CP255候选DLL为6A165BCAABD6A8DD2BCF98C2AAF29B6C029DD5AA5C1D5DFB61F450FB75FE3108；此前已完成限定验收的C#为CP253 B4F536…FC8DB，最新已验证前端为CP254。当前源码仍包含候选，正常构建会包含CP255改动，没有禁用开关或独立源码隔离，因此不得将当前树直接打成已验收包。此前已验证的CP254前端产物为dialog-BeoXyVTH.js，SHA256 18CB4FEAE47726CDA67A19B06C41AAF7ADCF3A1550105ACCD96717637A4CBEC9；已有实例可能缓存此前前端，未声称它们已加载该产物。heldR4、运行实例和旧CP235 NSIS未改；没有本轮启动的测试或Word进程。

## CP255 重音宽度候选：部分修复，仍未整体验收

03:16 UTC：原生位置依据审查结束。厂商明确字符EMBELL只能附着一个字符；HatBox翻译规则对应宽重音。当前固定组重音路径依据来自CP122人工构造的探针，并非MathType原生编辑生成的样本，旧保字测试没有证明垂直位置正确。缺少合法固定组重音的完整CHAR/字体/尺寸/模板及实际渲染样本，不能凭空偏移或换码位。此项保持真实视觉失败。原字节测试修改等待明确批准，未再启动测试或Word；当前运行实例、安装包未更新。

03:05 UTC：源码几何8/8、40专项通过；新FreezeR3 DLL6A165BCAABD6A8DD2BCF98C2AAF29B6C029DD5AA5C1D5DFB61F450FB75FE3108。完整汇总2324过1失败，原字节测试未改，保留未显式声明固定宽度与其EMBELL预期的冲突。新增独立固定字节反例通过，不放宽比较器。

6个隔离原生图均生成且已逐张查看，宽窄帽/波浪明显可区分；但多字符底数上的固定重音垂直距离过大，未通过像素保真。5/6原生载荷与CP253字节一致，含已有垂直距离缺陷；本轮仅wide单字符帽改为真实扩展模板。保持旧图、失败及前后载荷，不能称本阶段全部通过。

此前定位：

Office标准重音使用组合字符并随底数扩展，没有已找到的合法accPr grow开关；不使用未获依据的U005E/字体缩放替代。另当前DLL直接MathML→源码8例中3例宽重音被缩成普通命令，属于仍可保留信息时的恢复错误。正在核对MTEF固定EMBELL与扩展字形的结构依据，避免修源码时把原生固定重音误当宽重音；本阶段尚无生产改动。

## CP254 字体解析共同根因修复

最新02:37 UTC：符号整体加粗与字体选择分离，内层字体先解析、再按受支持字体轴合成bold；不在AST/产物保存隐藏字段。生产kernel字体作用域20/20、原CP253编辑器20/20均过，保留首次4/8与18/20失败。字体/运行时安全回归、TS及Office生产构建通过。C#仍CP253 B4F536…FC8DB/2302项通过，本阶段不重复Word/native矩阵，不代表六向或像素全通过。原bold-sans编辑器残项已在本轮解决。

此前基线：

8个当前生产编辑器小样4过4失败：boldsymbol的独立加粗作用域被内层mathsf/mathrm覆盖，另mathfrak错误保留外层mathbf字重。AST生成前已混同，不能给mathsf加保留bold特例。正在独立候选中验证符号加粗与字体选择两条作用域规则；尚无CP254生产修改或通过声明，原失败保留。

## CP253 控制字符样式双端与编辑器字体作用域修复

最新02:18 UTC：最终FreezeR3 DLL B4F536B93423255988FF65BA88CA3E31ADAE3B6C59C2B3675EDBF2508EAFC8DB，2302项managed通过。5式VT→OMML及保存OMML→VT均通过内容/持久源码、Undo/Redo和重开；独立前向角色/样式10/10、保存源码实际MathJax重解析5/5一致。编辑器使用当前项目kernel及完整Vite变换验证，基于AST字体族/字重的序列化替代字符猜测，删除被替代的Greek字符特例；20项19过1失败，剩余boldsymbol内mathsf在解析时已丢bold，未忽略。宽窄重音与Word像素尚未验收。性能冻结，旧安装包不变。

此前记录（已由上述最新状态更新）：

最新01:55 UTC：标准控制格式双端已修，14专项及2299汇总全过。FreezeR3 DLL6B413D0F324915A8B92A7290DD32588262EC711BE8271966894B1C4E075222B3。5个生产Word文档20阶段均保留重音normal/bold/italic/bold-italic及底数独立样式，尚无像素验收。源码恢复仍不完整：ConvertOver原有样式忽略；候选标准字体包装在MathLive编辑后被重写bm，2/4例污染底数字体，已保留失败且未集成该源码候选。当前继续查编辑器共享序列化合同，不报全链通过。

OP-only无应用标记的标准替代已核对：func会增加应用含义，Word opEmu是binary，不能用于偷换角色。另8个小样证明hat/widehat、tilde/widetilde、vec/overrightarrow的宽度区别在前向OMML已混同；加粗重音与仅加粗底数也混同。后者有标准m:ctrlPr独立格式合同，已启动3个自有Word原生控制验证，尚无修复/像素通过结论。生产DLL仍C799，只有诊断入口支持明确原生OMML输入；性能冻结，未扩大20或安装。

## CP252 函数源码恢复：构建通过，小样闭环与剩余边界

2026-10-04 01:28 UTC最新终态：用户明确批准7条旧源码字符串预期的更正，原被拒调用原样重试一次成功，原失败TRX保留。当前managed2285/2285、0失败/跳过（原4项native排除），FreezeR3完整构建DLL C7995FC64E4B83AE12B73FE487917A553A5A48F54FCB205F2986769FFB914605，heldR4未变。Office前端已构建，新隔离Companion70768/43131正常TLS启动，原Arc修复EXE不变，未安装/更新NSIS。

4式VT→OMML及保存OMML→VT均通过内容/Undo/Redo/重开；独立token/应用标记/脚本槽位审计前向6/8阶段一致，另2阶段为同一无应用标记operatorname{sin}被OMML读回拆成s/i/n。此项未通过，不能以内容gate绿色宣称整体保真。另外3式保存恢复源码保留sin的U2061、foo上标后的U2061，以及mi limit下标结构；MathLive真实源码编辑5项、权威MathJax12项、字体84项和TS已过。无当前像素验收；MTEF函数/应用标记碰撞及原F02/F09等仍未修。证据cp252-function-recovery/word-forward-audit.json、reverse-saved-source.json及两方向cp252-*function4目录。反向cached源码仅辅助证据，Word生命周期另核持久LaTeX。

剩余边界已进一步定位到Word之前：精确C799 DLL生成OMML时，OP标记的sin、普通直立多字mi sin、三个独立直立mi s/i/n三者字节相同；实际U2061则生成并保留m:func。证据omml-run-boundary.json。不能靠反向按名称猜测修复，当前候选只保证输入仍含角色/应用信息时不再丢弃或凭空添加。全部自有Word已退出，未运行20扩展或截图。

### CP252此前调查记录（以下历史状态由上文最新终态取代）

后续更正：真实自有headless编辑器set/get/插入后源码保留U2061，权威MathJax再解析5/5过；此前MathLive旁路MathML导出的异常不能证明产品源码损坏。7旧limit预期另有三方role证据：输入mi，旧内建命令改mo，新候选保留mi/munder。候选hunk已重新进入源码审查，但更新旧测试预期的动作被自动审批拒绝，未重试绕过，旧断言不变，源码当前未验收；运行实例/最新已验证DLL仍CP251。详见检查点与cp252-function-recovery。新增12项真实mo/mi正反例、原84字体/SVG回归及TS均过；旧7断言未改，等待该被拒编辑的明确授权，不报整体通过。

MathJax→C#→MathJax12项通过，但MathLive把显式U2061当mi并添加隐乘；另有7项limit-mi脚本回归。候选不能作为可编辑源码修复，原失败保留，未改旧断言。仅恢复本阶段自己新增的生产hunk到已验证CP251，候选代码/测试/构建作为未启用证据保留。继续只读核对编辑器实际支持的应用运算符合同，原生MTEF函数/应用标记碰撞未修。恢复后完整FreezeR3构建与CP251 DLL字节相同（24DC7AFBD3DEB17438965DD8F2D2969DE3F1FE29AF49A20B6938AC2068D22DC2），2273项managed重新全过；候选未进入运行实例/安装包。

## CP251 已解析 OMML 导入不再重复重建

重新核对磁盘，保留较新CP249/250全部修改。三份相同Word根因对照仅省去BuildUp，严格12阶段由6过6失败到12/12，固定括号/端点粗体经排版和重开仍保留。共享5个已解析OMML导入边界移除重复BuildUp；真正线性输入和旧编号删除修复不变，比较器/边界检查不变。

FreezeR3 DLL24DC7AFBD3DEB17438965DD8F2D2969DE3F1FE29AF49A20B6938AC2068D22DC2，heldR4未变。2273项managed全过，生产VT→OMML3式严格逐项/Undo/Redo/重开通过；20结构控制也已通过逐项内容、Undo/Redo及保存重开；MathType源3式→OMML与其VT→MT种子3式也通过同样生命周期。独立presentation初筛32/40一致、8告警，其中6为明确结构围栏两种表示差异，内固定括号实际保留；另2为accent固定标记未读回，几何仍未验收，原告警保留。不是六向全符号/像素验收，原F02/F09/hbar等失败保留，CP235安装包未更新。

## CP250 固定 run 反向修复；Word BuildUp 另有丢失（2026-10-04）

旧 CP230 读取缺口已通过明确审批的本机只读解除。实际 XSL 中 ParseMt 的普通 m:r/m:t 与重音调用上下文可区分；方程阵列另有专用 run tokenizer。仅在这两处普通 run 的 mo 发射中保留标准 stretchy=false，不按符号打补丁、不修改比较器、不影响结构化围栏/重音/大型运算符。改动 WordOmmlConverter.cs，新增 WordOmmlFixedRunProjectionTests.cs。

专项20基线11失败9通过；修正后含相邻92项共112/112。首次聚合2271过2失败是沙箱拒绝既有 SVG 测试的 AppData 临时写入；原报告保留，同一二进制正常审批复跑2273/2273、0失败/跳过，仍排除原4项原生服务测试。完整 FreezeR3 构建 cp250-fixed-run-geometry 成功，Word DLL 7A76E679CA8440D7653E370728D8ABC18EC88C6238E6A8963AD30335E234A415，held R4 未变。证据 cp250-fixed-run 与该构建目录。具体 Word 验收边界见以下终态，安装包仍 CP235。

本阶段终态：精确DLL对照的旧5项固定括号签名0/5→5/5；原20条纯投影仍20/20，前向OMML字节全部不变。真实Word3份文档共12阶段严格审计为6过6失败：右括号带上标例4阶段保持正确；另外两例插入正确，但BuildUp将普通括号run改成m:d，之后排版/重开仍丢固定几何及一例端点粗体表示。原始失败XML和独立审计保留，不能称三份Word保真全过。标准m:lit可能改变数学间距，未作为通用补丁，未运行该假设的原生探针。下一步查共享插入中BuildUp对已解析OMML的必要性与语义影响。

原20条实际已保存OMML→VT通过逐项内容、Undo/Redo及新进程保存重开；源仍含历史F02等损失，不是原意恢复或六向全通过。为确保隔离，验收Program.cs新增可选自有Word根目录入口，复用现有/x/Automation、PID/创建时间/自有文档验证工厂。最终构建cp250-fixed-run-owned的Word DLL仍为上述7A76…A415，验收EXE AEA8A25A12CE3DF1D0C1541073B8782A876706DE8E779C7518E607D1EC798F40。测试最后修正opEmu负控为合法box结构后专项20/20再过；2273聚合未在这个纯测试修正后重跑。没有当前版本截图或另五向/原生渲染重跑。

精确哈希与终态见cp250-fixed-run/final-audit.json，纯对照pure-before/after.json，Word失败word-stage-audit.json，日常回归cp250-Omml-VT-daily20。所有自建Word退出，原3PID及初始精确创建时间一致，无测试在运行；初次只读审计的JSON毫秒精度误差已纠正并保留原报告。HEAD仍fed066192373fc12364e822f5478acca9967f7b0，未暂存/提交/安装/推送，原修改全部保留。旧CP235安装包哈希重新核对B02C75CBDD09FD32B517DED6198B8916F4C1C7871338EF0CF92F57C722AA9929不变，不包含CP239后的修复。

## CP249 本机接手与六例函数结构定位（2026-10-04）

HEAD、最终 CP246 FreezeR3 DLL B4C56A2C37200B92F5D64B381C48F40B9EEBABE413D878D17F1AD6B66B840696、冻结 resolver 与 held R4 均重新核对一致；2253/2253 为旧 TRX 只读核验，四项原生排除保留，本阶段未构建或运行聚合。全部既有修改保留，无提交/安装。

当前有真正本机 Windows exec_command。受限 CIM 拒绝后，正常 require_escalated 审批成功，核实 Codex→PowerShell 本机执行链及三个原 Word 创建时间。原截图脚本仍为 CP248 的 11EC275E…68B1D6；新清单绑定历史 CP234/235 的实际 OMML/VT、20公式中 F09、转换日志及哈希。ValidateOnly 后同脚本正常审批执行，owned66652 一开始即为前台、SetForegroundWindow=true，成功1图并关闭退出，全部哈希不变。已查看图片，问号框仍在且 PNG 与旧图相同。这证明当前本机入口单次可用，不能称稳定修复前台或当前全20视觉通过。

六例精确 DLL 诊断证实 sin/cos 的普通函数与 operatorname 无 U+2061 版本产生完全相同 MTEF；直接源码恢复又仅凭名字选择内建命令，未知 operatorname{foo} 还会拆成 foox。这是两层丢失，不能只反向猜测或豁免新增 U+2061。未改生产规则/比较器。证据 cp249-handoff/{audit.json,function-result.json,historical-f09-capture/capture-summary.json}，新增脚本 cp249_function_application_probe.mts 与 Probe-Product.ps1。下一步检查普通 OMML run 的精确 token 发射上下文，争取推进 CP230 固定围栏共享修复；遇权限拒绝走正常审批，不绕路。既有 F02/F09/hbar/script-l 等失败保持；安装包仍 CP235。

## CP248 正常启动/应用激活对照

4个小样已终态：Shell启动首例截图成功，Direct对照、Shell复核和自有Word.Application.Activate均未取得前台。失败时SetForegroundWindow=false，同会话/desktop、Word可见未最小化，不是用户切走证据。首例F09仍与旧缺字图一致；所有自有Word干净退出，原3进程未动，没有继续20循环。

缺少的是前台执行方对精确测试PID/COM代理的正常Windows授权交接；后台脚本不能自授。当前连接器无此前台入口，修改桌面/连接器集成需要另行授权及可用源码，或明确交互启动。原截图检查全部保留，新增可选正常Shell/应用激活诊断、默认Direct不变。脚本SHA11EC275E85EB4C236EA48F142C593DDD47C2B4CA57BF11BC67BB1CD1FD68B1D6，证据cp248-shell-launch。不能称截图流程已稳定修复或完整20视觉通过。

## CP247 截图前台归因更正与 F09 字体对照

用户未操作 Windows；此前 foreground mismatch 不能证明用户切窗。原截图脚本仅补充激活返回值、窗口/会话/desktop/最小化状态日志，保留全部检查及原激活流程。预检后正常取得保存的 F09 文档级 Latin Modern 对照图，实际像素与原图字节相同，问号框仍在，不能宣称换字体已修复。本次 Word 从初始检查即为前台，SetForegroundWindow=true，没有复现过去失焦原因。自有84028已退出、原3Word受保护、全部输入哈希不变。记录/原脚本备份/新图见 cp247-foreground-diagnostic；新脚本哈希 C26549BB8D24C12FECDC01D48141F0986D1649B78E4B84696B8C6FA28854C1D8。图是明确标注的历史 CP234/235 字体诊断，不是当前CP246六向视觉验收。

随后一次完整20正常尝试复现：从首次记录起前台就为ChatGPT，SetForegroundWindow=false，document.Activate前后均未切入Word；会话/desktop相同，Word可见且非最小化。这次是没有取得前台，不是用户切走的证据。具体Windows拒绝条件未唯一定位；没有启用线程输入连接、假按键或关闭检查。0图，owned19816正常退出，完整20仍未验收。官方Windows契约和实际后台connector进程链说明当前路线没有已证明的稳定前台授权传递；AllowSetForegroundWindow也需有资格的前台方授权，脚本不能自授。具体旧失败条件仍不能唯一确定，不是杀毒归因。下一次稳定自动化需正常前台执行器handoff，或明确交互激活自有测试窗口，不能靠重复抢焦点/延长等待声称解决。

## CP246 原生 glyph 与基本样式选择顺序（2026-10-03）

最终补充：生产序列化省略单 mi 的默认 italic，已在既有 legacy glyph 资格范围统一读取有效样式；过宽候选触发4个字节负控后收窄，原失败与断言保留，53项负控及2253项聚合全部通过。最终 FreezeR3 DLL B4C56A2C37200B92F5D64B381C48F40B9EEBABE413D878D17F1AD6B66B840696。ℓ 四基本样式真实 VT→MT 与 MT→VT 均完成严格逐项、Undo/Redo、保存重开验收，测试 Word 全退出，仅原3进程。以下20原生图来自前一 B72C 显式样式候选，不能混称新DLL全20像素验收；本轮没有Word截图。

删除 ℓ 与 script l 的错误签名合并，保留精确码点和有效样式。生产编码改为先选择既有权威 Symbol/MT Extra 字形，再绑定基本样式；没有新逐字符映射。15 项定向及 2249 项 managed 聚合通过，FreezeR3 DLL B72C7477476CC3C4C337CC5BD27A0F6795E199ACFD0229EE319831DA76A5FFF4，heldR4 未变。

同组 20 原生预览全部返回并逐张查看。ℓ 的 normal/italic/bold/bold-italic 均正确可见，粗体/粗斜体不再是缺字符。对照 CP245，17 张墨迹相同，3 张 ℓ 样式改善；其余 16 控制保留原像素，其中 lowercase script l 两张旧空白仍失败。不能把 20 返回当成 20 保真通过。证据 cp246-native-style/rendered-r2；Word 像素/完整六向本轮未重跑，OMML 身份碰撞及 F02/F09/hbar 等边界仍在，安装包仍为 CP235。

### 当前剩余验证/证据边界

- 固定围栏：CP250已获正常读取权限并修复普通run的精确token反向几何；真实Word仍会在部分BuildUp情境把固定括号改成m:d，前向/宿主保真未完成，不能全局给括号加属性
- escaped punctuation 的 mi/mo：既有对照产生相同 OMML，反向无法猜回来源角色；需无语义变化的标准可持久表示证据
- 任意/负间距：当前 vendor nominal code/NUDGE 尚无精确 advance 合同；需官方编码说明或原生可编辑精确宽度样本，不能近似冒充无损
- script 小写、hbar alternate、特殊单侧/组围栏与单位：需原生可编辑样本及字体偏好，或确切字体/编码/槽位权威。script 小写已证实误选只含大写的 Euclid 字体，尚无合法替代元组
- F02 与 ℓ 的 OMML 存储身份碰撞仍失败；F09 文档数学字体对照已在CP247截图，和局部字体对照一样未改善问号框；CP249本机截图单例也未改善
- 复杂脚本既有结构/角色修复有定向证据；视觉基线、完整六向及最终统一构建大覆盖仍未完成。本轮没有发现另一项已有充分证据可立即修改的独立结构根因，不重复已知失败矩阵

## CP244–CP245 两项通用根因修正（2026-10-03）

修正已证明属于导入的段落标记删除方式：原Range.Delete额外吞原文空格，改精确Text清空，严格边界比较不变。2项VT源和4项MathType源→OMML的内容/Undo/Redo/重开通过，20项已真实完成转换，随后停在ℓ→script l身份/样式差异；未放宽该失败。

反向ParseMt按UTF-16单元拆补充平面字符会产生非法代理项，现按Unicode标量执行该模板的substring/string-length，保持XPath1边界行为、码点和样式。21项定向、2233项managed聚合通过；精确新产品DLL六项对照异常2→0、码点/样式6/6保留。CP245 DLL SHA38F3CAA7A2E5598621016006B91CBF178F1189D3B03FC86AAB32153524E9EA8C，FreezeR3，heldR4未变；现NSIS不含这些新修复。最终浮点舍入边缘修正后22项定向、2234项聚合通过；最终DLL更新为FDC97376552067E29A93F4BDA537C0A23953FB5E9A7A09107DFA5DC116325A26。真实Word U+20000汉字在插入、BuildUp、排版、保存重开四阶段精确保留码点/mi/normal，无代理项异常。3份字母表对照也完成，ℓ身份碰撞发生在Word插入阶段；没有新的像素通过。

ℓ在纯XSL中不丢，实际Word存储改为script l；Unicode标准区分U2113与U1D4C1，不能用NFKD或旧签名捷径判等。F02、F09文档级字体像素和hbar原生alternate仍开放；本阶段不是全六向视觉通过。

## CP243 最新20联合验收终态（2026-10-03）

经授权清理并确认所有自有残留退出，重跑20 VT↔MT的内容/有效样式、Undo/Redo及新进程保存重开均完成。正向独立presentation证据40条34过6差异，为3项hbar alternate丢失的前后两次记录。反向保存源码经当前前端再解析16/20一致；3项hbar身份损失，1项operatorname→sin新增U+2061函数应用标记，严格保留差异，不称整体保真通过。最新4项VT→OMML样式/生命周期通过，20项在检查。

F09文档级数学字体副本的截图仍因ChatGPT占前台被正确阻止，未得新像素；没有继续抢焦点。局部ctrlPr改字体无效的真实图片结论不变。所有运行结果见CP243；当前NSIS仍不含CP239。

## CP241 当前联合验收与字体反证（2026-10-03）

最新DLL与新构建前端的4项VT↔MT内容/有效样式、Undo/Redo及保存重开均通过，但hbar alternate在MathType链仍丢，不能称整体保真通过。20项共享样式/Greek/嵌套/脚本回归写入保存后被60秒工具窗口中断，未完成逐项/Undo/重开，不能报通过。剩余自有Word85832已只读核实仅持有该保存测试文档，待清理授权，新Word启动暂停。CP242已保存。独立presentation报告明确正向8条证据4过4失败（hbar alternate丢失）；反向无同类JSON为0条，非通过。新Companion42504/43130正常TLS，未安装。

F09 R2已被自有Word成功打开，修正之前“打不开”的推断；局部ctrlPr改Latin Modern后实际截图与Cambria原图字节相同，问号仍在。文档级OMathFontName改Latin Modern的新副本已保存并核实20公式，截图因前台归属不匹配停止，效果仍未验证。3个早期诊断脚本错误产生的自有空白PID22648/106232/2540待用户授权清理，自动审批拒绝后未绕过；原用户Word未动。详见CP241及cp240-font-open证据。

## CP239 显式identifier样式共享合同修正（2026-10-03）

已修双端信息丢失：C#保留mi显式normal，前端按显式基本字体作用于单字符identifier，不再以Unicode Sm筛选；保护Bbbk等内在字母表、函数角色、特殊mathit字体及嵌套重置。实际MathLive支持alpha normal/bold语义，候选与其对齐。源码20由9过11失败到20/20，20项通用控制、58operator回归、Math-variant和Office TS检查通过，184相关纯测试全过。三条旧字符串预期差异经实际前端语义等价验证后更新，原失败报告保留。SVG一致性不等于normal/italic内在轮廓不同，也不是Word像素通过。2026-10-03 05:50 UTC补充：当前树FreezeR3完整C#验收构建成功，Word DLL SHA256 9330F9D367A6FEDF17EE7DEBAEEE5431802874FFF72B8660285D6206CCEFFC57，held R4哈希不变。完整managed聚合首次2209过3失败，三项是normal mi的旧裸hslash/hbar/angle字符串预期；实际前端验证前两者裸命令丢normal而显式mathrm正确、angle两种写法token/normal等价后更新测试预期。原失败TRX保留，复跑2212/2212通过，0失败/跳过，仍沿用4项原生测试排除。报告shared-contract-aggregate/cp239-final-integrated-r2.trx。Word候选及最新前端联合验收尚未完成，现有CP235安装包不含此改动；F09缺字及F02原生样式问题仍开放。主checkpoint读取一次取消后按用户重试授权原样重试成功，以上结果已保存为CP240。


## CP238 F09字形覆盖/伸缩能力边界（2026-10-03）

F09问号处保存的OMML仍是U+23D0，控制字体Cambria Math；源码未变成问号。Cambria缺该码点，Latin Modern/STIX有字形但无该字符MATH竖向variant/assembly；MathJax另有U+2223组件伸缩数据，因此不能用换竖线或覆盖字体宣称完整修好。仅测试副本的字体对照两次未获得Word document对象，无新像素结论；第一份测试XML丢命名空间，第二份保留声明后仍打开失败，原因未定。仅确证自有测试PID清理，原用户Word未动。相关C#两次读取被取消且未绕路重试，生产未改。继续独立前端hslash/hbar共享样式作用域调查，先建立函数/Greek/字母表反例。


## CP237 当前构建4图成功，完整20因前台归属检查停止（2026-10-03）

用户授权重试后，同一原脚本与正常入口成功取得F02/F05/F07/F09的实际保存OMML截图。独立Word50948关闭退出、文档哈希不变、无清理错误。已逐张查看：F02平行线仍直立，F05角度与F07角括号可见，F09第二组出现问号框；这是可见缺陷证据，未推断新根因，也非完整保真通过。完整20预检通过，实际Word55872因前台属于另一PID78092而拒绝截图，0张，正常清理退出；未重试或改检查。截图目录为cp236-current-word-visual/key-capture-retry-20261003T0430，完整20失败证据为all20-capture-20261003T0432。没有新的六向或生命周期验收，CP236严格F02失败继续保留。


## CP236 当前构建截图预检通过，执行请求被取消（2026-10-03）

已核实修正版原脚本37AA…FA5C与CP235-S2。精确CP234/CP235同一Word DLL新生成20项VT→OMML并保存，F02斜体运算符严格语义仍失败，未进入Undo/Redo和重开。当前输入清单已记录两份文档、构建与转换日志哈希；F02/F05/F07/F09截图ValidateOnly通过。实际正常Devspace执行返回`user cancelled MCP tool call`，没有PowerShell输出；不重试、不换路线。只读确认新截图目录不存在，Word仍仅原有88900/93828/95356，无新截图可供验收。该结果不能归因为Defender或登录会话错误。证据在artifacts/conversion-symbol-audit-20261002/cp236-current-word-visual/和cp236-VisualTeX-Omml-current20-r2/。生产代码与安全配置未改；当前构建20项及六向视觉仍未通过。


## CP235-S2 真实 Word 截图功能已恢复（2026-10-03）

修正原路径 Capture-CP198R1Style.ps1 的实际功能问题：改用显式输入清单与文档/构建证据哈希；校验公式数量、编号和新输出目录；失败及清理异常返回非零且保留原始错误；移除可能发往其他应用的全局 Alt 按键；补齐异常路径资源释放。保持原 CopyFromScreen API、自有进程/文档/前台验证，不改脚本名、不编码执行、不换截图 API，不关闭防护或添加豁免。修正后 SHA25637AA594D89A278CEFC5A762341C840595F9AC3931329AB4E4CB4B8E66402FA5C，原始版本、完整差异和证据均保留。

语法、正向预检和5/5异常输入拒绝检查通过。默认受限执行环境实际启动 Word 时返回“指定的登录会话不存在”，正确以1退出，未启动 Word；通过工具审批后，在正常用户执行环境运行同一脚本成功：独立 PID88632 取得历史 CP198 R1 的 F02 真正 Word 截图并关闭退出，返回0，原有三个 Word PID及输入/副本文档哈希不变。实际 PNG 已查看。此为截图功能验证，不能计为 CP234/CP235 或最新六向保真通过；原公式的既有样式/间距问题仍未解决。

成功后只读确认 Defender 实时保护、行为监控、防篡改等均开启，情报版本仍1.459.491.0；本轮未查到针对该脚本的新检测/处置事件。旧 AMSI 命中特征仍未知，不宣称已确认某行导致误报。证据根目录：artifacts/conversion-symbol-audit-20261002/cp235-screenshot-script-repair/；入口 repair-result.json，实际图 historical-F02-interactive/F02-Word-OMML.png。下一步可以使用当前实际产物及明确构建来源，在正常用户执行环境继续自有隔离 Word 的视觉验收。没有运行新的转换矩阵或产品构建。

## CP235-S1 截图阻断核查（2026-10-03，独立于 CP235 打包）

本轮只读审查截图脚本及 Defender 日志，没有执行或修改截图代码。最新匹配事件仍是 2026-10-02 16:09 UTC（北京时间 10 月 3 日 00:09）的 AMSI 检测，名称 HackTool:PowerShell/EmpireGetScreenshot.C；1116/1117 分别记录检测及成功隔离，未给出具体命中行。CP234 没有新的截图尝试记录。本次查到两个历史检测/处置事件对，原脚本哈希仍为 64DEFB1B821F53FEB032327EA68205339B013BEAA8F0D53091D5E1F027542B2D；目录仍无 owned 文档、summary 或 PNG。

脚本第 121 行是实际 CopyFromScreen 调用，不能据此认定它就是命中特征。静态审查未见外传、下载、持久化安装或修改防护的代码；这不等于已确认误报。另有独立验收版本问题：脚本固定读取 CP198 R1 旧文档，只截 F05/F06/F02，不能验收 CP234。后续应先走正常安全复核，再为当前产物明确输入文档、构建哈希、公式集合和新输出目录；不通过改名、编码、更换执行路线或截图机制躲避检测。

补充历史对照：用户关于此前能截图的说明已由落盘证据确认。Capture-F20ShowAllOff.ps1 在北京时间 10 月 2 日 19:01 成功；与约 11 分钟后被拦的 CP198 脚本相比，第 10–198 行核心完全相同，仅输出名称、诊断标签及公式选择等有差异。旧 PNG 哈希与原 summary 一致，独立 Word 关闭退出、原文档未变。这支持正常验收流程的合法用途，与可能误报一致，不能将截图本身当成不允许的操作；具体检测触发点仍未知。本轮未执行新旧截图脚本。证据：cp235-screenshot-security-review/historical-success-comparison.json。

证据：artifacts/conversion-symbol-audit-20261002/cp235-screenshot-security-review/audit.json、defender-events.json。只新增审查证据及更新两份文档；未上传文件、修改安全设置、恢复隔离、启动 Word 或运行构建/转换/截图。最新已记录构建仍为 CP234，所有既有通过/失败范围不变，没有新增视觉通过。微软误报复核入口：https://www.microsoft.com/en-us/wdsi/filesubmission 。

状态（2026-10-03，CP225）：Companion 并发所有权修复保留；围栏结构、源识别及完整活动delimiter恢复合同已修正，最新2093纯测试通过，实际OMML原20项→VT通过生命周期。F02明确斜体在Word OMML中仍丢失；MTEF F07/F08/F09伸缩结构仍严格失败。历史CP220四过两失败矩阵不代表最新候选的完整六向验收。当前不可发布，性能工作仍暂停。

## 根因与改动

- 统一 MathML token 的声明样式与有效样式，替换编解码、源码恢复、OMML 各自推断默认值的逻辑。数学斜体与 TeX mathit 字母表来源分开处理
- 原生字体、样式和编码位置作为关联信息处理：未修改的原生字节保留；样式变更重新生成依赖字体的编码，防止旧 Symbol 槽位带入新字体
- OMML 在合并文本前按有效样式和角色分组，保留函数边界。没有修改 Office 安装目录的 XSL
- 共享精确长度契约用于源码恢复、MTEF 间距和 OMML 间距。已验证的 1em/2em 使用原生表示，未知宽度不近似舍入
- 扩充共同 glyph authority 的 shortmid、nshortmid、smallsetminus、varsubsetneq、varsupsetneq 数据；源码恢复也保留对应 alternate 身份
- 补齐保存重开后的逐项验证，分别记录数学内容、样式/字形/间距描述。内容签名不再被当作像素保真证明

## 已核实结果

- CP217 最终纯测试：1992/1992，0 失败、0 跳过；4 个原生运行方法明确排除。首轮 1983/1992 的 9 个旧错误角色预期经原生证据核对后最小更新，原失败 TRX 保留；原有编译/分析警告仍存在
- OMML spacing 定向测试：240/240，含 37 个新增案例
- 三份实际 Word 文档：1em、2em、零宽在插入、BuildUp、字号处理、保存重开后正确持久化并恢复；尚未测量实际像素间距
- 新五系列：40/40 原生编解码对照，122/122 相关测试，20/20 隔离 MathType 真实 WMF/PNG 检查；源码恢复 54/54 纯测试及 40/40 生产 MathJax/SVG 对照
- 精确源码宽度：68 项导出、102 项实际编辑器修改后源码往返；字母表回归另有 17 项生产 SVG 轮廓/位置/尺寸检查
- 阶段候选的 VT→MT、MT→VT、OMML→VT、OMML→MT 各 19 项分别完成转换、Undo/Redo 与重开逐项检查。VT 目标检查持久 LaTeX，MT 目标检查 MathML。这些不是最终统一候选的六向全验收
- CP217 交接时同一 C# 构建矩阵：VT→MT、MT→VT、OMML→MT 各 19 项通过转换、Undo/Redo、重开逐项核对；MT→OMML 已写入 19 项但 F02 的 4 个 italic 运算符读回不符，未进入 Undo/重开；VT→OMML 在 prewarm 因 43129 拒绝连接失败，未替换；OMML→VT 尚未运行。OMML 出向以实际已保存 OMML 为源，不能恢复先前丢失的意图

## 明确保留的未完成项

- F09 已通过 VT/OMML 定向三项双向生命周期，并纳入实际 OMML20→VT 通过集合；MTEF 的 F07/F08/F09 伸缩结构仍失败，不能称完整六向20通过
- VT→OMML、MT→OMML 的 italic 运算符仍未通过原始意图保真；不能从普通 Word 运算符的 Italic=-1 推断原始显式斜体意图
- Word 的普通原生运算符也会报告 Italic=-1；不能据此强行恢复用户的斜体意图。没有加入隐藏标记、替换码点或比较豁免
- U+2AFD varparallel 的 MathType 原生元组、单边 ceiling 崩溃及部分括号族，需要可靠原生样本；现有隔离渲染器没有可调用的导入入口
- vDash/models 原生元组冲突、部分 alternate 区别仍在清单。CP216 后续补齐 nmid/subsetneq/supsetneq 三条普通共享字形数据，9 个生成/读回 mi/mo 角色差异已通过定向测试；后续 9 项实际原生 WMF/PNG 的字体槽位、样式与普通/alternate 视觉区别也通过。旧泛用字体字节迁移及 Word 显示未验收
- 任意/负原生间距、其他尺寸/换行能力，以及多行复制 latex-expanded 的三个极端小数问题未修复
- CP235-S2 已在防护开启时恢复单例真实 Word 截图；最新构建完整 Word 截图、全量脚本和行内基线视觉验收仍未完成。CP217 的历史 Defender 阻断记录保留，不与本轮受限环境的 Word 登录会话错误混淆
- 原始 1479 项语料及全组合视觉检查不能由上述小范围通过替代

## CP234 时点的源码与构建（历史记录，最新见文首 CP250）

- 最新构建：artifacts/omml-edit-audit-20261001/build-cp234-hbar-source-contract，DLL SHA256 3F74E748DA6D44150AD655ED0A2C09EE46ECED41BC5C3A95CFA39ED253C057F1；FreezeR3，held R4未变。2209 managed通过，source20项alternate身份通过但有效样式仅9项通过；全六向/像素验收未完成
- 最新已验收Word DLL SHA256：4CE54B31FDA3B28F21EB4DDA42D7BDBA6F8DB0FC0378539F8FC5568719968CDF；历史CP220矩阵使用CP217的86900B64…1269F，不能混称同一构建
- 本地提交仍为 fed066192373fc12364e822f5478acca9967f7b0；之后源码、测试和文档修改尚未提交，包含新增共享文件
- 使用冻结 R3 resolver 的隔离构建；held R4 与既有无关 dirty 文件保留。没有发布、安装或推送本阶段修复

## 证据入口

- 完整阶段记录：docs/mathtype-history-checkpoint-20260928.md，CP194–CP220
- CP217 汇总：artifacts/conversion-symbol-audit-20261002/shared-contract-aggregate/cp217-stage-aggregate-r2.trx；原失败日志 cp217-stage-aggregate.trx
- CP218 接手与离线崩溃证据：artifacts/conversion-symbol-audit-20261002/cp218-handoff/。24412 转储确认 C0000005 写入 0x1010，异常栈在 Tauri 事件循环 plugins.lock；10532 转储在后台 HTTP 请求克隆 AppHandle 时进入 Tao 的 Rc<EventLoopRunner> 克隆。依赖的主线程上下文跨线程克隆非原子引用计数；CP219 已修正共享所有权边界，旧依赖并发回归崩溃、修正版三轮通过，CP220 同一进程完成全部矩阵
- runtime.json 的 24412 只是原始启动 PID；客户端后续自动启动过 80236、74084，接手时 74084 实际监听 43129 且 HTTPS health 通过。服务当下健康不等于崩溃已解决
- CP215 全量纯回归：artifacts/conversion-symbol-audit-20261002/shared-contract-aggregate/cp215-stage-aggregate.trx
- CP216 后续定向回归：149/149，artifacts/conversion-symbol-audit-20261002/alternate-five-tests/results/ordinary-authority-after-20261002T153200Z.trx。CP215 构建不包含这次三条后续数据修正
- CP216 新普通字形原生 9 项：artifacts/conversion-symbol-audit-20261002/ordinary-role-native9-20261002T153600Z/，325 ms，9/9 实际字体/槽位/样式及 PNG 对照通过
- Word 间距阶段 XML：artifacts/conversion-symbol-audit-20261002/cp215-word-space-probe/stages.json
- 新 glyph 原始图片：artifacts/conversion-symbol-audit-20261002/alternate-five-native20-20261002T143200Z/
- 新 glyph 综合证据：artifacts/conversion-symbol-audit-20261002/alternate-five-integrated-verification-20261002T143900Z.json
- 新 glyph 源码证据：artifacts/conversion-symbol-audit-20261002/alternate-five-source-verification-20261002T150000Z.json
- 独立 OMML 源检查：cp214-Omml-VisualTeX-semantic19、cp214-Omml-MathType-semantic19（同上 artifact 根目录）

下一步：继续 OMML italic 运算符表达边界和 F09 的共享结构/原生权威规则，并以当前实际产物和构建证据补齐 Word 视觉验收。CP235-S2 单例实际截图成功，正常用户执行环境可用；保持防护开启。可靠 MathType varparallel/单边括号样本仍是原生表示缺口的证据输入。CP220历史六方向结果不代表最新构建全矩阵通过。

## CP219 环境修复进展（2026-10-03）

共享 Tauri 运行时主线程上下文已改为 Arc 所有权，避免后台克隆 AppHandle 时操作 Tao 内部 Rc；固定原版本 2.11.4 通过仓库 vendor 补丁复现，不修改全局 Cargo 缓存。相同无窗口并发回归：旧依赖原生崩溃 C0000409；修正版 3/3 轮、48 万次克隆/释放和 24 次主线程回调通过。新 Companion PID46904/43129 使用全新隔离 profile，EXE SHA25681184AB7DEFFFFCE891E5C12FCE5B122BF82C49FED684F5CEEA84C0F61EE5891；Word DLL 仍为 CP217。F01/F02 新种子及 VT→MT 完整语义生命周期 2/2 通过；新 19 条种子已生成 19 个对象，后续矩阵正在验证。上述均不构成 Word 像素或全符号验收。证据：cp218-handoff、cp218-environment2、cp218-VisualTeX-MathType19（同一 artifacts/conversion-symbol-audit-20261002 根目录）。

## CP235 Windows NSIS 验收包

已按用户要求完成构建和最终解包静态校验，未安装。文件：artifacts/nsis-acceptance-cp235-20261003/VisualTeX_1.2.8_x64_CP235_Acceptance_20261003.exe；版本1.2.8，376914288字节；SHA256 B02C75CBDD09FD32B517DED6198B8916F4C1C7871338EF0CF92F57C722AA9929。
实际嵌入x64 Word与CP234完全一致；使用FreezeR3，held R4未动。两架构MSI各19项、23项Office前端、74项OCR资源和其他依赖均从最终包核对；Companion含已编译Arc补丁及最新前端。NSIS类型戳仅改变3字节，其他应用字节完全一致。报告：artifacts/nsis-acceptance-cp235-20261003/verified-embedded-r3/report.json；同目录README-ACCEPTANCE.md给出验收范围。安装包未数字签名，未改安全设置。未提交/发布，以下未决保真/性能问题没有因打包而自动解决。

## CP234 源码身份修正与未完成边界

普通U+210F恢复为hslash，token alternate=1恢复为hbar，仅修改现有普通/alternate共享表两条数据。25新增source-only案例与152定向相邻测试通过，最终managed2209通过/0失败/0跳过（四项原生仍排除）。实际20项C#→生产前端：scalar/mi角色/alternate及规范glyph身份从基线9项升至20项通过；有效样式两版均只9项通过，同样11项normal/bold-italic在现有命令字体处理边界丢失，无新增样式回归。完整严格回转从4/20升至9/20，仍失败，不宣称完整保真。证据hslash-hbar-source-comparison-cp234-20261002T2039.json。未把此源码映射改动算作native hbar/hslash字形修复。

当前剩余分类（CP231旧差异40项复算）：
- MTEF16项围栏/混合bra-ket：当前适配器模板/字形权威不足，需要合法原生结构样本或同等官方编码依据
- MTEF4项单位：原mi/mo载荷碰撞；Times+原MTCode的Ω native探针失败，需要确切字体/编码槽位权威，不能用µ单项成功推广
- MTEF4项hbar/hslash：样式签名与内在glyph设计关系未定；其中2项alternate仍有真实原生混同，需要区分字形依据
- OMML8项标点：normal mi/mo输出字节相同，reader不能恢复已丢意图
- OMML5项固定围栏：有结构恢复方向，但vendor token发射上下文尚因路径权限不可读，不做全局猜测
- 上述内容计数不包含全部alternate/spacing/像素能力缺口；F02显式italic mo与普通Word表示碰撞、命名/负间距等仍开放；截图功能已在CP235-S2恢复单例，完整视觉验收仍待进行

HEAD保持fed066192373fc12364e822f5478acca9967f7b0，无暂存/提交/发布。当前工作区有既有及本任务混合dirty和artifact；未清理或覆盖原有改动。

最小后续证据需求：
1. 可编辑MathType样本：单边/组/混合围栏、hbar/hslash不同字形，并保留原字体偏好；用于获取合法模板/槽位，不猜字节
2. 正体单位Ω/µ的可编辑原生样本或同字体/编码官方映射，须保留原码点与identifier含义；Times探针已反证，不能直接推广
3. 有权提供并置于授权工作区的reverse-token XSL片段，解除精确token发射点检查缺口；不绕开安装目录读取拒绝
4. 依据CP235-S2已验证的截图流程，用当前保存产物的显式清单继续自有隔离Word视觉验收；需要正常用户执行环境，保留防护及自有窗口检查

前端11项样式漏传已定位到共同hook覆盖范围，但尚无经验证的通用扩大判据；不将所有mi粗暴纳入，也不按ℏ单字符补丁。F02和标点表示碰撞并非保证有样本即可解决。当前所有已启动测试均终态，无重复矩阵排队；此阶段可独立且有充分依据的修正已完成，未解决项保持开放。

## CP231 等宽数学字母表修复已定向验证

已修复两个共同边界：前向漏发原生 m:scr=monospace/m:sty=p，反向直立数字默认规则误覆盖非Roman脚本。原基线17项中8失败；最终21专项、100定向/相邻、2184 managed汇总全部通过（0失败/跳过，保留原四项原生排除）。4份owned Word文档共16插入/BuildUp/字号/新进程重开阶段保持精确role/text/style：Latin x、数字09、混合普通字母、脚本中的monospace均持久。未用m:nor、字体覆盖、隐藏标记或比较豁免。证据cp231-word-monospace/stages.json及shared-contract-aggregate/cp231-final-integrated.trx。这不是Word像素或所有Unicode字符等宽支持证明。

CP230的普通run固定括号恢复仅完成只读方案；vendor XSL精确token发射位置因路径权限未读，未实施猜测性全局括号规则。CP232用精确CP231构建复算旧40项差异集合：40项均执行，direct MTEF仍24项差异，OMML由16降至13；C01293/C01470/C01471新一致，子集内无新增差异。C01434/35角色已修，fixed-fence差异仍在。证据cp232-residual-readonly/results.json；不是重新完整1476验收。单位Ω/µ的mi与mo在旧writer相同Symbol记录发生碰撞。显式字体合成20项虽全部完成渲染，实际Ω在普通/斜体/粗体/上标四情境均被MathType选成MT Extra/6E缺字图形，µ四情境能画；8项合成managed读回却均正确，表明纯回读不能代替字形。已否定这条Times资格扩展，未加µ单字符例外或修改生产资格。证据cp232-unit-font-probe/native20-exact。其余F02、原生围栏能力、alternate/spacing及截图阻塞边界仍保留。

## CP228 变体阶段结果；CP229脚本角色修复已定向验证

四个剩余subset/superset普通/变体族通过共享8条native＋4条source数据修复：32/32 codec与实际frontend、native20/20字体槽位/样式及逐张像素、2141/2141集成纯测试通过。20复合式VT↔MT双向转换、Undo/Redo、重开逐项也通过；另40份前向/重开记录的88个相关operator和48个alternate标记一致。最新验收DLL4CE54B31FDA3B28F21EB4DDA42D7BDBA6F8DB0FC0378539F8FC5568719968CDF。

同输入1476 managed差分无新增签名回归，只有这8个身份的24条native输出改变；MT alternate丢失68→56。MT→OMML丢失计数增加12，是上游现能保留而下游仍丢，OMML输出本身未改变。剩余MT24/OMML16已分类，不等于全部视觉缺陷。

CP229已修复脚本基底恢复把m:sty=p/m:nor一律当mi的共同角色错误。原18项基线13失败；扩充22项角色、混合结构、同文本不同样式对应与identifier负控全过，265相邻定向全过，最终managed汇总2163通过/0失败/0跳过（保留四项原生排除）。4份独立Word文档在插入、BuildUp、字号处理和新进程重开共16阶段，均保持严格role/text/effective-style：括号/双竖线mo，abc mi，普通文本符号mtext。首轮探针误把合法单子mrow包装当失败，已保留证据并仅修正探针的无属性单节点包装识别。证据cp229-word-script-role-r2/stages.json；未放宽产品比较器。两个脚本case仍有独立fixed-fence签名差异，不宣称全语料或全六向通过，也不代表可发布。

## CP225 已完成结果

完整活动delimiter合同已补齐：实际生产51词法入口、32非空码点加空侧，368/368生产检查与2093/2093集成纯测试通过（同4项native排除）。规范表仍只用于delimiter角色，普通运算符/正文不改变。

最新DLL D64871466F4E30CB2B43B9F3E0D534856C6646290C84FBA5EE9C3923245EA549。实际已保存OMML原20项→VT通过转换、Undo/Redo及重开逐项持久LaTeX，包含F09；这是从实际OMML恢复，不能恢复之前F02已丢的明确斜体意图，也不是Word像素验收。证据cp225-Omml-VisualTeX-actual20、shared-contract-aggregate/cp225-final-integrated.trx。

CP226合法native围栏审计已完成：3/3已成功渲染过的既有载荷、7/7围栏记录，经过读取与原prefix重写后selector/variation/alignment/顺序/CHAR字节都保留。没有证明现有默认样本丢失，故未加生产保存层；非默认样式/对齐仍未覆盖。CP227命名/负间距审查已完成：当前0.167/0.222/0.278/-0.167em都落到相同ASCII space并读回0.2em，原生精确advance合同缺失；已有EF02/EF22测量与简单别名映射不符，vendor importer后段不透明。因此未猜测码位或新增防御层。latex-expanded极端小数问题仅在独立多行复制路径，本任务不修改。下一步核对已有vendor权威的剩余subset/superset变体族。当前新源码未提交/安装，性能保持冻结。

## CP224 历史边界与后续纠正

F07写入回归已修复：原20项全部MTEF载荷与CP218逐字节相同，2088/2088纯测试通过。当前DLL524E1A22946507FCB20A2901FD95080CD45DAFD363E5396400B2ABE4ED0327A9。OMML原20项纯结构签名20/20；MTEF17/20，F07/F08/F09旧伸缩结构损失仍严格失败。

最新Word VT→OMML已写入原20项，仍因F02明确斜体失败；实际OMML→VT20在修改源前因AMS角括号的raw Unicode delimiter源码失败。先前28身份/146断言覆盖完整base包，不等于AllPackages实际启用集合。正在扩充统一delimiter合同到实际active handlers，保留此次失败。CP223三项双向生命周期通过仍是有效的小样证据。

## CP223 历史中间状态：新增围栏回归（CP224已纠正）

共享明确围栏识别与规范delimiter源码恢复已完成，2081/2081集成纯测试通过；原始F09加F12/F13三项在VT→OMML→VT两个方向完成转换、Undo/Redo、重开逐项检查。它们仅是定向闭环。

随后原20项纯审计发现：OMML20/20签名一致，MTEF17/20。F08/F09字节与旧版相同，严格识别揭示旧伸缩损失；F07字节却改变，源结构识别误入无原生模板能力的写入归一化，丢弃了corner的既有字体元组。这是实际回归，CP223当前构建不得视为可接受版本。正在将源语义识别与原生模板物化能力分开，保留严格失败和旧正确字节；不放宽比较器。详见CP223与cp223-representative20-pure。

## CP222 当前修复与验证

通用明确围栏行归一化已实现，2017/2017纯测试通过（同4项native排除）。CP222 DLL为3BAFE23A5A7DA25EC7334322067740DA618314EB4ED24041E03185C1E6AB3BB7。原始F09经生产投影进入Word，在插入、BuildUp、字号处理、保存重开后都保留2个m:d及原始组括号/arrowvert空右侧；像素未验收。源signature仍将结构识别与MTEF模板能力混合，正在按同一明确源结构合同修正，尚不称完整gate通过。

F02两项标准w:i对照已结束：搭配m:sty=i或省略m:sty，Word都在插入时规范化为与普通运算符相同的记录，重开不保留明确斜体意图。没有使用文本角色、二元运算符或隐藏数据替代。证据cp222-word-fence-structure、cp222-word-run-italic。

## CP221 接手后已确认结果

重新核对 HEAD、修正版 Companion、CP217 DLL 和 held R4 哈希一致；PID46904仍拥有43129，正常TLS认证和证书pin一致。没有重跑或覆盖Codex交接的六方向结果。

新增3份自有Word结构探针证明：组括号U27EE/U27EF、单侧arrowvert U23D0与空右侧、圆括号对照，均可通过标准mfenced→m:d在插入、BuildUp、字号处理及保存重开后保留准确围栏结构。证据cp221-word-fence-structure/stages.json。这不是像素验收；生产中的明确围栏归一化修正仍在实施。

MTEF仍缺完整F09模板权威：vendor的Euclid Symbol/BD/F8E6只证明arrowvertex延长组件，不能证明完整伸缩围栏；现有解码还按selector重建围栏。不能用新增字符表或替代selector冒充修复。F02的opEmu方案改变关系角色为二元运算符，未采用；两项标准w:i保存对照待执行。

## CP220 当前矩阵与边界

| 方向 | 当前结果 |
|---|---|
| VT→MathType | 19 条转换、Undo/Redo、保存重开逐项 MathML 通过；使用新生成种子 |
| MathType→VT | 19 条转换、Undo/Redo、保存重开逐项持久 LaTeX 通过 |
| OMML→MathType | 19 条转换、Undo/Redo、保存重开逐项 MathML 通过 |
| OMML→VT | 19 条转换、Undo/Redo、保存重开逐项持久 LaTeX 通过 |
| VT→OMML | 19 条写入完成；F02 四个 italic 运算符读回 normal，严格失败；未进入 Undo/重开 |
| MathType→OMML | 19 条写入完成；同一 F02 样式失败；未进入 Undo/重开 |

以上四项通过是代表集合的语义及持久化检查，不是全部符号、间距、变体或像素保真。两个 OMML 出向只对实际已保存源负责，不补回原始 VT 已丢的信息。F09 未进入 Word 批次；CP217 原始 20 项纯审计再次确认 MTEF/OMML 各 19 个 codec 签名相等、F09 仍失配，不能称 20/20。纯审计的 WordSignature 不可用（比较器编入验收 EXE），该字段无通过主张。

证据：artifacts/conversion-symbol-audit-20261002/cp218-handoff/matrix-summary.json、final-runtime-audit.json、cp218-representative20-pure/，以及 cp218-{源}-{目标}19 各目录。Word DLL、冻结 R3 及 held R4 哈希均再次核对一致。新 Companion PID46904 全程存活，未见启动后新增源码崩溃事件；原有三个 Word 进程及已安装 Companion 保留。C# 1992/1992 是本次核对的既有结果，没有声称重新执行。未提交、安装或推送；性能及截图限制保持。

## 2026-10-05 CP269–CP270 架构根因修复补充

共享FormulaXml已接入所列公开转换入口，解析时执行4M字符/256层/250k元素边界；移除不分所有权的全token零宽清理，并按mtext/ms文字角色保留字符/空白。新增23项离线检查通过；当前离线汇总2367/2368，仍保留CP255原字节失败，未扩展为全部符号或宿主视觉通过。最终frozen-R3 Word C1F790B2…C58FB3、PPT30957E9D…86D406，完整SHA/产物/根因在docs/architecture-root-remediation-20261005.md及cp270-final-state.json。旧source中已丢失的信息、本轮未运行的六向视觉、F02/F09/U+2AFD/固定组重音/复杂脚本/任意间距仍按本报告原范围保留。没有重开搁置探针、部署或提交。
## CP272 补充（2026-10-05）

共享MTEF读取已补齐完整根终止检查、原生递归256层和生成MathML树256层/250k元素边界，防止首根后内容被静默忽略及无界嵌套。新增独立字节测试修改前2/8、修改后8/8；相邻37/37；离线汇总2371/2372，四项未启用原生服务测试明确排除，唯一失败仍CP255原字节契约。精确net472新旧DLL只读解包归档19个MathType载荷，读回逐字19/19一致，不是新宿主/视觉矩阵。实际候选build-cp271-native-structure-final-r1，Word41F42D5F…A14F87、PPT8F9E7794…3215AC，完整证据和SHA见docs/architecture-root-remediation-20261005.md与artifacts/architecture-root-20261005/cp272-final-state.json。

固定组重音仍缺真正原生布局依据；本机厂商MathML输入规则也未提供保固定宽度的组模板，未据其扩展输出更改预期。用户已取消单次Undo阻塞要求；CP255、R08、F02/F09及原先六向视觉缺口没有据此关闭。此阶段未提交/部署/注册表/用户Office操作。
