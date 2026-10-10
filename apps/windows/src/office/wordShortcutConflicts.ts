// Word's own bindings for the chords the Word formula shortcut setting accepts.
// The Word add-in always leaves these to Word (WordFormulaShortcutPolicy), so a
// formula shortcut on one of them never fires. Captured with FindKey from a
// default Normal.dotm (Word 16, Microsoft 365); categories are WdKeyCategory
// values. Keys a user customizes, or that macros and add-in templates such as
// MathType bind, are reported by the add-in at run time instead. Word reports
// Ctrl+Alt+T (™) as category 2; it is listed as a symbol so it reads naturally.
const WORD_DEFAULT_BINDINGS_DATA =
  "Ctrl+A=1:EditSelectAll|Ctrl+B=1:Bold|Ctrl+C=1:EditCopy|Ctrl+D=1:FormatFont|" +
  "Ctrl+E=1:CenterPara|Ctrl+F=1:SmartFind|Ctrl+G=1:EditGoTo|Ctrl+H=1:EditReplace|" +
  "Ctrl+I=1:Italic|Ctrl+J=1:JustifyPara|Ctrl+K=1:InsertHyperlink|Ctrl+L=1:LeftPara|" +
  "Ctrl+M=1:Indent|Ctrl+N=1:FileNewDefault|Ctrl+O=1:FileOpenUsingBackstage|" +
  "Ctrl+P=1:PrintPreviewAndPrint|Ctrl+Q=1:ResetPara|Ctrl+R=1:RightPara|Ctrl+S=1:FileSave|" +
  "Ctrl+T=1:HangingIndent|Ctrl+U=1:Underline|Ctrl+V=1:EditPaste|Ctrl+W=1:DocClose|" +
  "Ctrl+X=1:EditCut|Ctrl+Y=1:EditRedoOrRepeat|Ctrl+Z=1:EditUndo|Ctrl+0=1:OpenOrCloseUpPara|" +
  "Ctrl+1=1:SpacePara1|Ctrl+2=1:SpacePara2|Ctrl+3=1:LockFields|Ctrl+4=1:UnlockFields|" +
  "Ctrl+5=1:SpacePara15|Ctrl+6=1:UnlinkFields|Ctrl+F1=1:ToggleRibbon|" +
  "Ctrl+F2=1:PrintPreviewAndPrint|Ctrl+F3=1:Spike|Ctrl+F4=1:DocClose|Ctrl+F5=1:DocRestore|" +
  "Ctrl+F6=1:NextWindow|Ctrl+F8=1:DocSize|Ctrl+F9=1:InsertFieldChars|Ctrl+F10=1:DocMaximize|" +
  "Ctrl+F11=1:LockFields|Ctrl+F12=1:FileOpen|Ctrl+Shift+A=1:AllCaps|Ctrl+Shift+B=1:Bold|" +
  "Ctrl+Shift+C=1:CopyFormat|Ctrl+Shift+D=1:DoubleUnderline|" +
  "Ctrl+Shift+E=1:ToolsRevisionMarksToggle|Ctrl+Shift+F=1:Font|" +
  "Ctrl+Shift+G=1:ToolsWordCountList|Ctrl+Shift+H=1:Hidden|Ctrl+Shift+I=1:Italic|" +
  "Ctrl+Shift+J=1:DistributePara|Ctrl+Shift+K=1:SmallCaps|Ctrl+Shift+L=1:ApplyListBullet|" +
  "Ctrl+Shift+M=1:UnIndent|Ctrl+Shift+N=1:NormalStyle|Ctrl+Shift+O=1:ResearchLookup|" +
  "Ctrl+Shift+P=1:FontSizeSelect|Ctrl+Shift+Q=1:SymbolFont|" +
  "Ctrl+Shift+R=1:ToolsWordCountRecount|Ctrl+Shift+S=1:StyleApplyPane|Ctrl+Shift+T=1:UnHang|" +
  "Ctrl+Shift+U=1:Underline|Ctrl+Shift+V=1:PasteFormat|Ctrl+Shift+W=1:WordUnderline|" +
  "Ctrl+Shift+X=1:ToggleXMLTagView|Ctrl+Shift+Z=1:ResetChar|Ctrl+Shift+2=7:|Ctrl+Shift+6=7:|" +
  "Ctrl+Shift+7=7:|Ctrl+Shift+8=1:ShowAll|Ctrl+Shift+F3=1:InsertSpike|" +
  "Ctrl+Shift+F5=1:EditBookmark|Ctrl+Shift+F6=1:PrevWindow|Ctrl+Shift+F7=1:UpdateSource|" +
  "Ctrl+Shift+F8=1:ColumnSelect|Ctrl+Shift+F9=1:UnlinkFields|Ctrl+Shift+F10=1:WW2_RulerMode|" +
  "Ctrl+Shift+F11=1:UnlockFields|Ctrl+Shift+F12=1:PrintPreviewAndPrint|Ctrl+Alt+C=6:©|" +
  "Ctrl+Alt+D=1:InsertEndnoteNow|Ctrl+Alt+E=6:€|Ctrl+Alt+F=1:InsertFootnoteNow|" +
  "Ctrl+Alt+H=1:Highlight|Ctrl+Alt+I=1:PrintPreviewAndPrint|Ctrl+Alt+K=1:FormatAutoFormat|" +
  "Ctrl+Alt+L=1:InsertListNumField|Ctrl+Alt+M=1:InsertAnnotation|Ctrl+Alt+N=1:ViewNormal|" +
  "Ctrl+Alt+O=1:ViewOutline|Ctrl+Alt+P=1:ViewPage|Ctrl+Alt+R=6:®|Ctrl+Alt+S=1:DocSplit|" +
  "Ctrl+Alt+T=6:™|Ctrl+Alt+U=1:TableUpdateAutoFormat|Ctrl+Alt+V=1:EditPasteSpecial|" +
  "Ctrl+Alt+Y=1:RepeatFind|Ctrl+Alt+Z=1:GoBack|Ctrl+Alt+1=1:ApplyHeading1|" +
  "Ctrl+Alt+2=1:ApplyHeading2|Ctrl+Alt+3=1:ApplyHeading3|Ctrl+Alt+5=1:GoToObjectLayer|" +
  "Ctrl+Alt+F1=1:MicrosoftSystemInfo|Ctrl+Alt+F2=1:FileOpen|Ctrl+Alt+F7=1:ToolsHHC|" +
  "Ctrl+Alt+Shift+S=1:FormattingPane|Ctrl+Alt+Shift+V=1:InsertAutoText|" +
  "Ctrl+Alt+Shift+1=6:¡|Ctrl+Alt+Shift+6=7:";

// Chinese names as Word's Chinese UI shows them; other commands keep Word's name.
const WORD_COMMAND_NAMES_ZH: Readonly<Record<string, string>> = {
  EditSelectAll: "全选",
  Bold: "加粗",
  EditCopy: "复制",
  FormatFont: "字体",
  CenterPara: "居中",
  SmartFind: "查找（导航窗格）",
  EditGoTo: "定位",
  EditReplace: "替换",
  Italic: "倾斜",
  JustifyPara: "两端对齐",
  InsertHyperlink: "插入链接",
  LeftPara: "左对齐",
  Indent: "增加缩进",
  FileNewDefault: "新建文档",
  FileOpenUsingBackstage: "打开",
  FileOpen: "打开",
  PrintPreviewAndPrint: "打印",
  ResetPara: "清除段落格式",
  RightPara: "右对齐",
  FileSave: "保存",
  HangingIndent: "悬挂缩进",
  Underline: "下划线",
  EditPaste: "粘贴",
  DocClose: "关闭文档",
  EditCut: "剪切",
  EditRedoOrRepeat: "恢复",
  EditUndo: "撤销",
  OpenOrCloseUpPara: "段前间距",
  SpacePara1: "单倍行距",
  SpacePara2: "双倍行距",
  SpacePara15: "1.5 倍行距",
  LockFields: "锁定域",
  UnlockFields: "解除域锁定",
  UnlinkFields: "取消域链接",
  ToggleRibbon: "折叠功能区",
  InsertFieldChars: "插入域",
  AllCaps: "全部大写",
  CopyFormat: "复制格式",
  DoubleUnderline: "双下划线",
  ToolsRevisionMarksToggle: "修订",
  Font: "字体",
  ToolsWordCountList: "字数统计",
  Hidden: "隐藏文字",
  DistributePara: "分散对齐",
  SmallCaps: "小型大写字母",
  ApplyListBullet: "项目符号",
  UnIndent: "减少缩进",
  NormalStyle: "正文样式",
  FontSizeSelect: "字号",
  SymbolFont: "Symbol 字体",
  StyleApplyPane: "应用样式",
  UnHang: "取消悬挂缩进",
  PasteFormat: "粘贴格式",
  ResetChar: "清除字符格式",
  ShowAll: "显示编辑标记",
  EditBookmark: "书签",
  InsertEndnoteNow: "插入尾注",
  InsertFootnoteNow: "插入脚注",
  Highlight: "突出显示",
  InsertAnnotation: "插入批注",
  ViewNormal: "草稿视图",
  ViewOutline: "大纲视图",
  ViewPage: "页面视图",
  DocSplit: "拆分窗口",
  EditPasteSpecial: "选择性粘贴",
  RepeatFind: "重复查找",
  GoBack: "返回上次编辑位置",
  ApplyHeading1: "标题 1",
  ApplyHeading2: "标题 2",
  ApplyHeading3: "标题 3",
  FormattingPane: "显示格式",
  InsertAutoText: "自动图文集",
};

export interface WordShortcutBinding {
  /** WdKeyCategory: 0 disabled, 1 command, 2 macro, 3 font, 4 AutoText, 5 style, 6 symbol, 7 prefix. */
  category: number;
  name: string;
}

function parseBindings(data: string, separator: string) {
  const bindings = new Map<string, WordShortcutBinding>();
  for (const part of data.split(separator)) {
    const match = /^([^=]+)=(-?\d+):(.*)$/.exec(part.trim());
    if (!match) continue;
    const category = Number(match[2]);
    if (category < 0) continue;
    bindings.set(match[1], { category, name: match[3] });
  }
  return bindings;
}

const WORD_DEFAULT_BINDINGS = parseBindings(WORD_DEFAULT_BINDINGS_DATA, "|");

/** Shortcuts no default Word binding uses, offered when a capture is refused. */
export const WORD_FREE_SHORTCUT_SUGGESTIONS = [
  "Ctrl+Alt+Shift+I",
  "Ctrl+Alt+Shift+D",
  "Ctrl+Alt+Shift+M",
  "Ctrl+F7",
] as const;

export function wordDefaultBinding(shortcut: string | null) {
  return shortcut ? WORD_DEFAULT_BINDINGS.get(shortcut) ?? null : null;
}

/** Parses the add-in's WordFormulaShortcutConflicts registry value. */
export function parseWordShortcutConflicts(value: unknown) {
  return parseBindings(typeof value === "string" ? value : "", ";");
}

export function describeWordShortcutBinding(binding: WordShortcutBinding, isEn: boolean) {
  const name = binding.name;
  switch (binding.category) {
    case 0:
      return isEn ? "Word (the key is disabled in Word)" : "Word（此键在 Word 中被设为禁用）";
    case 2:
      return isEn ? `the Word macro “${name}”` : `Word 宏「${name}」`;
    case 3:
      return isEn ? `the Word font “${name}”` : `Word 字体「${name}」`;
    case 4:
      return isEn ? `the Word AutoText “${name}”` : `Word 自动图文集「${name}」`;
    case 5:
      return isEn ? `the Word style “${name}”` : `Word 样式「${name}」`;
    case 6:
      return isEn ? `the Word symbol shortcut “${name}”` : `Word 插入符号「${name}」`;
    case 7:
      return isEn
        ? "a Word two-key sequence (accented letters)"
        : "Word 两段式快捷键（输入带重音的字母）";
    default:
      if (!name) return isEn ? "a Word command" : "Word 命令";
      return isEn
        ? `the Word command “${name}”`
        : `Word 的「${WORD_COMMAND_NAMES_ZH[name] ?? name}」`;
  }
}
