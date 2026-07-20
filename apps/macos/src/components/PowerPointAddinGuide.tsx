import { Check, FolderOpen, Plus } from "lucide-react";
import type { Language } from "../stores/editorStore";

interface Props {
  language: Language;
  compact?: boolean;
  loaded?: boolean;
}

export function PowerPointAddinGuide({ language, compact = false, loaded = false }: Props) {
  const isEn = language === "en";

  return (
    <div className={`powerpoint-native-guide${compact ? " is-compact" : ""}`}>
      <div className="powerpoint-native-app">
        <div className="powerpoint-native-titlebar">
          <span className="powerpoint-window-controls" aria-hidden="true">
            <i />
            <i />
            <i />
          </span>
          <strong>Microsoft PowerPoint</strong>
        </div>
        <div className="powerpoint-native-menubar">
          <span>PowerPoint</span>
          <span>{isEn ? "File" : "文件"}</span>
          <span>{isEn ? "Edit" : "编辑"}</span>
          <span>{isEn ? "View" : "视图"}</span>
          <span>{isEn ? "Insert" : "插入"}</span>
          <span className="is-active">{isEn ? "Tools" : "工具"}</span>
        </div>
        <div className="powerpoint-native-tools-menu">
          <span>{isEn ? "Spelling…" : "拼写…"}</span>
          <span>{isEn ? "Language…" : "语言…"}</span>
          <span className="is-highlighted">
            {isEn ? "PowerPoint Add-ins…" : "PowerPoint 加载项…"}
          </span>
        </div>
      </div>

      <div className="powerpoint-addins-dialog-mock">
        <header>
          <strong>{isEn ? "PowerPoint Add-ins" : "PowerPoint 加载项"}</strong>
        </header>
        <div className={`powerpoint-addins-list${loaded ? " is-loaded" : " is-empty"}`}>
          {loaded ? (
            <span className="is-selected">
              <i><Check size={12} /></i>
              <strong>VisualTeX</strong>
            </span>
          ) : (
            <span className="powerpoint-addins-empty-copy">
              <strong>{isEn ? "VisualTeX is not listed yet" : "此时列表中没有 VisualTeX 是正常的"}</strong>
              <small>{isEn ? "Click + below and select VisualTeX.ppam" : "请点击下方＋，再选择 VisualTeX.ppam"}</small>
            </span>
          )}
        </div>
        <div className="powerpoint-addins-controls">
          <span className={loaded ? "" : "is-next-action"} aria-hidden="true"><Plus size={13} /></span>
          <span aria-hidden="true">−</span>
          <small>
            {loaded
              ? isEn ? "VisualTeX is registered" : "VisualTeX 已登记"
              : isEn ? "Start here: click +" : "从这里开始：点击＋"}
          </small>
        </div>
      </div>

      <div className="powerpoint-guide-steps">
        <span><b>1</b>{isEn ? "Open Tools → PowerPoint Add-ins" : "打开“工具 → PowerPoint 加载项”"}</span>
        <span><b>2</b><Plus size={13} />{isEn ? "Click +. VisualTeX will not be in the list before this step" : "点击＋；执行这一步前，列表里不会出现 VisualTeX"}</span>
        <span><b>3</b><Check size={13} />{isEn ? "Keep VisualTeX checked, then restart PowerPoint" : "保持 VisualTeX 勾选，然后重启 PowerPoint"}</span>
        <span className="powerpoint-guide-path"><FolderOpen size={13} />~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam</span>
      </div>
    </div>
  );
}
