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
              <strong>{isEn ? "VisualTeX not listed" : "未列出 VisualTeX"}</strong>
            </span>
          )}
        </div>
        <div className="powerpoint-addins-controls">
          <span className={loaded ? "" : "is-next-action"} aria-hidden="true"><Plus size={13} /></span>
          <span aria-hidden="true">−</span>

        </div>
      </div>

      <div className="powerpoint-guide-steps">
        <span><b>1</b>{isEn ? "Open Tools → PowerPoint Add-ins" : "打开“工具 → PowerPoint 加载项”"}</span>
        <span><b>2</b><Plus size={13} />{isEn ? "Click +" : "点击＋"}</span>
        <span><b>3</b><Check size={13} />{isEn ? "Check VisualTeX and restart PowerPoint" : "勾选 VisualTeX 并重启 PowerPoint"}</span>
        <span className="powerpoint-guide-path"><FolderOpen size={13} />~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam</span>
      </div>
    </div>
  );
}
