import { Clock3, Trash2, X } from "lucide-react";
import { MathPreview } from "./MathPreview";
import { useEditorStore } from "../stores/editorStore";

interface Props {
  open: boolean;
  onClose: () => void;
  onRestore: (latex: string) => void;
}

export function HistoryPanel({ open, onClose, onRestore }: Props) {
  const history = useEditorStore((state) => state.history);
  const clearHistory = useEditorStore((state) => state.clearHistory);
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";

  const timeLabel = (time: number) =>
    new Intl.DateTimeFormat(isEn ? "en-US" : "zh-CN", {
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }).format(time);

  return (
    <aside className={"history-panel " + (open ? "is-open" : "")} aria-hidden={!open}>
      <header className="history-header">
        <div>
          <span className="eyebrow">RECENT</span>
          <h2>{isEn ? "Formula history" : "公式历史"}</h2>
        </div>
        <button
          type="button"
          className="icon-button"
          onClick={onClose}
          aria-label={isEn ? "Close history" : "关闭历史"}
        >
          <X size={18} />
        </button>
      </header>

      <div className="history-list">
        {history.length === 0 ? (
          <div className="empty-state">
            <span className="empty-icon"><Clock3 size={24} /></span>
            <h3>{isEn ? "No formula history yet" : "还没有历史公式"}</h3>
            <p>
              {isEn
                ? "The current formula will be saved here after you pause editing."
                : "停止编辑片刻后，当前公式会自动保存在这里。"}
            </p>
          </div>
        ) : (
          history.map((item) => {
            const itemLines = item.latex.split("\n").filter((line) => line.trim());
            return (
              <button
                type="button"
                className="history-item"
                key={item.id}
                onClick={() => onRestore(item.latex)}
              >
                <span className="history-formula-stack">
                  {itemLines.slice(0, 3).map((line, index) => (
                    <MathPreview latex={line} key={index} />
                  ))}
                  {itemLines.length > 3 && (
                    <small>+{itemLines.length - 3} {isEn ? "lines" : "行"}</small>
                  )}
                </span>
                <span>{timeLabel(item.createdAt)}</span>
              </button>
            );
          })
        )}
      </div>

      {history.length > 0 && (
        <footer className="history-footer">
          <button type="button" className="text-button danger-text" onClick={clearHistory}>
            <Trash2 size={14} /> {isEn ? "Clear history" : "清空历史"}
          </button>
        </footer>
      )}
    </aside>
  );
}
