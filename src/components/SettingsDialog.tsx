import {
  BrainCircuit,
  Languages,
  Moon,
  RotateCcw,
  SlidersHorizontal,
  Sun,
  X,
} from "lucide-react";
import { useEditorStore } from "../stores/editorStore";

interface Props {
  open: boolean;
  onClose: () => void;
}

export function SettingsDialog({ open, onClose }: Props) {
  const theme = useEditorStore((state) => state.theme);
  const setTheme = useEditorStore((state) => state.setTheme);
  const language = useEditorStore((state) => state.language);
  const setLanguage = useEditorStore((state) => state.setLanguage);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const personalize = useEditorStore((state) => state.personalize);
  const setPersonalize = useEditorStore((state) => state.setPersonalize);
  const suggestionCount = useEditorStore((state) => state.suggestionCount);
  const setSuggestionCount = useEditorStore((state) => state.setSuggestionCount);
  const resetUsage = useEditorStore((state) => state.resetUsage);
  const isEn = language === "en";

  if (!open) return null;

  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        className="settings-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="settings-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <span className="eyebrow">PREFERENCES</span>
            <h2 id="settings-title">{isEn ? "Settings" : "设置"}</h2>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={onClose}
            aria-label={isEn ? "Close settings" : "关闭设置"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="settings-content">
          <div className="settings-section">
            <div className="settings-section-title">
              <BrainCircuit size={18} />
              <div>
                <h3>{isEn ? "Personalized commands" : "个性化命令推荐"}</h3>
                <p>
                  {isEn
                    ? "Rank suggestions using frequency, accepted prefixes and recency."
                    : "根据使用频率、前缀选择和最近使用时间调整候选顺序。"}
                </p>
              </div>
            </div>
            <label className="switch-row">
              <span>
                <strong>{isEn ? "Enable personalized ranking" : "启用个性化排序"}</strong>
                <small>
                  {isEn ? "Turn off to restore the default order" : "关闭后恢复系统默认顺序"}
                </small>
              </span>
              <input
                type="checkbox"
                checked={personalize}
                onChange={(event) => setPersonalize(event.target.checked)}
              />
              <span className="switch-control" />
            </label>
            <label className="range-setting">
              <span>
                <strong>{isEn ? "Suggestion count" : "候选项数量"}</strong>
                <small>
                  {suggestionCount} {isEn ? "items" : "项"}
                </small>
              </span>
              <input
                type="range"
                min="3"
                max="10"
                value={suggestionCount}
                onChange={(event) => setSuggestionCount(Number(event.target.value))}
              />
            </label>
            <button
              type="button"
              className="secondary-button danger-subtle"
              onClick={resetUsage}
            >
              <RotateCcw size={15} />
              {isEn ? "Reset recommendation history" : "重置推荐记录"}
            </button>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <SlidersHorizontal size={18} />
              <div>
                <h3>{isEn ? "Appearance & editor" : "外观与编辑"}</h3>
                <p>
                  {isEn
                    ? "Appearance settings are saved automatically."
                    : "外观设置会自动保存在当前设备。"}
                </p>
              </div>
            </div>
            <div className="theme-segment">
              <button
                type="button"
                className={theme === "light" ? "is-active" : ""}
                onClick={() => setTheme("light")}
              >
                <Sun size={16} /> {isEn ? "Light" : "浅色"}
              </button>
              <button
                type="button"
                className={theme === "dark" ? "is-active" : ""}
                onClick={() => setTheme("dark")}
              >
                <Moon size={16} /> {isEn ? "Dark" : "深色"}
              </button>
            </div>
            <label className="range-setting">
              <span>
                <strong>{isEn ? "Formula zoom" : "公式显示缩放"}</strong>
                <small>{Math.round(zoom * 100)}%</small>
              </span>
              <input
                type="range"
                min="0.7"
                max="1.6"
                step="0.1"
                value={zoom}
                onChange={(event) => setZoom(Number(event.target.value))}
              />
            </label>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Languages size={18} />
              <div>
                <h3>{isEn ? "Interface language" : "界面语言"}</h3>
                <p>{isEn ? "Switch between English and Chinese." : "切换中文或英文界面。"}</p>
              </div>
            </div>
            <div className="theme-segment">
              <button
                type="button"
                className={language === "cn" ? "is-active" : ""}
                onClick={() => setLanguage("cn")}
              >
                中文
              </button>
              <button
                type="button"
                className={language === "en" ? "is-active" : ""}
                onClick={() => setLanguage("en")}
              >
                English
              </button>
            </div>
          </div>
        </div>

        <footer className="dialog-footer">
          <span>{isEn ? "Settings saved automatically" : "设置已自动保存"}</span>
          <button type="button" className="primary-button" onClick={onClose}>
            {isEn ? "Done" : "完成"}
          </button>
        </footer>
      </section>
    </div>
  );
}
