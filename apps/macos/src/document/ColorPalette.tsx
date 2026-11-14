// The formula editor's colour palette (EditorWorkspace's text and background colour popover) for the document:
// the same fixed colours, the same custom colours — kept in the same storage, so a colour saved in one editor is
// there in the other — and the same markup and look.
import { useLayoutEffect, useRef, useState, type CSSProperties } from "react";
import { createPortal } from "react-dom";
import { Highlighter, Palette, Plus, X } from "lucide-react";
import { normalizeCustomFormulaColor } from "../workspace/formulaColor";

export type ColorKind = "color" | "backgroundColor";
type Lang = "en" | "zh";
const t = (lang: Lang, en: string, zh: string) => (lang === "en" ? en : zh);

const PRESETS: Record<ColorKind, readonly string[]> = {
  color: ["#111827", "#dc2626", "#ea580c", "#ca8a04", "#16a34a", "#0891b2", "#2563eb", "#7c3aed", "#db2777"],
  backgroundColor: ["#fef3c7", "#fed7aa", "#fecaca", "#fbcfe8", "#ddd6fe", "#bfdbfe", "#a5f3fc", "#bbf7d0", "#e5e7eb"],
};
const STORAGE: Record<ColorKind, string> = {
  color: "visualtex-custom-formula-text-colors",
  backgroundColor: "visualtex-custom-formula-background-colors",
};
const MAXIMUM = 8;

function loadCustom(kind: ColorKind): string[] {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE[kind]) ?? "[]");
    if (!Array.isArray(parsed)) return [];
    return Array.from(new Set(parsed.map(normalizeCustomFormulaColor).filter(Boolean))).slice(0, MAXIMUM) as string[];
  } catch {
    return [];
  }
}

function saveCustom(kind: ColorKind, colors: string[]) {
  try {
    localStorage.setItem(STORAGE[kind], JSON.stringify(colors));
  } catch {
    // The custom colours still work for this session.
  }
}

/**
 * A colour button with the formula editor's palette: fixed colours, "+" to save a colour of one's own, the saved
 * colours (each can be deleted), and for backgrounds "no background". The editor keeps the focus.
 */
export function ColorMenuButton(props: {
  kind: ColorKind;
  lang: Lang;
  /** the colour of the selection now (drawn under the icon) */
  value: string | null;
  label: string;
  title: string;
  onApply: (color: string | null) => void;
}) {
  const { kind, lang, value, label, title, onApply } = props;
  const [open, setOpen] = useState(false);
  const [custom, setCustom] = useState<string[]>(() => loadCustom(kind));
  const [picker, setPicker] = useState(value && /^#[0-9a-f]{6}$/i.test(value) ? value : PRESETS[kind][0]);
  const draft = useRef<string | null>(null);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const [place, setPlace] = useState<{ left: number; top: number } | null>(null);

  useLayoutEffect(() => {
    if (!open) {
      setPlace(null);
      return;
    }
    setCustom(loadCustom(kind));
    const button = buttonRef.current?.getBoundingClientRect();
    const panel = panelRef.current;
    if (!button || !panel) return;
    const below = window.innerHeight - button.bottom - 7;
    const top = below >= panel.offsetHeight || button.top < panel.offsetHeight + 7 ? button.bottom + 7 : button.top - panel.offsetHeight - 7;
    const left = Math.max(4, Math.min(button.right - panel.offsetWidth, window.innerWidth - panel.offsetWidth - 4));
    setPlace({ left, top: Math.max(4, Math.min(top, window.innerHeight - panel.offsetHeight - 4)) });
  }, [open, kind]);

  const apply = (color: string | null) => {
    setOpen(false);
    onApply(color);
  };
  const save = (raw: string) => {
    const color = normalizeCustomFormulaColor(raw);
    if (!color) return;
    setPicker(color);
    if (PRESETS[kind].includes(color)) return;
    const previous = draft.current;
    const next = [color, ...custom.filter((c) => c !== color && c !== previous)].slice(0, MAXIMUM);
    draft.current = color;
    setCustom(next);
    saveCustom(kind, next);
  };
  const remove = (color: string) => {
    const next = custom.filter((c) => c !== color);
    setCustom(next);
    saveCustom(kind, next);
    if (draft.current === color) draft.current = null;
  };
  const keep = (event: { preventDefault: () => void }) => event.preventDefault();

  return (
    <>
      <button
        ref={buttonRef}
        type="button"
        className={"icon-button compact formula-formatting-button is-color-action" + (open ? " is-active" : "")}
        style={{ "--formula-format-color": value ?? (kind === "color" ? "#111827" : "transparent") } as CSSProperties}
        aria-label={label}
        title={title}
        aria-pressed={open}
        onMouseDown={keep}
        onClick={() => setOpen((o) => !o)}
      >
        {kind === "color" ? <Palette size={15} strokeWidth={2} /> : <Highlighter size={15} strokeWidth={2} />}
      </button>
      {open &&
        createPortal(
          <>
            <div className="document-control-dismiss" onMouseDown={() => setOpen(false)} />
            <div
              ref={panelRef}
              className="formula-color-popover"
              style={{ position: "fixed", ...(place ?? { left: -9999, top: -9999 }) }}
              data-visualtex-floating-layer
              role="dialog"
              aria-label={label}
            >
              <strong>{kind === "color" ? t(lang, "Text color", "字体颜色") : t(lang, "Background color", "背景颜色")}</strong>
              <div className="formula-color-content">
                <section className="formula-color-presets">
                  <span className="formula-color-section-label">{t(lang, "Preset", "固定颜色")}</span>
                  <button type="button" className="formula-color-none-action" onMouseDown={keep} onClick={() => apply(null)}>
                    <span className="formula-color-none-icon" aria-hidden="true" />
                    <span>{kind === "color" ? t(lang, "Default", "默认颜色") : t(lang, "None", "无背景")}</span>
                  </button>
                  <div className="formula-color-swatches" role="group">
                    {PRESETS[kind].map((color) => (
                      <button
                        key={color}
                        type="button"
                        className="formula-color-swatch"
                        style={{ backgroundColor: color }}
                        aria-label={`${t(lang, "Use", "使用")} ${color}`}
                        title={color}
                        onMouseDown={keep}
                        onClick={() => apply(color)}
                      />
                    ))}
                    <label className="formula-custom-color" title={t(lang, "Add custom color", "添加自定义颜色")}>
                      <input
                        type="color"
                        value={picker}
                        aria-label={t(lang, "Add custom color", "添加自定义颜色")}
                        onPointerDown={() => {
                          draft.current = null;
                        }}
                        onInput={(event) => save(event.currentTarget.value)}
                        onChange={(event) => save(event.currentTarget.value)}
                      />
                      <Plus size={13} />
                    </label>
                  </div>
                </section>
                <section className="formula-custom-colors-panel">
                  <div className="formula-custom-colors-heading">
                    <span>{t(lang, "Custom", "自定义颜色")}</span>
                    <small>
                      {custom.length}/{MAXIMUM}
                    </small>
                  </div>
                  {custom.length > 0 ? (
                    <div className="formula-custom-colors-grid">
                      {custom.map((color) => (
                        <div key={color} className="formula-custom-color-item">
                          <button
                            type="button"
                            className="formula-color-swatch"
                            style={{ backgroundColor: color }}
                            aria-label={`${t(lang, "Use custom", "使用自定义颜色")} ${color}`}
                            title={color}
                            onMouseDown={keep}
                            onClick={() => apply(color)}
                          />
                          <button
                            type="button"
                            className="formula-custom-color-delete"
                            aria-label={`${t(lang, "Delete custom", "删除自定义颜色")} ${color}`}
                            title={t(lang, "Delete", "删除")}
                            onPointerDown={(event) => {
                              event.preventDefault();
                              event.stopPropagation();
                            }}
                            onClick={(event) => {
                              event.stopPropagation();
                              remove(color);
                            }}
                          >
                            <X size={9} strokeWidth={2.4} />
                          </button>
                        </div>
                      ))}
                    </div>
                  ) : null}
                </section>
              </div>
            </div>
          </>,
          document.body,
        )}
    </>
  );
}
