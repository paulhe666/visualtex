// Export to Word: the formulas' form — Word equations (OMML), VisualTeX formulas (pictures that VisualTeX opens
// again), MathType objects (Windows, with MathType installed). The choice is remembered.
import { useState } from "react";
import { createPortal } from "react-dom";
import { useDocumentFormatStore, type WordFormulaFormat } from "../stores/documentFormatStore";

type Lang = "en" | "zh";

export const WORD_FORMULA_OPTIONS: Array<{ id: WordFormulaFormat; en: string; zh: string; windowsOnly?: boolean }> = [
  { id: "omml", en: "Word equations (OMML)", zh: "Word 公式（OMML）" },
  { id: "visualtex", en: "VisualTeX formulas", zh: "VisualTeX 公式" },
  { id: "mathtype", en: "MathType formulas", zh: "MathType 公式", windowsOnly: true },
];

export const isWindowsApp = () => typeof navigator !== "undefined" && /Windows/i.test(navigator.userAgent);

/** MathType objects need the app's MathType writer (not in this build yet). */
const MATHTYPE_EXPORT = false;

export function wordFormulaOptions() {
  const windows = isWindowsApp() && MATHTYPE_EXPORT;
  return WORD_FORMULA_OPTIONS.filter((option) => windows || !option.windowsOnly);
}

export function WordExportDialog(props: { lang: Lang; busy: string | null; onExport: (format: WordFormulaFormat) => void; onClose: () => void }) {
  const { lang, busy } = props;
  const options = wordFormulaOptions();
  const stored = useDocumentFormatStore((state) => state.wordFormulas);
  const [format, setFormat] = useState<WordFormulaFormat>(options.some((o) => o.id === stored) ? stored : "omml");
  const t = (en: string, zh: string) => (lang === "en" ? en : zh);
  return createPortal(
    <>
      <div className="document-control-dismiss vtx-export-dismiss" onMouseDown={() => !busy && props.onClose()} />
      <form
        className="document-control-panel vtx-word-export"
        role="dialog"
        aria-label={t("Export Word", "导出 Word")}
        onSubmit={(event) => {
          event.preventDefault();
          if (busy) return;
          useDocumentFormatStore.getState().setWordFormulas(format);
          props.onExport(format);
        }}
        onKeyDown={(event) => {
          if (event.key === "Escape" && !busy) {
            event.preventDefault();
            props.onClose();
          }
        }}
      >
        <div className="vtx-word-export-title">{t("Export Word", "导出 Word")}</div>
        <div className="vtx-word-export-label">{t("Formulas", "公式")}</div>
        <div className="vtx-word-export-options" role="radiogroup">
          {options.map((option) => (
            <label key={option.id} className={option.id === format ? "is-active" : ""}>
              <input type="radio" name="vtx-word-formulas" checked={option.id === format} disabled={!!busy} onChange={() => setFormat(option.id)} />
              <span>{lang === "en" ? option.en : option.zh}</span>
            </label>
          ))}
        </div>
        <div className="vtx-link-dialog-actions">
          <span className="vtx-word-export-status">{busy ?? ""}</span>
          <button type="button" className="secondary-button" disabled={!!busy} onClick={props.onClose}>
            {t("Cancel", "取消")}
          </button>
          <button type="submit" className="primary-small-button" disabled={!!busy} autoFocus>
            {t("Export", "导出")}
          </button>
        </div>
      </form>
    </>,
    document.body,
  );
}
