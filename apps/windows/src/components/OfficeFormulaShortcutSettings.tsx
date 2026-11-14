import { useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { Keyboard, X } from "lucide-react";
import {
  officeFormulaShortcutFromEvent,
  type OfficeFormulaShortcutCapture,
} from "../office/officeFormulaShortcuts";
import {
  describeWordShortcutBinding,
  parseWordShortcutConflicts,
  WORD_FREE_SHORTCUT_SUGGESTIONS,
  wordDefaultBinding,
  type WordShortcutBinding,
} from "../office/wordShortcutConflicts";
import { ANYWHERE_HOTKEY_ACTIONS, anywhereHotkeyLabel } from "../shortcuts/anywhereHotkeys";
import { useAnywhereHotkeyStore } from "../stores/anywhereHotkeyStore";
import { useEditorStore } from "../stores/editorStore";
import type {
  PowerPointFormulaShortcutAction,
  PowerPointFormulaShortcutSettings,
  WordFormulaShortcutSettings,
} from "../../../shared/visualtex-app/src/types/formula";

type WordShortcutAction = keyof WordFormulaShortcutSettings;
/** Word and WPS Writer: inline and display formulas with a format each. */
type DocumentHost = "word" | "wpsWriter";
/** PowerPoint and WPS Presentation: one shortcut per formula action. */
type SlideHost = "powerPoint" | "wpsPresentation";
type ShortcutTarget =
  | { host: DocumentHost; action: WordShortcutAction }
  | { host: SlideHost; action: PowerPointFormulaShortcutAction };

const HOST_NAMES: Record<DocumentHost | SlideHost, { en: string; zh: string }> = {
  word: { en: "Word", zh: "Word" },
  wpsWriter: { en: "WPS Writer", zh: "WPS 文字" },
  powerPoint: { en: "PowerPoint", zh: "PowerPoint" },
  wpsPresentation: { en: "WPS Presentation", zh: "WPS 演示" },
};

const powerPointActions: ReadonlyArray<{
  action: PowerPointFormulaShortcutAction;
  labelEn: string;
  labelZh: string;
}> = [
  {
    action: "svg",
    labelEn: "New SVG formula",
    labelZh: "新建 SVG 公式",
  },
  {
    action: "ole",
    labelEn: "New OLE formula",
    labelZh: "新建 OLE 公式",
  },
  {
    action: "inlineOmml",
    labelEn: "Insert inline OMML formula",
    labelZh: "插入行内 OMML 公式",
  },
  {
    action: "displayOmml",
    labelEn: "Insert display OMML formula",
    labelZh: "插入行间 OMML 公式",
  },
];

const wordRows: ReadonlyArray<{
  action: WordShortcutAction;
  labelEn: string;
  labelZh: string;
}> = [
  { action: "inline", labelEn: "Insert inline formula", labelZh: "插入行内公式" },
  { action: "display", labelEn: "Insert display formula", labelZh: "插入行间公式" },
];

function captureErrorText(
  result: Extract<OfficeFormulaShortcutCapture, { error: string }>,
  isEn: boolean,
) {
  switch (result.error) {
    case "modifier":
      return isEn
        ? "Requires Ctrl or Alt."
        : "须包含 Ctrl 或 Alt。";
    case "reserved":
      return isEn
        ? "Ctrl+Alt+O is reserved for Quick OCR."
        : "Ctrl+Alt+O 已用于快速 OCR。";
    default:
      return isEn
        ? "Ctrl or Alt + letter, digit or F1–F24."
        : "Ctrl 或 Alt + 字母、数字或 F1–F24。";
  }
}

function sameTarget(left: ShortcutTarget | null, right: ShortcutTarget) {
  return left?.host === right.host && left.action === right.action;
}

function isDocumentHost(host: ShortcutTarget["host"]): host is DocumentHost {
  return host === "word" || host === "wpsWriter";
}

export function OfficeFormulaShortcutSettingsSection() {
  const language = useEditorStore((state) => state.language);
  const wordSettings = useEditorStore((state) => state.wordFormulaShortcuts);
  const wpsWriterSettings = useEditorStore((state) => state.wpsWriterFormulaShortcuts);
  const powerPointSettings = useEditorStore(
    (state) => state.powerPointFormulaShortcuts,
  );
  const wpsPresentationSettings = useEditorStore(
    (state) => state.wpsPresentationFormulaShortcuts,
  );
  const setWordShortcut = useEditorStore(
    (state) => state.setWordFormulaShortcut,
  );
  const setWpsWriterShortcut = useEditorStore(
    (state) => state.setWpsWriterFormulaShortcut,
  );
  const setPowerPointShortcut = useEditorStore(
    (state) => state.setPowerPointFormulaShortcut,
  );
  const setWpsPresentationShortcut = useEditorStore(
    (state) => state.setWpsPresentationFormulaShortcut,
  );
  const anywhereBindings = useAnywhereHotkeyStore((state) => state.bindings);
  const [capturing, setCapturing] = useState<ShortcutTarget | null>(null);
  const [message, setMessage] = useState("");
  // The row the message is about: shown under it, where the keys were pressed
  // (at the section's end it was below the window's edge).
  const [messageTarget, setMessageTarget] = useState<ShortcutTarget | null>(null);
  const [wordReportedConflicts, setWordReportedConflicts] = useState(
    () => new Map<string, WordShortcutBinding>(),
  );
  const isEn = language === "en";

  const documentSettings = (host: DocumentHost): WordFormulaShortcutSettings =>
    host === "word" ? wordSettings : wpsWriterSettings;
  const slideSettings = (host: SlideHost): PowerPointFormulaShortcutSettings =>
    host === "powerPoint" ? powerPointSettings : wpsPresentationSettings;
  const hostName = (host: ShortcutTarget["host"]) =>
    isEn ? HOST_NAMES[host].en : HOST_NAMES[host].zh;

  // Conflicts only the running Word knows (custom keys, macros, MathType's
  // template) are recorded by the add-in when it leaves a chord to Word.
  const refreshWordReportedConflicts = useCallback(() => {
    invoke<string>("get_word_formula_shortcut_conflicts")
      .then((value) => setWordReportedConflicts(parseWordShortcutConflicts(value)))
      .catch(() => setWordReportedConflicts(new Map()));
  }, []);

  useEffect(() => {
    refreshWordReportedConflicts();
    window.addEventListener("focus", refreshWordReportedConflicts);
    return () => window.removeEventListener("focus", refreshWordReportedConflicts);
  }, [refreshWordReportedConflicts]);

  const wordConflictText = (shortcut: string | null) => {
    const binding = wordDefaultBinding(shortcut)
      ?? (shortcut ? wordReportedConflicts.get(shortcut) : undefined);
    return binding ? describeWordShortcutBinding(binding, isEn) : null;
  };

  const assignShortcut = (target: ShortcutTarget, shortcut: string | null) => {
    switch (target.host) {
      case "word":
        setWordShortcut(target.action, { shortcut });
        break;
      case "wpsWriter":
        setWpsWriterShortcut(target.action, { shortcut });
        break;
      case "powerPoint":
        setPowerPointShortcut(target.action, shortcut);
        break;
      case "wpsPresentation":
        setWpsPresentationShortcut(target.action, shortcut);
        break;
    }
  };

  useEffect(() => {
    if (!capturing) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (event.repeat || event.isComposing || event.key === "Process") return;
      if (
        event.key === "Escape" &&
        !event.ctrlKey &&
        !event.altKey &&
        !event.shiftKey &&
        !event.metaKey
      ) {
        setCapturing(null);
        setMessage("");
        return;
      }
      // WebView2 reports Ctrl+Alt as AltGraph (Windows types AltGr characters
      // with Ctrl+Alt): only an AltGr without both reported is refused.
      if (event.getModifierState("AltGraph") && !(event.ctrlKey && event.altKey)) {
        setMessage(
          isEn
            ? "AltGr is used for text input and cannot be assigned."
            : "AltGr 用于文字输入，不能设置为快捷键。",
        );
        return;
      }
      const result = officeFormulaShortcutFromEvent(event);
      if ("error" in result) {
        setMessage(captureErrorText(result, isEn));
        return;
      }
      // A system-wide VisualTeX Anywhere shortcut never reaches the Office window.
      const anywhere = ANYWHERE_HOTKEY_ACTIONS.find(
        (definition) =>
          definition.global && anywhereHotkeyLabel(anywhereBindings[definition.id]) === result.shortcut,
      );
      if (anywhere) {
        setMessage(
          isEn
            ? `${result.shortcut} is used by VisualTeX Anywhere (“${anywhere.labelEn}”).`
            : `${result.shortcut} 已用于 VisualTeX Anywhere（${anywhere.labelZh}）。`,
        );
        return;
      }
      if (isDocumentHost(capturing.host)) {
        const name = hostName(capturing.host);
        if (!event.ctrlKey) {
          setMessage(isEn
            ? `${name}: requires Ctrl.`
            : `${name}：须包含 Ctrl。`);
          return;
        }
        if (/^F(?:1[7-9]|2[0-4])$/.test(event.code)) {
          setMessage(isEn ? `${name}: F1–F16 only.` : `${name}：仅支持 F1–F16。`);
          return;
        }
      }
      const wordBinding =
        capturing.host === "word" ? wordDefaultBinding(result.shortcut) : null;
      if (wordBinding) {
        // Word always keeps its own bindings, so this chord could never fire.
        const owner = describeWordShortcutBinding(wordBinding, isEn);
        const suggestions = WORD_FREE_SHORTCUT_SUGGESTIONS.filter(
          (candidate) =>
            !Object.values(wordSettings).some((setting) => setting.shortcut === candidate),
        )
          .slice(0, 2)
          .join(isEn ? " or " : " 或 ");
        setMessage(
          isEn
            ? `${result.shortcut} is used by ${owner}. Available: ${suggestions}.`
            : `${result.shortcut} 已被 ${owner}占用。可用：${suggestions}。`,
        );
        return;
      }
      const moved = isDocumentHost(capturing.host)
        ? Object.entries(documentSettings(capturing.host)).some(
            ([action, setting]) =>
              action !== capturing.action && setting.shortcut === result.shortcut,
          )
        : Object.entries(slideSettings(capturing.host)).some(
            ([action, shortcut]) =>
              action !== capturing.action && shortcut === result.shortcut,
          );
      assignShortcut(capturing, result.shortcut);
      setCapturing(null);
      setMessage(
        moved
          ? isEn
            ? `${result.shortcut}: moved from another action.`
            : `${result.shortcut}：已从其他操作移至此处。`
          : "",
      );
    };
    document.addEventListener("keydown", handleKeyDown, true);
    return () => document.removeEventListener("keydown", handleKeyDown, true);
  }, [capturing, isEn, anywhereBindings, wordSettings, wpsWriterSettings, powerPointSettings, wpsPresentationSettings]);

  const shortcutControls = (
    target: ShortcutTarget,
    label: string,
    shortcut: string | null,
  ) => {
    const isCapturing = sameTarget(capturing, target);
    return (
      <>
        <button
          type="button"
          className={`office-formula-shortcut-capture${isCapturing ? " is-capturing" : ""}`}
          aria-pressed={isCapturing}
          onClick={() => {
            setMessage("");
            setMessageTarget(target);
            setCapturing(isCapturing ? null : target);
          }}
        >
          <kbd>
            {isCapturing
              ? isEn
                ? "Press keys…"
                : "请按键…"
              : shortcut ?? (isEn ? "None" : "无")}
          </kbd>
        </button>
        <button
          type="button"
          className="office-formula-shortcut-clear"
          disabled={!shortcut && !isCapturing}
          onClick={() => {
            assignShortcut(target, null);
            if (isCapturing) setCapturing(null);
            setMessage("");
            setMessageTarget(null);
          }}
          aria-label={isEn ? `Clear ${label}` : `清除${label}快捷键`}
          title={isEn ? "Set to None" : "设为无"}
        >
          <X size={14} />
        </button>
      </>
    );
  };

  const rowMessage = (target: ShortcutTarget) =>
    message && sameTarget(messageTarget, target) ? (
      <small className="office-formula-shortcut-warning" role="status" data-office-formula-shortcut-message>
        {message}
      </small>
    ) : null;

  const documentHostSection = (host: DocumentHost) => {
    const settings = documentSettings(host);
    const setShortcut = host === "word" ? setWordShortcut : setWpsWriterShortcut;
    const name = hostName(host);
    return (
      <section
        className="office-formula-shortcut-host"
        data-office-formula-shortcut-host={host}
      >
        <header className="office-formula-shortcut-host-header">
          <strong>{name}</strong>
        </header>
        <div className="office-formula-shortcut-list">
          {wordRows.map(({ action, labelEn, labelZh }) => {
            const setting = settings[action];
            const target = { host, action } as const;
            const label = isEn ? labelEn : labelZh;
            const conflict = host === "word" ? wordConflictText(setting.shortcut) : null;
            return (
              <div
                className="office-formula-shortcut-row"
                data-office-formula-shortcut={action}
                key={action}
              >
                <div className="office-formula-shortcut-copy">
                  <strong>{label}</strong>
                  {conflict && (
                    <small
                      className="office-formula-shortcut-warning"
                      data-office-formula-shortcut-conflict
                    >
                      {isEn
                        ? `${setting.shortcut} is used by ${conflict}.`
                        : `${setting.shortcut} 已被 ${conflict}占用。`}
                    </small>
                  )}
                  {rowMessage(target)}
                </div>
                {shortcutControls(target, `${name} ${label}`, setting.shortcut)}
                <label className="office-formula-format-select">
                  <span>{isEn ? "Format" : "格式"}</span>
                  <select
                    value={setting.format}
                    onChange={(event) =>
                      setShortcut(action, {
                        format: event.target.value === "omml" ? "omml" : "ole",
                      })
                    }
                    aria-label={isEn ? `${name} ${label} format` : `${name} ${label}格式`}
                  >
                    <option value="ole">OLE</option>
                    <option value="omml">OMML</option>
                  </select>
                </label>
              </div>
            );
          })}
        </div>
      </section>
    );
  };

  const slideHostSection = (host: SlideHost) => {
    const settings = slideSettings(host);
    const name = hostName(host);
    return (
      <section
        className="office-formula-shortcut-host"
        data-office-formula-shortcut-host={host}
      >
        <header className="office-formula-shortcut-host-header">
          <strong>{name}</strong>
        </header>
        <div className="office-formula-shortcut-list">
          {powerPointActions.map(({ action, labelEn, labelZh }) => {
            const target = { host, action } as const;
            const label = isEn ? labelEn : labelZh;
            return (
              <div
                className="office-formula-shortcut-row is-format-fixed"
                data-office-formula-shortcut={action}
                key={action}
              >
                <div className="office-formula-shortcut-copy">
                  <strong>{label}</strong>
                  {rowMessage(target)}
                </div>
                {shortcutControls(target, `${name} ${label}`, settings[action])}
              </div>
            );
          })}
        </div>
      </section>
    );
  };

  return (
    <section className="shortcut-settings-group" data-office-formula-shortcuts>
      <div className="settings-section-title">
        <Keyboard size={18} />
        <div>
          <h3>{isEn ? "Office shortcuts" : "Office 快捷键"}</h3>
        </div>
      </div>

      <div className="office-formula-shortcut-hosts">
        {documentHostSection("word")}
        {slideHostSection("powerPoint")}
        {documentHostSection("wpsWriter")}
        {slideHostSection("wpsPresentation")}
      </div>

      {message && !messageTarget && <p className="office-formula-shortcut-help" role="status">{message}</p>}
    </section>
  );
}
