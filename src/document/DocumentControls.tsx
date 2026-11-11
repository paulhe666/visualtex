// The document's formatting controls in the workspace's bottom bar, laid out as the formula controls that formula
// mode puts there (docs/document-editor/DESIGN.md §6.4). In the text: alignment, section, lists, indent, table,
// size, bold / italic / underline / strike, colours. While a formula is open: its display style, its colours and
// its LaTeX. Undo / redo are the header's (DocumentHistoryButtons), as in formula mode. They read the editor through the document store and run the shared commands.
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from "react";
import { createPortal } from "react-dom";
import type { Command, EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { redo, redoDepth, undo, undoDepth } from "prosemirror-history";
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  ChevronDown,
  ClipboardCopy,
  IndentDecrease,
  IndentIncrease,
  Italic,
  List,
  ListOrdered,
  Redo2,
  SquareCode,
  Strikethrough,
  Table,
  Image as ImageIcon,
  Underline,
  Undo2,
  Link2,
  BookMarked,
  FileCog,
  Hash,
  Superscript,
} from "lucide-react";
import {
  FONT_SIZES,
  currentAlign,
  lineListKind,
  toggleLineList,
  indentLines,
  outdentLines,
  insertCodeFence,
  insertTable,
  requestImage,
  isInTable,
  openFormulaId,
  insertSourceLink,
  sectionLevel,
  setAlign,
  setSourceHeading,
  sourceFormatActive,
  toggleDocumentBold,
  toggleSourceFormat,
  CHINESE_FONT_SIZES,
  chineseSizeName,
  bodyLayout,
  documentSettings,
  insertFootnote,
  LINE_SPACINGS,
  PAPER_SIZES,
  MARGIN_PRESETS,
  POINT_FONT_SIZES,
  clampMargin,
  type MarginName,
  setDocumentSettings,
  type DocumentSettings,
  type EquationBrackets,
  type EquationNumbering,
  type HeadingNumbering,
  type PaperName,
} from "../../shared/visualtex-document/src/index.ts";
import { fontLabel, PREFERRED_CJK, PREFERRED_LATIN, systemFontFamilies, type SystemFontFamily } from "./systemFonts";
import { CitationPicker } from "./CitationPicker";
import { useDocumentEditorStore } from "../stores/documentEditorStore";
import { useEditorStore } from "../stores/editorStore";
import { openDocumentFormula } from "./openDocumentFormula";
import { ColorMenuButton } from "./ColorPalette";
import { setMathLivePersistentTypingStyle, type MathLivePersistentTypingStyle } from "../editor/mathLiveKernelCommands";
import { portableFormulaLatex } from "./DocumentEditor";

type Lang = "en" | "zh";
const t = (lang: Lang, en: string, zh: string) => (lang === "en" ? en : zh);

function run(view: EditorView | null, command: Command) {
  if (!view) return false;
  const done = command(view.state, view.dispatch, view);
  view.focus();
  return done;
}

/** Bold: an open formula's own, in the text the source's "**" (sourceMarkdown.ts). */
const sourceBold: Command = (state, dispatch, view) =>
  openFormulaId(state) ? toggleDocumentBold(state, dispatch, view) : toggleSourceFormat("strong")(state, dispatch, view);

/** The value of an attribute mark (colour, size) at the selection, if one is set. */
function markValue(state: EditorState, name: string, attr: string): string | null {
  const type = state.schema.marks[name];
  const { $from, empty, from, to } = state.selection;
  if (empty) return (type.isInSet(state.storedMarks ?? $from.marks())?.attrs[attr] as string | undefined) ?? null;
  let value: string | null = null;
  state.doc.nodesBetween(from, to, (node) => {
    const mark = type.isInSet(node.marks);
    if (mark && value === null) value = mark.attrs[attr];
  });
  return value;
}

/** Set (or with null, remove) an attribute mark on the selection or the typing style. */
function setAttrMark(name: string, attrs: Record<string, string> | null): Command {
  return (state, dispatch) => {
    const type = state.schema.marks[name];
    const { from, to, empty } = state.selection;
    if (!dispatch) return true;
    if (empty) {
      const marks = (state.storedMarks ?? state.selection.$from.marks()).filter((m) => m.type !== type);
      dispatch(state.tr.setStoredMarks(attrs ? [...marks, type.create(attrs)] : marks));
      return true;
    }
    const tr = state.tr.removeMark(from, to, type);
    if (attrs) tr.addMark(from, to, type.create(attrs));
    dispatch(tr);
    return true;
  };
}

// The controls never take the focus: the editor keeps its selection.
const keepFocus = (event: { preventDefault: () => void }) => event.preventDefault();

function Button(props: {
  label: string;
  onClick: () => void;
  active?: boolean;
  disabled?: boolean;
  className?: string;
  style?: CSSProperties;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      className={"icon-button compact formula-formatting-button" + (props.active ? " is-active" : "") + (props.className ? ` ${props.className}` : "")}
      style={props.style}
      title={props.label}
      aria-label={props.label}
      aria-pressed={props.active}
      disabled={props.disabled}
      onMouseDown={keepFocus}
      onClick={props.onClick}
    >
      {props.children}
    </button>
  );
}

const Divider = () => <span className="formula-formatting-divider" aria-hidden="true" />;

/**
 * A button with a panel: the panel is placed on the page (no bar clips it), under the button or above it when the
 * window has no room below, and kept inside the window.
 */
function Popover(props: {
  label: string;
  button: ReactNode;
  disabled?: boolean;
  className?: string;
  style?: CSSProperties;
  wide?: boolean;
  children: (close: () => void) => ReactNode;
}) {
  const [open, setOpen] = useState(false);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const [place, setPlace] = useState<{ left: number; top: number } | null>(null);
  const close = () => setOpen(false);
  useLayoutEffect(() => {
    if (!open) {
      setPlace(null);
      return;
    }
    const button = buttonRef.current?.getBoundingClientRect();
    const panel = panelRef.current;
    if (!button || !panel) return;
    const width = panel.offsetWidth;
    const height = panel.offsetHeight;
    const below = window.innerHeight - button.bottom - 6;
    const top = below >= height || button.top < height + 6 ? button.bottom + 4 : button.top - height - 4;
    const left = Math.max(4, Math.min(button.left, window.innerWidth - width - 4));
    setPlace({ left, top: Math.max(4, Math.min(top, window.innerHeight - height - 4)) });
  }, [open]);
  return (
    <>
      <button
        ref={buttonRef}
        type="button"
        className={
          "icon-button compact formula-formatting-button document-control-popover-button" +
          (props.wide ? " is-wide" : "") +
          (open ? " is-active" : "") +
          (props.className ? ` ${props.className}` : "")
        }
        style={props.style}
        title={props.label}
        aria-label={props.label}
        aria-expanded={open}
        disabled={props.disabled}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        {props.button}
        <ChevronDown size={11} className="document-control-caret" />
      </button>
      {open &&
        createPortal(
          <>
            <div className="document-control-dismiss" onMouseDown={close} />
            <div
              ref={panelRef}
              className="document-control-panel"
              role="menu"
              style={place ? { left: place.left, top: place.top } : { left: -9999, top: -9999 }}
              onMouseDown={keepFocus}
            >
              {props.children(close)}
            </div>
          </>,
          document.body,
        )}
    </>
  );
}

/** The table size picker: rows x columns on an 8 x 8 grid. */
function TablePicker(props: { onPick: (rows: number, cols: number) => void }) {
  const [hover, setHover] = useState<[number, number]>([3, 3]);
  return (
    <div className="document-table-picker">
      <div className="document-table-grid">
        {Array.from({ length: 8 * 8 }, (_, k) => {
          const r = Math.floor(k / 8) + 1;
          const c = (k % 8) + 1;
          return (
            <span
              key={k}
              className={"document-table-cell" + (r <= hover[0] && c <= hover[1] ? " is-on" : "")}
              onMouseEnter={() => setHover([r, c])}
              onClick={() => props.onPick(r, c)}
            />
          );
        })}
      </div>
      <div className="document-table-size">
        {hover[0]} × {hover[1]}
      </div>
    </div>
  );
}

// ------------------------------------------------------------- page and body --

/** A press in a panel's field: it may take the focus (the panels otherwise keep the editor's). */
const letFocus = (event: { stopPropagation: () => void }) => event.stopPropagation();

const sameFamily = (a: string, b: string) => a.replace(/\s+/g, "").toLowerCase() === b.replace(/\s+/g, "").toLowerCase();

/** The system's fonts as a menu: a search field, the usual ones first, then all of them, each in its own face. */
function FontList(props: { lang: Lang; cjkOnly: boolean; value: string; onPick: (family: string) => void }) {
  const [fonts, setFonts] = useState<SystemFontFamily[] | null>(null);
  const [query, setQuery] = useState("");
  useEffect(() => {
    let live = true;
    // the list at hand first; then, in a browser, the fonts it may use (asked once: the user opened this menu)
    const take = (list: SystemFontFamily[]) => live && setFonts((shown) => (shown && shown.length > list.length ? shown : list));
    void systemFontFamilies().then(take);
    void systemFontFamilies({ ask: true }).then(take);
    return () => {
      live = false;
    };
  }, []);
  const isEn = props.lang === "en";
  const { preferred, rest } = useMemo(() => {
    const list = (fonts ?? []).filter((font) => !props.cjkOnly || font.cjk);
    const q = query.trim().toLowerCase();
    const matching = q ? list.filter((font) => font.family.toLowerCase().includes(q) || (font.localized ?? "").toLowerCase().includes(q)) : list;
    const order = props.cjkOnly ? PREFERRED_CJK : PREFERRED_LATIN;
    const preferred = order.map((name) => matching.find((font) => sameFamily(font.family, name))).filter((font): font is SystemFontFamily => !!font);
    return { preferred, rest: matching.filter((font) => !preferred.includes(font)) };
  }, [fonts, query, props.cjkOnly]);
  const item = (font: SystemFontFamily) => (
    <button
      key={font.family}
      type="button"
      className={"document-menu-item document-font-item" + (sameFamily(font.family, props.value) ? " is-current" : "")}
      style={{ fontFamily: `"${font.family.replace(/"/g, "")}", var(--font-ui, sans-serif)` }}
      title={font.localized ? `${font.localized} (${font.family})` : font.family}
      onClick={() => props.onPick(font.family)}
    >
      {fontLabel(font, isEn)}
    </button>
  );
  return (
    <div className="document-font-list">
      <input
        className="document-font-search"
        autoFocus
        value={query}
        onMouseDown={letFocus}
        onChange={(event) => setQuery(event.target.value)}
        placeholder={t(props.lang, "Search fonts", "搜索字体")}
        spellCheck={false}
      />
      {!fonts && <div className="document-menu-note">{t(props.lang, "Reading the system's fonts…", "正在读取系统字体…")}</div>}
      {preferred.length > 0 && <div className="document-menu-heading">{t(props.lang, "Common", "常用")}</div>}
      {preferred.map(item)}
      {rest.length > 0 && <div className="document-menu-heading">{t(props.lang, "All fonts", "全部字体")}</div>}
      {rest.map(item)}
      {fonts && !preferred.length && !rest.length && <div className="document-menu-note">{t(props.lang, "No font matches", "没有匹配的字体")}</div>}
    </div>
  );
}

function FontButton(props: { lang: Lang; label: string; caption: string; value: string; cjkOnly: boolean; onPick: (family: string) => void }) {
  const [fonts, setFonts] = useState<SystemFontFamily[] | null>(null);
  useEffect(() => {
    let live = true;
    void systemFontFamilies().then((list) => live && setFonts(list));
    return () => {
      live = false;
    };
  }, []);
  const known = fonts?.find((font) => sameFamily(font.family, props.value));
  return (
    <Popover
      label={props.label}
      wide
      className="document-font-button"
      button={
        <span className="document-font-value">
          <span className="document-font-caption">{props.caption}</span>
          <span style={{ fontFamily: `"${props.value.replace(/"/g, "")}"` }}>{fontLabel(known ?? { family: props.value, localized: null }, props.lang === "en")}</span>
        </span>
      }
    >
      {(close) => (
        <FontList
          lang={props.lang}
          cjkOnly={props.cjkOnly}
          value={props.value}
          onPick={(family) => {
            close();
            props.onPick(family);
          }}
        />
      )}
    </Popover>
  );
}

const formatPt = (pt: number) => `${Math.round(pt * 10) / 10}`;

/** The body size: Chinese named sizes and point sizes. */
/** The one font-size control: the body text's size (the whole document: Chinese named sizes or points) and, for the
 *  selected text, LaTeX's sizes relative to it (\small … \Huge, each with the size it gives). */
function SizeButton(props: { lang: Lang; body: number; mark: string; onBody: (pt: number) => void; onMark: (name: string) => void }) {
  const { lang, body, mark } = props;
  const named = chineseSizeName(body);
  const bodyLabel = lang === "en" ? `${formatPt(body)} pt` : (named ?? `${formatPt(body)} 磅`);
  const shown = mark !== "normalsize" ? `\\${mark}` : bodyLabel;
  return (
    <Popover label={t(lang, "Font size", "字号")} wide className="document-body-size-button" button={<span className="document-size-value">{shown}</span>}>
      {(close) => (
        <div className="document-size-columns">
          <div>
            <div className="document-menu-heading">{t(lang, "Body · Chinese sizes", "正文 · 中文字号")}</div>
            {CHINESE_FONT_SIZES.filter((size) => size.pt >= 7.5 && size.pt <= 26).map((size) => (
              <button
                key={size.name}
                type="button"
                className={"document-menu-item" + (Math.abs(size.pt - body) < 0.01 ? " is-current" : "")}
                onClick={() => {
                  close();
                  props.onBody(size.pt);
                }}
              >
                {size.name}
                <span className="document-menu-hint">{formatPt(size.pt)} pt</span>
              </button>
            ))}
          </div>
          <div>
            <div className="document-menu-heading">{t(lang, "Body · points", "正文 · 磅值")}</div>
            {POINT_FONT_SIZES.map((pt) => (
              <button
                key={pt}
                type="button"
                className={"document-menu-item" + (Math.abs(pt - body) < 0.01 ? " is-current" : "")}
                onClick={() => {
                  close();
                  props.onBody(pt);
                }}
              >
                {formatPt(pt)}
              </button>
            ))}
          </div>
          <div>
            <div className="document-menu-heading">{t(lang, "Selected text", "所选文字")}</div>
            {FONT_SIZES.map((s) => (
              <button
                key={s.name}
                type="button"
                className={"document-menu-item" + (s.name === mark ? " is-current" : "")}
                onClick={() => {
                  close();
                  props.onMark(s.name);
                }}
              >
                {`\\${s.name}`}
                <span className="document-menu-hint">{formatPt(s.ratio * body)} pt</span>
              </button>
            ))}
          </div>
        </div>
      )}
    </Popover>
  );
}

/** Equation and section numbers. */
function NumberingSetup(props: { lang: Lang; settings: DocumentSettings; apply: (patch: Partial<DocumentSettings>) => void }) {
  const { lang, settings, apply } = props;
  const open = settings.equationBrackets === "[]" ? "[" : "(";
  const close = settings.equationBrackets === "[]" ? "]" : ")";
  const row = <T extends string>(choices: Array<[T, string]>, current: T, set: (value: T) => void) => (
    <div className="document-page-papers">
      {choices.map(([value, label]) => (
        <button key={value} type="button" className={"document-menu-item" + (current === value ? " is-current" : "")} onClick={() => set(value)}>
          {label}
        </button>
      ))}
    </div>
  );
  return (
    <div className="document-page-setup">
      <div className="document-menu-heading">{t(lang, "Equation numbers", "公式编号")}</div>
      {row<EquationNumbering>(
        [
          ["1", `${open}1${close}`],
          ["1.1", `${open}1.1${close}`],
          ["1-1", `${open}1-1${close}`],
        ],
        settings.equationNumbering,
        (equationNumbering) => apply({ equationNumbering }),
      )}
      <div className="document-menu-heading">{t(lang, "Brackets", "括号")}</div>
      {row<EquationBrackets>(
        [
          ["()", "( )"],
          ["[]", "[ ]"],
        ],
        settings.equationBrackets,
        (equationBrackets) => apply({ equationBrackets }),
      )}
      <div className="document-menu-heading">{t(lang, "Section numbers", "章节编号")}</div>
      {row<HeadingNumbering>(
        [
          ["1", "1  1.1"],
          ["1.", "1.  1.1."],
          ["none", t(lang, "None", "无")],
        ],
        settings.headingNumbering,
        (headingNumbering) => apply({ headingNumbering }),
      )}
    </div>
  );
}

/** Paper, margins and line spacing. */
/** A margin typed in mm: shows the document's value (also after a drag on the sheet); Enter or leaving the field
 *  sets it, Esc puts the value back, ↑/↓ step 0.1 mm (Shift: 1 mm). */
function MarginField(props: { name: MarginName; label: string; settings: DocumentSettings; apply: (patch: Partial<DocumentSettings>) => void }) {
  const { name, settings, apply } = props;
  const current = Math.round(settings[name] * 10) / 10;
  const [draft, setDraft] = useState<string | null>(null);
  const commit = (text: string) => {
    setDraft(null);
    const value = Number(text.replace(",", "."));
    if (!Number.isFinite(value)) return;
    const mm = clampMargin(settings, name, value);
    if (Math.abs(mm - settings[name]) > 0.01) apply({ [name]: mm });
  };
  return (
    <label className="document-page-field">
      <span>{props.label}</span>
      <input
        type="text"
        inputMode="decimal"
        data-margin={name}
        value={draft ?? String(current)}
        onMouseDown={letFocus}
        onChange={(event) => setDraft(event.target.value)}
        onBlur={(event) => commit(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            commit(event.currentTarget.value);
          } else if (event.key === "Escape") {
            event.stopPropagation();
            setDraft(null);
          } else if (event.key === "ArrowUp" || event.key === "ArrowDown") {
            event.preventDefault();
            const step = (event.shiftKey ? 1 : 0.1) * (event.key === "ArrowUp" ? 1 : -1);
            commit(String((Number(event.currentTarget.value) || current) + step));
          }
        }}
      />
      <span className="document-page-unit">mm</span>
    </label>
  );
}

function PageSetup(props: { lang: Lang; settings: DocumentSettings; apply: (patch: Partial<DocumentSettings>) => void; type: (patch: Partial<DocumentSettings>) => void }) {
  const { lang, settings, apply } = props;
  const preset = MARGIN_PRESETS.find(
    (p) => Math.abs(p.top - settings.marginTop) < 0.05 && Math.abs(p.bottom - settings.marginBottom) < 0.05 && Math.abs(p.left - settings.marginLeft) < 0.05 && Math.abs(p.right - settings.marginRight) < 0.05,
  );
  return (
    <div className="document-page-setup">
      <div className="document-menu-heading">{t(lang, "Paper", "纸张")}</div>
      <div className="document-page-papers">
        {(Object.keys(PAPER_SIZES) as PaperName[]).map((name) => (
          <button
            key={name}
            type="button"
            className={"document-menu-item" + (settings.paper === name ? " is-current" : "")}
            onClick={() => apply({ paper: name })}
            title={`${PAPER_SIZES[name].width} × ${PAPER_SIZES[name].height} mm`}
          >
            {PAPER_SIZES[name].label}
          </button>
        ))}
      </div>
      <div className="document-menu-heading">{t(lang, "Margins", "页边距")}</div>
      <div className="document-page-margins">
        <MarginField name="marginTop" label={t(lang, "Top", "上")} settings={settings} apply={props.type} />
        <MarginField name="marginBottom" label={t(lang, "Bottom", "下")} settings={settings} apply={props.type} />
        <MarginField name="marginLeft" label={t(lang, "Left", "左")} settings={settings} apply={props.type} />
        <MarginField name="marginRight" label={t(lang, "Right", "右")} settings={settings} apply={props.type} />
      </div>
      <div className="document-page-papers document-margin-presets">
        {MARGIN_PRESETS.map((p) => (
          <button
            key={p.name}
            type="button"
            className={"document-menu-item" + (preset === p ? " is-current" : "")}
            title={`${p.top} / ${p.bottom} / ${p.left} / ${p.right} mm`}
            onClick={() => apply({ marginTop: p.top, marginBottom: p.bottom, marginLeft: p.left, marginRight: p.right })}
          >
            {t(lang, p.en, p.zh)}
          </button>
        ))}
      </div>
      <div className="document-menu-heading">{t(lang, "Line spacing", "行距")}</div>
      <div className="document-page-papers">
        {LINE_SPACINGS.map((spacing) => (
          <button
            key={spacing}
            type="button"
            className={"document-menu-item" + (Math.abs(settings.lineSpacing - spacing) < 0.001 ? " is-current" : "")}
            onClick={() => apply({ lineSpacing: spacing })}
          >
            {spacing === 1 ? t(lang, "Single", "单倍") : `${spacing}`}
          </button>
        ))}
      </div>
    </div>
  );
}

// ----------------------------------------------------------- the open formula --

function toggleDisplayStyle() {
  const target = openDocumentFormula();
  if (!target) return;
  const { field, mutate } = target;
  field.focus();
  mutate(() => {
    const latex = field.value;
    const next = /^\s*\\displaystyle\b/.test(latex) ? latex.replace(/^\s*\\displaystyle\s*/, "") : `\\displaystyle ${latex}`;
    field.setValue(next, { mode: "math", format: "latex", insertionMode: "replaceAll", selectionMode: "after", silenceNotifications: true });
    return true;
  });
}

/** Colour the open formula's selection (no selection: what is typed next). */
function setFormulaStyle(style: { color?: string; backgroundColor?: string }) {
  const target = openDocumentFormula();
  if (!target) return;
  const { field, mutate } = target;
  field.focus();
  mutate(() => {
    const [[start, end]] = field.selection.ranges;
    const range: [number, number] = [Math.min(start, end), Math.max(start, end)];
    if (range[0] === range[1]) {
      // what is typed next takes its colours from the field's typing style (the kernel ignores applyStyle here)
      const current =
        (field as unknown as { visualTexPersistentTypingStyle?: Partial<MathLivePersistentTypingStyle> }).visualTexPersistentTypingStyle ?? {};
      const pick = (value: string | undefined, kept: string | null | undefined) =>
        value === undefined ? (kept ?? null) : value === "none" ? null : value;
      setMathLivePersistentTypingStyle(field, {
        bold: false,
        italic: null,
        color: pick(style.color, current.color),
        backgroundColor: pick(style.backgroundColor, current.backgroundColor),
      });
      return false;
    }
    field.applyStyle(style, { range, operation: "set" });
    return true;
  });
}

function copyFormulaLatex() {
  const target = openDocumentFormula();
  if (!target) return;
  void navigator.clipboard.writeText(portableFormulaLatex(target.field.value));
  target.field.focus();
}

// ------------------------------------------------------------------ controls --

function useLang(): Lang {
  const language = useEditorStore((store) => store.language);
  return language === "en" ? "en" : "zh";
}

/**
 * The bottom strip in document mode: the open formula's controls (display style, the formula editor's colour
 * palette, copy LaTeX), as in formula mode; the document's own buttons are at the top (DocumentTextToolbar).
 */
export function DocumentControls() {
  const state = useDocumentEditorStore((store) => store.state);
  const lang = useLang();
  if (!state || !openFormulaId(state)) return <div className="formula-alignment-controls document-controls" />;
  const display = /^\s*\\displaystyle\b/.test(openDocumentFormula()?.field.value ?? "");
  return (
    <div className="formula-alignment-controls document-controls" role="toolbar" aria-label={t(lang, "Formula", "公式")}>
      <Button label={t(lang, "Display style (\\displaystyle)", "显示样式 (\\displaystyle)")} active={display} className="is-text-action" onClick={toggleDisplayStyle}>
        <span className="document-control-text">Σ→<span className="document-big-sigma">Σ</span></span>
      </Button>
      <Divider />
      <ColorMenuButton
        kind="color"
        lang={lang}
        value={null}
        label={t(lang, "Formula text color", "公式文字颜色")}
        title={t(lang, "Text color", "文字颜色")}
        onApply={(color) => setFormulaStyle({ color: color ?? "none" })}
      />
      <ColorMenuButton
        kind="backgroundColor"
        lang={lang}
        value={null}
        label={t(lang, "Formula background color", "公式背景颜色")}
        title={t(lang, "Background color", "背景颜色")}
        onApply={(color) => setFormulaStyle({ backgroundColor: color ?? "none" })}
      />
      <Divider />
      <Button label={t(lang, "Copy LaTeX", "复制 LaTeX")} onClick={copyFormulaLatex}>
        <ClipboardCopy size={15} />
      </Button>
    </div>
  );
}

/** The document's own buttons, above the document: paragraph, lists, table, code, size, marks, colours. */
export function DocumentTextToolbar(props: { leading?: ReactNode; trailing?: ReactNode } = {}) {
  const view = useDocumentEditorStore((store) => store.view);
  const state = useDocumentEditorStore((store) => store.state);
  const lang = useLang();
  if (!state) return <div className="document-controls document-top-toolbar">{props.leading}{props.trailing}</div>;
  const inMath = !!openFormulaId(state);
  const align = currentAlign(state);
  const list = lineListKind(state);
  const level = sectionLevel(state);
  const size = markValue(state, "fontSize", "size") ?? "normalsize";
  const color = markValue(state, "textColor", "color");
  const highlight = markValue(state, "highlight", "color");
  const settings = documentSettings(state.doc);
  const applySettings = (patch: Partial<DocumentSettings>) => run(view, setDocumentSettings(patch));
  // a value typed in a field: set without taking the focus from the field
  const typeSettings = (patch: Partial<DocumentSettings>) => {
    if (view) setDocumentSettings(patch)(view.state, view.dispatch);
  };
  return (
    <div className={"document-controls document-top-toolbar" + (inMath ? " is-in-formula" : "")} role="toolbar" aria-label={t(lang, "Document", "文档")}>
      {props.leading}
      <FontButton
        lang={lang}
        label={t(lang, "Chinese font", "中文字体")}
        caption={t(lang, "CJK", "中")}
        value={settings.cjkFont}
        cjkOnly
        onPick={(cjkFont) => applySettings({ cjkFont })}
      />
      <FontButton
        lang={lang}
        label={t(lang, "Latin font", "西文字体")}
        caption={t(lang, "Latin", "西")}
        value={settings.latinFont}
        cjkOnly={false}
        onPick={(latinFont) => applySettings({ latinFont })}
      />
      <SizeButton
        lang={lang}
        body={bodyLayout(state.doc).fontSize}
        mark={size}
        onBody={(fontSize) => applySettings({ fontSize })}
        onMark={(name) => run(view, setAttrMark("fontSize", name === "normalsize" ? null : { size: name }))}
      />
      <Popover label={t(lang, "Page setup", "页面设置")} button={<FileCog size={16} strokeWidth={1.8} />}>
        {() => <PageSetup lang={lang} settings={settings} apply={applySettings} type={typeSettings} />}
      </Popover>
      <Popover label={t(lang, "Numbering", "编号格式")} button={<Hash size={16} strokeWidth={1.8} />} className="document-numbering-button">
        {() => <NumberingSetup lang={lang} settings={settings} apply={applySettings} />}
      </Popover>
      <Divider />
      {(
        [
          ["left", AlignLeft, t(lang, "Align left (Ctrl+Shift+L)", "左对齐 (Ctrl+Shift+L)")],
          ["center", AlignCenter, t(lang, "Centre (Ctrl+Shift+E)", "居中 (Ctrl+Shift+E)")],
          ["right", AlignRight, t(lang, "Align right (Ctrl+Shift+R)", "右对齐 (Ctrl+Shift+R)")],
          ["justify", AlignJustify, t(lang, "Justify (Ctrl+Shift+J)", "两端对齐 (Ctrl+Shift+J)")],
        ] as const
      ).map(([value, Icon, label]) => (
        <Button key={value} label={label} active={align === value} className="formula-alignment-button" onClick={() => run(view, setAlign(value))}>
          <Icon size={16} strokeWidth={2} />
        </Button>
      ))}
      <Divider />
      <Button label={t(lang, "Heading (Ctrl+1–6)", "标题 (Ctrl+1–6)")} active={level > 0} className="is-text-action" onClick={() => run(view, setSourceHeading(level || 1))}>
        <span className="document-section-icon">
          1.
          <br />
          1.2.
        </span>
      </Button>
      <Button label={t(lang, "Numbered list (Ctrl+Shift+[)", "有序列表 (Ctrl+Shift+[)")} active={list === "ordered"} onClick={() => run(view, toggleLineList("ordered"))}>
        <ListOrdered size={16} />
      </Button>
      <Button label={t(lang, "Bulleted list (Ctrl+Shift+])", "无序列表 (Ctrl+Shift+])")} active={list === "bullet"} onClick={() => run(view, toggleLineList("bullet"))}>
        <List size={16} />
      </Button>
      <Button label={t(lang, "Indent", "增加缩进")} onClick={() => run(view, indentLines)}>
        <IndentIncrease size={16} />
      </Button>
      <Button label={t(lang, "Outdent", "减少缩进")} onClick={() => run(view, outdentLines)}>
        <IndentDecrease size={16} />
      </Button>
      <Divider />
      <Popover label={t(lang, "Table (Ctrl+T)", "表格 (Ctrl+T)")} disabled={isInTable(state)} button={<Table size={16} strokeWidth={1.8} />}>
        {(close) => (
          <TablePicker
            onPick={(rows, cols) => {
              close();
              run(view, insertTable(rows, cols));
            }}
          />
        )}
      </Popover>
      <Button label={t(lang, "Image", "图片")} onClick={() => view && requestImage(view)}>
        <ImageIcon size={16} strokeWidth={1.8} />
      </Button>
      <Button label={t(lang, "Code block", "代码块")} onClick={() => run(view, insertCodeFence)}>
        <SquareCode size={16} strokeWidth={1.8} />
      </Button>
      <Button label={t(lang, "Link (Ctrl+K)", "链接 (Ctrl+K)")} onClick={() => run(view, insertSourceLink)}>
        <Link2 size={16} strokeWidth={1.9} />
      </Button>
      <Button label={t(lang, "Footnote (Ctrl+Alt+F)", "脚注 (Ctrl+Alt+F)")} onClick={() => run(view, insertFootnote())}>
        <Superscript size={16} strokeWidth={1.9} />
      </Button>
      <Popover label={t(lang, "Citation", "引用")} button={<BookMarked size={16} strokeWidth={1.8} />}>
        {(close) => <CitationPicker lang={lang} view={view} state={state} close={close} />}
      </Popover>
      <Divider />
      <Button label={t(lang, "Bold (Ctrl+B)", "加粗 (Ctrl+B)")} active={sourceFormatActive(state, "strong")} onClick={() => run(view, sourceBold)}>
        <Bold size={15} strokeWidth={2.2} />
      </Button>
      <Button label={t(lang, "Italic (Ctrl+I)", "斜体 (Ctrl+I)")} active={sourceFormatActive(state, "em")} onClick={() => run(view, toggleSourceFormat("em"))}>
        <Italic size={15} strokeWidth={2.2} />
      </Button>
      <Button label={t(lang, "Underline (Ctrl+U)", "下划线 (Ctrl+U)")} active={sourceFormatActive(state, "underline")} onClick={() => run(view, toggleSourceFormat("underline"))}>
        <Underline size={15} strokeWidth={2.2} />
      </Button>
      <Button label={t(lang, "Strikethrough (Alt+Shift+5)", "删除线 (Alt+Shift+5)")} active={sourceFormatActive(state, "strike")} onClick={() => run(view, toggleSourceFormat("strike"))}>
        <Strikethrough size={15} strokeWidth={2.2} />
      </Button>
      <ColorMenuButton
        kind="color"
        lang={lang}
        value={color}
        label={t(lang, "Text color", "文字颜色")}
        title={t(lang, "Text color", "文字颜色")}
        onApply={(c) => run(view, setAttrMark("textColor", c ? { color: c } : null))}
      />
      <ColorMenuButton
        kind="backgroundColor"
        lang={lang}
        value={highlight}
        label={t(lang, "Text background", "文字背景颜色")}
        title={t(lang, "Text background", "文字背景颜色")}
        onApply={(c) => run(view, setAttrMark("highlight", c ? { color: c } : null))}
      />
      {props.trailing}
    </div>
  );
}

/** Undo / redo of the document, for the main window's header (where formula mode has its own). */
export function DocumentHistoryButtons(props: { isEn: boolean }) {
  const view = useDocumentEditorStore((store) => store.view);
  const state = useDocumentEditorStore((store) => store.state);
  const { isEn } = props;
  return (
    <>
      <button
        type="button"
        className="icon-button"
        onMouseDown={keepFocus}
        onClick={() => run(view, undo)}
        disabled={!state || undoDepth(state) === 0}
        aria-label={isEn ? "Undo" : "撤销"}
        title={isEn ? "Undo (Ctrl+Z)" : "撤销 (Ctrl+Z)"}
      >
        <Undo2 size={17} />
      </button>
      <button
        type="button"
        className="icon-button"
        onMouseDown={keepFocus}
        onClick={() => run(view, redo)}
        disabled={!state || redoDepth(state) === 0}
        aria-label={isEn ? "Redo" : "重做"}
        title={isEn ? "Redo (Ctrl+Y)" : "重做 (Ctrl+Y)"}
      >
        <Redo2 size={17} />
      </button>
    </>
  );
}
