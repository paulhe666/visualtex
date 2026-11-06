// The page margins on the sheet (DESIGN.md §8.2): the text area's four edges can be dragged, on every page, as on
// Word's ruler. While dragging, the sheet shows the new margin and its value; the release sets it (one undo step).
// Each edge's grip lies in the margin just outside the text area (never over the text, where a press selects).
// A double-click on a grip opens a field there for the exact value in mm.
import { useEffect, useRef, useState, type ReactElement } from "react";
import { PAPER_SIZES, clampMargin, type DocumentSettings, type MarginName } from "../../shared/visualtex-document/src/index.ts";

type Lang = "en" | "zh";

const LABELS: Record<MarginName, { en: string; zh: string }> = {
  marginTop: { en: "Top margin", zh: "上边距" },
  marginBottom: { en: "Bottom margin", zh: "下边距" },
  marginLeft: { en: "Left margin", zh: "左边距" },
  marginRight: { en: "Right margin", zh: "右边距" },
};

const round = (mm: number) => Math.round(mm * 10) / 10;
/** the grip's width (px), outside the text area */
const GRIP = 6;

export function MarginGuides(props: {
  sheet: HTMLElement | null;
  settings: DocumentSettings;
  /** screen px per mm */
  mm: number;
  gap: number;
  lang: Lang;
  onPreview: (patch: Partial<DocumentSettings> | null) => void;
  onCommit: (patch: Partial<DocumentSettings>) => void;
}) {
  const { sheet, settings, mm, gap, lang } = props;
  const paper = PAPER_SIZES[settings.paper];
  const stride = paper.height * mm + gap;
  // as many pages as the sheet holds
  const [pages, setPages] = useState(1);
  useEffect(() => {
    if (!sheet) return;
    const count = () => setPages(Math.max(1, Math.round((sheet.offsetHeight + gap) / stride)));
    count();
    const observer = new ResizeObserver(count);
    observer.observe(sheet);
    return () => observer.disconnect();
  }, [sheet, stride, gap]);
  const [drag, setDrag] = useState<{ name: MarginName; value: number; page: number } | null>(null);
  // the exact-value field a double-click opens (at, in the layer's px)
  const [editing, setEditing] = useState<{ name: MarginName; left: number; top: number } | null>(null);
  const field = useRef<HTMLInputElement>(null);
  const editOpen = useRef(false);
  useEffect(() => {
    field.current?.focus();
    field.current?.select();
  }, [editing]);
  const edit = (name: MarginName) => (event: React.MouseEvent<HTMLDivElement>) => {
    if (!sheet) return;
    const rect = sheet.getBoundingClientRect();
    editOpen.current = true;
    setEditing({ name, left: event.clientX - rect.left + 10, top: event.clientY - rect.top + 10 });
  };
  const commitEdit = (text: string) => {
    if (!editing || !editOpen.current) return;
    editOpen.current = false;
    const value = Number(text.replace(",", "."));
    setEditing(null);
    if (!Number.isFinite(value)) return;
    const mm = clampMargin(settings, editing.name, value);
    if (Math.abs(mm - settings[editing.name]) > 0.01) props.onCommit({ [editing.name]: mm });
  };

  const start = (name: MarginName, page: number) => (event: React.PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0 || !sheet) return;
    event.preventDefault();
    event.stopPropagation();
    const target = event.currentTarget;
    target.setPointerCapture(event.pointerId);
    const rect = sheet.getBoundingClientRect();
    const pageTop = rect.top + page * stride;
    const rawAt = (x: number, y: number) =>
      name === "marginLeft" ? (x - rect.left) / mm
      : name === "marginRight" ? (rect.right - x) / mm
      : name === "marginTop" ? (y - pageTop) / mm
      : (pageTop + paper.height * mm - y) / mm;
    // the edge keeps its distance to the pointer (the grip is beside the edge, not on it)
    const grab = rawAt(event.clientX, event.clientY) - settings[name];
    const valueAt = (x: number, y: number) => clampMargin(settings, name, rawAt(x, y) - grab);
    let value = settings[name];
    const move = (e: PointerEvent) => {
      value = valueAt(e.clientX, e.clientY);
      setDrag({ name, value, page });
      props.onPreview({ [name]: value });
    };
    const up = () => {
      target.removeEventListener("pointermove", move);
      target.removeEventListener("pointerup", up);
      target.removeEventListener("pointercancel", up);
      setDrag(null);
      props.onPreview(null);
      if (Math.abs(value - settings[name]) >= 0.05) props.onCommit({ [name]: value });
    };
    target.addEventListener("pointermove", move);
    target.addEventListener("pointerup", up);
    target.addEventListener("pointercancel", up);
    setDrag({ name, value, page });
  };

  const width = paper.width * mm;
  const guides: ReactElement[] = [];
  const title = (name: MarginName) => `${LABELS[name][lang]} ${round(settings[name])} mm`;
  guides.push(
    <div key="l" className="vtx-margin-guide is-vertical is-start" style={{ left: settings.marginLeft * mm - GRIP }} title={title("marginLeft")} onPointerDown={start("marginLeft", 0)} onDoubleClick={edit("marginLeft")} />,
    <div key="r" className="vtx-margin-guide is-vertical is-end" style={{ left: width - settings.marginRight * mm }} title={title("marginRight")} onPointerDown={start("marginRight", 0)} onDoubleClick={edit("marginRight")} />,
  );
  for (let page = 0; page < pages; page++) {
    const top = page * stride;
    guides.push(
      <div key={`t${page}`} className="vtx-margin-guide is-horizontal is-start" style={{ top: top + settings.marginTop * mm - GRIP }} title={title("marginTop")} onPointerDown={start("marginTop", page)} onDoubleClick={edit("marginTop")} />,
      <div key={`b${page}`} className="vtx-margin-guide is-horizontal is-end" style={{ top: top + (paper.height - settings.marginBottom) * mm }} title={title("marginBottom")} onPointerDown={start("marginBottom", page)} onDoubleClick={edit("marginBottom")} />,
    );
  }
  let line: ReactElement | null = null;
  if (drag) {
    const vertical = drag.name === "marginLeft" || drag.name === "marginRight";
    const pageTop = drag.page * stride;
    const at =
      drag.name === "marginLeft" ? drag.value * mm
      : drag.name === "marginRight" ? width - drag.value * mm
      : drag.name === "marginTop" ? pageTop + drag.value * mm
      : pageTop + (paper.height - drag.value) * mm;
    line = (
      <>
        <div className={"vtx-margin-line " + (vertical ? "is-vertical" : "is-horizontal")} style={vertical ? { left: at } : { top: at }} />
        <div className="vtx-margin-value" style={vertical ? { left: at + 6, top: pageTop + 8 + (sheet ? Math.max(0, -sheet.getBoundingClientRect().top) : 0) } : { left: 8, top: at + 6 }}>
          {LABELS[drag.name][lang]} {drag.value.toFixed(1)} mm
        </div>
      </>
    );
  }
  return (
    <div className={"vtx-margin-guides" + (drag ? " is-dragging" : "")}>
      {guides}
      {line}
      {editing && (
        <form
          className="vtx-margin-edit"
          style={{ left: editing.left, top: editing.top }}
          onMouseDown={(event) => event.stopPropagation()}
          onPointerDown={(event) => event.stopPropagation()}
          onSubmit={(event) => {
            event.preventDefault();
            commitEdit(field.current?.value ?? "");
          }}
        >
          <span>{LABELS[editing.name][lang]}</span>
          <input
            ref={field}
            type="text"
            inputMode="decimal"
            defaultValue={String(round(settings[editing.name]))}
            onBlur={(event) => commitEdit(event.target.value)}
            onKeyDown={(event) => {
              event.stopPropagation();
              if (event.key === "Escape") {
                event.preventDefault();
                editOpen.current = false;
                setEditing(null);
              }
            }}
          />
          <span>mm</span>
        </form>
      )}
    </div>
  );
}
