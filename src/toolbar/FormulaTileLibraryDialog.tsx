import { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { Check, Plus, Search, X } from "lucide-react";
import { MathPreview } from "../components/MathPreview";
import {
  formulaTileLevelLabels,
  formulaTileLibrary,
  formulaTileLibraryCount,
  formulaTileSubjects,
  type FormulaTileLibraryGroup,
  type FormulaTileSubject,
} from "./formulaTileLibrary";

// Long formulas get a double-width card so they stay readable.
const wideFormulaLength = 72;

interface Props {
  open: boolean;
  language: string;
  /** Ids currently shown in the Common tab. */
  selectedIds: string[];
  onToggle: (tileId: string) => void;
  onClose: () => void;
}

// Picks built-in formula tiles for the Common tab. Clicking a tile adds it,
// clicking it again removes it; the tab updates while the dialog stays open.
export function FormulaTileLibraryDialog({
  open,
  language,
  selectedIds,
  onToggle,
  onClose,
}: Props) {
  const isEn = language === "en";
  const [subject, setSubject] = useState<FormulaTileSubject>("math");
  const [groupId, setGroupId] = useState(formulaTileLibrary[0].id);
  const [query, setQuery] = useState("");

  useEffect(() => {
    if (!open) return;
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      event.stopPropagation();
      onClose();
    };
    window.addEventListener("keydown", closeOnEscape, true);
    return () => window.removeEventListener("keydown", closeOnEscape, true);
  }, [onClose, open]);

  const selected = useMemo(() => new Set(selectedIds), [selectedIds]);
  const subjectGroups = useMemo(
    () => formulaTileLibrary.filter((group) => group.subject === subject),
    [subject],
  );
  const activeGroup =
    subjectGroups.find((group) => group.id === groupId) ?? subjectGroups[0];

  const needle = query.trim().toLowerCase();
  const visibleGroups = useMemo<FormulaTileLibraryGroup[]>(() => {
    if (!needle) return activeGroup ? [activeGroup] : [];
    return formulaTileLibrary
      .map((group) => ({
        ...group,
        tiles: group.tiles.filter(
          (tile) =>
            tile.labelZh.toLowerCase().includes(needle) ||
            tile.labelEn.toLowerCase().includes(needle) ||
            tile.latex.toLowerCase().includes(needle),
        ),
      }))
      .filter((group) => group.tiles.length > 0);
  }, [activeGroup, needle]);

  if (!open) return null;

  const selectSubject = (next: FormulaTileSubject) => {
    setSubject(next);
    setGroupId(
      formulaTileLibrary.find((group) => group.subject === next)?.id ?? groupId,
    );
  };

  return createPortal(
    <div
      className="modal-backdrop formula-tile-library-backdrop"
      role="presentation"
      onMouseDown={onClose}
    >
      <section
        className="formula-tile-library-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="formula-tile-library-title"
        data-formula-tile-library
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <span className="eyebrow">FORMULA TILES</span>
            <h2 id="formula-tile-library-title">
              {isEn ? "Add common tiles" : "添加常用磁贴"}
            </h2>
          </div>
          <label className="formula-tile-library-search">
            <Search size={14} aria-hidden="true" />
            <input
              type="search"
              value={query}
              placeholder={
                isEn
                  ? `Search ${formulaTileLibraryCount} formulas`
                  : `搜索 ${formulaTileLibraryCount} 条公式`
              }
              aria-label={isEn ? "Search formulas" : "搜索公式"}
              onChange={(event) => setQuery(event.currentTarget.value)}
            />
          </label>
          <button
            type="button"
            className="icon-button"
            onClick={onClose}
            aria-label={isEn ? "Close" : "关闭"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="formula-tile-library-body">
          <nav
            className="formula-tile-library-nav"
            aria-label={isEn ? "Formula categories" : "公式分类"}
          >
            <div className="formula-tile-library-subjects" role="tablist">
              {formulaTileSubjects.map((item) => (
                <button
                  key={item.id}
                  type="button"
                  role="tab"
                  aria-selected={!needle && subject === item.id}
                  className={!needle && subject === item.id ? "is-active" : ""}
                  data-formula-tile-subject={item.id}
                  onClick={() => {
                    setQuery("");
                    selectSubject(item.id);
                  }}
                >
                  {isEn ? item.labelEn : item.labelZh}
                </button>
              ))}
            </div>
            <div className="formula-tile-library-groups">
              {subjectGroups.map((group) => {
                const added = group.tiles.filter((tile) =>
                  selected.has(tile.id),
                ).length;
                return (
                  <button
                    key={group.id}
                    type="button"
                    className={
                      !needle && activeGroup?.id === group.id ? "is-active" : ""
                    }
                    data-formula-tile-group={group.id}
                    onClick={() => {
                      setQuery("");
                      setGroupId(group.id);
                    }}
                  >
                    <span className="formula-tile-library-group-name">
                      {isEn ? group.labelEn : group.labelZh}
                    </span>
                    <span className="formula-tile-library-group-meta">
                      {isEn
                        ? formulaTileLevelLabels[group.level].labelEn
                        : formulaTileLevelLabels[group.level].labelZh}
                      {" · "}
                      {added > 0 ? `${added}/` : ""}
                      {group.tiles.length}
                    </span>
                  </button>
                );
              })}
            </div>
          </nav>

          <div className="formula-tile-library-content">
            {visibleGroups.length === 0 && (
              <p className="formula-tile-library-empty">
                {isEn ? "No formula matches." : "没有匹配的公式。"}
              </p>
            )}
            {visibleGroups.map((group) => (
              <section key={group.id} className="formula-tile-library-section">
                <h3>
                  {isEn ? group.labelEn : group.labelZh}
                  <span>
                    {isEn
                      ? formulaTileLevelLabels[group.level].labelEn
                      : formulaTileLevelLabels[group.level].labelZh}
                  </span>
                </h3>
                <div className="formula-tile-library-grid">
                  {group.tiles.map((tile) => {
                    const isAdded = selected.has(tile.id);
                    const label = isEn ? tile.labelEn : tile.labelZh;
                    return (
                      <button
                        key={tile.id}
                        type="button"
                        className={
                          "formula-tile-library-item" +
                          (tile.latex.length > wideFormulaLength ? " is-wide" : "") +
                          (isAdded ? " is-added" : "")
                        }
                        aria-pressed={isAdded}
                        data-formula-library-tile={tile.id}
                        title={
                          isAdded
                            ? isEn
                              ? `${label} · click to remove`
                              : `${label} · 点击移除`
                            : isEn
                              ? `${label} · click to add`
                              : `${label} · 点击添加`
                        }
                        onClick={() => onToggle(tile.id)}
                      >
                        <MathPreview
                          latex={tile.latex}
                          className="formula-tile-library-preview"
                          fit
                          maximumFitScale={1}
                        />
                        <span className="formula-tile-library-item-label">
                          {label}
                        </span>
                        <span
                          className="formula-tile-library-item-state"
                          aria-hidden="true"
                        >
                          {isAdded ? <Check size={12} /> : <Plus size={12} />}
                        </span>
                      </button>
                    );
                  })}
                </div>
              </section>
            ))}
          </div>
        </div>

        <footer className="formula-tile-library-footer">
          <span>
            {isEn
              ? `${selectedIds.length} tiles in Common · click a formula to add or remove it`
              : `常用中已有 ${selectedIds.length} 个磁贴 · 点击公式即可添加或移除`}
          </span>
          <button type="button" className="secondary-button" onClick={onClose}>
            {isEn ? "Done" : "完成"}
          </button>
        </footer>
      </section>
    </div>,
    document.body,
  );
}
