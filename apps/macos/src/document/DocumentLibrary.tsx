// The document library beside the page (document mode): folders, the documents newest first, search over titles
// and text, and the recently deleted ones. DocumentWorkspace owns the library; this draws it.
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Copy,
  Download,
  FilePlus2,
  FileText,
  FileUp,
  Folder,
  FolderInput,
  FolderPlus,
  FolderSearch,
  MoreHorizontal,
  PanelLeft,
  PanelLeftClose,
  Pencil,
  RotateCcw,
  Search,
  Trash2,
  X,
} from "lucide-react";
import { libraryDocumentName, useDocumentLibraryStore, type Library, type LibraryDocument } from "./documentLibraryApi";

export interface DocumentLibraryActions {
  open(id: string): void;
  create(folder: string | null): void;
  importFile(): void;
  rename(id: string, title: string): void;
  duplicate(id: string): void;
  move(id: string, folder: string | null): void;
  trash(id: string): void;
  restore(id: string): void;
  remove(id: string): void;
  emptyTrash(): void;
  /** Resolves to the new folder's id. */
  createFolder(): Promise<string | null>;
  renameFolder(id: string, name: string): void;
  deleteFolder(id: string): void;
  /** the document to a file: its own format, Markdown or PDF */
  exportFile(id: string, format?: DocumentExportFormat): void;
  /** Shows the file in Finder (the app only). */
  reveal: ((id: string) => void) | null;
  search(query: string): Promise<string[]>;
}

const DRAG_TYPE = "application/x-visualtex-library-document";

function shortTime(ms: number, isEn: boolean): string {
  const date = new Date(ms);
  const now = new Date();
  if (date.toDateString() === now.toDateString())
    return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
  if (date.getFullYear() === now.getFullYear())
    return isEn ? date.toLocaleDateString("en", { month: "short", day: "numeric" }) : `${date.getMonth() + 1}月${date.getDate()}日`;
  return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
}

/** The toolbar's first button while the library is hidden: shows it. */
export function DocumentLibraryToggle(props: { isEn: boolean }) {
  const setOpen = useDocumentLibraryStore((state) => state.setOpen);
  const label = props.isEn ? "Documents" : "文档库";
  return (
    <>
      <button
        type="button"
        className="icon-button compact formula-formatting-button document-library-toggle"
        title={label}
        aria-label={label}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => setOpen(true)}
      >
        <PanelLeft size={16} strokeWidth={2} />
      </button>
      <span className="formula-formatting-divider" aria-hidden="true" />
    </>
  );
}

function InlineName(props: { value: string; placeholder: string; onDone: (value: string | null) => void }) {
  const ref = useRef<HTMLInputElement>(null);
  const done = useRef(false);
  useLayoutEffect(() => {
    ref.current?.focus();
    ref.current?.select();
  }, []);
  const finish = (value: string | null) => {
    if (done.current) return;
    done.current = true;
    props.onDone(value);
  };
  return (
    <input
      ref={ref}
      className="document-library-rename"
      defaultValue={props.value}
      placeholder={props.placeholder}
      onKeyDown={(event) => {
        if (event.key === "Enter") finish(event.currentTarget.value);
        else if (event.key === "Escape") finish(null);
        event.stopPropagation();
      }}
      onBlur={(event) => finish(event.currentTarget.value)}
      onClick={(event) => event.stopPropagation()}
    />
  );
}

type MenuTarget = { kind: "document" | "folder"; id: string; x: number; y: number; moving?: boolean; exporting?: boolean };

export type DocumentExportFormat = "vtxdoc" | "markdown" | "pdf";

const EXPORT_FORMATS: Array<{ id: DocumentExportFormat; en: string; zh: string; extension: string }> = [
  { id: "pdf", en: "PDF", zh: "PDF", extension: ".pdf" },
  { id: "markdown", en: "Markdown", zh: "Markdown", extension: ".md" },
  { id: "vtxdoc", en: "VisualTeX document", zh: "VisualTeX 文档", extension: ".vtxdoc" },
];

function Menu(props: { target: MenuTarget; onClose: () => void; children: ReactNode }) {
  const ref = useRef<HTMLDivElement>(null);
  const [place, setPlace] = useState({ left: props.target.x, top: props.target.y });
  useLayoutEffect(() => {
    const panel = ref.current;
    if (!panel) return;
    const { offsetWidth: width, offsetHeight: height } = panel;
    setPlace({
      left: Math.max(4, Math.min(props.target.x, window.innerWidth - width - 4)),
      top: Math.max(4, Math.min(props.target.y, window.innerHeight - height - 4)),
    });
  }, [props.target, props.children]);
  useEffect(() => {
    const away = (event: MouseEvent) => {
      if (!ref.current?.contains(event.target as globalThis.Node)) props.onClose();
    };
    const key = (event: KeyboardEvent) => event.key === "Escape" && props.onClose();
    window.addEventListener("mousedown", away, true);
    window.addEventListener("keydown", key, true);
    window.addEventListener("blur", props.onClose);
    return () => {
      window.removeEventListener("mousedown", away, true);
      window.removeEventListener("keydown", key, true);
      window.removeEventListener("blur", props.onClose);
    };
  }, [props.onClose]);
  return (
    <div ref={ref} className="document-library-menu" role="menu" style={{ left: place.left, top: place.top }}>
      {props.children}
    </div>
  );
}

function MenuItem(props: { icon?: ReactNode; label: string; onClick: () => void; danger?: boolean; checked?: boolean }) {
  return (
    <button
      type="button"
      role="menuitem"
      className={"document-library-menu-item" + (props.danger ? " is-danger" : "") + (props.checked ? " is-checked" : "")}
      onClick={props.onClick}
    >
      <span className="document-library-menu-icon">{props.icon}</span>
      <span>{props.label}</span>
    </button>
  );
}

export function DocumentLibrary(props: {
  isEn: boolean;
  library: Library | null;
  currentId: string | null;
  actions: DocumentLibraryActions;
}) {
  const { isEn, library, currentId, actions } = props;
  const t = (en: string, zh: string) => (isEn ? en : zh);
  const setOpen = useDocumentLibraryStore((state) => state.setOpen);
  const collapsed = useDocumentLibraryStore((state) => state.collapsed);
  const toggleFolder = useDocumentLibraryStore((state) => state.toggleFolder);
  const [query, setQuery] = useState("");
  const [found, setFound] = useState<string[] | null>(null);
  const [renaming, setRenaming] = useState<{ kind: "document" | "folder"; id: string } | null>(null);
  const [menu, setMenu] = useState<MenuTarget | null>(null);
  const [trashOpen, setTrashOpen] = useState(false);
  const [dropTarget, setDropTarget] = useState<string | null>(null);
  const listRef = useRef<HTMLDivElement>(null);

  // the text search runs a moment after typing stops; titles match at once
  useEffect(() => {
    const words = query.trim();
    if (!words) {
      setFound(null);
      return;
    }
    let live = true;
    const timer = window.setTimeout(() => {
      void actions.search(words).then(
        (ids) => live && setFound(ids),
        () => live && setFound([]),
      );
    }, 200);
    return () => {
      live = false;
      window.clearTimeout(timer);
    };
    // the actions object changes every render; the query decides
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query, library]);

  const documents = useMemo(() => library?.documents.filter((doc) => doc.trashed === null) ?? [], [library]);
  const trashed = useMemo(
    () => (library?.documents.filter((doc) => doc.trashed !== null) ?? []).sort((a, b) => (b.trashed ?? 0) - (a.trashed ?? 0)),
    [library],
  );
  const folders = useMemo(
    () => [...(library?.folders ?? [])].sort((a, b) => a.name.localeCompare(b.name, isEn ? "en" : "zh-Hans-CN")),
    [library, isEn],
  );
  const results = useMemo(() => {
    const words = query.trim().toLowerCase().split(/\s+/).filter(Boolean);
    if (!words.length) return null;
    return documents.filter(
      (doc) => found?.includes(doc.id) || words.every((word) => libraryDocumentName(doc, isEn).toLowerCase().includes(word)),
    );
  }, [documents, found, query, isEn]);

  // the open document stays in view
  useEffect(() => {
    listRef.current?.querySelector(".document-library-item.is-current")?.scrollIntoView({ block: "nearest" });
  }, [currentId, library]);

  const closeMenu = () => setMenu(null);
  const openMenu = (event: React.MouseEvent, kind: MenuTarget["kind"], id: string) => {
    event.preventDefault();
    event.stopPropagation();
    const anchor = (event.currentTarget as HTMLElement).getBoundingClientRect();
    setMenu(event.type === "contextmenu" ? { kind, id, x: event.clientX, y: event.clientY } : { kind, id, x: anchor.left, y: anchor.bottom + 2 });
  };

  const drop = (folder: string | null) => ({
    onDragOver: (event: React.DragEvent) => {
      if (!event.dataTransfer.types.includes(DRAG_TYPE)) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      setDropTarget(folder ?? "");
    },
    onDragLeave: (event: React.DragEvent) => {
      if (!(event.currentTarget as HTMLElement).contains(event.relatedTarget as globalThis.Node | null)) setDropTarget(null);
    },
    onDrop: (event: React.DragEvent) => {
      const id = event.dataTransfer.getData(DRAG_TYPE);
      setDropTarget(null);
      if (!id) return;
      event.preventDefault();
      event.stopPropagation();
      actions.move(id, folder);
    },
  });

  const documentRow = (doc: LibraryDocument, nested: boolean) => {
    const current = doc.id === currentId;
    const name = libraryDocumentName(doc, isEn);
    return (
      <div
        key={doc.id}
        role="treeitem"
        aria-selected={current}
        className={"document-library-item" + (current ? " is-current" : "") + (nested ? " is-nested" : "") + (doc.title ? "" : " is-untitled")}
        draggable={renaming?.id !== doc.id}
        onDragStart={(event) => {
          event.dataTransfer.setData(DRAG_TYPE, doc.id);
          event.dataTransfer.effectAllowed = "move";
        }}
        onDragEnd={() => setDropTarget(null)}
        onClick={() => {
          actions.open(doc.id);
          // over the page (a narrow window) the library steps aside
          if (window.matchMedia("(max-width: 760px)").matches) setOpen(false);
        }}
        onDoubleClick={() => setRenaming({ kind: "document", id: doc.id })}
        onContextMenu={(event) => openMenu(event, "document", doc.id)}
        title={name}
      >
        <FileText size={14} className="document-library-item-icon" />
        {renaming?.kind === "document" && renaming.id === doc.id ? (
          <InlineName
            value={doc.titleSet ? doc.title : ""}
            placeholder={doc.title || t("Untitled", "未命名文档")}
            onDone={(value) => {
              setRenaming(null);
              if (value !== null && value.trim() !== (doc.titleSet ? doc.title : "")) actions.rename(doc.id, value.trim());
            }}
          />
        ) : (
          <>
            <span className="document-library-item-name">{name}</span>
            <span className="document-library-item-time">{shortTime(doc.modified, isEn)}</span>
            <button
              type="button"
              className="document-library-item-more"
              aria-label={t("More", "更多")}
              title={t("More", "更多")}
              onClick={(event) => openMenu(event, "document", doc.id)}
            >
              <MoreHorizontal size={14} />
            </button>
          </>
        )}
      </div>
    );
  };

  const menuDocument = menu?.kind === "document" ? library?.documents.find((doc) => doc.id === menu.id) : undefined;
  const menuFolder = menu?.kind === "folder" ? library?.folders.find((folder) => folder.id === menu.id) : undefined;

  return (
    <aside className="document-library" aria-label={t("Documents", "文档库")}>
      <div className="document-library-head">
        <span className="document-library-heading">{t("Documents", "文档库")}</span>
        <button type="button" className="document-library-icon" title={t("New document (⌘N)", "新建文档 (⌘N)")} aria-label={t("New document", "新建文档")} onClick={() => actions.create(null)}>
          <FilePlus2 size={15} />
        </button>
        <button
          type="button"
          className="document-library-icon"
          title={t("New folder", "新建文件夹")}
          aria-label={t("New folder", "新建文件夹")}
          onClick={() => void actions.createFolder().then((id) => id && setRenaming({ kind: "folder", id }))}
        >
          <FolderPlus size={15} />
        </button>
        <button type="button" className="document-library-icon" title={t("Import (⌘O)", "导入 (⌘O)")} aria-label={t("Import", "导入")} onClick={() => actions.importFile()}>
          <FileUp size={15} />
        </button>
        <button type="button" className="document-library-icon" title={t("Hide", "收起")} aria-label={t("Hide", "收起")} onClick={() => setOpen(false)}>
          <PanelLeftClose size={15} />
        </button>
      </div>
      <label className="document-library-search">
        <Search size={13} />
        <input
          type="search"
          value={query}
          placeholder={t("Search", "搜索")}
          aria-label={t("Search", "搜索")}
          onChange={(event) => setQuery(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Escape") setQuery("");
            else if (event.key === "Enter" && results?.length) actions.open(results[0].id);
          }}
        />
        {query && (
          <button type="button" className="document-library-search-clear" aria-label={t("Clear", "清除")} title={t("Clear", "清除")} onClick={() => setQuery("")}>
            <X size={12} />
          </button>
        )}
      </label>
      <div className="document-library-list" role="tree" ref={listRef}>
        {results ? (
          results.length ? (
            results.map((doc) => documentRow(doc, false))
          ) : (
            <div className="document-library-empty">{t("No results", "无结果")}</div>
          )
        ) : (
          <>
            {folders.map((folder) => {
              const open = !collapsed.includes(folder.id);
              const inside = documents.filter((doc) => doc.folder === folder.id);
              return (
                <div key={folder.id} className="document-library-group" role="group" {...drop(folder.id)}>
                  <div
                    role="treeitem"
                    aria-expanded={open}
                    className={"document-library-folder" + (dropTarget === folder.id ? " is-drop" : "")}
                    onClick={() => renaming?.id !== folder.id && toggleFolder(folder.id)}
                    onDoubleClick={() => setRenaming({ kind: "folder", id: folder.id })}
                    onContextMenu={(event) => openMenu(event, "folder", folder.id)}
                    title={folder.name}
                  >
                    {open ? <ChevronDown size={13} className="document-library-chevron" /> : <ChevronRight size={13} className="document-library-chevron" />}
                    <Folder size={14} className="document-library-item-icon" />
                    {renaming?.kind === "folder" && renaming.id === folder.id ? (
                      <InlineName
                        value={folder.name}
                        placeholder={t("Folder", "文件夹")}
                        onDone={(value) => {
                          setRenaming(null);
                          if (value?.trim() && value.trim() !== folder.name) actions.renameFolder(folder.id, value.trim());
                        }}
                      />
                    ) : (
                      <>
                        <span className="document-library-item-name">{folder.name}</span>
                        <span className="document-library-item-time">{inside.length || ""}</span>
                        <button
                          type="button"
                          className="document-library-item-more"
                          aria-label={t("More", "更多")}
                          title={t("More", "更多")}
                          onClick={(event) => openMenu(event, "folder", folder.id)}
                        >
                          <MoreHorizontal size={14} />
                        </button>
                      </>
                    )}
                  </div>
                  {open && inside.map((doc) => documentRow(doc, true))}
                </div>
              );
            })}
            <div className={"document-library-root" + (dropTarget === "" ? " is-drop" : "")} {...drop(null)}>
              {documents.filter((doc) => !doc.folder).map((doc) => documentRow(doc, false))}
            </div>
          </>
        )}
      </div>
      {trashed.length > 0 && !results && (
        <div className={"document-library-trash" + (trashOpen ? " is-open" : "")}>
          <div className="document-library-folder document-library-trash-head" role="button" aria-expanded={trashOpen} onClick={() => setTrashOpen(!trashOpen)}>
            {trashOpen ? <ChevronDown size={13} className="document-library-chevron" /> : <ChevronRight size={13} className="document-library-chevron" />}
            <Trash2 size={14} className="document-library-item-icon" />
            <span className="document-library-item-name">{t("Recently deleted", "最近删除")}</span>
            <span className="document-library-item-time">{trashed.length}</span>
            {trashOpen && (
              <button
                type="button"
                className="document-library-text-button"
                onClick={(event) => {
                  event.stopPropagation();
                  actions.emptyTrash();
                }}
              >
                {t("Empty", "清空")}
              </button>
            )}
          </div>
          {trashOpen && (
            <div className="document-library-trash-list">
              {trashed.map((doc) => (
                <div key={doc.id} className="document-library-item is-nested is-trashed" title={libraryDocumentName(doc, isEn)}>
                  <FileText size={14} className="document-library-item-icon" />
                  <span className="document-library-item-name">{libraryDocumentName(doc, isEn)}</span>
                  <button type="button" className="document-library-item-action" title={t("Restore", "恢复")} aria-label={t("Restore", "恢复")} onClick={() => actions.restore(doc.id)}>
                    <RotateCcw size={13} />
                  </button>
                  <button type="button" className="document-library-item-action" title={t("Delete now", "立即删除")} aria-label={t("Delete now", "立即删除")} onClick={() => actions.remove(doc.id)}>
                    <X size={13} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
      {menu && menuDocument && (
        <Menu target={menu} onClose={closeMenu}>
          {menu.exporting ? (
            <>
              <MenuItem icon={<ChevronLeft size={14} />} label={t("Export", "导出")} onClick={() => setMenu({ ...menu, exporting: false })} />
              <div className="document-library-menu-rule" />
              {EXPORT_FORMATS.map((format) => (
                <MenuItem
                  key={format.id}
                  label={`${t(format.en, format.zh)} (${format.extension})`}
                  onClick={() => {
                    closeMenu();
                    actions.exportFile(menuDocument.id, format.id);
                  }}
                />
              ))}
            </>
          ) : menu.moving ? (
            <>
              <MenuItem icon={<ChevronLeft size={14} />} label={t("Move to", "移到")} onClick={() => setMenu({ ...menu, moving: false })} />
              <div className="document-library-menu-rule" />
              <MenuItem
                label={t("No folder", "不放入文件夹")}
                checked={!menuDocument.folder}
                onClick={() => {
                  closeMenu();
                  actions.move(menuDocument.id, null);
                }}
              />
              {folders.map((folder) => (
                <MenuItem
                  key={folder.id}
                  icon={<Folder size={14} />}
                  label={folder.name}
                  checked={menuDocument.folder === folder.id}
                  onClick={() => {
                    closeMenu();
                    actions.move(menuDocument.id, folder.id);
                  }}
                />
              ))}
            </>
          ) : (
            <>
              <MenuItem
                icon={<Pencil size={14} />}
                label={t("Rename", "重命名")}
                onClick={() => {
                  closeMenu();
                  setRenaming({ kind: "document", id: menuDocument.id });
                }}
              />
              <MenuItem
                icon={<Copy size={14} />}
                label={t("Duplicate", "创建副本")}
                onClick={() => {
                  closeMenu();
                  actions.duplicate(menuDocument.id);
                }}
              />
              {folders.length > 0 && <MenuItem icon={<FolderInput size={14} />} label={t("Move to", "移到")} onClick={() => setMenu({ ...menu, moving: true })} />}
              <MenuItem
                icon={<Download size={14} />}
                label={t("Export", "导出")}
                onClick={() => setMenu({ ...menu, exporting: true })}
              />
              {actions.reveal && (
                <MenuItem
                  icon={<FolderSearch size={14} />}
                  label={t("Show in Finder", "在访达中显示")}
                  onClick={() => {
                    closeMenu();
                    actions.reveal?.(menuDocument.id);
                  }}
                />
              )}
              <div className="document-library-menu-rule" />
              <MenuItem
                icon={<Trash2 size={14} />}
                label={t("Delete", "删除")}
                danger
                onClick={() => {
                  closeMenu();
                  actions.trash(menuDocument.id);
                }}
              />
            </>
          )}
        </Menu>
      )}
      {menu && menuFolder && (
        <Menu target={menu} onClose={closeMenu}>
          <MenuItem
            icon={<FilePlus2 size={14} />}
            label={t("New document", "新建文档")}
            onClick={() => {
              closeMenu();
              actions.create(menuFolder.id);
            }}
          />
          <MenuItem
            icon={<Pencil size={14} />}
            label={t("Rename", "重命名")}
            onClick={() => {
              closeMenu();
              setRenaming({ kind: "folder", id: menuFolder.id });
            }}
          />
          <div className="document-library-menu-rule" />
          <MenuItem
            icon={<Trash2 size={14} />}
            label={t("Delete folder", "删除文件夹")}
            danger
            onClick={() => {
              closeMenu();
              actions.deleteFolder(menuFolder.id);
            }}
          />
        </Menu>
      )}
    </aside>
  );
}
