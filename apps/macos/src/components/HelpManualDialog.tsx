import {
  Fragment,
  useEffect,
  useMemo,
  useRef,
  type ReactNode,
} from "react";
import { BookOpenText, X } from "lucide-react";
import manualMarkdown from "../../docs/help/VisualTeX_帮助手册.md?raw";

interface Props {
  open: boolean;
  language: "cn" | "en";
  onClose: () => void;
}

type MarkdownBlock =
  | { type: "heading"; level: number; text: string; id: string }
  | { type: "paragraph"; text: string }
  | { type: "unordered-list"; items: string[] }
  | { type: "ordered-list"; items: string[] }
  | { type: "code"; language: string; text: string }
  | { type: "table"; rows: string[][] }
  | { type: "rule" };

interface TocEntry {
  id: string;
  level: number;
  text: string;
}

function stripInlineMarkdown(value: string) {
  return value
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .trim();
}

function slugifyHeading(value: string, index: number) {
  const normalized = stripInlineMarkdown(value)
    .toLocaleLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "-")
    .replace(/^-+|-+$/g, "");
  return normalized ? `help-${normalized}-${index}` : `help-section-${index}`;
}

function parseTableRow(line: string) {
  return line
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

function isTableDivider(line: string) {
  const cells = parseTableRow(line);
  return (
    cells.length > 0 &&
    cells.every((cell) => /^:?-{3,}:?$/.test(cell.replace(/\s+/g, "")))
  );
}

function isBlockStart(lines: string[], index: number) {
  const trimmed = lines[index]?.trim() ?? "";
  if (!trimmed) return true;
  if (/^(#{1,6})\s+/.test(trimmed)) return true;
  if (/^```/.test(trimmed)) return true;
  if (/^---+$/.test(trimmed)) return true;
  if (/^[-*]\s+/.test(trimmed)) return true;
  if (/^\d+\.\s+/.test(trimmed)) return true;
  if (
    trimmed.startsWith("|") &&
    index + 1 < lines.length &&
    isTableDivider(lines[index + 1].trim())
  ) {
    return true;
  }
  return false;
}

function parseManual(markdown: string): MarkdownBlock[] {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  const blocks: MarkdownBlock[] = [];
  let index = 0;
  let headingIndex = 0;

  while (index < lines.length) {
    const trimmed = lines[index].trim();

    if (!trimmed) {
      index += 1;
      continue;
    }

    if (/^```/.test(trimmed)) {
      const language = trimmed.slice(3).trim();
      const code: string[] = [];
      index += 1;
      while (index < lines.length && !/^```/.test(lines[index].trim())) {
        code.push(lines[index]);
        index += 1;
      }
      if (index < lines.length) index += 1;
      blocks.push({ type: "code", language, text: code.join("\n") });
      continue;
    }

    const heading = /^(#{1,6})\s+(.+)$/.exec(trimmed);
    if (heading) {
      headingIndex += 1;
      blocks.push({
        type: "heading",
        level: heading[1].length,
        text: heading[2].trim(),
        id: slugifyHeading(heading[2], headingIndex),
      });
      index += 1;
      continue;
    }

    if (/^---+$/.test(trimmed)) {
      blocks.push({ type: "rule" });
      index += 1;
      continue;
    }

    if (/^[-*]\s+/.test(trimmed)) {
      const items: string[] = [];
      while (index < lines.length) {
        const match = /^[-*]\s+(.+)$/.exec(lines[index].trim());
        if (!match) break;
        items.push(match[1].trim());
        index += 1;
      }
      blocks.push({ type: "unordered-list", items });
      continue;
    }

    if (/^\d+\.\s+/.test(trimmed)) {
      const items: string[] = [];
      while (index < lines.length) {
        const match = /^\d+\.\s+(.+)$/.exec(lines[index].trim());
        if (!match) break;
        items.push(match[1].trim());
        index += 1;
      }
      blocks.push({ type: "ordered-list", items });
      continue;
    }

    if (
      trimmed.startsWith("|") &&
      index + 1 < lines.length &&
      isTableDivider(lines[index + 1].trim())
    ) {
      const rows = [parseTableRow(lines[index])];
      index += 2;
      while (index < lines.length && lines[index].trim().startsWith("|")) {
        rows.push(parseTableRow(lines[index]));
        index += 1;
      }
      blocks.push({ type: "table", rows });
      continue;
    }

    const paragraph = [trimmed];
    index += 1;
    while (index < lines.length && !isBlockStart(lines, index)) {
      paragraph.push(lines[index].trim());
      index += 1;
    }
    blocks.push({ type: "paragraph", text: paragraph.join(" ") });
  }

  return blocks;
}

function renderInline(value: string): ReactNode[] {
  const tokens = value.split(/(\*\*.*?\*\*|`[^`]*`)/g).filter(Boolean);
  return tokens.map((token, index) => {
    if (token.startsWith("**") && token.endsWith("**")) {
      return <strong key={`${token}-${index}`}>{token.slice(2, -2)}</strong>;
    }
    if (token.startsWith("`") && token.endsWith("`")) {
      return <code key={`${token}-${index}`}>{token.slice(1, -1)}</code>;
    }
    return <Fragment key={`${token}-${index}`}>{token}</Fragment>;
  });
}

function ManualBlock({ block }: { block: MarkdownBlock }) {
  if (block.type === "heading") {
    const content = renderInline(block.text);
    if (block.level === 1) return <h1 id={block.id}>{content}</h1>;
    if (block.level === 2) return <h2 id={block.id}>{content}</h2>;
    if (block.level === 3) return <h3 id={block.id}>{content}</h3>;
    return <h4 id={block.id}>{content}</h4>;
  }
  if (block.type === "paragraph") {
    return <p>{renderInline(block.text)}</p>;
  }
  if (block.type === "unordered-list") {
    return (
      <ul>
        {block.items.map((item, index) => (
          <li key={`${item}-${index}`}>{renderInline(item)}</li>
        ))}
      </ul>
    );
  }
  if (block.type === "ordered-list") {
    return (
      <ol>
        {block.items.map((item, index) => (
          <li key={`${item}-${index}`}>{renderInline(item)}</li>
        ))}
      </ol>
    );
  }
  if (block.type === "code") {
    return (
      <pre data-language={block.language || undefined}>
        <code>{block.text}</code>
      </pre>
    );
  }
  if (block.type === "table") {
    const [header, ...rows] = block.rows;
    return (
      <div className="help-manual-table-wrap">
        <table>
          <thead>
            <tr>
              {header.map((cell, index) => (
                <th key={`${cell}-${index}`}>{renderInline(cell)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, cellIndex) => (
                  <td key={`${cell}-${cellIndex}`}>{renderInline(cell)}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
  return <hr />;
}

export function HelpManualDialog({ open, language, onClose }: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const blocks = useMemo(() => parseManual(manualMarkdown), []);
  const toc = useMemo<TocEntry[]>(
    () =>
      blocks
        .filter(
          (block): block is Extract<MarkdownBlock, { type: "heading" }> =>
            block.type === "heading" && block.level >= 1 && block.level <= 3,
        )
        .map(({ id, level, text }) => ({ id, level, text: stripInlineMarkdown(text) })),
    [blocks],
  );
  const isEn = language === "en";

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => dialogRef.current?.focus());
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      event.preventDefault();
      onClose();
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => {
      window.cancelAnimationFrame(frame);
      window.removeEventListener("keydown", onKeyDown, true);
      previousFocusRef.current?.focus({ preventScroll: true });
    };
  }, [onClose, open]);

  if (!open) return null;

  const scrollToSection = (id: string) => {
    const target = document.getElementById(id);
    target?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  return (
    <div className="modal-backdrop help-manual-backdrop" role="presentation">
      <section
        ref={dialogRef}
        className="help-manual-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={isEn ? "VisualTeX Help Manual" : "VisualTeX 帮助手册"}
        tabIndex={-1}
      >
        <header className="help-manual-header">
          <div>
            <BookOpenText size={18} />
            <span>
              <strong>{isEn ? "Help Manual" : "帮助手册"}</strong>
              <small>VisualTeX 1.2.6 · macOS</small>
            </span>
          </div>
          <button
            type="button"
            className="icon-button"
            aria-label={isEn ? "Close help manual" : "关闭帮助手册"}
            onClick={onClose}
          >
            <X size={17} />
          </button>
        </header>

        <div className="help-manual-layout">
          <nav className="help-manual-toc" aria-label={isEn ? "Contents" : "目录"}>
            <strong>{isEn ? "Contents" : "目录"}</strong>
            <div>
              {toc.map((entry) => (
                <button
                  type="button"
                  key={entry.id}
                  data-level={entry.level}
                  onClick={() => scrollToSection(entry.id)}
                >
                  {entry.text}
                </button>
              ))}
            </div>
          </nav>

          <article className="help-manual-content">
            {blocks.map((block, index) => (
              <ManualBlock block={block} key={`${block.type}-${index}`} />
            ))}
          </article>
        </div>
      </section>
    </div>
  );
}
