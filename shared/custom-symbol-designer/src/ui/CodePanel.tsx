import { useEffect, useState } from "react";
import { Check, ChevronDown, ChevronUp, Copy } from "lucide-react";

interface Props {
  tex: string;
  svg: string;
  omissions: string[];
  isEn: boolean;
}

async function copyText(text: string) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    const area = document.createElement("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.append(area);
    area.select();
    const copied = document.execCommand("copy");
    area.remove();
    return copied;
  }
}

/** The symbol as code anyone can paste: a LaTeX \newcommand, or SVG. */
export function CodePanel({ tex, svg, omissions, isEn }: Props) {
  const [open, setOpen] = useState(true);
  const [tab, setTab] = useState<"tex" | "svg">("tex");
  const [copied, setCopied] = useState(false);
  const code = tab === "tex" ? tex : svg;

  useEffect(() => {
    if (!copied) return;
    const timer = window.setTimeout(() => setCopied(false), 1400);
    return () => window.clearTimeout(timer);
  }, [copied]);

  return (
    <section className={`csd-code${open ? " is-open" : ""}`} data-csd-code>
      <header>
        <div className="csd-tabs is-inline" role="tablist">
          <button type="button" role="tab" aria-selected={tab === "tex"} className={tab === "tex" ? "is-active" : ""} data-csd-code-tab="tex"
            onClick={() => { setTab("tex"); setOpen(true); }}>
            LaTeX \newcommand
          </button>
          <button type="button" role="tab" aria-selected={tab === "svg"} className={tab === "svg" ? "is-active" : ""} data-csd-code-tab="svg"
            onClick={() => { setTab("svg"); setOpen(true); }}>
            SVG
          </button>
        </div>
        <span className="csd-code-note">
          {tab === "tex"
            ? isEn ? "picture + graphicx, packages on the second line · follows the font size" : "picture + graphicx，所需宏包见第二行 · 随字号缩放"
            : isEn ? "1000 units per em · baseline-aligned" : "每 em 1000 单位 · 与基线对齐"}
        </span>
        <button type="button" className="csd-copy" disabled={!code} data-csd-copy
          onClick={() => void copyText(code).then(setCopied)}>
          {copied ? <Check size={13} /> : <Copy size={13} />}
          {copied ? (isEn ? "Copied" : "已复制") : isEn ? "Copy" : "复制"}
        </button>
        <button type="button" className="csd-code-toggle" aria-expanded={open} title={open ? (isEn ? "Collapse" : "收起") : isEn ? "Expand" : "展开"}
          onClick={() => setOpen(!open)}>
          {open ? <ChevronDown size={14} /> : <ChevronUp size={14} />}
        </button>
      </header>
      {open ? (
        <>
          <pre data-csd-code-text={tab}><code>{code || (isEn ? "Add a symbol to see its code." : "加入符号后在这里显示代码。")}</code></pre>
          {tab === "tex" && omissions.length ? (
            <p className="csd-code-omissions">
              {isEn ? "Not in the LaTeX code: " : "LaTeX 代码不含："}
              {omissions.join("; ")}
            </p>
          ) : null}
        </>
      ) : null}
    </section>
  );
}
