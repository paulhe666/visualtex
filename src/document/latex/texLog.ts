// The messages of a TeX log that belong on source lines: errors ("file:line: message" — the engine runs with
// -file-line-error), warnings with "on input line N" (their file from the log's open-file parentheses), and
// overfull/underfull boxes. Paths come back as the project's (the warm engine's files are the main file).

export interface TexMessage {
  kind: "error" | "warning" | "box";
  /** the project path, or "" when the message is not in a project file */
  file: string;
  /** 1-based, 0: unknown */
  line: number;
  message: string;
  /** the source text TeX showed (l.23 …) */
  context?: string;
}

/**
 * `toProject`: a path as TeX printed it (./chap1.tex, /abs/…/x.sty) → the project's path, or null when it is not
 * a project file.
 */
export function parseTexLog(log: string, toProject: (path: string) => string | null): TexMessage[] {
  const out: TexMessage[] = [];
  const lines = log.replace(/\r\n?/g, "\n").split("\n");
  // the open files: "(" + a path opens, ")" closes
  const stack: string[] = [];
  const current = () => {
    for (let i = stack.length - 1; i >= 0; i--) {
      const p = toProject(stack[i]);
      if (p) return p;
    }
    return "";
  };
  const fileOpen = /\((\.{0,2}\/[^\s()]+|[A-Za-z]:[\\/][^\s()]+|[^\s()/]+\.(?:tex|sty|cls|cfg|def|clo|fd|ldf|bbl|aux|toc|ltx))/g;
  for (let i = 0; i < lines.length; i++) {
    const text = lines[i];
    const error = /^(.+?\.[A-Za-z0-9]+):(\d+): (.*)$/.exec(text);
    if (error) {
      const file = toProject(error[1]) ?? "";
      let message = error[3];
      // the message may continue on the next lines (up to a blank line or the "l.N" context)
      let context: string | undefined;
      for (let j = i + 1; j < Math.min(lines.length, i + 12); j++) {
        const next = lines[j];
        const at = /^l\.\d+\s?(.*)$/.exec(next);
        if (at) {
          context = (at[1] + " " + (lines[j + 1] ?? "").trim()).trim();
          break;
        }
        if (!next.trim() || next.startsWith("See the ") || next.startsWith("Type  H")) break;
        if (j === i + 1 && !/^[A-Z!]/.test(next)) message += " " + next.trim();
      }
      out.push({ kind: "error", file, line: Number(error[2]), message: message.replace(/^LaTeX Error: /, ""), context });
      continue;
    }
    const warning = /^(?:LaTeX|Package (\S+)|Class (\S+)) Warning: (.*)$/.exec(text);
    if (warning) {
      let message = warning[3];
      for (let j = i + 1; j < Math.min(lines.length, i + 6) && /^\(\S+\)\s/.test(lines[j]); j++) message += " " + lines[j].replace(/^\(\S+\)\s+/, "");
      const at = /on input line (\d+)\.?/.exec(message);
      out.push({ kind: "warning", file: at ? current() : "", line: at ? Number(at[1]) : 0, message: message.replace(/\s*on input line \d+\.?/, "").trim() });
    }
    const box = /^(Overfull|Underfull) \\[hv]box \((.*?)\) (?:in paragraph |detected )?at lines? (\d+)/.exec(text);
    if (box) out.push({ kind: "box", file: current(), line: Number(box[3]), message: `${box[1]} box (${box[2]})` });
    // the file stack, after the messages of the line
    let depthText = text;
    fileOpen.lastIndex = 0;
    let m: RegExpExecArray | null;
    const opens: Array<{ index: number; path: string }> = [];
    while ((m = fileOpen.exec(text))) opens.push({ index: m.index, path: m[1] });
    let o = 0;
    for (let k = 0; k < depthText.length; k++) {
      if (o < opens.length && opens[o].index === k) {
        stack.push(opens[o].path);
        k += opens[o].path.length;
        o++;
      } else if (depthText[k] === "(") stack.push("");
      else if (depthText[k] === ")" && stack.length) stack.pop();
    }
    depthText = "";
  }
  // one entry per place and text
  const seen = new Set<string>();
  return out.filter((m) => {
    const key = `${m.kind}|${m.file}|${m.line}|${m.message}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
