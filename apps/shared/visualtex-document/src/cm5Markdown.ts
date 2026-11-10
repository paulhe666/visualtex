// The Markdown tokenizer of the source editor: CodeMirror 5's "markdown" mode (MIT licence, copyright (c) Marijn
// Haverbeke and others; https://codemirror.net/5/LICENSE), the tokenizer Obsidian's editor is built on, ported to
// TypeScript with a small line runner in place of CodeMirror. Settings as Obsidian has them: formatting characters as
// tokens of their own ("formatting formatting-strong strong"), task lists, strikethrough. Changed: `~~` only (one "~"
// is text, as in Obsidian), `==highlight==` added, code spans classed "inline-code" and code block lines
// "hmd-codeblock" (Obsidian's names), no HTML mode (inline HTML is the source editor's own), and next to Chinese,
// Japanese or Korean text punctuation does not stop "*" / "_" from opening or closing ("前**粗体。**后" is bold —
// the one place the source editor reads Markdown otherwise than Obsidian, on the user's word, 2026-11-10).

/* eslint-disable @typescript-eslint/no-explicit-any */

/** CodeMirror's StringStream: one line, a cursor in it. */
export class StringStream {
  pos = 0;
  start = 0;
  constructor(
    public string: string,
    private next_lines: (n: number) => string | undefined,
    public tabSize = 4,
  ) {}
  eol() {
    return this.pos >= this.string.length;
  }
  sol() {
    return this.pos === 0;
  }
  peek(): string | undefined {
    return this.string.charAt(this.pos) || undefined;
  }
  next(): string | undefined {
    if (this.pos < this.string.length) return this.string.charAt(this.pos++);
    return undefined;
  }
  eat(match: string | RegExp | ((ch: string) => boolean)): string | undefined {
    const ch = this.string.charAt(this.pos);
    const ok = typeof match === "string" ? ch === match : ch && (match instanceof RegExp ? match.test(ch) : match(ch));
    if (ok) {
      this.pos++;
      return ch;
    }
    return undefined;
  }
  eatWhile(match: string | RegExp | ((ch: string) => boolean)): boolean {
    const start = this.pos;
    while (this.eat(match));
    return this.pos > start;
  }
  eatSpace(): boolean {
    const start = this.pos;
    while (/[\s ]/.test(this.string.charAt(this.pos))) this.pos++;
    return this.pos > start;
  }
  skipToEnd() {
    this.pos = this.string.length;
  }
  backUp(n: number) {
    this.pos -= n;
  }
  column(): number {
    let col = 0;
    for (let i = 0; i < this.start; i++) col += this.string[i] === "\t" ? this.tabSize - (col % this.tabSize) : 1;
    return col;
  }
  match(pattern: string | RegExp, consume = true, caseInsensitive = false): any {
    if (typeof pattern === "string") {
      const cased = (s: string) => (caseInsensitive ? s.toLowerCase() : s);
      const substr = this.string.substr(this.pos, pattern.length);
      if (cased(substr) === cased(pattern)) {
        if (consume !== false) this.pos += pattern.length;
        return true;
      }
      return null;
    }
    const m = this.string.slice(this.pos).match(pattern);
    if (m && m.index! > 0) return null;
    if (m && consume !== false) this.pos += m[0].length;
    return m;
  }
  current(): string {
    return this.string.slice(this.start, this.pos);
  }
  lookAhead(n: number): string | undefined {
    return this.next_lines(n);
  }
}

const CJK = /[ᄀ-ᇿ⺀-⿿　-〿぀-ヿ㄀-ㇿ㈀-㏿㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿︰-﹏＀-￯]|[\ud840-\ud87f]/;

const hrRE = /^([*\-_])(?:\s*\1){2,}\s*$/;
const listRE = /^(?:[*\-+]|^[0-9]+([.)]))\s+/;
const taskListRE = /^\[(x| )\](?=\s)/i;
const atxHeaderRE = /^(#+)(?: |$)/;
const setextHeaderRE = /^ {0,3}(?:={1,}|-{2,})\s*$/;
const textRE = /^[^#!\[\]*_\\<>` "'(~:=]+/;
const fencedCodeRE = /^(~~~+|```+)[ \t]*([\w\/+#-]*)[^\n`]*$/;
const linkDefRE = /^\s*\[[^\]]+?\]:.*$/;
const punctuation = /[!"#$%&'()*+,\-.\/:;<=>?@\[\\\]^_`{|}~\xA1\xA7\xAB\xB6\xB7\xBB\xBF\u037E\u0387\u055A-\u055F\u0589\u058A\u05BE\u05C0\u05C3\u05C6\u05F3\u05F4\u0609\u060A\u060C\u060D\u061B\u061E\u061F\u066A-\u066D\u06D4\u0700-\u070D\u07F7-\u07F9\u0830-\u083E\u085E\u0964\u0965\u0970\u0AF0\u0DF4\u0E4F\u0E5A\u0E5B\u0F04-\u0F12\u0F14\u0F3A-\u0F3D\u0F85\u0FD0-\u0FD4\u0FD9\u0FDA\u104A-\u104F\u10FB\u1360-\u1368\u1400\u166D\u166E\u169B\u169C\u16EB-\u16ED\u1735\u1736\u17D4-\u17D6\u17D8-\u17DA\u1800-\u180A\u1944\u1945\u1A1E\u1A1F\u1AA0-\u1AA6\u1AA8-\u1AAD\u1B5A-\u1B60\u1BFC-\u1BFF\u1C3B-\u1C3F\u1C7E\u1C7F\u1CC0-\u1CC7\u1CD3\u2010-\u2027\u2030-\u2043\u2045-\u2051\u2053-\u205E\u207D\u207E\u208D\u208E\u2308-\u230B\u2329\u232A\u2768-\u2775\u27C5\u27C6\u27E6-\u27EF\u2983-\u2998\u29D8-\u29DB\u29FC\u29FD\u2CF9-\u2CFC\u2CFE\u2CFF\u2D70\u2E00-\u2E2E\u2E30-\u2E42\u3001-\u3003\u3008-\u3011\u3014-\u301F\u3030\u303D\u30A0\u30FB\uA4FE\uA4FF\uA60D-\uA60F\uA673\uA67E\uA6F2-\uA6F7\uA874-\uA877\uA8CE\uA8CF\uA8F8-\uA8FA\uA8FC\uA92E\uA92F\uA95F\uA9C1-\uA9CD\uA9DE\uA9DF\uAA5C-\uAA5F\uAADE\uAADF\uAAF0\uAAF1\uABEB\uFD3E\uFD3F\uFE10-\uFE19\uFE30-\uFE52\uFE54-\uFE61\uFE63\uFE68\uFE6A\uFE6B\uFF01-\uFF03\uFF05-\uFF0A\uFF0C-\uFF0F\uFF1A\uFF1B\uFF1F\uFF20\uFF3B-\uFF3D\uFF3F\uFF5B\uFF5D\uFF5F-\uFF65]|\uD800[\uDD00-\uDD02\uDF9F\uDFD0]|\uD801\uDD6F|\uD802[\uDC57\uDD1F\uDD3F\uDE50-\uDE58\uDE7F\uDEF0-\uDEF6\uDF39-\uDF3F\uDF99-\uDF9C]|\uD804[\uDC47-\uDC4D\uDCBB\uDCBC\uDCBE-\uDCC1\uDD40-\uDD43\uDD74\uDD75\uDDC5-\uDDC9\uDDCD\uDDDB\uDDDD-\uDDDF\uDE38-\uDE3D\uDEA9]|\uD805[\uDCC6\uDDC1-\uDDD7\uDE41-\uDE43\uDF3C-\uDF3E]|\uD809[\uDC70-\uDC74]|\uD81A[\uDE6E\uDE6F\uDEF5\uDF37-\uDF3B\uDF44]|\uD82F\uDC9F|\uD836[\uDE87-\uDE8B]/;
const expandedTab = "    ";

/** The tokenizer's state between lines (and tokens). */
export type MarkdownState = any;

function switchInline(stream: StringStream, state: MarkdownState, f: any) {
  state.f = state.inline = f;
  return f(stream, state);
}

function switchBlock(stream: StringStream, state: MarkdownState, f: any) {
  state.f = state.block = f;
  return f(stream, state);
}

function lineIsEmpty(line: StringStream | null) {
  return !line || !/\S/.test(line.string);
}

export function blankLine(state: MarkdownState) {
  state.linkTitle = false;
  state.linkHref = false;
  state.linkText = false;
  state.em = false;
  state.strong = false;
  state.strikethrough = false;
  state.highlight = false;
  state.quote = 0;
  state.indentedCode = false;
  state.trailingSpace = 0;
  state.trailingSpaceNewLine = false;
  state.prevLine = state.thisLine;
  state.thisLine = { stream: null };
  return null;
}

function blockNormal(stream: StringStream, state: MarkdownState): string | null {
  const firstTokenOnLine = stream.column() === state.indentation;
  const prevLineLineIsEmpty = lineIsEmpty(state.prevLine.stream);
  const prevLineIsIndentedCode = state.indentedCode;
  const prevLineIsHr = state.prevLine.hr;
  const prevLineIsList = state.list !== false;
  const maxNonCodeIndentation = (state.listStack[state.listStack.length - 1] || 0) + 3;
  state.indentedCode = false;
  const lineIndentation = state.indentation;
  if (state.indentationDiff === null) {
    state.indentationDiff = state.indentation;
    if (prevLineIsList) {
      state.list = null;
      while (lineIndentation < state.listStack[state.listStack.length - 1]) {
        state.listStack.pop();
        if (state.listStack.length) state.indentation = state.listStack[state.listStack.length - 1];
        else state.list = false;
      }
      if (state.list !== false) state.indentationDiff = lineIndentation - state.listStack[state.listStack.length - 1];
    }
  }
  const allowsInlineContinuation =
    !prevLineLineIsEmpty && !prevLineIsHr && !state.prevLine.header && (!prevLineIsList || !prevLineIsIndentedCode) && !state.prevLine.fencedCodeEnd;
  const isHr = (state.list === false || prevLineIsHr || prevLineLineIsEmpty) && state.indentation <= maxNonCodeIndentation && stream.match(hrRE);
  let match: any = null;
  if (state.indentationDiff >= 4 && (prevLineIsIndentedCode || state.prevLine.fencedCodeEnd || state.prevLine.header || prevLineLineIsEmpty)) {
    stream.skipToEnd();
    state.indentedCode = true;
    return "comment hmd-codeblock hmd-indented-code";
  } else if (stream.eatSpace()) {
    return null;
  } else if (firstTokenOnLine && state.indentation <= maxNonCodeIndentation && (match = stream.match(atxHeaderRE)) && match[1].length <= 6) {
    state.quote = 0;
    state.header = match[1].length;
    state.thisLine.header = true;
    state.formatting = "header";
    state.f = state.inline;
    return getType(state);
  } else if (state.indentation <= maxNonCodeIndentation && stream.eat(">")) {
    state.quote = firstTokenOnLine ? 1 : state.quote + 1;
    state.formatting = "quote";
    stream.eatSpace();
    return getType(state);
  } else if (!isHr && !state.setext && firstTokenOnLine && state.indentation <= maxNonCodeIndentation && (match = stream.match(listRE))) {
    const listType = match[1] ? "ol" : "ul";
    state.indentation = lineIndentation + stream.current().length;
    state.list = true;
    state.quote = 0;
    state.listStack.push(state.indentation);
    state.em = false;
    state.strong = false;
    state.code = false;
    state.strikethrough = false;
    state.highlight = false;
    if (stream.match(taskListRE, false)) state.taskList = true;
    state.f = state.inline;
    state.formatting = ["list", "list-" + listType];
    return getType(state);
  } else if (firstTokenOnLine && state.indentation <= maxNonCodeIndentation && (match = stream.match(fencedCodeRE, true))) {
    state.quote = 0;
    state.fencedEndRE = new RegExp(match[1] + "+ *$");
    state.fencedLang = match[2] || "";
    state.f = state.block = local;
    state.formatting = "code-block";
    state.code = -1;
    state.thisLine.fencedCodeBegin = true;
    return getType(state);
  } else if (
    state.setext ||
    ((!allowsInlineContinuation || !prevLineIsList) &&
      !state.quote &&
      state.list === false &&
      !state.code &&
      !isHr &&
      !linkDefRE.test(stream.string) &&
      (match = stream.lookAhead(1)) &&
      (match = match.match(setextHeaderRE)))
  ) {
    if (!state.setext) {
      state.header = match[0].charAt(0) === "=" ? 1 : 2;
      state.setext = state.header;
    } else {
      state.header = state.setext;
      state.setext = 0;
      stream.skipToEnd();
      state.formatting = "header";
    }
    state.thisLine.header = true;
    state.f = state.inline;
    return getType(state);
  } else if (isHr) {
    stream.skipToEnd();
    state.hr = true;
    state.thisLine.hr = true;
    return "hr";
  } else if (stream.peek() === "[") {
    return switchInline(stream, state, footnoteLink);
  }
  return switchInline(stream, state, state.inline);
}

function local(stream: StringStream, state: MarkdownState): string | null {
  const currListInd = state.listStack[state.listStack.length - 1] || 0;
  const hasExitedList = state.indentation < currListInd;
  const maxFencedEndInd = currListInd + 3;
  if (state.fencedEndRE && state.indentation <= maxFencedEndInd && (hasExitedList || stream.match(state.fencedEndRE))) {
    state.formatting = "code-block";
    let returnType: string | null = null;
    if (!hasExitedList) returnType = getType(state);
    state.block = blockNormal;
    state.f = inlineNormal;
    state.fencedEndRE = null;
    state.code = 0;
    state.thisLine.fencedCodeEnd = true;
    if (hasExitedList) return switchBlock(stream, state, state.block);
    return returnType;
  }
  stream.skipToEnd();
  return "comment hmd-codeblock";
}

export function getType(state: MarkdownState): string | null {
  const styles: string[] = [];
  if (state.formatting) {
    styles.push("formatting");
    if (typeof state.formatting === "string") state.formatting = [state.formatting];
    for (const f of state.formatting) {
      styles.push("formatting-" + f);
      if (f === "header") styles.push("formatting-header-" + state.header);
      if (f === "quote") styles.push("formatting-quote-" + state.quote);
    }
  }
  if (state.taskOpen) {
    styles.push("meta");
    return styles.join(" ");
  }
  if (state.taskClosed) {
    styles.push("property");
    return styles.join(" ");
  }
  if (state.linkHref) styles.push("string", "url");
  else {
    if (state.strong) styles.push("strong");
    if (state.em) styles.push("em");
    if (state.strikethrough) styles.push("strikethrough");
    if (state.highlight) styles.push("highlight");
    if (state.linkText) styles.push("link");
    if (state.code === -1) styles.push("comment", "hmd-codeblock");
    else if (state.code) styles.push("inline-code");
    if (state.image) styles.push("image");
    if (state.imageAltText) styles.push("image-alt-text", "link");
    if (state.imageMarker) styles.push("image-marker");
  }
  if (state.header) styles.push("header", "header-" + state.header);
  if (state.quote) styles.push("quote", "quote-" + state.quote);
  if (state.list !== false) styles.push("list-" + state.listStack.length);
  return styles.length ? styles.join(" ") : null;
}

function handleText(stream: StringStream, state: MarkdownState) {
  if (stream.match(textRE, true)) return getType(state);
  return undefined;
}

/** CommonMark's flanking, with the source editor's one change: CJK neighbours make punctuation count as letters. */
function flanking(before: string, after: string) {
  const cjk = CJK.test(before) || CJK.test(after);
  const pBefore = !cjk && punctuation.test(before);
  const pAfter = !cjk && punctuation.test(after);
  const left = !/\s/.test(after) && (!pAfter || /\s/.test(before) || pBefore);
  const right = !/\s/.test(before) && (!pBefore || /\s/.test(after) || pAfter);
  return { left, right, pBefore, pAfter };
}

function inlineNormal(stream: StringStream, state: MarkdownState): string | null {
  const style = state.text(stream, state);
  if (typeof style !== "undefined") return style;
  if (state.list) {
    state.list = null;
    return getType(state);
  }
  if (state.taskList) {
    const taskOpen = stream.match(taskListRE, true)[1] === " ";
    if (taskOpen) state.taskOpen = true;
    else state.taskClosed = true;
    state.formatting = "task";
    state.taskList = false;
    return getType(state);
  }
  state.taskOpen = false;
  state.taskClosed = false;
  if (state.header && stream.match(/^#+$/, true)) {
    state.formatting = "header";
    return getType(state);
  }
  const ch = stream.next();
  if (state.linkTitle) {
    state.linkTitle = false;
    let matchCh = ch;
    if (ch === "(") matchCh = ")";
    matchCh = (matchCh + "").replace(/([.?*+^\[\]\\(){}|-])/g, "\\$1");
    const regex = "^\\s*(?:[^" + matchCh + "\\\\]+|\\\\\\\\|\\\\.)" + matchCh;
    if (stream.match(new RegExp(regex), true)) return "string";
  }
  if (ch === "`") {
    const previousFormatting = state.formatting;
    state.formatting = "code";
    stream.eatWhile("`");
    const count = stream.current().length;
    if (state.code == 0 && (!state.quote || count == 1)) {
      state.code = count;
      return getType(state);
    } else if (count == state.code) {
      const t = getType(state);
      state.code = 0;
      return t;
    } else {
      state.formatting = previousFormatting;
      return getType(state);
    }
  } else if (state.code) {
    return getType(state);
  }
  // (an escape is "\" before ASCII punctuation, as CommonMark and Obsidian have it; "\begin" is text)
  if (ch === "\\" && /[!-\/:-@\[-`{-~]/.test(stream.peek() ?? "")) {
    stream.next();
    const type = getType(state);
    return type ? type + " formatting-escape" : "formatting-escape";
  }
  if (ch === "!" && stream.match(/\[[^\]]*\] ?(?:\(|\[)/, false)) {
    state.imageMarker = true;
    state.image = true;
    state.formatting = "image";
    return getType(state);
  }
  if (ch === "[" && state.imageMarker && stream.match(/[^\]]*\](\(.*?\)| ?\[.*?\])/, false)) {
    state.imageMarker = false;
    state.imageAltText = true;
    state.formatting = "image";
    return getType(state);
  }
  if (ch === "]" && state.imageAltText) {
    state.formatting = "image";
    const type = getType(state);
    state.imageAltText = false;
    state.image = false;
    state.inline = state.f = linkHref;
    return type;
  }
  if (ch === "[" && !state.image) {
    if (state.linkText && stream.match(/^.*?\]/)) return getType(state);
    state.linkText = true;
    state.formatting = "link";
    return getType(state);
  }
  if (ch === "]" && state.linkText) {
    state.formatting = "link";
    const type = getType(state);
    state.linkText = false;
    state.inline = state.f = stream.match(/\(.*?\)| ?\[.*?\]/, false) ? linkHref : inlineNormal;
    return type;
  }
  if (ch === "<" && stream.match(/^(https?|ftps?):\/\/(?:[^\\>]|\\.)+>/, false)) {
    state.f = state.inline = linkInline;
    state.formatting = "link";
    const type = getType(state);
    return (type ? type + " " : "") + "link";
  }
  if (ch === "<" && stream.match(/^[^> \\]+@(?:[^\\>]|\\.)+>/, false)) {
    state.f = state.inline = linkInline;
    state.formatting = "link";
    const type = getType(state);
    return (type ? type + " " : "") + "link";
  }
  if (ch === "*" || ch === "_") {
    let len = 1;
    const before = stream.pos == 1 ? " " : stream.string.charAt(stream.pos - 2);
    while (len < 3 && stream.eat(ch)) len++;
    const after = stream.peek() || " ";
    const { left: leftFlanking, right: rightFlanking, pBefore, pAfter } = flanking(before, after);
    let setEm: boolean | null = null;
    let setStrong: boolean | null = null;
    if (len % 2) {
      if (!state.em && leftFlanking && (ch === "*" || !rightFlanking || pBefore)) setEm = true;
      else if (state.em == ch && rightFlanking && (ch === "*" || !leftFlanking || pAfter)) setEm = false;
    }
    if (len > 1) {
      if (!state.strong && leftFlanking && (ch === "*" || !rightFlanking || pBefore)) setStrong = true;
      else if (state.strong == ch && rightFlanking && (ch === "*" || !leftFlanking || pAfter)) setStrong = false;
    }
    if (setStrong != null || setEm != null) {
      state.formatting = setEm == null ? "strong" : setStrong == null ? "em" : "strong em";
      if (setEm === true) state.em = ch;
      if (setStrong === true) state.strong = ch;
      const t = getType(state);
      if (setEm === false) state.em = false;
      if (setStrong === false) state.strong = false;
      return t;
    }
  } else if (ch === " ") {
    if (stream.eat("*") || stream.eat("_")) {
      if (stream.peek() === " ") return getType(state);
      stream.backUp(1);
    }
  }
  // "~~" (exactly two) and "==": strikethrough and highlight
  if ((ch === "~" || ch === "=") && stream.peek() === ch) {
    const start = stream.pos - 1;
    stream.eatWhile(ch);
    const run = stream.pos - start;
    const key = ch === "~" ? "strikethrough" : "highlight";
    if (run === 2) {
      if (state[key]) {
        state.formatting = key;
        const t = getType(state);
        state[key] = false;
        return t;
      } else if (stream.match(/^[^\s]/, false)) {
        state[key] = true;
        state.formatting = key;
        return getType(state);
      }
    }
    return getType(state);
  } else if (ch === " ") {
    if (stream.match("~~", true) || stream.match("==", true)) {
      if (stream.peek() === " ") return getType(state);
      stream.backUp(2);
    }
  }
  if (ch === " ") {
    if (stream.match(/^ +$/, false)) state.trailingSpace++;
    else if (state.trailingSpace) state.trailingSpaceNewLine = true;
  }
  return getType(state);
}

function linkInline(stream: StringStream, state: MarkdownState): string | null {
  const ch = stream.next();
  if (ch === ">") {
    state.f = state.inline = inlineNormal;
    state.formatting = "link";
    const type = getType(state);
    return (type ? type + " " : "") + "link";
  }
  stream.match(/^[^>]+/, true);
  return "link url";
}

function linkHref(stream: StringStream, state: MarkdownState): string | null {
  if (stream.eatSpace()) return null;
  const ch = stream.next();
  if (ch === "(" || ch === "[") {
    state.f = state.inline = getLinkHrefInside(ch === "(" ? ")" : "]");
    state.formatting = "link-string";
    state.linkHref = true;
    return getType(state);
  }
  return "error";
}

const linkRE: Record<string, RegExp> = {
  ")": /^(?:[^\\\(\)]|\\.|\((?:[^\\\(\)]|\\.)*\))*?(?=\))/,
  "]": /^(?:[^\\\[\]]|\\.|\[(?:[^\\\[\]]|\\.)*\])*?(?=\])/,
};

function getLinkHrefInside(endChar: string) {
  return function inside(stream: StringStream, state: MarkdownState): string | null {
    const ch = stream.next();
    if (ch === endChar) {
      state.f = state.inline = inlineNormal;
      state.formatting = "link-string";
      const returnState = getType(state);
      state.linkHref = false;
      return returnState;
    }
    stream.match(linkRE[endChar]);
    state.linkHref = true;
    return getType(state);
  };
}

function footnoteLink(stream: StringStream, state: MarkdownState): string | null {
  if (stream.match(/^([^\]\\]|\\.)*\]:/, false)) {
    state.f = footnoteLinkInside;
    stream.next();
    state.formatting = "link";
    state.linkText = true;
    return getType(state);
  }
  return switchInline(stream, state, inlineNormal);
}

function footnoteLinkInside(stream: StringStream, state: MarkdownState): string | null {
  if (stream.match("]:", true)) {
    state.f = state.inline = footnoteUrl;
    state.formatting = "link";
    const returnType = getType(state);
    state.linkText = false;
    return returnType;
  }
  stream.match(/^([^\]\\]|\\.)+/, true);
  return "link";
}

function footnoteUrl(stream: StringStream, state: MarkdownState): string | null {
  if (stream.eatSpace()) return null;
  stream.match(/^[^\s]+/, true);
  if (stream.peek() === undefined) state.linkTitle = true;
  else stream.match(/^(?:\s+(?:"(?:[^"\\]|\\.)+"|'(?:[^'\\]|\\.)+'|\((?:[^)\\]|\\.)+\)))?/, true);
  state.f = state.inline = inlineNormal;
  return "string url";
}

export function startState(): MarkdownState {
  return {
    f: blockNormal,
    prevLine: { stream: null },
    thisLine: { stream: null },
    block: blockNormal,
    indentation: 0,
    inline: inlineNormal,
    text: handleText,
    formatting: false,
    linkText: false,
    linkHref: false,
    linkTitle: false,
    code: 0,
    em: false,
    strong: false,
    header: 0,
    setext: 0,
    hr: false,
    taskList: false,
    list: false,
    listStack: [],
    quote: 0,
    trailingSpace: 0,
    trailingSpaceNewLine: false,
    strikethrough: false,
    highlight: false,
    fencedEndRE: null,
    fencedLang: "",
  };
}

export function copyState(s: MarkdownState): MarkdownState {
  return {
    f: s.f,
    prevLine: s.prevLine,
    thisLine: s.thisLine,
    block: s.block,
    indentation: s.indentation,
    inline: s.inline,
    text: s.text,
    formatting: false,
    linkText: s.linkText,
    linkTitle: s.linkTitle,
    linkHref: s.linkHref,
    code: s.code,
    em: s.em,
    strong: s.strong,
    strikethrough: s.strikethrough,
    highlight: s.highlight,
    header: s.header,
    setext: s.setext,
    hr: s.hr,
    taskList: s.taskList,
    list: s.list,
    listStack: s.listStack.slice(0),
    quote: s.quote,
    indentedCode: s.indentedCode,
    trailingSpace: s.trailingSpace,
    trailingSpaceNewLine: s.trailingSpaceNewLine,
    fencedEndRE: s.fencedEndRE,
    fencedLang: s.fencedLang,
  };
}

function token(stream: StringStream, state: MarkdownState): string | null {
  state.formatting = false;
  if (stream !== state.thisLine.stream) {
    state.header = 0;
    state.hr = false;
    if (stream.match(/^\s*$/, true)) {
      blankLine(state);
      return null;
    }
    state.prevLine = state.thisLine;
    state.thisLine = { stream };
    state.taskList = false;
    // (Obsidian's tokenizer: a code span of one or two backticks does not go on into the next line)
    if (state.code === 1 || state.code === 2) state.code = 0;
    state.trailingSpace = 0;
    state.trailingSpaceNewLine = false;
    state.f = state.block;
    const indentation = stream.match(/^\s*/, true)[0].replace(/\t/g, expandedTab).length;
    state.indentation = indentation;
    state.indentationDiff = null;
    if (indentation > 0) return "hmd-indent";
  }
  return state.f(stream, state);
}

export interface MarkdownToken {
  from: number;
  to: number;
  /** the token's classes ("formatting formatting-strong strong") */
  type: string;
}

/** What the tokenizer knew of a line once it was read. */
export interface LineFacts {
  header: number;
  /** list depth (0: not in a list) */
  list: number;
  quote: number;
  /** a fenced code block: this line opens / closes it, or is inside it */
  codeBegin: boolean;
  codeEnd: boolean;
  codeInside: boolean;
  /** an indented code line */
  indentedCode: boolean;
  codeLang: string;
  hr: boolean;
  blank: boolean;
  /** a line of the document's front matter (set by the document's reading, not the tokenizer) */
  frontMatter?: boolean;
}

/**
 * One line read with the state the lines before it left (the state is changed in place): its tokens and facts.
 * `ahead(n)` gives the n-th line after it (setext headings look one line ahead).
 */
export function tokenizeLine(text: string, state: MarkdownState, ahead: (n: number) => string | undefined): { tokens: MarkdownToken[]; facts: LineFacts } {
  const tokens: MarkdownToken[] = [];
  const inCodeBefore = state.code === -1;
  const lang = state.fencedLang;
  if (!/\S/.test(text)) {
    // CodeMirror reads a blank line with blankLine (inside a fenced block it stays a code line)
    blankLine(state);
    return { tokens, facts: { header: 0, list: 0, quote: 0, codeBegin: false, codeEnd: false, codeInside: inCodeBefore, indentedCode: false, codeLang: lang, hr: false, blank: true } };
  }
  const stream = new StringStream(text, ahead);
  let header = 0;
  let quote = 0;
  while (!stream.eol()) {
    stream.start = stream.pos;
    // (as CodeMirror's readToken: a step that reads nothing changed the state; it is tried again, ten times at most)
    let type: string | null = null;
    for (let i = 0; i < 10; i++) {
      type = token(stream, state);
      if (stream.pos > stream.start) break;
    }
    if (stream.pos <= stream.start) stream.pos = stream.start + 1;
    header = Math.max(header, state.header || 0);
    quote = Math.max(quote, state.quote || 0);
    const last = tokens[tokens.length - 1];
    const t = type ?? "";
    if (last && last.type === t && last.to === stream.start && !/formatting/.test(t)) last.to = stream.pos;
    else tokens.push({ from: stream.start, to: stream.pos, type: t });
  }
  const thisLine = state.thisLine;
  return {
    tokens,
    facts: {
      header,
      list: state.list !== false ? state.listStack.length : 0,
      quote,
      codeBegin: !!thisLine.fencedCodeBegin,
      codeEnd: !!thisLine.fencedCodeEnd,
      codeInside: inCodeBefore && !thisLine.fencedCodeEnd,
      indentedCode: !!state.indentedCode,
      codeLang: thisLine.fencedCodeBegin ? state.fencedLang : lang,
      hr: !!thisLine.hr,
      blank: false,
    },
  };
}

/** A signature of the state between lines: two lines read alike when it and their text are the same. */
export function stateKey(state: MarkdownState): string {
  return [
    state.block === local ? "L" : "N",
    state.code,
    state.em || "",
    state.strong || "",
    state.strikethrough ? 1 : 0,
    state.highlight ? 1 : 0,
    state.list === false ? "f" : state.list === null ? "n" : "t",
    state.listStack.join(","),
    state.quote,
    state.indentation,
    state.indentedCode ? 1 : 0,
    state.linkTitle ? 1 : 0,
    state.linkText ? 1 : 0,
    state.linkHref ? 1 : 0,
    state.setext,
    state.fencedEndRE ? state.fencedEndRE.source : "",
    state.fencedLang,
    lineIsEmpty(state.thisLine.stream) ? "b" : "",
    state.thisLine.header ? "h" : "",
    state.thisLine.hr ? "r" : "",
    state.thisLine.fencedCodeEnd ? "e" : "",
  ].join("|");
}
