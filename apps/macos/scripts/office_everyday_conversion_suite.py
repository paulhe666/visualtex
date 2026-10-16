#!/usr/bin/env python3
"""Everyday Word conversion suite for the installed VisualTeX on macOS.

Drives the real Microsoft Word with the installed VisualTeX add-in and the
running VisualTeX app, the way a user would: it builds .docx fixtures, presses
the VisualTeX ribbon through accessibility, waits for Word to settle, saves a
snapshot and checks the saved OOXML. It never replaces the installed add-ins
or the app.

Test files live in a folder inside the Office group container (the only place
sandboxed Word can save without a file-access prompt). Snapshots are copied to
--results and the folder is removed at the end unless --keep is given.

The Mac must be idle while it runs: Word is brought to the front for every
ribbon press. Copy/paste cases save and restore the clipboard with
--clipboard-tool (a vtclip binary); without one they are skipped.

usage: office_everyday_conversion_suite.py [--list] [--only ID[,ID]] [--keep]
                                           [--results DIR] [--clipboard-tool PATH]
"""
import argparse
import base64
import html
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
import zlib

HOME = os.path.expanduser("~")
OUT = os.path.join(HOME, "Library/Group Containers/UBF8T346G9.Office/VisualTeXConversionTests")
WORD = "Microsoft Word"

# Ribbon labels (customUI14.xml of the Word add-in).
REDRAW_DOC = ("重绘全文",)
REDRAW_DOC_IMAGE = REDRAW_DOC + ("全文重绘为 VisualTeX 图片公式",)
REDRAW_DOC_OMML = REDRAW_DOC + ("全文重绘为 Word OMML",)
RESTORE_DOC_IMAGE = REDRAW_DOC + ("将全文图片公式转为 LaTeX",)
RESTORE_DOC_OMML = REDRAW_DOC + ("将全文 OMML 公式转为 LaTeX",)
REDRAW_SEL = ("重绘所选",)
REDRAW_SEL_IMAGE = REDRAW_SEL + ("重绘为 VisualTeX 图片公式",)
REDRAW_SEL_OMML = REDRAW_SEL + ("重绘为 Word OMML",)
RESTORE_SEL_IMAGE = REDRAW_SEL + ("将图片公式转为 LaTeX",)
RESTORE_SEL_OMML = REDRAW_SEL + ("将 OMML 公式转为 LaTeX",)
CONVERT_DOC = ("全文格式转换",)
DOC_IMAGES_TO_OMML = CONVERT_DOC + ("全文图片公式 → Word OMML",)
DOC_OMML_TO_IMAGES = CONVERT_DOC + ("全文 OMML → 图片公式",)
SEL_TO_OMML = ("转为 Word OMML",)
SEL_TO_IMAGE = ("转为图片公式",)


# --------------------------------------------------------------------------
# osascript plumbing


def osa(source, *args, timeout=60, lang="AppleScript"):
    cmd = ["osascript", "-l", lang, "-", *[str(a) for a in args]]
    p = subprocess.run(cmd, input=source, capture_output=True, text=True, timeout=timeout)
    if p.returncode != 0:
        raise RuntimeError(p.stderr.strip() or f"osascript exit {p.returncode}")
    return p.stdout.rstrip("\n")


RIBBON_JXA = r"""
function run(argv) {
  const [appName, controlName, itemName] = argv;
  const norm = (s) => String(s || "").normalize("NFC");
  Application(appName).activate();
  const se = Application("System Events");
  const proc = se.processes.byName(appName);
  proc.frontmost = true;
  const roles = new Set(["AXButton", "AXMenuButton", "AXMenuItem", "AXPopUpButton"]);
  function find(el, wanted, depth) {
    let kids = [];
    try { kids = el.uiElements(); } catch (e) { return null; }
    for (const k of kids) {
      let role = "", name = "";
      try { role = k.role(); } catch (e) {}
      try { name = norm(k.name()); } catch (e) {}
      if (roles.has(role) && name === norm(wanted)) return k;
      if (depth < 8) { const f = find(k, wanted, depth + 1); if (f) return f; }
    }
    return null;
  }
  // Word rebuilds the ribbon after a VisualTeX command (Invalidate), so look
  // the window, tab and control up again on every attempt.
  let target = null, last = "";
  for (let attempt = 0; attempt < 6 && !target; attempt++) {
    if (attempt) delay(1);
    try {
      const standard = proc.windows().find((w) => { try { return w.subrole() === "AXStandardWindow" && w.tabGroups.length > 0; } catch (e) { return false; } });
      if (!standard) { last = "no document window"; continue; }
      const ribbon = standard.tabGroups[0];
      const tab = ribbon.radioButtons().find((b) => norm(b.name()) === "VisualTeX");
      if (!tab) { last = "no VisualTeX tab"; continue; }
      tab.actions.byName("AXPress").perform();
      for (let i = 0; i < 8 && !target; i++) { delay(0.3); target = find(ribbon, controlName, 0); }
      if (!target) last = "ribbon control not found: " + controlName;
    } catch (e) { last = String(e); }
  }
  if (!target) throw new Error(last);
  target.actions.byName("AXPress").perform();
  if (!itemName) return "pressed " + controlName;
  let item = null;
  for (let i = 0; i < 10 && !item; i++) {
    delay(0.3);
    for (const w of proc.windows()) { item = find(w, itemName, 0); if (item) break; }
  }
  if (!item) { se.keyCode(53); throw new Error("menu item not found: " + itemName); }
  item.actions.byName("AXPress").perform();
  return "pressed " + controlName + " > " + itemName;
}
"""

# Lists dialogs/sheets of a process; with "dismiss" presses the most neutral
# button. VisualTeX's own windows are only ever listed, never pressed.
DIALOG_JXA = r"""
function run(argv) {
  const [appName, mode] = argv;
  const se = Application("System Events");
  if (!se.processes.byName(appName).exists()) return "[]";
  const proc = se.processes.byName(appName);
  const out = [];
  function walk(el, depth, acc) {
    let kids = [];
    try { kids = el.uiElements(); } catch (e) { return; }
    for (const k of kids) {
      let role = "";
      try { role = k.role(); } catch (e) {}
      if (role === "AXStaticText" || role === "AXTextArea") {
        let v = ""; try { v = k.value(); } catch (e) {}
        if (v) acc.texts.push(String(v));
      }
      if (role === "AXButton") {
        let n = ""; try { n = k.name() || k.title() || k.description(); } catch (e) {}
        acc.buttons.push({ n: String(n || ""), el: k });
      }
      if (depth < 4) walk(k, depth + 1, acc);
    }
  }
  for (const w of proc.windows()) {
    let sr = ""; try { sr = w.subrole(); } catch (e) {}
    const cands = [];
    if (sr === "AXDialog" || sr === "AXSystemDialog") cands.push(w);
    try { for (const s of w.sheets()) cands.push(s); } catch (e) {}
    for (const c of cands) {
      const acc = { texts: [], buttons: [] };
      walk(c, 0, acc);
      // Window chrome is not a dialog (VisualTeX parks its editor as a panel).
      acc.buttons = acc.buttons.filter((b) => !/^(close|zoom|minimize|full ?screen) button$/i.test(b.n));
      if (!acc.buttons.length) continue;
      let title = ""; try { title = c.name() || ""; } catch (e) {}
      const entry = { title: String(title), texts: acc.texts, buttons: acc.buttons.map((b) => b.n) };
      if (mode === "dismiss") {
        const pref = ["确定", "OK", "好", "关闭", "Close", "取消", "Cancel", "否", "No", "不保存", "Don't Save"];
        let pick = null;
        for (const p of pref) { pick = acc.buttons.find((b) => b.n === p); if (pick) break; }
        if (pick) { pick.el.actions.byName("AXPress").perform(); entry.pressed = pick.n; delay(0.4); }
      }
      out.push(entry);
    }
  }
  return JSON.stringify(out);
}
"""


def press(path):
    return osa(RIBBON_JXA, WORD, *path, lang="JavaScript", timeout=60)


def dialogs(app=WORD, dismiss=True):
    try:
        return json.loads(osa(DIALOG_JXA, app, "dismiss" if dismiss else "list", lang="JavaScript", timeout=30))
    except Exception as error:  # accessibility hiccups must not abort a case
        return [{"error": str(error)}]


def word(body, *args, timeout=60):
    return osa('on run argv\ntell application "Microsoft Word"\n' + body + "\nend tell\nend run", *args, timeout=timeout)


def word_running():
    return osa('tell application "System Events" to return (exists process "Microsoft Word") as text') == "true"


def word_open(path):
    word("activate\nopen (POSIX file (item 1 of argv))\ndelay 1", path)
    dialogs()
    return word("return name of active document")


def word_new(text):
    return word(
        "activate\nset d to make new document\nset content of text object of d to (item 1 of argv)\nreturn name of d",
        text.replace("\n", "\r"),
    )


def word_close_all():
    try:
        word("close every document saving no")
    except Exception:
        pass
    dialogs()


def word_save_as(path):
    word("save as document 1 file name (item 1 of argv) file format format document", path)


def word_sig():
    try:
        return word(
            'set d to active document\nreturn ((count of inline shapes of d) as text) & "|" & ((count of math objects of d) as text) & "|" & (content of text object of d)',
            timeout=20,
        )
    except Exception as error:
        return "busy:" + str(error)[:80]


def word_select_paragraph(index):
    word("select (text object of paragraph (item 1 of argv as integer) of active document)", index)


def word_select_shape(index):
    word("select (text object of inline shape (item 1 of argv as integer) of active document)", index)


def word_select_math(index):
    word("select (text range of math object (item 1 of argv as integer) of active document)", index)


def word_undo(times=1):
    return word("return (undo active document times (item 1 of argv as integer)) as text", times)


def word_copy_shape_to_end(index):
    word(
        "set d to active document\n"
        "copy object (text object of inline shape (item 1 of argv as integer) of d)\n"
        "set e to (end of content of text object of d) - 1\n"
        "paste object (create range d start e end e)",
        index,
    )


# --------------------------------------------------------------------------
# .docx fixtures


def esc(text):
    return html.escape(text, quote=False)


def run_xml(text, bold=False):
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def para_xml(parts, style=None, numbered=None, center=False):
    ppr = ""
    if style:
        ppr += f'<w:pStyle w:val="{style}"/>'
    if numbered:
        ppr += f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{numbered}"/></w:numPr>'
    if center:
        ppr += '<w:jc w:val="center"/>'
    body = ""
    for part in parts if isinstance(parts, list) else [parts]:
        if isinstance(part, str):
            body += run_xml(part)
        elif part[0] == "b":
            body += run_xml(part[1], bold=True)
        elif part[0] == "m":
            body += f"<m:oMath>{part[1]}</m:oMath>"
        elif part[0] == "mp":
            body += f"<m:oMathPara><m:oMath>{part[1]}</m:oMath></m:oMathPara>"
    return f"<w:p>{'<w:pPr>' + ppr + '</w:pPr>' if ppr else ''}{body}</w:p>"


def table_xml(rows):
    cols = max(len(r) for r in rows)
    border = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join('<w:gridCol w:w="3000"/>' for _ in range(cols))
    trs = ""
    for row in rows:
        tcs = "".join(f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{para_xml(cell)}</w:tc>' for cell in row)
        trs += f"<w:tr>{tcs}</w:tr>"
    return f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>{border}</w:tblBorders></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>'


NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)
STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
    "</w:styles>"
)
NUMBERING = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS}>'
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
    '<w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
    '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
    "</w:numbering>"
)


def write_docx(path, blocks):
    body = "".join(blocks)
    document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}<w:sectPr/></w:body></w:document>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
            "</Types>",
        )
        z.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            "</Relationships>",
        )
        z.writestr(
            "word/_rels/document.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
            "</Relationships>",
        )
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)


# OMML builders for equations as Word's own equation editor writes them.
def mr(text, normal=False):
    rpr = "<m:rPr><m:nor/></m:rPr>" if normal else ""
    return f'<m:r>{rpr}<m:t xml:space="preserve">{esc(text)}</m:t></m:r>'


def mf(num, den):
    return f"<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>"


def msup(base, sup):
    return f"<m:sSup><m:e>{base}</m:e><m:sup>{sup}</m:sup></m:sSup>"


def msub(base, sub):
    return f"<m:sSub><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>"


def mrad(e):
    return f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{e}</m:e></m:rad>'


def mnary(chr_, sub, sup, e, und_ovr=False):
    loc = '<m:limLoc m:val="undOvr"/>' if und_ovr else '<m:limLoc m:val="subSup"/>'
    return f'<m:nary><m:naryPr><m:chr m:val="{chr_}"/>{loc}</m:naryPr><m:sub>{sub}</m:sub><m:sup>{sup}</m:sup><m:e>{e}</m:e></m:nary>'


def md(e, beg="(", end=")"):
    return f'<m:d><m:dPr><m:begChr m:val="{beg}"/><m:endChr m:val="{end}"/></m:dPr><m:e>{e}</m:e></m:d>'


def mfunc(name_xml, arg):
    return f"<m:func><m:fName>{name_xml}</m:fName><m:e>{arg}</m:e></m:func>"


def mfn(name):
    return f'<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>{name}</m:t></m:r>'


def macc(chr_, e):
    return f'<m:acc><m:accPr><m:chr m:val="{chr_}"/></m:accPr><m:e>{e}</m:e></m:acc>'


def mbar(e):
    return f'<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>{e}</m:e></m:bar>'


def mmat(rows):
    return "<m:m>" + "".join("<m:mr>" + "".join(f"<m:e>{c}</m:e>" for c in row) + "</m:mr>" for row in rows) + "</m:m>"


def mlimlow(e, lim):
    return f"<m:limLow><m:e>{e}</m:e><m:lim>{lim}</m:lim></m:limLow>"


# --------------------------------------------------------------------------
# reading saved documents


def decode_descr(descr):
    try:
        s = descr.split(":", 3)[3]
        raw = base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))
        return json.loads(zlib.decompress(raw, -15).decode("utf8"))
    except Exception:
        return None


STRUCT = {"f", "sSup", "sSub", "sSubSup", "rad", "nary", "d", "m", "acc", "bar", "func", "limLow", "limUpp", "eqArr", "groupChr", "box", "borderBox", "sPre"}


def omml_info(xml):
    names = [n for n in re.findall(r"<m:(\w+)[ >/]", xml) if n in STRUCT]
    text = "".join(html.unescape(t) for t in re.findall(r"<m:t(?: [^>]*)?>([^<]*)</m:t>", xml))
    return {"struct": names, "text": text}


TOKEN = re.compile(
    r"<m:oMathPara\b.*?</m:oMathPara>|<m:oMath\b.*?</m:oMath>|<w:drawing>.*?</w:drawing>|<w:t(?: [^>]*)?>([^<]*)</w:t>|<w:tab/>",
    re.S,
)


def read_docx(path):
    """Paragraphs of word/document.xml in order: context, text with formula
    markers, and the formulas (image metadata or OMML structure)."""
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf8")
    tables = [(m.start(), m.end()) for m in re.finditer(r"<w:tbl>.*?</w:tbl>", xml, re.S)]
    paras = []
    for m in re.finditer(r"<w:p\b(?:(?!<w:p\b).)*?</w:p>", xml, re.S):
        p = m.group(0)
        ctx = "table" if any(a <= m.start() < b for a, b in tables) else "body"
        if "<w:numPr>" in p:
            ctx += "/list"
        style = re.search(r'<w:pStyle w:val="([^"]+)"', p)
        if style and style.group(1).lower().startswith("heading"):
            ctx += "/heading"
        text, formulas = "", []
        for t in TOKEN.finditer(p):
            s = t.group(0)
            if s.startswith("<m:oMathPara"):
                formulas.append({"kind": "omml", "mode": "display", **omml_info(s)})
                text += "⟦F⟧"
            elif s.startswith("<m:oMath"):
                formulas.append({"kind": "omml", "mode": "inline", **omml_info(s)})
                text += "⟦F⟧"
            elif s.startswith("<w:drawing"):
                docpr = re.search(r"<wp:docPr\b[^>]*>", s)
                attrs = {k: html.unescape(v) for k, v in re.findall(r'(\w+)="([^"]*)"', docpr.group(0))} if docpr else {}
                meta = decode_descr(attrs.get("descr", "")) or {}
                formulas.append({
                    "kind": "image",
                    "mode": "display" if meta.get("displayMode") == "block" else ("inline" if meta else "?"),
                    "latex": meta.get("latex"),
                    "id": meta.get("formulaId"),
                    "title": attrs.get("title", ""),
                })
                text += "⟦F⟧"
            elif s == "<w:tab/>":
                text += "\t"
            else:
                text += html.unescape(t.group(1) or "")
        paras.append({"ctx": ctx, "text": text, "formulas": formulas})
    return paras


# --------------------------------------------------------------------------
# expectations from LaTeX source text

SRC_FORMULA = re.compile(r"\$\$(.+?)\$\$|\\\[(.+?)\\\]|\$(.+?)\$|\\\((.+?)\\\)")


def source_formulas(text):
    out = []
    for m in SRC_FORMULA.finditer(text):
        if m.group(1) is not None or m.group(2) is not None:
            out.append(("display", m.group(1) if m.group(1) is not None else m.group(2)))
        else:
            out.append(("inline", m.group(3) if m.group(3) is not None else m.group(4)))
    return out


def source_marked(text):
    return SRC_FORMULA.sub("⟦F⟧", text)


def norm_latex(s):
    return re.sub(r"\s+", "", s or "")


def canon_latex(s):
    """LaTeX with rewrites that do not change the rendered meaning folded away:
    spacing, braces around one token, \\mathrm{ab} vs \\mathrm{a}\\mathrm{b},
    \\to/\\rightarrow, pmatrix/cases written as \\left..\\begin{matrix},
    f^{\\prime}/f', and \\lim\\limits_ (how VisualTeX reads Word limits back,
    approved 2026-10-15) vs \\lim_."""
    s = (s or "").strip()
    s = re.sub(r"^\\\[(.*)\\\]$", r"\1", s, flags=re.S)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"\\[,;:!]", "", s)
    for a, b in (("\\limits", ""), ("\\rightarrow", "\\to"), ("\\geq", "\\ge"), ("\\leq", "\\le"), ("\\neq", "\\ne"), ("⋅", "\\cdot"), ("−", "-")):
        s = s.replace(a, b)
    s = re.sub(r"\\mathrm\{([^{}\\]*)\}", lambda m: "".join(f"\\mathrm{{{ch}}}" if ch.isalpha() else ch for ch in m.group(1)), s)
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"([\^_])\{(\\[A-Za-z]+|[^{}\\])\}", r"\1\2", s)
    s = re.sub(r"\\left\(\\begin\{matrix\}(.*?)\\end\{matrix\}\\right\)", r"\\begin{pmatrix}\1\\end{pmatrix}", s)
    s = re.sub(r"\\left\\\{\\begin\{matrix\}(.*?)\\end\{matrix\}\\right\.", r"\\begin{cases}\1\\end{cases}", s)
    s = re.sub(r"\^(?:\{((?:\\prime)+)\}|((?:\\prime)+)(?![A-Za-z]))", lambda m: "'" * (m.group(1) or m.group(2)).count("\\prime"), s)
    return s.replace("^'", "'")


def latex_diffs(pairs, canonical):
    """(source, got) pairs that differ, exactly or after canon_latex."""
    f = canon_latex if canonical else norm_latex
    return [(a, b) for a, b in pairs if f(a) != f(b)]


def show(pairs, limit=8):
    return "; ".join(f"{a!r} -> {b!r}" for a, b in pairs[:limit]) + (f" (+{len(pairs) - limit} more)" if len(pairs) > limit else "")


def canonical_delimiters(text):
    """Source text with \\(..\\) as $..$ and \\[..\\] as $$..$$, whitespace removed."""
    def rep(m):
        if m.group(1) is not None or m.group(2) is not None:
            return "$$" + (m.group(1) if m.group(1) is not None else m.group(2)) + "$$"
        return "$" + (m.group(3) if m.group(3) is not None else m.group(4)) + "$"
    return norm_latex(SRC_FORMULA.sub(rep, text))


# --------------------------------------------------------------------------
# test context


class Case:
    def __init__(self, cid, title, results_dir):
        self.id, self.title, self.results_dir = cid, title, results_dir
        self.checks, self.notes, self.steps = [], [], []
        self.snap_index = 0

    def check(self, name, ok, detail=""):
        self.checks.append({"name": name, "ok": bool(ok), "detail": detail})
        mark = "PASS" if ok else "FAIL"
        print(f"    [{mark}] {name}" + (f" — {detail}" if detail and not ok else ""), flush=True)
        return ok

    def note(self, text):
        self.notes.append(text)
        print(f"    note: {text}", flush=True)

    def snapshot(self, label):
        self.snap_index += 1
        path = os.path.join(OUT, f"{self.id}-{self.snap_index:02d}-{label}.docx")
        word_save_as(path)
        time.sleep(0.3)
        shutil.copy2(path, os.path.join(self.results_dir, os.path.basename(path)))
        return read_docx(path)

    def act(self, label, path=None, action=None, timeout=60, expect_change=True, quiet=2.5):
        """Press a ribbon command (or run action) and wait until Word settles."""
        before = word_sig()
        t0 = time.time()
        pressed = press(path) if path else (action() or "")
        changed, seen, last, last_change = False, [], before, time.time()
        while time.time() - t0 < timeout:
            time.sleep(0.5)
            d = [x for x in dialogs() if x.get("buttons") or x.get("error")]
            if d:
                seen.extend(d)
                last_change = time.time()
            sig = word_sig()
            if sig != last:
                last, last_change = sig, time.time()
            if sig != before and not sig.startswith("busy:"):
                changed = True
            if changed and time.time() - last_change >= quiet:
                break
            if not expect_change and time.time() - t0 >= 2 * quiet and time.time() - last_change >= quiet:
                break
        elapsed = round(time.time() - t0, 1)
        vt = [x for x in dialogs("VisualTeX", dismiss=False) if x.get("buttons")]
        step = {"label": label, "pressed": pressed, "changed": changed, "seconds": elapsed, "wordDialogs": seen, "visualtexDialogs": vt}
        self.steps.append(step)
        print(f"  · {label}: changed={changed} {elapsed}s" + (f" dialogs={json.dumps(seen, ensure_ascii=False)}" if seen else "") + (f" visualtex={json.dumps(vt, ensure_ascii=False)}" if vt else ""), flush=True)
        return step


def formulas_of(paras):
    return [f for p in paras for f in p["formulas"]]


def check_converted(case, paras, sources, kind, label):
    """sources: list of source paragraph texts in document order (one per paragraph)."""
    expected = [f for s in sources for f in source_formulas(s)]
    got = formulas_of(paras)
    case.check(f"{label}: {len(expected)} formulas became {kind}",
               len(got) == len(expected) and all(f["kind"] == kind for f in got),
               f"got {[(f['kind'], f['mode']) for f in got]}")
    modes_ok = [g["mode"] == e[0] for g, e in zip(got, expected)]
    case.check(f"{label}: inline/display kept", all(modes_ok) and len(got) == len(expected),
               "; ".join(f"{e[1][:30]} expected {e[0]} got {g['mode']}" for g, e, ok in zip(got, expected, modes_ok) if not ok))
    texts = [p["text"] for p in paras if p["text"].strip() or p["formulas"]]
    want = [source_marked(s) for s in sources]
    case.check(f"{label}: surrounding text unchanged", texts == want,
               "; ".join(f"{w!r} -> {t!r}" for w, t in zip(want, texts) if w != t) or f"{len(want)} vs {len(texts)} paragraphs")
    if kind == "image":
        pairs = [(e[1], g.get("latex")) for g, e in zip(got, expected)]
        case.check(f"{label}: image metadata keeps the source LaTeX exactly", not latex_diffs(pairs, False), show(latex_diffs(pairs, False)))
        ids = [g.get("id") for g in got]
        case.check(f"{label}: every image has its own formula id", len(set(ids)) == len(ids) and None not in ids, str(ids))
    return got


def check_restored(case, paras, sources, label, exact):
    """exact: the LaTeX must come back verbatim (it was stored in the image);
    otherwise only its meaning is compared and exact differences are noted."""
    texts = [p["text"] for p in paras if p["text"].strip() or p["formulas"]]
    left = formulas_of(paras)
    case.check(f"{label}: no formulas left", not left, f"{len(left)} left")
    case.check(f"{label}: text around the formulas unchanged", [source_marked(t) for t in texts] == [source_marked(s) for s in sources],
               show([(s, t) for s, t in zip(sources, texts) if source_marked(s) != source_marked(t)]) or f"{len(sources)} vs {len(texts)} paragraphs")
    want = [f for s in sources for f in source_formulas(s)]
    got = [f for t in texts for f in source_formulas(t)]
    case.check(f"{label}: inline/display delimiters kept", [m for m, _ in want] == [m for m, _ in got],
               show([(f"{a[0]}:{a[1]}", f"{b[0]}:{b[1]}") for a, b in zip(want, got) if a[0] != b[0]]))
    pairs = [(a[1], b[1]) for a, b in zip(want, got)]
    if exact:
        case.check(f"{label}: LaTeX comes back exactly as written", not latex_diffs(pairs, False) and len(want) == len(got), show(latex_diffs(pairs, False)))
    else:
        case.check(f"{label}: LaTeX keeps its meaning", not latex_diffs(pairs, True) and len(want) == len(got), show(latex_diffs(pairs, True)))
        cosmetic = [p for p in latex_diffs(pairs, False) if p not in latex_diffs(pairs, True)]
        if cosmetic:
            case.note(f"{label}: {len(cosmetic)} formulas rewritten without changing meaning, e.g. {show(cosmetic, 3)}")


def omml_signature(f):
    return (f["mode"], tuple(f["struct"]), norm_latex(f["text"]).replace("−", "-"))


# --------------------------------------------------------------------------
# fixtures


EVERYDAY = [
    r"已知速度 $v=\frac{s}{t}$，其中 $s$ 为路程，$t$ 为时间。",
    r"Pythagoras: $x^2+y^2=z^2$, and Euler: $e^{i\pi}+1=0$.",
    r"$$\int_0^1 x^2\,\mathrm{d}x=\frac{1}{3}$$",
    r"若 $a\le b\ne c$ 且 $\alpha+\beta=\gamma$，则成立。",
    r"\[\sum_{i=1}^{n} i=\frac{n(n+1)}{2}\]",
    r"Inline with parentheses: \(\sqrt[3]{x+1}\) and \(a_{ij}\).",
    r"$$f(x)=\begin{cases}x,&x\ge 0\\-x,&x<0\end{cases}$$",
    r"单位：$v=9.8\,\mathrm{m/s^2}$，水的化学式 $\mathrm{H_2O}$。",
    r"$$\begin{pmatrix}1&2\\3&4\end{pmatrix}$$",
    r"极限 $\lim_{x\to 0}\frac{\sin x}{x}=1$ 与 $\sin^2\theta+\cos^2\theta=1$。",
    r"$$x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}$$",
    r"向量 $\vec{F}=m\vec{a}$，线段 $\overline{AB}$，集合 $\mathbb{R}$ 与 $\{1,2\}$，绝对值 $|x|$。",
    r"$$\text{速度}=\frac{s}{t}$$",
    r"$$\begin{aligned}a&=b+c\\&=d\end{aligned}$$",
    r"$$\left(\frac{a}{b}\right)^2+\binom{n}{k}$$",
    r"导数 $f'(x)$ 与 $2\times 3\cdot 4\pm 1$。",
]

NATIVE = [
    ("inline", msup(mr("x"), mr("2")) + mr("+") + msup(mr("y"), mr("2")) + mr("=") + msup(mr("z"), mr("2"))),
    ("inline", mf(mr("a"), mr("b"))),
    ("inline", mrad(msup(mr("b"), mr("2")) + mr("−4ac"))),
    ("display", mnary("∫", mr("0"), mr("1"), msup(mr("x"), mr("2")) + mr("d") + mr("x")) + mr("=") + mf(mr("1"), mr("3"))),
    ("display", mnary("∑", mr("i=1"), mr("n"), mr("i"), und_ovr=True) + mr("=") + mf(mr("n") + md(mr("n+1")), mr("2"))),
    ("inline", mfunc(msup(mfn("sin"), mr("2")), mr("θ")) + mr("+") + mfunc(msup(mfn("cos"), mr("2")), mr("θ")) + mr("=1")),
    ("display", md(mmat([[mr("1"), mr("2")], [mr("3"), mr("4")]]))),
    ("inline", macc("⃗", mr("F")) + mr("=m") + macc("⃗", mr("a"))),
    ("display", mfunc(mlimlow(mfn("lim"), mr("x→0")), mf(mfunc(mfn("sin"), mr("x")), mr("x"))) + mr("=1")),
    ("inline", msub(mr("a"), mr("ij"))),
    ("inline", mbar(mr("AB"))),
    ("display", mr("速度", normal=True) + mr("=") + mf(mr("s"), mr("t"))),
]


def native_blocks():
    blocks = []
    for i, (mode, xml) in enumerate(NATIVE, 1):
        if mode == "inline":
            blocks.append(para_xml([f"公式 {i}：", ("m", xml), "。"]))
        else:
            blocks.append(para_xml([("mp", xml)]))
    return blocks


# --------------------------------------------------------------------------
# cases


def fixture(case, name, blocks):
    path = os.path.join(OUT, f"{case.id}-{name}.docx")
    write_docx(path, blocks)
    word_close_all()
    word_open(path)
    time.sleep(1)
    return path


def case_everyday_image_restore(c):
    fixture(c, "source", [para_xml(s) for s in EVERYDAY])
    c.act("redraw whole document to images", REDRAW_DOC_IMAGE, timeout=120)
    check_converted(c, c.snapshot("images"), EVERYDAY, "image", "redraw to images")
    c.act("restore all image formulas to LaTeX", RESTORE_DOC_IMAGE, timeout=120)
    check_restored(c, c.snapshot("latex"), EVERYDAY, "images straight back to LaTeX", exact=True)


def case_everyday_image_roundtrip(c):
    fixture(c, "source", [para_xml(s) for s in EVERYDAY])
    c.act("redraw whole document to images", REDRAW_DOC_IMAGE, timeout=120)
    c.snapshot("images")
    c.act("all image formulas to OMML", DOC_IMAGES_TO_OMML, timeout=120)
    paras = c.snapshot("omml")
    got = formulas_of(paras)
    n = len([f for s in EVERYDAY for f in source_formulas(s)])
    c.check("images to OMML: every formula is OMML", len(got) == n and all(f["kind"] == "omml" for f in got), f"{[(f['kind']) for f in got]}")
    c.check("images to OMML: inline/display kept", [f["mode"] for f in got] == [m for s in EVERYDAY for m, _ in source_formulas(s)],
            str([f["mode"] for f in got]))
    c.check("images to OMML: surrounding text unchanged",
            [p["text"] for p in paras if p["text"].strip() or p["formulas"]] == [source_marked(s) for s in EVERYDAY])
    c.act("all OMML to image formulas", DOC_OMML_TO_IMAGES, timeout=120)
    paras = c.snapshot("images-again")
    got = formulas_of(paras)
    c.check("OMML back to images: every formula is an image", len(got) == n and all(f["kind"] == "image" for f in got))
    pairs = [(e, g.get("latex")) for (m, e), g in zip([x for s in EVERYDAY for x in source_formulas(s)], got)]
    c.check("image→OMML→image keeps each formula's meaning", not latex_diffs(pairs, True), show(latex_diffs(pairs, True)))
    c.act("restore all image formulas to LaTeX", RESTORE_DOC_IMAGE, timeout=120)
    check_restored(c, c.snapshot("latex"), EVERYDAY, "after the OMML round trip, restore to LaTeX", exact=False)


def case_everyday_omml_roundtrip(c):
    fixture(c, "source", [para_xml(s) for s in EVERYDAY])
    c.act("redraw whole document to OMML", REDRAW_DOC_OMML, timeout=120)
    check_converted(c, c.snapshot("omml"), EVERYDAY, "omml", "redraw to OMML")
    c.act("restore all OMML to LaTeX", RESTORE_DOC_OMML, timeout=120)
    check_restored(c, c.snapshot("latex"), EVERYDAY, "restore OMML to LaTeX", exact=False)


def case_native_equations(c):
    fixture(c, "source", native_blocks())
    original = formulas_of(read_docx(os.path.join(OUT, f"{c.id}-source.docx")))
    c.act("Word equations to image formulas", DOC_OMML_TO_IMAGES, timeout=120)
    got = formulas_of(c.snapshot("images"))
    c.check(f"{len(NATIVE)} Word equations became images", len(got) == len(NATIVE) and all(f["kind"] == "image" for f in got),
            str([f["kind"] for f in got]))
    c.check("inline/display kept", [f["mode"] for f in got] == [m for m, _ in NATIVE], str([f["mode"] for f in got]))
    c.check("every image carries LaTeX", all(f.get("latex") for f in got))
    for f in got:
        c.note(f"LaTeX from Word equation: {f.get('latex')!r}")
    c.act("image formulas back to Word equations", DOC_IMAGES_TO_OMML, timeout=120)
    back = formulas_of(c.snapshot("omml"))
    diffs = [(omml_signature(a), omml_signature(b)) for a, b in zip(original, back) if omml_signature(a) != omml_signature(b)]
    c.check("Word equation → image → Word equation keeps structure and text", len(back) == len(original) and not diffs,
            "; ".join(f"{a} -> {b}" for a, b in diffs[:6]) or f"{len(original)} vs {len(back)}")


SELECTION = [
    r"第一段 $a+b$ 不应改变。",
    r"第二段 $\frac{1}{2}$ 与 $x^2$ 需要重绘。",
    r"第三段 $c-d$ 不应改变。",
]


def case_selection_scope(c):
    fixture(c, "source", [para_xml(s) for s in SELECTION])
    word_select_paragraph(2)
    c.act("redraw selected paragraph to images", REDRAW_SEL_IMAGE)
    paras = c.snapshot("sel-images")
    c.check("only the selected paragraph changed",
            [p["text"] for p in paras] == [SELECTION[0], source_marked(SELECTION[1]), SELECTION[2]],
            str([p["text"] for p in paras]))
    c.check("selected paragraph has 2 image formulas", [f["kind"] for f in paras[1]["formulas"]] == ["image", "image"])
    word_select_shape(1)
    c.act("convert the selected image to OMML", SEL_TO_OMML)
    paras = c.snapshot("one-omml")
    c.check("only the selected image became OMML", [f["kind"] for f in paras[1]["formulas"]] == ["omml", "image"],
            str([f["kind"] for f in paras[1]["formulas"]]))
    word_select_math(1)
    c.act("convert the selected OMML back to an image", SEL_TO_IMAGE)
    paras = c.snapshot("back-image")
    c.check("selected OMML became an image again", [f["kind"] for f in paras[1]["formulas"]] == ["image", "image"],
            str([f["kind"] for f in paras[1]["formulas"]]))
    c.check("…with its LaTeX", [canon_latex(f.get("latex")) for f in paras[1]["formulas"]] == [r"\frac{1}{2}", "x^2"],
            str([f.get("latex") for f in paras[1]["formulas"]]))
    word_select_paragraph(2)
    c.act("restore the selected paragraph to LaTeX", RESTORE_SEL_IMAGE)
    paras = c.snapshot("sel-latex")
    c.check("selected paragraph restored to LaTeX", [canon_latex(p["text"]) for p in paras] == [canon_latex(s) for s in SELECTION], str([p["text"] for p in paras]))
    word_select_paragraph(1)
    c.act("redraw selected paragraph to OMML", REDRAW_SEL_OMML)
    paras = c.snapshot("sel-omml")
    c.check("only paragraph 1 became OMML", [len(p["formulas"]) for p in paras] == [1, 0, 0] and paras[0]["formulas"][0]["kind"] == "omml",
            str([[f["kind"] for f in p["formulas"]] for p in paras]))
    word_select_paragraph(1)
    c.act("restore selected OMML to LaTeX", RESTORE_SEL_OMML)
    paras = c.snapshot("sel-omml-latex")
    c.check("paragraph 1 restored", canon_latex(paras[0]["text"]) == canon_latex(SELECTION[0]), repr(paras[0]["text"]))


STRUCTURE_TABLE = [[r"$a^2$", r"面积 $S=\pi r^2$"], [r"$\frac{1}{x}$", r"$$\sqrt{2}$$"]]
STRUCTURE_LIST = [r"条件一：$x>0$", r"条件二：$y\in\mathbb{R}$"]
STRUCTURE_HEADING = r"1.2 匀速运动 $v=\frac{s}{t}$"
STRUCTURE_BOLD = ("b", r"定理：$a^2+b^2=c^2$ 成立。")


def structure_blocks():
    return [
        para_xml(STRUCTURE_HEADING, style="Heading1"),
        para_xml([STRUCTURE_BOLD]),
        para_xml(STRUCTURE_LIST[0], numbered=1),
        para_xml(STRUCTURE_LIST[1], numbered=1),
        table_xml(STRUCTURE_TABLE),
        para_xml(r"表后 $e^x$ 结束。"),
    ]


STRUCTURE_SOURCES = [STRUCTURE_HEADING, STRUCTURE_BOLD[1], *STRUCTURE_LIST, *[cell for row in STRUCTURE_TABLE for cell in row], r"表后 $e^x$ 结束。"]
STRUCTURE_CTX = ["body/heading", "body", "body/list", "body/list", "table", "table", "table", "table", "body"]


def case_structures(c):
    fixture(c, "source", structure_blocks())
    c.act("redraw whole document to images", REDRAW_DOC_IMAGE, timeout=90)
    paras = [p for p in c.snapshot("images") if p["text"].strip() or p["formulas"]]
    c.check("heading, list, table and body keep their place", [p["ctx"] for p in paras] == STRUCTURE_CTX, str([p["ctx"] for p in paras]))
    check_converted(c, paras, STRUCTURE_SOURCES, "image", "redraw in heading/list/table")
    c.act("all image formulas to OMML", DOC_IMAGES_TO_OMML, timeout=90)
    paras = [p for p in c.snapshot("omml") if p["text"].strip() or p["formulas"]]
    got = formulas_of(paras)
    c.check("image→OMML reaches heading, list and table", len(got) == 9 and all(f["kind"] == "omml" for f in got),
            str([(p["ctx"], [f["kind"] for f in p["formulas"]]) for p in paras]))
    c.act("restore all OMML to LaTeX", RESTORE_DOC_OMML, timeout=90)
    check_restored(c, [p for p in c.snapshot("latex") if p["text"].strip() or p["formulas"]], STRUCTURE_SOURCES, "restore in heading/list/table", exact=False)


def case_save_reopen(c):
    sources = EVERYDAY[:6]
    fixture(c, "source", [para_xml(s) for s in sources])
    c.act("redraw whole document to images", REDRAW_DOC_IMAGE, timeout=90)
    c.snapshot("images")
    saved = os.path.join(OUT, f"{c.id}-saved.docx")
    word_save_as(saved)
    word_close_all()
    word_open(saved)
    time.sleep(1)
    c.act("after reopening: images to OMML", DOC_IMAGES_TO_OMML, timeout=90)
    paras = c.snapshot("reopened-omml")
    c.check("reopened document converts every image", len(formulas_of(paras)) == sum(len(source_formulas(s)) for s in sources)
            and all(f["kind"] == "omml" for f in formulas_of(paras)))
    word_save_as(saved)
    word_close_all()
    word_open(saved)
    time.sleep(1)
    c.act("after reopening: OMML to images", DOC_OMML_TO_IMAGES, timeout=90)
    c.act("restore images to LaTeX", RESTORE_DOC_IMAGE, timeout=90)
    check_restored(c, c.snapshot("reopened-latex"), sources, "after two save/reopen cycles", exact=False)


def case_copy_paste(c, clip_tool):
    if not clip_tool:
        c.note("skipped: no --clipboard-tool to save/restore the clipboard")
        return
    sources = [r"原式 $\frac{a+b}{2}$ 与 $\sqrt{x}$。", "复制到这里："]
    fixture(c, "source", [para_xml(s) for s in sources])
    c.act("redraw whole document to images", REDRAW_DOC_IMAGE)
    c.act("copy formula 1 and paste it at the end", action=lambda: word_copy_shape_to_end(1), quiet=1.5)
    paras = c.snapshot("pasted")
    got = formulas_of(paras)
    c.check("pasted copy is an image formula with the same LaTeX", len(got) == 3 and norm_latex(got[2].get("latex")) == r"\frac{a+b}{2}",
            str([(f["kind"], f.get("latex")) for f in got]))
    c.note(f"formula ids after paste: {[f.get('id') for f in got]}")
    c.act("all image formulas to OMML", DOC_IMAGES_TO_OMML)
    got = formulas_of(c.snapshot("omml"))
    c.check("original and pasted copy both convert to OMML", len(got) == 3 and all(f["kind"] == "omml" for f in got), str([f["kind"] for f in got]))
    c.act("all OMML back to images", DOC_OMML_TO_IMAGES)
    got = formulas_of(c.snapshot("images"))
    ids = [f.get("id") for f in got]
    c.check("after the round trip every formula has its own id", len(got) == 3 and len(set(ids)) == 3, str(ids))


def case_undo(c):
    sources = [r"撤销测试 $x+y$ 与 $\frac{1}{2}$。", r"$$E=mc^2$$"]
    fixture(c, "source", [para_xml(s) for s in sources])
    c.act("redraw whole document to images", REDRAW_DOC_IMAGE)
    before_undo = word_sig()
    undone = 0
    for i in range(1, 9):
        word_undo(1)
        undone = i
        time.sleep(0.3)
        sig = word_sig()
        if "$" in sig and sig.split("|")[0] == "0":
            break
    c.note(f"undo steps to get the LaTeX back: {undone}")
    paras = c.snapshot("undone")
    texts = [p["text"] for p in paras if p["text"].strip() or p["formulas"]]
    c.check("Undo brings back the LaTeX text", texts == sources, f"{texts} (sig before undo {before_undo[:40]!r})")
    c.check("one Undo is enough", undone == 1, f"needed {undone}")


# Each example is its own document: (paragraphs, expected paragraphs with ⟦F⟧, expected LaTeX).
SAFETY = [
    (["门票 $5，套餐 $10。"], ["门票 $5，套餐 $10。"], []),
    (["It costs $5 and $10 per person."], ["It costs $5 and $10 per person."], []),
    (["Price: US$ 20, then $x$ is the unknown."], ["Price: US$ 20, then ⟦F⟧ is the unknown."], ["x"]),
    (["Run export PATH=$PATH:$HOME/bin first."], ["Run export PATH=$PATH:$HOME/bin first."], []),
    (["增长 50%，利润率 12.5%。"], ["增长 50%，利润率 12.5%。"], []),
    (["Tickets cost $5 each.", "The energy is $E=mc^2$."], ["Tickets cost $5 each.", "The energy is ⟦F⟧."], ["E=mc^2"]),
    (["单价 $5。", "公式 $a+b$ 在此。"], ["单价 $5。", "公式 ⟦F⟧ 在此。"], ["a+b"]),
    (["未写完的 $a", "b$ 结束。"], ["未写完的 $a", "b$ 结束。"], []),
]


def case_text_safety(c):
    for i, (sources, want, latex) in enumerate(SAFETY, 1):
        fixture(c, f"source{i}", [para_xml(s) for s in sources])
        c.act(f"redraw example {i}", REDRAW_DOC_IMAGE, timeout=30, expect_change=bool(latex))
        paras = [p for p in c.snapshot(f"example{i}") if p["text"].strip() or p["formulas"]]
        got_latex = [f.get("latex") for f in formulas_of(paras)]
        texts = [p["text"] for p in paras]
        c.check(f"text kept: {' ¶ '.join(sources)!r}", texts == want and [norm_latex(x) for x in got_latex] == [norm_latex(x) for x in latex],
                f"became {' ¶ '.join(texts)!r} with formulas {got_latex}")
        word_close_all()


def case_noop_commands(c):
    plain = ["只有文字，没有公式。", "Plain text only."]
    fixture(c, "source", [para_xml(s) for s in plain])
    for label, path in [("redraw whole document with no LaTeX", REDRAW_DOC_IMAGE),
                        ("images to OMML with no images", DOC_IMAGES_TO_OMML),
                        ("OMML to images with no OMML", DOC_OMML_TO_IMAGES),
                        ("restore images with no images", RESTORE_DOC_IMAGE)]:
        step = c.act(label, path, timeout=15, expect_change=False)
        c.check(f"{label}: document unchanged", not step["changed"])
        c.note(f"{label}: Word dialogs {json.dumps(step['wordDialogs'], ensure_ascii=False)}")
    paras = c.snapshot("after")
    c.check("text intact after no-op commands", [p["text"] for p in paras if p["text"].strip()] == plain)


def case_repeat_redraw(c):
    sources = EVERYDAY[:4]
    fixture(c, "source", [para_xml(s) for s in sources])
    c.act("redraw whole document to images", REDRAW_DOC_IMAGE)
    first = c.snapshot("first")
    step = c.act("redraw whole document again", REDRAW_DOC_IMAGE, timeout=20, expect_change=False)
    second = c.snapshot("second")
    c.check("second redraw changes nothing", [p["text"] for p in first] == [p["text"] for p in second]
            and [f.get("id") for f in formulas_of(first)] == [f.get("id") for f in formulas_of(second)], f"changed={step['changed']}")
    c.act("redraw images to OMML while images exist", REDRAW_DOC_OMML, timeout=20, expect_change=False)
    third = c.snapshot("third")
    c.note(f"redraw-to-OMML over existing images: kinds {[f['kind'] for f in formulas_of(third)]}")


def case_mixed_formats(c):
    sources = [r"图片 $a^2$ 与 OMML $b^2$ 混排。", r"$$\frac{1}{2}$$", r"再一个 $c_1$。"]
    fixture(c, "source", [para_xml(s) for s in sources])
    word_select_paragraph(1)
    c.act("paragraph 1 to images", REDRAW_SEL_IMAGE)
    word_select_paragraph(2)
    c.act("paragraph 2 to OMML", REDRAW_SEL_OMML)
    word_select_paragraph(3)
    c.act("paragraph 3 to images", REDRAW_SEL_IMAGE)
    kinds = [f["kind"] for f in formulas_of(c.snapshot("mixed"))]
    c.check("document now mixes images and OMML", kinds == ["image", "image", "omml", "image"], str(kinds))
    c.act("all images to OMML", DOC_IMAGES_TO_OMML)
    kinds = [f["kind"] for f in formulas_of(c.snapshot("all-omml"))]
    c.check("mixed document all OMML", kinds == ["omml"] * 4, str(kinds))
    c.act("all OMML to images", DOC_OMML_TO_IMAGES)
    got = formulas_of(c.snapshot("all-images"))
    c.check("mixed document all images", [f["kind"] for f in got] == ["image"] * 4, str([f["kind"] for f in got]))
    c.check("LaTeX survives the mixed round trip", [canon_latex(f.get("latex")) for f in got] == ["a^2", "b^2", r"\frac{1}{2}", "c_1"],
            str([f.get("latex") for f in got]))


def case_large_document(c):
    sources = [rf"第 {i} 行：$x_{{{i}}}^2+\frac{{{i}}}{{{i + 1}}}=y_{{{i}}}$，以及 $\sqrt{{{i}}}$。" for i in range(1, 51)]
    fixture(c, "source", [para_xml(s) for s in sources])
    step = c.act("redraw 100 formulas to images", REDRAW_DOC_IMAGE, timeout=300, quiet=4)
    paras = c.snapshot("images")
    got = formulas_of(paras)
    c.check("100 formulas became images", len(got) == 100 and all(f["kind"] == "image" for f in got), f"{len(got)} formulas")
    c.note(f"redraw 100 formulas: {step['seconds']} s")
    step = c.act("100 images to OMML", DOC_IMAGES_TO_OMML, timeout=300, quiet=4)
    got = formulas_of(c.snapshot("omml"))
    c.check("100 images became OMML", len(got) == 100 and all(f["kind"] == "omml" for f in got), f"{len(got)}")
    c.note(f"images to OMML for 100 formulas: {step['seconds']} s")


def case_native_restore_all_or_nothing(c):
    simple = EVERYDAY[:2]
    fixture(c, "simple", [para_xml(s) for s in simple])
    c.act("simple formulas: redraw to OMML", REDRAW_DOC_OMML)
    c.act("simple formulas: restore OMML to LaTeX", RESTORE_DOC_OMML)
    check_restored(c, c.snapshot("simple-latex"), simple, "VisualTeX-made OMML, simple formulas only", exact=True)
    word_close_all()
    mixed = simple + [r"$$\begin{aligned}a&=b+c\\&=d\end{aligned}$$"]
    fixture(c, "plus-aligned", [para_xml(s) for s in mixed])
    c.act("same plus one aligned block: redraw to OMML", REDRAW_DOC_OMML)
    c.act("same plus one aligned block: restore OMML to LaTeX", RESTORE_DOC_OMML)
    paras = [p for p in c.snapshot("plus-aligned-latex") if p["text"].strip()]
    first = [p["text"] for p in paras[:2]]
    c.check("one aligned block does not change how the other formulas restore",
            [norm_latex(t) for t in first] == [canonical_delimiters(s) for s in simple], show(list(zip(simple, first))))


CASES = [
    ("W01", "Everyday mixed document: LaTeX → images → LaTeX", case_everyday_image_restore),
    ("W02", "Everyday mixed document: LaTeX → images → OMML → images → LaTeX", case_everyday_image_roundtrip),
    ("W03", "Everyday mixed document: LaTeX → OMML → LaTeX", case_everyday_omml_roundtrip),
    ("W04", "Equations made in Word's own equation editor → images → equations", case_native_equations),
    ("W05", "Selection commands only touch the selection", case_selection_scope),
    ("W06", "Formulas in a heading, bold text, a list and a table", case_structures),
    ("W07", "Save, close and reopen between conversions", case_save_reopen),
    ("W08", "Copy and paste a formula inside Word, then convert", case_copy_paste),
    ("W09", "Undo after a whole-document redraw", case_undo),
    ("W10", "Everyday text with $ signs is not turned into formulas", case_text_safety),
    ("W11", "Commands on a document with nothing to convert", case_noop_commands),
    ("W12", "Running a redraw twice", case_repeat_redraw),
    ("W13", "Images and OMML mixed in one document", case_mixed_formats),
    ("W14", "100 formulas in one document", case_large_document),
    ("W15", "Restoring VisualTeX-made OMML: one odd formula vs. the rest", case_native_restore_all_or_nothing),
]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--only", default="")
    ap.add_argument("--keep", action="store_true", help="keep the test folder in the Office group container")
    ap.add_argument("--results", default=os.path.join(os.getcwd(), "office-everyday-results"))
    ap.add_argument("--clipboard-tool", default="")
    args = ap.parse_args()
    if args.list:
        for cid, title, _ in CASES:
            print(cid, title)
        return 0
    only = {x.strip() for x in args.only.split(",") if x.strip()}
    os.makedirs(OUT, mode=0o700, exist_ok=True)
    os.makedirs(args.results, exist_ok=True)
    was_running = word_running()
    clip_dir = os.path.join(args.results, "clipboard-backup")
    clip_tool = args.clipboard_tool if args.clipboard_tool and os.access(args.clipboard_tool, os.X_OK) else ""
    if clip_tool:
        subprocess.run([clip_tool, "save", clip_dir], check=True)
    report = []
    try:
        for cid, title, fn in CASES:
            if only and cid not in only:
                continue
            print(f"\n== {cid} {title}", flush=True)
            case = Case(cid, title, args.results)
            started = time.time()
            try:
                if fn is case_copy_paste:
                    fn(case, clip_tool)
                else:
                    fn(case)
            except Exception as error:
                case.check("case ran to the end", False, f"{type(error).__name__}: {error}")
            finally:
                word_close_all()
            report.append({"id": cid, "title": title, "seconds": round(time.time() - started, 1),
                           "checks": case.checks, "notes": case.notes, "steps": case.steps})
    finally:
        word_close_all()
        if clip_tool:
            subprocess.run([clip_tool, "restore", clip_dir])
        if not was_running and word_running():
            try:
                osa('tell application "Microsoft Word" to quit saving no')
            except Exception:
                pass
        if not args.keep:
            shutil.rmtree(OUT, ignore_errors=True)
        with open(os.path.join(args.results, "report.json"), "w") as f:
            json.dump(report, f, ensure_ascii=False, indent=1)
    total = sum(len(r["checks"]) for r in report)
    failed = [(r["id"], ch["name"], ch["detail"]) for r in report for ch in r["checks"] if not ch["ok"]]
    print(f"\n{total - len(failed)}/{total} checks passed")
    for cid, name, detail in failed:
        print(f"  FAIL {cid} {name}" + (f" — {detail}" if detail else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
